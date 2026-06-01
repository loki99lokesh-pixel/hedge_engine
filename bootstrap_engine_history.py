"""
bootstrap_engine_history.py
============================
Pre-computes Engine_Hedge_Target and Engine_Stage for every row in
MARKET_DATA.csv from 2015-01-01 onwards, then writes them back to the CSV.

Run this script:
  - Once initially to populate the columns
  - After appending new rows to MARKET_DATA.csv (incremental — only
    processes rows where Engine_Hedge_Target is NaN/empty)

Usage:
    python bootstrap_engine_history.py
    python bootstrap_engine_history.py --full     # force full recompute from 2015
    python bootstrap_engine_history.py --from 2020-01-01  # recompute from a date

The script saves a checkpoint file (engine_state_checkpoint.json) alongside
the CSV. This checkpoint stores the engine state at the last computed row so
that incremental runs can resume exactly from where the previous run ended
without re-processing the entire history.

Architecture:
  CSV  → permanent historical record (committed to repo)
  Redis → live buffer: engine state + today's chart points (ephemeral)

On server boot, backend.py reads Engine_Hedge_Target / Engine_Stage directly
from the CSV — no engine replay needed. Redis is only used for live daily
appends beyond the last CSV date.
"""

import os
import sys
import json
import argparse
import math
from datetime import datetime

import numpy as np
import pandas as pd

# ── Locate files ────────────────────────────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
CSV_PATH     = os.path.join(SCRIPT_DIR, 'MARKET_DATA.csv')
CHECKPOINT   = os.path.join(SCRIPT_DIR, 'engine_state_checkpoint.json')
START_YEAR   = 2015   # rows before this year are left blank (pre-VIX era)

# ── Engine defaults (mirrors backend.py _V3_STATE_DEFAULTS) ────────────────
ENGINE_DEFAULTS = {
    'peak_price'         : 0.0,
    'trough_price'       : float('inf'),
    'days_in_dd'         : 0,
    'low_vol_days'       : 0,
    'prev_hedge'         : 0.0,
    'last_date'          : '',
    'ema_vix'            : 18.5,
    'sod_dd_pct'         : 0.0,
    'sod_natural_target' : 0.0,
    'sod_buffer'         : 5.0,
    'sod_recovery_ratio' : 0.0,
}

# ── Inline engine (copy of portfolio_hedge_engine.py core) ──────────────────
# Duplicated here so this script has zero import dependencies beyond pandas/numpy.

def _sigmoid(x):
    return 1.0 / (1.0 + math.exp(-x))

def _get_continuous_notional(score):
    n  = 10.0
    n += 30.0 * _sigmoid((score - 25.0) / 3.5)
    n += 30.0 * _sigmoid((score - 50.0) / 3.5)
    n += 30.0 * _sigmoid((score - 70.0) / 3.5)
    return min(100.0, max(10.0, n))

def _norm_fpi(fpi):
    if fpi < 0:
        return min(100.0, (-fpi) / 8000.0 * 100.0)
    else:
        return max(-50.0, (-fpi) / 8000.0 * 100.0)

def _run_engine_day(vix, rv20d, fpi_net, gap_pct, current_price,
                    state, ret_5d=0.0, new_calendar_day=True):
    """
    Single-day engine evaluation.
    Returns (hedge_target, active_stage, new_state).
    Mirrors calculate_v3_magnitude_hedge() in portfolio_hedge_engine.py v3.2.
    """
    # VIX EMA smoothing
    alpha        = 0.5
    prev_ema_vix = state.get('ema_vix', vix)
    vix          = vix * alpha + prev_ema_vix * (1 - alpha)

    # 1. Peak / trough / drawdown
    peak    = max(state.get('peak_price', current_price), current_price)
    dd_pct  = abs((current_price - peak) / peak) * 100 if peak > 0 else 0.0
    trough  = min(state.get('trough_price', current_price), current_price)
    if dd_pct == 0:
        trough = current_price

    days_in_dd = (
        state.get('days_in_dd', 0) + (1 if new_calendar_day else 0)
        if dd_pct > 0 else 0
    )

    # 2. Stage 1 onset score
    rvN  = min(100.0, max(0.0, (rv20d  -  5.0) / 25.0 * 100.0))
    gapN = min(100.0, max(0.0, (-gap_pct)       / 20.0 * 100.0))
    fpiN = _norm_fpi(fpi_net)
    vixN = min(100.0, max(0.0, (vix    - 10.0)  / 30.0 * 100.0))

    msf_active = (rv20d > 18.0 and ret_5d > -3.0)
    fpi_weight = 0.10 * (0.8 if msf_active else 1.0)

    onset_score = 0.48*rvN + 0.25*gapN + fpi_weight*fpiN + 0.17*vixN
    s1_target   = _get_continuous_notional(onset_score)

    # 3. Stage 2 active phase
    penalty   = min(0.08 * days_in_dd, 15.0)
    s2_target = 12.5 + 0.52*vix + 0.48*dd_pct + 5.10*(vix*dd_pct/100.0) - penalty
    s2_target = min(90.0, max(10.0, s2_target))

    # 4. Transition
    final_target = max(s1_target, s2_target)
    active_stage = "Stage 1 — Onset" if s1_target >= s2_target else "Stage 2 — Active Phase"

    # Recovery ratio
    recovery_ratio_live = 0.0
    if peak - trough > 0:
        recovery_ratio_live = (current_price - trough) / (peak - trough)

    # SOD snapshots
    if new_calendar_day or 'sod_natural_target' not in state:
        sod_natural_target  = final_target
        sod_buffer          = 5.0 + max(0.0, ((vix - 15.0) / 10.0) * 3.0)
        sod_dd_pct          = dd_pct
        sod_recovery_ratio  = recovery_ratio_live
    else:
        sod_natural_target  = state.get('sod_natural_target', final_target)
        sod_buffer          = state.get('sod_buffer', 5.0)
        sod_dd_pct          = state.get('sod_dd_pct', dd_pct)
        sod_recovery_ratio  = state.get('sod_recovery_ratio', recovery_ratio_live)

    # 5. Stage 3 escalation (only after 3 days in drawdown)
    if days_in_dd >= 3:
        if dd_pct >= 20:
            final_target = max(final_target, 75.0)
            active_stage = "Stage 3: Escalated (MAJOR minimum)"
        if dd_pct >= 15 and (rv20d > 28 or vix > 28):
            final_target = 100.0
            active_stage = "Stage 3: Escalated (SEVERE)"
        if dd_pct >= 25:
            final_target = 100.0
            active_stage = "Stage 3: Escalated (SEVERE unconditional)"

    # 6. De-escalation gate
    low_vol_days = state.get('low_vol_days', 0)
    if new_calendar_day:
        low_vol_days = low_vol_days + 1 if (rv20d < 22 and vix < 20) else 0

    incoming_prev_hedge = state.get('prev_hedge', 0)
    if dd_pct == 0:
        incoming_prev_hedge = 0
        low_vol_days        = 0
        days_in_dd          = 0

    final_target       = round(final_target, 1)
    gate_threshold_check = sod_natural_target + sod_buffer

    if incoming_prev_hedge > gate_threshold_check and "Stage 3" not in active_stage:
        if sod_dd_pct >= 2.0:
            cleared_by_price = sod_recovery_ratio >= 0.40
            cleared_by_vol   = low_vol_days >= 5
            gate_cleared     = (cleared_by_price or cleared_by_vol) and days_in_dd >= 3
            if not gate_cleared:
                final_target = incoming_prev_hedge
                active_stage = "Stage 3: De-escalation Blocked (Gate closed)"

    new_prev_hedge = max(incoming_prev_hedge, final_target) if dd_pct > 0 else 0.0

    new_state = {
        'peak_price'         : peak,
        'trough_price'       : trough,
        'days_in_dd'         : days_in_dd,
        'low_vol_days'       : low_vol_days,
        'prev_hedge'         : new_prev_hedge,
        'ema_vix'            : vix,
        'sod_dd_pct'         : sod_dd_pct,
        'sod_natural_target' : sod_natural_target,
        'sod_buffer'         : sod_buffer,
        'sod_recovery_ratio' : sod_recovery_ratio,
    }
    return final_target, active_stage, new_state


# ── Date parsing ─────────────────────────────────────────────────────────────
def _parse_dates(series):
    for fmt in ['%d-%m-%Y', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
        try:
            parsed = pd.to_datetime(series, format=fmt)
            if parsed.notna().sum() > len(series) * 0.9:
                return parsed
        except Exception:
            continue
    return pd.to_datetime(series, infer_datetime_format=True, errors='coerce')


# ── Checkpoint helpers ───────────────────────────────────────────────────────
def load_checkpoint():
    if not os.path.exists(CHECKPOINT):
        return None
    try:
        with open(CHECKPOINT, 'r') as f:
            data = json.load(f)
        if data.get('trough_price', 0) >= 1e14:
            data['trough_price'] = float('inf')
        print(f"[Checkpoint] Loaded — last date: {data.get('last_date', '?')}")
        return data
    except Exception as e:
        print(f"[Checkpoint] Failed to load: {e}")
        return None

def save_checkpoint(state, last_date):
    data = dict(state)
    data['last_date'] = last_date
    if data.get('trough_price') == float('inf'):
        data['trough_price'] = 1e15
    with open(CHECKPOINT, 'w') as f:
        json.dump(data, f, indent=2)
    print(f"[Checkpoint] Saved — last date: {last_date}")


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description='Pre-compute engine history into MARKET_DATA.csv')
    parser.add_argument('--full',  action='store_true',
                        help='Force full recompute from START_YEAR even if columns exist')
    parser.add_argument('--from',  dest='from_date', default=None,
                        help='Recompute from this date (YYYY-MM-DD), discarding checkpoint')
    args = parser.parse_args()

    if not os.path.exists(CSV_PATH):
        print(f"[ERROR] CSV not found: {CSV_PATH}")
        sys.exit(1)

    # ── Load CSV ──────────────────────────────────────────────────────────────
    print(f"[Load] Reading {CSV_PATH} …")
    df = pd.read_csv(CSV_PATH, low_memory=False)
    df['Date'] = _parse_dates(df['Date'])
    df = df.dropna(subset=['Date']).sort_values('Date').reset_index(drop=True)

    # Ensure engine columns exist
    if 'Engine_Hedge_Target' not in df.columns:
        df['Engine_Hedge_Target'] = float('nan')
    if 'Engine_Stage' not in df.columns:
        df['Engine_Stage'] = ''

    print(f"[Load] {len(df)} rows  {df['Date'].iloc[0].date()} → {df['Date'].iloc[-1].date()}")

    # ── Determine start row ───────────────────────────────────────────────────
    force_from = None
    if args.from_date:
        force_from = pd.to_datetime(args.from_date)
        print(f"[Mode] Recompute from {force_from.date()} (--from flag)")
    elif args.full:
        print(f"[Mode] Full recompute from {START_YEAR} (--full flag)")
    else:
        print(f"[Mode] Incremental — filling missing Engine_Hedge_Target rows only")

    # Rows eligible for computation: 2015 onwards (VIX data available)
    cutoff_date = force_from if force_from else pd.Timestamp(f'{START_YEAR}-01-01')

    # For incremental: find first row with missing target on or after cutoff
    if not args.full and force_from is None:
        mask_missing = (
            df['Date'] >= cutoff_date
        ) & (
            df['Engine_Hedge_Target'].isna() | (df['Engine_Hedge_Target'] == '')
        )
        if not mask_missing.any():
            print("[Done] All rows already computed. Nothing to do.")
            print("       Add new rows to the CSV and re-run to process them.")
            return
        first_missing_date = df.loc[mask_missing, 'Date'].iloc[0]
        # Roll back 30 rows before first missing to give engine state warm-up
        first_missing_idx = df.index[df['Date'] == first_missing_date][0]
        warmup_start_idx  = max(0, first_missing_idx - 30)
        compute_from_date = df.loc[warmup_start_idx, 'Date']
        print(f"[Incremental] First missing: {first_missing_date.date()} | "
              f"Warm-up from: {compute_from_date.date()}")
    else:
        compute_from_date = cutoff_date

    # ── Load engine state ─────────────────────────────────────────────────────
    if args.full or force_from is not None:
        # Discard checkpoint — fresh state from compute_from_date
        state = dict(ENGINE_DEFAULTS)
        state['trough_price'] = float('inf')
        print(f"[State] Starting fresh engine state from {compute_from_date.date()}")
    else:
        # Try checkpoint first (only valid if checkpoint.last_date < compute_from_date)
        ckpt = load_checkpoint()
        if (ckpt and ckpt.get('last_date')
                and pd.to_datetime(ckpt['last_date']) < compute_from_date):
            state = ckpt
            print(f"[State] Resuming from checkpoint: {ckpt.get('last_date')}")
        else:
            state = dict(ENGINE_DEFAULTS)
            state['trough_price'] = float('inf')
            print(f"[State] No valid checkpoint — starting fresh")

    # ── Slice rows to process ─────────────────────────────────────────────────
    process_mask = df['Date'] >= compute_from_date
    process_df   = df[process_mask].copy()
    print(f"[Engine] Processing {len(process_df)} rows "
          f"({compute_from_date.date()} → {df['Date'].iloc[-1].date()}) …")

    # ── Replay ───────────────────────────────────────────────────────────────
    rolling_peak  = state.get('peak_price', 0.0)
    targets       = {}   # date → (hedge_target, stage)
    rows_computed = 0
    last_date_str = ''

    for _, row in process_df.iterrows():
        price = row['Nifty_Close']
        if pd.isna(price) or price <= 0:
            continue

        price    = float(price)
        vix_val  = float(row['VIX'])              if not pd.isna(row['VIX'])              else 18.0
        rv20d    = float(row['Nifty_RealVol20d']) if not pd.isna(row['Nifty_RealVol20d']) else vix_val * 0.75
        fpi_daily= float(row['FPI_Net_Equity_Cr'])if not pd.isna(row['FPI_Net_Equity_Cr'])else -1200.0
        fpi_weekly = fpi_daily * 5

        rolling_peak = max(rolling_peak, price)
        gap_pct      = ((price / rolling_peak) - 1) * 100 if rolling_peak > 0 else 0.0

        hedge_target, active_stage, state = _run_engine_day(
            vix=vix_val, rv20d=rv20d, fpi_net=fpi_weekly,
            gap_pct=gap_pct, current_price=price,
            state=state, ret_5d=0.0, new_calendar_day=True,
        )

        date_ts      = row['Date']
        last_date_str= date_ts.strftime('%Y-%m-%d')
        targets[date_ts] = (round(hedge_target, 1), active_stage)
        rows_computed += 1

    print(f"[Engine] Computed {rows_computed} rows")

    # ── Write results back to DataFrame ──────────────────────────────────────
    written = 0
    for date_ts, (ht, stage) in targets.items():
        mask = df['Date'] == date_ts
        # Only overwrite NaN rows in incremental mode; always overwrite in full/from mode
        if args.full or force_from is not None:
            df.loc[mask, 'Engine_Hedge_Target'] = ht
            df.loc[mask, 'Engine_Stage']         = stage
            written += 1
        else:
            # Incremental: only write if currently blank
            if df.loc[mask, 'Engine_Hedge_Target'].isna().any():
                df.loc[mask, 'Engine_Hedge_Target'] = ht
                df.loc[mask, 'Engine_Stage']         = stage
                written += 1

    print(f"[Write] {written} rows updated in DataFrame")

    # ── Save CSV ──────────────────────────────────────────────────────────────
    # Write dates back in original DD-MM-YYYY format to match the source file
    df_out = df.copy()
    df_out['Date'] = df_out['Date'].dt.strftime('%d-%m-%Y')
    df_out.to_csv(CSV_PATH, index=False, float_format='%.4f')
    print(f"[Save] CSV written: {CSV_PATH}")

    # ── Save checkpoint ───────────────────────────────────────────────────────
    if last_date_str:
        save_checkpoint(state, last_date_str)

    # ── Summary ───────────────────────────────────────────────────────────────
    filled = df['Engine_Hedge_Target'].notna().sum()
    total  = len(df)
    print(f"\n{'='*60}")
    print(f"  DONE")
    print(f"  Rows with Engine_Hedge_Target: {filled}/{total}")
    print(f"  Last computed date           : {last_date_str}")
    print(f"  Checkpoint saved             : {CHECKPOINT}")
    print(f"{'='*60}")
    print(f"\n  Next steps:")
    print(f"  1. Verify a few values look reasonable")
    print(f"  2. git add MARKET_DATA.csv engine_state_checkpoint.json")
    print(f"  3. git commit -m 'pre-compute engine history'")
    print(f"  4. git push  →  Render picks up new CSV on next deploy")
    print(f"\n  For future CSV updates:")
    print(f"  → Add new rows to MARKET_DATA.csv")
    print(f"  → Run: python bootstrap_engine_history.py   (incremental)")
    print(f"  → Push to repo")
    print(f"{'='*60}\n")


if __name__ == '__main__':
    main()
