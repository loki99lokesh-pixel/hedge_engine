from datetime import datetime, timezone
import os
import time
import threading
import io

import json
import redis

import pandas as pd
import requests
import yfinance as yf
from flask import Flask, jsonify, send_from_directory, request, send_file, Response

import portfolio_hedge_engine as phe

app = Flask(__name__, static_folder='.', static_url_path='')

try:
    from flask_cors import CORS
    CORS(app)
except ImportError:
    pass

V3_STATE_KEY = "v3_engine_state"

_V3_STATE_DEFAULTS = {
    'peak_price'  : 0,
    'trough_price': float('inf'),
    'days_in_dd'  : 0,
    'low_vol_days': 0,
    'prev_hedge'  : 0,
    'last_date'   : '',
    'ema_vix'     : 18.5, 
    'sod_dd_pct'  : 0.0,
    'sod_natural_target': 0.0,
    'sod_buffer'  : 5.0,
    'sod_recovery_ratio': 0.0
}

def load_v3_state():
    """Load v3 engine state from Redis (or LOCAL_CACHE in dev).
    Falls back to safe defaults if the key is missing or corrupt.
    NOTE: trough_price is stored as a large finite sentinel (1e15) because
    JSON cannot round-trip float('inf'). It is converted back on load."""
    try:
        data = get_state(V3_STATE_KEY)
        if data and isinstance(data, dict):
            # Restore sentinel back to inf so engine comparisons work correctly
            if data.get('trough_price', 0) >= 1e14:
                data['trough_price'] = float('inf')
            return data
    except Exception as e:
        print(f"[load_v3_state] Error reading state: {e}")
    return dict(_V3_STATE_DEFAULTS)

def save_v3_state(state):
    """Persist v3 engine state to Redis (or LOCAL_CACHE in dev).
    float('inf') is replaced with a large finite sentinel (1e15) because
    JSON does not support Infinity natively."""
    try:
        serialisable = dict(state)
        if serialisable.get('trough_price') == float('inf'):
            serialisable['trough_price'] = 1e15   # sentinel: restored on load
        save_state(V3_STATE_KEY, serialisable)
    except Exception as e:
        print(f"[save_v3_state] Error saving state: {e}")

@app.errorhandler(405)
def method_not_allowed(e):
    return Response('{"error": "Method not allowed"}', status=405, mimetype='application/json')

# ══════════════════════════════════════════════════════════════
# DYNAMIC BASELINES & LKG MEMORY
# ══════════════════════════════════════════════════════════════
LKG_DMA = 23521  # Seed — overwritten on first successful fetch

def get_dynamic_jan1_baseline():
    """Fetches the first trading day close for the current year."""
    current_year = datetime.now().year
    try:
        df = yf.download('^NSEI', start=f'{current_year}-01-01', end=f'{current_year}-01-08',
                         progress=False, threads=False)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
        if df.empty:
            raise ValueError("Empty data")
        return float(df[close_col].iloc[0])
    except Exception:
        return 26329  # Hard fallback

NIFTY_JAN1 = get_dynamic_jan1_baseline()

# ══════════════════════════════════════════════════════════════
# DATABASE CONNECTION (REDIS)
# ══════════════════════════════════════════════════════════════
REDIS_URL = os.environ.get('REDIS_URL')

if REDIS_URL:
    db = redis.from_url(REDIS_URL, decode_responses=True)
else:
    # Failsafe for local testing — simple in-memory dict mirrors Redis behaviour.
    # Seeded with safe fallback values so local runs are meaningful without Redis.
    db = None
    LOCAL_CACHE = {
        'dma'              : 23521.0,   # approximate 100-DMA seed
        'rolling_high_180d': 26277.0,   # recent 6M peak — gives max caution locally
        # Seed live_market_data so local runs without Redis get the same
        # field structure as Render. The background thread overwrites this
        # within seconds of startup; these values are only used if the
        # first /api/v3_hedge call races ahead of the first fetch.
        'live_market_data' : {
            'nifty'           : 24000.0,
            'vix'             : 18.5,
            'dd_from_peak_pct': -8.5,   # representative current drawdown
            'dmaGap'          : -3.5,
            'ret_5d'          : -1.2,
        },
    }

def save_state(key, data):
    if db:
        db.set(key, json.dumps(data))
    else:
        LOCAL_CACHE[key] = data   # uniform — no special cases

def get_state(key):
    if db:
        val = db.get(key)
        return json.loads(val) if val else None
    else:
        return LOCAL_CACHE.get(key)  # returns None if key missing — callers handle this


# ══════════════════════════════════════════════════════════════
# DATA FETCHERS
# ══════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════
# ROBUST DATA FETCHERS (STOOQ PRIMARY + YFINANCE FALLBACK)
# ══════════════════════════════════════════════════════════════
def attempt_yf(symbol, name):
    # Disguise removed. We let yfinance use its native curl_cffi bypass.
    df = yf.download(symbol, period='1d', interval='1d', progress=False, threads=False)
    if df.empty:
        raise ValueError(f"yfinance returned empty for {name}")
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
    return float(df[close_col].iloc[-1])

def attempt_stooq(symbol, name):
    url = f"https://stooq.com/q/l/?s={symbol}&f=sd2t2ohlcv&h&e=csv"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers, timeout=5)
    response.raise_for_status()
    lines = response.text.strip().split('\n')
    if len(lines) < 2:
        raise ValueError(f"Stooq returned no data rows for {name}")
    cols = lines[1].split(',')
    raw_val = cols[5] if len(cols) >= 6 else cols[-1]
    
    # Gracefully handle the N/D (No Data) string during weekends/holidays
    if raw_val.strip() == 'N/D':
        raise ValueError(f"Stooq returned N/D for {name}, shifting to fallback.")
        
    val = float(raw_val)
    if val <= 0:
        raise ValueError(f"Stooq returned invalid value {val} for {name}")
    return val

def get_prev_close(symbol: str) -> float:
    """Returns the previous trading day's close — used as the day-open baseline for change arrows."""
    df = yf.download(symbol, period='5d', interval='1d', progress=False, threads=False)
    if df.empty:
        raise ValueError(f'No history for {symbol}')
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
    series = df[close_col].dropna()
    if len(series) < 2:
        raise ValueError(f'Not enough rows for prev close: {symbol}')
    return float(series.iloc[-2])

def get_dma_only() -> tuple:
    """Fetches 1 year of history for the 100-DMA and 180-day rolling high calculations."""
    df = yf.download('^NSEI', period='1y', interval='1d', progress=False, threads=False)
    if df.empty:
        raise ValueError('No yfinance history for ^NSEI')
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
    series = df[close_col].dropna()

    # 100-DMA — trend direction signal (used in checklist)
    dma_100 = float(series.rolling(window=100).mean().iloc[-1])
    if pd.isna(dma_100):
        raise ValueError('Calculated DMA is NaN')

    # 180-day rolling high — how deep are we in the current drawdown
    # 180 calendar days ≈ 126 trading days; we use last 126 rows of daily data
    rolling_high_180d = float(series.iloc[-126:].max())

    save_state('dma', dma_100)
    save_state('rolling_high_180d', rolling_high_180d)
    return dma_100, rolling_high_180d

def fetch_live_data():
    data = {}
    sources = {}

    # 1. NIFTY LIVE PRICE (Stooq Primary -> yfinance Fallback)
    try:
        data['nifty'] = attempt_stooq('^nsei', 'nifty')
        sources['nifty'] = 'stooq'
    except Exception as e:
        print(f"Stooq Nifty failed: {e}")
        data['nifty'] = attempt_yf('^NSEI', 'nifty')
        sources['nifty'] = 'yfinance_fallback'

    # 2. NIFTY 100-DMA + 180-day rolling high (yfinance background -> LKG Memory Fallback)
    try:
        dynamic_dma, rolling_high_180d = get_dma_only()
    except Exception as e:
        print(f"DMA math failed, using LKG: {e}")
        dynamic_dma      = get_state('dma') or 23521.0
        rolling_high_180d = get_state('rolling_high_180d') or 26277.0

    # 3. INDIA VIX (Stooq Primary -> yfinance Fallback)
    try:
        data['vix'] = attempt_stooq('^indiavix', 'vix')
        sources['vix'] = 'stooq'
    except Exception:
        data['vix'] = attempt_yf('^INDIAVIX', 'vix')
        sources['vix'] = 'yfinance_fallback'

    # 4. USD/INR (Frankfurter Primary -> yfinance Fallback)
    try:
        res = requests.get('https://api.frankfurter.app/latest?from=USD&to=INR', timeout=5).json()
        data['usdinr'] = res['rates']['INR']
        sources['usdinr'] = 'frankfurter'
    except Exception:
        try:
            data['usdinr'] = attempt_yf('INR=X', 'usdinr')
            sources['usdinr'] = 'yfinance_fallback'
        except Exception:
            data['usdinr'] = 84.50

    # 5. GOLD (Stooq Primary -> yfinance Fallback)
    try:
        data['goldUSD'] = attempt_stooq('xauusd', 'goldUSD')
        sources['goldUSD'] = 'stooq'
    except Exception:
        data['goldUSD'] = attempt_yf('GC=F', 'goldUSD')
        sources['goldUSD'] = 'yfinance_fallback'

    # 6. BRENT CRUDE (Stooq Primary -> yfinance Fallback)
    try:
        data['brent'] = attempt_stooq('cb.f', 'brent')
        sources['brent'] = 'stooq'
    except Exception:
        data['brent'] = attempt_yf('BZ=F', 'brent')
        sources['brent'] = 'yfinance_fallback'

    # Previous-day closes — used by dashboard for day-change arrows
    # Each falls back to None silently; frontend handles missing values gracefully
    try:
        data['prevClose_nifty']  = get_prev_close('^NSEI')
    except Exception:
        data['prevClose_nifty']  = None

    try:
        data['prevClose_vix']    = get_prev_close('^INDIAVIX')
    except Exception:
        data['prevClose_vix']    = None

    try:
        data['prevClose_goldUSD'] = get_prev_close('GC=F')
    except Exception:
        data['prevClose_goldUSD'] = None

    try:
        data['prevClose_brent']  = get_prev_close('BZ=F')
    except Exception:
        data['prevClose_brent']  = None

    # USD/INR and Gold INR prev close derived from the above
    try:
        data['prevClose_usdinr'] = get_prev_close('USDINR=X')
    except Exception:
        data['prevClose_usdinr'] = None

    if data.get('prevClose_goldUSD') and data.get('prevClose_usdinr'):
        data['prevClose_goldINR'] = round(
            data['prevClose_goldUSD'] * data['prevClose_usdinr'] / 31.1035, 2
        )
    else:
        data['prevClose_goldINR'] = None

    # 5-day return for Mixed Signal Filter in v3.2 engine
    try:
        df5 = yf.download('^NSEI', period='10d', interval='1d', progress=False, threads=False)
        if isinstance(df5.columns, pd.MultiIndex):
            df5.columns = df5.columns.get_level_values(0)
        close_col5 = 'Close' if 'Close' in df5.columns else df5.columns[-1]
        s5 = df5[close_col5].dropna()
        if len(s5) >= 6:
            data['ret_5d'] = round(((float(s5.iloc[-1]) / float(s5.iloc[-6])) - 1) * 100, 2)
        else:
            data['ret_5d'] = 0.0
    except Exception:
        data['ret_5d'] = 0.0

    # Final Calculations
    data['goldINR'] = round(data['goldUSD'] * data['usdinr'] / 31.1035, 2)
    data['ytd'] = round(((data['nifty'] / NIFTY_JAN1) - 1) * 100, 2)
    data['dmaGap'] = round(((data['nifty'] / dynamic_dma) - 1) * 100, 2)
    data['dma_used'] = round(dynamic_dma, 2)
    data['rolling_high_180d'] = round(rolling_high_180d, 2)
    data['dd_from_peak_pct'] = round(((data['nifty'] / rolling_high_180d) - 1) * 100, 2)
    
    from datetime import datetime, timezone
    data['timestamp'] = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
    data['sources'] = sources

    return data

def update_cache_loop():
    while True:
        try:
            payload = fetch_live_data()
            save_state('live_market_data', payload)
        except Exception as e:
            print(f"[Background Sync Error] {e}")
        time.sleep(300) # Re-fetch every 5 minutes

def _parse_date_column(series: pd.Series) -> pd.Series:
    """
    Robustly parses a date Series by trying common formats explicitly.
    FIX: parse_dates=['Date'] silently fails on DD-MM-YYYY (common in Indian
    exports), leaving string dtype and causing AttributeError on .date calls.
    """
    for fmt in ['%d-%m-%Y', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
        try:
            parsed = pd.to_datetime(series, format=fmt)
            if parsed.notna().sum() > len(series) * 0.9:
                return parsed
        except Exception:
            continue
    return pd.to_datetime(series, infer_datetime_format=True, errors='coerce')


def _load_csv_history() -> pd.DataFrame:
    """
    Loads MARKET_DATA.csv from the project directory.

    FAST PATH (preferred): If the CSV has Engine_Hedge_Target and Engine_Stage
    columns (pre-computed by bootstrap_engine_history.py), we read them directly
    and build chart_history with zero engine replay. This is instant.

    SLOW PATH (fallback): If pre-computed columns are missing, falls back to the
    original live engine replay over the last 1 year of CSV data.

    Returns a DataFrame indexed by Date with at minimum:
      Nifty_Close, VIX, Nifty_RealVol20d, FPI_Net_Equity_Cr,
      Engine_Hedge_Target (float | NaN), Engine_Stage (str | '')
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(script_dir, 'MARKET_DATA.csv'),
        'MARKET_DATA.csv',
        os.path.join(script_dir, 'MASTER_CLEAN_UPDATED.csv'),
        'MASTER_CLEAN_UPDATED.csv',
    ]
    needed = ['Nifty_Close', 'VIX', 'Nifty_RealVol20d', 'FPI_Net_Equity_Cr']

    for path in candidates:
        if not os.path.exists(path):
            continue
        try:
            # Read WITHOUT parse_dates — parse Date column explicitly below
            df = pd.read_csv(path, low_memory=False)

            if 'Date' not in df.columns:
                raise ValueError(f"'Date' column missing. Found: {df.columns.tolist()}")

            df['Date'] = _parse_date_column(df['Date'])
            invalid = df['Date'].isna().sum()
            if invalid > len(df) * 0.1:
                raise ValueError(f"Date parsing produced {invalid} NaT values")

            df = df.dropna(subset=['Date']).sort_values('Date').set_index('Date')

            for col in needed:
                if col not in df.columns:
                    df[col] = float('nan')

            # Ensure engine columns exist (may be absent in older CSV versions)
            if 'Engine_Hedge_Target' not in df.columns:
                df['Engine_Hedge_Target'] = float('nan')
            if 'Engine_Stage' not in df.columns:
                df['Engine_Stage'] = ''

            df[needed] = df[needed].ffill().bfill()

            fname = os.path.basename(path)
            pre_computed = df['Engine_Hedge_Target'].notna().sum()
            print(f"[Bootstrap] {fname} loaded: {len(df)} rows "
                  f"{df.index[0].date()} → {df.index[-1].date()} "
                  f"| Pre-computed rows: {pre_computed}")
            return df
        except Exception as e:
            print(f"[Bootstrap] {path} failed: {e}")

    print("[Bootstrap] No CSV found — will use yfinance fallback")
    return pd.DataFrame()


def bootstrap_v3_state():
    """
    Initialises v3 engine state and chart history in Redis on first boot.

    FAST PATH (CSV has Engine_Hedge_Target column):
      Reads pre-computed hedge targets directly from CSV — no engine replay.
      Loads up to 3,650 days (10 years) of history from 2015 onwards.
      Completes in under 1 second.

    SLOW PATH (CSV missing pre-computed columns):
      Falls back to live engine replay over last 1 year of CSV data.
      Run bootstrap_engine_history.py to pre-compute and avoid this path.

    YFINANCE FALLBACK (no CSV at all):
      Downloads 1 year of Nifty data and replays with neutral VIX/FPI seeds.

    On subsequent boots with healthy Redis state: skips everything immediately.
    """
    state = load_v3_state()
    if state.get('peak_price', 0) != 0:
        print("[Bootstrap] State already initialised — skipping.")
        return

    print("[Bootstrap] Fresh state — initialising from CSV...")
    try:
        csv_df = _load_csv_history()

        # ── Determine start date: 2015-01-01 or earliest available ──────────
        start_date = pd.Timestamp('2015-01-01')

        if not csv_df.empty:
            has_precomputed = (
                'Engine_Hedge_Target' in csv_df.columns
                and csv_df['Engine_Hedge_Target'].notna().sum() > 100
            )

            if has_precomputed:
                # ════════════════════════════════════════════════════════════
                # FAST PATH — read pre-computed columns directly
                # ════════════════════════════════════════════════════════════
                print("[Bootstrap] Fast path — reading pre-computed engine columns...")

                # Slice from 2015 onwards (where Engine_Hedge_Target is filled)
                recent = csv_df[csv_df.index >= start_date].copy()
                recent = recent[recent['Engine_Hedge_Target'].notna()]

                if len(recent) < 10:
                    raise ValueError("Pre-computed column exists but has < 10 valid rows")

                chart_history = []
                rolling_peak  = 0.0

                for date_idx, row in recent.iterrows():
                    price_f = float(row['Nifty_Close'])
                    if pd.isna(price_f) or price_f <= 0:
                        continue

                    rolling_peak = max(rolling_peak, price_f)
                    gap_pct = ((price_f / rolling_peak) - 1) * 100 if rolling_peak > 0 else 0.0

                    date_str = (date_idx.strftime('%Y-%m-%d')
                                if hasattr(date_idx, 'strftime') else str(date_idx)[:10])
                    chart_history.append({
                        'date'        : date_str,
                        'hedge_target': float(row['Engine_Hedge_Target']),
                        'drawdown'    : round(gap_pct, 2),
                        'stage'       : str(row.get('Engine_Stage', 'Stage 1 — Onset')),
                        'nifty'       : round(price_f, 0),
                    })

                # Build engine state from the last row
                last_row   = recent.iloc[-1]
                last_price = float(last_row['Nifty_Close'])
                bootstrap_state = dict(_V3_STATE_DEFAULTS)
                bootstrap_state['trough_price'] = float('inf')
                bootstrap_state['peak_price']   = rolling_peak
                bootstrap_state['last_date']    = (
                    recent.index[-1].strftime('%Y-%m-%d')
                    if hasattr(recent.index[-1], 'strftime')
                    else str(recent.index[-1])[:10]
                )

                print(f"[Bootstrap] Fast path done. "
                      f"Rows loaded: {len(chart_history)} | "
                      f"HWM: {rolling_peak:,.0f} | "
                      f"Last date: {bootstrap_state['last_date']}")

            else:
                # ════════════════════════════════════════════════════════════
                # SLOW PATH — live engine replay (pre-computed columns missing)
                # ════════════════════════════════════════════════════════════
                print("[Bootstrap] Pre-computed columns not found — running engine replay...")
                print("[Bootstrap] Tip: run bootstrap_engine_history.py to pre-compute "
                      "and make future boots instant.")

                from datetime import timedelta
                cutoff = datetime.now(timezone.utc).date() - timedelta(days=365)
                recent = csv_df[csv_df.index.date >= cutoff]
                if len(recent) < 50:
                    recent = csv_df.tail(252)

                bootstrap_state = dict(_V3_STATE_DEFAULTS)
                bootstrap_state['trough_price'] = float('inf')
                replay_rolling_peak = 0.0
                chart_history = []

                for date_idx, row in recent.iterrows():
                    price_f = float(row['Nifty_Close'])
                    if pd.isna(price_f) or price_f <= 0:
                        continue
                    vix_val   = float(row['VIX']) if not pd.isna(row['VIX']) else 18.0
                    rv20d_val = (float(row['Nifty_RealVol20d'])
                                 if not pd.isna(row['Nifty_RealVol20d']) else vix_val * 0.75)
                    fpi_daily = (float(row['FPI_Net_Equity_Cr'])
                                 if not pd.isna(row['FPI_Net_Equity_Cr']) else -1200.0)
                    fpi_weekly = fpi_daily * 5

                    replay_rolling_peak = max(replay_rolling_peak, price_f)
                    replay_gap = ((price_f / replay_rolling_peak) - 1) * 100 if replay_rolling_peak > 0 else 0.0

                    hedge_target, active_stage, _, bootstrap_state = phe.calculate_v3_magnitude_hedge(
                        vix=vix_val, rv20d=rv20d_val, fpi_net=fpi_weekly,
                        gap_pct=replay_gap, current_price=price_f,
                        state=bootstrap_state, ret_5d=0.0, new_calendar_day=True
                    )
                    date_str = (date_idx.strftime('%Y-%m-%d')
                                if hasattr(date_idx, 'strftime') else str(date_idx)[:10])
                    chart_history.append({
                        'date'        : date_str,
                        'hedge_target': round(float(hedge_target), 1),
                        'drawdown'    : round(float(replay_gap), 2),
                        'stage'       : active_stage,
                        'nifty'       : round(price_f, 0),
                    })

                print(f"[Bootstrap] Replay done. "
                      f"Peak: {bootstrap_state['peak_price']:,.0f} | "
                      f"Chart points: {len(chart_history)}")

        else:
            # ════════════════════════════════════════════════════════════════
            # YFINANCE FALLBACK — no CSV available
            # ════════════════════════════════════════════════════════════════
            print("[Bootstrap] No CSV — falling back to yfinance with neutral seeds...")
            yf_df = yf.download('^NSEI', period='1y', interval='1d',
                                 progress=False, threads=False)
            if isinstance(yf_df.columns, pd.MultiIndex):
                yf_df.columns = yf_df.columns.get_level_values(0)
            close_col = 'Close' if 'Close' in yf_df.columns else yf_df.columns[-1]
            recent = pd.DataFrame({'Nifty_Close': yf_df[close_col].dropna()})
            recent['VIX']               = 18.0
            recent['Nifty_RealVol20d']  = 14.0
            recent['FPI_Net_Equity_Cr'] = -1200.0

            bootstrap_state = dict(_V3_STATE_DEFAULTS)
            bootstrap_state['trough_price'] = float('inf')
            replay_rolling_peak = 0.0
            chart_history = []

            for date_idx, row in recent.iterrows():
                price_f = float(row['Nifty_Close'])
                if pd.isna(price_f) or price_f <= 0:
                    continue
                replay_rolling_peak = max(replay_rolling_peak, price_f)
                replay_gap = ((price_f / replay_rolling_peak) - 1) * 100 if replay_rolling_peak > 0 else 0.0
                hedge_target, active_stage, _, bootstrap_state = phe.calculate_v3_magnitude_hedge(
                    vix=18.0, rv20d=14.0, fpi_net=-6000.0,
                    gap_pct=replay_gap, current_price=price_f,
                    state=bootstrap_state, ret_5d=0.0, new_calendar_day=True
                )
                date_str = (date_idx.strftime('%Y-%m-%d')
                            if hasattr(date_idx, 'strftime') else str(date_idx)[:10])
                chart_history.append({
                    'date'        : date_str,
                    'hedge_target': round(float(hedge_target), 1),
                    'drawdown'    : round(float(replay_gap), 2),
                    'stage'       : active_stage,
                    'nifty'       : round(price_f, 0),
                })
            print(f"[Bootstrap] yfinance fallback done. Points: {len(chart_history)}")
            bootstrap_state['peak_price'] = replay_rolling_peak

        # ── Persist to Redis ──────────────────────────────────────────────────
        bootstrap_state['last_date'] = datetime.now(timezone.utc).strftime('%Y-%m-%d')
        save_v3_state(bootstrap_state)
        # Store up to 3,650 days (10 years) — covers 2015 → present
        save_state('v3_chart_history', chart_history[-3650:])
        print(f"[Bootstrap] Redis seeded. Chart points stored: {len(chart_history[-3650:])}")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[Bootstrap] Failed — engine will self-correct over coming days: {e}")

# Run bootstrap synchronously FIRST — blocks until complete so that
# the first /api/v3_hedge call always sees a fully seeded state.
# Only takes time on first boot or after a reset (peak_price == 0).
# On normal restarts with healthy Redis state it returns in milliseconds.
bootstrap_v3_state()

# Background loop starts AFTER bootstrap is done
threading.Thread(target=update_cache_loop, daemon=True).start()

# ══════════════════════════════════════════════════════════════
# HELPER: Build portfolio DataFrame from frontend payload
# ══════════════════════════════════════════════════════════════
def build_portfolio_df(holdings):
    """
    Converts the JSON holdings list from the frontend into a DataFrame
    matching what portfolio_hedge_engine expects.

    Frontend sends: [{name, type, alloc, beta, cat}, ...]
      name  = display name
      type  = asset_type key (e.g. 'large_cap_stock')
      alloc = allocation %
      beta  = resolved beta value (already computed by frontend JS)
      cat   = asset category ('equity', 'debt', 'gold', 'intl', 'cash')

    Engine expects columns:
      Holding_Name, Asset_Type, Allocation_Pct, Beta, Sector,
      Weight, Weighted_Beta
    """
    df = pd.DataFrame(holdings)

    # Map frontend asset-category ('equity', 'debt', …) to readable sector labels
    CAT_LABEL = {
        'equity': 'Equity',
        'debt'  : 'Debt / Liquid',
        'gold'  : 'Gold / Hedge',
        'intl'  : 'International',
        'cash'  : 'Cash / FD',
    }

    df.rename(columns={
        'name' : 'Holding_Name',
        'type' : 'Asset_Type',
        'alloc': 'Allocation_Pct',
        'beta' : 'Beta',
    }, inplace=True)

    # Use readable sector label; fall back to the raw cat value if unknown
    df['Sector'] = df['cat'].apply(lambda c: CAT_LABEL.get(str(c).lower(), str(c).title()))
    df.drop(columns=['cat'], inplace=True)

    df['Allocation_Pct'] = pd.to_numeric(df['Allocation_Pct'], errors='coerce').fillna(0)
    df['Beta'] = pd.to_numeric(df['Beta'], errors='coerce').fillna(1.0)

    total = df['Allocation_Pct'].sum()
    if total > 0 and abs(total - 100) > 0.5:
        df['Allocation_Pct'] = df['Allocation_Pct'] / total * 100

    df['Weight'] = df['Allocation_Pct'] / 100
    df['Weighted_Beta'] = df['Weight'] * df['Beta']

    return df

# ══════════════════════════════════════════════════════════════
# WEB ROUTES
# ══════════════════════════════════════════════════════════════
@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'dashboard.html')

@app.route('/portfolio')
def portfolio_route():
    return send_from_directory(app.static_folder, 'portfolio_hedge_dashboard.html')

@app.route('/api/live')
def api_live():
    cached = get_state('live_market_data')
    if cached:
        return jsonify({'status': 'Success', 'data': cached, 'cached': True})
    
    # Failsafe if the database is completely empty on first boot
    try:
        payload = fetch_live_data()
        save_state('live_market_data', payload)
        return jsonify({'status': 'Success', 'data': payload, 'cached': False})
    except Exception as exc:
        return jsonify({'status': 'Error', 'message': str(exc)}), 500

# ══════════════════════════════════════════════════════════════
# ENGINE BRIDGES
# ══════════════════════════════════════════════════════════════
@app.route('/api/generate-excel', methods=['POST','OPTIONS'])
@app.route('/api/generate-excel/', methods=['POST','OPTIONS'])
def generate_excel():
    """Generates and streams the Excel workbook."""
    if request.method == 'OPTIONS':
        return '', 204
    try:
        payload = request.json
        holdings = payload.get('holdings', [])
        port_value = payload.get('portfolio_value')

        if not holdings:
            return jsonify({'error': 'No holdings provided'}), 400

        df = build_portfolio_df(holdings)
        metrics = phe.compute_portfolio_metrics(df, port_value)
        hedge_sizing = phe.compute_hedge_sizing(metrics["portfolio_beta"], port_value)
        scenario_df = phe.compute_scenario_analysis(metrics["portfolio_beta"], port_value)

        output = io.BytesIO()
        phe.build_excel(df, metrics, hedge_sizing, scenario_df, output)
        output.seek(0)

        return send_file(
            output,
            download_name="Institutional_Hedge_Report.xlsx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        print(f"Engine Error (Excel): {e}")
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate-text', methods=['POST','OPTIONS'])
@app.route('/api/generate-text/', methods=['POST','OPTIONS'])
def generate_text():
    """
    Generates and streams the text report.

    FIX: Now writes directly to io.StringIO — no temp file needed.
    The engine's generate_text_report() already supports file-like objects
    via hasattr(output_path, 'write') check.
    """
    if request.method == 'OPTIONS':
        return '', 204
    try:
        payload = request.json
        holdings = payload.get('holdings', [])
        port_value = payload.get('portfolio_value')

        if not holdings:
            return jsonify({'error': 'No holdings provided'}), 400

        df = build_portfolio_df(holdings)
        metrics = phe.compute_portfolio_metrics(df, port_value)
        hedge_sizing = phe.compute_hedge_sizing(metrics["portfolio_beta"], port_value)
        scenario_df = phe.compute_scenario_analysis(metrics["portfolio_beta"], port_value)

        # Use the already-cached live data (avoids a second 6-month yfinance download
        # in the same request and prevents the endpoint from hanging on slow connections)
        cached = get_state('live_market_data')

        live_nifty = None
        if cached:
            try:
                live_nifty = {
                    "close"          : cached["nifty"],
                    "dma100"         : cached["dma_used"],
                    "gap_pct"        : cached["dmaGap"],
                    "trigger_active" : cached["dmaGap"] <= -2.0,
                }
            except Exception:
                live_nifty = None

        # Write to StringIO buffer — no disk I/O needed
        text_buffer = io.StringIO()
        phe.generate_text_report(metrics, hedge_sizing, scenario_df, text_buffer, live_nifty)

        # Encode to UTF-8 bytes and stream
        memory_file = io.BytesIO(text_buffer.getvalue().encode('utf-8'))
        memory_file.seek(0)

        return send_file(
            memory_file,
            download_name="Institutional_Hedge_Report.txt",
            as_attachment=True,
            mimetype='text/plain; charset=utf-8'
        )
    except Exception as e:
        print(f"Engine Error (Text): {e}")
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-word', methods=['POST','OPTIONS'])
@app.route('/api/generate-word/', methods=['POST','OPTIONS'])
def generate_word():
    """Generates and streams the Word document report."""
    if request.method == 'OPTIONS':
        return '', 204
    try:
        payload = request.json
        holdings = payload.get('holdings', [])
        port_value = payload.get('portfolio_value')

        if not holdings:
            return jsonify({'error': 'No holdings provided'}), 400

        df = build_portfolio_df(holdings)
        metrics = phe.compute_portfolio_metrics(df, port_value)
        hedge_sizing = phe.compute_hedge_sizing(metrics["portfolio_beta"], port_value)
        scenario_df = phe.compute_scenario_analysis(metrics["portfolio_beta"], port_value)

        cached = get_state('live_market_data')

        live_nifty = None
        if cached:
            try:
                live_nifty = {
                    "close"          : cached["nifty"],
                    "dma100"         : cached["dma_used"],
                    "gap_pct"        : cached["dmaGap"],
                    "trigger_active" : cached["dmaGap"] <= -2.0,
                }
            except Exception:
                live_nifty = None

        output = io.BytesIO()
        phe.generate_docx_report(metrics, hedge_sizing, scenario_df, output, live_nifty)
        output.seek(0)

        return send_file(
            output,
            download_name="Institutional_Hedge_Report.docx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        print(f"Engine Error (Word): {e}")
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/v3_hedge', methods=['GET'])
def get_v3_magnitude_hedge():
    try:
        # 1. Pull all live inputs from the cached market data (populated by background thread)
        cached = get_state('live_market_data') or {}
        nifty_close      = cached.get('nifty', 24000)
        # Use dd_from_peak_pct (6M rolling high) as the drawdown signal — this correctly
        # captures drawdowns that haven't yet crossed the 100-DMA (e.g. current ~7% fall).
        # dmaGap is kept for the checklist signal only (trend direction).
        dd_from_peak_pct = cached.get('dd_from_peak_pct', 0.0)  # negative when below peak
        gap_pct          = cached.get('dmaGap', 0)               # kept for inputs_used log
        vix              = cached.get('vix', 18.5)

        # BUG FIX: ret_5d MUST be read from cache before rv20d is computed.
        # The original code placed this assignment AFTER the rv20d try-block,
        # so ret_5d was always an undefined name inside the try, causing a
        # NameError that silently fell through to 'except Exception' on every
        # call. Result: rv20d was always vix*0.65 (the floor), never the actual
        # realised-vol from price movement. On Render this inflated rv20d vs
        # local because Render had a bootstrapped state with real VIX data.
        ret_5d = cached.get('ret_5d', 0.0)

        # rv20d: proper annualised realised-vol proxy from the 5-day return.
        # Formula: daily_move = |ret_5d| / sqrt(5), annualised = daily_move * sqrt(252).
        # Floored at vix * 0.65 (RV is typically ~65-80% of IV) so it never
        # collapses to zero on a quiet 5-day window.
        try:
            daily_move  = abs(ret_5d) / (5 ** 0.5)
            rv20d_raw   = daily_move * (252 ** 0.5)
            rv20d_floor = vix * 0.65
            rv20d       = round(max(rv20d_floor, rv20d_raw), 2)
        except Exception:
            rv20d = vix * 0.65

        # 2. FPI weekly outflow — read from query param sent by the dashboard
        # Dashboard sends live.fpiWeekly (user-entered via the ✎ edit field)
        try:
            fpi = float(request.args.get('fpi', -1200))
        except (TypeError, ValueError):
            fpi = -1200

        # 3. portfolio_beta — defaults to 1.0 for the main dashboard (pure Nifty exposure).
        # The portfolio hedge page passes the user's actual computed beta via ?beta=X
        # so the engine returns a beta-adjusted target specific to that portfolio.
        try:
            portfolio_beta = float(request.args.get('beta', 1.0))
            portfolio_beta = max(0.01, min(3.0, portfolio_beta))   # clamp to sane range
        except (TypeError, ValueError):
            portfolio_beta = 1.0

        # 4. Load persisted state memory (peak/trough/days_in_dd/low_vol_days)
        state = load_v3_state()

        # If peak_price is 0 (fresh state after reset or first boot), seed it
        # from the live 6M rolling high so the engine's internal dd_pct aligns
        # immediately with dd_from_peak_pct. Without this, the engine starts
        # from today's price as the new peak, producing dd_pct=0 and showing
        # 0% in the "DD from 6M Peak" tile right after a state reset.
        if state.get('peak_price', 0) == 0:
            rh = cached.get('rolling_high_180d', nifty_close)
            state['peak_price'] = float(rh)

        # Fix: low_vol_days must count calendar days, not browser refreshes.
        # We record the UTC date of the last engine call and only advance
        # low_vol_days when a genuinely new calendar date is observed.
        today_str = datetime.now(timezone.utc).strftime('%Y-%m-%d')
        last_date = state.get('last_date', '')
        new_calendar_day = (today_str != last_date)

        # 5. Execute the v3.2 magnitude engine
        # gap_pct passed is now dd_from_peak_pct (6M rolling high based) so the
        # engine's gapN component correctly reflects the real current drawdown depth.
        nifty_hedge, active_stage, diag, new_state = phe.calculate_v3_magnitude_hedge(
            vix, rv20d, fpi, dd_from_peak_pct, nifty_close, state, ret_5d,
            new_calendar_day=new_calendar_day
        )

        # 6. Scale by portfolio beta (1.0 for dashboard = no scaling, pure Nifty target)
        adjusted_hedge = min(100.0, nifty_hedge * portfolio_beta)

        # 7. Persist updated state + append live chart point
        new_state['last_date'] = today_str
        save_v3_state(new_state)

        # Append today's live point to chart history (once per calendar day)
        if new_calendar_day:
            try:
                history = get_state('v3_chart_history') or []
                # Avoid duplicate dates
                if not history or history[-1].get('date') != today_str:
                    history.append({
                        'date'        : today_str,
                        'hedge_target': round(adjusted_hedge, 1),
                        'drawdown'    : round(dd_from_peak_pct, 2),
                        'stage'       : active_stage,
                        'nifty'       : round(nifty_close, 0),
                    })
                    save_state('v3_chart_history', history[-3650:])  # rolling 10 years
            except Exception as e:
                print(f"[Chart History] Append failed: {e}")

        return jsonify({
            'status': 'success',
            'data': {
                'beta_adjusted_target': round(adjusted_hedge, 1),
                'nifty_base_target':    round(nifty_hedge, 1),
                'active_stage':         active_stage,
                'diagnostics':          diag,
                'inputs_used': {
                    'vix':              round(vix, 1),
                    'rv20d':            round(rv20d, 1),
                    'fpi_weekly':       fpi,
                    'dd_from_peak_pct': round(dd_from_peak_pct, 2),  # used by engine
                    'dma_gap_pct':      round(gap_pct, 2),            # checklist only
                    'ret_5d':           round(ret_5d, 2),
                    'nifty_close':      round(nifty_close, 0),
                }
            }
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/v3_reset', methods=['POST'])
def reset_v3_state(): 
    # Clears the persisted v3.2 state file so the engine starts fresh.
    try: 
        fresh_state = {
            'peak_price'   : 0, 
            'trough_price' : float('inf'), 
            'days_in_dd'   : 0, 
            'low_vol_days' : 0, 
            'prev_hedge'   : 0,
            'last_date'    : '',
            'ema_vix'      : 18.5, 
            'sod_dd_pct'   : 0.0,
            'sod_natural_target': 0.0,
            'sod_buffer'   : 5.0,
            'sod_recovery_ratio': 0.0
        } 
        save_v3_state(fresh_state) 
        return jsonify({'status': 'success', 'message': 'v3 state reset to fresh defaults.'}) 
    except Exception as e: 
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/v3_chart', methods=['GET'])
def get_v3_chart():
    """Returns up to 3,650 days (10 years) of engine performance history.
    Each point: {date, hedge_target, drawdown, stage, nifty}
    Pre-computed from 2015 via bootstrap_engine_history.py; live-appended daily.
    """
    try:
        history = get_state('v3_chart_history') or []
        return jsonify({'status': 'success', 'data': history})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/v3_state', methods=['GET'])
def get_v3_state_debug():
    """Returns the current raw v3 state — useful for diagnosing Stage 3 persistence."""
    try:
        state = load_v3_state()
        return jsonify({'status': 'success', 'state': state})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'ok'}), 200

if __name__ == '__main__':
    print("=" * 60)
    print("  Starting Institutional Risk Engine Server")
    print("  Main Dashboard:  http://127.0.0.1:5000/")
    print("  Portfolio Tool:  http://127.0.0.1:5000/portfolio")
    print("=" * 60)
    app.run(debug=True, port=5000, use_reloader=False)
