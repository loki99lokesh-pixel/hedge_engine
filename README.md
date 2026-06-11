# Nifty 50 — Institutional Risk Engine

A full-stack quantitative risk monitoring and portfolio hedging system for Indian equity markets. Tracks live Nifty 50 conditions, runs a multi-stage continuous hedge sizing engine, and serves two frontends: a main institutional dashboard and a portfolio-specific hedging tool.

---

## Project structure

```
├── backend.py                      # Flask server — API routes, data fetching, engine orchestration
├── portfolio_hedge_engine.py       # v3.2 magnitude engine + portfolio analytics + report builders
├── bootstrap_engine_history.py     # Pre-computes engine output for all historical CSV rows
├── dashboard.html                  # Main institutional risk dashboard  (served at /)
├── portfolio_hedge_dashboard.html  # Portfolio-specific hedging tool    (served at /portfolio)
└── MARKET_DATA.csv                 # Historical daily data 1996–present — read-only source of truth
```

---

## Architecture

### Data flow

```
MARKET_DATA.csv ──► bootstrap_engine_history.py ──► Redis (v3_chart_history)
                                                            ▲
yfinance / Stooq / Frankfurter ──► fetch_live_data() ──► Redis (live_market_data)
                                                            │
                                                            ▼
                                          /api/v3_hedge ──► Engine ──► Dashboard
```

- `MARKET_DATA.csv` is a **read-only** permanent historical record. The server never writes to it.
- Redis manages all live state: engine state, live market data cache, and chart history.
- A background thread re-fetches all live market data every 5 minutes and writes to Redis.
- Local development without `REDIS_URL` automatically falls back to an in-memory `LOCAL_CACHE` dict that mirrors Redis behaviour exactly.

### Redis key map

| Key | Contents | Written by |
|---|---|---|
| `live_market_data` | Nifty, VIX, USD/INR, Gold, Brent, DMA, drawdown %, 5d return, prev closes | Background thread every 5 min |
| `v3_engine_state` | Peak price, trough, days in drawdown, EMA VIX, SOD snapshots, gate vars | `/api/v3_hedge` on each call |
| `v3_chart_history` | Rolling 10-year array of daily engine outputs | Bootstrap on boot + `/api/v3_hedge` once per day |
| `dma` | Last-known-good 100-DMA value | `get_dma_only()` |
| `rolling_high_180d` | Last-known-good 6-month rolling high | `get_dma_only()` |

### Bootstrap sequence on server start

1. `bootstrap_v3_state()` runs **synchronously** before the first request is accepted.
2. **Fast path** (normal): CSV has pre-computed `Engine_Hedge_Target` column → reads directly, seeds Redis in under 1 second. No engine replay.
3. **Slow path** (fallback): pre-computed column missing → replays live engine over last 1 year of CSV data.
4. **yfinance fallback**: no CSV at all → downloads 1 year of Nifty data and replays with neutral VIX/FPI seeds.
5. On subsequent boots with a healthy Redis state (`peak_price != 0`): skips everything in milliseconds.

---

## Engine — v3.2 magnitude hedge

`portfolio_hedge_engine.py` → `calculate_v3_magnitude_hedge()`

The engine is a **continuous, magnitude-based hedge sizing model**. It does not use simple threshold switches. Instead, three layered sigmoid-based stages activate smoothly as market stress increases, with hard overrides only for extreme confirmed drawdowns.

### Inputs

| Param | Source | Description |
|---|---|---|
| `vix` | Stooq/yfinance | India VIX — EMA-smoothed before use |
| `rv20d` | Derived | Annualised realised vol proxy from 5-day return (`\|ret_5d\| / √5 × √252`), floored at `vix × 0.65` |
| `fpi_net` | User-entered | FPI weekly net equity flows in ₹ Cr (editable via dashboard) |
| `gap_pct` | Derived | `(nifty / rolling_high_180d − 1) × 100` — drawdown from 6-month peak |
| `current_price` | Stooq/yfinance | Live Nifty 50 close |
| `state` | Redis | Persisted state from previous call |
| `ret_5d` | yfinance | 5-day Nifty return % |

### Stage 1 — Onset score

Computes a normalised 0–100 composite from four signals:

```
onset_score = 0.48 × rvN + 0.25 × gapN + fpi_weight × fpiN + 0.12 × vixN
```

Each signal is normalised to 0–100:

| Signal | Formula |
|---|---|
| `rvN`  | `(rv20d − 5) / 25 × 100`, clamped 0–100 |
| `gapN` | `(−gap_pct) / 20 × 100`, clamped 0–100 |
| `fpiN` | outflow: `min(100, −fpi / 8000 × 100)` · inflow: `max(−50, −fpi / 8000 × 100)` |
| `vixN` | `(vix − 10) / 30 × 100`, clamped 0–100 |

`fpi_weight` is normally `0.15` but reduced to `0.12` when the Mixed Signal Filter is active (see below).

The onset score feeds `get_continuous_notional()`, a **triple-sigmoid** function:

```python
def get_continuous_notional(score):
    n = 10.0
    n += 30.0 * sigmoid((score - 25.0) / 3.5)   # mild stress ramp
    n += 30.0 * sigmoid((score - 50.0) / 3.5)   # elevated stress ramp
    n += 30.0 * sigmoid((score - 70.0) / 3.5)   # crisis ramp
    return min(100.0, max(10.0, n))
```

Range: 10%–100%. Inflection points at scores 25, 50, 70. Hedge rises continuously — no cliff-edge jumps.

### Stage 2 — Active phase (Lasso-derived)

Engages once a drawdown is confirmed. Uses the backtest-validated Lasso regression formula:

```
s2_target = 12.5 + 0.52 × vix + 0.48 × dd_pct + 5.10 × (vix × dd_pct / 100) − penalty
```

`penalty = min(0.08 × days_in_drawdown, 15)` — gradually reduces the target for prolonged drawdowns to prevent over-hedging stale positions. Clamped 10%–90%.

### Transition

```
final_target = max(s1_target, s2_target)
```

Whichever stage demands more protection wins. The `active_stage` label reflects the dominant stage.

### Stage 3 — Hard escalation overrides

Fire unconditionally, but only after `days_in_dd ≥ 3` (3-day shock window guard prevents flash crashes from triggering them):

| Condition | Override |
|---|---|
| `dd_pct ≥ 20%` | `final_target = max(final_target, 75%)` |
| `dd_pct ≥ 15%` AND (`rv20d > 28` OR `vix > 28`) | `final_target = 100%` |
| `dd_pct ≥ 25%` (unconditional) | `final_target = 100%` |

### De-escalation gate

Prevents premature hedge removal after a crisis. When `prev_hedge > sod_natural_target + sod_buffer`, the gate holds the previous hedge level until **at least one** escape route clears AND `days_in_dd ≥ 3`:

- **Escape route 1 — price recovery**: SOD `recovery_ratio ≥ 0.40` (price has recovered 40% of peak-to-trough range)
- **Escape route 2 — volatility collapse**: `rv20d < 22` AND `vix < 20` for ≥ 5 consecutive calendar days

`sod_buffer` is dynamic: `5.0 + max(0, (vix − 15) / 10 × 3)` — wider buffer at higher VIX.

### VIX EMA smoothing

Raw VIX is smoothed with α=0.5 before any engine use:

```
ema_vix = vix × 0.5 + prev_ema_vix × 0.5
```

Prevents intraday VIX noise from triggering false escalations.

### Mixed Signal Filter (MSF)

When `rv20d > 18` AND `ret_5d > −3%` simultaneously, FPI weight is reduced to 80% of normal (0.15 → 0.12). Prevents a single bearish FPI print from dominating during a genuine recovery phase where price is actually climbing.

---

## API routes

| Method | Route | Description |
|---|---|---|
| `GET` | `/` | Serves `dashboard.html` |
| `GET` | `/portfolio` | Serves `portfolio_hedge_dashboard.html` |
| `GET` | `/api/live` | Live market data — Nifty, VIX, Gold, Brent, USD/INR, DMA, drawdown, prev closes |
| `GET` | `/api/v3_hedge` | Engine output — hedge target %, stage, diagnostics, inputs used |
| `GET` | `/api/v3_chart` | Up to 3,650 days of engine + drawdown history for the chart |
| `GET` | `/api/v3_state` | Raw Redis engine state — for debugging |
| `POST` | `/api/v3_reset` | Resets engine state to factory defaults |
| `POST` | `/api/generate-excel` | Streams portfolio hedge analysis as `.xlsx` (5 sheets) |
| `POST` | `/api/generate-text` | Streams portfolio hedge report as `.txt` |
| `POST` | `/api/generate-word` | Streams portfolio hedge report as `.docx` |
| `GET` | `/health` | Health check — `{"status": "ok"}` |

### `/api/v3_hedge` query parameters

| Parameter | Default | Description |
|---|---|---|
| `fpi` | `-1200` | FPI weekly net equity flow in ₹ Cr (user-entered via dashboard edit field) |
| `beta` | `1.0` | Portfolio beta — scales hedge target for non-Nifty portfolios |

### `/api/v3_hedge` response shape

```json
{
  "status": "success",
  "data": {
    "beta_adjusted_target": 42.5,
    "nifty_base_target": 42.5,
    "active_stage": "Stage 2 — Active Phase",
    "diagnostics": {
      "onset_score": 38.2,
      "s1_target": 28.4,
      "s2_penalty": 1.6,
      "s2_target": 42.5,
      "current_dd": 8.31,
      "gate_locked": false,
      "gate_prev_hedge": 38.0,
      "recovery_ratio": 0.142,
      "low_vol_days": 0
    },
    "peak_date": "2026-01-02",
    "vix_at_hwm": 9.4,
    "vix_now": 17.6,
    "inputs_used": { ... }
  }
}
```

---

## Data sources

| Asset | Primary | Fallback |
|---|---|---|
| Nifty 50 live price | Stooq `^nsei` | yfinance `^NSEI` |
| India VIX | Stooq `^indiavix` | yfinance `^INDIAVIX` |
| USD/INR | Frankfurter API | yfinance `INR=X` → hardcoded 84.50 |
| Gold (USD/oz) | Stooq `xauusd` | yfinance `GC=F` |
| Brent Crude | Stooq `cb.f` | yfinance `BZ=F` |
| 100-DMA + 6M rolling high | yfinance 1-year daily | LKG Redis values |

Stooq `N/D` responses (weekends/holidays) are caught and re-raised to trigger the yfinance fallback chain.

---

## Bootstrap script

`bootstrap_engine_history.py` pre-computes `Engine_Hedge_Target` and `Engine_Stage` for every row in `MARKET_DATA.csv` from 2015 onwards and writes results back to the CSV.

```bash
python bootstrap_engine_history.py                     # incremental — unprocessed rows only
python bootstrap_engine_history.py --full              # force full recompute from 2015
python bootstrap_engine_history.py --from 2022-01-01   # recompute from a specific date
```

A checkpoint file (`engine_state_checkpoint.json`) is saved alongside the CSV so incremental runs can resume exactly from where the last run ended. The bootstrap script has **zero import dependencies on the Flask app** — the engine logic is inlined so it can run standalone.

---

## MARKET_DATA.csv

7,816 rows of daily Indian market data from 1996-01-01 to 2026-05-29.

| Column | Description |
|---|---|
| `Date` | Trading date in DD-MM-YYYY format |
| `Nifty_Close` | Nifty 50 closing price |
| `VIX` | India VIX (implied volatility index) |
| `Nifty_RealVol20d` | 20-day realised volatility, annualised |
| `FPI_Net_Equity_Cr` | Daily FPI net equity flows in ₹ Crore |
| `usdinr` | USD/INR exchange rate |
| `tri` | Nifty Total Return Index |
| `Engine_Hedge_Target` | Pre-computed hedge % from bootstrap (2015 onwards) |
| `Engine_Stage` | Pre-computed stage label from bootstrap |

**The server never writes to this file.** It is a read-only historical record committed to the repo.

---

## Portfolio hedge tool

`portfolio_hedge_dashboard.html` + engine functions in `portfolio_hedge_engine.py`.

Users enter their portfolio holdings (asset type, allocation %, optional beta override). The frontend computes portfolio beta and POSTs to `/api/v3_hedge?beta=X` to get a beta-adjusted hedge target. Three report formats are available for download: Excel (5 sheets), plain text, and Word document.

### Beta defaults

24 asset types with hardcoded default betas (e.g. `large_cap_stock: 1.05`, `mid_cap_stock: 1.35`, `gold_etf_sgb: −0.20`, `debt_liquid: 0.05`). Users can override per holding.

### Report output (Excel, 5 sheets)

1. Portfolio holdings + beta analysis + summary
2. Hedge sizing recommendations ranked by composite score
3. Historical episode scenario analysis (9 episodes, 8 strategies)
4. VIX-level scenario table
5. Strategy notes and methodology

---

## Setup and running

### Local development

```bash
pip install flask flask-cors pandas yfinance requests openpyxl redis numpy scipy python-docx

# No Redis needed — falls back to in-memory dict automatically
python backend.py
```

Server starts at `http://127.0.0.1:5000`.

### Environment variables

| Variable | Required | Description |
|---|---|---|
| `REDIS_URL` | Production only | Redis connection string (e.g. `redis://...`) |
| `PORT` | Optional | Port override — defaults to `5000` |

### Render deployment

1. Set `REDIS_URL` in the Render environment variables panel.
2. Run `python bootstrap_engine_history.py` locally once to populate `Engine_Hedge_Target` in `MARKET_DATA.csv`.
3. Commit the updated CSV and push — Render picks it up on next deploy.
4. The background thread handles all live data updates automatically after boot.

---

## Key design decisions

**Why two separate drawdown metrics?**
`dmaGap` (Nifty vs 100-DMA) is used only in the early-warning checklist as a trend direction signal. `dd_from_peak_pct` (Nifty vs 6-month rolling high) is the engine's actual drawdown input. A market can be 7% off its peak while still above the DMA — `dmaGap` would read near zero and the engine would see no drawdown. They serve distinct roles and are not interchangeable.

**Why triple-sigmoid instead of step-function stages?**
`get_continuous_notional()` ensures hedge allocation rises smoothly and proportionally as onset score increases. There are no cliff-edges at fixed thresholds. The three inflection points at scores 25, 50, and 70 loosely correspond to mild / elevated / crisis stress, but the transition between them is gradual.

**Why freeze SOD values for the gate?**
The de-escalation gate's `recovery_ratio`, `natural_target`, `buffer`, and `dd_pct` are frozen at start-of-day. Without this freeze, intraday price swings would trigger and clear the gate multiple times per session, causing the hedge target to oscillate. Decisions are locked to once per calendar day.

**Why is `float('inf')` replaced with `1e15` in Redis?**
JSON does not support `Infinity`. `trough_price` is stored as the sentinel `1e15` and restored to `float('inf')` on load (`>= 1e14` check).

**Why is `NIFTY_JAN1` dynamic?**
The YTD return tile uses the first trading day close of the current calendar year as its base. Hard-coding this would require a code change every January. The dynamic fetch runs once on server start with a hardcoded fallback.
