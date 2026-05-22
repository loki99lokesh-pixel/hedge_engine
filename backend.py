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

import yfinance as yf
from flask import Flask, jsonify, send_from_directory, request, send_file, Response

import portfolio_hedge_engine as phe

app = Flask(__name__, static_folder='.', static_url_path='')

try:
    from flask_cors import CORS
    CORS(app)
except ImportError:
    pass

STATE_FILE = "v3_state.json"

def load_v3_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {'peak_price': 0, 'trough_price': float('inf'), 'days_in_dd': 0, 'low_vol_days': 0, 'prev_hedge': 0}

def save_v3_state(state):
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f)

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
    # Failsafe for local testing on your laptop
    db = None 
    LOCAL_CACHE = {}
    LOCAL_DMA = 23521.0

def save_state(key, data):
    if db:
        db.set(key, json.dumps(data))
    else:
        if key == 'dma': 
            global LOCAL_DMA; LOCAL_DMA = data
        else: 
            global LOCAL_CACHE; LOCAL_CACHE[key] = data

def get_state(key):
    if db:
        val = db.get(key)
        return json.loads(val) if val else None
    else:
        if key == 'dma': 
            return LOCAL_DMA
        else: 
            return LOCAL_CACHE.get(key)


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

def get_dma_only() -> float:
    """Fetches 1 year of history specifically for the 100-DMA calculation."""
    df = yf.download('^NSEI', period='1y', interval='1d', progress=False, threads=False)
    if df.empty:
        raise ValueError('No yfinance history for ^NSEI')
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
    
    # Calculate DMA and update database
    dma_100 = float(df[close_col].dropna().rolling(window=100).mean().iloc[-1])
    if pd.isna(dma_100):
         raise ValueError('Calculated DMA is NaN')
         
    save_state('dma', dma_100)
    return dma_100

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

    # 2. NIFTY 100-DMA (yfinance background -> LKG Memory Fallback)
    try:
        dynamic_dma = get_dma_only()
    except Exception as e:
        print(f"DMA math failed, using LKG: {e}")
        dynamic_dma = get_state('dma') or 23521.0

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
        nifty_close = cached.get('nifty', 24000)
        gap_pct     = cached.get('dmaGap', 0)
        vix         = cached.get('vix', 18.5)

        # rv20d: use India VIX as the best available realised-vol proxy
        # (actual 20-day realised vol is not fetched separately; VIX is highly correlated)
        rv20d = vix

        # 5-day return: derive from current price vs 5-day-ago close if available,
        # otherwise fall back to 0 (neutral — no Mixed Signal Filter bias)
        ret_5d = cached.get('ret_5d', 0.0)

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

        # 5. Execute the v3.2 magnitude engine
        nifty_hedge, active_stage, diag, new_state = phe.calculate_v3_magnitude_hedge(
            vix, rv20d, fpi, gap_pct, nifty_close, state, ret_5d
        )

        # 6. Scale by portfolio beta (1.0 for dashboard = no scaling, pure Nifty target)
        adjusted_hedge = min(100.0, nifty_hedge * portfolio_beta)

        # 7. Persist updated state for next call
        save_v3_state(new_state)

        return jsonify({
            'status': 'success',
            'data': {
                'beta_adjusted_target': round(adjusted_hedge, 1),
                'nifty_base_target':    round(nifty_hedge, 1),
                'active_stage':         active_stage,
                'diagnostics':          diag,
                'inputs_used': {
                    'vix':         round(vix, 1),
                    'rv20d':       round(rv20d, 1),
                    'fpi_weekly':  fpi,
                    'gap_pct':     round(gap_pct, 2),
                    'ret_5d':      round(ret_5d, 2),
                    'nifty_close': round(nifty_close, 0),
                }
            }
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/v3_reset', methods=['POST'])
def reset_v3_state(): 
    #Clears the persisted v3.2 state file so the engine starts fresh.
    #Call this whenever stale prev_hedge values are contaminating the output
    #(e.g. after the old hardcoded-input era, or after a deployment change).
    try: 
        fresh_state = {
            'peak_price' : 0, 
            'trough_price': float('inf'), 
            'days_in_dd' : 0, 
            'low_vol_days': 0, 
            'prev_hedge' : 0, 
        } 
        save_v3_state(fresh_state) 
        return jsonify({'status': 'success', 'message': 'v3 state reset to fresh defaults.'}) 
    except Exception as e: 
        return jsonify({'status': 'error', 'message': str(e)}), 500 


if __name__ == '__main__':
    print("=" * 60)
    print("  Starting Institutional Risk Engine Server")
    print("  Main Dashboard:  http://127.0.0.1:5000/")
    print("  Portfolio Tool:  http://127.0.0.1:5000/portfolio")
    print("=" * 60)
    app.run(debug=True, port=5000, use_reloader=False)
