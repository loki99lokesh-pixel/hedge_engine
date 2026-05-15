from datetime import datetime, timezone
import os
import time
import threading
import io

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

CACHE = None
CACHE_TIMESTAMP = 0
CACHE_TTL = 300
CACHE_LOCK = threading.Lock()

# ══════════════════════════════════════════════════════════════
# DATA FETCHERS
# ══════════════════════════════════════════════════════════════
def attempt_yf(symbol, name):
    df = yf.download(symbol, period='2d', interval='1d', progress=False, threads=False)
    if df.empty:
        raise ValueError(f"yfinance returned empty for {name}")
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
    return float(df[close_col].dropna().iloc[-1])

def attempt_stooq(symbol, name):
    url = f"https://stooq.com/q/l/?s={symbol}&f=sd2t2ohlcv&h&e=csv"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers, timeout=8)
    response.raise_for_status()
    lines = response.text.strip().split('\n')
    if len(lines) < 2:
        raise ValueError(f"Stooq returned no data rows for {name}")
    cols = lines[1].split(',')
    # Stooq CSV: Symbol,Date,Time,Open,High,Low,Close,Volume
    val = float(cols[6] if len(cols) >= 8 else cols[5])
    if val <= 0:
        raise ValueError(f"Stooq returned invalid value {val} for {name}")
    return val

def get_nifty_live_and_dma() -> tuple:
    global LKG_DMA
    df = yf.download('^NSEI', period='1y', interval='1d', progress=False, threads=False)
    if df.empty:
        raise ValueError('No yfinance history for ^NSEI')
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
    series = df[close_col].dropna()
    close = float(series.iloc[-1])
    dma_100 = float(series.rolling(window=100).mean().iloc[-1])
    if pd.isna(close) or pd.isna(dma_100):
        raise ValueError('Calculated Nifty or DMA is NaN')
    LKG_DMA = dma_100
    return close, dma_100

def fetch_live_data():
    data = {}
    sources = {}

    try:
        nifty_close, dynamic_dma = get_nifty_live_and_dma()
        data['nifty'] = nifty_close
        sources['nifty'] = 'yfinance'
    except Exception as e:
        print(f"Nifty primary fetch failed: {e}")
        try:
            data['nifty'] = attempt_stooq('^nsei', 'nifty')
            sources['nifty'] = 'stooq_fallback'
        except Exception as e2:
            print(f"Nifty stooq fallback failed: {e2}")
            data['nifty'] = 24386
            sources['nifty'] = 'static_fallback'
        dynamic_dma = LKG_DMA

    try:
        data['vix'] = attempt_yf('^INDIAVIX', 'vix')
        sources['vix'] = 'yfinance'
    except Exception:
        try:
            data['vix'] = attempt_stooq('^ind.vix', 'vix')
            sources['vix'] = 'stooq_fallback'
        except Exception:
            data['vix'] = 17.9
            sources['vix'] = 'static_fallback'

    try:
        data['usdinr'] = attempt_yf('USDINR=X', 'usdinr')
        sources['usdinr'] = 'yfinance'
    except Exception:
        try:
            res = requests.get('https://api.frankfurter.app/latest?from=USD&to=INR', timeout=8).json()
            data['usdinr'] = res['rates']['INR']
            sources['usdinr'] = 'frankfurter'
        except Exception:
            try:
                res = requests.get('https://open.er-api.com/v6/latest/USD', timeout=8).json()
                data['usdinr'] = res['rates']['INR']
                sources['usdinr'] = 'erapi'
            except Exception:
                data['usdinr'] = 84.50
                sources['usdinr'] = 'static_fallback'

    try:
        data['goldUSD'] = attempt_yf('GC=F', 'goldUSD')
        sources['goldUSD'] = 'yfinance'
    except Exception:
        try:
            data['goldUSD'] = attempt_stooq('xauusd', 'goldUSD')
            sources['goldUSD'] = 'stooq_fallback'
        except Exception:
            data['goldUSD'] = 3295
            sources['goldUSD'] = 'static_fallback'

    try:
        data['brent'] = attempt_yf('BZ=F', 'brent')
        sources['brent'] = 'yfinance'
    except Exception:
        try:
            data['brent'] = attempt_stooq('cb.f', 'brent')
            sources['brent'] = 'stooq_fallback'
        except Exception:
            data['brent'] = 65.0
            sources['brent'] = 'static_fallback'

    data['goldINR'] = round(data['goldUSD'] * data['usdinr'] / 31.1035, 2)
    data['ytd'] = round(((data['nifty'] / NIFTY_JAN1) - 1) * 100, 2)
    data['dmaGap'] = round(((data['nifty'] / dynamic_dma) - 1) * 100, 2)
    data['dma_used'] = round(dynamic_dma, 2)
    data['timestamp'] = datetime.now(timezone.utc).isoformat()
    data['sources'] = sources
    return data

def update_cache_loop():
    global CACHE, CACHE_TIMESTAMP
    while True:
        try:
            payload = fetch_live_data()
            with CACHE_LOCK:
                CACHE = payload
                CACHE_TIMESTAMP = time.time()
        except Exception as e:
            print(f"[Background Sync Error] {e}")
        time.sleep(CACHE_TTL)

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
    global CACHE, CACHE_TIMESTAMP
    with CACHE_LOCK:
        if CACHE is not None and (time.time() - CACHE_TIMESTAMP) < CACHE_TTL:
            return jsonify({'status': 'Success', 'data': CACHE, 'cached': True})
        try:
            payload = fetch_live_data()
            CACHE = payload
            CACHE_TIMESTAMP = time.time()
            return jsonify({'status': 'Success', 'data': payload, 'cached': False})
        except Exception as exc:
            if CACHE is not None:
                return jsonify({'status': 'Success', 'data': CACHE, 'cached': 'Stale',
                                'error_msg': str(exc)})
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
        with CACHE_LOCK:
            cached = CACHE

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

        with CACHE_LOCK:
            cached = CACHE

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

if __name__ == '__main__':
    print("=" * 60)
    print("  Starting Institutional Risk Engine Server")
    print("  Main Dashboard:  http://127.0.0.1:5000/")
    print("  Portfolio Tool:  http://127.0.0.1:5000/portfolio")
    print("=" * 60)
    app.run(debug=True, port=5000, use_reloader=False)