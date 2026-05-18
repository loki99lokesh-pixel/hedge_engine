from datetime import datetime, timezone
import os
import time
import threading
import io

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

def get_dma_only() -> float:
    """Fetches 1 year of history specifically for the 100-DMA calculation."""
    global LKG_DMA
    df = yf.download('^NSEI', period='1y', interval='1d', progress=False, threads=False)
    if df.empty:
        raise ValueError('No yfinance history for ^NSEI')
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    close_col = 'Close' if 'Close' in df.columns else df.columns[-1]
    
    # Calculate DMA and update Last Known Good memory
    dma_100 = float(df[close_col].dropna().rolling(window=100).mean().iloc[-1])
    if pd.isna(dma_100):
         raise ValueError('Calculated DMA is NaN')
         
    LKG_DMA = dma_100
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
        dynamic_dma = LKG_DMA

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

    Frontend now sends (new schema):
      name        = display name
      type        = asset_type key  (e.g. 'equity_banking_large' or 'debt_liquid')
      alloc       = allocation %
      beta        = resolved beta mid value (computed by frontend JS)
      betaLow     = beta low bound from matrix (or same as beta for override/flat)
      betaHigh    = beta high bound from matrix
      betaSource  = 'matrix' | 'override' | 'flat' | 'default'
      cat         = asset category ('equity', 'debt', 'gold', 'intl', 'cash')
      sector      = equity sector key (e.g. 'banking') — empty for non-equity
      cap         = cap tier key (e.g. 'large') — empty for non-equity

    Engine expects columns:
      Holding_Name, Asset_Type, Allocation_Pct, Beta, Beta_Low, Beta_High,
      Beta_Source, Sector, Weight, Weighted_Beta
    """
    df = pd.DataFrame(holdings)

    # Readable sector label for non-equity; for equity use the sector key with title-case
    CAT_LABEL = {
        'equity': 'Equity',
        'debt'  : 'Debt / Liquid',
        'gold'  : 'Gold / Hedge',
        'intl'  : 'International',
        'cash'  : 'Cash / FD',
    }
    SECTOR_LABEL = {
        'banking': 'Banking / Finance', 'it': 'IT / Tech', 'fmcg': 'FMCG / Consumer',
        'pharma': 'Pharma', 'auto': 'Auto / EV', 'infra': 'Infra / Capital Goods',
        'metals': 'Metals / Mining', 'energy': 'Energy / Oil & Gas',
        'realty': 'Realty', 'conglomerate': 'Conglomerate',
        'multisector': 'Multi-sector', 'nifty_index': 'Nifty Index',
    }

    df.rename(columns={
        'name'      : 'Holding_Name',
        'type'      : 'Asset_Type',
        'alloc'     : 'Allocation_Pct',
        'beta'      : 'Beta',
        'betaLow'   : 'Beta_Low',
        'betaHigh'  : 'Beta_High',
        'betaSource': 'Beta_Source',
    }, inplace=True)

    # Build readable Sector column: equity rows use sector label, others use cat label
    def resolve_sector(row):
        cat = str(row.get('cat', '')).lower()
        if cat == 'equity':
            sec = str(row.get('sector', '')).lower()
            return SECTOR_LABEL.get(sec, sec.replace('_', ' ').title() if sec else 'Equity')
        return CAT_LABEL.get(cat, cat.title())

    df['Sector'] = df.apply(resolve_sector, axis=1)

    # Drop frontend-only fields no longer needed
    for col in ['cat', 'sector', 'cap']:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Ensure numeric types
    df['Allocation_Pct'] = pd.to_numeric(df['Allocation_Pct'], errors='coerce').fillna(0)
    df['Beta']           = pd.to_numeric(df['Beta'],           errors='coerce').fillna(1.0)
    df['Beta_Low']       = pd.to_numeric(df.get('Beta_Low',  df['Beta']), errors='coerce').fillna(df['Beta'])
    df['Beta_High']      = pd.to_numeric(df.get('Beta_High', df['Beta']), errors='coerce').fillna(df['Beta'])

    # Normalise beta source label to title-case for Excel display
    if 'Beta_Source' in df.columns:
        df['Beta_Source'] = df['Beta_Source'].apply(
            lambda s: {'matrix':'Matrix','override':'Override','flat':'Flat',
                       'default':'Default'}.get(str(s).lower(), str(s).title())
        )
    else:
        df['Beta_Source'] = 'Default'

    total = df['Allocation_Pct'].sum()
    if total > 0 and abs(total - 100) > 0.5:
        df['Allocation_Pct'] = df['Allocation_Pct'] / total * 100

    df['Weight']        = df['Allocation_Pct'] / 100
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
