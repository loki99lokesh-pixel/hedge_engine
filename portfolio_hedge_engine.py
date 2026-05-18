"""
portfolio_hedge_engine.py  v2.0
================================
Applies the Nifty 50 drawdown hedge analysis to YOUR actual portfolio.
Now features dynamic 100-DMA calculation via live market data.

HOW IT WORKS:
  1. You define your portfolio as a list of holdings (asset type, sector, % allocation)
  2. Each holding gets a Nifty Beta — the sensitivity of that holding to Nifty moves
  3. Beta-weighted Nifty Exposure is computed for the whole portfolio
  4. All 8 hedge strategies are re-scaled to YOUR portfolio's actual exposure
  5. If you provide a total portfolio value (optional), all outputs are in ₹ as well
"""

import os, sys, argparse, textwrap
from datetime import date
import warnings
import numpy as np
import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side, GradientFill)
from openpyxl.utils import get_column_letter

# Suppress yfinance timezone warnings for a cleaner CLI
warnings.filterwarnings('ignore', category=FutureWarning)

# ══════════════════════════════════════════════════════════════
# LIVE DATA MODULE
# ══════════════════════════════════════════════════════════════
def fetch_live_nifty_status():
    """Dynamically calculates the Nifty 100-DMA and checks the hedge trigger."""
    print("  [Live API] Fetching Nifty 50 data for dynamic 100-DMA calculation...")
    try:
        # Fetch 6 months of daily data to ensure we have 100 trading sessions
        ticker = yf.Ticker("^NSEI")
        hist = ticker.history(period="6mo")
        
        if hist.empty or len(hist) < 100:
            print("  Warning: Not enough data fetched for 100-DMA calculation.")
            return None

        # Calculate the 100-Day Simple Moving Average
        hist['100_DMA'] = hist['Close'].rolling(window=100).mean()
        
        latest_close = hist['Close'].iloc[-1]
        latest_100dma = hist['100_DMA'].iloc[-1]
        
        # Calculate the percentage gap between current price and 100-DMA
        gap_pct = ((latest_close - latest_100dma) / latest_100dma) * 100
        
        # Determine if the tactical trigger is met (Nifty is below 100-DMA by at least 2%)
        trigger_active = gap_pct <= -2.0 

        return {
            "close": round(latest_close, 2),
            "dma100": round(latest_100dma, 2),
            "gap_pct": round(gap_pct, 2),
            "trigger_active": trigger_active
        }
    except Exception as e:
        print(f"  Warning: Failed to fetch live Nifty data: {e}")
        return None

# ══════════════════════════════════════════════════════════════
# BETA MATRIX — Sector × Market Cap (NSE historical, 2005-2025)
# Each cell: (mid, low, high)  — mid is the point estimate used
# Non-equity asset types use flat lookup below
# ══════════════════════════════════════════════════════════════
BETA_MATRIX = {
    # sector       : { cap   : (mid,  low,  high) }
    "banking"      : { "large":(1.15, 0.95, 1.35), "mid":(1.35, 1.10, 1.60), "small":(1.55, 1.25, 1.80), "micro":(1.65, 1.35, 1.95) },
    "it"           : { "large":(1.10, 0.90, 1.30), "mid":(1.30, 1.05, 1.55), "small":(1.50, 1.20, 1.75), "micro":(1.60, 1.30, 1.90) },
    "fmcg"         : { "large":(0.70, 0.55, 0.85), "mid":(0.85, 0.65, 1.05), "small":(1.00, 0.75, 1.25), "micro":(1.10, 0.85, 1.35) },
    "pharma"       : { "large":(0.75, 0.60, 0.90), "mid":(0.90, 0.70, 1.10), "small":(1.10, 0.85, 1.35), "micro":(1.20, 0.90, 1.50) },
    "auto"         : { "large":(1.10, 0.90, 1.30), "mid":(1.30, 1.05, 1.55), "small":(1.50, 1.20, 1.75), "micro":(1.60, 1.30, 1.90) },
    "infra"        : { "large":(1.20, 1.00, 1.40), "mid":(1.45, 1.15, 1.70), "small":(1.65, 1.35, 1.90), "micro":(1.75, 1.45, 2.00) },
    "metals"       : { "large":(1.25, 1.05, 1.50), "mid":(1.50, 1.20, 1.75), "small":(1.70, 1.40, 2.00), "micro":(1.80, 1.50, 2.10) },
    "energy"       : { "large":(0.95, 0.75, 1.15), "mid":(1.15, 0.90, 1.40), "small":(1.35, 1.05, 1.65), "micro":(1.45, 1.15, 1.75) },
    "realty"       : { "large":(1.30, 1.10, 1.55), "mid":(1.55, 1.25, 1.80), "small":(1.75, 1.45, 2.05), "micro":(1.85, 1.55, 2.15) },
    "conglomerate" : { "large":(1.00, 0.80, 1.20), "mid":(1.20, 0.95, 1.45), "small":(1.40, 1.10, 1.70), "micro":(1.50, 1.20, 1.80) },
    "multisector"  : { "large":(1.05, 0.85, 1.25), "mid":(1.25, 1.00, 1.50), "small":(1.45, 1.15, 1.75), "micro":(1.55, 1.25, 1.85) },
    "nifty_index"  : { "large":(1.00, 0.97, 1.03), "mid":(1.00, 0.97, 1.03), "small":(1.00, 0.97, 1.03), "micro":(1.00, 0.97, 1.03) },
}

# Flat beta lookup for non-equity / special asset classes
# Each entry: (mid, low, high)
BETA_FLAT = {
    "debt_liquid"  : ( 0.05,  0.02,  0.07),
    "debt_short"   : ( 0.08,  0.05,  0.12),
    "debt_gilt"    : ( 0.12,  0.08,  0.18),
    "gold_etf"     : (-0.20, -0.35, -0.05),
    "gold_fund"    : (-0.18, -0.32, -0.04),
    "reit_invit"   : ( 0.55,  0.40,  0.70),
    "intl_fund"    : ( 0.35,  0.20,  0.50),
    "us_tech"      : ( 0.30,  0.15,  0.45),
    "cash_fd"      : ( 0.00,  0.00,  0.00),
    # Legacy keys from old CSV format — kept for backward-compat
    "nifty_etf"           : (1.00, 0.97, 1.03),
    "large_cap_stock"     : (1.05, 0.85, 1.25),
    "mid_cap_stock"       : (1.35, 1.05, 1.65),
    "small_cap_stock"     : (1.55, 1.25, 1.85),
    "multi_cap_fund"      : (1.10, 0.88, 1.32),
    "flexi_cap_fund"      : (1.05, 0.85, 1.25),
    "sectoral_it"         : (1.20, 0.98, 1.42),
    "sectoral_banking"    : (1.15, 0.95, 1.35),
    "sectoral_pharma"     : (0.80, 0.62, 0.98),
    "sectoral_fmcg"       : (0.70, 0.55, 0.85),
    "sectoral_auto"       : (1.25, 1.00, 1.50),
    "sectoral_infra"      : (1.30, 1.05, 1.55),
    "sectoral_realty"     : (1.45, 1.18, 1.72),
    "sectoral_metals"     : (1.40, 1.12, 1.68),
    "debt_shortterm"      : (0.08, 0.05, 0.12),
    "debt_liquid"         : (0.05, 0.02, 0.07),
    "debt_shortterm"      : (0.08, 0.05, 0.12),
    "gold_etf_sgb"        : (-0.20, -0.35, -0.05),
    "gold_fund"           : (-0.18, -0.32, -0.04),
    "international_fund"  : (0.35, 0.20, 0.50),
    "us_tech_fund"        : (0.30, 0.15, 0.45),
    "pms_aif"             : (1.10, 0.85, 1.35),
}

# Keep BETA_DEFAULTS as a simple mid-value dict for any code that still uses it
BETA_DEFAULTS = {k: v[0] for k, v in BETA_FLAT.items()}


def resolve_beta(asset_type: str, sector: str = "", cap: str = "large") -> tuple:
    """
    Resolve (mid, low, high) beta for a holding.

    Priority:
      1. equity_<sector>_<cap> composite key  (from frontend new-schema)
      2. BETA_MATRIX[sector][cap]             (direct matrix lookup)
      3. BETA_FLAT[asset_type]                (flat non-equity lookup)
      4. Fallback (1.0, 0.8, 1.2)
    """
    atype = str(asset_type).strip().lower()

    # New composite key from frontend: "equity_banking_large"
    if atype.startswith("equity_"):
        parts = atype.split("_", 2)
        if len(parts) == 3:
            _, sec, cp = parts
            row = BETA_MATRIX.get(sec)
            if row:
                cell = row.get(cp) or row.get("large")
                if cell:
                    return cell  # (mid, low, high)

    # Direct BETA_MATRIX lookup using separate sector/cap args
    if sector and sector in BETA_MATRIX:
        cap_key = cap if cap in ("large", "mid", "small", "micro") else "large"
        cell = BETA_MATRIX[sector].get(cap_key)
        if cell:
            return cell

    # Flat lookup (non-equity types + legacy keys)
    if atype in BETA_FLAT:
        return BETA_FLAT[atype]

    return (1.00, 0.80, 1.20)  # safe fallback

# ══════════════════════════════════════════════════════════════
# HEDGE STRATEGY RESULTS FROM BACK-TEST (30yr, v3.1)
# ══════════════════════════════════════════════════════════════
STRATEGIES = {
    "S1_ShortFutures": {
        "name"        : "Short Nifty Futures (100%)",
        "short_name"  : "Short Futures",
        "win_rate"    : 1.00,
        "avg_pnl_pct" : 23.34,
        "worst_pct"   : 10.17,
        "best_pct"    : 59.86,
        "avg_cost_pct": 0.20,
        "monthly_drag": 0.20,
        "liquidity"   : 9,
        "simplicity"  : 8,
        "score"       : 77.5,
        "requires"    : "Futures account (F&O enabled)",
        "trigger"     : "VIX > 28 (Strategy C)",
        "note"        : "Full offset but unlimited upside sacrifice. Needs 25% margin buffer.",
        "beta_scale"  : True,
    },
    "S2_ATM_Put": {
        "name"        : "ATM Protective Put",
        "short_name"  : "ATM Put",
        "win_rate"    : 1.00,
        "avg_pnl_pct" : 21.73,
        "worst_pct"   : 8.25,
        "best_pct"    : 57.33,
        "avg_cost_pct": 1.20,
        "monthly_drag": 1.20,
        "liquidity"   : 7,
        "simplicity"  : 6,
        "score"       : 71.6,
        "requires"    : "F&O account",
        "trigger"     : "2-of-3 signals (Strategy B)",
        "note"        : "Best protection for known catalyst. Cost scales with VIX.",
        "beta_scale"  : True,
    },
    "S3_OTM5_Put": {
        "name"        : "5% OTM Protective Put",
        "short_name"  : "5% OTM Put",
        "win_rate"    : 1.00,
        "avg_pnl_pct" : 18.01,
        "worst_pct"   : 3.25,
        "best_pct"    : 54.06,
        "avg_cost_pct": 0.70,
        "monthly_drag": 0.70,
        "liquidity"   : 6,
        "simplicity"  : 6,
        "score"       : 65.0,
        "requires"    : "F&O account",
        "trigger"     : "2-of-3 signals (Strategy B backup)",
        "note"        : "Cheaper than ATM. First 5% loss unhedged. Good for deep crashes.",
        "beta_scale"  : True,
    },
    "S4_Gold15": {
        "name"        : "Gold Allocation (15%)",
        "short_name"  : "Gold 15%",
        "win_rate"    : 0.688,
        "avg_pnl_pct" : 0.90,
        "worst_pct"   : -2.75,
        "best_pct"    : 5.77,
        "avg_cost_pct": 0.05,
        "monthly_drag": 0.05,
        "liquidity"   : 8,
        "simplicity"  : 9,
        "score"       : 38.4,
        "requires"    : "Demat (Gold ETF / SGB)",
        "trigger"     : "Always-on (Strategy A)",
        "note"        : "Positive in 69% of drawdowns. Low cost. Global crisis hedge.",
        "beta_scale"  : False,
    },
    "S5_USDINR10": {
        "name"        : "USD/INR Long (10%)",
        "short_name"  : "USD/INR 10%",
        "win_rate"    : 0.938,
        "avg_pnl_pct" : 0.59,
        "worst_pct"   : -0.04,
        "best_pct"    : 2.68,
        "avg_cost_pct": 0.10,
        "monthly_drag": 0.10,
        "liquidity"   : 7,
        "simplicity"  : 7,
        "score"       : 41.8,
        "requires"    : "International fund / USD FD",
        "trigger"     : "INR stress (USD/INR > 88)",
        "note"        : "Useful for FX-driven drawdowns. 94% win rate across episodes.",
        "beta_scale"  : False,
    },
    "S6_Debt30": {
        "name"        : "Debt Rotation (30%)",
        "short_name"  : "Debt 30%",
        "win_rate"    : 1.00,
        "avg_pnl_pct" : 7.92,
        "worst_pct"   : 3.20,
        "best_pct"    : 19.52,
        "avg_cost_pct": 0.30,
        "monthly_drag": 0.10,
        "liquidity"   : 9,
        "simplicity"  : 9,
        "score"       : 57.1,
        "requires"    : "Liquid / overnight debt fund",
        "trigger"     : "Always-on (Strategy A core)",
        "note"        : "100% win rate. Reduces equity beta + earns 6.5% p.a. Simplest hedge.",
        "beta_scale"  : False,
    },
    "S7_Collar": {
        "name"        : "Collar (ATM Put + OTM Call)",
        "short_name"  : "Collar",
        "win_rate"    : 1.00,
        "avg_pnl_pct" : 23.03,
        "worst_pct"   : 9.45,
        "best_pct"    : 58.53,
        "avg_cost_pct": 0.60,
        "monthly_drag": 0.60,
        "liquidity"   : 5,
        "simplicity"  : 4,
        "score"       : 70.0,
        "requires"    : "F&O account (2 legs)",
        "trigger"     : "Pre-known events (elections, budgets)",
        "note"        : "Sells upside to pay for put. Ideal when you need cheap crash protection.",
        "beta_scale"  : True,
    },
    "S8_Combined": {
        "name"        : "Combined (Put 50% + Gold 15% + Debt 20%)",
        "short_name"  : "Combined",
        "win_rate"    : 1.00,
        "avg_pnl_pct" : 16.83,
        "worst_pct"   : 6.26,
        "best_pct"    : 42.76,
        "avg_cost_pct": 0.80,
        "monthly_drag": 0.80,
        "liquidity"   : 6,
        "simplicity"  : 5,
        "score"       : 63.2,
        "requires"    : "F&O + Demat + Debt fund",
        "trigger"     : "2-of-3 signals active",
        "note"        : "Diversified hedge. No single point of failure. Recommended for HNIs.",
        "beta_scale"  : True,
    },
}

# Historical episodes for per-episode scenario analysis
EPISODES = [
    {"label":"2000 Dot-com",    "dd":-51.36, "dur":588, "rec":818},
    {"label":"2004 Election",   "dd":-29.94, "dur":124, "rec":199},
    {"label":"2008 GFC",        "dd":-59.86, "dur":293, "rec":739},
    {"label":"2010 Euro Debt",  "dd":-28.01, "dur":410, "rec":684},
    {"label":"2015 China",      "dd":-22.52, "dur":359, "rec":383},
    {"label":"2018 IL&FS",      "dd":-10.17, "dur":53,  "rec":123},
    {"label":"2020 COVID",      "dd":-38.44, "dur":69,  "rec":231},
    {"label":"2024 FII Exodus", "dd":-15.77, "dur":159, "rec":304},
    {"label":"2026 Curr. DD",   "dd":-15.18, "dur":87,  "rec":None},
]

# Actual hedge P&L (net) per episode per strategy
EPISODE_PNL = {
    "2000 Dot-com"   : [51.36, 48.75, 45.50,  None,  None, 18.55, 49.95, 36.74],
    "2004 Election"  : [29.94, 28.16, 24.44, -1.47,  0.02,  9.64, 29.36, 19.04],
    "2008 GFC"       : [59.86, 57.33, 54.06,  1.07,  2.68, 19.52, 58.53, 42.76],
    "2010 Euro Debt" : [28.01, 26.04, 22.57,  5.77,  1.97, 10.59, 28.01, 25.85],
    "2015 China"     : [22.52, 20.99, 17.31,  2.04,  1.05,  8.67, 22.52, 18.32],
    "2018 IL&FS"     : [10.17,  8.36,  4.82,  0.35,  0.25,  3.33, 10.17,  6.75],
    "2020 COVID"     : [38.44, 37.09, 33.30,  1.29,  0.69, 11.90, 38.44, 27.77],
    "2024 FII Exodus": [15.77, 14.63, 10.70,  2.07,  0.44,  5.58, 15.77, 13.11],
    "2026 Curr. DD"  : [15.18, 14.33, 10.16,  1.58,  0.54,  5.02, 15.18, 12.09],
}

STRAT_KEYS = list(STRATEGIES.keys())

# ══════════════════════════════════════════════════════════════
# SAMPLE PORTFOLIO TEMPLATE
# ══════════════════════════════════════════════════════════════
SAMPLE_PORTFOLIO = [
    ("Nifty BeES ETF",           "nifty_etf",         "Index",         20.0,   None),
    ("HDFC Bank",                "large_cap_stock",   "Banking",        8.0,   1.10),
    ("Infosys",                  "large_cap_stock",   "IT",             7.0,   1.15),
    ("Reliance Industries",      "large_cap_stock",   "Conglomerate",   6.0,   1.05),
    ("ICICI Pru Mid Cap Fund",   "mid_cap_stock",     "Multi-sector",   5.0,   None),
    ("SBI Small Cap Fund",       "small_cap_stock",   "Multi-sector",   5.0,   None),
    ("Titan Company",            "large_cap_stock",   "Consumer",       4.0,   0.95),
    ("Adani Ports",              "large_cap_stock",   "Infra",          4.0,   1.20),
    ("Nippon IT ETF",            "sectoral_it",       "IT",             3.0,   None),
    ("Kotak Flexicap Fund",      "flexi_cap_fund",    "Multi-sector",   5.0,   None),
    ("HDFC Liquid Fund",         "debt_liquid",       "Debt",          10.0,   None),
    ("SBI Short Term Debt Fund", "debt_shortterm",    "Debt",           5.0,   None),
    ("Sovereign Gold Bond 2028", "gold_etf_sgb",      "Gold",           5.0,   None),
    ("Nippon Gold ETF",          "gold_etf_sgb",      "Gold",           3.0,   None),
    ("Embassy REIT",             "reit_invit",        "Real Estate",    3.0,   None),
    ("Mirae US Tech Fund",       "us_tech_fund",      "International",  4.0,   None),
    ("Cash / FD",                "cash_fd",           "Cash",           3.0,   None),
]

def generate_template_csv(path: str):
    rows = []
    for name, atype, sector, alloc, beta in SAMPLE_PORTFOLIO:
        rows.append({
            "Holding_Name"   : name,
            "Asset_Type"     : atype,
            "Sector"         : sector,
            "Allocation_Pct" : alloc,
            "Beta_Override"  : beta if beta is not None else "",
        })
    df = pd.DataFrame(rows)
    df.to_csv(path, index=False)
    print(f"  Template portfolio saved: {path}")

def load_portfolio(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    total = df["Allocation_Pct"].sum()
    if abs(total - 100) > 0.5:
        print(f"  Warning: Allocations sum to {total:.1f}% — normalising to 100%")
        df["Allocation_Pct"] = df["Allocation_Pct"] / total * 100

    def get_beta_row(row):
        """Return (mid, low, high, source) for a holding row."""
        # 1. Manual override wins
        ov = row.get("Beta_Override", "")
        if ov != "" and not pd.isna(ov):
            try:
                v = float(ov)
                return v, v, v, "Override"
            except Exception:
                pass
        # 2. Matrix / flat resolution
        atype  = str(row.get("Asset_Type", "")).strip().lower()
        sector = str(row.get("Sector", "")).strip().lower().replace(" ", "_").replace("/","_").replace("&","_")
        cap    = str(row.get("Cap_Tier", "large")).strip().lower()
        mid, low, high = resolve_beta(atype, sector, cap)
        # Determine source label
        if atype.startswith("equity_") or (sector in BETA_MATRIX and atype not in BETA_FLAT):
            source = "Matrix"
        elif atype in BETA_FLAT:
            source = "Flat"
        else:
            source = "Default"
        return mid, low, high, source

    beta_data = df.apply(get_beta_row, axis=1, result_type="expand")
    beta_data.columns = ["Beta", "Beta_Low", "Beta_High", "Beta_Source"]
    df = pd.concat([df, beta_data], axis=1)

    df["Weight"]       = df["Allocation_Pct"] / 100
    df["Weighted_Beta"]= df["Weight"] * df["Beta"]
    return df

def compute_portfolio_metrics(df: pd.DataFrame, portfolio_value: float = None):
    portfolio_beta   = df["Weighted_Beta"].sum()
    # Beta confidence band — weighted sum of each holding's low/high
    beta_low  = (df["Weight"] * df.get("Beta_Low",  df["Beta"])).sum()
    beta_high = (df["Weight"] * df.get("Beta_High", df["Beta"])).sum()

    equity_weight    = df[df["Beta"] > 0.3]["Weight"].sum()
    hedge_weight     = df[df["Beta"] <= 0]["Weight"].sum()
    debt_weight      = df[(df["Beta"] > 0) & (df["Beta"] < 0.3)]["Weight"].sum()
    sector_breakdown = df.groupby("Sector")["Allocation_Pct"].sum().sort_values(ascending=False)
    top3_pct         = df.nlargest(3, "Allocation_Pct")["Allocation_Pct"].sum()

    metrics = {
        "portfolio_beta"    : round(portfolio_beta, 3),
        "beta_low"          : round(beta_low, 3),
        "beta_high"         : round(beta_high, 3),
        "equity_weight"     : round(equity_weight * 100, 1),
        "debt_weight"       : round(debt_weight * 100, 1),
        "hedge_weight"      : round(hedge_weight * 100, 1),
        "top3_concentration": round(top3_pct, 1),
        "sector_breakdown"  : sector_breakdown,
        "portfolio_value"   : portfolio_value,
        "num_holdings"      : len(df),
    }
    return metrics

def compute_hedge_sizing(portfolio_beta: float, portfolio_value: float = None):
    results = {}
    for key, strat in STRATEGIES.items():
        is_beta_scaled = strat.get("beta_scale", True) 

        if is_beta_scaled:
            effective_coverage_pct = strat["avg_pnl_pct"] * portfolio_beta
            worst_pct              = strat["worst_pct"]   * portfolio_beta
            best_pct               = strat["best_pct"]    * portfolio_beta
            monthly_drag_pct       = strat["monthly_drag"] * portfolio_beta
            drag_pa                = strat["monthly_drag"] * 12 * portfolio_beta
        else:
            effective_coverage_pct = strat["avg_pnl_pct"]
            worst_pct              = strat["worst_pct"]
            best_pct               = strat["best_pct"]
            monthly_drag_pct       = strat["monthly_drag"]
            drag_pa                = strat["monthly_drag"] * 12

        nifty_equiv_pct = portfolio_beta * 100 

        r = {
            "strategy_name"           : strat["name"],
            "short_name"              : strat["short_name"],
            "portfolio_beta"          : portfolio_beta,
            "is_beta_scaled"          : is_beta_scaled,
            "nifty_equiv_exposure_pct": round(nifty_equiv_pct, 1),
            "avg_protection_pct"      : round(effective_coverage_pct, 2),
            "worst_protection_pct"    : round(worst_pct, 2),
            "best_protection_pct"     : round(best_pct, 2),
            "monthly_drag_pct"        : round(monthly_drag_pct, 3),
            "annual_drag_pct"         : round(drag_pa, 2),
            "win_rate_pct"            : round(strat["win_rate"] * 100, 1),
            "composite_score"         : strat["score"],
            "requires"                : strat["requires"],
            "trigger"                 : strat["trigger"],
            "note"                    : strat["note"],
        }
        if portfolio_value:
            nifty_notional = portfolio_value * portfolio_beta
            r["nifty_equiv_exposure_inr"] = round(nifty_notional, 0)
            r["avg_protection_inr"]        = round(portfolio_value * effective_coverage_pct / 100, 0)
            r["monthly_drag_inr"]          = round(portfolio_value * monthly_drag_pct / 100, 0)
            r["annual_drag_inr"]           = round(portfolio_value * drag_pa / 100, 0)

        results[key] = r
    return results

def compute_scenario_analysis(portfolio_beta: float, portfolio_value: float = None):
    rows = []
    strat_list = list(STRATEGIES.keys())
    for ep in EPISODES:
        label = ep["label"]
        dd    = ep["dd"]
        port_loss_pct = max(dd * portfolio_beta, -100.0)
        row = {
            "Episode"           : label,
            "Nifty_DD_Pct"      : dd,
            "Portfolio_Loss_Pct": round(port_loss_pct, 2),
            "Duration_Days"     : ep["dur"],
            "Recovery_Days"     : ep["rec"] if ep["rec"] else "Ongoing",
        }
        ep_pnl = EPISODE_PNL.get(label, [None]*8)
        for i, sk in enumerate(strat_list):
            strat        = STRATEGIES[sk]
            is_beta_scaled = strat.get("beta_scale", True)
            raw_pnl      = ep_pnl[i]

            if raw_pnl is None:
                pnl_scaled   = None
                net_position = None
                offset_pct   = None
            else:
                pnl_scaled   = raw_pnl * portfolio_beta if is_beta_scaled else raw_pnl
                net_position = port_loss_pct + pnl_scaled
                offset_pct   = (pnl_scaled / abs(port_loss_pct) * 100) if port_loss_pct != 0 else 0

            sn = strat["short_name"]
            row[f"{sn}_HedgePnL_Pct"] = round(pnl_scaled, 2)   if pnl_scaled   is not None else None
            row[f"{sn}_NetPos_Pct"]   = round(net_position, 2)  if net_position is not None else None
            row[f"{sn}_Offset_Pct"]   = round(offset_pct, 1) if offset_pct is not None else None
            if portfolio_value:
                row[f"{sn}_HedgePnL_INR"] = round(portfolio_value * pnl_scaled / 100, 0) if pnl_scaled is not None else None
        if portfolio_value:
            row["Portfolio_Loss_INR"] = round(portfolio_value * port_loss_pct / 100, 0)
        rows.append(row)
    return pd.DataFrame(rows)

# ══════════════════════════════════════════════════════════════
# EXCEL WORKBOOK BUILDER
# ══════════════════════════════════════════════════════════════
def fmt_inr(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "—"
    val = float(val)
    if abs(val) >= 1e7:
        return f"₹{val/1e7:.2f}Cr"
    elif abs(val) >= 1e5:
        return f"₹{val/1e5:.2f}L"
    else:
        return f"₹{val:,.0f}"

def _style_header(cell, bg="1E3A5F", fg="FFFFFF", bold=True, size=10, wrap=False):
    cell.font = Font(name="Calibri", bold=bold, color=fg, size=size)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=wrap)

def _border(cell, style="thin"):
    s = Side(style=style, color="D0D0D0")
    cell.border = Border(left=s, right=s, top=s, bottom=s)

def _num_cell(cell, value, fmt="0.00", bold=False, color="000000"):
    cell.value = value if not isinstance(value, str) else value
    cell.number_format = fmt
    cell.font = Font(name="Calibri", size=10, bold=bold, color=color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    _border(cell)

def build_excel(df_portfolio, metrics, hedge_sizing, scenario_df, output_path):
    wb = Workbook()

    # ── Sheet 1: Portfolio Holdings ──────────────────────────
    ws1 = wb.active
    ws1.title = "1_Portfolio Holdings"
    ws1.freeze_panes = "A3"

    ws1.merge_cells("A1:J1")
    t = ws1["A1"]
    t.value = "PORTFOLIO HEDGE ENGINE — Holdings & Beta Matrix Analysis"
    t.font  = Font(name="Calibri", bold=True, size=14, color="1E3A5F")
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[1].height = 28

    headers = ["#", "Holding Name", "Asset Type", "Sector",
               "Allocation %", "Beta (Mid)", "Beta Low", "Beta High", "Beta Source", "Weighted Beta"]
    col_w   = [4, 28, 22, 18, 13, 11, 10, 10, 12, 15]
    for j, (h, w) in enumerate(zip(headers, col_w), 1):
        c = ws1.cell(row=2, column=j, value=h)
        _style_header(c)
        ws1.column_dimensions[get_column_letter(j)].width = w

    for i, row in df_portfolio.iterrows():
        r = i + 3
        beta       = row["Beta"]
        beta_low   = row.get("Beta_Low",  beta)
        beta_high  = row.get("Beta_High", beta)
        beta_src   = row.get("Beta_Source", "—")
        if beta < 0:
            bg = "E8F5E9"
        elif beta < 0.1:
            bg = "F5F5F5"
        else:
            bg = "FFFFFF"

        data = [i+1, row["Holding_Name"], row["Asset_Type"], row.get("Sector",""),
                row["Allocation_Pct"], beta, beta_low, beta_high, beta_src, row["Weighted_Beta"]]
        for j, val in enumerate(data, 1):
            c = ws1.cell(row=r, column=j, value=val)
            c.font = Font(name="Calibri", size=10)
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="center" if j != 2 else "left",
                                    vertical="center")
            _border(c)
            if j == 5:
                c.number_format = "0.0\"%\""
            elif j in (6, 7, 8, 10):
                c.number_format = "0.00"
            # Colour the Beta Source cell
            if j == 9:
                if str(val) == "Override":
                    c.font = Font(name="Calibri", size=10, color="D97706", bold=True)
                elif str(val) == "Matrix":
                    c.font = Font(name="Calibri", size=10, color="2563EB")

    sr = len(df_portfolio) + 5
    ws1.merge_cells(f"A{sr}:B{sr}")
    ws1[f"A{sr}"].value = "PORTFOLIO SUMMARY"
    ws1[f"A{sr}"].font  = Font(name="Calibri", bold=True, size=11, color="1E3A5F")

    summary_data = [
        ("Portfolio Beta (β)",       f"{metrics['portfolio_beta']:.3f}"),
        ("Beta Range (low – high)",  f"{metrics.get('beta_low', metrics['portfolio_beta']):.3f} – {metrics.get('beta_high', metrics['portfolio_beta']):.3f}"),
        ("⚠ Beta Creep Note",       "Mid/small cap betas rise during drawdowns — use Beta High for stress sizing"),
        ("Equity Exposure",          f"{metrics['equity_weight']:.1f}%"),
        ("Debt / Liquid",            f"{metrics['debt_weight']:.1f}%"),
        ("Gold / Inverse Hedge",     f"{metrics['hedge_weight']:.1f}%"),
        ("Top-3 Concentration",      f"{metrics['top3_concentration']:.1f}%"),
        ("Number of Holdings",       str(metrics["num_holdings"])),
        ("Portfolio Value",          fmt_inr(metrics["portfolio_value"]) if metrics["portfolio_value"] else "Not provided"),
    ]
    for k, (lbl, val) in enumerate(summary_data):
        row_n = sr + 1 + k
        c1 = ws1.cell(row=row_n, column=1, value=lbl)
        c2 = ws1.cell(row=row_n, column=2, value=val)
        # Special styling for beta-related rows
        if "Range" in lbl:
            c1.font = Font(name="Calibri", bold=True, size=10, color="2563EB")
            c2.font = Font(name="Calibri", size=10, color="2563EB")
        elif "Creep" in lbl:
            c1.font = Font(name="Calibri", bold=True, size=10, color="D97706")
            c2.font = Font(name="Calibri", size=10,  color="D97706")
            c1.fill = PatternFill("solid", fgColor="FFFBEB")
            c2.fill = PatternFill("solid", fgColor="FFFBEB")
        else:
            c1.font = Font(name="Calibri", bold=True, size=10)
            c2.font = Font(name="Calibri", size=10, color="1E3A5F")
        for c in (c1, c2):
            c.alignment = Alignment(vertical="center", wrap_text=True)
            _border(c)
        ws1.row_dimensions[row_n].height = 18

    # ── Sheet 2: Hedge Sizing Recommendations ────────────────
    ws2 = wb.create_sheet("2_Hedge Recommendations")
    ws2.freeze_panes = "A3"

    ws2.merge_cells("A1:K1")
    t2 = ws2["A1"]
    t2.value = (f"HEDGE SIZING RECOMMENDATIONS  |  Portfolio β = {metrics['portfolio_beta']:.3f}  |  "
                f"Nifty-Equivalent Exposure: {metrics['portfolio_beta']*100:.0f}% of portfolio")
    t2.font = Font(name="Calibri", bold=True, size=12, color="1E3A5F")
    t2.alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[1].height = 24

    h2 = ["Rank", "Strategy", "Win Rate", "Avg Hedge Gain", "Worst Case",
          "Best Case", "Monthly Drag", "Annual Drag", "Score", "Trigger", "Requires"]
    cw2 = [6, 32, 10, 16, 14, 14, 13, 13, 9, 30, 30]
    for j, (h, w) in enumerate(zip(h2, cw2), 1):
        c = ws2.cell(row=2, column=j, value=h)
        _style_header(c)
        ws2.column_dimensions[get_column_letter(j)].width = w

    sorted_keys = sorted(STRATEGIES.keys(), key=lambda k: STRATEGIES[k]["score"], reverse=True)
    for rank, sk in enumerate(sorted_keys, 1):
        r = rank + 2
        hs = hedge_sizing[sk]
        st = STRATEGIES[sk]
        pv = metrics["portfolio_value"]

        def pct_or_inr(pct_key, inr_key):
            if pv and inr_key in hs:
                return f"{hs[pct_key]:.2f}%  ({fmt_inr(hs[inr_key])})"
            return f"{hs[pct_key]:.2f}%"

        row_data = [
            rank,
            st["name"],
            f"{hs['win_rate_pct']:.0f}%",
            pct_or_inr("avg_protection_pct", "avg_protection_inr"),
            f"{hs['worst_protection_pct']:.2f}%",
            f"{hs['best_protection_pct']:.2f}%",
            pct_or_inr("monthly_drag_pct", "monthly_drag_inr"),
            f"{hs['annual_drag_pct']:.2f}%",
            hs["composite_score"],
            st["trigger"],
            st["requires"],
        ]
        bg = "E8F4FD" if rank <= 3 else ("FFF8E1" if rank <= 5 else "FFFFFF")
        for j, val in enumerate(row_data, 1):
            c = ws2.cell(row=r, column=j, value=val)
            c.font = Font(name="Calibri", size=10, bold=(rank <= 3))
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="center" if j != 2 else "left",
                                    vertical="center", wrap_text=(j >= 10))
            _border(c)
            ws2.row_dimensions[r].height = 24

    pr = len(STRATEGIES) + 5
    ws2.merge_cells(f"A{pr}:K{pr}")
    ws2[f"A{pr}"].value = "RECOMMENDED PLAYBOOK FOR THIS PORTFOLIO"
    ws2[f"A{pr}"].font  = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    ws2[f"A{pr}"].fill  = PatternFill("solid", fgColor="1E3A5F")
    ws2[f"A{pr}"].alignment = Alignment(vertical="center")
    ws2.row_dimensions[pr].height = 20

    pb = metrics["portfolio_beta"]
    playbook = [
        ("STRATEGY A — Always-On",
         f"Keep 20% Liquid Debt + 10% Gold. Your β={pb:.2f} means each 10% Nifty fall = ~{10*pb:.1f}% portfolio loss."),
        ("STRATEGY B — Tactical",
         f"Trigger: 2-of-3 signals (VIX>20, Nifty<100-DMA by -2%, FII outflow). Buy ATM Nifty Puts on {pb*100:.0f}% notional."),
        ("STRATEGY C — Crisis",
         f"VIX > 28. Short Nifty Futures covering {pb*100:.0f}% of portfolio notional. Needs F&O account + 25% margin."),
    ]
    for k, (title, desc) in enumerate(playbook):
        rn = pr + 1 + k
        ws2.merge_cells(f"C{rn}:K{rn}")
        c1 = ws2.cell(row=rn, column=1, value=title)
        ws2.merge_cells(f"A{rn}:B{rn}")
        c1 = ws2.cell(row=rn, column=1, value=title)
        c1.font = Font(name="Calibri", bold=True, size=10, color="1E3A5F")
        c1.alignment = Alignment(vertical="center")
        c2 = ws2.cell(row=rn, column=3, value=desc)
        c2.font = Font(name="Calibri", size=10)
        c2.alignment = Alignment(vertical="center", wrap_text=True)
        ws2.row_dimensions[rn].height = 30
        for j in range(1, 12):
            _border(ws2.cell(row=rn, column=j))

    # ── Sheet 3: Scenario Analysis ────────────────────────────
    ws3 = wb.create_sheet("3_Scenario Analysis")
    ws3.freeze_panes = "C3"

    ws3.merge_cells("A1:E1")
    t3 = ws3["A1"]
    t3.value = (f"SCENARIO ANALYSIS — How Each Hedge Would Have Protected Your Portfolio  "
                f"(β={metrics['portfolio_beta']:.3f})")
    t3.font = Font(name="Calibri", bold=True, size=12, color="1E3A5F")
    t3.alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[1].height = 24

    base_hdrs = ["Episode", "Nifty DD%", "Your Portfolio Loss%"]
    if metrics["portfolio_value"]:
        base_hdrs += ["Portfolio Loss (₹)"]
    for sk in STRAT_KEYS:
        sn = STRATEGIES[sk]["short_name"]
        base_hdrs += [f"{sn} Hedge Gain%", f"{sn} Net Pos%", f"{sn} Offset%"]

    for j, h in enumerate(base_hdrs, 1):
        c = ws3.cell(row=2, column=j, value=h)
        _style_header(c, wrap=True)
        ws3.column_dimensions[get_column_letter(j)].width = 15 if j <= 4 else 13
    ws3.row_dimensions[2].height = 32

    for i, ep_row in scenario_df.iterrows():
        r = i + 3
        port_loss = ep_row["Portfolio_Loss_Pct"]
        row_vals = [ep_row["Episode"], ep_row["Nifty_DD_Pct"], port_loss]
        if metrics["portfolio_value"]:
            row_vals.append(ep_row.get("Portfolio_Loss_INR", ""))
        for sk in STRAT_KEYS:
            sn = STRATEGIES[sk]["short_name"]
            row_vals += [
                ep_row.get(f"{sn}_HedgePnL_Pct", ""),
                ep_row.get(f"{sn}_NetPos_Pct", ""),
                ep_row.get(f"{sn}_Offset_Pct", ""),
            ]

        for j, val in enumerate(row_vals, 1):
            c = ws3.cell(row=r, column=j, value=val)
            c.font = Font(name="Calibri", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center")
            _border(c)
            if isinstance(val, (int, float)):
                if j == 3 or (metrics["portfolio_value"] and j == 4):
                    c.font = Font(name="Calibri", size=10, color="C62828", bold=True)
                elif j > 4 and (j - (5 if metrics["portfolio_value"] else 4)) % 3 == 1:
                    if val > 0:
                        c.font = Font(name="Calibri", size=10, color="2E7D32")
                    elif val < 0:
                        c.font = Font(name="Calibri", size=10, color="C62828")

    # ── Sheet 4: Strategy Detail ──────────────────────────────
    ws4 = wb.create_sheet("4_Strategy Detail")

    ws4.merge_cells("A1:F1")
    t4 = ws4["A1"]
    t4.value = "STRATEGY DEEP-DIVE — How Each Hedge Works"
    t4.font = Font(name="Calibri", bold=True, size=12, color="1E3A5F")
    t4.alignment = Alignment(horizontal="left", vertical="center")
    ws4.row_dimensions[1].height = 24

    col_w4 = [5, 30, 60, 25, 25, 25]
    h4 = ["#", "Strategy", "How it Works / When to Use", "Requires", "Trigger", "Monthly Cost"]
    for j, (h, w) in enumerate(zip(h4, col_w4), 1):
        c = ws4.cell(row=2, column=j, value=h)
        _style_header(c)
        ws4.column_dimensions[get_column_letter(j)].width = w

    for rank, sk in enumerate(sorted_keys, 1):
        r = rank + 2
        st = STRATEGIES[sk]
        row_data = [rank, st["name"], st["note"], st["requires"], st["trigger"],
                    f"{st['monthly_drag']:.2f}% / month"]
        bg = "E8F4FD" if rank <= 3 else "FFFFFF"
        for j, val in enumerate(row_data, 1):
            c = ws4.cell(row=r, column=j, value=val)
            c.font = Font(name="Calibri", size=10, bold=(rank <= 3 and j == 2))
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="left" if j in (2,3,4,5) else "center",
                                    vertical="center", wrap_text=True)
            _border(c)
        ws4.row_dimensions[r].height = 40

    # ── Sheet 5: Instructions ─────────────────────────────────
    #ws5 = wb.create_sheet("5_How to Use")
    #ws5.column_dimensions["A"].width = 100

    #lines = [
    #    ("PORTFOLIO HEDGE ENGINE — USER GUIDE", 14, True, "1E3A5F"),
    #    ("", 10, False, "000000"),
    #    ("HOW TO UPDATE THIS FILE WITH YOUR ACTUAL PORTFOLIO", 11, True, "1E3A5F"),
    #    ("1. Go to sheet '1_Portfolio Holdings'", 10, False, "000000"),
    #    ("2. Replace the sample holdings with your own investments.", 10, False, "000000"),
    #    ("   - Asset_Type must match one of the types in the Beta Defaults table below.", 10, False, "000000"),
    #    ("   - Beta_Override: leave blank to use the default, or enter a custom β value.", 10, False, "000000"),
    #    ("   - Allocation_Pct: your % weight in that holding.", 10, False, "000000"),
    #    ("3. Re-run portfolio_hedge_engine.py with --file your_portfolio.csv to refresh all outputs.", 10, False, "000000"),
    #    ("", 10, False, "000000"),
    #    ("WHAT IS PORTFOLIO BETA?", 11, True, "1E3A5F"),
    #    ("  β = 1.0 → Your portfolio moves exactly with the Nifty.", 10, False, "000000"),
    #    ("  β = 1.2 → A 10% Nifty fall causes ~12% portfolio loss.", 10, False, "000000"),
    #    ("", 10, False, "000000"),
    #    ("BETA DEFAULTS TABLE", 11, True, "1E3A5F"),
    #]
    #for row_n, (text, size, bold, color) in enumerate(lines, 1):
    #    c = ws5.cell(row=row_n, column=1, value=text)
    #    c.font = Font(name="Calibri", size=size, bold=bold, color=color)
    #    c.alignment = Alignment(vertical="center", wrap_text=True)
    #    ws5.row_dimensions[row_n].height = 18

    #start_r = len(lines) + 2
    #ws5.cell(row=start_r, column=1, value="Asset Type").font = Font(bold=True)
    #ws5.cell(row=start_r, column=2, value="Default Beta").font = Font(bold=True)
    #ws5.column_dimensions["B"].width = 15
    #for k, (atype, beta) in enumerate(BETA_DEFAULTS.items(), 1):
    #    c1 = ws5.cell(row=start_r + k, column=1, value=atype)
    #    c2 = ws5.cell(row=start_r + k, column=2, value=beta)
    #    c1.font = Font(name="Calibri", size=10)
    #    c2.font = Font(name="Calibri", size=10,
    #                   color="2E7D32" if beta < 0 else ("1565C0" if beta < 0.1 else "000000"))

    wb.save(output_path)
    print(f"  Excel saved: {output_path}")

# ══════════════════════════════════════════════════════════════
# TEXT REPORT
# ══════════════════════════════════════════════════════════════
def generate_text_report(metrics, hedge_sizing, scenario_df, output_path, live_nifty=None):
    pv = metrics["portfolio_value"]
    pb = metrics["portfolio_beta"]

    lines = [
        "═" * 68,
        "  PORTFOLIO HEDGE ENGINE — ANALYSIS REPORT",
        f"  Generated: {date.today().strftime('%d %B %Y')}",
        "═" * 68,
        "",
        "PORTFOLIO SUMMARY",
        f"  Portfolio Beta (β):       {pb:.3f}",
        f"  Beta Range (low–high):    {metrics.get('beta_low', pb):.3f} – {metrics.get('beta_high', pb):.3f}",
        f"  ⚠ Beta Creep: Mid/small cap betas rise in drawdowns — use high-end of range for stress sizing.",
        f"  Equity Exposure:          {metrics['equity_weight']:.1f}%",
        f"  Debt / Liquid:            {metrics['debt_weight']:.1f}%",
        f"  Gold / Inverse Hedge:     {metrics['hedge_weight']:.1f}%",
        f"  Top-3 Concentration:      {metrics['top3_concentration']:.1f}%",
        f"  Holdings:                 {metrics['num_holdings']}",
        f"  Portfolio Value:          {fmt_inr(pv) if pv else 'Not provided (% outputs only)'}",
        "",
        "WHAT YOUR BETA MEANS",
        f"  A 10% Nifty fall  →  ~{10*pb:.1f}% loss  (range {10*metrics.get('beta_low',pb):.1f}% – {10*metrics.get('beta_high',pb):.1f}%)",
        f"  A 20% Nifty fall  →  ~{20*pb:.1f}% loss  (range {20*metrics.get('beta_low',pb):.1f}% – {20*metrics.get('beta_high',pb):.1f}%)",
        f"  A 40% Nifty fall  →  ~{40*pb:.1f}% loss  (range {40*metrics.get('beta_low',pb):.1f}% – {40*metrics.get('beta_high',pb):.1f}%)",
        f"  A 59% Nifty fall  →  ~{min(59*pb,100):.1f}% loss (2008 worst case)",
        "",
        "SECTOR BREAKDOWN",
    ]
    for sector, pct in metrics["sector_breakdown"].items():
        lines.append(f"  {sector:<25} {pct:>6.1f}%")

    lines += ["", "═" * 68, "HEDGE RECOMMENDATIONS (ranked by composite score)", "═" * 68, ""]

    sorted_keys = sorted(STRATEGIES.keys(), key=lambda k: STRATEGIES[k]["score"], reverse=True)
    for rank, sk in enumerate(sorted_keys, 1):
        hs = hedge_sizing[sk]
        st = STRATEGIES[sk]
        lines.append(f"  #{rank}  {st['name']}")
        lines.append(f"       Win Rate:       {hs['win_rate_pct']:.0f}%")
        lines.append(f"       Avg Protection: {hs['avg_protection_pct']:.2f}%"
                     + (f"  ({fmt_inr(hs.get('avg_protection_inr'))})" if pv else ""))
        lines.append(f"       Worst / Best:   {hs['worst_protection_pct']:.2f}% / {hs['best_protection_pct']:.2f}%")
        lines.append(f"       Annual Drag:    {hs['annual_drag_pct']:.2f}%"
                     + (f"  ({fmt_inr(hs.get('annual_drag_inr'))})" if pv else ""))
        lines.append(f"       Trigger:        {st['trigger']}")
        lines.append(f"       Requires:       {st['requires']}")
        lines.append(f"       Note:           {textwrap.fill(st['note'], 55, subsequent_indent=' '*23)}")
        lines.append("")

    lines += ["═" * 68, "SCENARIO ANALYSIS — Key Historical Episodes", "═" * 68, ""]
    for _, ep_row in scenario_df.iterrows():
        lines.append(f"  {ep_row['Episode']}")
        lines.append(f"    Nifty Drawdown:      {ep_row['Nifty_DD_Pct']:.1f}%")
        lines.append(f"    Your Portfolio Loss: {ep_row['Portfolio_Loss_Pct']:.1f}%"
                     + (f"  ({fmt_inr(ep_row.get('Portfolio_Loss_INR'))})" if pv else ""))
        lines.append("    Best hedge offset:")
        offsets = {}
        for sk in STRAT_KEYS:
            sn  = STRATEGIES[sk]["short_name"]
            val = ep_row.get(f"{sn}_Offset_Pct", None)
            if val is not None:
                offsets[sn] = float(val)
        if offsets:
            best_sn = max(offsets, key=offsets.get)
            lines.append(f"      → {best_sn}: {offsets[best_sn]:.0f}% of loss offset")
        else:
            lines.append("      → N/A (no data for this episode)")
        lines.append("")

    # --- INJECTING LIVE 100-DMA LOGIC HERE ---
    lines += ["═" * 68, "LIVE MARKET STATUS (Dynamic 100-DMA)", "═" * 68, ""]
    
    if live_nifty:
        lines.append(f"  Nifty 50 Current:   {live_nifty['close']:,.2f}")
        lines.append(f"  Nifty 100-DMA:      {live_nifty['dma100']:,.2f}")
        lines.append(f"  Gap to 100-DMA:     {live_nifty['gap_pct']:+.2f}%")
        
        if live_nifty['trigger_active']:
            lines.append("  ► TRIGGER STATUS:   [ACTIVE] Nifty is > 2% below 100-DMA.")
            lines.append("                      Deploy Strategy B hedges immediately.")
        else:
            lines.append("  ► TRIGGER STATUS:   [INACTIVE] Nifty is holding above threshold.")
            lines.append("                      Do not overpay for ATM puts today.")
    else:
        lines.append("  [Live data unavailable. Run script with internet connection.]")
    
    lines.append("")
    # ----------------------------------------

    lines += [
        "═" * 68,
        "RECOMMENDED PLAYBOOK",
        "═" * 68,
        "",
        "  STRATEGY A (Always-On):",
        "    Keep 20% Liquid Debt + 10% Gold at all times.",
        f"    This immediately reduces your effective β from {pb:.2f} to ~{pb*0.70:.2f}.",
        "",
        "  STRATEGY B (Tactical):",
        "    Trigger: 2-of-3 signals (VIX>20, Nifty<100-DMA by -2%, FII outflow)",
        f"    Buy ATM Nifty Puts on {pb*100:.0f}% portfolio notional.",
        "    Hold for 1 month, roll if signals still active.",
        "",
        "  STRATEGY C (Crisis):",
        f"    Short Nifty Futures covering {pb*100:.0f}% notional.",
        "    Combine with Strategy B collars to reduce cost.",
        "",
        "═" * 68,
        "CAVEATS",
        "═" * 68,
        "  • All P&L figures are pre-tax (STT, CGT, GST excluded).",
        "  • Beta estimates are approximate; actual correlations vary by market regime.",
        "  • Annual drag figures assume the hedge is held 100% of the time.",
        "    In practice, tactical hedges (B/C) are deployed only when signals fire.",
        "  • Past back-test performance does not guarantee future results.",
        "  • Short Nifty Futures require a 25% cash margin buffer for MTM margin calls.",
        "  • Pre-2008 VIX is a 20-day realised volatility proxy, not the actual NSE VIX.",
        "═" * 68,
    ]

    if hasattr(output_path, 'write'):
        output_path.write("\n".join(lines))
    else:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    print(f"  Report saved: {output_path}")
# ==============================================================
# word document builder
# ==============================================================
def generate_docx_report(metrics, hedge_sizing, scenario_df, output_path, live_nifty=None):
    from docx import Document
    from datetime import date

    doc = Document()
    
    # Helper for formatting INR locally
    def local_fmt_inr(val):
        if not val: return "—"
        if abs(val) >= 1e7: return f"₹{val/1e7:.2f}Cr"
        if abs(val) >= 1e5: return f"₹{val/1e5:.2f}L"
        return f"₹{val:,.0f}"

    pv = metrics.get("portfolio_value")
    pb = metrics.get("portfolio_beta", 1.0)

    # 1. Header
    doc.add_heading('INSTITUTIONAL RISK ENGINE — ANALYSIS REPORT', 0)
    doc.add_paragraph(f"Generated: {date.today().strftime('%d %B %Y')}")
    
    # 2. Summary
    doc.add_heading('PORTFOLIO SUMMARY', level=1)
    p1 = doc.add_paragraph()
    p1.add_run(f"Portfolio Beta (β): {pb:.3f}\n").bold = True
    bl = metrics.get('beta_low', pb)
    bh = metrics.get('beta_high', pb)
    p1.add_run(f"Beta Range: {bl:.3f} – {bh:.3f}\n").bold = True
    p1.add_run("⚠ Beta Creep: Mid/small cap betas rise in drawdowns — use the high-end of range for stress sizing.\n")
    p1.add_run(f"Equity Exposure: {metrics.get('equity_weight', 0):.1f}%\n")
    p1.add_run(f"Debt / Liquid: {metrics.get('debt_weight', 0):.1f}%\n")
    p1.add_run(f"Gold / Inverse: {metrics.get('hedge_weight', 0):.1f}%\n")
    p1.add_run(f"Holdings: {metrics.get('num_holdings', 0)}\n")
    if pv:
        p1.add_run(f"Portfolio Value: {local_fmt_inr(pv)}\n")
        
    # 3. Live Data
    doc.add_heading('LIVE MARKET STATUS (Dynamic 100-DMA)', level=1)
    p_live = doc.add_paragraph()
    if live_nifty:
        p_live.add_run(f"Nifty 50 Current: {live_nifty['close']:,.2f}\n")
        p_live.add_run(f"Nifty 100-DMA: {live_nifty['dma100']:,.2f}\n")
        p_live.add_run(f"Gap to 100-DMA: {live_nifty['gap_pct']:+.2f}%\n")
        if live_nifty['trigger_active']:
            p_live.add_run("► TRIGGER STATUS: [ACTIVE] Nifty is > 2% below 100-DMA. Deploy Strategy B.\n").bold = True
        else:
            p_live.add_run("► TRIGGER STATUS: [INACTIVE] Nifty is holding above threshold.\n")
    else:
        p_live.add_run("[Live data unavailable]")

    # 4. Hedge Recommendations
    doc.add_heading('HEDGE RECOMMENDATIONS', level=1)
    sorted_keys = sorted(STRATEGIES.keys(), key=lambda k: STRATEGIES[k]["score"], reverse=True)
    for rank, sk in enumerate(sorted_keys, 1):
        hs = hedge_sizing[sk]
        st = STRATEGIES[sk]
        p2 = doc.add_paragraph()
        p2.add_run(f"#{rank} {st['name']}\n").bold = True
        p2.add_run(f"Win Rate: {hs['win_rate_pct']:.0f}%\n")
        
        pnl_str = f"{hs['avg_protection_pct']:.2f}%"
        if pv and 'avg_protection_inr' in hs: 
            pnl_str += f" ({local_fmt_inr(hs['avg_protection_inr'])})"
        p2.add_run(f"Avg Protection: {pnl_str}\n")
        
        drag_str = f"{hs['annual_drag_pct']:.2f}%"
        if pv and 'annual_drag_inr' in hs: 
            drag_str += f" ({local_fmt_inr(hs['annual_drag_inr'])})"
        p2.add_run(f"Annual Drag: {drag_str}\n")
        p2.add_run(f"Trigger: {st['trigger']}\n")
        p2.add_run(f"Note: {st['note']}")

    doc.save(output_path)


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description="Portfolio Hedge Engine v2.0")
    parser.add_argument("--file",  default="portfolio_input.csv",
                        help="Path to your portfolio CSV (default: portfolio_input.csv)")
    parser.add_argument("--value", type=float, default=None,
                        help="Total portfolio value in ₹ (optional). E.g. --value 5000000 for ₹50L")
    parser.add_argument("--output-dir", default="./portfolio_output",
                        help="Output directory (default: ./portfolio_output)")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    portfolio_file = args.file

    print("\n" + "═"*60)
    print("  PORTFOLIO HEDGE ENGINE  v2.0 (Live Data Edition)")
    print("═"*60)

    if not os.path.exists(portfolio_file):
        print(f"\n  Portfolio file '{portfolio_file}' not found.")
        print("  Generating sample portfolio template...")
        generate_template_csv(portfolio_file)

    print(f"\n  Loading portfolio: {portfolio_file}")
    df_port = load_portfolio(portfolio_file)
    print(f"  Holdings: {len(df_port)} | Total allocation: {df_port['Allocation_Pct'].sum():.1f}%")

    metrics       = compute_portfolio_metrics(df_port, args.value)
    hedge_sizing  = compute_hedge_sizing(metrics["portfolio_beta"], args.value)
    scenario_df   = compute_scenario_analysis(metrics["portfolio_beta"], args.value)
    
    # --- NEW LIVE DATA FETCH ---
    live_nifty = fetch_live_nifty_status()
    # ---------------------------

    print(f"\n  Portfolio Beta: {metrics['portfolio_beta']:.3f}")
    if args.value:
        print(f"  Portfolio Value: {fmt_inr(args.value)}")

    xlsx_path = os.path.join(args.output_dir, "portfolio_hedge_analysis.xlsx")
    txt_path  = os.path.join(args.output_dir, "portfolio_hedge_report.txt")

    print("\n  Building Excel workbook (5 sheets)...")
    build_excel(df_port, metrics, hedge_sizing, scenario_df, xlsx_path)

    print("  Building text report...")
    generate_text_report(metrics, hedge_sizing, scenario_df, txt_path, live_nifty)

    print("═"*60 + "\n")

if __name__ == "__main__":
    main()
