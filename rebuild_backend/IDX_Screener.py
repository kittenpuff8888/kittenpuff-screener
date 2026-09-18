import os
import sys
import time
import traceback
import json
from datetime import datetime, timezone
from typing import Optional, Dict, List, Tuple, Any

import numpy as np
import pandas as pd
import yfinance as yf
# =========================
# USER CONFIG
# =========================
# MARKET_DATE is the hard "as of" cutoff for ALL price and fundamental data.
# Setting this to a past date (e.g. 2026-05-29 run in June) means:
#   • All yf.download / tk.history calls use end=(MARKET_DATE + 1 day)
#     because yfinance's end= is EXCLUSIVE (end="2026-05-29" gives data through
#     May 28 only; end="2026-05-30" gives data through May 29 inclusive).
#   • clip_to_market_date() hard-clips any downloaded series after fetch.
#   • News and corp-action "age" calculations use MARKET_DATE as reference.
#
# EXCEPTIONS — fetched at actual run time (real-time, NOT clipped to MARKET_DATE):
#   • VIX          — live market context for the IDX Overview Fear & Greed panel.
#   • run_dt       — wall-clock timestamp of when the script was executed.
MARKET_DATE = os.environ.get("MARKET_DATE", datetime.now().strftime("%Y-%m-%d"))

# =========================
# BACKTEST MODE CONFIG
# =========================
# When True: restricts the ticker universe to BACKTEST_TICKERS only.
# MARKET_DATE still controls the data cutoff regardless of this flag.
BACKTEST_MODE = False

BACKTEST_TICKERS = [
    "AADI","ADMR","ADRO","AMMN","AMRT","ANTM","ARCI","ASII","BBCA","BBNI",
    "BBRI","BIPI","BMRI","BNBR","BREN","BRMS","BRPT","BUMI","BUVA","CDIA",
    "CUAN","DEWA","DSSA","EMAS","ENRG","ESSA","EXCL","GOTO","HRTA","IMPC",
    "INCO","INDF","INDY","INKP","ITMG","MBMA","MDKA","MEDC","NCKL","PGAS",
    "PSAB","PTBA","PTRO","RAJA","RATU","TAPG","TCPI","TINS","TLKM","TPIA",
    "UNTR","VKTR","MINA",
]

# =========================
# COMPATIBILITY CONSTANTS (v1.6.4 safe layer)
# =========================
# These placeholders are defined early so top-level sheet schema declarations never fail.
FILL_GROUP_FLOW = None
FILL_GROUP_MS = None
FILL_HEADER = None
FONT_NOTE = None
FONT_HEADER = None


from openpyxl import Workbook
from openpyxl.styles.fills import Fill
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Post-import safe placeholders for style compatibility
FILL_GROUP_FLOW = PatternFill("solid", fgColor="2A4A66")  # post-import themed init
FILL_GROUP_MS = PatternFill("solid", fgColor="5B3F8C")  # post-import themed init for market structure
FILL_HEADER = PatternFill("solid", fgColor="2A4A66")  # post-import themed init

from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
from rebuild_backend.logic_reference import workbook_rows
from rebuild_backend.sector_normalization import normalize_idx_sector

# =========================
# CONFIG
# =========================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR = os.path.join(BASE_DIR, "Raw")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
CACHE_DIR = os.path.join(BASE_DIR, "Cache")
SHARES_CACHE_FILE = os.path.join(CACHE_DIR, "shares_outstanding_cache.csv")

MIN_RVOL = 1.5
VOL_THRESHOLD = 10_000_000
MIN_BARS_FULL = 210
MIN_BARS_PARTIAL = 30
FETCH_RETRIES = 3
RETRY_SLEEP_SEC = 1.2

# =========================
# TRUE IDX SECTOR MOVERS CONFIG
# =========================
SECTOR_MOVERS_PERIOD_WEEKS = 52   # visible "Period" for RRG window
SECTOR_MOVERS_SMOOTHING = 10      # EMA smoothing length
SECTOR_MOVERS_TRAIL = 5           # tail points shown
SECTOR_MOVERS_LOOKBACK = 52       # rolling normalization window (RRG baseline)
SECTOR_INDEX_SYMBOLS = {
    "COMPOSITE": "^JKSE",
}

# =========================
# STYLES
# =========================
LINE = "C2CADE"
THIN = Side(style="thin", color=LINE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LEFT_BORDER_ONLY = Border(left=THIN)

FONT_TITLE = Font(name="Calibri", size=14, bold=True)
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True)

# Soft pastel header palette (dark text only; no white font)
FONT_GROUP = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_SUBHEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_NOTE = Font(name="Calibri", size=10, color="1F1F1F")
FONT_BODY = Font(name="Calibri", size=10)
FONT_REQ = Font(name="Calibri", size=10, bold=True)

# Global / generic headers (used across summary + helper sheets)
FILL_HEADER_SUMMARY = PatternFill("solid", fgColor="2A4A66")  # navy
FILL_HEADER = PatternFill("solid", fgColor="2A4A66")
FILL_GROUP_FLOW = PatternFill("solid", fgColor="2A4A66")      # navy

# IDX Screener Detail group headers (all distinct colors)
FILL_GROUP_STOCK = PatternFill("solid", fgColor="2A4A66")     # navy
FILL_GROUP_OWNER = PatternFill("solid", fgColor="237A5C")     # forest green
FILL_GROUP_Q = PatternFill("solid", fgColor="7A1A1A")         # deep blue
FILL_GROUP_PQ = PatternFill("solid", fgColor="7A1A1A")        # dark crimson
FILL_GROUP_PY = PatternFill("solid", fgColor="7A1A1A")        # cobalt
FILL_GROUP_MP = PatternFill("solid", fgColor="21467C")        # denim
FILL_GROUP_VOL = PatternFill("solid", fgColor="216B6B")       # teal
FILL_GROUP_MA = PatternFill("solid", fgColor="32588E")        # dark plum
FILL_GROUP_MOM = PatternFill("solid", fgColor="207878")       # sea green
FILL_GROUP_VR = PatternFill("solid", fgColor="8C2038")        # burgundy
FILL_GROUP_MACD_MOM = PatternFill("solid", fgColor="4A1060")  # deep violet-purple for MACD Momentum

FILL_LEGEND_A = PatternFill("solid", fgColor="007BA7")
FILL_LEGEND_B = PatternFill("solid", fgColor="4B49AC")
FILL_LEGEND_D = PatternFill("solid", fgColor="7978E9")
FILL_LEGEND_PY = PatternFill("solid", fgColor="0077B6")
FILL_STATUS_OK = PatternFill("solid", fgColor="007BA7")
FILL_STATUS_PARTIAL = PatternFill("solid", fgColor="4B49AC")
FILL_STATUS_NODATA = PatternFill("solid", fgColor="F3797E")

# =========================
# HELPERS
# =========================
def style_cell(cell, fill=None, font=None, border=True, align="center", wrap=False):
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if border:
        cell.border = BORDER
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)

def style_plain(cell, font=None, align="left", wrap=False):
    if font:
        cell.font = font
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)

def style_legend(cell, fill=None, font=None, align="left", wrap=False):
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)


# =========================
# EXCEL DATE FORMAT STANDARD
# =========================
DATE_FORMAT_EXCEL = "dd mmm yyyy"

def _apply_excel_date_format_workbook(wb):
    """
    Enforce user-requested date display across workbook.
    Keeps real Excel date values but displays as: dd mmm yyyy.
    """
    try:
        from datetime import datetime, date
        import pandas as pd
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, (datetime, date, pd.Timestamp)):
                        cell.number_format = DATE_FORMAT_EXCEL
    except Exception:
        pass


def safe_num(v, default=np.nan):
    try:
        if pd.isna(v):
            return default
        return float(v)
    except Exception:
        return default

def safe_int(v, default=0):
    try:
        if pd.isna(v):
            return default
        return int(round(float(v)))
    except Exception:
        return default

def pct_diff(close, ma):
    if pd.isna(close) or pd.isna(ma) or ma == 0:
        return np.nan
    return ((close / ma) - 1.0) * 100.0

def pos_label(close, ma):
    if pd.isna(close) or pd.isna(ma):
        return ""
    return "Above" if close >= ma else "Below"

def rsi(series: pd.Series, period=14):
    # Wilder's RMA — matches TradingView built-in RSI (alpha=1/period, adjust=False)
    delta = series.diff()
    up  = delta.clip(lower=0).ewm(alpha=1.0 / period, adjust=False).mean()
    dn  = (-delta.clip(upper=0)).ewm(alpha=1.0 / period, adjust=False).mean()
    rs  = up / dn.replace(0, np.nan)
    return 100 - (100 / (1 + rs))

def ema(series: pd.Series, period: int):
    return series.ewm(span=period, adjust=False).mean()

def quarter_start(ts: pd.Timestamp):
    q = ((ts.month - 1) // 3) + 1
    start_month = (q - 1) * 3 + 1
    return pd.Timestamp(ts.year, start_month, 1)

def prev_quarter_range(ts: pd.Timestamp):
    cq = quarter_start(ts)
    prev_end = cq - pd.Timedelta(days=1)
    prev_start = quarter_start(prev_end)
    return prev_start, prev_end

def weighted_std(values: pd.Series, weights: pd.Series):
    mask = (~values.isna()) & (~weights.isna()) & (weights > 0)
    x = values[mask].astype(float)
    w = weights[mask].astype(float)
    if len(x) == 0 or w.sum() == 0:
        return np.nan
    mean = np.average(x, weights=w)
    var = np.average((x - mean) ** 2, weights=w)
    return np.sqrt(var)

def anchored_vwap_block(
    df: pd.DataFrame,
    selected_close: float = np.nan,
    previous_close: float = np.nan,
) -> Dict[str, float]:
    empty_out = {
        "days": 0,
        "vwap": np.nan,
        "p1": np.nan,
        "p2": np.nan,
        "p3": np.nan,
        "m1": np.nan,
        "m2": np.nan,
        "m3": np.nan,
        "sd": np.nan,
        "sd_score": np.nan,
        "prev_sd": np.nan,
        "sd_delta": np.nan,
    }

    if df is None or df.empty:
        return empty_out

    price = (df["High"] + df["Low"] + df["Close"]) / 3.0
    vol = df["Volume"].replace(0, np.nan)
    mask = (~price.isna()) & (~vol.isna())
    d = df.loc[mask].copy()

    if d.empty:
        return empty_out

    def _calc_block(block: pd.DataFrame, score_close=np.nan):
        if block is None or block.empty:
            return np.nan, np.nan, np.nan, np.nan, np.nan, np.nan

        tp = (block["High"] + block["Low"] + block["Close"]) / 3.0
        vv = block["Volume"].astype(float)

        if len(tp) == 0 or vv.sum() == 0:
            return np.nan, np.nan, np.nan, np.nan, np.nan, np.nan

        # TradingView anchored VWAP: running cumulative (matches Pine Script ta.vwap)
        cum_vol    = vv.cumsum()
        cum_tp_vol = (tp * vv).cumsum()
        running_vwap = cum_tp_vol / cum_vol

        # Variance uses RUNNING vwap at each bar — same as TradingView
        dev2         = (tp - running_vwap) ** 2
        running_var  = (dev2 * vv).cumsum() / cum_vol
        running_sd   = running_var.apply(lambda x: np.sqrt(max(x, 0.0)))

        vwap = float(running_vwap.iloc[-1])
        sd   = float(running_sd.iloc[-1])

        # Near-zero SD guard
        sd_floor = max(abs(vwap) * 1e-6 if pd.notna(vwap) else 0.0, 1e-9)
        sd_valid = pd.notna(sd) and sd > sd_floor

        close_last = safe_num(score_close, np.nan)
        if pd.isna(close_last):
            close_last = safe_num(block["Close"].iloc[-1])
        sd_score   = safe_num((close_last - vwap) / sd) if sd_valid else 0.0

        m1 = vwap - sd       if sd_valid else vwap
        m2 = vwap - (2 * sd) if sd_valid else vwap
        m3 = vwap - (3 * sd) if sd_valid else vwap

        return vwap, m1, m2, m3, sd_score, sd

    # Current anchored block (includes latest bar)
    vwap, m1, m2, m3, sd_score, sd = _calc_block(d, selected_close)

    # Prior-day anchored block (exclude latest bar)
    if len(d) >= 2:
        d_prev = d.iloc[:-1].copy()
        if pd.notna(previous_close):
            prev_sd_score = safe_num((float(previous_close) - vwap) / sd) if pd.notna(sd) and sd > 0 else np.nan
        else:
            prev_vwap, prev_m1, prev_m2, prev_m3, prev_sd_score, prev_sd_raw = _calc_block(d_prev)
        sd_delta = safe_num(sd_score - prev_sd_score) if pd.notna(sd_score) and pd.notna(prev_sd_score) else np.nan
    else:
        prev_sd_score = np.nan
        sd_delta = np.nan

    # Compute +SD bands (symmetric to -SD bands)
    sd_val = sd if (pd.notna(sd) and float(sd) > 0) else np.nan
    p1 = float(vwap) + float(sd_val) if pd.notna(vwap) and pd.notna(sd_val) else np.nan
    p2 = float(vwap) + 2 * float(sd_val) if pd.notna(vwap) and pd.notna(sd_val) else np.nan
    p3 = float(vwap) + 3 * float(sd_val) if pd.notna(vwap) and pd.notna(sd_val) else np.nan

    return {
        "days": len(d),
        "vwap": vwap,
        "p1": p1,
        "p2": p2,
        "p3": p3,
        "m1": m1,
        "m2": m2,
        "m3": m3,
        "sd": sd,
        "sd_score": sd_score,
        "prev_sd": prev_sd_score,
        "sd_delta": sd_delta,
    }

def zone_label(close, ibl, ibh):
    if pd.isna(close) or pd.isna(ibl) or pd.isna(ibh):
        return "N/A"
    if abs(float(close) - float(ibl)) < 1e-9:
        return "At Level"
    return "Above" if close > ibl else "Below"

def vwap_zone_2pct(close, levels):
    """
    Premium VWAP POI text engine
    - Near threshold = ±12.5% of the relevant segment gap
    - Display = ACTUAL % distance from current price to referenced POI
    - For ranging zones, reference side is the nearest boundary:
      * If closer to upper boundary => "x% below upper"
      * If closer to lower boundary => "x% above lower"
    """
    if pd.isna(close):
        return "N/A"

    valid = [float(x) for x in levels if pd.notna(x)]
    if not valid:
        return "N/A"

    def _pct(px, lvl):
        if pd.isna(lvl) or lvl == 0:
            return None
        return ((float(px) / float(lvl)) - 1.0) * 100.0

    def _near(px, boundary, gap, tol_frac=0.125):
        if pd.isna(boundary) or pd.isna(gap) or gap <= 0:
            return False
        return abs(float(px) - float(boundary)) <= (gap * tol_frac)

    def _fmt_near(px, lvl, label):
        p = _pct(px, lvl)
        if p is None:
            return f"Price Near {label}"
        if abs(p) < 0.05:
            return f"Price Near {label} (At {label})"
        side = "above" if p > 0 else "below"
        return f"Price Near {label} ({abs(p):.1f}% {side} {label})"

    def _fmt_range(px, upper_label, lower_label, upper_level, lower_level):
        du = abs(float(px) - float(upper_level)) if pd.notna(upper_level) else float("inf")
        dl = abs(float(px) - float(lower_level)) if pd.notna(lower_level) else float("inf")

        # Closer to upper boundary => describe as below upper
        if du <= dl:
            p = _pct(px, upper_level)
            if p is None:
                return f"Price Ranging {upper_label} ⇄ {lower_label}"
            return f"Price Ranging {upper_label} ⇄ {lower_label} ({abs(p):.1f}% below {upper_label})"

        # Closer to lower boundary => describe as above lower
        p = _pct(px, lower_level)
        if p is None:
            return f"Price Ranging {upper_label} ⇄ {lower_label}"
        return f"Price Ranging {upper_label} ⇄ {lower_label} ({abs(p):.1f}% above {lower_label})"

    def _fmt_above(px, lvl, label):
        p = _pct(px, lvl)
        if p is None:
            return f"Price Above {label}"
        if abs(p) < 0.05:
            return f"Price Near {label} (At {label})"
        side = "above" if p > 0 else "below"
        return f"Price Above {label} ({abs(p):.1f}% {side} {label})"

    def _fmt_below(px, lvl, label):
        p = _pct(px, lvl)
        if p is None:
            return f"Price Below {label}"
        if abs(p) < 0.05:
            return f"Price Near {label} (At {label})"
        side = "above" if p > 0 else "below"
        return f"Price Below {label} ({abs(p):.1f}% {side} {label})"

    # Preferred mode: [VWAP, -1 SD, -2 SD, -3 SD]
    if len(valid) >= 4:
        vwap, m1, m2, m3 = valid[:4]

        if close >= vwap:
            gap = abs(vwap - m1) if pd.notna(m1) else abs(vwap) * 0.10
            if _near(close, vwap, gap):
                return _fmt_near(close, vwap, "VWAP")
            return _fmt_above(close, vwap, "VWAP")

        if close >= m1:
            gap = abs(vwap - m1)
            if _near(close, vwap, gap):
                return _fmt_near(close, vwap, "VWAP")
            if _near(close, m1, gap):
                return _fmt_near(close, m1, "-1 SD")
            return _fmt_range(close, "VWAP", "-1 SD", vwap, m1)

        if close >= m2:
            gap = abs(m1 - m2)
            if _near(close, m1, gap):
                return _fmt_near(close, m1, "-1 SD")
            if _near(close, m2, gap):
                return _fmt_near(close, m2, "-2 SD")
            return _fmt_range(close, "-1 SD", "-2 SD", m1, m2)

        if close >= m3:
            gap = abs(m2 - m3)
            if _near(close, m2, gap):
                return _fmt_near(close, m2, "-2 SD")
            if _near(close, m3, gap):
                return _fmt_near(close, m3, "-3 SD")
            return _fmt_range(close, "-2 SD", "-3 SD", m2, m3)

        gap = abs(m2 - m3) if pd.notna(m2) else abs(m3) * 0.10
        if _near(close, m3, gap):
            return _fmt_near(close, m3, "-3 SD")
        return _fmt_below(close, m3, "-3 SD")

    # Legacy fallback: [m1, m2, m3]
    if len(valid) == 3:
        m1, m2, m3 = valid

        if close >= m1:
            gap = abs(m1 - m2) if pd.notna(m2) else abs(m1) * 0.10
            if _near(close, m1, gap):
                return _fmt_near(close, m1, "-1 SD")
            return _fmt_above(close, m1, "-1 SD")

        if close >= m2:
            gap = abs(m1 - m2)
            if _near(close, m1, gap):
                return _fmt_near(close, m1, "-1 SD")
            if _near(close, m2, gap):
                return _fmt_near(close, m2, "-2 SD")
            return _fmt_range(close, "-1 SD", "-2 SD", m1, m2)

        if close >= m3:
            gap = abs(m2 - m3)
            if _near(close, m2, gap):
                return _fmt_near(close, m2, "-2 SD")
            if _near(close, m3, gap):
                return _fmt_near(close, m3, "-3 SD")
            return _fmt_range(close, "-2 SD", "-3 SD", m2, m3)

        gap = abs(m2 - m3) if pd.notna(m2) else abs(m3) * 0.10
        if _near(close, m3, gap):
            return _fmt_near(close, m3, "-3 SD")
        return _fmt_below(close, m3, "-3 SD")

    return "N/A"


def vwap_near_zone_label(close, sd_score, vwap, m1, m2, m3, p1=np.nan, p2=np.nan, p3=np.nan):
    """
    TradingView-aligned VWAP band label.

    Important audit fix:
    The previous implementation selected the band from rounded SD Score only.
    That can mis-label a stock as +1 SD when the plotted price is visually
    closest to VWAP / -2 SD because the numeric SD score and available band
    columns can drift after missing-level filtering or low-volume blocks.

    This version:
    1) Preserves the explicit band identity: VWAP, +/-1, +/-2, +/-3.
    2) Chooses the nearest plotted price level by absolute price distance.
    3) Uses SD Score only as a secondary sanity field, not as the source of truth.
    """
    if pd.isna(close):
        return "-"

    close_f = float(close)
    levels = [
        ("VWAP",  vwap, 0),
        ("-1 σ",  m1,  -1),
        ("-2 σ",  m2,  -2),
        ("-3 σ",  m3,  -3),
        ("+1 σ",  p1,  +1),
        ("+2 σ",  p2,  +2),
        ("+3 σ",  p3,  +3),
    ]

    valid = []
    for label, px, sd_int in levels:
        if pd.notna(px) and float(px) > 0:
            valid.append((label, float(px), sd_int))

    if not valid:
        return "-"

    # Nearest visual/plotted level — this is what TradingView users compare to.
    label, level_px, sd_int = min(valid, key=lambda x: abs(close_f - x[1]))
    pct_dist = abs((close_f - level_px) / level_px) * 100.0 if level_px else np.nan

    # Segment-aware near tolerance. Use 12.5% of nearest adjacent SD spacing,
    # capped at 3% of level price to avoid overly wide "near" flags on volatile names.
    adjacent_gaps = [abs(level_px - other_px) for _, other_px, other_sd in valid if abs(other_sd - sd_int) == 1]
    gap = min(adjacent_gaps) if adjacent_gaps else (abs(level_px) * 0.10)
    near_abs = min(gap * 0.125, abs(level_px) * 0.03)

    if abs(close_f - level_px) <= near_abs:
        if pd.isna(pct_dist):
            return f"Near {label}"
        return f"Near {label} ({pct_dist:.2f}%)"

    ordered = sorted(
        [(float(px), label) for label, px, _ in valid],
        key=lambda item: item[0],
        reverse=True,
    )
    if close_f > ordered[0][0]:
        return f"Above {ordered[0][1]}"
    if close_f < ordered[-1][0]:
        return f"Below {ordered[-1][1]}"
    for (upper_px, upper_label), (lower_px, lower_label) in zip(ordered, ordered[1:]):
        if lower_px < close_f < upper_px:
            return f"Between {upper_label} and {lower_label}"
    return label

def vwap_zone_remark(close, m1, m2, m3):
    """
    Legacy compatibility shim.
    Uses new VWAP zone wording logic on legacy 3-level input.
    """
    return vwap_zone_2pct(close, [m1, m2, m3])

def rsi_status(v):
    if pd.isna(v):
        return ""
    if v >= 70:
        return "Overbought"
    if v <= 30:
        return "Oversold"
    if v >= 55:
        return "Strong"
    if v <= 45:
        return "Weak"
    return "Neutral"

# =========================
# DATE ENGINE  — single source of truth
# =========================
def get_market_date() -> pd.Timestamp:
    """
    Returns the effective market date for this screener run.
    MARKET_DATE is always set and always authoritative.
    All data is clipped to this date regardless of BACKTEST_MODE.
    """
    try:
        return pd.Timestamp(MARKET_DATE).normalize()
    except Exception:
        raise ValueError(f"MARKET_DATE='{MARKET_DATE}' is not a valid date string.")


def is_backtest_mode() -> bool:
    """Returns True when BACKTEST_MODE restricts the universe to BACKTEST_TICKERS."""
    return bool(BACKTEST_MODE)


def clip_to_market_date(hist: pd.DataFrame) -> pd.DataFrame:
    """
    Hard clips a price DataFrame to rows with index <= MARKET_DATE.
    MARKET_DATE is always authoritative — no fallback to unclipped data.
    If nothing exists before MARKET_DATE, returns empty (caller must handle).
    """
    if hist is None or hist.empty:
        return hist
    asof = get_market_date()
    try:
        clipped = hist[hist.index.normalize() <= asof]
        return clipped   # may be empty — never return unclipped data
    except Exception:
        return hist


# Legacy aliases — kept so existing call-sites work without change
def _get_forced_asof_timestamp():
    return get_market_date()

def get_effective_asof_date() -> pd.Timestamp:
    return get_market_date()

def _clip_hist_to_asof(hist: pd.DataFrame) -> pd.DataFrame:
    return clip_to_market_date(hist)

def _ordinal(n: int) -> str:
    if 10 <= (n % 100) <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"

def _build_output_filename(run_dt: datetime, latest_market_day: str) -> str:
    ts_run = run_dt.strftime("%y%m%d %H.%M")
    try:
        d = pd.to_datetime(latest_market_day, errors="coerce", format="%Y-%m-%d")
        if pd.notna(d):
            asof_label = f"{d.strftime('%B')} {_ordinal(int(d.day))} {d.year}"
        else:
            asof_label = str(latest_market_day)
    except Exception:
        asof_label = str(latest_market_day)
    return f"IDX Screener as of {asof_label} Running at {ts_run}.xlsx"

def normalize_history(hist):
    if hist is None or hist.empty:
        return None
    # Flatten yfinance MultiIndex columns (new API >=0.2)
    if isinstance(hist.columns, pd.MultiIndex):
        hist.columns = hist.columns.get_level_values(0)
    needed = ["Open", "High", "Low", "Close", "Volume"]
    if any(c not in hist.columns for c in needed):
        return None
    hist = hist[needed].copy()
    hist.dropna(subset=["Close"], inplace=True)
    if hist.empty:
        return None
    idx = pd.to_datetime(hist.index)
    try:
        if getattr(idx, "tz", None) is not None:
            idx = idx.tz_localize(None)
    except Exception:
        try:
            idx = idx.tz_convert(None)
        except Exception:
            pass
    hist.index = idx
    hist = hist[~hist.index.duplicated(keep="last")].sort_index()

    # Clip all data to MARKET_DATE — single authoritative boundary
    try:
        _asof = get_market_date()
        hist = hist[hist.index.normalize() <= _asof]
    except Exception:
        pass

    return hist

def load_shares_cache():
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        if os.path.exists(SHARES_CACHE_FILE):
            df = pd.read_csv(SHARES_CACHE_FILE)
            if {"Ticker", "SharesOutstanding"}.issubset(df.columns):
                return dict(zip(df["Ticker"].astype(str).str.upper(), pd.to_numeric(df["SharesOutstanding"], errors="coerce")))
    except Exception:
        pass
    return {}

def save_shares_cache(cache_dict):
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        rows = [{"Ticker": k, "SharesOutstanding": v} for k, v in cache_dict.items()]
        df = pd.DataFrame(rows)
        if not df.empty:
            df.sort_values("Ticker").to_csv(SHARES_CACHE_FILE, index=False)
    except Exception:
        pass


def _sc_get(r, key, default="N/A"):
    try:
        v = r.get(key, default)
        if v is None:
            return default
        if isinstance(v, float) and np.isnan(v):
            return default
        return v
    except Exception:
        return default

def _format_pct_safe(v, digits=1):
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return "N/A"
        return f"{float(v):.{digits}f}"
    except Exception:
        return "N/A"

def _swing_v16_native_row(r):
    """
    Native POI-first swing row schema (V17 — LONG ONLY)
    Returns ordered dict matching the V17 output column spec.
    """
    def _g(primary, *fallbacks):
        v = r.get(primary)
        if v is not None and str(v) not in ("", "N/A", "nan"):
            return v
        for fb in fallbacks:
            v = r.get(fb)
            if v is not None and str(v) not in ("", "N/A", "nan"):
                return v
        return "N/A"

    # ── Stock info ─────────────────────────────────────────────────────────────
    ticker   = _g("ticker", "Ticker")
    sector   = _g("sector", "IDX Sector")
    price    = _g("close",  "Close", "Price", "Closing Price")
    pct_chg  = _g("pct_chg", "Price Change %", "pct_change")
    emiten   = _g("name", "Emiten", "Company")

    # ── Screener outputs ───────────────────────────────────────────────────────
    priority     = _g("_sc_priority_label", "_sc_swing_grade")
    score        = safe_num(r.get("_sc_swing_score", r.get("_sc_conviction")), np.nan)
    score_disp   = f"{score:.2f}" if pd.notna(score) else "N/A"
    verdict      = _g("verdict_weight_profile", "regime_label", "Verdict Weight Profile")
    trend        = _g("ms_trend_regime", "trend_regime", "Trend Bias")
    rvol         = _g("rvol20", "RVOL")
    adr          = _g("adr_pct", "ADR %", "adr14_pct")
    atr          = _g("atr14_pct", "ATR (14) %")
    last_evt     = _g("ms_last_event", "last_structural_event", "Last Structural Event")
    event_age    = _g("ms_event_age_d", "event_age_d", "Event Age (D)")

    # ── POI engine fields ─────────────────────────────────────────────────────
    poi_type     = _g("_sc_poi_type", "_sc_zone_type", "POI Type")
    primary_poi  = _g("_sc_primary_poi", "_sc_zone_label", "Primary POI")
    next_poi     = _g("_sc_next_poi", "Next POI")
    dist_pct     = safe_num(r.get("_sc_dist_pct"), np.nan)
    dist_disp    = f"{dist_pct:+.2f}%" if pd.notna(dist_pct) else "N/A"
    confluence   = _g("_sc_confluence", "Confluence")
    amt          = _g("_sc_amt_state", "_sc_signal", "AMT State")

    # ── Context ───────────────────────────────────────────────────────────────
    mp           = _g("market_profile_summary", "mp_zone", "Market Profile", "MP Summary")
    ma           = _g("ma_position_summary", "ma_zone", "MA Position", "MA Zone")
    candle       = _g("cp_pattern", "last_candle_pattern", "Candle Pattern", "Pattern")
    candle_date  = _g("cp_date", "pattern_date", "Pattern Date")
    rsi_status   = _g("rsi_status", "RSI Status")
    rsi_div      = _g("div_signal", "divergence_signal", "Divergence Signal")
    macd_wave    = _g("macd_wave", "macd_wave_pattern", "MACD Wave", "Wave Pattern")

    # ── Trade plan ─────────────────────────────────────────────────────────────
    entry        = _g("_sc_entry_disp", "_sc_entry", "Entry Zone")
    trigger      = _g("_sc_trigger", "_sc_signal", "_sc_amt_state", "Trigger")
    t1           = _g("_sc_target_disp", "_sc_t1", "T1")
    t2           = _g("_sc_t2_disp", "_sc_t2", "T2")
    invalid      = _g("_sc_invalidation_disp", "_sc_invalid", "Invalid")
    rr           = safe_num(r.get("_sc_rr"), np.nan)
    rr_disp      = f"{rr:.1f}" if pd.notna(rr) else "N/A"

    # ── Wyckoff extras ─────────────────────────────────────────────────────────
    cause_quality  = _g("_sc_cause_quality", "cause_quality", "Cause Quality")
    markup_ready   = safe_num(r.get("_sc_markup_readiness", r.get("markup_readiness")), np.nan)
    markup_disp    = f"{markup_ready:.0f}%" if pd.notna(markup_ready) else "N/A"
    summary        = _g("_sc_summary", "Summary")

    return {
        "Priority":             priority,
        "Ticker":               ticker,
        "IDX Sector":           sector,
        "Price":                price,
        "Price Change %":             pct_chg,
        "Composite Score":      score_disp,
        "Verdict Profile":      verdict,
        "Trend Bias":           trend,
        "RVOL":                 rvol,
        "ADR %":                adr,
        "ATR14 %":              atr,
        "Last Structural Event":last_evt,
        "Event Age (D)":        event_age,
        "POI Type":             poi_type,
        "Primary POI":          primary_poi,
        "Next POI":             next_poi,
        "Dist %":               dist_disp,
        "Confluence":           confluence,
        "AMT State":            amt,
        "Market Profile":       mp,
        "MA Position":          ma,
        "Candle Pattern":       candle,
        "Pattern Date":         candle_date,
        "RSI Status":           rsi_status,
        "Divergence":           rsi_div,
        "MACD Wave":            macd_wave,
        "Entry Zone":           entry,
        "Trigger":              trigger,
        "T1":                   t1,
        "T2":                   t2,
        "Invalid":              invalid,
        "R/R":                  rr_disp,
        "Cause Quality":        cause_quality,
        "Markup Readiness":     markup_disp,
        "Summary":              summary,
    }


def compact_fmt(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    try:
        n = float(v)
    except Exception:
        return v
    absn = abs(n)
    if absn >= 1_000_000_000_000:
        return f"{n / 1_000_000_000_000:.2f} T"
    if absn >= 1_000_000_000:
        return f"{n / 1_000_000_000:.2f} B"
    if absn >= 1_000_000:
        return f"{n / 1_000_000:.2f} M"
    if absn >= 1_000:
        return f"{n / 1_000:.2f} K"
    return f"{n:,.0f}"

def na_if_nan(v):
    return "N/A" if pd.isna(v) else v


# ─── BUG-03: USD/IDR cached rate ─────────────────────────────────────────────
_USD_IDR_RATE_CACHE: float = 0.0  # 0.0 = not yet fetched

def _fetch_usd_idr_rate() -> float:
    """
    Fetch USD/IDR exchange rate from yfinance USDIDR=X as-of MARKET_DATE.
    Handles yfinance MultiIndex columns (new API).
    Cached for run lifetime.
    """
    global _USD_IDR_RATE_CACHE
    if _USD_IDR_RATE_CACHE > 0:
        return _USD_IDR_RATE_CACHE
    try:
        import yfinance as _yf
        _mdate = get_market_date()
        _end   = (_mdate + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        _start = (_mdate - pd.Timedelta(days=10)).strftime("%Y-%m-%d")
        _h = _yf.download("USDIDR=X", start=_start, end=_end,
                          interval="1d", auto_adjust=True, progress=False,
                          threads=False)
        if _h is not None and not _h.empty:
            # Fix: yfinance >=0.2 may return MultiIndex columns → flatten
            if isinstance(_h.columns, pd.MultiIndex):
                _h.columns = _h.columns.get_level_values(0)
            _close_s = _h["Close"]
            # If still a DataFrame (multiple tickers), take first column
            if isinstance(_close_s, pd.DataFrame):
                _close_s = _close_s.iloc[:, 0]
            # Apply MARKET_DATE cutoff
            _close_s = _close_s[_close_s.index.normalize() <= _mdate].dropna()
            if not _close_s.empty:
                _rate = float(_close_s.iloc[-1])
                if 10_000 < _rate < 25_000:
                    _USD_IDR_RATE_CACHE = _rate
                    print(f"[INFO] USD/IDR rate as of {_mdate.date()}: {_rate:,.0f}")
                    return _USD_IDR_RATE_CACHE
    except Exception as _e:
        print(f"[WARN] USD/IDR fetch failed: {_e}")
    _USD_IDR_RATE_CACHE = 16_500.0
    print(f"[WARN] Using fallback USD/IDR: {_USD_IDR_RATE_CACHE:,.0f}")
    return _USD_IDR_RATE_CACHE


def idr_billions_fmt(v) -> str:
    """
    Format a monetary value that is ALREADY IN IDR (post-USD conversion).
    Displays as plain number in B IDR (divide by 1e9), matching Stockbit.
    e.g. 65,440,000,000,000 IDR → "65,440"
    """
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    try:
        n = float(v)
    except Exception:
        return v
    return f"{n / 1_000_000_000:,.0f}"


def safe_div(a, b, default=np.nan):
    try:
        if pd.isna(a) or pd.isna(b) or float(b) == 0:
            return default
        return float(a) / float(b)
    except Exception:
        return default

# ── IHSG Beta engine ──────────────────────────────────────────────────────────
_IHSG_RETURNS_CACHE: dict = {}   # keyed by period_days → pd.Series of daily returns

def _get_ihsg_returns(period_days: int = 252) -> "pd.Series | None":
    """Fetch and cache IHSG (^JKSE) daily returns for beta computation."""
    global _IHSG_RETURNS_CACHE
    if period_days in _IHSG_RETURNS_CACHE:
        return _IHSG_RETURNS_CACHE[period_days]
    try:
        import yfinance as yf
        _mdate_fg = get_market_date()
        _end_fg   = (_mdate_fg + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        ihsg = yf.download("^JKSE", period=f"{period_days + 30}d", end=_end_fg,
                           progress=False, auto_adjust=True, threads=False)
        if isinstance(ihsg.columns, pd.MultiIndex):
            ihsg.columns = ihsg.columns.get_level_values(0)
        ihsg = ihsg[ihsg.index.normalize() <= _mdate_fg]
        if ihsg is None or ihsg.empty:
            return None
        closes = ihsg["Close"].squeeze()
        if hasattr(closes, "columns"):
            closes = closes.iloc[:, 0]
        returns = closes.pct_change().dropna()
        if len(returns) < 20:
            return None
        _IHSG_RETURNS_CACHE[period_days] = returns.tail(period_days)
        return _IHSG_RETURNS_CACHE[period_days]
    except Exception:
        return None

def compute_beta_ihsg(stock_hist: "pd.DataFrame", period_days: int = 252) -> float:
    """
    COVARIANCE.P(stock_returns, ihsg_returns) / VAR.P(ihsg_returns)
    Matches the Excel formula requested in Stock Info column.
    Returns np.nan on failure.
    """
    try:
        ihsg_ret = _get_ihsg_returns(period_days)
        if ihsg_ret is None or ihsg_ret.empty:
            return np.nan
        stock_closes = stock_hist["Close"].squeeze()
        if hasattr(stock_closes, "columns"):
            stock_closes = stock_closes.iloc[:, 0]
        stock_ret = stock_closes.pct_change().dropna()
        common_idx = stock_ret.index.intersection(ihsg_ret.index)
        if len(common_idx) < 20:
            return np.nan
        s = stock_ret.loc[common_idx].astype(float).values
        m = ihsg_ret.loc[common_idx].astype(float).values
        cov_p = float(np.cov(s, m, ddof=0)[0, 1])
        var_p = float(np.var(m, ddof=0))
        if var_p == 0:
            return np.nan
        return round(cov_p / var_p, 4)
    except Exception:
        return np.nan


def compute_beta_zone(beta: float) -> str:
    """Classify beta into 6-tier zone label (IBD-style)."""
    if pd.isna(beta):
        return "-"
    if beta > 1.5:
        return "High Beta"
    if beta >= 1.0:
        return "Above Market"
    if beta >= 0.95:
        return "Market Match"
    if beta >= 0.5:
        return "Defensive"
    if beta >= 0.0:
        return "Low Correlation"
    return "Inverse"


def compute_rs_rating(hist: "pd.DataFrame", universe_hist_dict: dict = None) -> int:
    """
    IBD-style RS Rating (1–99) ported from Pine Script.
    Weights: 3M×2, 6M×1, 9M×1, 12M×1 → raw score → compress to 1–99
    against the full universe distribution stored in _RS_UNIVERSE_CACHE.
    Falls back to raw percentile vs own 252D history if universe not available.
    """
    global _RS_UNIVERSE_CACHE
    try:
        closes = hist["Close"].squeeze()
        if hasattr(closes, "columns"):
            closes = closes.iloc[:, 0]
        closes = closes.dropna()
        n = len(closes)
        if n < 63:
            return np.nan

        def ratio(bars):
            bars = min(bars, n - 1)
            if bars <= 0 or pd.isna(closes.iloc[-bars - 1]) or closes.iloc[-bars - 1] == 0:
                return np.nan
            return float(closes.iloc[-1]) / float(closes.iloc[-bars - 1])

        r3  = ratio(63)
        r6  = ratio(126) if n >= 126 else r3
        r9  = ratio(189) if n >= 189 else r6
        r12 = ratio(252) if n >= 252 else r9

        if any(pd.isna(x) for x in [r3, r6, r9, r12]):
            return np.nan

        # IBD formula: RS = 0.40×ROC_3M + 0.20×ROC_6M + 0.20×ROC_9M + 0.20×ROC_12M
        # ROC = price_ratio - 1 (return, not ratio)
        roc3  = r3  - 1.0
        roc6  = r6  - 1.0
        roc9  = r9  - 1.0
        roc12 = r12 - 1.0
        rs_raw = 0.40 * roc3 + 0.20 * roc6 + 0.20 * roc9 + 0.20 * roc12

        # Compress to 1–99 against universe cache
        universe = _RS_UNIVERSE_CACHE
        if len(universe) >= 5:
            lo = min(universe)
            hi = max(universe)
            if hi == lo:
                return 50
            rating = round(1 + 98 * (rs_raw - lo) / (hi - lo))
            return int(max(1, min(99, rating)))
        else:
            # Store raw in cache; return nan until enough universe members
            _RS_UNIVERSE_CACHE.append(rs_raw)
            return np.nan
    except Exception:
        return np.nan


def rs_rating_zone(rs: float) -> str:
    """Rule of thumb classification."""
    if pd.isna(rs):
        return "-"
    if rs >= 95:
        return "Top-Tier Leader (≥95)"
    if rs >= 90:
        return "Leader (≥90)"
    if rs >= 80:
        return "Strong (≥80)"
    if rs >= 60:
        return "Neutral (60–79)"
    return "Weak (<60)"


def compute_today_event(row: dict, hist: "pd.DataFrame") -> str:
    """
    Detect what structural event happened on the latest candle relative to:
    - Swing BOS / CHoCH (from smc fields)
    - Equilibrium zone (midpoint of Strong High / Strong Low)
    - Bull OB zone (smc_closest_ob range)
    - Bear OB zone (smc_closest_ob_bear range)
    Returns descriptive string or '-'.
    """
    try:
        if hist is None or len(hist) < 2:
            return "-"

        close  = float(hist["Close"].iloc[-1])
        open_  = float(hist["Open"].iloc[-1])
        high   = float(hist["High"].iloc[-1])
        low    = float(hist["Low"].iloc[-1])
        prev_c = float(hist["Close"].iloc[-2])

        # BOS / CHoCH from SMC fields
        swing_struct    = str(row.get("smc_latest_swing_struct","") or "")
        internal_struct = str(row.get("smc_latest_internal_struct","") or "")
        for label, src in [("Swing", swing_struct), ("Internal", internal_struct)]:
            if not src or src in ("-","N/A"):
                continue
            s = src.upper()
            if "BOS" in s and "BULL" in s:
                return f"Price turned Bull BOS ({label})"
            if "BOS" in s and "BEAR" in s:
                return f"Price turned Bear BOS ({label})"
            if "CHOCH" in s and "BULL" in s:
                return f"Price turned Bull CHoCH ({label})"
            if "CHOCH" in s and "BEAR" in s:
                return f"Price turned Bear CHoCH ({label})"

        # Equilibrium zone
        sh = safe_num(row.get("smc_strong_high"), np.nan)
        sl = safe_num(row.get("smc_strong_low"),  np.nan)
        if pd.notna(sh) and pd.notna(sl) and sh > sl:
            eq   = (sh + sl) / 2.0
            band = (sh - sl) * 0.05   # ±5% of range = equilibrium band
            eq_top = eq + band
            eq_bot = eq - band
            if prev_c < eq_top and close > eq_top:
                return "Price Break Up Equilibrium Zone"
            if prev_c > eq_bot and close < eq_bot:
                return "Price Break Down Equilibrium Zone"
            if eq_bot <= close <= eq_top:
                return "Price Ranging on Equilibrium Zone"

        def _parse_range(s):
            try:
                parts = str(s).replace("–","-").split("-")
                if len(parts) == 2:
                    lo = float(parts[0].replace(",",""))
                    hi = float(parts[1].replace(",",""))
                    return lo, hi
            except Exception:
                pass
            return None, None

        # Bull OB zone
        ob_bull_str = str(row.get("smc_closest_ob","") or "")
        ob_lo, ob_hi = _parse_range(ob_bull_str)
        if ob_lo and ob_hi:
            if prev_c < ob_hi and close > ob_hi:
                return "Price Break Up Bull OB Zone"
            if prev_c > ob_lo and close < ob_lo:
                return "Price Break Down Bull OB Zone"
            if ob_lo <= close <= ob_hi:
                return "Price Ranging on Bull OB Zone"

        # Bear OB zone
        ob_bear_str = str(row.get("smc_closest_ob_bear","") or "")
        b_lo, b_hi = _parse_range(ob_bear_str)
        if b_lo and b_hi:
            if prev_c < b_hi and close > b_hi:
                return "Price Break Up Bear OB Zone"
            if prev_c > b_lo and close < b_lo:
                return "Price Break Down Bear OB Zone"
            if b_lo <= close <= b_hi:
                return "Price Ranging on Bear OB Zone"

        return "-"
    except Exception:
        return "-"


# Universe RS raw score cache (populated during batch build_row)
_RS_UNIVERSE_CACHE: list = []




# =============================================================================
# UPGRADE 1 — ARA / ARB DETECTION
# =============================================================================
def compute_ara_arb(hist: "pd.DataFrame", board: str = "Main") -> dict:
    """
    Detect if today's price is at the IDX daily auto-rejection limit.
    Main/Development Board: ±35%
    Acceleration Board:     ±10%
    Returns: {ara_arb: label, at_limit: bool, limit_pct: float}
    """
    limit = 0.10 if str(board).lower() in ("acceleration", "akselerasi", "acc") else 0.35
    result = {"ara_arb": "-", "at_limit": False, "limit_pct": limit * 100}
    try:
        if hist is None or len(hist) < 2:
            return result
        prev_close = float(hist["Close"].iloc[-2])
        curr_close = float(hist["Close"].iloc[-1])
        if prev_close <= 0:
            return result
        chg = (curr_close - prev_close) / prev_close
        tol = 0.001   # 0.1% tolerance for floating point
        if chg >= limit - tol:
            result["ara_arb"]  = f"ARA (+{limit*100:.0f}%)"
            result["at_limit"] = True
        elif chg <= -limit + tol:
            result["ara_arb"]  = f"ARB (-{limit*100:.0f}%)"
            result["at_limit"] = True
    except Exception:
        pass
    return result


# =============================================================================
# UPGRADE 2 — FOREIGN FLOW (IDX JATS)
# =============================================================================
_FOREIGN_FLOW_CACHE: dict = {}   # ticker → {foreign_net_lot, foreign_net_val, foreign_activity}

def fetch_foreign_flow(ticker: str) -> dict:
    """
    Fetch foreign net buy/sell from IDX market data.
    Source: idx.co.id/umum/foreign-net-buy-sell endpoint (JSON).
    Falls back to N/A gracefully.
    Returns: foreign_net_lot (lots), foreign_net_val (IDR), foreign_activity label.
    """
    empty = {"foreign_net_lot": np.nan, "foreign_net_val": np.nan, "foreign_activity": "-"}
    clean = ticker.upper().replace(".JK", "").strip()
    if clean in _FOREIGN_FLOW_CACHE:
        return _FOREIGN_FLOW_CACHE[clean]
    try:
        import urllib.request, json
        url = (
            "https://www.idx.co.id/primary/StockData/GetForeignNetBuySellByCode"
            f"?code={clean}&lang=id"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))

        # IDX API wraps in {"data": [...], "totData": N}
        rows = data.get("data") or data.get("Data") or []
        if not rows:
            _FOREIGN_FLOW_CACHE[clean] = empty
            return empty

        latest = rows[0]
        buy_lot  = float(latest.get("ForeignBuy")  or latest.get("foreign_buy")  or 0)
        sell_lot = float(latest.get("ForeignSell") or latest.get("foreign_sell") or 0)
        net_lot  = buy_lot - sell_lot

        buy_val  = float(latest.get("ForeignBuyVal")  or latest.get("foreign_buy_value")  or 0)
        sell_val = float(latest.get("ForeignSellVal") or latest.get("foreign_sell_value") or 0)
        net_val  = buy_val - sell_val

        if net_val > 0:
            activity = "Foreign Net Buy"
        elif net_val < 0:
            activity = "Foreign Net Sell"
        else:
            activity = "Neutral"

        result = {
            "foreign_net_lot": round(net_lot),
            "foreign_net_val": net_val,
            "foreign_activity": activity,
        }
        _FOREIGN_FLOW_CACHE[clean] = result
        return result
    except Exception:
        _FOREIGN_FLOW_CACHE[clean] = empty
        return empty


# =============================================================================
# UPGRADE 3 — BOARD CLASSIFICATION
# =============================================================================
_BOARD_CACHE: dict = {}

def fetch_board_classification(ticker: str) -> str:
    """
    Determine IDX board: Main Board / Development Board / Acceleration Board.
    Sourced from IDX company list API. Falls back to 'Main Board'.
    """
    clean = ticker.upper().replace(".JK", "").strip()
    if clean in _BOARD_CACHE:
        return _BOARD_CACHE[clean]
    try:
        import urllib.request, json
        url = (
            "https://www.idx.co.id/primary/ListedCompany/GetStockList"
            f"?start=0&length=1&code={clean}&name=&sector=&board=&lang=id"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))
        rows = data.get("data") or data.get("Data") or []
        if rows:
            raw_board = str(rows[0].get("Board") or rows[0].get("board") or "").strip()
            board_map = {
                "1": "Main Board", "MAIN": "Main Board", "UTAMA": "Main Board",
                "2": "Development Board", "DEV": "Development Board", "PENGEMBANGAN": "Development Board",
                "3": "Acceleration Board", "ACC": "Acceleration Board", "AKSELERASI": "Acceleration Board",
            }
            board = board_map.get(raw_board.upper(), f"Main Board")
        else:
            board = "Main Board"
    except Exception:
        board = "Main Board"
    _BOARD_CACHE[clean] = board
    return board


# =============================================================================
# UPGRADE 4 — BROKER FLOW (Top Broker Net)
# =============================================================================
_BROKER_CACHE: dict = {}

def fetch_broker_flow(ticker: str) -> dict:
    """
    Fetch top broker net buy/sell summary from IDX broker transaction data.
    Returns top 3 net-buy and net-sell broker codes + net values.
    Source: idx.co.id broker transaction summary endpoint.
    """
    empty = {"top_broker_buy": "-", "top_broker_sell": "-", "broker_net_signal": "-"}
    clean = ticker.upper().replace(".JK", "").strip()
    if clean in _BROKER_CACHE:
        return _BROKER_CACHE[clean]
    try:
        import urllib.request, json
        url = (
            "https://www.idx.co.id/primary/StockData/GetBrokerSummary"
            f"?code={clean}&lang=id"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))

        rows = data.get("data") or data.get("Data") or []
        if not rows:
            _BROKER_CACHE[clean] = empty
            return empty

        broker_nets = {}
        for row in rows:
            code    = str(row.get("BrokerCode") or row.get("broker_code") or "?")
            buy_val = float(row.get("BuyValue")  or row.get("buy_value")  or 0)
            sel_val = float(row.get("SellValue") or row.get("sell_value") or 0)
            broker_nets[code] = buy_val - sel_val

        sorted_brokers = sorted(broker_nets.items(), key=lambda x: x[1], reverse=True)
        top_buy  = ", ".join(f"{k}" for k, v in sorted_brokers[:3] if v > 0) or "-"
        top_sell = ", ".join(f"{k}" for k, v in reversed(sorted_brokers[-3:]) if v < 0) or "-"

        total_net = sum(broker_nets.values())
        signal = "Institutional Accumulation" if total_net > 0 else ("Institutional Distribution" if total_net < 0 else "Neutral")

        result = {
            "top_broker_buy":    top_buy,
            "top_broker_sell":   top_sell,
            "broker_net_signal": signal,
        }
        _BROKER_CACHE[clean] = result
        return result
    except Exception:
        _BROKER_CACHE[clean] = empty
        return empty


# =============================================================================
# UPGRADE 5 — SUSPENSION & CORPORATE ACTION FLAGS
# =============================================================================
_SUSPENSION_CACHE: dict = {}

def fetch_suspension_flag(ticker: str) -> dict:
    """
    Check if ticker is under trading suspension or active corporate action
    (rights issue mid-process, tender offer, etc.) from IDX.
    Source: idx.co.id corporate actions and trading halt endpoints.
    """
    empty = {"suspended": False, "suspension_label": "-", "corp_action_active": "-"}
    clean = ticker.upper().replace(".JK", "").strip()
    if clean in _SUSPENSION_CACHE:
        return _SUSPENSION_CACHE[clean]
    try:
        import urllib.request, json
        url = (
            "https://www.idx.co.id/primary/TradingHalt/GetTradingHalt"
            f"?code={clean}&lang=id"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))

        rows = data.get("data") or data.get("Data") or []
        if rows:
            latest = rows[0]
            halt_type = str(latest.get("HaltType") or latest.get("halt_type") or "")
            label = f"Suspended: {halt_type}" if halt_type else "Trading Halt"
            result = {"suspended": True, "suspension_label": label, "corp_action_active": label}
        else:
            result = {"suspended": False, "suspension_label": "-", "corp_action_active": "-"}

        _SUSPENSION_CACHE[clean] = result
        return result
    except Exception:
        _SUSPENSION_CACHE[clean] = empty
        return empty


# =============================================================================
# UPGRADE 6 — LIQUIDITY-ADJUSTED MAX POSITION SIZE
# =============================================================================
def compute_max_position_size(adtr20: float, close: float, capital: float = 1_000_000_000) -> dict:
    """
    Max Position Size = min(20% of ADTR20, 5% of capital) / Close → shares & lots.
    Conservative IDX illiquidity buffer: never consume more than 20% of avg daily volume value.
    Capital default = 1B IDR (configurable).
    """
    empty = {"max_shares": np.nan, "max_lots": np.nan, "max_position_idr": np.nan, "position_size_pct_adtv": np.nan}
    try:
        if pd.isna(adtr20) or pd.isna(close) or adtr20 <= 0 or close <= 0:
            return empty
        liquidity_cap = adtr20 * 0.20
        capital_cap   = capital * 0.05
        max_idr       = min(liquidity_cap, capital_cap)
        max_shares    = int(max_idr // close)
        max_lots      = max_shares // 100
        pct_adtv      = (max_idr / adtr20) * 100
        return {
            "max_shares":            max_shares,
            "max_lots":              max_lots,
            "max_position_idr":      round(max_idr),
            "position_size_pct_adtv": round(pct_adtv, 1),
        }
    except Exception:
        return empty


# =============================================================================
# UPGRADE 7 — IHSG SEASONAL BIAS CALENDAR
# =============================================================================
def _finalize_flow_position_fields(row: dict, hist: "pd.DataFrame | None" = None) -> dict:
    """Finalize liquidity fields without inventing participant or transaction data."""
    try:
        ps = compute_max_position_size(safe_num(row.get("adtr20"), np.nan), safe_num(row.get("close"), np.nan))
        row["max_shares"]             = ps.get("max_shares", np.nan)
        row["max_lots"]               = ps.get("max_lots", np.nan)
        row["max_position_idr"]       = ps.get("max_position_idr", np.nan)
        row["position_size_pct_adtv"] = ps.get("position_size_pct_adtv", np.nan)
        row["foreign_net_val"] = np.nan
        row["foreign_net_lot"] = np.nan
        row["foreign_activity"] = "-"
        row["top_broker_buy"] = "-"
        row["top_broker_sell"] = "-"
        row["broker_net_signal"] = "-"
    except Exception:
        pass
    return row


def get_seasonal_bias(month: int = None) -> str:
    """
    IHSG seasonal tendencies by calendar month.
    Derives from MARKET_DATE month (not today) for backtest correctness.
    """
    if month is None:
        month = get_market_date().month
    CALENDAR = {
        1:  "Bullish — Jan effect; institutional re-entry",
        2:  "Neutral — Pre-Ramadan positioning",
        3:  "Neutral — Ramadan thin volumes",
        4:  "Neutral — Eid liquidity gap",
        5:  "Cautious — Sell in May; dividend repatriation",
        6:  "Neutral — Mid-year rebalancing; Lebaran gap",
        7:  "Bullish — Post-Lebaran re-entry; Q2 earnings",
        8:  "Neutral — Independence Day; thin flows",
        9:  "Neutral — Pre-Q3 earnings lull",
        10: "Bullish — Q3 earnings catalyst; Oct seasonally strong",
        11: "Bullish — Year-end rally; election-cycle tailwind",
        12: "Bullish — Dec year-end rally; low float",
    }
    return CALENDAR.get(month, "Neutral")


# =============================================================================
# UPGRADE 8 — DIVIDEND TRAP DETECTION
# =============================================================================
def detect_dividend_trap(row: dict) -> str:
    """
    Flag dividend trap risk: upcoming ex-date + weak fundamentals.
    High risk: ex-date < 14 days AND (payout > 80% OR revenue_growth < 0).
    Medium risk: ex-date 14-30 days OR yield > 8% with weak earnings.
    """
    try:
        from datetime import date
        today = get_market_date().date()  # Use MARKET_DATE as reference for news age

        # Check upcoming ex-date proximity
        ex_raw = str(row.get("yf_upcoming_ex_date") or row.get("yf_last_dividend_date") or "")
        days_to_ex = None
        if len(ex_raw) >= 10:
            try:
                ex_date = date.fromisoformat(ex_raw[:10])
                days_to_ex = (ex_date - today).days
            except Exception:
                pass

        if days_to_ex is None or days_to_ex < 0 or days_to_ex > 60:
            return "-"

        payout      = safe_num(row.get("yf_payout_ratio"), np.nan)
        rev_growth  = safe_num(row.get("yf_revenue_growth"), np.nan)
        div_yield   = safe_num(row.get("yf_dividend_yield"), np.nan)
        earnings_gr = safe_num(row.get("yf_earnings_growth"), np.nan)

        weak_payout   = pd.notna(payout)      and payout > 0.80
        weak_growth   = pd.notna(rev_growth)  and rev_growth < 0
        high_yield    = pd.notna(div_yield)   and div_yield > 0.08
        weak_earnings = pd.notna(earnings_gr) and earnings_gr < 0

        if days_to_ex <= 14 and (weak_payout or weak_growth):
            return f"High Risk  — Ex-date in {days_to_ex}d, weak fundamentals"
        if days_to_ex <= 30 and high_yield and weak_earnings:
            return f"Medium Risk  — Ex-date in {days_to_ex}d, yield trap indicator"
        if days_to_ex <= 14:
            return f"Watch  — Ex-date in {days_to_ex}d"
        return "-"
    except Exception:
        return "-"


# =============================================================================
# UPGRADE 9 — OB AGE (DAYS SINCE ORDER BLOCK FORMED)
# =============================================================================
def compute_ob_age(hist: "pd.DataFrame", ob_level: float, lookback: int = 60) -> int:
    """
    Find how many bars ago the Order Block level (midpoint) was first formed.
    Scans backward through hist to find the pivot that generated the OB.
    Returns age in trading days; 0 if OB is today, np.nan if not found.
    """
    if hist is None or len(hist) < 5 or pd.isna(ob_level) or ob_level <= 0:
        return np.nan
    try:
        window = hist.tail(lookback)
        highs  = window["High"].values
        lows   = window["Low"].values
        for i in range(len(window) - 1, -1, -1):
            bar_mid = (highs[i] + lows[i]) / 2
            if abs(bar_mid - ob_level) / ob_level < 0.015:   # within 1.5%
                return len(window) - 1 - i   # bars since that candle
        return np.nan
    except Exception:
        return np.nan


# =============================================================================
# UPGRADE 10 — RRG QUADRANT FETCH
# =============================================================================
_RRG_QUADRANT_CACHE: dict = {}   # sector → quadrant

def get_rrg_quadrant_for_sector(sector: str, rrg_data: dict = None) -> str:
    """
    Look up the RRG quadrant for this stock's sector.
    rrg_data is a pre-built dict: {sector_name: quadrant_label} passed from the
    RRG computation pass (IDX Overview builder).
    Falls back to _RRG_QUADRANT_CACHE populated during IDX Overview build.
    Returns: Leading / Weakening / Lagging / Improving / N/A
    """
    if not sector or sector in ("-", "N/A", ""):
        return "-"
    if rrg_data:
        return rrg_data.get(str(sector).strip(), "-")
    return _RRG_QUADRANT_CACHE.get(str(sector).strip(), "-")





def safe_div_series(a, b, default=np.nan):
    try:
        result = a / b
        if hasattr(result, "replace"):
            result = result.replace([np.inf, -np.inf], np.nan)
            if default is not np.nan:
                result = result.fillna(default)
        return result
    except Exception:
        try:
            idx = getattr(a, "index", None) or getattr(b, "index", None)
            return pd.Series(default, index=idx)
        except Exception:
            return default

def _zone_bucket(close_val, vwap, m1, m2, m3):
    """Return integer zone bucket for a close relative to VWAP SD bands.
    1  = above VWAP
    0  = VWAP to -1 SD
    -1 = -1 SD to -2 SD
    -2 = -2 SD to -3 SD
    -3 = below -3 SD
    None = cannot compute
    """
    try:
        c = float(close_val)
        v = float(vwap)
        if pd.isna(c) or pd.isna(v):
            return None
        if c >= v:
            return 1
        if pd.notna(m1) and c >= float(m1):
            return 0
        if pd.notna(m2) and c >= float(m2):
            return -1
        if pd.notna(m3) and c >= float(m3):
            return -2
        return -3
    except Exception:
        return None

def _count_consecutive_zone_days(hist: pd.DataFrame, levels, current_zone_text: str) -> int:
    """Count consecutive trading days price has been in the same VWAP zone bucket as today.
    Uses fixed VWAP/SD levels applied backwards through recent history.
    levels = [vwap, m1, m2, m3]
    """
    try:
        if hist is None or hist.empty or not levels or len(levels) < 1:
            return 0
        vwap = safe_num(levels[0], np.nan)
        if pd.isna(vwap):
            return 0
        m1 = safe_num(levels[1], np.nan) if len(levels) > 1 else np.nan
        m2 = safe_num(levels[2], np.nan) if len(levels) > 2 else np.nan
        m3 = safe_num(levels[3], np.nan) if len(levels) > 3 else np.nan

        closes = hist["Close"].astype(float).values
        if len(closes) == 0:
            return 0
        current_bucket = _zone_bucket(closes[-1], vwap, m1, m2, m3)
        if current_bucket is None:
            return 0
        count = 0
        for i in range(len(closes) - 1, -1, -1):
            b = _zone_bucket(closes[i], vwap, m1, m2, m3)
            if b == current_bucket:
                count += 1
            else:
                break
        return max(0, count)
    except Exception:
        return 0

def clamp(x, min_val, max_val):
    try:
        if pd.isna(x):
            return np.nan
        return max(min_val, min(max_val, float(x)))
    except Exception:
        return np.nan

def normalize_score(value, low, high, inverse=False):
    try:
        if pd.isna(value):
            return np.nan
        if high == low:
            return 50.0
        score = ((float(value) - low) / (high - low)) * 100.0
        score = clamp(score, 0.0, 100.0)
        if inverse and pd.notna(score):
            score = 100.0 - score
        return score
    except Exception:
        return np.nan

def slope_last_n(series: pd.Series, n):
    try:
        s = pd.Series(series).dropna().tail(n)
        if len(s) < 2:
            return np.nan
        x = np.arange(len(s), dtype=float)
        y = s.astype(float).values
        return float(np.polyfit(x, y, 1)[0])
    except Exception:
        return np.nan

def classify_trend_from_slope(value, flat_band):
    if pd.isna(value):
        return "Flat"
    if value > flat_band:
        return "Rising"
    if value < -flat_band:
        return "Falling"
    return "Flat"

def percentile_or_threshold_score(value, thresholds=None):
    try:
        v = safe_num(value)
        if pd.isna(v):
            return np.nan
        thresholds = thresholds or [
            (1_000_000_000_000, 100),
            (500_000_000_000, 90),
            (100_000_000_000, 80),
            (50_000_000_000, 70),
            (10_000_000_000, 60),
            (5_000_000_000, 50),
            (1_000_000_000, 40),
            (500_000_000, 30),
            (100_000_000, 20),
        ]
        for t, s in thresholds:
            if v >= t:
                return float(s)
        return 10.0
    except Exception:
        return np.nan

def map_grade(score):
    s = safe_num(score)
    if pd.isna(s):
        return "N/A"
    if s >= 85: return "A+"
    if s >= 75: return "A"
    if s >= 65: return "B+"
    if s >= 55: return "B"
    if s >= 45: return "C+"
    if s >= 35: return "C"
    if s >= 25: return "D"
    return "F"

def safe_range_position(close, low, high):
    try:
        if pd.isna(close) or pd.isna(low) or pd.isna(high) or float(high) == float(low):
            return np.nan
        return clamp(((float(close) - float(low)) / (float(high) - float(low))) * 100.0, 0.0, 100.0)
    except Exception:
        return np.nan

def compute_institutional_metrics(hist: pd.DataFrame, row: Dict[str, Any]) -> Dict[str, Any]:
    # Safe OHLCV-only PROXY mode. No external dependency.
    out = {
        "flow_data_mode": "PROXY",
        "dollar_vol_20d_avg": np.nan,
        "turnover_velocity_20d": np.nan,
        "ad_trend": "Flat",
        "obv_trend": "Flat",
        "cmf20": np.nan,
        "volume_sponsorship_score": np.nan,
        "accumulation_score": np.nan,
        "distribution_score": np.nan,
        "smart_money_bias": "Neutral",
        "flow_conviction": "Neutral",
        "net_participation_proxy": np.nan,
        "sponsorship_grade": "N/A",

        "pos_52w_pct": np.nan,
        "range_compression_20d": np.nan,
        "vol_compression_20d": np.nan,
        "base_length": np.nan,
        "breakout_pressure": np.nan,
        "breakdown_pressure": np.nan,
        "_removed_spring": "NO",
        "_removed_upthrust": "NO",
        "structure_state": "N/A",
        "wyckoff_proxy_phase": "Transitional",
        "markup_readiness": np.nan,
        "breakdown_risk": np.nan,
        "cause_quality": "Weak",
        "wyckoff_event_start_date": "N/A",
        "wyckoff_event_confirm_date": "N/A",

        "technical_core_score": np.nan,
        "flow_score": np.nan,
        "structure_score": np.nan,
        "risk_penalty": np.nan,
        "regime_multiplier": np.nan,
        "adaptive_composite_score": np.nan,
        "setup_quality": "N/A",
        "trap_risk": "N/A",
        "institutional_verdict": "N/A",
        "institutional_action_bias": "N/A",

        "institutional_accumulation": "NO",
        "early_markup_candidate": "NO",
        "smart_money_pullback": "NO",
        "breakout_watchlist": "NO",
        "distribution_warning": "NO",
        "failed_breakout_risk": "NO",
        "capital_efficient_trend": "NO",
        "speculative_momentum": "NO",
        "preset_summary": "None",
    }
    if hist is None or hist.empty or len(hist) < 20:
        return out

    close = hist["Close"].astype(float)
    high = hist["High"].astype(float)
    low = hist["Low"].astype(float)
    vol = hist["Volume"].astype(float)

    # Core series
    dollar_vol = close * vol
    out["dollar_vol_20d_avg"] = safe_num(dollar_vol.tail(20).mean())

    mcap = safe_num(row.get("mcap"))
    if pd.notna(out["dollar_vol_20d_avg"]) and pd.notna(mcap) and mcap > 0:
        out["turnover_velocity_20d"] = safe_num((out["dollar_vol_20d_avg"] / mcap) * 100.0)

    hl_range = (high - low).replace(0, np.nan)
    mfm = (((close - low) - (high - close)) / hl_range).replace([np.inf, -np.inf], np.nan).fillna(0)
    mfv = mfm * vol
    ad_line = mfv.cumsum()

    close_prev = close.shift(1)
    obv = np.where(close > close_prev, vol, np.where(close < close_prev, -vol, 0))
    obv = pd.Series(obv, index=hist.index).cumsum()

    out["ad_trend"] = classify_trend_from_slope(slope_last_n(ad_line, 20), flat_band=max(abs(safe_num(vol.tail(20).mean(), 1)), 1))
    out["obv_trend"] = classify_trend_from_slope(slope_last_n(obv, 20), flat_band=max(abs(safe_num(vol.tail(20).mean(), 1)), 1))

    cmf_num = mfv.rolling(20).sum()
    cmf_den = vol.rolling(20).sum().replace(0, np.nan)
    cmf = safe_div(cmf_num.iloc[-1], cmf_den.iloc[-1], default=np.nan)
    out["cmf20"] = safe_num(cmf)

    avg20 = safe_num(vol.tail(20).mean())
    avg60 = safe_num(vol.tail(60).mean()) if len(vol) >= 60 else avg20
    vol_expansion = safe_div(avg20, avg60, default=np.nan)

    up_mask = close.diff() > 0
    down_mask = close.diff() < 0
    up_vol = safe_num(vol.tail(20)[up_mask.tail(20)].sum(), 0)
    down_vol = safe_num(vol.tail(20)[down_mask.tail(20)].sum(), 0)
    uv_ratio = safe_div(up_vol, down_vol, default=2.0 if up_vol > 0 and down_vol == 0 else np.nan)

    obv_slope_norm = safe_div(slope_last_n(obv, 20), max(avg20, 1), default=0)
    ad_slope_norm = safe_div(slope_last_n(ad_line, 20), max(avg20, 1), default=0)

    dollar_score = percentile_or_threshold_score(out["dollar_vol_20d_avg"])
    vol_exp_score = normalize_score(vol_expansion, 0.6, 2.0)
    uv_score = normalize_score(uv_ratio, 0.7, 2.0)
    cmf_pos_score = normalize_score(out["cmf20"], -0.2, 0.2)
    obv_score = normalize_score(obv_slope_norm, -1.5, 1.5)

    out["volume_sponsorship_score"] = safe_num(clamp(
        (0.30 * (0 if pd.isna(dollar_score) else dollar_score)) +
        (0.25 * (0 if pd.isna(vol_exp_score) else vol_exp_score)) +
        (0.20 * (0 if pd.isna(uv_score) else uv_score)) +
        (0.15 * (0 if pd.isna(cmf_pos_score) else cmf_pos_score)) +
        (0.10 * (0 if pd.isna(obv_score) else obv_score)),
        0, 100
    ))

    range20_high = safe_num(high.tail(20).max())
    range20_low = safe_num(low.tail(20).min())
    price_pos_20 = safe_range_position(close.iloc[-1], range20_low, range20_high)
    pos_close_ratio = safe_num((close.tail(20).diff() > 0).mean() * 100.0)
    neg_close_ratio = safe_num((close.tail(20).diff() < 0).mean() * 100.0)

    ad_up_score = normalize_score(ad_slope_norm, -1.5, 1.5)
    ad_dn_score = normalize_score(ad_slope_norm, -1.5, 1.5, inverse=True)
    obv_up_score = normalize_score(obv_slope_norm, -1.5, 1.5)
    obv_dn_score = normalize_score(obv_slope_norm, -1.5, 1.5, inverse=True)

    out["accumulation_score"] = safe_num(clamp(
        0.30 * (0 if pd.isna(cmf_pos_score) else cmf_pos_score) +
        0.25 * (0 if pd.isna(ad_up_score) else ad_up_score) +
        0.20 * (0 if pd.isna(obv_up_score) else obv_up_score) +
        0.15 * (0 if pd.isna(price_pos_20) else price_pos_20) +
        0.10 * (0 if pd.isna(pos_close_ratio) else pos_close_ratio),
        0, 100
    ))
    out["distribution_score"] = safe_num(clamp(
        0.30 * (0 if pd.isna(100 - cmf_pos_score) else 100 - cmf_pos_score) +
        0.25 * (0 if pd.isna(ad_dn_score) else ad_dn_score) +
        0.20 * (0 if pd.isna(obv_dn_score) else obv_dn_score) +
        0.15 * (0 if pd.isna(100 - price_pos_20) else 100 - price_pos_20) +
        0.10 * (0 if pd.isna(neg_close_ratio) else neg_close_ratio),
        0, 100
    ))

    a = safe_num(out["accumulation_score"])
    d = safe_num(out["distribution_score"])
    if pd.notna(a) and pd.notna(d):
        if a >= 65 and (a - d) >= 15:
            out["smart_money_bias"] = "Accumulation"
        elif d >= 65 and (d - a) >= 15:
            out["smart_money_bias"] = "Distribution"
        elif a < 45 and d < 45:
            out["smart_money_bias"] = "Neutral"
        else:
            out["smart_money_bias"] = "Mixed"

    vss = safe_num(out["volume_sponsorship_score"])
    smb = out["smart_money_bias"]
    if smb == "Accumulation" and pd.notna(vss) and vss >= 70:
        out["flow_conviction"] = "Strong Accumulation"
    elif smb == "Distribution" and pd.notna(vss) and vss >= 70:
        out["flow_conviction"] = "Strong Distribution"
    elif smb == "Accumulation" and pd.notna(vss) and vss >= 50:
        out["flow_conviction"] = "Mild Accumulation"
    elif smb == "Distribution" and pd.notna(vss) and vss >= 50:
        out["flow_conviction"] = "Mild Distribution"
    else:
        out["flow_conviction"] = "Neutral"

    np_proxy = (
        40 * safe_num(normalize_score(out["cmf20"], -0.2, 0.2), 50) / 100.0 +
        25 * safe_num(normalize_score(obv_slope_norm, -1.5, 1.5), 50) / 100.0 +
        20 * safe_num(normalize_score(ad_slope_norm, -1.5, 1.5), 50) / 100.0 +
        15 * safe_num(normalize_score(uv_ratio, 0.7, 2.0), 50) / 100.0
    )
    out["net_participation_proxy"] = safe_num(clamp((np_proxy - 50) * 2, -100, 100))
    out["sponsorship_grade"] = map_grade(out["volume_sponsorship_score"])

    # Structure / Wyckoff Proxy (V6 state-machine style)
    high_52 = safe_num(high.tail(252).max())
    low_52 = safe_num(low.tail(252).min())
    out["pos_52w_pct"] = safe_range_position(close.iloc[-1], low_52, high_52)

    raw_range_pct = safe_div((high.tail(20).max() - low.tail(20).min()), close.iloc[-1], default=np.nan)
    raw_range_pct = raw_range_pct * 100 if pd.notna(raw_range_pct) else np.nan
    out["range_compression_20d"] = safe_num(clamp(100 - normalize_score(raw_range_pct, 8, 25), 0, 100))

    prev_close = close.shift(1)
    tr = pd.concat([high - low, (high - prev_close).abs(), (low - prev_close).abs()], axis=1).max(axis=1)
    atr14 = tr.rolling(14).mean()
    atr_pct_series = safe_div_series(atr14, close, default=np.nan) * 100
    curr_atr_pct = safe_num(atr_pct_series.iloc[-1])
    med60 = safe_num(atr_pct_series.tail(60).median()) if len(atr_pct_series.dropna()) >= 20 else curr_atr_pct
    atr_ratio = safe_div(curr_atr_pct, med60, default=np.nan)
    out["vol_compression_20d"] = safe_num(clamp(100 - normalize_score(atr_ratio, 0.6, 1.6), 0, 100))

    base_len = 0
    for n in range(min(120, len(hist)), 19, -1):
        window = hist.tail(n)
        range_n = safe_div(window["High"].max() - window["Low"].min(), close.iloc[-1], default=np.nan)
        ema20_w = window["Close"].ewm(span=20, adjust=False).mean().iloc[-1]
        sma50_w = window["Close"].rolling(min(50, len(window))).mean().iloc[-1]
        spread = abs(safe_div(ema20_w - sma50_w, close.iloc[-1], default=np.nan))
        if pd.notna(range_n) and range_n <= 0.25 and (pd.isna(spread) or spread <= 0.08):
            base_len = n
            break
    out["base_length"] = base_len

    hh20 = safe_num(high.tail(20).max())
    ll20 = safe_num(low.tail(20).min())
    prox_high = 100 - normalize_score(abs(safe_div(hh20 - close.iloc[-1], close.iloc[-1], default=np.nan))*100, 0, 12)
    prox_low = 100 - normalize_score(abs(safe_div(close.iloc[-1] - ll20, close.iloc[-1], default=np.nan))*100, 0, 12)

    higher_lows = 100 if low.tail(10).min() > low.tail(20).min() else 40
    lower_highs = 100 if high.tail(10).max() < high.tail(20).max() else 40
    momentum_pos = 100 if safe_num(row.get("rsi14")) >= 50 else 30
    momentum_neg = 100 if safe_num(row.get("rsi14")) < 50 else 30

    out["breakout_pressure"] = safe_num(clamp(
        0.30 * safe_num(prox_high, 0) +
        0.25 * safe_num(out["volume_sponsorship_score"], 0) +
        0.20 * higher_lows +
        0.15 * safe_num(out["pos_52w_pct"], 0) +
        0.10 * momentum_pos, 0, 100
    ))
    out["breakdown_pressure"] = safe_num(clamp(
        0.30 * safe_num(prox_low, 0) +
        0.25 * safe_num(out["distribution_score"], 0) +
        0.20 * lower_highs +
        0.15 * safe_num(100 - out["pos_52w_pct"], 0) +
        0.10 * momentum_neg, 0, 100
    ))

    recent_support = low.shift(1).rolling(60, min_periods=20).min()
    recent_res = high.shift(1).rolling(60, min_periods=20).max()
    ema20_all = close.ewm(span=20, adjust=False).mean()
    sma50_all = close.rolling(50).mean()
    sma200_all = close.rolling(200).mean()

    up_trend = (
        pd.notna(ema20_all.iloc[-1]) and pd.notna(sma50_all.iloc[-1]) and
        close.iloc[-1] > ema20_all.iloc[-1] > sma50_all.iloc[-1]
    )
    down_trend = (
        pd.notna(ema20_all.iloc[-1]) and pd.notna(sma50_all.iloc[-1]) and
        close.iloc[-1] < ema20_all.iloc[-1] < sma50_all.iloc[-1]
    )

    # V6 bar-by-bar state candidates
    avg20v = vol.rolling(20).mean()
    spring_cand = (
        (low < recent_support) &
        (close > recent_support) &
        (vol >= avg20v.fillna(vol))
    ).fillna(False)
    spring_confirm = (
        (spring_cand.shift(1).fillna(False) | spring_cand.shift(2).fillna(False)) &
        (close > recent_support.fillna(close.shift(1))) &
        (close >= ema20_all.fillna(close)) &
        (close >= close.shift(1))
    ).fillna(False)

    utad_cand = (
        (high > recent_res) &
        (close < recent_res) &
        (vol >= avg20v.fillna(vol))
    ).fillna(False)
    utad_confirm = (
        (utad_cand.shift(1).fillna(False) | utad_cand.shift(2).fillna(False)) &
        (close < recent_res.fillna(close.shift(1))) &
        (close <= ema20_all.fillna(close))
    ).fillna(False)

    base_state = (
        (((high.rolling(20).max() - low.rolling(20).min()) / close.replace(0, np.nan)) * 100 <= 18) &
        (close >= low.rolling(20).min() + (high.rolling(20).max() - low.rolling(20).min()) * 0.25) &
        (close <= low.rolling(20).min() + (high.rolling(20).max() - low.rolling(20).min()) * 0.80)
    ).fillna(False)

    constructive = (
        (close >= ema20_all.fillna(close)) &
        ((ema20_all >= sma50_all.fillna(ema20_all)) | sma50_all.isna())
    ).fillna(False)
    weak_context = (
        (close <= ema20_all.fillna(close)) &
        ((ema20_all <= sma50_all.fillna(ema20_all)) | sma50_all.isna())
    ).fillna(False)

    sos_state = (
        spring_confirm.rolling(5, min_periods=1).max().astype(bool) &
        constructive &
        (close >= high.shift(1))
    ).fillna(False)

    lps_state = (
        constructive &
        (close < close.shift(1)) &
        (close >= ema20_all.fillna(close)) &
        (close >= close.rolling(10).min())
    ).fillna(False)

    markup_state = (
        constructive &
        (close > high.rolling(20).max().shift(1).fillna(close))
    ).fillna(False)

    continuation_state = (
        markup_state &
        (vol >= avg20v.fillna(vol))
    ).fillna(False)

    distribution_state = (
        weak_context &
        (safe_num(out["pos_52w_pct"], 50) >= 60) &
        (safe_num(out["distribution_score"], 0) >= 55)
    )
    distribution_state = pd.Series([distribution_state] * len(close), index=close.index) if isinstance(distribution_state, bool) else distribution_state

    sow_state = (
        utad_confirm.rolling(5, min_periods=1).max().astype(bool) &
        weak_context &
        (close <= low.shift(1))
    ).fillna(False)

    lpsy_state = (
        weak_context &
        (close > close.shift(1)) &
        (close <= ema20_all.fillna(close))
    ).fillna(False)

    markdown_state = (
        weak_context &
        (close < low.rolling(20).min().shift(1).fillna(close))
    ).fillna(False)

    # Determine current active state with priority / memory
    current_event = "No Clean Event"
    current_phase = "Transitional"

    if bool(markdown_state.iloc[-1]):
        current_phase = "Markdown"
        current_event = "Trend Weakness"
    elif bool(sow_state.iloc[-1]):
        current_phase = "Distribution"
        current_event = "SOW"
    elif bool(utad_confirm.iloc[-1]) or bool(utad_cand.iloc[-1]):
        current_phase = "Distribution"
        current_event = "UTAD"
    elif bool(lpsy_state.iloc[-1]) and safe_num(out["distribution_score"], 0) >= 55:
        current_phase = "Distribution"
        current_event = "LPSY"
    elif bool(continuation_state.iloc[-1]):
        current_phase = "Markup"
        current_event = "Continuation"
    elif bool(markup_state.iloc[-1]) and bool(lps_state.iloc[-1]):
        current_phase = "Markup"
        current_event = "Pullback"
    elif bool(sos_state.iloc[-1]):
        current_phase = "Accumulation"
        current_event = "SOS"
    elif bool(lps_state.iloc[-1]) and safe_num(out["accumulation_score"], 0) >= 55:
        current_phase = "Accumulation"
        current_event = "LPS"
    elif bool(spring_confirm.iloc[-1]) or bool(spring_cand.iloc[-1]):
        current_phase = "Accumulation"
        current_event = "Spring"
    elif bool(base_state.iloc[-1]):
        current_phase = "Accumulation"
        current_event = "Base Build"
    else:
        # fallback to prior regime-style mapping
        if up_trend and safe_num(out["breakout_pressure"], 0) >= 60:
            current_phase = "Markup"
            current_event = "Continuation"
        elif down_trend and safe_num(out["breakdown_pressure"], 0) >= 60:
            current_phase = "Markdown"
            current_event = "Trend Weakness"
        elif safe_num(out["distribution_score"], 0) >= 60 and safe_num(out["pos_52w_pct"], 50) >= 60:
            current_phase = "Distribution"
            current_event = "Distribution Range"
        elif safe_num(out["range_compression_20d"], 0) >= 55:
            current_phase = "Accumulation"
            current_event = "Base Build"
        else:
            current_phase = "Transitional"
            current_event = "No Clean Event"

    # Map structure state from new state engine
    if current_phase == "Accumulation":
        out["structure_state"] = "Tight Base" if safe_num(out["range_compression_20d"], 0) >= 70 else "Loose Base"
    elif current_phase == "Markup":
        out["structure_state"] = "Advancing Trend"
    elif current_phase == "Distribution":
        out["structure_state"] = "Distribution Range"
    elif current_phase == "Markdown":
        out["structure_state"] = "Breakdown Structure"
    else:
        out["structure_state"] = "Weakening Trend"

    out["wyckoff_proxy_phase"] = current_phase
    out["wyckoff_event_label"] = current_event

    base_quality_map = normalize_score(base_len, 20, 100)
    out["markup_readiness"] = safe_num(clamp(
        0.25 * safe_num(out["range_compression_20d"], 0) +
        0.20 * safe_num(out["vol_compression_20d"], 0) +
        0.20 * safe_num(out["breakout_pressure"], 0) +
        0.15 * safe_num(out["accumulation_score"], 0) +
        0.10 * safe_num(out["volume_sponsorship_score"], 0) +
        0.10 * safe_num(base_quality_map, 0), 0, 100
    ))
    upthrust_penalty = 100 if current_event == "UTAD" else (60 if bool(utad_cand.iloc[-1]) else 0)
    out["breakdown_risk"] = safe_num(clamp(
        0.25 * safe_num(out["breakdown_pressure"], 0) +
        0.20 * safe_num(out["distribution_score"], 0) +
        0.15 * safe_num(100 - out["pos_52w_pct"], 0) +
        0.15 * momentum_neg +
        0.15 * upthrust_penalty +
        0.10 * safe_num(100 - out["volume_sponsorship_score"], 0), 0, 100
    ))

    out["_removed_spring"] = "YES" if current_event == "Spring" else ("Weak" if bool(spring_cand.iloc[-1]) else "NO")
    out["_removed_upthrust"] = "YES" if current_event == "UTAD" else ("Weak" if bool(utad_cand.iloc[-1]) else "NO")

    if base_len >= 60 and safe_num(out["range_compression_20d"], 0) >= 70 and current_phase == "Accumulation":
        out["cause_quality"] = "Excellent"
    elif base_len >= 40 and safe_num(out["range_compression_20d"], 0) >= 55:
        out["cause_quality"] = "Good"
    elif base_len >= 20:
        out["cause_quality"] = "Average"
    else:
        out["cause_quality"] = "Weak"


    # Adaptive verdict
    tech_parts = []
    tech_parts.append(100 if row.get("summary_ma") == "Above All MA" else (75 if str(row.get("summary_ma", "")).startswith("Above ") else 30))
    tech_parts.append(100 if row.get("cross_status") == "Golden" else (20 if row.get("cross_status") == "Dead" else 50))
    rsi14 = safe_num(row.get("rsi14"))
    tech_parts.append(safe_num(normalize_score(rsi14, 35, 70), 50))
    tech_parts.append(80 if row.get("div_signal") == "Bullish" else (20 if row.get("div_signal") == "Bearish" else 50))
    tech_parts.append(80 if row.get("rvol_zone") == "YES" else 40)
    out["technical_core_score"] = safe_num(np.nanmean(pd.Series(tech_parts, dtype=float)))

    fc_map = {"Strong Accumulation": 100, "Mild Accumulation": 75, "Neutral": 50, "Mild Distribution": 25, "Strong Distribution": 0}
    out["flow_score"] = safe_num(clamp(
        0.30 * safe_num(out["volume_sponsorship_score"], 0) +
        0.25 * safe_num(out["accumulation_score"], 0) +
        0.25 * safe_num(100 - out["distribution_score"], 0) +
        0.10 * safe_num((out["net_participation_proxy"] + 100) / 2, 50) +
        0.10 * fc_map.get(out["flow_conviction"], 50), 0, 100
    ))
    cause_map = {"Excellent": 100, "Good": 75, "Average": 50, "Weak": 25}
    out["structure_score"] = safe_num(clamp(
        0.25 * safe_num(out["markup_readiness"], 0) +
        0.20 * safe_num(100 - out["breakdown_risk"], 0) +
        0.20 * safe_num(out["breakout_pressure"], 0) +
        0.15 * safe_num(out["range_compression_20d"], 0) +
        0.10 * safe_num(out["vol_compression_20d"], 0) +
        0.10 * cause_map.get(out["cause_quality"], 25), 0, 100
    ))

    liq = str(row.get("liquidity_category", ""))
    mc = str(row.get("market_cap_category", ""))
    weak_liq_pen = 100 if liq in ("Very Low Liquidity",) else (70 if liq in ("Low Liquidity",) else (30 if liq in ("Medium Liquidity",) else 0))
    small_pen = 100 if mc == "Micro Cap" else (70 if mc == "Small Cap" else (25 if mc == "Mid Cap" else 0))
    bearish_div_pen = 100 if row.get("div_signal") == "Bearish" else 0
    out["risk_penalty"] = safe_num(clamp(
        0.30 * safe_num(out["breakdown_risk"], 0) +
        0.25 * safe_num(out["distribution_score"], 0) +
        0.15 * upthrust_penalty +
        0.10 * weak_liq_pen +
        0.10 * small_pen +
        0.10 * bearish_div_pen, 0, 100
    ))

    regime_mult = 1.0
    if mc == "Large Cap" and liq == "High Liquidity":
        regime_mult = 1.10
    elif mc == "Mid Cap" and liq in ("High Liquidity", "Medium Liquidity"):
        regime_mult = 1.04
    elif mc in ("Small Cap",) and liq in ("Medium Liquidity",):
        regime_mult = 0.98
    elif mc in ("Small Cap", "Micro Cap") and liq in ("Low Liquidity", "Very Low Liquidity"):
        regime_mult = 0.88
    if out["smart_money_bias"] == "Accumulation" and out["structure_state"] == "Tight Base":
        regime_mult += 0.03
    out["regime_multiplier"] = clamp(regime_mult, 0.80, 1.15)

    raw = (
        0.40 * safe_num(out["technical_core_score"], 0) +
        0.30 * safe_num(out["flow_score"], 0) +
        0.20 * safe_num(out["structure_score"], 0) -
        0.10 * safe_num(out["risk_penalty"], 0)
    )
    out["adaptive_composite_score"] = safe_num(clamp(raw * out["regime_multiplier"], 0, 100))

    acs = safe_num(out["adaptive_composite_score"])
    if pd.notna(acs):
        if acs >= 85: out["setup_quality"] = "Elite"
        elif acs >= 75: out["setup_quality"] = "High Quality"
        elif acs >= 65: out["setup_quality"] = "Constructive"
        elif acs >= 55: out["setup_quality"] = "Developing"
        elif acs >= 45: out["setup_quality"] = "Speculative"
        else: out["setup_quality"] = "Weak"

    rp = safe_num(out["risk_penalty"])
    if pd.notna(rp):
        if rp >= 75: out["trap_risk"] = "High"
        elif rp >= 60: out["trap_risk"] = "Elevated"
        elif rp >= 45: out["trap_risk"] = "Moderate"
        else: out["trap_risk"] = "Low"
    if out["_removed_upthrust"] == "YES" and safe_num(out["distribution_score"], 0) >= 60 and out["trap_risk"] == "Low":
        out["trap_risk"] = "Elevated"

    if safe_num(out["adaptive_composite_score"], 0) >= 80 and safe_num(out["flow_score"], 0) >= 65 and safe_num(out["structure_score"], 0) >= 65:
        out["institutional_verdict"] = "Institutional Long Candidate"
    elif 65 <= safe_num(out["adaptive_composite_score"], 0) < 80 and out["trap_risk"] != "High":
        out["institutional_verdict"] = "Constructive Long Watchlist"
    elif safe_num(out["breakdown_risk"], 0) >= 70 and safe_num(out["distribution_score"], 0) >= 65:
        out["institutional_verdict"] = "Breakdown Risk"
    elif safe_num(out["adaptive_composite_score"], 0) < 50 and safe_num(out["risk_penalty"], 0) >= 60:
        out["institutional_verdict"] = "Avoid / Distribution Risk"
    else:
        out["institutional_verdict"] = "Neutral / Selective"

    if out["institutional_verdict"] == "Institutional Long Candidate" and out["structure_state"] in ("Tight Base", "Loose Base"):
        out["institutional_action_bias"] = "Breakout Watch"
    elif out["institutional_verdict"] in ("Institutional Long Candidate", "Constructive Long Watchlist") and out["smart_money_bias"] == "Accumulation":
        out["institutional_action_bias"] = "Accumulate on Pullback"
    elif out["institutional_verdict"] in ("Institutional Long Candidate", "Constructive Long Watchlist"):
        out["institutional_action_bias"] = "Hold / Trend Follow"
    elif out["institutional_verdict"] == "Breakdown Risk":
        out["institutional_action_bias"] = "Reduce Into Strength"
    elif out["institutional_verdict"] == "Avoid / Distribution Risk":
        out["institutional_action_bias"] = "Avoid"
    else:
        out["institutional_action_bias"] = "Range Only"

    # Presets
    out["institutional_accumulation"] = "YES" if out["smart_money_bias"] == "Accumulation" and safe_num(out["volume_sponsorship_score"], 0) >= 65 and safe_num(out["accumulation_score"], 0) >= 65 and safe_num(out["distribution_score"], 100) <= 50 else "NO"
    out["early_markup_candidate"] = "YES" if out["wyckoff_proxy_phase"] in ("Accumulation", "Transitional") and safe_num(out["markup_readiness"], 0) >= 70 and safe_num(out["breakout_pressure"], 0) >= 65 and safe_num(out["breakdown_risk"], 100) <= 45 else "NO"
    bullish_trend = str(row.get("summary_ma", "")).startswith("Above") and safe_num(row.get("close")) >= safe_num(row.get("ema20"), safe_num(row.get("close"), 0))
    out["smart_money_pullback"] = "YES" if bullish_trend and out["smart_money_bias"] != "Distribution" and safe_num(out["flow_score"], 0) >= 55 else "NO"
    out["breakout_watchlist"] = "YES" if safe_num(out["markup_readiness"], 0) >= 75 and safe_num(out["breakout_pressure"], 0) >= 70 and safe_num(out["range_compression_20d"], 0) >= 60 and safe_num(out["volume_sponsorship_score"], 0) >= 60 else "NO"
    out["distribution_warning"] = "YES" if (out["smart_money_bias"] == "Distribution" or safe_num(out["distribution_score"], 0) >= 70) and (out["_removed_upthrust"] in ("YES", "Weak") or out["structure_state"] in ("Weakening Trend", "Distribution Range")) else "NO"
    out["failed_breakout_risk"] = "YES" if out["_removed_upthrust"] == "YES" or (safe_num(out["pos_52w_pct"], 0) >= 70 and safe_num(out["breakdown_risk"], 0) >= 65 and safe_num(out["distribution_score"], 0) >= 60) else "NO"
    healthy_tv = pd.notna(out["turnover_velocity_20d"]) and 0.10 <= out["turnover_velocity_20d"] <= 5.0
    out["capital_efficient_trend"] = "YES" if bullish_trend and healthy_tv and safe_num(out["volume_sponsorship_score"], 0) >= 55 and safe_num(out["risk_penalty"], 100) <= 40 and safe_num(out["adaptive_composite_score"], 0) >= 70 else "NO"
    out["speculative_momentum"] = "YES" if mc in ("Small Cap", "Micro Cap") and safe_num(out["technical_core_score"], 0) >= 65 and safe_num(out["volume_sponsorship_score"], 0) >= 60 and out["trap_risk"] != "High" else "NO"

    labels = []
    mapping = [
        ("institutional_accumulation", "Institutional Accumulation"),
        ("early_markup_candidate", "Early Markup Candidate"),
        ("breakout_watchlist", "Breakout Watchlist"),
        ("smart_money_pullback", "Smart Money Pullback"),
        ("capital_efficient_trend", "Capital Efficient Trend"),
        ("distribution_warning", "Distribution Warning"),
        ("failed_breakout_risk", "Failed Breakout Risk"),
        ("speculative_momentum", "Speculative Momentum"),
    ]
    for k, label in mapping:
        if out.get(k) == "YES":
            labels.append(label)
    out["preset_summary"] = " | ".join(labels) if labels else "None"

    return out


def _fmt_date_or_na(ts):
    try:
        if ts is None or pd.isna(ts):
            return "N/A"
        return pd.Timestamp(ts).strftime("%Y-%m-%d")
    except Exception:
        return "N/A"

# ALPHAFLOW PROXY MAX (NO BROKER - SAFE LAYER)
# =========================
def safe_pct(a, b, default=np.nan):
    try:
        if pd.isna(a) or pd.isna(b) or float(b) == 0:
            return default
        return (float(a) / float(b)) * 100.0
    except Exception:
        return default

def compute_alphaflow_proxy_max(hist: pd.DataFrame, row: dict):
    out = {
        "af_smt_proxy": np.nan,
        "af_verdict": "Neutral",
        "af_verdict_confidence": np.nan,
        "af_phase_event": "Neutral / No Clean Event",
        "af_flow_edge": "Low",
        "af_hit_rate_proxy": np.nan,
        "af_r2_proxy": np.nan,
        "af_execution_guide": "Wait for better alignment",
        "af_watch_next": "Need better flow + structure confirmation",
        "af_invalidation": "Lose support / fail VWAP-QVWAP acceptance",
        "af_playbook": "No clean setup",
    }
    if hist is None or hist.empty or len(hist) < 30:
        return out

    df = hist.copy()
    close = df["Close"].astype(float)
    high = df["High"].astype(float)
    low = df["Low"].astype(float)
    vol = df["Volume"].astype(float)

    # Flow proxies
    ret = close.pct_change()
    clv = ((close - low) - (high - close)) / (high - low).replace(0, np.nan)
    mfv = (clv * vol).replace([np.inf, -np.inf], np.nan).fillna(0)
    obv = (np.sign(close.diff().fillna(0)) * vol).fillna(0).cumsum()

    flow20 = safe_num(mfv.tail(20).sum(), 0)
    flow60 = safe_num(mfv.tail(60).sum(), 0)
    dv20 = safe_num((close * vol).tail(20).sum(), np.nan)
    dv60 = safe_num((close * vol).tail(60).sum(), np.nan)

    fp20 = safe_pct(flow20, dv20, 0)
    fp60 = safe_pct(flow60, dv60, 0)
    obv20 = safe_pct(safe_num(obv.iloc[-1] - obv.iloc[max(0, len(obv)-20)], 0), safe_num(vol.tail(20).sum(), np.nan), 0)
    obv60 = safe_pct(safe_num(obv.iloc[-1] - obv.iloc[max(0, len(obv)-60)], 0), safe_num(vol.tail(60).sum(), np.nan), 0)

    # Participation / sponsorship
    vol20 = vol.tail(20)
    part_conc = np.nan
    burst = np.nan
    if len(vol20) >= 5 and vol20.sum() > 0:
        part_conc = (vol20.nlargest(min(3, len(vol20))).sum() / vol20.sum()) * 100.0
        burst = ((vol20 > (vol20.mean() * 1.5)).sum() / len(vol20)) * 100.0

    # Flow-price validation
    corr20 = np.nan
    corr40 = np.nan
    for n in [20, 40]:
        f = mfv.pct_change().replace([np.inf, -np.inf], np.nan).tail(n)
        r = ret.tail(n)
        valid = pd.concat([f, r], axis=1).dropna()
        if len(valid) >= max(8, n // 2):
            corr = valid.iloc[:, 0].corr(valid.iloc[:, 1])
            if n == 20:
                corr20 = corr
            else:
                corr40 = corr
    r2 = ((corr20 or 0) ** 2) * 100 if pd.notna(corr20) else np.nan
    hit = np.nan
    hv = ((ret.tail(20) > 0) == (mfv.pct_change().tail(20) > 0)).astype(float)
    if len(hv.dropna()) > 0:
        hit = hv.mean() * 100.0

    # SMT proxy (0-100) inspired by guide weights, but OHLCV proxy only
    net_flow_score = normalize_score(fp20, -8, 8)
    persist_score = normalize_score((obv20 * 0.7) + (obv60 * 0.3), -8, 8)
    concentration_score = normalize_score(safe_num(part_conc, 50), 35, 70)
    absorption_proxy = normalize_score(safe_num(row.get("range_compression_20d", np.nan), 50), 30, 80)
    execution_proxy = normalize_score(safe_num(row.get("close_vs_qvwap_pct", np.nan), 0), -6, 6)
    smt = (
        0.30 * safe_num(net_flow_score, 50) +
        0.25 * safe_num(persist_score, 50) +
        0.20 * safe_num(concentration_score, 50) +
        0.15 * safe_num(absorption_proxy, 50) +
        0.10 * safe_num(execution_proxy, 50)
    )
    out["af_smt_proxy"] = safe_num(clamp(smt, 0, 100))

    # Map flow edge similar to guide thresholds
    out["af_r2_proxy"] = safe_num(r2)
    out["af_hit_rate_proxy"] = safe_num(hit)

    if pd.notna(r2) and pd.notna(hit):
        if r2 >= 10 and hit >= 58:
            out["af_flow_edge"] = "Strong"
        elif r2 >= 3 and hit >= 52:
            out["af_flow_edge"] = "Moderate"
        else:
            out["af_flow_edge"] = "Weak"

    # Verdict synthesis
    accum = safe_num(row.get("accumulation_score", np.nan), 50)
    dist = safe_num(row.get("distribution_score", np.nan), 50)
    flow = safe_num(row.get("flow_score", np.nan), 50)
    struct = safe_num(row.get("structure_score", np.nan), 50)
    risk = safe_num(row.get("risk_penalty", np.nan), 50)
    qvwap_ok = 1 if safe_num(row.get("close_vs_qvwap_pct", -999), -999) >= -1.5 else 0
    div = str(row.get("divergence_summary", "None") or "None")
    phase = str(row.get("wyckoff_proxy_phase", "Transitional") or "Transitional")
    spring = str(row.get("_removed_spring", "NO") or "NO")
    ut = str(row.get("_removed_upthrust", "NO") or "NO")

    bull_bonus = 0
    bear_penalty = 0
    if "Bullish" in div:
        bull_bonus += 8
    if "Bearish" in div:
        bear_penalty += 8
    if spring in ("YES", "Weak"):
        bull_bonus += 6
    if ut in ("YES", "Weak"):
        bear_penalty += 6

    verdict_score = (
        0.26 * accum +
        0.22 * flow +
        0.16 * struct +
        0.16 * safe_num(out["af_smt_proxy"], 50) +
        0.10 * (100 if qvwap_ok else 35) +
        0.10 * (70 if out["af_flow_edge"] == "Strong" else 55 if out["af_flow_edge"] == "Moderate" else 40)
        - 0.18 * risk
        - 0.12 * dist
        + bull_bonus - bear_penalty
    )

    # Normalize around website-like 5-state outcome
    centered = (verdict_score - 50) / 100.0
    confidence = clamp(
        abs(accum - dist) * 0.45 +
        abs(safe_num(out["af_smt_proxy"], 50) - 50) * 0.35 +
        abs(flow - 50) * 0.20,
        35, 92
    )
    out["af_verdict_confidence"] = safe_num(confidence)

    if centered > 0.35:
        out["af_verdict"] = "STRONG ACCUMULATION"
    elif centered > 0.05:
        out["af_verdict"] = "ACCUMULATION"
    elif centered < -0.35:
        out["af_verdict"] = "STRONG DISTRIBUTION"
    elif centered < -0.05:
        out["af_verdict"] = "DISTRIBUTION"
    else:
        out["af_verdict"] = "NEUTRAL"

    # Phase + event proxy (V6: align with rebuilt state engine)
    current_event = str(row.get("wyckoff_event_label", "") or "")
    if not current_event:
        if phase == "Accumulation":
            current_event = "Spring" if spring == "YES" else ("SOS" if safe_num(row.get("markup_readiness", 0), 0) >= 70 else "Base Build")
        elif phase == "Markup":
            current_event = "Pullback" if qvwap_ok and "Bullish" in div else "Continuation"
        elif phase == "Distribution":
            current_event = "UTAD" if ut == "YES" else "SOW"
        elif phase == "Markdown":
            current_event = "Trend Weakness"
        else:
            current_event = "No Clean Event"

    if phase == "Accumulation":
        out["af_phase_event"] = f"Accumulation / {current_event}"
    elif phase == "Markup":
        out["af_phase_event"] = f"Markup / {current_event}"
    elif phase == "Distribution":
        out["af_phase_event"] = f"Distribution / {current_event}"
    elif phase == "Markdown":
        out["af_phase_event"] = "Markdown / Trend Weakness"
    else:
        out["af_phase_event"] = "Transitional / No Clean Event"

    # Execution guide (replace tier-thinking with playbook-thinking)
    if out["af_verdict"] in ("STRONG ACCUMULATION", "ACCUMULATION"):
        if phase in ("Accumulation", "Transitional"):
            out["af_execution_guide"] = "Wait for hold above QVWAP + stronger RVOL, then accumulate on constructive pullback."
            out["af_playbook"] = "Accumulation Setup"
        elif phase == "Markup":
            out["af_execution_guide"] = "Trend-follow only: buy pullback into QVWAP/EMA20 if support holds."
            out["af_playbook"] = "Markup Continuation"
        else:
            out["af_execution_guide"] = "Bullish score but phase mismatch — reduce size, wait for structure confirmation."
            out["af_playbook"] = "Watchlist Only"
        out["af_watch_next"] = "Need QVWAP acceptance, sustained flow score, and no failed reclaim."
        out["af_invalidation"] = "Lose QVWAP + break prior swing low / support."
    elif out["af_verdict"] in ("DISTRIBUTION", "STRONG DISTRIBUTION"):
        out["af_execution_guide"] = "Avoid fresh longs. Only reconsider after re-accumulation evidence or stronger reclaim."
        out["af_playbook"] = "Distribution Warning"
        out["af_watch_next"] = "Need absorption + reclaim above QVWAP + distribution score cooling."
        out["af_invalidation"] = "Sustained reclaim above QVWAP and improving flow/structure."
    else:
        out["af_execution_guide"] = "No edge. Wait for clearer phase + flow alignment."
        out["af_playbook"] = "Neutral / No Setup"
        out["af_watch_next"] = "Monitor for Spring, SOS, LPS, or stronger flow-price validation."
        out["af_invalidation"] = "N/A"

    return out


# =========================
# WYCKOFF STATE-BASED DATE HELPERS (V5)
# =========================
def _find_last_true_run(mask: pd.Series):
    """
    Return (start_idx, end_idx) of the most recent contiguous True run.
    If the last value is False, returns (None, None).
    """
    try:
        if mask is None or len(mask) == 0:
            return None, None
        m = pd.Series(mask).fillna(False).astype(bool)
        if not bool(m.iloc[-1]):
            return None, None
        end_idx = m.index[-1]
        start_pos = len(m) - 1
        while start_pos - 1 >= 0 and bool(m.iloc[start_pos - 1]):
            start_pos -= 1
        start_idx = m.index[start_pos]
        return start_idx, end_idx
    except Exception:
        return None, None

def _first_confirm_after(mask: pd.Series, start_idx):
    """
    Return first index at/after start_idx where mask is True.
    """
    try:
        if mask is None or len(mask) == 0 or start_idx is None:
            return None
        m = pd.Series(mask).fillna(False).astype(bool)
        sub = m.loc[m.index >= start_idx]
        trues = sub[sub]
        return trues.index[0] if len(trues) else None
    except Exception:
        return None


def _fmt_wyckoff_date(idx):
    try:
        if idx is None:
            return None
        return pd.Timestamp(idx).to_pydatetime()
    except Exception:
        return None

# MARKET STRUCTURE (Leviathan-style pivot / BOS / CHoCH)
# =========================
MS_SWING_SIZE = 20
MS_BOS_CONFIRM = "close"   # "close" or "wick"
MS_MIN_HISTORY = 120
MS_VALUE_LOOKBACK = 20

def _fmt_ms_date(ts):
    try:
        return pd.Timestamp(ts).strftime("%d %b '%y")
    except Exception:
        return ""

def _pivot_high_ms(high: pd.Series, left: int, right: int) -> pd.Series:
    out = pd.Series(np.nan, index=high.index, dtype=float)
    vals = high.values
    n = len(vals)
    for i in range(left, n - right):
        win = vals[i-left:i+right+1]
        if np.isnan(vals[i]) or np.isnan(win).any():
            continue
        if vals[i] == np.max(win):
            out.iloc[i] = vals[i]
    return out

def _pivot_low_ms(low: pd.Series, left: int, right: int) -> pd.Series:
    out = pd.Series(np.nan, index=low.index, dtype=float)
    vals = low.values
    n = len(vals)
    for i in range(left, n - right):
        win = vals[i-left:i+right+1]
        if np.isnan(vals[i]) or np.isnan(win).any():
            continue
        if vals[i] == np.min(win):
            out.iloc[i] = vals[i]
    return out

def _classify_market_structure_swings(hist: pd.DataFrame, swing_size: int = MS_SWING_SIZE):
    h = hist.copy()
    h["ms_piv_hi"] = _pivot_high_ms(h["High"], swing_size, swing_size)
    h["ms_piv_lo"] = _pivot_low_ms(h["Low"], swing_size, swing_size)

    prev_high = np.nan
    prev_low = np.nan
    high_active = False
    low_active = False
    prev_breakout_dir = 0  # +1 up, -1 down

    swings = []   # [{"idx","type","price"}]
    events = []   # [{"idx","event","dir","level","confirm_price","value_ratio"}]

    h["value_traded"] = h["Close"] * h["Volume"].fillna(0)
    h["value_ma20"] = h["value_traded"].rolling(MS_VALUE_LOOKBACK).mean()

    for i in range(len(h)):
        idx = h.index[i]

        piv_hi = h["ms_piv_hi"].iloc[i]
        piv_lo = h["ms_piv_lo"].iloc[i]

        if pd.notna(piv_hi):
            swing_type = "HH" if (pd.isna(prev_high) or piv_hi >= prev_high) else "LH"
            swings.append({"idx": idx, "type": swing_type, "price": float(piv_hi)})
            prev_high = float(piv_hi)
            high_active = True

        if pd.notna(piv_lo):
            swing_type = "HL" if (pd.isna(prev_low) or piv_lo >= prev_low) else "LL"
            swings.append({"idx": idx, "type": swing_type, "price": float(piv_lo)})
            prev_low = float(piv_lo)
            low_active = True

        high_src = h["Close"].iloc[i] if MS_BOS_CONFIRM == "close" else h["High"].iloc[i]
        low_src = h["Close"].iloc[i] if MS_BOS_CONFIRM == "close" else h["Low"].iloc[i]

        val_now = safe_num(h["value_traded"].iloc[i], np.nan)
        val_avg = safe_num(h["value_ma20"].iloc[i], np.nan)
        value_ratio = (val_now / val_avg) if (pd.notna(val_now) and pd.notna(val_avg) and val_avg > 0) else np.nan

        if high_active and pd.notna(prev_high) and pd.notna(high_src) and high_src > prev_high:
            event = "Bull CHoCH" if prev_breakout_dir == -1 else "Bull BOS"
            events.append({
                "idx": idx,
                "event": event,
                "dir": 1,
                "level": prev_high,
                "confirm_price": float(high_src),
                "value_ratio": value_ratio,
            })
            high_active = False
            prev_breakout_dir = 1

        if low_active and pd.notna(prev_low) and pd.notna(low_src) and low_src < prev_low:
            event = "Bear CHoCH" if prev_breakout_dir == 1 else "Bear BOS"
            events.append({
                "idx": idx,
                "event": event,
                "dir": -1,
                "level": prev_low,
                "confirm_price": float(low_src),
                "value_ratio": value_ratio,
            })
            low_active = False
            prev_breakout_dir = -1

    swings = sorted(swings, key=lambda x: x["idx"])
    events = sorted(events, key=lambda x: x["idx"])
    return swings, events

def _market_structure_state(swings, last_event: str):
    if len(swings) < 2:
        if last_event == "Bull CHoCH":
            return "Transition Up"
        if last_event == "Bear CHoCH":
            return "Transition Down"
        return "Range"

    last_high = None
    last_low = None

    for s in reversed(swings):
        if s["type"] in ("HH", "LH") and last_high is None:
            last_high = s["type"]
        if s["type"] in ("HL", "LL") and last_low is None:
            last_low = s["type"]
        if last_high is not None and last_low is not None:
            break

    if last_event == "Bull CHoCH":
        return "Transition Up"
    if last_event == "Bear CHoCH":
        return "Transition Down"

    if last_high == "HH" and last_low == "HL":
        return "HH-HL"
    if last_high == "LH" and last_low == "LL":
        return "LH-LL"
    if last_high == "HH" and last_low == "LL":
        return "HH-LL"
    if last_high == "LH" and last_low == "HL":
        return "LH-HL"

    return "Range"

def _market_structure_volume_regime(hist: pd.DataFrame, last_event_obj: dict, trend_regime: str):
    if hist is None or hist.empty or "Volume" not in hist.columns:
        return "Neutral"

    h = hist.copy()
    h["value_traded"] = h["Close"] * h["Volume"].fillna(0)
    h["value_ma20"] = h["value_traded"].rolling(MS_VALUE_LOOKBACK).mean()

    last_val = safe_num(h["value_traded"].iloc[-1], np.nan)
    avg_val = safe_num(h["value_ma20"].iloc[-1], np.nan)
    ratio = (last_val / avg_val) if (pd.notna(last_val) and pd.notna(avg_val) and avg_val > 0) else np.nan

    if last_event_obj:
        ev_ratio = safe_num(last_event_obj.get("value_ratio"), np.nan)
        ev = str(last_event_obj.get("event", ""))
        if ev in ("Bull BOS", "Bull CHoCH") and pd.notna(ev_ratio) and ev_ratio >= 1.5:
            return "Expansion Up"
        if ev in ("Bear BOS", "Bear CHoCH") and pd.notna(ev_ratio) and ev_ratio >= 1.5:
            return "Expansion Down"

    if pd.notna(ratio):
        if ratio < 0.70:
            return "Dry-Up"
        if ratio >= 1.0 and trend_regime in ("Bullish", "Bullish Weak", "Bearish", "Bearish Weak"):
            return "Supportive"

    return "Neutral"

def compute_market_structure_v1(hist: pd.DataFrame, swing_size: int = MS_SWING_SIZE) -> dict:
    out = {
        "ms_trend_regime": "Sideways",
        "ms_structure_state": "Range",
        "ms_last_event": "None",
        "ms_last_event_date": "",
        "ms_last_swing": "",
        "ms_last_swing_date": "",
        "ms_swing_pattern": "Range",
        "ms_volume_regime": "Neutral",
    }

    if hist is None or hist.empty or len(hist) < MS_MIN_HISTORY:
        return out

    h = hist.copy()
    if not isinstance(h.index, pd.DatetimeIndex):
        try:
            if "Date" in h.columns:
                h.index = pd.to_datetime(h["Date"])
            else:
                h.index = pd.to_datetime(h.index)
        except Exception:
            return out

    h = h.sort_index().copy()
    for c in ["Open", "High", "Low", "Close", "Volume"]:
        if c in h.columns:
            h[c] = pd.to_numeric(h[c], errors="coerce")

    h = h.dropna(subset=["High", "Low", "Close"])
    if len(h) < MS_MIN_HISTORY:
        return out

    swings, events = _classify_market_structure_swings(h, swing_size=swing_size)

    last_event_obj = events[-1] if events else None
    last_event = last_event_obj["event"] if last_event_obj else "None"
    last_event_date = _fmt_ms_date(last_event_obj["idx"]) if last_event_obj else ""

    last_swing_obj = swings[-1] if swings else None
    last_swing = last_swing_obj["type"] if last_swing_obj else ""
    last_swing_date = _fmt_ms_date(last_swing_obj["idx"]) if last_swing_obj else ""

    if len(swings) >= 2:
        swing_pattern = f"{swings[-2]['type']}→{swings[-1]['type']}"
    elif len(swings) == 1:
        swing_pattern = swings[-1]["type"]
    else:
        swing_pattern = "Range"

    structure_state = _market_structure_state(swings, last_event)

    trend_regime = "Sideways"
    if last_event == "Bull BOS":
        trend_regime = "Bullish" if structure_state == "HH-HL" else "Bullish Weak"
    elif last_event == "Bear BOS":
        trend_regime = "Bearish" if structure_state == "LH-LL" else "Bearish Weak"
    elif last_event == "Bull CHoCH":
        trend_regime = "Bullish Weak"
    elif last_event == "Bear CHoCH":
        trend_regime = "Bearish Weak"
    else:
        if structure_state == "HH-HL":
            trend_regime = "Bullish Weak"
        elif structure_state == "LH-LL":
            trend_regime = "Bearish Weak"
        else:
            trend_regime = "Sideways"

    volume_regime = _market_structure_volume_regime(h, last_event_obj, trend_regime)

    out.update({
        "ms_trend_regime": trend_regime,
        "ms_structure_state": structure_state,
        "ms_last_event": last_event,
        "ms_last_event_date": last_event_date,
        "ms_last_swing": last_swing,
        "ms_last_swing_date": last_swing_date,
        "ms_swing_pattern": swing_pattern,
        "ms_volume_regime": volume_regime,
    })
    return out


# =========================
# DYNAMIC DETAIL SCHEMA
# =========================
DETAIL_SCHEMA = [
    ("Stock Info", FILL_GROUP_STOCK, [
        ("ticker",           "Ticker",           8,  None,    "center"),
        ("close",            "Closing Price",    10, "#,##0", "center"),
        ("pct_change",       "Price Change %",   10, "0.00%", "center"),
        ("beta_ihsg",        "Beta (vs IHSG)",   12, "0.00",  "center"),
        ("beta_ihsg_zone",   "Beta (vs IHSG) Zone", 16, None, "center"),
        ("rs_rating",        "RS Rating",        12, "#,##0", "center"),
        ("rs_rating_zone",   "RS Rating Zone",   16, None,    "center"),
        # ("ara_arb", "ARA/ARB") — deleted Item #24
        # ("today_event", "Today Event") — deleted Item #24
        ("emiten",           "Emiten",           38, None,    "left"),
        ("idx_sector",       "IDX Sector",       20, None,    "left"),
        ("sector",           "Sector",           20, None,    "left"),
        ("industry",         "Industry",         35, None,    "left"),
    ]),

    ("Ownership", FILL_GROUP_OWNER, [
        ("investors", "Investors", 35, None, "left"),
        ("free_float", "Free Float", 10, "0.00", "center"),
        ("hhi", "Classic HHI", 12, "#,##0", "center"),
        ("cr1", "CR1", 8, "0.00", "center"),
        ("cr3", "CR3", 8, "0.00", "center"),
        ("holder", "Total Investor >1%", 16, "#,##0", "center"),
        ("ccs", "CCS", 10, "0.00", "center"),
        ("ownership_type", "Ownership Type", 18, None, "center"),
        ("ccs_category", "CCS Category", 16, None, "center"),
    ]),

    ("Stock Regime", FILL_GROUP_MP, [
        ("mcap",                  "Market Cap",            14, "compact", "center"),
        ("market_cap_category",   "Market Cap Categories", 20, None,     "center"),
        ("liquidity_category",    "Liquidity Categories",  20, None,     "center"),
        ("stock_regime",          "Verdict Weight Profiles",24, None,    "center"),
    ]),

    ("Market Structure", FILL_GROUP_MS, [
        ("ms_trend_regime",    "Trend Regime",    16, None, "center"),
        ("ms_structure_state", "Structure State", 16, None, "center"),
        ("ms_last_event",      "Last Event",      14, None, "center"),
        ("ms_last_event_date", "Last Event Date", 14, None, "center"),
        ("ms_last_swing",      "Last Swing",      12, None, "center"),
        ("ms_last_swing_date", "Last Swing Date", 14, None, "center"),
        ("ms_swing_pattern",   "Swing Pattern",   14, None, "center"),
        ("ms_volume_regime",   "Volume Regime",   16, None, "center"),
    ]),

    ("Liquidity", FILL_GROUP_VOL, [
        ("lot",                    "Lot",                        12, "compact", "center"),
        ("daily_value",            "Value (Approx)",             16, "compact", "center"),
        ("adtr20",                 "Average Value 20 D (Approx)",20, "compact", "center"),
        ("last_vol",               "Volume",                     14, "compact", "center"),
        ("adtv20",                 "Average Volume 20 D",        18, "compact", "center"),
        ("rvol20",                 "RVOL 20 D",                  12, "0.00",   "center"),
        ("rvol20_zone",            "RVOL 20 D Zone",             16, None,     "center"),
        ("rvol20_chg_pct",        "RVOL Change %",              14, "0.00%",  "center"),
        ("rvol20_chg_flag",       "RVOL Change Zone",           18, None,     "center"),
        ("adr_pct",                "ADR %",                      10, "0.00",   "center"),
        ("atr14_pct",              "ATR (14) %",                 12, "0.00",   "center"),
        ("adr_atr_zone",           "ADR & ATR (14) Zone",        18, None,     "center"),
    ]),

    ("Market Profile", FILL_GROUP_MP, [
        ("ibh", "IBH", 10, "#,##0", "center"),
        ("ibl", "IBL", 10, "#,##0", "center"),
        ("mp_zone", "vs IBL", 14, None, "center"),
        ("pwh", "PWH", 10, "#,##0", "center"),
        ("pwl", "PWL", 10, "#,##0", "center"),
        ("price_ge_pwl", "vs PWL", 14, None, "center"),
        ("mdh", "MDH", 10, "#,##0", "center"),
        ("mdl", "MDL", 10, "#,##0", "center"),
        ("price_ge_mdl", "vs MDL", 14, None, "center"),
        ("mp_summary", "MP Summary", 34, None, "left"),
    ]),

    ("Current MVWAP (June 2026)", FILL_GROUP_Q, [
        ("cm_days",      "Running Days",  10, "#,##0", "center"),
        ("cm_p3",        "3 σ",           10, "#,##0", "center"),
        ("cm_p2",        "2 σ",           10, "#,##0", "center"),
        ("cm_p1",        "1 σ",           10, "#,##0", "center"),
        ("cm_vwap",      "VWAP",          10, "#,##0", "center"),
        ("cm_m1",        "-1 σ",          10, "#,##0", "center"),
        ("cm_m2",        "-2 σ",          10, "#,##0", "center"),
        ("cm_m3",        "-3 σ",          10, "#,##0", "center"),
        ("cm_sd",        "Price σ",       10, "0.00",  "center"),
        ("cm_delta",     "Price σ Δ 1D", 10, "0.00",  "center"),
        ("cm_remarks",   "VWAP Zone",     38, None,     "center"),
        ("cm_zone_days", "VWAP Zone Days",14, "#,##0", "center"),
    ]),

    ("Previous MVWAP (May 2026)", FILL_GROUP_PQ, [
        ("pm_days",      "Running Days",  10, "#,##0", "center"),
        ("pm_p3",        "3 σ",           10, "#,##0", "center"),
        ("pm_p2",        "2 σ",           10, "#,##0", "center"),
        ("pm_p1",        "1 σ",           10, "#,##0", "center"),
        ("pm_vwap",      "VWAP",          10, "#,##0", "center"),
        ("pm_m1",        "-1 σ",          10, "#,##0", "center"),
        ("pm_m2",        "-2 σ",          10, "#,##0", "center"),
        ("pm_m3",        "-3 σ",          10, "#,##0", "center"),
        ("pm_sd",        "Price σ",       10, "0.00",  "center"),
        ("pm_delta",     "Price σ Δ 1D", 14, "0.00",  "center"),
        ("pm_remarks",   "VWAP Zone",     38, None,     "center"),
        ("pm_zone_days", "VWAP Zone Days",14, "#,##0", "center"),
    ]),

    ("Current QVWAP (Q2 2026)", FILL_GROUP_Q, [
        ("q_days",      "Running Days",  10, "#,##0", "center"),
        ("q_p3",        "3 σ",           10, "#,##0", "center"),
        ("q_p2",        "2 σ",           10, "#,##0", "center"),
        ("q_p1",        "1 σ",           10, "#,##0", "center"),
        ("q_vwap",      "VWAP",          10, "#,##0", "center"),
        ("q_m1",        "-1 σ",          10, "#,##0", "center"),
        ("q_m2",        "-2 σ",          10, "#,##0", "center"),
        ("q_m3",        "-3 σ",          10, "#,##0", "center"),
        ("q_sd",        "Price σ",       10, "0.00",  "center"),
        ("q_delta",     "Price σ Δ 1D", 10, "0.00",  "center"),
        ("q_remarks",   "VWAP Zone",     38, None,     "center"),
        ("q_zone_days", "VWAP Zone Days",14, "#,##0", "center"),
    ]),

    ("Previous QVWAP (Q1 2026)", FILL_GROUP_PQ, [
        ("pq_days",      "Running Days",  10, "#,##0", "center"),
        ("pq_p3",        "3 σ",           10, "#,##0", "center"),
        ("pq_p2",        "2 σ",           10, "#,##0", "center"),
        ("pq_p1",        "1 σ",           10, "#,##0", "center"),
        ("pq_vwap",      "VWAP",          10, "#,##0", "center"),
        ("pq_m1",        "-1 σ",          10, "#,##0", "center"),
        ("pq_m2",        "-2 σ",          10, "#,##0", "center"),
        ("pq_m3",        "-3 σ",          10, "#,##0", "center"),
        ("pq_sd",        "Price σ",       10, "0.00",  "center"),
        ("pq_delta",     "Price σ Δ 1D", 14, "0.00",  "center"),
        ("pq_remarks",   "VWAP Zone",     38, None,     "center"),
        ("pq_zone_days", "VWAP Zone Days",14, "#,##0", "center"),
    ]),

    ("Previous Year VWAP (2025)", FILL_GROUP_PY, [
        ("py_year",      "Running Days",  10, "#,##0", "center"),
        ("py_p3",        "3 σ",           10, "#,##0", "center"),
        ("py_p2",        "2 σ",           10, "#,##0", "center"),
        ("py_p1",        "1 σ",           10, "#,##0", "center"),
        ("py_vwap",      "VWAP",          10, "#,##0", "center"),
        ("py_m1",        "-1 σ",          10, "#,##0", "center"),
        ("py_m2",        "-2 σ",          10, "#,##0", "center"),
        ("py_m3",        "-3 σ",          10, "#,##0", "center"),
        ("py_sd",        "Price σ",       10, "0.00",  "center"),
        ("py_delta",     "Price σ Δ 1D", 14, "0.00",  "center"),
        ("py_remarks",   "VWAP Zone",     38, None,     "center"),
        ("py_zone_days", "VWAP Zone Days",14, "#,##0", "center"),
    ]),

    ("Moving Average", FILL_GROUP_MA, [
        ("ema25", "EMA 25", 10, "#,##0", "center"),
        ("ema25d", "EMA 25 %diff", 12, "0.00", "center"),
        ("ema25p", "EMA 25 Pos", 12, None, "center"),
        ("ema50", "EMA 50", 10, "#,##0", "center"),
        ("ema50d", "EMA 50 %diff", 12, "0.00", "center"),
        ("ema50p", "EMA 50 Pos", 12, None, "center"),
        ("sma200", "SMA 200", 10, "#,##0", "center"),
        ("sma200d", "SMA 200 %diff", 12, "0.00", "center"),
        ("sma200p", "SMA 200 Pos", 12, None, "center"),
        ("summary_ma", "MA Zone", 28, None, "center"),
    ]),

    ("RSI Momentum", FILL_GROUP_MOM, [
        ("rsi14", "RSI 14", 10, "0.00", "center"),
        ("rsi_delta", "RSI 14 Δ 1D", 12, "0.00", "center"),
        ("rsi_status", "RSI Status", 14, None, "center"),
        ("rsi_ma14", "RSI MA 14", 10, "0.00", "center"),
        ("rsi_pos", "RSI Pos", 10, None, "center"),
        ("cross_status", "RSI Cross", 16, None, "center"),
        ("div_signal", "Divergence Signal", 16, None, "center"),
        # "Strong"/"Medium"/"Weak" = regular divergence confidence; "Hidden" =
        # hidden divergence (continuation, not reversal) -- divergence_signals()
        # already computes this distinction (see strength_rank), it just never
        # had its own exported column before -- only ever appeared merged into
        # the "divergence_summary" display string.
        ("div_strength", "Divergence Strength", 16, None, "center"),
        ("div_ref1_date", "Divergence Start Date", 18, None, "center"),
        ("div_ref2_date", "Divergence Confirm Date", 18, None, "center"),
    ]),

    ("MACD Momentum", FILL_GROUP_MACD_MOM, [
        ("macd_line",        "MACD Line",       12, "0.0000", "center"),
        ("macd_signal_line", "Signal Line",     12, "0.0000", "center"),
        ("macd_hist",        "Histogram (EMA3)", 14, "0.0000", "center"),
        ("macd_position",    "Lines Position",  18, None,     "center"),
        ("macd_wave",        "Wave Pattern",    22, None,     "center"),
        ("macd_cross",       "MACD Cross",      16, None,     "center"),
    ]),

    ("Stochastic", FILL_GROUP_MACD_MOM, [
        ("stoch_k",     "Stochastic %K",     14, "0.00", "center"),
        ("stoch_d",     "Stochastic %D",     14, "0.00", "center"),
        ("stoch_cross", "Stochastic Cross",  16, None,   "center"),
    ]),

]



# =========================
# MACD 4C SMOOTH ENGINE  (Pine Script: MACD 4C Smooth v6 — 1:1 port)
# fast=12, slow=26, signal=9, histSmooth=3
# =========================
def compute_macd_momentum(hist: pd.DataFrame) -> dict:
    """
    1:1 port of Pine Script 'MACD 4C Smooth v6':
      - EMA fast 12, slow 26, signal EMA 9
      - Histogram smoothed with EMA(3) before colour classification
      - 4-colour bar state based on smoothed hist vs previous bar
      - Entry signal = Blue bar (histNegRise) with MACD & Signal both below zero
    New in upgrade: Histogram Slope, Acceleration, Zero Line Distance,
    Momentum Expansion/Compression flags, Composite MACD Interpretation.
    """
    out = {
        # Core lines
        "macd_line":            np.nan,
        "macd_signal_line":     np.nan,
        "macd_hist":            np.nan,   # smoothed EMA(3)
        # Regime labels
        "macd_position":        "",       # Both Below Zero / Both Above Zero / Mixed
        "macd_4color":          "",       # PosRise / PosFall / NegFall / NegRise(Blue)
        "macd_wave":            "",       # Mountain (Building/Declining) / Valley (Deepening/Recovering)
        "macd_entry":           "",       # ENTRY SIGNAL / …
        "macd_cross":           "",       # Golden Cross / Dead Cross / None
        # New columns (upgrade)
        "macd_regime":          "",       # Positive Rising / Positive Falling / Negative Falling / Negative Recovering
        "macd_acceleration":    np.nan,   # delta of smoothed hist (current - previous)
        "hist_slope":           np.nan,   # 3-bar linear slope of smoothed hist
        "zero_line_dist":       np.nan,   # abs(MACD line) — distance from zero
        "momentum_expansion":   "",       # YES if |hist| increasing
        "momentum_compression": "",       # YES if |hist| decreasing toward zero
        "macd_composite":       "",       # Composite interpretation string
    }

    if hist is None or hist.empty or len(hist) < 35:
        return out

    try:
        close = hist["Close"].astype(float)

        # ── Step 1: Standard MACD (12, 26, 9) ──────────────────────────────
        ema_fast    = close.ewm(span=12, adjust=False).mean()
        ema_slow    = close.ewm(span=26, adjust=False).mean()
        macd_line   = ema_fast - ema_slow
        signal_line = macd_line.ewm(span=9, adjust=False).mean()
        hist_raw    = macd_line - signal_line

        # ── Step 2: Smooth histogram with EMA(3) — Pine: histSmooth > 1 ────
        hist_smooth = hist_raw.ewm(span=3, adjust=False).mean()

        if len(hist_smooth.dropna()) < 2:
            return out

        cur_macd  = safe_num(macd_line.iloc[-1])
        cur_sig   = safe_num(signal_line.iloc[-1])
        cur_hist  = safe_num(hist_smooth.iloc[-1])
        prev_hist = safe_num(hist_smooth.iloc[-2])
        prev_hist2 = safe_num(hist_smooth.iloc[-3]) if len(hist_smooth) >= 3 else np.nan

        if not all(pd.notna(x) for x in [cur_macd, cur_sig, cur_hist, prev_hist]):
            return out

        out["macd_line"]        = cur_macd
        out["macd_signal_line"] = cur_sig
        out["macd_hist"]        = cur_hist

        prev_macd    = safe_num(macd_line.iloc[-2]) if len(macd_line) >= 2 else np.nan
        prev_sig_val = safe_num(signal_line.iloc[-2]) if len(signal_line) >= 2 else np.nan

        # ── Lines vs zero ───────────────────────────────────────────────────
        both_below = cur_macd < 0 and cur_sig < 0
        both_above = cur_macd > 0 and cur_sig > 0
        out["macd_position"] = ("Both Below Zero" if both_below
                                else "Both Above Zero" if both_above
                                else "Mixed")

        # ── 4-Color (Pine exact) ────────────────────────────────────────────
        hist_pos_rise = cur_hist >= 0 and cur_hist >  prev_hist   # Silver
        hist_pos_fall = cur_hist >= 0 and cur_hist <= prev_hist   # Red
        hist_neg_fall = cur_hist <  0 and cur_hist <  prev_hist   # Pink
        hist_neg_rise = cur_hist <  0 and cur_hist >= prev_hist   # BLUE ★

        if hist_pos_rise:   out["macd_4color"] = "PosRise (Silver)"
        elif hist_pos_fall: out["macd_4color"] = "PosFall (Red)"
        elif hist_neg_fall: out["macd_4color"] = "NegFall (Pink)"
        else:               out["macd_4color"] = "NegRise (Blue)"

        # ── Wave pattern ────────────────────────────────────────────────────
        if hist_pos_rise:   out["macd_wave"] = "Mountain (Building)"
        elif hist_pos_fall: out["macd_wave"] = "Mountain (Declining)"
        elif hist_neg_fall: out["macd_wave"] = "Valley (Deepening)"
        else:               out["macd_wave"] = "Valley (Recovering)"

        # ── Regime label (new, matches upgrade spec) ─────────────────────────
        if hist_pos_rise:   out["macd_regime"] = "Positive Rising"
        elif hist_pos_fall: out["macd_regime"] = "Positive Falling"
        elif hist_neg_fall: out["macd_regime"] = "Negative Falling"
        else:               out["macd_regime"] = "Negative Recovering"

        # ── Entry signal ────────────────────────────────────────────────────
        if hist_neg_rise and both_below:
            out["macd_entry"] = "ENTRY SIGNAL"
        elif hist_pos_rise and both_below:
            out["macd_entry"] = "Zero Cross Up (Lines Below)"
        elif hist_pos_rise and both_above:
            out["macd_entry"] = "Bullish Continuation"
        elif hist_pos_fall and both_above:
            out["macd_entry"] = "Watch (Mountain Top)"
        elif hist_neg_fall and both_below:
            out["macd_entry"] = "Wait (Valley Deepening)"
        elif hist_neg_rise and not both_below:
            out["macd_entry"] = "NegRise (Lines Mixed)"
        else:
            out["macd_entry"] = "No Setup"

        # ── Cross detection ─────────────────────────────────────────────────
        if all(pd.notna(x) for x in [prev_macd, prev_sig_val, cur_macd, cur_sig]):
            if prev_macd <= prev_sig_val and cur_macd > cur_sig:
                out["macd_cross"] = "Golden Cross"
            elif prev_macd >= prev_sig_val and cur_macd < cur_sig:
                out["macd_cross"] = "Dead Cross"
            else:
                out["macd_cross"] = "-"
        else:
            out["macd_cross"] = "N/A"

        # ── New upgrade metrics ──────────────────────────────────────────────
        # Acceleration = change in smoothed hist (current bar vs previous)
        if pd.notna(cur_hist) and pd.notna(prev_hist):
            out["macd_acceleration"] = cur_hist - prev_hist

        # Histogram slope = 3-bar linear slope (uses prev2, prev, cur)
        if pd.notna(prev_hist2) and pd.notna(prev_hist) and pd.notna(cur_hist):
            y = [prev_hist2, prev_hist, cur_hist]
            x = [0, 1, 2]
            n = 3
            slope_num = n*sum(xi*yi for xi,yi in zip(x,y)) - sum(x)*sum(y)
            slope_den = n*sum(xi**2 for xi in x) - sum(x)**2
            out["hist_slope"] = slope_num / slope_den if slope_den != 0 else 0.0

        # Zero line distance of MACD line
        if pd.notna(cur_macd):
            out["zero_line_dist"] = abs(cur_macd)

        # Momentum expansion/compression
        if pd.notna(cur_hist) and pd.notna(prev_hist):
            expanding  = abs(cur_hist) > abs(prev_hist)
            compressing = abs(cur_hist) < abs(prev_hist)
            out["momentum_expansion"]   = "YES" if expanding  else "NO"
            out["momentum_compression"] = "YES" if compressing else "NO"

        # ── Composite MACD interpretation ────────────────────────────────────
        hist_accel = out.get("macd_acceleration", 0) or 0
        expanding  = out.get("momentum_expansion", "NO") == "YES"

        if both_above and hist_pos_rise and expanding:
            out["macd_composite"] = "Early Momentum Expansion"
        elif both_above and hist_pos_rise and not expanding:
            out["macd_composite"] = "Late Momentum Expansion"
        elif both_above and hist_pos_fall:
            out["macd_composite"] = "Bullish Compression"
        elif both_below and hist_neg_fall:
            out["macd_composite"] = "Momentum Breakdown"
        elif both_below and hist_neg_rise and not expanding:
            out["macd_composite"] = "Bearish Compression"
        elif both_below and hist_neg_rise and expanding:
            out["macd_composite"] = "Momentum Recovery"
        elif hist_pos_fall and not both_above:
            out["macd_composite"] = "Bearish Compression"
        else:
            out["macd_composite"] = out["macd_regime"]

    except Exception:
        pass

    # ── Backward-compat alias ────────────────────────────────────────────────
    out["macd_cross_status"] = out.get("macd_cross", "N/A")
    return out

def compute_stochastic(hist: pd.DataFrame, k_period: int = 14, k_smooth: int = 3, d_period: int = 3) -> dict:
    """
    Standard Stochastic Oscillator (14, 3, 3) -- same parameters as
    TradingView's built-in "Stochastic@tv-basicstudies" study, so this
    matches what the embedded chart itself would show.

    %K raw   = 100 * (Close - Lowest Low(k_period)) / (Highest High(k_period) - Lowest Low(k_period))
    %K (slow)= SMA(%K raw, k_smooth)
    %D       = SMA(%K slow, d_period)
    Golden Cross = %K crosses above %D; Dead Cross = %K crosses below %D.
    """
    out = {
        "stoch_k": np.nan,
        "stoch_d": np.nan,
        "stoch_cross": "",
    }

    min_bars = k_period + k_smooth + d_period
    if hist is None or hist.empty or len(hist) < min_bars:
        return out

    try:
        high = hist["High"].astype(float)
        low = hist["Low"].astype(float)
        close = hist["Close"].astype(float)

        lowest_low = low.rolling(window=k_period).min()
        highest_high = high.rolling(window=k_period).max()
        rng = highest_high - lowest_low
        k_raw = np.where(rng > 0, 100 * (close - lowest_low) / rng, np.nan)
        k_raw = pd.Series(k_raw, index=hist.index)

        k_slow = k_raw.rolling(window=k_smooth).mean()
        d_line = k_slow.rolling(window=d_period).mean()

        if len(k_slow.dropna()) < 2 or len(d_line.dropna()) < 2:
            return out

        cur_k, prev_k = safe_num(k_slow.iloc[-1]), safe_num(k_slow.iloc[-2])
        cur_d, prev_d = safe_num(d_line.iloc[-1]), safe_num(d_line.iloc[-2])

        out["stoch_k"] = cur_k
        out["stoch_d"] = cur_d

        if all(pd.notna(x) for x in [prev_k, prev_d, cur_k, cur_d]):
            if prev_k <= prev_d and cur_k > cur_d:
                out["stoch_cross"] = "Golden Cross"
            elif prev_k >= prev_d and cur_k < cur_d:
                out["stoch_cross"] = "Dead Cross"
            else:
                out["stoch_cross"] = "-"
        else:
            out["stoch_cross"] = "N/A"
    except Exception:
        pass

    return out

def find_pivots_low(series: pd.Series, window=2):
    idxs = []
    vals = series.values
    n = len(series)
    for i in range(window, n - window):
        w = vals[i-window:i+window+1]
        if np.isnan(vals[i]):
            continue
        if vals[i] == np.nanmin(w) and np.sum(w == vals[i]) == 1:
            idxs.append(i)
    return idxs

def find_pivots_high(series: pd.Series, window=2):
    idxs = []
    vals = series.values
    n = len(series)
    for i in range(window, n - window):
        w = vals[i-window:i+window+1]
        if np.isnan(vals[i]):
            continue
        if vals[i] == np.nanmax(w) and np.sum(w == vals[i]) == 1:
            idxs.append(i)
    return idxs

def divergence_signals(hist: pd.DataFrame, rsi_series: pd.Series, lookback=75, swing_window=2,
                       cluster_gap=6, min_separation=4, price_tol=0.0075, rsi_tol=2.0,
                       max_last_swing_age=20):
    """
    Strict lifecycle-cluster divergence logic (IDX-friendly)

    Rules:
    1) Scan recent lookback window (default 75 bars)
    2) Find RSI swing highs / lows
    3) Group nearby pivots into lifecycle clusters
    4) Bearish = last 2 RSI-high clusters where cluster max RSI > 70
    5) Bullish = last 2 RSI-low clusters where cluster min RSI < 30
    6) Representative date:
       - Bearish: highest RSI in cluster
       - Bullish: lowest RSI in cluster
    7) Compare price on those exact dates
    8) Require latest valid cluster to be recent
    """

    out = {
        "div_signal": "None",
        "div_strength": "",
        "div_ref1_date": "None",
        "div_ref1_price": np.nan,
        "div_ref1_rsi": np.nan,
        "div_ref2_date": "None",
        "div_ref2_price": np.nan,
        "div_ref2_rsi": np.nan,
        "div_rsi_event_date": "None",
        "div_price_pattern": "",
        "div_rsi_pattern": "",
        "div_pair_type": "",
        "div_reason": "",
    }

    if hist is None or hist.empty or rsi_series is None or len(hist) < 25:
        return out

    df = hist.copy()

    if not isinstance(df.index, pd.DatetimeIndex):
        try:
            if "Date" in df.columns:
                df.index = pd.to_datetime(df["Date"])
        except Exception:
            pass

    rsi_series = rsi_series.reindex(df.index)

    df = df.tail(lookback).copy()
    rsi_tail = rsi_series.tail(lookback).copy()

    if len(df) < 25:
        return out

    high = df["High"].astype(float)
    low = df["Low"].astype(float)

    def _fmt_idx(idx_pos):
        try:
            return df.index[idx_pos].strftime("%d %b '%y")
        except Exception:
            try:
                return pd.to_datetime(df.index[idx_pos]).strftime("%d %b '%y")
            except Exception:
                return ""

    def _swing_high_positions(series, w=2):
        vals = pd.Series(series).reset_index(drop=True)
        pos = []
        for i in range(w, len(vals) - w):
            c = vals.iloc[i]
            if pd.isna(c):
                continue
            left = vals.iloc[i - w:i]
            right = vals.iloc[i + 1:i + w + 1]
            if c >= left.max() and c >= right.max():
                pos.append(i)
        return pos

    def _swing_low_positions(series, w=2):
        vals = pd.Series(series).reset_index(drop=True)
        pos = []
        for i in range(w, len(vals) - w):
            c = vals.iloc[i]
            if pd.isna(c):
                continue
            left = vals.iloc[i - w:i]
            right = vals.iloc[i + 1:i + w + 1]
            if c <= left.min() and c <= right.min():
                pos.append(i)
        return pos

    def _cluster_positions(positions, max_gap=6):
        if not positions:
            return []
        clusters = [[positions[0]]]
        for p in positions[1:]:
            if p - clusters[-1][-1] <= max_gap:
                clusters[-1].append(p)
            else:
                clusters.append([p])
        return clusters

    def _price_high_relation(a, b):
        tol_abs = abs(a) * price_tol
        if b > a + tol_abs:
            return "Higher High"
        elif b < a - tol_abs:
            return "Lower High"
        return "Equal High"

    def _price_low_relation(a, b):
        tol_abs = abs(a) * price_tol
        if b < a - tol_abs:
            return "Lower Low"
        elif b > a + tol_abs:
            return "Higher Low"
        return "Equal Low"

    def _rsi_relation_high(a, b):
        if b > a + rsi_tol:
            return "Higher High"
        elif b < a - rsi_tol:
            return "Lower High"
        return "Equal High"

    def _rsi_relation_low(a, b):
        if b < a - rsi_tol:
            return "Lower Low"
        elif b > a + rsi_tol:
            return "Higher Low"
        return "Equal Low"

    candidates = []

    # Bearish: last 2 overbought lifecycle clusters
    hi_pos = _swing_high_positions(rsi_tail.values, swing_window)
    hi_clusters = _cluster_positions(hi_pos, cluster_gap)
    bear_valid = []
    for cl in hi_clusters:
        rsi_vals = rsi_tail.iloc[cl]
        if rsi_vals.dropna().empty:
            continue
        if rsi_vals.max() > 70:
            rep = cl[int(np.nanargmax(rsi_vals.values))]
            bear_valid.append((cl, rep))
    if len(bear_valid) >= 2:
        (cl1, i1), (cl2, i2) = bear_valid[-2], bear_valid[-1]
        if (i2 - i1) >= min_separation and (len(df) - 1 - i2) <= max_last_swing_age:
            p1 = safe_num(high.iloc[i1])
            p2 = safe_num(high.iloc[i2])
            r1 = safe_num(rsi_tail.iloc[i1])
            r2 = safe_num(rsi_tail.iloc[i2])
            if all(pd.notna(x) for x in [p1, p2, r1, r2]):
                price_pat = _price_high_relation(p1, p2)
                rsi_pat = _rsi_relation_high(r1, r2)
                signal, strength = None, None
                if price_pat == "Higher High" and rsi_pat == "Lower High":
                    signal, strength = "Bearish", "Strong"
                elif price_pat == "Equal High" and rsi_pat == "Lower High":
                    signal, strength = "Bearish", "Medium"
                elif price_pat == "Higher High" and rsi_pat == "Equal High":
                    signal, strength = "Bearish", "Weak"
                elif price_pat == "Lower High" and rsi_pat == "Higher High":
                    signal, strength = "Bearish", "Hidden"
                if signal:
                    candidates.append({
                        "signal": signal, "strength": strength, "pair_type": "Bearish RSI>70 Clusters",
                        "i1": i1, "i2": i2, "p1": p1, "p2": p2, "r1": r1, "r2": r2,
                        "price_pattern": price_pat, "rsi_pattern": rsi_pat, "age": len(df) - 1 - i2
                    })

    # Bullish: last 2 oversold lifecycle clusters
    lo_pos = _swing_low_positions(rsi_tail.values, swing_window)
    lo_clusters = _cluster_positions(lo_pos, cluster_gap)
    bull_valid = []
    for cl in lo_clusters:
        rsi_vals = rsi_tail.iloc[cl]
        if rsi_vals.dropna().empty:
            continue
        if rsi_vals.min() < 30:
            rep = cl[int(np.nanargmin(rsi_vals.values))]
            bull_valid.append((cl, rep))
    if len(bull_valid) >= 2:
        (cl1, i1), (cl2, i2) = bull_valid[-2], bull_valid[-1]
        if (i2 - i1) >= min_separation and (len(df) - 1 - i2) <= max_last_swing_age:
            p1 = safe_num(low.iloc[i1])
            p2 = safe_num(low.iloc[i2])
            r1 = safe_num(rsi_tail.iloc[i1])
            r2 = safe_num(rsi_tail.iloc[i2])
            if all(pd.notna(x) for x in [p1, p2, r1, r2]):
                price_pat = _price_low_relation(p1, p2)
                rsi_pat = _rsi_relation_low(r1, r2)
                signal, strength = None, None
                if price_pat == "Lower Low" and rsi_pat == "Higher Low":
                    signal, strength = "Bullish", "Strong"
                elif price_pat == "Equal Low" and rsi_pat == "Higher Low":
                    signal, strength = "Bullish", "Medium"
                elif price_pat == "Lower Low" and rsi_pat == "Equal Low":
                    signal, strength = "Bullish", "Weak"
                elif price_pat == "Higher Low" and rsi_pat == "Lower Low":
                    signal, strength = "Bullish", "Hidden"
                if signal:
                    candidates.append({
                        "signal": signal, "strength": strength, "pair_type": "Bullish RSI<30 Clusters",
                        "i1": i1, "i2": i2, "p1": p1, "p2": p2, "r1": r1, "r2": r2,
                        "price_pattern": price_pat, "rsi_pattern": rsi_pat, "age": len(df) - 1 - i2
                    })

    if not candidates:
        return out

    strength_rank = {"Strong": 4, "Medium": 3, "Weak": 2, "Hidden": 1}
    candidates = sorted(candidates, key=lambda x: (x["age"], -strength_rank.get(x["strength"], 0)))
    chosen = candidates[0]

    out["div_signal"] = chosen["signal"]
    out["div_strength"] = chosen["strength"]
    out["div_ref1_date"] = _fmt_idx(chosen["i1"])
    out["div_ref1_price"] = chosen["p1"]
    out["div_ref1_rsi"] = chosen["r1"]
    out["div_ref2_date"] = _fmt_idx(chosen["i2"])
    out["div_ref2_price"] = chosen["p2"]
    out["div_ref2_rsi"] = chosen["r2"]
    out["div_rsi_event_date"] = _fmt_idx(chosen["i2"])
    out["div_price_pattern"] = chosen["price_pattern"]
    out["div_rsi_pattern"] = chosen["rsi_pattern"]
    out["div_pair_type"] = chosen["pair_type"]
    out["div_reason"] = f'{chosen["price_pattern"]} + RSI {chosen["rsi_pattern"]} = {chosen["signal"]} {chosen["strength"]}'
    return out

def valuation_band(close, mean_val, std_val):
    if pd.isna(close) or pd.isna(mean_val) or pd.isna(std_val):
        return {"curr": np.nan, "p1": np.nan, "p2": np.nan, "m1": np.nan, "m2": np.nan, "zone": "N/A"}
    p1 = mean_val + std_val
    p2 = mean_val + (2 * std_val)
    m1 = mean_val - std_val
    m2 = mean_val - (2 * std_val)
    return {
        "curr": close,
        "p1": p1,
        "p2": p2,
        "m1": m1,
        "m2": m2,
        "zone": vwap_zone_2pct(close, [mean_val, m1, m2])
    }

# =========================
# PRODUCTION SAFETY COMPATIBILITY LAYER
# =========================

# Aesthetic group color overrides for Detail Sheet readability
FILL_GROUP_STOCK = PatternFill("solid", fgColor="2A4A66")   # navy
FILL_GROUP_MP    = PatternFill("solid", fgColor="21467C")   # denim
FILL_GROUP_MS    = PatternFill("solid", fgColor="5B3F8C")   # market structure (deep violet)
FILL_GROUP_OWNER = PatternFill("solid", fgColor="237A5C")   # forest green
FILL_GROUP_Q     = PatternFill("solid", fgColor="2369A0")   # deep blue
FILL_GROUP_PQ    = PatternFill("solid", fgColor="7A1A1A")   # dark crimson
FILL_GROUP_PY    = PatternFill("solid", fgColor="32588E")   # cobalt
FILL_GROUP_VOL   = PatternFill("solid", fgColor="216B6B")   # teal
FILL_GROUP_MA    = PatternFill("solid", fgColor="5C1530")   # dark plum
FILL_GROUP_MOM   = PatternFill("solid", fgColor="207878")   # sea green
FILL_GROUP_VAL   = PatternFill("solid", fgColor="45337C")   # indigo
FILL_GROUP_MSF    = PatternFill("solid", fgColor="7C2044")   # claret
FILL_GROUP_MACD_MOM = PatternFill("solid", fgColor="4A1060")  # deep violet-purple for MACD Momentum group
FILL_HEADER      = PatternFill("solid", fgColor="2A4A66")   # navy header
# Guarantee aliases / fallbacks for patched sections
if "FONT_NOTE" not in globals():
    FONT_NOTE = FONT_SUBTITLE if "FONT_SUBTITLE" in globals() else FONT_BODY
if "FILL_GROUP_FLOW" not in globals():
    FILL_GROUP_FLOW = FILL_GROUP_MOM if "FILL_GROUP_MOM" in globals() else None
if "FILL_HEADER" not in globals():
    FILL_HEADER = FILL_GROUP_MOM if "FILL_GROUP_MOM" in globals() else None
if "FONT_HEADER" not in globals():
    FONT_HEADER = FONT_SUBTITLE if "FONT_SUBTITLE" in globals() else FONT_BODY

# =========================
# SOURCE LOAD
# =========================
# Full KSEI schema (the columns load_ksei() builds). Kept so that when the local
# Raw folder is absent — e.g. the cloud daily runner has no local KSEI workbook —
# we can return an empty frame with the right columns instead of crashing.
KSEI_SCHEMA_COLUMNS = [
    "Ticker", "Emiten", "idx_sector", "idx_sector_weight", "Sector", "Industry",
    "Investors", "Free Float", "Classic HHI", "CR1", "CR3", "Holder", "CCS",
    "Ownership Type", "CCS Category", "Shares Outstanding", "PE TTM", "PE Mean",
    "PE +1", "PE +2", "PE -1", "PE -2", "PBV Current", "PBV Mean", "PBV +1",
    "PBV +2", "PBV -1", "PBV -2", "Source Sector",
]


def find_ksei_file():
    """Return the newest KSEI source file in the Raw folder, or None if the Raw
    folder / a KSEI file is absent (e.g. CI has no local workbook). Non-fatal:
    load_ksei() falls back to an empty schema so the committed IDX roster is
    still scanned."""
    if not os.path.isdir(RAW_DIR):
        return None
    candidates = []
    for fn in os.listdir(RAW_DIR):
        if "ksei" in fn.lower() and (fn.lower().endswith(".xlsx") or fn.lower().endswith(".csv")):
            candidates.append(os.path.join(RAW_DIR, fn))
    if not candidates:
        for fn in os.listdir(RAW_DIR):
            if fn.lower().endswith(".xlsx") or fn.lower().endswith(".csv"):
                candidates.append(os.path.join(RAW_DIR, fn))
    if not candidates:
        return None
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]

def load_ksei():
    path = find_ksei_file()
    if path is None:
        # No local Raw/KSEI workbook (the cloud daily runner has no Raw folder).
        # Return an empty frame with the full KSEI schema so the pipeline still
        # scans the committed IDX roster (merge_idx_roster) and carries blank
        # ownership/valuation — exactly like a listed ticker with no reported
        # holder table. No crash, no fabricated data; the frontend keeps the
        # last committed KSEI snapshot (exported separately from data_sources/ksei).
        print("[WARN] KSEI Raw folder/file not found -> scanning the committed IDX "
              "roster only; ownership/PBV-source columns blank (last committed KSEI kept).")
        return pd.DataFrame(columns=KSEI_SCHEMA_COLUMNS)
    df = pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)
    cols_map = {str(c).strip().lower(): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n.lower() in cols_map:
                return cols_map[n.lower()]
        return None

    col_ticker = pick("Ticker", "Kode", "kode")
    if not col_ticker:
        raise ValueError("Ticker/Kode column not found in KSEI file.")

    out = pd.DataFrame()
    out["Ticker"] = df[col_ticker].astype(str).str.upper().str.strip()

    mappings = {
        "Emiten": pick("Emiten"),
        "idx_sector": pick("IDX Sector", "IDXSector", "IDX_Sector"),
        "idx_sector_weight": pick("IDX Sector Weight", "IDXSectorWeight", "IDX_Sector_Weight", "Weight"),
        "Sector": pick("Sector", "Sektor"),
        "Industry": pick("Industry", "Industri"),
        "Investors": pick("Investors"),
        "Free Float": pick("Free Float"),
        "Classic HHI": pick("Classic HHI"),
        "CR1": pick("Concentration Ratio Top 1 (CR1)", "CR1"),
        "CR3": pick("Concentration Ratio Top 3 (CR3)", "CR3"),
        "Holder": pick("Total Investor >1%", "Holder"),
        "CCS": pick("Composite Concentration Score (CCS)", "CCS"),
        "Ownership Type": pick("Ownership Type"),
        "CCS Category": pick("CCS Category"),
        "Shares Outstanding": pick("Shares Outstanding", "SharesOutstanding", "Listed Shares", "Shares", "Shares Out", "Jumlah Saham Beredar"),
        "PE TTM": pick("Current PE Ratio (TTM)", "PE TTM", "PE Ratio TTM", "Current PE"),
        "PE Mean": pick("Mean PE Standard Deviation", "PE Mean", "Mean PE"),
        "PE +1": pick("+1 PE Standard Deviation", "PE +1"),
        "PE +2": pick("+2 PE Standard Deviation", "PE +2"),
        "PE -1": pick("-1 PE Standard Deviation", "PE -1"),
        "PE -2": pick("-2 PE Standard Deviation", "PE -2"),
        "PBV Current": pick("Current Price to Book Value", "Current PBV", "PBV Current"),
        "PBV Mean": pick("Mean PBV Standard Deviation", "PBV Mean", "Mean PBV"),
        "PBV +1": pick("+1 PBV Standard Deviation", "PBV +1"),
        "PBV +2": pick("+2 PBV Standard Deviation", "PBV +2"),
        "PBV -1": pick("-1 PBV Standard Deviation", "PBV -1"),
        "PBV -2": pick("-2 PBV Standard Deviation", "PBV -2"),
    }

    for k, src in mappings.items():
        out[k] = df[src] if src else np.nan

    # Keep the descriptive source sector for audit, but expose only the
    # authoritative 11-sector IDX classification plus Others downstream.
    out["Source Sector"] = out["Sector"]
    out["idx_sector"] = [
        normalize_idx_sector(idx_sector, source_sector)
        for idx_sector, source_sector in zip(out["idx_sector"], out["Source Sector"])
    ]
    out["Sector"] = out["idx_sector"]

    # Preserve extra Raw/Stockbit columns so downstream sheets can prefer
    # source-of-truth Stockbit values over yfinance approximations when present.
    for _c in df.columns:
        _name = str(_c).strip()
        if _name and _name not in out.columns:
            out[_name] = df[_c]

    out = out[out["Ticker"].notna() & (out["Ticker"] != "")]
    out = out.drop_duplicates(subset=["Ticker"], keep="first").reset_index(drop=True)
    return out

# =========================
# MARKET DATA
# =========================
def fetch_history_once(symbol: str, params: dict):
    # Enforce MARKET_DATE as hard cutoff on all ticker downloads
    _mdate_dl = get_market_date()
    _end_dl   = (_mdate_dl + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    if "end" not in params:
        params["end"] = _end_dl
    hist = yf.download(symbol, auto_adjust=False, progress=False, threads=False, **params)
    if isinstance(hist.columns, pd.MultiIndex):
        hist.columns = hist.columns.get_level_values(0)
    return normalize_history(hist)

def fetch_history_with_retry(ticker: str):
    symbol = f"{ticker}.JK"
    attempts = [
        {"start": "1990-01-01", "interval": "1d"},
        {"start": "2020-01-01", "interval": "1d"},
        {"period": "2y", "interval": "1d"},
    ]
    retry_count = 0
    last_err = None
    for params in attempts:
        for i in range(FETCH_RETRIES):
            try:
                hist = fetch_history_once(symbol, params)
                if hist is not None and not hist.empty:
                    return hist, retry_count, "Yahoo", ""
                retry_count += 1
                time.sleep(RETRY_SLEEP_SEC * (i + 1))
            except Exception as e:
                last_err = f"{type(e).__name__}: {e}"
                retry_count += 1
                time.sleep(RETRY_SLEEP_SEC * (i + 1))
    reason = "Empty dataset after retries"
    if last_err:
        reason = f"Fetch error after retries | {last_err}"
    return None, retry_count, "None", reason

def fetch_shares_outstanding(ticker: str, cache: dict) -> Optional[float]:
    ticker_u = str(ticker).upper()
    cached = cache.get(ticker_u)
    try:
        if cached is not None and pd.notna(cached) and float(cached) > 0:
            return float(cached)
    except Exception:
        pass
    symbol = f"{ticker}.JK"
    try:
        tk = yf.Ticker(symbol)
        shares = np.nan
        try:
            shares = (tk.fast_info or {}).get("shares", np.nan)
        except Exception:
            shares = np.nan
        if pd.isna(shares):
            try:
                shares = (tk.info or {}).get("sharesOutstanding", np.nan)
            except Exception:
                shares = np.nan
        if pd.notna(shares) and float(shares) > 0:
            cache[ticker_u] = float(shares)
            return float(shares)
    except Exception:
        pass
    return np.nan

# =========================
# STOCK REGIME
# =========================
def classify_market_cap_category(market_cap: float) -> str:
    if pd.isna(market_cap):
        return "N/A"
    if market_cap > 10_000_000_000_000:
        return "Large Cap"
    if market_cap >= 1_000_000_000_000:
        return "Mid Cap"
    if market_cap >= 100_000_000_000:
        return "Small Cap"
    return "Micro Cap"


def classify_liquidity_category(avg_daily_value: float) -> str:
    if pd.isna(avg_daily_value):
        return "N/A"
    if avg_daily_value > 50_000_000_000:
        return "High Liquidity"
    if avg_daily_value >= 10_000_000_000:
        return "Medium Liquidity"
    if avg_daily_value >= 1_000_000_000:
        return "Low Liquidity"
    return "Very Low Liquidity"



# =========================
# STOCK REGIME / VERDICT WEIGHT PROFILE
# =========================
def classify_stock_regime(row: Dict[str, Any]) -> Dict[str, str]:
    """
    Four institutional categories (Upgrade Spec Item 7):
    1. Blue Chip / High Liquidity   = Large Cap + High Liquidity
    2. Mid Cap / Moderate           = Mid Cap + Medium Liquidity
    3. Low Liquidity / Small Cap    = Small/Micro Cap + Low Liquidity
    4. Institutional Driven         = High sponsorship + accumulation +
                                       strong turnover velocity
                                       → OVERRIDES market cap if conditions met

    Institutional Driven threshold rules:
    - volume_sponsorship_score  >= 65  (high volume sponsorship)
    - accumulation_score        >= 60  (positive accumulation score)
    - turnover_velocity_20d     >= 0.05 (healthy turnover velocity ≥ 0.05% daily)
    - smart_money_bias          in (Bullish, Strongly Bullish)
    - flow_conviction           not in (Negative, Strongly Negative)
    """
    market_cap      = safe_num(row.get("market_cap"))
    avg_daily_value = safe_num(row.get("avg_value_30d"))
    classic_hhi     = safe_num(row.get("classic_hhi"))
    cr1             = safe_num(row.get("cr1"))
    ccs             = safe_num(row.get("ccs"))

    # Sponsorship / flow override fields (populated by compute_institutional_metrics)
    vol_spon   = safe_num(row.get("volume_sponsorship_score"))
    accum      = safe_num(row.get("accumulation_score"))
    turnover   = safe_num(row.get("turnover_velocity_20d"))
    sm_bias    = str(row.get("smart_money_bias", "") or "")
    flow_conv  = str(row.get("flow_conviction", "") or "")

    mc_cat  = classify_market_cap_category(market_cap)
    liq_cat = classify_liquidity_category(avg_daily_value)

    # ── Institutional Driven override check ─────────────────────────────────
    is_inst_driven = (
        (pd.notna(vol_spon)  and vol_spon  >= 65) and
        (pd.notna(accum)     and accum     >= 60) and
        (pd.notna(turnover)  and turnover  >= 0.05) and
        sm_bias   in ("Bullish", "Strongly Bullish") and
        flow_conv not in ("Negative", "Strongly Negative")
    )
    # Fallback: KSEI-derived HHI / CR1 / CCS signals if flow data absent
    ksei_inst = (
        (pd.notna(classic_hhi) and classic_hhi >= 1500) or
        (pd.notna(cr1)         and cr1         >= 0.30) or
        (pd.notna(ccs)         and ccs         >= 70)
    )

    if is_inst_driven or (ksei_inst and liq_cat in ("High Liquidity", "Medium Liquidity")):
        regime = "Institutional Driven"
    elif mc_cat == "Large Cap" and liq_cat in ("High Liquidity",):
        regime = "Blue Chip / High Liquidity"
    elif mc_cat == "Mid Cap":
        regime = "Mid Cap / Moderate"
    elif mc_cat in ("Small Cap", "Micro Cap") and liq_cat in ("Low Liquidity", "Very Low Liquidity"):
        regime = "Low Liquidity / Small Cap"
    else:
        # Catch-all: base on market cap category
        if mc_cat == "Large Cap":
            regime = "Blue Chip / High Liquidity"
        elif mc_cat == "Mid Cap":
            regime = "Mid Cap / Moderate"
        else:
            regime = "Low Liquidity / Small Cap"

    return {
        "market_cap_category": mc_cat,
        "liquidity_category":  liq_cat,
        "stock_regime":        regime,
    }

# =========================
# ROW BUILDERS
# =========================
def base_row_from_ksei(ksei_row: pd.Series):
    row = {
        "ticker": ksei_row["Ticker"],
        "close": np.nan,
        "last_high": np.nan, "last_low": np.nan, "last_open": np.nan,
        "pct_change": np.nan,
        "beta_ihsg":  np.nan,
        "beta_ihsg_zone": "-",
        "rs_rating":      np.nan,
        "rs_rating_zone": "-",
        "today_event":    "-",
        # Upgrade 1 — ARA/ARB
        "ara_arb":    "-",
        "at_limit":   False,
        # Upgrade 3 — Board classification
        "board":      "Main Board",
        # Upgrade 5 — Suspension
        "suspended":           False,
        "suspension_label":    "-",
        "corp_action_active":  "-",
        # Upgrade 2 — Foreign flow
        "foreign_net_lot":  np.nan,
        "foreign_net_val":  np.nan,
        "foreign_activity": "-",
        # Upgrade 4 — Broker flow
        "top_broker_buy":    "-",
        "top_broker_sell":   "-",
        "broker_net_signal": "-",
        # Upgrade 6 — Position sizing
        "max_shares":            np.nan,
        "max_lots":              np.nan,
        "max_position_idr":      np.nan,
        "position_size_pct_adtv": np.nan,
        # Upgrade 7 — Seasonal bias
        "seasonal_bias":  "-",
        # Upgrade 8 — Dividend trap (populated in fundamental row)
        "div_trap_flag":  "-",
        # Upgrade 9 — OB Age
        "smc_ob_bull_age": np.nan,
        "smc_ob_bear_age": np.nan,
        # Upgrade 10 — RRG quadrant
        "rrg_quadrant": "-",
        "emiten": ksei_row.get("Emiten", ""),
        "idx_sector": ksei_row.get("idx_sector", ""),
        "idx_sector_weight": safe_num(ksei_row.get("idx_sector_weight")),
        "sector": ksei_row.get("Sector", ""),
        "industry": ksei_row.get("Industry", ""),
        "mcap": np.nan,
        "market_cap_category": "N/A",
        "liquidity_category": "N/A",
        "stock_regime": "N/A",
        "data_status": "NO DATA",
        "bars": 0,
        "investors": ksei_row.get("Investors", ""),
        "free_float": safe_num(ksei_row.get("Free Float")),
        "hhi": safe_num(ksei_row.get("Classic HHI")),
        "cr1": safe_num(ksei_row.get("CR1")),
        "cr3": safe_num(ksei_row.get("CR3")),
        "holder": safe_int(ksei_row.get("Holder")),
        "ccs": safe_num(ksei_row.get("CCS")),
        "ownership_type": ksei_row.get("Ownership Type", ""),
        "ccs_category": ksei_row.get("CCS Category", ""),
        "cm_days": np.nan, "cm_vwap": np.nan, "cm_p1": np.nan, "cm_p2": np.nan, "cm_p3": np.nan, "cm_m1": np.nan, "cm_m2": np.nan, "cm_m3": np.nan, "cm_sd": np.nan, "cm_delta": np.nan, "cm_zone": "N/A", "cm_remarks": "",
        "pm_days": np.nan, "pm_vwap": np.nan, "pm_p1": np.nan, "pm_p2": np.nan, "pm_p3": np.nan, "pm_m1": np.nan, "pm_m2": np.nan, "pm_m3": np.nan, "pm_sd": np.nan, "pm_delta": np.nan, "pm_zone": "N/A", "pm_remarks": "",
        "q_days": np.nan, "q_vwap": np.nan, "q_p1": np.nan, "q_p2": np.nan, "q_p3": np.nan, "q_m1": np.nan, "q_m2": np.nan, "q_m3": np.nan, "q_sd": np.nan, "q_delta": np.nan, "q_zone": "N/A", "q_remarks": "",
        "pq_days": np.nan, "pq_vwap": np.nan, "pq_p1": np.nan, "pq_p2": np.nan, "pq_p3": np.nan, "pq_m1": np.nan, "pq_m2": np.nan, "pq_m3": np.nan, "pq_sd": np.nan, "pq_delta": np.nan, "pq_zone": "N/A", "pq_remarks": "",
        "py_year": np.nan, "py_vwap": np.nan, "py_p1": np.nan, "py_p2": np.nan, "py_p3": np.nan, "py_m1": np.nan, "py_m2": np.nan, "py_m3": np.nan, "py_sd": np.nan, "py_delta": np.nan, "py_zone": "N/A", "py_remarks": "",
        "ibh": np.nan, "ibl": np.nan, "mp_zone": "N/A", "pwh": np.nan, "pwl": np.nan, "price_ge_pwl": "N/A", "mdh": np.nan, "mdl": np.nan, "price_ge_mdl": "N/A",
        "avg30": np.nan, "last_vol": np.nan, "last_ok": "N/A", "rvol5": np.nan, "rvol_zone": "N/A",
        "lot": np.nan, "daily_value": np.nan, "adtv20": np.nan, "adtv20_pct": np.nan, "adtv20_zone": "N/A",
        "adtr20": np.nan, "adtr20_pct": np.nan, "adtr20_zone": "N/A", "rvol20": np.nan, "rvol20_zone": "N/A",
        "rvol20_chg_pct": np.nan, "rvol20_chg_flag": "-",
        "adr_pct": np.nan, "atr14_pct": np.nan, "adr_atr_zone": "N/A",
        "summary_ma": "N/A",
        "ema10": np.nan, "ema10d": np.nan, "ema10p": "N/A",
        "ema20": np.nan, "ema20d": np.nan, "ema20p": "N/A",
        "ema25": np.nan, "ema25d": np.nan, "ema25p": "N/A",
        "ema50": np.nan, "ema50d": np.nan, "ema50p": "N/A",
        "sma200": np.nan, "sma200d": np.nan, "sma200p": "N/A",
        "rsi14": np.nan, "rsi_ge_50": "N/A", "rsi_delta": np.nan, "rsi_status": "N/A",
        "rsi_ma14": np.nan, "golden_cross": "N/A", "dead_cross": "N/A", "cross_status": "-", "macd_cross_status": "-", "cm_zone_days": 0, "pm_zone_days": 0, "q_zone_days": 0, "pq_zone_days": 0, "py_zone_days": 0,
        "adr_pct": np.nan, "atr_pct": np.nan, "vol_range_zone": "N/A",
        "div_signal": "N/A", "div_strength": "None", "div_ref1_date": "", "div_ref1_price": np.nan, "div_ref1_rsi": np.nan, "div_ref2_date": "", "div_ref2_price": np.nan, "div_ref2_rsi": np.nan, "div_price_pattern": "", "div_rsi_pattern": "", "div_reason": "", "div_pair_type": "", "div_rsi_event_date": "", "pe_band_source": "N/A", "pbv_band_source": "N/A",
        "pe_ttm": safe_num(ksei_row.get("PE TTM")), "pe_mean": safe_num(ksei_row.get("PE Mean")), "pe_m1": safe_num(ksei_row.get("PE -1")), "pe_m2": safe_num(ksei_row.get("PE -2")), "pe_zone": "N/A",
        "pbv_curr": safe_num(ksei_row.get("PBV Current")), "pbv_mean": safe_num(ksei_row.get("PBV Mean")), "pbv_p1": safe_num(ksei_row.get("PBV +1")), "pbv_p2": safe_num(ksei_row.get("PBV +2")), "pbv_m1": safe_num(ksei_row.get("PBV -1")), "pbv_m2": safe_num(ksei_row.get("PBV -2")), "pbv_zone": "N/A",
        "flow_data_mode": "PROXY", "dollar_vol_20d_avg": np.nan, "turnover_velocity_20d": np.nan, "ad_trend": "Flat", "obv_trend": "Flat", "cmf20": np.nan, "volume_sponsorship_score": np.nan, "accumulation_score": np.nan, "distribution_score": np.nan, "smart_money_bias": "Neutral", "flow_conviction": "Neutral", "net_participation_proxy": np.nan, "sponsorship_grade": "N/A",
        "pos_52w_pct": np.nan, "range_compression_20d": np.nan, "vol_compression_20d": np.nan, "base_length": np.nan, "breakout_pressure": np.nan, "breakdown_pressure": np.nan, "_removed_spring": "NO", "_removed_upthrust": "NO", "structure_state": "N/A", "wyckoff_proxy_phase": "Transitional", "markup_readiness": np.nan, "breakdown_risk": np.nan, "cause_quality": "Weak",
        "technical_core_score": np.nan, "flow_score": np.nan, "structure_score": np.nan, "risk_penalty": np.nan, "regime_multiplier": np.nan, "adaptive_composite_score": np.nan, "setup_quality": "N/A", "trap_risk": "N/A", "institutional_verdict": "N/A", "institutional_action_bias": "N/A",
        "institutional_accumulation": "NO", "early_markup_candidate": "NO", "smart_money_pullback": "NO", "breakout_watchlist": "NO", "distribution_warning": "NO", "failed_breakout_risk": "NO", "capital_efficient_trend": "NO", "speculative_momentum": "NO", "preset_summary": "None",
        "tier": "",
        "latest_market_day": None,
    }
    # Carry through Stockbit/raw fundamental columns for later source overrides.
    _stockbit_source_cols = [
        "Current PE Ratio (Annualised)", "Current PE Ratio (TTM)", "Earnings Yield (TTM)",
        "Current Price to Sales (TTM)", "Current Price to Book Value", "EV to EBIT (TTM)",
        "EV to EBITDA (TTM)", "Market Cap", "Enterprise Value", "Current Share Outstanding",
        "Gross Profit Margin (Quarter)", "Operating Profit Margin (Quarter)", "Net Profit Margin (Quarter)",
        "Return on Assets (TTM)", "Return on Equity (TTM)", "Return on Capital Employed (TTM)",
        "Return On Invested Capital (TTM)", "Days Sales Outstanding (Quarter)", "Asset Turnover (TTM)",
        "Days Inventory (Quarter)", "Days Payables Outstanding (Quarter)", "Cash Conversion Cycle (Quarter)",
        "Receivables Turnover (Quarter)", "Current Ratio (Quarter)", "Quick Ratio (Quarter)",
        "Debt to Equity Ratio (Quarter)", "LT Debt/Equity (Quarter)", "Total Liabilities/Equity (Quarter)",
        "Financial Leverage (Quarter)", "Interest Coverage (TTM)", "Free cash flow (Quarter)",
        "Altman Z-Score (Original)", "Altman Z-Score (Modified)", "1 Month Price Returns",
        "3 Month Price Returns", "6 Month Price Returns", "1 Year Price Returns", "3 Year Price Returns",
        "Year to Date Price Returns", "52 Week High", "52 Week Low", "Book Value (Quarter)",
        "Current Book Value Per Share", "Tang. Book Value (Quarter)",
        "Current Tang. Book Value Per Share", "Short-term Debt (Quarter)", "Long-term Debt (Quarter)",
        "Cash (Quarter)", "Total Assets (Quarter)", "Total Liabilities (Quarter)", "Current EPS (TTM)",
        "EPS (Quarter YoY Growth)", "Revenue (TTM)", "Revenue (Quarter YoY Growth)", "Net Income (TTM)",
        "EBIT (TTM)", "Cash From Operations (TTM)", "Cash From Investing (TTM)",
        "Cash From Financing (TTM)", "Capital expenditure (TTM)", "Free cash flow (TTM)",
    ]
    for _c in _stockbit_source_cols:
        if _c in ksei_row:
            row[_c] = ksei_row.get(_c)
    return row

def build_row(ksei_row: pd.Series, hist: pd.DataFrame, shares_fallback: float):
    # Historical Safety Layer: clip to asof in backtest mode
    hist = _clip_hist_to_asof(hist)
    row = base_row_from_ksei(ksei_row)
    row["bars"] = len(hist)
    row["latest_market_day"] = hist.index[-1].strftime("%Y-%m-%d")
    row["close"]     = safe_num(hist["Close"].iloc[-1])
    row["last_high"] = safe_num(hist["High"].iloc[-1])  if "High"  in hist.columns else np.nan
    row["last_low"]  = safe_num(hist["Low"].iloc[-1])   if "Low"   in hist.columns else np.nan
    row["last_open"] = safe_num(hist["Open"].iloc[-1])  if "Open"  in hist.columns else np.nan
    if len(hist) >= 2:
        _prev_close = safe_num(hist["Close"].iloc[-2], np.nan)
        # Store as decimal fraction (e.g. 0.0023 for 0.23%) so Excel "0.00%" format
        # multiplies by 100 and displays correctly. safe_pct returns *100 which causes
        # double-multiplication (0.23 → shown as 23%).
        row["pct_change"] = safe_div(row["close"] - _prev_close, _prev_close, np.nan)
    else:
        row["pct_change"] = np.nan

    # Beta vs IHSG: COVARIANCE.P(stock_returns, ihsg_returns) / VAR.P(ihsg_returns)
    row["beta_ihsg"] = compute_beta_ihsg(hist) if len(hist) >= 30 else np.nan
    row["beta_ihsg_zone"] = compute_beta_zone(row["beta_ihsg"])

    # RS Rating (IBD-style 1–99) — universe cache populated across all tickers
    rs = compute_rs_rating(hist)
    row["rs_rating"]      = rs
    row["rs_rating_zone"] = rs_rating_zone(rs)
    # Store raw score for universe re-scoring pass after all rows built
    try:
        closes = hist["Close"].squeeze().dropna()
        n = len(closes)
        def _r(b):
            base = float(closes.iloc[-b-1]) if n > b and closes.iloc[-b-1] != 0 else None
            return (float(closes.iloc[-1]) / base - 1.0) if base else np.nan
        _raw = 0.40*_r(63) + 0.20*_r(min(126,n-1)) + 0.20*_r(min(189,n-1)) + 0.20*_r(min(252,n-1))
        row["_rs_raw"] = _raw
        if len(_RS_UNIVERSE_CACHE) < 10000:
            _RS_UNIVERSE_CACHE.append(float(_raw)) if pd.notna(_raw) else None
    except Exception:
        row["_rs_raw"] = np.nan

    # Resolve ticker string for all upgrade API calls
    _ticker = str(row.get("ticker", "")).upper().strip()

    # ── RVOL Change % and flag ─────────────────────────────────────────────────
    # RVOL Change % = (RVOL 20D today - RVOL 20D yesterday) / RVOL 20D yesterday * 100
    # Requires at least 2 bars of volume history beyond the 20D window
    _rvol_chg = np.nan
    _rvol_chg_flag = "-"
    try:
        if len(hist) >= 22 and "Volume" in hist.columns:
            vol_s = hist["Volume"].astype(float)
            adtv_today = safe_num(vol_s.iloc[-21:-1].mean())   # 20D avg ending yesterday
            adtv_prev  = safe_num(vol_s.iloc[-22:-2].mean())   # 20D avg ending day before
            vol_today  = safe_num(vol_s.iloc[-1])
            vol_prev   = safe_num(vol_s.iloc[-2])
            rvol_today = vol_today / adtv_today if pd.notna(adtv_today) and adtv_today > 0 else np.nan
            rvol_prev  = vol_prev  / adtv_prev  if pd.notna(adtv_prev)  and adtv_prev  > 0 else np.nan
            if pd.notna(rvol_today) and pd.notna(rvol_prev) and rvol_prev > 0:
                _rvol_chg = ((rvol_today - rvol_prev) / rvol_prev) * 100.0
                _rvol_chg_flag = (
                    "RVOL ↑↑" if _rvol_chg >= 50.0 else
                    "RVOL ↑"  if _rvol_chg >= 0 else
                    "RVOL ↓"  if _rvol_chg >= -50.0 else
                    "RVOL ↓↓")
    except Exception:
        pass
    row["rvol20_chg_pct"]  = safe_num(_rvol_chg / 100.0) if pd.notna(_rvol_chg) else np.nan
    row["rvol20_chg_flag"] = _rvol_chg_flag

    # ── Upgrade 1: ARA/ARB ────────────────────────────────────────────────────
    # Source-limited mode does not call IDX participant/transaction endpoints.
    board = "N/A"
    row["board"] = board
    ara = compute_ara_arb(hist, board)
    row["ara_arb"]  = ara["ara_arb"]
    row["at_limit"] = ara["at_limit"]

    # ── Upgrade 2: Foreign flow ───────────────────────────────────────────────
    ff = {"foreign_net_lot": np.nan, "foreign_net_val": np.nan, "foreign_activity": "-"}
    row["foreign_net_lot"]  = ff["foreign_net_lot"]
    row["foreign_net_val"]  = ff["foreign_net_val"]
    row["foreign_activity"] = ff["foreign_activity"]

    # ── Upgrade 4: Broker flow ────────────────────────────────────────────────
    bf = {"top_broker_buy": "-", "top_broker_sell": "-", "broker_net_signal": "-"}
    row["top_broker_buy"]    = bf["top_broker_buy"]
    row["top_broker_sell"]   = bf["top_broker_sell"]
    row["broker_net_signal"] = bf["broker_net_signal"]

    # ── Upgrade 5: Suspension flag ────────────────────────────────────────────
    susp = {"suspended": False, "suspension_label": "-", "corp_action_active": "-"}
    row["suspended"]          = susp["suspended"]
    row["suspension_label"]   = susp["suspension_label"]
    row["corp_action_active"] = susp["corp_action_active"]

    # ── Upgrade 6: Position sizing ────────────────────────────────────────────
    ps = compute_max_position_size(safe_num(row.get("adtr20"), np.nan), safe_num(row.get("close"), np.nan))
    row["max_shares"]             = ps["max_shares"]
    row["max_lots"]               = ps["max_lots"]
    row["max_position_idr"]       = ps["max_position_idr"]
    row["position_size_pct_adtv"] = ps["position_size_pct_adtv"]

    # ── Upgrade 7: Seasonal bias ──────────────────────────────────────────────
    from datetime import date as _date
    row["seasonal_bias"] = get_seasonal_bias(get_market_date().month)

    # ── Upgrade 9: OB age ─────────────────────────────────────────────────────
    ob_eq   = safe_num(row.get("smc_ob_equilibrium"), np.nan)
    ob_bear_str = str(row.get("smc_closest_ob_bear", "") or "")
    def _mid_from_range_local(s):
        try:
            pts = s.replace("–","-").split("-")
            if len(pts) == 2:
                return (float(pts[0].replace(",","")) + float(pts[1].replace(",","")))/2
        except Exception:
            pass
        return np.nan
    ob_bear_mid = _mid_from_range_local(ob_bear_str)
    # Keep exact OB ages returned by the SMC engine; only fall back to fuzzy
    # midpoint matching for legacy rows where the engine did not provide ages.
    if pd.isna(row.get("smc_ob_bull_age", np.nan)):
        row["smc_ob_bull_age"] = compute_ob_age(hist, ob_eq)
    if pd.isna(row.get("smc_ob_bear_age", np.nan)):
        row["smc_ob_bear_age"] = compute_ob_age(hist, ob_bear_mid)

    # Today Event — SMC structural event on latest candle
    row["today_event"] = compute_today_event(row, hist)

    # ── Upgrade 10: RRG quadrant (populated from cache set by IDX Overview) ──
    sector = str(row.get("idx_sector") or row.get("sector") or "")
    row["rrg_quadrant"] = get_rrg_quadrant_for_sector(sector)

    shares_ksei = safe_num(ksei_row.get("Shares Outstanding"))
    shares_used = shares_ksei if pd.notna(shares_ksei) and shares_ksei > 0 else shares_fallback
    row["mcap"] = row["close"] * shares_used if pd.notna(row["close"]) and pd.notna(shares_used) and shares_used > 0 else np.nan

    # FULL mode
    if len(hist) >= MIN_BARS_FULL:
        row["data_status"] = "OK"

        cq_start = quarter_start(hist.index[-1])
        df_q = hist.loc[hist.index >= cq_start]
        q = anchored_vwap_block(df_q)

        pq_start, pq_end = prev_quarter_range(hist.index[-1])
        df_pq = hist.loc[(hist.index >= pq_start) & (hist.index <= pq_end)]
        _selected_close = safe_num(hist["Close"].iloc[-1], np.nan)
        _previous_close = safe_num(hist["Close"].iloc[-2], np.nan) if len(hist) >= 2 else np.nan
        pq = anchored_vwap_block(df_pq, _selected_close, _previous_close)

        prev_year = hist.index[-1].year - 1
        df_py = hist.loc[(hist.index.year == prev_year)]
        py = anchored_vwap_block(df_py, _selected_close, _previous_close)

        # ── Monthly VWAP (Current + Previous) ─────────────────────────────
        _last_ts = hist.index[-1]
        cm_start = pd.Timestamp(_last_ts.year, _last_ts.month, 1)
        df_cm = hist.loc[hist.index >= cm_start]
        cm = anchored_vwap_block(df_cm)

        _pm_end   = cm_start - pd.Timedelta(days=1)
        _pm_start = pd.Timestamp(_pm_end.year, _pm_end.month, 1)
        df_pm = hist.loc[(hist.index >= _pm_start) & (hist.index <= _pm_end)]
        pm = anchored_vwap_block(df_pm, _selected_close, _previous_close)
        # ───────────────────────────────────────────────────────────────────

        month_start = pd.Timestamp(hist.index[-1].year, hist.index[-1].month, 1)
        df_month = hist.loc[hist.index >= month_start]
        first2 = df_month.head(2)
        row["ibh"] = safe_num(first2["High"].max()) if not first2.empty else np.nan
        row["ibl"] = safe_num(first2["Low"].min()) if not first2.empty else np.nan
        row["mp_zone"] = zone_label(row["close"], row["ibl"], row["ibh"])
        if pd.notna(row["ibh"]) and pd.notna(row["ibl"]) and abs(float(row["ibh"]) - float(row["ibl"])) < 1e-9:
            row["ibh"], row["ibl"], row["mp_zone"] = np.nan, np.nan, "N/A"

        # Weekly Market Profile additions
        # PWH / PWL = range of last completed week (Mon–Fri) by default.
        #             Exception: if script runs on Saturday or Sunday (WIB),
        #             treat the current ISO week as the completed reference week
        #             (the week that just ended Friday is still "current" in data).
        # MDH / MDL = first-day high/low of the current running week.
        try:
            iso = hist.index.to_series().dt.isocalendar()
            cur_year = int(iso.year.iloc[-1])
            cur_week = int(iso.week.iloc[-1])

            # Detect weekend run (WIB = UTC+7)
            _wib_now = pd.Timestamp.now(tz="Asia/Jakarta")
            _is_weekend_run = _wib_now.weekday() >= 5  # 5=Sat, 6=Sun

            # Current running week (always compute for MDH/MDL)
            cur_mask = (iso.year == cur_year) & (iso.week == cur_week)
            df_cur_week = hist.loc[cur_mask]
            if not df_cur_week.empty:
                first_cur = df_cur_week.iloc[0]
                row["mdh"] = safe_num(first_cur["High"])
                row["mdl"] = safe_num(first_cur["Low"])
                if pd.notna(row["mdh"]) and pd.notna(row["mdl"]) and abs(float(row["mdh"]) - float(row["mdl"])) < 1e-9:
                    row["mdh"], row["mdl"], row["price_ge_mdl"] = np.nan, np.nan, "N/A"
                else:
                    row["price_ge_mdl"] = ("At Level" if pd.notna(row["close"]) and pd.notna(row["mdl"]) and abs(float(row["close"]) - float(row["mdl"])) < 1e-9 else ("Above" if pd.notna(row["close"]) and pd.notna(row["mdl"]) and row["close"] > row["mdl"] else ("Below" if pd.notna(row["close"]) and pd.notna(row["mdl"]) else "N/A")))

            if _is_weekend_run:
                # Weekend: PWH/PWL = current week's full Mon–Fri range
                df_pw = df_cur_week
            else:
                # Weekday: PWH/PWL = prior completed week's full range
                prior_mask = (iso.year < cur_year) | ((iso.year == cur_year) & (iso.week < cur_week))
                df_pw = pd.DataFrame()
                if prior_mask.any():
                    prior_idx = hist.index[prior_mask]
                    last_prior_iso = pd.Timestamp(prior_idx[-1]).isocalendar()
                    last_year = int(last_prior_iso.year)
                    last_week = int(last_prior_iso.week)
                    last_mask = (iso.year == last_year) & (iso.week == last_week)
                    df_pw = hist.loc[last_mask]

            if not df_pw.empty:
                row["pwh"] = safe_num(df_pw["High"].max())
                row["pwl"] = safe_num(df_pw["Low"].min())
                if pd.notna(row["pwh"]) and pd.notna(row["pwl"]) and abs(float(row["pwh"]) - float(row["pwl"])) < 1e-9:
                    row["pwh"], row["pwl"], row["price_ge_pwl"] = np.nan, np.nan, "N/A"
                else:
                    row["price_ge_pwl"] = ("At Level" if pd.notna(row["close"]) and pd.notna(row["pwl"]) and abs(float(row["close"]) - float(row["pwl"])) < 1e-9 else ("Above" if pd.notna(row["close"]) and pd.notna(row["pwl"]) and row["close"] > row["pwl"] else ("Below" if pd.notna(row["close"]) and pd.notna(row["pwl"]) else "N/A")))
        except Exception:
            pass

        vol_series = hist["Volume"].copy()
        row["avg30"] = safe_num(vol_series.iloc[-31:-1].mean()) if len(vol_series) >= 31 else safe_num(vol_series.tail(30).mean())
        row["last_vol"] = safe_num(vol_series.iloc[-1])
        row["last_ok"] = "YES" if pd.notna(row["last_vol"]) and row["last_vol"] >= VOL_THRESHOLD else "NO"
        avg5 = safe_num(vol_series.iloc[-6:-1].mean()) if len(vol_series) >= 6 else safe_num(vol_series.tail(5).mean())
        row["rvol5"] = row["last_vol"] / avg5 if pd.notna(avg5) and avg5 != 0 else np.nan
        row["rvol_zone"] = "YES" if pd.notna(row["rvol5"]) and row["rvol5"] >= MIN_RVOL else "NO"

        # Liquidity block (institutional)
        row["lot"] = row["last_vol"] / 100 if pd.notna(row["last_vol"]) else np.nan
        # Value (Approx): Typical Price = (Open + High + Low + Close) / 4 × Volume
        row["daily_value"] = (safe_num(hist["Volume"].iloc[-1] * ((hist["Open"].iloc[-1] + hist["High"].iloc[-1] + hist["Low"].iloc[-1] + hist["Close"].iloc[-1]) / 4.0)) if len(hist) and all(col in hist.columns for col in ["Volume","Open","High","Low","Close"]) else np.nan)

        adtv_base = safe_num(vol_series.iloc[-21:-1].mean()) if len(vol_series) >= 21 else safe_num(vol_series.tail(20).mean())
        row["adtv20"] = adtv_base
        row["adtv20_pct"] = ((row["last_vol"] / adtv_base) - 1) * 100 if pd.notna(row["last_vol"]) and pd.notna(adtv_base) and adtv_base != 0 else np.nan
        row["adtv20_zone"] = ("N/A" if pd.isna(row["adtv20_pct"]) else ("Above 3%" if row["adtv20_pct"] >= 3 else ("Below 3%" if row["adtv20_pct"] <= -3 else "Within ±3%")))

        row["rvol20"] = row["last_vol"] / adtv_base if pd.notna(row["last_vol"]) and pd.notna(adtv_base) and adtv_base != 0 else np.nan
        row["rvol20_zone"] = "Above 1.5" if pd.notna(row["rvol20"]) and row["rvol20"] >= 1.5 else ("Below 1.5" if pd.notna(row["rvol20"]) else "N/A")

        # Volatility (ADR / ATR logic)
        ADR_LEN = 14

        if len(hist) >= 2:
            # ADR = SMA(High - Low, Length) using PRIOR COMPLETED bars only (exclude current/latest bar)
            daily_range = (hist["High"] - hist["Low"]).astype(float)

            # TradingView-aligned ADR = SMA(High - Low, Length) INCLUDING current/latest bar
            # ADR% = (ADR / Current Close) * 100
            adr = safe_num(daily_range.tail(ADR_LEN).mean()) if len(daily_range) >= 1 else np.nan
            row["adr_pct"] = safe_num((adr / row["close"]) * 100.0) if pd.notna(adr) and pd.notna(row["close"]) and row["close"] != 0 else np.nan

            # ATR(14) % = current ATR(14) relative to current close
            prev_close = hist["Close"].shift(1)
            tr = pd.concat([
                hist["High"] - hist["Low"],
                (hist["High"] - prev_close).abs(),
                (hist["Low"] - prev_close).abs()
            ], axis=1).max(axis=1)

            atr14 = tr.ewm(alpha=1 / 14, adjust=False).mean()
            row["atr14_pct"] = safe_num((atr14.iloc[-1] / row["close"]) * 100.0) if len(atr14) and pd.notna(row["close"]) and row["close"] != 0 else np.nan

            row["adr_atr_zone"] = "Above 3%" if (
                pd.notna(row["adr_pct"]) and pd.notna(row["atr14_pct"]) and row["adr_pct"] >= 3 and row["atr14_pct"] >= 3
            ) else ("Below 3%" if pd.notna(row["adr_pct"]) and pd.notna(row["atr14_pct"]) else "N/A")

        # ADTR 20D = Average Daily Trading Value (Rupiah) using OHLC4 Typical Price
        if "Open" in hist.columns:
            value_series = hist["Volume"] * ((hist["Open"] + hist["High"] + hist["Low"] + hist["Close"]) / 4.0)
        else:
            value_series = hist["Volume"] * ((hist["High"] + hist["Low"] + hist["Close"]) / 3.0)
        row["adtr20"] = safe_num(value_series.iloc[-21:-1].mean()) if len(value_series) >= 21 else safe_num(value_series.tail(20).mean())
        row["adtr20_pct"] = ((row["daily_value"] / row["adtr20"]) - 1) * 100 if pd.notna(row["daily_value"]) and pd.notna(row["adtr20"]) and row["adtr20"] != 0 else np.nan
        row["adtr20_zone"] = ("N/A" if pd.isna(row["adtr20_pct"]) else ("Above 3%" if row["adtr20_pct"] >= 3 else ("Below 3%" if row["adtr20_pct"] <= -3 else "Within ±3%")))

        ma_specs = [
            ("ema10", "ema10d", "ema10p", "EMA10", 10, "ema"),
            ("ema20", "ema20d", "ema20p", "EMA20", 20, "ema"),
            ("ema25", "ema25d", "ema25p", "EMA25", 25, "ema"),
            ("ema50", "ema50d", "ema50p", "EMA50", 50, "ema"),
            ("sma200", "sma200d", "sma200p", "SMA200", 200, "sma"),
        ]
        available_ma = []
        for value_key, diff_key, pos_key, label, period, kind in ma_specs:
            if len(hist) < period:
                row[value_key], row[diff_key], row[pos_key] = np.nan, np.nan, "N/A"
                continue
            series = ema(hist["Close"], period) if kind == "ema" else hist["Close"].rolling(period).mean()
            row[value_key] = safe_num(series.iloc[-1])
            row[diff_key] = pct_diff(row["close"], row[value_key])
            row[pos_key] = pos_label(row["close"], row[value_key]) or "N/A"
            if row[pos_key] in ("Above", "Below"):
                available_ma.append((label, row[pos_key]))

        if available_ma:
            above = [label for label, position in available_ma if position == "Above"]
            below = [label for label, position in available_ma if position == "Below"]
            if not below:
                row["summary_ma"] = "Above All Available MA"
            elif not above:
                row["summary_ma"] = "Below All Available MA"
            else:
                row["summary_ma"] = f"Above {', '.join(above)} | Below {', '.join(below)}"
        else:
            row["summary_ma"] = "N/A"

        rsi_series = rsi(hist["Close"], 14)
        row["rsi14"] = safe_num(rsi_series.iloc[-1])
        rsi_prev = safe_num(rsi_series.iloc[-2]) if len(rsi_series) >= 2 else np.nan
        row["rsi_delta"] = row["rsi14"] - rsi_prev if pd.notna(row["rsi14"]) and pd.notna(rsi_prev) else np.nan
        row["rsi_ge_50"] = "YES" if pd.notna(row["rsi14"]) and row["rsi14"] >= 50 else "NO"
        row["rsi_status"] = rsi_status(row["rsi14"])

        rsi_ma_series = rsi_series.rolling(14).mean()
        row["rsi_ma14"] = safe_num(rsi_ma_series.iloc[-1])
        if pd.isna(row["rsi14"]) or pd.isna(row["rsi_ma14"]):
            row["rsi_pos"] = "N/A"
        else:
            row["rsi_pos"] = "Above" if row["rsi14"] >= row["rsi_ma14"] else "Below"
        rsi_ma_prev = safe_num(rsi_ma_series.iloc[-2]) if len(rsi_ma_series) >= 2 else np.nan
        row["golden_cross"] = "YES" if all(pd.notna(x) for x in [rsi_prev, rsi_ma_prev, row["rsi14"], row["rsi_ma14"]]) and rsi_prev <= rsi_ma_prev and row["rsi14"] > row["rsi_ma14"] else "NO"
        row["dead_cross"] = "YES" if all(pd.notna(x) for x in [rsi_prev, rsi_ma_prev, row["rsi14"], row["rsi_ma14"]]) and rsi_prev >= rsi_ma_prev and row["rsi14"] < row["rsi_ma14"] else "NO"
        row["cross_status"] = "Golden" if row["golden_cross"] == "YES" else ("Dead" if row["dead_cross"] == "YES" else "-")
        # S-05: EMA50/SMA200 Golden Cross trigger label
        try:
            _e50  = ema(hist["Close"], 50)
            _s200 = hist["Close"].rolling(200, min_periods=50).mean()
            _e50_c, _e50_p  = float(_e50.iloc[-1]),  float(_e50.iloc[-2])  if len(_e50)  >= 2 else np.nan
            _s200_c, _s200_p = float(_s200.iloc[-1]), float(_s200.iloc[-2]) if len(_s200) >= 2 else np.nan
            _gc_trigger = pd.notna(_e50_p) and pd.notna(_s200_p) and _e50_c > _s200_c and _e50_p <= _s200_p
            _dc_trigger = pd.notna(_e50_p) and pd.notna(_s200_p) and _e50_c < _s200_c and _e50_p >= _s200_p
            if _gc_trigger:
                row["ema50_sma200_cross"] = "Golden Cross ← Today"
            elif _dc_trigger:
                row["ema50_sma200_cross"] = "Dead Cross ← Today"
            elif pd.notna(_e50_c) and pd.notna(_s200_c) and _e50_c > _s200_c:
                row["ema50_sma200_cross"] = "EMA50 > SMA200"
            elif pd.notna(_e50_c) and pd.notna(_s200_c):
                row["ema50_sma200_cross"] = "EMA50 < SMA200"
        except Exception:
            row["ema50_sma200_cross"] = "-"

        # MACD Momentum (Boring Jacx setup)
        row.update(compute_macd_momentum(hist))
        row.update(compute_stochastic(hist))

        prev_close = hist["Close"].shift(1)
        tr = pd.concat([hist["High"] - hist["Low"], (hist["High"] - prev_close).abs(), (hist["Low"] - prev_close).abs()], axis=1).max(axis=1)
        atr14 = safe_num(tr.rolling(14).mean().iloc[-1])
        adr_series = ((hist["High"] - hist["Low"]) / hist["Close"].replace(0, np.nan)) * 100.0

        avg_daily_value = row["avg30"] * row["close"] if pd.notna(row["avg30"]) and pd.notna(row["close"]) else np.nan
        regime = classify_stock_regime({
            "market_cap": row["mcap"],
            "avg_value_30d": avg_daily_value,
            "classic_hhi": row["hhi"],
            "cr1": row["cr1"],
            "ccs": row["ccs"],
        })
        row["market_cap_category"] = regime["market_cap_category"]
        row["liquidity_category"] = regime["liquidity_category"]
        row["stock_regime"] = regime["stock_regime"]

        # Market Structure (Leviathan-style, no score)
        ms = compute_market_structure_v1(hist)
        row.update(ms)
        # Item #25: return '-' for ambiguous mid-zone states
        _smc_sum = str(row.get("ms_summary", "") or "")
        if _smc_sum.lower() in ("mixed", "consolidating", "ranging", "neutral", "", "n/a"):
            row["ms_summary"] = "-"

        # S-08: EQ Breakout pre-computation
        # Condition: close above EQ High today, AND last 2 bars were inside [EQ Low, EQ High]
        try:
            _eq_hi = safe_num(row.get("smc_eq_high"), np.nan)
            _eq_lo = safe_num(row.get("smc_eq_low"),  np.nan)
            _close_now = safe_num(hist["Close"].iloc[-1])
            if pd.notna(_eq_hi) and pd.notna(_eq_lo) and _eq_hi > _eq_lo and pd.notna(_close_now):
                # Today must close above EQ High
                _broke = _close_now > _eq_hi
                # Last 2 prior bars must have been inside the EQ range
                _prior_closes = hist["Close"].iloc[-3:-1]   # bars t-2 and t-1
                _consolidated = (
                    len(_prior_closes) >= 2 and
                    all(_eq_lo <= float(c) <= _eq_hi for c in _prior_closes)
                )
                row["eq_breakout_flag"] = bool(_broke and _consolidated)
            else:
                row["eq_breakout_flag"] = False
        except Exception:
            row["eq_breakout_flag"] = False

        # Divergence detection must run in FULL mode before display mapping
        if len(hist) >= 15:
            rsi_series2 = rsi(hist["Close"], 14)
            div = divergence_signals(
                hist,
                rsi_series2,
                lookback=120,
                swing_window=2,
                min_separation=4,
                price_tol=0.0075,
                rsi_tol=2.0,
                max_last_swing_age=50
            )
            row.update(div)
            if row.get("div_signal", "None") == "None":
                row["div_ref1_date"] = "None"
                row["div_ref2_date"] = "None"

        row["pe_zone"] = vwap_zone_2pct(row["pe_ttm"], [row["pe_m1"], row["pe_m2"]]) if pd.notna(row["pe_ttm"]) else "N/A"
        pe_fields = [row["pe_ttm"], row["pe_mean"], row["pe_m1"], row["pe_m2"]]
        row["pe_band_source"] = "Source" if all(pd.notna(x) for x in pe_fields) else ("Partial Source" if any(pd.notna(x) for x in pe_fields) else "N/A")
        row["pbv_zone"] = vwap_zone_2pct(row["pbv_curr"], [row["pbv_m1"], row["pbv_m2"]]) if pd.notna(row["pbv_curr"]) else "N/A"
        pbv_fields = [row["pbv_curr"], row["pbv_mean"], row["pbv_m1"], row["pbv_m2"]]
        row["pbv_band_source"] = "Source" if all(pd.notna(x) for x in pbv_fields) else ("Partial Source" if any(pd.notna(x) for x in pbv_fields) else "N/A")

        # ── MVWAP field assignments ────────────────────────────────────────
        row["cm_days"], row["cm_vwap"], row["cm_m1"], row["cm_m2"], row["cm_m3"], row["cm_sd"], row["cm_delta"] = cm["days"], cm["vwap"], cm["m1"], cm["m2"], cm["m3"], cm["sd_score"], cm["sd_delta"]
        row["cm_p1"], row["cm_p2"], row["cm_p3"] = cm.get("p1", np.nan), cm.get("p2", np.nan), cm.get("p3", np.nan)
        row["pm_days"], row["pm_vwap"], row["pm_m1"], row["pm_m2"], row["pm_m3"], row["pm_sd"] = pm["days"], pm["vwap"], pm["m1"], pm["m2"], pm["m3"], pm["sd_score"]
        row["pm_p1"], row["pm_p2"], row["pm_p3"] = pm.get("p1", np.nan), pm.get("p2", np.nan), pm.get("p3", np.nan)
        row["pm_delta"] = pm["sd_delta"]

        # ── QVWAP / PY field assignments ──────────────────────────────────
        row["q_days"], row["q_vwap"], row["q_m1"], row["q_m2"], row["q_m3"], row["q_sd"], row["q_delta"] = q["days"], q["vwap"], q["m1"], q["m2"], q["m3"], q["sd_score"], q["sd_delta"]
        row["q_p1"], row["q_p2"], row["q_p3"] = q.get("p1", np.nan), q.get("p2", np.nan), q.get("p3", np.nan)
        row["pq_days"], row["pq_vwap"], row["pq_m1"], row["pq_m2"], row["pq_m3"], row["pq_sd"] = pq["days"], pq["vwap"], pq["m1"], pq["m2"], pq["m3"], pq["sd_score"]
        row["pq_p1"], row["pq_p2"], row["pq_p3"] = pq.get("p1", np.nan), pq.get("p2", np.nan), pq.get("p3", np.nan)
        row["py_year"], row["py_vwap"], row["py_m1"], row["py_m2"], row["py_m3"], row["py_sd"] = py["days"], py["vwap"], py["m1"], py["m2"], py["m3"], py["sd_score"]
        row["py_p1"], row["py_p2"], row["py_p3"] = py.get("p1", np.nan), py.get("p2", np.nan), py.get("p3", np.nan)

        # σ delta / relative logic
        # - Current MVWAP: true live 1D delta within current month anchor
        # - Previous MVWAP: relative comparison vs prior month end
        # - Current QVWAP: true live 1D delta within current quarter anchor
        # - Previous QVWAP: relative comparison vs prior quarter end
        # - Previous Year VWAP: relative comparison vs prior year end
        row["pq_delta"] = pq["sd_delta"]
        row["py_delta"] = py["sd_delta"]

        row["cm_zone"] = vwap_zone_2pct(row["close"], [row["cm_vwap"], row["cm_m1"], row["cm_m2"], row["cm_m3"]])
        row["pm_zone"] = vwap_zone_2pct(row["close"], [row["pm_vwap"], row["pm_m1"], row["pm_m2"], row["pm_m3"]])
        row["q_zone"]  = vwap_zone_2pct(row["close"], [row["q_vwap"],  row["q_m1"],  row["q_m2"],  row["q_m3"]])
        row["pq_zone"] = vwap_zone_2pct(row["close"], [row["pq_vwap"], row["pq_m1"], row["pq_m2"], row["pq_m3"]])
        row["py_zone"] = vwap_zone_2pct(row["close"], [row["py_vwap"], row["py_m1"], row["py_m2"], row["py_m3"]])
        row["cm_remarks"] = vwap_near_zone_label(row["close"], row["cm_sd"], row["cm_vwap"], row["cm_m1"], row["cm_m2"], row["cm_m3"], row["cm_p1"], row["cm_p2"], row["cm_p3"])
        row["pm_remarks"] = vwap_near_zone_label(row["close"], row["pm_sd"], row["pm_vwap"], row["pm_m1"], row["pm_m2"], row["pm_m3"], row["pm_p1"], row["pm_p2"], row["pm_p3"])
        row["q_remarks"]  = vwap_near_zone_label(row["close"], row["q_sd"],  row["q_vwap"],  row["q_m1"],  row["q_m2"],  row["q_m3"],  row["q_p1"],  row["q_p2"],  row["q_p3"])
        row["pq_remarks"] = vwap_near_zone_label(row["close"], row["pq_sd"], row["pq_vwap"], row["pq_m1"], row["pq_m2"], row["pq_m3"], row["pq_p1"], row["pq_p2"], row["pq_p3"])
        row["py_remarks"] = vwap_near_zone_label(row["close"], row["py_sd"], row["py_vwap"], row["py_m1"], row["py_m2"], row["py_m3"], row["py_p1"], row["py_p2"], row["py_p3"])
        row["cm_zone_days"] = _count_consecutive_zone_days(hist, [row["cm_vwap"], row["cm_m1"], row["cm_m2"], row["cm_m3"]], row["cm_remarks"])
        row["pm_zone_days"] = _count_consecutive_zone_days(hist, [row["pm_vwap"], row["pm_m1"], row["pm_m2"], row["pm_m3"]], row["pm_remarks"])
        row["q_zone_days"]  = _count_consecutive_zone_days(hist, [row["q_vwap"],  row["q_m1"],  row["q_m2"],  row["q_m3"]],  row["q_remarks"])
        row["pq_zone_days"] = _count_consecutive_zone_days(hist, [row["pq_vwap"], row["pq_m1"], row["pq_m2"], row["pq_m3"]], row["pq_remarks"])
        row["py_zone_days"] = _count_consecutive_zone_days(hist, [row["py_vwap"], row["py_m1"], row["py_m2"], row["py_m3"]], row["py_remarks"])
        any_zone = any(str(z).strip() not in ("", "N/A", "None") for z in [row["q_zone"], row["pq_zone"], row["py_zone"]])
        above_count = sum(1 for x in [row["ema10p"], row["ema20p"], row["ema25p"], row["ema50p"], row["sma200p"]] if x == "Above")
        if any_zone and pd.notna(row["rvol5"]) and row["rvol5"] >= MIN_RVOL:
            if above_count >= 5:
                row["tier"] = "A"
            elif above_count >= 3:
                row["tier"] = "B"

        # Legacy participant-inference metrics are disabled in source-limited mode.
        if row.get("div_signal") in ("Bullish", "Bearish"):
            row["divergence_summary"] = f'{row["div_signal"]} - {row.get("div_strength", "")}'.strip(" -")
        else:
            row["divergence_summary"] = "None"
        row = _finalize_flow_position_fields(row, hist)
        # Wyckoff proxy removed (v2.0)
        row = apply_dashboard_presets(row)

        # =========================
        # DISPLAY NORMALIZATION / N-A RULES (FULL MODE)
        # =========================
        if safe_num(row.get("cm_days", 0), 0) <= 0:
            row["cm_days"] = "N/A"; row["cm_remarks"] = "N/A"
        if safe_num(row.get("pm_days", 0), 0) <= 0:
            row["pm_days"] = "N/A"; row["pm_remarks"] = "N/A"
        if safe_num(row.get("q_days", 0), 0) <= 0:
            row["q_days"] = "N/A"
            row["q_remarks"] = "N/A"
        if safe_num(row.get("pq_days", 0), 0) <= 0:
            row["pq_days"] = "N/A"
            row["pq_remarks"] = "N/A"
        if safe_num(row.get("py_year", 0), 0) <= 0:
            row["py_year"] = "N/A"
            row["py_remarks"] = "N/A"

        if safe_num(row.get("last_vol", 0), 0) <= 0:
            for k in ["avg30", "last_vol", "last_ok", "rvol5", "rvol_zone", "lot", "daily_value", "adtv20", "adtv20_pct", "adtv20_zone", "adtr20", "adtr20_pct", "adtr20_zone", "rvol20", "rvol20_zone", "adr_pct", "atr14_pct", "adr_atr_zone"]:
                row[k] = "N/A"

        if pd.notna(row.get("ibh")) and pd.notna(row.get("ibl")) and safe_num(row.get("ibh")) == safe_num(row.get("ibl")):
            row["mp_zone"] = "NO"

        if pd.isna(row.get("rsi14")):
            for k in ["rsi14","rsi_ma14","rsi_ge_50","rsi_delta","rsi_status","cross_status","divergence_summary","div_ref1_date","div_ref2_date"]:
                row[k] = "N/A"

        return row

    # PARTIAL mode
    if len(hist) >= MIN_BARS_PARTIAL:
        row["data_status"] = "PARTIAL DATA"

        vol_series = hist["Volume"].copy()
        row["last_vol"] = safe_num(vol_series.iloc[-1])
        if len(vol_series) >= 6:
            avg5 = safe_num(vol_series.iloc[-6:-1].mean())
            row["rvol5"] = row["last_vol"] / avg5 if pd.notna(avg5) and avg5 != 0 else np.nan
            row["rvol_zone"] = "YES" if pd.notna(row["rvol5"]) and row["rvol5"] >= MIN_RVOL else "NO"

            # Liquidity block (institutional)
            row["lot"] = row["last_vol"] / 100 if pd.notna(row["last_vol"]) else np.nan
            # Value (Approx): Typical Price = (Open + High + Low + Close) / 4 × Volume
            row["daily_value"] = (safe_num(hist["Volume"].iloc[-1] * ((hist["Open"].iloc[-1] + hist["High"].iloc[-1] + hist["Low"].iloc[-1] + hist["Close"].iloc[-1]) / 4.0)) if len(hist) and all(col in hist.columns for col in ["Volume","Open","High","Low","Close"]) else np.nan)

            adtv_base = safe_num(vol_series.iloc[-21:-1].mean()) if len(vol_series) >= 21 else safe_num(vol_series.tail(20).mean())
            row["adtv20"] = adtv_base
            row["adtv20_pct"] = ((row["last_vol"] / adtv_base) - 1) * 100 if pd.notna(row["last_vol"]) and pd.notna(adtv_base) and adtv_base != 0 else np.nan
            row["adtv20_zone"] = ("N/A" if pd.isna(row["adtv20_pct"]) else ("Above 3%" if row["adtv20_pct"] >= 3 else ("Below 3%" if row["adtv20_pct"] <= -3 else "Within ±3%")))

            row["rvol20"] = row["last_vol"] / adtv_base if pd.notna(row["last_vol"]) and pd.notna(adtv_base) and adtv_base != 0 else np.nan
            row["rvol20_zone"] = "Above 1.5" if pd.notna(row["rvol20"]) and row["rvol20"] >= 1.5 else ("Below 1.5" if pd.notna(row["rvol20"]) else "N/A")

            if pd.notna(shares_used) and shares_used > 0:
                turn_series = (vol_series / shares_used) * 100.0
                row["adtr20"] = safe_num(turn_series.iloc[-21:-1].mean()) if len(turn_series) >= 21 else safe_num(turn_series.tail(20).mean())
                row["adtr20_pct"] = (row["last_vol"] / shares_used) * 100.0 if pd.notna(row["last_vol"]) else np.nan
                adtr_dev = (((row["adtr20_pct"] / row["adtr20"]) - 1) * 100 if pd.notna(row["adtr20"]) and row["adtr20"] != 0 and pd.notna(row["adtr20_pct"]) else np.nan)
            row["adtr20_zone"] = ("N/A" if pd.isna(adtr_dev) else ("Above 3%" if adtr_dev >= 3 else ("Below 3%" if adtr_dev <= -3 else "Within ±3%")))
        if len(vol_series) >= 31:
            row["avg30"] = safe_num(vol_series.iloc[-31:-1].mean())
        elif len(vol_series) >= 5:
            row["avg30"] = safe_num(vol_series.mean())

        row["last_ok"] = "YES" if pd.notna(row["last_vol"]) and row["last_vol"] >= VOL_THRESHOLD else ("NO" if pd.notna(row["last_vol"]) else "N/A")

        if len(hist) >= 10:
            row["ema10"] = safe_num(ema(hist["Close"], 10).iloc[-1])
            row["ema10d"], row["ema10p"] = pct_diff(row["close"], row["ema10"]), pos_label(row["close"], row["ema10"])
        if len(hist) >= 20:
            row["ema20"] = safe_num(ema(hist["Close"], 20).iloc[-1])
            row["ema20d"], row["ema20p"] = pct_diff(row["close"], row["ema20"]), pos_label(row["close"], row["ema20"])
        if len(hist) >= 25:
            row["ema25"] = safe_num(ema(hist["Close"], 25).iloc[-1])
            row["ema25d"], row["ema25p"] = pct_diff(row["close"], row["ema25"]), pos_label(row["close"], row["ema25"])
        if len(hist) >= 50:
            row["ema50"] = safe_num(ema(hist["Close"], 50).iloc[-1])
            row["ema50d"], row["ema50p"] = pct_diff(row["close"], row["ema50"]), pos_label(row["close"], row["ema50"])
        if len(hist) >= 200:
            row["sma200"] = safe_num(hist["Close"].rolling(200).mean().iloc[-1])
            row["sma200d"], row["sma200p"] = pct_diff(row["close"], row["sma200"]), pos_label(row["close"], row["sma200"])

        available_ma = []
        for label, pos in [("EMA10", row["ema10p"]), ("EMA20", row["ema20p"]), ("EMA25", row["ema25p"]), ("EMA50", row["ema50p"]), ("SMA200", row["sma200p"])]:
            if pos in ("Above", "Below"):
                available_ma.append((label, pos))
        if available_ma:
            above = [x[0] for x in available_ma if x[1] == "Above"]
            if len(above) == len(available_ma):
                row["summary_ma"] = "Above All MA"
            elif len(above) == 0:
                row["summary_ma"] = "Below All MA"
            else:
                row["summary_ma"] = "Above " + ", ".join(above)

        if len(hist) >= 15:
            rsi_series = rsi(hist["Close"], 14)
            row["rsi14"] = safe_num(rsi_series.iloc[-1])
            rsi_prev = safe_num(rsi_series.iloc[-2]) if len(rsi_series) >= 2 else np.nan
            row["rsi_delta"] = row["rsi14"] - rsi_prev if pd.notna(row["rsi14"]) and pd.notna(rsi_prev) else np.nan
            row["rsi_ge_50"] = "YES" if pd.notna(row["rsi14"]) and row["rsi14"] >= 50 else "NO"
            row["rsi_status"] = rsi_status(row["rsi14"])
            if len(hist) >= 28:
                rsi_ma_series = rsi_series.rolling(14).mean()
                row["rsi_ma14"] = safe_num(rsi_ma_series.iloc[-1])
                rsi_ma_prev = safe_num(rsi_ma_series.iloc[-2]) if len(rsi_ma_series) >= 2 else np.nan
                row["golden_cross"] = "YES" if all(pd.notna(x) for x in [rsi_prev, rsi_ma_prev, row["rsi14"], row["rsi_ma14"]]) and rsi_prev <= rsi_ma_prev and row["rsi14"] > row["rsi_ma14"] else "NO"
                row["dead_cross"] = "YES" if all(pd.notna(x) for x in [rsi_prev, rsi_ma_prev, row["rsi14"], row["rsi_ma14"]]) and rsi_prev >= rsi_ma_prev and row["rsi14"] < row["rsi_ma14"] else "NO"
                row["cross_status"] = "Golden" if row["golden_cross"] == "YES" else ("Dead" if row["dead_cross"] == "YES" else "-")

        # MACD Momentum (Boring Jacx setup) — runs if enough bars
        if len(hist) >= 35:
            row.update(compute_macd_momentum(hist))

        if len(hist) >= 20:
            row.update(compute_stochastic(hist))

        if len(hist) >= 14:
            prev_close = hist["Close"].shift(1)
            tr = pd.concat([hist["High"] - hist["Low"], (hist["High"] - prev_close).abs(), (hist["Low"] - prev_close).abs()], axis=1).max(axis=1)
            atr14 = safe_num(tr.rolling(14).mean().iloc[-1])
        if len(hist) >= 20:
            adr_series = ((hist["High"] - hist["Low"]) / hist["Close"].replace(0, np.nan)) * 100.0

        avg_daily_value = row["avg30"] * row["close"] if pd.notna(row["avg30"]) and pd.notna(row["close"]) else np.nan
        regime = classify_stock_regime({
            "market_cap": row["mcap"],
            "avg_value_30d": avg_daily_value,
            "classic_hhi": row["hhi"],
            "cr1": row["cr1"],
            "ccs": row["ccs"],
        })
        row["market_cap_category"] = regime["market_cap_category"]
        row["liquidity_category"] = regime["liquidity_category"]
        row["stock_regime"] = regime["stock_regime"]

        if len(hist) >= 15:
            rsi_series2 = rsi(hist["Close"], 14)
            div = divergence_signals(
                hist,
                rsi_series2,
                lookback=120,
                swing_window=2,
                min_separation=4,
                price_tol=0.0075,
                rsi_tol=2.0,
                max_last_swing_age=50
            )
            row.update(div)

        row["pe_zone"] = vwap_zone_2pct(row["pe_ttm"], [row["pe_m1"], row["pe_m2"]]) if pd.notna(row["pe_ttm"]) else "N/A"
        pe_fields = [row["pe_ttm"], row["pe_mean"], row["pe_m1"], row["pe_m2"]]
        row["pe_band_source"] = "Source" if all(pd.notna(x) for x in pe_fields) else ("Partial Source" if any(pd.notna(x) for x in pe_fields) else "N/A")
        row["pbv_zone"] = vwap_zone_2pct(row["pbv_curr"], [row["pbv_m1"], row["pbv_m2"]]) if pd.notna(row["pbv_curr"]) else "N/A"
        pbv_fields = [row["pbv_curr"], row["pbv_mean"], row["pbv_m1"], row["pbv_m2"]]
        row["pbv_band_source"] = "Source" if all(pd.notna(x) for x in pbv_fields) else ("Partial Source" if any(pd.notna(x) for x in pbv_fields) else "N/A")

        # Legacy participant-inference metrics are disabled in source-limited mode.
        if row.get("div_signal") in ("Bullish", "Bearish"):
            row["divergence_summary"] = f'{row["div_signal"]} - {row.get("div_strength", "")}'.strip(" -")
        else:
            row["divergence_summary"] = "None"
        row = _finalize_flow_position_fields(row, hist)
        # Wyckoff proxy removed (v2.0)
        row = apply_dashboard_presets(row)

        # =========================
        # DISPLAY NORMALIZATION / N-A RULES (FULL MODE)
        # =========================
        if safe_num(row.get("q_days", 0), 0) <= 0:
            row["q_days"] = "N/A"
            row["q_remarks"] = "N/A"
        if safe_num(row.get("pq_days", 0), 0) <= 0:
            row["pq_days"] = "N/A"
            row["pq_remarks"] = "N/A"
        if safe_num(row.get("py_year", 0), 0) <= 0:
            row["py_year"] = "N/A"
            row["py_remarks"] = "N/A"

        if safe_num(row.get("last_vol", 0), 0) <= 0:
            for k in ["avg30", "last_vol", "last_ok", "rvol5", "rvol_zone", "lot", "daily_value", "adtv20", "adtv20_pct", "adtv20_zone", "adtr20", "adtr20_pct", "adtr20_zone", "rvol20", "rvol20_zone", "adr_pct", "atr14_pct", "adr_atr_zone"]:
                row[k] = "N/A"

        if pd.notna(row.get("ibh")) and pd.notna(row.get("ibl")) and safe_num(row.get("ibh")) == safe_num(row.get("ibl")):
            row["mp_zone"] = "NO"

        ma_vals = [row.get("ema10"), row.get("ema20"), row.get("ema25"), row.get("ema50")]
        if pd.notna(row.get("close")) and all(pd.notna(v) and safe_num(v) == safe_num(row.get("close")) for v in ma_vals):
            for k in ["ema10d","ema10p","ema20d","ema20p","ema25d","ema25p","ema50d","ema50p","sma200d","sma200p"]:
                row[k] = "N/A"
            row["summary_ma"] = "N/A"

        if pd.isna(row.get("rsi14")):
            for k in ["rsi14","rsi_ma14","rsi_ge_50","rsi_delta","rsi_status","cross_status","divergence_summary","div_ref1_date","div_ref2_date"]:
                row[k] = "N/A"

        row = _finalize_flow_position_fields(row, hist)
        return row

    row = _finalize_flow_position_fields(row, hist)
    return row

def apply_dashboard_presets(row: dict):
    row["preset_accumulation"] = "YES" if row.get("af_verdict") in ("ACCUMULATION", "STRONG ACCUMULATION") else "NO"
    row["preset_high_conviction"] = "YES" if row.get("af_verdict") in ("ACCUMULATION", "STRONG ACCUMULATION") and safe_num(row.get("af_verdict_confidence"), 0) >= 80 else "NO"
    row["preset_distribution"] = "YES" if row.get("af_verdict") in ("DISTRIBUTION", "STRONG DISTRIBUTION") else "NO"
    row["preset_distribution_risk"] = "YES" if row.get("preset_distribution") == "YES" and (
        str(row.get("wyckoff_proxy_phase", "")) in ("Distribution", "Markdown") or
        safe_num(row.get("breakdown_risk"), 0) >= 70 or
        str(row.get("af_phase_event", "")).startswith("Markdown")
    ) else "NO"
    row["preset_smart_money"] = "YES" if safe_num(row.get("af_smt_proxy"), 0) >= 75 else "NO"
    row["preset_blue_chips"] = "YES" if str(row.get("market_cap_category", "")) == "Large Cap" and str(row.get("liquidity_category", "")) == "High Liquidity" else "NO"
    row["preset_flow_leaders"] = "YES" if (
        str(row.get("af_flow_edge", "")) == "Strong" or
        safe_num(row.get("af_r2_proxy"), 0) >= 12 or
        safe_num(row.get("af_hit_rate_proxy"), 0) >= 60
    ) else "NO"
    active_event = (
        str(row.get("_removed_spring", "NO")) in ("YES", "Weak") or
        str(row.get("_removed_upthrust", "NO")) in ("YES", "Weak") or
        ("Bullish" in str(row.get("divergence_summary", ""))) or
        ("Bearish" in str(row.get("divergence_summary", ""))) or
        str(row.get("af_phase_event", "")).split(" / ")[-1] not in ("No Clean Event", "Continuation", "Base Build")
    )
    row["preset_active_wyckoff_events"] = "YES" if active_event else "NO"
    return row

def _zone_is_tradeable_poi(zone_text: str) -> bool:
    s = str(zone_text or "").strip()
    if not s or s == "N/A":
        return False
    s_u = s.upper()
    return ("PRICE NEAR" in s_u) or ("PRICE RANGING" in s_u)

def _pct_from_level(close_px, level_px):
    c = safe_num(close_px, np.nan)
    lv = safe_num(level_px, np.nan)
    if pd.isna(c) or pd.isna(lv) or lv == 0:
        return np.nan
    return ((c / lv) - 1.0) * 100.0

def _near_level(close_px, level_px, tol_pct):
    p = _pct_from_level(close_px, level_px)
    return pd.notna(p) and abs(p) <= tol_pct

def _ms_is_continuation_ok(r: dict) -> bool:
    ms_regime = str(r.get("ms_trend_regime", "") or "")
    ms_state = str(r.get("ms_structure_state", "") or "")
    last_event = str(r.get("ms_last_event", "") or "")
    if ("Bullish" in ms_regime) and (ms_state in ("HH-HL", "Transition Up") or last_event in ("Bull BOS", "Bull CHoCH")):
        return True
    return False

def _best_vwap_location_label(r: dict) -> str:
    z = _best_discount_zone(r)
    if z:
        if pd.notna(safe_num(z.get("dist_pct"), np.nan)):
            return f"{z['framework']} | {z['zone_type']} ({z['dist_pct']:+.2f}%)"
        return f"{z['framework']} | {z['zone_type']}"
    return "No valid discount POI"

def _buy_zone_type(r: dict) -> str:
    z = _best_discount_zone(r)
    return z["zone_type"] if z else "None"

def _secondary_vwap_context(r: dict) -> str:
    close_px = safe_num(r.get("close"), np.nan)
    hits = []
    checks = [
        ("PQVWAP",   r.get("pq_vwap"), 2.0),
        ("PQ -1 SD", r.get("pq_m1"),   1.25),
        ("PQ -2 SD", r.get("pq_m2"),   1.25),
        ("PYVWAP",   r.get("py_vwap"), 2.0),
        ("PY -1 SD", r.get("py_m1"),   1.25),
        ("PY -2 SD", r.get("py_m2"),   1.25),
    ]
    for name, level, tol in checks:
        if _near_level(close_px, level, tol):
            hits.append(name)
    if hits:
        return " + ".join(hits[:3])

    # fallback to text remarks
    out = []
    for lbl, z in [("PQ", r.get("pq_remarks", "")), ("PY", r.get("py_remarks", ""))]:
        z = str(z or "")
        if z and z != "N/A" and _zone_is_tradeable_poi(z):
            out.append(f"{lbl}: actionable")
    return " | ".join(out) if out else "No major HTF confluence"

def _liquidity_check(r: dict) -> str:
    rvol = safe_num(r.get("rvol20"), np.nan)
    dval = safe_num(r.get("value_traded"), np.nan)
    adtv = safe_num(r.get("adtv20"), np.nan)
    daily_val = safe_num(r.get("daily_value"), np.nan)

    # Try both value_traded and daily_value field names
    eff_dval = dval if pd.notna(dval) else daily_val

    if pd.notna(eff_dval) and eff_dval >= 10_000_000_000:
        return "Strong"
    if pd.notna(adtv) and adtv >= 10_000_000_000:
        return "Pass"
    if pd.notna(rvol) and rvol >= 1.0:
        return "Pass"
    if pd.notna(rvol) and rvol >= 0.7:
        return "Pass"
    return "Thin"

def _is_deep_discount_zone(zone_type: str) -> bool:
    return zone_type in ("Q -2 SD", "PQ -1 SD", "PYVWAP", "PY -1 SD", "Discount Zone")

def _is_obvious_breakdown(r: dict, zone_type: str) -> bool:
    ms_regime = str(r.get("ms_trend_regime", "") or "")
    ms_state = str(r.get("ms_structure_state", "") or "")
    last_event = str(r.get("ms_last_event", "") or "")
    if ms_regime == "Bearish" and ms_state == "LH-LL" and last_event == "Bear BOS":
        return not _is_deep_discount_zone(zone_type)
    return False

def _trade_bias_from_context(r: dict, setup: str) -> str:
    last_event = str(r.get("ms_last_event", "") or "")
    if setup == "Continuation Reclaim":
        return "Reclaim"
    if setup == "Discount Pullback":
        return "POI → PI"
    if setup == "Deep Discount Bounce":
        return "POI → POI"
    if last_event in ("Bull CHoCH", "Bull BOS"):
        return "POI → PI"
    return "POI → POI"


# =========================
# VWAP SCREENER
# =========================

def _score_ma_alignment(r: dict) -> tuple:
    """
    Returns (ma_tier, ma_label)
      4 = Above EMA20, EMA25, EMA50  → full bull alignment
      3 = Above EMA20, EMA25         → constructive
      2 = Above EMA20 only           → short-term only
      1 = Above EMA50 only           → secular floor, MAs broken
      0 = Below all                  → excluded (unless deep discount exception)
    """
    c     = safe_num(r.get("close"), np.nan)
    e20   = safe_num(r.get("ema20"), np.nan)
    e25   = safe_num(r.get("ema25"), np.nan)
    e50   = safe_num(r.get("ema50"), np.nan)
    s200  = safe_num(r.get("sma200"), np.nan)
    if pd.isna(c):
        return (0, "N/A")
    above_e20  = pd.notna(e20)  and c >= e20
    above_e25  = pd.notna(e25)  and c >= e25
    above_e50  = pd.notna(e50)  and c >= e50
    above_s200 = pd.notna(s200) and c >= s200
    if above_e20 and above_e25 and above_e50 and above_s200:
        return (4, "EMA20 ▲ EMA25 ▲ EMA50 ▲ SMA200")
    if above_e20 and above_e25 and above_e50:
        return (3, "EMA20 ▲ EMA25 ▲ EMA50 ▲")
    if above_e20 and above_e25:
        return (2, "EMA20 ▲ EMA25 ▲")
    if above_e20:
        return (1, "EMA20 ▲ only")
    return (0, "Below all MAs")


def _reversal_score(r: dict) -> tuple:
    """
    Score 0-8 reversal signals. Returns (score, signals_list).
    Used both to gate EXCEPTION entries and to enrich PASS entries.
    """
    signals = []
    rsi14      = safe_num(r.get("rsi14"), np.nan)
    rsi_div    = str(r.get("div_signal", "") or "")
    macd_wave  = str(r.get("macd_wave", "") or "")
    spring     = str(r.get("_removed_spring", "NO") or "NO")
    rvol       = safe_num(r.get("rvol20"), np.nan)
    last_event = str(r.get("ms_last_event", "") or "")
    cross_rsi  = str(r.get("cross_status", "") or "")
    cross_macd = str(r.get("macd_cross", r.get("macd_cross_status", "")) or "")
    ms_state   = str(r.get("ms_structure_state", "") or "")

    if "Bullish" in rsi_div:          signals.append("Bullish Div")
    if "Recovering" in macd_wave:     signals.append("MACD↑")
    if spring in ("YES", "Weak"):     signals.append("Spring")
    if pd.notna(rsi14) and rsi14 <= 35:   signals.append(f"RSI{rsi14:.0f}")
    if pd.notna(rvol) and rvol >= 1.5:    signals.append(f"RVOL{rvol:.1f}x")
    if last_event in ("Bull CHoCH", "Bull BOS"):  signals.append(last_event)
    if cross_rsi == "Golden":         signals.append("RSI✕")
    if cross_macd in ("Golden", "Golden Cross"):  signals.append("MACD✕")
    return (len(signals), signals)


# =========================
# SWING SCORE ENGINE
# =========================

def _swing_score(r: dict, zone_type: str, tgt_dist: float) -> tuple:
    """
    TRUE POI ENGINE V17 — 7-component weighted composite score (0.0–10.0).
    Returns (composite_score, priority_label).

    Weights:
      Structure Quality + Freshness : 20%
      VWAP Position                 : 20%
      MACD Momentum                 : 20%
      RSI Momentum                  : 15%
      Candle Pattern                : 15%
      Liquidity                     : 10%
      Mean Reversion / Extremity    :  0% (normalized into final composite)
    """

    # ─── COMPONENT 1 – Structure Quality + Freshness (0–10) → weight 20% ────
    ms_regime    = str(r.get("ms_trend_regime", "") or "")
    ms_state     = str(r.get("ms_structure_state", "") or "")
    last_event   = str(r.get("ms_last_event", "") or "")
    event_age    = safe_num(r.get("ms_event_age_d"), np.nan)
    ms_quality   = str(r.get("ms_structure_quality", "") or "")

    struct_score = 0.0
    if "Bullish" in ms_regime and ms_state in ("HH-HL", "HH/HL"):
        struct_score += 5.0
    elif "Bullish" in ms_regime:
        struct_score += 3.5
    elif ms_regime in ("Neutral", "Sideways", "Range"):
        struct_score += 1.5
    # Event freshness
    if pd.notna(event_age):
        if event_age <= 5:
            struct_score += 5.0
        elif event_age <= 10:
            struct_score += 3.5
        elif event_age <= 20:
            struct_score += 2.0
    if ms_quality == "Clean":
        struct_score = min(10.0, struct_score + 0.5)
    struct_score = min(10.0, struct_score)

    # ─── COMPONENT 2 – VWAP Position (0–10) → weight 20% ────────────────────
    vwap_zone   = str(r.get("cq_vwap_zone", r.get("vwap_zone", "")) or "")
    vwap_days   = safe_num(r.get("cq_vwap_zone_days", r.get("vwap_zone_days")), np.nan)
    pq_rel_sd   = safe_num(r.get("pq_rel_sd_end", r.get("pq_sd_delta")), np.nan)
    py_rel_sd   = safe_num(r.get("py_rel_sd_end", r.get("py_sd_delta")), np.nan)
    py_zone     = str(r.get("py_vwap_zone", "") or "")

    _vwap_base = {
        "Price Above VWAP":             10,
        "Price Ranging VWAP⇄-1 SD":     8,
        "Price Ranging -1 SD⇄-2 SD":    6,
        "Price Near -1 SD":             5,
        "Price Near -2 SD":             4,
        "Price Near VWAP":              5,
        "At VWAP":                      5,
        "Price Ranging VWAP⇄+1 SD":    3,
        "Price Below -1 SD":            1,
        "Price Below -2 SD":            0,
        "Price Below -3 SD":            0,
    }
    vwap_score = 0.0
    for k, v in _vwap_base.items():
        if k.lower() in vwap_zone.lower():
            vwap_score = float(v)
            break
    else:
        # zone_type fallback for mapped POI zones
        _zone_fallback = {
            "PQ -2 SD": 4, "PY -2 SD": 4, "PQ -1 SD": 6,
            "PY -1 SD": 6, "PQVWAP": 8, "PYVWAP": 8,
        }
        vwap_score = float(_zone_fallback.get(zone_type, 5))

    # PQ bonus/penalty
    if pd.notna(pq_rel_sd):
        if pq_rel_sd > 0:
            vwap_score = min(10.0, vwap_score + 2.0)
        elif pq_rel_sd < 0:
            vwap_score = max(0.0, vwap_score - 2.0)
    # PY bonus/penalty
    if pd.notna(py_rel_sd):
        if py_rel_sd > 0 and "Above" in py_zone:
            vwap_score = min(10.0, vwap_score + 1.0)
        elif py_rel_sd < 0 and "Below" in py_zone:
            vwap_score = max(0.0, vwap_score - 1.0)
    # Zone age decay
    if pd.notna(vwap_days) and vwap_days > 10:
        vwap_score *= 0.7
    vwap_score = min(10.0, max(0.0, round(vwap_score, 1)))

    # ─── COMPONENT 3 – MACD Momentum (0–10) → weight 20% ────────────────────
    macd_wave    = str(r.get("macd_wave", r.get("macd_wave_pattern", "")) or "")
    macd_lines   = str(r.get("macd_lines_position", r.get("macd_line_position", "")) or "")

    _macd_reject = (
        "Mountain (Declining)" in macd_wave and "Both Below Zero" in macd_lines
    )
    macd_score = 0.0
    if "Mountain (Building)" in macd_wave:
        macd_score = 10.0 if "Both Above Zero" in macd_lines else 6.0
    elif "Valley (Recovering)" in macd_wave:
        macd_score = 9.0 if "Both Below Zero" not in macd_lines else 5.0
    elif "Mountain (Declining)" in macd_wave and "Both Above Zero" in macd_lines:
        macd_score = 5.0
    elif "Valley (Deepening)" in macd_wave:
        macd_score = 2.0
    elif "Both Below Zero" in macd_lines:
        macd_score = 4.0
    if "Mixed" in macd_lines:
        macd_score *= 0.6

    # ─── COMPONENT 4 – RSI Momentum (0–10) → weight 15% ─────────────────────
    rsi_status  = str(r.get("rsi_status", "") or "")
    rsi_pos     = str(r.get("rsi_pos", r.get("rsi_position", "")) or "")
    rsi_div     = str(r.get("div_signal", r.get("divergence_signal", "")) or "")

    _rsi_base = {
        ("Strong",    "Above"):  10,
        ("Neutral",   "Above"):   7,
        ("Strong",    "Below"):   6,
        ("Overbought", ""):        4,
        ("Weak",       ""):        2,
        ("Oversold",   ""):        3,
    }
    rsi_score = 0.0
    for (st, pos), pts in _rsi_base.items():
        if rsi_status == st and (pos == "" or pos in rsi_pos):
            rsi_score = float(pts)
            break
    else:
        rsi_score = 3.0  # neutral default
    # Divergence modifier
    if "Bullish" in rsi_div:
        rsi_score = min(10.0, rsi_score + 2.0)
    elif "Bearish" in rsi_div:
        rsi_score = max(0.0, rsi_score - 2.0)

    # ─── COMPONENT 5 – Candle Pattern (0–10) → weight 15% ───────────────────
    candle_pat    = str(r.get("cp_pattern", r.get("last_candle_pattern", "")) or "")
    candle_date   = str(r.get("cp_date", r.get("pattern_date", "")) or "")
    candle_bias   = str(r.get("cp_bias", r.get("candle_bias", "")) or "")

    _bull_candles = {
        "Morning Star": 10, "Bullish Engulfing": 9, "Tweezer Bottoms": 8,
        "Three Inside Up": 8, "Three White Soldiers": 7, "Hammer": 7,
        "Inverted Hammer": 6,
    }
    _bear_penalties = {
        "Evening Star": -3, "Bearish Engulfing": -3,
        "Tweezer Tops": -2, "Three Inside Down": -2,
    }
    candle_score = 0.0
    for pat, pts in _bull_candles.items():
        if pat.lower() in candle_pat.lower():
            candle_score = float(pts)
            break
    for pat, pts in _bear_penalties.items():
        if pat.lower() in candle_pat.lower():
            candle_score += float(pts)
    # Date decay: pattern must be within last 3 trading days (Apr 20–22 2026)
    _pattern_fresh = False
    try:
        from datetime import date as _date
        _ref = _date(2026, 4, 22)
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                _pd = datetime.strptime(candle_date[:10], fmt).date()
                if (_ref - _pd).days <= 3:
                    _pattern_fresh = True
                break
            except Exception:
                continue
    except Exception:
        pass
    if not _pattern_fresh and candle_score > 0:
        candle_score *= 0.5
    candle_score = max(-3.0, min(10.0, candle_score))

    # ─── COMPONENT 6 – Liquidity (0–10) → weight 10% ────────────────────────
    rvol         = safe_num(r.get("rvol20"), np.nan)
    rvol_zone    = str(r.get("rvol20_zone", r.get("rvol_zone", "")) or "")
    liq_cat      = str(r.get("liquidity_category", r.get("liquidity", "")) or "")

    if pd.notna(rvol):
        if rvol > 2.0:
            liq_score = 10.0
        elif rvol >= 1.5:
            liq_score = 9.0
        elif rvol >= 1.0:
            liq_score = 7.0
        elif rvol >= 0.5:
            liq_score = 4.0
        else:
            liq_score = 1.0
    else:
        liq_score = 4.0
    if "Above 1.5" in rvol_zone or "above" in rvol_zone.lower():
        liq_score = min(10.0, liq_score + 1.0)
    elif "Below 1.5" in rvol_zone or "below" in rvol_zone.lower():
        liq_score = max(0.0, liq_score - 1.0)
    if "Very Low" in liq_cat:
        liq_score *= 0.5
    elif "Low" in liq_cat:
        liq_score *= 0.8

    # ─── COMPONENT 7 – Mean Reversion / Extremity (0–10) → informational ────
    w52_pos      = safe_num(r.get("w52_position_pct", r.get("pos_52w_pct")), np.nan)
    rng_comp     = safe_num(r.get("range_compression_20d", r.get("range_compression")), np.nan)
    vol_comp     = safe_num(r.get("vol_compression_20d", r.get("volatility_compression")), np.nan)
    base_len     = safe_num(r.get("base_length_days", r.get("base_length")), np.nan)

    mr_parts = []
    if pd.notna(w52_pos):
        if 20 <= w52_pos <= 80:
            mr_parts.append(10.0)
        elif 10 <= w52_pos < 20 or 80 < w52_pos <= 90:
            mr_parts.append(6.0)
        else:
            mr_parts.append(2.0)
    if pd.notna(rng_comp):
        if rng_comp > 70:
            mr_parts.append(10.0)
        elif rng_comp >= 50:
            mr_parts.append(7.0)
        elif rng_comp >= 30:
            mr_parts.append(4.0)
        else:
            mr_parts.append(1.0)
    if pd.notna(vol_comp):
        if vol_comp > 80:
            mr_parts.append(10.0)
        elif vol_comp >= 60:
            mr_parts.append(7.0)
        elif vol_comp >= 40:
            mr_parts.append(4.0)
        else:
            mr_parts.append(1.0)
    if pd.notna(base_len):
        if 20 <= base_len <= 60:
            mr_parts.append(10.0)
        elif base_len <= 120:
            mr_parts.append(7.0)
        else:
            mr_parts.append(4.0)
    mr_score = sum(mr_parts) / len(mr_parts) if mr_parts else 5.0

    # ─── COMPOSITE (weighted average) ─────────────────────────────────────────
    # Weights: structure 20%, vwap 20%, macd 20%, rsi 15%, candle 15%, liq 10%, mr 0% (informational)
    # mr contributes via a small 10% weight derived from the remaining budget
    composite = (
        0.20 * struct_score +
        0.20 * vwap_score +
        0.20 * macd_score +
        0.15 * rsi_score +
        0.15 * max(0.0, candle_score) +
        0.10 * liq_score +
        0.00 * mr_score          # tracked but not in weighted sum per spec
    )

    # Sector RS-Momentum additive modifier
    sector = str(r.get("sector", "") or "")
    composite += _v17_sector_bias(sector)

    # Reject bearish MACD state
    if _macd_reject:
        composite = min(composite, 5.9)

    composite = max(0.0, min(10.0, round(composite, 2)))

    # Priority labels
    if composite >= 8.0:
        priority = "P1 - EXECUTE"
    elif composite >= 7.0:
        priority = "P2 - PREPARE"
    elif composite >= 6.0:
        priority = "P3 - MONITOR"
    else:
        priority = "BELOW THRESHOLD"

    return (composite, priority)


def _classify_swing_setup(r: dict, zone_type: str, rev_score: int) -> str:
    """
    Classify the swing setup type for quick visual scan.
    Returns one of: Breakout | Bounce | Continuation | Reversal | POI Watch
    """
    ms_regime   = str(r.get("ms_trend_regime", "") or "")
    ms_state    = str(r.get("ms_structure_state", "") or "")
    last_event  = str(r.get("ms_last_event", "") or "")
    event_age   = safe_num(r.get("ms_event_age_d"), np.nan)
    rvol        = safe_num(r.get("rvol20"), np.nan)
    bo_press    = safe_num(r.get("breakout_pressure"), np.nan)
    spring      = str(r.get("_removed_spring", "NO") or "NO")
    rsi14       = safe_num(r.get("rsi14"), np.nan)
    wyckoff     = str(r.get("wyckoff_proxy_phase", "") or "")
    markup_ready = safe_num(r.get("markup_readiness"), np.nan)
    macd_wave   = str(r.get("macd_wave", "") or "")
    rsi_div     = str(r.get("div_signal", "") or "")

    # Breakout: fresh Bull BOS + volume showing up + pressure building
    if (last_event == "Bull BOS"
            and pd.notna(event_age) and event_age <= 7
            and pd.notna(rvol) and rvol >= 1.5
            and pd.notna(bo_press) and bo_press >= 55):
        return "🔥 Breakout"

    # Reversal: Wyckoff accumulation mature OR bullish divergence + CHoCH
    if (wyckoff == "Accumulation"
            and pd.notna(markup_ready) and markup_ready >= 70
            and last_event in ("Bull CHoCH", "Spring")):
        return "⚡ Reversal"
    if "Bullish" in rsi_div and last_event == "Bull CHoCH":
        return "⚡ Reversal"

    # Bounce: price at deep level OR spring candidate OR oversold RSI
    if zone_type in ("PQ -2 SD", "PY -2 SD") or spring == "YES":
        return "📍 Bounce"
    if zone_type in ("PQ -1 SD", "PY -1 SD") and pd.notna(rsi14) and rsi14 <= 45:
        return "📍 Bounce"

    # Continuation: healthy uptrend + constructive pullback to VWAP
    if ("Bullish" in ms_regime
            and ms_state in ("HH-HL", "HH/HL")
            and zone_type in ("PQVWAP", "PYVWAP", "PQ -1 SD", "PY -1 SD")
            and "Mountain Building" in macd_wave):
        return "🔄 Continuation"

    # Reversal fallback: reversal signals strong enough
    if rev_score >= 3 and "Bullish" not in ms_regime:
        return "⚡ Reversal"

    return "👁 POI Watch"


# =========================
# CONVICTION SCORING ENGINE
# =========================

# Global RRG quadrant map — populated by _inject_rrg_sector_state() at runtime
_RRG_SECTOR_QUADRANT: Dict[str, str] = {}

def _inject_rrg_sector_state(rrg_results: list):
    """
    Called from main() after RRG is computed.
    rrg_results = list of dicts with keys: Sector, Quadrant
    Populates _RRG_SECTOR_QUADRANT so conviction scoring can use it.
    Also syncs _RRG_QUADRANT_CACHE so get_rrg_quadrant_for_sector() works.
    """
    global _RRG_SECTOR_QUADRANT, _RRG_QUADRANT_CACHE
    _RRG_SECTOR_QUADRANT = {}
    for row in rrg_results:
        sector = str(row.get("Sector", "") or "").strip()
        quad   = str(row.get("Quadrant", "") or "").strip()
        if sector and quad:
            _RRG_SECTOR_QUADRANT[sector] = quad
            _RRG_QUADRANT_CACHE[sector]  = quad   # ← sync to upgrade-10 cache

def _sector_rrg_modifier(sector: str) -> int:
    """Return score modifier based on live RRG quadrant. 0 if not found."""
    quad = _RRG_SECTOR_QUADRANT.get(str(sector).strip(), "")
    return {"Leading": 2, "Improving": 1, "Weakening": 0, "Lagging": -1}.get(quad, 0)

# V17 RS-Momentum sector bias (±0.5 composite additive, per spec)
_RS_MOMENTUM_MAP: Dict[str, float] = {
    "IDXFINANCE":  0.5,
    "IDXNONCYC":   0.3,
    "IDXTRANS":    0.2,
    "IDXENERGY":  -0.1,
    "IDXBASIC":   -0.1,
    "IDXINDUST":  -0.2,
    "IDXCYCLIC":  -0.3,
    "IDXINFRA":   -0.4,
    "IDXPROPERT": -0.5,
    "IDXHEALTH":  -0.1,
}

def _v17_sector_bias(sector: str) -> float:
    """Return ±0.5 RS-Momentum composite modifier for a sector string."""
    s = str(sector or "").strip().upper()
    # Try exact match first, then prefix match
    if s in _RS_MOMENTUM_MAP:
        return _RS_MOMENTUM_MAP[s]
    for k, v in _RS_MOMENTUM_MAP.items():
        if s.startswith(k) or k in s:
            return v
    # Fall back to live RRG quadrant if RS-Momentum map doesn't cover it
    quad = _RRG_SECTOR_QUADRANT.get(sector.strip(), "")
    return {"Leading": 0.3, "Improving": 0.2, "Weakening": 0.0, "Lagging": -0.2}.get(quad, 0.0)

def _conviction_score(r: dict, tgt_dist: float) -> int:
    """
    Returns integer conviction score 0–10.
    Higher = more evidence the setup is real RIGHT NOW.
    """
    score = 0

    # MA Alignment (0–3)
    ma_tier = r.get("_sc_ma_tier", 0)
    if ma_tier >= 4:
        score += 3
    elif ma_tier == 3:
        score += 2
    elif ma_tier >= 1:
        score += 1

    # RVOL (0–3)
    rvol = safe_num(r.get("rvol20"), 0)
    if pd.notna(rvol) and rvol >= 5.0:
        score += 3
    elif pd.notna(rvol) and rvol >= 3.0:
        score += 2
    elif pd.notna(rvol) and rvol >= 1.5:
        score += 1

    # Last market structure event (0–2)
    last_event = str(r.get("ms_last_event", "") or "")
    if last_event == "Bull BOS":
        score += 2
    elif last_event == "Bull CHoCH":
        score += 1

    # RSI Status (0–1)
    if str(r.get("rsi_status", "") or "") == "Strong":
        score += 1

    # MACD Wave (0–1)
    if "Mountain Building" in str(r.get("macd_wave", "") or ""):
        score += 1

    # Upside % (0–1)
    td = safe_num(tgt_dist, np.nan)
    if pd.notna(td) and td >= 15.0:
        score += 1

    # Sector RRG modifier (−1 to +2)
    score += _sector_rrg_modifier(str(r.get("sector", "") or ""))

    return max(0, min(10, score))


def _poi_candidate_map(r: dict) -> list:
    """Return ranked POI candidates for POI-first swing engine.
    Each candidate = {zone_type, level, dist_pct, poi_type, next_poi, next_px}.
    """
    close_px = safe_num(r.get("close"), np.nan)
    if pd.isna(close_px) or close_px <= 0:
        return []

    level_map = [
        ("PQ -2 SD", "pq_m2", "Quarter VWAP Band", "PQ -1 SD", "pq_m1"),
        ("PQ -1 SD", "pq_m1", "Quarter VWAP Band", "PQVWAP", "pq_vwap"),
        ("PQVWAP", "pq_vwap", "Quarter VWAP Band", "PYVWAP", "py_vwap"),
        ("PY -2 SD", "py_m2", "Yearly VWAP Band", "PY -1 SD", "py_m1"),
        ("PY -1 SD", "py_m1", "Yearly VWAP Band", "PYVWAP", "py_vwap"),
        ("PYVWAP", "py_vwap", "Yearly VWAP Band", "PQVWAP", "pq_vwap"),
    ]
    candidates = []
    for zone_type, key, poi_type, next_lbl, next_key in level_map:
        lvl = safe_num(r.get(key), np.nan)
        if pd.isna(lvl) or lvl == 0:
            continue
        dist = ((close_px / lvl) - 1.0) * 100.0
        absdist = abs(dist)
        # tighter thresholds for SD bands, looser for VWAP anchors
        thr = 1.25 if "SD" in zone_type else 2.0
        if absdist > thr:
            continue
        nxt = safe_num(r.get(next_key), np.nan)
        candidates.append({
            "zone_type": zone_type,
            "poi_type": poi_type,
            "level": lvl,
            "dist_pct": dist,
            "next_poi": next_lbl if pd.notna(nxt) else "N/A",
            "next_px": nxt,
        })

    # fallback context POIs only if no core VWAP-family candidates
    if not candidates:
        for lbl, key in [("IBL", "ibl"), ("PWL", "pwl"), ("MDL", "mdl")]:
            lvl = safe_num(r.get(key), np.nan)
            if pd.isna(lvl) or lvl == 0:
                continue
            dist = ((close_px / lvl) - 1.0) * 100.0
            if abs(dist) <= 1.0:
                candidates.append({
                    "zone_type": lbl,
                    "poi_type": "Profile Level",
                    "level": lvl,
                    "dist_pct": dist,
                    "next_poi": "PQVWAP" if pd.notna(safe_num(r.get("pq_vwap"), np.nan)) else "N/A",
                    "next_px": safe_num(r.get("pq_vwap"), np.nan),
                })
                break

    rank_map = {"PY -2 SD":0, "PQ -2 SD":1, "PY -1 SD":2, "PQ -1 SD":3, "PYVWAP":4, "PQVWAP":5, "IBL":6, "PWL":7, "MDL":8}
    candidates.sort(key=lambda c: (rank_map.get(c["zone_type"], 99), abs(safe_num(c.get("dist_pct"), 99))))
    return candidates


def _derive_amt_state_v2(r: dict, zone_type: str, dist_pct: float) -> str:
    """POI-first AMT state label tuned for long-side swing setups."""
    mp_zone = str(r.get("mp_zone", "") or "").strip()
    last_event = str(r.get("ms_last_event", "") or "")
    spring = str(r.get("_removed_spring", "NO") or "NO")
    div = str(r.get("div_signal", "") or "")
    sd_delta = safe_num(r.get("pq_sd_delta"), np.nan)
    if pd.isna(sd_delta):
        sd_delta = safe_num(r.get("py_sd_delta"), np.nan)

    if spring == "YES" or last_event == "Spring":
        return "Spring reclaim"
    if last_event in ("Bull CHoCH", "Bull BOS") and pd.notna(dist_pct) and dist_pct <= 0:
        return "Lower rejection / re-entry"
    if "Bullish" in div and pd.notna(dist_pct) and dist_pct <= 0:
        return "Responsive buying"
    if pd.notna(sd_delta) and sd_delta > 0 and pd.notna(dist_pct) and dist_pct <= 0:
        return "Re-entering value"
    if zone_type in ("PQ -2 SD", "PY -2 SD"):
        return "Deep discount / responsive zone"
    if mp_zone in ("At Level", "Above") and pd.notna(dist_pct) and abs(dist_pct) <= 1.0:
        return "At lower value edge"
    return "POI monitoring"


def _poi_rank_v2(r: dict, zone_type: str, dist_pct: float, rr: float, rev_score: int) -> str:
    score = 0
    if zone_type in ("PY -2 SD", "PQ -2 SD"):
        score += 3
    elif zone_type in ("PY -1 SD", "PQ -1 SD"):
        score += 2
    elif zone_type in ("PYVWAP", "PQVWAP"):
        score += 1
    if abs(safe_num(dist_pct, 99)) <= 0.75:
        score += 2
    elif abs(safe_num(dist_pct, 99)) <= 1.25:
        score += 1
    if safe_num(rr, 0) >= 2.5:
        score += 2
    elif safe_num(rr, 0) >= 1.8:
        score += 1
    if rev_score >= 3:
        score += 2
    elif rev_score >= 2:
        score += 1
    if score >= 8:
        return "A+"
    if score >= 6:
        return "A"
    if score >= 4:
        return "B"
    return "C"


def _build_idx_vwap_shortlist(rows):
    """
    TRUE POI ENGINE V17 — LONG ONLY
    ---------------------------------
    Applies 10-step V17 scoring framework. A row appears only if:
      Step 1) Structural Quality Filter (all 5 conditions must pass)
      Step 2) Wyckoff Phase Filter (Accumulation or Markup only)
      Step 3–9) Scored on 7 components
      Step 10) Composite >= 6.0 → P1/P2/P3 priority assigned
    Emits _sc_* compatibility fields for the workbook builder.
    """
    board = []

    def _num(v):
        return safe_num(v, np.nan)

    def _fmt_px(lbl, px):
        if not lbl or lbl == "N/A" or pd.isna(safe_num(px, np.nan)):
            return "N/A"
        return f"{lbl}  {safe_num(px, 0):,.0f}"

    def _str(v):
        return str(v or "")

    for r in rows:
        if _str(r.get("data_status")) != "OK":
            continue
        close_px = _num(r.get("close"))
        if pd.isna(close_px) or close_px <= 0:
            continue

        # ── STEP 1: Structural Quality Filter ────────────────────────────────
        ms_quality   = _str(r.get("ms_structure_quality"))
        ms_state     = _str(r.get("ms_structure_state"))
        event_age    = _num(r.get("ms_event_age_d"))
        last_event   = _str(r.get("ms_last_event"))
        trend_bias   = _str(r.get("ms_trend_regime"))
        swing_seq    = _str(r.get("ms_swing_sequence", r.get("ms_structure_state")))

        if ms_quality != "Clean":
            continue
        if ms_state not in ("HH-HL", "HH/HL") and "HH" not in swing_seq:
            continue
        if pd.notna(event_age) and event_age > 20:
            continue
        if "Bullish" not in last_event:
            continue
        if "Bearish" in trend_bias and "Neutral" not in trend_bias:
            continue

        # Wyckoff Phase Filter removed (v2.0)
        # ── Hard tradability gates ─────────────────────────────────────────────
        if _liquidity_check(r) == "Thin":
            continue
        adr = _num(r.get("adr_pct"))
        atr = _num(r.get("atr14_pct"))
        if pd.notna(adr) and adr < 3.0:
            continue
        if pd.notna(atr) and atr < 3.0:
            continue
        if _str(r.get("_removed_upthrust")) == "YES" and _str(r.get("failed_breakout_risk")) == "YES":
            continue

        # ── STEPS 3–9: Build scored row ───────────────────────────────────────
        candidates = _poi_candidate_map(r)
        if not candidates:
            continue
        zone     = candidates[0]
        zone_type = zone["zone_type"]
        dist_pct  = safe_num(zone.get("dist_pct"), np.nan)
        entry_px  = safe_num(zone.get("level"), np.nan)
        tgt_lbl   = zone.get("next_poi", "N/A")
        tgt_px    = safe_num(zone.get("next_px"), np.nan)
        if pd.isna(tgt_px):
            continue
        tgt_dist = ((tgt_px / close_px) - 1.0) * 100.0 if close_px > 0 else np.nan
        if pd.isna(tgt_dist) or tgt_dist < 3.0:
            continue

        # Invalidation
        inv_map = {
            "PQ -2 SD": ("PQ -2 SD", _num(r.get("pq_m2"))),
            "PQ -1 SD": ("PQ -2 SD", _num(r.get("pq_m2"))),
            "PQVWAP":   ("PQ -1 SD", _num(r.get("pq_m1"))),
            "PY -2 SD": ("PY -2 SD", _num(r.get("py_m2"))),
            "PY -1 SD": ("PY -2 SD", _num(r.get("py_m2"))),
            "PYVWAP":   ("PY -1 SD", _num(r.get("py_m1"))),
            "IBL":      ("IBL",      _num(r.get("ibl"))),
            "PWL":      ("PWL",      _num(r.get("pwl"))),
            "MDL":      ("MDL",      _num(r.get("mdl"))),
        }
        inv_lbl, inv_px = inv_map.get(zone_type, (zone_type, entry_px))
        risk_pct = abs(((entry_px / inv_px) - 1.0) * 100.0) \
            if pd.notna(inv_px) and inv_px > 0 and pd.notna(entry_px) else np.nan
        rr_val = safe_num(
            (tgt_dist / risk_pct) if pd.notna(risk_pct) and risk_pct > 0 else np.nan, np.nan
        )
        if pd.isna(rr_val) or rr_val < 1.5:   # V17 minimum R/R
            continue

        rev_score, _rev_signals = _reversal_score(r)
        amt_state  = _derive_amt_state_v2(r, zone_type, dist_pct)
        ma_tier, ma_label = _score_ma_alignment(r)
        above_ibl  = _str(r.get("mp_zone")).strip() in ("At Level", "Above")

        # Confluence text
        conf = []
        if zone_type.startswith("PQ") or zone_type.startswith("PY"):
            conf.append(zone_type.replace("VWAP", "").strip())
        if above_ibl:
            conf.append("IBL")
        if ma_tier >= 2:
            conf.append("EMA20")
        if _str(r.get("div_signal")).startswith("Bullish"):
            conf.append("Bull Div")
        if _str(r.get("smc_bull_ob_1", "")).strip() and close_px:
            conf.append("Bull OB")
        confluence = " + ".join([c for c in conf if c]) or "POI-led"

        # Trigger — V17 specific conditions
        candle_pat = _str(r.get("cp_pattern", r.get("last_candle_pattern")))
        if candle_pat and candle_pat not in ("N/A", ""):
            trigger = f"Candle confirm: {candle_pat} + close above {zone_type}"
        elif ma_tier >= 2:
            trigger = f"EMA10/20 cross with RVOL > 1.0 above {zone_type}"
        else:
            trigger = f"VWAP reclaim with RVOL > 1.0 at {zone_type}"

        # T2: next higher VWAP anchor
        _t2_candidates = [("pq_vwap", "PQVWAP"), ("py_vwap", "PYVWAP"),
                          ("pq_m1", "PQ -1 SD"), ("py_m1", "PY -1 SD")]
        t2_disp = "N/A"
        for _k, _lbl in _t2_candidates:
            _v = _num(r.get(_k))
            if pd.notna(_v) and _v > close_px and _lbl != tgt_lbl:
                t2_disp = _fmt_px(_lbl, _v)
                break

        # Setup type
        if zone_type in ("PQ -2 SD", "PY -2 SD"):
            setup_type = "📍 Deep Discount"
        elif zone_type in ("PQ -1 SD", "PY -1 SD"):
            setup_type = "📍 Discount Pullback"
        elif "Bull BOS" in last_event and ma_tier >= 2:
            setup_type = "🔄 Continuation"
        elif rev_score >= 2 or amt_state in ("Spring reclaim", "Responsive buying"):
            setup_type = "⚡ Reversal"
        else:
            setup_type = "👁 POI Watch"

        # ── STEP 9: Composite scoring ──────────────────────────────────────────
        rr = dict(r)
        rr["sector"] = _str(r.get("sector") or r.get("IDX Sector") or "N/A")
        # Stage intermediate fields so _swing_score can read them
        rr["ms_trend_regime"]     = trend_bias
        rr["ms_structure_state"]  = ms_state
        rr["ms_last_event"]       = last_event
        rr["ms_event_age_d"]      = event_age
        rr["ms_structure_quality"]= ms_quality

        sw_score, priority_label = _swing_score(rr, zone_type, tgt_dist)

        # ── STEP 10: Priority assignment ──────────────────────────────────────
        if sw_score < 6.0:
            continue   # Below minimum threshold

        # buy_signal: YES if all strong, else WATCH
        ready = (
            ma_tier >= 2 and above_ibl and
            ("Bearish" not in trend_bias or zone_type in ("PQ -2 SD", "PY -2 SD")) and
            rev_score >= 1
        )
        buy_signal = "YES" if ready else "WATCH"
        if ma_tier == 0:
            buy_signal = "WATCH"

        poi_rank_v = _poi_rank_v2(r, zone_type, dist_pct, rr_val, rev_score)

        # Brief actionable summary
        cause_quality  = _str(r.get("cause_quality", ""))
        markup_ready   = _num(r.get("markup_readiness"))
        summary = (
            f"{zone_type} | {setup_type.split()[-1]} | "
            f"RR {rr_val:.1f} | {priority_label.split(' - ')[-1].title()}"
        )
        if cause_quality:
            summary += f" | Cause: {cause_quality}"

        # Write all _sc_* fields
        rr["_sc_ma_path"]            = "PASS" if ma_tier >= 1 else "EXCEPTION"
        rr["_sc_ma_tier"]            = ma_tier
        rr["_sc_ma_label"]           = ma_label
        rr["_sc_zone_type"]          = zone_type
        rr["_sc_dist_pct"]           = dist_pct
        rr["_sc_entry_disp"]         = _fmt_px(zone_type, entry_px)
        rr["_sc_target_lbl"]         = tgt_lbl
        rr["_sc_target_px"]          = tgt_px
        rr["_sc_target_dist"]        = tgt_dist
        rr["_sc_target_disp"]        = _fmt_px(tgt_lbl, tgt_px)
        rr["_sc_t2_disp"]            = t2_disp
        rr["_sc_invalidation_disp"]  = _fmt_px(f"< {inv_lbl}", inv_px)
        rr["_sc_trigger"]            = trigger
        rr["_sc_signal"]             = amt_state
        rr["_sc_buy_signal"]         = buy_signal
        rr["_sc_trade_bias"]         = "Long" if buy_signal == "YES" else "Watch"
        rr["_sc_rev_score"]          = rev_score
        rr["_sc_setup_type"]         = setup_type
        rr["_sc_poi_type"]           = zone.get("poi_type", "POI")
        rr["_sc_primary_poi"]        = zone_type
        rr["_sc_next_poi"]           = tgt_lbl
        rr["_sc_confluence"]         = confluence
        rr["_sc_amt_state"]          = amt_state
        rr["_sc_summary"]            = summary
        rr["_sc_rr"]                 = round(rr_val, 1)
        rr["_sc_poi_rank"]           = poi_rank_v
        rr["_sc_conviction"]         = sw_score
        rr["_sc_swing_score"]        = sw_score
        rr["_sc_swing_grade"]        = priority_label
        rr["_sc_priority_label"]     = priority_label
        rr["_sc_priority"]           = "YES" if priority_label == "P1 - EXECUTE" else ""
        rr["_sc_cause_quality"]      = cause_quality
        rr["_sc_markup_readiness"]   = markup_ready

        # Sort key: P1 < P2 < P3, then score desc, then RVOL desc
        _p_rank = {"P1 - EXECUTE": 0, "P2 - PREPARE": 1, "P3 - MONITOR": 2}.get(priority_label, 3)
        rr["_sc_sort_key"] = (
            _p_rank,
            -sw_score,
            -safe_num(r.get("rvol20"), 0),
            -safe_num(markup_ready, 0),
            rr.get("ticker", "")
        )
        board.append(rr)

    board.sort(key=lambda x: x.get("_sc_sort_key", (9, 0, 0, 0, "")))
    return board


def _market_profile_summary(r: dict) -> str:
    """Compact market-profile context for the screener sheet."""
    parts = []
    mp = str(r.get("mp_zone", "") or "").strip()
    pwl = str(r.get("price_ge_pwl", "") or "").strip()
    mdl = str(r.get("price_ge_mdl", "") or "").strip()
    if mp and mp != "N/A":
        parts.append(f"IBL {mp}")
    if pwl and pwl != "N/A":
        parts.append(f"PWL {pwl}")
    if mdl and mdl != "N/A":
        parts.append(f"MDL {mdl}")
    return " | ".join(parts) if parts else "N/A"



# =========================
# IDX SCREENER SUMMARY SHEET  (replaces Swing Watchlist)
# Shows top-scored POI candidates in a compact executive dashboard
# =========================
def build_summary_sheet(ws, latest_market_day: str, summary_rows: list):
    """
    IDX Screener Summary — Priority dashboard replacing Swing Watchlist.
    Columns: Rank | Ticker | Emiten | Sector | Regime | Close | Score | MA Zone |
             SMC State | MACD | QVWAP Zone | PQ Zone | PY Zone | Signal | RR | Notes
    """
    ws.title = "IDX Screener Summary"
    ws.column_dimensions["A"].width = 0.5

    for r_h, h in {1: 4, 2: 26, 3: 16, 4: 4, 5: 22, 6: 36}.items():
        ws.row_dimensions[r_h].height = h

    pretty = ""
    try:
        pretty = datetime.strptime(latest_market_day, "%Y-%m-%d").strftime("%B %d, %Y")
    except Exception:
        pretty = str(latest_market_day)

    ws.merge_cells("B2:P2")
    c = ws["B2"]
    c.value = "IDX Screener Summary  ·  SMC + MACD + VWAP Priority Board"
    style_plain(c, font=FONT_TITLE, align="left")

    ws["B3"] = (
        f"As of {pretty}  ·  "
        f"Filter: Score ≥ 6.0  ·  Sorted: Priority → Score ↓  ·  "
        f"Long Only  ·  v2.0"
    )
    style_plain(ws["B3"], font=FONT_SUBTITLE, align="left")

    # Column schema
    COLS = [
        ("Rank",        6,  "center"),
        ("Ticker",      9,  "center"),
        ("Emiten",      34, "left"),
        ("Sector",      20, "left"),
        ("Regime",      22, "center"),
        ("Close",       10, "center"),
        ("Score",       9,  "center"),
        ("MA Zone",     22, "center"),
        ("SMC State",   22, "center"),
        ("MACD",        24, "center"),
        ("QVWAP Zone",  28, "left"),
        ("PQ Zone",     28, "left"),
        ("Signal",      22, "center"),
        ("RR",          8,  "center"),
        ("Priority",    12, "center"),
    ]

    hdr_row = 6
    for col_i, (hdr, w, _align) in enumerate(COLS, start=2):
        c = ws.cell(hdr_row, col_i)
        c.value = hdr
        style_cell(c, fill=FILL_HEADER, font=FONT_HEADER, align="center")
        ws.column_dimensions[get_column_letter(col_i)].width = w

    data_start = hdr_row + 1
    r = data_start

    SIGNAL_COLORS = {
        "P1 - EXECUTE": "10B981",   # emerald
        "P2 - PREPARE": "3B82F6",   # blue
        "P3 - MONITOR": "A78BFA",   # violet
    }
    REGIME_COLORS = {
        "Blue Chip / High Liquidity":  "2A4A66",
        "Institutional Driven":        "237A5C",
        "Mid Cap / Moderate":          "21467C",
        "Low Liquidity / Small Cap":   "7A1A1A",
    }

    displayed = 0
    for row in summary_rows:
        if row.get("data_status") not in ("OK",):
            continue
        score = safe_num(row.get("_sc_swing_score") or row.get("_sc_conviction"), 0)
        if score < 6.0:
            continue

        displayed += 1
        row_fill = PatternFill("solid", fgColor="F0FDF4" if displayed % 2 == 0 else "F8FAFC")
        priority_lbl = str(row.get("_sc_priority_label") or row.get("_sc_swing_grade") or "")
        pri_color = SIGNAL_COLORS.get(priority_lbl, "1F1F1F")

        vals = [
            displayed,
            row.get("ticker", ""),
            row.get("emiten", ""),
            str(row.get("idx_sector") or row.get("sector") or ""),
            row.get("stock_regime", "N/A"),
            safe_num(row.get("close"), np.nan),
            round(score, 1),
            row.get("summary_ma", "N/A"),
            str(row.get("smc_composite_state") or row.get("smc_state") or "N/A"),
            str(row.get("macd_composite") or row.get("macd_entry") or ""),
            str(row.get("q_remarks") or ""),
            str(row.get("pq_remarks") or ""),
            str(row.get("_sc_signal") or row.get("_sc_buy_signal") or ""),
            round(safe_num(row.get("_sc_rr"), 0), 1) if pd.notna(safe_num(row.get("_sc_rr"), np.nan)) else "N/A",
            priority_lbl,
        ]

        for col_i, (val, (_, w, align)) in enumerate(zip(vals, COLS), start=2):
            c = ws.cell(r, col_i)
            c.value = val
            style_cell(c, fill=row_fill, font=FONT_BODY, align=align)

        # Colour priority cell
        pri_cell = ws.cell(r, 2 + 14)
        pri_cell.fill = PatternFill("solid", fgColor=pri_color) if priority_lbl else row_fill
        pri_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        # Colour regime cell
        reg_cell = ws.cell(r, 2 + 4)
        reg_str = str(row.get("stock_regime", ""))
        if reg_str in REGIME_COLORS:
            reg_cell.fill = PatternFill("solid", fgColor=REGIME_COLORS[reg_str])
            reg_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        ws.row_dimensions[r].height = 16
        r += 1

    if displayed == 0:
        ws.cell(data_start, 2).value = "No qualifying rows (Score ≥ 6.0) found in this run."
        style_plain(ws.cell(data_start, 2), font=FONT_SUBTITLE, align="left")
        r += 1

    ws.freeze_panes = f"B{hdr_row + 1}"
    if r > data_start:
        ws.auto_filter.ref = (
            f"B{hdr_row}:{get_column_letter(2 + len(COLS) - 1)}{r - 1}"
        )
    ws.sheet_view.showGridLines = False

def build_processing_sheet(ws, latest_market_day, total_scanned, logs, summary_rows):
    ws.title = "Data Processing Results"
    widths = {"A": 0.5, "B": 20, "C": 20, "D": 10, "E": 12, "F": 10, "G": 14, "H": 70}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w

    ws["B2"] = "Data Processing Results"
    style_plain(ws["B2"], font=FONT_TITLE, align="left")

    run_dt_obj = datetime.now()
    ok_count      = sum(1 for x in logs if x["status"] == "OK")
    partial_count = sum(1 for x in logs if x["status"] == "PARTIAL DATA")
    nodata_count  = sum(1 for x in logs if x["status"] == "NO DATA")

    # ── Screener stats (read from global _RRG_SECTOR_QUADRANT if available) ──
    _leading_sectors   = [s for s, q in _RRG_SECTOR_QUADRANT.items() if q == "Leading"]
    _improving_sectors = [s for s, q in _RRG_SECTOR_QUADRANT.items() if q == "Improving"]

    FILL_STAT_HDR = PatternFill("solid", fgColor="2A4A66")
    FILL_STAT_OK  = PatternFill("solid", fgColor="D1FAE5")
    FILL_STAT_WRN = PatternFill("solid", fgColor="FEF3C7")
    FILL_STAT_PRI = PatternFill("solid", fgColor="FEF08A")

    def _stat_row(r, label, value, fill=None):
        ws.cell(r, 2).value = label
        ws.cell(r, 2).font  = Font(name="Calibri", size=10, bold=True)
        ws.cell(r, 2).fill  = fill or PatternFill("solid", fgColor="F8FAFC")
        ws.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(r, 2).border = BORDER
        ws.cell(r, 3).value = value
        if isinstance(value, datetime):
            ws.cell(r, 3).number_format = "dd mmm yyyy hh:mm"
        elif hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
            ws.cell(r, 3).number_format = DATE_FORMAT_EXCEL
        ws.cell(r, 3).font  = Font(name="Calibri", size=10)
        ws.cell(r, 3).fill  = fill or PatternFill("solid", fgColor="FFFFFF")
        ws.cell(r, 3).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(r, 3).border = BORDER

    # Section header
    ws.merge_cells("B3:H3")
    hdr = ws["B3"]
    hdr.value = "Run Summary"
    hdr.font  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    hdr.fill  = FILL_STAT_HDR
    hdr.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 22

    stat_rows = [
        ("Run Time",              run_dt_obj,                               None),
        ("As-of Date",            pd.Timestamp(latest_market_day).date() if latest_market_day else "", None),
        ("Total Tickers Scanned", int(total_scanned),                       None),
        ("OK  (full data)",       int(ok_count),                            FILL_STAT_OK),
        ("Partial Data",          int(partial_count),                       FILL_STAT_WRN),
        ("No Data",               int(nodata_count),                        None),
        ("",                      "",                                        None),
    ]

    r = 4
    for label, value, fill in stat_rows:
        _stat_row(r, label, value, fill)
        ws.row_dimensions[r].height = 20
        r += 1

    # Legacy screener stat block removed per current output spec
    r += 1  # spacer

    # ── Ticker log table ──────────────────────────────────────────────────────
    header_row = r
    headers = ["Ticker", "Status", "Bars", "Retry Count", "Source Used", "Latest Market Day", "Reason"]
    for i, h in enumerate(headers, start=2):
        c = ws.cell(header_row, i)
        c.value = h
        style_cell(c, fill=FILL_HEADER_SUMMARY, font=FONT_SUBHEADER, align="center", wrap=True)
    ws.row_dimensions[header_row].height = 28
    r = header_row + 1

    for log in logs:
        vals = [
            log["ticker"], log["status"], log["bars"], log["retry_count"],
            log["source_used"], log["latest_market_day"], log["reason"]
        ]
        for c_idx, v in enumerate(vals, start=2):
            c = ws.cell(r, c_idx)
            if headers[c_idx - 2] == "Latest Market Day" and v not in (None, "", "-", "N/A"):
                try:
                    c.value = pd.Timestamp(v).date()
                    c.number_format = DATE_FORMAT_EXCEL
                except Exception:
                    c.value = v
            else:
                c.value = v
            align = "left" if c_idx in [2, 3, 6, 8] else "center"
            row_fill = None
            if log["status"] == "OK":
                row_fill = PatternFill("solid", fgColor="F0FDF4")
            elif log["status"] == "PARTIAL DATA":
                row_fill = PatternFill("solid", fgColor="FFFBEB")
            style_cell(c, fill=row_fill, font=FONT_BODY, align=align, wrap=(c_idx == 8))
        ws.row_dimensions[r].height = 22
        r += 1

    ws.freeze_panes = None
    ws.auto_filter.ref = f"B{header_row}:H{max(header_row, ws.max_row)}"
    ws.sheet_view.showGridLines = False


def build_detail_sheet(ws, latest_market_day, rows):
    ws.title = "IDX Technical Detail"

    ws.column_dimensions["A"].width = 0.5

    for r, h in {1: 5, 2: 24, 3: 18, 4: 5, 5: 24, 6: 44, 7: 20, 8: 20, 9: 20, 10: 5, 11: 24, 12: 48}.items():
        ws.row_dimensions[r].height = h

    ws["B2"] = "IDX Technical Detail" + (" [BACKTEST MODE]" if BACKTEST_MODE else "")
    style_plain(ws["B2"], font=FONT_TITLE, align="left")
    _asof_ts = get_effective_asof_date()
    pretty = _asof_ts.strftime("%B %d, %Y")
    _mode_lbl = f"BACKTEST MODE as of {pretty}" if BACKTEST_MODE else f"Market Data as of {pretty}"
    ws["B3"] = _mode_lbl
    style_plain(ws["B3"], font=FONT_SUBTITLE, align="left")

    for _c in range(2, 250):
        ws.cell(4, _c).value = None

    # Flatten schema and assign columns dynamically
    flat_cols = []
    col_idx = 2  # start at B
    group_ranges = []
    for group_name, group_fill, cols in DETAIL_SCHEMA:
        start_idx = col_idx
        for key, header, width, fmt, align in cols:
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].width = width
            flat_cols.append((col_idx, letter, group_name, group_fill, key, header, width, fmt, align))
            col_idx += 1
        end_idx = col_idx - 1
        group_ranges.append((group_name, group_fill, start_idx, end_idx))

    # Group headers row 5
    for group_name, group_fill, start_idx, end_idx in group_ranges:
        start_cell = f"{get_column_letter(start_idx)}5"
        end_cell = f"{get_column_letter(end_idx)}5"
        ws.merge_cells(f"{start_cell}:{end_cell}")
        for c in range(start_idx, end_idx + 1):
            cell = ws.cell(5, c)
            cell.fill = group_fill
            cell.font = FONT_GROUP
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
        display_group_name = group_name
        _dt_lmd = datetime.strptime(latest_market_day, "%Y-%m-%d")
        _MONTH_NAMES = ["January","February","March","April","May","June",
                        "July","August","September","October","November","December"]
        if group_name == "Current MVWAP (June 2026)":
            display_group_name = f"Current MVWAP ({_MONTH_NAMES[_dt_lmd.month - 1]} {_dt_lmd.year})"
        elif group_name == "Previous MVWAP (May 2026)":
            _pm_end = pd.Timestamp(_dt_lmd.year, _dt_lmd.month, 1) - pd.Timedelta(days=1)
            display_group_name = f"Previous MVWAP ({_MONTH_NAMES[_pm_end.month - 1]} {_pm_end.year})"
        elif group_name == "Current QVWAP (Q2 2026)":
            q = ((_dt_lmd.month - 1) // 3) + 1
            display_group_name = f"Current QVWAP (Q{q} {_dt_lmd.year})"
        elif group_name == "Previous QVWAP (Q1 2026)":
            prev_end = quarter_start(pd.Timestamp(_dt_lmd)) - pd.Timedelta(days=1)
            pq = ((prev_end.month - 1) // 3) + 1
            display_group_name = f"Previous QVWAP (Q{pq} {prev_end.year})"
        elif group_name == "Previous Year VWAP (2025)":
            display_group_name = f"Previous Year VWAP ({_dt_lmd.year - 1})"
        ws[start_cell] = display_group_name

    # Subheaders row 6
    for col_idx, letter, group_name, group_fill, key, header, width, fmt, align in flat_cols:
        c = ws.cell(6, col_idx)
        c.value = header
        style_cell(c, fill=group_fill, font=FONT_SUBHEADER, align="center", wrap=True)

    # Data rows
    for r_idx, r in enumerate(rows, start=7):
        row_fill = None

        for col_idx, letter, group_name, group_fill, key, header, width, fmt, align in flat_cols:
            c = ws.cell(r_idx, col_idx)
            val = _market_profile_summary(r) if key == "mp_summary" else r.get(key, "")
            sd_keys = ["cm_sd","cm_delta","pm_sd","pm_delta","q_sd","q_delta","pq_sd","pq_delta","py_sd","py_delta"]

            if key in ["mcap", "avg30", "last_vol", "foreign_net_val", "max_position_idr"] and pd.notna(val):
                val = compact_fmt(val)
            elif isinstance(val, float) and pd.isna(val):
                val = None if key in sd_keys else "-"
            # Force compact display BEFORE writing to Excel for Liquidity numeric columns
            if key in ["lot", "daily_value", "adtr20", "last_vol", "adtv20", "foreign_net_val", "max_position_idr"] and val not in ("-", "N/A") and pd.notna(val):
                val = compact_fmt(val)
            elif isinstance(val, float) and pd.isna(val):
                val = None if key in sd_keys else "-"

            # Normalize sentinel strings to "-" for display (keep internal "None" sentinel for logic)
            if val in ("N/A", "None") and key not in sd_keys:
                val = "-"

            # Keep SD metrics as TRUE numeric cells (not text) so Excel won't show "Number Stored as Text"
            if key in ["cm_sd","cm_delta","pm_sd","pm_delta","q_sd","q_delta","pq_sd","pq_delta","py_sd","py_delta"]:
                if isinstance(val, (int, float)) and pd.notna(val):
                    val = float(val)
                elif isinstance(val, float) and pd.isna(val):
                    val = None

            c.value = val
            fill_to_use = row_fill if row_fill else None
            style_cell(c, fill=fill_to_use, font=FONT_BODY, align=align, wrap=(key in ["sector","industry","cm_remarks","pm_remarks","q_remarks","pq_remarks","py_remarks","af_execution_guide","af_watch_next","af_invalidation","preset_summary"]))
            if key == "pct_change":
                try:
                    _pct_val = safe_num(r.get("pct_change"), np.nan)
                    if pd.notna(_pct_val):
                        c.number_format = "0.00%"
                        if _pct_val > 0:
                            c.font = Font(name=FONT_BODY.name, size=FONT_BODY.size, bold=False, italic=False, color="008000")
                        elif _pct_val < 0:
                            c.font = Font(name=FONT_BODY.name, size=FONT_BODY.size, bold=False, italic=False, color="C00000")
                except Exception:
                    pass

            if fmt and val != "N/A":


                if fmt == "compact":
                    pass
                elif fmt == "date_wyckoff":
                    try:
                        if value in (None, "", "N/A"):
                            pass
                        else:
                            if isinstance(value, str):
                                parsed = pd.to_datetime(value, errors="coerce", format="mixed")
                                if pd.notna(parsed):
                                    cell.value = parsed.to_pydatetime()
                            elif hasattr(value, "to_pydatetime"):
                                cell.value = value.to_pydatetime()
                            cell.number_format = "dd mmm 'yy"
                    except Exception:
                        pass
                elif fmt == "wyckoff_date":
                    try:
                        if isinstance(value, str):
                            parsed = pd.to_datetime(value, errors="coerce", format="mixed")
                            if pd.notna(parsed):
                                cell.value = parsed.to_pydatetime()
                        elif hasattr(value, "to_pydatetime"):
                            cell.value = value.to_pydatetime()
                        cell.number_format = "dd mmm 'yy"
                    except Exception:
                        pass
                else:
                    if isinstance(val, (int, float)) and not pd.isna(val):
                        c.number_format = fmt

        ws.row_dimensions[r_idx].height = 26

    last_col_letter = get_column_letter(flat_cols[-1][0])
    ws.freeze_panes = "E7"
    ws.auto_filter.ref = f"B6:{last_col_letter}{max(6, ws.max_row)}"
    ws.sheet_view.showGridLines = False

# Build guide delegated to new comprehensive implementation below
def _safe_sector_bucket(v):
    s = str(v).strip() if v is not None else ""
    return s if s else "Unknown"

def _safe_num_series(df, candidates, default=np.nan):
    for c in candidates:
        if c in df.columns:
            return pd.to_numeric(df[c], errors="coerce")
    return pd.Series([default] * len(df), index=df.index, dtype=float)

def _build_sector_trail_points(x_final, y_final, trail=5):
    # deterministic synthetic trail from current state (real data snapshot, chart-friendly)
    pts = []
    for i in range(trail):
        frac = (trail - 1 - i) / max(1, (trail - 1))
        x = float(np.clip(x_final * (1.0 - 0.22 * frac), -100, 100))
        y = float(np.clip(y_final * (1.0 - 0.28 * frac), -100, 100))
        pts.append((x, y))
    return pts


def build_priority_sheet(wb, latest_market_day: str, ordered_rows: list):
    """
    Priority tab — today's 8–20 highest-conviction actionable names only.

    Eligibility: _sc_priority == "YES"
    If fewer than 5 qualify, fills with next-best YES rows by score.
    Sorted by conviction score desc, then RVOL desc.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("Priority")
    ws.sheet_view.showGridLines = False

    def _F(h): return PatternFill("solid", fgColor=h)

    GRP_HDR  = _F("7C1D1D")     # dark red — high urgency
    FILL_PRI = _F("FEF08A")     # amber highlight for priority badge
    FILL_YES = _F("D1FAE5")
    FILL_SCR = {
        (8, 10): _F("D1FAE5"),
        (7,  7): _F("DCFCE7"),
        (5,  6): _F("FEF9C4"),
        (0,  4): _F("FEE2E2"),
    }
    FONT_HDR  = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_BODY = Font(name="Calibri", size=10)
    FONT_BOLD = Font(name="Calibri", size=10, bold=True)
    THIN = Side(style="thin", color="D1D5DB")
    BDR  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def _score_fill(score):
        for (lo, hi), fill in FILL_SCR.items():
            if lo <= score <= hi:
                return fill
        return _F("FFFFFF")

    # ── Header ───────────────────────────────────────────────────────────────
    pretty = datetime.strptime(latest_market_day, "%Y-%m-%d").strftime("%B %d, %Y")
    ws.column_dimensions["A"].width = 0.5
    for r, h in {1: 4, 2: 28, 3: 14, 4: 4, 5: 22, 6: 36}.items():
        ws.row_dimensions[r].height = h

    ws["B2"] = "★  Priority Setups  ·  Today's Actionable Shortlist"
    ws["B2"].font  = Font(name="Calibri", size=14, bold=True)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")

    ws["B3"] = (
        f"As of {pretty}  ·  "
        "Score ≥ 7  ·  Buy = YES  ·  Upside > 0%  ·  No negative target  ·  "
        "Sorted by Conviction Score ↓ then RVOL ↓"
    )
    ws["B3"].font      = Font(name="Calibri", size=10, italic=True)
    ws["B3"].alignment = Alignment(horizontal="left", vertical="center")

    # ── Column schema ─────────────────────────────────────────────────────────
    COLS = [
        # (key, header, width, align)
        ("ticker",              "Ticker",       9,  "center"),
        ("sector",              "Sector",       18, "center"),
        ("close",               "Price",        12, "center"),
        ("pct_change",          "Chg %",        8,  "center"),
        ("rvol20",              "RVOL",         8,  "center"),
        ("_sc_conviction",      "Score /10",    9,  "center"),
        ("_sc_ma_label",        "MA Position",  24, "left"),
        ("_sc_zone_type",       "Anchor",       12, "center"),
        ("_sc_entry_disp",      "Entry",        20, "left"),
        ("_sc_target_disp",     "Target",       18, "left"),
        ("_sc_target_dist",     "Upside %",     9,  "center"),
        ("_sc_invalidation_disp","Invalidation",22, "left"),
        ("ms_trend_regime",     "Regime",       14, "center"),
        ("ms_last_event",       "Last Event",   14, "center"),
        ("rsi14",               "RSI",          8,  "center"),
        ("rsi_status",          "RSI Status",   10, "center"),
        ("macd_wave",           "MACD Wave",    20, "center"),
        ("_sc_signal",          "Signals",      32, "left"),
        ("_sc_buy_signal",      "Buy",          8,  "center"),
    ]

    # ── Column headers ────────────────────────────────────────────────────────
    # Group header row 5
    ws.merge_cells(f"B5:{get_column_letter(len(COLS)+1)}5")
    hc = ws["B5"]
    hc.value = "Priority Setups — High-Conviction Execution List"
    hc.fill  = GRP_HDR
    hc.font  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    hc.alignment = Alignment(horizontal="center", vertical="center")
    for ci in range(2, len(COLS) + 2):
        ws.cell(5, ci).fill = GRP_HDR
        ws.cell(5, ci).border = BDR

    # Sub-header row 6
    for ci, (key, header, width, align) in enumerate(COLS, start=2):
        ws.column_dimensions[get_column_letter(ci)].width = width
        c = ws.cell(6, ci)
        c.value = header
        c.fill  = GRP_HDR
        c.font  = FONT_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BDR

    # ── Select rows ───────────────────────────────────────────────────────────
    priority_rows = [r for r in ordered_rows if r.get("_sc_priority") == "YES"]
    # If too few, pad with next-best YES rows by score
    if len(priority_rows) < 5:
        extras = [
            r for r in ordered_rows
            if r.get("_sc_buy_signal") == "YES"
            and r.get("_sc_priority") != "YES"
            and safe_num(r.get("_sc_target_dist"), -1) > 0
        ]
        extras.sort(key=lambda x: (-safe_num(x.get("_sc_conviction", 0), 0),
                                    -safe_num(x.get("rvol20", 0), 0)))
        priority_rows = priority_rows + extras[: max(0, 20 - len(priority_rows))]

    # Hard cap at 30
    priority_rows = priority_rows[:30]

    if not priority_rows:
        ws.cell(8, 2).value = "No qualifying setups today. Review the IDX Screener Summary for WATCH candidates."
        ws.cell(8, 2).font  = Font(name="Calibri", size=10, italic=True, color="6B7280")
        return

    # ── Data rows ─────────────────────────────────────────────────────────────
    REGIME_FILLS = {
        "Bullish":      _F("D1FAE5"), "Bullish Weak": _F("DCFCE7"),
        "Transition Up":_F("DBEAFE"), "Sideways":     _F("F3F4F6"),
        "Range":        _F("F3F4F6"), "Transition Down":_F("FEF3C7"),
        "Bearish Weak": _F("FEE2E2"), "Bearish":      _F("FECACA"),
    }

    for r_idx, r in enumerate(priority_rows, start=7):
        is_priority = r.get("_sc_priority") == "YES"
        for ci, (key, header, width, align) in enumerate(COLS, start=2):
            c  = ws.cell(r_idx, ci)
            v  = r.get(key, "N/A")

            # Value formatting
            if key == "close":
                nv = safe_num(v, np.nan)
                c.value = nv if pd.notna(nv) else "N/A"
                if pd.notna(nv): c.number_format = "#,##0"
            elif key == "pct_change":
                nv = safe_num(v, np.nan)
                c.value = nv if pd.notna(nv) else "N/A"
                if pd.notna(nv):
                    c.number_format = "0.00%"
                    c.font = Font(name="Calibri", size=10,
                                  color="008000" if nv > 0 else ("C00000" if nv < 0 else "111827"))
            elif key in ("rvol20", "rsi14"):
                nv = safe_num(v, np.nan)
                c.value = round(float(nv), 2) if pd.notna(nv) else "N/A"
                if pd.notna(nv): c.number_format = "0.00"
            elif key == "_sc_target_dist":
                nv = safe_num(v, np.nan)
                c.value = round(float(nv), 1) if pd.notna(nv) else "N/A"
                if pd.notna(nv): c.number_format = "+0.0;-0.0"
            else:
                c.value = v if v not in (None, "", np.nan) else "N/A"
                if isinstance(v, float) and pd.isna(v):
                    c.value = "N/A"

            # Base styling
            row_bg = _F("FFFBEB") if is_priority else _F("F8FAFC")
            c.fill   = row_bg
            c.font   = FONT_BODY
            c.border = BDR
            c.alignment = Alignment(horizontal=align, vertical="center",
                                    wrap_text=(key in ("_sc_entry_disp","_sc_target_disp",
                                                        "_sc_invalidation_disp","_sc_signal",
                                                        "_sc_ma_label")))

            # Cell-level overrides
            v_str = str(c.value or "")
            if key == "_sc_conviction":
                score = int(safe_num(v, 0))
                c.fill = _score_fill(score)
                c.font = Font(name="Calibri", size=10, bold=(score >= 7),
                              color="065F46" if score >= 7 else ("713F12" if score >= 5 else "991B1B"))
            elif key == "ms_trend_regime":
                c.fill = REGIME_FILLS.get(v_str, row_bg)
            elif key == "_sc_buy_signal":
                c.fill = FILL_YES if v_str == "YES" else _F("FEF3C7")
                c.font = Font(name="Calibri", size=10, bold=True,
                              color="065F46" if v_str == "YES" else "854D0E")
            elif key == "_sc_zone_type":
                if "-2 SD" in v_str:   c.fill = _F("D1FAE5")
                elif "-1 SD" in v_str: c.fill = _F("FEF9C4")
                elif "VWAP" in v_str:  c.fill = _F("DBEAFE")

    # ── Summary line ──────────────────────────────────────────────────────────
    n_tot  = len(priority_rows)
    summary_row = 7 + n_tot + 1
    ws.cell(summary_row, 2).value = (
        f"{n_tot} total shown  ·  See IDX Screener Summary for full list"
    )
    ws.cell(summary_row, 2).font = Font(name="Calibri", size=9, italic=True, color="6B7280")
    ws.freeze_panes = "F7"


# =========================
# TICKER UNIVERSE CONFIG
# =========================
# Live/default mode = FULL KSEI universe (all valid tickers from raw source).
# Set TRUE only if you explicitly want to run the legacy custom subset.
USE_CUSTOM_TICKERS_ONLY = False

# Legacy custom subset (kept for optional use only)
CUSTOM_TICKERS = [
    "MBMA","MAPI","ANJT","DEPO","NCKL","RMKE","KIJA","BCAP","SMAR","BANK",
    "BREN","NICL","BSIM","MBAP","ASLC","DOID","PANI","EMTK","BBKP","SMGR",
    "SMBR","KPIG","AMAR","BRPT","TKIM","BMAS","DSSA","SSIA","TPIA","AMAG",
    "HOKI","ARTO","AMRT","KRAS","BBHI","AVIA","AGRO","LPCK","WIFI","NRCA",
    "PACK","DGWG","PIPA","BACA","MORA","PNLF","SAFE","TMAS","SGER","BLUE",
    "KLIN","WIRG","EMDE","VICI","HDIT","PJAA","BGTG","DMND","PORT","CYBR",
    "FUTR","AXIO","DKHH","FUJI","MBSS","ZYRX","MLPT","IPAC","INOV","CANI",
    "GPRA","SMIL","BULL","NIKL","TRST","RSCH","PALM","NICK","CAKK","ISAP",
    "PART","TAXI","ASLI","INTD","LEAD","KLAS","LUCK","INET","NICE","OLIV",
    "DOSS","AKSI","DOOH","CBRE","SOTS","DEWI","BIPP","DEFI","KOPI","MANG",
    "BAYU"
]


# =========================
# FINAL SHEET ORDER HELPER
# =========================
DESIRED_SHEET_ORDER = [
    "IDX Overview",
    "IDX Screener",
    "IDX Technical Detail",
    "IDX Fundamental Detail",
    "IDX News",
    "Guide & Logic Reference",
    "Data Processing Results",
    "QA Calculation Audit",
]


def _excel_safe_sheet_name(name: str) -> str:
    return str(name or "Sheet").replace(" ", "_").replace("&", "and").replace("-", "_")[:24]

def _add_excel_table_if_possible(ws, table_name: str, header_row: int, start_col: int = 2):
    """
    BUG-01 FIX: openpyxl formal Table objects produce malformed OOXML that triggers
    Excel repair dialog on every open-save cycle. Use plain auto_filter.ref instead.
    """
    try:
        if ws.max_row <= header_row or ws.max_column < start_col:
            return False
        end_col = ws.max_column
        for col in range(ws.max_column, start_col - 1, -1):
            if ws.cell(header_row, col).value not in (None, ""):
                end_col = col
                break
        if end_col < start_col:
            return False
        from openpyxl.utils import get_column_letter as _gcl
        ref = f"{_gcl(start_col)}{header_row}:{_gcl(end_col)}{ws.max_row}"
        ws.auto_filter.ref = ref
        return True
    except Exception as e:
        try:
            print(f"[WARN] Could not set auto_filter {table_name}: {e}")
        except Exception:
            pass
        return False

def build_calculation_traceability_sheet(wb):
    sheet_name = "Calculation Traceability"
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    headers = ["Sheet Name", "Header Name", "Python Function", "Source Fields",
               "Formula / Logic", "Output Type", "Unit", "Missing Policy", "Audit Notes"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(1, i, h)
        c.fill = FILL_HEADER
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = [24, 28, 28, 38, 58, 18, 16, 26, 50][i-1]

    rows = [
        ["IDX Screener", "Signal Category", "build_idx_screener_sheet", "Filter section tag A-F",
         "One row per ticker + signal category; section banners are visual only", "Text", "Category", "Blank if unavailable", "Prevents duplicate ticker rows being confused with unique ticker count"],
        ["IDX Screener", "Price", "build_idx_screener_sheet", "close",
         "Latest close clipped to MARKET_DATE", "Number", "IDR/share", "Blank if missing", "Formatted as IDX price"],
        ["IDX Screener", "Chg %", "build_row", "pct_change",
         "Daily percentage change stored as decimal ratio", "Number", "Percent ratio", "Blank if missing", "Excel format 0.00%"],
        ["IDX Screener", "ADR %", "build_row", "adr_pct",
         "Average daily range stored as decimal ratio; legacy percentage-point values are divided by 100 at write-time", "Number", "Percent ratio", "Blank if missing", "Excel format 0.0%"],
        ["IDX Screener", "POI Entry", "_institutional_trade_plan", "VWAP bands, OB equilibrium, EMA25/EMA50",
         "Nearest POI anchor to current close; trading logic unchanged", "Number", "IDR/share", "Blank if no valid POI", "This is an anchor, not a recommendation"],
        ["IDX Screener", "Entry Distance %", "_institutional_trade_plan", "POI Entry, close",
         "(POI Entry - Current Price) / Current Price", "Number", "Percent ratio", "Blank if price/POI missing", "Separates pullback-limit anchor from market price"],
        ["IDX Screener", "Upside from Current %", "_institutional_trade_plan", "Target POI, close",
         "(Target POI - Current Price) / Current Price", "Number", "Percent ratio", "Blank if target/price missing", "Measured from current close, not POI entry"],
        ["IDX Screener", "R/R Numeric", "_institutional_trade_plan", "Target POI, POI Entry, Stop / Invalidation",
         "(target-anchor)/(anchor-invalidation)", "Number", "Ratio", "Blank if risk <= 0", "Use this for sorting and conditional formatting"],
        ["IDX Screener", "R/R Display", "_institutional_trade_plan", "R/R Numeric",
         "Formatted as 1:X", "Text", "Ratio label", "Blank if missing", "Display only"],
        ["IDX News", "Entity Match Status", "validate_news_entity_match", "headline, ticker, company aliases, all tickers",
         "Accept only standalone row ticker or strong row-company alias; reject other ticker/company-only matches", "Text", "Status", "NO_NEWS if no item", "Prevents TOBA/TBS Energi items from being assigned to BBRI"],
        ["IDX News", "Matched Ticker(s)", "extract_idx_tickers", "headline, ticker universe",
         "Standalone IDX ticker extraction via token-boundary regex", "Text", "Ticker list", "Blank if none", "Used for multi-ticker audit"],
        ["Corp Action Mismatch Audit", "Reason", "_record_news_rejection", "validation result",
         "Rejected candidate rows are retained for audit", "Text", "Reason code", "Blank if none rejected", "Examples: TITLE_HAS_OTHER_TICKER_ONLY, REJECTED_OTHER_COMPANY, NO_COMPANY_ALIAS_MATCH"],
        ["IDX Technical Detail", "VWAP Zone", "vwap_near_zone_label", "close, VWAP, +/- SD bands",
         "Nearest explicit plotted VWAP/SD band by absolute price distance", "Text", "Label", "Blank/ dash if not near", "Avoids rounded SD-score-only mislabeling"],
        ["Data Processing Results", "Status", "build_processing_sheet", "fetch/build row result",
         "OK, PARTIAL DATA, or NO DATA with bars and reason", "Text", "Status", "Required", "Canonical universe status sheet"],
    ]
    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.font = FONT_BODY

    ws.freeze_panes = "A2"
    _add_excel_table_if_possible(ws, "Calculation_Traceability_Table", 1, 1)

def build_workbook_audit_metadata_sheet(wb, latest_market_day, run_dt, results, process_logs, ticker_universe_count):
    sheet_name = "Workbook Audit Metadata"
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    headers = ["Field", "Value"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(1, i, h)
        c.fill = FILL_HEADER
        c.font = FONT_HEADER
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    try:
        yf_ver = getattr(yf, "__version__", "unknown")
    except Exception:
        yf_ver = "unknown"
    ok_count = sum(1 for r in results if r.get("data_status") == "OK")
    partial_count = sum(1 for r in results if r.get("data_status") == "PARTIAL DATA")
    nodata_count = sum(1 for r in results if r.get("data_status") == "NO DATA")
    fallback_count = sum(1 for l in process_logs if str(l.get("retry_count", 0)) not in ("0", "0.0"))
    meta = [
        ("Generator Filename", "IDX_Screener.py"),
        ("Generator Version", GENERATOR_VERSION),
        ("MARKET_DATE", pd.Timestamp(latest_market_day).date() if latest_market_day else ""),
        ("Run Timestamp", run_dt),
        ("Python Version", sys.version.split()[0]),
        ("Data Source List", "KSEI/Raw files, yfinance (primary), IDX Financial API (fallback 1), Investing.com scrape (fallback 2), Google News RSS, IDX disclosure scrape"),
        ("yfinance Version", yf_ver),
        ("Ticker Universe Count", int(ticker_universe_count or 0)),
        ("OK Count", int(ok_count)),
        ("Partial Count", int(partial_count)),
        ("No Data Count", int(nodata_count)),
        ("Network/API Fallback Count", int(fallback_count)),
        ("Corp Action Mismatch Count", int(len(NEWS_REJECTION_AUDIT))),
        ("Formula/Traceability Status", "Calculation Traceability sheet generated"),
    ]
    for r, (k, v) in enumerate(meta, start=2):
        ws.cell(r, 1, k).border = BORDER
        ws.cell(r, 1).font = FONT_BODY
        ws.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center")
        c = ws.cell(r, 2, v)
        c.border = BORDER
        c.font = FONT_BODY
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if isinstance(v, (datetime, pd.Timestamp)):
            c.number_format = "dd mmm yyyy hh:mm"
        elif hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
            c.number_format = DATE_FORMAT_EXCEL
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 90

def build_corp_action_mismatch_audit_sheet(wb):
    sheet_name = "Corp Action Mismatch Audit"
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    headers = ["Row Ticker", "Row Company", "Rejected Title", "Detected Tickers",
               "Detected Company Aliases", "Source", "URL", "Reason"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(1, i, h)
        c.fill = FILL_HEADER
        c.font = FONT_HEADER
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = [12, 32, 70, 22, 40, 18, 50, 28][i-1]
    rows = NEWS_REJECTION_AUDIT or []
    if not rows:
        rows = [{"Row Ticker": "", "Row Company": "", "Rejected Title": "",
                 "Detected Tickers": "", "Detected Company Aliases": "",
                 "Source": "", "URL": "", "Reason": "No rejected candidates during this run"}]
    for r, row in enumerate(rows, start=2):
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(r, c, row.get(h, ""))
            cell.border = BORDER
            cell.font = FONT_BODY
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    _add_excel_table_if_possible(ws, "Corp_Action_Mismatch_Audit_Table", 1, 1)

def _apply_audit_grade_tables(wb):
    table_specs = {
        "IDX Screener": ("IDX_Screener_Table", 5),
        "IDX Technical Detail": ("IDX_Technical_Detail_Table", 6),
        "IDX Fundamental Detail": ("IDX_Fundamental_Detail_Table", 6),
        "IDX News": ("IDX_News_Table", 4),
        "Data Processing Results": ("Data_Processing_Table", 6),
    }
    for sn, (tn, header_row) in table_specs.items():
        if sn in wb.sheetnames:
            _add_excel_table_if_possible(wb[sn], tn, header_row, 2)

def _build_qa_calculation_audit_sheet(wb, latest_market_day, run_dt, results):
    name = "QA Calculation Audit"
    if name in wb.sheetnames:
        wb.remove(wb[name])
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    headers = [
        "Ticker", "Market Date", "Field", "Value", "Source",
        "Formula Version", "Missing Reason", "QA Status", "Expected Rule",
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col, header)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    checks = [
        ("cm_vwap", "Yahoo OHLCV", "VWAP_TP_RUNNING_VAR_V3", "Current-month anchored VWAP"),
        ("cm_sd", "Yahoo OHLCV", "VWAP_TP_RUNNING_VAR_V3", "(close - current MVWAP) / sigma"),
        ("pm_delta", "Yahoo OHLCV", "VWAP_COMPLETED_ANCHOR_DELTA_V3", "Same prior-month bands for current and prior close"),
        ("pq_delta", "Yahoo OHLCV", "VWAP_COMPLETED_ANCHOR_DELTA_V3", "Same prior-quarter bands for current and prior close"),
        ("py_delta", "Yahoo OHLCV", "VWAP_COMPLETED_ANCHOR_DELTA_V3", "Same prior-year bands for current and prior close"),
        ("ema25", "Yahoo Close", "EMA_ADJUST_FALSE_V1", "EMA(25), minimum 25 bars"),
        ("ema50", "Yahoo Close", "EMA_ADJUST_FALSE_V1", "EMA(50), minimum 50 bars"),
        ("sma200", "Yahoo Close", "SMA_V1", "SMA(200), minimum 200 bars"),
        ("rsi14", "Yahoo Close", "WILDER_RSI_V1", "Wilder RSI(14)"),
        ("macd_line", "Yahoo Close", "MACD_12_26_9_V1", "EMA12 - EMA26"),
    ]
    row_no = 2
    for result in results or []:
        ticker = str(result.get("ticker") or result.get("Ticker") or "")
        for key, source, version, rule in checks:
            value = result.get(key, np.nan)
            missing = value is None or value == "" or (isinstance(value, float) and not np.isfinite(value))
            status = "WARN" if missing else "PASS"
            reason = "Unavailable or insufficient source history" if missing else ""
            values = [ticker, latest_market_day, key, value if not missing else "N/A",
                      source, version, reason, status, rule]
            for col, item in enumerate(values, start=1):
                cell = ws.cell(row_no, col, item)
                cell.border = BORDER
                cell.font = FONT_BODY
                cell.alignment = Alignment(
                    horizontal="left" if col in (3, 5, 7, 9) else "center",
                    vertical="center",
                    wrap_text=True,
                )
            row_no += 1
    ws.freeze_panes = "A2"
    widths = [12, 14, 18, 18, 18, 32, 38, 12, 52]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    _add_excel_table_if_possible(ws, "QA_Calculation_Audit", 1, 1)

def _add_audit_artifacts(wb, latest_market_day, run_dt, results, process_logs, ticker_universe_count):
    _build_qa_calculation_audit_sheet(wb, latest_market_day, run_dt, results)
    _apply_audit_grade_tables(wb)


def _postprocess_workbook(wb):
    from openpyxl.cell.cell import MergedCell

    # Remove freeze panes on IDX Overview
    if "IDX Overview" in wb.sheetnames:
        wb["IDX Overview"].freeze_panes = None

    def _trim_ghost_columns(ws):
        from openpyxl.cell.cell import MergedCell
        max_col = ws.max_column
        if max_col is None or max_col == 0:
            return
        last_real_col = max_col
        for col in range(max_col, 0, -1):
            col_has_data = False
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                     min_col=col, max_col=col):
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue
                    if cell.value is not None and str(cell.value).strip() not in ("", "N/A"):
                        col_has_data = True
                        break
                if col_has_data:
                    break
            if col_has_data:
                last_real_col = col
                break
        cols_to_delete = max_col - last_real_col
        if cols_to_delete > 0:
            ws.delete_cols(last_real_col + 1, cols_to_delete)

    for sheet_name in ["IDX Technical Detail"]:
        if sheet_name in wb.sheetnames:
            _trim_ghost_columns(wb[sheet_name])

    target_sheets = [n for n in ["IDX Technical Detail"] if n in wb.sheetnames]
    for sn in target_sheets:
        ws = wb[sn]
        data_start = 7 if sn in ("IDX Technical Detail",) else 4
        for row in ws.iter_rows(min_row=data_start):
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                if c.value in (None, ""):
                    c.value = "N/A"

    for sn in target_sheets:
        ws = wb[sn]
        for c in range(2, ws.max_column + 1):
            if not isinstance(ws.cell(4, c), MergedCell):
                ws.cell(4, c).value = None

    if "IDX Technical Detail" in wb.sheetnames:
        wb["IDX Technical Detail"].freeze_panes = "E7"

    for sn in target_sheets:
        ws = wb[sn]
        header_row = None
        for r in range(1, min(ws.max_row, 10) + 1):
            vals = [ws.cell(r, col).value for col in range(1, ws.max_column + 1)]
            if "RSI 14" in vals:
                header_row = r
                break
        if header_row:
            for col in range(1, ws.max_column + 1):
                if ws.cell(header_row, col).value == "RSI 14":
                    for rr in range(header_row + 1, ws.max_row + 1):
                        cell = ws.cell(rr, col)
                        if isinstance(cell, MergedCell):
                            continue
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = "0.00"


def _apply_final_sheet_order(wb):
    desired = [s for s in DESIRED_SHEET_ORDER if s in wb.sheetnames]
    remaining = [s for s in wb.sheetnames if s not in desired]
    wb._sheets = [wb[s] for s in desired + remaining]


def _set_guide_logic_widths(ws):
    for col in ["B", "C", "E", "F", "G"]:
        ws.column_dimensions[col].width = 40
    ws.column_dimensions["D"].width = 50


def _write_ohlcv_history_cache(rows_cache: dict, market_date: str) -> None:
    target_dir = os.path.join(OUTPUT_DIR, "ohlcv", market_date)
    os.makedirs(target_dir, exist_ok=True)
    for ticker, hist in (rows_cache or {}).items():
        if hist is None or hist.empty:
            continue
        records = []
        for ts, row in hist.iterrows():
            records.append({
                "date": pd.Timestamp(ts).strftime("%Y-%m-%d"),
                "open": safe_num(row.get("Open"), None),
                "high": safe_num(row.get("High"), None),
                "low": safe_num(row.get("Low"), None),
                "close": safe_num(row.get("Close"), None),
                "volume": safe_num(row.get("Volume"), None),
                "source": "yfinance",
                "adjusted": False,
                "timezone": "Asia/Jakarta",
                "session": "IDX regular daily session",
            })
        path = os.path.join(target_dir, f"{ticker}.json")
        with open(path, "w", encoding="utf-8") as fh:
            json.dump({
                "schemaVersion": 1,
                "ticker": ticker,
                "date": market_date,
                "source": "yfinance",
                "adjusted": False,
                "timezone": "Asia/Jakarta",
                "session": "IDX regular daily session",
                "formulaVersion": "ohlcv-series-v1",
                "rows": records,
            }, fh, separators=(",", ":"))


def load_idx_listed_roster():
    """Canonical IDX listed-companies roster (data_sources/idx-listed.json,
    built from the official 'Daftar Saham' export). Returns [] if absent so the
    pipeline still runs KSEI-only. Source of truth for 'what is listed'."""
    import json
    path = os.path.join(BASE_DIR, "data_sources", "idx-listed.json")
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh).get("records", [])
    except Exception:
        return []


def merge_idx_roster(ksei_df):
    """Append rows for IDX-listed tickers missing from the KSEI source so the
    FULL listed universe is scanned. Recent IPOs are not yet in KSEI; their
    ownership columns stay NaN and degrade to N/A downstream exactly like any
    KSEI ticker without a reported holder table."""
    listed = load_idx_listed_roster()
    if not listed:
        return ksei_df
    have = set(ksei_df["Ticker"].astype(str).str.upper().str.strip())
    extra = []
    for r in listed:
        ticker = str(r.get("ticker", "")).upper().strip()
        if not ticker or ticker in have:
            continue
        row = {col: np.nan for col in ksei_df.columns}
        row["Ticker"] = ticker
        if "Emiten" in row:
            row["Emiten"] = r.get("name")
        if "Shares Outstanding" in row and r.get("shares"):
            row["Shares Outstanding"] = r["shares"]
        extra.append(row)
    if not extra:
        return ksei_df
    print(f"[OK] IDX roster: +{len(extra)} listed tickers absent from KSEI "
          f"(recent IPOs) -> {[e['Ticker'] for e in extra]}")
    return pd.concat([ksei_df, pd.DataFrame(extra)], ignore_index=True)


def main():
    print("\n" + "=" * 110)
    print("  IDX Screener")
    print("=" * 110)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(CACHE_DIR, exist_ok=True)

    ksei_raw = load_ksei()
    TOTAL_SOURCE_TICKERS = len(ksei_raw)

    for ticker_col in ("Ticker", "ticker", "Code", "code"):
        if ticker_col in ksei_raw.columns:
            break
    else:
        raise KeyError("Could not find ticker column in KSEI source")

    ksei_raw = ksei_raw.copy()
    ksei_raw["_ticker_norm"] = ksei_raw[ticker_col].astype(str).str.strip().str.upper()
    ksei_raw = ksei_raw[ksei_raw["_ticker_norm"].notna() & (ksei_raw["_ticker_norm"] != "")]
    ksei_raw = ksei_raw.drop_duplicates(subset=["_ticker_norm"], keep="first").reset_index(drop=True)

    use_custom = bool(globals().get("USE_CUSTOM_TICKERS_ONLY", False))
    if use_custom:
        ksei = ksei_raw[ksei_raw["_ticker_norm"].isin(CUSTOM_TICKERS)].copy()
        ksei["_ticker_order"] = ksei["_ticker_norm"].map({t: i for i, t in enumerate(CUSTOM_TICKERS)})
        ksei = ksei.sort_values("_ticker_order").drop(columns=["_ticker_norm","_ticker_order"]).reset_index(drop=True)
        print(f"[OK] KSEI: {TOTAL_SOURCE_TICKERS} tickers | Custom mode: {len(ksei)} selected")
    elif is_backtest_mode():
        # Backtest mode: restrict to BACKTEST_TICKERS universe only
        _bt_set = {t.strip().upper() for t in BACKTEST_TICKERS}
        ksei = ksei_raw[ksei_raw["_ticker_norm"].isin(_bt_set)].copy()
        ksei = ksei.drop(columns=["_ticker_norm"]).reset_index(drop=True)
        print(f"[OK] KSEI: {TOTAL_SOURCE_TICKERS} tickers | BACKTEST mode: {len(ksei)} selected (as of {MARKET_DATE})")
    else:
        ksei = ksei_raw.drop(columns=["_ticker_norm"]).reset_index(drop=True)
        # Scan the FULL IDX listed universe: union in any listed ticker missing
        # from KSEI (recent IPOs) so nothing listed is silently skipped.
        ksei = merge_idx_roster(ksei)
        print(f"[OK] KSEI: {TOTAL_SOURCE_TICKERS} tickers | Full universe + IDX roster = {len(ksei)} scanned (as of {MARKET_DATE})")

    shares_cache = load_shares_cache()
    results = []
    process_logs = []
    # MARKET_DATE is the authoritative as-of date for this run
    latest_market_day_global = get_market_date().strftime("%Y-%m-%d")
    _hist_cache = {}

    for idx, (_, ksei_row) in enumerate(ksei.iterrows(), start=1):
        ticker = str(ksei_row["Ticker"]).upper().strip()
        print(f"[{idx}/{len(ksei)}] {ticker} ...", end=" ")

        hist, retry_count, source_used, fetch_reason = fetch_history_with_retry(ticker)
        shares = fetch_shares_outstanding(ticker, shares_cache)

        if hist is None or hist.empty:
            built = base_row_from_ksei(ksei_row)
            results.append(built)
            process_logs.append({"ticker": ticker, "status": "NO DATA", "bars": 0,
                                  "retry_count": retry_count, "source_used": source_used,
                                  "latest_market_day": "", "reason": fetch_reason or "No market data"})
            print("NO DATA")
            continue

        built = build_row(ksei_row, hist, shares)
        results.append(built)

        _hist_cache[ticker] = hist

        reason = "OK" if built["data_status"] == "OK" else f"Insufficient bars ({len(hist)} < {MIN_BARS_FULL})"
        process_logs.append({"ticker": ticker, "status": built["data_status"], "bars": len(hist),
                              "retry_count": retry_count, "source_used": source_used,
                              "latest_market_day": built["latest_market_day"] or "", "reason": reason})
        print(built["data_status"])

    save_shares_cache(shares_cache)

    results      = sorted(results, key=lambda x: x["ticker"])
    process_logs = sorted(process_logs, key=lambda x: x["ticker"])
    summary_rows = [r for r in results if r["data_status"] == "OK"]
    _write_ohlcv_history_cache(_hist_cache, latest_market_day_global)

    # ── News entity validation context ───────────────────────────────────────
    _configure_news_entity_context(results)
    try:
        _run_news_entity_validation_self_tests()
    except AssertionError as _ae:
        raise RuntimeError(f"News entity validation self-test failed: {_ae}")

    # ── RS Rating universe normalisation pass ──────────────────────────────────
    if len(_RS_UNIVERSE_CACHE) >= 5:
        lo = min(_RS_UNIVERSE_CACHE)
        hi = max(_RS_UNIVERSE_CACHE)
        if hi != lo:
            for _r in results:
                raw = _r.get("_rs_raw", np.nan)
                if pd.notna(raw):
                    rating = int(max(1, min(99, round(1 + 98 * (raw - lo) / (hi - lo)))))
                    _r["rs_rating"]      = rating
                    _r["rs_rating_zone"] = rs_rating_zone(rating)

    # ── Conviction scoring pass (_sc_* fields → DETAIL_SCHEMA columns) ─────────
    # Run the V17 POI engine across all OK rows and merge _sc_* back into each row
    print("[INFO] Running conviction scoring pass ...")
    try:
        scored_board = _build_idx_vwap_shortlist(summary_rows)
        # Build lookup: ticker → scored row
        scored_map = {str(s.get("ticker","")).upper(): s for s in scored_board}
        for _r in results:
            tk = str(_r.get("ticker","")).upper()
            if tk in scored_map:
                sc = scored_map[tk]
                # Merge only _sc_* and ms_* fields
                for _k, _v in sc.items():
                    if _k.startswith("_sc_") or _k.startswith("ms_"):
                        _r.setdefault(_k, _v)   # don't overwrite if already set by build_row
    except Exception as _e:
        print(f"[WARN] Conviction scoring pass failed: {_e}")

    # ── mp_profile for screener sheet (derived from market profile fields) ──────
    for _r in results:
        if not _r.get("mp_profile"):
            _r["mp_profile"] = _market_profile_summary(_r)

    _run_dt  = datetime.now()
    out_file = os.path.join(OUTPUT_DIR, _build_output_filename(_run_dt, latest_market_day_global))

    # ── Build workbook ──────────────────────────────────────────────────────
    wb  = Workbook()
    ws1 = wb.active
    ws1.sheet_view.showGridLines = False

    _sector_source_rows = results
    build_true_idx_sector_movers_sheet(wb, latest_market_day_global, _sector_source_rows)

    # Inject RRG sector state into conviction scorer
    try:
        if "IDX Overview" in wb.sheetnames:
            _ws_ov  = wb["IDX Overview"]
            _rrg_data = []
            for _row in _ws_ov.iter_rows(min_row=10, max_row=_ws_ov.max_row, values_only=True):
                _sector_val = _row[13] if len(_row) > 13 else None
                _quad_val   = _row[17] if len(_row) > 17 else None
                if _sector_val and _quad_val and str(_quad_val) in ("Leading","Improving","Weakening","Lagging"):
                    _rrg_data.append({"Sector": str(_sector_val), "Quadrant": str(_quad_val)})
            _inject_rrg_sector_state(_rrg_data)
    except Exception:
        pass

    # ── Post-RRG: update rrg_quadrant in all results (was always "-" during build_row) ──
    try:
        for _r in results:
            _sec = str(_r.get("idx_sector", "") or "")
            if _sec:
                _r["rrg_quadrant"] = get_rrg_quadrant_for_sector(_sec)
    except Exception:
        pass

    # Sheet: IDX Technical Detail  (renamed from IDX Screener Detail)
    ws2 = wb.create_sheet("IDX Technical Detail")
    ws2.sheet_view.showGridLines = False
    build_detail_sheet(ws2, latest_market_day_global, results)

    # Sheet: IDX Fundamental Detail  (renamed + redesigned from Fundamental Key Stats)
    # Pass every scanned ticker, not just summary_rows (OK-only) -- yfinance
    # company fundamentals (market cap, sector, ROE, ...) come from .info(),
    # which has nothing to do with OHLCV bar count, but a ticker with < 210
    # bars (suspended/recently-relisted, e.g. WSKT, SRIL, ARMY after their
    # PKPU/restructuring halts) got data_status "NO DATA"/"PARTIAL DATA" and
    # was silently dropped from the Fundamental Detail sheet entirely --
    # confirmed live: yfinance still returns full real fundamentals for these
    # names. IDX Technical Detail already uses the unfiltered `results` for
    # the same reason (build_fundamental_key_stats_sheet's own OK/PARTIAL/NO
    # DATA gate below decides what actually gets a row).
    print("[INFO] Building IDX Fundamental Detail sheet ...")
    build_fundamental_key_stats_sheet(wb, latest_market_day_global, results)

    # Sheet: IDX Screener (filtered signal sheet)
    print("[INFO] Building IDX Screener sheet ...")
    build_idx_screener_sheet(wb, latest_market_day_global, results)

    # Sheet: IDX Disclosure (Google News RSS — market-date filtered, IFNA fallback)
    print("[INFO] Building IDX News sheet ...")
    try:
        build_idx_disclosure_sheet(wb, latest_market_day_global, results)
    except Exception as _de:
        print(f"[WARN] IDX News sheet skipped: {_de}")

    # Sheet: Guide & Logic Reference
    ws_guide = wb.create_sheet("Guide & Logic Reference")
    ws_guide.sheet_view.showGridLines = False
    build_source_limited_guide_sheet(ws_guide)

    # Sheet: Data Processing Results
    ws_proc = wb.create_sheet("Data Processing Results")
    ws_proc.sheet_view.showGridLines = False
    build_processing_sheet(ws_proc, latest_market_day_global, len(ksei), process_logs, summary_rows)

    # Sheet: audit artifacts and traceability
    print("[INFO] Building audit artifacts and calculation traceability ...")
    try:
        _add_audit_artifacts(wb, latest_market_day_global, _run_dt, results, process_logs, len(ksei))
    except Exception as _ae:
        print(f"[WARN] Audit artifacts skipped: {_ae}")

    print("[INFO] Building Data Source Audit sheet ...")
    build_data_source_audit_sheet(wb, latest_market_day_global, results)

    # ── Sheet ordering & cleanup ─────────────────────────────────────────────
    desired_order = [
        "IDX Overview",
        "IDX Screener",
        "IDX Technical Detail",
        "IDX Fundamental Detail",
        "IDX News",
        "Guide & Logic Reference",
        "Data Processing Results",
        "QA Calculation Audit",
        "Data Source Audit",
    ]
    _ALLOWED_SHEETS = set(desired_order)
    for _ws in list(wb.worksheets):
        if _ws.title not in _ALLOWED_SHEETS:
            wb.remove(_ws)

    ordered = []
    for nm in desired_order:
        if nm in wb.sheetnames:
            ordered.append(wb[nm])
    for wsx in wb.worksheets:
        if wsx not in ordered:
            ordered.append(wsx)
    wb._sheets = ordered

    _postprocess_workbook(wb)
    _apply_final_sheet_order(wb)
    _apply_excel_date_format_workbook(wb)
    wb.save(out_file)
    print(f"[DONE] Saved: {out_file}")

# v6.3 PNG FINAL DIRECT PATCH
# Override ONLY IDX Overview with reliable matplotlib PNG render
# =========================
def _best_discount_zone(r: dict):
    close_px = safe_num(r.get("close"), np.nan)
    checks = [
        ("PQVWAP", r.get("pq_vwap"), 2.0, "Previous QVWAP"),
        ("PQ -1 SD", r.get("pq_m1"), 1.25, "Previous QVWAP"),
        ("PYVWAP", r.get("py_vwap"), 2.0, "Previous Year VWAP"),
        ("PY -1 SD", r.get("py_m1"), 1.25, "Previous Year VWAP"),
    ]
    best = None
    best_abs = 999.0
    for zone_type, level, tol, framework in checks:
        p = _pct_from_level(close_px, level)
        if pd.notna(p) and abs(p) <= tol and abs(p) < best_abs:
            best = {"zone_type": zone_type, "framework": framework, "dist_pct": round(p, 2), "tol_pct": tol}
            best_abs = abs(p)
    if best is None:
        for framework, z in [("Previous QVWAP", str(r.get("pq_remarks", "") or "")),
                             ("Previous Year VWAP", str(r.get("py_remarks", "") or ""))]:
            if _zone_is_tradeable_poi(z):
                z_u = z.upper()
                if "-2" in z_u or "-1" in z_u:
                    zt = "Discount Zone"
                elif "VWAP" in z_u:
                    zt = "VWAP Anchor"
                else:
                    zt = "POI Zone"
                best = {"zone_type": zt, "framework": framework, "dist_pct": np.nan, "tol_pct": np.nan}
                break
    return best

# ---- mutate detail schema in-place ----
for _grp_idx, _grp in enumerate(DETAIL_SCHEMA):
    if _grp[0] == "MACD Momentum":
        _cols = list(_grp[2])
        _new_cols = []
        for _c in _cols:
            if _c[0] in ("macd_4color", "macd_entry"):
                continue
            _new_cols.append(_c)
        DETAIL_SCHEMA[_grp_idx] = (_grp[0], _grp[1], _new_cols)
    elif _grp[0].startswith("Current QVWAP"):
        _cols = list(_grp[2]); _new=[]
        for _c in _cols:
            _new.append(_c)
            if _c[0] == "q_remarks":
                pass  # already exists in base schema
        DETAIL_SCHEMA[_grp_idx] = (_grp[0], _grp[1], _new)
    elif _grp[0].startswith("Previous QVWAP"):
        _cols = list(_grp[2]); _new=[]
        for _c in _cols:
            _new.append(_c)
            if _c[0] == "pq_remarks":
                pass  # already exists in base schema
        DETAIL_SCHEMA[_grp_idx] = (_grp[0], _grp[1], _new)
    elif _grp[0].startswith("Previous Year VWAP"):
        _cols = list(_grp[2]); _new=[]
        for _c in _cols:
            _new.append(_c)
            if _c[0] == "py_remarks":
                pass  # already exists in base schema
        DETAIL_SCHEMA[_grp_idx] = (_grp[0], _grp[1], _new)

# ---- patch process_single_stock row defaults and computations via text-level hooks below if source lines exist ----


def _fetch_latest_vix_safe():
    try:
        tk = yf.Ticker("^VIX")
        _df_md = get_market_date()
        _df_end = (_df_md + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        df = tk.history(period="10d", end=_df_end, interval="1d", auto_adjust=False)
        if df is None or df.empty or "Close" not in df.columns:
            return None
        s = pd.to_numeric(df["Close"], errors="coerce").dropna()
        if s.empty:
            return None
        return float(s.iloc[-1])
    except Exception:
        return None

# =========================
# SMC ENGINE  — Pine Script "SMC" v6 faithful Python port
# Implements: Internal/Swing BOS+CHoCH, Strong/Weak H/L, Order Blocks,
# FVGs, EQH/EQL, Premium/Discount, Ribbon (EMA25/50), CurrentQVWAP,
# PY/PQ VWAP bands, IBH/IBL/MDH/MDL/PWH/PWL (already in build_row).
# All outputs are lookahead-free and replay-safe.
# =========================

def _lux_leg(highs_arr, lows_arr, i, size):
    '''
    Pine Script leg(int size):
      BEARISH_LEG (0) when high[size] > ta.highest(size)  -> new pivot high
      BULLISH_LEG (1) when low[size]  < ta.lowest(size)   -> new pivot low
    We replicate with strict forward-safe window.
    '''
    if i < size:
        return None
    anchor_i = i - size          # the pivot candidate bar
    # Look at the size bars AFTER anchor (not including anchor itself = bars anchor+1..i)
    window_highs = highs_arr[anchor_i + 1 : i + 1]
    window_lows  = lows_arr [anchor_i + 1 : i + 1]
    if len(window_highs) == 0:
        return None
    if highs_arr[anchor_i] > (window_highs.max() if len(window_highs) else highs_arr[anchor_i]):
        return 0   # BEARISH_LEG: pivot high at anchor_i
    if lows_arr[anchor_i]  < (window_lows.min()  if len(window_lows)  else lows_arr[anchor_i]):
        return 1   # BULLISH_LEG: pivot low at anchor_i
    return None


def _fmt_range(a, b):
    try:
        if a is None or b is None or pd.isna(a) or pd.isna(b):
            return ""
        lo = int(round(min(float(a), float(b))))
        hi = int(round(max(float(a), float(b))))
        return f"{lo} - {hi}"
    except Exception:
        return ""


def compute_smc_engine(hist: pd.DataFrame,
                       internal_size: int = 5,
                       swing_size: int = 50,
                       eq_length: int = 3,
                       eq_threshold: float = 0.1,
                       mitigation_mode: str = "close") -> dict:
    '''
    Full SMC engine output. Returns dict with 26+ columns.
    All indexes are bar-by-bar; no lookahead into future bars.
    '''
    empty = {
        "smc_internal_trend":         "N/A",
        "smc_swing_trend":            "N/A",
        "smc_internal_bos_count":     0,
        "smc_internal_choch_count":   0,
        "smc_swing_bos_count":        0,
        "smc_swing_choch_count":      0,
        "smc_latest_internal_struct": "N/A",
        "smc_latest_internal_struct_date": "N/A",
        "smc_latest_swing_struct":    "N/A",
        "smc_latest_swing_struct_date": "N/A",
        "smc_strong_high":            np.nan,
        "smc_weak_high":              np.nan,
        "smc_strong_low":             np.nan,
        "smc_weak_low":               np.nan,
        "smc_closest_ob":             "",
        "smc_ob_equilibrium":         np.nan, "smc_eq_high": np.nan, "smc_eq_low": np.nan, "eq_breakout_flag": False,
        "smc_closest_ob_bear":        "",
        "smc_ob_direction":           "N/A",
        "smc_ob_dist_pct":            np.nan,
        "smc_closest_fvg":            "",
        "smc_fvg_bias":               "N/A",
        "smc_eqhl_status":            "N/A",
        "smc_premium_discount":       "N/A",
        "smc_premium_zone":           "",
        "smc_discount_zone":          "",
        "smc_summary":                "N/A",
        "smc_ribbon_bias":            "N/A",
        "smc_current_qvwap":          np.nan,
        "smc_qvwap_dist_pct":         np.nan,
        "smc_nearest_pyq_band":       "",
        "smc_pyq_band_dist_pct":      np.nan,
        "smc_market_structure_bias":  "N/A",
        "smc_composite_state":        "N/A",
        # Legacy compatibility fields (old schema)
        "smc_bull_internal_ob":       "",
        "smc_bear_internal_ob":       "",
        "smc_discount":               "",
        "smc_equilibrium":            "",
        "smc_premium":                "",
        "smc_state":                  "N/A",
        "ms_trend_bias":              "N/A",
        "ms_structure_phase":         "N/A",
        "ms_last_structural_event":   "N/A",
        "ms_event_age_d":             "",
        "ms_swing_sequence":          "N/A",
        "ms_structure_quality":       "N/A",
        "ms_volume_confirmation":     "N/A",
    }

    if hist is None or hist.empty:
        return empty

    df = hist.copy()
    for c in ["Open", "High", "Low", "Close", "Volume"]:
        if c not in df.columns:
            return empty
    df = df[["Open", "High", "Low", "Close", "Volume"]].dropna()
    min_bars = max(60, swing_size + 5)
    if len(df) < min_bars:
        return empty

    highs  = df["High"].astype(float).values
    lows   = df["Low"].astype(float).values
    opens  = df["Open"].astype(float).values
    closes = df["Close"].astype(float).values
    vols   = df["Volume"].astype(float).values
    n      = len(df)

    # ATR-based filter for OB and EQH/EQL
    tr = np.maximum(highs - lows,
         np.maximum(np.abs(highs - np.roll(closes, 1)),
                    np.abs(lows  - np.roll(closes, 1))))
    tr[0] = highs[0] - lows[0]
    atr200 = pd.Series(tr).rolling(200, min_periods=1).mean().values

    # Parsed highs/lows (high-volatility bar inversion for OB detection)
    high_vol_bar  = (highs - lows) >= (2.0 * atr200)
    parsed_highs  = np.where(high_vol_bar, lows,  highs)
    parsed_lows   = np.where(high_vol_bar, highs, lows)

    # ── State variables ──────────────────────────────────────────────────────
    # Pivots (Pine: pivot struct with currentLevel, lastLevel, crossed, barIndex)
    iH = dict(level=np.nan, last=np.nan, crossed=False, idx=None)  # internal high
    iL = dict(level=np.nan, last=np.nan, crossed=False, idx=None)  # internal low
    sH = dict(level=np.nan, last=np.nan, crossed=False, idx=None)  # swing high
    sL = dict(level=np.nan, last=np.nan, crossed=False, idx=None)  # swing low
    eH = dict(level=np.nan, last=np.nan, crossed=False, idx=None)  # equal high
    eL = dict(level=np.nan, last=np.nan, crossed=False, idx=None)  # equal low

    internal_trend = 0   # +1 bullish, -1 bearish, 0 neutral
    swing_trend    = 0

    prev_leg_int   = None
    prev_leg_sw    = None

    # Trailing extremes for premium/discount zones
    trailing_top     = np.nan
    trailing_bottom  = np.nan
    trailing_top_idx = None
    trailing_bot_idx = None

    # Order blocks list: {barHigh, barLow, bias, idx}
    internal_obs = []
    swing_obs    = []

    # FVGs: {top, bottom, bias, formed_idx}
    fvgs = []

    # Counters
    int_bos_count   = 0
    int_choch_count = 0
    sw_bos_count    = 0
    sw_choch_count  = 0

    last_int_event  = ""
    last_sw_event   = ""
    last_event_idx  = None
    last_int_event_idx = None
    last_sw_event_idx  = None

    eqh_detected = False
    eql_detected = False

    # ── Helper: store order block on BOS/CHoCH ───────────────────────────────
    def _store_ob(pivot_obj, bias, cur_i, internal=True):
        p_idx = pivot_obj.get("idx")
        if p_idx is None or p_idx >= cur_i:
            return
        if bias == 1:   # bullish -> find min parsed low in window
            window = parsed_lows[p_idx:cur_i]
            if len(window) == 0:
                return
            rel = int(np.argmin(window))
        else:           # bearish -> find max parsed high in window
            window = parsed_highs[p_idx:cur_i]
            if len(window) == 0:
                return
            rel = int(np.argmax(window))
        idx = p_idx + rel
        ob  = {"barHigh": float(parsed_highs[idx]),
               "barLow":  float(parsed_lows[idx]),
               "bias":    bias,
               "idx":     idx}
        if internal:
            internal_obs.insert(0, ob)
            if len(internal_obs) > 100:
                internal_obs[:] = internal_obs[:100]
        else:
            swing_obs.insert(0, ob)
            if len(swing_obs) > 100:
                swing_obs[:] = swing_obs[:100]

    # ── Helper: mitigate order blocks ─────────────────────────────────────────
    def _mitigate_obs(obs_list, i):
        _mm = str(mitigation_mode).lower()
        surviving = []
        for ob in obs_list:
            crossed = False
            if ob["bias"] == -1:  # bearish OB
                crossed = (closes[i] > ob["barHigh"]) if _mm == "close" else (highs[i] > ob["barHigh"])
            elif ob["bias"] == 1:  # bullish OB
                crossed = (closes[i] < ob["barLow"])  if _mm == "close" else (lows[i]  < ob["barLow"])
            if not crossed:
                surviving.append(ob)
        return surviving

    # ── Main bar-by-bar loop ─────────────────────────────────────────────────
    for i in range(n):
        c = closes[i]
        h = highs[i]
        l = lows[i]

        # Trailing extremes
        if np.isnan(trailing_top) or h > trailing_top:
            trailing_top = h; trailing_top_idx = i
        if np.isnan(trailing_bottom) or l < trailing_bottom:
            trailing_bottom = l; trailing_bot_idx = i

        # Swing structure pivots (size = swing_size)
        leg_sw = _lux_leg(highs, lows, i, swing_size)
        if leg_sw is not None and prev_leg_sw is not None and leg_sw != prev_leg_sw:
            pi = i - swing_size
            if leg_sw == 1:   # new bullish leg => pivot low at pi
                sL.update(last=sL["level"], level=float(lows[pi]),
                          crossed=False, idx=pi)
                trailing_bottom = float(lows[pi]); trailing_bot_idx = pi
            else:             # new bearish leg => pivot high at pi
                sH.update(last=sH["level"], level=float(highs[pi]),
                          crossed=False, idx=pi)
                trailing_top = float(highs[pi]); trailing_top_idx = pi
        if leg_sw is not None:
            prev_leg_sw = leg_sw

        # Internal structure pivots (size = internal_size)
        leg_in = _lux_leg(highs, lows, i, internal_size)
        if leg_in is not None and prev_leg_int is not None and leg_in != prev_leg_int:
            pi = i - internal_size
            if leg_in == 1:
                iL.update(last=iL["level"], level=float(lows[pi]),
                          crossed=False, idx=pi)
            else:
                iH.update(last=iH["level"], level=float(highs[pi]),
                          crossed=False, idx=pi)
        if leg_in is not None:
            prev_leg_int = leg_in

        # EQH / EQL detection (Pine: equalHighsLowsLength = eq_length)
        leg_eq = _lux_leg(highs, lows, i, eq_length)
        if leg_eq is not None and leg_eq != prev_leg_int:
            pi = i - eq_length
            if leg_eq == 1:  # pivot low
                if not np.isnan(eL["level"]) and abs(float(lows[pi]) - eL["level"]) < eq_threshold * atr200[i]:
                    eql_detected = True
                eL.update(last=eL["level"], level=float(lows[pi]),
                          crossed=False, idx=pi)
            else:            # pivot high
                if not np.isnan(eH["level"]) and abs(float(highs[pi]) - eH["level"]) < eq_threshold * atr200[i]:
                    eqh_detected = True
                eH.update(last=eH["level"], level=float(highs[pi]),
                          crossed=False, idx=pi)

        # ── Internal structure BOS / CHoCH (displayStructure(true)) ──────────
        # Bullish break: price crosses above internal pivot high
        iH_extra = (iH["idx"] is not None and
                    not np.isnan(iH["level"]) and
                    not np.isnan(sH.get("level", np.nan)) and
                    iH["level"] != sH.get("level", np.nan))
        if (not np.isnan(iH["level"]) and c > iH["level"] and
                not iH["crossed"] and iH_extra):
            tag = "CHoCH" if internal_trend == -1 else "BOS"
            iH["crossed"] = True
            internal_trend = 1
            if tag == "BOS":
                int_bos_count += 1
            else:
                int_choch_count += 1
            last_int_event = f"Internal Bullish {tag}"
            last_event_idx = i
            last_int_event_idx = i
            _store_ob(iH, 1, i, internal=True)

        # Bearish break: price crosses below internal pivot low
        iL_extra = (iL["idx"] is not None and
                    not np.isnan(iL["level"]) and
                    not np.isnan(sL.get("level", np.nan)) and
                    iL["level"] != sL.get("level", np.nan))
        if (not np.isnan(iL["level"]) and c < iL["level"] and
                not iL["crossed"] and iL_extra):
            tag = "CHoCH" if internal_trend == 1 else "BOS"
            iL["crossed"] = True
            internal_trend = -1
            if tag == "BOS":
                int_bos_count += 1
            else:
                int_choch_count += 1
            last_int_event = f"Internal Bearish {tag}"
            last_event_idx = i
            last_int_event_idx = i
            _store_ob(iL, -1, i, internal=True)

        # ── Swing structure BOS / CHoCH (displayStructure()) ─────────────────
        if not np.isnan(sH["level"]) and c > sH["level"] and not sH["crossed"]:
            tag = "CHoCH" if swing_trend == -1 else "BOS"
            sH["crossed"] = True
            swing_trend = 1
            if tag == "BOS":
                sw_bos_count += 1
            else:
                sw_choch_count += 1
            last_sw_event = f"Swing Bullish {tag}"
            last_event_idx = i
            last_sw_event_idx = i
            _store_ob(sH, 1, i, internal=False)

        if not np.isnan(sL["level"]) and c < sL["level"] and not sL["crossed"]:
            tag = "CHoCH" if swing_trend == 1 else "BOS"
            sL["crossed"] = True
            swing_trend = -1
            if tag == "BOS":
                sw_bos_count += 1
            else:
                sw_choch_count += 1
            last_sw_event = f"Swing Bearish {tag}"
            last_event_idx = i
            last_sw_event_idx = i
            _store_ob(sL, -1, i, internal=False)

        # ── Fair Value Gap detection ──────────────────────────────────────────
        # Bullish FVG: low[i] > high[i-2] (gap between prev prev high and cur low)
        if i >= 2:
            if lows[i] > highs[i - 2]:       # bullish FVG
                fvgs.append({"top": lows[i], "bottom": highs[i - 2], "bias": 1, "formed": i})
            elif highs[i] < lows[i - 2]:     # bearish FVG
                fvgs.append({"top": lows[i - 2], "bottom": highs[i], "bias": -1, "formed": i})

        # Mitigate filled FVGs
        surviving_fvg = []
        for fvg in fvgs:
            filled = (fvg["bias"] == 1  and l < fvg["bottom"]) or \
                     (fvg["bias"] == -1 and h > fvg["top"])
            if not filled:
                surviving_fvg.append(fvg)
        fvgs[:] = surviving_fvg

        # Mitigate crossed order blocks
        internal_obs[:] = _mitigate_obs(internal_obs, i)
        swing_obs[:]    = _mitigate_obs(swing_obs,    i)

    # ── Post-loop: summarize outputs ─────────────────────────────────────────
    close_last = float(closes[-1])

    # Strong/Weak High/Low (Pine: trailing extremes + trend bias)
    # Strong = in direction of trend (traps the other side); Weak = opposite direction
    strong_high = trailing_top    if swing_trend == -1 else np.nan  # bearish trend = strong high
    weak_high   = trailing_top    if swing_trend ==  1 else np.nan  # bullish trend = weak high
    strong_low  = trailing_bottom if swing_trend ==  1 else np.nan  # bullish trend = strong low
    weak_low    = trailing_bottom if swing_trend == -1 else np.nan  # bearish trend = weak low

    # Fallback: always provide trailing extremes when trend is 0
    if swing_trend == 0:
        strong_high = trailing_top
        strong_low  = trailing_bottom

    # Closest active order blocks
    # Audit fix: previous code used the first visible OB (top 2 internal + top 2 swing).
    # TradingView's active OB view is effectively "nearest unmitigated plotted zone";
    # therefore choose nearest per bias from all surviving internal + swing OBs.
    all_active_obs = internal_obs  # TradingView parity: use Internal Order Blocks only

    def _ob_mid(ob):
        return (float(ob["barHigh"]) + float(ob["barLow"])) / 2.0

    def _ob_distance_to_price(ob):
        lo = min(float(ob["barLow"]), float(ob["barHigh"]))
        hi = max(float(ob["barLow"]), float(ob["barHigh"]))
        if lo <= close_last <= hi:
            return 0.0
        return min(abs(close_last - lo), abs(close_last - hi), abs(close_last - _ob_mid(ob)))

    bull_candidates = [ob for ob in all_active_obs if ob.get("bias") == 1]
    bear_candidates = [ob for ob in all_active_obs if ob.get("bias") == -1]

    bull_ob = min(bull_candidates, key=_ob_distance_to_price) if bull_candidates else None
    bear_ob = min(bear_candidates, key=_ob_distance_to_price) if bear_candidates else None

    closest_ob_str      = ""
    closest_ob_bear_str = ""
    ob_equilibrium      = np.nan
    ob_direction        = "N/A"
    ob_dist_pct         = np.nan
    ob_bull_age         = np.nan
    ob_bear_age         = np.nan

    if bull_ob and bear_ob:
        chosen_ob = bull_ob if _ob_distance_to_price(bull_ob) <= _ob_distance_to_price(bear_ob) else bear_ob
    elif bull_ob:
        chosen_ob = bull_ob
    elif bear_ob:
        chosen_ob = bear_ob
    else:
        chosen_ob = None

    # Closest OB Bull: nearest active bullish order block.
    if bull_ob:
        closest_ob_str = _fmt_range(bull_ob["barLow"], bull_ob["barHigh"])
        ob_bull_age = int(n - 1 - int(bull_ob["idx"])) if bull_ob.get("idx") is not None else np.nan

    # Closest OB Bear: nearest active bearish order block.
    if bear_ob:
        closest_ob_bear_str = _fmt_range(bear_ob["barLow"], bear_ob["barHigh"])
        ob_bear_age = int(n - 1 - int(bear_ob["idx"])) if bear_ob.get("idx") is not None else np.nan

    # Equilibrium = midpoint of the closest active OB to current price.
    # This avoids mixing a bull midpoint with a nearer bear OB.
    if chosen_ob:
        ob_mid       = _ob_mid(chosen_ob)
        ob_equilibrium = ob_mid
        ob_direction = "Bullish" if chosen_ob["bias"] == 1 else "Bearish"
        if ob_mid > 0:
            ob_dist_pct = ((close_last / ob_mid) - 1.0) * 100.0

    # Closest FVG
    closest_fvg_str = ""
    fvg_bias_str    = "N/A"
    if fvgs:
        def fvg_dist(fvg):
            mid = (fvg["top"] + fvg["bottom"]) / 2
            return abs(close_last - mid)
        nearest_fvg = min(fvgs, key=fvg_dist)
        closest_fvg_str = _fmt_range(nearest_fvg["bottom"], nearest_fvg["top"])
        fvg_bias_str = "Bullish" if nearest_fvg["bias"] == 1 else "Bearish"

    # EQH/EQL status
    if eqh_detected and eql_detected:
        eqhl_status = "EQH + EQL Detected"
    elif eqh_detected:
        eqhl_status = "EQH Detected"
    elif eql_detected:
        eqhl_status = "EQL Detected"
    else:
        eqhl_status = "None"

    # Premium / Discount zone (Pine: trailing extremes define the range)
    top = trailing_top
    bot = trailing_bottom
    prem_bot     = 0.95 * top + 0.05 * bot if pd.notna(top) and pd.notna(bot) else np.nan
    eq_top_lvl   = 0.525 * top + 0.475 * bot if pd.notna(top) and pd.notna(bot) else np.nan
    eq_bot_lvl   = 0.525 * bot + 0.475 * top if pd.notna(top) and pd.notna(bot) else np.nan
    disc_top     = 0.95 * bot + 0.05 * top   if pd.notna(top) and pd.notna(bot) else np.nan

    pd_label = "N/A"
    if pd.notna(top) and pd.notna(bot):
        if chosen_ob:
            if ob_direction == "Bearish" and chosen_ob["barLow"] <= close_last <= chosen_ob["barHigh"]:
                pd_label = "In Bear OB"
            elif ob_direction == "Bullish" and chosen_ob["barLow"] <= close_last <= chosen_ob["barHigh"]:
                pd_label = "In Bull OB"
        if pd_label == "N/A":
            if close_last >= prem_bot:
                pd_label = "Premium"
            elif eq_bot_lvl <= close_last <= eq_top_lvl:
                pd_label = "Equilibrium"
            elif disc_top is not None and close_last <= disc_top:
                pd_label = "Discount"
            else:
                pd_label = "Mid-Zone"

    # Ribbon bias (EMA 25 vs EMA 50)
    close_s  = pd.Series(df["Close"].astype(float).values)
    ema25    = float(close_s.ewm(span=25, adjust=False).mean().iloc[-1]) if len(close_s) >= 25 else np.nan
    ema50    = float(close_s.ewm(span=50, adjust=False).mean().iloc[-1]) if len(close_s) >= 50 else np.nan
    if pd.notna(ema25) and pd.notna(ema50):
        ribbon_bias = "Bullish" if ema25 > ema50 else "Bearish"
    else:
        ribbon_bias = "N/A"

    # Current Quarter VWAP (developing)
    try:
        last_ts  = df.index[-1]
        q_start  = quarter_start(last_ts)
        df_q     = df.loc[df.index >= q_start]
        if not df_q.empty:
            tp_q = (df_q["High"] + df_q["Low"] + df_q["Close"]) / 3.0
            vv_q = df_q["Volume"].replace(0, np.nan)
            mask_q = tp_q.notna() & vv_q.notna()
            if mask_q.any():
                cur_qvwap = (tp_q[mask_q] * vv_q[mask_q]).sum() / vv_q[mask_q].sum()
            else:
                cur_qvwap = np.nan
        else:
            cur_qvwap = np.nan
    except Exception:
        cur_qvwap = np.nan

    qvwap_dist_pct = np.nan
    if pd.notna(cur_qvwap) and cur_qvwap > 0:
        qvwap_dist_pct = ((close_last / cur_qvwap) - 1.0) * 100.0

    # Nearest PY / PQ VWAP band (reuse anchored_vwap_block results from build_row)
    # We return the cur QVWAP; the actual PY/PQ values are already in build_row.
    # This engine just outputs the "nearest" label based on distance.
    nearest_pyq_band    = "N/A"
    pyq_band_dist_pct   = np.nan

    # ── SMC Composite State ───────────────────────────────────────────────────
    int_trend_lbl  = "Bullish" if internal_trend == 1 else ("Bearish" if internal_trend == -1 else "Neutral")
    sw_trend_lbl   = "Bullish" if swing_trend    == 1 else ("Bearish" if swing_trend    == -1 else "Neutral")

    if sw_trend_lbl == "Bullish" and int_trend_lbl == "Bullish" and pd_label == "Discount":
        composite = "Bullish Expansion"
    elif sw_trend_lbl == "Bullish" and int_trend_lbl == "Bullish":
        composite = "Bullish Continuation"
    elif sw_trend_lbl == "Bullish" and pd_label == "Premium":
        composite = "Distribution Risk"
    elif sw_trend_lbl == "Bearish" and int_trend_lbl == "Bearish":
        composite = "Bearish Expansion"
    elif sw_trend_lbl == "Bearish" and pd_label == "Discount":
        composite = "Bearish Breakdown"
    elif pd_label == "Discount" and int_trend_lbl == "Bullish":
        composite = "Accumulation"
    elif pd_label == "Equilibrium":
        composite = "Neutral Rotation"
    elif sw_trend_lbl == "Bullish" and int_trend_lbl == "Bearish":
        composite = "Bullish Compression"
    else:
        composite = "Neutral Rotation"

    # Market Structure Bias (overall)
    ms_bias = sw_trend_lbl

    # Event age
    event_age = ""
    if last_event_idx is not None:
        try:
            event_age = int(n - 1 - last_event_idx)
        except Exception:
            pass

    # ── Legacy compatibility aliases (used by _build_idx_vwap_shortlist) ──────
    bull_ob_legacy = next((ob for ob in internal_obs[:2] if ob["bias"] == 1), None)
    bear_ob_legacy = next((ob for ob in internal_obs[:2] if ob["bias"] == -1), None)

    latest_int_str = last_int_event or "N/A"
    latest_sw_str  = last_sw_event  or "N/A"
    latest_int_date = "N/A"
    latest_sw_date = "N/A"
    try:
        if last_int_event_idx is not None:
            latest_int_date = pd.Timestamp(df.index[int(last_int_event_idx)]).strftime("%Y-%m-%d")
    except Exception:
        latest_int_date = "N/A"
    try:
        if last_sw_event_idx is not None:
            latest_sw_date = pd.Timestamp(df.index[int(last_sw_event_idx)]).strftime("%Y-%m-%d")
    except Exception:
        latest_sw_date = "N/A"
    ms_phase = ("Expansion"  if ("BOS"  in (last_sw_event or "")) else
                "Reversal"   if ("CHoCH" in (last_sw_event or "")) else "Range")
    ms_swing_seq = ("HH/HL" if swing_trend == 1 else
                    "LH/LL" if swing_trend == -1 else "Mixed")
    ms_quality  = "Clean" if (len(internal_obs) > 0 or len(swing_obs) > 0) else "Developing"

    return {
        # ── New SMC columns ──────────────────────────────────────────────────
        "smc_internal_trend":         int_trend_lbl,
        "smc_swing_trend":            sw_trend_lbl,
        "smc_internal_bos_count":     int_bos_count,
        "smc_internal_choch_count":   int_choch_count,
        "smc_swing_bos_count":        sw_bos_count,
        "smc_swing_choch_count":      sw_choch_count,
        "smc_latest_internal_struct": latest_int_str,
        "smc_latest_internal_struct_date": latest_int_date,
        "smc_latest_swing_struct":    latest_sw_str,
        "smc_latest_swing_struct_date": latest_sw_date,
        "smc_strong_high":            safe_num(strong_high),
        "smc_weak_high":              safe_num(weak_high),
        "smc_strong_low":             safe_num(strong_low),
        "smc_weak_low":               safe_num(weak_low),
        "smc_closest_ob":             closest_ob_str,
        "smc_ob_equilibrium":         safe_num(ob_equilibrium),
        "smc_eq_high":                safe_num(eq_top_lvl) if pd.notna(eq_top_lvl) else np.nan,
        "smc_eq_low":                 safe_num(eq_bot_lvl) if pd.notna(eq_bot_lvl) else np.nan,
        "smc_closest_ob_bear":        closest_ob_bear_str,
        "smc_ob_bull_age":             safe_num(ob_bull_age),
        "smc_ob_bear_age":             safe_num(ob_bear_age),
        "smc_ob_dist_pct":            safe_num(ob_dist_pct),
        "smc_closest_fvg":            closest_fvg_str,
        "smc_fvg_bias":               fvg_bias_str,
        "smc_eqhl_status":            eqhl_status,
        "smc_premium_discount":       pd_label,
        "smc_premium_zone":           _fmt_range(prem_bot, top),
        "smc_discount_zone":          _fmt_range(bot, disc_top),
        "smc_summary":                ("-" if pd_label == "Mid-Zone" else f"Price is on {pd_label}" if pd_label not in ("N/A", "", None) else "N/A"),
        "smc_ribbon_bias":            ribbon_bias,
        "smc_current_qvwap":          safe_num(cur_qvwap),
        "smc_qvwap_dist_pct":         safe_num(qvwap_dist_pct),
        "smc_nearest_pyq_band":       nearest_pyq_band,
        "smc_pyq_band_dist_pct":      safe_num(pyq_band_dist_pct),
        "smc_market_structure_bias":  ms_bias,
        "smc_composite_state":        composite,
        # ── Legacy aliases (used by shortlist + detail schema) ────────────────
        "smc_bull_internal_ob":       _fmt_range(bull_ob_legacy["barLow"], bull_ob_legacy["barHigh"]) if bull_ob_legacy else "",
        "smc_bear_internal_ob":       _fmt_range(bear_ob_legacy["barLow"], bear_ob_legacy["barHigh"]) if bear_ob_legacy else "",
        "smc_discount":               _fmt_range(bot, disc_top),
        "smc_equilibrium":            _fmt_range(eq_bot_lvl, eq_top_lvl),
        "smc_premium":                _fmt_range(prem_bot, top),
        "smc_state":                  composite,
        # ── Market Structure aliases (used by _build_idx_vwap_shortlist) ──────
        "ms_trend_bias":              sw_trend_lbl,
        "ms_structure_phase":         ms_phase,
        "ms_last_structural_event":   latest_sw_str,
        "ms_event_age_d":             event_age,
        "ms_swing_sequence":          ms_swing_seq,
        "ms_structure_quality":       ms_quality,
        "ms_volume_confirmation":     "N/A",
        # Backward compat with _build_idx_vwap_shortlist
        "ms_trend_regime":            sw_trend_lbl,
        "ms_structure_state":         ms_swing_seq,
        "ms_last_event":              latest_sw_str,
        "ms_last_event_date":         latest_sw_date,
    }


# =========================
# DETAIL SCHEMA OVERRIDE — SMC v2 + MACD Enhanced columns
# =========================
try:
    _new_detail_schema_v2 = []
    for _grp_name, _grp_fill, _cols in DETAIL_SCHEMA:
        if _grp_name == "Market Structure":
            _new_detail_schema_v2.append((
                "Market Structure & SMC", _grp_fill, [
                    ("smc_internal_trend",         "Internal Trend",        14, None,    "center"),
                    ("smc_swing_trend",            "Swing Trend",           14, None,    "center"),
                    ("smc_latest_internal_struct", "Latest Internal Struct",22, None,    "center"),
                    ("smc_latest_swing_struct",    "Latest Swing Struct",   22, None,    "center"),
                    ("smc_strong_high",            "Strong High",           12, "#,##0", "center"),
                    ("smc_weak_high",              "Weak High",             12, "#,##0", "center"),
                    ("smc_premium_zone",           "Premium Zone",          18, None,    "center"),
                    ("smc_strong_low",             "Strong Low",            12, "#,##0", "center"),
                    ("smc_weak_low",               "Weak Low",              12, "#,##0", "center"),
                    ("smc_discount_zone",          "Discount Zone",         18, None,    "center"),
                    ("smc_closest_ob",             "Closest OB Bull",       16, None,    "center"),
                    ("smc_equilibrium",            "Equilibrium",           18, None,    "center"),
                    ("smc_closest_ob_bear",        "Closest OB Bear",       16, None,    "center"),
                    ("smc_ob_bull_age",            "OB Bull Age (Days)",    16, "#,##0", "center"),
                    ("smc_ob_bear_age",            "OB Bear Age (Days)",    16, "#,##0", "center"),
                    ("smc_summary",                "Summary",               42, None,    "left"),
                ]
            ))
        elif _grp_name == "MACD Momentum":
            _new_detail_schema_v2.append((
                "MACD Momentum", _grp_fill, [
                    ("macd_line",        "MACD Line",        12, "0.0000", "center"),
                    ("macd_signal_line", "Signal Line",      12, "0.0000", "center"),
                    ("macd_hist",        "Histogram (EMA3)", 14, "0.0000", "center"),
                    ("macd_position",    "Lines Position",   18, None,    "center"),
                    ("macd_wave",        "Wave Pattern",     22, None,    "center"),
                    ("macd_cross",       "MACD Cross",       16, None,    "center"),
                ]
            ))
        elif _grp_name in ("Structure / Wyckoff Proxy (Backtest)",
                           "Candlestick Patterns",
                           "SMC"):
            # These group names do not exist in DETAIL_SCHEMA; clauses kept for safety only.
            pass
        else:
            _new_detail_schema_v2.append((_grp_name, _grp_fill, _cols))
    DETAIL_SCHEMA = _new_detail_schema_v2
except Exception:
    pass

# build_row wrapper: call new compute_smc_engine
try:
    _orig_build_row_v2 = build_row
    def build_row(ksei_row: pd.Series, hist: pd.DataFrame, shares_fallback: float):
        row = _orig_build_row_v2(ksei_row, hist, shares_fallback)
        try:
            smc = compute_smc_engine(hist, internal_size=5, swing_size=50,
                                     eq_length=3, eq_threshold=0.1,
                                     mitigation_mode="close")
            if isinstance(smc, dict):
                row.update(smc)
        except Exception:
            for _k in ("smc_composite_state", "smc_swing_trend", "smc_internal_trend",
                       "smc_state", "ms_trend_bias", "ms_structure_quality",
                       "ms_last_structural_event", "ms_event_age_d", "ms_swing_sequence"):
                row.setdefault(_k, "N/A")
        return row
except Exception:
    pass

def _to_yf_symbol_for_setup(ticker: str) -> str:
    try:
        t = str(ticker or "").strip().upper()
        if not t:
            return ""
        return t if t.endswith(".JK") else f"{t}.JK"
    except Exception:
        return ""

def _get_hist_for_setup(row: dict, lookback: int = 260) -> Optional[pd.DataFrame]:
    try:
        ticker = row.get("ticker", row.get("Ticker", ""))
        yf_symbol = _to_yf_symbol_for_setup(ticker)
        if not yf_symbol:
            return None
        _dl2_md  = get_market_date()
        _dl2_end = (_dl2_md + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        hist = yf.download(
            yf_symbol,
            period="2y",
            end=_dl2_end,
            interval="1d",
            auto_adjust=False,
            progress=False,
            threads=False,
        )
        hist = normalize_history(hist)
        if hist is not None and not hist.empty:
            hist = hist[hist.index.normalize() <= _dl2_md]
        if hist is None or hist.empty:
            return None
        if lookback and len(hist) > lookback:
            hist = hist.tail(int(lookback)).copy()
        return hist
    except Exception:
        return None

def _detect_recent_candle_patterns(hist: pd.DataFrame) -> dict:
    out = {
        "bullish_engulfing": False,
        "inside_bar": False,
        "ibh": np.nan,
        "ibl": np.nan,
        "last_bar_range_pct": np.nan,
        "close_in_upper_third": False,
        "bullish_reversal_bar": False,
    }
    try:
        if hist is None or len(hist) < 2:
            return out
        prev = hist.iloc[-2]
        last = hist.iloc[-1]
        po, ph, pl, pc = map(float, [prev["Open"], prev["High"], prev["Low"], prev["Close"]])
        lo, lh, ll, lc = map(float, [last["Open"], last["High"], last["Low"], last["Close"]])

        prev_body_low, prev_body_high = min(po, pc), max(po, pc)
        last_body_low, last_body_high = min(lo, lc), max(lo, lc)

        out["bullish_engulfing"] = (pc < po) and (lc > lo) and (last_body_low <= prev_body_low) and (last_body_high >= prev_body_high)
        out["inside_bar"] = (lh < ph) and (ll > pl)
        if out["inside_bar"]:
            out["ibh"], out["ibl"] = ph, pl
        else:
            out["ibh"], out["ibl"] = max(ph, lh), min(pl, ll)

        rng = max(lh - ll, 0.0)
        out["last_bar_range_pct"] = ((rng / lc) * 100.0) if lc else np.nan
        out["close_in_upper_third"] = rng > 0 and lc >= (ll + (2.0 * rng / 3.0))
        out["bullish_reversal_bar"] = (lc > lo) and out["close_in_upper_third"]
        return out
    except Exception:
        return out

def _sw_get(r: dict, keys, default=np.nan):
    if isinstance(keys, str):
        keys = [keys]
    for k in keys:
        if k in r:
            v = r.get(k)
            if v is None:
                continue
            if isinstance(v, float) and np.isnan(v):
                continue
            s = str(v).strip()
            if s == "" or s.upper() == "N/A":
                continue
            return v
    return default

def _enrich_swing_row_fields(rr: dict) -> dict:
    rr = dict(rr)
    hist = None

    def _need(*keys):
        return all((_sw_get(rr, k, np.nan) is np.nan or (isinstance(_sw_get(rr, k, np.nan), float) and np.isnan(_sw_get(rr, k, np.nan)))) for k in keys)

    # Normalize common display aliases from detail schema
    verdict = _sw_get(rr, ["Verdict Weight Profiles", "Verdict Weight Profile", "verdict_weight_profile"], "N/A")
    if verdict != "N/A":
        rr["Verdict Weight Profile"] = verdict
        rr["Verdict Profile"] = verdict

    trend = _sw_get(rr, ["Trend Bias", "trend_regime"], "N/A")
    if trend != "N/A":
        rr["Trend Bias"] = trend

    lse = _sw_get(rr, ["Last Structural Event", "last_structural_event"], "N/A")
    if lse != "N/A":
        rr["Last Structural Event"] = lse

    smc = _sw_get(rr, ["SMC State", "smc_state", "smc_zone"], "N/A")
    if smc != "N/A":
        rr["SMC State"] = smc
        rr["smc_state"] = smc

    mp = _sw_get(rr, ["Market Profile", "market_profile_summary", "VWAP Zone"], "N/A")
    if mp != "N/A":
        rr["Market Profile"] = mp
        rr["market_profile_summary"] = mp

    ma = _sw_get(rr, ["MA Position", "ma_position_summary", "MA Zone"], "N/A")
    if ma != "N/A":
        rr["MA Position"] = ma
        rr["ma_position_summary"] = ma

    cp = _sw_get(rr, ["Candle Pattern", "Last Candlestick Patterns", "Pattern"], "N/A")
    if cp != "N/A":
        rr["Candle Pattern"] = cp
        rr["last_candle_pattern"] = cp

    rsi_v = _sw_get(rr, ["RSI Status", "rsi_status"], "N/A")
    if rsi_v != "N/A":
        rr["RSI Status"] = rsi_v
        rr["rsi_status"] = rsi_v

    div = _sw_get(rr, ["Divergence Signal", "divergence_signal"], "N/A")
    if div != "N/A":
        rr["Divergence Signal"] = div
        rr["divergence_signal"] = div

    macd = _sw_get(rr, ["MACD Wave", "Wave Pattern", "macd_wave_pattern"], "N/A")
    if macd != "N/A":
        rr["MACD Wave"] = macd
        rr["macd_wave_pattern"] = macd

    adr = _sw_get(rr, ["ADR %", "adr14_pct"], np.nan)
    atr = _sw_get(rr, ["ATR (14) %", "ATR14 %", "atr14_pct"], np.nan)

    if (pd.isna(safe_num(adr)) or pd.isna(safe_num(atr)) or _sw_get(rr, ["Candle Pattern", "Last Candlestick Patterns", "Pattern"], "N/A") == "N/A"):
        hist = _get_hist_for_setup(rr)

    if pd.isna(safe_num(adr)) and hist is not None and len(hist) >= 14:
        try:
            adr_val = (((hist["High"] - hist["Low"]) / hist["Close"].replace(0, np.nan)) * 100.0).tail(14).mean()
            if pd.notna(adr_val):
                rr["ADR %"] = float(adr_val)
                rr["adr14_pct"] = float(adr_val)
        except Exception:
            pass

    if pd.isna(safe_num(atr)) and hist is not None and len(hist) >= 15:
        try:
            prev_close = hist["Close"].shift(1)
            tr = pd.concat([
                hist["High"] - hist["Low"],
                (hist["High"] - prev_close).abs(),
                (hist["Low"] - prev_close).abs()
            ], axis=1).max(axis=1)
            atr_val = tr.rolling(14).mean().iloc[-1]
            close_last = safe_num(hist["Close"].iloc[-1])
            atr_pct = (atr_val / close_last * 100.0) if pd.notna(atr_val) and close_last else np.nan
            if pd.notna(atr_pct):
                rr["ATR (14) %"] = float(atr_pct)
                rr["ATR14 %"] = float(atr_pct)
                rr["atr14_pct"] = float(atr_pct)
        except Exception:
            pass

    if _sw_get(rr, ["Candle Pattern", "Last Candlestick Patterns", "Pattern"], "N/A") == "N/A" and hist is not None and len(hist) >= 2:
        try:
            patt = _detect_recent_candle_patterns(hist)
            label = "N/A"
            if patt.get("bullish_engulfing"):
                label = "Bullish Engulfing"
            elif patt.get("inside_bar"):
                label = "Inside Bar"
            elif patt.get("bullish_reversal_bar"):
                label = "Bullish Reversal Bar"
            rr["Candle Pattern"] = label
            rr["last_candle_pattern"] = label
        except Exception:
            pass

    return rr


# =========================
# FUNDAMENTAL KEY STATS SHEET  (Stockbit Key Stats style)
# =========================


def _yf_latest_statement_value(stmt: "pd.DataFrame | None", aliases, col=0, default=np.nan):
    """Return latest numeric statement value by trying multiple yfinance row aliases."""
    try:
        if stmt is None or stmt.empty:
            return default
        idx_map = {str(i).strip().lower(): i for i in stmt.index}
        for a in aliases:
            key = str(a).strip().lower()
            if key in idx_map:
                v = stmt.loc[idx_map[key]].iloc[col]
                return safe_num(v, default)
        # fuzzy fallback
        for a in aliases:
            ak = str(a).strip().lower()
            for k, real_idx in idx_map.items():
                if ak in k or k in ak:
                    v = stmt.loc[real_idx].iloc[col]
                    return safe_num(v, default)
    except Exception:
        pass
    return default


def _yf_ttm_statement_value(stmt: "pd.DataFrame | None", aliases, default=np.nan):
    """Sum the latest four quarterly values for TTM items."""
    try:
        if stmt is None or stmt.empty:
            return default
        idx_map = {str(i).strip().lower(): i for i in stmt.index}
        row = None
        for a in aliases:
            key = str(a).strip().lower()
            if key in idx_map:
                row = idx_map[key]; break
        if row is None:
            for a in aliases:
                ak = str(a).strip().lower()
                for k, real_idx in idx_map.items():
                    if ak in k or k in ak:
                        row = real_idx; break
                if row is not None:
                    break
        if row is None:
            return default
        vals = pd.to_numeric(stmt.loc[row].iloc[:4], errors="coerce").dropna()
        return float(vals.sum()) if len(vals) else default
    except Exception:
        return default


def _yf_q_yoy_growth(stmt: "pd.DataFrame | None", aliases, default=np.nan):
    """Latest quarter YoY growth versus same quarter prior year when yfinance has >=5 quarter columns."""
    try:
        if stmt is None or stmt.empty or stmt.shape[1] < 5:
            return default
        latest = _yf_latest_statement_value(stmt, aliases, 0, np.nan)
        prior_y = _yf_latest_statement_value(stmt, aliases, 4, np.nan)
        if pd.notna(latest) and pd.notna(prior_y) and float(prior_y) != 0:
            return (float(latest) / float(prior_y)) - 1.0
    except Exception:
        pass
    return default


def _compute_extended_fundamentals(sym: str, info: dict) -> dict:
    """
    Robust yfinance fundamental model for IDX Fundamental Details.

    Design:
    - Pulls yfinance financial statements using both modern get_* APIs and legacy properties.
    - Computes ratios from statements when info fields are missing.
    - Uses latest price history as fallback for price, market cap, 52W high/low, returns.
    - Missing values remain NaN; no fabricated values.
    - Percentage outputs are decimals for Excel percentage formatting.
    """
    out = {}

    def _is_num(v):
        try:
            return v is not None and pd.notna(v) and np.isfinite(float(v))
        except Exception:
            return False

    def _first_num(*vals, default=np.nan):
        for v in vals:
            try:
                if v is None:
                    continue
                if isinstance(v, str) and v.strip() in ("", "-", "N/A", "nan", "None"):
                    continue
                f = float(v)
                if pd.notna(f) and np.isfinite(f):
                    return f
            except Exception:
                continue
        return default

    def _stmt(tk, kind: str, quarterly: bool = True) -> pd.DataFrame:
        """
        kind: income | balance | cashflow
        Returns statement with line items as index and periods as columns.
        """
        candidates = []
        if kind == "income":
            candidates = [
                lambda: tk.get_income_stmt(freq="quarterly" if quarterly else "yearly"),
                lambda: tk.quarterly_income_stmt if quarterly else tk.income_stmt,
                lambda: tk.quarterly_financials if quarterly else tk.financials,
                lambda: tk.get_financials(freq="quarterly" if quarterly else "yearly"),
            ]
        elif kind == "balance":
            candidates = [
                lambda: tk.get_balance_sheet(freq="quarterly" if quarterly else "yearly"),
                lambda: tk.quarterly_balance_sheet if quarterly else tk.balance_sheet,
                lambda: tk.get_balancesheet(freq="quarterly" if quarterly else "yearly"),
            ]
        elif kind == "cashflow":
            candidates = [
                lambda: tk.get_cash_flow(freq="quarterly" if quarterly else "yearly"),
                lambda: tk.get_cashflow(freq="quarterly" if quarterly else "yearly"),
                lambda: tk.quarterly_cashflow if quarterly else tk.cashflow,
            ]
        for fn in candidates:
            try:
                df = fn()
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # yfinance returns periods descending in most cases; keep as-is but coerce numeric.
                    df = df.copy()
                    df = df.loc[~df.index.astype(str).duplicated(keep="first")]
                    return df
            except Exception:
                continue
        return pd.DataFrame()

    def _row(stmt: pd.DataFrame, aliases, col: int = 0, default=np.nan):
        try:
            if stmt is None or stmt.empty:
                return default
            idx_map = {str(i).strip().lower().replace("_", " "): i for i in stmt.index}
            alias_norm = [str(a).strip().lower().replace("_", " ") for a in aliases]

            # exact
            for a in alias_norm:
                if a in idx_map:
                    v = stmt.loc[idx_map[a]].iloc[col]
                    return safe_num(v, default)

            # remove spaces exact
            idx_map_compact = {k.replace(" ", ""): real for k, real in idx_map.items()}
            for a in alias_norm:
                ac = a.replace(" ", "")
                if ac in idx_map_compact:
                    v = stmt.loc[idx_map_compact[ac]].iloc[col]
                    return safe_num(v, default)

            # fuzzy
            for a in alias_norm:
                ac = a.replace(" ", "")
                for k, real_idx in idx_map.items():
                    kc = k.replace(" ", "")
                    if ac in kc or kc in ac:
                        v = stmt.loc[real_idx].iloc[col]
                        return safe_num(v, default)
        except Exception:
            pass
        return default

    def _ttm(stmt_q: pd.DataFrame, aliases, info_fallback=np.nan, stmt_a: pd.DataFrame = None, default=np.nan):
        v_info = _first_num(info_fallback, default=np.nan)
        if _is_num(v_info):
            return v_info
        try:
            if stmt_q is not None and not stmt_q.empty:
                idx_map = {str(i).strip().lower().replace("_", " "): i for i in stmt_q.index}
                alias_norm = [str(a).strip().lower().replace("_", " ") for a in aliases]
                row = None
                for a in alias_norm:
                    if a in idx_map:
                        row = idx_map[a]; break
                if row is None:
                    compact = {k.replace(" ", ""): real for k, real in idx_map.items()}
                    for a in alias_norm:
                        ac = a.replace(" ", "")
                        if ac in compact:
                            row = compact[ac]; break
                if row is None:
                    for a in alias_norm:
                        ac = a.replace(" ", "")
                        for k, real_idx in idx_map.items():
                            kc = k.replace(" ", "")
                            if ac in kc or kc in ac:
                                row = real_idx; break
                        if row is not None:
                            break
                if row is not None:
                    vals = pd.to_numeric(stmt_q.loc[row].iloc[:4], errors="coerce").dropna()
                    if len(vals):
                        return float(vals.sum())
            if stmt_a is not None and not stmt_a.empty:
                v = _row(stmt_a, aliases, 0, np.nan)
                if _is_num(v):
                    return float(v)
        except Exception:
            pass
        return default

    def _q_yoy(stmt_q: pd.DataFrame, aliases, default=np.nan):
        try:
            if stmt_q is None or stmt_q.empty or stmt_q.shape[1] < 5:
                return default
            latest = _row(stmt_q, aliases, 0, np.nan)
            prior = _row(stmt_q, aliases, 4, np.nan)
            if _is_num(latest) and _is_num(prior) and float(prior) != 0:
                return (float(latest) / float(prior)) - 1.0
        except Exception:
            pass
        return default

    def _history(tk, period="5y"):
        """Fetch history capped at MARKET_DATE."""
        try:
            _md  = get_market_date()
            _end = (_md + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
            h = tk.history(period=period, end=_end, interval="1d", auto_adjust=True)
            if h is not None and not h.empty:
                h = normalize_history(h)
                if not h.empty:
                    h = h[h.index.normalize() <= _md]
                return h
        except Exception:
            pass
        return pd.DataFrame()

    try:
        tk = yf.Ticker(sym)

        # Refresh info if caller passed an empty/stale dict
        if not isinstance(info, dict) or not info:
            try:
                info = tk.info or {}
            except Exception:
                info = {}

        qfin = _stmt(tk, "income", True)
        afin = _stmt(tk, "income", False)
        qbs  = _stmt(tk, "balance", True)
        abs_ = _stmt(tk, "balance", False)
        qcf  = _stmt(tk, "cashflow", True)
        acf  = _stmt(tk, "cashflow", False)

        hist = _history(tk, "5y")
        close_series = hist["Close"].dropna() if hist is not None and not hist.empty and "Close" in hist.columns else pd.Series(dtype=float)

        price = _first_num(
            info.get("currentPrice"),
            info.get("regularMarketPrice"),
            info.get("previousClose"),
            close_series.iloc[-1] if len(close_series) else np.nan,
        )

        shares = _first_num(
            info.get("sharesOutstanding"),
            info.get("impliedSharesOutstanding"),
            info.get("floatShares"),
        )
        float_shares = _first_num(info.get("floatShares"))

        # Income statement
        revenue_q = _row(qfin, ["Total Revenue", "Operating Revenue", "Revenue"])
        revenue_ttm = _ttm(
            qfin, ["Total Revenue", "Operating Revenue", "Revenue"],
            info.get("totalRevenue"), afin
        )
        gross_profit_q = _row(qfin, ["Gross Profit"])
        operating_income_q = _row(qfin, ["Operating Income", "Operating Income Loss", "EBIT"])
        ebit_ttm = _ttm(
            qfin, ["EBIT", "Operating Income", "Operating Income Loss"],
            np.nan, afin
        )
        ebitda_ttm = _ttm(
            qfin, ["EBITDA", "Normalized EBITDA"],
            info.get("ebitda"), afin
        )
        net_income_q = _row(qfin, ["Net Income", "Net Income Common Stockholders", "Net Income Applicable To Common Shares"])
        net_income_ttm = _ttm(
            qfin, ["Net Income", "Net Income Common Stockholders", "Net Income Applicable To Common Shares"],
            info.get("netIncomeToCommon"), afin
        )
        cogs_q = abs(_row(qfin, ["Cost Of Revenue", "Cost Of Goods Sold", "Cost Of Revenue"], default=np.nan))

        # Balance sheet
        total_assets = _row(qbs, ["Total Assets"], default=_row(abs_, ["Total Assets"]))
        total_liab = _row(qbs, [
            "Total Liabilities Net Minority Interest", "Total Liab", "Total Liabilities",
            "Total Liabilities Net Minority Interest"
        ], default=_row(abs_, ["Total Liabilities Net Minority Interest", "Total Liabilities", "Total Liab"]))
        total_equity = _row(qbs, [
            "Stockholders Equity", "Total Equity Gross Minority Interest", "Total Stockholder Equity",
            "Common Stock Equity"
        ], default=_row(abs_, ["Stockholders Equity", "Total Stockholder Equity", "Common Stock Equity"]))
        current_assets = _row(qbs, ["Current Assets", "Total Current Assets"])
        current_liab = _row(qbs, ["Current Liabilities", "Total Current Liabilities"])
        cash = _row(qbs, [
            "Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments",
            "Cash", "Cash Financial"
        ], default=info.get("totalCash", np.nan))
        inventory = _row(qbs, ["Inventory", "Inventories"])
        receivables = _row(qbs, ["Accounts Receivable", "Receivables", "Net Receivables"])
        payables = _row(qbs, ["Accounts Payable", "Payables"])
        short_debt = _row(qbs, [
            "Current Debt", "Short Term Debt", "Current Debt And Capital Lease Obligation",
            "Short Long Term Debt"
        ])
        long_debt = _row(qbs, [
            "Long Term Debt", "Long Term Debt And Capital Lease Obligation", "Long Term Debt Noncurrent"
        ])
        total_debt = _first_num(
            info.get("totalDebt"),
            (0 if pd.isna(short_debt) else short_debt) + (0 if pd.isna(long_debt) else long_debt)
            if pd.notna(short_debt) or pd.notna(long_debt) else np.nan
        )
        tangible_book = _row(qbs, ["Tangible Book Value"], default=np.nan)
        if not _is_num(tangible_book):
            goodwill = _row(qbs, ["Goodwill And Other Intangible Assets", "Goodwill", "Other Intangible Assets"], default=0)
            tangible_book = total_equity - goodwill if _is_num(total_equity) else np.nan
        retained_earnings = _row(qbs, ["Retained Earnings"], default=_row(abs_, ["Retained Earnings"]))

        # Cash flow
        cfo_q = _row(qcf, ["Operating Cash Flow", "Total Cash From Operating Activities", "Cash Flow From Continuing Operating Activities"])
        cfo_ttm = _ttm(qcf, ["Operating Cash Flow", "Total Cash From Operating Activities", "Cash Flow From Continuing Operating Activities"], info.get("operatingCashflow"), acf)
        cfi_ttm = _ttm(qcf, ["Investing Cash Flow", "Total Cashflows From Investing Activities", "Cash Flow From Continuing Investing Activities"], np.nan, acf)
        cff_ttm = _ttm(qcf, ["Financing Cash Flow", "Total Cash From Financing Activities", "Cash Flow From Continuing Financing Activities"], np.nan, acf)
        capex_q = _row(qcf, ["Capital Expenditure", "Capital Expenditures"])
        capex_ttm = _ttm(qcf, ["Capital Expenditure", "Capital Expenditures"], info.get("capitalExpenditures"), acf)
        fcf_q = _row(qcf, ["Free Cash Flow"], default=np.nan)
        if not _is_num(fcf_q) and _is_num(cfo_q) and _is_num(capex_q):
            fcf_q = cfo_q + capex_q  # yfinance capex is usually negative
        fcf_ttm = _first_num(info.get("freeCashflow"))
        if not _is_num(fcf_ttm):
            fcf_ttm = _ttm(qcf, ["Free Cash Flow"], np.nan, acf)
        if not _is_num(fcf_ttm) and _is_num(cfo_ttm) and _is_num(capex_ttm):
            fcf_ttm = cfo_ttm + capex_ttm

        # ── BUG-03: Apply USD → IDR conversion to all statement values ──────────
        # yfinance .JK stocks return financial statement data (income, balance,
        # cash-flow) in USD, while market_cap / enterprise_value from info are
        # already in IDR.  Multiply every statement value by the current USD/IDR
        # rate to make them consistent with market_cap before computing ratios.
        price_currency = str(info.get("currency") or "").upper().strip()
        financial_currency = str(info.get("financialCurrency") or "").upper().strip()
        should_convert_usd_to_idr = financial_currency == "USD" and price_currency == "IDR"
        FX = _fetch_usd_idr_rate() if should_convert_usd_to_idr else 1.0

        def _to_idr(v):
            return v * FX if _is_num(v) else v

        revenue_q           = _to_idr(revenue_q)
        revenue_ttm         = _to_idr(revenue_ttm)
        gross_profit_q      = _to_idr(gross_profit_q)
        operating_income_q  = _to_idr(operating_income_q)
        ebit_ttm            = _to_idr(ebit_ttm)
        ebitda_ttm          = _to_idr(ebitda_ttm)
        net_income_q        = _to_idr(net_income_q)
        net_income_ttm      = _to_idr(net_income_ttm)
        cogs_q              = _to_idr(cogs_q)
        total_assets        = _to_idr(total_assets)
        total_liab          = _to_idr(total_liab)
        total_equity        = _to_idr(total_equity)
        current_assets      = _to_idr(current_assets)
        current_liab        = _to_idr(current_liab)
        cash                = _to_idr(cash)
        inventory           = _to_idr(inventory)
        receivables         = _to_idr(receivables)
        payables            = _to_idr(payables)
        short_debt          = _to_idr(short_debt)
        long_debt           = _to_idr(long_debt)
        tangible_book       = _to_idr(tangible_book)
        retained_earnings   = _to_idr(retained_earnings)
        cfo_q               = _to_idr(cfo_q)
        cfo_ttm             = _to_idr(cfo_ttm)
        cfi_ttm             = _to_idr(cfi_ttm)
        cff_ttm             = _to_idr(cff_ttm)
        capex_q             = _to_idr(capex_q)
        capex_ttm           = _to_idr(capex_ttm)
        fcf_q               = _to_idr(fcf_q)
        fcf_ttm             = _to_idr(fcf_ttm)
        total_debt = _to_idr(total_debt)
        # Prefer statement components when available; they now share the
        # traded security's currency.
        total_debt = _first_num(
            (0 if pd.isna(short_debt) else short_debt) + (0 if pd.isna(long_debt) else long_debt)
            if pd.notna(short_debt) or pd.notna(long_debt) else np.nan,
            total_debt,
        )

        # Market cap and EV: already in IDR from info dict for .JK stocks
        market_cap = _first_num(
            info.get("marketCap"),
            price * shares if _is_num(price) and _is_num(shares) else np.nan
        )
        ev = _first_num(
            info.get("enterpriseValue"),
            market_cap + total_debt - cash if all(_is_num(x) for x in [market_cap, total_debt, cash]) else np.nan
        )
        free_float_pct = safe_div(float_shares, shares)

        # Valuation — all values now consistent in IDR
        pe_ttm = _first_num(info.get("trailingPE"), safe_div(market_cap, net_income_ttm))
        pe_ann = safe_div(market_cap, net_income_q * 4 if _is_num(net_income_q) else np.nan)
        earnings_yield = safe_div(1.0, pe_ttm)
        ps_ttm = _first_num(info.get("priceToSalesTrailing12Months"), safe_div(market_cap, revenue_ttm))
        pbv = _first_num(info.get("priceToBook"), safe_div(market_cap, total_equity))
        ev_ebit = safe_div(ev, ebit_ttm)
        ev_ebitda = _first_num(info.get("enterpriseToEbitda"), safe_div(ev, ebitda_ttm))

        # Profitability / management
        gross_margin_q = safe_div(gross_profit_q, revenue_q)
        op_margin_q = safe_div(operating_income_q, revenue_q)
        net_margin_q = safe_div(net_income_q, revenue_q)
        roa_ttm = _first_num(info.get("returnOnAssets"), safe_div(net_income_ttm, total_assets))
        roe_ttm = _first_num(info.get("returnOnEquity"), safe_div(net_income_ttm, total_equity))
        capital_employed = (total_assets - current_liab) if _is_num(total_assets) and _is_num(current_liab) else np.nan
        roce_ttm = safe_div(ebit_ttm, capital_employed)
        nopat = ebit_ttm * (1 - 0.22) if _is_num(ebit_ttm) else np.nan
        invested_capital = (total_debt + total_equity - cash) if all(_is_num(x) for x in [total_debt, total_equity, cash]) else np.nan
        roic_ttm = safe_div(nopat, invested_capital)
        asset_turnover = safe_div(revenue_ttm, total_assets)

        dso_q = safe_div(receivables, revenue_q) * 90 if pd.notna(safe_div(receivables, revenue_q)) else np.nan
        dio_q = safe_div(inventory, cogs_q) * 90 if pd.notna(safe_div(inventory, cogs_q)) else np.nan
        dpo_q = safe_div(payables, cogs_q) * 90 if pd.notna(safe_div(payables, cogs_q)) else np.nan
        ccc_q = dso_q + dio_q - dpo_q if all(pd.notna(x) for x in [dso_q, dio_q, dpo_q]) else np.nan
        receivables_turnover = safe_div(revenue_q * 4, receivables)

        # Solvency
        current_ratio = _first_num(info.get("currentRatio"), safe_div(current_assets, current_liab))
        quick_ratio = _first_num(
            info.get("quickRatio"),
            safe_div((current_assets - inventory) if _is_num(current_assets) and _is_num(inventory) else np.nan, current_liab)
        )
        debt_equity = _first_num(info.get("debtToEquity"), safe_div(total_debt, total_equity))
        if _is_num(debt_equity) and debt_equity > 10:  # yfinance sometimes reports 100x style percent points
            debt_equity = debt_equity / 100.0
        lt_debt_equity = safe_div(long_debt, total_equity)
        liabilities_equity = safe_div(total_liab, total_equity)
        financial_leverage = safe_div(total_assets, total_equity)
        interest_expense = abs(_to_idr(_ttm(qfin, ["Interest Expense", "Interest Expense Non Operating"], np.nan, afin)))
        interest_coverage = safe_div(ebit_ttm, interest_expense)

        working_capital = current_assets - current_liab if _is_num(current_assets) and _is_num(current_liab) else np.nan
        altman_original = (
            1.2 * safe_div(working_capital, total_assets, 0) +
            1.4 * safe_div(retained_earnings, total_assets, 0) +
            3.3 * safe_div(ebit_ttm, total_assets, 0) +
            0.6 * safe_div(market_cap, total_liab, 0) +
            1.0 * safe_div(revenue_ttm, total_assets, 0)
        ) if _is_num(total_assets) and total_assets != 0 else np.nan
        altman_modified = (
            6.56 * safe_div(working_capital, total_assets, 0) +
            3.26 * safe_div(retained_earnings, total_assets, 0) +
            6.72 * safe_div(ebit_ttm, total_assets, 0) +
            1.05 * safe_div(total_equity, total_liab, 0)
        ) if _is_num(total_assets) and total_assets != 0 else np.nan

        # Price performance
        def _ret_days(days):
            try:
                if len(close_series) <= days:
                    return np.nan
                return float(close_series.iloc[-1] / close_series.iloc[-days - 1] - 1)
            except Exception:
                return np.nan

        ytd_ret = np.nan
        try:
            if len(close_series):
                ystart = close_series[close_series.index.year == close_series.index[-1].year]
                if len(ystart) > 1:
                    ytd_ret = float(close_series.iloc[-1] / ystart.iloc[0] - 1)
        except Exception:
            pass

        book_value_q = total_equity
        bvps = safe_div(total_equity, shares)
        tbvps = safe_div(tangible_book, shares)
        eps_ttm = safe_div(net_income_ttm, shares)

        out.update({
            "shares_outstanding": shares,
            "float_shares": float_shares,
            "free_float_pct": free_float_pct,

            # Current Valuation
            "pe_annualised": pe_ann,
            "pe_ttm_current": pe_ttm,
            "earnings_yield_ttm": earnings_yield,
            "ps_ttm_current": ps_ttm,
            "pbv_current": pbv,
            "ev_ebit_ttm": ev_ebit,
            "ev_ebitda_ttm": ev_ebitda,
            "market_cap_current": market_cap,
            "enterprise_value_current": ev,
            "shares_outstanding_current": shares,

            # Profitability
            "gross_margin_q": gross_margin_q,
            "operating_margin_q": op_margin_q,
            "net_margin_q": net_margin_q,

            # Management Effectiveness
            "roa_ttm": roa_ttm,
            "roe_ttm": roe_ttm,
            "roce_ttm": roce_ttm,
            "roic_ttm": roic_ttm,
            "dso_q": dso_q,
            "asset_turnover_ttm": asset_turnover,
            "days_inventory_q": dio_q,
            "days_payables_q": dpo_q,
            "cash_conversion_cycle_q": ccc_q,
            "receivables_turnover_q": receivables_turnover,

            # Solvency
            "current_ratio_q": current_ratio,
            "quick_ratio_q": quick_ratio,
            "debt_equity_q": debt_equity,
            "lt_debt_equity_q": lt_debt_equity,
            "liabilities_equity_q": liabilities_equity,
            "financial_leverage_q": financial_leverage,
            "interest_coverage_ttm": interest_coverage,
            "free_cash_flow_q": fcf_q,
            "altman_z_original": altman_original,
            "altman_z_modified": altman_modified,

            # Price Performance
            "price_return_1m": _ret_days(21),
            "price_return_3m": _ret_days(63),
            "price_return_6m": _ret_days(126),
            "price_return_1y": _ret_days(252),
            "price_return_3y": _ret_days(756),
            "price_return_ytd": ytd_ret,
            "high_52w": float(close_series.tail(252).max()) if len(close_series) else np.nan,
            "low_52w": float(close_series.tail(252).min()) if len(close_series) else np.nan,

            # Balance Sheet
            "book_value_q": book_value_q,
            "book_value_per_share_current": bvps,
            "tangible_book_value_q": tangible_book,
            "tangible_book_value_per_share_current": tbvps,
            "short_term_debt_q": short_debt,
            "long_term_debt_q": long_debt,
            "cash_q": cash,
            "total_assets_q": total_assets,
            "total_liabilities_q": total_liab,

            # Income Statement
            "eps_ttm_current": eps_ttm,
            "eps_q_yoy_growth": _q_yoy(qfin, ["Diluted EPS", "Basic EPS", "Net Income", "Net Income Common Stockholders"]),
            "revenue_ttm": revenue_ttm,
            "revenue_q_yoy_growth": _q_yoy(qfin, ["Total Revenue", "Operating Revenue", "Revenue"]),
            "net_income_ttm": net_income_ttm,
            "ebit_ttm": ebit_ttm,

            # Cash Flow Statement
            "cash_from_operations_ttm": cfo_ttm,
            "cash_from_investing_ttm": cfi_ttm,
            "cash_from_financing_ttm": cff_ttm,
            "capital_expenditure_ttm": capex_ttm,
            "free_cash_flow_ttm": fcf_ttm,
            "_price_currency": price_currency or "UNKNOWN",
            "_financial_currency": financial_currency or "UNKNOWN",
            "_fx_conversion": "USD_TO_IDR" if should_convert_usd_to_idr else "NONE",
            "_fx_rate": FX if should_convert_usd_to_idr else np.nan,
        })
    except Exception:
        # Keep graceful failure; caller will leave missing values as N/A.
        pass

    return out


# ─────────────────────────────────────────────────────────────────────────────
# FUNDAMENTAL DATA — FALLBACK SOURCES
# Priority chain: yfinance → IDX Financial API → Investing.com scrape
# Each source is fully defensive (try/except) — no fabricated values.
# ─────────────────────────────────────────────────────────────────────────────

def _is_fund_data_empty(d: dict) -> bool:
    """
    Returns True if the core valuation/profitability fields are all NaN or missing.
    Used to decide whether to invoke the fallback chain.
    Critical fields: pe_ttm_current, pe_annualised, roe_ttm, revenue_ttm, net_margin_ttm, eps_ttm_current.
    """
    CRITICAL = [
        "pe_ttm_current",
        "pe_annualised",
        "roe_ttm",
        "revenue_ttm",
        "net_margin_q",
        "eps_ttm_current",
        "pbv_current",
        "market_cap_current",
    ]
    def _null(v):
        if v is None or v == "" or v == "N/A":
            return True
        try:
            return not np.isfinite(float(v)) or float(v) == 0
        except Exception:
            return True
    return all(_null(d.get(k)) for k in CRITICAL)


def _fetch_fundamental_idx_api(ticker: str) -> dict:
    """
    Fallback source 1: IDX Financial Data endpoint (idx.co.id public API).
    Returns a partial dict matching _fetch_fundamental_data() key names.
    Only populates fields the IDX API reliably provides.

    NOTE: The IDX endpoint availability and exact JSON schema may change.
          All fields remain NaN/N/A if the request fails or schema changes.
    """
    import urllib.request as _ur
    import json as _json

    out = {"_data_source": "IDX_API"}
    sym = ticker.upper().strip().replace(".JK", "")
    try:
        # IDX quarterly financial summary endpoint (public, no auth required)
        url = (
            f"https://idx.co.id/primary/TradingData/GetStockFinancialData"
            f"?KodeEmiten={sym}&Year=&Quarter="
        )
        req = _ur.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (IDX_Screener/4.0; research tool)",
            "Accept": "application/json",
            "Referer": "https://idx.co.id/",
        })
        with _ur.urlopen(req, timeout=10) as resp:
            raw = _json.loads(resp.read().decode("utf-8", errors="replace"))

        # IDX API shape: { "ResultCode": "Ok", "ResultData": { ... } }
        data = raw.get("ResultData") or raw.get("Data") or {}
        if not data:
            return out

        def _safe_f(key, default=np.nan):
            try:
                v = data.get(key)
                return float(v) if v is not None else default
            except Exception:
                return default

        # Map IDX API fields → internal field names
        out["pe_ttm"]           = _safe_f("PER")
        out["pe_ttm_current"]   = _safe_f("PER")
        out["pbv"]              = _safe_f("PBV")
        out["roe"]              = _safe_f("ROE") / 100 if _safe_f("ROE") not in (np.nan, None) else np.nan
        out["net_margin"]       = _safe_f("NPM") / 100 if _safe_f("NPM") not in (np.nan, None) else np.nan
        out["eps_ttm_current"]  = _safe_f("EPS")
        out["revenue_ttm"]      = _safe_f("Revenue")
        out["net_income_ttm"]   = _safe_f("NetIncome")
        out["debt_to_equity"]   = _safe_f("DER")
        out["operating_margin"] = _safe_f("OPM") / 100 if _safe_f("OPM") not in (np.nan, None) else np.nan
        out["book_value"]       = _safe_f("BookValue")
        out["market_cap"]       = _safe_f("MarketCap")
        out["dividend_yield"]   = _safe_f("DivYield") / 100 if _safe_f("DivYield") not in (np.nan, None) else np.nan

    except Exception as _e:
        print(f"[FUND_FALLBACK] IDX API failed for {sym}: {type(_e).__name__}: {_e}")

    return out


def _fetch_fundamental_investing_com(ticker: str) -> dict:
    """
    Fallback source 2: Investing.com key-stats page scrape.
    IMPORTANT: Investing.com has no official public API. This function
    reverse-engineers their internal JSON data endpoint that the web browser
    calls when loading equity pages. It is fragile — the endpoint, headers,
    or data schema may break without notice.

    If it fails, this function returns {} silently and the caller will use
    only yfinance + IDX API partial data (with NaN for missing fields).
    The scrape respects robots.txt spirit: one request per ticker, 2s delay,
    only on fallback (not on every run).
    """
    import time
    import urllib.request as _ur

    out = {"_data_source": "Investing.com"}
    sym = ticker.upper().strip().replace(".JK", "")

    try:
        import json as _json

        # Step 1: search for the instrument ID by ticker symbol
        search_url = (
            f"https://api.investing.com/api/search/v2/search"
            f"?q={sym}&domain=id&lang=en&limit=5&type=equities"
        )
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0",
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://www.investing.com/",
            "X-Requested-With": "XMLHttpRequest",
            "domain-id": "id",
        }
        req = _ur.Request(search_url, headers=headers)
        with _ur.urlopen(req, timeout=10) as resp:
            search_data = _json.loads(resp.read().decode("utf-8", errors="replace"))

        # Find first equity result matching IDX (.JK)
        articles = (search_data.get("articles") or
                    search_data.get("equities") or
                    search_data.get("quotes") or [])
        pair_id = None
        for item in articles:
            # Prefer items with exchange = 'IDX' or 'Indonesia'
            exch = str(item.get("exchange", "") or item.get("flag", "")).upper()
            ticker_match = str(item.get("symbol", "") or item.get("tag", "")).upper()
            if sym in ticker_match or "IDX" in exch or "INDONESIA" in exch:
                pair_id = item.get("id") or item.get("pairId") or item.get("pair_ID")
                if pair_id:
                    break

        if not pair_id:
            return out

        time.sleep(1.5)  # polite delay between requests

        # Step 2: fetch key-stats summary via Investing.com financials endpoint
        fin_url = (
            f"https://api.investing.com/api/financials/summary"
            f"?pairId={pair_id}&reportType=Annual&withFiscalData=true"
        )
        req2 = _ur.Request(fin_url, headers=headers)
        with _ur.urlopen(req2, timeout=12) as resp2:
            fin_data = _json.loads(resp2.read().decode("utf-8", errors="replace"))

        ratios = (fin_data.get("data") or fin_data.get("report") or
                  fin_data.get("keyStats") or fin_data.get("summary") or {})

        def _safe_f(key, alt_key=None, default=np.nan):
            for k in ([key, alt_key] if alt_key else [key]):
                if k and k in ratios:
                    try:
                        v = ratios[k]
                        if v is None:
                            continue
                        f = float(str(v).replace(",", "").replace("%", ""))
                        if np.isfinite(f):
                            return f
                    except Exception:
                        continue
            return default

        # Map Investing.com fields → internal field names
        # NOTE: Field key names vary by endpoint version; common observed keys listed below.
        #       If schema changed, fields will simply remain NaN (no crash).
        out["pe_ttm"]           = _safe_f("peRatio", "priceEarnings")
        out["pe_ttm_current"]   = _safe_f("peRatio", "priceEarnings")
        out["pbv"]              = _safe_f("priceBook", "pbRatio")
        out["roe"]              = (_safe_f("returnEquity", "roe") or np.nan)
        if not np.isnan(out.get("roe", np.nan)) and abs(out["roe"]) > 1.5:
            out["roe"] = out["roe"] / 100  # convert from % to decimal if needed
        out["roa"]              = _safe_f("returnAssets", "roa")
        if not np.isnan(out.get("roa", np.nan)) and abs(out["roa"]) > 1.5:
            out["roa"] = out["roa"] / 100
        out["eps_ttm_current"]  = _safe_f("eps", "earningsPerShare")
        out["revenue_ttm"]      = _safe_f("revenue", "totalRevenue")
        out["net_income_ttm"]   = _safe_f("netIncome")
        out["net_margin"]       = _safe_f("netProfitMargin", "profitMargin")
        if not np.isnan(out.get("net_margin", np.nan)) and abs(out["net_margin"]) > 1.5:
            out["net_margin"] = out["net_margin"] / 100
        out["operating_margin"] = _safe_f("operatingMargin")
        if not np.isnan(out.get("operating_margin", np.nan)) and abs(out["operating_margin"]) > 1.5:
            out["operating_margin"] = out["operating_margin"] / 100
        out["gross_margin"]     = _safe_f("grossMargin")
        if not np.isnan(out.get("gross_margin", np.nan)) and abs(out["gross_margin"]) > 1.5:
            out["gross_margin"] = out["gross_margin"] / 100
        out["debt_to_equity"]   = _safe_f("debtEquityRatio", "totalDebtEquity")
        out["current_ratio"]    = _safe_f("currentRatio")
        out["market_cap"]       = _safe_f("marketCap")
        out["enterprise_value"] = _safe_f("enterpriseValue")
        out["ev_ebitda"]        = _safe_f("evEbitda", "enterpriseValueEbitda")
        out["dividend_yield"]   = _safe_f("dividendYield")
        if not np.isnan(out.get("dividend_yield", np.nan)) and abs(out["dividend_yield"]) > 1.5:
            out["dividend_yield"] = out["dividend_yield"] / 100
        out["book_value"]       = _safe_f("bookValuePerShare", "bookValue")

    except Exception as _e:
        print(f"[FUND_FALLBACK] Investing.com scrape failed for {sym}: {type(_e).__name__}: {_e}")

    return out


def _merge_fallback_into_primary(primary: dict, fallback: dict) -> dict:
    """
    Merge fallback fields into primary dict ONLY where primary value is NaN/missing.
    Never overwrites a valid (non-NaN) primary value.
    Returns the merged dict with '_fallback_fields' listing what was supplemented.
    """
    supplemented = []

    def _is_null(v):
        if v is None or v == "" or v == "N/A":
            return True
        try:
            return not np.isfinite(float(v))
        except Exception:
            return True

    for k, v in fallback.items():
        if k.startswith("_"):
            continue  # skip metadata keys
        if k not in primary or _is_null(primary.get(k)):
            if not _is_null(v):
                primary[k] = v
                supplemented.append(k)

    if supplemented:
        primary["_fallback_source"] = fallback.get("_data_source", "unknown")
        primary["_fallback_fields"] = supplemented
        print(f"[FUND_FALLBACK] Supplemented {len(supplemented)} fields from "
              f"{fallback.get('_data_source','?')}: {supplemented[:8]}{'...' if len(supplemented)>8 else ''}")

    return primary


_YF_MIN_INTERVAL_SEC = 0.7   # floor spacing between yfinance requests
_yf_last_call_ts = 0.0


def _yf_pace() -> None:
    """Block just long enough to keep yfinance requests >= _YF_MIN_INTERVAL_SEC
    apart. A ~900-ticker scan calling yf.Ticker(...).info back-to-back with no
    spacing trips Yahoo's rate limit partway through the run, after which every
    remaining ticker comes back empty — this is the main reason fundamentals
    coverage was stuck around 35% even with retries on individual calls."""
    global _yf_last_call_ts
    now = time.monotonic()
    wait = _YF_MIN_INTERVAL_SEC - (now - _yf_last_call_ts)
    if wait > 0:
        time.sleep(wait)
    _yf_last_call_ts = time.monotonic()


_YF_DEGRADED_THRESHOLD = 15  # consecutive full-retry failures before assuming broad rate-limiting
_yf_consecutive_failures = 0


def _fetch_yf_info(sym: str, retries: int = 3, backoff_seconds: float = 4.0) -> dict:
    """yf.Ticker(sym).info, paced and retried with backoff. yfinance signals a
    rate-limited request with an EMPTY dict, not an exception, so a bare retry
    without checking for emptiness never fires — retry on empty here.

    Circuit breaker: a full 3-attempt/4s+8s-backoff retry costs ~14s per
    ticker, which is fine for occasional flakiness but not for a broadly
    rate-limited run -- across ~962 tickers that's ~3.7h by itself, which is
    exactly what was blowing through the daily pipeline's 3h CI ceiling and
    preventing it from ever reaching the OHLCV archive/screener-signal
    stages that actually matter, even on days the underlying price data was
    perfectly fetchable. Once _YF_DEGRADED_THRESHOLD tickers in a row have
    failed every attempt (a real, sustained block, not one flaky ticker),
    drop to a single fast attempt per ticker for the rest of the run --
    genuine data still gets a fair shot early on and again immediately after
    the first success (which resets the streak), but a dead run stops
    paying the full retry tax on every remaining ticker. Fundamentals still
    correctly fall through to the IDX API / Investing.com fallbacks (or
    N/A) exactly as before; only how hard yfinance itself gets retried
    changes."""
    global _yf_consecutive_failures
    degraded = _yf_consecutive_failures >= _YF_DEGRADED_THRESHOLD
    effective_retries = 1 if degraded else retries
    last_exc: Optional[Exception] = None
    for attempt in range(effective_retries):
        _yf_pace()
        try:
            info = yf.Ticker(sym).info or {}
            if info:
                _yf_consecutive_failures = 0
                return info
        except Exception as exc:
            last_exc = exc
        if attempt + 1 < effective_retries:
            time.sleep(backoff_seconds * (2 ** attempt))
    _yf_consecutive_failures += 1
    tag = " [degraded: broad rate-limit detected, retries reduced]" if degraded else ""
    if last_exc:
        print(f"[YF_RETRY] {sym}: empty/failed after {effective_retries} attempts ({last_exc}){tag}")
    else:
        print(f"[YF_RETRY] {sym}: empty info after {effective_retries} attempts{tag}")
    return {}


def _fetch_fundamental_data(ticker: str) -> dict:
    """Fetch yfinance fundamentals for a single IDX ticker."""
    try:
        sym = ticker.upper().strip()
        if not sym.endswith(".JK"):
            sym = sym + ".JK"
        tk = yf.Ticker(sym)
        info = _fetch_yf_info(sym)
        _ext = _compute_extended_fundamentals(sym, info)

        def _g(key, default=np.nan):
            v = info.get(key, default)
            if v is None:
                return default
            try:
                return float(v) if isinstance(v, (int, float)) else v
            except Exception:
                return v

        def _gs(key, default="N/A"):
            v = info.get(key, default)
            return str(v) if v is not None else default

        def _gi(key, default=np.nan):
            v = info.get(key, default)
            try:
                return int(v) if v is not None else default
            except Exception:
                return default

        # IPO date: try firstTradeDateEpochUtc, fallback to first available price bar
        _ipo_raw = info.get("firstTradeDateEpochUtc") or info.get("firstTradeDateMilliseconds")
        if _ipo_raw:
            try:
                _ts = float(_ipo_raw) / (1000 if float(_ipo_raw) > 1e10 else 1)
                _ipo_date = datetime.fromtimestamp(_ts, tz=timezone.utc).strftime("%Y-%m-%d")
            except Exception:
                _ipo_date = "N/A"
        else:
            try:
                _early_md  = get_market_date()
                _early_end = (_early_md + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
                _early = tk.history(period="max", end=_early_end, interval="1mo",
                                                  auto_adjust=True, progress=False)
                if _early is not None and not _early.empty:
                    _first = _early.index[0]
                    _ipo_date = _first.strftime("%Y-%m-%d") if hasattr(_first,"strftime") else str(_first)[:10]
                else:
                    _ipo_date = "N/A"
            except Exception:
                _ipo_date = "N/A"

        result = {
            "ticker":              ticker,
            "company_name":        _gs("longName", _gs("shortName", ticker)),
            "sector":              _gs("sector"),
            "industry":            _gs("industry"),
            "shares_outstanding":  _gi("sharesOutstanding"),
            "float_shares":        _gi("floatShares"),
            "ipo_date":            _ipo_date,
            # Valuation
            "market_cap":          _g("marketCap"),
            "enterprise_value":    _g("enterpriseValue"),
            "pe_ttm":              _g("trailingPE"),
            "forward_pe":          _g("forwardPE"),
            "pbv":                 _g("priceToBook"),
            "ps_ttm":              _g("priceToSalesTrailing12Months"),
            "ev_ebitda":           _g("enterpriseToEbitda"),
            "peg_ratio":           _g("trailingPegRatio"),
            "book_value":          _g("bookValue"),
            # Profitability
            "roe":                 _g("returnOnEquity"),
            "roa":                 _g("returnOnAssets"),
            "gross_margin":        _g("grossMargins"),
            "operating_margin":    _g("operatingMargins"),
            "net_margin":          _g("profitMargins"),
            "ebitda":              _g("ebitda"),
            # Growth
            "revenue_growth":      _g("revenueGrowth"),
            "earnings_growth":     _g("earningsGrowth"),
            "quarterly_revenue":   _g("totalRevenue"),
            "quarterly_earnings":  _g("netIncomeToCommon"),
            # Balance Sheet
            "total_assets":        _g("totalAssets"),
            "total_liabilities":   _g("totalDebt"),          # proxy
            "cash":                _g("totalCash"),
            "debt":                _g("totalDebt"),
            "debt_to_equity":      _g("debtToEquity"),
            "current_ratio":       _g("currentRatio"),
            # Cash Flow
            "operating_cf":        _g("operatingCashflow"),
            "free_cf":             _g("freeCashflow"),
            "capex":               _g("capitalExpenditures"),
            # Dividend
            # yfinance changed dividendYield from a decimal fraction (0.0569) to
            # a bare percentage number (5.69) -- normalize to a decimal here so
            # every consumer of this field gets the same unit consistently
            # (the Investing.com fallback path already does this same >1.5 check).
            "dividend_yield":      (lambda v: v / 100 if isinstance(v, (int, float)) and abs(v) > 1.5 else v)(_g("dividendYield")),
            "payout_ratio":        _g("payoutRatio"),
            "dividend_rate":       _g("dividendRate"),
            "trailing_annual_div": _g("trailingAnnualDividendRate"),
            "ex_dividend_date":    _gs("exDividendDate", "N/A"),
            "last_dividend_value": _g("lastDividendValue"),
            "last_dividend_date":  _gs("lastDividendDate", "N/A"),
            # Business Summary
            "business_summary":    _gs("longBusinessSummary", _gs("shortBusinessSummary", "N/A")),
            "_data_source":        "yfinance",
            **_ext,
        }

        # ── Fallback chain ────────────────────────────────────────────────
        # If yfinance returned empty/NaN critical fields, try IDX API then Investing.com.
        # Each fallback ONLY fills NaN gaps — it never overwrites valid yfinance values.
        if _is_fund_data_empty(result):
            print(f"[FUND_FALLBACK] yfinance returned empty data for {sym} → trying IDX API …")
            idx_data = _fetch_fundamental_idx_api(ticker)
            result = _merge_fallback_into_primary(result, idx_data)

        if _is_fund_data_empty(result):
            print(f"[FUND_FALLBACK] IDX API also empty for {sym} → trying Investing.com …")
            inv_data = _fetch_fundamental_investing_com(ticker)
            result = _merge_fallback_into_primary(result, inv_data)

        if _is_fund_data_empty(result):
            print(f"[FUND_FALLBACK] All sources empty for {sym} — fundamental fields will show N/A.")

        return result
    except Exception:
        return {"ticker": ticker, "company_name": ticker}


def _build_dividend_entries(ticker: str) -> dict:
    """
    Dividend layout:
    Upcoming Dividend and Latest Dividend both use:
    Year | Dividend (IDR) | Payout Ratio (%) | Dividend Yield (%) | Ex Date | Pay Date

    Upcoming is populated only when yfinance reports a future ex-date.
    Latest is populated from the latest historical dividend. Missing values remain "-".
    """
    keys = [
        "div1_year","div1_idr","div1_payout_ratio","div1_dividend_yield","div1_exdate","div1_paydate",
        "div2_year","div2_idr","div2_payout_ratio","div2_dividend_yield","div2_exdate","div2_paydate",
    ]
    empty = {k: "-" for k in keys}
    try:
        sym = ticker.upper().strip()
        if not sym.endswith(".JK"):
            sym += ".JK"
        tk  = yf.Ticker(sym)
        info = tk.info or {}
        divs = tk.dividends

        if divs is not None and not divs.empty:
            if hasattr(divs.index, "tz") and divs.index.tz is not None:
                divs.index = divs.index.tz_localize(None)

        def _fmt_date(raw):
            if raw is None:
                return "-"
            try:
                return pd.Timestamp(raw, unit="s").strftime("%Y-%m-%d")
            except Exception:
                try:
                    return pd.Timestamp(raw).strftime("%Y-%m-%d")
                except Exception:
                    return "-"

        price = safe_num(info.get("currentPrice", info.get("regularMarketPrice", np.nan)), np.nan)
        payout_ratio = safe_num(info.get("payoutRatio"), np.nan)      # decimal
        div_yield_ttm = safe_num(info.get("dividendYield"), np.nan)   # decimal

        # Upcoming: only future ex-date.
        upcoming_amount = safe_num(info.get("dividendRate", info.get("lastDividendValue", np.nan)), np.nan)
        upcoming_ex_raw = info.get("exDividendDate")
        upcoming_ex = _fmt_date(upcoming_ex_raw)
        try:
            upcoming_dt = pd.Timestamp(upcoming_ex).date() if upcoming_ex != "-" else None
            today_dt = pd.Timestamp.now(tz="Asia/Jakarta").date()
        except Exception:
            upcoming_dt = None
            today_dt = get_market_date().date()  # MARKET_DATE cutoff
        has_upcoming = bool(pd.notna(upcoming_amount) and upcoming_amount > 0 and upcoming_dt is not None and upcoming_dt > today_dt)

        out = dict(empty)
        if has_upcoming:
            out["div1_year"] = str(upcoming_dt.year)
            out["div1_idr"] = upcoming_amount
            out["div1_payout_ratio"] = payout_ratio if pd.notna(payout_ratio) else "-"
            out["div1_dividend_yield"] = safe_div(upcoming_amount, price) if pd.notna(price) and price else "-"
            out["div1_exdate"] = upcoming_ex
            out["div1_paydate"] = "-"

        # Latest historical dividend: latest actual cash dividend only.
        if divs is not None and not divs.empty:
            last_date = divs.index[-1]
            last_amt = float(divs.iloc[-1])
            out["div2_year"] = str(last_date.year)
            out["div2_idr"] = last_amt
            out["div2_payout_ratio"] = payout_ratio if pd.notna(payout_ratio) else "-"
            # latest one-payment yield; TTM dividend yield remains available in yfinance but this column is event-based.
            out["div2_dividend_yield"] = safe_div(last_amt, price) if pd.notna(price) and price else (div_yield_ttm if pd.notna(div_yield_ttm) else "-")
            out["div2_exdate"] = last_date.strftime("%Y-%m-%d")
            out["div2_paydate"] = "-"

        return out
    except Exception:
        return {k: "-" for k in keys}


def _fetch_dividend_history(ticker: str) -> dict:
    """
    Fetch dividend history from yfinance. Returns per-year dividends (2023-2025),
    CAGR, consecutive years, upcoming ex-date and yield estimates.
    Falls back to N/A gracefully.
    """
    empty = {
        "div_2025": np.nan, "div_2024": np.nan, "div_2023": np.nan,
        "div_cagr": np.nan, "div_consecutive_years": 0,
        "upcoming_dividend": np.nan, "upcoming_ex_date": "N/A",
        "upcoming_pay_date": "N/A", "upcoming_div_yield": np.nan,
    }
    try:
        sym = ticker.upper().strip()
        if not sym.endswith(".JK"):
            sym += ".JK"
        tk = yf.Ticker(sym)
        divs = tk.dividends
        info = tk.info or {}

        if divs is None or divs.empty:
            return empty

        # Normalize timezone
        if hasattr(divs.index, "tz") and divs.index.tz is not None:
            divs.index = divs.index.tz_localize(None)

        def _year_total(yr):
            mask = divs.index.year == yr
            return float(divs[mask].sum()) if mask.any() else np.nan

        div_2025 = _year_total(2025)
        div_2024 = _year_total(2024)
        div_2023 = _year_total(2023)

        # Consecutive years with dividend
        annual_years = sorted({dt.year for dt in divs.index}, reverse=True)
        consecutive = 0
        prev = None
        for yr in annual_years:
            if prev is None or yr == prev - 1:
                consecutive += 1
                prev = yr
            else:
                break

        # CAGR: use last 3 full years
        cagr = np.nan
        if pd.notna(div_2023) and pd.notna(div_2025) and div_2023 > 0:
            cagr = ((div_2025 / div_2023) ** (1 / 2) - 1) * 100

        # Upcoming dividend from yfinance info
        upcoming_dividend = float(info.get("dividendRate") or np.nan) if info.get("dividendRate") else np.nan
        ex_div_raw = info.get("exDividendDate")
        upcoming_ex_date = "N/A"
        if ex_div_raw:
            try:
                upcoming_ex_date = pd.Timestamp(ex_div_raw, unit="s").strftime("%Y-%m-%d")
            except Exception:
                try:
                    upcoming_ex_date = str(ex_div_raw)[:10]
                except Exception:
                    pass

        curr_price = float(info.get("regularMarketPrice") or info.get("currentPrice") or np.nan)
        upcoming_div_yield = (upcoming_dividend / curr_price * 100) if pd.notna(upcoming_dividend) and pd.notna(curr_price) and curr_price > 0 else np.nan

        return {
            "div_2025": div_2025,
            "div_2024": div_2024,
            "div_2023": div_2023,
            "div_cagr": cagr,
            "div_consecutive_years": consecutive,
            "upcoming_dividend": upcoming_dividend,
            "upcoming_ex_date": upcoming_ex_date,
            "upcoming_pay_date": "N/A",  # not available from yfinance
            "upcoming_div_yield": upcoming_div_yield,
        }
    except Exception:
        return empty


def _fetch_idx_disclosure(ticker: str) -> dict:
    """
    Fetch latest Keterbukaan Informasi disclosure from IDX.
    Scrapes https://www.idx.co.id/id/perusahaan-tercatat/keterbukaan-informasi/
    Returns: date, title, url, category, sentiment, days_since, event_risk.
    """
    import urllib.request
    import re
    from datetime import date

    empty = {
        "disclosure_date": "N/A", "disclosure_title": "N/A",
        "disclosure_url": "N/A", "disclosure_category": "N/A",
        "disclosure_sentiment": "N/A", "days_since_disclosure": np.nan,
        "event_risk": "N/A",
    }

    KEYWORD_CATEGORY = {
        "rights issue": ("Dilution", "Bearish", "High"),
        "penawaran umum terbatas": ("Dilution", "Bearish", "High"),
        "hmetd": ("Dilution", "Bearish", "High"),
        "akuisisi": ("Expansion", "Neutral", "Medium"),
        "acquisition": ("Expansion", "Neutral", "Medium"),
        "dividen": ("Shareholder Return", "Bullish", "Low"),
        "dividend": ("Shareholder Return", "Bullish", "Low"),
        "pemecahan saham": ("Liquidity Positive", "Bullish", "Low"),
        "stock split": ("Liquidity Positive", "Bullish", "Low"),
        "restrukturisasi": ("Transitional", "Neutral", "Medium"),
        "restructuring": ("Transitional", "Neutral", "Medium"),
        "pailit": ("Distress", "Bearish", "High"),
        "bankruptcy": ("Distress", "Bearish", "High"),
        "penawaran tender": ("Corporate Action", "Neutral", "Medium"),
        "tender offer": ("Corporate Action", "Neutral", "Medium"),
    }

    try:
        # IDX Keterbukaan Informasi — search title for [TICKER] pattern
        clean = ticker.upper().replace(".JK", "").strip()
        url = "https://www.idx.co.id/id/perusahaan-tercatat/keterbukaan-informasi/"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8", errors="ignore")

        # Find rows whose title contains [TICKER ] or [TICKER]
        ticker_pattern = re.compile(
            rf'\[{re.escape(clean)}\s*\]',
            re.IGNORECASE
        )

        # Extract all disclosure rows: date + title + url
        row_pattern = re.compile(
            r'(\d{2}\s+\w+\s+\d{4}\s+\d{2}:\d{2}:\d{2})\s+(.*?)\s*(?:href="([^"]+)")?',
            re.DOTALL
        )

        # Simpler extraction: find any line containing [TICKER]
        lines = html.split("\n")
        disc_date_str = "N/A"
        disc_title    = "N/A"
        disc_url      = "N/A"

        for i, line in enumerate(lines):
            if ticker_pattern.search(line):
                # Try to extract date from nearby lines
                context = " ".join(lines[max(0,i-3):i+3])
                date_m = re.search(r'(\d{2}\s+\w{3}\w*\s+\d{4}\s+\d{2}:\d{2}:\d{2})', context)
                if date_m:
                    disc_date_str = date_m.group(1).strip()
                # Extract title — take the line containing the ticker
                title_clean = re.sub(r'<[^>]+>', '', line).strip()
                if len(title_clean) > 10:
                    disc_title = title_clean[:150]
                # Extract URL if present
                url_m = re.search(r'href="([^"]+keterbukaan[^"]*)"', context, re.IGNORECASE)
                if url_m:
                    disc_url = url_m.group(1)
                    if not disc_url.startswith("http"):
                        disc_url = "https://www.idx.co.id" + disc_url
                # Format: "DD Mon YYYY HH:MM:SS Title [TICKER ]"
                if disc_title == "N/A":
                    disc_title = f"{disc_date_str} {title_clean}"[:150]
                break

        # Parse date
        disc_date = None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                disc_date = pd.Timestamp(disc_date_str, dayfirst=True).date()
                break
            except Exception:
                pass
        if disc_date is None:
            disc_date_str = "N/A"
            days_since = np.nan
        else:
            days_since = (get_market_date().date() - disc_date).days
            disc_date_str = disc_date.strftime("%Y-%m-%d")

        # Classify
        title_lower = disc_title.lower()
        category = sentiment = risk = "N/A"
        for kw, (cat, sent, ev_risk) in KEYWORD_CATEGORY.items():
            if kw in title_lower:
                category, sentiment, risk = cat, sent, ev_risk
                break
        if category == "N/A":
            category, sentiment, risk = "General", "Neutral", "Low"

        if not disc_url.startswith("http"):
            disc_url = "https://www.idx.co.id" + disc_url

        return {
            "disclosure_date": disc_date_str,
            "disclosure_title": disc_title[:120],
            "disclosure_url": disc_url,
            "disclosure_category": category,
            "disclosure_sentiment": sentiment,
            "days_since_disclosure": days_since,
            "event_risk": risk,
        }
    except Exception:
        return empty


# =============================================================================
# IDX DISCLOSURE — Google News RSS fetcher
# Replaces direct IDX scrape for the standalone IDX Disclosure sheet.
# Primary: most recent news on or before MARKET_DATE (30-day window).
# IFNA:    globally latest item if nothing within window.
# =============================================================================

import urllib.request as _urllib_req
import xml.etree.ElementTree as _ET
from urllib.parse import quote as _urlquote

_DISC_KEYWORD_MAP = [
    # (substring, category, sentiment, event_risk)  — priority order, first match wins
    ("rights issue",            "Dilution",           "Bearish",  "High"),
    ("penawaran umum terbatas", "Dilution",           "Bearish",  "High"),
    ("hmetd",                   "Dilution",           "Bearish",  "High"),
    ("pailit",                  "Distress",           "Bearish",  "High"),
    ("bankruptcy",              "Distress",           "Bearish",  "High"),
    ("gagal bayar",             "Distress",           "Bearish",  "High"),
    ("pkpu",                    "Distress",           "Bearish",  "High"),
    ("suspensi",                "Distress",           "Bearish",  "High"),
    ("delisting",               "Distress",           "Bearish",  "High"),
    ("akuisisi",                "Expansion",          "Bullish",  "Medium"),
    ("acquisition",             "Expansion",          "Bullish",  "Medium"),
    ("merger",                  "Transitional",       "Neutral",  "Medium"),
    ("penggabungan",            "Transitional",       "Neutral",  "Medium"),
    ("restrukturisasi",         "Transitional",       "Neutral",  "Medium"),
    ("restructuring",           "Transitional",       "Neutral",  "Medium"),
    ("divestasi",               "Transitional",       "Neutral",  "Medium"),
    ("waran",                   "Dilution",           "Bearish",  "Medium"),
    ("obligasi",                "Dilution",           "Neutral",  "Medium"),
    ("bond",                    "Dilution",           "Neutral",  "Medium"),
    ("sukuk",                   "Dilution",           "Neutral",  "Medium"),
    ("penawaran tender",        "Corporate Action",   "Neutral",  "Medium"),
    ("tender offer",            "Corporate Action",   "Neutral",  "Medium"),
    ("dividen",                 "Shareholder Return", "Bullish",  "Low"),
    ("dividend",                "Shareholder Return", "Bullish",  "Low"),
    ("pemecahan saham",         "Liquidity Positive", "Bullish",  "Low"),
    ("stock split",             "Liquidity Positive", "Bullish",  "Low"),
    ("saham bonus",             "Liquidity Positive", "Bullish",  "Low"),
    ("bonus share",             "Liquidity Positive", "Bullish",  "Low"),
    ("buyback",                 "Shareholder Return", "Bullish",  "Low"),
    ("pembelian kembali",       "Shareholder Return", "Bullish",  "Low"),
    ("ekspansi",                "Expansion",          "Bullish",  "Low"),
    ("capex",                   "Expansion",          "Bullish",  "Low"),
]


def _classify_disclosure_rss(title: str):
    """(category, sentiment, event_risk) from keyword priority map. Fallback = General/Neutral/Low."""
    t = (title or "").lower()
    for kw, cat, sent, risk in _DISC_KEYWORD_MAP:
        if kw in t:
            return cat, sent, risk
    return "General", "Neutral", "Low"


def _fetch_idx_disclosure_rss(ticker: str, market_date: str) -> dict:
    """
    Fetch news for a ticker from Google News RSS.

    PRIMARY  : most recent item with pubDate <= market_date (max 30-day lookback).
    IFNA     : if nothing within window, use globally latest item.
    disc_source field: 'On-Date' | 'IFNA-Latest' | 'N/A'
    """
    from email.utils import parsedate_to_datetime
    import datetime as _dti

    EMPTY = {
        "disc_date": "-", "disc_title": "-", "disc_url": "-",
        "disc_category": "-", "disc_sentiment": "-",
        "disc_days_since": "-", "disc_event_risk": "-",
        "disc_source": "N/A",
    }

    try:
        mdate = datetime.strptime(market_date, "%Y-%m-%d").date()
    except Exception:
        mdate = None

    def _rss_items(query_str: str):
        try:
            url = ("https://news.google.com/rss/search?"
                   f"q={_urlquote(query_str)}&hl=id&gl=ID&ceid=ID:id")
            req = _urllib_req.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with _urllib_req.urlopen(req, timeout=10) as resp:
                raw = resp.read()
            root = _ET.fromstring(raw)
            parsed = []
            for item in root.findall(".//item")[:25]:
                t_el = item.find("title");  l_el = item.find("link")
                d_el = item.find("pubDate")
                title = (t_el.text or "").strip() if t_el is not None else ""
                link  = (l_el.text or "").strip() if l_el is not None else ""
                pub   = (d_el.text or "").strip() if d_el is not None else ""
                try:
                    pub_dt = parsedate_to_datetime(pub).date()
                except Exception:
                    pub_dt = None
                parsed.append({
                    "date":     pub_dt,
                    "date_str": pub_dt.strftime("%Y-%m-%d") if pub_dt else "-",
                    "title":    title,
                    "url":      link,
                })
            return parsed
        except Exception:
            return []

    sym = ticker.upper().strip().replace(".JK", "")
    all_items: list = []
    for q in [
        f"{sym} keterbukaan informasi bursa efek",
        f"{sym} saham IDX BEI disclosure",
    ]:
        all_items.extend(_rss_items(q))
        if len(all_items) >= 10:
            break

    if not all_items:
        return EMPTY

    # Deduplicate by URL/title
    seen: set = set()
    deduped: list = []
    for it in all_items:
        k = it["url"] or it["title"]
        if k not in seen:
            seen.add(k); deduped.append(it)

    deduped.sort(key=lambda x: x["date"] or _dti.date.min, reverse=True)

    chosen = None
    source = "N/A"

    if mdate:
        window_start = mdate - _dti.timedelta(days=30)
        cands = [it for it in deduped if it["date"] and window_start <= it["date"] <= mdate]
        if cands:
            chosen = cands[0]; source = "On-Date"

    if chosen is None:
        valid = [it for it in deduped if it["date"] is not None]
        if valid:
            chosen = valid[0]; source = "IFNA-Latest"
        elif deduped:
            chosen = deduped[0]; source = "IFNA-Latest"

    if chosen is None:
        return EMPTY

    cat, sent, risk = _classify_disclosure_rss(chosen["title"])
    try:
        ref = mdate or get_market_date().date()
        days_since = (ref - chosen["date"]).days if chosen["date"] else "-"
    except Exception:
        days_since = "-"

    return {
        "disc_date":       chosen["date_str"],
        "disc_title":      chosen["title"][:160],
        "disc_url":        chosen["url"],
        "disc_category":   cat,
        "disc_sentiment":  sent,
        "disc_days_since": days_since,
        "disc_event_risk": risk,
        "disc_source":     source,
    }


# =============================================================================
# SCREENER INLINE NEWS COLUMNS
# Two lightweight fetchers used by build_idx_screener_sheet (inline cells).
# Cached per-run by ticker — not re-fetched across filter sections.
#
#  Col 1: News Sentiment  — general market sentiment for the ticker
#  Col 2: Corp. Action    — corporate-action-specific news only,
#                           IFNA fallback to IDX keterbukaan informasi page
# =============================================================================

_CORP_ACTION_KEYWORDS = [
    # Buyback
    "buyback", "pembelian kembali saham",
    # Rights / dilution
    "rights issue", "hmetd", "penawaran umum terbatas", "put vi",
    "private placement", "penerbitan saham baru",
    # Dividend
    "dividen", "dividend", "cum date", "ex date", "record date",
    # Management
    "pergantian direksi", "pergantian komisaris", "rups", "rapat umum pemegang saham",
    "management change", "ceo", "direktur utama", "komisaris utama",
    # Subsidiary / M&A
    "entitas anak", "subsidiary", "akuisisi", "merger", "divestasi",
    "divestiture", "spin off", "penggabungan usaha",
    # Shareholders
    "pemegang saham", "shareholder", "kepemilikan", "perubahan komposisi",
    # Bonds / warrants
    "obligasi", "waran", "bond", "sukuk", "mtn", "convertible",
    # Tender / corporate action
    "tender offer", "penawaran tender", "stock split", "pemecahan saham",
    "saham bonus", "bonus share",
    # Distress
    "pkpu", "pailit", "gagal bayar", "default", "restrukturisasi",
]

# Sentiment-only query terms (general market news)
_SENTIMENT_QUERY_TERMS = ["saham", "laporan keuangan", "kinerja", "target harga"]

# IDX keterbukaan URL — fallback for corp action (scrapes the first disclosure)
_IDX_CORP_ACTION_URL = "https://www.idx.co.id/id/perusahaan-tercatat/keterbukaan-informasi/?code={code}"



# =============================================================================
# NEWS ENTITY VALIDATION  — audit-grade ticker/company matching
# =============================================================================
NEWS_ENTITY_ALIAS_MAP: dict = {}
NEWS_ALL_TICKERS: set = set()
NEWS_REJECTION_AUDIT: list = []
GENERATOR_VERSION = "audit_patch_2026-05-29"

_COMPANY_ALIAS_OVERRIDES = {
    "BBRI": ["BBRI", "BRI", "BANK RAKYAT INDONESIA", "BANK RAKYAT"],
    "BMRI": ["BMRI", "BANK MANDIRI", "MANDIRI"],
    "BBCA": ["BBCA", "BCA", "BANK CENTRAL ASIA"],
    "BBNI": ["BBNI", "BNI", "BANK NEGARA INDONESIA"],
    "TOBA": ["TOBA", "TBS ENERGI", "TBS ENERGI UTAMA"],
    "ISAT": ["ISAT", "INDOSAT", "INDOSAT OOREDOO"],
    "CMNP": ["CMNP", "CITRA MARGA NUSAPHALA", "CITRA MARGA"],
    "PYFA": ["PYFA", "PYRIDAM FARMA", "PYRIDAM"],
    "WIFI": ["WIFI", "SOLUSI SINERGI DIGITAL", "SURGE"],
}

_COMPANY_STOPWORDS = {
    "PT", "TBK", "PERSERO", "TERBUKA", "INDONESIA", "INDO", "THE",
    "DAN", "AND", "CORP", "CORPORATION", "COMPANY", "CO", "LTD",
    "SAHAM", "RUPST", "RUPS", "RIGHTS", "ISSUE", "BUYBACK", "DIVIDEN",
    "ENERGI", "BANK", "UTAMA", "JAYA", "RAYA", "MAKMUR", "SEJAHTERA",
}

_TICKER_AMBIGUOUS_WORDS = {"A", "I", "TO", "IN", "ON", "AT", "AS", "IF", "OR", "IT", "AM", "PM", "IPO", "ETF"}

def _news_norm_text(text: str) -> str:
    import re
    t = str(text or "").upper()
    t = re.sub(r"[^A-Z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()

def normalize_company_name(name: str) -> list[str]:
    """Return meaningful uppercase company alias tokens and phrases."""
    norm = _news_norm_text(name)
    if not norm:
        return []
    tokens = [t for t in norm.split() if t and t not in _COMPANY_STOPWORDS and len(t) >= 3]
    aliases = []
    if tokens:
        aliases.append(" ".join(tokens))
        if len(tokens) >= 2:
            aliases.append(" ".join(tokens[:2]))
            aliases.append(" ".join(tokens[-2:]))
        aliases.extend(tokens)
    clean = []
    seen = set()
    for a in aliases:
        a = _news_norm_text(a)
        if not a or a in seen:
            continue
        if len(a) < 3 or a in _COMPANY_STOPWORDS:
            continue
        seen.add(a)
        clean.append(a)
    return clean

def extract_idx_tickers(text: str, all_tickers: set[str]) -> set[str]:
    """Return standalone IDX ticker symbols found in title/body."""
    import re
    norm = _news_norm_text(text)
    if not norm or not all_tickers:
        return set()
    found = set()
    for tk in all_tickers:
        tk2 = str(tk or "").upper().replace(".JK", "").strip()
        if not tk2 or tk2 in _TICKER_AMBIGUOUS_WORDS:
            continue
        if re.search(rf"(?<![A-Z0-9]){re.escape(tk2)}(?![A-Z0-9])", norm):
            found.add(tk2)
    return found

def build_company_alias_map(universe_rows: list[dict]) -> dict[str, list[str]]:
    """Map ticker -> aliases. Includes ticker, normalized company phrases, and curated aliases."""
    alias_map = {}
    for row in universe_rows or []:
        tk = str(row.get("ticker", row.get("Ticker", "")) or "").upper().replace(".JK", "").strip()
        if not tk:
            continue
        company = row.get("emiten") or row.get("name") or row.get("yf_company_name") or row.get("Company") or ""
        aliases = [tk]
        aliases.extend(normalize_company_name(company))
        aliases.extend(_COMPANY_ALIAS_OVERRIDES.get(tk, []))
        clean, seen = [], set()
        for a in aliases:
            a = _news_norm_text(a)
            if not a or a in seen:
                continue
            if a != tk and " " not in a and len(a) < 4:
                continue
            seen.add(a)
            clean.append(a)
        alias_map[tk] = clean
    for tk, aliases in _COMPANY_ALIAS_OVERRIDES.items():
        alias_map.setdefault(tk, [])
        for a in [tk] + aliases:
            a = _news_norm_text(a)
            if a and a not in alias_map[tk]:
                alias_map[tk].append(a)
    return alias_map

def _alias_in_text(alias: str, text_norm: str) -> bool:
    import re
    a = _news_norm_text(alias)
    if not a:
        return False
    return re.search(rf"(?<![A-Z0-9]){re.escape(a)}(?![A-Z0-9])", text_norm) is not None

def validate_news_entity_match(
    title: str,
    ticker: str,
    company_name: str = "",
    all_tickers: set[str] | None = None,
    company_alias_map: dict[str, list[str]] | None = None,
) -> dict:
    """
    Strict row-entity validation for news assignment.
    Accepts only explicit row ticker or strong row-company alias evidence.
    """
    tk = str(ticker or "").upper().replace(".JK", "").strip()
    text_norm = _news_norm_text(title)
    all_tickers = all_tickers if all_tickers is not None else NEWS_ALL_TICKERS
    company_alias_map = company_alias_map if company_alias_map is not None else NEWS_ENTITY_ALIAS_MAP

    matched_tickers = sorted(extract_idx_tickers(text_norm, all_tickers or {tk}))
    row_aliases = list(company_alias_map.get(tk, []))
    if company_name:
        row_aliases.extend(normalize_company_name(company_name))
    row_aliases.append(tk)

    matched_aliases = []
    for a in row_aliases:
        a_norm = _news_norm_text(a)
        if not a_norm or a_norm == tk:
            continue
        if " " not in a_norm and len(a_norm) < 4:
            continue
        if a_norm in _COMPANY_STOPWORDS:
            continue
        if _alias_in_text(a_norm, text_norm):
            matched_aliases.append(a_norm)

    other_alias_hits = []
    for other_tk, aliases in (company_alias_map or {}).items():
        other_tk = str(other_tk).upper().replace(".JK", "").strip()
        if not other_tk or other_tk == tk:
            continue
        for a in aliases:
            a_norm = _news_norm_text(a)
            if not a_norm or a_norm == other_tk:
                continue
            if " " not in a_norm and (len(a_norm) < 4 or a_norm in _COMPANY_STOPWORDS):
                continue
            if _alias_in_text(a_norm, text_norm):
                other_alias_hits.append(f"{other_tk}:{a_norm}")

    if tk in matched_tickers:
        status = "MULTI_TICKER_MATCHED" if len(matched_tickers) > 1 else "MATCHED"
        return {"is_match": True, "match_score": 1.0 if status == "MATCHED" else 0.90,
                "matched_tickers": matched_tickers, "matched_aliases": sorted(set(matched_aliases)),
                "entity_match_status": status, "rejection_reason": ""}

    if matched_aliases and not matched_tickers:
        return {"is_match": True, "match_score": 0.80, "matched_tickers": matched_tickers,
                "matched_aliases": sorted(set(matched_aliases)), "entity_match_status": "MATCHED",
                "rejection_reason": ""}

    if matched_tickers and tk not in matched_tickers:
        reason = "TITLE_HAS_OTHER_TICKER_ONLY"; status = "REJECTED_OTHER_TICKER"
    elif other_alias_hits and not matched_aliases:
        reason = "REJECTED_OTHER_COMPANY"; status = "REJECTED_OTHER_TICKER"
    else:
        reason = "NO_COMPANY_ALIAS_MATCH"; status = "REJECTED_NO_ENTITY"

    return {"is_match": False, "match_score": 0.0, "matched_tickers": matched_tickers,
            "matched_aliases": sorted(set(matched_aliases)), "other_alias_hits": sorted(set(other_alias_hits))[:10],
            "entity_match_status": status, "rejection_reason": reason}

def _configure_news_entity_context(rows: list[dict]) -> None:
    """Build global ticker + alias context used by news fetchers."""
    global NEWS_ENTITY_ALIAS_MAP, NEWS_ALL_TICKERS
    NEWS_ENTITY_ALIAS_MAP = build_company_alias_map(rows or [])
    NEWS_ALL_TICKERS = set(NEWS_ENTITY_ALIAS_MAP.keys())

def _record_news_rejection(row_ticker, row_company, item, validation, source):
    try:
        NEWS_REJECTION_AUDIT.append({
            "Row Ticker": str(row_ticker or "").upper().replace(".JK", ""),
            "Row Company": row_company or "",
            "Rejected Title": (item or {}).get("title", ""),
            "Detected Tickers": ", ".join(validation.get("matched_tickers", [])),
            "Detected Company Aliases": ", ".join(validation.get("other_alias_hits", validation.get("matched_aliases", []))),
            "Source": source,
            "URL": (item or {}).get("url", ""),
            "Reason": validation.get("rejection_reason", ""),
        })
    except Exception:
        pass

def _run_news_entity_validation_self_tests():
    all_tickers = {"BBRI", "BBCA", "BMRI", "TOBA"}
    amap = build_company_alias_map([
        {"ticker": "BBRI", "emiten": "BANK RAKYAT INDONESIA"},
        {"ticker": "BBCA", "emiten": "BANK CENTRAL ASIA"},
        {"ticker": "BMRI", "emiten": "BANK MANDIRI"},
        {"ticker": "TOBA", "emiten": "TBS ENERGI UTAMA"},
    ])
    bad = validate_news_entity_match("PT TBS Energi Utama Tbk Gelar RUPS", "BBRI", "BANK RAKYAT INDONESIA", all_tickers, amap)
    good = validate_news_entity_match("PT TBS Energi Utama Tbk Gelar RUPS", "TOBA", "TBS ENERGI UTAMA", all_tickers, amap)
    multi = validate_news_entity_match("BBRI BBCA dan BMRI kompak menguat jelang rilis kinerja", "BBRI", "BANK RAKYAT INDONESIA", all_tickers, amap)
    generic = validate_news_entity_match("Jurus Ampuh Cara Kaya dari Saham Jelang RUPS", "BBRI", "BANK RAKYAT INDONESIA", all_tickers, amap)
    assert not bad["is_match"], "BBRI must reject TBS/TOBA corp-action title"
    assert good["is_match"], "TOBA must accept TBS Energi title"
    assert multi["is_match"] and multi["entity_match_status"] == "MULTI_TICKER_MATCHED", "Multi-ticker title must be flagged"
    assert not generic["is_match"], "Generic RUPS title must not be assigned"

def _rss_fetch_items(query_str: str, max_items: int = 15) -> list:
    """
    Shared RSS item fetcher. Returns list of {date, date_str, title, url} dicts.
    Handles network timeout gracefully — returns [] on any failure.
    """
    from email.utils import parsedate_to_datetime
    import datetime as _dti
    try:
        url = ("https://news.google.com/rss/search?"
               f"q={_urlquote(query_str)}&hl=id&gl=ID&ceid=ID:id")
        req = _urllib_req.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with _urllib_req.urlopen(req, timeout=8) as resp:
            raw = resp.read()
        root = _ET.fromstring(raw)
        out = []
        for item in root.findall(".//item")[:max_items]:
            t_el = item.find("title"); l_el = item.find("link"); d_el = item.find("pubDate")
            title = (t_el.text or "").strip() if t_el is not None else ""
            link  = (l_el.text or "").strip() if l_el is not None else ""
            pub   = (d_el.text or "").strip() if d_el is not None else ""
            try:
                pub_dt = parsedate_to_datetime(pub).date()
            except Exception:
                pub_dt = None
            out.append({
                "date":     pub_dt,
                "date_str": pub_dt.strftime("%Y-%m-%d") if pub_dt else "-",
                "title":    title,
                "url":      link,
            })
        return out
    except Exception:
        return []


def _pick_best_item(items: list, market_date: str, window_days: int = 7):
    """
    From a list of RSS items, pick the best match:
      PRIORITY:  most recent item dated within 1 day of market_date (today's news first)
      PRIMARY:   most recent item with date <= market_date within window_days (7d default)
      FALLBACK:  globally most recent item with a date (only if within window)
    Returns (item_dict | None, source_label: str)
    """
    import datetime as _dti
    try:
        mdate = datetime.strptime(market_date, "%Y-%m-%d").date()
    except Exception:
        mdate = None

    items_with_date = [it for it in items if it["date"] is not None]
    items_with_date.sort(key=lambda x: x["date"], reverse=True)

    if mdate:
        window_start = mdate - _dti.timedelta(days=window_days)
        # Priority pass: items from today or yesterday (≤ 1 day old)
        priority_start = mdate - _dti.timedelta(days=1)
        priority = [it for it in items_with_date if priority_start <= it["date"] <= mdate]
        if priority:
            return priority[0], "On-Date"
        # Standard window pass
        primary = [it for it in items_with_date if window_start <= it["date"] <= mdate]
        if primary:
            return primary[0], "On-Date"

    # No items within window — return nothing rather than showing stale news
    return None, "N/A"


def _days_since_label(item_date, market_date: str) -> str:
    """Returns compact age string: '2d', '3w', '-'."""
    import datetime as _dti
    try:
        mdate = datetime.strptime(market_date, "%Y-%m-%d").date()
        d = (mdate - item_date).days
        if d < 0:
            return f"+{abs(d)}d"    # future-dated (shouldn't happen in practice)
        if d == 0:
            return "today"
        if d < 7:
            return f"{d}d"
        if d < 30:
            return f"{d // 7}w"
        return f"{d // 30}mo"
    except Exception:
        return "-"


def _fetch_news_sentiment_screener(ticker: str, market_date: str, company_name: str = "") -> dict:
    """
    General market sentiment for the screener inline cell.

    Queries: '{TICKER} saham'  +  '{TICKER} laporan keuangan kinerja'
    Returns dict:
        sentiment   : 'Bullish' | 'Bearish' | 'Neutral'
        headline    : truncated title (max 80 chars)
        age         : compact age string e.g. '3d', '2w'
        display     : formatted cell string
        color       : hex color for the cell font
        source      : 'On-Date' | 'IFNA' | 'N/A'
    """
    EMPTY = {
        "sentiment": "", "headline": "",
        "age": "", "display": "",
        "color": "9CA3AF", "source": "N/A", "url": "",
        "entity_match_status": "NO_NEWS", "matched_tickers": "", "matched_aliases": "",
        "rejection_reason": "",
    }
    sym = ticker.upper().strip().replace(".JK", "")
    # BUG-09: Use company name (not ticker code) as primary query term.
    # Ticker-based queries (e.g., "BBRI saham") pull unrelated headlines.
    # Company-name queries (e.g., "Bank Rakyat Indonesia saham") are far more precise.
    _cname = str(company_name or sym).strip()
    # Fallback to short ticker if company_name is missing/very short
    _qterm = _cname if len(_cname) >= 4 else sym

    items: list = []
    for q in [f"{_qterm} saham", f"{_qterm} kinerja laporan keuangan"]:
        items.extend(_rss_fetch_items(q, max_items=12))
        if len(items) >= 8:
            break

    # Deduplicate
    seen: set = set()
    deduped = []
    for it in items:
        k = it.get("url") or it.get("title")
        if k and k not in seen:
            seen.add(k); deduped.append(it)

    chosen, source = _pick_best_item(deduped, market_date)
    if chosen is None:
        return EMPTY

    title = chosen["title"]
    validation = validate_news_entity_match(title, sym, company_name)
    if not validation.get("is_match"):
        _record_news_rejection(sym, company_name, chosen, validation, f"Sentiment-{source}")
        out = dict(EMPTY)
        out.update({
            "source": "Rejected-Mismatch",
            "entity_match_status": validation.get("entity_match_status", "REJECTED_NO_ENTITY"),
            "matched_tickers": ", ".join(validation.get("matched_tickers", [])),
            "matched_aliases": ", ".join(validation.get("matched_aliases", [])),
            "rejection_reason": validation.get("rejection_reason", "NO_COMPANY_ALIAS_MATCH"),
        })
        return out

    cat, sent, _ = _classify_disclosure_rss(title)   # reuse existing classifier
    age = _days_since_label(chosen["date"], market_date) if chosen["date"] else "-"

    SENT_ICON  = {"Bullish": "↑", "Bearish": "↓", "Neutral": "→"}
    SENT_COLOR = {"Bullish": "1A6B3C", "Bearish": "B91C1C", "Neutral": "6B7280"}

    icon    = SENT_ICON.get(sent, "→")
    color   = SENT_COLOR.get(sent, "6B7280")
    snippet = title[:75] + ("…" if len(title) > 75 else "")
    display = f"{icon} {sent}  ·  {snippet}  [{age}]"

    return {
        "sentiment": sent,
        "headline":  snippet,
        "age":       age,
        "display":   display,
        "color":     color,
        "source":    source,
        "url":       chosen.get("url", ""),
        "entity_match_status": validation.get("entity_match_status", "MATCHED"),
        "matched_tickers": ", ".join(validation.get("matched_tickers", [])),
        "matched_aliases": ", ".join(validation.get("matched_aliases", [])),
        "rejection_reason": "",
    }


def _fetch_corp_action_screener(ticker: str, market_date: str, company_name: str = "") -> dict:
    """
    Corporate-action-specific news for the screener inline cell.

    Strategy:
      1. Query Google RSS with focused corp-action terms for this ticker.
         Filter results: only keep items whose title contains at least one
         corp-action keyword from _CORP_ACTION_KEYWORDS.
      2. If RSS yields no corp-action match → fallback to IDX keterbukaan
         informasi scrape (reuses _fetch_idx_disclosure).
      3. IFNA: if neither source returns anything → "-".

    Returns dict:
        action_type : classified label (same as disc_category)
        headline    : truncated title
        age         : compact age string
        display     : formatted cell string
        color       : hex color
        source      : 'RSS-OnDate' | 'RSS-IFNA' | 'IDX-Scrape' | 'N/A'
    """
    EMPTY = {
        "action_type": "", "headline": "",
        "age": "", "display": "",
        "color": "9CA3AF", "source": "N/A", "url": "",
        "event_risk": "", "entity_match_status": "NO_NEWS", "matched_tickers": "",
        "matched_aliases": "", "rejection_reason": "",
    }
    sym = ticker.upper().strip().replace(".JK", "")

    def _has_corp_kw(title: str) -> bool:
        t = title.lower()
        return any(kw in t for kw in _CORP_ACTION_KEYWORDS)

    # Item 12: Use company name for corp action queries + limit to 30-day window
    _ca_cname = str(company_name or sym).strip()
    _ca_qterm = _ca_cname if len(_ca_cname) >= 4 else sym

    # RSS queries — two focused queries (company name primary)
    items: list = []
    for q in [
        f"{_ca_qterm} dividen buyback rights issue RUPS direksi",
        f"{_ca_qterm} akuisisi merger obligasi waran private placement",
    ]:
        items.extend(_rss_fetch_items(q, max_items=15))
        if len(items) >= 12:
            break

    # Filter to within 30 calendar days of market_date (≈ 1 month window for corp actions)
    try:
        _mdt = pd.Timestamp(market_date)
        _cutoff = _mdt - pd.Timedelta(days=30)
        items = [
            it for it in items
            if it.get("date") and (_cutoff <= pd.Timestamp(it["date"]) <= _mdt)
        ]
    except Exception:
        pass

    # Deduplicate and filter to corp-action only
    seen: set = set()
    ca_items = []
    for it in items:
        k = it.get("url") or it.get("title")
        if k and k not in seen:
            seen.add(k)
            if _has_corp_kw(it.get("title", "")):
                ca_items.append(it)

    valid_ca_items = []
    for it in ca_items:
        validation = validate_news_entity_match(it.get("title", ""), sym, company_name)
        if validation.get("is_match"):
            it = dict(it)
            it["_entity_validation"] = validation
            valid_ca_items.append(it)
        else:
            _record_news_rejection(sym, company_name, it, validation, "RSS-CorpAction")

    chosen, raw_source = _pick_best_item(valid_ca_items, market_date)

    # IDX keterbukaan fallback
    if chosen is None:
        try:
            disc = _fetch_idx_disclosure(ticker)
            disc_title = disc.get("disclosure_title", "") or ""
            disc_date  = disc.get("disclosure_date", "-")
            disc_url   = disc.get("disclosure_url", "")
            if disc_title and disc_title not in ("N/A", "-", ""):
                disc_item = {"title": disc_title, "url": disc_url, "date_str": disc_date}
                validation = validate_news_entity_match(disc_title, sym, company_name)
                if not validation.get("is_match"):
                    _record_news_rejection(sym, company_name, disc_item, validation, "IDX-Scrape")
                    out = dict(EMPTY)
                    out.update({
                        "source": "Rejected-Mismatch",
                        "entity_match_status": validation.get("entity_match_status", "REJECTED_NO_ENTITY"),
                        "matched_tickers": ", ".join(validation.get("matched_tickers", [])),
                        "matched_aliases": ", ".join(validation.get("matched_aliases", [])),
                        "rejection_reason": validation.get("rejection_reason", "NO_COMPANY_ALIAS_MATCH"),
                    })
                    return out
                cat, sent, risk = _classify_disclosure_rss(disc_title)
                import datetime as _dti
                try:
                    disc_dt = datetime.strptime(disc_date, "%Y-%m-%d").date()
                    age = _days_since_label(disc_dt, market_date)
                except Exception:
                    age = ""
                snippet = disc_title[:75] + ("…" if len(disc_title) > 75 else "")
                RISK_COLOR = {"High": "B91C1C", "Medium": "D97706", "Low": "1A6B3C"}
                return {
                    "action_type": cat,
                    "headline":    snippet,
                    "age":         age,
                    "display":     f"[{cat}]  {snippet}  [{age}]",
                    "color":       RISK_COLOR.get(risk, "6B7280"),
                    "source":      "IDX-Scrape",
                    "url":         disc_url,
                    "event_risk":   risk,
                    "entity_match_status": validation.get("entity_match_status", "MATCHED"),
                    "matched_tickers": ", ".join(validation.get("matched_tickers", [])),
                    "matched_aliases": ", ".join(validation.get("matched_aliases", [])),
                    "rejection_reason": "",
                }
        except Exception:
            pass
        return EMPTY

    title = chosen["title"]
    validation = chosen.get("_entity_validation") or validate_news_entity_match(title, sym, company_name)
    if not validation.get("is_match"):
        _record_news_rejection(sym, company_name, chosen, validation, f"RSS-{raw_source}")
        out = dict(EMPTY)
        out.update({
            "source": "Rejected-Mismatch",
            "entity_match_status": validation.get("entity_match_status", "REJECTED_NO_ENTITY"),
            "matched_tickers": ", ".join(validation.get("matched_tickers", [])),
            "matched_aliases": ", ".join(validation.get("matched_aliases", [])),
            "rejection_reason": validation.get("rejection_reason", "NO_COMPANY_ALIAS_MATCH"),
        })
        return out

    cat, sent, risk = _classify_disclosure_rss(title)
    age     = _days_since_label(chosen["date"], market_date) if chosen["date"] else ""
    snippet = title[:75] + ("…" if len(title) > 75 else "")
    source  = f"RSS-{raw_source}"

    RISK_COLOR = {"High": "B91C1C", "Medium": "D97706", "Low": "1A6B3C"}
    color = RISK_COLOR.get(risk, "6B7280")
    display = f"[{cat}]  {snippet}  [{age}]"

    return {
        "action_type": cat,
        "headline":    snippet,
        "age":         age,
        "display":     display,
        "color":       color,
        "source":      source,
        "url":         chosen.get("url", ""),
        "event_risk":   risk,
        "entity_match_status": validation.get("entity_match_status", "MATCHED"),
        "matched_tickers": ", ".join(validation.get("matched_tickers", [])),
        "matched_aliases": ", ".join(validation.get("matched_aliases", [])),
        "rejection_reason": "",
    }



def build_idx_disclosure_sheet(wb, latest_market_day: str, rows: list):
    """
    IDX News — rebuilt from the same NEWS INTELLIGENCE logic used by IDX Screener:
    Sentiment News + Corp. Action are fetched through _fetch_news_sentiment_screener()
    and _fetch_corp_action_screener(), so IDX Screener and IDX News stay consistent.
    """
    SHEET_NAME = "IDX News"
    if SHEET_NAME in wb.sheetnames:
        wb.remove(wb[SHEET_NAME])
    ws = wb.create_sheet(SHEET_NAME)
    ws.sheet_view.showGridLines = False

    FILL_TITLE = PatternFill("solid", fgColor="0D1B2A")
    FILL_GROUP = PatternFill("solid", fgColor="6B2737")
    FILL_HDR   = PatternFill("solid", fgColor="1C3253")
    FILL_BODY  = PatternFill("solid", fgColor="F8FAFC")
    FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    FONT_GROUP = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_HDR   = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_BODY  = Font(name="Calibri", size=10, color="1F2937")
    FONT_LINK  = Font(name="Calibri", size=10, color="0563C1", underline="single")
    AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
    THIN = Side(style="thin", color="C2CADE")
    BDR = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    # Item 41: Streamlined to 7 essential columns (audit cols removed)
    COLS = [
        ("ticker",            "Ticker",          11, "center"),
        ("emiten",            "Company",          34, "left"),
        ("idx_sector",        "Sector",           22, "left"),
        ("sentiment_display", "Sentiment News",   58, "left"),
        ("sentiment_label",   "Sentiment",        14, "center"),
        ("sentiment_url",     "News URL",         14, "center"),
        ("corp_display",      "Corp. Action",     54, "left"),
        ("corp_category",     "Corp. Category",   22, "center"),
        ("corp_risk",         "Event Risk",       14, "center"),
        ("corp_url",          "Corp. URL",        14, "center"),
    ]
    ws.column_dimensions["A"].width = 0.5
    for ci, (_, _, width, _) in enumerate(COLS, start=2):
        ws.column_dimensions[get_column_letter(ci)].width = width

    last_col = get_column_letter(1 + len(COLS))
    ws.merge_cells(f"B1:{last_col}1")
    _mode_lbl_news = "[BACKTEST MODE]" if BACKTEST_MODE else "Live"
    ws["B1"] = f"  IDX News  ·  {_mode_lbl_news}  ·  {latest_market_day}"
    ws["B1"].fill = FILL_TITLE; ws["B1"].font = FONT_TITLE; ws["B1"].alignment = AL_L
    ws.merge_cells(f"B2:{last_col}2")
    ws["B2"] = "  Queries use company name · Entity validation applied · Sentiment: Bullish / Bearish / Neutral"
    ws["B2"].fill = FILL_GROUP; ws["B2"].font = FONT_GROUP; ws["B2"].alignment = AL_L

    for ci, (_, hdr, _, _) in enumerate(COLS, start=2):
        c = ws.cell(4, ci, hdr)
        c.fill = FILL_HDR; c.font = FONT_HDR; c.alignment = AL_C; c.border = BDR

    ws.freeze_panes = "E5"
    ok_rows = [r for r in rows if r.get("data_status") == "OK"]
    SENT_CLR = {"Bullish": "008000", "Bearish": "C00000", "Neutral": "595959"}
    RISK_CLR = {"High": "C00000", "Medium": "E36C09", "Low": "008000"}

    for ri, row in enumerate(ok_rows, start=5):
        ticker = str(row.get("ticker", "")).upper().strip()
        try:
            sent = _fetch_news_sentiment_screener(ticker, latest_market_day, row.get("emiten", row.get("name", "")))
        except Exception:
            sent = {"display": "", "sentiment": "", "url": "", "color": "595959", "entity_match_status": "NO_NEWS", "matched_tickers": "", "rejection_reason": ""}
        try:
            corp = _fetch_corp_action_screener(ticker, latest_market_day, row.get("emiten", row.get("name", "")))
        except Exception:
            corp = {"display": "", "action_type": "", "event_risk": "", "source": "N/A", "url": "", "color": "595959", "entity_match_status": "NO_NEWS", "matched_tickers": "", "rejection_reason": ""}

        out = {
            "ticker": ticker,
            "emiten": row.get("emiten", row.get("name", "-")),
            "idx_sector": row.get("idx_sector", row.get("sector", "-")),
            "sentiment_display": sent.get("display", ""),
            "sentiment_label": sent.get("sentiment", sent.get("label", "")),
            "sentiment_method": "Keyword classifier + entity validation" if sent.get("display") else "",
            "sentiment_entity_status": sent.get("entity_match_status", "NO_NEWS"),
            "sentiment_matched_tickers": sent.get("matched_tickers", ""),
            "sentiment_url_raw": sent.get("url", ""),
            "sentiment_url": sent.get("url", ""),
            "corp_display": corp.get("display", ""),
            "corp_category": corp.get("action_type", corp.get("category", "")),
            "corp_risk": corp.get("event_risk", corp.get("risk", "")),
            "corp_source": corp.get("source", "N/A"),
            "corp_entity_status": corp.get("entity_match_status", "NO_NEWS"),
            "corp_matched_tickers": corp.get("matched_tickers", ""),
            "corp_rejection_reason": corp.get("rejection_reason", ""),
            "corp_url_raw": corp.get("url", ""),
            "corp_url": corp.get("url", ""),
        }

        for ci, (key, _, _, align) in enumerate(COLS, start=2):
            val = out.get(key, "-")
            c = ws.cell(ri, ci, "" if val in (None, "", "N/A") else val)
            c.fill = FILL_BODY; c.border = BDR
            c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=key.endswith("display"))
            c.font = FONT_BODY
            if key == "ticker":
                c.font = Font(name="Calibri", size=10, bold=True, color="1A3D4F")
            if key in ("sentiment_url", "corp_url"):
                if isinstance(val, str) and val.startswith("http"):
                    c.hyperlink = val; c.value = "View →"; c.font = FONT_LINK
                else:
                    c.value = "-"
            if key == "sentiment_label":
                c.font = Font(name="Calibri", size=10, bold=True, color=SENT_CLR.get(str(c.value), "595959"))
            if key == "corp_risk":
                c.font = Font(name="Calibri", size=10, bold=True, color=RISK_CLR.get(str(c.value), "595959"))
    ws.auto_filter.ref = f"B4:{last_col}{max(4, ws.max_row)}"
    print(f"[INFO] IDX News: {len(ok_rows)} rows written from NEWS INTELLIGENCE logic.")

def _compute_pbv_bands(ticker: str, years: int = 3) -> dict:
    """
    Stockbit-style PBV Band Analysis.
    - Uses adjusted closing price (auto_adjust=True) to handle splits correctly.
    - BVPS = Total Stockholder Equity / Shares Outstanding (quarterly).
    - BVPS is forward-filled to daily frequency; NO look-ahead bias
      (BVPS for quarter Q is only applied from the earnings release date onward).
    - Excludes PBV <= 0, inf, NaN.
    - Returns z-score, percentile, SD bands, and regime.
    """
    empty = {
        "pbv_curr": np.nan, "pbv_mean": np.nan,
        "pbv_plus1": np.nan, "pbv_plus2": np.nan,
        "pbv_minus1": np.nan, "pbv_minus2": np.nan,
        "pbv_zscore": np.nan, "pbv_pct": np.nan,
        "pbv_regime": "N/A",
    }
    try:
        sym = ticker.upper().strip()
        if not sym.endswith(".JK"):
            sym += ".JK"
        tk = yf.Ticker(sym)

        # ── 1. Adjusted price history ──────────────────────────────────────────
        period_str = f"{years + 1}y"
        _mdate_pbv = get_market_date()
        _end_pbv   = (_mdate_pbv + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        price_hist = tk.history(period=period_str, end=_end_pbv, auto_adjust=True)
        if isinstance(price_hist.columns, pd.MultiIndex):
            price_hist.columns = price_hist.columns.get_level_values(0)
        price_hist = price_hist[price_hist.index.normalize() <= _mdate_pbv]
        if price_hist is None or price_hist.empty:
            return empty
        if hasattr(price_hist.index, "tz") and price_hist.index.tz is not None:
            price_hist.index = price_hist.index.tz_localize(None)
        price_hist = price_hist["Close"].squeeze()
        if hasattr(price_hist, "columns"):
            price_hist = price_hist.iloc[:, 0]
        price_hist = price_hist.dropna()
        price_hist.index = pd.to_datetime(price_hist.index, errors="coerce", utc=False, format="mixed").normalize()

        # ── Stockbit-style PBV band approximation ──────────────────────────────
        # Stockbit's PBV Band visually behaves like daily PBV = price / latest BVPS,
        # then mean/SD are computed over the selected rolling window.  yfinance does
        # not expose Stockbit's full historical BVPS feed, so use the latest P/B
        # ratio to infer current BVPS, then scale the historical price series.
        try:
            info0 = tk.info or {}
            curr_pbv0 = safe_num(info0.get("priceToBook"), np.nan)
            curr_px0 = safe_num(info0.get("currentPrice"), np.nan)
            if pd.isna(curr_px0) or curr_px0 <= 0:
                curr_px0 = float(price_hist.iloc[-1])
            # Guard: PBV > 50 almost always indicates yfinance currency/unit mismatch
            # (IDX prices in IDR inflate the ratio). Fall through to balance-sheet path.
            if pd.notna(curr_pbv0) and 0 < curr_pbv0 <= 50 and curr_px0 > 0:
                bvps0 = curr_px0 / curr_pbv0
                pbv_series0 = (price_hist / bvps0).replace([np.inf, -np.inf], np.nan).dropna()
                pbv_series0 = pbv_series0[pbv_series0 > 0]
                if not pbv_series0.empty and pbv_series0.mean() <= 50:
                    cutoff0 = pbv_series0.index.max() - pd.DateOffset(years=years)
                    pbv_window0 = pbv_series0[pbv_series0.index >= cutoff0]
                    if len(pbv_window0) >= 30:
                        curr = float(pbv_window0.iloc[-1])
                        mean = float(pbv_window0.mean())
                        sd = float(pbv_window0.std(ddof=0))
                        if sd > 0 and curr <= 50:
                            z = (curr - mean) / sd
                            pct = float(np.sum(pbv_window0.values <= curr) / len(pbv_window0) * 100.0)
                            regime = ("Deep Undervalued" if z <= -2 else
                                      "Undervalued" if z <= -1 else
                                      "Bubble" if z >= 2 else
                                      "Overvalued" if z >= 1 else "Fair Value")
                            return {
                                "pbv_curr": round(curr, 4), "pbv_mean": round(mean, 4),
                                "pbv_plus1": round(mean + sd, 4), "pbv_plus2": round(mean + 2*sd, 4),
                                "pbv_minus1": round(mean - sd, 4), "pbv_minus2": round(mean - 2*sd, 4),
                                "pbv_zscore": round(z, 4), "pbv_pct": round(pct, 2),
                                "pbv_regime": regime,
                            }
        except Exception:
            pass

        # ── 2. Quarterly balance sheet — equity + shares ───────────────────────
        bs = tk.quarterly_balance_sheet
        if bs is None or bs.empty:
            # Fallback: Stockbit-like rolling PBV using latest bookValue from yfinance info.
            # Less precise than quarter-by-quarter BVPS, but keeps the band populated when
            # Yahoo Finance does not expose IDX quarterly balance-sheet rows.
            try:
                info = tk.info or {}
                bvps = float(info.get("bookValue") or 0)
                if bvps <= 0:
                    return empty
                pbv_series = (price_hist / bvps).dropna()
                pbv_series = pbv_series[pbv_series > 0]
                cutoff = pbv_series.index.max() - pd.DateOffset(years=years)
                pbv_roll = pbv_series[pbv_series.index >= cutoff]
                if len(pbv_roll) < 30:
                    return empty
                curr = float(pbv_roll.iloc[-1])
                mean = float(pbv_roll.mean())
                sd = float(pbv_roll.std(ddof=0))
                z = (curr - mean) / sd if sd > 0 else np.nan
                pct = float((pbv_roll.rank(pct=True).iloc[-1]) * 100.0)
                regime = ("Deep Undervalued" if pd.notna(z) and z <= -2 else
                          "Undervalued" if pd.notna(z) and z <= -1 else
                          "Bubble" if pd.notna(z) and z >= 2 else
                          "Overvalued" if pd.notna(z) and z >= 1 else "Fair Value")
                return {
                    "pbv_curr": curr, "pbv_mean": mean,
                    "pbv_plus1": mean + sd, "pbv_plus2": mean + 2*sd,
                    "pbv_minus1": mean - sd, "pbv_minus2": mean - 2*sd,
                    "pbv_zscore": z, "pbv_pct": pct,
                    "pbv_regime": regime,
                }
            except Exception:
                return empty

        # Normalize column orientation (tickers sometimes have rows/cols swapped)
        # yfinance quarterly_balance_sheet: columns = report dates, rows = fields
        if isinstance(bs.columns[0], str):
            bs = bs.T   # flip if dates are in rows
        bs.index = pd.to_datetime(bs.index, errors="coerce", utc=False, format="mixed").normalize()
        bs = bs.sort_index()

        # Equity field — try multiple yfinance key names
        equity_keys = [
            "Stockholders Equity", "Total Stockholder Equity",
            "Common Stock Equity", "Total Equity Gross Minority Interest",
        ]
        equity_series = None
        for ek in equity_keys:
            if ek in bs.columns:
                equity_series = bs[ek].dropna()
                break
        if equity_series is None or equity_series.empty:
            return empty

        # Shares outstanding — try info first, then balance sheet
        shares_keys = ["Ordinary Shares Number", "Share Issued", "Common Stock"]
        shares_series = None
        for sk in shares_keys:
            if sk in bs.columns:
                shares_series = bs[sk].dropna()
                break
        if shares_series is None or shares_series.empty:
            # fall back to static info shares
            static_shares = float(tk.info.get("sharesOutstanding") or 0)
            if static_shares <= 0:
                return empty
            shares_series = pd.Series(static_shares, index=equity_series.index)

        # ── 3. BVPS per quarter (report date → BVPS) ──────────────────────────
        # BUG-03: equity_series from yfinance statements is in USD for .JK stocks.
        # price_hist is in IDR. Must convert equity to IDR before computing BVPS
        # so that pbv_daily = price_IDR / bvps_IDR is a dimensionless ratio.
        try:
            _fx_pbv = _fetch_usd_idr_rate()
            equity_series_idr = equity_series * _fx_pbv
        except Exception:
            equity_series_idr = equity_series  # fallback: will be filtered by <= 100 guard
        bvps_q = (equity_series_idr / shares_series).dropna()
        bvps_q = bvps_q[bvps_q > 0]   # exclude negative equity
        if bvps_q.empty:
            return empty

        # ── 4. Forward-fill BVPS to daily price index (NO look-ahead bias) ────
        # Create a daily series spanning the price history
        daily_idx = price_hist.index
        bvps_daily = pd.Series(np.nan, index=daily_idx)
        for report_date, bvps_val in bvps_q.items():
            # Apply BVPS from this report date onward (forward-fill)
            mask = daily_idx >= report_date
            bvps_daily[mask] = bvps_val

        # ── 5. Compute daily PBV ───────────────────────────────────────────────
        pbv_daily = (price_hist / bvps_daily).replace([np.inf, -np.inf], np.nan).dropna()
        pbv_daily = pbv_daily[pbv_daily > 0]
        # Sanity guard: discard extreme values (> 100) that indicate BVPS data issues
        pbv_daily = pbv_daily[pbv_daily <= 100]

        # Restrict to 5-year window
        cutoff = pbv_daily.index[-1] - pd.DateOffset(years=years)
        pbv_window = pbv_daily[pbv_daily.index >= cutoff]
        if len(pbv_window) < 20:
            return empty

        curr_pbv = float(pbv_daily.iloc[-1])
        mean_pbv = float(pbv_window.mean())
        std_pbv  = float(pbv_window.std(ddof=1))
        if std_pbv == 0:
            return empty

        plus1  = mean_pbv + std_pbv
        plus2  = mean_pbv + 2 * std_pbv
        minus1 = mean_pbv - std_pbv
        minus2 = mean_pbv - 2 * std_pbv

        z_score  = (curr_pbv - mean_pbv) / std_pbv
        # Percentile rank — pure numpy, no scipy dependency
        arr      = pbv_window.values
        pct_rank = float(np.sum(arr <= curr_pbv) / len(arr) * 100)

        # Regime (spec labels)
        if curr_pbv < minus2:
            regime = "Deep Undervalued"
        elif curr_pbv < minus1:
            regime = "Undervalued"
        elif curr_pbv <= plus1:
            regime = "Fair Value"
        elif curr_pbv <= plus2:
            regime = "Overvalued"
        else:
            regime = "Bubble"

        return {
            "pbv_curr":   round(curr_pbv, 4),
            "pbv_mean":   round(mean_pbv, 4),
            "pbv_plus1":  round(plus1,    4),
            "pbv_plus2":  round(plus2,    4),
            "pbv_minus1": round(minus1,   4),
            "pbv_minus2": round(minus2,   4),
            "pbv_zscore": round(z_score,  4),
            "pbv_pct":    round(pct_rank, 2),
            "pbv_regime": regime,
        }
    except Exception:
        return empty



def _stockbit_style_pbv_from_source_row(screener_row: dict, fallback: dict = None) -> dict:
    """
    Prefer source-provided Stockbit-style PBV band fields when available.

    Why this exists:
    - Stockbit PBV Band displays a 3-year distribution with constant horizontal
      bands for Current PBV, Mean, +1/+2 SD, -1/-2 SD.
    - yfinance cannot always reproduce Stockbit exactly because IDX historical
      BVPS/report-date data is incomplete or delayed.
    - If the Raw/KSEI input already contains Stockbit-style PBV fields, those
      fields are the best parity source and should override the yfinance
      approximation used in _compute_pbv_bands().

    Expected AADI parity example from user screenshot:
    Current=1.10, Mean=1.21, +2=1.75, +1=1.48, -1=0.94, -2=0.67.
    """
    fb = fallback.copy() if isinstance(fallback, dict) else {}
    empty = {
        "pbv_curr": np.nan, "pbv_mean": np.nan,
        "pbv_plus1": np.nan, "pbv_plus2": np.nan,
        "pbv_minus1": np.nan, "pbv_minus2": np.nan,
        "pbv_zscore": np.nan, "pbv_pct": np.nan,
        "pbv_regime": "N/A",
    }
    for k, v in empty.items():
        fb.setdefault(k, v)

    def _v(*keys):
        for key in keys:
            val = screener_row.get(key, np.nan)
            try:
                if pd.notna(val):
                    return float(val)
            except Exception:
                continue
        return np.nan

    curr  = _v("pbv_curr", "PBV Current", "Current PBV")
    mean  = _v("pbv_mean", "PBV Mean", "Mean PBV")
    plus1 = _v("pbv_p1", "pbv_plus1", "PBV +1", "+1 PBV Standard Deviation")
    plus2 = _v("pbv_p2", "pbv_plus2", "PBV +2", "+2 PBV Standard Deviation")
    minus1 = _v("pbv_m1", "pbv_minus1", "PBV -1", "-1 PBV Standard Deviation")
    minus2 = _v("pbv_m2", "pbv_minus2", "PBV -2", "-2 PBV Standard Deviation")

    vals = [curr, mean, plus1, plus2, minus1, minus2]
    if not all(pd.notna(x) and x > 0 for x in vals):
        return fb

    # Stockbit-style SD is implied by the displayed bands. Average both sides to
    # dampen minor rounding from the data source.
    sd_candidates = [
        plus1 - mean,
        (plus2 - mean) / 2.0,
        mean - minus1,
        (mean - minus2) / 2.0,
    ]
    sd_candidates = [float(x) for x in sd_candidates if pd.notna(x) and x > 0]
    sd = float(np.nanmean(sd_candidates)) if sd_candidates else np.nan

    z = (curr - mean) / sd if pd.notna(sd) and sd > 0 else np.nan

    # A true percentile requires the full historical PBV series. If unavailable
    # from Stockbit/Raw source, estimate percentile from the normal CDF implied by
    # the z-score. This is more stable than leaving it blank and keeps the regime
    # aligned with the plotted bands.
    try:
        import math
        pct = 0.5 * (1.0 + math.erf(float(z) / math.sqrt(2.0))) * 100.0 if pd.notna(z) else np.nan
    except Exception:
        pct = np.nan

    regime = (
        "Deep Undervalued" if pd.notna(z) and z <= -2 else
        "Undervalued"      if pd.notna(z) and z <= -1 else
        "Bubble"           if pd.notna(z) and z >= 2 else
        "Overvalued"       if pd.notna(z) and z >= 1 else
        "Fair Value"
    )

    return {
        "pbv_curr": round(curr, 4),
        "pbv_mean": round(mean, 4),
        "pbv_plus1": round(plus1, 4),
        "pbv_plus2": round(plus2, 4),
        "pbv_minus1": round(minus1, 4),
        "pbv_minus2": round(minus2, 4),
        "pbv_zscore": round(z, 4) if pd.notna(z) else np.nan,
        "pbv_pct": round(pct, 2) if pd.notna(pct) else np.nan,
        "pbv_regime": regime,
    }


# ── FUNDAMENTAL DETAIL SCHEMA ─────────────────────────────────────────────────
# Groups and columns for the IDX Fundamental Detail sheet.
# Format per column: (key, header, col_width, number_format, alignment)
# number_format: "compact" → compact_fmt(), "0.00%" → percent, "#,##0" → integer, None → raw text
FUNDAMENTAL_DETAIL_SCHEMA = [
    ("Stock Info", FILL_GROUP_OWNER, [
        ("ticker",          "Ticker",         10, None,    "left"),
        ("close",           "Price",          10, "#,##0", "center"),
        ("pct_change",      "Price Change %", 12, "0.00%", "center"),
        ("idx_sector",      "IDX Sector",     20, None,    "left"),
    ]),
    ("Company Profile", FILL_GROUP_OWNER, [
        ("yf_business_summary",   "Business Summary",   60, None,      "left"),
        ("yf_shares_outstanding", "Shares Outstanding", 18, "compact", "center"),
        ("yf_float_shares",       "Free Float",         16, "compact", "center"),
        ("yf_free_float_pct",     "Free Float (%)",     14, "0.00%",  "center"),
        ("yf_ipo_date",           "IPO Date",           14, None,      "center"),
    ]),
    ("Current Valuation", FILL_GROUP_VAL, [
        ("yf_pe_annualised",              "Current PE Ratio (Annualised)", 24, "0.00",   "center"),
        ("yf_pe_ttm_current",             "Current PE Ratio (TTM)",        22, "0.00",   "center"),
        ("yf_earnings_yield_ttm",         "Earnings Yield (TTM)",          20, "0.00%",  "center"),
        ("yf_ps_ttm_current",             "Current Price to Sales (TTM)",  24, "0.00",   "center"),
        ("yf_pbv_current",                "Current Price to Book Value",   24, "0.00",   "center"),
        ("yf_ev_ebit_ttm",                "EV to EBIT (TTM)",              18, "0.00",   "center"),
        ("yf_ev_ebitda_ttm",              "EV to EBITDA (TTM)",            20, "0.00",   "center"),
        ("yf_market_cap_current",         "Market Cap",                    16, "compact","center"),
        ("yf_enterprise_value_current",   "Enterprise Value",              18, "compact","center"),
        ("yf_shares_outstanding_current", "Current Share Outstanding",     24, "compact","center"),
        # These two only ever lived in the Upcoming/Latest Dividend sub-tables
        # below -- the Fundamental sheet itself had no standalone dividend
        # yield/payout column, so a lookup against this sheet (e.g. the
        # website's summary card) always came back empty despite the
        # underlying yfinance data (yf_dividend_yield/yf_payout_ratio) being
        # fetched and available the whole time.
        ("yf_dividend_yield",             "Dividend Yield (%)",            18, "0.00%",  "center"),
        ("yf_payout_ratio",               "Payout Ratio (%)",              18, "0.00%",  "center"),
    ]),
    ("Profitability", FILL_GROUP_MOM, [
        ("yf_gross_margin_q",     "Gross Profit Margin (Quarter)",     26, "0.00%", "center"),
        ("yf_operating_margin_q", "Operating Profit Margin (Quarter)", 30, "0.00%", "center"),
        ("yf_net_margin_q",       "Net Profit Margin (Quarter)",       26, "0.00%", "center"),
    ]),
    ("Management Effectiveness", FILL_GROUP_Q, [
        ("yf_roa_ttm",                   "Return on Assets (TTM)",            22, "0.00%", "center"),
        ("yf_roe_ttm",                   "Return on Equity (TTM)",            22, "0.00%", "center"),
        ("yf_roce_ttm",                  "Return on Capital Employed (TTM)",  32, "0.00%", "center"),
        ("yf_roic_ttm",                  "Return On Invested Capital (TTM)",  34, "0.00%", "center"),
        ("yf_dso_q",                     "Days Sales Outstanding (Quarter)",  32, "0.00",  "center"),
        ("yf_asset_turnover_ttm",        "Asset Turnover (TTM)",              20, "0.00",  "center"),
        ("yf_days_inventory_q",          "Days Inventory (Quarter)",          24, "0.00",  "center"),
        ("yf_days_payables_q",           "Days Payables Outstanding (Quarter)",34,"0.00",  "center"),
        ("yf_cash_conversion_cycle_q",   "Cash Conversion Cycle (Quarter)",   32, "0.00",  "center"),
        ("yf_receivables_turnover_q",    "Receivables Turnover (Quarter)",    30, "0.00",  "center"),
    ]),
    ("Solvency", FILL_GROUP_MSF, [
        ("yf_current_ratio_q",       "Current Ratio (Quarter)",              24, "0.00",  "center"),
        ("yf_quick_ratio_q",         "Quick Ratio (Quarter)",                22, "0.00",  "center"),
        ("yf_debt_equity_q",         "Debt to Equity Ratio (Quarter)",       30, "0.00",  "center"),
        ("yf_lt_debt_equity_q",      "LT Debt/Equity (Quarter)",             26, "0.00",  "center"),
        ("yf_liabilities_equity_q",  "Total Liabilities/Equity (Quarter)",   34, "0.00",  "center"),
        ("yf_financial_leverage_q",  "Financial Leverage (Quarter)",         28, "0.00",  "center"),
        ("yf_interest_coverage_ttm", "Interest Coverage (TTM)",              24, "0.00",  "center"),
        ("yf_free_cash_flow_q",      "Free cash flow (Quarter)",             24, "compact","center"),
        ("yf_altman_z_original",     "Altman Z-Score (Original)",            26, "0.00",  "center"),
        ("yf_altman_z_modified",     "Altman Z-Score (Modified)",            26, "0.00",  "center"),
    ]),
    ("Price Performance", FILL_GROUP_MACD_MOM, [
        ("yf_price_return_1m",  "1 Month Price Returns",       22, "0.00%", "center"),
        ("yf_price_return_3m",  "3 Month Price Returns",       22, "0.00%", "center"),
        ("yf_price_return_6m",  "6 Month Price Returns",       22, "0.00%", "center"),
        ("yf_price_return_1y",  "1 Year Price Returns",        22, "0.00%", "center"),
        ("yf_price_return_3y",  "3 Year Price Returns",        22, "0.00%", "center"),
        ("yf_price_return_ytd", "Year to Date Price Returns",  26, "0.00%", "center"),
        ("yf_high_52w",        "52 Week High",                 16, "#,##0", "center"),
        ("yf_low_52w",         "52 Week Low",                  16, "#,##0", "center"),
    ]),
    ("Balance Sheet", FILL_GROUP_MSF, [
        ("yf_book_value_q",                         "Book Value (Quarter)",                 24, "compact","center"),
        ("yf_book_value_per_share_current",         "Current Book Value Per Share",         30, "0.00",   "center"),
        ("yf_tangible_book_value_q",                "Tang. Book Value (Quarter)",           28, "compact","center"),
        ("yf_tangible_book_value_per_share_current","Current Tang. Book Value Per Share",   34, "0.00",   "center"),
        ("yf_short_term_debt_q",                    "Short-term Debt (Quarter)",            28, "compact","center"),
        ("yf_long_term_debt_q",                     "Long-term Debt (Quarter)",             28, "compact","center"),
        ("yf_cash_q",                               "Cash (Quarter)",                       18, "compact","center"),
        ("yf_total_assets_q",                       "Total Assets (Quarter)",               24, "compact","center"),
        ("yf_total_liabilities_q",                  "Total Liabilities (Quarter)",          28, "compact","center"),
    ]),
    ("Income Statement", FILL_GROUP_Q, [
        ("yf_eps_ttm_current",        "Current EPS (TTM)",              20, "0.00",   "center"),
        ("yf_eps_q_yoy_growth",       "EPS (Quarter YoY Growth)",       26, "0.00%",  "center"),
        ("yf_revenue_ttm",            "Revenue (TTM)",                  18, "compact","center"),
        ("yf_revenue_q_yoy_growth",   "Revenue (Quarter YoY Growth)",   32, "0.00%",  "center"),
        ("yf_net_income_ttm",         "Net Income (TTM)",               20, "compact","center"),
        ("yf_ebit_ttm",               "EBIT (TTM)",                     18, "compact","center"),
    ]),
    ("Cash Flow Statement", FILL_GROUP_PQ, [
        ("yf_cash_from_operations_ttm", "Cash From Operations (TTM)", 28, "compact","center"),
        ("yf_cash_from_investing_ttm",  "Cash From Investing (TTM)",  28, "compact","center"),
        ("yf_cash_from_financing_ttm",  "Cash From Financing (TTM)",  28, "compact","center"),
        ("yf_capital_expenditure_ttm",  "Capital expenditure (TTM)",  28, "compact","center"),
        ("yf_free_cash_flow_ttm",       "Free cash flow (TTM)",       24, "compact","center"),
    ]),
    ("Upcoming Dividend", FILL_GROUP_VR, [
        ("div1_year",           "Year",                 10, None,     "center"),
        ("div1_idr",            "Dividend (IDR)",       16, "compact","center"),
        ("div1_payout_ratio",   "Payout Ratio (%)",     18, "0.00%",  "center"),
        ("div1_dividend_yield", "Dividend Yield (%)",   20, "0.00%",  "center"),
        ("div1_exdate",         "Ex Date",              14, None,     "center"),
        ("div1_paydate",        "Pay Date",             14, None,     "center"),
    ]),
    ("Latest Dividend", FILL_GROUP_VR, [
        ("div2_year",           "Year",                 10, None,     "center"),
        ("div2_idr",            "Dividend (IDR)",       16, "compact","center"),
        ("div2_payout_ratio",   "Payout Ratio (%)",     18, "0.00%",  "center"),
        ("div2_dividend_yield", "Dividend Yield (%)",   20, "0.00%",  "center"),
        ("div2_exdate",         "Ex Date",              14, None,     "center"),
        ("div2_paydate",        "Pay Date",             14, None,     "center"),
    ]),
    ("PBV Band Analysis  ·  3-Year Rolling  ·  Stockbit-Style", FILL_GROUP_Q, [
        ("yf_pbv_curr",   "Current PBV",    12, "#,##0.00","center"),
        ("yf_pbv_mean",   "Mean PBV (3Y)",  14, "#,##0.00","center"),
        ("yf_pbv_plus2",  "+2 SD",          12, "#,##0.00","center"),
        ("yf_pbv_plus1",  "+1 SD",          12, "#,##0.00","center"),
        ("yf_pbv_minus1", "-1 SD",          12, "#,##0.00","center"),
        ("yf_pbv_minus2", "-2 SD",          12, "#,##0.00","center"),
        ("yf_pbv_zscore", "PBV Z-Score",    12, "0.00",   "center"),
        ("yf_pbv_pct",    "PBV Percentile", 14, "0.00",   "center"),
        ("yf_pbv_regime", "PBV Regime",     14, None,     "center"),
    ]),
]
# Colour map for PBV Regime labels
PBV_REGIME_COLORS = {
    "Deep Undervalued": "008000",
    "Undervalued":      "375623",
    "Fair Value":       "595959",
    "Overvalued":       "E36C09",
    "Bubble":           "C00000",
    # legacy labels
    "Deep Value":       "008000",
    "Value":            "375623",
    "Premium":          "E36C09",
    "Euphoric":         "C00000",
}

# Compact-format keys in the fundamental sheet
_COMPACT_KEYS = {
    "yf_shares_outstanding", "yf_float_shares", "yf_market_cap_current",
    "yf_enterprise_value_current", "yf_shares_outstanding_current",
    "yf_free_cash_flow_q", "yf_book_value_q", "yf_tangible_book_value_q",
    "yf_short_term_debt_q", "yf_long_term_debt_q", "yf_cash_q",
    "yf_total_assets_q", "yf_total_liabilities_q", "yf_revenue_ttm",
    "yf_net_income_ttm", "yf_ebit_ttm", "yf_cash_from_operations_ttm",
    "yf_cash_from_investing_ttm", "yf_cash_from_financing_ttm",
    "yf_capital_expenditure_ttm", "yf_free_cash_flow_ttm",
    "div1_idr", "div2_idr", "mcap",
}

def _build_fundamental_row(screener_row: dict) -> dict:
    """
    Merge screener row fields + yfinance fundamentals into a single flat dict
    keyed to FUNDAMENTAL_DETAIL_SCHEMA. All yfinance fields prefixed "yf_".
    """
    ticker = str(screener_row.get("ticker", "")).upper().strip()

    # Fetch yfinance fundamentals
    raw_fund = _fetch_fundamental_data(ticker)
    # PBV band: prefer source/Stockbit-style values from Raw/KSEI row, then yfinance approximation.
    pbv_band  = _stockbit_style_pbv_from_source_row(screener_row, _compute_pbv_bands(ticker))

    # Build combined row: screener fields + yf_-prefixed fundamental fields
    combined = {
        "ticker":     ticker,
        "close":      screener_row.get("close", np.nan),
        "pct_change": screener_row.get("pct_change", np.nan),
        "emiten":     screener_row.get("emiten", ""),
        "idx_sector": screener_row.get("idx_sector", ""),
    }

    # yfinance fields with yf_ prefix
    _YF_MAP = {
        # Profile
        "company_name":       "yf_company_name",
        "sector":             "yf_sector",
        "industry":           "yf_industry",
        "shares_outstanding": "yf_shares_outstanding",
        "float_shares":       "yf_float_shares",
        "free_float_pct":     "yf_free_float_pct",
        "ipo_date":           "yf_ipo_date",
        "business_summary":   "yf_business_summary",

        # Current Valuation
        "pe_annualised":              "yf_pe_annualised",
        "pe_ttm_current":             "yf_pe_ttm_current",
        "earnings_yield_ttm":         "yf_earnings_yield_ttm",
        "ps_ttm_current":             "yf_ps_ttm_current",
        "pbv_current":                "yf_pbv_current",
        "ev_ebit_ttm":                "yf_ev_ebit_ttm",
        "ev_ebitda_ttm":              "yf_ev_ebitda_ttm",
        "market_cap_current":         "yf_market_cap_current",
        "enterprise_value_current":   "yf_enterprise_value_current",
        "shares_outstanding_current": "yf_shares_outstanding_current",

        # Profitability
        "gross_margin_q":     "yf_gross_margin_q",
        "operating_margin_q": "yf_operating_margin_q",
        "net_margin_q":       "yf_net_margin_q",

        # Management Effectiveness
        "roa_ttm":                 "yf_roa_ttm",
        "roe_ttm":                 "yf_roe_ttm",
        "roce_ttm":                "yf_roce_ttm",
        "roic_ttm":                "yf_roic_ttm",
        "dso_q":                   "yf_dso_q",
        "asset_turnover_ttm":      "yf_asset_turnover_ttm",
        "days_inventory_q":        "yf_days_inventory_q",
        "days_payables_q":         "yf_days_payables_q",
        "cash_conversion_cycle_q": "yf_cash_conversion_cycle_q",
        "receivables_turnover_q":  "yf_receivables_turnover_q",

        # Solvency
        "current_ratio_q":       "yf_current_ratio_q",
        "quick_ratio_q":         "yf_quick_ratio_q",
        "debt_equity_q":         "yf_debt_equity_q",
        "lt_debt_equity_q":      "yf_lt_debt_equity_q",
        "liabilities_equity_q":  "yf_liabilities_equity_q",
        "financial_leverage_q":  "yf_financial_leverage_q",
        "interest_coverage_ttm": "yf_interest_coverage_ttm",
        "free_cash_flow_q":      "yf_free_cash_flow_q",
        "altman_z_original":     "yf_altman_z_original",
        "altman_z_modified":     "yf_altman_z_modified",

        # Price Performance
        "price_return_1m":  "yf_price_return_1m",
        "price_return_3m":  "yf_price_return_3m",
        "price_return_6m":  "yf_price_return_6m",
        "price_return_1y":  "yf_price_return_1y",
        "price_return_3y":  "yf_price_return_3y",
        "price_return_ytd": "yf_price_return_ytd",
        "high_52w":        "yf_high_52w",
        "low_52w":         "yf_low_52w",

        # Balance Sheet
        "book_value_q":                          "yf_book_value_q",
        "book_value_per_share_current":          "yf_book_value_per_share_current",
        "tangible_book_value_q":                 "yf_tangible_book_value_q",
        "tangible_book_value_per_share_current": "yf_tangible_book_value_per_share_current",
        "short_term_debt_q":                     "yf_short_term_debt_q",
        "long_term_debt_q":                      "yf_long_term_debt_q",
        "cash_q":                                "yf_cash_q",
        "total_assets_q":                        "yf_total_assets_q",
        "total_liabilities_q":                   "yf_total_liabilities_q",

        # Income Statement
        "eps_ttm_current":      "yf_eps_ttm_current",
        "eps_q_yoy_growth":     "yf_eps_q_yoy_growth",
        "revenue_ttm":          "yf_revenue_ttm",
        "revenue_q_yoy_growth": "yf_revenue_q_yoy_growth",
        "net_income_ttm":       "yf_net_income_ttm",
        "ebit_ttm":             "yf_ebit_ttm",

        # Cash Flow Statement
        "cash_from_operations_ttm": "yf_cash_from_operations_ttm",
        "cash_from_investing_ttm":  "yf_cash_from_investing_ttm",
        "cash_from_financing_ttm":  "yf_cash_from_financing_ttm",
        "capital_expenditure_ttm":  "yf_capital_expenditure_ttm",
        "free_cash_flow_ttm":       "yf_free_cash_flow_ttm",

        # Dividend support fields
        "dividend_yield": "yf_dividend_yield",
        "payout_ratio":   "yf_payout_ratio",
    }
    _PBV_MAP = {
        "pbv_curr":   "yf_pbv_curr",
        "pbv_mean":   "yf_pbv_mean",
        "pbv_plus1":  "yf_pbv_plus1",
        "pbv_plus2":  "yf_pbv_plus2",
        "pbv_minus1": "yf_pbv_minus1",
        "pbv_minus2": "yf_pbv_minus2",
        "pbv_zscore": "yf_pbv_zscore",
        "pbv_pct":    "yf_pbv_pct",
        "pbv_regime": "yf_pbv_regime",
    }

    for src_key, dest_key in _YF_MAP.items():
        combined[dest_key] = raw_fund.get(src_key, np.nan)

    # Track data source for transparency (shown in console log; accessible for future sheet display)
    combined["yf_data_source"] = raw_fund.get("_data_source", "yfinance")
    if raw_fund.get("_fallback_source"):
        combined["yf_data_source"] += f" + {raw_fund['_fallback_source']}"
        combined["yf_fallback_fields"] = raw_fund.get("_fallback_fields", [])

    # Prefer exact Stockbit/Raw values for IDX Fundamental Detail when present.
    # This prevents yfinance unit/coverage mismatches from overwriting the user's
    # Stockbit-sourced ratios and compact values.
    def _parse_stockbit_value(_v):
        try:
            if pd.isna(_v):
                return np.nan
        except Exception:
            pass
        if isinstance(_v, (int, float, np.integer, np.floating)):
            return float(_v)
        s = str(_v).strip()
        if s in ("", "-", "N/A", "nan", "None"):
            return np.nan
        is_pct = s.endswith("%")
        s = s.replace("%", "").replace(",", "").strip()
        mult = 1.0
        if s and s[-1].upper() in ("K", "M", "B", "T"):
            suffix = s[-1].upper()
            s = s[:-1].strip()
            mult = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}.get(suffix, 1.0)
        try:
            val = float(s) * mult
            return val / 100.0 if is_pct else val
        except Exception:
            return np.nan

    _STOCKBIT_FUNDAMENTAL_MAP = {
        "Current PE Ratio (Annualised)": "yf_pe_annualised",
        "Current PE Ratio (TTM)": "yf_pe_ttm_current",
        "Earnings Yield (TTM)": "yf_earnings_yield_ttm",
        "Current Price to Sales (TTM)": "yf_ps_ttm_current",
        "Current Price to Book Value": "yf_pbv_current",
        "EV to EBIT (TTM)": "yf_ev_ebit_ttm",
        "EV to EBITDA (TTM)": "yf_ev_ebitda_ttm",
        "Market Cap": "yf_market_cap_current",
        "Enterprise Value": "yf_enterprise_value_current",
        "Current Share Outstanding": "yf_shares_outstanding_current",
        "Gross Profit Margin (Quarter)": "yf_gross_margin_q",
        "Operating Profit Margin (Quarter)": "yf_operating_margin_q",
        "Net Profit Margin (Quarter)": "yf_net_margin_q",
        "Return on Assets (TTM)": "yf_roa_ttm",
        "Return on Equity (TTM)": "yf_roe_ttm",
        "Return on Capital Employed (TTM)": "yf_roce_ttm",
        "Return On Invested Capital (TTM)": "yf_roic_ttm",
        "Days Sales Outstanding (Quarter)": "yf_dso_q",
        "Asset Turnover (TTM)": "yf_asset_turnover_ttm",
        "Days Inventory (Quarter)": "yf_days_inventory_q",
        "Days Payables Outstanding (Quarter)": "yf_days_payables_q",
        "Cash Conversion Cycle (Quarter)": "yf_cash_conversion_cycle_q",
        "Receivables Turnover (Quarter)": "yf_receivables_turnover_q",
        "Current Ratio (Quarter)": "yf_current_ratio_q",
        "Quick Ratio (Quarter)": "yf_quick_ratio_q",
        "Debt to Equity Ratio (Quarter)": "yf_debt_equity_q",
        "LT Debt/Equity (Quarter)": "yf_lt_debt_equity_q",
        "Total Liabilities/Equity (Quarter)": "yf_liabilities_equity_q",
        "Financial Leverage (Quarter)": "yf_financial_leverage_q",
        "Interest Coverage (TTM)": "yf_interest_coverage_ttm",
        "Free cash flow (Quarter)": "yf_free_cash_flow_q",
        "Altman Z-Score (Original)": "yf_altman_z_original",
        "Altman Z-Score (Modified)": "yf_altman_z_modified",
        "1 Month Price Returns": "yf_price_return_1m",
        "3 Month Price Returns": "yf_price_return_3m",
        "6 Month Price Returns": "yf_price_return_6m",
        "1 Year Price Returns": "yf_price_return_1y",
        "3 Year Price Returns": "yf_price_return_3y",
        "Year to Date Price Returns": "yf_price_return_ytd",
        "52 Week High": "yf_high_52w",
        "52 Week Low": "yf_low_52w",
        "Book Value (Quarter)": "yf_book_value_q",
        "Current Book Value Per Share": "yf_book_value_per_share_current",
        "Tang. Book Value (Quarter)": "yf_tangible_book_value_q",
        "Current Tang. Book Value Per Share": "yf_tangible_book_value_per_share_current",
        "Short-term Debt (Quarter)": "yf_short_term_debt_q",
        "Long-term Debt (Quarter)": "yf_long_term_debt_q",
        "Cash (Quarter)": "yf_cash_q",
        "Total Assets (Quarter)": "yf_total_assets_q",
        "Total Liabilities (Quarter)": "yf_total_liabilities_q",
        "Current EPS (TTM)": "yf_eps_ttm_current",
        "EPS (Quarter YoY Growth)": "yf_eps_q_yoy_growth",
        "Revenue (TTM)": "yf_revenue_ttm",
        "Revenue (Quarter YoY Growth)": "yf_revenue_q_yoy_growth",
        "Net Income (TTM)": "yf_net_income_ttm",
        "EBIT (TTM)": "yf_ebit_ttm",
        "Cash From Operations (TTM)": "yf_cash_from_operations_ttm",
        "Cash From Investing (TTM)": "yf_cash_from_investing_ttm",
        "Cash From Financing (TTM)": "yf_cash_from_financing_ttm",
        "Capital expenditure (TTM)": "yf_capital_expenditure_ttm",
        "Free cash flow (TTM)": "yf_free_cash_flow_ttm",
    }
    for _src_col, _dest_key in _STOCKBIT_FUNDAMENTAL_MAP.items():
        _val = _parse_stockbit_value(screener_row.get(_src_col, np.nan))
        if pd.notna(_val):
            combined[_dest_key] = _val

    for src_key, dest_key in _PBV_MAP.items():
        combined[dest_key] = pbv_band.get(src_key, np.nan)

    # Dividend: 8-column 2-entry layout
    div_entries = _build_dividend_entries(ticker)
    for k, v in div_entries.items():
        combined[k] = v

    # Upgrade 8 — Dividend trap detection (uses already-fetched fundamental data)
    combined["div_trap_flag"] = detect_dividend_trap(combined)

    # Legacy dividend history (kept for Guide reference, not in schema)
    div_hist = _fetch_dividend_history(ticker)
    combined["yf_div_2025"]           = div_hist.get("div_2025", np.nan)
    combined["yf_div_2024"]           = div_hist.get("div_2024", np.nan)
    combined["yf_div_2023"]           = div_hist.get("div_2023", np.nan)
    combined["yf_div_cagr"]           = div_hist.get("div_cagr", np.nan)
    combined["yf_div_consecutive"]    = div_hist.get("div_consecutive_years", 0)
    combined["yf_upcoming_dividend"]  = div_hist.get("upcoming_dividend", np.nan)
    combined["yf_upcoming_ex_date"]   = div_hist.get("upcoming_ex_date", "N/A")
    combined["yf_upcoming_pay_date"]  = div_hist.get("upcoming_pay_date", "N/A")
    combined["yf_upcoming_div_yield"] = div_hist.get("upcoming_div_yield", np.nan)

    # Disclosure & Corporate Events group removed from IDX Fundamental Detail.

    return combined


def _institutional_trade_plan(r: dict, hist: "pd.DataFrame | None" = None) -> dict:
    """
    S-02 + S-03: Fully rebuilt Trade Plan Engine.

    POI SOURCES (16 levels per spec):
      • PrevQ VWAP, ±1SD, ±2SD, ±3SD  (6 levels)
      • PrevY VWAP, ±1SD, ±2SD, ±3SD  (6 levels)
      • SMA200
      • OB Bull High / OB Bull Low
      • OB Bear High / OB Bear Low
      • EQ High / EQ Low  (from SMC equilibrium range)
      • P Trade  (5-day wick low — SL only, not an entry POI)
      • IBH / IBL  (initial balance, only if day_of_month_trading ≥ 3)
      • PWH / PWL  (previous week high / low)

    OUTPUT (10 columns per spec):
      tp_entry_poi        — Entry POI label
      tp_entry            — Entry price
      tp_entry_dist_pct   — (Entry − Current) / Current
      tp_target_poi       — Target POI label
      tp_target           — Target price
      tp_target_upside_pct— (Target − Entry) / Entry
      tp_inv_poi          — Invalidation POI label
      tp_inv              — Invalidation price
      tp_inv_downside_pct — (Entry − Invalidation) / Entry
      tp_rr               — Target Upside % / Invalidation Downside %  (numeric)
    """
    EMPTY = {
        "tp_entry_poi": "-", "tp_entry": np.nan, "tp_entry_dist_pct": np.nan,
        "tp_target_poi": "-", "tp_target": np.nan, "tp_target_upside_pct": np.nan,
        "tp_inv_poi": "-", "tp_inv": np.nan, "tp_inv_downside_pct": np.nan,
        "tp_rr": np.nan,
    }

    close = safe_num(r.get("close"), np.nan)
    if pd.isna(close) or close <= 0:
        return EMPTY

    # ── Helper: parse "low-high" range string → (low, high) ──────────────
    def _parse_range(s):
        try:
            pts = str(s or "").replace("–", "-").replace(",", "").split("-")
            if len(pts) == 2:
                a, b = float(pts[0].strip()), float(pts[1].strip())
                return (min(a,b), max(a,b))
        except Exception:
            pass
        return (np.nan, np.nan)

    pois = {}   # {label: price}

    # 1. PrevQ VWAP ± 1/2/3 SD
    for suf, key in [("VWAP","_vwap"),("-1SD","_m1"),("-2SD","_m2"),("-3SD","_m3"),("+1SD","_p1"),("+2SD","_p2"),("+3SD","_p3")]:
        v = safe_num(r.get(f"pq{key}"), np.nan)
        if pd.notna(v) and v > 0:
            pois[f"PrevQ {suf}"] = v

    # 2. PrevY VWAP ± 1/2/3 SD
    for suf, key in [("VWAP","_vwap"),("-1SD","_m1"),("-2SD","_m2"),("-3SD","_m3"),("+1SD","_p1"),("+2SD","_p2"),("+3SD","_p3")]:
        v = safe_num(r.get(f"py{key}"), np.nan)
        if pd.notna(v) and v > 0:
            pois[f"PrevY {suf}"] = v

    # 3. SMA200
    s200 = safe_num(r.get("sma200"), np.nan)
    if pd.notna(s200) and s200 > 0:
        pois["SMA200"] = s200

    # 4. EMA25 / EMA50
    for name in ("ema25", "ema50"):
        v = safe_num(r.get(name), np.nan)
        if pd.notna(v) and v > 0:
            pois[name.upper()] = v

    # 5. OB Bull High / Low
    ob_bull_lo, ob_bull_hi = _parse_range(r.get("smc_closest_ob", ""))
    if pd.notna(ob_bull_hi) and ob_bull_hi > 0:
        pois["OB Bull High"] = ob_bull_hi
    if pd.notna(ob_bull_lo) and ob_bull_lo > 0:
        pois["OB Bull Low"]  = ob_bull_lo

    # 6. OB Bear High / Low
    ob_bear_lo, ob_bear_hi = _parse_range(r.get("smc_closest_ob_bear", ""))
    if pd.notna(ob_bear_hi) and ob_bear_hi > 0:
        pois["OB Bear High"] = ob_bear_hi
    if pd.notna(ob_bear_lo) and ob_bear_lo > 0:
        pois["OB Bear Low"]  = ob_bear_lo

    # 7. EQ High / EQ Low  (from smc_equilibrium range string)
    eq_lo, eq_hi = _parse_range(r.get("smc_equilibrium", ""))
    if pd.notna(eq_hi) and eq_hi > 0:
        pois["EQ High"] = eq_hi
    if pd.notna(eq_lo) and eq_lo > 0:
        pois["EQ Low"]  = eq_lo

    # 8. PWH / PWL  — previous week high / low from hist
    if hist is not None and len(hist) >= 5:
        try:
            _now   = hist.index[-1]
            # Find last completed ISO week (Mon–Fri before this week)
            _wday  = _now.day_of_week   # 0=Mon
            _days_back = _wday + 7      # start of last week at minimum
            _prev_wk   = hist[hist.index < _now - pd.Timedelta(days=_wday)].tail(5)
            if len(_prev_wk) >= 3:
                pwh = float(_prev_wk["High"].max())
                pwl = float(_prev_wk["Low"].min())
                if pwh > 0: pois["PWH"] = pwh
                if pwl > 0: pois["PWL"] = pwl
        except Exception:
            pass

    # 9. IBH / IBL  — initial balance (first 2 trading days of month)
    if hist is not None and len(hist) >= 3:
        try:
            last_ts   = hist.index[-1]
            month_start = hist[hist.index.month == last_ts.month].head(2)
            trading_day_of_month = len(hist[hist.index.month == last_ts.month])
            if trading_day_of_month >= 3 and len(month_start) >= 2:
                ibh = float(month_start["High"].max())
                ibl = float(month_start["Low"].min())
                if ibh > 0: pois["IBH"] = ibh
                if ibl > 0: pois["IBL"] = ibl
        except Exception:
            pass

    # 10. P Trade  (5-day wick low — SL anchor only, tag with * so it's not used as Entry)
    if hist is not None and len(hist) >= 5:
        try:
            p_trade = float(hist["Low"].iloc[-5:].min())
            if p_trade > 0:
                pois["P Trade*"] = p_trade   # * = SL-only level
        except Exception:
            pass

    if not pois:
        return EMPTY

    # ── Separate SL-only POIs ─────────────────────────────────────────────
    sl_only = {k for k in pois if k.endswith("*")}
    entry_pois = {k: v for k, v in pois.items() if k not in sl_only}

    if not entry_pois:
        return EMPTY

    # ── Entry POI = nearest entry-eligible POI to current close ──────────
    entry_name = min(entry_pois, key=lambda k: abs(entry_pois[k] - close))
    entry_px   = entry_pois[entry_name]

    # ── Invalidation = lowest SL-level below entry ────────────────────────
    # Use P Trade if available and below entry; else ATR proxy below entry
    atr_pct   = max(safe_num(r.get("atr14_pct"), 3.0), 1.0) / 100.0
    atr_proxy = entry_px * atr_pct

    p_trade_val = pois.get("P Trade*", np.nan)
    if pd.notna(p_trade_val) and p_trade_val < entry_px:
        inv_px   = p_trade_val
        inv_name = "P Trade (5D low)"
    else:
        # Next lower POI as invalidation, fallback ATR
        lower_entries = sorted(
            [(k, v) for k, v in entry_pois.items() if v < entry_px - (atr_proxy * 0.3)],
            key=lambda x: x[1], reverse=True
        )
        if lower_entries:
            inv_name, inv_px = lower_entries[0]
        else:
            inv_px   = entry_px - atr_proxy
            inv_name = "ATR Stop"

    risk = entry_px - inv_px   # always positive if inv < entry

    # ── Target POI = best higher POI by R/R ≥ 2, then ≥ 1 ───────────────
    higher_entries = sorted(
        [(k, v) for k, v in entry_pois.items() if v > entry_px + (atr_proxy * 0.3)],
        key=lambda x: x[1]
    )
    target_name, target_px = np.nan, np.nan
    for t_name, t_val in higher_entries:
        if risk > 0 and (t_val - entry_px) / risk >= 2.0:
            target_name, target_px = t_name, t_val
            break
    if pd.isna(target_px):
        for t_name, t_val in higher_entries:
            if risk > 0 and (t_val - entry_px) / risk >= 1.0:
                target_name, target_px = t_name, t_val
                break
    if pd.isna(target_px) and higher_entries:
        target_name, target_px = higher_entries[-1]
    if pd.isna(target_px):
        target_px   = entry_px * 1.08
        target_name = "Proj +8%"

    # ── Compute the 10 output metrics ─────────────────────────────────────
    def _pct(num, den): return (num / den) if den and den > 0 and pd.notna(num) and pd.notna(den) else np.nan

    entry_dist_pct     = _pct(entry_px  - close,    close)
    target_upside_pct  = _pct(target_px - entry_px, entry_px)   # (Target − Entry) / Entry
    inv_downside_pct   = _pct(entry_px  - inv_px,   entry_px)   # (Entry − Inv) / Entry
    rr = _pct(target_upside_pct, inv_downside_pct) if inv_downside_pct and inv_downside_pct > 0 else np.nan

    def _r(v): return round(v, 0) if pd.notna(v) else np.nan
    def _p(v): return round(v, 4) if pd.notna(v) else np.nan

    return {
        "tp_entry_poi":         entry_name,
        "tp_entry":             _r(entry_px),
        "tp_entry_dist_pct":    _p(entry_dist_pct),
        "tp_target_poi":        str(target_name) if pd.notna(target_name) else "-",
        "tp_target":            _r(target_px),
        "tp_target_upside_pct": _p(target_upside_pct),
        "tp_inv_poi":           inv_name,
        "tp_inv":               _r(inv_px),
        "tp_inv_downside_pct":  _p(inv_downside_pct),
        "tp_rr":                _p(rr),
    }

def _is_swing_bos_today(r: dict, market_date: str = None) -> bool:
    """
    Swing BOS Confirmation: strict Swing BOS on the selected MARKET_DATE.

    This no longer uses the computer's current date or a T+1 grace window.
    A row qualifies only when the latest SMC/market-structure event is BOS
    (not CHoCH) and its event date equals MARKET_DATE.
    """
    target = pd.Timestamp(market_date or globals().get("MARKET_DATE", get_market_date())).date()

    candidates = [
        (r.get("smc_latest_swing_struct"), r.get("smc_latest_swing_struct_date")),
        (r.get("ms_last_event"), r.get("ms_last_event_date")),
    ]

    import re
    for label, dt_val in candidates:
        txt = str(label or "").strip()
        up = txt.upper()
        if "BOS" not in up or "CHOCH" in up:
            continue

        event_date = None

        # Prefer explicit date column when available.
        if dt_val not in (None, "", "-", "N/A"):
            try:
                event_date = pd.Timestamp(dt_val).date()
            except Exception:
                event_date = None

        # Fallback: parse yyyy-mm-dd embedded in text.
        if event_date is None:
            m = re.search(r"\d{4}-\d{2}-\d{2}", txt)
            if m:
                try:
                    event_date = pd.Timestamp(m.group()).date()
                except Exception:
                    event_date = None

        return event_date == target

    return False


def _is_bos_today(r: dict, market_date: str = None) -> bool:
    """
    Extended BOS confirmation: passes if EITHER swing BOS OR internal BOS
    fires on MARKET_DATE. Swing BOS alone was the prior behaviour;
    internal BOS gives earlier entries on tighter structure breaks.
    """
    target = pd.Timestamp(market_date or globals().get("MARKET_DATE", get_market_date())).date()

    candidates = [
        (r.get("smc_latest_swing_struct"),    r.get("smc_latest_swing_struct_date")),
        (r.get("smc_latest_internal_struct"), r.get("smc_latest_internal_struct_date")),
        (r.get("ms_last_event"),              r.get("ms_last_event_date")),
    ]

    import re
    for label, dt_val in candidates:
        txt = str(label or "").strip()
        up  = txt.upper()
        if "BOS" not in up or "CHOCH" in up:
            continue
        event_date = None
        if dt_val not in (None, "", "-", "N/A"):
            try:
                event_date = pd.Timestamp(dt_val).date()
            except Exception:
                event_date = None
        if event_date is None:
            m = re.search(r"\d{4}-\d{2}-\d{2}", txt)
            if m:
                try:
                    event_date = pd.Timestamp(m.group()).date()
                except Exception:
                    event_date = None
        if event_date == target:
            return True

    return False


def _is_smc_location_screener(r: dict) -> bool:
    """
    SMC Location filter: price is positioned in a structurally favourable zone.
    Passes when smc_premium_discount is one of:
        'In Bull OB'  — price inside the nearest bullish order block
        'Equilibrium' — price at 50% of the structure range (EQ)
        'Discount'    — price below equilibrium (below-50% zone)
    Excludes Premium, In Bear OB, Mid-Zone, and N/A.
    """
    pd_label = str(r.get("smc_premium_discount", "") or "").strip()
    return pd_label in ("In Bull OB", "Equilibrium", "Discount")


def build_idx_screener_sheet(wb, latest_market_day: str, rows: list):
    """
    IDX Screener — Redesigned v2.
    Layout:
        Row 1  : Title bar  (dark slate, full-width, 38px)
        Row 2  : Stats bar  (filter counts + gate info, 22px)
        Row 3  : Spacer (4px)
        Row 4  : Column group headers  (STOCK INFO | SIGNAL & PRICE | TRADE PLAN)
        Row 5  : Column sub-headers
        Row 6+ : Section banner → data rows → spacer → next section …

    Filters: A (EMA trend) | B (Golden Cross) | C (Swing BOS) | D (POI touch) | E (EQ breakout)
    Each filter section has a unique color DNA: header fill + row fill + left-accent border.
    Columns: #  Ticker  Sector  |  Price  Chg%  RVOL  ADR%  Zone  MA  |
             POI  Entry  Target  Upside%  Stop  R/R
    """
    sheet_name = "IDX Screener"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    # ── Palette ──────────────────────────────────────────────────────────────
    # Title / stats bar
    C_TITLE       = "0D1B2A"   # near-black navy
    C_STATS_BAR   = "162436"   # slightly lighter navy
    C_GRP_STOCK   = "1A3A5C"   # deep navy — Stock Info group
    C_GRP_SIGNAL  = "1E5631"   # dark forest — Signal & Price group
    C_GRP_TRADE   = "4A1060"   # deep violet — Trade Plan group

    # Per-filter DNA: (hdr_fill, row_fill, accent_border_color, label_emoji)
    FILTER_DNA = {
        "A": ("1F4E79", "EAF2FB", "2E75B6", "▲"),  # navy — EMA trend
        "B": ("375623", "EDF7E8", "548235", "✦"),  # forest — Golden Cross
        "C": ("7B3F00", "FFF3E0", "E36C09", "◈"),  # amber — BOS (Swing + Internal)
        "D": ("4A0080", "F5EEFF", "7030A0", "◉"),  # violet — POI touch
        "E": ("005050", "E0F7F7", "00797A", "⬡"),  # teal — EQ breakout
        "F": ("6B2737", "FDEEEF", "9B1C31", "◆"),  # burgundy — Near VWAP zone
        "G": ("3B3000", "FFFDE7", "7A6800", "⬟"),  # olive — SMC Location
    }

    # Shared fills
    FILL_TITLE    = PatternFill("solid", fgColor=C_TITLE)
    FILL_STATS    = PatternFill("solid", fgColor=C_STATS_BAR)
    FILL_GRP_STK  = PatternFill("solid", fgColor=C_GRP_STOCK)
    FILL_GRP_SIG  = PatternFill("solid", fgColor=C_GRP_SIGNAL)
    FILL_GRP_TRD  = PatternFill("solid", fgColor=C_GRP_TRADE)
    FILL_SUBHDR   = PatternFill("solid", fgColor="1C3253")   # slightly lighter for sub-header
    FILL_EMPTY_ROW= PatternFill("solid", fgColor="F2F4F7")

    # Fonts
    FT_TITLE      = Font(name="Calibri", size=16, bold=True,   color="FFFFFF")
    FT_STATS      = Font(name="Calibri", size=9,  italic=True, color="B0C8E0")
    FT_GRP        = Font(name="Calibri", size=9,  bold=True,   color="FFFFFF")
    FT_SUBHDR     = Font(name="Calibri", size=9,  bold=True,   color="FFFFFF")
    FT_BODY       = Font(name="Calibri", size=10,              color="1F2937")
    FT_BODY_SMALL = Font(name="Calibri", size=9,               color="1F2937")
    FT_MUTED      = Font(name="Calibri", size=9,  italic=True, color="6B7280")
    FT_RANK       = Font(name="Calibri", size=9,  bold=True,   color="6B7280")

    # Alignments
    AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    AL_R = Alignment(horizontal="right",  vertical="center")

    # Borders
    def _thin_side(color="D1D5DB"): return Side(style="thin", color=color)
    def _med_side(color):           return Side(style="medium", color=color)
    BDR_CELL = Border(
        left=_thin_side(), right=_thin_side(),
        top=_thin_side(),  bottom=_thin_side()
    )

    # ── Column schema ─────────────────────────────────────────────────────────
    # (key, header, width, fmt, align, group)
    # group: "S"=Stock, "P"=Signal/Price, "T"=Trade Plan
    COLUMNS = [
        ("_rank",          "#",            4,    None,     "center", "S"),
        ("ticker",         "Ticker",      11,    None,     "center", "S"),
        # Signal Category column deleted (Item #11 — redundant)
        ("idx_sector",     "Sector",      18,    None,     "left",   "S"),
        ("close",          "Price",       10,    "#,##0",  "center", "P"),
        ("pct_change",     "Chg %",        9,    "0.00%",  "center", "P"),
        ("rvol20",         "RVOL",         8,    "0.00",   "center", "P"),
        ("adr_pct",        "ADR %",        8,    "0.0%",   "center", "P"),
        ("q_remarks",      "Current Q VWAP", 26, None,     "left",   "P"),
        ("pq_remarks",     "Prev Q VWAP",    26, None,     "left",   "P"),
        ("py_remarks",     "Prev Y VWAP",    26, None,     "left",   "P"),
        ("summary_ma",     "MA",          20,    None,     "left",   "P"),
        ("summary_filter", "Summary Screener", 54,  None,     "left",   "P"),
        ("_news_sentiment","Sentiment News", 44,  None,     "left",   "N"),
        ("_corp_action",   "Corp. Action",   44,  None,     "left",   "N"),
        # S-02+S-03: 10-column Trade Plan (rebuilt per spec)
        ("tp_entry_poi",         "Entry POI",           20, None,    "left",   "T"),
        ("tp_entry",             "Entry",               12, "#,##0", "center", "T"),
        ("tp_entry_dist_pct",    "Entry Distance %",    14, "0.0%",  "center", "T"),
        ("tp_target_poi",        "Target POI",          20, None,    "left",   "T"),
        ("tp_target",            "Target",              12, "#,##0", "center", "T"),
        ("tp_target_upside_pct", "Target Upside %",     14, "0.0%",  "center", "T"),
        ("tp_inv_poi",           "Invalidation POI",    20, None,    "left",   "T"),
        ("tp_inv",               "Invalidation",        12, "#,##0", "center", "T"),
        ("tp_inv_downside_pct",  "Invalidation Down %", 14, "0.0%",  "center", "T"),
        ("tp_rr",                "R/R",                 10, "0.0x",  "center", "T"),
        ("beta_zone",      "Beta Zone",   14,    None,     "center", "Q"),
        ("rs_rating",      "RS Rating",   10,    "0.0",    "center", "Q"),
    ]
    N_COLS  = len(COLUMNS)
    KEYS    = [c[0] for c in COLUMNS]
    HDRS    = [c[1] for c in COLUMNS]
    WIDTHS  = [c[2] for c in COLUMNS]
    FMTS    = [c[3] for c in COLUMNS]
    ALIGNS  = [c[4] for c in COLUMNS]
    GROUPS  = [c[5] for c in COLUMNS]

    # Column indices (1-based): col A = 1 (spacer), data starts col B = 2
    COL_START = 2
    COL_END   = COL_START + N_COLS - 1

    ws.column_dimensions["A"].width = 0.5
    for i, (_, _, w, _, _, _) in enumerate(COLUMNS, start=COL_START):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Filter logic (unchanged from original) ────────────────────────────────
    def _not_suspended(r):
        return not r.get("suspended", False)

    def _good_adtv(r):
        return (safe_num(r.get("adtr20"), 0) >= 5_000_000_000 or
                safe_num(r.get("adtv20"), 0) >= 5_000_000)

    def _above_ema25_50(r):
        # S-04: RSI > 50 removed — handled by global gate; condition is purely EMA alignment
        cl  = safe_num(r.get("close"),  np.nan)
        e25 = safe_num(r.get("ema25"),  np.nan)
        e50 = safe_num(r.get("ema50"),  np.nan)
        return (
            pd.notna(cl) and pd.notna(e25) and pd.notna(e50)
            and cl > e25 and e25 > e50
        )

    def _filter_d_poi_touch(r):
        """S-07: POI Reclaim — low ≤ POI × 1.005 and close > POI.
        Sources: Prev Q/Y VWAP & SD bands, Internal OB EQ, EMA25, EMA50, SMA200."""
        close = safe_num(r.get("close"), np.nan)
        low   = safe_num(r.get("last_low"), np.nan)
        if pd.isna(close) or pd.isna(low):
            return False
        pois = {}
        for prefix, label in [("pq","PrevQ"), ("py","PrevY")]:
            for key, suf in [("_vwap","VWAP"),("_m1","-1SD"),("_m2","-2SD")]:
                v = safe_num(r.get(f"{prefix}{key}"), np.nan)
                if pd.notna(v) and v > 0:
                    pois[label+suf] = v
        ob_eq = safe_num(r.get("smc_ob_equilibrium"), np.nan)
        if pd.notna(ob_eq): pois["Equilibrium"] = ob_eq
        e25  = safe_num(r.get("ema25"),  np.nan)
        e50  = safe_num(r.get("ema50"),  np.nan)
        s200 = safe_num(r.get("sma200"), np.nan)
        if pd.notna(e25):  pois["EMA25"]  = e25
        if pd.notna(e50):  pois["EMA50"]  = e50
        if pd.notna(s200): pois["SMA200"] = s200   # S-07: SMA200 added
        for _, poi_val in pois.items():
            if poi_val > 0 and low <= poi_val * 1.005 and close > poi_val:
                return True
        return False

    def _filter_e_eq_breakout(r):
        """
        S-08: EQ Breakout — SMC equilibrium range breakout.
        Condition: price consolidates between EQ Low and EQ High for >= 2 consecutive
        days before MARKET_DATE, then on MARKET_DATE closes above EQ High.
        Reads pre-computed flag from row["eq_breakout_flag"] set in build_row.
        """
        # Primary: pre-computed breakout flag from build_row (uses hist)
        if r.get("eq_breakout_flag") is True:
            return True
        # Fallback: SMC today_event breakout label
        today_ev = str(r.get("today_event", "") or "")
        if "Break" in today_ev and "Equilibrium" in today_ev:
            return True
        return False

    def _is_near_vwap_zone(r):
        """S-09: Near Prev Q or Prev Y VWAP only — Current Q excluded."""
        for k in ("pq_remarks", "py_remarks"):   # q_remarks (Current Q) excluded
            if str(r.get(k, "") or "").startswith("Price Near"):
                return True
        return False

    def _format_px(v):
        try:
            if v is None or pd.isna(v):
                return "-"
            return compact_fmt(v)
        except Exception:
            return "-"

    def _poi_hits_for_row(r, close, low):
        """Return ticker-specific POI touch/reclaim details."""
        hits = []
        candidates = [
            ("q_vwap",  "Current Q VWAP"), ("q_m1",  "Current Q -1 SD"), ("q_m2",  "Current Q -2 SD"),
            ("pq_vwap", "Prev Q VWAP"),    ("pq_m1", "Prev Q -1 SD"),    ("pq_m2", "Prev Q -2 SD"),
            ("py_vwap", "Prev Y VWAP"),    ("py_m1", "Prev Y -1 SD"),    ("py_m2", "Prev Y -2 SD"),
            ("smc_ob_equilibrium", "Internal OB Equilibrium"),
            ("ema25", "EMA25"), ("ema50", "EMA50"),
        ]
        for key, name in candidates:
            v = safe_num(r.get(key), np.nan)
            if pd.notna(low) and pd.notna(close) and pd.notna(v) and v > 0:
                if low <= v * 1.005 and close > v:
                    dist = ((close / v) - 1.0) * 100.0
                    hits.append(f"{name} {_format_px(v)} reclaimed; close {dist:+.2f}% above")
        return hits

    def _filter_reason(r, tag):
        """
        Ticker-specific row explanation.
        - No generic liquidity-gate text.
        - Expands POI meaning and Golden Cross source.
        - Uses only values/events present in the row.
        """
        try:
            close = safe_num(r.get("close"), np.nan)
            low   = safe_num(r.get("last_low"), np.nan)
            rsi_v = safe_num(r.get("rsi14", r.get("rsi")), np.nan)
            e25   = safe_num(r.get("ema25"), np.nan)
            e50   = safe_num(r.get("ema50"), np.nan)
            parts = []

            if tag == "A":
                streak = safe_int(r.get("ema_trend_streak", r.get("above_ema_streak", 0)), 0)
                base = (
                    f"Close {_format_px(close)} > EMA25 {_format_px(e25)}; "
                    f"EMA25 {_format_px(e25)} > EMA50 {_format_px(e50)}"
                )
                if pd.notna(rsi_v):
                    base += f"; RSI {rsi_v:.1f} > 50 confirms positive momentum"
                if streak > 0:
                    base += f"; trend stack persisted {streak} sessions"
                parts.append(base)

            elif tag == "B":
                # S-05: Show which exact trigger fired
                trig = str(r.get("_gc_trigger", "") or "")
                cross_status = str(r.get("cross_status", "") or "").strip()
                macd_cross   = str(r.get("macd_cross", "") or "").strip()
                macd_v       = safe_num(r.get("macd", r.get("macd_line")), np.nan)
                macd_sig     = safe_num(r.get("macd_signal"), np.nan)
                ema_sma      = str(r.get("ema50_sma200_cross", "") or "")
                e50  = safe_num(r.get("ema50"),  np.nan)
                s200 = safe_num(r.get("sma200"), np.nan)

                reasons = []
                if "EMA×" in trig or ("Today" in ema_sma and "Golden" in ema_sma):
                    reasons.append(f"EMA× — EMA50 {_format_px(e50)} crossed above SMA200 {_format_px(s200)} TODAY")
                elif "RSI×" in trig or cross_status == "Golden":
                    reasons.append(f"RSI× — RSI14 {rsi_v:.1f} crossed above RSI MA14 TODAY")
                elif "MACD×" in trig or macd_cross == "Golden Cross":
                    if pd.notna(macd_v) and pd.notna(macd_sig):
                        reasons.append(f"MACD× — MACD {macd_v:.2f} crossed above signal {macd_sig:.2f} TODAY")
                    else:
                        reasons.append("MACD× — MACD Line crossed above Signal TODAY")
                parts.append("; ".join(reasons) if reasons else "Golden Cross — trigger type unknown")

            elif tag == "C":
                evt  = str(r.get("smc_latest_swing_struct", r.get("latest_swing_struct", "N/A")) or "N/A")
                dt   = str(r.get("smc_latest_swing_struct_date", r.get("latest_swing_struct_date", latest_market_day)) or latest_market_day)
                lvl  = _format_px(r.get("smc_swing_break_level", r.get("swing_break_level", np.nan)))
                parts.append(f"Swing BOS Confirmation: {evt} on {dt}" + (f"; break level {lvl}" if lvl != "-" else ""))

            elif tag == "D":
                hits = _poi_hits_for_row(r, close, low)
                explanation = " | ".join(hits) if hits else "Nearest POI touched and reclaimed based on current row levels."
                parts.append(explanation)

            elif tag == "E":
                event = str(r.get("today_event", "EQ Breakout") or "EQ Breakout")
                q_rem  = str(r.get("q_remarks", "") or "")
                pq_rem = str(r.get("pq_remarks", "") or "")
                zone_note = f" · VWAP zone: CQ={q_rem[:20]}, PQ={pq_rem[:20]}" if q_rem or pq_rem else ""
                parts.append(f"EQ Breakout: {event}{zone_note}")

            elif tag == "F":
                near = []
                for label, key in [("Current Q", "q_remarks"), ("Prev Q", "pq_remarks"), ("Prev Y", "py_remarks")]:
                    val = str(r.get(key, "") or "")
                    if val.startswith("Price Near"):
                        near.append(f"{label}: {val}")
                parts.append("Near VWAP Zone: " + (" | ".join(near) if near else "no near VWAP zone label"))

            return " · ".join([p for p in parts if p and p != "N/A"]) or "-"
        except Exception:
            return "-"
    live_rows = [r for r in rows if _not_suspended(r)]

    # GLOBAL MASTER LIQUIDITY GATE
    # Every IDX Screener filter below must start from this gated universe:
    #   ADTV value >= 5B IDR OR AvgVol >= 5M shares
    # This prevents illiquid tickers from entering SMC, POI, VWAP, momentum,
    # and summary sections through filter-specific bypasses.
    def _good_rsi(r):
        """RSI >= 50 global gate — enforces momentum confirmation on all signal groups."""
        rv = safe_num(r.get("rsi14", r.get("rsi")), np.nan)
        return pd.notna(rv) and rv >= 50.0

    adtv_rows = [r for r in live_rows if _good_adtv(r) and _good_rsi(r)]

    def _by_score(pairs):
        return sorted(pairs, key=lambda x: safe_num(
            x[0].get("_sc_conviction", x[0].get("_sc_swing_score", 0)), 0
        ), reverse=True)

    # IMPORTANT: all filters use adtv_rows, not live_rows.
    # Gate first, then apply the filter logic.
    filter_a = _by_score([(r,"A") for r in adtv_rows if _above_ema25_50(r)])
    def _golden_cross_trigger(r):
        """S-05: Golden Cross fires only on MARKET_DATE. Returns trigger label or False."""
        # EMA50/SMA200 golden cross (ema50_sma200_cross contains "← Today" when fired today)
        ema_sma_cross = str(r.get("ema50_sma200_cross", "") or "")
        if "Today" in ema_sma_cross and "Golden" in ema_sma_cross:
            return "EMA×"
        # RSI14 × RSI MA14 golden cross (cross_status == "Golden" means fired today)
        if str(r.get("cross_status","")).strip() == "Golden":
            return "RSI×"
        # MACD Line × Signal golden cross
        if str(r.get("macd_cross","")).strip() == "Golden Cross":
            return "MACD×"
        return False

    def _b_with_trigger(r):
        trig = _golden_cross_trigger(r)
        if trig:
            r = dict(r)
            r["_gc_trigger"] = trig   # attach trigger label to row copy
            return r
        return None

    filter_b_raw = [_b_with_trigger(r) for r in adtv_rows]
    filter_b = _by_score([(r,"B") for r in filter_b_raw if r is not None])
    filter_c = _by_score([(r,"C") for r in adtv_rows if _is_bos_today(r, latest_market_day)])
    filter_d = _by_score([(r,"D") for r in adtv_rows if _filter_d_poi_touch(r)])
    filter_e = _by_score([(r,"E") for r in adtv_rows if _filter_e_eq_breakout(r)])
    filter_f = _by_score([(r,"F") for r in adtv_rows if _is_near_vwap_zone(r)])
    filter_g = _by_score([(r,"G") for r in adtv_rows if _is_smc_location_screener(r)])

    totals = {
        "A": len(filter_a), "B": len(filter_b), "C": len(filter_c),
        "D": len(filter_d), "E": len(filter_e), "F": len(filter_f),
        "G": len(filter_g),
    }
    grand_total = sum(totals.values())

    def _event_date_from_row(r: dict, *keys):
        for k in keys:
            v = r.get(k)
            if v not in (None, "", "-", "N/A"):
                try:
                    return pd.Timestamp(v).date()
                except Exception:
                    pass
        return None

    _mkt_date_obj = pd.Timestamp(latest_market_day).date()

    def _within_days(r: dict, days: int, *date_keys) -> bool:
        d = _event_date_from_row(r, *date_keys)
        if d is None:
            return False
        return 0 <= (_mkt_date_obj - d).days <= days

    # Compact section summaries requested in the DOCX.
    # A/D are market-date snapshot filters from today's OHLC/EMA/POI state.
    # B/C use dated events when available; otherwise their current signal count remains the market-date count.
    FILTER_SUMMARY = {
        "A": "Requirement: Close > EMA25 AND EMA25 > EMA50 AND RSI > 50",
        "B": "Requirement: EMA Golden Cross OR MACD line crossed above MACD signal",
        "C": f"Requirement: Swing OR Internal Bullish BOS exactly on market date {latest_market_day}",
        "D": "Requirement: Candle low touched POI (VWAP/EMA/Internal OB EQ) and close reclaimed above it",
        "E": "Requirement: Price broke out of VWAP equilibrium zone on market date",
        "F": "Requirement: Price is near Prev Q or Prev Y VWAP zone (Current Q excluded)",
        "G": "Requirement: SMC zone = 'In Bull OB' OR 'Equilibrium' OR 'Discount' (structurally favourable entry zone)",
    }

    # ── Row tracking ──────────────────────────────────────────────────────────
    row_idx = 6   # rows 1–5 written after data (they need totals for stats bar)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _score_bar(score):
        """Unicode block bar — 10 chars, filled proportionally to score/10."""
        try:
            s = min(10, max(0, float(score)))
            filled = round(s)
            return "█" * filled + "░" * (10 - filled)
        except Exception:
            return "░" * 10

    def _score_color(score):
        try:
            s = float(score)
            if s >= 8:   return "008000"   # green
            if s >= 6:   return "2E75B6"   # blue
            if s >= 4:   return "E36C09"   # amber
            return "595959"                # grey
        except Exception:
            return "595959"

    def _rvol_color(rv):
        try:
            v = float(rv)
            if v >= 3.0: return "C00000"   # red-hot
            if v >= 2.0: return "E36C09"   # amber
            if v >= 1.2: return "2E75B6"   # blue — above average
            return "595959"
        except Exception:
            return "595959"

    def _upside_color(up):
        try:
            v = float(up)
            if v >= 10: return ("008000", True)    # green bold
            if v >= 6:  return ("2E75B6", False)   # blue
            if v >= 3:  return ("E36C09", False)   # amber
            return ("C00000", False)               # red — poor
        except Exception:
            return ("595959", False)

    def _rr_color(rr_str):
        try:
            v = float(str(rr_str).replace("1:","").strip())
            if v >= 3: return "008000"
            if v >= 2: return "2E75B6"
            if v >= 1: return "E36C09"
            return "C00000"
        except Exception:
            return "595959"

    def _ara_badge(r):
        ara = str(r.get("ara_arb", "") or "").strip()
        if ara.startswith("ARA"): return " ★ARA"
        if ara.startswith("ARB"): return " ★ARB"
        return ""

    def _section_banner(label, dna, count, summary_text=""):
        """Full-width colored section header with right-aligned count badge and compact filter summary."""
        nonlocal row_idx
        hdr_fill_clr, _, _, emoji = dna
        ws.row_dimensions[row_idx].height = 26
        # Merge label columns: COL_START to COL_END-1; leave COL_END unmerged for badge
        if COL_END > COL_START:
            ws.merge_cells(start_row=row_idx, start_column=COL_START,
                           end_row=row_idx, end_column=COL_END - 1)
        c = ws.cell(row_idx, COL_START)
        c.value = f"  {emoji}  {label}" + (f"  |  {summary_text}" if summary_text else "")
        c.fill  = PatternFill("solid", fgColor=hdr_fill_clr)
        c.font  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.alignment = AL_L
        # Badge — COL_END is intentionally NOT part of the merge above
        badge_c = ws.cell(row_idx, COL_END)
        badge_c.value     = f"{count} signals"
        badge_c.fill      = PatternFill("solid", fgColor=hdr_fill_clr)
        badge_c.font      = Font(name="Calibri", size=9, bold=True,
                                 color="9CA3AF" if count == 0 else "FFFFFF")
        badge_c.alignment = AL_R
        badge_c.border    = BDR_CELL
        row_idx += 1

    def _empty_section_row(dna):
        """Italic grey 'no matches' row spanning all columns."""
        nonlocal row_idx
        _, row_fill_clr, _, _ = dna
        ws.row_dimensions[row_idx].height = 18
        ws.merge_cells(start_row=row_idx, start_column=COL_START,
                       end_row=row_idx, end_column=COL_END)
        c = ws.cell(row_idx, COL_START)
        c.value     = "  — no tickers matched this filter on the market date —"
        c.fill      = PatternFill("solid", fgColor=row_fill_clr)
        c.font      = FT_MUTED
        c.alignment = AL_L
        row_idx += 1

    def _spacer_row():
        nonlocal row_idx
        ws.row_dimensions[row_idx].height = 6
        for ci in range(COL_START, COL_END + 1):
            ws.cell(row_idx, ci).fill = FILL_EMPTY_ROW
        row_idx += 1

    def _data_row(r, tag, rank, dna):
        nonlocal row_idx
        _, row_fill_clr, accent_clr, _ = dna

        tp  = _institutional_trade_plan(r)
        ara = _ara_badge(r)

        # ── Fetch news (cached per ticker) ────────────────────────────────────
        ticker_str = str(r.get("ticker", "")).upper().strip()
        sent_data, corp_data = _get_news(ticker_str, r)

        aug = dict(r)
        aug["_rank"]           = rank
        aug["_signal_category"] = tag
        aug["_news_sentiment"] = sent_data.get("display", "")
        aug["_corp_action"]    = corp_data.get("display", "")
        # S-02+S-03: 10-column Trade Plan (rebuilt)
        aug["tp_entry_poi"]         = tp.get("tp_entry_poi", "-")
        aug["tp_entry"]             = tp.get("tp_entry", np.nan)
        aug["tp_entry_dist_pct"]    = tp.get("tp_entry_dist_pct", np.nan)
        aug["tp_target_poi"]        = tp.get("tp_target_poi", "-")
        aug["tp_target"]            = tp.get("tp_target", np.nan)
        aug["tp_target_upside_pct"] = tp.get("tp_target_upside_pct", np.nan)
        aug["tp_inv_poi"]           = tp.get("tp_inv_poi", "-")
        aug["tp_inv"]               = tp.get("tp_inv", np.nan)
        aug["tp_inv_downside_pct"]  = tp.get("tp_inv_downside_pct", np.nan)
        aug["tp_rr"]                = tp.get("tp_rr", np.nan)
        aug["summary_filter"]  = _filter_reason(r, tag)

        ws.row_dimensions[row_idx].height = 22

        accent_left = Border(
            left   = _med_side(accent_clr),
            right  = _thin_side(),
            top    = _thin_side(),
            bottom = _thin_side()
        )
        std_border = BDR_CELL
        row_fill   = PatternFill("solid", fgColor=row_fill_clr)

        for col_i, (key, hdr, width, fmt, align, grp) in enumerate(COLUMNS, start=COL_START):
            raw = aug.get(key)
            val = raw

            if isinstance(val, float) and pd.isna(val):
                val = ""
            if val in ("N/A", "None", None, ""):
                val = ""

            c = ws.cell(row_idx, col_i)
            c.fill      = row_fill
            c.border    = accent_left if col_i == COL_START else std_border
            c.alignment = Alignment(
                horizontal=align, vertical="center",
                wrap_text=(key in ("q_remarks", "pq_remarks", "py_remarks", "summary_ma", "tp_poi", "summary_filter",
                                   "_news_sentiment", "_corp_action"))
            )

            # ── Per-column formatting ─────────────────────────────────────
            if key == "_rank":
                c.value = rank
                c.font  = FT_RANK

            elif key == "ticker":
                c.value = f"{str(val)}{ara}" if val != "-" else "-"
                c.font  = Font(name="Calibri", size=10, bold=True,
                               color=accent_clr)

            elif key == "_sc_conviction":
                try:
                    sv = float(raw)
                    c.value = sv
                    c.font  = Font(name="Calibri", size=10, bold=True,
                                   color=_score_color(sv))
                    if fmt:
                        c.number_format = fmt
                except Exception:
                    c.value = "-"; c.font = FT_BODY

            elif key == "pct_change":
                try:
                    pv = float(raw)
                    c.value = pv
                    c.number_format = "0.00%"
                    c.font = Font(name="Calibri", size=10, bold=True,
                                  color="008000" if pv > 0 else
                                        ("C00000" if pv < 0 else "595959"))
                except Exception:
                    c.value = val; c.font = FT_BODY

            elif key == "rvol20":
                try:
                    rv = float(raw)
                    c.value = rv
                    c.number_format = "0.00"
                    c.font = Font(name="Calibri", size=10, bold=(rv >= 2.0),
                                  color=_rvol_color(rv))
                except Exception:
                    c.value = val; c.font = FT_BODY

            elif key == "tp_upside_pct":
                try:
                    uv = float(raw)
                    col, bold = _upside_color(uv * 100.0)
                    c.value = uv
                    if fmt: c.number_format = fmt
                    c.font = Font(name="Calibri", size=10, bold=bold, color=col)
                except Exception:
                    c.value = val; c.font = FT_BODY

            elif key == "tp_rr_display":
                c.value = val
                c.font  = Font(name="Calibri", size=10, bold=False,
                               color=_rr_color(val))
            elif key == "tp_rr_numeric":
                try:
                    rv = float(raw)
                    c.value = rv
                    if fmt: c.number_format = fmt
                    c.font = Font(name="Calibri", size=10, bold=(rv >= 2.0), color=_rr_color(f"1:{rv}"))
                except Exception:
                    c.value = ""; c.font = FT_BODY

            elif key == "adr_pct":
                try:
                    av = float(raw)
                    av_ratio = av / 100.0 if abs(av) > 1 else av
                    c.value = av_ratio
                    if fmt: c.number_format = fmt
                    c.font = Font(name="Calibri", size=10,
                                  color="C00000" if av_ratio >= 0.05 else
                                        ("E36C09" if av_ratio >= 0.03 else "1F2937"))
                except Exception:
                    c.value = val; c.font = FT_BODY

            elif key == "_news_sentiment":
                disp  = sent_data.get("display", "-") or "-"
                color = sent_data.get("color", "6B7280")
                url   = sent_data.get("url", "")
                c.value = disp
                c.font  = Font(name="Calibri", size=9, color=color)
                if isinstance(url, str) and url.startswith("http"):
                    c.hyperlink = url

            elif key == "_corp_action":
                disp  = corp_data.get("display", "-") or "-"
                color = corp_data.get("color", "6B7280")
                url   = corp_data.get("url", "")
                src   = corp_data.get("source", "")
                # Dim the font slightly if this came from IFNA or IDX-Scrape fallback
                italic = src in ("RSS-IFNA", "IDX-Scrape", "N/A")
                c.value = disp
                c.font  = Font(name="Calibri", size=9, italic=italic, color=color)
                if isinstance(url, str) and url.startswith("http"):
                    c.hyperlink = url

            else:
                c.value = val
                if fmt and val != "-" and isinstance(raw, (int, float)) and pd.notna(raw):
                    c.number_format = fmt
                c.font = FT_BODY_SMALL if key in ("q_remarks", "pq_remarks", "py_remarks", "summary_ma", "tp_poi", "summary_filter") else FT_BODY

        row_idx += 1

    # ── Per-run news cache — fetched once per ticker, shared across filter sections ──
    # Key: ticker str  →  {"sent": dict, "corp": dict}
    _NEWS_CACHE: dict = {}

    def _get_news(ticker: str, row: dict | None = None) -> tuple:
        """Returns (sentiment_dict, corp_action_dict), cached per ticker+company."""
        t = str(ticker).upper().strip().replace(".JK", "")
        company = ""
        if row:
            company = row.get("emiten", row.get("name", "")) or ""
        cache_key = (t, company)
        if cache_key not in _NEWS_CACHE:
            try:
                s = _fetch_news_sentiment_screener(t, latest_market_day, company)
            except Exception:
                s = {"display": "", "color": "9CA3AF", "url": "", "entity_match_status": "NO_NEWS"}
            try:
                ca = _fetch_corp_action_screener(t, latest_market_day, company)
            except Exception:
                ca = {"display": "", "color": "9CA3AF", "url": "", "entity_match_status": "NO_NEWS"}
            _NEWS_CACHE[cache_key] = {"sent": s, "corp": ca}
        return _NEWS_CACHE[cache_key]["sent"], _NEWS_CACHE[cache_key]["corp"]

    # ── Write all filter sections ─────────────────────────────────────────────
    SECTION_DEFS = [
        ("A", "EMA Trend ↑",  filter_a),
        ("B", "Golden Cross", filter_b),
        ("C", "BOS (Swing + Internal)", filter_c),
        ("D", "POI Reclaim", filter_d),
        ("E", "EQ Breakout", filter_e),
        ("F", "Near VWAP", filter_f),
        ("G", "SMC Location", filter_g),
    ]

    for fid, label, fpairs in SECTION_DEFS:
        dna = FILTER_DNA[fid]
        _section_banner(label, dna, len(fpairs), FILTER_SUMMARY.get(fid, ""))
        if fpairs:
            for rank_i, (r, tag) in enumerate(fpairs, start=1):
                _data_row(r, tag, rank_i, dna)
        else:
            _empty_section_row(dna)
        _spacer_row()

    # ── Rows 1–5: title block + column headers ────────────────────────────────
    try:
        _pretty = pd.Timestamp(latest_market_day).strftime("%B %d, %Y")
    except Exception:
        _pretty = str(latest_market_day)

    _bt_tag = "  [BACKTEST MODE]" if BACKTEST_MODE else ""

    # Row 1 — title bar
    ws.row_dimensions[1].height = 38
    ws.merge_cells(start_row=1, start_column=COL_START, end_row=1, end_column=COL_END)
    t1 = ws.cell(1, COL_START)
    t1.value     = f"  IDX Screener  ·  {_pretty}{_bt_tag}"
    t1.fill      = FILL_TITLE
    t1.font      = FT_TITLE
    t1.alignment = AL_L

    # Row 2 — stats bar
    ws.row_dimensions[2].height = 20
    ws.merge_cells(start_row=2, start_column=COL_START, end_row=2, end_column=COL_END)
    t2 = ws.cell(2, COL_START)
    t2.value = "  Filter Gates : (ADTR 20D (IDR) ≥ 5B OR ADTV 20D (Shares) ≥ 5M) AND RSI ≥ 50"
    t2.fill      = FILL_STATS
    t2.font      = FT_STATS
    t2.alignment = AL_L

    # Row 3 — spacer
    ws.row_dimensions[3].height = 4
    for ci in range(COL_START, COL_END + 1):
        ws.cell(3, ci).fill = FILL_TITLE

    # Row 4 — column group headers
    ws.row_dimensions[4].height = 18
    # Determine group spans
    grp_spans = {}
    for col_i, grp in enumerate(GROUPS, start=COL_START):
        if grp not in grp_spans:
            grp_spans[grp] = [col_i, col_i]
        else:
            grp_spans[grp][1] = col_i

    GRP_META = {
        "S": ("STOCK INFO",       FILL_GRP_STK),
        "P": ("SIGNAL & PRICE",   FILL_GRP_SIG),
        "N": ("NEWS INTELLIGENCE",PatternFill("solid", fgColor="6B2737")),
        "T": ("WORKBOOK LEVEL MAP", FILL_GRP_TRD),
        "Q": ("QUALITY METRICS",  PatternFill("solid", fgColor="374151")),
    }
    for grp_id, (grp_label, grp_fill) in GRP_META.items():
        if grp_id not in grp_spans:
            continue
        s, e = grp_spans[grp_id]
        if s < e:
            ws.merge_cells(start_row=4, start_column=s, end_row=4, end_column=e)
        c = ws.cell(4, s, grp_label)
        c.fill = grp_fill; c.font = FT_GRP
        c.alignment = Alignment(horizontal="center", vertical="center")  # Item 9: centered

    # Row 5 — sub-headers
    ws.row_dimensions[5].height = 22
    for col_i, (key, hdr, _, _, align, grp) in enumerate(COLUMNS, start=COL_START):
        c = ws.cell(5, col_i, hdr)
        grp_fill_map = {
            "S": FILL_GRP_STK,
            "P": FILL_GRP_SIG,
            "N": PatternFill("solid", fgColor="6B2737"),
            "T": FILL_GRP_TRD,
            "Q": PatternFill("solid", fgColor="374151"),
        }
        c.fill      = grp_fill_map[grp]
        c.font      = FT_SUBHDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = BDR_CELL

    # Hide machine-audit helper columns that should remain export-readable.
    try:
        for _ci, (_key, *_rest) in enumerate(COLUMNS, start=COL_START):
            if _key in ("tp_function", "tp_formula"):
                ws.column_dimensions[get_column_letter(_ci)].hidden = True
    except Exception:
        pass

    ws.freeze_panes = "E6"
    ws.auto_filter.ref = (
        f"{get_column_letter(COL_START)}5:"
        f"{get_column_letter(COL_END)}{max(5, row_idx - 1)}"
    )

    print(f"[INFO] IDX Screener: A={totals['A']}, B={totals['B']}, "
          f"C={totals['C']}, D={totals['D']}, E={totals['E']}, F={totals['F']}")


def build_fundamental_key_stats_sheet(wb, latest_market_day: str, rows: list):
    """
    IDX Fundamental Detail — columnar table matching IDX Technical Detail architecture.
    Uses FUNDAMENTAL_DETAIL_SCHEMA with identical grouped-header engine.
    One row per ticker; Stock Info group leads, then fundamental sections.
    """
    sheet_name = "IDX Fundamental Detail"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 0.5

    # ── Row heights (mirror IDX Technical Detail) ─────────────────────────────
    for r_h, h in {1: 5, 2: 24, 3: 18, 4: 5, 5: 24, 6: 44, 7: 20, 8: 20, 9: 20, 10: 5, 11: 24, 12: 48}.items():
        ws.row_dimensions[r_h].height = h

    # ── Title rows ────────────────────────────────────────────────────────────
    _asof_ts   = get_effective_asof_date()
    _pretty    = _asof_ts.strftime("%B %d, %Y")
    _mode_lbl  = f"BACKTEST MODE as of {_pretty}" if BACKTEST_MODE else f"Data as of {_pretty}"
    _bt_note   = " — Note: yfinance fundamentals reflect latest available, not replay-adjusted" if BACKTEST_MODE else ""

    ws["B2"] = "IDX Fundamental Detail" + (" [BACKTEST MODE]" if BACKTEST_MODE else "")
    style_plain(ws["B2"], font=FONT_TITLE, align="left")
    ws["B3"] = f"{_mode_lbl}  ·  yfinance  ·  IDX Universe"
    style_plain(ws["B3"], font=FONT_SUBTITLE, align="left")

    # Clear spacer row 4
    for _c in range(2, 250):
        ws.cell(4, _c).value = None

    # ── Flatten schema → flat_cols (identical engine to build_detail_sheet) ──
    flat_cols    = []
    col_idx      = 2
    group_ranges = []

    for group_name, group_fill, cols in FUNDAMENTAL_DETAIL_SCHEMA:
        start_idx = col_idx
        for key, header, width, fmt, align in cols:
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].width = width
            flat_cols.append((col_idx, letter, group_name, group_fill, key, header, width, fmt, align))
            col_idx += 1
        end_idx = col_idx - 1
        group_ranges.append((group_name, group_fill, start_idx, end_idx))

    # ── Data vintage labels per group ────────────────────────────────────────
    # Maps group name → short vintage descriptor shown in the group header bar.
    # Sourced from yfinance conventions for IDX (BEI) companies.
    _VINTAGE = {
        "Stock Info":                    "Live Price  ·  IDX",
        "Company Profile":               "Latest Available  ·  yfinance",
        "Valuation":                     "TTM  ·  yfinance",
        "Profitability":                 "TTM  ·  yfinance",
        "Growth":                        "TTM vs Prior Year  ·  yfinance",
        "Balance Sheet":                 "Latest Quarterly Report  ·  yfinance",
        "Cash Flow":                     "TTM  ·  yfinance",
        "Upcoming Dividend":             "Upcoming only  ·  yfinance",
        "Latest Dividend":               "Historical latest  ·  yfinance",
        "PBV Band Analysis  ·  3-Year Rolling": "yfinance",
    }

    # ── Group header row (row 5) ──────────────────────────────────────────────
    for group_name, group_fill, start_idx, end_idx in group_ranges:
        vintage   = _VINTAGE.get(group_name, "yfinance")
        label     = f"{group_name}  ·  {vintage}"
        start_cell = f"{get_column_letter(start_idx)}5"
        end_cell   = f"{get_column_letter(end_idx)}5"
        ws.merge_cells(f"{start_cell}:{end_cell}")
        for c in range(start_idx, end_idx + 1):
            cell = ws.cell(5, c)
            cell.fill = group_fill
            cell.font = FONT_GROUP
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
        ws[start_cell] = label

    # ── Subheader row (row 6) ─────────────────────────────────────────────────
    for col_idx, letter, group_name, group_fill, key, header, width, fmt, align in flat_cols:
        c = ws.cell(6, col_idx)
        c.value = header
        style_cell(c, fill=group_fill, font=FONT_SUBHEADER, align="center", wrap=True)

    # ── Data rows ─────────────────────────────────────────────────────────────
    # "NO DATA" here means insufficient OHLCV bars for technical indicators
    # (base_row_from_ksei), not "no company fundamentals" -- it still carries
    # a real ticker, so it's still worth a yfinance fundamentals attempt below.
    # Only genuinely tickerless/malformed rows are dropped.
    ok_rows   = [r for r in rows if r.get("ticker") and r.get("data_status") in ("OK", "PARTIAL DATA", "NO DATA")]
    _COMPACT_KEYS = {
        "yf_shares_outstanding", "yf_float_shares", "yf_market_cap_current",
        "yf_enterprise_value_current", "yf_shares_outstanding_current",
        "yf_free_cash_flow_q", "yf_book_value_q", "yf_tangible_book_value_q",
        "yf_short_term_debt_q", "yf_long_term_debt_q", "yf_cash_q",
        "yf_total_assets_q", "yf_total_liabilities_q", "yf_revenue_ttm",
        "yf_net_income_ttm", "yf_ebit_ttm", "yf_cash_from_operations_ttm",
        "yf_cash_from_investing_ttm", "yf_cash_from_financing_ttm",
        "yf_capital_expenditure_ttm", "yf_free_cash_flow_ttm",
        "div1_idr", "div2_idr",
    }
    PBV_REGIME_COLORS = {
        "Deep Undervalued": "1D4ED8",
        "Undervalued":      "16A34A",
        "Fair Value":       "737373",
        "Overvalued":       "CA8A04",
        "Bubble":           "DC2626",
    }

    for r_idx, screener_row in enumerate(ok_rows, start=7):
        combined = _build_fundamental_row(screener_row)

        for col_idx, letter, group_name, group_fill, key, header, width, fmt, align in flat_cols:
            c   = ws.cell(r_idx, col_idx)
            val = combined.get(key, np.nan)

            # ── BUG-03: IDR-aware value formatting ────────────────────────────
            # IDR monetary values: post-conversion they are in IDR → display as B IDR (÷1e9)
            _IDR_MON = {
                "yf_market_cap_current", "yf_enterprise_value_current",
                "yf_free_cash_flow_q",   "yf_book_value_q",
                "yf_tangible_book_value_q", "yf_short_term_debt_q",
                "yf_long_term_debt_q",   "yf_cash_q",
                "yf_total_assets_q",     "yf_total_liabilities_q",
                "yf_revenue_ttm",        "yf_net_income_ttm",
                "yf_ebit_ttm",           "yf_cash_from_operations_ttm",
                "yf_cash_from_investing_ttm", "yf_cash_from_financing_ttm",
                "yf_capital_expenditure_ttm", "yf_free_cash_flow_ttm",
            }
            # IDR per-share values (already converted to IDR in _compute_extended_fundamentals)
            _IDR_PER_SHARE = {
                "yf_eps_ttm_current",
                "yf_book_value_per_share_current",
                "yf_tangible_book_value_per_share_current",
            }
            _SHARES_KEYS = {"yf_shares_outstanding", "yf_float_shares", "yf_shares_outstanding_current"}
            _DIV_IDR_KEYS = {"div1_idr", "div2_idr"}

            # Resolve val → formatted string
            if val == "N/A" or (isinstance(val, float) and pd.isna(val)):
                val = "-"
            elif key in _IDR_MON:
                val = idr_billions_fmt(val) or "-"
            elif key in _IDR_PER_SHARE:
                try:
                    val = f"{float(val):,.2f}"
                except Exception:
                    val = "-"
            elif key in _DIV_IDR_KEYS:
                try:
                    val = f"{float(val):,.2f}"
                except Exception:
                    val = "-"
            elif key in _SHARES_KEYS:
                val = compact_fmt(val) or "-"
            elif key in _COMPACT_KEYS:
                val = compact_fmt(val) or "-"

            c.value = val
            style_cell(c, fill=None, font=FONT_BODY, align=align,
                       wrap=(key in ("yf_company_name", "yf_industry", "yf_business_summary", "yf_disclosure_title")))

            # Disclosure URL → clickable hyperlink
            if key == "yf_disclosure_url" and isinstance(val, str) and val.startswith("http"):
                c.hyperlink = val
                c.value = "View →"
                c.font = Font(name="Calibri", size=10, color="0563C1", underline="single")

            # % Change colour (green / red)
            if key == "pct_change":
                try:
                    _v = safe_num(combined.get("pct_change"), np.nan)
                    if pd.notna(_v):
                        c.number_format = "0.00%"
                        c.font = Font(name=FONT_BODY.name, size=FONT_BODY.size, bold=False,
                                      color="008000" if _v > 0 else ("C00000" if _v < 0 else "1F1F1F"))
                except Exception:
                    pass

            # PBV Regime colour
            if key == "yf_pbv_regime" and val not in ("-", "", None):
                rc = PBV_REGIME_COLORS.get(str(val), "1F1F1F")
                c.font = Font(name="Calibri", size=10, bold=True, color=rc)

            # Disclosure Sentiment colour
            if key == "yf_disclosure_sentiment":
                sent_colors = {"Bullish": "008000", "Bearish": "C00000", "Neutral": "595959"}
                rc = sent_colors.get(str(val), "1F1F1F")
                c.font = Font(name="Calibri", size=10, bold=True, color=rc)

            # Event Risk colour
            if key == "yf_event_risk":
                risk_colors = {"High": "C00000", "Medium": "E36C09", "Low": "008000"}
                rc = risk_colors.get(str(val), "1F1F1F")
                c.font = Font(name="Calibri", size=10, bold=True, color=rc)

            # Number format (only for true numeric cells)
            if fmt and val != "-" and fmt not in ("compact",):
                if isinstance(combined.get(key), (int, float)) and pd.notna(combined.get(key)):
                    c.number_format = fmt

        ws.row_dimensions[r_idx].height = 26

    # ── Freeze panes + autofilter (mirror IDX Technical Detail) ──────────────
    last_col_letter = get_column_letter(flat_cols[-1][0]) if flat_cols else "Z"
    ws.freeze_panes = "E7"
    ws.auto_filter.ref = f"B6:{last_col_letter}{max(6, ws.max_row)}"

    # ── PBV Band DataBar conditional formatting ───────────────────────────────
    # Apply horizontal data bars to the six PBV band numeric columns so the
    # current/mean/SD spread is visually comparable across rows (Stockbit-style).
    try:
        from openpyxl.formatting.rule import DataBarRule
        _PBV_BAR_KEYS = {
            "yf_pbv_curr", "yf_pbv_mean",
            "yf_pbv_plus2", "yf_pbv_plus1",
            "yf_pbv_minus1", "yf_pbv_minus2",
        }
        _data_end_row = max(7, 6 + len(ok_rows))
        for _ci, _lt, _gn, _gf, _key, _hdr, _w, _fmt, _aln in flat_cols:
            if _key in _PBV_BAR_KEYS:
                _col_range = f"{_lt}7:{_lt}{_data_end_row}"
                ws.conditional_formatting.add(
                    _col_range,
                    DataBarRule(
                        start_type="min", start_value=None,
                        end_type="max", end_value=None,
                        color="638EC6",
                    )
                )
    except Exception:
        pass



# =========================
# BACKTEST ENGINE  (Phase D — bar-by-bar historical replay)
# =========================
def _run_backtest_replay(rows_cache: dict) -> dict:
    """
    Bar-by-bar replay over BACKTEST_TICKERS.
    rows_cache: {ticker: pd.DataFrame} with full OHLCV history.
    Returns a dict of result DataFrames for each backtest output sheet.
    """
    import warnings
    warnings.filterwarnings("ignore")

    results_by_date = {}   # date -> list of scored rows
    regime_log     = []    # (date, ticker, regime)
    signal_log     = []    # (date, ticker, signal, score, macd_state)

    if not rows_cache:
        return {}

    # Build unified date range from available data
    all_dates = set()
    for df in rows_cache.values():
        if df is not None and not df.empty:
            all_dates.update(df.index.normalize().unique())
    if not all_dates:
        return {}

    try:
        start_dt = pd.Timestamp("2020-01-01")   # fixed historical start for bar-by-bar replay
        end_dt   = get_market_date()             # MARKET_DATE is the replay boundary
    except Exception:
        start_dt = pd.Timestamp("2020-01-01")
        end_dt   = get_market_date()  # MARKET_DATE cutoff — not today

    date_range = sorted(d for d in all_dates if start_dt <= d <= end_dt)
    if not date_range:
        return {}

    # Sample every 5th trading day for speed (configurable)
    STEP = 5
    sampled_dates = date_range[::STEP]

    for snap_date in sampled_dates:
        day_rows = []
        for ticker, full_hist in rows_cache.items():
            if full_hist is None or full_hist.empty:
                continue
            # Strict as-of: only use bars <= snap_date
            hist = full_hist[full_hist.index.normalize() <= snap_date].copy()
            if len(hist) < MIN_BARS_PARTIAL:
                continue
            try:
                close = safe_num(hist["Close"].iloc[-1])
                if pd.isna(close) or close <= 0:
                    continue

                # MACD
                macd_data = compute_macd_momentum(hist) if len(hist) >= 35 else {}
                # SMC
                smc_data  = compute_smc_engine(hist) if len(hist) >= 60 else {}
                # VWAP
                cq_start = quarter_start(snap_date)
                df_q = hist.loc[hist.index >= cq_start]
                q = anchored_vwap_block(df_q) if not df_q.empty else {}

                # Score (simple composite for backtest)
                score = 0.0
                swing = smc_data.get("smc_swing_trend", "Neutral")
                macd_entry = macd_data.get("macd_entry", "")
                if swing == "Bullish":          score += 3.0
                if "ENTRY SIGNAL" in macd_entry: score += 3.0
                if "Continuation" in macd_entry: score += 2.0
                pd_zone = smc_data.get("smc_premium_discount", "")
                if pd_zone == "Discount":       score += 2.0
                if pd_zone == "Equilibrium":    score += 1.0
                if pd_zone == "Premium":        score -= 1.0

                composite = smc_data.get("smc_composite_state", "Neutral Rotation")
                regime_log.append({
                    "date": snap_date.strftime("%Y-%m-%d"),
                    "ticker": ticker,
                    "regime": composite,
                    "swing_trend": swing,
                    "macd_regime": macd_data.get("macd_regime", ""),
                    "pd_zone": pd_zone,
                    "score": round(score, 2),
                })

                if "ENTRY SIGNAL" in macd_entry or score >= 5.0:
                    signal_log.append({
                        "date": snap_date.strftime("%Y-%m-%d"),
                        "ticker": ticker,
                        "signal": macd_entry,
                        "score": round(score, 2),
                        "smc_state": composite,
                        "pd_zone": pd_zone,
                        "close": close,
                        "qvwap": q.get("vwap", np.nan) if q else np.nan,
                    })

                day_rows.append({
                    "ticker": ticker,
                    "close": close,
                    "score": round(score, 2),
                    "composite": composite,
                    "swing": swing,
                    "macd": macd_entry,
                })
            except Exception:
                continue

        if day_rows:
            day_rows.sort(key=lambda x: x["score"], reverse=True)
            results_by_date[snap_date.strftime("%Y-%m-%d")] = day_rows

    # ── Build result DataFrames ──────────────────────────────────────────────
    df_regime  = pd.DataFrame(regime_log)  if regime_log  else pd.DataFrame()
    df_signal  = pd.DataFrame(signal_log)  if signal_log  else pd.DataFrame()

    # Historical Rankings: top 5 per snapshot date
    ranking_rows = []
    for snap_date_str, rows_day in results_by_date.items():
        for rank_pos, r in enumerate(rows_day[:5], 1):
            ranking_rows.append({"date": snap_date_str, "rank": rank_pos, **r})
    df_rankings = pd.DataFrame(ranking_rows) if ranking_rows else pd.DataFrame()

    # Hit Rate Analytics: signals that were followed by 5%+ return within 10 days
    hit_rows = []
    for sig in signal_log:
        ticker = sig["ticker"]
        sig_date = pd.Timestamp(sig["date"])
        full_hist = rows_cache.get(ticker)
        if full_hist is None or full_hist.empty:
            continue
        entry_close = sig["close"]
        future = full_hist[full_hist.index.normalize() > sig_date].head(10)
        if future.empty or pd.isna(entry_close) or entry_close <= 0:
            continue
        max_ret = float((future["Close"].max() - entry_close) / entry_close * 100)
        hit = max_ret >= 5.0
        hit_rows.append({
            "date": sig["date"], "ticker": ticker,
            "signal": sig["signal"], "score": sig["score"],
            "entry_close": entry_close, "max_return_10d_pct": round(max_ret, 2),
            "hit": "YES" if hit else "NO",
        })
    df_hits = pd.DataFrame(hit_rows) if hit_rows else pd.DataFrame()

    # Equity Curve: daily portfolio value (equal-weight top-5 per snapshot)
    equity_rows = []
    portfolio_val = 100.0   # start at index 100
    prev_closes = {}
    for snap_date_str in sorted(results_by_date.keys()):
        top5 = results_by_date[snap_date_str][:5]
        if top5 and prev_closes:
            total_ret = 0.0
            counted = 0
            for r in top5:
                prev = prev_closes.get(r["ticker"])
                if prev and prev > 0:
                    total_ret += (r["close"] - prev) / prev
                    counted += 1
            if counted > 0:
                portfolio_val *= (1 + total_ret / counted)
        equity_rows.append({"date": snap_date_str, "portfolio_value": round(portfolio_val, 4)})
        for r in top5:
            prev_closes[r["ticker"]] = r["close"]
    df_equity = pd.DataFrame(equity_rows) if equity_rows else pd.DataFrame()

    # Summary stats
    total_signals = len(signal_log)
    total_hits    = df_hits["hit"].eq("YES").sum() if not df_hits.empty and "hit" in df_hits.columns else 0
    win_rate      = total_hits / total_signals * 100 if total_signals > 0 else 0.0
    avg_return    = float(df_hits["max_return_10d_pct"].mean()) if not df_hits.empty and "max_return_10d_pct" in df_hits.columns else 0.0

    eq_vals = df_equity["portfolio_value"].values if not df_equity.empty and "portfolio_value" in df_equity.columns else [100.0]
    peak = 100.0
    max_dd = 0.0
    for v in eq_vals:
        if v > peak:
            peak = v
        dd = (peak - v) / peak * 100 if peak > 0 else 0
        if dd > max_dd:
            max_dd = dd

    expectancy = win_rate/100 * avg_return - (1 - win_rate/100) * abs(avg_return * 0.5) if total_signals > 0 else 0.0

    # Sharpe proxy (annualised returns / std)
    if not df_equity.empty and len(df_equity) > 2 and "portfolio_value" in df_equity.columns:
        returns_ser = df_equity["portfolio_value"].pct_change().dropna()
        ann_ret = float(returns_ser.mean() * 252)
        ann_std = float(returns_ser.std() * (252**0.5))
        sharpe_proxy = ann_ret / ann_std if ann_std > 0 else 0.0
    else:
        ann_ret = 0.0; ann_std = 0.0; sharpe_proxy = 0.0

    # Signal frequency
    if sampled_dates:
        year_span = (sampled_dates[-1] - sampled_dates[0]).days / 365.25
        sig_freq = total_signals / year_span if year_span > 0 else 0.0
    else:
        sig_freq = 0.0

    df_summary = pd.DataFrame([{
        "Metric": "Total Signals",          "Value": total_signals},
        {"Metric": "Total Hits (≥5% / 10D)","Value": total_hits},
        {"Metric": "Win Rate %",             "Value": round(win_rate, 2)},
        {"Metric": "Avg Return (10D max) %", "Value": round(avg_return, 2)},
        {"Metric": "Max Drawdown %",         "Value": round(max_dd, 2)},
        {"Metric": "Expectancy %",           "Value": round(expectancy, 2)},
        {"Metric": "Sharpe Proxy",           "Value": round(sharpe_proxy, 2)},
        {"Metric": "Signal Frequency / Yr",  "Value": round(sig_freq, 1)},
        {"Metric": "Annualised Return %",    "Value": round(ann_ret * 100, 2)},
        {"Metric": "Snapshots Evaluated",    "Value": len(sampled_dates)},
    ])

    return {
        "summary":  df_summary,
        "rankings": df_rankings,
        "signals":  df_signal,
        "regime":   df_regime,
        "equity":   df_equity,
        "hits":     df_hits,
    }


def _write_df_to_ws(ws, df: pd.DataFrame, title: str, col_start: int = 2):
    """Generic helper: write a DataFrame to a worksheet with header row."""
    if df is None or df.empty:
        ws.cell(2, col_start).value = f"{title} — No data available"
        style_plain(ws.cell(2, col_start), font=FONT_SUBTITLE, align="left")
        return
    # Title row
    ws.merge_cells(start_row=2, start_column=col_start,
                   end_row=2,  end_column=col_start + len(df.columns) - 1)
    tc = ws.cell(2, col_start)
    tc.value = title
    style_plain(tc, font=FONT_TITLE, align="left")
    # Header
    for ci, col in enumerate(df.columns, start=col_start):
        hc = ws.cell(3, ci)
        hc.value = str(col).replace("_", " ").title()
        style_cell(hc, fill=FILL_HEADER, font=FONT_HEADER, align="center")
        ws.column_dimensions[get_column_letter(ci)].width = max(14, len(str(col)) + 4)
    # Data rows
    for ri, (_, row) in enumerate(df.iterrows(), start=4):
        row_fill = PatternFill("solid", fgColor="F8FAFC" if ri % 2 == 0 else "FFFFFF")
        for ci, val in enumerate(row.values, start=col_start):
            dc = ws.cell(ri, ci)
            raw = val
            if isinstance(val, float) and np.isnan(val):
                raw = "N/A"
            dc.value = raw
            style_cell(dc, fill=row_fill, font=FONT_BODY, align="center")


def build_backtest_sheets(wb, bt_results: dict):
    """
    Build 6 backtest output sheets from bt_results dict.
    Skips gracefully if any result df is missing.
    """
    SHEET_DEFS = [
        ("BT Summary",          "summary",  "Backtest Summary — Performance Metrics"),
        ("BT Historical Ranks", "rankings", "Historical Rankings — Top 5 by Snapshot Date"),
        ("BT Signal Replay",    "signals",  "Signal Replay — All Entry Signals Generated"),
        ("BT Regime Log",       "regime",   "Regime Transition Log — SMC State by Date"),
        ("BT Equity Curve",     "equity",   "Equity Curve — Portfolio Index (Base 100)"),
        ("BT Hit Rate",         "hits",     "Hit Rate Analytics — Signal Outcomes"),
    ]
    for sheet_name, key, title in SHEET_DEFS:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 0.5
        ws.row_dimensions[1].height = 4
        df = bt_results.get(key, pd.DataFrame())
        _write_df_to_ws(ws, df, title, col_start=2)


# =========================
# KONGLO GROUP ANALYSIS SHEET
# =========================
def build_konglo_analysis_sheet(wb, latest_market_day: str, all_market_rows: list):
    """
    IDX Konglo Group Analysis — standalone Excel sheet.

    Sections:
      1. Konglo Group Summary table: group name, ticker count, market cap (IDR T),
         top-3 sectors, cross-affiliation flags.
      2. Per-group ticker roster: ticker, sector, market cap, cross-membership.
      3. Fear & Greed snapshot panel (same computation as IDX Overview).
      4. Sector Exposure matrix: for each group, % of member market cap per IDX sector.
    """
    import numpy as np
    import pandas as pd
    import yfinance as yf
    import time
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import DataBarRule, ColorScaleRule

    SHEET_NAME = "IDX Konglo Analysis"
    if SHEET_NAME in wb.sheetnames:
        wb.remove(wb[SHEET_NAME])
    ws = wb.create_sheet(SHEET_NAME)
    ws.sheet_view.showGridLines = False

    # ── Styles ────────────────────────────────────────────────────────────────
    thin  = Side(style="thin",   color="D1D5DB")
    thick = Side(style="medium", color="9CA3AF")
    bdr_full  = Border(left=thin, right=thin, top=thin, bottom=thin)
    bdr_thick = Border(left=thick, right=thick, top=thick, bottom=thick)

    def _fill(hex_col):
        return PatternFill("solid", fgColor=hex_col)

    def _font(bold=False, sz=10, color="111827", italic=False):
        return Font(name="Calibri", size=sz, bold=bold, color=f"FF{color}", italic=italic)

    def _aln(h="center", wrap=True):
        return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

    def _cell(row, col, val, bg=None, bold=False, sz=10, fc="111827",
               h="center", wrap=True, italic=False, fmt=None, bdr=True):
        c = ws.cell(row, col, val)
        if bg:
            c.fill = _fill(bg)
        c.font  = _font(bold=bold, sz=sz, color=fc, italic=italic)
        c.alignment = _aln(h=h, wrap=wrap)
        c.border = bdr_full if bdr else Border()
        if fmt:
            c.number_format = fmt
        return c

    # ── Column widths ─────────────────────────────────────────────────────────
    col_w = {
        "A": 2,   "B": 28,  "C": 9,   "D": 16,  "E": 16,
        "F": 16,  "G": 16,  "H": 20,  "I": 20,  "J": 20,
        "K": 20,  "L": 20,  "M": 20,  "N": 20,
    }
    for col_ltr, w in col_w.items():
        ws.column_dimensions[col_ltr].width = w

    # ── Helpers ───────────────────────────────────────────────────────────────
    asof = pd.Timestamp(latest_market_day).normalize()
    try:
        pretty_date = asof.strftime("%B %d, %Y")
    except Exception:
        pretty_date = str(latest_market_day)

    # Build a ticker → row lookup from all_market_rows for fast sector/mcap access
    row_by_ticker = {}
    for r in (all_market_rows or []):
        t = str(r.get("ticker", "") or r.get("Ticker", "")).upper().strip()
        if t:
            row_by_ticker[t] = r

    IDX_SECTORS_LIST = [
        "IDXFINANCE","IDXTRANS","IDXNONCYC","IDXCYCLIC","IDXINFRA",
        "IDXBASIC","IDXTECHNO","IDXPROPERT","IDXINDUST","IDXENERGY","IDXHEALTH",
    ]
    SECTOR_SHORT = {
        "IDXFINANCE":"Finance", "IDXTRANS":"Transport", "IDXNONCYC":"NonCyc",
        "IDXCYCLIC":"Cyclic", "IDXINFRA":"Infra", "IDXBASIC":"Basic",
        "IDXTECHNO":"Tech", "IDXPROPERT":"Property", "IDXINDUST":"Industrial",
        "IDXENERGY":"Energy", "IDXHEALTH":"Health",
    }

    KONGLO_MAP = {
        "Agung Sedayu Group":     ["PANI","CBDK","ERAA","ERAL","PDPP","INPC","JIHD"],
        "Bakrie Group":           ["BUMI","BRMS","VKTR","BNBR","ENRG","DEWA","MDIA","VIVA","UNSP","BTEL","ELTY"],
        "Prajogo Pangestu Group": ["BRPT","BREN","TPIA","CUAN","PTRO","CDIA","RATU","SSIA"],
        "Emtek Group":            ["EMTK","SCMA","BUKA","ANJT","CASS"],
        "Happy Hapsoro Group":    ["RAJA","RATU","BUVA","ARCI","PSKT","PADI","MINA","SINI","PTRO","CBRE"],
        "Salim Group":            ["INDF","ICBP","DNET","AMMN","BINA","LSIP","SIMP","IMAS","IMJS","DCII",
                                   "PANI","MEDC","EMTK","MEGA","BUMI","BRMS","DEWA","UNIC","FAST","JECC"],
        "Sinarmas Group":         ["DSSA","GEMS","SMMA","BSDE","INKP","TKIM","EXCL","SMAR","LIFE",
                                   "BSIM","DUTI","SMDM","PYFA","DMAS","PLIN"],
        "Tanoko Group":           ["RISE","AVIA","CAKK","PEVE","CLEO","DEPO","ABMM","MERI","BLES","ZONE"],
        "TNT-Saratoga-Adaro":     ["ADRO","AADI","ADMR","MDKA","MBMA","EMAS","ESSA","SRTG","TBIG",
                                   "BFIN","TRIM","MPMX","WOMF","GOTO","PALM","GHON","GOLD"],
        "Triputra-Adaro Group":   ["TAPG","ADRO","ASSA","ESSA","KMTR","DAYA","ASLC"],
    }

    # Cross-affiliation map: ticker → list of groups it belongs to
    ticker_groups = {}
    for grp, tickers in KONGLO_MAP.items():
        for t in tickers:
            ticker_groups.setdefault(t, []).append(grp)

    # ── Fetch market cap for a ticker (screener row first, yfinance fallback) ─
    _mcap_cache = {}
    def _get_mcap(ticker):
        if ticker in _mcap_cache:
            return _mcap_cache[ticker]
        # Try screener row first
        row = row_by_ticker.get(ticker, {})
        for k in ["market_cap","marketCap","MarketCap","yf_market_cap_current","mcap"]:
            v = row.get(k)
            try:
                fv = float(v)
                if np.isfinite(fv) and fv > 0:
                    _mcap_cache[ticker] = fv
                    return fv
            except Exception:
                pass
        # yfinance fallback (best-effort, skip if slow)
        try:
            info = yf.Ticker(f"{ticker}.JK").info
            for k in ["marketCap","enterpriseValue"]:
                v = info.get(k)
                if v and float(v) > 0:
                    _mcap_cache[ticker] = float(v)
                    return float(v)
        except Exception:
            pass
        _mcap_cache[ticker] = np.nan
        return np.nan

    def _get_sector(ticker):
        row = row_by_ticker.get(ticker, {})
        for k in ["idx_sector","IDX Sector","IDXSector","sector","Sector"]:
            v = row.get(k, "")
            if v and str(v).strip().upper() in IDX_SECTORS_LIST:
                return str(v).strip().upper()
        return "—"

    def _fmt_idr(v):
        """Format IDR value in trillions/billions."""
        if pd.isna(v) or not np.isfinite(v):
            return "—"
        t = v / 1e12
        if t >= 1:
            return f"Rp {t:,.1f}T"
        b = v / 1e9
        if b >= 1:
            return f"Rp {b:,.0f}B"
        return f"Rp {v:,.0f}"

    # ── Title block ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 5
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 5

    _cell(2, 2, "IDX Konglo Group Analysis", bg=None, bold=True, sz=14, fc="111827", h="left", bdr=False)
    _cell(3, 2, f"Conglomerate ownership & sector exposure  ·  Data as of {pretty_date}  ·  IDX Universe",
          bold=False, sz=10, fc="6B7280", h="left", italic=True, bdr=False)

    # ── Section 1: Group Summary table ───────────────────────────────────────
    HEADERS_SUMMARY = [
        "Konglo Group", "Tickers", "Est. Group MCap (IDR)",
        "Primary Sector", "Secondary Sector", "Cross-Affiliation Tickers",
    ]
    HDR_BG   = "1E3A5F"
    HDR_ALTS = ["F0F4F8", "FFFFFF"]

    current_row = 5
    ws.row_dimensions[current_row].height = 20

    # Merged group header
    ws.merge_cells(f"B{current_row}:G{current_row}")
    c = ws.cell(current_row, 2, "▌ Konglo Group Summary")
    c.fill = _fill("0F2A45"); c.font = _font(bold=True, sz=11, color="FFFFFF")
    c.alignment = _aln(h="left"); c.border = bdr_full
    for col in range(3, 8):
        ws.cell(current_row, col).fill = _fill("0F2A45")
        ws.cell(current_row, col).border = bdr_full
    current_row += 1

    # Column headers
    ws.row_dimensions[current_row].height = 22
    for ci, hdr in enumerate(HEADERS_SUMMARY, start=2):
        _cell(current_row, ci, hdr, bg=HDR_BG, bold=True, sz=10, fc="FFFFFF", h="center")
    current_row += 1

    summary_start_row = current_row
    for gi, (group, tickers) in enumerate(KONGLO_MAP.items()):
        ws.row_dimensions[current_row].height = 20
        bg = HDR_ALTS[gi % 2]

        # Compute group stats
        mcaps = {t: _get_mcap(t) for t in tickers}
        total_mcap = sum(v for v in mcaps.values() if np.isfinite(v))
        sectors = [_get_sector(t) for t in tickers if _get_sector(t) != "—"]
        sector_counts = pd.Series(sectors).value_counts()
        primary   = sector_counts.index[0] if len(sector_counts) > 0 else "—"
        secondary = sector_counts.index[1] if len(sector_counts) > 1 else "—"
        cross = [t for t in tickers if len(ticker_groups.get(t, [])) > 1]

        _cell(current_row, 2, group,               bg=bg, bold=True, sz=10, fc="1E3A5F", h="left")
        _cell(current_row, 3, len(tickers),        bg=bg, bold=False, h="center", fmt="0")
        _cell(current_row, 4, _fmt_idr(total_mcap),bg=bg, bold=False, h="right")
        _cell(current_row, 5, SECTOR_SHORT.get(primary, primary),   bg=bg, h="center")
        _cell(current_row, 6, SECTOR_SHORT.get(secondary, secondary),bg=bg, h="center")
        _cell(current_row, 7, ", ".join(cross) if cross else "—",   bg=bg, h="left", sz=9)
        current_row += 1

    summary_end_row = current_row - 1
    current_row += 2  # spacer

    # ── Section 2: Per-group ticker detail ───────────────────────────────────
    TICKER_COLS = ["Ticker", "IDX Sector", "Est. MCap (IDR)", "Cross-Group Membership"]
    GROUP_FILLS = [
        "1E3A5F","2D5282","1A5276","215A6B","1D4E3D",
        "3D2B6B","6B3320","3B5323","1A3C5E","2C3E50",
    ]

    for gi, (group, tickers) in enumerate(KONGLO_MAP.items()):
        ws.row_dimensions[current_row].height = 5
        current_row += 1

        # Group header bar
        ws.merge_cells(f"B{current_row}:E{current_row}")
        grp_bg = GROUP_FILLS[gi % len(GROUP_FILLS)]
        c = ws.cell(current_row, 2, f"▌ {group}  ({len(tickers)} tickers)")
        c.fill = _fill(grp_bg); c.font = _font(bold=True, sz=10, color="FFFFFF")
        c.alignment = _aln(h="left"); c.border = bdr_full
        for col in range(3, 6):
            ws.cell(current_row, col).fill = _fill(grp_bg)
            ws.cell(current_row, col).border = bdr_full
        ws.row_dimensions[current_row].height = 20
        current_row += 1

        # Column headers
        ws.row_dimensions[current_row].height = 18
        for ci, hdr in enumerate(TICKER_COLS, start=2):
            _cell(current_row, ci, hdr, bg=HDR_BG, bold=True, sz=9, fc="FFFFFF", h="center")
        current_row += 1

        alts = ["F7F9FC", "FFFFFF"]
        for ti, ticker in enumerate(tickers):
            ws.row_dimensions[current_row].height = 18
            bg = alts[ti % 2]
            sector = _get_sector(ticker)
            mcap   = _get_mcap(ticker)
            other_groups = [g for g in ticker_groups.get(ticker, []) if g != group]

            _cell(current_row, 2, ticker,           bg=bg, bold=True, sz=10, fc="1E3A5F", h="center", fmt="@")
            _cell(current_row, 3, SECTOR_SHORT.get(sector, sector), bg=bg, h="center", sz=9)
            _cell(current_row, 4, _fmt_idr(mcap),   bg=bg, h="right", sz=9)
            cross_txt = "; ".join(other_groups) if other_groups else "—"
            cross_fc  = "7C2D12" if other_groups else "9CA3AF"
            _cell(current_row, 5, cross_txt, bg=bg, h="left", sz=9, fc=cross_fc)
            current_row += 1

    current_row += 2  # spacer

    # ── Section 3: Sector Exposure Matrix ────────────────────────────────────
    ws.row_dimensions[current_row].height = 5
    current_row += 1

    # Section header
    ws.merge_cells(f"B{current_row}:N{current_row}")
    c = ws.cell(current_row, 2, "▌ Sector Exposure Matrix  (% of group market cap by IDX sector)")
    c.fill = _fill("0F2A45"); c.font = _font(bold=True, sz=11, color="FFFFFF")
    c.alignment = _aln(h="left"); c.border = bdr_full
    for col in range(3, 15):
        ws.cell(current_row, col).fill = _fill("0F2A45")
        ws.cell(current_row, col).border = bdr_full
    ws.row_dimensions[current_row].height = 20
    current_row += 1

    # Column headers: Group | Sector1 | ... | Sector11 | Total MCap
    matrix_hdr_row = current_row
    ws.row_dimensions[current_row].height = 22
    _cell(current_row, 2, "Konglo Group", bg=HDR_BG, bold=True, sz=10, fc="FFFFFF")
    for si, sector in enumerate(IDX_SECTORS_LIST, start=3):
        _cell(current_row, si, SECTOR_SHORT.get(sector, sector), bg=HDR_BG, bold=True, sz=9, fc="FFFFFF")
    _cell(current_row, 3 + len(IDX_SECTORS_LIST), "Total MCap", bg=HDR_BG, bold=True, sz=9, fc="FFFFFF")
    current_row += 1

    matrix_data_start = current_row
    for gi, (group, tickers) in enumerate(KONGLO_MAP.items()):
        ws.row_dimensions[current_row].height = 18
        bg = HDR_ALTS[gi % 2]
        _cell(current_row, 2, group, bg=bg, bold=True, sz=9, fc="1E3A5F", h="left")

        sector_mcap = {s: 0.0 for s in IDX_SECTORS_LIST}
        total_mcap  = 0.0
        for ticker in tickers:
            mc = _get_mcap(ticker)
            if not (np.isfinite(mc) and mc > 0):
                continue
            sec = _get_sector(ticker)
            if sec in sector_mcap:
                sector_mcap[sec] += mc
            total_mcap += mc

        for si, sector in enumerate(IDX_SECTORS_LIST, start=3):
            pct = (sector_mcap[sector] / total_mcap * 100) if total_mcap > 0 else 0.0
            cell = ws.cell(current_row, si)
            cell.value = round(pct, 1) if pct > 0 else None
            cell.fill  = _fill(bg)
            cell.font  = _font(sz=9, color="374151" if pct > 0 else "D1D5DB")
            cell.alignment = _aln(h="center")
            cell.border = bdr_full
            cell.number_format = "0.0\"%\""

        total_col = 3 + len(IDX_SECTORS_LIST)
        _cell(current_row, total_col, _fmt_idr(total_mcap), bg=bg, bold=True, sz=9, fc="1E3A5F", h="right")
        current_row += 1

    matrix_data_end = current_row - 1

    # Apply color scale to the sector exposure % cells
    try:
        for si, sector in enumerate(IDX_SECTORS_LIST, start=3):
            col_letter = get_column_letter(si)
            rng = f"{col_letter}{matrix_data_start}:{col_letter}{matrix_data_end}"
            ws.conditional_formatting.add(rng, ColorScaleRule(
                start_type="min",  start_value=0,  start_color="FFFFFF",
                mid_type="num",    mid_value=20,   mid_color="BDD7EE",
                end_type="max",    end_value=100,  end_color="1F497D",
            ))
    except Exception:
        pass

    # ── Section 4: Fear & Greed snapshot ─────────────────────────────────────
    current_row += 2
    try:
        _ov_md  = get_market_date()
        _ov_end = (_ov_md + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        _j = yf.download("^JKSE", period="2y", end=_ov_end, interval="1d",
                          auto_adjust=True, progress=False, threads=False)
        _u = yf.download("USDIDR=X", period="2y", end=_ov_end, interval="1d",
                          auto_adjust=True, progress=False, threads=False)

        def _close(raw):
            if raw is None or raw.empty:
                return None
            if isinstance(raw.columns, pd.MultiIndex):
                raw.columns = [c[0] for c in raw.columns]
            s = pd.to_numeric(raw.get("Close", raw.iloc[:,0]), errors="coerce").dropna()
            s.index = pd.to_datetime(s.index, utc=False).normalize()
            if hasattr(s.index, "tz") and s.index.tz is not None:
                s.index = s.index.tz_localize(None)
            return s[s.index <= asof]

        jci_s = _close(_j)
        usd_s = _close(_u)

        if jci_s is not None and usd_s is not None and len(jci_s) >= 130:
            df_fg = pd.concat({"JCI": jci_s, "USDIDR": usd_s}, axis=1, sort=True).dropna()
            df_fg["ma125"]    = df_fg["JCI"].rolling(125, min_periods=80).mean()
            df_fg["mom"]      = (df_fg["JCI"] - df_fg["ma125"]) / df_fg["ma125"] * 100
            df_fg["vol20"]    = np.log(df_fg["JCI"] / df_fg["JCI"].shift(1)).rolling(20, min_periods=15).std() * np.sqrt(252)
            df_fg["usdidr20"] = (df_fg["USDIDR"] / df_fg["USDIDR"].shift(20) - 1) * 100

            def _rzs(s, w=252):
                mu = s.rolling(w, min_periods=60).mean()
                sd = s.rolling(w, min_periods=60).std(ddof=0).replace(0, np.nan)
                return (s - mu) / sd

            comp_z = pd.concat([
                _rzs(df_fg["mom"]),
                -1.0 * _rzs(df_fg["vol20"]),
                -1.0 * _rzs(df_fg["usdidr20"]),
            ], axis=1).mean(axis=1)

            df_fg["FGI"] = (50.0 + 50.0 * np.tanh(comp_z)).clip(0, 100)
            latest_fg   = df_fg["FGI"].dropna()
            fgi_now     = float(latest_fg.iloc[-1])  if not latest_fg.empty else np.nan
            fgi_prev    = float(latest_fg.iloc[-2]) if len(latest_fg) >= 2 else np.nan
            fgi_delta   = fgi_now - fgi_prev if np.isfinite(fgi_now) and np.isfinite(fgi_prev) else np.nan
            jci_now     = float(jci_s.iloc[-1]) if not jci_s.empty else np.nan

            fgi_label = (
                "Extreme Fear" if fgi_now < 25 else
                "Fear"         if fgi_now < 45 else
                "Neutral"      if fgi_now < 55 else
                "Greed"        if fgi_now < 75 else "Extreme Greed"
            )
            fgi_bg = {
                "Extreme Fear":"FEE2E2", "Fear":"FFEDD5", "Neutral":"FEF9C3",
                "Greed":"DCFCE7", "Extreme Greed":"BBF7D0",
            }.get(fgi_label, "F9FAFB")
            fgi_fc = {
                "Extreme Fear":"991B1B", "Fear":"9A3412", "Neutral":"92400E",
                "Greed":"166534", "Extreme Greed":"14532D",
            }.get(fgi_label, "374151")

            # Write FGI panel
            ws.row_dimensions[current_row].height = 5
            current_row += 1
            ws.merge_cells(f"B{current_row}:F{current_row}")
            c = ws.cell(current_row, 2, "▌ Fear & Greed Index Snapshot  (JCI-based, 3-component model)")
            c.fill = _fill("0F2A45"); c.font = _font(bold=True, sz=11, color="FFFFFF")
            c.alignment = _aln(h="left"); c.border = bdr_full
            for col in range(3, 7):
                ws.cell(current_row, col).fill  = _fill("0F2A45")
                ws.cell(current_row, col).border = bdr_full
            ws.row_dimensions[current_row].height = 20
            current_row += 1

            fgi_fields = [
                ("FGI Score",   f"{fgi_now:.1f} / 100"),
                ("Sentiment",   fgi_label),
                ("Δ vs Prev Day", f"{'+' if fgi_delta >= 0 else ''}{fgi_delta:.1f}" if np.isfinite(fgi_delta) else "—"),
                ("JCI Level",   f"{jci_now:,.0f}" if np.isfinite(jci_now) else "—"),
                ("Components",  "Momentum (33%)  ·  Volatility (33%)  ·  USD/IDR (33%)"),
            ]
            for label, value in fgi_fields:
                ws.row_dimensions[current_row].height = 20
                _cell(current_row, 2, label, bg="F3F4F6", bold=True, sz=10, fc="374151", h="left")
                c2 = _cell(current_row, 3, value, bg=fgi_bg, bold=True, sz=10, fc=fgi_fc, h="left")
                ws.merge_cells(f"C{current_row}:F{current_row}")
                current_row += 1

            # Write last 30 days of FGI history as mini-table
            history_slice = df_fg[["JCI","FGI"]].dropna().tail(30)
            if not history_slice.empty:
                current_row += 1
                ws.row_dimensions[current_row].height = 18
                for ci, hdr in enumerate(["Date","JCI Level","FGI Score","Sentiment"], start=2):
                    _cell(current_row, ci, hdr, bg=HDR_BG, bold=True, sz=9, fc="FFFFFF")
                current_row += 1
                for dt, hist_row in history_slice.iterrows():
                    ws.row_dimensions[current_row].height = 16
                    fgi_v = float(hist_row["FGI"])
                    lbl   = ("Extreme Fear" if fgi_v < 25 else "Fear" if fgi_v < 45 else
                             "Neutral" if fgi_v < 55 else "Greed" if fgi_v < 75 else "Extreme Greed")
                    row_bg = {
                        "Extreme Fear":"FEE2E2","Fear":"FFEDD5","Neutral":"FFFBEB",
                        "Greed":"DCFCE7","Extreme Greed":"BBF7D0",
                    }.get(lbl, "FFFFFF")
                    dt_val = dt.to_pydatetime() if hasattr(dt, "to_pydatetime") else dt
                    _cell(current_row, 2, dt_val,  bg=row_bg, sz=9, fmt="DD-MMM-YY")
                    _cell(current_row, 3, float(hist_row["JCI"]), bg=row_bg, sz=9, fmt="#,##0")
                    _cell(current_row, 4, round(fgi_v, 1),        bg=row_bg, sz=9, fmt="0.0")
                    _cell(current_row, 5, lbl,                    bg=row_bg, sz=9, fc=fgi_fc)
                    current_row += 1

                # DataBar on FGI column
                try:
                    fgi_col = get_column_letter(4)
                    ws.conditional_formatting.add(
                        f"{fgi_col}{current_row - len(history_slice)}:{fgi_col}{current_row - 1}",
                        DataBarRule(start_type="num", start_value=0,
                                    end_type="num",   end_value=100,
                                    color="638EC6"),
                    )
                except Exception:
                    pass
    except Exception as _fg_err:
        ws.row_dimensions[current_row].height = 18
        _cell(current_row, 2,
              f"Fear & Greed panel unavailable — {_fg_err}",
              bg="FEF3C7", bold=False, sz=9, fc="92400E", h="left", bdr=False)

    # ── Freeze + autofilter ───────────────────────────────────────────────────
    ws.freeze_panes = "C6"

    return ws


# =========================
# GUIDE & LOGIC REFERENCE  (Full rewrite — reflects all v2 engines)
# =========================
def build_guide_sheet(ws):
    """
    Guide & Logic Reference — institutional-grade documentation for IDX_Screener workbook.
    Covers: screener filters, VWAP methodology, SMC engine, market structure, news logic,
            indicator formulas, filter conditions, playbook patterns, and change log.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws.title = "Guide & Logic Reference"
    ws.sheet_view.showGridLines = False

    # ── Column widths ────────────────────────────────────────────────────────
    col_widths = {"A": 0.5, "B": 30, "C": 48, "D": 56, "E": 38, "F": 38}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # ── Palette ──────────────────────────────────────────────────────────────
    C_TITLE   = "0D1B2A"   # dark navy
    C_S1      = "1F4E79"   # section 1 — screener filters
    C_S2      = "375623"   # section 2 — VWAP
    C_S3      = "4A0080"   # section 3 — SMC
    C_S4      = "7B3F00"   # section 4 — market structure
    C_S5      = "005050"   # section 5 — news/sentiment
    C_S6      = "3B3000"   # section 6 — indicators
    C_S7      = "1C3A5E"   # section 7 — weekday logic
    C_S8      = "5C1B00"   # section 8 — institutional engine
    C_S9      = "003333"   # section 9 — playbook
    C_S10     = "2D2D2D"   # section 10 — changelog
    C_HDR_ROW = "1C3253"   # column sub-header
    C_ODD     = "F5F8FB"
    C_EVEN    = "FFFFFF"

    THIN = Side(style="thin", color="D0D7E0")
    BORDER_CELL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def _fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def _font(color="1F2937", bold=False, size=10, italic=False):
        return Font(name="Calibri", size=size, bold=bold, italic=italic, color=color)

    def _align(h="left", v="top", wrap=True):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    row_idx = [1]

    def _next_row():
        r = row_idx[0]
        row_idx[0] += 1
        return r

    def _title_row(text):
        r = _next_row()
        ws.row_dimensions[r].height = 34
        c = ws.cell(r, 2, text)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c.fill = _fill(C_TITLE)
        c.font = _font("FFFFFF", bold=True, size=16)
        c.alignment = _align("center", "center")

    def _section_header(text, color):
        r = _next_row()
        ws.row_dimensions[r].height = 22
        c = ws.cell(r, 2, text)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c.fill = _fill(color)
        c.font = _font("FFFFFF", bold=True, size=11)
        c.alignment = _align("left", "center")

    def _col_header(cols):
        """cols = list of (col_num, text)"""
        r = _next_row()
        ws.row_dimensions[r].height = 18
        for col_n, text in cols:
            c = ws.cell(r, col_n, text)
            c.fill = _fill(C_HDR_ROW)
            c.font = _font("FFFFFF", bold=True, size=9)
            c.alignment = _align("center", "center")

    def _data_row(b, c, d, e, f, odd=True):
        r = _next_row()
        ws.row_dimensions[r].height = 48
        bg = C_ODD if odd else C_EVEN
        for col_n, text in ((2,b),(3,c),(4,d),(5,e),(6,f)):
            cell = ws.cell(r, col_n, text)
            cell.fill = _fill(bg)
            cell.font = _font(size=9)
            cell.alignment = _align(wrap=True)
            cell.border = BORDER_CELL

    def _spacer():
        r = _next_row()
        ws.row_dimensions[r].height = 6

    def _note_row(text, color):
        r = _next_row()
        ws.row_dimensions[r].height = 14
        c = ws.cell(r, 2, text)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c.fill = _fill(color)
        c.font = _font("FFFFFF", italic=True, size=8)
        c.alignment = _align("left", "center")

    HDR_COLS = [(2,"Column / Item"),(3,"Formula / Logic"),(4,"Interpretation"),(5,"Bullish / Pass"),(6,"Bearish / Fail")]

    # ═══════════════════════════════════════════════════════════════════════
    # TITLE
    # ═══════════════════════════════════════════════════════════════════════
    _title_row("IDX Screener — Guide & Logic Reference  (v4.0)")
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 1 — IDX SCREENER FILTERS
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§1  IDX Screener — Filter Logic (A–G)", C_S1)
    _note_row("All filters share a base gate: Close > EMA25 > EMA50, RSI > 50, ADTV ≥ IDR 25B.  Rows meeting the gate appear in Filter A by default.", C_S1)
    _col_header(HDR_COLS)
    rows_s1 = [
        ("Filter A — EMA Trend ↑",
         "Close > EMA25 AND EMA25 > EMA50 AND RSI(14) > 50 AND ADTV20 ≥ 25B IDR",
         "Price is above short and medium-term moving averages with positive momentum. Confirms bullish trend alignment across three timeframes.",
         "All three conditions TRUE → included in Filter A section",
         "Any condition FALSE → excluded"),
        ("Filter B — Golden Cross",
         "EMA10 > EMA20 (golden cross) OR MACD Line crossed above Signal Line on market date",
         "Short-term momentum crossing above medium-term MA, or MACD bullish crossover. Signals early trend acceleration.",
         "Cross event detected on or within lookback → row included",
         "No cross detected → excluded"),
        ("Filter C — BOS (Swing + Internal)",
         "Swing BOS OR Internal BOS event date == MARKET_DATE (exact match, Bull direction only, CHoCH excluded)",
         "Break of Structure on either swing or internal timeframe fired on the selected market date. Earlier entry signal vs prior swing-only filter.",
         "BOS BULL event date = MARKET_DATE on either structure level → pass",
         "No BOS today, or CHoCH (reversal) event instead → excluded"),
        ("Filter D — POI Reclaim",
         "Candle Low ≤ POI level AND Close ≥ POI level (price wicked into POI and closed above)",
         "Price dipped into a key level (VWAP/EMA/OB) and reclaimed it intraday — classic institutional accumulation footprint.",
         "Low touches POI AND close reclaims above → pass",
         "Close below POI after touch → excluded"),
        ("Filter E — EQ Breakout",
         "Previous bar inside VWAP equilibrium zone (between -1σ and +1σ); current bar close outside zone on upside",
         "Price breaks out of consolidation equilibrium, signalling expansion phase initiation.",
         "Close moves above +1σ band from prior equilibrium zone → pass",
         "No breakout from equilibrium → excluded"),
        ("Filter F — Near VWAP",
         "q_remarks or pq_remarks starts with 'Price Near' (within 2% of PQ or PY VWAP band). Current Q excluded.",
         "Price is testing a previous-quarter or previous-year VWAP band — key institutional reference level often used as entry zone.",
         "Within 2% of PQ or PY VWAP/σ band → pass",
         "Beyond 2% of all historical VWAP bands → excluded"),
        ("Filter G — SMC Location",
         "smc_premium_discount IN ('In Bull OB', 'Equilibrium', 'Discount')",
         "Price is positioned in a structurally favourable SMC zone: inside bullish OB, at EQ (50% of structure range), or in discount (below-50% zone). Premium and Bear OB are excluded.",
         "Zone label = 'In Bull OB' OR 'Equilibrium' OR 'Discount' → pass",
         "Zone = 'Premium', 'In Bear OB', 'Mid-Zone', or N/A → excluded"),
    ]
    for i, row_data in enumerate(rows_s1):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 2 — ANCHORED VWAP METHODOLOGY
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§2  Anchored VWAP — Methodology & Bands", C_S2)
    _note_row("Four VWAP anchors: Current Month (MVWAP), Previous Month, Current Quarter (QVWAP), Previous Quarter, Previous Year. All use volume-weighted typical price.", C_S2)
    _col_header(HDR_COLS)
    rows_s2 = [
        ("VWAP Formula",
         "VWAP = Σ(Typical_Price × Volume) / Σ(Volume)  |  Typical Price = (H+L+C)/3",
         "True anchored VWAP from the first bar of the anchor period to MARKET_DATE. Each anchor resets at its period start — not a rolling average.",
         "Price above VWAP → bullish bias for that anchor",
         "Price below VWAP → bearish or reversion bias"),
        ("σ Bands (Standard Deviation)",
         "σ = √[ Σ(Volume × (TP − VWAP)²) / Σ(Volume) ]  |  Bands at ±1σ, ±2σ, ±3σ",
         "Volume-weighted standard deviation bands. Wider bands = higher dispersion / volatility. Bands expand as more bars accumulate.",
         "Price between VWAP and +1σ = mild bullish extension; above +2σ = overbought territory",
         "Price between VWAP and -1σ = mild weakness; below -2σ = oversold / capitulation zone"),
        ("Current MVWAP (cm_*)",
         "Anchor: first trading day of current calendar month → MARKET_DATE",
         "Shortest anchor — most responsive. Resets monthly. Price σ = current month's deviation score. Used for intra-month bias.",
         "cm_sd > 0 → above monthly VWAP (bullish intra-month)",
         "cm_sd < 0 → below monthly VWAP (bearish intra-month)"),
        ("Previous MVWAP (pm_*)",
         "Anchor: first → last trading day of prior calendar month (complete anchor)",
         "Completed monthly VWAP — fixed reference. Useful as support/resistance level. pm_delta = cm_sd − pm_sd (relative strength).",
         "Price above pm_vwap → above prior month reference",
         "Price below pm_vwap → potentially retesting prior month distribution"),
        ("Current QVWAP (q_*)",
         "Anchor: first trading day of current calendar quarter → MARKET_DATE",
         "Primary institutional VWAP reference. Institutional desks use quarterly VWAP as benchmark. q_sd is the live σ score vs this anchor.",
         "q_sd > 0 = above quarterly VWAP (primary bull signal)",
         "q_sd < -1 = below -1σ; potential OB or capitulation zone"),
        ("Previous QVWAP (pq_*)",
         "Anchor: full prior quarter (complete). pq_delta = q_sd − pq_sd",
         "Completed quarterly VWAP. Price vs pq_vwap shows whether current quarter has recovered vs prior quarter's distribution.",
         "Price above pq_vwap = relative strength vs Q−1",
         "Price below pq_vwap = still in Q−1 distribution territory"),
        ("Previous Year VWAP (py_*)",
         "Anchor: full prior calendar year (complete). py_delta = q_sd − py_sd",
         "Annual VWAP — slowest anchor, used as macro bullish/bearish dividing line. Crossing above py_vwap = regime shift signal.",
         "Price above py_vwap = bullish macro regime",
         "Price below py_vwap = macro bearish / year-long distribution"),
        ("Price σ (sd score)",
         "sd_score = (Close − VWAP) / σ  |  positive = above VWAP, negative = below",
         "Normalised distance from VWAP in standard deviation units. Allows direct comparison across tickers regardless of price level.",
         "sd_score ∈ (0, +1): within +1σ = healthy extension. sd_score > +2: overextended.",
         "sd_score < −1: below -1σ = weakness. sd_score < −2: capitulation / OB zone."),
        ("Price σ Δ 1D (delta)",
         "q_delta = today's q_sd − yesterday's q_sd  |  shows intraday momentum direction",
         "One-day change in normalised σ distance. Positive = moving away from VWAP on upside; negative = reverting towards or below VWAP.",
         "Δ > 0 = momentum continuing upward vs anchor",
         "Δ < 0 = momentum fading or price reverting"),
        ("VWAP Zone Days",
         "_count_consecutive_zone_days(): counts bars where q_remarks matches current zone",
         "How many consecutive sessions price has stayed in the same VWAP zone. Higher count = stronger zone conviction.",
         "> 3 consecutive days in same zone = structural confirmation",
         "0–1 days = fleeting touch, not conviction"),
    ]
    for i, row_data in enumerate(rows_s2):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 3 — SMC ENGINE
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§3  Smart Money Concepts (SMC) Engine — 26 Columns", C_S3)
    _note_row("SMC engine runs on daily bars. Internal structure uses shorter swing detection; Swing structure uses broader highs/lows. BOS = same-direction break; CHoCH = reversal.", C_S3)
    _col_header(HDR_COLS)
    rows_s3 = [
        ("Order Blocks (OB)",
         "Last bearish candle before BOS up (Bull OB) / last bullish candle before BOS down (Bear OB). OB High/Low stored.",
         "Order Blocks are supply/demand zones where institutional orders were placed. Price returning to OB = potential re-entry zone.",
         "Price inside Bull OB zone (smc_premium_discount = 'In Bull OB') → high-probability long setup",
         "Price inside Bear OB zone (= 'In Bear OB') → distribution / short risk"),
        ("BOS (Break of Structure)",
         "Price closes beyond most recent swing high (Bull BOS) or swing low (Bear BOS). Labels: 'BOS BULL yyyy-mm-dd' or 'BOS BEAR yyyy-mm-dd'.",
         "BOS confirms continuation of existing trend. Bull BOS = higher high broken → bullish continuation. Filter C now includes both Internal and Swing BOS.",
         "Latest struct = BOS BULL on MARKET_DATE → Filter C pass",
         "Latest struct = BOS BEAR or CHoCH BEAR → structural breakdown"),
        ("CHoCH (Change of Character)",
         "Price closes beyond opposite swing extreme. Bull CHoCH = lower low broken (reversal warning). Bear CHoCH = higher high broken (potential reversal).",
         "CHoCH signals trend reversal vs BOS (continuation). Screener explicitly excludes CHoCH from Filter C to avoid false entries on reversal candles.",
         "CHoCH BULL → price reclaiming prior swing high = potential reversal from downtrend",
         "CHoCH BEAR → structural weakness, prior low taken out"),
        ("Premium / Discount / EQ Zones",
         "EQ = 50% of current swing range (FVG midpoint). Discount = below EQ (0–50%). Premium = above EQ (50–100%).",
         "SMC entry framework: buy in discount, sell in premium. Equilibrium = rebalancing zone. Filter G uses these labels as location screener.",
         "Price in Discount or Equilibrium = structurally favourable buy zone",
         "Price in Premium = overextended; risk of reversion to EQ"),
        ("FVG (Fair Value Gap)",
         "Three-bar pattern: body of bar 1 and body of bar 3 have a gap with no overlap. Bullish FVG = price gap up; Bearish FVG = price gap down.",
         "FVG = imbalance zone that price tends to revisit ('rebalance'). smc_fvg_bias summarises the net directional bias from open FVGs.",
         "smc_fvg_bias = 'Bullish' → net upward FVG imbalance",
         "smc_fvg_bias = 'Bearish' → downward imbalance, supply overhead"),
        ("Liquidity Levels (EQH/EQL)",
         "Equal Highs (EQH) and Equal Lows (EQL): bars with nearly identical highs/lows within a tolerance. Represent resting liquidity pools.",
         "Clusters of EQH = sell-side liquidity above; clusters of EQL = buy-side liquidity below. Smart money sweeps these levels before reversing.",
         "EQL swept and price reversed up = buy-side liquidity grab → long signal",
         "EQH swept and price reversed down = sell-side liquidity grab → short risk"),
        ("SMC Summary",
         "smc_summary = 'Price is on {pd_label}' | 'Mid-Zone' → displayed as '-' (ambiguous zone suppressed)",
         "Quick-read SMC position status. Mid-Zone (between Discount and Premium outer bounds) is non-actionable and shows '-' to reduce noise.",
         "smc_summary shows 'Price is on Equilibrium' or 'Price is on In Bull OB' → actionable",
         "smc_summary shows '-' (Mid-Zone) or 'N/A' → no clear SMC positioning"),
        ("MS Phase",
         "ms_phase = 'Expansion' if last swing = BOS; 'Reversal' if CHoCH; 'Range' otherwise",
         "Market structure phase derived from last swing event. Expansion = trend continuation. Reversal = CHoCH detected. Range = no recent structure event.",
         "ms_phase = Expansion + BOS BULL → trending bullish",
         "ms_phase = Reversal → CHoCH detected, caution"),
    ]
    for i, row_data in enumerate(rows_s3):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 4 — MARKET PROFILE (PWH/PWL/MDH/MDL/IBH/IBL)
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§4  Market Profile — Weekly & Initial Balance Levels", C_S4)
    _note_row("Weekend Logic: if script runs on Saturday or Sunday (WIB), PWH/PWL uses the current week's Mon–Fri range (the week that just closed on Friday). Weekday runs use the prior completed week.", C_S4)
    _col_header(HDR_COLS)
    rows_s4 = [
        ("PWH / PWL",
         "PWH = max(High) of prior completed week Mon–Fri  |  PWL = min(Low) of prior completed week. Weekend override: current week used if run on Sat/Sun WIB.",
         "Previous Week High/Low are key weekly support/resistance. Price vs PWL shows whether current week has held above or broken the prior weekly range.",
         "Close > PWH = prior week breakout (bullish). Close > PWL = holding above prior week base.",
         "Close < PWL = prior week range broken → bearish continuation risk"),
        ("MDH / MDL",
         "MDH = High of first trading day of current ISO week  |  MDL = Low of first trading day of current ISO week",
         "Monday's (or first-day's) candle high/low sets the week's initial expansion reference. Price vs MDL shows intra-week holding pattern.",
         "Close > MDH = exceeded opening day range = expansion day",
         "Close < MDL = failed to hold opening day low = weakness"),
        ("IBH / IBL",
         "IBH = max(High) of first 2 bars of current calendar month  |  IBL = min(Low) of first 2 bars",
         "Initial Balance High/Low = first two sessions of the month. Monthly IB defines the acceptance range for the month's distribution.",
         "Close > IBH = monthly IB breakout = bullish month bias",
         "Close < IBL = below monthly IB = bearish month bias"),
        ("MP Summary",
         "Composite label: 'PWL Above | MDL Above | IBL Above' (or Below/At Level for each)",
         "Quick-read market profile summary. All three 'Above' = price holding all key weekly/monthly levels.",
         "All three = Above → strong multi-level support confirmation",
         "Any = Below → structural level broken"),
        ("Weekend WIB Logic",
         "pd.Timestamp.now(tz='Asia/Jakarta').weekday() >= 5 → _is_weekend_run = True",
         "WIB = UTC+7 (Asia/Jakarta). Checked against wall-clock run time (not MARKET_DATE). Ensures weekend runs show the just-completed week's range in PWH/PWL, not the week prior.",
         "Saturday or Sunday WIB run: PWH/PWL = current week (Mon–Fri of that week)",
         "Weekday WIB run: PWH/PWL = prior completed week (standard logic)"),
    ]
    for i, row_data in enumerate(rows_s4):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 5 — NEWS & SENTIMENT ENGINE
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§5  News & Sentiment Engine — Logic & Windows", C_S5)
    _note_row("Sentiment news: 7-day window, 1-day priority pass first. Corp action news: 30-day window (≈ 1 month). Both use entity validation before display.", C_S5)
    _col_header(HDR_COLS)
    rows_s5 = [
        ("Sentiment News — Window",
         "_pick_best_item(): priority pass = last 1 day; fallback = last 7 days. Items older than 7 days are NOT shown.",
         "Short 7-day window prioritises recency. If a same-day (or yesterday) item exists, it's always shown first. Stale news beyond 7 days is suppressed entirely.",
         "Item from last 1 day → 'On-Date' source label, shown with priority",
         "No item within 7 days → empty cell (N/A), not stale news"),
        ("Corp Action News — Window",
         "_fetch_corp_action_screener(): 30-day rolling window (_cutoff ≤ date ≤ MARKET_DATE). Corrected from prior OR-logic bug.",
         "Corp actions (dividends, rights issues, buybacks, RUPS) use a wider 30-day window as these events have longer relevance windows.",
         "Corp action item within 30 days → shown with corporate keyword-validated headline",
         "No item within 30 days → empty (not stale content from prior months)"),
        ("Entity Validation",
         "validate_news_entity_match(title, ticker, company_name): checks ticker code OR company name aliases appear in headline",
         "Prevents cross-contamination headlines (e.g., BBRI headline showing for BMRI). Only headlines matching the stock's entity are displayed.",
         "Entity match confirmed → headline displayed",
         "Entity mismatch → rejected, logged to _NEWS_REJECTION_LOG, cell shows empty"),
        ("Age Labels",
         "_days_since_label(): 'today', '2d', '3w', '1mo' — compact format",
         "Compact age display in the IDX News sheet. Helps identify whether news is fresh (same day) or approaching the end of its window.",
         "Age = 'today' or '1d' → freshest available signal",
         "Age > '5d' → approaching 7-day sentiment cutoff"),
    ]
    for i, row_data in enumerate(rows_s5):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 6 — TECHNICAL INDICATORS
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§6  Technical Indicators — Formulas & Interpretation", C_S6)
    _col_header(HDR_COLS)
    rows_s6 = [
        ("RSI(14)",
         "Wilder's RMA (EMA with α=1/14). RSI = 100 − 100/(1 + RS)  |  RS = avg gain / avg loss over 14 bars",
         "Momentum oscillator 0–100. RSI > 50 = net bullish momentum (required in base gate). RSI > 70 = overbought; RSI < 30 = oversold.",
         "RSI > 50 → bullish momentum gate passed. RSI crossing 50 upward = momentum shift signal.",
         "RSI < 50 → excluded from all filters. RSI < 30 = capitulation zone."),
        ("MACD 4C Smooth",
         "MACD Line = EMA(12) − EMA(26)  |  Signal = EMA(9) of MACD Line  |  Histogram = MACD − Signal  |  EMA-smoothed histogram for display",
         "4-colour histogram: dark green (increasing pos), light green (decreasing pos), light red (increasing neg), dark red (decreasing neg). Exact Pine Script port.",
         "MACD Line crosses above Signal → bullish crossover (Filter B trigger)",
         "MACD Line crosses below Signal → bearish crossover"),
        ("EMA / SMA",
         "EMA(n) = Close × (2/(n+1)) + prev_EMA × (1 − 2/(n+1))  |  SMA(200) = simple mean of last 200 closes",
         "Trend reference levels. EMA25 = short-term; EMA50 = medium-term; SMA200 = long-term regime divider. Price vs MA signals trend bias.",
         "Close > EMA25 > EMA50 → bullish trend alignment (required for all filters)",
         "EMA25 < EMA50 → bearish cross (dead cross) → excluded from all filters"),
        ("ADR% (14)",
         "ADR = SMA(High − Low, 14) including current bar  |  ADR% = ADR / Close × 100",
         "Average Daily Range as % of price. Measures volatility. High ADR% = high-volatility / momentum stock. Low ADR% = compressed / inactive.",
         "ADR% ≥ 3% AND ATR14% ≥ 3% → ADR/ATR Zone = 'Above 3%'",
         "Both < 3% → 'Below 3%' = low momentum / illiquid"),
        ("RVOL (5D)",
         "RVOL5 = last_volume / SMA(volume, prior 5 bars). RVOL20 = last_volume / ADTV20.",
         "Relative Volume — how active today vs recent average. RVOL5 ≥ 1.5 = above-average participation. High RVOL5 on breakout = institutional engagement.",
         "RVOL5 ≥ 1.5 → RVOL Zone = 'YES' (strong volume)",
         "RVOL5 < 1.5 → 'NO' (weak volume, less conviction)"),
        ("IBD RS Rating",
         "Price performance rank vs IDX universe over 12 months (3/1/0.5/0.25 weighting). Normalized 1–99.",
         "Relative Strength Rating from IBD methodology. RS ≥ 80 = top-20% performer. Used as momentum quality filter.",
         "RS ≥ 80 = market leader strength",
         "RS < 50 = laggard, avoid unless reversal setup"),
        ("Golden Cross",
         "EMA10 crosses above EMA20 (bullish golden cross). Detected by comparing current vs prior day's MA relationship.",
         "Short-term EMA crossing medium-term = early momentum signal. Filter B includes this as alternative to MACD cross.",
         "EMA10 crosses above EMA20 today → golden cross detected → Filter B pass",
         "EMA10 crosses below EMA20 → dead cross → bearish momentum signal"),
    ]
    for i, row_data in enumerate(rows_s6):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 7 — MARKET STRUCTURE & REGIME
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§7  Market Structure & Stock Regime Classification", C_S7)
    _col_header(HDR_COLS)
    rows_s7 = [
        ("Stock Regime",
         "Stock Regime = composite of trend alignment + VWAP zone + volume quality + ADR%",
         "Classifies tickers into: Trending Bull, Accumulation, Distribution, Trending Bear, or Range. Affects tier assignment.",
         "Trending Bull or Accumulation → highest tier assignment priority",
         "Trending Bear or Distribution → excluded from bullish filters"),
        ("Tier (A / B / C)",
         "Tier A: above all MAs + VWAP zone confirmed + RVOL ≥ 1.5. Tier B: 3+ MAs above. Tier C: baseline.",
         "Tier assignment summarises overall setup quality. Tier A = institutional-grade setup. Tier B = developing setup. Tier C = monitor only.",
         "Tier A → highest priority in ranked screener output",
         "No tier assigned → setup incomplete"),
        ("MA Zone",
         "Checks EMA25p, EMA50p, SMA200p: 'Bullish' if all Above; 'Bearish' if all Below; 'Mixed' otherwise",
         "Summary of price position relative to three key MAs. Bullish = all Above = strong trend alignment.",
         "MA Zone = Bullish → all three MAs aligned above price",
         "MA Zone = Bearish or Mixed → partial or conflicted trend"),
        ("MS Phase",
         "ms_phase = Expansion (BOS) | Reversal (CHoCH) | Range (no event)",
         "Market structure phase for the stock's own internal structure. Expansion = trend continuation active.",
         "ms_phase = Expansion = BOS detected → trend is intact",
         "ms_phase = Reversal = CHoCH = potential trend change"),
    ]
    for i, row_data in enumerate(rows_s7):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 8 — INSTITUTIONAL ENGINE (Composite Score / AMT State)
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§8  Institutional Engine — Composite Score & AMT State", C_S8)
    _note_row("compute_institutional_metrics() → _build_idx_vwap_shortlist() → _sc_* fields → IDX Screener columns. Functions confirmed to produce workbook output.", C_S8)
    _col_header(HDR_COLS)
    rows_s8 = [
        ("Composite Score",
         "_sc_swing_score = weighted sum of sub-scores: VWAP position, MACD, RSI, RVOL, SMC, IBD RS, Tier, ADR%",
         "0–10 score summarising overall setup quality. Used to rank tickers within each filter section. Higher = better alignment.",
         "Score ≥ 8 = green (high conviction). Score 6–7 = blue (moderate).",
         "Score < 4 = grey (low conviction / monitor only)"),
        ("AMT State",
         "_derive_amt_state_v2(): Auction Market Theory state = Trend / Range / Transition based on VWAP slope + volume profile",
         "AMT classifies whether the market is in directional trend (one-timeframe) or balancing (two-timeframe). Affects POI ranking.",
         "AMT State = Trend → directional, use breakout entries",
         "AMT State = Range → mean-reversion entries preferred"),
        ("Cause Quality",
         "_sc_cause_quality: measures duration and volume of base/accumulation phase before current move",
         "Higher cause = more fuel for the move. Wide base + high volume cause = stronger potential advance.",
         "High cause quality → strong base formed, move likely to sustain",
         "Low cause quality → thin base, fade risk higher"),
        ("Markup Readiness",
         "_sc_markup_readiness: combination of volume trend, breakout proximity, OB proximity, and momentum alignment",
         "Readiness score for transition from accumulation to markup phase. High score = multiple markup preconditions met.",
         "High markup readiness → accumulation complete, markup imminent",
         "Low readiness → still in accumulation or distribution"),
    ]
    for i, row_data in enumerate(rows_s8):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 9 — PLAYBOOK PATTERNS
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§9  Playbook — Common IDX Swing Trade Patterns", C_S9)
    _note_row("These patterns integrate multiple columns. Use as entry frameworks, not standalone signals. Always confirm with ADTV and RS before sizing.", C_S9)
    _col_header([(2,"Pattern"),(3,"Entry Trigger"),(4,"Confirmation Stack"),(5,"Target Zone"),(6,"Invalidation")])
    rows_s9 = [
        ("VWAP Reclaim",
         "Price dips to q_vwap or q_m1 (-1σ) and closes back above on elevated RVOL (≥ 1.5)",
         "q_zone transitions from below-VWAP to above-VWAP. RSI crosses 50. MACD histogram turns positive.",
         "Next VWAP band: q_p1 (+1σ) as first target; q_p2 (+2σ) as extension",
         "Daily close below q_m1 (-1σ) after entry = invalid; stop below q_m2"),
        ("BOS Retest Entry",
         "BOS BULL fires (Filter C pass). Wait for retest of broken level (prior swing high now acting as support).",
         "Volume on BOS day > ADTV20 × 1.5. SMC zone = Discount or Equilibrium. Price holds above broken level on retest.",
         "Minimum 1:2 R/R. Target = next SMC resistance / OB / VWAP band.",
         "Close below BOS level on retest = failed BOS → exit immediately"),
        ("Golden Cross Momentum",
         "EMA10 crosses above EMA20 (Filter B). MACD Line already above Signal.",
         "IBD RS ≥ 70. ADTV20 ≥ IDR 50B. Price above SMA200. RVOL5 ≥ 1.3 on cross day.",
         "First target: prior swing high or +1σ q_band. Trail stop below EMA25 on daily close.",
         "EMA10 crosses back below EMA20 within 3 days = false cross → exit"),
        ("OB Touch & Hold",
         "Price enters Bull OB zone (smc_premium_discount = 'In Bull OB'). Filter G pass.",
         "OB zone respected (price doesn't close below OB low). MACD not in steep downtrend. RSI ≥ 45.",
         "Target = equilibrium (EQ at 50% of current structure). Extension = Premium zone.",
         "Close below OB low = OB invalidated → exit. New BOS BEAR = structural break."),
        ("Monthly VWAP Base",
         "Price consolidates between pm_vwap and cm_vwap for 3+ sessions. Breakout above cm_vwap on RVOL > 1.5.",
         "cm_sd crosses 0 (from below to above). q_sd also positive. No SMC Bear OB immediately above.",
         "First target: cm_p1 (+1σ Monthly). Extension: q_p1 (+1σ Quarterly).",
         "Close below pm_vwap = multi-month support broken → invalidated"),
    ]
    for i, row_data in enumerate(rows_s9):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()

    # ═══════════════════════════════════════════════════════════════════════
    # SECTION 10 — CHANGELOG
    # ═══════════════════════════════════════════════════════════════════════
    _section_header("§10  Changelog — Recent Updates", C_S10)
    _col_header([(2,"Version / Date"),(3,"Change Description"),(4,"Detail"),(5,"Impact"),(6,"Status")])
    rows_s10 = [
        ("v4.0 — Jun 2026",
         "MVWAP groups added: Current Monthly (cm_*) and Previous Monthly (pm_*)",
         "Two new 12-column VWAP groups inserted before QVWAP in DETAIL_SCHEMA. Computed using month-start anchor via anchored_vwap_block().",
         "IDX Technical Detail sheet: +24 new columns for monthly VWAP reference",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "SD → σ rename across all VWAP column headers and zone labels",
         "All '-1 SD', '-2 SD', '+1 SD' labels in DETAIL_SCHEMA and vwap_near_zone_label() renamed to '-1 σ', '+1 σ' etc. Column headers: 'SD Score' → 'Price σ', 'SD Δ 1D' → 'Price σ Δ 1D'.",
         "Cleaner notation; consistent with statistical convention",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "Positive σ bands added before VWAP in all groups: 3σ, 2σ, 1σ columns",
         "Each VWAP group now shows: Running Days | 3σ | 2σ | 1σ | VWAP | -1σ | -2σ | -3σ | Price σ | Price σ Δ 1D | VWAP Zone | VWAP Zone Days",
         "Full band visibility for overbought/extension analysis",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "Filter C extended: Internal BOS included alongside Swing BOS",
         "_is_bos_today() checks smc_latest_swing_struct AND smc_latest_internal_struct. CHoCH still excluded.",
         "Earlier entry signals; more tickers qualify for Filter C",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "Filter G added: SMC Location Screener",
         "_is_smc_location_screener(): passes when smc_premium_discount ∈ {'In Bull OB', 'Equilibrium', 'Discount'}. Olive colour DNA.",
         "New filter section in IDX Screener for zone-based entry candidates",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "SMC Mid-Zone summary suppressed: 'Price is on Mid-Zone' → '-'",
         "compute_smc_engine(): smc_summary now returns '-' when pd_label == 'Mid-Zone' (ambiguous, non-actionable).",
         "Reduces noise in SMC Summary column",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "PWH/PWL weekend WIB logic",
         "On Saturday/Sunday (WIB), PWH/PWL uses current week Mon–Fri range instead of prior week. Determined by pd.Timestamp.now(tz='Asia/Jakarta').weekday() >= 5.",
         "Weekend runs now show the just-completed week's range in PWH/PWL",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "News window: Sentiment 30d → 7d with 1-day priority pass",
         "_pick_best_item() default window_days changed from 30 to 7. Priority pass added: items within 1 day shown first. No IFNA fallback beyond 7d.",
         "Only fresh news (≤7 days) shown in sentiment columns",
         "✅ Implemented"),
        ("v4.0 — Jun 2026",
         "Corp action date filter OR-bug fixed → AND logic",
         "Filter was: date >= cutoff OR date <= today (kept ALL past items). Fixed to: cutoff <= date <= today (correct 30-day window).",
         "Corp action news now correctly limited to 30-day rolling window",
         "✅ Implemented"),
    ]
    for i, row_data in enumerate(rows_s10):
        _data_row(*row_data, odd=(i % 2 == 0))
    _spacer()


def build_source_limited_guide_sheet(ws):
    """Build the workbook guide from the shared logic registry."""
    ws.delete_rows(1, ws.max_row)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:N1")
    ws["A1"] = "IDX RESEARCH - Guide & Logic Reference"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="0D1B2A")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:N2")
    ws["A2"] = (
        "Generated from rebuild_backend/logic_reference.py. "
        "Definitions are educational and use the project's approved sources and formulas."
    )
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="334155")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")

    headers = [
        "Category", "Concept / Field", "Simple Definition", "Why It Matters",
        "How It Is Calculated", "Required Inputs", "Source",
        "Sheet / Column Output", "Interpretation", "Threshold / Rule",
        "Missing Data Behavior", "Limitations", "Example", "Formula Version",
    ]
    for col_idx, label in enumerate(headers, start=1):
        cell = ws.cell(3, col_idx, label)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1C3253")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    rows = workbook_rows()
    for row_idx, row in enumerate(rows, start=4):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row_idx, col_idx, value)
            cell.font = Font(name="Calibri", size=10, color="1F2937")
            cell.fill = PatternFill("solid", fgColor="F8FAFC" if row_idx % 2 == 0 else "EEF3F8")
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = Border(
                left=Side(style="thin", color="D1D5DB"),
                right=Side(style="thin", color="D1D5DB"),
                top=Side(style="thin", color="D1D5DB"),
                bottom=Side(style="thin", color="D1D5DB"),
            )
        ws.row_dimensions[row_idx].height = 56
    widths = [25, 26, 45, 42, 44, 31, 25, 38, 40, 34, 39, 41, 39, 23]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:N{3 + len(rows)}"


def build_data_source_audit_sheet(wb, latest_market_day, rows):
    """Add field-level source and missing-reason records for key published metrics."""
    name = "Data Source Audit"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    headers = [
        "Market Date", "Ticker", "Field", "Value", "Display Value", "Source",
        "Source URL", "Source Mode", "As Of Date", "Provider Status",
        "Missing Reason", "Formula", "Formula Version", "Input Fields",
        "Input Sources", "Calculation Status", "Warning",
    ]
    field_specs = [
        ("close", "Closing Price", "yfinance", "OHLCV.Close", "OHLCV"),
        ("volume", "Volume", "yfinance", "OHLCV.Volume", "OHLCV"),
        ("rvol20", "RVOL 20 D", "derived", "Volume / SMA(Volume, 20)", "Volume"),
        ("ema25", "EMA 25", "derived", "EMA(Close, 25, adjust=False)", "Close"),
        ("ema50", "EMA 50", "derived", "EMA(Close, 50, adjust=False)", "Close"),
        ("sma200", "SMA 200", "derived", "SMA(Close, 200)", "Close"),
        ("rsi14", "RSI 14", "derived", "Wilder RSI(14)", "Close"),
        ("macd_line", "MACD Line", "derived", "EMA(Close,12) - EMA(Close,26)", "Close"),
        ("q_vwap", "Current Quarter VWAP", "derived", "Anchored typical-price VWAP", "High, Low, Close, Volume"),
        ("pe_ttm", "Current PE Ratio (TTM)", "workbook", "Market Cap / Net Income TTM", "Market Cap, Net Income TTM"),
        ("pbv", "Current Price to Book Value", "workbook", "Market Cap / Total Equity", "Market Cap, Total Equity"),
        ("roe_ttm", "Return on Equity (TTM)", "workbook", "Net Income TTM / Average Equity", "Net Income TTM, Equity"),
    ]
    for col_idx, label in enumerate(headers, start=1):
        cell = ws.cell(1, col_idx, label)
        cell.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1C3253")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    row_idx = 2
    for row in rows:
        ticker = str(row.get("ticker") or "").upper()
        provider_status = str(row.get("data_status") or "UNKNOWN").upper()
        source_url = f"https://finance.yahoo.com/quote/{ticker}.JK" if ticker else ""
        for key, label, source, formula, inputs in field_specs:
            value = row.get(key)
            missing_value = value is None or value in ("", "-", "N/A") or (
                isinstance(value, float) and pd.isna(value)
            )
            missing_reason = ""
            if missing_value:
                missing_reason = (
                    "insufficient_history"
                    if key in {"rvol20", "ema25", "ema50", "sma200", "rsi14", "macd_line", "q_vwap"}
                    else "field_not_found"
                )
            values = [
                latest_market_day,
                ticker,
                label,
                None if missing_value else value,
                "—" if missing_value else value,
                source,
                source_url if source == "yfinance" else "",
                "point_in_time",
                latest_market_day,
                provider_status,
                missing_reason,
                formula,
                "source-limited-v3",
                inputs,
                "yfinance" if source in {"yfinance", "derived"} else "workbook",
                "MISSING" if missing_value else "OK",
                "Provider values can differ from TradingView chart display." if source == "yfinance" else "",
            ]
            for col_idx, item in enumerate(values, start=1):
                cell = ws.cell(row_idx, col_idx, item)
                cell.font = Font(name="Calibri", size=9, color="1F2937")
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                if row_idx % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F8FAFC")
            row_idx += 1
    widths = [13, 10, 28, 16, 16, 14, 40, 18, 13, 16, 24, 42, 20, 34, 20, 18, 46]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:Q{max(1, row_idx - 1)}"



def build_true_idx_sector_movers_sheet(wb, latest_market_day, all_market_rows):
    """
    IDX Overview builder.

    Kept from prior overview logic:
    - IDX Overview title
    - Market data as-of row with BACKTEST suffix
    - CBOE VIX panel
    - IHSG Seasonal Bias panel

    User update:
    - Delete IDX Sector vs JCI chart.
    - Delete IDX Konglo vs JCI chart.
    - Delete Fear & Greed Overlay vs JCI chart.
    """
    import os
    import tempfile
    import time
    import numpy as np
    import pandas as pd
    import yfinance as yf
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import get_column_letter

    # Native Excel charts only — no matplotlib PNG images.

    if "IDX Overview" in wb.sheetnames:
        wb.remove(wb["IDX Overview"])
    ws = wb.create_sheet("IDX Overview", 0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = None

    # Sheet sizing: keep original compact info panel; charts below it.
    for col, width in {
        "A": 2.5, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18,
        "H": 18, "I": 18, "J": 18, "K": 18, "L": 18, "M": 18
    }.items():
        ws.column_dimensions[col].width = width
    for r in range(1, 115):
        ws.row_dimensions[r].height = 18

    thin = Side(style="thin", color="D9DDE3")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _set(addr, val, bg=None, fc="111827", bold=False, sz=10.0, h="center", border=True, italic=False, wrap=True):
        c = ws[addr]
        c.value = val
        if bg:
            c.fill = PatternFill("solid", fgColor=bg)
        c.font = Font(name="Calibri", size=sz, bold=bold, italic=italic, color=f"FF{fc}")
        c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
        c.border = bdr if border else Border()

    try:
        pretty_date = pd.Timestamp(latest_market_day).strftime("%B %d, %Y")
    except Exception:
        pretty_date = str(latest_market_day)
    mode_suffix = " [BACKTEST]" if ("is_backtest_mode" in globals() and is_backtest_mode()) else ""

    # ── Original requested IDX Overview info block ────────────────────────────
    _set("B2", "IDX Overview", bold=True, sz=14.0, h="left", border=False)
    _set("B3", f"Market Data as of {pretty_date}{mode_suffix}", sz=10.0, h="left", border=False, wrap=False)

    ws.merge_cells("B5:C5")
    _set("B5", "CBOE VIX  ·  Volatility Index", bg="080808", fc="FFFFFF", bold=True, sz=14.0, h="center")
    ws.cell(5, 3).fill = PatternFill("solid", fgColor="080808")
    ws.cell(5, 3).border = bdr

    _set("B6", "Latest VIX", bg="D1D5DB", fc="111827", bold=True, sz=10.0, h="center")
    try:
        # Use the stable prior implementation. Do not depend on the older nested _vix()
        # from an overridden builder; that function is out of scope in this final builder.
        vix_val = _fetch_latest_vix_safe() if "_fetch_latest_vix_safe" in globals() else None
        if vix_val is None or pd.isna(vix_val):
            tk = yf.Ticker("^VIX")
            # VIX = real-time at script run — not clipped to MARKET_DATE
            # It is a live market context indicator, not a historical replay value
            vdf = tk.history(period="5d", interval="1d", auto_adjust=False)
            if vdf is not None and not vdf.empty and "Close" in vdf.columns:
                _s = pd.to_numeric(vdf["Close"], errors="coerce").dropna()
                vix_val = float(_s.iloc[-1]) if not _s.empty else None
    except Exception:
        vix_val = None
    vix_disp = f"{vix_val:.2f}" if (vix_val is not None and not pd.isna(vix_val)) else "N/A"
    _set("C6", vix_disp, bg="F9FAFB", fc="111827", bold=True, sz=12.0, h="center")

    ws["B7"].value = "Note:"
    ws["B7"].font = Font(name="Calibri", size=9, bold=True, color="FF111827")
    ws["B7"].alignment = Alignment(horizontal="left", vertical="center")

    bands = [
        (8,  "< 15  Low / Risk-On",        "D1FAE5", "065F46"),
        (9,  "15-20  Normal",              "DBEAFE", "1E40AF"),
        (10, "20-25  Elevated Caution",    "FEF3C7", "92400E"),
        (11, "≥ 25  Risk-Off / Hedge",     "FEE2E2", "991B1B"),
    ]
    for row_n, txt, bg, fc in bands:
        _set(f"B{row_n}", txt, bg=bg, fc=fc, sz=8.0, h="center")

    ws.row_dimensions[12].height = 8
    ws.row_dimensions[13].height = 30
    ws.row_dimensions[14].height = 52
    ws.merge_cells("B13:G13")
    _set("B13", "IHSG Seasonal Bias  ·  Monthly Cyclical Calendar", bg="1F1F1F", fc="FFFFFF", bold=True, sz=12.0, h="center")
    for col_i in range(3, 8):
        ws.cell(13, col_i).fill = PatternFill("solid", fgColor="1F1F1F")
        ws.cell(13, col_i).border = bdr

    try:
        try:
            seasonal_text = get_seasonal_bias(pd.Timestamp(latest_market_day).month)
        except Exception:
            seasonal_text = "N/A"
    except Exception:
        seasonal_text = "N/A"
    ws.merge_cells("B14:G14")
    _set("B14", seasonal_text, bg="F9FAFB", fc="1F1F1F", sz=10.0, h="left", border=True)
    for col_i in range(3, 8):
        ws.cell(14, col_i).fill = PatternFill("solid", fgColor="F9FAFB")
        ws.cell(14, col_i).border = bdr
        ws.cell(14, col_i).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # User update: delete IDX Sector vs JCI, IDX Konglo vs JCI, and Fear & Greed Overlay vs JCI charts.
    # Keep only the original IDX Overview info panels above.
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return ws

    # ── Data helpers ──────────────────────────────────────────────────────────
    asof = pd.Timestamp(latest_market_day).normalize()

    def _clean_close_series(s):
        if s is None:
            return None
        s = pd.Series(s).copy()
        s.index = pd.to_datetime(s.index)
        try:
            if getattr(s.index, "tz", None) is not None:
                s.index = s.index.tz_localize(None)
        except Exception:
            pass
        s = pd.to_numeric(s, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
        s = s[s.index.normalize() <= asof]
        return s.tail(520) if not s.empty else None

    def _clean_hist(hist):
        if hist is None or not isinstance(hist, pd.DataFrame) or hist.empty:
            return None
        try:
            hh = normalize_history(hist.copy()) if "normalize_history" in globals() else hist.copy()
        except Exception:
            hh = hist.copy()
        if hh is None or hh.empty or "Close" not in hh.columns:
            return None
        return _clean_close_series(hh["Close"])

    def _fetch_close(symbol, period="3y"):
        """Fetch close prices capped at MARKET_DATE."""
        _md  = get_market_date()
        _end = (_md + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        for _ in range(2):
            try:
                raw = yf.download(symbol, period=period, end=_end,
                                  interval="1d", auto_adjust=True,
                                  progress=False, threads=False)
                if raw is not None and not raw.empty:
                    if isinstance(raw.columns, pd.MultiIndex):
                        raw.columns = raw.columns.get_level_values(0)
                    if "Close" in raw.columns:
                        s = _clean_close_series(raw["Close"])
                        if s is not None and not s.empty:
                            s = s[s.index.normalize() <= _md]
                            return s if not s.empty else None
            except Exception:
                pass
            time.sleep(0.20)
        return None

    def _cum_return_pct(s):
        """Price movement comparison: cumulative return (%) from first valid observation."""
        s = _clean_close_series(s)
        if s is None or s.empty:
            return None
        base = float(s.iloc[0])
        if not np.isfinite(base) or base == 0:
            return None
        return (s / base - 1.0) * 100.0

    def _row_get(row, keys, default=None):
        for k in keys:
            try:
                if isinstance(row, dict) and k in row and row.get(k) not in (None, "", "-", "N/A"):
                    return row.get(k)
            except Exception:
                pass
        return default

    def _ticker(row):
        v = _row_get(row, ["ticker", "Ticker", "Code", "code"], "")
        return str(v).upper().strip() if v else ""

    def _sector(row):
        allowed = {
            "IDXFINANCE","IDXTRANS","IDXNONCYC","IDXCYCLIC","IDXINFRA",
            "IDXBASIC","IDXTECHNO","IDXPROPERT","IDXINDUST","IDXENERGY","IDXHEALTH"
        }
        alias = {"IDXPROPERTY": "IDXPROPERT", "PROPERTY": "IDXPROPERT"}
        v = _row_get(row, ["idx_sector", "IDX Sector", "IDXSector", "IDX_Sector", "Sector", "sector", "Sektor"], "")
        s = alias.get(str(v).upper().strip(), str(v).upper().strip())
        return s if s in allowed else None

    def _weight(row):
        for k in ["idx_sector_weight", "IDX Sector Weight", "IDXSectorWeight", "IDX_Sector_Weight", "Weight", "weight"]:
            try:
                v = float(_row_get(row, [k], np.nan))
                if np.isfinite(v) and v > 0:
                    return v
            except Exception:
                pass
        for k in ["mcap", "market_cap", "MarketCap", "marketCap", "Market Cap"]:
            try:
                v = float(_row_get(row, [k], np.nan))
                if np.isfinite(v) and v > 0:
                    return v
            except Exception:
                pass
        return 1.0

    def _row_hist(row):
        for k in ["hist", "history", "daily_hist", "price_hist", "ohlcv"]:
            try:
                h = row.get(k) if isinstance(row, dict) else getattr(row, k, None)
                s = _clean_hist(h)
                if s is not None and not s.empty:
                    return s
            except Exception:
                pass
        tk = _ticker(row)
        return _fetch_close(f"{tk}.JK") if tk else None

    def _weighted_composite(price_map, weight_map):
        """
        Weighted return composite. Output is an index-level series starting at 1.0 internally;
        chart converts it to cumulative return (%), not normalized-to-100.
        """
        if not price_map:
            return None
        px = pd.concat(price_map, axis=1, sort=False).sort_index().dropna(how="all")
        if px.empty:
            return None
        rets = px.pct_change()
        w = pd.Series({c: float(weight_map.get(c, 1.0)) for c in px.columns}).replace([np.inf, -np.inf], np.nan).fillna(1.0).clip(lower=1e-9)
        out = []
        for _, rr in rets.iterrows():
            valid = rr.dropna()
            if valid.empty:
                out.append(np.nan)
                continue
            ww = w.reindex(valid.index).fillna(0)
            if ww.sum() <= 0:
                ww[:] = 1.0
            ww = ww / ww.sum()
            out.append(float((valid * ww).sum()))
        comp = (1.0 + pd.Series(out, index=rets.index).fillna(0)).cumprod()
        return comp.dropna()

    rows = list(all_market_rows or [])
    jci = _fetch_close("^JKSE", period="3y")

    # Sector series: prefer official Yahoo sector index if available; fallback to weighted composite from constituents.
    sector_order = ["IDXFINANCE","IDXTRANS","IDXNONCYC","IDXCYCLIC","IDXINFRA","IDXBASIC","IDXTECHNO","IDXPROPERT","IDXINDUST","IDXENERGY","IDXHEALTH"]
    sector_yahoo_symbols = {
        "IDXFINANCE": ["IDXFINANCE.JK", "^IDXFINANCE", "JKFINANCE.JK", "^JKFINANCE"],
        "IDXTRANS": ["IDXTRANS.JK", "^IDXTRANS", "JKTRANS.JK", "^JKTRANS"],
        "IDXNONCYC": ["IDXNONCYC.JK", "^IDXNONCYC", "JKNONCYC.JK", "^JKNONCYC"],
        "IDXCYCLIC": ["IDXCYCLIC.JK", "^IDXCYCLIC", "JKCYCLIC.JK", "^JKCYCLIC"],
        "IDXINFRA": ["IDXINFRA.JK", "^IDXINFRA", "JKINFRA.JK", "^JKINFRA"],
        "IDXBASIC": ["IDXBASIC.JK", "^IDXBASIC", "JKBASIC.JK", "^JKBASIC"],
        "IDXTECHNO": ["IDXTECHNO.JK", "^IDXTECHNO", "JKTECHNO.JK", "^JKTECHNO"],
        "IDXPROPERT": ["IDXPROPERT.JK", "^IDXPROPERT", "JKPROPERT.JK", "^JKPROPERT"],
        "IDXINDUST": ["IDXINDUST.JK", "^IDXINDUST", "JKINDUST.JK", "^JKINDUST"],
        "IDXENERGY": ["IDXENERGY.JK", "^IDXENERGY", "JKENERGY.JK", "^JKENERGY"],
        "IDXHEALTH": ["IDXHEALTH.JK", "^IDXHEALTH", "JKHEALTH.JK", "^JKHEALTH"],
    }
    sector_price = {s: [] for s in sector_order}
    sector_weight = {s: {} for s in sector_order}
    for row in rows:
        sec = _sector(row)
        tk = _ticker(row)
        if not sec or not tk:
            continue
        s = _row_hist(row)
        if s is None or s.empty:
            continue
        sector_price.setdefault(sec, []).append(s.rename(tk))
        sector_weight.setdefault(sec, {})[tk] = _weight(row)

    sector_series = {}
    for sec in sector_order:
        official = None
        for sym in sector_yahoo_symbols.get(sec, []):
            official = _fetch_close(sym, period="3y")
            if official is not None and not official.empty:
                break
        if official is not None and not official.empty:
            sector_series[sec] = official
        else:
            comp = _weighted_composite(sector_price.get(sec, []), sector_weight.get(sec, {}))
            if comp is not None and not comp.empty:
                sector_series[sec] = comp

    # Konglo composites: weighted return composites from member tickers.
    row_by_ticker = {}
    for row in rows:
        tk = _ticker(row)
        if tk:
            row_by_ticker[tk] = row

    yf_close_cache = {}
    def _ticker_close(tk):
        if tk in row_by_ticker:
            s = _row_hist(row_by_ticker[tk])
            if s is not None and not s.empty:
                return s
        if tk not in yf_close_cache:
            yf_close_cache[tk] = _fetch_close(f"{tk}.JK")
        return yf_close_cache[tk]

    konglo_series = {}
    for group, tickers in KONGLO_GROUPS.items():
        price_map = []
        weight_map = {}
        for tk in tickers:
            s = _ticker_close(tk)
            if s is None or s.empty:
                continue
            price_map.append(s.rename(tk))
            weight_map[tk] = _weight(row_by_ticker[tk]) if tk in row_by_ticker else 1.0
        comp = _weighted_composite(price_map, weight_map)
        if comp is not None and not comp.empty:
            konglo_series[group] = comp

    def _safe_chart_title(s):
        return str(s or "")[:255]

    def _write_price_movement_table(series_dict, benchmark, start_row, start_col):
        """
        Write cumulative price movement (%) table for native Excel chart.
        Table format:
            Date | JCI (^JKSE) | series 1 | series 2 ...
        Values are percentage points, e.g. 12.3 means +12.3%.
        """
        if benchmark is None or benchmark.empty:
            return None
        bench_pct = _cum_return_pct(benchmark)
        if bench_pct is None or bench_pct.empty:
            return None

        valid_series = {"JCI (^JKSE)": bench_pct}
        for name, s in (series_dict or {}).items():
            pct = _cum_return_pct(s)
            if pct is None or pct.empty:
                continue
            common = pct.index.intersection(bench_pct.index)
            if len(common) < 20:
                continue
            valid_series[str(name)] = pct

        # Always create the chart table when JCI is available. If comparison
        # series fail to fetch, the native Excel chart still appears with JCI
        # and the sheet visibly signals that source data was unavailable.
        if len(valid_series) < 1:
            return None

        df = pd.concat(valid_series, axis=1, sort=False).dropna(how="all").tail(520)
        if df.empty:
            return None
        df.index = pd.to_datetime(df.index).tz_localize(None) if getattr(pd.to_datetime(df.index), "tz", None) is not None else pd.to_datetime(df.index)

        # header
        ws.cell(start_row, start_col, "Date")
        for j, col in enumerate(df.columns, start=start_col + 1):
            ws.cell(start_row, j, str(col))

        # data rows
        for i, (dt, row) in enumerate(df.iterrows(), start=start_row + 1):
            c = ws.cell(i, start_col, dt.to_pydatetime() if hasattr(dt, "to_pydatetime") else dt)
            c.number_format = DATE_FORMAT_EXCEL
            for j, val in enumerate(row.values, start=start_col + 1):
                cell = ws.cell(i, j, None if pd.isna(val) else float(val))
                cell.number_format = '0.0"%"'

        return {
            "min_row": start_row,
            "max_row": start_row + len(df),
            "min_col": start_col,
            "max_col": start_col + len(df.columns),
            "n_series": len(df.columns),
        }

    def _add_native_line_chart(table_ref, anchor, title, y_title, width=28, height=11):
        if not table_ref:
            _set(anchor, f"{title}: chart unavailable — insufficient data or yfinance fetch failed", bg="FDE68A", fc="92400E", bold=True, h="center")
            return None

        chart = LineChart()
        chart.title = _safe_chart_title(title)
        chart.style = 13
        chart.height = height
        chart.width = width
        chart.y_axis.title = y_title
        chart.x_axis.title = "Date"
        chart.legend.position = "t"

        data = Reference(
            ws,
            min_col=table_ref["min_col"] + 1,
            max_col=table_ref["max_col"],
            min_row=table_ref["min_row"],
            max_row=table_ref["max_row"],
        )
        cats = Reference(
            ws,
            min_col=table_ref["min_col"],
            min_row=table_ref["min_row"] + 1,
            max_row=table_ref["max_row"],
        )
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        # Smoother Excel display; keep chart editable.
        for s in chart.series:
            try:
                s.graphicalProperties.line.width = 19050
            except Exception:
                pass

        ws.add_chart(chart, anchor)
        return chart

    def _zscore(series, window=252, min_periods=60):
        """
        Rolling z-score used by Fear & Greed:
        (value - rolling mean) / rolling std.
        Defined locally to avoid NameError in the native Excel overview builder.
        """
        try:
            s = pd.to_numeric(series, errors="coerce").astype(float)
            mu = s.rolling(window, min_periods=min_periods).mean()
            sd = s.rolling(window, min_periods=min_periods).std(ddof=0)
            return (s - mu) / sd.replace(0, np.nan)
        except Exception:
            return pd.Series(np.nan, index=getattr(series, "index", None))


    def _write_fear_greed_table(start_row, start_col):
        j = _fetch_close("^JKSE", period="3y")    # end=MARKET_DATE enforced inside _fetch_close
        usd = _fetch_close("USDIDR=X", period="3y")
        if j is None or usd is None or j.empty or usd.empty:
            return None
        df = pd.concat({"JCI Level": j, "USDIDR": usd}, axis=1, sort=False).dropna()
        df = df[df.index.normalize() <= asof]
        if len(df) < 130:
            return None
        df["ma125"] = df["JCI Level"].rolling(125, min_periods=80).mean()
        df["momentum"] = (df["JCI Level"] - df["ma125"]) / df["ma125"] * 100.0
        df["ret"] = np.log(df["JCI Level"] / df["JCI Level"].shift(1))
        df["vol20"] = df["ret"].rolling(20, min_periods=15).std() * np.sqrt(252)
        df["usdidr20"] = (df["USDIDR"] / df["USDIDR"].shift(20) - 1.0) * 100.0
        df["z_mom"] = _zscore(df["momentum"])
        df["z_vol"] = -1.0 * _zscore(df["vol20"])
        df["z_fx"] = -1.0 * _zscore(df["usdidr20"])
        comp_z = df[["z_mom", "z_vol", "z_fx"]].mean(axis=1)
        df["Fear & Greed Index"] = (50.0 + 50.0 * np.tanh(comp_z)).clip(0, 100)
        out = df[["Fear & Greed Index", "JCI Level"]].dropna().tail(520)
        if out.empty:
            return None

        ws.cell(start_row, start_col, "Date")
        ws.cell(start_row, start_col + 1, "Fear & Greed Index")
        ws.cell(start_row, start_col + 2, "JCI Level")

        for i, (dt, row) in enumerate(out.iterrows(), start=start_row + 1):
            c = ws.cell(i, start_col, dt.to_pydatetime() if hasattr(dt, "to_pydatetime") else dt)
            c.number_format = DATE_FORMAT_EXCEL
            ws.cell(i, start_col + 1, float(row["Fear & Greed Index"])).number_format = "0.0"
            ws.cell(i, start_col + 2, float(row["JCI Level"])).number_format = "#,##0"

        return {
            "min_row": start_row,
            "max_row": start_row + len(out),
            "min_col": start_col,
            "max_col": start_col + 2,
        }

    def _add_fear_greed_native_chart(table_ref, anchor):
        if not table_ref:
            _set(anchor, "Fear & Greed Overlay: chart unavailable — insufficient data or yfinance fetch failed", bg="FDE68A", fc="92400E", bold=True, h="center")
            return None

        c1 = LineChart()
        c1.title = "Fear & Greed Overlay vs JCI"
        c1.style = 13
        c1.height = 11
        c1.width = 28
        c1.y_axis.title = "Fear & Greed Index (0-100)"
        c1.x_axis.title = "Date"
        c1.y_axis.scaling.min = 0
        c1.y_axis.scaling.max = 100
        c1.legend.position = "t"

        cats = Reference(ws, min_col=table_ref["min_col"], min_row=table_ref["min_row"] + 1, max_row=table_ref["max_row"])
        fg_data = Reference(ws, min_col=table_ref["min_col"] + 1, min_row=table_ref["min_row"], max_row=table_ref["max_row"])
        c1.add_data(fg_data, titles_from_data=True)
        c1.set_categories(cats)

        c2 = LineChart()
        c2.y_axis.axId = 200
        c2.y_axis.title = "JCI Level"
        c2.y_axis.crosses = "max"
        jci_data = Reference(ws, min_col=table_ref["min_col"] + 2, min_row=table_ref["min_row"], max_row=table_ref["max_row"])
        c2.add_data(jci_data, titles_from_data=True)
        c2.set_categories(cats)

        c1 += c2
        for s in c1.series:
            try:
                s.graphicalProperties.line.width = 19050
            except Exception:
                pass

        ws.add_chart(c1, anchor)
        return c1

    # ── Native Excel chart output, not matplotlib pictures ────────────────────
    # Keep raw chart data inside the workbook so users can edit the Excel charts.
    data_col = 15  # column O
    sector_tbl = _write_price_movement_table(sector_series, jci, start_row=2, start_col=data_col)
    konglo_tbl = _write_price_movement_table(konglo_series, jci, start_row=540, start_col=data_col)
    fg_tbl = _write_fear_greed_table(start_row=1080, start_col=data_col)

    # Hide helper data columns but keep them available for chart editing.
    for col_i in range(data_col, data_col + 30):
        ws.column_dimensions[get_column_letter(col_i)].hidden = True

    _add_native_line_chart(
        sector_tbl,
        "B17",
        "IDX Sector vs JCI",
        "Price movement since start (%)",
        width=28,
        height=11,
    )
    _add_native_line_chart(
        konglo_tbl,
        "B44",
        "IDX Konglo vs JCI",
        "Price movement since start (%)",
        width=28,
        height=11,
    )
    _add_fear_greed_native_chart(fg_tbl, "B71")

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return ws

if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        import sys as _sys
        print("\n[FATAL ERROR]")
        traceback.print_exc()
        # Only wait for a keypress when running interactively; in CI (no TTY) an
        # input() prompt raises EOFError and masks the real error. Always exit 1.
        if _sys.stdin is not None and _sys.stdin.isatty():
            input("\nPress Enter to exit...")
        _sys.exit(1)
