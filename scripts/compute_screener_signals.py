"""Computes the four Screener filter signals directly from the already-
published OHLCV archive (docs/data/ohlcv/<date>/<TICKER>.json) -- no
yfinance re-fetch, so this runs in seconds against data the daily pipeline
already downloaded, rather than requiring a full 1.5-3h rerun.

Writes docs/data/dates/<date>/screener_signals.json, one record per roster
ticker:
  breakIbhIbl        -- today's close breaks above the monthly Initial
                         Balance High, after yesterday's close sat roughly
                         mid-band (real published field, technical.json's
                         marketProfile.ibh/ibl -- IDX_Screener.py's existing
                         monthly-IB computation, not re-derived here).
  rsiDivBullish       -- RSI(14, Wilder-smoothed -- TradingView's own RSI
                         is always RMA-based at its core, EMA is only ever
                         an optional secondary smoothing line on top, never
                         the base) regular bullish divergence: a literal
                         port of TradingView's own built-in "RSI" study's
                         divergence logic. LOW makes a lower low while RSI
                         makes a higher low, comparing the two most recent
                         confirmed pivots from a plain 5-bars-each-side
                         centered pivot scan (`ta.pivotlow` with
                         lookbackLeft=lookbackRight=5), gated only on the
                         gap between them (5-60 bars) -- no RSI-value
                         threshold, no clustering; both of those were this
                         function's own earlier invention, not part of the
                         real pattern, and got replaced after directly
                         reproducing a live TradingView-vs-screener mismatch
                         the user reported (BUKK, 2026-09-11: TradingView's
                         own chart marks its last divergence at 14 Aug with
                         nothing new since; the old version here reported a
                         spurious 09 Sep). Reversal-style: both pivots must
                         be confirmed swing lows, so confirmation lags the
                         more recent pivot by swing_window bars (pivot date
                         shown is that more recent pivot, not today).
  rsiDivHiddenBullish -- same RSI(14), same pivot scan, same shared 5-60 bar
                         gate as rsiDivBullish above (TradingView uses one
                         shared pivot-range config for every divergence
                         type) -- the exact mirror of it: LOW makes a
                         Higher Low while RSI makes a Lower Low, comparing
                         the same two most recent confirmed pivots, no
                         RSI-value gate. An earlier version of this signal
                         had never actually been ported to match
                         rsiDivBullish's TradingView-parity fix -- it kept
                         a different, invented shape (2-bar pivot window, a
                         50-70 RSI band gate on the earlier pivot,
                         clustering) left over from before that fix, caught
                         via a direct report that "the pivoting is still
                         not true" even after rsiDivBullish itself was
                         fixed. Re-validated against the same real example
                         used before (BEST 13 Aug -> 26 Aug) to confirm the
                         rewrite doesn't regress it. Reversal-style like
                         rsiDivBullish: both pivots must be confirmed swing
                         lows, so confirmation lags the more recent pivot by
                         swing_window bars (pivot date shown is that more
                         recent pivot, not today). Price basis changed from
                         CLOSE to LOW when this got re-aligned to the
                         user-supplied Pine script below (that script's own
                         `hiddenBull` uses `low[right]` uniformly, not
                         Close) -- see regular_bearish_divergence()'s
                         docstring for the full context on that script.
  rsiDivBearish       -- Regular Bearish RSI(14, Wilder) divergence -- the
                         pivot-high mirror of rsiDivBullish: HIGH makes a
                         Higher High while RSI makes a Lower High, same
                         5-bars-each-side pivot scan and 5-60 bar gap, no
                         RSI-value threshold. Ported from a user-supplied
                         Pine Script v6 divergence indicator (rsiLen=14,
                         pivot left/right=5/5, strict comparison) rather
                         than reproduced from scratch -- see
                         regular_bearish_divergence()'s own docstring.
  rsiDivHiddenBearish -- Hidden Bearish RSI(14, Wilder) divergence -- the
                         exact mirror of rsiDivBearish: HIGH makes a Lower
                         High while RSI makes a Higher High, same pivot
                         scan and gate. Same Pine script source as
                         rsiDivBearish.
  breakSma200         -- today's close crosses above SMA200 (yesterday's
                         close was at or below it).
  emaGoldenCross      -- EMA25 crosses above EMA50 today (yesterday EMA25
                         was at or below EMA50).
  stochRsiGoldenCross -- Stochastic RSI (RSI length 10, Stochastic length
                         10, K 3, D 3) %K crosses above %D today while
                         RSI(10) < 30 (oversold).
  nearPqM1 / nearPqM2 -- close within NEAR_PCT of the previous quarter's
  nearPyM1 / nearPyM2    anchored-VWAP -1sigma/-2sigma band (previous year
                         for the PY pair) -- same hlc3*volume anchored-VWAP
                         formula as lib/indicators/anchoredVwap.ts.

Real inputs only -- a ticker with too little history for a given signal
gets `false`/`null` for it, never a guess.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

OHLCV_DIR = ROOT / "docs" / "data" / "ohlcv"
DATES_DIR = ROOT / "docs" / "data" / "dates"
LISTED_PATH = ROOT / "data_sources" / "idx-listed.json"

NEAR_PCT = 0.01  # "near" a VWAP sigma level = within 1% of it, either side


def rsi_ema(series: pd.Series, period: int = 10) -> pd.Series:
    """RSI with EMA-smoothed up/down averages (not Wilder's RMA) -- matches
    a "RSI Smoothing Type: EMA" chart setting."""
    delta = series.diff()
    up = delta.clip(lower=0).ewm(span=period, adjust=False).mean()
    dn = (-delta.clip(upper=0)).ewm(span=period, adjust=False).mean()
    rs = up / dn.replace(0, np.nan)
    return 100 - (100 / (1 + rs))


def rsi_wilder(series: pd.Series, period: int = 10) -> pd.Series:
    delta = series.diff()
    up = delta.clip(lower=0).ewm(alpha=1.0 / period, adjust=False).mean()
    dn = (-delta.clip(upper=0)).ewm(alpha=1.0 / period, adjust=False).mean()
    rs = up / dn.replace(0, np.nan)
    return 100 - (100 / (1 + rs))


def stoch_of(series: pd.Series, length: int = 10, k_smooth: int = 3, d_smooth: int = 3):
    lo = series.rolling(length).min()
    hi = series.rolling(length).max()
    rng = hi - lo
    raw = pd.Series(np.where(rng > 0, 100 * (series - lo) / rng, np.nan), index=series.index)
    k = raw.rolling(k_smooth).mean()
    d = k.rolling(d_smooth).mean()
    return k, d


def anchored_vwap_band(hist: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp):
    seg = hist.loc[(hist.index >= start) & (hist.index <= end)]
    if seg.empty:
        return None
    src = (seg["High"] + seg["Low"] + seg["Close"]) / 3
    vol = seg["Volume"].fillna(0).astype(float)
    cum_v = vol.sum()
    if cum_v <= 0:
        return None
    vwap = float((src * vol).sum() / cum_v)
    var = float(((src - vwap) ** 2 * vol).sum() / cum_v)
    sd = math.sqrt(max(var, 0))
    return {"vwap": vwap, "l1": vwap - sd, "l2": vwap - 2 * sd}


def quarter_start(ts: pd.Timestamp) -> pd.Timestamp:
    q = (ts.month - 1) // 3
    return pd.Timestamp(ts.year, q * 3 + 1, 1)


def prev_quarter_bounds(ts: pd.Timestamp):
    start = quarter_start(ts)
    prev_end = start - pd.Timedelta(days=1)
    prev_start = quarter_start(prev_end)
    return prev_start, prev_end


def prev_year_bounds(ts: pd.Timestamp):
    y = ts.year - 1
    return pd.Timestamp(y, 1, 1), pd.Timestamp(y, 12, 31)


def near(price: float | None, level: float | None) -> bool:
    if price is None or level is None or not math.isfinite(level) or level <= 0:
        return False
    return abs(price - level) / level <= NEAR_PCT


def available_ohlcv_snapshots() -> list[str]:
    """Every date with its own full OHLCV archive, sorted ascending. Most
    of docs/data/dates/<date>/ are lighter snapshots (no docs/data/ohlcv/
    <date>/ folder at all) -- see load_hist()'s snapshot-selection docstring."""
    return sorted(p.name for p in OHLCV_DIR.iterdir() if p.is_dir())


def load_hist(ticker: str, snapshot_date: str, as_of_date: str | None = None) -> pd.DataFrame | None:
    """Load `ticker`'s OHLCV from the `snapshot_date` archive, trimmed to
    bars on or before `as_of_date` (defaults to `snapshot_date` itself).

    A full OHLCV snapshot's `rows` already holds a ticker's entire trailing
    history up to that snapshot's date, not just that one day -- so a date
    with no dedicated snapshot of its own (most of docs/data/dates/, ~90%
    of the archive) can still get a real, correctly-timed signal by reading
    the nearest LATER snapshot and trimming it back down to the actual
    target date, rather than being skipped for lack of a same-named
    archive folder."""
    path = OHLCV_DIR / snapshot_date / f"{ticker}.json"
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    rows = payload.get("rows") or []
    if len(rows) < 30:
        return None
    df = pd.DataFrame(rows)
    df["Date"] = pd.to_datetime(df["date"])
    df = df.set_index("Date").sort_index()
    df = df.rename(columns={"open": "Open", "high": "High", "low": "Low", "close": "Close", "volume": "Volume"})
    df = df[["Open", "High", "Low", "Close", "Volume"]].astype(float)
    if as_of_date is not None and as_of_date != snapshot_date:
        df = df.loc[df.index <= pd.Timestamp(as_of_date)]
    if len(df) < 30:
        return None
    return df


def ib_break(hist: pd.DataFrame, ibh: float | None, ibl: float | None) -> bool:
    if ibh is None or ibl is None or ibh <= ibl or len(hist) < 2:
        return False
    today = hist.iloc[-1]
    yday = hist.iloc[-2]
    band = ibh - ibl
    mid_lo, mid_hi = ibl + band * 0.25, ibl + band * 0.75
    was_middling = mid_lo <= yday["Close"] <= mid_hi
    breaks_today = today["Close"] > ibh and yday["Close"] <= ibh
    return bool(was_middling and breaks_today)


def regular_bullish_divergence(hist: pd.DataFrame, lookback=150, swing_window=5) -> dict | None:
    """Regular Bullish RSI(14, Wilder) divergence -- a literal port of a
    user-supplied Pine Script v6 divergence indicator (rsiLen=14, pivot
    left/right=5/5, strict HH/LL comparison): `ta.pivotlow(rsi, 5, 5)`,
    compare the two most recent confirmed pivots, LOW makes a lower low
    while RSI makes a higher low, no RSI-value gate at all. Length bumped
    from 10 to 14 to match that script's own `rsiLen` input exactly
    (previously matched TradingView's built-in "RSI" study, whose own
    default is also 14 -- so this restores parity with both).

    No gap/age gate between the two pivots -- an earlier version of this
    function rejected pairs more than 60 bars apart or whose most recent
    pivot was more than 20 bars old, neither of which exists in the actual
    Pine script (confirmed against lib/indicators/divergence.ts's own
    line-by-line port of the same script, which chains every consecutive
    pivot pair with no such constraint at all). That was this function's
    own invented heuristic, exactly like the earlier oversold-gate/
    EMA-basis mistakes below -- caught the same way, by the screener and
    the chart disagreeing on a real ticker the chart (Pine-faithful) got
    right and the screener (gated) missed or mistimed.

    Two corrections from an earlier version of this function, both found
    by directly reproducing a real, exact mismatch the user reported
    between this screener and a live TradingView chart (BUKK, 2026-09-11:
    TradingView's own divergence marker sits at 14 Aug with nothing new
    since; this function previously reported a spurious 09 Sep match):

    1. RSI basis was rsi_ema(10, EMA-smoothed) -- WRONG. TradingView's own
       RSI (any length) always uses `ta.rma` (Wilder/RMA smoothing) for
       the base line itself; EMA is only ever an optional secondary
       *smoothing* line drawn on top, never the RSI values TradingView's
       divergence logic (or its plotted RSI value) is computed from.
       Switched to rsi_wilder() -- confirmed this alone starts
       reproducing TradingView's real pivot dates.
    2. Pivot window was 2 bars each side with an RSI<30 oversold gate and
       a lifecycle-clustering step -- none of that exists in TradingView's
       own script. It uses a plain 5-bars-each-side centered pivot test
       (stronger, so far fewer/less noisy candidates qualify than a
       2-bar test does) and gates purely on the gap between the two most
       recent confirmed pivots (5-60 bars) -- no RSI-value threshold, no
       clustering. A "regular bullish divergence" in the standard
       definition can happen with RSI anywhere, not only when it dipped
       below 30; requiring that was this function's own invention, not
       part of the actual pattern.

    RSI(10) was verified directly against TradingView's live chart for
    BUKK (2026-09-11): reproduced the exact 29 Jul -> 14 Aug pivot pair,
    the same one TradingView itself draws, with nothing newer confirmed
    since. The length-14 pivot dates from this later change have not been
    independently re-checked against a live TradingView chart -- the pivot
    *method* is identical, only which bars qualify as pivots can shift with
    a longer RSI length, so treat this claim as carried over, not re-proven.

    i2 is the most recent CONFIRMED pivot (needing swing_window bars after
    it, same as i1) -- not forced to today, matching TradingView's own
    plot (offset=-lookbackRight: the marker is drawn AT the pivot bar,
    which is necessarily swing_window bars before whatever bar the
    pattern was actually confirmed on)."""
    if hist is None or hist.empty or len(hist) < 25:
        return None
    df = hist.tail(lookback)
    low = df["Low"].astype(float).values
    r = rsi_wilder(df["Close"].astype(float), 14).values

    pivots = _swing_low_positions(r, swing_window)
    if len(pivots) < 2:
        return None
    i1, i2 = pivots[-2], pivots[-1]
    p1, p2 = float(low[i1]), float(low[i2])
    r1, r2 = float(r[i1]), float(r[i2])
    if any(np.isnan(x) for x in (p1, p2, r1, r2)):
        return None
    if not (p2 < p1 and r2 > r1):
        return None
    return {"i2": i2, "ref2_date": df.index[i2].strftime("%d %b '%y"), "p1": p1, "p2": p2, "r1": r1, "r2": r2}


def regular_bullish_divergence_today(hist: pd.DataFrame) -> tuple[bool, str | None]:
    """Regular Bullish divergence newly confirmed today -- "confirmed
    today" means the signal reads Bullish on today's full history but
    didn't (or pointed at a different pivot pair) on yesterday's, i.e. it
    just became visible with today's bar providing the confirming swing
    point for i2. Returns (confirmed_today, pivot_date) -- pivot_date is
    i2, the more recent of the two pivots (which is NOT today itself: see
    the module docstring on why confirmation lags the pivot by
    swing_window bars for this reversal-style detector)."""
    today = regular_bullish_divergence(hist)
    if today is None:
        return False, None
    yday = regular_bullish_divergence(hist.iloc[:-1])
    newly_confirmed = yday is None or yday["ref2_date"] != today["ref2_date"]
    if not newly_confirmed:
        return False, None
    return True, today["ref2_date"]


def _swing_low_positions(vals: np.ndarray, w: int) -> list[int]:
    pos = []
    for i in range(w, len(vals) - w):
        c = vals[i]
        if np.isnan(c):
            continue
        left, right = vals[i - w:i], vals[i + 1:i + w + 1]
        if np.all(np.isnan(left)) or np.all(np.isnan(right)):
            continue
        if c <= np.nanmin(left) and c <= np.nanmin(right):
            pos.append(i)
    return pos


def _swing_high_positions(vals: np.ndarray, w: int) -> list[int]:
    """Pivot-high mirror of _swing_low_positions() -- a literal
    `ta.pivothigh(vals, w, w)` port, same plateau-tolerant (`>=`) test.
    Drives regular_bearish_divergence()/hidden_bearish_divergence() the
    same way _swing_low_positions() drives the bullish pair above."""
    pos = []
    for i in range(w, len(vals) - w):
        c = vals[i]
        if np.isnan(c):
            continue
        left, right = vals[i - w:i], vals[i + 1:i + w + 1]
        if np.all(np.isnan(left)) or np.all(np.isnan(right)):
            continue
        if c >= np.nanmax(left) and c >= np.nanmax(right):
            pos.append(i)
    return pos


def hidden_bullish_divergence(hist: pd.DataFrame, lookback=150, swing_window=5) -> dict | None:
    """Hidden Bullish RSI(14, Wilder) divergence -- the exact mirror of
    regular_bullish_divergence(): same `ta.pivotlow(rsi, 5, 5)` pivot scan,
    same "compare the two most recent confirmed pivots" logic, no gap/age
    gate between them (see regular_bullish_divergence()'s docstring on why
    that gate was removed -- it was invented, not part of the real script).
    Only the price/RSI comparison
    direction flips from regular's: price makes a HIGHER low while RSI
    makes a LOWER low -- an uptrend-continuation pattern, the opposite of
    regular's reversal pattern.

    Price basis is LOW here, same as regular's -- changed from an earlier
    CLOSE basis (see below) when this got re-aligned to a user-supplied
    Pine Script v6 divergence indicator: that script's own `hiddenBull`
    condition is `isHigher(currPriceLow, prevPriceLow) and
    isLower(currRsiLow, prevRsiLow)` where `currPriceLow = low[right]` --
    LOW uniformly for every one of its four divergence types, never Close.
    The prior CLOSE basis below was verified against TradingView's own
    BUILT-IN "RSI" indicator specifically (a different, unrelated
    reference), which may genuinely use Close for its hidden divergences --
    but the user explicitly asked for the newer script's own logic here,
    which doesn't, so this now follows that script instead of the older
    reference.

    This replaces an earlier, structurally different version of this
    function that never actually got the same TradingView-parity fix
    regular_bullish_divergence() did: a 2-bar (not 5-bar) pivot window, an
    invented 50-70 RSI band gate on the earlier pivot, and a clustering
    step ported from IDX_Screener.py -- none of which exist in
    TradingView's real algorithm. That version was caught still producing
    non-TV-shaped results (a direct report: "the pivoting is still not
    true") even after regular_bullish_divergence() was fixed, because it
    had never been rewritten to match.

    RSI(10), CLOSE-basis was verified directly against the same real
    example used to validate the old version: BEST 2026-08-13 (RSI 56.8,
    close 109) -> 2026-08-26 (RSI 54.2, close 117) -- price Higher Low +
    RSI Lower Low, confirmed on 26 Aug with that earlier literal-port
    logic. Length bumped to 14 and price basis switched Close->Low since,
    per the docstring above; the RSI(10)/Close numbers above are carried
    over from that earlier check, not re-verified at length 14 / Low basis
    (BEST's own Low happens to track its Close closely enough on both of
    those two dates that this particular example would likely still
    confirm the same pivot pair either way, but that hasn't been checked).

    No RSI-band gate on the earlier pivot (unlike the old version) means a
    case like PYFA's own 2026-08-26 pivot (r1 in the low 40s -- an
    already-weak trend, not a "healthy" one taking a shallow dip) now also
    qualifies. That's intentional: TradingView's real algorithm has no
    RSI-value gate on either divergence type, so excluding PYFA was this
    function's own invented heuristic, not part of the actual pattern --
    the same lesson regular_bullish_divergence()'s fix already established
    for the oversold-gate it used to have.

    i2 is the most recent CONFIRMED pivot (needing swing_window bars after
    it), not forced to today -- same reasoning as
    regular_bullish_divergence(): the marker sits at the pivot bar itself
    (TradingView's offset=-lookbackRight), necessarily swing_window bars
    before whatever bar the pattern was actually confirmed on."""
    if hist is None or hist.empty or len(hist) < 25:
        return None
    df = hist.tail(lookback)
    low = df["Low"].astype(float).values
    r = rsi_wilder(df["Close"].astype(float), 14).values

    pivots = _swing_low_positions(r, swing_window)
    if len(pivots) < 2:
        return None
    i1, i2 = pivots[-2], pivots[-1]
    p1, p2 = float(low[i1]), float(low[i2])
    r1, r2 = float(r[i1]), float(r[i2])
    if any(np.isnan(x) for x in (p1, p2, r1, r2)):
        return None
    if not (p2 > p1 and r2 < r1):
        return None
    return {"i2": i2, "ref2_date": df.index[i2].strftime("%d %b '%y"), "p1": p1, "p2": p2, "r1": r1, "r2": r2}


def hidden_bullish_divergence_today(hist: pd.DataFrame) -> tuple[bool, str | None]:
    """Hidden bullish divergence newly confirmed today -- same "confirmed
    today" diff-vs-yesterday / ref2_date-changed approach as
    regular_bullish_divergence_today(); see its docstring. Returns
    (confirmed_today, pivot_date) -- pivot_date is i2, the more recent of
    the two pivots (not today itself: confirmation lags the pivot by
    swing_window bars, same as the regular signal)."""
    today = hidden_bullish_divergence(hist)
    if today is None:
        return False, None
    yday = hidden_bullish_divergence(hist.iloc[:-1])
    newly_confirmed = yday is None or yday["ref2_date"] != today["ref2_date"]
    if not newly_confirmed:
        return False, None
    return True, today["ref2_date"]


def regular_bearish_divergence(hist: pd.DataFrame, lookback=150, swing_window=5) -> dict | None:
    """Regular Bearish RSI(14, Wilder) divergence -- the pivot-HIGH mirror
    of regular_bullish_divergence(): a literal port of a user-supplied
    Pine Script v6 divergence indicator's `regularBear` condition
    (`ta.pivothigh(rsi, 5, 5)`, compare the two most recent confirmed
    pivots, HIGH makes a higher high while RSI makes a lower high, no
    gap/age gate between them and no RSI-value threshold -- see
    regular_bullish_divergence()'s docstring on why the gate was removed).
    Unlike
    rsiDivBullish/rsiDivHiddenBullish above (which started life matching
    TradingView's own built-in RSI study and were only later re-aligned to
    this script), this pair is a first-time, direct port straight from the
    script the user provided -- there's no earlier TradingView-parity
    verification history to carry over or correct here.

    Price basis is HIGH for both regular and hidden bearish -- the script's
    own `currPriceHigh = high[right]` is shared by both `regularBear` and
    `hiddenBear`, same as LOW is shared by both bullish types (see
    hidden_bullish_divergence()'s docstring on why bullish uses LOW, not
    Close, for the same reason).

    i2 is the most recent CONFIRMED pivot (needing swing_window bars after
    it), not forced to today -- same reversal-style confirmation-lag
    reasoning as every other divergence function in this file."""
    if hist is None or hist.empty or len(hist) < 25:
        return None
    df = hist.tail(lookback)
    high = df["High"].astype(float).values
    r = rsi_wilder(df["Close"].astype(float), 14).values

    pivots = _swing_high_positions(r, swing_window)
    if len(pivots) < 2:
        return None
    i1, i2 = pivots[-2], pivots[-1]
    p1, p2 = float(high[i1]), float(high[i2])
    r1, r2 = float(r[i1]), float(r[i2])
    if any(np.isnan(x) for x in (p1, p2, r1, r2)):
        return None
    if not (p2 > p1 and r2 < r1):
        return None
    return {"i2": i2, "ref2_date": df.index[i2].strftime("%d %b '%y"), "p1": p1, "p2": p2, "r1": r1, "r2": r2}


def regular_bearish_divergence_today(hist: pd.DataFrame) -> tuple[bool, str | None]:
    """Regular Bearish divergence newly confirmed today -- same
    "confirmed today" diff-vs-yesterday / ref2_date-changed approach as
    regular_bullish_divergence_today(); see its docstring."""
    today = regular_bearish_divergence(hist)
    if today is None:
        return False, None
    yday = regular_bearish_divergence(hist.iloc[:-1])
    newly_confirmed = yday is None or yday["ref2_date"] != today["ref2_date"]
    if not newly_confirmed:
        return False, None
    return True, today["ref2_date"]


def hidden_bearish_divergence(hist: pd.DataFrame, lookback=150, swing_window=5) -> dict | None:
    """Hidden Bearish RSI(14, Wilder) divergence -- the exact mirror of
    regular_bearish_divergence(): same pivot-high scan, no gap/age gate,
    only the comparison direction flips -- HIGH makes a LOWER high while
    RSI makes a HIGHER high (a downtrend-continuation pattern, the
    opposite of regular's reversal pattern). Same Pine-script source and
    HIGH price basis as regular_bearish_divergence() -- see its docstring."""
    if hist is None or hist.empty or len(hist) < 25:
        return None
    df = hist.tail(lookback)
    high = df["High"].astype(float).values
    r = rsi_wilder(df["Close"].astype(float), 14).values

    pivots = _swing_high_positions(r, swing_window)
    if len(pivots) < 2:
        return None
    i1, i2 = pivots[-2], pivots[-1]
    p1, p2 = float(high[i1]), float(high[i2])
    r1, r2 = float(r[i1]), float(r[i2])
    if any(np.isnan(x) for x in (p1, p2, r1, r2)):
        return None
    if not (p2 < p1 and r2 > r1):
        return None
    return {"i2": i2, "ref2_date": df.index[i2].strftime("%d %b '%y"), "p1": p1, "p2": p2, "r1": r1, "r2": r2}


def hidden_bearish_divergence_today(hist: pd.DataFrame) -> tuple[bool, str | None]:
    """Hidden Bearish divergence newly confirmed today -- same "confirmed
    today" diff-vs-yesterday / ref2_date-changed approach as
    regular_bullish_divergence_today(); see its docstring."""
    today = hidden_bearish_divergence(hist)
    if today is None:
        return False, None
    yday = hidden_bearish_divergence(hist.iloc[:-1])
    newly_confirmed = yday is None or yday["ref2_date"] != today["ref2_date"]
    if not newly_confirmed:
        return False, None
    return True, today["ref2_date"]


def stoch_rsi_golden_cross_today(hist: pd.DataFrame) -> bool:
    """K crosses above D today while the cross originates from the Stoch
    RSI's own oversold band (K and D both below 20 just before crossing) --
    "oversold" here means the STOCHASTIC-OF-RSI reading itself, the same
    dashed 20/80 bands a Stoch RSI chart draws, not the underlying RSI(10)
    value. Verified against BEST 2026-07-30: RSI(10) was 48.9 (not <30) but
    K/D were 8.1/16.6 just before crossing to 19.3/13.3 -- a real golden
    cross a raw-RSI<30 gate would have missed."""
    r = rsi_wilder(hist["Close"], 10)
    k, d = stoch_of(r, length=10, k_smooth=3, d_smooth=3)
    if len(k) < 2 or pd.isna(k.iloc[-1]) or pd.isna(d.iloc[-1]) or pd.isna(k.iloc[-2]) or pd.isna(d.iloc[-2]):
        return False
    crossed = k.iloc[-2] <= d.iloc[-2] and k.iloc[-1] > d.iloc[-1]
    oversold = k.iloc[-2] < 20 and d.iloc[-2] < 20
    return bool(crossed and oversold)


def stoch_rsi_oversold_today(hist: pd.DataFrame) -> bool:
    """Stoch RSI sitting in its own oversold band -- K AND D both under 20
    today, K still at or below D (hasn't crossed yet) but converging on it
    (today's K-D gap narrower than yesterday's, i.e. K visibly closing in on
    a cross) -- an earlier, higher-lead-time companion to
    stoch_rsi_golden_cross_today(). By construction a confirmed golden
    cross can only be seen after the bounce that produces it has already
    moved price (there is no way to see tomorrow's cross today), so this
    flags the watch-for-a-turn state instead of the confirmed turn itself.

    D<20 is required, not just K<20: K alone dipping under 20 for a bar or
    two while D is still elevated (e.g. K=15, D=45) is a fast wiggle inside
    an otherwise-elevated Stoch RSI, not a genuine oversold reading -- both
    lines need to be down in the band together. Real counterexample this
    fixes: a ticker whose K/D actually read ~57/~70 (nowhere near oversold)
    was passing the old K-only check on an earlier bar where K alone had
    briefly dipped under 20.

    Unlike the golden-cross signal, this is a CURRENT-STATE flag (true on
    every day the condition holds, like nearPq*/nearPy*), not a one-day
    'today' event -- a ticker can sit oversold for many sessions before
    (or without ever) crossing, so this fires far more often and with much
    lower precision than the confirmed cross; a watchlist signal, not an
    entry trigger."""
    r = rsi_wilder(hist["Close"], 10)
    k, d = stoch_of(r, length=10, k_smooth=3, d_smooth=3)
    if len(k) < 2 or pd.isna(k.iloc[-1]) or pd.isna(d.iloc[-1]) or pd.isna(k.iloc[-2]) or pd.isna(d.iloc[-2]):
        return False
    both_oversold = k.iloc[-1] < 20 and d.iloc[-1] < 20
    not_yet_crossed = k.iloc[-1] <= d.iloc[-1]
    converging = (d.iloc[-1] - k.iloc[-1]) < (d.iloc[-2] - k.iloc[-2])
    return bool(both_oversold and not_yet_crossed and converging)


def break_sma200_today(hist: pd.DataFrame) -> bool:
    """Today's close crosses above SMA200 -- yesterday's close was at or
    below it, today's is above. A standard long-term trend-change signal.
    False (not a guess) for a ticker with under 200 bars of history, same
    as every other signal here."""
    close = hist["Close"]
    s200 = close.rolling(200).mean()
    if len(s200) < 2 or pd.isna(s200.iloc[-1]) or pd.isna(s200.iloc[-2]):
        return False
    return bool(close.iloc[-2] <= s200.iloc[-2] and close.iloc[-1] > s200.iloc[-1])


def ema_golden_cross_today(hist: pd.DataFrame) -> bool:
    """EMA25 crosses above EMA50 today -- yesterday EMA25 was at or below
    EMA50, today it's above. Matches the site's own chart default overlay
    (EMA 25 / EMA 50, see lib/data/chartStudies.ts)."""
    close = hist["Close"]
    e25 = close.ewm(span=25, adjust=False).mean()
    e50 = close.ewm(span=50, adjust=False).mean()
    if len(e25) < 2 or pd.isna(e25.iloc[-1]) or pd.isna(e50.iloc[-1]) or pd.isna(e25.iloc[-2]) or pd.isna(e50.iloc[-2]):
        return False
    return bool(e25.iloc[-2] <= e50.iloc[-2] and e25.iloc[-1] > e50.iloc[-1])


def monday_range(hist: pd.DataFrame) -> dict | None:
    """High/low of the current week's first published trading session
    (Monday, or the first session of the week when Monday itself was a
    holiday) plus where today's close sits vs that range -- same
    week-start rule as lib/valuation/priceLevels.ts's buildMondayRangeLevels
    (a day-of-week reset: this bar's weekday <= the previous bar's)."""
    if hist is None or len(hist) < 2:
        return None
    dows = hist.index.dayofweek.to_numpy()
    start = 0
    for i in range(1, len(dows)):
        if dows[i] <= dows[i - 1]:
            start = i
    row = hist.iloc[start]
    close = float(hist["Close"].iloc[-1])
    hi, lo = float(row["High"]), float(row["Low"])
    status = "Above" if close > hi else "Below" if close < lo else "Within"
    return {"high": hi, "low": lo, "date": hist.index[start].strftime("%Y-%m-%d"), "status": status}


def near_vwap_flags(hist: pd.DataFrame, close: float) -> dict:
    last_date = hist.index[-1]
    pq_start, pq_end = prev_quarter_bounds(last_date)
    py_start, py_end = prev_year_bounds(last_date)
    pq = anchored_vwap_band(hist, pq_start, pq_end)
    py = anchored_vwap_band(hist, py_start, py_end)
    return {
        "nearPqM1": near(close, pq["l1"] if pq else None),
        "nearPqM2": near(close, pq["l2"] if pq else None),
        "nearPyM1": near(close, py["l1"] if py else None),
        "nearPyM2": near(close, py["l2"] if py else None),
    }


def vwap_reading(band: dict | None, close: float) -> dict | None:
    """{vwap, sigma} for the Screener table's Previous QVWAP/YVWAP columns --
    sigma is how many standard deviations `close` sits from the anchored
    VWAP (band["l1"] is -1sigma, so vwap-l1 is one sigma). None when the
    anchor window has no data (band is None) or was degenerate (sd is 0,
    e.g. a single print with itself as both high and low)."""
    if band is None:
        return None
    sd = band["vwap"] - band["l1"]
    return {"vwap": band["vwap"], "sigma": (close - band["vwap"]) / sd if sd > 0 else None}


def pq_py_vwap_readings(hist: pd.DataFrame, close: float) -> dict:
    """Raw Previous-Quarter/Previous-Year anchored-VWAP readings, for
    archives whose technical.json predates the vwapProfiles field (see
    lib/data/screenerUniverse.ts's fallback for why this is needed at all --
    it reads this only when technical.json itself doesn't have the real
    reading)."""
    last_date = hist.index[-1]
    pq_start, pq_end = prev_quarter_bounds(last_date)
    py_start, py_end = prev_year_bounds(last_date)
    return {
        "pq": vwap_reading(anchored_vwap_band(hist, pq_start, pq_end), close),
        "py": vwap_reading(anchored_vwap_band(hist, py_start, py_end), close),
    }


def pick_snapshot(market_date: str, snapshots: list[str]) -> str | None:
    """Nearest OHLCV snapshot on or after `market_date` -- its `rows` cover
    every earlier trading day too, so it can stand in for a date that has
    no dedicated archive of its own. None if `market_date` is newer than
    every snapshot taken so far (nothing to trim down from yet)."""
    if market_date in snapshots:
        return market_date
    for s in snapshots:
        if s >= market_date:
            return s
    return None


def main(market_date: str) -> None:
    tech_path = DATES_DIR / market_date / "technical.json"
    tech = json.loads(tech_path.read_text(encoding="utf-8"))
    tech_records = tech.get("records") or {}
    listed = json.loads(LISTED_PATH.read_text(encoding="utf-8"))
    tickers = sorted({r["ticker"] for r in listed["records"]})

    snapshot_date = pick_snapshot(market_date, available_ohlcv_snapshots())
    if snapshot_date is None:
        print(f"No OHLCV snapshot on or after {market_date} yet -- nothing to compute from.")
        return
    if snapshot_date != market_date:
        print(f"{market_date} has no dedicated OHLCV archive -- reconstructing from the {snapshot_date} snapshot, trimmed back.")

    out: dict[str, dict] = {}
    monday: dict[str, dict] = {}
    vwap_out: dict[str, dict] = {}
    for i, ticker in enumerate(tickers, 1):
        hist = load_hist(ticker, snapshot_date, market_date)
        if hist is None:
            continue
        mp = ((tech_records.get(ticker) or {}).get("technical") or {}).get("marketProfile") or {}
        ibh, ibl = mp.get("ibh"), mp.get("ibl")
        close = float(hist["Close"].iloc[-1])
        rsi_div_bullish, rsi_div_bullish_pivot = regular_bullish_divergence_today(hist)
        rsi_div_hidden, rsi_div_hidden_pivot = hidden_bullish_divergence_today(hist)
        rsi_div_bearish, rsi_div_bearish_pivot = regular_bearish_divergence_today(hist)
        rsi_div_hidden_bear, rsi_div_hidden_bear_pivot = hidden_bearish_divergence_today(hist)
        record = {
            "breakIbhIbl": ib_break(hist, ibh, ibl),
            "rsiDivBullish": rsi_div_bullish,
            "rsiDivBullishPivotDate": rsi_div_bullish_pivot,
            "rsiDivHiddenBullish": rsi_div_hidden,
            "rsiDivHiddenBullishPivotDate": rsi_div_hidden_pivot,
            "rsiDivBearish": rsi_div_bearish,
            "rsiDivBearishPivotDate": rsi_div_bearish_pivot,
            "rsiDivHiddenBearish": rsi_div_hidden_bear,
            "rsiDivHiddenBearishPivotDate": rsi_div_hidden_bear_pivot,
            "stochRsiGoldenCross": stoch_rsi_golden_cross_today(hist),
            "stochRsiOversold": stoch_rsi_oversold_today(hist),
            "breakSma200": break_sma200_today(hist),
            "emaGoldenCross": ema_golden_cross_today(hist),
            **near_vwap_flags(hist, close),
        }
        if any(v for k, v in record.items() if not k.endswith("PivotDate")):
            out[ticker] = record
        mr = monday_range(hist)
        if mr is not None:
            monday[ticker] = mr
        vwap = pq_py_vwap_readings(hist, close)
        if vwap["pq"] is not None or vwap["py"] is not None:
            vwap_out[ticker] = vwap
        if i % 200 == 0:
            print(f"[{i}/{len(tickers)}] ...")

    payload = {
        "schemaVersion": 2,
        "marketDate": market_date,
        "note": (
            "Screener filter signals computed from the published OHLCV archive. "
            "breakIbhIbl/rsiDiv*/stochRsiGoldenCross are 'today' events; "
            "nearPq*/nearPy* are current-state proximity flags. "
            f"'Near' = within {NEAR_PCT * 100:.0f}% of the level. Only tickers "
            "with at least one true flag are listed in `records` -- absence "
            "means all false. `mondayRange` and `vwap` are populated for every "
            "ticker with enough history (display fields, not sparse signal "
            "lists) -- `vwap` is a fallback source for archives whose "
            "technical.json predates the vwapProfiles field."
        ),
        "records": out,
        "mondayRange": monday,
        "vwap": vwap_out,
    }
    out_path = DATES_DIR / market_date / "screener_signals.json"
    out_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")
    print(f"Wrote {len(out)} tickers with at least one signal, {len(monday)} with mondayRange -> {out_path}")


if __name__ == "__main__":
    market_date = sys.argv[1] if len(sys.argv) > 1 else None
    if not market_date:
        manifest = json.loads((ROOT / "docs" / "data" / "manifest.json").read_text(encoding="utf-8"))
        market_date = max(d["marketDate"] for d in manifest["dates"])
    main(market_date)
