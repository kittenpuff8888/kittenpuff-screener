// Setups-screener interpretation layer over the REAL IDX workbook
// (docs/data/dates/<date>/screener.json + setups.json). This is the typed
// INGEST boundary: the fragile string-parsing the workbook emits — VWAP sigma
// out of "Near +1 σ (0.06%)", RSI out of the summary sentence, the EMA stack
// out of the MA blob, the Foreign/Local flow (a labelled KSEI proxy, since no
// broker-level column exists) — is done ONCE here and exposed as clean typed
// fields. The UI consumes `UniverseRow`, never the raw strings.
//
// Nothing is fabricated. Genuinely-absent fields return {available:false} so
// the UI can render an explicit "no data". BOS has no dedicated workbook
// column, so it is an INFERRED predicate (flagged `inferred: true`).

import { fetchJson } from "./client";
import { IDX_SECTOR_MAP } from "@/lib/domain/sectors";
import type { IndexPayload } from "@/lib/domain/types";

export type Tone = "up" | "down" | "flat";

/** "71.57 M" / "199.75 K" style, matching the pre-formatted string the
    current pipeline writes to technical.liquidity.averageVolume20 -- used
    only as a fallback for older archives that carry the raw number at
    technical.averageVolume20 but not that pre-formatted nested field. */
function formatVolumeCompact(n: number): string {
  if (n >= 1e9) return `${(n / 1e9).toFixed(2)} B`;
  if (n >= 1e6) return `${(n / 1e6).toFixed(2)} M`;
  if (n >= 1e3) return `${(n / 1e3).toFixed(2)} K`;
  return n.toFixed(0);
}

/** Reality of a predicate/field against the workbook. */
export type Backing = true | false | "sparse";

/** One interpreted signal cell (Trend / Structure / VWAP / Liquidity / Flow). */
export type Cell =
  | { available: false }
  | {
      available: true;
      label: string;
      tone: Tone;
      /** Display value, e.g. "RSI 63.2" or "−1.4σ". */
      val: string;
      /** Raw provenance string this was derived from. */
      raw: string;
      /** Sort key for the column. */
      sort: number;
      /** True when the value is a labelled proxy (KSEI flow, not broker data). */
      proxy?: boolean;
    };

export type CapTier = "Large cap" | "Mid cap" | "Small cap";

/** A fully-typed, one-per-ticker universe row. */
export interface UniverseRow {
  ticker: string;
  sectorCode: string;
  sectorLabel: string;
  price: number | null;
  chg: number; // ratio (0.0176 = +1.76%)
  rvol: number | null;
  adr: number | null;
  rr: number | null;
  rsRating: number | null;
  signalType: string;
  summary: string;
  maRaw: string;
  news: string;
  entry: number | null;
  target: number | null;
  invalidation: number | null;
  entryPOI: string;
  targetPOI: string;
  // parsed-from-string, now typed:
  rsi: number | null;
  vwapSigma: number | null;
  vwapRaw: string | null;
  emaAboveKey: boolean;
  emaBelowKey: boolean;
  reclaim: boolean;
  // monthly Initial Balance (technical.json marketProfile join). null until
  // the month's first 2 sessions lock the band (IDX_Screener.py "ibh"/"ibl").
  ibh: number | null;
  ibl: number | null;
  // technical.json macdDetail/rsiDetail join -- real backend-computed
  // crossover/divergence fields, not re-derived from price here.
  macdCross: string | null; // "Golden Cross" | "Dead Cross" | "N/A" | null
  rsiDivergence: string | null; // "Bullish" | "Bearish" | null
  // true when the backend classified the current divergence as "Hidden"
  // (continuation pattern) rather than regular Strong/Medium/Weak (reversal).
  rsiDivergenceHidden: boolean;
  stochCross: string | null; // "Golden Cross" | "Dead Cross" | null
  // real technical.json fields for the redesigned results table (2026-09
  // column set) -- straight joins, nothing derived/guessed.
  beta: number | null;
  avgVolume20: string | null; // pre-formatted, e.g. "100.66 M"
  structureInternal: string | null;
  structureSwing: string | null;
  rsi14: number | null;
  rsiMa14: number | null; // RSI-based MA (technical.rsiDetail.average14)
  maZoneReal: string | null; // technical.maZone -- MA-only, no RSI mixed in
  vwapPq: VwapReading | null;
  vwapPy: VwapReading | null;
  mondayRange: MondayRange | null;
  // scripts/compute_screener_signals.py join -- the four screener filters
  // below, computed fresh from the published OHLCV archive (see that
  // script's docstring for exact parameters). Absent for a ticker with too
  // little history for a given signal, never guessed.
  breakIbhIbl: boolean;
  rsiDivBullish: boolean;
  rsiDivBullishPivotDate: string | null; // pivot bar's own date -- NOT the confirmation date; see compute_screener_signals.py
  rsiDivHiddenBullish: boolean;
  rsiDivHiddenBullishPivotDate: string | null;
  rsiDivBearish: boolean;
  rsiDivBearishPivotDate: string | null;
  rsiDivHiddenBearish: boolean;
  rsiDivHiddenBearishPivotDate: string | null;
  stochRsiGoldenCross: boolean;
  stochRsiOversold: boolean; // earlier, lower-precision companion: K and D both <20, K converging on D but not yet crossed
  breakSma200: boolean;
  emaGoldenCross: boolean;
  nearPqM1: boolean;
  nearPqM2: boolean;
  nearPyM1: boolean;
  nearPyM2: boolean;
  // engine (setups.json) join:
  score: number | null;
  isRecentIpo: boolean | null;
  medianValueTraded: number | null;
  kseiDelta: number | null; // percentage points, labelled proxy
  kseiAccum: boolean | null;
  // derived:
  capTier: CapTier;
  smcZone: string;
  setupsMatched: string[];
  setupsBear: string[];
  freshRank: number;
  cellTrend: Cell;
  cellStructure: Cell;
  cellVwap: Cell;
  cellLiquidity: Cell;
  cellFlow: Cell;
}

export interface Universe {
  marketDate: string;
  totalUniverse: number;
  scanned: number;
  rows: UniverseRow[];
}

// ── sector labels (workbook uses IDX* codes) ──
const SECTOR_LABEL: Record<string, string> = {
  ...IDX_SECTOR_MAP,
  Others: "Others",
};

// ── raw workbook record (only the fields we read) ──
type ScreenerRecord = Record<string, unknown>;
type ScreenerDoc = { marketDate?: string; records?: ScreenerRecord[]; uniqueTickerCount?: number };
type EngineSetup = {
  ticker: string;
  score?: number;
  isRecentIpo?: boolean;
  medianValueTraded20?: number;
  kseiFootprint?: { available?: boolean; netDeltaPP?: number; accumulation?: boolean } | null;
};
type SetupsDoc = { setups?: EngineSetup[]; scanned?: number };
type TechnicalRecord = {
  beta?: unknown;
  sector?: unknown;
  lastPrice?: unknown;
  changePercent?: unknown;
  rvol?: unknown;
  // Present on every archived date (old and new schema alike) -- unlike the
  // richer technical.* fields below, which only exist on dates generated
  // after the technical-schema upgrade. Used as fallbacks for those.
  averageVolume20?: unknown;
  movingAverages?: { zone?: unknown };
  technical?: {
    marketProfile?: { ibh?: unknown; ibl?: unknown };
    macdDetail?: { cross?: unknown };
    rsiDetail?: { cross?: unknown; divergenceSignal?: unknown; divergenceStrength?: unknown; average14?: unknown };
    stochDetail?: { cross?: unknown };
    rsi14?: unknown;
    // Old-schema archives carry RSI's own MA directly under technical
    // (no rsiDetail wrapper existed yet); new ones nest it in rsiDetail.
    rsiMa14?: unknown;
    maZone?: unknown;
    liquidity?: { averageVolume20?: unknown };
    vwapProfiles?: {
      previousQuarter?: { vwap?: unknown; zone?: unknown; priceSigma?: unknown };
      previousYear?: { vwap?: unknown; zone?: unknown; priceSigma?: unknown };
    };
  };
  structure?: { internal?: unknown; swing?: unknown };
  // On incremental-day snapshots (build_historical_snapshots.py's lighter
  // path, roughly half of all archived dates), ibh/ibl live here instead of
  // technical.marketProfile -- that nested object doesn't exist at all on
  // those days. Read both; prefer marketProfile, fall back to levels.
  levels?: { ibh?: unknown; ibl?: unknown };
};
type TechnicalDoc = { records?: Record<string, TechnicalRecord> };
type VwapReading = { vwap: number; sigma: number | null; zone: string };
type TechExtra = {
  ibh: number | null;
  ibl: number | null;
  macdCross: string | null;
  rsiDivergence: string | null;
  rsiDivergenceHidden: boolean;
  stochCross: string | null;
  beta: number | null;
  avgVolume20: string | null;
  structureInternal: string | null;
  structureSwing: string | null;
  rsi14: number | null;
  rsiMa14: number | null;
  maZoneReal: string | null;
  vwapPq: VwapReading | null;
  vwapPy: VwapReading | null;
  // Real fallback fields for tickers stubbed in (present in technical.json /
  // screener_signals.json but not screener.json's own ~278-ticker legacy
  // roster) -- without these a stub row's Price/Chg%/RVOL/Sector rendered
  // as an incorrect "no data" dash even though real values existed.
  lastPrice: number | null;
  changePercent: number | null;
  rvolReal: number | null;
  sectorReal: string | null;
};
type IbMap = Record<string, TechExtra>;
export type MondayRange = { high: number; low: number; date: string; status: "Above" | "Within" | "Below" };

function str(v: unknown): string {
  return v === null || v === undefined ? "" : String(v);
}
function num(v: unknown): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "string" && v.trim() && v.trim() !== "-") {
    const n = Number(v.replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

// ── parse "Near +1 σ (0.06%)", "Between +2 σ and +1 σ", "Below -3 σ" → numeric sigma ──
export function parseSigma(input: unknown): number | null {
  const s = str(input);
  if (!s || s === "-") return null;
  const nums = (s.match(/[+-]?\d+(?:\.\d+)?(?=\s*σ)/g) || []).map(Number);
  if (/^Above/.test(s)) return 3.2;
  if (/^Below/.test(s)) return -3.2;
  if (/^Near/.test(s)) return nums[0] ?? 0;
  if (/^Between/.test(s)) return nums.length >= 2 ? (nums[0] + nums[1]) / 2 : nums[0] ?? 0;
  return nums[0] ?? 0;
}

// ── RSI out of the Summary Screener sentence ("… RSI 63.2 > 50 …") ──
export function parseRsi(summary: unknown): number | null {
  const m = str(summary).match(/RSI\s+([\d.]+)/i);
  return m ? Number(m[1]) : null;
}

// ── the preset setups — each a predicate over REAL fields, with a bear mirror
//    where derivable. `inferred` flags predicates with no dedicated column. ──
export interface SetupDef {
  key: string;
  label: string;
  icon: string;
  hasBear: boolean;
  req: string;
  inferred?: boolean;
  bull: (r: UniverseRow) => boolean;
  bear?: (r: UniverseRow) => boolean;
}

// Exactly the eight screener filters requested — nothing else. Each reads a
// field scripts/compute_screener_signals.py computed fresh from the
// published OHLCV archive that day (see that script's docstring for exact
// parameters/formulas); "today" fields (breakIbhIbl, the two RSI divergence
// variants, stochRsiGoldenCross) are true only on the session the event
// itself occurred, not "still in that state from days ago".
export const SETUPS: SetupDef[] = [
  {
    key: "break_ibh_ibl",
    label: "Break IBH/IBL",
    icon: "⇕",
    hasBear: false,
    req: "Yesterday's close sat mid-band inside the monthly Initial Balance (first 2 sessions of the month), and today's close breaks above the IBH",
    bull: (r) => r.breakIbhIbl,
  },
  {
    key: "rsi10_div_bullish",
    label: "RSI Bullish Divergence",
    icon: "⤢",
    hasBear: false,
    req: "Regular bullish RSI(14, Wilder-smoothed) divergence confirmed today — a literal port of a user-supplied Pine Script v6 divergence indicator: price (Low) makes a lower low while RSI makes a higher low, comparing the two most recent confirmed 5-bar RSI pivots, 5–60 bars apart, no RSI-value threshold",
    bull: (r) => r.rsiDivBullish,
  },
  {
    key: "rsi10_div_hidden_bullish",
    label: "RSI Hidden Bullish Divergence",
    icon: "⤢",
    hasBear: false,
    req: "Hidden bullish RSI(14, Wilder-smoothed) divergence confirmed today — the exact mirror of Regular above: price (Low) makes a higher low while RSI makes a lower low, same 5-bar confirmed-pivot logic and 5–60 bar gap, no RSI-value threshold",
    bull: (r) => r.rsiDivHiddenBullish,
  },
  {
    key: "rsi14_div_bearish",
    label: "RSI Bearish Divergence",
    icon: "⤡",
    hasBear: false,
    req: "Regular bearish RSI(14, Wilder-smoothed) divergence confirmed today — the pivot-high mirror of RSI Bullish Divergence, ported from the same Pine Script v6 indicator: price (High) makes a higher high while RSI makes a lower high, comparing the two most recent confirmed 5-bar RSI pivots, 5–60 bars apart, no RSI-value threshold",
    bull: (r) => r.rsiDivBearish,
  },
  {
    key: "rsi14_div_hidden_bearish",
    label: "RSI Hidden Bearish Divergence",
    icon: "⤡",
    hasBear: false,
    req: "Hidden bearish RSI(14, Wilder-smoothed) divergence confirmed today — the exact mirror of Regular Bearish above: price (High) makes a lower high while RSI makes a higher high, same 5-bar confirmed-pivot logic and 5–60 bar gap, no RSI-value threshold",
    bull: (r) => r.rsiDivHiddenBearish,
  },
  {
    key: "stoch_rsi_golden_cross",
    label: "Stoch RSI Golden Cross",
    icon: "✦",
    hasBear: false,
    req: "Stochastic RSI %K crosses above %D today (RSI length 10, Stochastic length 10, K 3, D 3, source Close), the cross originating from Stoch RSI's own oversold band (K and D both < 20 the day before)",
    bull: (r) => r.stochRsiGoldenCross,
  },
  {
    key: "stoch_rsi_oversold",
    label: "Stoch RSI Oversold",
    icon: "◐",
    hasBear: false,
    req: "Stochastic RSI %K and %D are both under 20 today, K converging on D but hasn't crossed above it yet — an earlier, lower-precision companion to Golden Cross: a genuine reversal confirmation can only appear after the bounce that produces it, so this flags the watch-for-a-turn state instead, at the cost of far more (and less reliable) matches",
    bull: (r) => r.stochRsiOversold,
  },
  {
    key: "break_sma200",
    label: "Break MA 200",
    icon: "▲",
    hasBear: false,
    req: "Today's close crosses above SMA200 — yesterday's close was at or below it",
    bull: (r) => r.breakSma200,
  },
  {
    key: "ema_golden_cross",
    label: "EMA 25/50 Golden Cross",
    icon: "✦",
    hasBear: false,
    req: "EMA25 crosses above EMA50 today — yesterday EMA25 was at or below EMA50",
    bull: (r) => r.emaGoldenCross,
  },
  {
    key: "near_vwap_pq_m1",
    label: "Near VWAP — PQ −1σ",
    icon: "≈",
    hasBear: false,
    req: "Close within 1% of the Previous Quarter anchored-VWAP −1σ band",
    bull: (r) => r.nearPqM1,
  },
  {
    key: "near_vwap_pq_m2",
    label: "Near VWAP — PQ −2σ",
    icon: "≈",
    hasBear: false,
    req: "Close within 1% of the Previous Quarter anchored-VWAP −2σ band",
    bull: (r) => r.nearPqM2,
  },
  {
    key: "near_vwap_py_m1",
    label: "Near VWAP — PY −1σ",
    icon: "≈",
    hasBear: false,
    req: "Close within 1% of the Previous Year anchored-VWAP −1σ band",
    bull: (r) => r.nearPyM1,
  },
  {
    key: "near_vwap_py_m2",
    label: "Near VWAP — PY −2σ",
    icon: "≈",
    hasBear: false,
    req: "Close within 1% of the Previous Year anchored-VWAP −2σ band",
    bull: (r) => r.nearPyM2,
  },
];

// UI-only grouping: the 4 Near-VWAP setups above are still 4 independent,
// separately-matched signals (setupsMatched/CSV/solo counts are unaffected)
// — this just tells the picker grid to render the ±1σ/±2σ pair for each
// period as ONE box with an in-box σ toggle, instead of 4 separate tiles.
export type VwapGroupVariant = { sigmaLabel: string; setupKey: string };
export type VwapGroup = { boxId: string; label: string; icon: string; req: string; variants: VwapGroupVariant[] };
export const VWAP_GROUPS: VwapGroup[] = [
  {
    boxId: "near_vwap_pq",
    label: "Holding Previous QVWAP",
    icon: "≈",
    req: "Close within the picked σ band of the Previous Quarter anchored-VWAP",
    variants: [
      { sigmaLabel: "−1σ", setupKey: "near_vwap_pq_m1" },
      { sigmaLabel: "−2σ", setupKey: "near_vwap_pq_m2" },
    ],
  },
  {
    boxId: "near_vwap_py",
    label: "Holding Previous YVWAP",
    icon: "≈",
    req: "Close within the picked σ band of the Previous Year anchored-VWAP",
    variants: [
      { sigmaLabel: "−1σ", setupKey: "near_vwap_py_m1" },
      { sigmaLabel: "−2σ", setupKey: "near_vwap_py_m2" },
    ],
  },
];
const VWAP_GROUP_KEYS = new Set(VWAP_GROUPS.flatMap((g) => g.variants.map((v) => v.setupKey)));
/** SETUPS entries that render as their own tile (i.e. not one of the 4
    Near-VWAP setups folded into VWAP_GROUPS above). */
export function ungroupedSetups(): SetupDef[] {
  return SETUPS.filter((s) => !VWAP_GROUP_KEYS.has(s.key));
}

const VWAP_VARIANT_DISPLAY: Record<string, string> = Object.fromEntries(
  VWAP_GROUPS.flatMap((g) => g.variants.map((v) => [v.setupKey, `${g.label} · ${v.sigmaLabel}`])),
);
/** The label a setup key should actually render as — the merged box name
    (e.g. "Holding Previous QVWAP · −1σ") for a Near-VWAP variant, or the
    plain SetupDef label for everything else. Used by the Setup-column
    bubbles and the filter pills so they match what the picker grid shows,
    not the pre-merge per-σ tile name. */
export function setupDisplayLabel(key: string): string {
  return VWAP_VARIANT_DISPLAY[key] || SETUP_BY_KEY[key]?.label || key;
}

const SETUP_BY_KEY: Record<string, SetupDef> = Object.fromEntries(SETUPS.map((s) => [s.key, s]));
export function setupByKey(key: string): SetupDef | undefined {
  return SETUP_BY_KEY[key];
}

// ── general top filters. Sector is REAL; Liquidity is a labelled RVOL/cap
//    proxy; Konglo groups are REAL (docs/data/indexes.json's "KONGLO INDEX"
//    section, one group per Indonesian conglomerate with real constituent
//    tickers) — join it by ticker rather than re-fetching per-date. ──
export const SECTOR_OPTIONS: Array<{ v: string; label: string }> = [
  { v: "", label: "All sectors" },
  ...Object.entries(SECTOR_LABEL).map(([v, label]) => ({ v, label })),
];
export const LIQUIDITY_OPTIONS: Array<{ v: string; label: string }> = [
  { v: "", label: "All liquidity" },
  { v: "surging", label: "Surging (RVOL ≥ 3×)" },
  { v: "busy", label: "Busy (RVOL ≥ 1.5×)" },
  { v: "avg", label: "Average" },
  { v: "quiet", label: "Quiet (RVOL < 0.5×)" },
];
export const PRICE_OPTIONS: Array<{ v: string; label: string }> = [
  { v: "", label: "Price: any" },
  { v: "under50", label: "Price < 50" },
  { v: "over50", label: "Price ≥ 50" },
];

const KONGLO_SECTION = "KONGLO INDEX";

/** ticker → the real konglomerate-group labels it belongs to (a ticker can
    sit in more than one group), built from indexes.json's own groups —
    nothing invented, absent from the map = genuinely not in any group. */
export function kongloGroupsByTicker(indexes: IndexPayload | null): Record<string, string[]> {
  const map: Record<string, string[]> = {};
  (indexes?.groups || []).forEach((g) => {
    if (g.section !== KONGLO_SECTION) return;
    (g.constituents || []).forEach((c) => {
      const t = String(c.ticker || "").toUpperCase();
      if (!t) return;
      (map[t] ||= []).push(g.label);
    });
  });
  return map;
}

/** v = the group's own label (also the key kongloGroupsByTicker's values
    use), so passKonglo needs no separate id→label lookup. */
export function kongloOptionsFromIndexes(indexes: IndexPayload | null): Array<{ v: string; label: string }> {
  const groups = (indexes?.groups || []).filter((g) => g.section === KONGLO_SECTION);
  const opts = groups
    .map((g) => ({ v: g.label, label: `${g.label} (${(g.constituents || []).length})` }))
    .sort((a, b) => a.label.localeCompare(b.label));
  return [{ v: "", label: opts.length ? "All konglo groups" : "All konglo groups — no data" }, ...opts];
}

export function passKonglo(kongloMap: Record<string, string[]>, ticker: string, key: string): boolean {
  if (!key) return true;
  return (kongloMap[ticker] || []).includes(key);
}

export function passPrice(r: UniverseRow, key: string): boolean {
  if (!key || r.price == null) return !key;
  if (key === "under50") return r.price < 50;
  if (key === "over50") return r.price >= 50;
  return true;
}

export function passLiquidity(r: UniverseRow, key: string): boolean {
  if (!key || r.rvol == null) return !key;
  if (key === "surging") return r.rvol >= 3;
  if (key === "busy") return r.rvol >= 1.5;
  if (key === "avg") return r.rvol >= 0.5 && r.rvol < 1.5;
  if (key === "quiet") return r.rvol < 0.5;
  return true;
}

// ── interpreted cells ──
function trendCell(r: UniverseRow): Cell {
  if (r.rsi == null && !r.maRaw) return { available: false };
  const bull = r.emaAboveKey && (r.rsi ?? 0) > 50;
  const bear = r.emaBelowKey || (r.rsi != null && r.rsi < 45);
  const label = bull ? "Above EMA25 & EMA50" : r.emaAboveKey ? "Above key EMAs" : "Below key EMAs";
  const tone: Tone = bull ? "up" : bear ? "down" : "flat";
  return {
    available: true,
    label,
    tone,
    val: r.rsi != null ? `RSI ${r.rsi.toFixed(1)}` : r.maRaw,
    raw: r.summary || r.maRaw || "",
    sort: (r.emaAboveKey ? 100 : 0) + (r.rsi ?? 0),
  };
}

function structureCell(r: UniverseRow): Cell {
  const poi = r.entryPOI;
  if (/OB Bull/i.test(poi)) return { available: true, label: "In bullish order block", tone: "up", val: poi, raw: r.summary, sort: 3 };
  if (/^EQ/.test(poi)) return { available: true, label: "Equilibrium reclaim", tone: "up", val: poi, raw: r.summary, sort: 2 };
  if (r.reclaim && r.emaAboveKey) return { available: true, label: "Bullish reclaim", tone: "up", val: "reclaim + EMA stack", raw: r.summary, sort: 2 };
  if (/OB Bear/i.test(poi)) return { available: true, label: "At bearish order block", tone: "down", val: poi, raw: r.summary, sort: -2 };
  if ((r.vwapSigma ?? 9) < -0.75) return { available: true, label: "SMC discount", tone: "up", val: `${(r.vwapSigma ?? 0).toFixed(1)}σ`, raw: r.vwapRaw || "", sort: 1 };
  return { available: true, label: "Neutral structure", tone: "flat", val: poi || "—", raw: r.summary, sort: 0 };
}

function vwapCell(r: UniverseRow): Cell {
  if (r.vwapSigma == null) return { available: false };
  const s = r.vwapSigma;
  if (/^EQ/.test(r.entryPOI) && r.reclaim) return { available: true, label: "Reclaimed EQ", tone: "up", val: r.vwapRaw || "", raw: r.vwapRaw || "", sort: 2 };
  let label: string;
  let tone: Tone;
  if (s <= -2) { label = "Deep discount to VWAP"; tone = "up"; }
  else if (s <= -0.6) { label = "Below VWAP (discount)"; tone = "up"; }
  else if (s < 0.6) { label = "Near VWAP"; tone = "flat"; }
  else if (s < 2) { label = "Above VWAP (premium)"; tone = "down"; }
  else { label = "Rich vs VWAP"; tone = "down"; }
  const val = `${s > 0 ? "+" : s < 0 ? "−" : ""}${Math.abs(s).toFixed(1)}σ`;
  return { available: true, label, tone, val, raw: r.vwapRaw || "", sort: -s };
}

function liquidityCell(r: UniverseRow): Cell {
  if (r.rvol == null) return { available: false };
  const busy = r.rvol >= 1.5;
  const thin = r.rvol < 0.5;
  const label = busy ? "Busy vs its norm" : thin ? "Quiet vs its norm" : "Average activity";
  const adtv = r.medianValueTraded != null ? ` · ADTV ${(r.medianValueTraded / 1e9).toFixed(1)}B` : "";
  return {
    available: true,
    label,
    tone: busy ? "up" : "flat",
    val: `RVOL ${r.rvol.toFixed(2)}×${adtv}`,
    raw: `RVOL ${r.rvol.toFixed(3)}${adtv ? " (ADTV real)" : " (ADTV: no data)"}`,
    sort: r.rvol,
  };
}

// Foreign/Local flow — a LABELLED KSEI proxy. Broker-level buy/sell does not
// exist in the data, so we never invent broker counts.
function flowCell(r: UniverseRow): Cell {
  if (r.kseiDelta == null) return { available: false };
  const d = r.kseiDelta;
  if (Math.abs(d) < 0.05) return { available: true, label: "Balanced flow", tone: "flat", val: "KSEI proxy 0.0pp", raw: `KSEI net ${d.toFixed(2)}pp (proxy)`, proxy: true, sort: 0 };
  const fb = d > 0;
  return {
    available: true,
    label: fb ? "Foreign buying, local selling" : "Foreign selling, local buying",
    tone: fb ? "up" : "down",
    val: `KSEI proxy ${d > 0 ? "+" : "−"}${Math.abs(d).toFixed(1)}pp`,
    raw: `KSEI net ${d.toFixed(2)}pp (proxy — no broker data)`,
    proxy: true,
    sort: d,
  };
}

export type ScreenerSignal = {
  breakIbhIbl?: boolean;
  rsiDivBullish?: boolean; rsiDivBullishPivotDate?: string | null;
  rsiDivHiddenBullish?: boolean; rsiDivHiddenBullishPivotDate?: string | null;
  rsiDivBearish?: boolean; rsiDivBearishPivotDate?: string | null;
  rsiDivHiddenBearish?: boolean; rsiDivHiddenBearishPivotDate?: string | null;
  stochRsiGoldenCross?: boolean; stochRsiOversold?: boolean;
  breakSma200?: boolean; emaGoldenCross?: boolean;
  nearPqM1?: boolean; nearPqM2?: boolean; nearPyM1?: boolean; nearPyM2?: boolean;
};
type RawVwapReading = { vwap: number; sigma: number | null };
export type ScreenerSignalsDoc = {
  records?: Record<string, ScreenerSignal>;
  mondayRange?: Record<string, MondayRange>;
  // Fallback source for vwapPq/vwapPy on archives whose technical.json
  // predates the vwapProfiles field -- see loadUniverse()'s ibMap join.
  vwap?: Record<string, { pq?: RawVwapReading | null; py?: RawVwapReading | null }>;
};

// ── build the typed universe from the raw workbook + engine docs ──
export function buildUniverse(scr: ScreenerDoc, setupsDoc: SetupsDoc, ibMap: IbMap = {}, signals: Record<string, ScreenerSignal> = {}, mondayRangeMap: Record<string, MondayRange> = {}): Universe {
  const setupByTicker: Record<string, EngineSetup> = {};
  (setupsDoc.setups || []).forEach((s) => (setupByTicker[s.ticker] = s));

  // one record per ticker — keep the best-R/R signal row. Every real IDX
  // ticker is exactly 4 letters/digits (verified against idx-listed.json,
  // 962/962); screener.json's own Excel-parsing pipeline occasionally spills
  // a neighbouring cell's text into the Ticker column (observed live:
  // "ISTRUCTURE BREAK", a fragment of a Structure-column sentence) --
  // reject anything that isn't a real ticker shape rather than let a
  // corrupted phantom row into the table.
  const TICKER_RE = /^[A-Z0-9]{4}$/;
  const byTicker: Record<string, ScreenerRecord> = {};
  (scr.records || []).forEach((rec) => {
    const t = str(rec.Ticker).toUpperCase();
    if (!TICKER_RE.test(t)) return;
    const cur = byTicker[t];
    if (!cur || (num(rec["R/R"]) ?? 0) > (num(cur["R/R"]) ?? 0)) byTicker[t] = rec;
  });
  // screener.json's roster is a ~400-ticker subset (the legacy engine's own
  // "scanned" set); technical.json and screener_signals.json are both
  // computed for the FULL ~962-ticker universe. Any ticker with real data
  // in either of those but no roster row (most often a thinly-traded name
  // outside the legacy engine's scan) must still get a row here, or its
  // real fields (and any of the 8 signals it triggers) can never surface in
  // the table — stub in the bare minimum (every other field degrades to its
  // existing "no data" / null handling; nothing is guessed).
  new Set([...Object.keys(signals), ...Object.keys(ibMap)]).forEach((t) => {
    const tu = t.toUpperCase();
    if (byTicker[tu]) return;
    // Real fallback data (technical.json covers the full universe, unlike
    // screener.json's own ~278-ticker legacy roster) -- without this a
    // stub row's Price/Chg%/RVOL/Sector rendered as an incorrect "no data"
    // dash even when technical.json had the real values right there.
    const te = ibMap[tu];
    byTicker[tu] = {
      Ticker: tu,
      Price: te?.lastPrice ?? undefined,
      "Chg %": te?.changePercent ?? undefined,
      RVOL: te?.rvolReal ?? undefined,
      Sector: te?.sectorReal ?? undefined,
    };
  });

  const rows: UniverseRow[] = Object.values(byTicker).map((rec) => {
    const su = setupByTicker[str(rec.Ticker).toUpperCase()];
    const maRaw = str(rec.MA);
    const summary = str(rec["Summary Screener"]);
    const emaAboveKey =
      /Above All Available MA/.test(maRaw) ||
      /Above[^|]*EMA25[^|]*EMA50/.test(maRaw) ||
      /Close .*> EMA25.*EMA25.*> EMA50/i.test(summary);
    const emaBelowKey = /Below[^|]*EMA50/.test(maRaw) && !emaAboveKey;
    const sectorCode = str(rec.Sector) || "Others";
    const price = num(rec.Price);
    const tickerU = str(rec.Ticker).toUpperCase();
    const sig = signals[tickerU] || {};
    const te = ibMap[tickerU];

    const r: UniverseRow = {
      ticker: str(rec.Ticker).toUpperCase(),
      sectorCode,
      sectorLabel: SECTOR_LABEL[sectorCode] || sectorCode,
      price,
      chg: num(rec["Chg %"]) ?? 0,
      rvol: num(rec.RVOL),
      adr: num(rec["ADR %"]),
      rr: num(rec["R/R"]),
      rsRating: num(rec["RS Rating"]),
      signalType: str(rec.signalType),
      summary,
      maRaw,
      news: str(rec["Sentiment News"]),
      entry: num(rec.Entry),
      target: num(rec.Target),
      invalidation: num(rec.Invalidation),
      entryPOI: str(rec["Entry POI"]),
      targetPOI: str(rec["Target POI"]),
      rsi: parseRsi(summary),
      vwapSigma: parseSigma(rec["Current Q VWAP"]),
      vwapRaw: str(rec["Current Q VWAP"]) || null,
      emaAboveKey,
      emaBelowKey,
      reclaim: /reclaim/i.test(summary),
      beta: te?.beta ?? null,
      avgVolume20: te?.avgVolume20 ?? null,
      structureInternal: te?.structureInternal ?? null,
      structureSwing: te?.structureSwing ?? null,
      rsi14: te?.rsi14 ?? null,
      rsiMa14: te?.rsiMa14 ?? null,
      maZoneReal: te?.maZoneReal ?? null,
      vwapPq: te?.vwapPq ?? null,
      vwapPy: te?.vwapPy ?? null,
      mondayRange: mondayRangeMap[tickerU] ?? null,
      breakIbhIbl: !!sig.breakIbhIbl,
      rsiDivBullish: !!sig.rsiDivBullish,
      rsiDivBullishPivotDate: sig.rsiDivBullishPivotDate ?? null,
      rsiDivHiddenBullish: !!sig.rsiDivHiddenBullish,
      rsiDivHiddenBullishPivotDate: sig.rsiDivHiddenBullishPivotDate ?? null,
      rsiDivBearish: !!sig.rsiDivBearish,
      rsiDivBearishPivotDate: sig.rsiDivBearishPivotDate ?? null,
      rsiDivHiddenBearish: !!sig.rsiDivHiddenBearish,
      rsiDivHiddenBearishPivotDate: sig.rsiDivHiddenBearishPivotDate ?? null,
      stochRsiGoldenCross: !!sig.stochRsiGoldenCross,
      stochRsiOversold: !!sig.stochRsiOversold,
      breakSma200: !!sig.breakSma200,
      emaGoldenCross: !!sig.emaGoldenCross,
      nearPqM1: !!sig.nearPqM1,
      nearPqM2: !!sig.nearPqM2,
      nearPyM1: !!sig.nearPyM1,
      nearPyM2: !!sig.nearPyM2,
      score: su && typeof su.score === "number" ? su.score : null,
      isRecentIpo: su ? !!su.isRecentIpo : null,
      medianValueTraded: su && typeof su.medianValueTraded20 === "number" ? su.medianValueTraded20 : null,
      kseiDelta: su && su.kseiFootprint && su.kseiFootprint.available ? su.kseiFootprint.netDeltaPP ?? null : null,
      kseiAccum: su && su.kseiFootprint && su.kseiFootprint.available ? !!su.kseiFootprint.accumulation : null,
      capTier: (price ?? 0) > 3000 ? "Large cap" : (price ?? 0) >= 200 ? "Mid cap" : "Small cap",
      ibh: ibMap[str(rec.Ticker).toUpperCase()]?.ibh ?? null,
      ibl: ibMap[str(rec.Ticker).toUpperCase()]?.ibl ?? null,
      macdCross: ibMap[str(rec.Ticker).toUpperCase()]?.macdCross ?? null,
      rsiDivergence: ibMap[str(rec.Ticker).toUpperCase()]?.rsiDivergence ?? null,
      rsiDivergenceHidden: ibMap[str(rec.Ticker).toUpperCase()]?.rsiDivergenceHidden ?? false,
      stochCross: ibMap[str(rec.Ticker).toUpperCase()]?.stochCross ?? null,
      smcZone: "",
      setupsMatched: [],
      setupsBear: [],
      freshRank: 0,
      cellTrend: { available: false },
      cellStructure: { available: false },
      cellVwap: { available: false },
      cellLiquidity: { available: false },
      cellFlow: { available: false },
    };

    r.setupsMatched = SETUPS.filter((s) => {
      try { return s.bull(r); } catch { return false; }
    }).map((s) => s.key);
    r.setupsBear = SETUPS.filter((s) => s.hasBear && s.bear)
      .filter((s) => {
        try { return s.bear!(r); } catch { return false; }
      })
      .map((s) => s.key);

    r.cellTrend = trendCell(r);
    r.cellStructure = structureCell(r);
    r.cellVwap = vwapCell(r);
    r.cellLiquidity = liquidityCell(r);
    r.cellFlow = flowCell(r);

    r.smcZone = /OB Bull/i.test(r.entryPOI)
      ? "Bull OB"
      : /^EQ/.test(r.entryPOI)
        ? "Equilibrium"
        : (r.vwapSigma ?? 9) < -0.75
          ? "Discount"
          : (r.vwapSigma ?? -9) > 0.75
            ? "Premium"
            : "—";

    r.freshRank =
      (su && typeof su.score === "number" ? 100000 + su.score * 100 : 0) +
      r.setupsMatched.length * 100 +
      (r.rvol ?? 0);
    return r;
  });

  // setupsDoc.scanned / scr.uniqueTickerCount are legacy engine fields with
  // their own, DIFFERENT meaning ("tickers that matched the engine's own
  // screening criteria that day", not "size of the universe") -- on lighter
  // historical snapshots that can be far smaller than rows.length (real
  // case hit live: uniqueTickerCount 96 on a day rows.length was 923,
  // rendering the nonsensical "Matching 923 of 96"). The real universe size
  // is never smaller than the rows we actually have data for.
  const totalUniverse = Math.max(rows.length, setupsDoc.scanned || 0, scr.uniqueTickerCount || 0);
  return {
    marketDate: str(scr.marketDate),
    totalUniverse,
    scanned: rows.length,
    rows,
  };
}

/** Precomputed solo (single-setup) match counts across the universe. */
export function soloCounts(rows: UniverseRow[]): Record<string, { bull: number; bear: number }> {
  const out: Record<string, { bull: number; bear: number }> = {};
  SETUPS.forEach((s) => {
    out[s.key] = {
      bull: rows.reduce((n, r) => n + (r.setupsMatched.includes(s.key) ? 1 : 0), 0),
      bear: rows.reduce((n, r) => n + (r.setupsBear.includes(s.key) ? 1 : 0), 0),
    };
  });
  return out;
}

// ── loader: fetch the workbook + engine docs for a market date ──
export async function loadUniverse(marketDate: string): Promise<Universe> {
  const base = `/data/dates/${marketDate}`;
  const [scr, setupsDoc, technicalDoc, signalsDoc] = await Promise.all([
    fetchJson<ScreenerDoc>(`${base}/screener.json`),
    fetchJson<SetupsDoc>(`${base}/setups.json`).catch(() => ({ setups: [] }) as SetupsDoc),
    fetchJson<TechnicalDoc>(`${base}/technical.json`).catch(() => ({ records: {} }) as TechnicalDoc),
    fetchJson<ScreenerSignalsDoc>(`${base}/screener_signals.json`).catch(() => ({ records: {} }) as ScreenerSignalsDoc),
  ]);
  const ibMap: IbMap = {};
  const vwapReading = (p?: { vwap?: unknown; zone?: unknown; priceSigma?: unknown }): VwapReading | null => {
    const vwap = num(p?.vwap);
    if (vwap == null) return null;
    return { vwap, sigma: num(p?.priceSigma), zone: str(p?.zone) };
  };
  // Fallback for archives whose technical.json predates the vwapProfiles
  // field: screener_signals.json's own `vwap` block computes the same
  // anchored-VWAP reading straight from OHLCV (see compute_screener_signals.py),
  // with no zone label (the table never renders it, see vwapCellText()).
  const fallbackVwapReading = (p?: { vwap?: number; sigma?: number | null } | null): VwapReading | null => {
    if (!p || p.vwap == null) return null;
    return { vwap: p.vwap, sigma: p.sigma ?? null, zone: "" };
  };
  Object.entries(technicalDoc.records || {}).forEach(([ticker, rec]) => {
    const mp = rec.technical?.marketProfile;
    const macdCrossRaw = str(rec.technical?.macdDetail?.cross);
    const rsiDivRaw = str(rec.technical?.rsiDetail?.divergenceSignal);
    ibMap[ticker.toUpperCase()] = {
      ibh: num(mp?.ibh) ?? num(rec.levels?.ibh),
      ibl: num(mp?.ibl) ?? num(rec.levels?.ibl),
      macdCross: macdCrossRaw && macdCrossRaw !== "N/A" ? macdCrossRaw : null,
      rsiDivergence: rsiDivRaw || null,
      rsiDivergenceHidden: str(rec.technical?.rsiDetail?.divergenceStrength) === "Hidden",
      stochCross: (() => {
        const v = str(rec.technical?.stochDetail?.cross);
        return v && v !== "N/A" && v !== "-" ? v : null;
      })(),
      beta: num(rec.beta),
      // technical.liquidity.averageVolume20 is a pre-formatted string that
      // only exists on dates generated after the technical-schema upgrade;
      // older archives still carry the raw 20-day average as a plain number
      // at the record's top level -- format that instead of showing "no
      // data" for a real, already-computed figure.
      avgVolume20:
        str(rec.technical?.liquidity?.averageVolume20) ||
        (() => {
          const n = num(rec.averageVolume20);
          return n != null ? formatVolumeCompact(n) : null;
        })(),
      structureInternal: str(rec.structure?.internal) || null,
      structureSwing: str(rec.structure?.swing) || null,
      rsi14: num(rec.technical?.rsi14),
      // Old-schema archives had RSI's own MA directly under technical, not
      // nested in rsiDetail yet -- same real figure, different path.
      rsiMa14: num(rec.technical?.rsiDetail?.average14) ?? num(rec.technical?.rsiMa14),
      // Old-schema archives compute the MA-zone label from only 3 MAs
      // (movingAverages.zone, via EMA25/EMA50/SMA200) instead of the current
      // 5-MA technical.maZone -- same kind of real, already-computed
      // classification; better to show it than "no data".
      maZoneReal: str(rec.technical?.maZone) || str(rec.movingAverages?.zone) || null,
      vwapPq:
        vwapReading(rec.technical?.vwapProfiles?.previousQuarter) ??
        fallbackVwapReading(signalsDoc.vwap?.[ticker.toUpperCase()]?.pq),
      vwapPy:
        vwapReading(rec.technical?.vwapProfiles?.previousYear) ??
        fallbackVwapReading(signalsDoc.vwap?.[ticker.toUpperCase()]?.py),
      lastPrice: num(rec.lastPrice),
      changePercent: num(rec.changePercent),
      rvolReal: num(rec.rvol),
      sectorReal: str(rec.sector) || null,
    };
  });
  const mondayRangeMap = (signalsDoc.mondayRange || {}) as Record<string, MondayRange>;
  return buildUniverse(scr, setupsDoc, ibMap, signalsDoc.records || {}, mondayRangeMap);
}
