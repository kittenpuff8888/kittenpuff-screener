"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import type { CSSProperties } from "react";
import { useApp } from "@/components/providers/AppProvider";
import { kseiSectorMap } from "@/lib/data/ksei";
import { asNumber, formatPercent, formatPrice } from "@/lib/format/number";
import { computeDcf, DEFAULT_ASSUMPTIONS, type DcfInputs } from "@/lib/valuation/dcf";
import { newsStories } from "@/lib/data/news";
import type { JsonRecord } from "@/lib/domain/types";
import {
  LIQUIDITY_OPTIONS,
  PRICE_OPTIONS,
  SECTOR_OPTIONS,
  VWAP_GROUPS,
  kongloGroupsByTicker,
  kongloOptionsFromIndexes,
  loadUniverse,
  passKonglo,
  passLiquidity,
  passPrice,
  setupByKey,
  setupDisplayLabel,
  soloCounts,
  ungroupedSetups,
  type Tone,
  type Universe,
  type UniverseRow,
} from "@/lib/data/screenerUniverse";

const MONO = "var(--mono, var(--font-mono))";
const CARD: CSSProperties = { background: "var(--panel)", border: "1px solid var(--border)", borderRadius: "var(--r)", boxShadow: "var(--sh, var(--shadow))" };

const PRESETS_KEY = "idxr:screenerPresets";
const GRID = "100px 88px 118px 200px 76px 68px 56px 76px 60px 160px 128px 128px 108px 140px 100px 168px 168px 220px";

type Conf = "AND" | "OR";
type Dir = "bull" | "bear";
type SavedPreset = {
  id: string;
  name: string;
  selectedSetups: string[];
  setupDir: Record<string, Dir>;
  conf: Conf;
};

function toneColor(t: Tone): string {
  return t === "up" ? "var(--up)" : t === "down" ? "var(--down)" : "var(--flat)";
}

// IBH / IBL zone: strictly derived from real ibh/ibl/price, no fabrication.
function ibZone(price: number | null, ibh: number | null, ibl: number | null): string | null {
  if (price == null || ibh == null || ibl == null) return null;
  if (price > ibh) return "Above IBH";
  if (price < ibl) return "Below IBL";
  return "Between";
}
function ibZoneTone(zone: string | null): Tone {
  return zone === "Above IBH" ? "up" : zone === "Below IBL" ? "down" : "flat";
}
function vwapCellText(v: { vwap: number; sigma: number | null; zone: string } | null): string | null {
  if (!v) return null;
  const sigmaTxt = v.sigma != null ? ` · ${v.sigma >= 0 ? "+" : ""}${v.sigma.toFixed(2)}σ` : "";
  return `${formatPrice(v.vwap)}${sigmaTxt}`;
}
function mondayRangeTone(status: string | undefined): Tone {
  return status === "Above" ? "up" : status === "Below" ? "down" : "flat";
}
const structureTone = (s: string | null): Tone => (s ? (/Bullish/i.test(s) ? "up" : /Bearish/i.test(s) ? "down" : "flat") : "flat");

function readPresets(): SavedPreset[] {
  if (typeof window === "undefined") return [];
  try {
    const parsed = JSON.parse(window.localStorage.getItem(PRESETS_KEY) || "[]");
    return Array.isArray(parsed) ? (parsed as SavedPreset[]) : [];
  } catch {
    return [];
  }
}

// Fixed assumptions for the screener's DCF column — a table has no room for
// per-row sliders, so every ticker uses the same macro assumptions (same
// defaults as the ticker-detail DCF panel); only the FCF-growth input varies
// per ticker, from that ticker's own real revenue growth. Open the ticker
// page to adjust assumptions interactively.
const SCREENER_DCF_ASSUMPTIONS = DEFAULT_ASSUMPTIONS;

export function ScreenerPage() {
  const { marketDate, openTicker, ksei, bundle, indexes } = useApp();
  // Real per-ticker sector (the workbook ships "IDX Sector" as "-").
  const kseiSec = useMemo(() => kseiSectorMap(ksei), [ksei]);
  const secLabel = (t: string, fallback: string) => kseiSec.get(t)?.label || fallback;
  const secCode = (t: string, fallback: string) => kseiSec.get(t)?.code || fallback;
  // Real konglomerate-group membership (docs/data/indexes.json's own
  // "KONGLO INDEX" groups) -- already loaded app-wide, no extra fetch.
  const kongloMap = useMemo(() => kongloGroupsByTicker(indexes), [indexes]);
  const kongloOptions = useMemo(() => kongloOptionsFromIndexes(indexes), [indexes]);
  // ticker → its single freshest headline, from the same News page data —
  // no separate fetch, News stories arrive on `bundle` already. newsStories()
  // returns freshest-first, so the first hit per ticker is what we want.
  const newsByTicker = useMemo(() => {
    const out = new Map<string, { title: string; tone: "up" | "down" | "flat"; when: string }>();
    newsStories(bundle).forEach((s) => {
      if (!out.has(s.ticker)) out.set(s.ticker, { title: s.title, tone: s.tone, when: s.when });
    });
    return out;
  }, [bundle]);
  const [universe, setUniverse] = useState<Universe | null>(null);
  const [error, setError] = useState<string | null>(null);

  // ticker → DCF upside%, from the same model as the ticker-detail panel —
  // real Free Cash Flow / beta / price / share count, fixed macro assumptions.
  const dcfByTicker = useMemo(() => {
    const m = new Map<string, { up: number; asOf?: string }>();
    if (!bundle) return m;
    bundle.fundamentals.forEach((fund, ticker) => {
      const stock = bundle.technical.get(ticker);
      const price = asNumber(stock?.lastPrice ?? fund["Price"]);
      const beta = asNumber(stock?.beta);
      const fcfTtmBn = asNumber(fund["Free cash flow (TTM)"]);
      const revenueGrowth = asNumber(fund["Revenue (Quarter YoY Growth)"]);
      const marketCapAbs = asNumber((stock?.fundamentals as JsonRecord | undefined)?.["marketCap"]);
      const sharesOutstanding = marketCapAbs != null && price != null && price > 0 ? marketCapAbs / price : null;
      const inputs: DcfInputs = { ticker, price, beta, fcfTtmBn, revenueGrowth, sharesOutstanding, week52High: null, week52Low: null, currency: "IDR" };
      const fcfGrowthRate = revenueGrowth == null || !isFinite(revenueGrowth) ? 0.05 : Math.max(-0.3, Math.min(0.4, revenueGrowth));
      const result = computeDcf(inputs, { ...SCREENER_DCF_ASSUMPTIONS, fcfGrowthRate });
      if (result.eligible) m.set(ticker, { up: result.upsidePct, asOf: fund["Fundamentals As Of"] as string | undefined });
    });
    return m;
  }, [bundle]);

  const [selectedSetups, setSelectedSetups] = useState<string[]>([]);
  const [setupDir, setSetupDir] = useState<Record<string, Dir>>({});
  const [conf, setConf] = useState<Conf>("AND");
  const [fTicker, setFTicker] = useState("");
  const [fSector, setFSector] = useState("");
  const [fLiq, setFLiq] = useState("");
  const [fDcf, setFDcf] = useState("");
  const [fKonglo, setFKonglo] = useState("");
  const [fPrice, setFPrice] = useState("");
  // Which σ variant each merged Near-VWAP box currently points at (boxId → "1"|"2").
  const [sigmaSel, setSigmaSel] = useState<Record<string, "1" | "2">>({});
  const [sortCol, setSortCol] = useState("fresh");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("desc");
  const [showN, setShowN] = useState(25);
  const [page, setPage] = useState(0);
  const [savedPresets, setSavedPresets] = useState<SavedPreset[]>([]);

  useEffect(() => setSavedPresets(readPresets()), []);

  useEffect(() => {
    if (!marketDate) return;
    let cancelled = false;
    setUniverse(null);
    setError(null);
    loadUniverse(marketDate)
      .then((u) => !cancelled && setUniverse(u))
      .catch(() => !cancelled && setError(`No screener universe for ${marketDate} — the workbook is published on the daily pipeline; older dates may not have one.`));
    return () => {
      cancelled = true;
    };
  }, [marketDate]);

  const rows = universe?.rows ?? [];
  const solo = useMemo(() => soloCounts(rows), [rows]);

  // ── matching ──
  const passGeneral = useCallback(
    (r: UniverseRow) => {
      if (fTicker && !r.ticker.toLowerCase().includes(fTicker.toLowerCase())) return false;
      if (fSector && secCode(r.ticker, r.sectorCode) !== fSector) return false;
      if (!passLiquidity(r, fLiq)) return false;
      if (!passKonglo(kongloMap, r.ticker, fKonglo)) return false;
      if (!passPrice(r, fPrice)) return false;
      if (fDcf) {
        const dcf = dcfByTicker.get(r.ticker);
        if (dcf === undefined) return false;
        if (fDcf === "under" && dcf.up < 0.1) return false;
        if (fDcf === "over" && dcf.up > -0.1) return false;
      }
      return true;
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps -- secCode is stable per kseiSec
    [fTicker, fSector, fLiq, fDcf, fKonglo, fPrice, dcfByTicker, kseiSec, kongloMap],
  );

  const filtered = useMemo(() => {
    let out = rows.filter(passGeneral);
    if (selectedSetups.length) {
      out = out.filter((r) => {
        const tests = selectedSetups.map((k) => {
          const dir = setupDir[k] || "bull";
          const s = setupByKey(k);
          if (!s) return false;
          return dir === "bear" && s.hasBear ? r.setupsBear.includes(k) : r.setupsMatched.includes(k);
        });
        return conf === "AND" ? tests.every(Boolean) : tests.some(Boolean);
      });
    }
    return out;
  }, [rows, passGeneral, selectedSetups, setupDir, conf]);

  const sorted = useMemo(() => {
    const dir = sortDir === "asc" ? 1 : -1;
    const key = (r: UniverseRow): number | string => {
      switch (sortCol) {
        case "ticker": return r.ticker;
        case "sector": return secLabel(r.ticker, r.sectorLabel);
        case "konglo": return (kongloMap[r.ticker] || []).join(",");
        case "setup": return r.setupsMatched.length;
        case "price": return r.price ?? -Infinity;
        case "chg": return r.chg;
        case "beta": return r.beta ?? -Infinity;
        case "rvol": return r.rvol ?? -Infinity;
        case "rsi": return r.rsi14 ?? -Infinity;
        default: return r.freshRank;
      }
    };
    return filtered.slice().sort((a, b) => {
      const ka = key(a);
      const kb = key(b);
      if (typeof ka === "string") return ka.localeCompare(kb as string) * dir;
      return ((ka as number) - (kb as number)) * dir;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps -- secLabel is stable per kseiSec
  }, [filtered, sortCol, sortDir, kongloMap, kseiSec]);

  const matchCount = sorted.length;
  const pages = Math.max(1, Math.ceil(matchCount / showN));
  const start = page * showN;
  const pageRows = sorted.slice(start, start + showN);

  // reset to first page when the result set changes size out from under us
  useEffect(() => {
    setPage(0);
  }, [selectedSetups, setupDir, conf, fTicker, fSector, fLiq, fKonglo, fPrice, showN]);

  // ── handlers ──
  const toggleSetup = (k: string) =>
    setSelectedSetups((s) => (s.includes(k) ? s.filter((x) => x !== k) : [...s, k]));
  const setDir = (k: string, d: Dir) => {
    setSetupDir((m) => ({ ...m, [k]: d }));
    setSelectedSetups((s) => (s.includes(k) ? s : [...s, k]));
  };
  const setSort = (col: string) => {
    if (sortCol === col && sortDir === "desc") setSortDir("asc");
    else {
      setSortCol(col);
      setSortDir("desc");
    }
  };
  const savePreset = () => {
    const name = typeof window !== "undefined" ? window.prompt("Preset name:", "My confluence") : null;
    if (!name) return;
    const p: SavedPreset = { id: `p${Date.now()}`, name, selectedSetups, setupDir, conf };
    const next = [...savedPresets.filter((x) => x.name !== name), p];
    try {
      window.localStorage.setItem(PRESETS_KEY, JSON.stringify(next));
    } catch {
      /* localStorage unavailable — preset is session-only */
    }
    setSavedPresets(next);
  };
  const loadPreset = (id: string) => {
    const p = savedPresets.find((x) => x.id === id);
    if (!p) return;
    setSelectedSetups(p.selectedSetups || []);
    setSetupDir(p.setupDir || {});
    setConf(p.conf || "AND");
  };

  const csv = () => {
    const head = ["Ticker", "Sector", "Konglo", "Setup", "Price", "Change %", "Beta", "Avg Vol 20D", "RVOL", "Moving Average", "Previous QVWAP", "Previous YVWAP", "Monday Range", "IBH", "IBL", "IBH/IBL Zone", "RSI", "RSI MA", "Internal Structure", "Swing Structure", "News"];
    const esc = (s: unknown) => `"${String(s ?? "").replace(/"/g, '""')}"`;
    const lines = sorted.map((r) => {
      const news = newsByTicker.get(r.ticker);
      const zone = ibZone(r.price, r.ibh, r.ibl);
      return [
        r.ticker, secLabel(r.ticker, r.sectorLabel), (kongloMap[r.ticker] || []).join("|") || "no data",
        r.setupsMatched.map(setupDisplayLabel).join("|") || "-",
        r.price ?? "", (r.chg * 100).toFixed(2), r.beta ?? "no data", r.avgVolume20 ?? "no data", r.rvol ?? "",
        r.maZoneReal ?? "no data",
        r.vwapPq ? `${r.vwapPq.vwap.toFixed(0)} (${r.vwapPq.sigma?.toFixed(2) ?? "?"}σ)` : "no data",
        r.vwapPy ? `${r.vwapPy.vwap.toFixed(0)} (${r.vwapPy.sigma?.toFixed(2) ?? "?"}σ)` : "no data",
        r.mondayRange ? `${r.mondayRange.status} (H${r.mondayRange.high} L${r.mondayRange.low})` : "no data",
        r.ibh ?? "no data", r.ibl ?? "no data", zone ?? "no data",
        r.rsi14?.toFixed(1) ?? "no data", r.rsiMa14?.toFixed(1) ?? "no data",
        r.structureInternal ?? "no data", r.structureSwing ?? "no data",
        news ? news.title : "no data",
      ].map(esc).join(",");
    });
    const blob = new Blob([`${head.map(esc).join(",")}\n${lines.join("\n")}`], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `idx-setups-${universe?.marketDate || marketDate}.csv`;
    a.click();
  };

  const pills: Array<{ label: string; color: string; bg: string; border: string; onRemove: () => void }> = [];
  selectedSetups.forEach((k) => {
    const s = setupByKey(k);
    if (!s) return;
    const bear = s.hasBear && (setupDir[k] || "bull") === "bear";
    pills.push({ label: setupDisplayLabel(k) + (s.hasBear ? ` · ${bear ? "bear" : "bull"}` : ""), color: bear ? "var(--down)" : "var(--up)", bg: bear ? "var(--downSoft)" : "var(--upSoft)", border: "transparent", onRemove: () => toggleSetup(k) });
  });
  if (fSector) {
    const so = SECTOR_OPTIONS.find((x) => x.v === fSector);
    pills.push({ label: `Sector: ${so ? so.label : fSector}`, color: "var(--muted)", bg: "var(--soft)", border: "var(--border)", onRemove: () => setFSector("") });
  }
  if (fLiq) {
    const lo = LIQUIDITY_OPTIONS.find((x) => x.v === fLiq);
    pills.push({ label: lo ? lo.label : fLiq, color: "var(--muted)", bg: "var(--soft)", border: "var(--border)", onRemove: () => setFLiq("") });
  }
  if (fKonglo) pills.push({ label: `Konglo: ${fKonglo}`, color: "var(--muted)", bg: "var(--soft)", border: "var(--border)", onRemove: () => setFKonglo("") });
  if (fPrice) pills.push({ label: fPrice === "under50" ? "Price < 50" : "Price ≥ 50", color: "var(--muted)", bg: "var(--soft)", border: "var(--border)", onRemove: () => setFPrice("") });
  if (fTicker) pills.push({ label: `Ticker: "${fTicker}"`, color: "var(--muted)", bg: "var(--soft)", border: "var(--border)", onRemove: () => setFTicker("") });
  if (fDcf) pills.push({ label: fDcf === "under" ? "DCF: Undervalued (≥10% upside)" : "DCF: Overvalued (≤10% downside)", color: "var(--accent)", bg: "var(--accentSoft)", border: "var(--accent-border)", onRemove: () => setFDcf("") });

  const headers: Array<[string, string, CSSProperties["justifyContent"], boolean]> = [
    ["ticker", "TICKER", "flex-start", true],
    ["sector", "SECTOR", "flex-start", false],
    ["konglo", "KONGLO", "flex-start", false],
    ["setup", "SETUP", "flex-start", false],
    ["price", "PRICE", "flex-end", false],
    ["chg", "CHANGE %", "flex-end", false],
    ["beta", "BETA (vs IHSG)", "flex-end", false],
    ["avgVol20", "AVG VOL 20D", "flex-end", false],
    ["rvol", "RVOL", "flex-end", false],
    ["ma", "MOVING AVERAGE", "flex-start", false],
    ["pqvwap", "PREVIOUS QVWAP", "flex-start", false],
    ["pyvwap", "PREVIOUS YVWAP", "flex-start", false],
    ["mondayRange", "MONDAY RANGE", "flex-start", false],
    ["ibhIbl", "IBH / IBL", "flex-start", false],
    ["rsi", "RSI", "flex-start", false],
    ["structInternal", "LATEST INTERNAL STRUCTURE", "flex-start", false],
    ["structSwing", "LATEST SWING STRUCTURE", "flex-start", false],
    ["news", "NEWS INTELLIGENCE", "flex-start", false],
  ];

  const total = universe?.totalUniverse ?? 0;

  return (
    <section>
      {/* title */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6, flexWrap: "wrap" }}>
          <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800, letterSpacing: "-.02em" }}>Setups Screener</h1>
          <span style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".1em", color: "var(--muted)", background: "var(--soft)", border: "1px solid var(--border)", borderRadius: 999, padding: "3px 9px" }}>CONFLUENCE</span>
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 11.5, color: "var(--faint)" }}>
            Matching <strong style={{ color: "var(--text)", fontFamily: MONO }}>{matchCount}</strong> of {total} · scanned {universe?.scanned ?? 0} · {universe?.marketDate || marketDate}
          </span>
        </div>
        <p style={{ margin: 0, color: "var(--muted)", fontSize: 13.5, maxWidth: 760, lineHeight: 1.5 }}>
          Pick the setups you trade or build a custom context — they combine into a confluence screen over the real workbook. Every cell is an interpreted judgment from a real field; hover for the raw value.
        </p>
      </div>

      {error ? <div style={{ ...CARD, padding: "16px 18px", color: "var(--muted)", fontSize: 13 }}>{error}</div> : null}

      {/* presets */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ flex: 1 }} />
        {savedPresets.length ? (
          <select onChange={(e) => e.target.value && loadPreset(e.target.value)} defaultValue="" style={{ fontSize: 11.5, fontWeight: 600, color: "var(--muted)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 9, padding: "7px 10px", cursor: "pointer" }}>
            <option value="">Saved presets…</option>
            {savedPresets.map((p) => (
              <option key={p.id} value={p.id}>{p.name}</option>
            ))}
          </select>
        ) : null}
        <button type="button" onClick={savePreset} style={{ fontSize: 11.5, fontWeight: 700, color: "var(--accent)", background: "var(--accentSoft)", border: "1px solid var(--accent-border)", borderRadius: 9, padding: "7px 13px", cursor: "pointer" }}>＋ Save preset</button>
      </div>

      {/* general filter bar */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 7, ...CARD, borderRadius: 10, padding: "7px 11px" }}>
          <span style={{ color: "var(--faint)", fontSize: 13 }}>⌕</span>
          <input value={fTicker} onChange={(e) => setFTicker(e.target.value)} placeholder="Ticker…" aria-label="Filter by ticker" style={{ border: "none", outline: "none", background: "transparent", fontFamily: MONO, fontSize: 12, color: "var(--text)", width: 120 }} />
        </div>
        <select value={fSector} onChange={(e) => setFSector(e.target.value)} aria-label="Filter by sector" style={{ fontSize: 12, fontWeight: 600, color: "var(--text)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 10, padding: "8px 11px", cursor: "pointer" }}>
          {SECTOR_OPTIONS.map((o) => (<option key={o.v} value={o.v}>{o.label}</option>))}
        </select>
        <select value={fKonglo} onChange={(e) => setFKonglo(e.target.value)} aria-label="Filter by konglomerate group" title="Real konglomerate-group membership (docs/data/indexes.json)" style={{ fontSize: 12, fontWeight: 600, color: "var(--text)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 10, padding: "8px 11px", cursor: "pointer", maxWidth: 220 }}>
          {kongloOptions.map((o) => (<option key={o.v} value={o.v}>{o.label}</option>))}
        </select>
        <select value={fLiq} onChange={(e) => setFLiq(e.target.value)} aria-label="Filter by liquidity" style={{ fontSize: 12, fontWeight: 600, color: "var(--text)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 10, padding: "8px 11px", cursor: "pointer" }}>
          {LIQUIDITY_OPTIONS.map((o) => (<option key={o.v} value={o.v}>{o.label}</option>))}
        </select>
        <select value={fPrice} onChange={(e) => setFPrice(e.target.value)} aria-label="Filter by price" style={{ fontSize: 12, fontWeight: 600, color: "var(--text)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 10, padding: "8px 11px", cursor: "pointer" }}>
          {PRICE_OPTIONS.map((o) => (<option key={o.v} value={o.v}>{o.label}</option>))}
        </select>
        <select value={fDcf} onChange={(e) => setFDcf(e.target.value)} aria-label="Filter by DCF valuation" title="DCF model: real Free Cash Flow / beta / price, fixed macro assumptions — open a ticker to adjust" style={{ fontSize: 12, fontWeight: 600, color: "var(--text)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 10, padding: "8px 11px", cursor: "pointer" }}>
          <option value="">DCF: any</option>
          <option value="under">DCF: Undervalued (≥10% upside)</option>
          <option value="over">DCF: Overvalued (≤10% downside)</option>
        </select>
        <div style={{ flex: 1 }} />
        <span style={{ fontSize: 10, color: "var(--faint)" }}>Sector, liquidity &amp; konglo groups are real. DCF uses fixed default assumptions here — open a ticker for adjustable sliders.</span>
      </div>

      {/* setup library */}
      <div>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 11, flexWrap: "wrap" }}>
            <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".12em", color: "var(--faint)" }}>ENTRY SETUPS · MULTI-SELECT</span>
            <span style={{ fontSize: 11, color: "var(--muted)" }}>pick the setups you trade — they combine into a confluence screen</span>
            <div style={{ flex: 1 }} />
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 10.5, fontWeight: 600, color: "var(--faint)" }}>Combine</span>
              <div style={{ display: "flex", gap: 3, background: "var(--soft)", border: "1px solid var(--border)", borderRadius: 8, padding: 3 }}>
                {(["AND", "OR"] as const).map((c) => (
                  <button key={c} type="button" onClick={() => setConf(c)} title={c === "AND" ? "All selected setups must match" : "Any selected setup matches"} style={{ fontFamily: MONO, fontSize: 10.5, fontWeight: 700, padding: "4px 11px", borderRadius: 6, border: "none", cursor: "pointer", background: conf === c ? "var(--accent)" : "transparent", color: conf === c ? "#fff" : "var(--muted)" }}>{c}</button>
                ))}
              </div>
              {selectedSetups.length ? (
                <button type="button" onClick={() => setSelectedSetups([])} style={{ fontSize: 10.5, fontWeight: 700, color: "var(--muted)", background: "transparent", border: "1px solid var(--border)", borderRadius: 8, padding: "5px 10px", cursor: "pointer" }}>Clear</button>
              ) : null}
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(210px,1fr))", gap: 8, marginBottom: 16 }}>
            {ungroupedSetups().map((s) => {
              const on = selectedSetups.includes(s.key);
              const dir = setupDir[s.key] || "bull";
              const cnt = solo[s.key] ? (dir === "bear" ? solo[s.key].bear : solo[s.key].bull) : 0;
              return (
                <div key={s.key} role="button" tabIndex={0} aria-pressed={on} title={s.req} onClick={() => toggleSetup(s.key)} onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggleSetup(s.key); } }} style={{ border: `1.5px solid ${on ? "var(--accent-border)" : "var(--border)"}`, background: on ? "var(--accentSoft)" : "var(--panel)", borderRadius: 11, padding: "9px 11px", boxShadow: "var(--sh, var(--shadow))", cursor: "pointer" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ width: 24, height: 24, flex: "none", borderRadius: 7, background: on ? "var(--accent)" : "var(--soft)", color: on ? "#fff" : "var(--muted)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12.5 }} aria-hidden>{s.icon}</span>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ fontSize: 12, fontWeight: 800, letterSpacing: "-.01em", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.label}</span>
                        {/* DESIGN_SPEC §3.3: evidence strength is explicit, not a glyph. */}
                        <span
                          title={s.inferred ? "INFERRED — reconstructed from other fields, no dedicated column" : "EXACT — reads a real column directly"}
                          style={{ flex: "none", fontSize: 8, fontWeight: 700, letterSpacing: ".06em", padding: "1px 5px", borderRadius: 5, background: s.inferred ? "var(--warnSoft)" : "var(--upSoft)", color: s.inferred ? "var(--warning)" : "var(--up)" }}
                        >
                          {s.inferred ? "INFERRED" : "EXACT"}
                        </span>
                      </div>
                      <div style={{ fontFamily: MONO, fontSize: 9.5, color: cnt ? "var(--accent)" : "var(--faint)", fontWeight: 700 }}>{cnt} {cnt === 1 ? "ticker" : "tickers"} today</div>
                    </div>
                    <span style={{ width: 18, height: 18, flex: "none", borderRadius: 6, border: `1.5px solid ${on ? "var(--accent)" : "var(--border)"}`, background: on ? "var(--accent)" : "transparent", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 11 }} aria-hidden>{on ? "✓" : ""}</span>
                  </div>
                  {s.hasBear ? (
                    <div style={{ display: "flex", gap: 3, marginTop: 7 }} onClick={(e) => e.stopPropagation()}>
                      {(["bull", "bear"] as const).map((d) => (
                        <button key={d} type="button" onClick={() => setDir(s.key, d)} style={{ flex: 1, fontSize: 9.5, fontWeight: 700, padding: "4px 6px", borderRadius: 6, border: `1px solid ${dir === d ? "transparent" : "var(--border)"}`, background: dir === d ? (d === "bull" ? "var(--upSoft)" : "var(--downSoft)") : "transparent", color: dir === d ? (d === "bull" ? "var(--up)" : "var(--down)") : "var(--muted)", cursor: "pointer" }}>{d === "bull" ? "Bullish" : "Bearish"}</button>
                      ))}
                    </div>
                  ) : null}
                </div>
              );
            })}
            {VWAP_GROUPS.map((g) => {
              const sigma = sigmaSel[g.boxId] || "1";
              const variant = g.variants.find((v) => v.sigmaLabel.includes(sigma)) || g.variants[0];
              const on = selectedSetups.includes(variant.setupKey);
              const cnt = solo[variant.setupKey]?.bull ?? 0;
              return (
                <div key={g.boxId} role="button" tabIndex={0} aria-pressed={on} title={g.req} onClick={() => toggleSetup(variant.setupKey)} onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggleSetup(variant.setupKey); } }} style={{ border: `1.5px solid ${on ? "var(--accent-border)" : "var(--border)"}`, background: on ? "var(--accentSoft)" : "var(--panel)", borderRadius: 11, padding: "9px 11px", boxShadow: "var(--sh, var(--shadow))", cursor: "pointer" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ width: 24, height: 24, flex: "none", borderRadius: 7, background: on ? "var(--accent)" : "var(--soft)", color: on ? "#fff" : "var(--muted)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12.5 }} aria-hidden>{g.icon}</span>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ fontSize: 12, fontWeight: 800, letterSpacing: "-.01em", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{g.label}</span>
                        <span title="EXACT — reads a real column directly" style={{ flex: "none", fontSize: 8, fontWeight: 700, letterSpacing: ".06em", padding: "1px 5px", borderRadius: 5, background: "var(--upSoft)", color: "var(--up)" }}>EXACT</span>
                      </div>
                      <div style={{ fontFamily: MONO, fontSize: 9.5, color: cnt ? "var(--accent)" : "var(--faint)", fontWeight: 700 }}>{cnt} {cnt === 1 ? "ticker" : "tickers"} today · {variant.sigmaLabel}</div>
                    </div>
                    <span style={{ width: 18, height: 18, flex: "none", borderRadius: 6, border: `1.5px solid ${on ? "var(--accent)" : "var(--border)"}`, background: on ? "var(--accent)" : "transparent", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 11 }} aria-hidden>{on ? "✓" : ""}</span>
                  </div>
                  <div style={{ display: "flex", gap: 3, marginTop: 7 }} onClick={(e) => e.stopPropagation()}>
                    {g.variants.map((v, vi) => {
                      const vSigma = vi === 0 ? "1" : "2";
                      const vOn = sigma === vSigma;
                      return (
                        <button
                          key={v.setupKey}
                          type="button"
                          onClick={() => {
                            // Switching σ on an already-selected box swaps which
                            // variant key is active in selectedSetups, so the
                            // filter follows the toggle instead of silently
                            // keeping the old σ's key selected underneath it.
                            setSigmaSel((m) => ({ ...m, [g.boxId]: vSigma }));
                            if (on) {
                              setSelectedSetups((s) => [...s.filter((k) => k !== variant.setupKey), v.setupKey]);
                            }
                          }}
                          style={{ flex: 1, fontSize: 9.5, fontWeight: 700, padding: "4px 6px", borderRadius: 6, border: `1px solid ${vOn ? "transparent" : "var(--border)"}`, background: vOn ? "var(--accentSoft)" : "transparent", color: vOn ? "var(--accent)" : "var(--muted)", cursor: "pointer" }}
                        >
                          {v.sigmaLabel}
                        </button>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
      </div>

      {/* pills + toolbar */}
      <div style={{ display: "flex", alignItems: "center", gap: 9, flexWrap: "wrap", marginBottom: 12 }}>
        {pills.length ? (
          pills.map((p, i) => (
            <span key={i} style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11, fontWeight: 700, color: p.color, background: p.bg, border: `1px solid ${p.border}`, borderRadius: 999, padding: "4px 6px 4px 11px" }}>
              {p.label}
              <button type="button" onClick={p.onRemove} aria-label={`Remove ${p.label}`} style={{ width: 16, height: 16, borderRadius: "50%", border: "none", background: "transparent", color: "inherit", cursor: "pointer", fontSize: 12, lineHeight: 1 }}>×</button>
            </span>
          ))
        ) : (
          <span style={{ fontSize: 11, color: "var(--faint)" }}>No filters — showing the full scanned universe, freshest signals first.</span>
        )}
        <div style={{ flex: 1 }} />
        <span style={{ fontSize: 11.5, color: "var(--muted)", fontFamily: MONO }}><strong style={{ color: "var(--text)" }}>{matchCount}</strong> of {total}</span>
        <select value={showN} onChange={(e) => setShowN(Number(e.target.value))} aria-label="Rows per page" style={{ fontSize: 11, fontWeight: 600, color: "var(--muted)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 8, padding: "6px 9px", cursor: "pointer" }}>
          {[25, 50, 100, 250].map((n) => (<option key={n} value={n}>Show {n}</option>))}
        </select>
        <button type="button" onClick={csv} style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 8, padding: "6px 12px", cursor: "pointer" }}>⭳ CSV</button>
      </div>

      {/* results table -- fixed-height viewport (~10 rows) with its own
          vertical scrollbar, so a large match count doesn't stretch the
          whole page; horizontal scroll stays separate (18 columns is too
          much to compress into one screen width without hurting
          readability, and per-column widths already vary by content). */}
      <div style={{ ...CARD, overflow: "hidden" }}>
        <div style={{ overflowX: "auto", overflowY: "auto", maxHeight: 645 }}>
          <div style={{ minWidth: 1428 }}>
            <div style={{ position: "sticky", top: 0, zIndex: 20, display: "grid", gridTemplateColumns: GRID, background: "var(--soft)", borderBottom: "1px solid var(--border)" }}>
              {headers.map(([col, label, justify, sticky]) => (
                <button key={col} type="button" onClick={() => setSort(col)} style={{ display: "flex", alignItems: "center", gap: 4, justifyContent: justify, padding: "10px 12px", border: "none", background: sticky ? "var(--soft)" : "transparent", cursor: "pointer", fontSize: 9, fontWeight: 700, letterSpacing: ".06em", color: sortCol === col ? "var(--text)" : "var(--faint)", textAlign: "left", ...(sticky ? { position: "sticky", left: 0, zIndex: 6 } : {}) }}>
                  {label}<span style={{ color: "var(--accent)" }}>{sortCol === col ? (sortDir === "desc" ? "▼" : "▲") : ""}</span>
                </button>
              ))}
            </div>
            {universe && !pageRows.length ? (
              <div style={{ padding: 44, textAlign: "center", color: "var(--faint)", fontSize: 13 }}>No tickers matched this filter on the market date ({universe.marketDate}). Loosen a condition or switch AND → OR.</div>
            ) : null}
            {!universe && !error ? (
              <div style={{ padding: 44, textAlign: "center", color: "var(--faint)", fontSize: 13 }}>Loading the workbook universe…</div>
            ) : null}
            {pageRows.map((r) => {
              const rowBg = "var(--panel)";
              const konglo = kongloMap[r.ticker] || [];
              const zone = ibZone(r.price, r.ibh, r.ibl);
              const news = newsByTicker.get(r.ticker);
              const NoData = () => <span style={{ fontSize: 10, color: "var(--faint)", fontStyle: "italic" }}>no data</span>;
              return (
                <div key={r.ticker} role="button" tabIndex={0} onClick={() => openTicker(r.ticker)} onKeyDown={(e) => { if (e.key === "Enter") openTicker(r.ticker); }} style={{ display: "grid", gridTemplateColumns: GRID, borderBottom: "1px solid var(--hair)", color: "var(--text)", background: rowBg, cursor: "pointer" }}>
                  {/* TICKER (sticky) */}
                  <div style={{ position: "sticky", left: 0, zIndex: 5, background: rowBg, padding: "9px 12px", borderRight: "1px solid var(--hair)" }}>
                    <div style={{ fontFamily: MONO, fontSize: 13, fontWeight: 800 }}>{r.ticker}</div>
                  </div>
                  {/* SECTOR */}
                  <div style={{ padding: "9px 10px", fontSize: 11, color: "var(--muted)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={secLabel(r.ticker, r.sectorLabel)}>{secLabel(r.ticker, r.sectorLabel)}</div>
                  {/* KONGLO */}
                  <div style={{ padding: "9px 10px", fontSize: 10.5, color: konglo.length ? "var(--text)" : "var(--faint)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={konglo.join(", ")}>{konglo.length ? konglo.join(", ") : "—"}</div>
                  {/* SETUP */}
                  <div style={{ padding: "9px 12px", display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap" }}>
                    {r.setupsMatched.map((k) => {
                      const label = setupDisplayLabel(k);
                      // Pivot date shown != "today" for both RSI divergence
                      // setups, same reason for each: both are reversal-style
                      // confirmed-swing pivots, so confirmation genuinely lags
                      // the more recent pivot by swing_window bars.
                      const pivot = k === "rsi_divergence" ? r.rsiDivBullishPivotDate : k === "rsi_divergence_hidden" ? r.rsiDivHiddenBullishPivotDate : null;
                      const tip = pivot ? `${label} (Bull) · pivot ${pivot} · confirmed today` : label;
                      return <span key={`${k}-bull`} title={tip} style={{ fontSize: 9, fontWeight: 700, color: "var(--up)", background: "var(--upSoft)", borderRadius: 5, padding: "2px 6px", whiteSpace: "nowrap" }}>{label}{pivot ? <span style={{ opacity: 0.7, fontWeight: 600 }}> · piv {pivot}</span> : null}</span>;
                    })}
                    {r.setupsBear.map((k) => {
                      const label = setupDisplayLabel(k);
                      const pivot = k === "rsi_divergence" ? r.rsiDivBearishPivotDate : k === "rsi_divergence_hidden" ? r.rsiDivHiddenBearishPivotDate : null;
                      const tip = pivot ? `${label} (Bear) · pivot ${pivot} · confirmed today` : `${label} (Bear)`;
                      return <span key={`${k}-bear`} title={tip} style={{ fontSize: 9, fontWeight: 700, color: "var(--down)", background: "var(--downSoft)", borderRadius: 5, padding: "2px 6px", whiteSpace: "nowrap" }}>{label}{pivot ? <span style={{ opacity: 0.7, fontWeight: 600 }}> · piv {pivot}</span> : null}</span>;
                    })}
                    {!r.setupsMatched.length && !r.setupsBear.length ? <span style={{ fontSize: 9.5, color: "var(--faint)" }}>—</span> : null}
                  </div>
                  {/* PRICE */}
                  <div style={{ padding: "9px 8px", textAlign: "right" }}><span style={{ fontFamily: MONO, fontSize: 11.5, fontWeight: 700 }}>{r.price == null ? "—" : formatPrice(r.price)}</span></div>
                  {/* CHANGE % */}
                  <div style={{ padding: "9px 8px", textAlign: "right" }}><span style={{ fontFamily: MONO, fontSize: 11.5, fontWeight: 700, color: r.chg > 0 ? "var(--up)" : r.chg < 0 ? "var(--down)" : "var(--flat)" }}>{r.chg > 0 ? "▲ " : r.chg < 0 ? "▼ " : ""}{formatPercent(r.chg)}</span></div>
                  {/* BETA */}
                  <div style={{ padding: "9px 8px", textAlign: "right" }}>{r.beta == null ? <NoData /> : <span style={{ fontFamily: MONO, fontSize: 11 }}>{r.beta.toFixed(2)}</span>}</div>
                  {/* AVG VOL 20D */}
                  <div style={{ padding: "9px 8px", textAlign: "right" }}>{r.avgVolume20 == null ? <NoData /> : <span style={{ fontFamily: MONO, fontSize: 11 }}>{r.avgVolume20}</span>}</div>
                  {/* RVOL */}
                  <div style={{ padding: "9px 8px", textAlign: "right" }}>{r.rvol == null ? <NoData /> : <span style={{ fontFamily: MONO, fontSize: 11, fontWeight: 700, color: r.rvol >= 1.5 ? "var(--up)" : "var(--muted)" }}>{r.rvol.toFixed(2)}×</span>}</div>
                  {/* MOVING AVERAGE (MA zone only, RSI split into its own column) */}
                  <div style={{ padding: "9px 10px", fontSize: 10.5, lineHeight: 1.3 }} title={r.maZoneReal ?? undefined}>{r.maZoneReal == null ? <NoData /> : <span style={{ color: /Above All/i.test(r.maZoneReal) ? "var(--up)" : /Below All/i.test(r.maZoneReal) ? "var(--down)" : "var(--text)" }}>{r.maZoneReal}</span>}</div>
                  {/* PREVIOUS QVWAP */}
                  <div style={{ padding: "9px 10px" }} title={r.vwapPq?.zone}>{!r.vwapPq ? <NoData /> : <><span style={{ fontFamily: MONO, fontSize: 11, fontWeight: 700, color: r.price != null && r.price >= r.vwapPq.vwap ? "var(--up)" : "var(--down)" }}>{vwapCellText(r.vwapPq)}</span>{r.vwapPq.zone ? <div style={{ fontSize: 9, color: "var(--faint)", marginTop: 1 }}>{r.vwapPq.zone}</div> : null}</>}</div>
                  {/* PREVIOUS YVWAP */}
                  <div style={{ padding: "9px 10px" }} title={r.vwapPy?.zone}>{!r.vwapPy ? <NoData /> : <><span style={{ fontFamily: MONO, fontSize: 11, fontWeight: 700, color: r.price != null && r.price >= r.vwapPy.vwap ? "var(--up)" : "var(--down)" }}>{vwapCellText(r.vwapPy)}</span>{r.vwapPy.zone ? <div style={{ fontSize: 9, color: "var(--faint)", marginTop: 1 }}>{r.vwapPy.zone}</div> : null}</>}</div>
                  {/* MONDAY RANGE */}
                  <div style={{ padding: "9px 10px" }} title={r.mondayRange ? `High ${formatPrice(r.mondayRange.high)} · Low ${formatPrice(r.mondayRange.low)} · ${r.mondayRange.date}` : undefined}>
                    {!r.mondayRange ? <NoData /> : (
                      <>
                        <span style={{ fontSize: 11, fontWeight: 700, color: toneColor(mondayRangeTone(r.mondayRange.status)) }}>{r.mondayRange.status}</span>
                        <div style={{ fontFamily: MONO, fontSize: 9, color: "var(--faint)", marginTop: 1 }}>{formatPrice(r.mondayRange.low)}–{formatPrice(r.mondayRange.high)}</div>
                      </>
                    )}
                  </div>
                  {/* IBH / IBL */}
                  <div style={{ padding: "9px 10px" }}>
                    {r.ibh == null || r.ibl == null ? <NoData /> : (
                      <>
                        <span style={{ fontFamily: MONO, fontSize: 10.5 }}>{formatPrice(r.ibh)} / {formatPrice(r.ibl)}</span>
                        <div style={{ fontSize: 9.5, fontWeight: 700, color: toneColor(ibZoneTone(zone)), marginTop: 1 }}>{zone}</div>
                      </>
                    )}
                  </div>
                  {/* RSI */}
                  <div style={{ padding: "9px 10px" }}>
                    {r.rsi14 == null ? <NoData /> : (
                      <>
                        <span style={{ fontFamily: MONO, fontSize: 12, fontWeight: 800, color: r.rsi14 >= 70 ? "var(--down)" : r.rsi14 <= 30 ? "var(--up)" : "var(--text)" }}>{r.rsi14.toFixed(1)}</span>
                        {r.rsiMa14 != null ? <span style={{ fontFamily: MONO, fontSize: 9.5, color: "var(--faint)", marginLeft: 5 }}>MA {r.rsiMa14.toFixed(1)}</span> : null}
                      </>
                    )}
                  </div>
                  {/* LATEST INTERNAL STRUCTURE */}
                  <div style={{ padding: "9px 10px", fontSize: 10.5, color: toneColor(structureTone(r.structureInternal)), lineHeight: 1.3 }}>{r.structureInternal ?? <NoData />}</div>
                  {/* LATEST SWING STRUCTURE */}
                  <div style={{ padding: "9px 10px", fontSize: 10.5, color: toneColor(structureTone(r.structureSwing)), lineHeight: 1.3 }}>{r.structureSwing ?? <NoData />}</div>
                  {/* NEWS INTELLIGENCE */}
                  <div style={{ padding: "9px 10px", fontSize: 10.5, lineHeight: 1.35, color: news ? toneColor(news.tone) : undefined, overflow: "hidden", textOverflow: "ellipsis", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical" }} title={news?.title}>
                    {news ? news.title : <NoData />}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
        {matchCount > showN ? (
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 16px", borderTop: "1px solid var(--border)" }}>
            <span style={{ fontSize: 11, color: "var(--muted)" }}>{matchCount ? `${start + 1}–${Math.min(start + showN, matchCount)} of ${matchCount}` : "0"}</span>
            <div style={{ flex: 1 }} />
            <button type="button" onClick={() => setPage((p) => Math.max(0, p - 1))} style={{ fontSize: 11, fontWeight: 700, color: page > 0 ? "var(--text)" : "var(--faint)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 8, padding: "6px 12px", cursor: "pointer" }}>‹ Prev</button>
            <span style={{ fontSize: 11, fontFamily: MONO, color: "var(--muted)" }}>{page + 1} / {pages}</span>
            <button type="button" onClick={() => setPage((p) => Math.min(pages - 1, p + 1))} style={{ fontSize: 11, fontWeight: 700, color: page < pages - 1 ? "var(--text)" : "var(--faint)", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 8, padding: "6px 12px", cursor: "pointer" }}>Next ›</button>
          </div>
        ) : null}
      </div>

      <div style={{ fontSize: 10.5, color: "var(--faint)", lineHeight: 1.5, maxWidth: 900, marginTop: 14 }}>
        Real IDX workbook · {universe?.marketDate || marketDate}. Every cell is a real published field — hover for extra detail, click a row to open the ticker. Konglo = real konglomerate-group membership; Monday Range = the current week&apos;s first session&apos;s high/low vs today&apos;s close; News Intelligence is joined from the News page&apos;s own data. Missing fields render an explicit &ldquo;no data&rdquo; rather than a guess. Entry/Invalidation/Target/R:R and DCF Upside moved off this table — open a ticker&apos;s own page for the full Trade Plan and adjustable DCF model (the DCF filter above still works).
      </div>
    </section>
  );
}
