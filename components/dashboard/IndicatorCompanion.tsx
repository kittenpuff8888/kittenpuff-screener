"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, ReactNode } from "react";
import {
  createChart, CandlestickSeries, HistogramSeries, LineSeries,
  LineStyle, CrosshairMode,
  type IChartApi, type Time, type MouseEventParams, type LineWidth,
} from "lightweight-charts";
import type { OhlcvPayload, OhlcvRow } from "@/lib/domain/types";
import { computeInitialBalance } from "@/lib/indicators/initialBalance";
import { computeAnchoredVwap, periodLabel, VWAP_SOURCE_FN, type AvwapPoint, type VwapAnchor, type VwapSource } from "@/lib/indicators/anchoredVwap";
import { ema, sma, rsiWilder, stochOf, macdOf, dmi } from "@/lib/indicators/oscillators";
import { computeDivergences, type DivergenceType } from "@/lib/indicators/divergence";
import { formatPrice, formatCompact, formatPercent } from "@/lib/format/number";
import { DrawingTool, type SelectionInfo } from "@/lib/charting/DrawingTool";
import { loadViewRange, saveViewRange } from "@/lib/charting/viewRangeStore";
import { BandFill } from "@/lib/charting/drawings/BandFill";
import { IbBoxes } from "@/lib/charting/drawings/IbBoxes";
import { VwapFill } from "@/lib/charting/drawings/VwapFill";
import { MarginLabels, type MarginLabelItem } from "@/lib/charting/drawings/MarginLabels";
import { LevelLines, type Seg } from "@/lib/charting/drawings/LevelLines";
import { ChartToolbar } from "@/components/dashboard/ChartToolbar";
import {
  loadStudySettings, saveStudySettings, subscribeStudySettings, type StudySettings,
  loadHiddenMap, saveHiddenMap, subscribeHiddenMap, type HiddenMap, type StudyId,
} from "@/lib/data/chartStudySettings";

// The ticker page's ONLY chart -- built on TradingView's own open-source
// lightweight-charts engine, replacing the old TradingView-embed-above +
// simpler-companion-below layout entirely (the redesign that drove this
// rewrite retires the embed: it can't run Pine, can't take per-study
// colors, and can't expose per-study settings reliably -- this chart is
// our own code, so all three are fully real here). One multi-pane chart
// instance (price + RSI + MACD + Stoch RSI as `chart.addSeries(def, opts,
// paneIndex)` panes, not four separate synced chart instances the way the
// design prototype's own sandbox was forced to) -- this repo already has a
// proven working pane-resize/multi-pane setup, which is the approach the
// redesign's own handoff doc says to prefer when available.
//
// Per-indicator visibility (`hidden`, an eye toggle on each legend row) and
// settings (`settings`, a gear opening a 3-tab Inputs/Style/Visibility
// panel) both live in lib/data/chartStudySettings.ts -- global per-viewer
// preferences, not per-ticker facts, same as the old chartStudies.ts
// picker this component no longer reads from (it's fully self-contained
// now; that picker still serves the Dashboard/Watchlist pages' own
// TradingView embeds unchanged).

const MONO = "var(--font-mono)";
const CARD: CSSProperties = { background: "var(--panel)", border: "1px solid var(--border)", borderRadius: "var(--r, 16px)", padding: "12px 14px 14px", marginBottom: 14 };
const OSC_H = 130;
const OSC_H_HIDDEN = 26;
const PRICE_H = 420;
// `pointer-events: none` on the row itself (only the eye/gear buttons opt
// back in with their own `auto`) -- a legend row's text/background must
// never intercept a mouse-drag meant to pan the chart underneath it. An
// earlier version set `auto` on the whole row, which silently ate any
// drag-to-pan gesture starting anywhere near a legend pill, not just on
// its buttons -- caught via a direct report that panning didn't work.
const LEGEND_ROW: CSSProperties = { display: "flex", alignItems: "center", gap: 5, height: 18, padding: "0 5px", borderRadius: 5, pointerEvents: "none" };
const EYE_BTN: CSSProperties = { width: 16, height: 16, flex: "none", display: "flex", alignItems: "center", justifyContent: "center", border: "none", borderRadius: 3, cursor: "pointer", background: "transparent", pointerEvents: "auto" };

const RANGE_BUTTONS = [
  { id: "1M", n: 22 }, { id: "3M", n: 66 }, { id: "6M", n: 130 }, { id: "1Y", n: 252 }, { id: "All", n: null },
] as const;

// Right-margin space reserved for the VWAP Suite's own price/%% readouts
// (MarginLabels.ts, drawn past the last bar) -- kept in PIXELS, not a fixed
// bar count: `rightOffset` (lightweight-charts' own margin unit) is bars,
// so at a fixed bar count the pixel gap balloons at high zoom and vanishes
// at low zoom. Deliberately modest -- "close to the last candle", not a
// wide TradingView-style reading column.
const MARGIN_PX = 56;

/** Recomputes `rightOffset` (bars) so the VWAP margin stays ~MARGIN_PX wide
    on screen through pans/zooms/range changes. Solved in closed form
    rather than iterating `rightOffset = MARGIN_PX / currentBarSpacing`: that
    naive version reads back a barSpacing the offset itself already
    shrank (more phantom bars packed into the same pixel width lowers
    every bar's width, real candles included), so each correction demands
    a bigger correction next time -- a genuine unbounded feedback loop,
    not just jitter, caught from a direct report that the margin kept
    growing and candles kept visibly compressing. Deriving barSpacing =
    plotWidth / (visibleRealBars + rightOffset) and solving
    rightOffset*barSpacing = MARGIN_PX for rightOffset directly (algebra
    below) has no such circularity: it depends only on the REAL bar count
    in view, never on a barSpacing the previous correction already
    distorted.
      barSpacing = plotWidth / (visibleRealBars + rightOffset)
      rightOffset * barSpacing = MARGIN_PX
      => rightOffset = MARGIN_PX * visibleRealBars / (plotWidth - MARGIN_PX) */
function makeMarginOffsetTracker(chart: IChartApi, lastBarIndex: number, marginPx: number) {
  const adjustingRef = { current: false };
  const apply = () => {
    if (adjustingRef.current) return;
    const plotWidth = chart.timeScale().width();
    const lr = chart.timeScale().getVisibleLogicalRange();
    if (!plotWidth || plotWidth <= marginPx || !lr) return;
    const realFrom = Math.max(0, lr.from);
    const realTo = Math.min(lastBarIndex, lr.to);
    const visibleRealBars = Math.max(1, realTo - realFrom);
    const wanted = Math.max(2, Math.round((marginPx * visibleRealBars) / (plotWidth - marginPx)));
    const current = chart.timeScale().options().rightOffset;
    if (Math.abs(current - wanted) < 1) return;
    adjustingRef.current = true;
    chart.timeScale().applyOptions({ rightOffset: wanted });
    // Cleared on a timeout rather than immediately after the call above --
    // protects against the resulting range-change event firing on a
    // microtask/next tick rather than synchronously within applyOptions.
    setTimeout(() => { adjustingRef.current = false; }, 0);
  };
  return apply;
}

/** Pixel width the VWAP margin needs to fit its own text labels without
    clipping -- MARGIN_PX alone (tuned tight, "close to the last candle")
    is only enough when there's no label text to show at all. Mirrors
    MarginLabelsRenderer's own two-column layout math (col1 right-aligned
    at the edge, col0 inboard of it with a 14px gap) so the reserved margin
    and the text actually drawn into it always agree. Measured with a
    throwaway canvas context (cheap -- a handful of measureText calls once
    per chart rebuild), never below MARGIN_PX so VWAP-off/label-off stays
    at the original tight default. */
function measureMarginWidth(items: MarginLabelItem[], font: string): number {
  if (!items.length || typeof document === "undefined") return MARGIN_PX;
  const ctx = document.createElement("canvas").getContext("2d");
  if (!ctx) return MARGIN_PX;
  ctx.font = font;
  const widest = (col: 0 | 1) => items.filter((it) => it.col === col).reduce((w, it) => Math.max(w, ctx.measureText(it.text).width), 0);
  const w0 = widest(0), w1 = widest(1);
  const needed = (w1 ? w0 + 14 + w1 : w0) + 8 + 16;
  return Math.max(MARGIN_PX, Math.ceil(needed));
}

// IDX's own published tick-size schedule (price bands -> minimum price
// increment) -- a real exchange rule, not a guess, used for the candle
// series' own price formatter/minMove so the axis rounds the way the
// exchange itself quotes this ticker, rather than hardcoding one band's
// tick (25, right for AADI's own price level) for every ticker.
function idxTickSize(price: number): number {
  if (price < 200) return 1;
  if (price < 500) return 2;
  if (price < 2000) return 5;
  if (price < 5000) return 10;
  return 25;
}

function themeColors() {
  const cs = typeof document !== "undefined" ? getComputedStyle(document.documentElement) : null;
  const v = (name: string, fallback: string) => (cs?.getPropertyValue(name).trim() || fallback);
  const dark = typeof document !== "undefined" && document.documentElement.dataset.theme === "dark";
  return {
    dark,
    panel: v("--panel", dark ? "#0c0f14" : "#ffffff"),
    border: v("--border", dark ? "#1e222d" : "#e7e9ee"),
    hair: v("--hair", dark ? "#191e27" : "#eff1f4"),
    soft: v("--soft", dark ? "#161b24" : "#f4f6f9"),
    text: v("--text", dark ? "#d9dde5" : "#0b0e14"),
    muted: v("--muted", dark ? "#a3a9b5" : "#5b6472"),
    faint: v("--faint", dark ? "#6b7280" : "#9aa1ad"),
    accent: v("--accent", dark ? "#5b8cff" : "#2563eb"),
    shadow: v("--shadow", dark ? "rgba(0,0,0,.55)" : "rgba(8,10,13,.20)"),
    // Redesign-specific tokens with no CSS-var home yet -- chart-canvas-only
    // values, kept local per the file's existing themeColors() precedent.
    // A light scrim for legend-text legibility, not a wall -- at the old
    // .82/.88 alpha this fully occluded candles/wicks sitting behind it
    // (confirmed live: a tall wick visibly cut off right at the block's
    // bottom edge), which is exactly the opposite of how TradingView's own
    // legend renders (text directly over the chart, no solid backing).
    paneTag: dark ? "rgba(12,15,20,.32)" : "rgba(255,255,255,.42)",
    oscFill: dark ? "rgba(28,45,98,.32)" : "rgba(41,98,255,.07)",
    // Candle up/down deliberately split from the UI accent in dark mode
    // (TradingView's own dark candle blue is more saturated than this
    // site's dark UI accent) -- down candles are hollow (white body,
    // colored border), matching TradingView's own default look.
    candleUp: dark ? "#2962FF" : "#2563eb",
    candleDown: "#ffffff",
    candleBorderDown: dark ? "#ffffff" : "#5b6472",
  };
}

function chip(label: string, color: string): CSSProperties {
  return { fontSize: 8.5, fontWeight: 800, letterSpacing: ".06em", color, background: "var(--soft)", border: "1px solid var(--border)", borderRadius: 5, padding: "2px 6px" };
}

function nz(v: number | undefined): number | null { return v == null || Number.isNaN(v) ? null : v; }
function lw(n: number): LineWidth { return Math.min(4, Math.max(1, Math.round(n))) as LineWidth; }
function fmtOsc(v: number): string { return Math.min(100, Math.max(0, v)).toFixed(1); }

// ---- Eye / gear icon buttons on each legend row --------------------------

function EyeButton({ hidden, onClick }: { hidden: boolean; onClick: () => void }) {
  return (
    <button type="button" onClick={onClick} title={hidden ? "Show" : "Hide"} style={{ ...EYE_BTN, color: "var(--faint)" }}>
      <svg style={{ flex: "none" }} viewBox="0 0 24 24" width={13} height={13} fill="none" stroke="currentColor" strokeWidth={1.7}>
        <path d="M2 12s3.6-6 10-6 10 6 10 6-3.6 6-10 6-10-6-10-6Z" />
        <circle cx="12" cy="12" r="2.5" />
        {hidden ? <path d="M4 20 20 4" /> : null}
      </svg>
    </button>
  );
}

function GearButton({ onClick }: { onClick: (e: React.MouseEvent) => void }) {
  return (
    <button type="button" onClick={onClick} title="Settings" style={{ ...EYE_BTN, color: "var(--faint)" }}>
      <svg style={{ flex: "none" }} viewBox="0 0 24 24" width={13} height={13} fill="none" stroke="currentColor" strokeWidth={1.7}>
        <circle cx="12" cy="12" r="3" />
        <path d="M12 3v2M12 19v2M3 12h2M19 12h2M5.6 5.6l1.4 1.4M17 17l1.4 1.4M18.4 5.6 17 7M7 17l-1.4 1.4" />
      </svg>
    </button>
  );
}

/** RSI legend row's own "Bull"/"Bear" divergence toggle -- independent of
    the row's Hide/Settings pair, same idea as VWAP's per-period on/off but
    surfaced as a direct pill (not buried in the gear panel) since the user
    asked for a one-click show/hide for each divergence direction. */
function DivToggleButton({ label, active, color, onClick }: { label: string; active: boolean; color: string; onClick: () => void }) {
  return (
    <button type="button" onClick={onClick} title={`${active ? "Hide" : "Show"} ${label} divergences`}
      style={{ flex: "none", pointerEvents: "auto", fontSize: 9.5, fontWeight: 700, padding: "1px 6px", borderRadius: 5, border: `1px solid ${active ? color : "var(--border)"}`, cursor: "pointer", background: active ? hexA(color, 0.14) : "transparent", color: active ? color : "var(--faint)" }}>
      {label}
    </button>
  );
}

// ---- 3-tab Settings panel (Inputs / Style / Visibility) -------------------

type NumField = { kind: "number"; key: keyof StudySettings; label: string; min: number; max: number };
type ColorField = { kind: "color"; key: keyof StudySettings; label: string };
type SelectField = { kind: "select"; key: keyof StudySettings; label: string; options: string[] };
type CheckField = { kind: "check"; key: keyof StudySettings; label: string };
type GearField = NumField | ColorField | SelectField | CheckField;

const SWATCHES = ["#2962FF", "#FF5050", "#D6A100", "#16A34A", "#FF9800", "#787b86"];

function GearPanel({ title, top, inputs, style, id, values, hidden, onChange, onToggleHidden, onClose, onReset }: {
  title: string; top: number; id: StudyId;
  inputs: GearField[]; style: GearField[];
  values: StudySettings; hidden: boolean;
  onChange: (patch: Partial<StudySettings>) => void;
  onToggleHidden: () => void;
  onClose: () => void; onReset: () => void;
}) {
  const [tab, setTab] = useState<"Inputs" | "Style" | "Visibility">("Inputs");
  const fields = tab === "Inputs" ? inputs : tab === "Style" ? style : [];

  return (
    <div style={{ position: "absolute", left: 20, top, zIndex: 30, width: 300, maxHeight: 480, display: "flex", flexDirection: "column", borderRadius: 12, overflow: "hidden", border: "1px solid var(--border)", background: "var(--panel)", boxShadow: "0 22px 52px var(--shadow)", pointerEvents: "auto" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "11px 13px 7px" }}>
        <span style={{ flex: "1 1 auto", minWidth: 0, fontSize: 13, fontWeight: 700 }}>{title}</span>
        <button type="button" onClick={onClose} style={{ width: 20, height: 20, flex: "none", display: "flex", alignItems: "center", justifyContent: "center", border: "none", borderRadius: 5, cursor: "pointer", background: "transparent", color: "var(--faint)" }}>
          <svg style={{ flex: "none" }} viewBox="0 0 24 24" width={13} height={13} fill="none" stroke="currentColor" strokeWidth={1.8}><path d="M6 6l12 12M18 6 6 18" /></svg>
        </button>
      </div>
      <div style={{ display: "flex", gap: 14, padding: "0 13px", borderBottom: "1px solid var(--hair)" }}>
        {(["Inputs", "Style", "Visibility"] as const).map((t) => (
          <button key={t} type="button" onClick={() => setTab(t)} style={{ padding: "5px 0 7px", fontSize: 11.5, fontWeight: 600, border: "none", borderBottom: `2px solid ${tab === t ? "var(--accent)" : "transparent"}`, cursor: "pointer", background: "transparent", color: tab === t ? "var(--text)" : "var(--muted)" }}>{t}</button>
        ))}
      </div>
      <div style={{ flex: "1 1 auto", minHeight: 0, overflow: "auto", display: "flex", flexDirection: "column", gap: 4, padding: "8px 13px 11px" }}>
        {tab === "Visibility" ? (
          <label style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, minHeight: 28, fontSize: 11.5, color: "var(--muted)" }}>
            <span>Visible on chart</span>
            <input type="checkbox" checked={!hidden} onChange={onToggleHidden} style={{ width: 15, height: 15 }} />
          </label>
        ) : fields.map((f) => (
          <div key={String(f.key)} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, minHeight: 28 }}>
            <span style={{ flex: "0 1 auto", minWidth: 0, fontSize: 11.5, color: "var(--muted)" }}>{f.label}</span>
            {f.kind === "number" ? (
              <input type="number" min={f.min} max={f.max} value={values[f.key] as number}
                onChange={(e) => {
                  const n = Math.round(Number(e.target.value));
                  if (Number.isNaN(n)) return;
                  onChange({ [f.key]: Math.min(f.max, Math.max(f.min, n)) } as Partial<StudySettings>);
                }}
                style={{ width: 72, flex: "none", fontFamily: MONO, fontSize: 11.5, textAlign: "right", background: "var(--soft)", color: "var(--text)", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 7px" }} />
            ) : f.kind === "color" ? (
              <div style={{ display: "flex", gap: 3, flex: "none" }}>
                {SWATCHES.map((c) => (
                  <button key={c} type="button" onClick={() => onChange({ [f.key]: c } as Partial<StudySettings>)}
                    style={{ width: 17, height: 17, borderRadius: 5, cursor: "pointer", background: c, border: `2px solid ${values[f.key] === c ? "var(--text)" : "transparent"}` }} />
                ))}
              </div>
            ) : f.kind === "select" ? (
              <div style={{ display: "flex", gap: 2, padding: 2, flex: "none", borderRadius: 7, background: "var(--soft)" }}>
                {f.options.map((opt) => (
                  <button key={opt} type="button" onClick={() => onChange({ [f.key]: opt } as Partial<StudySettings>)}
                    style={{ padding: "4px 8px", fontSize: 10, fontWeight: 700, border: "none", borderRadius: 5, cursor: "pointer", background: values[f.key] === opt ? "var(--accent)" : "transparent", color: values[f.key] === opt ? "#fff" : "var(--muted)" }}>{opt}</button>
                ))}
              </div>
            ) : (
              <button type="button" onClick={() => onChange({ [f.key]: !values[f.key] } as Partial<StudySettings>)}
                style={{ width: 17, height: 17, flex: "none", display: "flex", alignItems: "center", justifyContent: "center", border: "1px solid var(--border)", borderRadius: 4, cursor: "pointer", background: values[f.key] ? "var(--accent)" : "transparent" }}>
                {values[f.key] ? <svg viewBox="0 0 24 24" width={12} height={12} fill="none" stroke="#fff" strokeWidth={3} strokeLinecap="round" strokeLinejoin="round"><path d="M5 13l4 4L19 7" /></svg> : null}
              </button>
            )}
          </div>
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 13px", borderTop: "1px solid var(--hair)" }}>
        <button type="button" onClick={onReset} style={{ height: 27, padding: "0 11px", fontSize: 11, fontWeight: 600, border: "1px solid var(--border)", borderRadius: 7, cursor: "pointer", background: "transparent", color: "var(--muted)" }}>Defaults</button>
        <div style={{ flex: "1 1 8px" }} />
        <button type="button" onClick={onClose} style={{ height: 27, padding: "0 13px", fontSize: 11, fontWeight: 700, border: "none", borderRadius: 7, cursor: "pointer", background: "var(--accent)", color: "#fff" }}>Ok</button>
      </div>
    </div>
  );
}

// ---- Drawing-selection restyle toolbar ------------------------------------

function SelectionToolbar({ sel, onRestyle, onDelete, onClose }: {
  sel: NonNullable<SelectionInfo>;
  onRestyle: (color: string, width: number, style: "solid" | "dashed" | "dotted") => void;
  onDelete: () => void; onClose: () => void;
}) {
  return (
    <div style={{ position: "absolute", left: Math.max(8, sel.x + 12), top: Math.max(8, sel.y + 12), zIndex: 36, display: "flex", alignItems: "center", gap: 8, padding: "6px 8px", borderRadius: 9, border: "1px solid var(--border)", background: "var(--panel)", boxShadow: "0 14px 34px var(--shadow)" }}>
      <span style={{ fontSize: 11, fontWeight: 600, whiteSpace: "nowrap" }}>{sel.title}</span>
      <span style={{ width: 1, height: 16, background: "var(--hair)" }} />
      <div style={{ display: "flex", gap: 3 }}>
        {SWATCHES.map((c) => (
          <button key={c} type="button" onClick={() => onRestyle(c, sel.width, sel.style)} style={{ width: 16, height: 16, borderRadius: 5, cursor: "pointer", background: c, border: `2px solid ${sel.color === c ? "var(--text)" : "transparent"}` }} />
        ))}
      </div>
      <div style={{ display: "flex", gap: 2 }}>
        {[1, 2, 3].map((w) => (
          <button key={w} type="button" onClick={() => onRestyle(sel.color, w, sel.style)} style={{ minWidth: 20, height: 20, fontSize: 10, fontWeight: 700, border: "none", borderRadius: 5, cursor: "pointer", background: sel.width === w ? "var(--accent)" : "var(--soft)", color: sel.width === w ? "#fff" : "var(--muted)" }}>{w}</button>
        ))}
      </div>
      <div style={{ display: "flex", gap: 2 }}>
        {(["solid", "dashed", "dotted"] as const).map((s) => (
          <button key={s} type="button" onClick={() => onRestyle(sel.color, sel.width, s)} style={{ height: 20, padding: "0 7px", fontSize: 10, fontWeight: 600, border: "none", borderRadius: 5, cursor: "pointer", background: sel.style === s ? "var(--accent)" : "var(--soft)", color: sel.style === s ? "#fff" : "var(--muted)" }}>{s}</button>
        ))}
      </div>
      <span style={{ width: 1, height: 16, background: "var(--hair)" }} />
      <button type="button" onClick={onDelete} title="Delete drawing" style={{ width: 22, height: 22, flex: "none", display: "flex", alignItems: "center", justifyContent: "center", border: "none", borderRadius: 5, cursor: "pointer", background: "transparent", color: "var(--faint)" }}>
        <svg style={{ flex: "none" }} viewBox="0 0 24 24" width={14} height={14} fill="none" stroke="currentColor" strokeWidth={1.7} strokeLinejoin="round"><path d="M4 7h16M9 7V5h6v2m-9 0 1 13h8l1-13" /></svg>
      </button>
      <button type="button" onClick={onClose} style={{ width: 22, height: 22, flex: "none", display: "flex", alignItems: "center", justifyContent: "center", border: "none", borderRadius: 5, cursor: "pointer", background: "transparent", color: "var(--faint)" }}>
        <svg style={{ flex: "none" }} viewBox="0 0 24 24" width={13} height={13} fill="none" stroke="currentColor" strokeWidth={1.8}><path d="M6 6l12 12M18 6 6 18" /></svg>
      </button>
    </div>
  );
}

type Legend = {
  date: string; o: number; h: number; l: number; c: number; vol: number;
  chg: number | null; chgPct: number | null;
  ib: { ibHigh: number; ibLow: number } | null; vwap: AvwapPoint | null;
  ribbon1: number | null; ribbon2: number | null; sma200: number | null;
  rsi: number | null; rsiMa: number | null; rsiMaBull: boolean | null;
  macd: number | null; macdSignal: number | null; macdHist: number | null;
  stochK: number | null; stochD: number | null;
};

export function IndicatorCompanion({ ohlcv, symbol, companyName }: { ohlcv: OhlcvPayload | null; symbol?: string; companyName?: string }) {
  const [settings, setSettings] = useState<StudySettings>(() => loadStudySettings());
  const [hidden, setHidden] = useState<HiddenMap>(() => loadHiddenMap());
  const [legend, setLegend] = useState<Legend | null>(null);
  const [themeTick, setThemeTick] = useState(0);
  const [drawingTool, setDrawingTool] = useState<DrawingTool | null>(null);
  const [selection, setSelection] = useState<SelectionInfo>(null);
  const [openGear, setOpenGear] = useState<StudyId | null>(null);
  const [activeRangeId, setActiveRangeId] = useState<string>("3M");
  const containerRef = useRef<HTMLDivElement | null>(null);
  const chartRef = useRef<IChartApi | null>(null);
  // The chart-building effect below tears down and fully recreates the
  // chart on every settings/hidden change (toggling any indicator's eye,
  // Bull/Bear, or a gear-panel edit), since indicator series can't just be
  // added/removed from a live chart cleanly here. Restoring the view from
  // viewRangeStore's SAVED range alone isn't enough to make that invisible:
  // that save is 400ms-debounced, so a toggle clicked right after a pan/
  // zoom (well within 400ms, and toggling is a quick click) would read a
  // stale range and visibly "jump" -- exactly the "chart zooms in when I
  // hide/unhide an indicator" bug this fixes. Capturing the chart's own
  // live range synchronously in the cleanup below, in a ref (so it
  // survives the teardown/rebuild instead of resetting with component
  // state), and preferring it over the debounced localStorage value closes
  // that race entirely -- zoom now only ever changes when the user
  // actually changes it.
  const lastRangeRef = useRef<{ from: string; to: string } | null>(null);
  const lastSymbolRef = useRef<string | undefined>(undefined);
  if (lastSymbolRef.current !== symbol) { lastRangeRef.current = null; lastSymbolRef.current = symbol; }

  useEffect(() => subscribeStudySettings(setSettings), []);
  useEffect(() => subscribeHiddenMap(setHidden), []);
  useEffect(() => {
    const obs = new MutationObserver(() => setThemeTick((t) => t + 1));
    if (typeof document !== "undefined") obs.observe(document.documentElement, { attributes: true, attributeFilter: ["data-theme"] });
    return () => obs.disconnect();
  }, []);
  useEffect(() => {
    if (!drawingTool) return;
    return drawingTool.onSelectionChange(setSelection);
  }, [drawingTool]);

  function updateSettings(patch: Partial<StudySettings>) {
    setSettings((prev) => { const next = { ...prev, ...patch }; saveStudySettings(next); return next; });
  }
  function toggleHidden(id: StudyId | "rsiDivBull" | "rsiDivBear") {
    setHidden((prev) => { const next = { ...prev, [id]: !prev[id] }; saveHiddenMap(next); return next; });
  }

  const rows = ohlcv?.rows;

  function applyRange(id: string) {
    setActiveRangeId(id);
    if (!rows || !chartRef.current) return;
    const btn = RANGE_BUTTONS.find((r) => r.id === id);
    if (!btn || btn.n == null) { chartRef.current.timeScale().fitContent(); return; }
    const fromIdx = Math.max(0, rows.length - btn.n);
    chartRef.current.timeScale().setVisibleRange({ from: rows[fromIdx].date as Time, to: rows[rows.length - 1].date as Time });
  }

  const oscCount = (hidden.rsi ? 0 : 1) + (hidden.macd ? 0 : 1) + (hidden.stoch ? 0 : 1);
  const oscStripCount = (hidden.rsi ? 1 : 0) + (hidden.macd ? 1 : 0) + (hidden.stoch ? 1 : 0);
  const chartHeight = PRICE_H + OSC_H * oscCount;
  const totalHeight = chartHeight + OSC_H_HIDDEN * oscStripCount;
  // Corner-label offsets for whichever oscillator panes are actually live --
  // an oscillator that's hidden contributes no gap to the ones after it.
  const rsiTop = PRICE_H;
  const macdTop = PRICE_H + (hidden.rsi ? 0 : OSC_H);
  const stochTop = PRICE_H + (hidden.rsi ? 0 : OSC_H) + (hidden.macd ? 0 : OSC_H);

  useEffect(() => {
    const el = containerRef.current;
    if (!el || !rows || rows.length < 2) { setLegend(null); return; }

    const c = themeColors();
    const s = settings;
    const close = rows[rows.length - 1].close;
    const chart = createChart(el, {
      autoSize: true,
      layout: { background: { color: c.panel }, textColor: c.muted, fontFamily: MONO, panes: { separatorColor: c.hair, separatorHoverColor: c.border, enableResize: true } },
      grid: { vertLines: { visible: false }, horzLines: { visible: false } },
      rightPriceScale: { borderColor: c.border, minimumWidth: 74 },
      timeScale: { borderColor: c.border, rightOffset: 3 },
      crosshair: {
        mode: CrosshairMode.Normal,
        vertLine: { color: c.faint, width: 1, style: LineStyle.Dashed, labelBackgroundColor: c.muted },
        horzLine: { color: c.faint, width: 1, style: LineStyle.Dashed, labelBackgroundColor: c.muted },
      },
    });
    chartRef.current = chart;

    const tick = idxTickSize(close);
    const candles = chart.addSeries(CandlestickSeries, {
      upColor: c.candleUp, downColor: c.candleDown, borderVisible: true, borderUpColor: c.candleUp, borderDownColor: c.candleBorderDown,
      wickUpColor: c.candleUp, wickDownColor: c.candleBorderDown,
      priceFormat: { type: "custom", formatter: (p: number) => formatPrice(p), minMove: tick },
    });
    candles.setData(rows.map((r) => ({ time: r.date as Time, open: r.open, high: r.high, low: r.low, close: r.close })));

    const drawing = new DrawingTool(chart, candles, symbol || "", rows);
    setDrawingTool(drawing);

    const closes = rows.map((r) => r.close);
    const highs = rows.map((r) => r.high);
    const lows = rows.map((r) => r.low);
    const rsiArr = rsiWilder(closes, s.rsiLen);
    const rsiMaArr = ema(rsiArr, s.rsiMaLen);
    // Momentum state behind the EMA-of-RSI line's bull/bear coloring --
    // computed here (not inside the RSI pane's own `if` block below) so the
    // hover legend can also read it, even for the bar under the crosshair.
    const dmiResult = dmi(highs, lows, closes, s.rsiDmiLen, s.rsiAdxSmoothing);
    const emaBull = rsiMaArr.map((v, i) => {
      if (i === 0 || Number.isNaN(v) || Number.isNaN(rsiMaArr[i - 1])) return true;
      const rising = v > rsiMaArr[i - 1];
      if (!s.rsiUseDmiFilter) return rising;
      return rising && dmiResult.diPlus[i] >= dmiResult.diMinus[i];
    });
    const macdResult = macdOf(closes, s.macdFast, s.macdSlow, s.macdSignal, s.macdSmooth);
    const stochBase = rsiWilder(closes, s.stochRsiLen);
    const stoch = stochOf(stochBase, s.stochLen, s.stochSmoothK, s.stochSmoothD);
    const lineData = (vals: number[]) => rows.map((r, i) => ({ time: r.date as Time, value: vals[i] })).filter((p) => !Number.isNaN(p.value));

    // ── Volume -- overlaid on the price pane via a hidden scale, up/down
    //    tinted at low alpha per the redesign, + an MA line. ──
    if (!hidden.volume) {
      const volume = chart.addSeries(HistogramSeries, { priceScaleId: "vol", lastValueVisible: false, priceLineVisible: false });
      volume.priceScale().applyOptions({ scaleMargins: { top: 0.86, bottom: 0 }, visible: false });
      volume.setData(rows.map((r) => ({ time: r.date as Time, value: r.volume || 0, color: r.close >= r.open ? hexA(s.volUpColor, 0.38) : hexA(s.volDnColor, 0.38) })));
      const volMa = chart.addSeries(LineSeries, { color: s.volMaColor, lineWidth: lw(s.volMaWidth), priceScaleId: "vol", crosshairMarkerVisible: false, lastValueVisible: false, priceLineVisible: false });
      volMa.setData(lineData(sma(rows.map((r) => r.volume || 0), s.volMaLen)));
    }

    // ── Monthly Initial Balance -- one stroked+filled box per month. ──
    const bands = computeInitialBalance(rows, s.ibDays);
    if (!hidden.ib) {
      const box = new IbBoxes(bands, rows, s.ibColor, s.ibFill ? hexA(s.ibColor, 0.05) : "rgba(0,0,0,0)", s.ibWidth);
      candles.attachPrimitive(box);
    }

    // ── EMA Ribbon -- exactly two EMAs (25/50 by default), matching
    //    TradingView's own simplest "EMA Cross" convention rather than the
    //    original 8-line fanned-ribbon design. `autoscaleInfoProvider: null`
    //    on every overlay below (Ribbon, SMA, VWAP) -- same reasoning as the
    //    VWAP σ-band exclusion: a long-period MA can sit far from the
    //    recent candle range (a 200-day average lags hard after a big move),
    //    and letting that drive the visible price range squashes the
    //    candles vertically the moment that overlay is switched on. The
    //    candlesticks alone own the price scale now; toggling any overlay
    //    on/off no longer changes how big the candles themselves look --
    //    caught from a direct report that SMA/VWAP visibly compressed them. ──
    if (!hidden.ribbon) {
      const line1 = chart.addSeries(LineSeries, { color: s.ribbonColor, lineWidth: lw(s.ribbonWidth), crosshairMarkerVisible: false, lastValueVisible: true, priceLineVisible: false, title: `EMA ${s.ribbon1Len}`, autoscaleInfoProvider: () => null });
      line1.setData(lineData(ema(closes, s.ribbon1Len)));
      const line2 = chart.addSeries(LineSeries, { color: hexA(s.ribbonColor, 0.55), lineWidth: lw(s.ribbonWidth), crosshairMarkerVisible: false, lastValueVisible: true, priceLineVisible: false, title: `EMA ${s.ribbon2Len}`, autoscaleInfoProvider: () => null });
      line2.setData(lineData(ema(closes, s.ribbon2Len)));
    }

    // ── SMA -- single line, hidden by default. ──
    if (!hidden.sma200) {
      const smaLine = chart.addSeries(LineSeries, { color: s.smaColor, lineWidth: lw(s.smaWidth), crosshairMarkerVisible: false, lastValueVisible: true, priceLineVisible: false, title: `SMA ${s.sma200Len}`, autoscaleInfoProvider: () => null });
      smaLine.setData(lineData(sma(closes, s.sma200Len)));
    }

    // ── VWAP Suite v2 -- one center line per anchor period (breaking
    //    cleanly at each reset); the live period gets ±1σ/±2σ envelope
    //    fills + full margin labels, every CLOSED period forwards its own
    //    final VWAP as a dotted line to the last bar (capped at 60 periods,
    //    matching the old MAX_VWAP_SEGMENTS cap, so a fine anchor over a
    //    long history doesn't create hundreds of series), and only the ~4
    //    most recent closed periods get margin-label text (labelling all of
    //    them would be unreadable noise; forward lines still draw for
    //    every one within the cap). ──
    const vw = computeAnchoredVwap(rows, s.vwapAnchor as VwapAnchor, s.vwapMult1, s.vwapMult2, s.vwapMult2 + 1, s.vwapSource as VwapSource);
    let dynamicMarginPx = MARGIN_PX;
    if (!hidden.vwap) {
      const segs: Array<{ key: string; idxs: number[] }> = [];
      vw.points.forEach((p, i) => {
        if (!p) return;
        const last = segs[segs.length - 1];
        if (!last || last.key !== p.key) segs.push({ key: p.key, idxs: [i] });
        else last.idxs.push(i);
      });
      const visSegs = segs.slice(-60);
      const marginItems: MarginLabelItem[] = [];
      visSegs.forEach((seg, si) => {
        const isCurrent = si === visSegs.length - 1;
        const lastPt = vw.points[seg.idxs[seg.idxs.length - 1]] as AvwapPoint;
        const pLabel = periodLabel(seg.key, s.vwapAnchor as VwapAnchor);
        if (isCurrent) {
          const center = chart.addSeries(LineSeries, { color: s.vwapColor, lineWidth: lw(s.vwapWidth), crosshairMarkerVisible: true, lastValueVisible: true, priceLineVisible: false, title: `VWAP Suite v2.0 ${pLabel}`, autoscaleInfoProvider: () => null });
          center.setData(seg.idxs.map((i) => ({ time: rows[i].date as Time, value: (vw.points[i] as AvwapPoint).vwap })));
          if (s.vwapText) {
            marginItems.push({ text: `${pLabel}VWAP • ${formatPrice(lastPt.vwap)} • ${pct(lastPt.vwap, close)}`, price: lastPt.vwap, col: 0 });
          }
          if (s.vwapBands) {
            // Excluded from autoscale (null autoscaleInfoProvider): a fresh
            // period's first bar or two has almost no volume sampled yet, so
            // its volume-weighted variance -- and therefore ±2σ -- can spike
            // absurdly wide before settling; letting that drive the visible
            // price range would squash the actual candles into a sliver.
            // The bands still render at whatever value they carry, same as
            // TradingView's own anchored VWAP does through that same
            // early-period wobble.
            const bandLine = (vals: (p: AvwapPoint) => number, label: string) => {
              const bs = chart.addSeries(LineSeries, { color: s.vwapColor, lineWidth: 1, lineStyle: LineStyle.Dashed, crosshairMarkerVisible: false, lastValueVisible: false, priceLineVisible: false, autoscaleInfoProvider: () => null });
              bs.setData(seg.idxs.map((i) => ({ time: rows[i].date as Time, value: vals(vw.points[i] as AvwapPoint) })));
              if (s.vwapText) marginItems.push({ text: `${label} • ${formatPrice(vals(lastPt))} • ${pct(vals(lastPt), close)}`, price: vals(lastPt), col: 0 });
            };
            bandLine((p) => p.u1, "+1σ"); bandLine((p) => p.l1, "−1σ");
            bandLine((p) => p.u2, "+2σ"); bandLine((p) => p.l2, "−2σ");
            const fillTop = chart.addSeries(LineSeries, { color: "rgba(0,0,0,0)", lastValueVisible: false, priceLineVisible: false, crosshairMarkerVisible: false, autoscaleInfoProvider: () => null });
            fillTop.setData(seg.idxs.map((i) => ({ time: rows[i].date as Time, value: (vw.points[i] as AvwapPoint).u2 })));
            fillTop.attachPrimitive(new VwapFill(seg.idxs.map((i) => rows[i].date as Time), seg.idxs.map((i) => (vw.points[i] as AvwapPoint).u2), seg.idxs.map((i) => (vw.points[i] as AvwapPoint).u1), hexA(s.vwapColor, 0.07)));
            const fillBot = chart.addSeries(LineSeries, { color: "rgba(0,0,0,0)", lastValueVisible: false, priceLineVisible: false, crosshairMarkerVisible: false, autoscaleInfoProvider: () => null });
            fillBot.setData(seg.idxs.map((i) => ({ time: rows[i].date as Time, value: (vw.points[i] as AvwapPoint).l1 })));
            fillBot.attachPrimitive(new VwapFill(seg.idxs.map((i) => rows[i].date as Time), seg.idxs.map((i) => (vw.points[i] as AvwapPoint).l1), seg.idxs.map((i) => (vw.points[i] as AvwapPoint).l2), hexA(s.vwapColor, 0.1)));
          }
        } else if (s.vwapLines) {
          // Excluded from autoscale too -- an old closed period's VWAP can
          // sit far outside the recent candle range (this ticker may simply
          // have traded much higher or lower back then), and forwarding
          // that as a flat reference line is meant to show that distance,
          // not to also drag the whole price scale out to reach it.
          const forward = chart.addSeries(LineSeries, { color: c.faint, lineWidth: 1, lineStyle: LineStyle.Dotted, crosshairMarkerVisible: false, lastValueVisible: false, priceLineVisible: false, autoscaleInfoProvider: () => null });
          forward.setData([{ time: rows[seg.idxs[0]].date as Time, value: lastPt.vwap }, { time: rows[rows.length - 1].date as Time, value: lastPt.vwap }]);
          if (si === visSegs.length - 2 && s.vwapText) {
            marginItems.push({ text: `P${pLabel}VWAP • ${formatPrice(lastPt.vwap)} • ${pct(lastPt.vwap, close)}`, price: lastPt.vwap, col: 1 });
            marginItems.push({ text: `P +1σ • ${formatPrice(lastPt.u1)} • ${pct(lastPt.u1, close)}`, price: lastPt.u1, col: 1 });
            marginItems.push({ text: `P −1σ • ${formatPrice(lastPt.l1)} • ${pct(lastPt.l1, close)}`, price: lastPt.l1, col: 1 });
          } else if (si >= visSegs.length - 4 && s.vwapText) {
            marginItems.push({ text: `${pLabel} • ${formatPrice(lastPt.vwap)} • ${pct(lastPt.vwap, close)}`, price: lastPt.vwap, col: 1 });
          }
        }
      });
      if (marginItems.length) {
        const font = `10.5px ${MONO}, ui-monospace, monospace`;
        const host = chart.addSeries(LineSeries, { color: "rgba(0,0,0,0)", lastValueVisible: false, priceLineVisible: false, crosshairMarkerVisible: false });
        host.setData(rows.map((r) => ({ time: r.date as Time, value: r.close })));
        host.attachPrimitive(new MarginLabels(marginItems, rows[rows.length - 1].date as Time, c.muted, font));
        dynamicMarginPx = measureMarginWidth(marginItems, font);
      }
    }

    // ── RSI -- length rsiLen (Wilder, default 14) + an EMA-of-RSI companion
    //    line (default length 9), both matching a user-supplied Pine Script
    //    v6 divergence indicator. The companion line's color reflects
    //    current momentum exactly like that script's `emaColor`: bullish
    //    (rising, and -- when the DMI filter is on -- +DI >= -DI) draws in
    //    rsiMaBullColor, everything else in rsiMaBearColor; rendered as one
    //    short LineSeries per contiguous same-color run (same technique as
    //    VWAP Suite's per-period segments below) since lightweight-charts
    //    has no native per-point line color. Divergence lines/labels (all
    //    four Pine types) are computed from the SAME rsiArr the base RSI
    //    line plots, drawn as LevelLines primitives on this pane, and
    //    independently toggle-able via the "Bull"/"Bear" legend buttons. ──
    let nextPane = 1;
    if (!hidden.rsi) {
      const pane = nextPane++;
      const band = new BandFill(s.rsiUpper, s.rsiLower, c.oscFill);
      const rsiLine = chart.addSeries(LineSeries, {
        color: s.rsiColor, lineWidth: lw(s.rsiWidth), crosshairMarkerVisible: true, lastValueVisible: true, priceLineVisible: false, title: `RSI ${s.rsiLen}`,
        autoscaleInfoProvider: () => ({ priceRange: { minValue: 0, maxValue: 100 } }),
      }, pane);
      rsiLine.setData(lineData(rsiArr));
      if (s.rsiFill) rsiLine.attachPrimitive(band);

      let segStart = -1, segBull = true, havePrevSeg = false;
      for (let i = 0; i <= rsiMaArr.length; i++) {
        const atEnd = i === rsiMaArr.length;
        const nan = !atEnd && Number.isNaN(rsiMaArr[i]);
        if (segStart === -1) { if (!nan && !atEnd) { segStart = i; segBull = emaBull[i]; } continue; }
        if (atEnd || nan || emaBull[i] !== segBull) {
          // Overlap the start back by 1 bar so this segment's line visually
          // connects to the previous one with no gap -- only valid when a
          // previous segment actually plotted that bar (never for the very
          // first segment, where one bar back is still the NaN warm-up
          // prefix, not a real value).
          const dataStart = havePrevSeg ? segStart - 1 : segStart;
          const dataEnd = i - 1;
          const isLastSeg = atEnd || nan;
          const seg = chart.addSeries(LineSeries, {
            color: segBull ? s.rsiMaBullColor : s.rsiMaBearColor, lineWidth: lw(s.rsiMaWidth),
            crosshairMarkerVisible: true, lastValueVisible: isLastSeg, priceLineVisible: false,
            title: isLastSeg ? "EMA of RSI" : "",
          }, pane);
          seg.setData(rows.slice(dataStart, dataEnd + 1).map((r, k) => ({ time: r.date as Time, value: rsiMaArr[dataStart + k] })));
          havePrevSeg = true;
          if (!nan && !atEnd) { segStart = i; segBull = emaBull[i]; } else segStart = -1;
        }
      }

      rsiLine.createPriceLine({ price: s.rsiUpper, color: c.muted, lineWidth: 1, lineStyle: LineStyle.Dashed, axisLabelVisible: true, title: String(s.rsiUpper) });
      rsiLine.createPriceLine({ price: s.rsiLower, color: c.muted, lineWidth: 1, lineStyle: LineStyle.Dashed, axisLabelVisible: true, title: String(s.rsiLower) });

      if (!hidden.rsiDivBull || !hidden.rsiDivBear) {
        const divs = computeDivergences(highs, lows, rsiArr, s.rsiDivLeft, s.rsiDivRight, s.rsiDivStrict);
        const mkSeg = (d: { a: { idx: number; rsi: number }; b: { idx: number; rsi: number } }, label: string): Seg => ({
          a: { time: rows[d.a.idx].date as Time, price: d.a.rsi },
          b: { time: rows[d.b.idx].date as Time, price: d.b.rsi },
          label: s.rsiDivShowLabels ? label : "",
          emphasis: true,
        });
        const byType = (t: DivergenceType) => divs.filter((d) => d.type === t);
        const drawType = (t: DivergenceType, color: string, label: string) => {
          const segs = byType(t).map((d) => mkSeg(d, label));
          if (segs.length) rsiLine.attachPrimitive(new LevelLines(segs, color, s.rsiDivWidth));
        };
        if (!hidden.rsiDivBull) {
          drawType("regularBull", s.rsiDivBullColor, "Bull Div");
          drawType("hiddenBull", s.rsiDivHiddenBullColor, "Hidden Bull");
        }
        if (!hidden.rsiDivBear) {
          drawType("regularBear", s.rsiDivBearColor, "Bear Div");
          drawType("hiddenBear", s.rsiDivHiddenBearColor, "Hidden Bear");
        }
      }
    }

    // ── MACD 4C Smooth -- EMA(fast)-EMA(slow) (optionally re-smoothed),
    //    signal = EMA of that, 4-color histogram (up/down zone x
    //    rising/falling alpha). ──
    if (!hidden.macd) {
      const pane = nextPane++;
      const histUpFull = s.macdHistUp, histUpWeak = hexA(s.macdHistUp, 0.4);
      const histDnFull = s.macdHistDn, histDnWeak = hexA(s.macdHistDn, 0.4);
      const histData = rows.map((r, i) => {
        const v = macdResult.hist[i];
        if (Number.isNaN(v)) return { time: r.date as Time, value: 0, color: "rgba(0,0,0,0)" };
        const prev = macdResult.hist[i - 1];
        const rising = !Number.isNaN(prev) && v > prev;
        const color = v >= 0 ? (rising ? histUpFull : histUpWeak) : (rising ? histDnWeak : histDnFull);
        return { time: r.date as Time, value: v, color };
      });
      const hist = chart.addSeries(HistogramSeries, { lastValueVisible: true, priceLineVisible: false, title: "Histogram" }, pane);
      hist.setData(histData);
      const macdLine = chart.addSeries(LineSeries, { color: s.macdColor, lineWidth: lw(s.macdWidth), crosshairMarkerVisible: true, lastValueVisible: true, priceLineVisible: false, title: "MACD" }, pane);
      macdLine.setData(lineData(macdResult.macd));
      const signalLine = chart.addSeries(LineSeries, { color: s.macdSignalColor, lineWidth: lw(s.macdSignalWidth), crosshairMarkerVisible: true, lastValueVisible: true, priceLineVisible: false, title: "Signal" }, pane);
      signalLine.setData(lineData(macdResult.signal));
    }

    // ── Stoch RSI -- K/D over an RSI(stochRsiLen) base, smoothed by
    //    stochSmoothK/D. Band 80-20, dashed guides at 20/40/60/80. ──
    if (!hidden.stoch) {
      const pane = nextPane++;
      const band = new BandFill(s.stochUpper, s.stochLower, c.oscFill);
      const stochD = chart.addSeries(LineSeries, { color: s.stochDColor, lineWidth: lw(s.stochDWidth), crosshairMarkerVisible: true, lastValueVisible: true, priceLineVisible: false, title: "Stoch RSI D" }, pane);
      stochD.setData(lineData(stoch.d));
      const stochK = chart.addSeries(LineSeries, {
        color: s.stochKColor, lineWidth: lw(s.stochKWidth), crosshairMarkerVisible: true, lastValueVisible: true, priceLineVisible: false, title: "Stoch RSI K",
        autoscaleInfoProvider: () => ({ priceRange: { minValue: 0, maxValue: 100 } }),
      }, pane);
      stochK.setData(lineData(stoch.k));
      if (s.stochFill) stochK.attachPrimitive(band);
      const stochStep = (s.stochUpper - s.stochLower) / 3;
      [s.stochUpper, s.stochUpper - stochStep, s.stochLower + stochStep, s.stochLower].forEach((lvl) => stochK.createPriceLine({ price: lvl, color: c.muted, lineWidth: 1, lineStyle: LineStyle.Dashed, axisLabelVisible: lvl === s.stochUpper || lvl === s.stochLower, title: lvl === s.stochUpper || lvl === s.stochLower ? String(Math.round(lvl)) : "" }));
    }

    // Price pane keeps the majority of the height; each active oscillator
    // pane gets an equal, smaller share -- see the file-level note on why
    // stretch factors (not setHeight) are what survives the chart's own
    // layout pass.
    const panes = chart.panes();
    panes[0]?.setStretchFactor(PRICE_H);
    for (let i = 1; i < panes.length; i++) panes[i]?.setStretchFactor(OSC_H);

    // Restores wherever this ticker's own chart was last zoomed/panned to,
    // falling back to the 3M default only the first time a symbol is ever
    // opened. Prefers the in-memory range this same rebuild's own cleanup
    // just captured (lastRangeRef -- exact, synchronous) over
    // viewRangeStore's localStorage copy (debounced 400ms, so it can be
    // stale by the time an indicator toggle -- a quick click -- triggers
    // this rebuild); localStorage is still the fallback for the real
    // first-mount-after-a-hard-refresh case, where no in-memory value
    // exists yet.
    const savedRange = lastRangeRef.current ?? loadViewRange(symbol || "");
    const dateSet = new Set(rows.map((r) => r.date));
    const initialSessions = RANGE_BUTTONS.find((r) => r.id === "3M")?.n ?? rows.length;
    const fromIdx = Math.max(0, rows.length - initialSessions);
    const defaultRange = { from: rows[fromIdx].date as Time, to: rows[rows.length - 1].date as Time };
    const restored = savedRange && dateSet.has(savedRange.from) && dateSet.has(savedRange.to) ? savedRange : null;
    chart.timeScale().setVisibleRange(restored ? { from: restored.from as Time, to: restored.to as Time } : defaultRange);
    setActiveRangeId(restored ? matchRangeButton(rows, restored) : "3M");

    // Persists the visible range on every pan/zoom/range-button change (not
    // just on unmount -- a tab close or crash shouldn't lose it), and keeps
    // the range-button row honest about whether the current view still
    // matches one of the presets. Debounced: a live drag or scroll-wheel
    // zoom fires this many times a second, and only the settled end state
    // needs to hit localStorage.
    let saveTimer: ReturnType<typeof setTimeout> | null = null;
    const onTimeRangeChange = (r: { from: Time; to: Time } | null) => {
      if (!r) return;
      setActiveRangeId(matchRangeButton(rows, { from: String(r.from), to: String(r.to) }));
      if (saveTimer) clearTimeout(saveTimer);
      saveTimer = setTimeout(() => saveViewRange(symbol || "", { from: String(r.from), to: String(r.to) }), 400);
    };
    chart.timeScale().subscribeVisibleTimeRangeChange(onTimeRangeChange);

    // Always on, regardless of which indicators are visible -- gating this
    // behind `!hidden.vwap` (the margin only exists to make room for VWAP's
    // margin labels) meant toggling VWAP itself changed rightOffset, which
    // changes barSpacing for every bar including real candles: exactly the
    // "candles move when I hide/unhide an indicator" bug this is fixing.
    // Keeping the same fixed margin active at all times, VWAP on or off,
    // means the candles' own size is a function of (plotWidth, bar count,
    // MARGIN_PX) alone -- never of which indicators happen to be showing.
    const trackMarginOffset = makeMarginOffsetTracker(chart, rows.length - 1, dynamicMarginPx);
    trackMarginOffset();
    chart.timeScale().subscribeVisibleLogicalRangeChange(trackMarginOffset);

    const legendAt = (idx: number): Legend => {
      const r = rows[idx];
      const prevClose = idx > 0 ? rows[idx - 1].close : null;
      const chg = prevClose != null ? r.close - prevClose : null;
      const chgPct = prevClose ? chg! / prevClose : null;
      const ib = bands.find((b) => idx >= b.startIdx && idx <= Math.min(b.endIdx, rows.length - 1)) ?? null;
      return {
        date: r.date, o: r.open, h: r.high, l: r.low, c: r.close, vol: r.volume, chg, chgPct,
        ib: ib ? { ibHigh: ib.ibHigh, ibLow: ib.ibLow } : null, vwap: vw.points[idx] ?? null,
        ribbon1: nz(ema(closes, s.ribbon1Len)[idx]), ribbon2: nz(ema(closes, s.ribbon2Len)[idx]), sma200: nz(sma(closes, s.sma200Len)[idx]),
        rsi: nz(rsiArr[idx]), rsiMa: nz(rsiMaArr[idx]), rsiMaBull: Number.isNaN(rsiMaArr[idx]) ? null : emaBull[idx],
        macd: nz(macdResult.macd[idx]), macdSignal: nz(macdResult.signal[idx]), macdHist: nz(macdResult.hist[idx]),
        stochK: nz(stoch.k[idx]), stochD: nz(stoch.d[idx]),
      };
    };
    setLegend(legendAt(rows.length - 1));
    const onMove = (param: MouseEventParams) => {
      if (param.logical == null) { setLegend(legendAt(rows.length - 1)); return; }
      setLegend(legendAt(Math.max(0, Math.min(rows.length - 1, Math.round(param.logical)))));
    };
    chart.subscribeCrosshairMove(onMove);

    return () => {
      // Captured synchronously, before teardown -- see lastRangeRef's own
      // comment above for why this (not the debounced localStorage save)
      // is what the next rebuild's restore prefers.
      const live = chart.timeScale().getVisibleRange();
      if (live) lastRangeRef.current = { from: String(live.from), to: String(live.to) };
      chart.unsubscribeCrosshairMove(onMove);
      chart.timeScale().unsubscribeVisibleTimeRangeChange(onTimeRangeChange);
      if (saveTimer) clearTimeout(saveTimer);
      drawing.destroy();
      setDrawingTool(null);
      chart.remove();
      chartRef.current = null;
    };
  }, [rows, symbol, themeTick, settings, hidden]);

  // Plain call, not memoized -- cheap (a handful of getComputedStyle reads)
  // and needs to reflect the CURRENT theme on every render, including the
  // one where themeTick just flipped but the effect above hasn't rebuilt
  // the chart yet.
  const paneColors = themeColors();

  if (!rows || rows.length < 2) {
    return (
      <div style={CARD}>
        <p style={{ fontSize: 11.5, color: "var(--faint)", lineHeight: 1.5, margin: 0 }}>No local OHLCV history for {symbol || "this symbol"}.</p>
      </div>
    );
  }

  const chgUp = legend ? (legend.chgPct ?? 0) >= 0 : true;
  const priceCol = chgUp ? "var(--up)" : "var(--down)";
  const last = rows[rows.length - 1];

  const gearRow = (id: StudyId, title: string, inputs: GearField[], style: GearField[], top: number) => openGear === id ? (
    <GearPanel id={id} title={title} top={top} inputs={inputs} style={style} values={settings} hidden={hidden[id]}
      onChange={updateSettings} onToggleHidden={() => toggleHidden(id)} onClose={() => setOpenGear(null)}
      onReset={() => saveStudySettings(loadStudySettings())} />
  ) : null;

  return (
    <div style={CARD}>
      {/* No instrument header here (ticker/company/price/change/as-of) --
          the ticker page's own header above this card already shows all of
          it; repeating it here was a straight duplicate. Watchlist's own
          selected-ticker header plays the same role there. */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", padding: "0 2px 10px" }}>
        <span style={{ fontSize: 9.5, fontWeight: 700, letterSpacing: ".12em", color: "var(--faint)" }}>{legend?.date ?? last.date}</span>
        <div style={{ flex: "1 1 12px" }} />
        <div style={{ display: "flex", gap: 2, padding: 2, borderRadius: 9, background: "var(--soft)" }}>
          {RANGE_BUTTONS.map((r) => (
            <button key={r.id} type="button" onClick={() => applyRange(r.id)} style={{ fontSize: 10, fontWeight: 800, padding: "4px 10px", borderRadius: 6, border: "none", cursor: "pointer", background: activeRangeId === r.id ? "var(--accent)" : "transparent", color: activeRangeId === r.id ? "#fff" : "var(--muted)" }}>{r.id}</button>
          ))}
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "stretch" }}>
        <ChartToolbar tool={drawingTool} height={chartHeight} />
        <div style={{ position: "relative", flex: "1 1 auto", minWidth: 0 }}>
          <div ref={containerRef} style={{ width: "100%", height: chartHeight, borderRadius: 8, overflow: "hidden" }} />

          {/* Price-pane legend stack -- symbol/OHLC row + one row per overlay study. */}
          <div style={{ position: "absolute", left: 9, top: 6, maxWidth: "88%", zIndex: 5, pointerEvents: "none", display: "flex", flexDirection: "column", gap: 1, fontFamily: MONO, fontSize: 11.5, lineHeight: 1.3 }}>
            <div style={{ ...LEGEND_ROW, flexWrap: "wrap", background: paneColors.paneTag }}>
              <span style={{ fontFamily: "var(--font-sans, sans-serif)", fontSize: 12, fontWeight: 600, color: "var(--text)" }}>{companyName || symbol}</span>
              <span style={{ color: "var(--muted)" }}>· 1D · IDX</span>
              {legend ? (<>
                <span style={{ color: "var(--faint)" }}>O</span><span style={{ color: priceCol }}>{formatPrice(legend.o)}</span>
                <span style={{ color: "var(--faint)" }}>H</span><span style={{ color: priceCol }}>{formatPrice(legend.h)}</span>
                <span style={{ color: "var(--faint)" }}>L</span><span style={{ color: priceCol }}>{formatPrice(legend.l)}</span>
                <span style={{ color: "var(--faint)" }}>C</span><span style={{ color: priceCol }}>{formatPrice(legend.c)}</span>
                {legend.chg != null ? <span style={{ color: priceCol }}>{legend.chg >= 0 ? "+" : ""}{formatPrice(legend.chg)} ({formatPercent(legend.chgPct)})</span> : null}
              </>) : null}
            </div>

            <div style={{ ...LEGEND_ROW, background: paneColors.paneTag, opacity: hidden.volume ? 0.45 : 1 }}>
              <span style={{ color: "var(--muted)" }}>Vol <span style={{ color: "var(--faint)" }}>{settings.volMaLen}</span></span>
              {legend ? (<><b style={{ color: "var(--text)", fontWeight: 500 }}>{formatCompact(legend.vol)}</b><span style={{ color: settings.volMaColor }}>{formatCompact(0)}</span></>) : null}
              <EyeButton hidden={hidden.volume} onClick={() => toggleHidden("volume")} />
              <GearButton onClick={() => setOpenGear(openGear === "volume" ? null : "volume")} />
            </div>
            {gearRow("volume", "Volume", [{ kind: "number", key: "volMaLen", label: "MA Length", min: 1, max: 200 }], [{ kind: "color", key: "volUpColor", label: "Up Color" }, { kind: "color", key: "volDnColor", label: "Down Color" }, { kind: "color", key: "volMaColor", label: "MA Color" }, { kind: "number", key: "volMaWidth", label: "MA Width", min: 1, max: 4 }], 30)}

            <div style={{ ...LEGEND_ROW, background: paneColors.paneTag, opacity: hidden.vwap ? 0.45 : 1 }}>
              <span style={{ whiteSpace: "nowrap", color: "var(--muted)" }}>VWAP Suite v2.0 <span style={{ color: "var(--faint)" }}>{periodLabel(vwapCurrentKey(rows, settings), settings.vwapAnchor as VwapAnchor)}</span></span>
              <EyeButton hidden={hidden.vwap} onClick={() => toggleHidden("vwap")} />
              <GearButton onClick={() => setOpenGear(openGear === "vwap" ? null : "vwap")} />
            </div>
            {gearRow("vwap", "VWAP Suite", [{ kind: "select", key: "vwapAnchor", label: "Anchor", options: ["week", "month", "quarter", "year"] }, { kind: "select", key: "vwapSource", label: "Source", options: Object.keys(VWAP_SOURCE_FN) }, { kind: "number", key: "vwapMult1", label: "Band 1 σ", min: 1, max: 4 }, { kind: "number", key: "vwapMult2", label: "Band 2 σ", min: 1, max: 5 }], [{ kind: "color", key: "vwapColor", label: "Line Color" }, { kind: "number", key: "vwapWidth", label: "Width", min: 1, max: 4 }, { kind: "check", key: "vwapBands", label: "Band Fill" }, { kind: "check", key: "vwapLines", label: "Forward Lines" }, { kind: "check", key: "vwapText", label: "Margin Text" }], 54)}

            <div style={{ ...LEGEND_ROW, background: paneColors.paneTag, opacity: hidden.ib ? 0.45 : 1 }}>
              <span style={{ whiteSpace: "nowrap", color: "var(--muted)" }}>Monthly IBH ~ IBL <span style={{ color: "var(--faint)" }}>{settings.ibDays}</span></span>
              {legend?.ib ? <b style={{ color: settings.ibColor, fontWeight: 500 }}>{formatPrice(legend.ib.ibHigh)} ~ {formatPrice(legend.ib.ibLow)}</b> : null}
              <EyeButton hidden={hidden.ib} onClick={() => toggleHidden("ib")} />
              <GearButton onClick={() => setOpenGear(openGear === "ib" ? null : "ib")} />
            </div>
            {gearRow("ib", "Monthly IBH ~ IBL", [{ kind: "number", key: "ibDays", label: "IB Sessions", min: 1, max: 10 }], [{ kind: "color", key: "ibColor", label: "Box Color" }, { kind: "check", key: "ibFill", label: "Fill" }, { kind: "number", key: "ibWidth", label: "Border Width", min: 1, max: 4 }], 78)}

            <div style={{ ...LEGEND_ROW, background: paneColors.paneTag, opacity: hidden.ribbon ? 0.45 : 1 }}>
              <span style={{ whiteSpace: "nowrap", color: "var(--muted)" }}>EMA Ribbon <span style={{ color: "var(--faint)" }}>{settings.ribbon1Len} {settings.ribbon2Len} close</span></span>
              {legend?.ribbon1 != null ? <b style={{ color: settings.ribbonColor, fontWeight: 500 }}>{formatPrice(legend.ribbon1)}</b> : null}
              {legend?.ribbon2 != null ? <b style={{ color: hexA(settings.ribbonColor, 0.55), fontWeight: 500 }}>{formatPrice(legend.ribbon2)}</b> : null}
              <EyeButton hidden={hidden.ribbon} onClick={() => toggleHidden("ribbon")} />
              <GearButton onClick={() => setOpenGear(openGear === "ribbon" ? null : "ribbon")} />
            </div>
            {gearRow("ribbon", "EMA Ribbon", [{ kind: "number", key: "ribbon1Len", label: "EMA 1 Length", min: 2, max: 200 }, { kind: "number", key: "ribbon2Len", label: "EMA 2 Length", min: 2, max: 200 }], [{ kind: "color", key: "ribbonColor", label: "Color" }, { kind: "number", key: "ribbonWidth", label: "Width", min: 1, max: 4 }], 102)}

            <div style={{ ...LEGEND_ROW, background: paneColors.paneTag, opacity: hidden.sma200 ? 0.45 : 1 }}>
              <span style={{ color: "var(--muted)" }}>SMA <span style={{ color: "var(--faint)" }}>{settings.sma200Len} close</span></span>
              {legend?.sma200 != null ? <b style={{ color: "var(--text)", fontWeight: 500 }}>{formatPrice(legend.sma200)}</b> : null}
              <EyeButton hidden={hidden.sma200} onClick={() => toggleHidden("sma200")} />
              <GearButton onClick={() => setOpenGear(openGear === "sma200" ? null : "sma200")} />
            </div>
            {gearRow("sma200", "SMA", [{ kind: "number", key: "sma200Len", label: "Length", min: 1, max: 500 }], [{ kind: "color", key: "smaColor", label: "Color" }, { kind: "number", key: "smaWidth", label: "Width", min: 1, max: 4 }], 126)}
          </div>

          {/* Oscillator panes -- each shown live (overlay legend on its own
              canvas pane, matching the price-pane corner labels above) or,
              when hidden, collapsed to a short static row below everything
              so its eye toggle stays reachable without taking real height. */}
          {!hidden.rsi ? (
            <div style={{ position: "absolute", left: 9, top: rsiTop + 5, zIndex: 5, ...LEGEND_ROW, fontFamily: MONO, fontSize: 11.5, background: paneColors.paneTag }}>
              <span style={{ color: "var(--muted)" }}>RSI <span style={{ color: "var(--faint)" }}>{settings.rsiLen} close</span></span>
              {legend?.rsi != null ? <b style={{ color: settings.rsiColor, fontWeight: 500 }}>{fmtOsc(legend.rsi)}</b> : null}
              {legend?.rsiMa != null ? <b style={{ color: legend.rsiMaBull ? settings.rsiMaBullColor : settings.rsiMaBearColor, fontWeight: 500 }}>{fmtOsc(legend.rsiMa)}</b> : null}
              <DivToggleButton label="Bull" active={!hidden.rsiDivBull} color={settings.rsiDivBullColor} onClick={() => toggleHidden("rsiDivBull")} />
              <DivToggleButton label="Bear" active={!hidden.rsiDivBear} color={settings.rsiDivBearColor} onClick={() => toggleHidden("rsiDivBear")} />
              <EyeButton hidden={false} onClick={() => toggleHidden("rsi")} />
              <GearButton onClick={() => setOpenGear(openGear === "rsi" ? null : "rsi")} />
              {gearRow("rsi", "RSI", [
                { kind: "number", key: "rsiLen", label: "RSI Length", min: 2, max: 100 }, { kind: "number", key: "rsiMaLen", label: "EMA of RSI Length", min: 1, max: 100 },
                { kind: "number", key: "rsiUpper", label: "Upper Band", min: 51, max: 99 }, { kind: "number", key: "rsiLower", label: "Lower Band", min: 1, max: 49 },
                { kind: "check", key: "rsiUseDmiFilter", label: "Use DMI Filter For EMA Color" }, { kind: "number", key: "rsiDmiLen", label: "DMI Length", min: 1, max: 100 }, { kind: "number", key: "rsiAdxSmoothing", label: "ADX Smoothing", min: 1, max: 100 },
                { kind: "number", key: "rsiDivLeft", label: "Divergence Pivot Left", min: 1, max: 50 }, { kind: "number", key: "rsiDivRight", label: "Divergence Pivot Right", min: 1, max: 50 },
                { kind: "check", key: "rsiDivStrict", label: "Strict HH/LL Comparison" }, { kind: "check", key: "rsiDivShowLabels", label: "Show Divergence Labels" },
              ], [
                { kind: "color", key: "rsiColor", label: "RSI Color" }, { kind: "number", key: "rsiWidth", label: "RSI Width", min: 1, max: 4 },
                { kind: "color", key: "rsiMaBullColor", label: "EMA Bullish Color" }, { kind: "color", key: "rsiMaBearColor", label: "EMA Bearish Color" }, { kind: "number", key: "rsiMaWidth", label: "EMA Width", min: 1, max: 4 },
                { kind: "check", key: "rsiFill", label: "Band Fill" },
                { kind: "color", key: "rsiDivBullColor", label: "Regular Bullish" }, { kind: "color", key: "rsiDivHiddenBullColor", label: "Hidden Bullish" },
                { kind: "color", key: "rsiDivBearColor", label: "Regular Bearish" }, { kind: "color", key: "rsiDivHiddenBearColor", label: "Hidden Bearish" },
                { kind: "number", key: "rsiDivWidth", label: "Divergence Line Width", min: 1, max: 5 },
              ], rsiTop + 24)}
            </div>
          ) : null}
          {!hidden.macd ? (
            <div style={{ position: "absolute", left: 9, top: macdTop + 5, zIndex: 5, ...LEGEND_ROW, fontFamily: MONO, fontSize: 11.5, background: paneColors.paneTag }}>
              <span style={{ whiteSpace: "nowrap", color: "var(--muted)" }}>MACD 4C Smooth <span style={{ color: "var(--faint)" }}>{settings.macdFast} {settings.macdSlow} {settings.macdSignal} {settings.macdSmooth} close</span></span>
              {legend?.macdHist != null ? <b style={{ color: "var(--text)", fontWeight: 500 }}>{legend.macdHist.toFixed(0)}</b> : null}
              {legend?.macd != null ? <b style={{ color: settings.macdColor, fontWeight: 500 }}>{legend.macd.toFixed(0)}</b> : null}
              {legend?.macdSignal != null ? <b style={{ color: settings.macdSignalColor, fontWeight: 500 }}>{legend.macdSignal.toFixed(0)}</b> : null}
              <EyeButton hidden={false} onClick={() => toggleHidden("macd")} />
              <GearButton onClick={() => setOpenGear(openGear === "macd" ? null : "macd")} />
              {gearRow("macd", "MACD 4C Smooth", [{ kind: "number", key: "macdFast", label: "Fast Length", min: 1, max: 100 }, { kind: "number", key: "macdSlow", label: "Slow Length", min: 1, max: 200 }, { kind: "number", key: "macdSignal", label: "Signal Length", min: 1, max: 100 }, { kind: "number", key: "macdSmooth", label: "Smoothing", min: 1, max: 50 }], [{ kind: "color", key: "macdColor", label: "MACD Color" }, { kind: "number", key: "macdWidth", label: "MACD Width", min: 1, max: 4 }, { kind: "color", key: "macdSignalColor", label: "Signal Color" }, { kind: "number", key: "macdSignalWidth", label: "Signal Width", min: 1, max: 4 }, { kind: "color", key: "macdHistUp", label: "Hist Up" }, { kind: "color", key: "macdHistDn", label: "Hist Down" }], macdTop + 24)}
            </div>
          ) : null}
          {!hidden.stoch ? (
            <div style={{ position: "absolute", left: 9, top: stochTop + 5, zIndex: 5, ...LEGEND_ROW, fontFamily: MONO, fontSize: 11.5, background: paneColors.paneTag }}>
              <span style={{ whiteSpace: "nowrap", color: "var(--muted)" }}>Stoch RSI <span style={{ color: "var(--faint)" }}>{settings.stochSmoothK} {settings.stochSmoothD} {settings.stochRsiLen} {settings.stochLen} close</span></span>
              {legend?.stochK != null ? <b style={{ color: settings.stochKColor, fontWeight: 500 }}>{fmtOsc(legend.stochK)}</b> : null}
              {legend?.stochD != null ? <b style={{ color: settings.stochDColor, fontWeight: 500 }}>{fmtOsc(legend.stochD)}</b> : null}
              <EyeButton hidden={false} onClick={() => toggleHidden("stoch")} />
              <GearButton onClick={() => setOpenGear(openGear === "stoch" ? null : "stoch")} />
              {gearRow("stoch", "Stoch RSI", [{ kind: "number", key: "stochSmoothK", label: "Smooth K", min: 1, max: 50 }, { kind: "number", key: "stochSmoothD", label: "Smooth D", min: 1, max: 50 }, { kind: "number", key: "stochRsiLen", label: "RSI Length", min: 2, max: 100 }, { kind: "number", key: "stochLen", label: "Stochastic Length", min: 2, max: 100 }, { kind: "number", key: "stochUpper", label: "Upper Band", min: 51, max: 99 }, { kind: "number", key: "stochLower", label: "Lower Band", min: 1, max: 49 }], [{ kind: "color", key: "stochKColor", label: "K Color" }, { kind: "number", key: "stochKWidth", label: "K Width", min: 1, max: 4 }, { kind: "color", key: "stochDColor", label: "D Color" }, { kind: "number", key: "stochDWidth", label: "D Width", min: 1, max: 4 }, { kind: "check", key: "stochFill", label: "Band Fill" }], stochTop + 24)}
            </div>
          ) : null}

          {selection ? (
            <SelectionToolbar sel={selection}
              onRestyle={(color, width, style) => drawingTool?.restyleSelected(color, width, style)}
              onDelete={() => drawingTool?.deleteSelected()}
              onClose={() => setSelection(null)} />
          ) : null}
        </div>
      </div>

      {/* Collapsed strips for hidden oscillators -- kept mounted (not just
          removed) so the eye toggle to bring one back stays reachable. */}
      {(["rsi", "macd", "stoch"] as const).filter((id) => hidden[id]).map((id) => (
        <div key={id} style={{ display: "flex", alignItems: "center", gap: 6, height: OSC_H_HIDDEN, paddingLeft: 9, borderTop: "1px solid var(--hair)", fontFamily: MONO, fontSize: 11, color: "var(--faint)" }}>
          <span>{id === "rsi" ? `RSI ${settings.rsiLen}` : id === "macd" ? "MACD 4C Smooth" : "Stoch RSI"}</span>
          <EyeButton hidden onClick={() => toggleHidden(id)} />
        </div>
      ))}

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "baseline", fontSize: 10, lineHeight: 1.55, color: "var(--faint)", marginTop: 12, padding: "10px 2px 0", borderTop: "1px solid var(--hair)" }}>
        <span style={{ flex: "1 1 320px", minWidth: 0 }}>
          Monthly IBH ~ IBL draws each month&apos;s first {settings.ibDays}-session range as a box held to month end. VWAP Suite anchors on {settings.vwapSource}·volume per {settings.vwapAnchor}, labelling ±1σ/±2σ of the live period and ±1σ of the previous one with price and distance from close.
        </span>
        <span style={{ flex: "1 1 220px", minWidth: 0 }}>
          Panes: RSI {settings.rsiLen} with its RSI-based MA, MACD 4C Smooth {settings.macdFast} {settings.macdSlow} {settings.macdSignal} {settings.macdSmooth}, Stoch RSI {settings.stochSmoothK} {settings.stochSmoothD} {settings.stochRsiLen} {settings.stochLen}. Daily EOD bars — {rows.length} sessions to {last.date}, source yfinance EOD.
        </span>
      </div>
    </div>
  );
}

function hexA(hex: string, alpha: number): string {
  const n = parseInt(hex.replace("#", ""), 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}
/** Which range button (if any) a restored {from,to} view corresponds to --
    "" when the viewer left it zoomed/panned somewhere that isn't exactly
    one of the presets, so none of the buttons should read as active. */
function matchRangeButton(rows: OhlcvRow[], range: { from: string; to: string }): string {
  const toIdx = rows.findIndex((r) => r.date === range.to);
  const fromIdx = rows.findIndex((r) => r.date === range.from);
  if (toIdx !== rows.length - 1 || fromIdx === -1) return "";
  const span = toIdx - fromIdx;
  const hit = RANGE_BUTTONS.find((r) => r.n != null && Math.abs(r.n - span) <= 1);
  return hit ? hit.id : fromIdx === 0 ? "All" : "";
}

function pct(v: number, close: number): string {
  const p = close ? ((v - close) / close) * 100 : 0;
  return `${p >= 0 ? "+" : ""}${p.toFixed(2)}%`;
}
function vwapCurrentKey(rows: OhlcvRow[], s: StudySettings): string | null {
  if (!rows.length) return null;
  return computeAnchoredVwap(rows, s.vwapAnchor as VwapAnchor, 1, 2, 3, s.vwapSource as VwapSource).currentKey;
}
