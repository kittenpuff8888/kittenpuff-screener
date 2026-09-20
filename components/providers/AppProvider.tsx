"use client";

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from "react";
import { usePathname, useRouter } from "next/navigation";
import { buildKongloMembership, loadIndexes } from "@/lib/data/indexes";
import { loadKsei } from "@/lib/data/ksei";
import { loadMarketContext, type MarketContextPayload } from "@/lib/data/marketContext";
import { loadManifest, resolveMarketDate } from "@/lib/data/metadata";
import { loadResearchBundle } from "@/lib/data/screener";
import { fetchLive, isMarketOpen, LIVE_ENDPOINT, LIVE_POLL_MS, type LiveSnapshot } from "@/lib/data/live";
import { loadIdxIndex, type IdxIndexPayload } from "@/lib/data/idxIndex";
import { loadWatchlist, newGroupId, saveWatchlist } from "@/lib/data/watchlistStore";
import type { IndexPayload, KseiPayload, Manifest, ResearchBundle } from "@/lib/domain/types";

const WATCHLIST_KEY = "idx_watchlist_tickers";

type AppContextValue = {
  manifest: Manifest | null;
  marketDate: string;
  notice: string | null;
  bundle: ResearchBundle | null;
  indexes: IndexPayload | null;
  marketContext: MarketContextPayload | null;
  ksei: KseiPayload | null;
  idxIndex: IdxIndexPayload | null;
  live: LiveSnapshot | null;
  loading: boolean;
  error: string | null;
  selectedTicker: string | null;
  watchlist: string[];
  tickerOptions: Array<{ ticker: string; label: string; sector: string }>;
  setMarketDate: (date: string) => void;
  openTicker: (ticker: string) => void;
  closeTicker: () => void;
  toggleWatchlist: (ticker: string) => void;
  isWatched: (ticker: string) => boolean;
  reload: () => void;
};

const AppContext = createContext<AppContextValue | null>(null);

/** The Watchlist board (DESIGN_SPEC §3.4) owns the persisted model: named
    groups of rows under `idxr:watchlist`. The ticker page's Watch button is a
    second door into that same model, so both read and write it here rather
    than keeping a parallel flat list that the board would silently ignore.
    WATCHLIST_KEY below is the pre-redesign flat array, migrated once. */
const STARRED_GROUP = "Starred";

function readWatchlist(): string[] {
  if (typeof window === "undefined") return [];
  const state = loadWatchlist();
  const fromGroups = state.groups.flatMap((g) => g.rows.map((r) => r.symbol));
  if (fromGroups.length || state.groups.length) return fromGroups;
  // One-time migration off the old flat key.
  try {
    const parsed = JSON.parse(window.localStorage.getItem(WATCHLIST_KEY) || "[]");
    const legacy = Array.isArray(parsed) ? parsed.map((item) => String(item).toUpperCase()) : [];
    if (legacy.length) {
      saveWatchlist({
        groups: [{ id: newGroupId(), name: STARRED_GROUP, rows: legacy.map((s) => ({ symbol: s, addedAt: "", addedClose: null })) }],
        activeGroupId: null,
        selectedSymbol: null,
      });
      window.localStorage.removeItem(WATCHLIST_KEY);
    }
    return legacy;
  } catch {
    return [];
  }
}

/** Add or remove a symbol, keeping it in whichever group already holds it and
    falling back to a "Starred" group when the board is still empty. */
function writeWatchlistSymbol(symbol: string, add: boolean) {
  if (typeof window === "undefined") return;
  const state = loadWatchlist();
  if (!add) {
    saveWatchlist({ ...state, groups: state.groups.map((g) => ({ ...g, rows: g.rows.filter((r) => r.symbol !== symbol) })) });
    return;
  }
  if (state.groups.some((g) => g.rows.some((r) => r.symbol === symbol))) return;
  const target = state.groups.find((g) => g.id === state.activeGroupId) || state.groups[0];
  const row = { symbol, addedAt: new Date().toISOString().slice(0, 10), addedClose: null };
  if (target) {
    saveWatchlist({ ...state, groups: state.groups.map((g) => (g.id === target.id ? { ...g, rows: [...g.rows, row] } : g)) });
    return;
  }
  const id = newGroupId();
  saveWatchlist({ groups: [{ id, name: STARRED_GROUP, rows: [row] }], activeGroupId: id, selectedSymbol: symbol });
}

export function AppProvider({ children }: { children: ReactNode }) {
  const router = useRouter();
  const pathname = usePathname();
  const [manifest, setManifest] = useState<Manifest | null>(null);
  const [marketDate, setMarketDateState] = useState("");
  const [notice, setNotice] = useState<string | null>(null);
  const [bundle, setBundle] = useState<ResearchBundle | null>(null);
  const [indexes, setIndexes] = useState<IndexPayload | null>(null);
  const [marketContext, setMarketContext] = useState<MarketContextPayload | null>(null);
  const [ksei, setKsei] = useState<KseiPayload | null>(null);
  const [idxIndex, setIdxIndex] = useState<IdxIndexPayload | null>(null);
  const [live, setLive] = useState<LiveSnapshot | null>(null);
  const [selectedTicker, setSelectedTicker] = useState<string | null>(null);
  const [watchlist, setWatchlist] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [reloadToken, setReloadToken] = useState(0);

  useEffect(() => {
    setWatchlist(readWatchlist());
  }, []);

  useEffect(() => {
    let cancelled = false;
    async function boot() {
      setLoading(true);
      setError(null);
      try {
        const [nextManifest, nextIndexes, nextKsei, nextMarketContext, nextIdxIndex] = await Promise.all([
          loadManifest(),
          loadIndexes(),
          loadKsei(),
          loadMarketContext(),
          loadIdxIndex(),
        ]);
        if (cancelled) return;
        setIdxIndex(nextIdxIndex);
        const requested = typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("date") : null;
        const resolved = resolveMarketDate(nextManifest, requested);
        setManifest(nextManifest);
        setIndexes(nextIndexes);
        setMarketContext(nextMarketContext);
        setKsei(nextKsei);
        setMarketDateState(resolved.marketDate);
        setNotice(resolved.notice || null);
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : "Unable to initialize IDX Research.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    boot();
    return () => {
      cancelled = true;
    };
  }, [reloadToken]);

  useEffect(() => {
    if (!marketDate) return;
    let cancelled = false;
    async function load() {
      setLoading(true);
      setError(null);
      try {
        const membership = buildKongloMembership(indexes);
        const nextBundle = await loadResearchBundle(marketDate, membership, ksei);
        if (!cancelled) setBundle(nextBundle);
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : `Unable to load ${marketDate}.`);
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    load();
    return () => {
      cancelled = true;
    };
  }, [marketDate, indexes, ksei]);

  // Live quote overlay. Inert unless NEXT_PUBLIC_LIVE_ENDPOINT is configured.
  // Polls only while the IDX session is open; outside hours the last snapshot
  // stands and we stop hitting the endpoint. Failures leave `live` as-is, so
  // the UI silently falls back to the archived close (never a fabricated price).
  useEffect(() => {
    if (!LIVE_ENDPOINT) return;
    let cancelled = false;
    const controller = new AbortController();
    async function tick() {
      const snapshot = await fetchLive(controller.signal);
      if (!cancelled && snapshot) setLive(snapshot);
    }
    tick();
    const id = window.setInterval(() => {
      if (isMarketOpen()) tick();
    }, LIVE_POLL_MS);
    return () => {
      cancelled = true;
      controller.abort();
      window.clearInterval(id);
    };
  }, []);

  const setMarketDate = useCallback(
    (date: string) => {
      if (!manifest) return;
      const resolved = resolveMarketDate(manifest, date);
      setMarketDateState(resolved.marketDate);
      setNotice(resolved.notice || null);
      const params = new URLSearchParams(typeof window !== "undefined" ? window.location.search : "");
      params.set("date", resolved.marketDate);
      router.replace(`${pathname}?${params.toString()}`);
    },
    [manifest, pathname, router],
  );

  const openTicker = useCallback((ticker: string) => {
    const clean = ticker.trim().toUpperCase().replace(".JK", "");
    if (clean) window.open(`/ticker?ticker=${clean}`, "_blank", "noopener,noreferrer");
  }, []);

  const closeTicker = useCallback(() => setSelectedTicker(null), []);

  const toggleWatchlist = useCallback((ticker: string) => {
    const clean = ticker.trim().toUpperCase().replace(".JK", "");
    setWatchlist((current) => {
      const adding = !current.includes(clean);
      writeWatchlistSymbol(clean, adding);
      return adding ? [...current, clean].sort() : current.filter((item) => item !== clean);
    });
  }, []);

  const isWatched = useCallback((ticker: string) => watchlist.includes(ticker.toUpperCase()), [watchlist]);

  const tickerOptions = useMemo(() => {
    // Real sector name from the KSEI registry (the technical record's `sector`
    // and the workbook "IDX Sector" both ship as "-"/label → "Others").
    const secOf = new Map<string, string>();
    ksei?.records.forEach((r) => { if (r.sector && r.sector !== "Others") secOf.set(r.ticker, r.sector); });
    const map = new Map<string, { ticker: string; label: string; sector: string }>();
    bundle?.technical.forEach((stock, ticker) => {
      map.set(ticker, {
        ticker,
        label: stock.companyName || ticker,
        sector: secOf.get(ticker) || "Others",
      });
    });
    ksei?.records.forEach((issuer) => {
      if (!map.has(issuer.ticker)) {
        map.set(issuer.ticker, { ticker: issuer.ticker, label: issuer.companyName, sector: issuer.sector });
      }
    });
    return [...map.values()].sort((a, b) => a.ticker.localeCompare(b.ticker));
  }, [bundle, ksei]);

  const value = useMemo<AppContextValue>(
    () => ({
      manifest,
      marketDate,
      notice,
      bundle,
      indexes,
      marketContext,
      ksei,
      idxIndex,
      live,
      loading,
      error,
      selectedTicker,
      watchlist,
      tickerOptions,
      setMarketDate,
      openTicker,
      closeTicker,
      toggleWatchlist,
      isWatched,
      reload: () => setReloadToken((token) => token + 1),
    }),
    [
      manifest,
      marketDate,
      notice,
      bundle,
      indexes,
      marketContext,
      ksei,
      idxIndex,
      live,
      loading,
      error,
      selectedTicker,
      watchlist,
      tickerOptions,
      setMarketDate,
      openTicker,
      closeTicker,
      toggleWatchlist,
      isWatched,
    ],
  );

  return <AppContext.Provider value={value}>{children}</AppContext.Provider>;
}

export function useApp() {
  const context = useContext(AppContext);
  if (!context) throw new Error("useApp must be used inside AppProvider.");
  return context;
}
