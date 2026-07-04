/**
 * ============================================================================
 * Usage tracking service (PRD-12)
 * ----------------------------------------------------------------------------
 * Records per-day usage counters:
 *   - prompt_tokens / completion_tokens / total_tokens
 *   - number of API calls (requests)
 *   - which tools were invoked
 *   - which slash commands were used
 *   - which Excel features were used (chart, formula, mask, ...)
 *
 * Storage: localStorage with one entry per day, keyed by yyyy-mm-dd.
 * Old days are pruned on read (>30 days).
 *
 * Cost estimation: configurable price-per-million-tokens table per model.
 * Defaults reflect DeepSeek's published pricing.
 * ============================================================================
 */

const STORAGE_KEY = "dsx_usage_v1";
const MAX_DAYS = 30;

export interface UsageDay {
  date: string; // yyyy-mm-dd (local)
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  requests: number;
  errors: number;
  tools: Record<string, number>;
  slashCommands: Record<string, number>;
  features: Record<string, number>;
  model: Record<string, number>;
}

interface UsageStore {
  days: UsageDay[];
}

/* ---------------- Default price table ---------------- */

/** USD per 1M tokens. Editable from settings in a future PR. */
export const PRICE_TABLE: Record<string, { prompt: number; completion: number }> = {
  "deepseek-chat": { prompt: 0.14, completion: 0.28 },
  "deepseek-reasoner": { prompt: 0.55, completion: 2.19 },
};

export function estimateCost(model: string, prompt: number, completion: number): number {
  const p = PRICE_TABLE[model] || PRICE_TABLE["deepseek-chat"];
  return (prompt * p.prompt + completion * p.completion) / 1_000_000;
}

/* ---------------- Storage helpers ---------------- */

function loadStore(): UsageStore {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { days: [] };
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.days)) return { days: [] };
    return parsed as UsageStore;
  } catch {
    return { days: [] };
  }
}

function saveStore(store: UsageStore): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
  } catch {
    /* Quota - best effort. */
  }
}

function prune(store: UsageStore): void {
  if (store.days.length <= MAX_DAYS) return;
  store.days.sort((a, b) => (a.date < b.date ? -1 : 1));
  store.days = store.days.slice(-MAX_DAYS);
}

/** yyyy-mm-dd in local time. */
function today(): string {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${dd}`;
}

/** Get-or-create the entry for today. */
function getTodayDay(store: UsageStore): UsageDay {
  const d = today();
  let day = store.days.find((x) => x.date === d);
  if (!day) {
    day = {
      date: d,
      promptTokens: 0,
      completionTokens: 0,
      totalTokens: 0,
      requests: 0,
      errors: 0,
      tools: {},
      slashCommands: {},
      features: {},
      model: {},
    };
    store.days.push(day);
  }
  return day;
}

/* ---------------- Public recorders ---------------- */

export function recordRequest(opts: {
  model: string;
  promptTokens: number;
  completionTokens: number;
  errored?: boolean;
}): void {
  const store = loadStore();
  const day = getTodayDay(store);
  day.promptTokens += Math.max(0, opts.promptTokens || 0);
  day.completionTokens += Math.max(0, opts.completionTokens || 0);
  day.totalTokens += Math.max(0, (opts.promptTokens || 0) + (opts.completionTokens || 0));
  day.requests += 1;
  if (opts.errored) day.errors += 1;
  if (opts.model) {
    day.model[opts.model] = (day.model[opts.model] || 0) + 1;
  }
  prune(store);
  saveStore(store);
}

export function recordToolInvocation(name: string): void {
  const store = loadStore();
  const day = getTodayDay(store);
  day.tools[name] = (day.tools[name] || 0) + 1;
  saveStore(store);
}

export function recordSlashCommand(trigger: string): void {
  const store = loadStore();
  const day = getTodayDay(store);
  day.slashCommands[trigger] = (day.slashCommands[trigger] || 0) + 1;
  saveStore(store);
}

export function recordFeature(name: string): void {
  const store = loadStore();
  const day = getTodayDay(store);
  day.features[name] = (day.features[name] || 0) + 1;
  saveStore(store);
}

export function clearAllUsage(): void {
  saveStore({ days: [] });
}

/* ---------------- Aggregations ---------------- */

export function getAllDays(): UsageDay[] {
  const store = loadStore();
  return store.days.slice().sort((a, b) => (a.date < b.date ? -1 : 1));
}

export function getTodayStats(): UsageDay {
  const store = loadStore();
  return getTodayDay(store);
}

/** Top-N entries from a count map, descending. */
export function topEntries<T extends Record<string, number>>(map: T, n = 5): Array<{ key: string; count: number }> {
  return Object.entries(map || {})
    .map(([key, count]) => ({ key, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, n);
}

export function totalCostAcrossDays(model: string = "deepseek-chat"): number {
  return getAllDays().reduce(
    (n, d) => n + estimateCost(model, d.promptTokens, d.completionTokens),
    0
  );
}