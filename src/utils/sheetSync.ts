/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxbniKVw6UiaT-vCA6cnvL908DQ5Ezzs9x6Llz2PfwL7A53rXqj9SLfDZ05Rqg44jON8g/exec";

export type ModuleKey =
  | 'dse'
  | 'onlineInvestments'
  | 'sukuk'
  | 'mutualFunds'
  | 'fixedDeposits'
  | 'incomeExpense';

// ── Per-module sync metadata stored in localStorage ───────────────────────
export interface SyncMeta {
  lastModified: number;   // Last time local data was written (create/edit/delete/import)
  lastSyncedAt: number;   // Last time a successful push to Sheets completed
  isDirty: boolean;       // True if local is ahead of cloud
}

const META_KEY = (m: ModuleKey) => `syncmeta_${m}`;

// ── Safe localStorage wrapper (incognito-proof) ───────────────────────────
function lsGet(key: string): string | null {
  try { return localStorage.getItem(key); } catch { return null; }
}
function lsSet(key: string, value: string): void {
  try { localStorage.setItem(key, value); } catch {}
}

export function getMeta(module: ModuleKey): SyncMeta {
  try {
    const raw = lsGet(META_KEY(module));
    if (raw) return JSON.parse(raw);
  } catch {}
  return { lastModified: 0, lastSyncedAt: 0, isDirty: false };
}

export function setMeta(module: ModuleKey, meta: SyncMeta): void {
  lsSet(META_KEY(module), JSON.stringify(meta));
}

export function getCachedData(module: ModuleKey): any[] {
  try {
    const cached = lsGet(`sheet_cache_${module}`);
    if (!cached) return [];
    const parsed = JSON.parse(cached);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

// ── Event dispatch for real-time UI synchronization ───────────────────────
export function dispatchSyncEvent(detail: {
  module: ModuleKey;
  status: 'idle' | 'syncing' | 'success' | 'error';
  message?: string;
  lastSyncedAt?: number;
}): void {
  try {
    if (typeof window !== 'undefined') {
      setTimeout(() => {
        try {
          window.dispatchEvent(new CustomEvent('app-sync-status', { detail }));
        } catch {}
      }, 0);
    }
  } catch {}
}

// ── Serialization & Deserialization helpers for Income & Expense ──────────
export function serializeIncomeExpense(state: {
  incomeExpenseTransactions?: any[];
  incomeExpenseAccounts?: any[];
  incomeExpenseCategories?: any[];
  conversionRates?: any;
}): any[] {
  const transactions = state.incomeExpenseTransactions || [];
  const accounts = state.incomeExpenseAccounts || [];
  const categories = state.incomeExpenseCategories || [];
  const conversionRates = state.conversionRates || { USD_to_BDT: 120, LYD_to_BDT: 20 };

  return [
    ...transactions.map(t => ({
      id: t.id,
      type: 'transaction',
      data: t
    })),
    ...accounts.map(a => ({
      id: a.id,
      type: 'account',
      data: a
    })),
    ...categories.map(c => ({
      id: c.id,
      type: 'category',
      data: c
    })),
    {
      id: '__INCOME_EXPENSE_SETTINGS__',
      type: 'settings',
      data: {
        conversionRates
      }
    }
  ];
}

export function deserializeIncomeExpense(rows: any[]) {
  if (!Array.isArray(rows)) return null;

  const hasTypedRows = rows.some((r: any) => r && (r.type === 'transaction' || r.type === 'account' || r.type === 'category' || r.type === 'settings'));

  if (hasTypedRows) {
    const transactions = rows
      .filter((r: any) => r && r.type === 'transaction')
      .map((r: any) => r.data || r);

    const accounts = rows
      .filter((r: any) => r && r.type === 'account')
      .map((r: any) => r.data || r);

    const categories = rows
      .filter((r: any) => r && r.type === 'category')
      .map((r: any) => r.data || r);

    const settingsRow = rows.find((r: any) => r && r.type === 'settings');

    return {
      transactions,
      accounts,
      categories,
      conversionRates: settingsRow?.data?.conversionRates || settingsRow?.conversionRates
    };
  }

  // Fallback: if rows are raw transactions from direct spreadsheet edits
  const transactions = rows.filter((r: any) => r && (r.date || r.amount !== undefined));
  return {
    transactions,
    accounts: [],
    categories: [],
    conversionRates: undefined
  };
}

// ── Mark local data as dirty (call after every write, including import) ───
export function markDirty(module: ModuleKey): void {
  const meta = getMeta(module);
  setMeta(module, {
    ...meta,
    lastModified: Date.now(),
    isDirty: true,
  });
}

// ── Mark sync complete ─────────────────────────────────────────────────────
export function markSynced(module: ModuleKey, syncStartedAt: number): void {
  const meta = getMeta(module);

  if (meta.lastModified > syncStartedAt) {
    setMeta(module, {
      ...meta,
      lastSyncedAt: Date.now(),
      isDirty: true,
    });
    return;
  }

  setMeta(module, {
    ...meta,
    lastSyncedAt: Date.now(),
    isDirty: false,
  });
}

// ── Fetch all rows for a module from the sheet ─────────────────────────────
// Only called on app startup or manual refresh. Returns null on network failure.
export async function fetchFromSheet(module: ModuleKey): Promise<{ data: any[]; lastModified: number } | null> {
  try {
    const res = await fetch(`${SCRIPT_URL}?module=${module}&t=${Date.now()}`);
    const json = await res.json();
    if (json && json.success) {
      const data = Array.isArray(json.data?.transactions) ? json.data.transactions : [];
      const remoteLastModified: number = json.data?.lastModified || 0;
      return { data, lastModified: remoteLastModified };
    }
    return null;
  } catch {
    return null;
  }
}

// ── Push entire local dataset or single row to Sheets ─────────────────────
export async function pushToSheet(
  module: ModuleKey,
  action: 'update' | 'delete' | 'replaceAll',
  payload: { data?: any; id?: string }
): Promise<boolean> {
  try {
    const body: any = { action, module };
    if (payload.data !== undefined) body.data = payload.data;
    if (payload.id !== undefined) body.id = payload.id;

    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify(body),
    });
    const json = await res.json();
    const isSuccess = json?.success === true;

    if (isSuccess) {
      setMeta(module, {
        ...getMeta(module),
        lastSyncedAt: Date.now(),
        isDirty: false
      });
      dispatchSyncEvent({
        module,
        status: 'success',
        lastSyncedAt: Date.now(),
        message: 'Saved to Google Sheets.'
      });
    }

    return isSuccess;
  } catch {
    return false;
  }
}

// Track active sync promises to ensure latest state is always sent without conflicts
const inFlightSyncs: Partial<Record<ModuleKey, Promise<boolean>>> = {};
const pendingSyncData: Partial<Record<ModuleKey, any[]>> = {};

// ── Push full dataset for a module immediately without delay ──────────────
export async function pushModuleData(module: ModuleKey, data: any[]): Promise<boolean> {
  const syncStartedAt = Date.now();

  // Save to local cache immediately
  lsSet(`sheet_cache_${module}`, JSON.stringify(data));
  lsSet(`sheet_cache_time_${module}`, String(syncStartedAt));
  markDirty(module);

  // If there's an ongoing sync for this module, queue the latest data
  if (inFlightSyncs[module]) {
    pendingSyncData[module] = data;
    return inFlightSyncs[module]!;
  }

  dispatchSyncEvent({ module, status: 'syncing' });

  const syncPromise = (async () => {
    try {
      const success = await pushToSheet(module, 'replaceAll', { data });

      if (success) {
        markSynced(module, syncStartedAt);
        dispatchSyncEvent({
          module,
          status: 'success',
          lastSyncedAt: Date.now(),
          message: 'Saved to Google Sheets.'
        });
      } else {
        dispatchSyncEvent({
          module,
          status: 'error',
          message: 'Google Sheets sync failed.'
        });
      }

      delete inFlightSyncs[module];

      // If newer data arrived while this request was flying, immediately push the newest state
      if (pendingSyncData[module] !== undefined) {
        const nextData = pendingSyncData[module]!;
        delete pendingSyncData[module];
        return pushModuleData(module, nextData);
      }

      return success;
    } catch {
      delete inFlightSyncs[module];
      dispatchSyncEvent({
        module,
        status: 'error',
        message: 'Google Sheets sync failed.'
      });
      return false;
    }
  })();

  inFlightSyncs[module] = syncPromise;
  return syncPromise;
}

// ── Instant direct push wrapper (zero-delay) ──────────────────────────────
export function schedulePush(module: ModuleKey, getLatestData: () => any[]): void {
  const data = getLatestData();
  pushModuleData(module, data);
}

// ── On app startup: compare local vs cloud, resolve conflict ──────────────
// Returns the data that should be used (cloud wins if available, matching DSE Tracker).
export async function resolveOnStartup(
  module: ModuleKey,
  localData: any[],
  pushData: () => any[]
): Promise<any[] | null> {
  const meta = getMeta(module);

  const remote = await fetchFromSheet(module);
  if (!remote) {
    // Network failure — keep local data, do NOT push
    return null;
  }

  // If cloud has data, cloud data should always win on startup (matching DSE Tracker pattern)
  if (remote.data && remote.data.length > 0) {
    localStorage.setItem(`sheet_cache_${module}`, JSON.stringify(remote.data));
    localStorage.setItem(`sheet_cache_time_${module}`, String(Date.now()));
    setMeta(module, {
      lastModified: remote.lastModified || Date.now(),
      lastSyncedAt: Date.now(),
      isDirty: false,
    });
    return remote.data;
  } else if ((!remote.data || remote.data.length === 0) && localData && localData.length > 0 && meta.isDirty) {
    // Only push local data if cloud sheet is completely empty AND local is dirty
    pushModuleData(module, pushData());
    return null;
  }

  return null;
}

// ── Flush all dirty modules before app unloads ────────────────────────────
export function flushOnUnload(getDataMap: () => Partial<Record<ModuleKey, any[]>>): void {
  const modules: ModuleKey[] = [
    'onlineInvestments',
    'sukuk',
    'mutualFunds',
    'fixedDeposits',
    'incomeExpense',
    'dse'
  ];

  const handleFlush = () => {
    const dataMap = getDataMap();
    modules.forEach(m => {
      const meta = getMeta(m);
      if (meta.isDirty && meta.lastModified > meta.lastSyncedAt) {
        const data = dataMap[m];
        if (data && data.length > 0) {
          const body = JSON.stringify({ action: 'replaceAll', module: m, data });
          if (typeof navigator !== 'undefined' && navigator.sendBeacon) {
            navigator.sendBeacon(SCRIPT_URL, body);
          } else {
            fetch(SCRIPT_URL, { method: 'POST', keepalive: true, body }).catch(() => {});
          }
          markSynced(m, Date.now());
        }
      }
    });
  };

  window.addEventListener('beforeunload', handleFlush);
  if (typeof document !== 'undefined') {
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden') {
        handleFlush();
      }
    });
  }
}

// ── Manual & Automated sync-all: compare local vs cloud timestamps, pull or push ─
export async function syncAllModules(
  getDataMap: () => Partial<Record<ModuleKey, any[]>>,
  onModuleResolved: (module: ModuleKey, data: any[] | null) => void
): Promise<{ pushed: ModuleKey[]; pulled: ModuleKey[]; failed: ModuleKey[]; unchanged: ModuleKey[] }> {
  const modules: ModuleKey[] = [
    'onlineInvestments',
    'sukuk',
    'mutualFunds',
    'fixedDeposits',
    'incomeExpense',
    'dse'
  ];
  const pushed: ModuleKey[] = [];
  const pulled: ModuleKey[] = [];
  const failed: ModuleKey[] = [];
  const unchanged: ModuleKey[] = [];

  await Promise.all(modules.map(async (module) => {
    try {
      const meta = getMeta(module);
      const localData = getDataMap()[module] ?? [];
      const remote = await fetchFromSheet(module);

      if (!remote) {
        // Network failure — keep local data safely
        failed.push(module);
        return;
      }

      const remoteLastModified = remote.lastModified || 0;
      const localLastModified = meta.lastModified || 0;
      const isDirty = meta.isDirty === true;
      const remoteCount = Array.isArray(remote.data) ? remote.data.length : 0;
      const localCount = Array.isArray(localData) ? localData.length : 0;

      // Case 1: Remote is newer OR local was completely empty while remote has data
      if (
        (remoteCount > 0 && localCount === 0) ||
        (remoteCount > 0 && remoteLastModified > localLastModified) ||
        (remoteCount > 0 && !isDirty && localLastModified === 0)
      ) {
        setMeta(module, {
          lastModified: remoteLastModified || Date.now(),
          lastSyncedAt: Date.now(),
          isDirty: false,
        });
        localStorage.setItem(`sheet_cache_${module}`, JSON.stringify(remote.data));
        localStorage.setItem(`sheet_cache_time_${module}`, String(Date.now()));
        onModuleResolved(module, remote.data);
        pulled.push(module);
      }
      // Case 2: Local is newer and marked dirty, OR remote is empty while local has data
      else if (isDirty && localCount > 0 && (localLastModified >= remoteLastModified || remoteCount === 0)) {
        const ok = await pushModuleData(module, localData);
        if (ok) {
          pushed.push(module);
        } else {
          failed.push(module);
        }
        onModuleResolved(module, null);
      }
      // Case 3: In sync / unchanged
      else {
        unchanged.push(module);
        onModuleResolved(module, null);
      }
    } catch {
      failed.push(module);
    }
  }));

  return { pushed, pulled, failed, unchanged };
}

// ── Background Auto-Sync loop: every 5 minutes ─────────────────────────────
let autoSyncStarted = false;

export function startAutoSync(
  syncAction: () => Promise<any>
): void {
  if (autoSyncStarted) return;
  autoSyncStarted = true;

  const FIVE_MINUTES = 5 * 60 * 1000;
  setInterval(async () => {
    try {
      await syncAction();
    } catch (err) {
      console.error('5-minute auto sync error', err);
    }
  }, FIVE_MINUTES);
}
