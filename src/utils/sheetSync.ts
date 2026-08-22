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
interface SyncMeta {
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

function setMeta(module: ModuleKey, meta: SyncMeta): void {
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
function markSynced(module: ModuleKey, syncStartedAt: number): void {
  const meta = getMeta(module);

  // IMPORTANT:
  // If local data changed AFTER this sync started,
  // do NOT clear dirty state.
  // Otherwise older sync responses can overwrite newer changes.
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

// ── Get cached data synchronously (for initial render) ────────────────────
// (Duplicate getCachedData removed)

// ── Fetch all rows for a module from the sheet ─────────────────────────────
// Only called on app startup. Returns null on network failure.
export async function fetchFromSheet(module: ModuleKey): Promise<{ data: any[]; lastModified: number } | null> {
  try {
    const res = await fetch(`${SCRIPT_URL}?module=${module}&t=${Date.now()}`);
    const json = await res.json();
    if (json.success) {
      const data = Array.isArray(json.data?.transactions) ? json.data.transactions : [];
      const remoteLastModified: number = json.data?.lastModified || 0;
      return { data, lastModified: remoteLastModified };
    }
    return null;
  } catch {
    return null;
  }
}

// ── Push entire local dataset to Sheets ───────────────────────────────────
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
    return json.success === true;
  } catch {
    return false;
  }
}

// ── Push full dataset for a module and mark as synced ─────────────────────
export async function pushModuleData(module: ModuleKey, data: any[]): Promise<boolean> {

  // Capture the exact moment this sync started
  const syncStartedAt = Date.now();

  const success = await pushToSheet(module, 'replaceAll', { data });

  if (success) {
    lsSet(`sheet_cache_${module}`, JSON.stringify(data));

    // Only mark synced if no newer edits happened meanwhile
    markSynced(module, syncStartedAt);
  }

  return success;
}

// ── Debounce timers per module ─────────────────────────────────────────────
const debounceTimers: Partial<Record<ModuleKey, ReturnType<typeof setTimeout>>> = {};

// ── Schedule a debounced push (2s delay). getLatestData must return current state. ──
export function schedulePush(module: ModuleKey, getLatestData: () => any[]): void {
  if (debounceTimers[module]) clearTimeout(debounceTimers[module]!);

  debounceTimers[module] = setTimeout(async () => {
    const meta = getMeta(module);

    // nothing changed
    if (!meta.isDirty && meta.lastModified <= meta.lastSyncedAt) {
      return;
    }

    const latestData = getLatestData();

    const success = await pushModuleData(module, latestData);

    if (!success) {
      // keep dirty so next interval retries
      markDirty(module);
    }
  }, 500);
}

// ── On app startup: compare local vs cloud, resolve conflict ──────────────
// Returns the data that should be used (local wins if newer, cloud if newer).
// If local is newer, schedules an immediate push.
export async function resolveOnStartup(
  module: ModuleKey,
  localData: any[],
  pushData: () => any[]
): Promise<any[] | null> {
  const meta = getMeta(module);
  const localLastModified = meta.lastModified;

  const remote = await fetchFromSheet(module);
  if (!remote) {
    // Network failure — keep local, schedule push if dirty
    if (meta.isDirty || localLastModified > meta.lastSyncedAt) {
      schedulePush(module, pushData);
    }
    return null; // null = keep using local unchanged
  }

  const remoteLastModified = remote.lastModified;

  if (remoteLastModified > localLastModified) {
    // Cloud is newer → use cloud data, update local cache
    localStorage.setItem(`sheet_cache_${module}`, JSON.stringify(remote.data));
    setMeta(module, {
      lastModified: remoteLastModified,
      lastSyncedAt: Date.now(),
      isDirty: false,
    });
    return remote.data;
  } else {
    // Local is newer or equal → keep local, push to cloud
    localStorage.setItem(`sheet_cache_${module}`, JSON.stringify(localData));
    if (meta.isDirty || localLastModified > meta.lastSyncedAt) {
      // Push immediately (no debounce needed at startup)
      pushModuleData(module, pushData());
    }
    return null; // null = keep using local unchanged
  }
}

// ── Flush all dirty modules before app unloads ────────────────────────────
export function flushOnUnload(getDataMap: () => Partial<Record<ModuleKey, any[]>>): void {
  const modules: ModuleKey[] = [
  'onlineInvestments',
  'sukuk',
  'mutualFunds',
  'fixedDeposits',
  'dse',
  'incomeExpense'
];
  window.addEventListener('beforeunload', () => {
    const dataMap = getDataMap();
    modules.forEach(m => {
      const meta = getMeta(m);
      if (meta.isDirty || meta.lastModified > meta.lastSyncedAt) {
        const data = dataMap[m];
        if (data) {
          // Use sendBeacon for reliability on unload
          const body = JSON.stringify({ action: 'replaceAll', module: m, data });
          navigator.sendBeacon(SCRIPT_URL, body);
        }
      }
    });
  });
}

// ── Manual sync-all: compare local vs cloud for all modules, push or pull ─
export async function syncAllModules(
  getDataMap: () => Partial<Record<ModuleKey, any[]>>,
  onModuleResolved: (module: ModuleKey, data: any[] | null) => void
): Promise<{ pushed: ModuleKey[]; pulled: ModuleKey[]; failed: ModuleKey[] }> {
  const modules: ModuleKey[] = [
  'onlineInvestments',
  'sukuk',
  'mutualFunds',
  'fixedDeposits',
  'incomeExpense'
];
  const pushed: ModuleKey[] = [];
  const pulled: ModuleKey[] = [];
  const failed: ModuleKey[] = [];

  await Promise.all(modules.map(async (module) => {
    try {
      const meta = getMeta(module);
      const remote = await fetchFromSheet(module);

      if (!remote) {
        // Network failure — push local if dirty
        const localData = getDataMap()[module] ?? [];
        if (meta.isDirty || meta.lastModified > meta.lastSyncedAt) {
          const ok = await pushModuleData(module, localData);
          if (ok) pushed.push(module); else failed.push(module);
        }
        return;
      }

      if (remote.lastModified > meta.lastModified) {
        // Cloud is newer → pull
        setMeta(module, { lastModified: remote.lastModified, lastSyncedAt: Date.now(), isDirty: false });
        localStorage.setItem(`sheet_cache_${module}`, JSON.stringify(remote.data));
        onModuleResolved(module, remote.data);
        pulled.push(module);
      } else {
        // Local is newer or equal → push
        const localData = getDataMap()[module] ?? [];
        const ok = await pushModuleData(module, localData);
        if (ok) pushed.push(module); else failed.push(module);
        onModuleResolved(module, null);
      }
    } catch {
      failed.push(module);
    }
  }));

  return { pushed, pulled, failed };
}



// ─────────────────────────────────────────────────────────────
// AUTO PUSH EVERY 10 SECONDS (ONE WAY: APP → SHEETS)
// ─────────────────────────────────────────────────────────────

let autoSyncStarted = false;

export function startAutoSync(
  getDataMap: () => Partial<Record<ModuleKey, any[]>>
): void {
  if (autoSyncStarted) return;

  autoSyncStarted = true;

  const modules: ModuleKey[] = [
  'onlineInvestments',
  'sukuk',
  'mutualFunds',
  'fixedDeposits',
  'incomeExpense',
];
  setInterval(async () => {
    const dataMap = getDataMap();

    for (const module of modules) {
      const meta = getMeta(module);

      // skip if no changes
      if (!meta.isDirty && meta.lastModified <= meta.lastSyncedAt) {
        continue;
      }

      const data = dataMap[module] ?? [];

      await pushModuleData(module, data);
    }
  }, 10000);
}