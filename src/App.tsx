/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import { useAppState } from './hooks/useAppState';
import { AppState, BankAccount } from './types';
import { Sidebar, Navbar } from './components/layout/Navigation';
import { PinLogin } from './components/auth/PinLogin';
import { DashboardCharts, SummaryCard } from './components/dashboard/DashboardComponents';
import { OnlineInvestments } from './components/modules/OnlineInvestments';
import { SukukInvestments } from './components/modules/SukukInvestments';
import { DseTrackerModule } from './components/modules/DseTrackerModule';
import { MutualFundsModule } from './components/modules/MutualFundsModule';
import { FixedDepositsModule } from './components/modules/FixedDepositsModule';
import { SettingsModule } from './components/modules/SettingsModule';
import { IncomeExpenseModule, isBoAccount } from './components/modules/IncomeExpenseModule';
import { Wallet, TrendingUp, PieChart, Banknote, Landmark, Plus, Briefcase, Coins, DollarSign, Calendar, ChevronDown, ChevronLeft, ChevronRight, Settings, Download, Upload, AlertCircle, CheckCircle2, XCircle } from 'lucide-react';
import * as XLSX from 'xlsx';
import { formatBDT, cn } from './utils/formatters';
import { Button } from './components/ui/BaseComponents';
import { schedulePush, resolveOnStartup, flushOnUnload, markDirty, syncAllModules, startAutoSync, serializeIncomeExpense, deserializeIncomeExpense } from './utils/sheetSync';
import type { ModuleKey } from './utils/sheetSync';

// ─── DSE cache keys (must match DseTrackerModule) ─────────────────────────────
const DSE_CACHE_KEY = 'sheet_cache_dse';

// ─── Helpers ──────────────────────────────────────────────────────────────────
const toDateStr = (d: Date): string => {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
};

const getTodayStr = (): string => toDateStr(new Date());

const getFirstOfMonth = (date: Date): string => {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-01`;
};

const getLastOfMonth = (date: Date): string => {
  const last = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  return toDateStr(last);
};

export default function App() {
  const { state, updateState } = useAppState();
  const [activeModule, setActiveModule] = useState('dashboard');
  const [isCollapsed, setIsCollapsed] = useState(false);
  const [isMobileOpen, setIsMobileOpen] = useState(false);
  const [triggerAdd, setTriggerAdd] = useState(false);
    const [inheritedAddData, setInheritedAddData] = useState<{
    date: string;
    amount: number;
    type?: 'Transfer' | 'Income' | 'Expense' | 'Loan';
    category?: string;
    subCategory?: string;
    targetModule?: string;
    description?: string;
    linkedTxId?: string;
    accountId?: string;
    returnModule?: string;
  } | null>(null);
  const latestStateRef = useRef(state);
  // Prevent a Google Sheets pull from being immediately pushed back as a new local edit.
  const suppressNextIncomeExpenseSyncRef = useRef(false);
  const incomeExpenseSnapshotRef = useRef<string | null>(null);

  const handleClearInheritedData = useCallback(() => {
    setInheritedAddData(null);
  }, []);

  const handleNavigateToModule = (
    module: string,
    initialAddData?: {
      date: string;
      amount: number;
      type?: 'Transfer' | 'Income' | 'Expense' | 'Loan';
      category?: string;
      subCategory?: string;
      targetModule?: string;
      description?: string;
      linkedTxId?: string;
      accountId?: string;
      returnModule?: string;
    },
    openModal: boolean = true
  ) => {
    if (initialAddData) {
      setInheritedAddData(initialAddData);
    } else {
      setInheritedAddData(null);
    }
    setActiveModule(module);
    if (openModal) {
      setTriggerAdd(true);
    } else {
      setTriggerAdd(false);
    }
  };

  const handleCreateFdrAccount = useCallback((currency: 'BDT' | 'USD', accountName: string, initialDate: string) => {
    updateState(s => {
      const accs = s.incomeExpenseAccounts || [];
      const parentName = currency === 'USD' ? 'Bank USD' : 'Investments';
      const parent = accs.find(
        a => a.isParent && a.currency === currency && a.name.trim().toLowerCase() === parentName.toLowerCase()
      );

      const newAccount: BankAccount = {
        id: 'acc-' + crypto.randomUUID().slice(0, 8),
        name: accountName,
        currency,
        initialBalance: 0,
        initialDate,
        isParent: false,
        parentId: parent?.id
      };

      return {
        ...s,
        incomeExpenseAccounts: [...accs, newAccount]
      };
    });
  }, [updateState]);

useEffect(() => {
  latestStateRef.current = state;
}, [state]);

  // ─── Income & Expense ↔ Google Sheets: sync after every local change ───────
  // The I&E module writes through updateState(), so this centralized watcher
  // catches create/edit/delete/import changes as well as changes made by the
  // investment modules when they update linked I&E entries.
  useEffect(() => {
    const snapshot = JSON.stringify(serializeIncomeExpense(state));

    // Establish the baseline for the already-loaded local state. This avoids
    // treating the initial render itself as a user edit.
    if (incomeExpenseSnapshotRef.current === null) {
      incomeExpenseSnapshotRef.current = snapshot;
      return;
    }

    if (incomeExpenseSnapshotRef.current === snapshot) return;
    incomeExpenseSnapshotRef.current = snapshot;

    // A remote Sheets pull changes React state too, but it must NOT be pushed
    // straight back to Sheets. The pull is already the authoritative result.
    if (suppressNextIncomeExpenseSyncRef.current) {
      suppressNextIncomeExpenseSyncRef.current = false;
      return;
    }

    markDirty('incomeExpense');
    schedulePush('incomeExpense', () => serializeIncomeExpense(latestStateRef.current));
  }, [
    state.incomeExpenseTransactions,
    state.incomeExpenseAccounts,
    state.incomeExpenseCategories,
    state.conversionRates,
  ]);

  // Inactivity timer (3 minutes)
  useEffect(() => {
    if (state.isLocked) return;

    let timeoutId: ReturnType<typeof setTimeout>;

    const resetTimer = () => {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(() => {
        updateState(s => ({ ...s, isLocked: true }));
      }, 3 * 60 * 1000); // 3 minutes
    };

    const events = ['mousedown', 'mousemove', 'keypress', 'scroll', 'touchstart'];
    const handleInactivity = () => resetTimer();

    events.forEach(name => window.addEventListener(name, handleInactivity));
    
    resetTimer(); // Initialize timer

    return () => {
      clearTimeout(timeoutId);
      events.forEach(name => window.removeEventListener(name, handleInactivity));
    };
  }, [state.isLocked, updateState]);

  const [historyRange, setHistoryRange] = useState<'this' | 'fiscal' | 'custom'>('custom');
  const [historyCustomDates, setHistoryCustomDates] = useState(() => {
    return {
      start: '2023-01-01',
      end: getTodayStr(),
    };
  });
  const [historyThisMonthDate, setHistoryThisMonthDate] = useState(() => {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1);
  });
  const [historyFiscalStartYear, setHistoryFiscalStartYear] = useState(() => {
    const now = new Date();
    return now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
  });

  const [isRangeMenuOpen, setIsRangeMenuOpen] = useState(false);
  const [isSettingsMenuOpen, setIsSettingsMenuOpen] = useState(false);

  const [mutualFundsTitle, setMutualFundsTitle] = useState<React.ReactNode>('Mutual Funds');
  const [fixedDepositsTitle, setFixedDepositsTitle] = useState<React.ReactNode>('Fixed Deposits');
  const [onlineInvestmentsTitle, setOnlineInvestmentsTitle] = useState<React.ReactNode>('Online Invests');
  const [sukukTitle, setSukukTitle] = useState<React.ReactNode>('Sukuk Funds');
  const [dseTrackerTitle, setDseTrackerTitle] = useState<React.ReactNode>('DSE Tracker');

  // ─── Global sync state ────────────────────────────────────────────────────
  type SyncStatus = 'idle' | 'syncing' | 'success' | 'error';
  const [syncStatus, setSyncStatus] = useState<SyncStatus>('idle');
  const [lastSyncedAt, setLastSyncedAt] = useState<number>(() => {
    // Use the most recent lastSyncedAt across all modules as the global "last synced"
    const modules = ['onlineInvestments', 'sukuk', 'mutualFunds', 'fixedDeposits', 'incomeExpense', 'dse'] as const;
    const times = modules.map(m => {
      try {
        const raw = localStorage.getItem(`syncmeta_${m}`);
        return raw ? JSON.parse(raw).lastSyncedAt ?? 0 : 0;
      } catch { return 0; }
    });
    return Math.max(...times.filter(t => t > 0)) || 0;
  });
  const [syncSummary, setSyncSummary] = useState<string>('');

  // Listen to instant sync events from any module
  useEffect(() => {
    let timer: ReturnType<typeof setTimeout>;
    const handleSyncStatus = (e: Event) => {
      const detail = (e as CustomEvent).detail;
      if (!detail) return;
      if (detail.status === 'syncing') {
        setSyncStatus('syncing');
        setSyncSummary('Syncing with Google Sheets...');
      } else if (detail.status === 'success') {
        setSyncStatus('success');
        setSyncSummary(detail.message || 'Saved to Google Sheets.');
        if (detail.lastSyncedAt) {
          setLastSyncedAt(detail.lastSyncedAt);
        }
        clearTimeout(timer);
        timer = setTimeout(() => {
          setSyncStatus('idle');
          setSyncSummary('');
        }, 3000);
      } else if (detail.status === 'error') {
        setSyncStatus('error');
        setSyncSummary(detail.message || 'Google Sheets sync failed.');
        clearTimeout(timer);
        timer = setTimeout(() => {
          setSyncStatus('idle');
          setSyncSummary('');
        }, 4000);
      }
    };

    window.addEventListener('app-sync-status', handleSyncStatus);
    return () => {
      window.removeEventListener('app-sync-status', handleSyncStatus);
      clearTimeout(timer);
    };
  }, []);

const dashboardBackupRef = useRef<HTMLInputElement>(null);
  const [dashboardPendingRestore, setDashboardPendingRestore] = useState<AppState | null>(null);
  const [dashboardNotification, setDashboardNotification] = useState<{ type: 'success' | 'error', message: string } | null>(null);

  const showDashboardNotification = (type: 'success' | 'error', message: string) => {
    setDashboardNotification({ type, message });
    setTimeout(() => setDashboardNotification(null), 5000);
  };

  const handleDashboardBackup = () => {
    try {
      const dataStr = JSON.stringify(state, null, 2);
      const dataBlob = new Blob([dataStr], { type: 'application/json' });
      const url = URL.createObjectURL(dataBlob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `fintrack_pro_backup_${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      showDashboardNotification('success', 'Backup created and downloaded successfully!');
    } catch (error) {
      showDashboardNotification('error', 'Failed to create backup. Please try again.');
    }
  };

const handleDashboardRestoreFile = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const importedState = JSON.parse(content);

        if (!importedState || typeof importedState !== 'object') {
          throw new Error('Invalid backup file format.');
        }

        const coreKeys = ['onlineInvestments', 'sukuks', 'cashBalance'];
        const hasCoreKeys = coreKeys.some(key => key in importedState);
        if (!hasCoreKeys) {
          throw new Error('This file does not appear to be a valid FinTrack Pro backup.');
        }

        setDashboardPendingRestore(importedState);
      } catch (error) {
        showDashboardNotification('error', `Failed to restore data: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    };
    reader.onerror = () => {
      showDashboardNotification('error', 'Failed to read the file. Please try again.');
    };
    reader.readAsText(file);
    event.target.value = '';
  };

  const confirmDashboardRestore = () => {
    if (dashboardPendingRestore) {
      updateState(prev => ({
        ...prev,
        ...dashboardPendingRestore,
        isLocked: true,
      }));
      setDashboardPendingRestore(null);
      showDashboardNotification('success', 'Data restored successfully! The application is now locked for security.');
    }
  };



  const handleSyncAll = useCallback(async () => {
    if (syncStatus === 'syncing') return;
    setSyncStatus('syncing');
    setSyncSummary('Syncing with Google Sheets...');

    try {
      const demoIds = ['1', '2', 'completed-1', 'completed-2', 'mf1', 'mf2', 'fdr1', 'fdr2', 's1', 's2', 'ol1', 'ol2'];

      const { pushed, pulled, failed } = await syncAllModules(
        () => {
          const cur = latestStateRef.current;
          let dseData: any[] = [];
          try {
            const raw = localStorage.getItem('sheet_cache_dse');
            if (raw) dseData = JSON.parse(raw);
          } catch {}

          return {
            onlineInvestments: cur.onlineInvestments,
            sukuk:             cur.sukuks,
            mutualFunds:       cur.mutualFunds,
            fixedDeposits:     cur.fdrs,
            incomeExpense:     serializeIncomeExpense(cur),
            dse:               dseData,
          };
        },
        (module, data) => {
          if (data === null) return;
          // Cloud was newer — apply to local state without re-marking dirty
          if (module === 'onlineInvestments') {
            const filtered = data.filter((item: any) => !demoIds.includes(item.id));
            updateState(s => ({ ...s, onlineInvestments: filtered }), []);
          } else if (module === 'sukuk') {
            const filtered = data.filter((item: any) => !demoIds.includes(item.id));
            updateState(s => ({ ...s, sukuks: filtered }), []);
          } else if (module === 'mutualFunds') {
            const filtered = data.filter((item: any) => !demoIds.includes(item.id));
            updateState(s => ({ ...s, mutualFunds: filtered }), []);
          } else if (module === 'fixedDeposits') {
            const filtered = data.filter((item: any) => !demoIds.includes(item.id));
            updateState(s => ({ ...s, fdrs: filtered }), []);
          } else if (module === 'incomeExpense') {
            const parsed = deserializeIncomeExpense(data);
            if (parsed) {
              suppressNextIncomeExpenseSyncRef.current = true;
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: parsed.transactions || [],
                incomeExpenseAccounts: parsed.accounts && parsed.accounts.length > 0 ? parsed.accounts : s.incomeExpenseAccounts,
                incomeExpenseCategories: parsed.categories && parsed.categories.length > 0 ? parsed.categories : s.incomeExpenseCategories,
                conversionRates: parsed.conversionRates || s.conversionRates,
              }), []);
            }
          } else if (module === 'dse') {
            if (Array.isArray(data) && data.length > 0) {
              const settingsRow = data.find((r: any) => r && r.type === '__SETTINGS__');
              if (settingsRow?.settings) {
                localStorage.setItem('dse_settings', JSON.stringify(settingsRow.settings));
              }
              const cleanTransactions = data.filter((r: any) => r && r.type !== '__SETTINGS__');
              localStorage.setItem('sheet_cache_dse', JSON.stringify(cleanTransactions));
              setDseTransactions(cleanTransactions);
              window.dispatchEvent(new Event('storage'));
            }
          }
        }
      );

      const now = Date.now();
      setLastSyncedAt(now);

      const parts: string[] = [];
      if (pushed.length) parts.push(`↑ Pushed ${pushed.length}`);
      if (pulled.length) parts.push(`↓ Pulled ${pulled.length}`);
      if (failed.length) parts.push(`✗ Failed ${failed.length}`);
      setSyncSummary(parts.join('  ·  ') || 'Synced with Google Sheets');
      setSyncStatus(failed.length > 0 ? 'error' : 'success');
    } catch {
      setSyncSummary('Sync failed — check connection');
      setSyncStatus('error');
    }

    // Reset to idle after 4 seconds
    setTimeout(() => {
      setSyncStatus('idle');
      setSyncSummary('');
    }, 4000);
  }, [syncStatus, updateState]);

  // ─── Cleanup: Remove known demo data if present ──────────────────────────────
  useEffect(() => {
    const demoIds = ['1', '2', 'completed-1', 'completed-2', 'mf1', 'mf2', 'fdr1', 'fdr2', 's1', 's2', 'ol1', 'ol2'];
    updateState(s => {
      let changed = false;
      const newState = { ...s };

      const filter = (arr: any[]) => {
        const filtered = (arr || []).filter(item => !demoIds.includes(item.id));
        if (filtered.length !== arr?.length) changed = true;
        return filtered;
      };

      newState.onlineInvestments = filter(s.onlineInvestments);
      newState.mutualFunds = filter(s.mutualFunds);
      newState.fdrs = filter(s.fdrs);
      newState.sukuks = filter(s.sukuks);

      return changed ? newState : s;
    }, []);
  }, []);

  // ─── DSE transactions from localStorage ───────────────────────────────────
  const [dseTransactions, setDseTransactions] = useState<any[]>(() => {
    try {
      const cached = localStorage.getItem(DSE_CACHE_KEY);
      return cached ? JSON.parse(cached) : [];
    } catch { return []; }
  });

  useEffect(() => {
    try {
      const cached = localStorage.getItem(DSE_CACHE_KEY);
      if (cached) setDseTransactions(JSON.parse(cached));
    } catch {}
  }, [activeModule]);



useEffect(() => {
  const handleStorageChange = () => {
    try {
      const cached = localStorage.getItem(DSE_CACHE_KEY);
      // Only update if we actually got data — never reset to [] on failure
      if (cached) {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setDseTransactions(parsed);
        }
      }
    } catch {}
  };
  window.addEventListener('storage', handleStorageChange);
  // Poll every 5s but only update if data is non-empty
  const interval = setInterval(handleStorageChange, 5000);
  return () => {
    window.removeEventListener('storage', handleStorageChange);
    clearInterval(interval);
  };
}, [DSE_CACHE_KEY]);

  // ─── 1. Startup Synchronization & Close Flush ─────────────────────────────
  useEffect(() => {
    // When the app loads for the first time in any session:
    // Local data is loaded immediately from localStorage.
    // Simultaneously, compare local data with Google Sheets data and pull or push based on which has been updated most recently.
    handleSyncAll();

    // Just at the time of closing: check and push the latest updates if dirty
    flushOnUnload(() => {
      const cur = latestStateRef.current;
      let dseData: any[] = [];
      try {
        const raw = localStorage.getItem('sheet_cache_dse');
        if (raw) dseData = JSON.parse(raw);
      } catch {}

      return {
        onlineInvestments: cur.onlineInvestments,
        sukuk:             cur.sukuks,
        mutualFunds:       cur.mutualFunds,
        fixedDeposits:     cur.fdrs,
        incomeExpense:     serializeIncomeExpense(cur),
        dse:               dseData,
      };
    });
  }, []);

  // ─── 2. Auto-sync every 5 minutes: check and push/pull latest update ───────
  useEffect(() => {
    const FIVE_MINUTES = 5 * 60 * 1000;
    const interval = setInterval(() => {
      handleSyncAll();
    }, FIVE_MINUTES);

    return () => clearInterval(interval);
  }, [handleSyncAll]);





  const navigateHistoryCustomMonth = (offset: number) => {
    const currentStart = new Date(historyCustomDates.start);
    const nextMonth = new Date(currentStart.getFullYear(), currentStart.getMonth() + offset, 1);
    const now = new Date();
    const isCurrentMonth =
      nextMonth.getFullYear() === now.getFullYear() &&
      nextMonth.getMonth() === now.getMonth();

    setHistoryCustomDates({
      start: getFirstOfMonth(nextMonth),
      end: isCurrentMonth ? getTodayStr() : getLastOfMonth(nextMonth),
    });
  };

  const navigateHistoryThisMonth = (offset: number) => {
    setHistoryThisMonthDate(prev => new Date(prev.getFullYear(), prev.getMonth() + offset, 1));
  };

  const handleRangeChange = (newRange: 'this' | 'fiscal' | 'custom') => {
    if (newRange === 'custom') {
      setHistoryCustomDates({
        start: '2023-01-01',
        end: getTodayStr(),
      });
    }
    setHistoryRange(newRange);
    setIsRangeMenuOpen(false);
  };

  const formatHistoryThisMonthLabel = (d: Date) => {
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return `${months[d.getMonth()]} ${d.getFullYear()}`;
  };

  const onlineInvestments = state.onlineInvestments;
  const sukuks = state.sukuks;

  // ─── Shared date range resolver ────────────────────────────────────────────
  const rangeDates = useMemo(() => {
    if (historyRange === 'this') {
      return {
        start: new Date(
          historyThisMonthDate.getFullYear(),
          historyThisMonthDate.getMonth(),
          1
        ),
        end: new Date(
          historyThisMonthDate.getFullYear(),
          historyThisMonthDate.getMonth() + 1,
          0,
          23, 59, 59, 999
        )
      };
    }

    if (historyRange === 'fiscal') {
      return {
        start: new Date(historyFiscalStartYear, 6, 1),
        end: new Date(historyFiscalStartYear + 1, 5, 30, 23, 59, 59, 999)
      };
    }

    const start = historyCustomDates.start
      ? new Date(historyCustomDates.start)
      : new Date(0);
    const end = historyCustomDates.end
      ? new Date(historyCustomDates.end)
      : new Date();

    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);

    return { start, end };
  }, [
    historyRange,
    historyCustomDates,
    historyThisMonthDate,
    historyFiscalStartYear
  ]);

  const { startStr, endStr } = useMemo(() => {
    return {
      startStr: toDateStr(rangeDates.start),
      endStr: toDateStr(rangeDates.end),
    };
  }, [rangeDates]);

  // ─── "Same date last year" range for year-over-year comparison ────────────
  const { lyStartStr, lyEndStr } = useMemo(() => {
    // Shift both start and end back by exactly 1 year
    const shiftYear = (s: string, delta: number): string => {
      if (!s || s.startsWith('0000') || s.startsWith('9999')) return s;
      const d = new Date(s);
      d.setFullYear(d.getFullYear() + delta);
      return toDateStr(d);
    };
    return {
      lyStartStr: shiftYear(startStr, -1),
      lyEndStr: shiftYear(endStr, -1),
    };
  }, [startStr, endStr]);

  // ─── Core stats computation (reusable for both current and last-year) ──────
  const globalUsdRate = state.conversionRates?.USD_to_BDT ?? 120;
  const globalLydRate = state.conversionRates?.LYD_to_BDT ?? 20;

  const computeStats = (
    effectiveStartStr: string,
    effectiveEndStr: string,
    dseTxns: any[],
  ) => {
    const inRange = (date: string) => date >= effectiveStartStr && date <= effectiveEndStr;
    const beforeEnd = (date: string) => date <= effectiveEndStr;

    const toStr2 = (d: Date) => toDateStr(d);

    const getOnlineStatus = (inv: any): string => {
      const today = toDateStr(new Date());
      if (inv.installments && inv.installments.length > 0) {
        if (inv.installments.every((i: any) => i.isPaid)) return 'Completed';
        if (inv.installments.some((i: any) => !i.isPaid && i.date < today)) return 'Delayed';
        return 'Active';
      }
      const expectedTotal = inv.estimatedReturn || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100));
      if (inv.totalRepaid >= expectedTotal * 0.99) return 'Completed';
      if (inv.maturityDate < today) return 'Delayed';
      return 'Active';
    };

    // ── MUTUAL FUNDS ──
    const mfStats = state.mutualFunds.map(f => {
      const cumulativeDeposit = f.transactions.reduce((sum, t) => {
        if (!beforeEnd(t.date)) return sum;
        if ((t.type === 'Buy' || t.type === 'Dividend') && !t.isDividend) return sum + (t.sipAmount || 0);
        return sum;
      }, 0);
      const cumulativeDividends = f.transactions.reduce((sum, t) => {
        if (!beforeEnd(t.date)) return sum;
        if (t.isDividend || t.type === 'Dividend') return sum + (t.sipAmount || 0);
        return sum;
      }, 0);
      const cumulativeSellAmount = f.transactions.reduce((sum, t) => {
        if (!beforeEnd(t.date)) return sum;
        if (t.type === 'Sell') return sum + t.amount;
        return sum;
      }, 0);
      const cumulativeWithdrawn = f.transactions.reduce((sum, t) => {
        if (!beforeEnd(t.date)) return sum;
        if (t.type === 'Withdrawal') return sum + t.amount;
        return sum;
      }, 0);
      const currentHolding = cumulativeDeposit + cumulativeDividends + cumulativeSellAmount - cumulativeWithdrawn;
      const activeInvestment = cumulativeDeposit - cumulativeWithdrawn;
      const rangeDeposit = f.transactions.reduce((sum, t) => {
        if (!inRange(t.date)) return sum;
        if ((t.type === 'Buy' || t.type === 'Dividend') && !t.isDividend) return sum + (t.sipAmount || 0);
        return sum;
      }, 0);
      const rangeDividends = f.transactions.reduce((sum, t) => {
        if (!inRange(t.date)) return sum;
        if (t.isDividend || t.type === 'Dividend') return sum + (t.sipAmount || 0);
        return sum;
      }, 0);
      return { currentHolding, activeInvestment, rangeDeposit, rangeDividends };
    });

    const mfCurrentHolding = mfStats.reduce((sum, s) => sum + s.currentHolding, 0);
    const mfActiveInvested = mfStats.reduce((sum, s) => sum + s.activeInvestment, 0);
    const mfTotalInvested  = mfStats.reduce((sum, s) => sum + s.rangeDeposit, 0);
    const mfTotalProfit    = mfStats.reduce((sum, s) => sum + s.rangeDividends, 0);

    // ── FIXED DEPOSITS ──
    let fdrInvestedTotal = 0, fdrActiveTotal = 0, fdrHoldingTotal = 0, fdrProfitRangeTotal = 0;
    state.fdrs.forEach(f => {
      const rate = f.currency === 'USD' ? globalUsdRate : (f.currency === 'LYD' ? globalLydRate : 1);
      if (inRange(f.investmentDate)) fdrInvestedTotal += f.principal * rate;
      const openedBeforeEnd = beforeEnd(f.investmentDate);
      const isClosedByEnd = f.status === 'Closed' && f.closingDate && f.closingDate <= effectiveEndStr;
      if (openedBeforeEnd && !isClosedByEnd) fdrActiveTotal += f.principal * rate;
      if (openedBeforeEnd) {
        let fAddedProfitAsOfEnd = 0, fChargeAsOfEnd = 0, fProfitRange = 0, fChargeRange = 0;
        f.transactions.forEach(t => {
          if (!beforeEnd(t.date)) return;
          if (t.type === 'Profit') {
            if (t.handling === 'Added') fAddedProfitAsOfEnd += t.amount;
            if (inRange(t.date)) fProfitRange += t.amount;
          } else if (t.type === 'Charge') {
            fChargeAsOfEnd += t.amount;
            if (inRange(t.date)) fChargeRange += t.amount;
          }
        });
        const computedBal = f.principal + fAddedProfitAsOfEnd - fChargeAsOfEnd;
        const bal = (isClosedByEnd && f.withdrawBalance != null)
          ? Math.max(0, computedBal - f.withdrawBalance)
          : computedBal;
        if (bal > 0) fdrHoldingTotal += bal * rate;
        fdrProfitRangeTotal += (fProfitRange - fChargeRange) * rate;
      }
    });

    const fdrCurrentHolding = fdrHoldingTotal;
    const fdrActiveInvested = fdrActiveTotal;
    const fdrTotalInvested  = fdrInvestedTotal;
    const fdrTotalProfit    = fdrProfitRangeTotal;

    // ── ONLINE INVESTMENTS ──
    const olStats = state.onlineInvestments.map(inv => {
      const rate = inv.currency === 'USD' ? globalUsdRate : (inv.currency === 'LYD' ? globalLydRate : 1);
      let isHolding = false;
      if (beforeEnd(inv.investmentDate)) {
        const status = getOnlineStatus(inv);
        if (status !== 'Completed') {
          isHolding = true;
        } else {
          const lastInst = inv.installments?.[inv.installments.length - 1];
          const completionDate = lastInst ? (lastInst.actualDate || lastInst.date) : inv.maturityDate;
          if (completionDate > effectiveEndStr) isHolding = true;
        }
      }
      let profitRange = 0;
      if (getOnlineStatus(inv) === 'Completed') {
        const lastInst = inv.installments?.[inv.installments.length - 1];
        const completionDate = lastInst ? (lastInst.actualDate || lastInst.date) : (inv.actualMaturityDate || inv.maturityDate);
        if (inRange(completionDate)) {
          profitRange = (inv.actualProfit !== undefined ? inv.actualProfit : (inv.totalRepaid - inv.amount)) * rate;
        }
      }
      return {
        holding: isHolding ? inv.amount * rate : 0,
        investedRange: inRange(inv.investmentDate) ? inv.amount * rate : 0,
        profitRange
      };
    });

    const olCurrentHolding = olStats.reduce((sum, s) => sum + s.holding, 0);
    const olActiveInvested = olCurrentHolding;
    const olTotalInvested  = olStats.reduce((sum, s) => sum + s.investedRange, 0);
    const olTotalProfit    = olStats.reduce((sum, s) => sum + s.profitRange, 0);

    // ── SUKUK ──
    const skStats = state.sukuks.map(inv => {
      let holding = 0;
      if (beforeEnd(inv.issueDate)) {
        const maturityDate = new Date(inv.issueDate);
        maturityDate.setFullYear(maturityDate.getFullYear() + inv.durationYears);
        const maturityStr = toStr2(maturityDate);
        const closingStr = (inv as any).closingDate;
        const effectiveClosingStr = closingStr || maturityStr;
        if (effectiveClosingStr >= effectiveEndStr) {
          holding = inv.principalAmount;
        } else {
          const withdrawBalance = (inv as any).withdrawBalance;
          if (withdrawBalance === undefined || withdrawBalance < inv.principalAmount) {
            holding = inv.principalAmount - (withdrawBalance || 0);
          }
        }
      }
      const profitRange = (inv.installments || []).reduce((sum, inst) => {
        const payDate = inst.actualDate || inst.date;
        if (inst.isPaid && inRange(payDate)) return sum + (inst.actualAmount || inst.amount);
        return sum;
      }, 0);
      return {
        holding,
        active: holding,
        investedRange: inRange(inv.issueDate) ? inv.principalAmount : 0,
        profitRange
      };
    });

    const skCurrentHolding = skStats.reduce((sum, s) => sum + s.holding, 0);
    const skActiveInvested = skStats.reduce((sum, s) => sum + s.active, 0);
    const skTotalInvested  = skStats.reduce((sum, s) => sum + s.investedRange, 0);
    const skTotalProfit    = skStats.reduce((sum, s) => sum + s.profitRange, 0);

    // ── DSE ──
    const typePriority: Record<string, number> = { Deposit: 0, Buy: 1, Dividend: 2, Charge: 3, Sell: 4, Withdrawal: 5 };
    const dseSorted = [...dseTxns].sort((a, b) => {
      const dateDiff = new Date(a.date).getTime() - new Date(b.date).getTime();
      if (dateDiff !== 0) return dateDiff;
      return (typePriority[a.type] ?? 9) - (typePriority[b.type] ?? 9);
    });

    let dT_Dep = 0, dT_Wd = 0, dT_Div = 0, dT_Chg = 0, dT_PnL = 0;
    let dR_Dep = 0, dR_Div = 0, dR_Chg = 0, dR_PnL = 0;
    const dHoldings: Record<string, { qty: number; totalCost: number }> = {};

    dseSorted.forEach(t => {
      const key = `${t.portfolio}|${t.ticker}`;
      if (t.type === 'Deposit') {
        if (beforeEnd(t.date)) dT_Dep += Math.abs(t.total);
        if (inRange(t.date)) dR_Dep += Math.abs(t.total);
      } else if (t.type === 'Withdrawal') {
        if (beforeEnd(t.date)) dT_Wd += Math.abs(t.total);
      } else if (t.type === 'Charge') {
        if (beforeEnd(t.date)) dT_Chg += Math.abs(t.total);
        if (inRange(t.date)) dR_Chg += Math.abs(t.total);
      } else if (t.type === 'Dividend') {
        if (beforeEnd(t.date)) dT_Div += Math.abs(t.total);
        if (inRange(t.date)) dR_Div += Math.abs(t.total);
      } else if (t.type === 'Buy') {
        if (!beforeEnd(t.date)) return;
        if (!dHoldings[key]) dHoldings[key] = { qty: 0, totalCost: 0 };
        dHoldings[key].qty += t.qty;
        dHoldings[key].totalCost += t.total;
      } else if (t.type === 'Sell') {
        if (!beforeEnd(t.date)) return;
        if (!dHoldings[key]) dHoldings[key] = { qty: 0, totalCost: 0 };
        const h = dHoldings[key];
        if (h.qty > 0) {
          const avgCost = h.totalCost / h.qty;
          const costOfSold = Math.min(t.qty, h.qty) * avgCost;
          const pnl = t.total - costOfSold;
          dT_PnL += pnl;
          if (inRange(t.date)) dR_PnL += pnl;
          h.qty = Math.max(0, h.qty - t.qty);
          h.totalCost = Math.max(0, h.totalCost - costOfSold);
        }
      }
    });

    const dStockHoldingCost = Object.values(dHoldings).reduce((sum, h) => sum + h.totalCost, 0);
    const dCashBalance = dT_Dep - dT_Wd - dStockHoldingCost + dT_PnL + dT_Div - dT_Chg;
    const dseCurrentHolding = dStockHoldingCost + dCashBalance;
    const dseActiveInvested = dT_Dep - dT_Wd;
    const dseTotalInvested  = dR_Dep;
    const dseTotalProfit    = dR_PnL + dR_Div - dR_Chg;

    // ── INCOME & EXPENSE / ACCOUNTS BALANCE IN BDT ──
    const accounts = state.incomeExpenseAccounts || [];
    const transactions = state.incomeExpenseTransactions || [];
    const initialBalanceTransactions = accounts
      .filter(acc => !acc.isParent && acc.initialBalance && acc.initialBalance > 0)
      .map(acc => ({
        id: `init-bal-${acc.id}`,
        date: acc.initialDate || '2026-05-01',
        amount: acc.initialBalance,
        type: 'Income' as const,
        accountId: acc.id,
        category: 'Initial Balance',
        subCategory: 'Starting Balance',
        description: `Initial Balance - ${acc.name}`
      }));
    const allAccountTxs = [...transactions, ...initialBalanceTransactions];
    const USD_rate = globalUsdRate;
    const LYD_rate = globalLydRate;

    const balances: Record<string, number> = {};
    accounts.forEach(acc => {
      const txs = allAccountTxs.filter(t => beforeEnd(t.date) && (t.accountId === acc.id || t.toAccountId === acc.id));
      const income = txs.filter(t => t.type === 'Income' && t.accountId === acc.id).reduce((s, t) => s + t.amount, 0);
      const expense = txs.filter(t => t.type === 'Expense' && t.accountId === acc.id).reduce((s, t) => s + t.amount, 0);
      let transOut = 0;
      let transIn = 0;
      txs.forEach(t => {
        if (t.type === 'Transfer') {
          if (t.transferType) {
            if (t.accountId === acc.id) {
              if (t.transferType === 'from') transOut += t.amount;
              else if (t.transferType === 'to') transIn += t.amount;
            }
          } else {
            if (t.accountId === acc.id) transOut += t.amount;
            if (t.toAccountId === acc.id) transIn += (t.toAmount !== undefined ? t.toAmount : t.amount);
          }
        }
      });
      const loanInflow = txs.filter(t => (t.type === 'Loan' || t.category === 'Loan') && t.accountId === acc.id && (t.subCategory === 'Borrowed' || t.subCategory === 'Received Lent Money')).reduce((s, t) => s + t.amount, 0);
      const loanOutflow = txs.filter(t => (t.type === 'Loan' || t.category === 'Loan') && t.accountId === acc.id && (t.subCategory === 'Lent' || t.subCategory === 'Repaid Borrowed Money')).reduce((s, t) => s + t.amount, 0);
      balances[acc.id] = income - expense - transOut + transIn + loanInflow - loanOutflow;
    });

    accounts.forEach(acc => {
      if (acc.isParent) {
        const children = accounts.filter(a => a.parentId === acc.id);
        const childrenSum = children.reduce((s, c) => s + (balances[c.id] || 0), 0);
        balances[acc.id] = (balances[acc.id] || 0) + childrenSum;
      }
    });

    const totalBDT = accounts
      .filter(a => a.currency === 'BDT' && !a.parentId)
      .reduce((sum, a) => sum + (balances[a.id] || 0), 0);
    const totalLYD = accounts
      .filter(a => a.currency === 'LYD' && !a.parentId)
      .reduce((sum, a) => sum + (balances[a.id] || 0), 0);
    const totalUSD = accounts
      .filter(a => a.currency === 'USD' && !a.parentId)
      .reduce((sum, a) => sum + (balances[a.id] || 0), 0);

    const loanStatsByCurr = {
      BDT: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
      LYD: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
      USD: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
    };

    allAccountTxs.forEach(t => {
      if (!beforeEnd(t.date)) return;
      if (t.type !== 'Loan' && t.category !== 'Loan') return;
      const acc = accounts.find(a => a.id === t.accountId);
      if (!acc) return;
      const currency = acc.currency as 'BDT' | 'LYD' | 'USD';
      if (!loanStatsByCurr[currency]) return;
      if (t.subCategory === 'Borrowed') loanStatsByCurr[currency].borrowed += t.amount;
      else if (t.subCategory === 'Repaid Borrowed Money') loanStatsByCurr[currency].repaid += t.amount;
      else if (t.subCategory === 'Lent') loanStatsByCurr[currency].lent += t.amount;
      else if (t.subCategory === 'Received Lent Money') loanStatsByCurr[currency].received += t.amount;
    });

    const totalCurrentBorrowedBDT = Math.max(0, loanStatsByCurr.BDT.borrowed - loanStatsByCurr.BDT.repaid)
      + Math.max(0, loanStatsByCurr.LYD.borrowed - loanStatsByCurr.LYD.repaid) * LYD_rate
      + Math.max(0, loanStatsByCurr.USD.borrowed - loanStatsByCurr.USD.repaid) * USD_rate;

    const totalCurrentLentBDT = Math.max(0, loanStatsByCurr.BDT.lent - loanStatsByCurr.BDT.received)
      + Math.max(0, loanStatsByCurr.LYD.lent - loanStatsByCurr.LYD.received) * LYD_rate
      + Math.max(0, loanStatsByCurr.USD.lent - loanStatsByCurr.USD.received) * USD_rate;

    const rawSum = totalBDT + (totalLYD * LYD_rate) + (totalUSD * USD_rate);
    const totalBalanceInBDT = rawSum - totalCurrentBorrowedBDT + totalCurrentLentBDT;

    const totalInvestmentHolding = dseCurrentHolding + mfCurrentHolding + fdrCurrentHolding + olCurrentHolding + skCurrentHolding;
    const totalNetWorth         = totalBalanceInBDT;
    const totalActiveInvestment = dseActiveInvested + mfActiveInvested + fdrActiveInvested + olActiveInvested + skCurrentHolding;
    const totalInvested         = dseTotalInvested  + mfTotalInvested  + fdrTotalInvested  + olTotalInvested  + skTotalInvested;
    const totalProfit           = dseTotalProfit    + mfTotalProfit    + fdrTotalProfit    + olTotalProfit    + skTotalProfit;

    const isInvestmentAcc = (acc: any, allAccs: any[]) => {
      if (!acc) return false;
      if (acc.id === 'bdt-investments') return true;
      if (acc.parentId) {
        const parent的的 = allAccs.find(a => a.id === acc.parentId);
        if (parent的的 && (parent的的.id === 'bdt-investments' || parent的的.name.toLowerCase().includes('invest'))) return true;
      }
      const id = (acc.id || '').toLowerCase();
      const name = (acc.name || '').toLowerCase();
      return (
        id === 'bdt-investments' ||
        id.includes('investment') ||
        id.includes('bo-account') ||
        id.includes('mutual-fund') ||
        id.includes('sukuk') ||
        id.includes('fdr') ||
        id.includes('online-invest') ||
        name === 'investments' ||
        name.includes('investment') ||
        name.includes('bo account') ||
        name.includes('online investment') ||
        name.includes('mutual fund') ||
        name.includes('sukuk') ||
        name.includes('fdr')
      );
    };

    const totalExpenseBDT = allAccountTxs
      .filter(t => beforeEnd(t.date) && t.type === 'Expense')
      .reduce((sum, t) => {
        const acc = accounts.find(a => a.id === t.accountId);
        const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
        return sum + (t.amount * mult);
      }, 0);

    const totalIncomeBDT = allAccountTxs
      .filter(t => beforeEnd(t.date) && t.type === 'Income')
      .reduce((sum, t) => {
        const acc主管 = accounts.find(a => a.id === t.accountId);
        const mult = acc主管?.currency === 'USD' ? USD_rate : acc主管?.currency === 'LYD' ? LYD_rate : 1;
        return sum + (t.amount * mult);
      }, 0);

    const totalInvestmentAccBDT = accounts
      .filter(acc => !acc.isParent && isInvestmentAcc(acc, accounts))
      .reduce((sum, acc) => {
        const bal = balances[acc.id] || 0;
        const mult = acc.currency === 'USD' ? USD_rate : acc.currency === 'LYD' ? LYD_rate : 1;
        return sum + (bal * mult);
      }, 0);

    const totalSavingBDT翼 = Math.max(0, totalIncomeBDT - totalExpenseBDT - totalInvestmentAccBDT);

    const assetAllocationData = [
      { name: 'Expense', value: Math.max(0, totalExpenseBDT), color: '#ef4444' },
      { name: 'Saving', value: totalSavingBDT翼, color: '#2dd4bf' },
      { name: 'Investment', value: Math.max(0, totalInvestmentAccBDT), color: '#3b82f6' },
    ];

    return {
      totalNetWorth,
      totalBalanceInBDT,
      totalInvestmentHolding,
      totalActiveInvestment,
      totalInvested,
      totalProfit,
      assetAllocationData,
      dseCurrentHolding,
      mfCurrentHolding,
      fdrCurrentHolding,
      olCurrentHolding,
      skCurrentHolding,
      dseTotalInvested,
      mfTotalInvested,
      fdrTotalInvested,
      olTotalInvested,
      skTotalInvested,
      dseTotalProfit,
      mfTotalProfit,
      fdrTotalProfit,
      olTotalProfit,
      skTotalProfit,
      totalCurrentBorrowedBDT,
      totalCurrentLentBDT,
    };
  };

  const stats = useMemo(() => {
    const getOnlineStatus = (inv: any): string => {
      const today = toDateStr(new Date());
      if (inv.installments && inv.installments.length > 0) {
        if (inv.installments.every((i: any) => i.isPaid)) return 'Completed';
        if (inv.installments.some((i: any) => !i.isPaid && i.date < today)) return 'Delayed';
        return 'Active';
      }
      const expectedTotal = inv.estimatedReturn || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100));
      if (inv.totalRepaid >= expectedTotal * 0.99) return 'Completed';
      if (inv.maturityDate < today) return 'Delayed';
      return 'Active';
    };

    // ── Current period stats ──
    const current = computeStats(startStr, endStr, dseTransactions);

    // ── Total (all-time, not scoped to the selected range) loan stats ──
    const totalLoanStats = computeStats(startStr, getTodayStr(), dseTransactions);

    // ── Last year stats ──
    const lastYear = computeStats(lyStartStr, lyEndStr, dseTransactions);

    // ── YoY % change ──
    const yoyPct = (curr: number, prev: number): number | null => {
      if (prev === 0) return null;
      return ((curr - prev) / Math.abs(prev)) * 100;
    };

    const netWorthYoy          = yoyPct(current.totalNetWorth, lastYear.totalNetWorth);
    const totalBalanceInBdtYoy = yoyPct(current.totalBalanceInBDT, lastYear.totalBalanceInBDT);
    const investmentHoldingYoy = yoyPct(current.totalInvestmentHolding, lastYear.totalInvestmentHolding);
    const activeInvYoy         = yoyPct(current.totalActiveInvestment, lastYear.totalActiveInvestment);
    const totalInvYoy          = yoyPct(current.totalInvested, lastYear.totalInvested);
    const totalProfitYoy       = yoyPct(current.totalProfit, lastYear.totalProfit);

    // ── Pie chart ──
    const pieData = [
      { name: 'DSE', value: current.dseCurrentHolding },
      { name: 'Mutual Funds', value: current.mfCurrentHolding },
      { name: 'Fixed Deposits', value: current.fdrCurrentHolding },
      { name: 'Online', value: current.olCurrentHolding },
      { name: 'Sukuk', value: current.skCurrentHolding },
    ].filter(item => item.value > 0);

    const typePriority: Record<string, number> = { Deposit: 0, Buy: 1, Dividend: 2, Charge: 3, Sell: 4, Withdrawal: 5 };
    const dseSorted = [...dseTransactions].sort((a, b) => {
      const dateDiff = new Date(a.date).getTime() - new Date(b.date).getTime();
      if (dateDiff !== 0) return dateDiff;
      return (typePriority[a.type] ?? 9) - (typePriority[b.type] ?? 9);
    });

    // ── Net Worth Trend (24 monthly snapshots) ────────────────────────────────
    // FIX: Compute raw series then normalize for chart display
    // Each series gets its own Y-axis domain so it occupies ~half the chart height.
    // We return both raw values and normalized [0,1] values for the chart to use.
// ── Net Worth Trend (24 monthly snapshots) ────────────────────────────────
    const rawTrendData: { name: string; [key: string]: number | string }[] = [];

    for (let i = 0; i < 24; i++) {
      const snap = new Date();
      snap.setDate(1);
      snap.setMonth(snap.getMonth() - (23 - i) + 1);
      snap.setDate(0);
      const snapStr = toDateStr(snap);
      const label = snap.toLocaleString('default', { month: 'short', year: '2-digit' });

      const mfSnap = state.mutualFunds.reduce((sum, f) => {
        return sum + f.transactions.reduce((tSum, t) => {
          if (t.date > snapStr) return tSum;
          if (t.type === 'Buy' || t.type === 'Dividend') return tSum + t.amount;
          if (t.type === 'Sell' || t.type === 'Withdrawal') return tSum - t.amount;
          return tSum;
        }, 0);
      }, 0);

      const olSnap = state.onlineInvestments.filter(inv => inv.investmentDate <= snapStr).reduce((sum, inv) => {
        const status = getOnlineStatus(inv);
        if (status === 'Completed') {
          const lastInst = inv.installments?.[inv.installments.length - 1];
          const completionDate = lastInst ? (lastInst.actualDate || lastInst.date) : inv.maturityDate;
          if (completionDate < snapStr) return sum;
        }
        return sum + inv.amount;
      }, 0);

      const skSnap = state.sukuks.filter(s => s.issueDate <= snapStr).reduce((sum, s) => {
        const maturityDate = new Date(s.issueDate);
        maturityDate.setFullYear(maturityDate.getFullYear() + s.durationYears);
        const maturityStr = toDateStr(maturityDate);
        const closingStr = (s as any).closingDate as string | undefined;
        const effectiveClosingStr = closingStr || maturityStr;
        if (effectiveClosingStr < snapStr) {
          const withdrawBalance = (s as any).withdrawBalance as number | undefined;
          if (withdrawBalance !== undefined && withdrawBalance >= s.principalAmount) return sum;
          const netPrincipal = withdrawBalance !== undefined ? s.principalAmount - withdrawBalance : s.principalAmount;
          return sum + netPrincipal;
        }
        return sum + s.principalAmount;
      }, 0);

      const fdrSnap = state.fdrs.reduce((sum, f) => {
        if (f.investmentDate > snapStr) return sum;
        const wasClosedBySnap = f.status === 'Closed' && f.closingDate && f.closingDate <= snapStr;
        const addedProfit = f.transactions.filter(t => t.type === 'Profit' && t.handling === 'Added' && t.date <= snapStr).reduce((p, t) => p + t.amount, 0);
        const charges = f.transactions.filter(t => t.type === 'Charge' && t.date <= snapStr).reduce((c, t) => c + t.amount, 0);
        const computedBal = f.principal + addedProfit - charges;
        const bal = wasClosedBySnap && f.withdrawBalance != null ? computedBal - f.withdrawBalance : computedBal;
        if (bal <= 0) return sum;
        const rate = f.currency === 'USD' ? globalUsdRate : (f.currency === 'LYD' ? globalLydRate : 1);
        return sum + (bal * rate);
      }, 0);

      const dseSnapMap: Record<string, { qty: number; totalCost: number }> = {};
      dseSorted.forEach((t: any) => {
        const key = t.ticker ? `${t.portfolio}|${t.ticker}` : null;
        if (!key || t.date > snapStr) return;
        if (t.type === 'Buy') {
          if (!dseSnapMap[key]) dseSnapMap[key] = { qty: 0, totalCost: 0 };
          dseSnapMap[key].qty += t.qty;
          dseSnapMap[key].totalCost += t.total;
        } else if (t.type === 'Sell') {
          if (!dseSnapMap[key]) dseSnapMap[key] = { qty: 0, totalCost: 0 };
          const h = dseSnapMap[key];
          if (h.qty > 0) {
            const avgCost = h.totalCost / h.qty;
            const costOfSold = t.qty * avgCost;
            h.qty = Math.max(0, h.qty - t.qty);
            h.totalCost = Math.max(0, h.totalCost - costOfSold);
          }
        }
      });
      const dseSnap = Object.values(dseSnapMap).reduce((sum, h) => sum + h.totalCost, 0);

      rawTrendData.push({
        name: label,
        'DSE Tracker': dseSnap,
        'Mutual Funds': mfSnap,
        'Fixed Deposits': fdrSnap,
        'Sukuk Funds': skSnap,
        'Online Invests': olSnap,
        'Total': mfSnap + fdrSnap + olSnap + skSnap + dseSnap,
      });
    }

    const seriesKeys = ['Total', 'DSE Tracker', 'Mutual Funds', 'Fixed Deposits', 'Sukuk Funds', 'Online Invests'];
    const seriesMetadata: Record<string, { min: number; max: number; range: number }> = {};
    seriesKeys.forEach(key => {
      const vals = rawTrendData.map(d => d[key] as number);
      const min = Math.min(...vals);
      const max = Math.max(...vals);
      seriesMetadata[key] = { min, max, range: max - min };
    });

    // ── Build cumulative profit snapshots per series ───────────────────────────
    const profitSnapshots: Record<string, number[]> = {
      'Total': [], 'DSE Tracker': [], 'Mutual Funds': [],
      'Fixed Deposits': [], 'Sukuk Funds': [], 'Online Invests': [],
    };

    for (let i = 0; i < 24; i++) {
      const snap = new Date();
      snap.setDate(1);
      snap.setMonth(snap.getMonth() - (23 - i) + 1);
      snap.setDate(0);
      const snapStr = toDateStr(snap);

      // ── MF: cumulative dividend income up to snapStr ──
      const mfProfit = state.mutualFunds.reduce((sum, f) =>
        sum + f.transactions.reduce((ts, t) => {
          if (t.date > snapStr) return ts;
          if (t.isDividend) return ts + (t.sipAmount || t.amount || 0);
          return ts;
        }, 0), 0);

      // ── FDR: cumulative profit-type transactions up to snapStr ──
      const fdrProfit = state.fdrs.reduce((sum, f) => {
        if (f.investmentDate > snapStr) return sum;
        const rate = f.currency === 'USD' ? globalUsdRate : (f.currency === 'LYD' ? globalLydRate : 1);
        return sum + f.transactions
          .filter(t => t.type === 'Profit' && t.date <= snapStr)
          .reduce((s, t) => s + t.amount * rate, 0);
      }, 0);

      // ── Sukuk: cumulative paid installments up to snapStr ──
      const skProfit = state.sukuks.reduce((sum, s) =>
        sum + (s.installments || []).reduce((ss, inst) => {
          const payDate = inst.actualDate || inst.date;
          if (inst.isPaid && payDate <= snapStr) return ss + (inst.actualAmount || inst.amount);
          return ss;
        }, 0), 0);

      // ── Online: cumulative paid installment profit up to snapStr ──
      const olProfit = state.onlineInvestments.reduce((sum, inv) => {
        if (inv.investmentDate > snapStr) return sum;
        const rate = inv.currency === 'USD' ? globalUsdRate : (inv.currency === 'LYD' ? globalLydRate : 1);
        const paidAmt = (inv.installments || []).reduce((ss, inst) => {
          const payDate = inst.actualDate || inst.date;
          if (inst.isPaid && payDate <= snapStr) return ss + (inst.actualAmount || inst.amount);
          return ss;
        }, 0);
        // Only count profit portion (paid beyond principal)
        return sum + Math.max(0, paidAmt - inv.amount * rate);
      }, 0);

      // ── DSE: cumulative dividends + realized P&L up to snapStr ──
      let dseDivSnap = 0, dsePnLSnap = 0;
      const dseHoldingsForPnL: Record<string, { qty: number; totalCost: number }> = {};
      dseSorted.forEach((t: any) => {
        if (t.date > snapStr) return;
        const key = `${t.portfolio}|${t.ticker}`;
        if (t.type === 'Dividend') {
          dseDivSnap += Math.abs(t.total);
        } else if (t.type === 'Charge') {
          dseDivSnap -= Math.abs(t.total); // charges reduce profit
        } else if (t.type === 'Buy') {
          if (!dseHoldingsForPnL[key]) dseHoldingsForPnL[key] = { qty: 0, totalCost: 0 };
          dseHoldingsForPnL[key].qty += t.qty;
          dseHoldingsForPnL[key].totalCost += t.total;
        } else if (t.type === 'Sell') {
          if (!dseHoldingsForPnL[key]) dseHoldingsForPnL[key] = { qty: 0, totalCost: 0 };
          const h = dseHoldingsForPnL[key];
          if (h.qty > 0) {
            const avgCost = h.totalCost / h.qty;
            const costOfSold = t.qty * avgCost;
            dsePnLSnap += t.total - costOfSold;
            h.qty = Math.max(0, h.qty - t.qty);
            h.totalCost = Math.max(0, h.totalCost - costOfSold);
          }
        }
      });
      const dseProfit = dseDivSnap + dsePnLSnap;

      const totalProfit = mfProfit + fdrProfit + skProfit + olProfit + dseProfit;

      profitSnapshots['Total'].push(totalProfit);
      profitSnapshots['DSE Tracker'].push(dseProfit);
      profitSnapshots['Mutual Funds'].push(mfProfit);
      profitSnapshots['Fixed Deposits'].push(fdrProfit);
      profitSnapshots['Sukuk Funds'].push(skProfit);
      profitSnapshots['Online Invests'].push(olProfit);
    }

    // ── Build netWorthTrends with profit field attached ────────────────────────
    const netWorthTrends: Record<string, { name: string; value: number; rawValue: number; profit: number }[]> = {};
    seriesKeys.forEach(key => {
      netWorthTrends[key] = rawTrendData.map((d, i) => {
        const rawValue = d[key] as number;
        return {
          name: d.name as string,
          value: rawValue,
          rawValue,
          profit: profitSnapshots[key][i] ?? 0,
        };
      });
    });

    // ── Accounts Trend (24 monthly snapshots) ──────────────────────────────────
    const accounts = state.incomeExpenseAccounts || [];
    const transactions = state.incomeExpenseTransactions || [];
    const initialBalanceTransactions = accounts
      .filter(acc => !acc.isParent && acc.initialBalance && acc.initialBalance > 0)
      .map(acc => ({
        id: `init-bal-${acc.id}`,
        date: acc.initialDate || '2026-05-01',
        amount: acc.initialBalance,
        type: 'Income' as const,
        accountId: acc.id,
        category: 'Initial Balance',
        subCategory: 'Starting Balance',
        description: `Initial Balance - ${acc.name}`
      }));
    const allAccountTxs = [...transactions, ...initialBalanceTransactions];

    const conversionRates = state.conversionRates || { USD_to_BDT: 120, LYD_to_BDT: 20 };
    const USD_rate = conversionRates.USD_to_BDT || 120;
    const LYD_rate = conversionRates.LYD_to_BDT || 20;

    const accountsTrendData: { name: string; 'Total (BDT)': number; 'Savings Income (BDT)': number; 'Savings Expense (BDT)': number; 'City Islamic': number; 'IBBL Savings': number; 'City FCY': number }[] = [];

    const cityAcc = accounts.find(a => a.id === 'bdt-city-islamic' || a.name.toLowerCase().includes('city islamic'));
    const ibblAcc = accounts.find(a => a.id === 'bdt-ibbl-savings' || a.name.toLowerCase().includes('ibbl savings'));
    const cityFcyAcc = accounts.find(a => a.id === 'usd-city-fcy' || a.id.includes('city-fcy') || a.name.toLowerCase().includes('city fcy'));

    for (let i = 0; i < 24; i++) {
      const snap = new Date();
      snap.setDate(1);
      snap.setMonth(snap.getMonth() - (23 - i) + 1);
      snap.setDate(0);
      const snapStr = toDateStr(snap);
      const label = snap.toLocaleString('default', { month: 'short', year: '2-digit' });

      const balances: Record<string, number> = {};
      accounts.forEach(acc => {
        const txs = allAccountTxs.filter(t => t.date <= snapStr && (t.accountId === acc.id || t.toAccountId === acc.id));
        const income = txs.filter(t => t.type === 'Income' && t.accountId === acc.id).reduce((s, t) => s + t.amount, 0);
        const expense = txs.filter(t => t.type === 'Expense' && t.accountId === acc.id).reduce((s, t) => s + t.amount, 0);
        let transOut = 0;
        let transIn = 0;
        txs.forEach(t => {
          if (t.type === 'Transfer') {
            if (t.transferType) {
              if (t.accountId === acc.id) {
                if (t.transferType === 'from') transOut += t.amount;
                else if (t.transferType === 'to') transIn += t.amount;
              }
            } else {
              if (t.accountId === acc.id) transOut += t.amount;
              if (t.toAccountId === acc.id) transIn += (t.toAmount !== undefined ? t.toAmount : t.amount);
            }
          }
        });
        const loanInflow = txs.filter(t => (t.type === 'Loan' || t.category === 'Loan') && t.accountId === acc.id && (t.subCategory === 'Borrowed' || t.subCategory === 'Received Lent Money')).reduce((s, t) => s + t.amount, 0);
        const loanOutflow = txs.filter(t => (t.type === 'Loan' || t.category === 'Loan') && t.accountId === acc.id && (t.subCategory === 'Lent' || t.subCategory === 'Repaid Borrowed Money')).reduce((s, t) => s + t.amount, 0);
        balances[acc.id] = income - expense - transOut + transIn + loanInflow - loanOutflow;
      });

      accounts.forEach(acc => {
        if (acc.isParent) {
          const children = accounts.filter(a => a.parentId === acc.id);
          const childrenSum = children.reduce((s, c) => s + (balances[c.id] || 0), 0);
          balances[acc.id] = (balances[acc.id] || 0) + childrenSum;
        }
      });

      const totalBDT = accounts
        .filter(a => a.currency === 'BDT' && !a.parentId)
        .reduce((sum, a) => sum + (balances[a.id] || 0), 0);
      const totalLYD = accounts
        .filter(a => a.currency === 'LYD' && !a.parentId)
        .reduce((sum, a) => sum + (balances[a.id] || 0), 0);
      const totalUSD = accounts
        .filter(a => a.currency === 'USD' && !a.parentId)
        .reduce((sum, a) => sum + (balances[a.id] || 0), 0);

      const loanStatsByCurr = {
        BDT: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
        LYD: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
        USD: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
      };

      allAccountTxs.forEach(t => {
        if (t.date > snapStr) return;
        if (t.type !== 'Loan' && t.category !== 'Loan') return;
        const acc = accounts.find(a => a.id === t.accountId);
        if (!acc) return;
        const currency = acc.currency as 'BDT' | 'LYD' | 'USD';
        if (!loanStatsByCurr[currency]) return;
        if (t.subCategory === 'Borrowed') loanStatsByCurr[currency].borrowed += t.amount;
        else if (t.subCategory === 'Repaid Borrowed Money') loanStatsByCurr[currency].repaid += t.amount;
        else if (t.subCategory === 'Lent') loanStatsByCurr[currency].lent += t.amount;
        else if (t.subCategory === 'Received Lent Money') loanStatsByCurr[currency].received += t.amount;
      });

      const totalCurrentBorrowedBDT = Math.max(0, loanStatsByCurr.BDT.borrowed - loanStatsByCurr.BDT.repaid)
        + Math.max(0, loanStatsByCurr.LYD.borrowed - loanStatsByCurr.LYD.repaid) * LYD_rate
        + Math.max(0, loanStatsByCurr.USD.borrowed - loanStatsByCurr.USD.repaid) * USD_rate;

      const totalCurrentLentBDT = Math.max(0, loanStatsByCurr.BDT.lent - loanStatsByCurr.BDT.received)
        + Math.max(0, loanStatsByCurr.LYD.lent - loanStatsByCurr.LYD.received) * LYD_rate
        + Math.max(0, loanStatsByCurr.USD.lent - loanStatsByCurr.USD.received) * USD_rate;

      const rawSum = totalBDT + (totalLYD * LYD_rate) + (totalUSD * USD_rate);
      const totalConvertedBDT = rawSum - totalCurrentBorrowedBDT + totalCurrentLentBDT;

      const citySnap = cityAcc ? (balances[cityAcc.id] || 0) : 0;
      const ibblSnap = ibblAcc ? (balances[ibblAcc.id] || 0) : 0;
      const cityFcySnap = cityFcyAcc ? (balances[cityFcyAcc.id] || 0) : 0;

      const monthStart = `${snap.getFullYear()}-${String(snap.getMonth() + 1).padStart(2, '0')}-01`;
      const monthEnd = snapStr;

      let monthIncomeBDT = 0;
      let monthExpenseBDT = 0;

      transactions.forEach(t => {
        if (t.date >= monthStart && t.date <= monthEnd) {
          if (t.category === 'Initial Balance' || t.id.startsWith('init-bal-')) return;
          if (t.type === 'Loan' || t.category === 'Loan') return;
          const acc = accounts.find(a => a.id === t.accountId);
          const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
          const bdt = (t.amount || 0) * mult;
          if (t.type === 'Income') {
            monthIncomeBDT += bdt;
          } else if (t.type === 'Expense') {
            monthExpenseBDT += bdt;
          }
        }
      });

      accountsTrendData.push({
        name: label,
        'Total (BDT)': totalConvertedBDT,
        'Savings Income (BDT)': monthIncomeBDT,
        'Savings Expense (BDT)': monthExpenseBDT,
        'City Islamic': citySnap,
        'IBBL Savings': ibblSnap,
        'City FCY': cityFcySnap,
      });
    }

    const accountsTrends: Record<string, { name: string; value: number; rawValue: number; investment?: number; expense?: number; income?: number }[]> = {
      'Total (BDT)': accountsTrendData.map((d, i) => ({
        name: d.name,
        value: d['Total (BDT)'],
        rawValue: d['Total (BDT)'],
        investment: (rawTrendData[i]?.['Total'] as number) || 0,
      })),
      'Cash Flow (BDT)': accountsTrendData.map(d => ({
        name: d.name,
        value: (d as any)['Savings Income (BDT)'],
        rawValue: (d as any)['Savings Income (BDT)'],
        income: (d as any)['Savings Income (BDT)'],
        expense: (d as any)['Savings Expense (BDT)'],
      })),
      'Savings (BDT)': accountsTrendData.map(d => ({
        name: d.name,
        value: (d as any)['Savings Income (BDT)'],
        rawValue: (d as any)['Savings Income (BDT)'],
        income: (d as any)['Savings Income (BDT)'],
        expense: (d as any)['Savings Expense (BDT)'],
      })),
      'City Islamic': accountsTrendData.map(d => ({ name: d.name, value: d['City Islamic'], rawValue: d['City Islamic'] })),
      'IBBL Savings': accountsTrendData.map(d => ({ name: d.name, value: d['IBBL Savings'], rawValue: d['IBBL Savings'] })),
      'City FCY': accountsTrendData.map(d => ({ name: d.name, value: d['City FCY'], rawValue: d['City FCY'] })),
    };

    // ── Cash Flow Bar Chart ──
    const barData = (() => {
      const points: { name: string; income: number; expense: number }[] = [];
      for (let i = 0; i < 24; i++) {
        const snap = new Date();
        snap.setDate(1);
        snap.setMonth(snap.getMonth() - (23 - i));
        const year = snap.getFullYear();
        const month = snap.getMonth();
        const monthStart = `${year}-${String(month + 1).padStart(2, '0')}-01`;
        const lastDay = new Date(year, month + 1, 0).getDate();
        const monthEnd = `${year}-${String(month + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
        const label = snap.toLocaleString('default', { month: 'short', year: '2-digit' });

        let income = 0;
        income += state.fdrs.reduce((sum, f) => {
          const rate = f.currency === 'USD' ? globalUsdRate : (f.currency === 'LYD' ? globalLydRate : 1);
          return sum + f.transactions.filter(t => t.type === 'Profit' && t.handling !== 'Added' && t.date >= monthStart && t.date <= monthEnd).reduce((s, t) => s + (t.amount * rate), 0);
        }, 0);
        income += state.sukuks.reduce((sum, s) => sum + (s.installments || []).filter(inst => inst.isPaid && (inst.actualDate || inst.date) >= monthStart && (inst.actualDate || inst.date) <= monthEnd).reduce((s2, inst) => s2 + (inst.actualAmount || inst.amount), 0), 0);
        income += state.onlineInvestments.reduce((sum, inv) => {
          const rate = inv.currency === 'USD' ? globalUsdRate : (inv.currency === 'LYD' ? globalLydRate : 1);
          return sum + (inv.installments || []).filter(inst => inst.isPaid && (inst.actualDate || inst.date) >= monthStart && (inst.actualDate || inst.date) <= monthEnd).reduce((s2, inst) => s2 + ((inst.actualAmount || inst.amount) * rate), 0);
        }, 0);
        income += dseTransactions.filter((t: any) => (t.type === 'Dividend') && t.date >= monthStart && t.date <= monthEnd).reduce((s: number, t: any) => s + Math.abs(t.total), 0);

        let expense = 0;
        expense += state.mutualFunds.reduce((sum, f) => sum + f.transactions.filter(t => t.type === 'Buy' && !t.isDividend && t.date >= monthStart && t.date <= monthEnd).reduce((s, t) => s + (t.sipAmount || t.amount || 0), 0), 0);
        expense += state.onlineInvestments.filter(inv => inv.investmentDate >= monthStart && inv.investmentDate <= monthEnd).reduce((sum, inv) => {
          const rate = inv.currency === 'USD' ? globalUsdRate : (inv.currency === 'LYD' ? globalLydRate : 1);
          return sum + (inv.amount * rate);
        }, 0);
        expense += state.sukuks.filter(s => s.investmentDate >= monthStart && s.investmentDate <= monthEnd).reduce((sum, s) => sum + s.principalAmount, 0);
        expense += state.fdrs.filter(f => f.investmentDate >= monthStart && f.investmentDate <= monthEnd).reduce((sum, f) => {
          const rate = f.currency === 'USD' ? globalUsdRate : (f.currency === 'LYD' ? globalLydRate : 1);
          return sum + (f.principal * rate);
        }, 0);
        expense += dseTransactions.filter((t: any) => t.type === 'Buy' && t.date >= monthStart && t.date <= monthEnd).reduce((s: number, t: any) => s + Math.abs(t.total), 0);

        points.push({ name: label, income, expense });
      }
      return points;
    })();

    return {
      totalNetWorth: current.totalBalanceInBDT,
      totalBalanceInBDT: current.totalBalanceInBDT,
      totalInvestmentHolding: current.totalInvestmentHolding,
      totalActiveInvestment: current.totalActiveInvestment,
      totalInvested: current.totalInvested,
      totalProfit: current.totalProfit,
      netWorthYoy: totalBalanceInBdtYoy,
      totalBalanceInBdtYoy,
      investmentHoldingYoy,
      activeInvYoy,
      totalInvYoy,
      totalProfitYoy,
      pieData,
      assetAllocationData: current.assetAllocationData,
      netWorthTrend: netWorthTrends,
      accountsTrend: accountsTrends,
      barData,
      seriesMetadata,
      totalCurrentBorrowedBDT: totalLoanStats.totalCurrentBorrowedBDT,
      totalCurrentLentBDT: totalLoanStats.totalCurrentLentBDT,
    };
  }, [state, dseTransactions, startStr, endStr, lyStartStr, lyEndStr]);

  if (state.isLocked) {
    return (
      <PinLogin
        correctPin={state.pin}
        setPin={(pin) => updateState(s => ({ ...s, pin }))}
        onSuccess={() => updateState(s => ({ ...s, isLocked: false }))}
      />
    );
  }

  const formatYoy = (pct: number | null): number => {
    if (pct === null) return 0;
    return Math.round(pct * 10) / 10;
  };

  const renderModule = () => {
    switch (activeModule) {
      case 'dashboard':
        return (
          <div className="space-y-6 sm:space-y-8">
          {/* Dashboard Notification */}
          {dashboardNotification && (
            <div className={cn(
              "flex items-center gap-3 p-4 rounded-xl border animate-in fade-in slide-in-from-top-4 duration-300",
              dashboardNotification.type === 'success' ? "bg-teal-400/10 border-teal-400/20 text-teal-400" : "bg-rose-500/10 border-rose-500/20 text-rose-500"
            )}>
              {dashboardNotification.type === 'success' ? <CheckCircle2 size={20} /> : <XCircle size={20} />}
              <p className="text-body font-bold uppercase">{dashboardNotification.message}</p>
            </div>
          )}
          {/* Dashboard Range Selector */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 relative">
            <div className="flex flex-col sm:flex-row sm:items-center gap-3 w-full sm:w-auto">
              {/* LINE 1 (Mobile): Range + Settings */}
              <div className="flex items-center justify-between w-full sm:w-auto gap-2">
                {/* Range Selection */}
                <div className="relative flex-1 sm:flex-none">
                  {/* Mobile Dropdown */}
                  <div className="block sm:hidden">
                    <button
                      onClick={() => setIsRangeMenuOpen(!isRangeMenuOpen)}
                      className="flex items-center justify-between gap-4 bg-slate-950 border border-slate-800 rounded-lg px-4 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase w-full"
                    >
                      <div className="flex items-center gap-2">
                        <Calendar size={14} className="text-teal-400" />
                        {historyRange === 'this'
                          ? 'This month'
                          : historyRange === 'fiscal'
                            ? 'Fiscal'
                            : 'Custom'}
                      </div>
                      <ChevronDown
                        size={14}
                        className={cn(
                          "text-slate-500 transition-transform",
                          isRangeMenuOpen ? "rotate-180 text-teal-400" : ""
                        )}
                      />
                    </button>

                    {isRangeMenuOpen && (
                      <>
                        <div
                          className="fixed inset-0 z-40"
                          onClick={() => setIsRangeMenuOpen(false)}
                        />
                        <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95 backdrop-blur-md">
                          {(['this', 'fiscal', 'custom'] as const).map(id => (
                            <button
                              key={id}
                              onClick={() => {
                                handleRangeChange(id);
                                setIsRangeMenuOpen(false);
                              }}
                              className={cn(
                                "w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase",
                                historyRange === id
                                  ? "bg-teal-400 text-slate-950"
                                  : "text-slate-300 hover:bg-slate-800"
                              )}
                            >
                              {id === 'this'
                                ? 'This month'
                                : id === 'fiscal'
                                  ? 'Fiscal'
                                  : 'Custom'}
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                  </div>

                  {/* Desktop Tabs */}
                  <div className="hidden sm:flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1 font-sans">
                    {(['this', 'fiscal', 'custom'] as const).map(id => (
                      <button
                        key={id}
                        onClick={() => handleRangeChange(id)}
                        className={cn(
                          "px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all border",
                          historyRange === id
                            ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20"
                            : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white"
                        )}
                      >
                        {id === 'this'
                          ? 'This month'
                          : id === 'fiscal'
                            ? 'Fiscal'
                            : 'Custom'}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Settings Button - Mobile */}
                <div className="block sm:hidden">
                  <div className="relative">
                    <button
                      onClick={() => setIsSettingsMenuOpen(!isSettingsMenuOpen)}
                      className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/10 hover:bg-teal-300 hover:shadow-teal-400/20"
                    >
                      <Settings size={14} />
                      <ChevronDown size={14} className={cn("opacity-50 transition-transform", isSettingsMenuOpen ? "rotate-180" : "")} />
                    </button>

                    {isSettingsMenuOpen && (
                      <>
                        <div className="fixed inset-0 z-40" onClick={() => setIsSettingsMenuOpen(false)} />
                        <div className="absolute right-0 mt-2 w-48 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-2 animate-in fade-in zoom-in-95 backdrop-blur-xl transition-all">
                          <button
                            onClick={() => { handleDashboardBackup(); setIsSettingsMenuOpen(false); }}
                            className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                          >
                            <Download size={14} className="text-teal-400" />
                            BACKUP
                          </button>
                          <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
                            <Upload size={14} className="text-teal-400" />
                            RESTORE
                            <input
                              type="file"
                              className="hidden"
                              accept=".json"
                              onChange={(e) => {
                                setIsSettingsMenuOpen(false);
                                handleDashboardRestoreFile(e);
                              }}
                            />
                          </label>
                        </div>
                      </>
                    )}
                  </div>
                </div>
              </div>

              {/* Date Browsing Controls */}
              <div className="flex items-center gap-1.5 px-1 animate-in fade-in duration-300">
                <button
                  onClick={() => {
                    if (historyRange === 'this') {
                      navigateHistoryThisMonth(-1);
                    } else if (historyRange === 'fiscal') {
                      setHistoryFiscalStartYear(prev => prev - 1);
                    } else {
                      navigateHistoryCustomMonth(-1);
                    }
                  }}
                  className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
                >
                  <ChevronLeft size={14} />
                </button>

                {historyRange === 'custom' ? (
                  <div className="flex items-center gap-1.5">
                    <input
                      type="date"
                      value={historyCustomDates.start}
                      onChange={e => setHistoryCustomDates(prev => ({
                        ...prev,
                        start: e.target.value
                      }))}
                      className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer"
                    />
                    <span className="text-slate-700 font-bold text-[10px]">–</span>
                    <input
                      type="date"
                      value={historyCustomDates.end}
                      onChange={e => setHistoryCustomDates(prev => ({
                        ...prev,
                        end: e.target.value
                      }))}
                      className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer"
                    />
                  </div>
                ) : historyRange === 'fiscal' ? (
                  <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-white uppercase select-none tracking-wider whitespace-nowrap min-w-[160px] text-center">
                    July {historyFiscalStartYear} - June {historyFiscalStartYear + 1}
                  </div>
                ) : (
                  <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-teal-400 uppercase select-none tracking-wider whitespace-nowrap min-w-[120px] text-center">
                    {formatHistoryThisMonthLabel(historyThisMonthDate)}
                  </div>
                )}

                <button
                  onClick={() => {
                    if (historyRange === 'this') {
                      navigateHistoryThisMonth(1);
                    } else if (historyRange === 'fiscal') {
                      setHistoryFiscalStartYear(prev => prev + 1);
                    } else {
                      navigateHistoryCustomMonth(1);
                    }
                  }}
                  className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
                >
                  <ChevronRight size={14} />
                </button>
              </div>
            </div>

            {/* Settings Button - Desktop */}
            <div className="hidden sm:block relative">
              <button
                onClick={() => setIsSettingsMenuOpen(!isSettingsMenuOpen)}
                className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/10 hover:bg-teal-300 hover:shadow-teal-400/20"
              >
                <Settings size={14} />
                <span className="hidden sm:inline">Settings</span>
                <ChevronDown size={14} className={cn("opacity-50 transition-transform", isSettingsMenuOpen ? "rotate-180" : "")} />
              </button>

              {isSettingsMenuOpen && (
                <>
                  <div className="fixed inset-0 z-40" onClick={() => setIsSettingsMenuOpen(false)} />
                  <div className="absolute right-0 mt-2 w-48 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-2 backdrop-blur-xl">
                    <button
                      onClick={() => { handleDashboardBackup(); setIsSettingsMenuOpen(false); }}
                      className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                    >
                      <Download size={14} className="text-teal-400" />
                      BACKUP
                    </button>
                    <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
                      <Upload size={14} className="text-teal-400" />
                      RESTORE
                      <input
                        type="file"
                        className="hidden"
                        accept=".json"
                        onChange={(e) => {
                          setIsSettingsMenuOpen(false);
                          handleDashboardRestoreFile(e);
                        }}
                      />
                    </label>
                  </div>
                </>
              )}
            </div>
          </div>

            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 lg:gap-4">
              <SummaryCard
                className="col-span-1"
                title="Total Net Worth"
                subtitle="Total Balance in BDT"
                value={stats.totalBalanceInBDT}
                trend={formatYoy(stats.totalBalanceInBdtYoy)}
                trendLabel="vs Last Year"
                icon={DollarSign}
                color="teal"
                hideFooter
                borrowedValue={stats.totalCurrentBorrowedBDT}
                lentValue={stats.totalCurrentLentBDT}
              />
              <SummaryCard
                className="col-span-1"
                title="Total Investment Holding"
                subtitle="Balance in Investment A/C"
                value={stats.totalInvestmentHolding}
                trend={formatYoy(stats.investmentHoldingYoy)}
                trendLabel="vs Last Year"
                icon={Wallet}
                color="blue"
                hideFooter
              />
              <SummaryCard
                className="col-span-1"
                title="Active Investment"
                subtitle="Total Active Investment"
                value={stats.totalActiveInvestment}
                trend={formatYoy(stats.activeInvYoy)}
                trendLabel="vs Last Year"
                icon={Briefcase}
                color="purple"
                hideFooter
              />
              <SummaryCard
                className="col-span-1"
                title="Total Profit"
                subtitle="Total Profit Received"
                value={stats.totalProfit}
                trend={formatYoy(stats.totalProfitYoy)}
                trendLabel="vs Last Year"
                icon={TrendingUp}
                color="emerald"
                hideFooter
              />
            </div>

            <DashboardCharts
              pieData={stats.pieData}
              assetAllocationData={stats.assetAllocationData}
              lineDataMap={stats.netWorthTrend}
              accountsLineDataMap={stats.accountsTrend}
              barData={stats.barData}
              seriesMetadata={stats.seriesMetadata}
            />
          </div>
        );

      case 'dse':
      case 'dse-summary':
      case 'dse-transactions':
      case 'dse-holdings':
      case 'dse-analytics':
      case 'dse-settings':
        return (
          <DseTrackerModule
            holdings={state.dseHoldings}
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            onTitleChange={setDseTrackerTitle}
            activeTab={activeModule.startsWith('dse-') ? (activeModule.replace('dse-', '') as any) : 'summary'}
            onNavigateToModule={handleNavigateToModule}
            onDeleteLinkedTransfer={(groupId) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.filter(t => t.transferGroupId !== groupId)
              }));
            }}
            onUpdateLinkedTransfer={(groupId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.map(t =>
                  t.transferGroupId === groupId
                    ? {
                        ...t,
                        date: updates.date,
                        description: updates.description ?? t.description,
                        amount: updates.amount,
                        toAmount: updates.amount
                      }
                    : t
                )
              }));
            }}
            onDeleteLinkedDseTransaction={(dseIds) => {
              const ids = Array.isArray(dseIds) ? dseIds : [dseIds];
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.filter(t => !t.dseTxId || !ids.includes(t.dseTxId))
              }));
            }}
            onUpdateLinkedDseTransaction={(dseTxId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.map(t =>
                  t.dseTxId === dseTxId
                    ? {
                        ...t,
                        date: updates.date,
                        amount: updates.amount,
                        toAmount: updates.amount,
                        description: updates.description ?? t.description
                      }
                    : t
                )
              }));
            }}
            onSyncSellProfitLoss={(sellTx, pnl) => {
              updateState(s => {
                const existing = s.incomeExpenseTransactions.find(tx => tx.dseTxId === sellTx.id);
                const boAcc = s.incomeExpenseAccounts?.find(a => isBoAccount(a) || a.id === 'bdt-bo-account')
                  || s.incomeExpenseAccounts?.find(a => !a.isParent)
                  || { id: 'bdt-bo-account' };

                const isProfit = pnl >= 0;
                const absAmount = Math.round(Math.abs(pnl) * 100) / 100;

                let categoryName = 'Finance Income';
                let subCategoryName = isProfit ? 'Capital Gain' : 'Capital Loss';

                const finCat = s.incomeExpenseCategories?.find(c => c.name.toLowerCase().includes('finance') || c.name.toLowerCase().includes('investment'));
                if (finCat) {
                  categoryName = finCat.name;
                }

                const ticker = sellTx.ticker ? sellTx.ticker.toUpperCase() : '';
                const defaultSellDesc = ticker ? `Stocks P&L: ${ticker}` : 'Stocks P&L';
                const desc = sellTx.notes?.trim()
                  ? sellTx.notes
                  : defaultSellDesc;

                const updatedTx: any = {
                  id: existing ? existing.id : crypto.randomUUID(),
                  date: sellTx.date,
                  amount: absAmount,
                  type: isProfit ? 'Income' : 'Expense',
                  accountId: existing?.accountId || boAcc.id,
                  category: existing?.category || categoryName,
                  subCategory: subCategoryName,
                  description: desc,
                  dseTxId: sellTx.id,
                  autoSyncDse: true
                };

                const remaining = s.incomeExpenseTransactions.filter(tx => tx.dseTxId !== sellTx.id);
                return {
                  ...s,
                  incomeExpenseTransactions: [updatedTx, ...remaining]
                };
              });
            }}
            onAdd={(newHolding) => {
              updateState(s => ({ ...s, dseHoldings: [...s.dseHoldings, { ...newHolding, id: crypto.randomUUID() }] }));
            }}
            onUpdate={(id, updates) => {
              updateState(s => ({ ...s, dseHoldings: s.dseHoldings.map(h => h.id === id ? { ...h, ...updates } : h) }));
            }}
            onDelete={(id) => {
              updateState(s => ({ ...s, dseHoldings: s.dseHoldings.filter(h => h.id !== id) }));
            }}
            onReplaceAllHoldings={(newHoldings) => {
              updateState(s => ({
                ...s,
                dseHoldings: newHoldings.map(h => ({ ...h, id: h.id || crypto.randomUUID() }))
              }));
            }}
          />
        );

      case 'online':
        return (
          <OnlineInvestments
            investments={onlineInvestments}
            conversionRates={state.conversionRates}
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            inheritedData={inheritedAddData}
            onClearInheritedData={handleClearInheritedData}
            onTitleChange={setOnlineInvestmentsTitle}
            onNavigateToModule={handleNavigateToModule}
            onDeleteLinkedTransfer={(groupId) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.filter(t => t.transferGroupId !== groupId)
              }));
            }}
            onUpdateLinkedTransfer={(groupId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.map(t =>
                  t.transferGroupId === groupId
                    ? {
                        ...t,
                        date: updates.date,
                        description: updates.description ?? t.description,
                        amount: updates.amount,
                        toAmount: updates.amount
                      }
                    : t
                )
              }));
            }}
            onDeleteLinkedIncomeTx={(txId) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.filter(t => t.id !== txId && t.onlineTxId !== txId && t.transferGroupId !== txId)
              }));
            }}
            onUpdateLinkedIncomeTx={(txId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.map(t =>
                  (t.id === txId || t.onlineTxId === txId || t.transferGroupId === txId)
                    ? {
                        ...t,
                        date: updates.date,
                        description: updates.description ?? t.description,
                        amount: updates.amount,
                        toAmount: updates.amount
                      }
                    : t
                )
              }));
            }}
            onAdd={(newInv) => {
  const withId = {
    ...newInv,
    id: crypto.randomUUID()
  };

  updateState(s => {
    const updated = [...s.onlineInvestments, withId];

    markDirty('onlineInvestments');
    schedulePush('onlineInvestments', () => updated);

    return {
      ...s,
      onlineInvestments: updated
    };
  }, []);
}}

onUpdate={(id, updates) => {
  updateState(s => {
    const updated = s.onlineInvestments.map(inv =>
      inv.id === id ? { ...inv, ...updates } : inv
    );

    markDirty('onlineInvestments');
    schedulePush('onlineInvestments', () => updated);

    return {
      ...s,
      onlineInvestments: updated
    };
  }, []);
}}
onDelete={(id) => {
  updateState(s => {
    const updated = s.onlineInvestments.filter(inv => inv.id !== id);

    markDirty('onlineInvestments');
    schedulePush('onlineInvestments', () => updated);

    return {
      ...s,
      onlineInvestments: updated
    };
  }, []);
}}
onBatchDelete={(ids) => {
  updateState(s => {
    const updated = s.onlineInvestments.filter(inv => !ids.includes(inv.id));

    markDirty('onlineInvestments');
    schedulePush('onlineInvestments', () => updated);

    return {
      ...s,
      onlineInvestments: updated
    };
  }, []);
}}
                        onBatchAdd={(newInvs) => {
  updateState(s => {

    const existingIds = new Set(
      s.onlineInvestments.map(inv => inv.id)
    );

    const normalized = newInvs.map(inv => ({
      ...inv,
      id: inv.id || crypto.randomUUID()
    }));

    const uniqueNewInvs = normalized.filter(
      inv => !existingIds.has(inv.id)
    );

    const updated = [
      ...s.onlineInvestments,
      ...uniqueNewInvs
    ];

    markDirty('onlineInvestments');

    schedulePush('onlineInvestments', () => updated);

    return {
      ...s,
      onlineInvestments: updated
    };
  }, []);
}}
                        onReplaceAll={(newInvs) => {
                          updateState(s => {
                            const updated = newInvs.map(inv => ({
                              ...inv,
                              id: inv.id || crypto.randomUUID()
                            }));
                            markDirty('onlineInvestments');
                            schedulePush('onlineInvestments', () => updated);
                            return { ...s, onlineInvestments: updated };
                          }, []);
                        }}
          />
        );

            case 'mutual-funds':
        return (
          <MutualFundsModule
            investments={state.mutualFunds}
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            inheritedData={inheritedAddData}
            onClearInheritedData={handleClearInheritedData}
            onTitleChange={setMutualFundsTitle}
            onNavigateToModule={handleNavigateToModule}
            onDeleteLinkedTransfer={(groupId) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.filter(t => t.transferGroupId !== groupId && t.id !== groupId && t.mfTxId !== groupId)
              }));
            }}
            onUpdateLinkedTransfer={(groupId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.map(t =>
                  (t.transferGroupId === groupId || t.id === groupId || t.mfTxId === groupId)
                    ? {
                        ...t,
                        date: updates.date,
                        description: updates.description ?? t.description,
                        amount: updates.amount,
                        toAmount: updates.amount
                      }
                    : t
                )
              }));
            }}

            onAdd={(newInv) => {
  const withId = {
    ...newInv,
    id: crypto.randomUUID()
  };

  updateState(s => {
    const updated = [...s.mutualFunds, withId];

    markDirty('mutualFunds');
    schedulePush('mutualFunds', () => updated);

    return {
      ...s,
      mutualFunds: updated
    };
  }, []);
}}

onUpdate={(id, updates) => {
  updateState(s => {
    const updated = s.mutualFunds.map(inv =>
      inv.id === id ? { ...inv, ...updates } : inv
    );

    markDirty('mutualFunds');
    schedulePush('mutualFunds', () => updated);

    return {
      ...s,
      mutualFunds: updated
    };
  }, []);
}}


onDelete={(id) => {
  updateState(s => {
    const updated = s.mutualFunds.filter(inv => inv.id !== id);

    markDirty('mutualFunds');
    schedulePush('mutualFunds', () => updated);

    return {
      ...s,
      mutualFunds: updated
    };
  }, []);
}}


onBatchDelete={(ids) => {
  updateState(s => {
    const updated = s.mutualFunds.filter(inv => !ids.includes(inv.id));

    markDirty('mutualFunds');
    schedulePush('mutualFunds', () => updated);

    return {
      ...s,
      mutualFunds: updated
    };
  }, []);
}}

            onBatchAdd={(newInvs) => {
  updateState(s => {

    const existingIds = new Set(
      s.mutualFunds.map(inv => inv.id)
    );

    const normalized = newInvs.map(inv => ({
      ...inv,
      id: inv.id || crypto.randomUUID()
    }));

    const uniqueNewInvs = normalized.filter(
      inv => !existingIds.has(inv.id)
    );

    const updated = [
      ...s.mutualFunds,
      ...uniqueNewInvs
    ];

    markDirty('mutualFunds');

    schedulePush('mutualFunds', () => updated);

    return {
      ...s,
      mutualFunds: updated
    };
  }, []);
}}
            onReplaceAll={(newInvs) => {
              updateState(s => {
                const updated = newInvs.map(inv => ({
                  ...inv,
                  id: inv.id || crypto.randomUUID()
                }));
                markDirty('mutualFunds');
                schedulePush('mutualFunds', () => updated);
                return { ...s, mutualFunds: updated };
              }, []);
            }}
          />
        );

      case 'fdrs':
        return (
          <FixedDepositsModule
            investments={state.fdrs}
            conversionRates={state.conversionRates}
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            inheritedData={inheritedAddData}
            onClearInheritedData={handleClearInheritedData}
            onTitleChange={setFixedDepositsTitle}
            onNavigateToModule={handleNavigateToModule}
            onCreateFdrAccount={handleCreateFdrAccount}
            onDeleteLinkedIncome={(txId) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions?.filter(t => t.id !== txId && t.fdrTxId !== txId)
              }));
            }}
            onUpdateLinkedIncome={(txId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions?.map(t =>
                  (t.id === txId || t.fdrTxId === txId)
                    ? {
                        ...t,
                        date: updates.date,
                        amount: updates.amount,
                        ...(updates.description ? { description: updates.description } : {})
                      }
                    : t
                )
              }));
            }}
            onAdd={(newInv) => {
  const withId = {
    ...newInv,
    id: crypto.randomUUID()
  };

  updateState(s => {
    const updated = [...s.fdrs, withId];

    markDirty('fixedDeposits');
    schedulePush('fixedDeposits', () => updated);

    return {
      ...s,
      fdrs: updated
    };
  }, []);
}}





onUpdate={(id, updates) => {
  updateState(s => {
    const updated = s.fdrs.map(inv =>
      inv.id === id ? { ...inv, ...updates } : inv
    );

    markDirty('fixedDeposits');
    schedulePush('fixedDeposits', () => updated);

    return {
      ...s,
      fdrs: updated
    };
  }, []);
}}





onDelete={(id) => {
  updateState(s => {
    const updated = s.fdrs.filter(inv => inv.id !== id);

    markDirty('fixedDeposits');
    schedulePush('fixedDeposits', () => updated);

    return {
      ...s,
      fdrs: updated
    };
  }, []);
}}





onBatchDelete={(ids) => {
  updateState(s => {
    const updated = s.fdrs.filter(inv => !ids.includes(inv.id));

    markDirty('fixedDeposits');
    schedulePush('fixedDeposits', () => updated);

    return {
      ...s,
      fdrs: updated
    };
  }, []);
}}

            onBatchAdd={(newInvs) => {
  updateState(s => {

    const existingIds = new Set(
      s.fdrs.map(inv => inv.id)
    );

    const normalized = newInvs.map(inv => ({
      ...inv,
      id: inv.id || crypto.randomUUID()
    }));

    const uniqueNewInvs = normalized.filter(
      inv => !existingIds.has(inv.id)
    );

    const updated = [
      ...s.fdrs,
      ...uniqueNewInvs
    ];

    markDirty('fixedDeposits');

    schedulePush('fixedDeposits', () => updated);

    return {
      ...s,
      fdrs: updated
    };
  }, []);
}}
            onReplaceAll={(newInvs) => {
              updateState(s => {
                const updated = newInvs.map(inv => ({
                  ...inv,
                  id: inv.id || crypto.randomUUID()
                }));
                markDirty('fixedDeposits');
                schedulePush('fixedDeposits', () => updated);
                return { ...s, fdrs: updated };
              }, []);
            }}
          />
        );

      case 'sukuk':
        return (
          <SukukInvestments
            investments={sukuks}
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            inheritedData={inheritedAddData}
            onClearInheritedData={handleClearInheritedData}
            onTitleChange={setSukukTitle}
            onNavigateToModule={handleNavigateToModule}
            onDeleteLinkedTransfer={(groupId) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.filter(t => t.transferGroupId !== groupId)
              }));
            }}
            onUpdateLinkedTransfer={(groupId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.map(t =>
                  t.transferGroupId === groupId
                    ? {
                        ...t,
                        date: updates.date,
                        description: updates.description ?? t.description,
                        amount: updates.amount,
                        toAmount: updates.amount
                      }
                    : t
                )
              }));
            }}
            onDeleteLinkedIncomeTx={(txId) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.filter(t => t.id !== txId && t.sukukTxId !== txId)
              }));
            }}
            onUpdateLinkedIncomeTx={(txId, updates) => {
              updateState(s => ({
                ...s,
                incomeExpenseTransactions: s.incomeExpenseTransactions.map(t =>
                  (t.id === txId || t.sukukTxId === txId)
                    ? {
                        ...t,
                        date: updates.date,
                        description: updates.description ?? t.description,
                        amount: updates.amount
                      }
                    : t
                )
              }));
            }}
            
            onAdd={(newInv) => {
  const withId = {
    ...newInv,
    id: crypto.randomUUID()
  };

  updateState(s => {
    const updated = [...s.sukuks, withId];

    markDirty('sukuk');
    schedulePush('sukuk', () => updated);

    return {
      ...s,
      sukuks: updated
    };
  }, []);
}}





onUpdate={(id, updates) => {
  updateState(s => {
    const updated = s.sukuks.map(inv =>
      inv.id === id ? { ...inv, ...updates } : inv
    );

    markDirty('sukuk');
    schedulePush('sukuk', () => updated);

    return {
      ...s,
      sukuks: updated
    };
  }, []);
}}





onDelete={(id) => {
  updateState(s => {
    const updated = s.sukuks.filter(inv => inv.id !== id);

    markDirty('sukuk');
    schedulePush('sukuk', () => updated);

    return {
      ...s,
      sukuks: updated
    };
  }, []);
}}





onBatchDelete={(ids) => {
  updateState(s => {
    const updated = s.sukuks.filter(inv => !ids.includes(inv.id));

    markDirty('sukuk');
    schedulePush('sukuk', () => updated);

    return {
      ...s,
      sukuks: updated
    };
  }, []);
}}

            onBatchAdd={(newInvs) => {
  updateState(s => {

    const existingIds = new Set(
      s.sukuks.map(inv => inv.id)
    );

    const normalized = newInvs.map(inv => ({
      ...inv,
      id: inv.id || crypto.randomUUID()
    }));

    const uniqueNewInvs = normalized.filter(
      inv => !existingIds.has(inv.id)
    );

    const updated = [
      ...s.sukuks,
      ...uniqueNewInvs
    ];

    markDirty('sukuk');

    schedulePush('sukuk', () => updated);

    return {
      ...s,
      sukuks: updated
    };
  }, []);
}}
            onReplaceAll={(newInvs) => {
              updateState(s => {
                const updated = newInvs.map(inv => ({
                  ...inv,
                  id: inv.id || crypto.randomUUID()
                }));
                markDirty('sukuk');
                schedulePush('sukuk', () => updated);
                return { ...s, sukuks: updated };
              }, []);
            }}
          />
        );

      case 'income-expense':
      case 'income-expense-summary':
      case 'income-expense-transactions':
      case 'income-expense-analytics':
      case 'income-expense-settings':
        return (
          <IncomeExpenseModule
            state={state}
            updateState={updateState}
            activeTab={activeModule.startsWith('income-expense-') ? (activeModule.replace('income-expense-', '') as any) : 'summary'}
            setActiveTab={(tab) => {
              setActiveModule('income-expense-' + tab);
            }}
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            onNavigateToModule={handleNavigateToModule}
            inheritedData={inheritedAddData}
            onClearInheritedData={handleClearInheritedData}
          />
        );

      case 'settings':
        return <SettingsModule state={state} updateState={updateState} />;

      default:
        return (
          <div className="flex flex-col items-center justify-center py-20 text-slate-500">
            <h3 className="text-heading font-bold mb-2 uppercase">Module Under Construction</h3>
            <p className="text-body">The {activeModule} module is coming soon in the next update.</p>
          </div>
        );
    }
  };

  return (
    <div className="flex h-screen bg-[#020617] text-white font-sans overflow-hidden">
      <Sidebar
        activeModule={activeModule}
        setActiveModule={setActiveModule}
        isCollapsed={isCollapsed}
        setIsCollapsed={setIsCollapsed}
        isMobileOpen={isMobileOpen}
        setIsMobileOpen={setIsMobileOpen}
      />
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <Navbar
          title={
            activeModule.startsWith('dse') ? dseTrackerTitle :
            activeModule === 'dashboard' ? (
              <span className="flex items-center gap-2">
                DASHBOARD <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">/ {rangeDates.start.toLocaleString('default', { month: 'short', year: 'numeric' })} - {rangeDates.end.toLocaleString('default', { month: 'short', year: 'numeric' })}</span>
              </span>
            ) :
            activeModule === 'mutual-funds' ? mutualFundsTitle :
            activeModule.startsWith('income-expense') ? (
              <span className="flex items-center gap-2">
                INCOME & EXPENSE <span className="text-teal-400 font-display text-xs font-bold tracking-widest leading-none">/ {activeModule.replace('income-expense-', '').replace('income-expense', 'summary').toUpperCase()}</span>
              </span>
            ) :
            activeModule === 'fdrs' ? fixedDepositsTitle :
            activeModule === 'online' ? onlineInvestmentsTitle :
            activeModule === 'sukuk' ? sukukTitle :
            activeModule === 'analytics' ? 'Global Metrics' :
            activeModule.charAt(0).toUpperCase() + activeModule.slice(1).replace('-', ' ')
          }
          onAdd={['online', 'sukuk', 'mutual-funds', 'fdrs', 'income-expense'].includes(activeModule) || activeModule.startsWith('dse') || activeModule.startsWith('income-expense') ? () => { setInheritedAddData(null); setTriggerAdd(true); } : undefined}
          addLabel={activeModule === 'fdrs' ? "New Profit" : (activeModule === 'dse' || activeModule.startsWith('dse-') || activeModule === 'income-expense' || activeModule.startsWith('income-expense-')) ? "New Transaction" : undefined}
          onToggleSidebar={() => setIsMobileOpen(!isMobileOpen)}
          syncStatus={syncStatus}
          syncSummary={syncSummary}
          lastSyncedAt={lastSyncedAt}
          onSync={handleSyncAll}
        />
        <div className="flex-1 p-4 lg:p-8 overflow-y-auto">
          {renderModule()}
        </div>
      </main>
      
      {/* Dashboard Restore Confirmation Modal */}
      {dashboardPendingRestore && (
        <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/70 backdrop-blur-sm animate-in fade-in">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl shadow-2xl p-6 max-w-md w-full mx-4 animate-in zoom-in-95">
            <h3 className="text-subheading font-bold text-white uppercase mb-4">Confirm Restore</h3>
            <div className="flex items-start gap-4 p-4 bg-amber-500/10 border border-amber-500/20 rounded-xl mb-6">
              <AlertCircle className="text-amber-500 w-6 h-6 shrink-0 mt-1" />
              <div>
                <p className="text-body font-bold text-white mb-1 uppercase">Overwrite Existing Data?</p>
                <p className="text-label text-slate-400 leading-relaxed">
                  Restoring will replace all your current investments, transactions, and settings with the data from the backup file. This action cannot be undone.
                </p>
              </div>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => setDashboardPendingRestore(null)}
                className="flex-1 px-4 py-2 rounded-lg bg-slate-800 border border-slate-700 text-slate-300 text-label font-bold uppercase hover:bg-slate-700 transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={confirmDashboardRestore}
                className="flex-1 px-4 py-2 rounded-lg bg-teal-400 text-slate-950 text-label font-bold uppercase hover:bg-teal-300 transition-colors"
              >
                Confirm Restore
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}













