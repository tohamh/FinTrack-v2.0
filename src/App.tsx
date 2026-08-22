/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import { useAppState } from './hooks/useAppState';
import { AppState } from './types';
import { Sidebar, Navbar } from './components/layout/Navigation';
import { PinLogin } from './components/auth/PinLogin';
import { DashboardCharts, SummaryCard } from './components/dashboard/DashboardComponents';
import { OnlineInvestments } from './components/modules/OnlineInvestments';
import { SukukInvestments } from './components/modules/SukukInvestments';
import { DseTrackerModule } from './components/modules/DseTrackerModule';
import { MutualFundsModule } from './components/modules/MutualFundsModule';
import { FixedDepositsModule } from './components/modules/FixedDepositsModule';
import { SettingsModule } from './components/modules/SettingsModule';
import { IncomeExpenseModule } from './components/modules/IncomeExpenseModule';
import { Wallet, TrendingUp, PieChart, Banknote, Landmark, Plus, Briefcase, Coins, DollarSign, Calendar, ChevronDown, ChevronLeft, ChevronRight, Settings, Download, Upload, AlertCircle, CheckCircle2, XCircle } from 'lucide-react';
import * as XLSX from 'xlsx';
import { formatBDT, cn } from './utils/formatters';
import { Button } from './components/ui/BaseComponents';
import { schedulePush, resolveOnStartup, flushOnUnload, markDirty, syncAllModules, startAutoSync } from './utils/sheetSync';
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
    targetModule?: string;
    description?: string;
  } | null>(null);
  const latestStateRef = useRef(state);

  const handleNavigateToModule = (
    module: string,
    initialAddData?: {
      date: string;
      amount: number;
      type?: 'Transfer' | 'Income' | 'Expense' | 'Loan';
      targetModule?: string;
      description?: string;
    }
  ) => {
    if (initialAddData) {
      setInheritedAddData(initialAddData);
    }
    setActiveModule(module);
    setTriggerAdd(true);
  };

useEffect(() => {
  latestStateRef.current = state;
}, [state]);

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

  const [historyRange, setHistoryRange] = useState<'all' | 'last12m' | 'fiscal' | 'custom'>('all');
  // Custom range: default to first of current month → today
  const [historyCustomDates, setHistoryCustomDates] = useState(() => {
    const now = new Date();
    return {
      start: getFirstOfMonth(now),
      end: getTodayStr(),
    };
  });
  // Track the "anchor" month for < > navigation (year + month index)
  const [customNavMonth, setCustomNavMonth] = useState<{ year: number; month: number }>(() => {
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() };
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
    // Use the oldest lastSyncedAt across all modules as the global "last synced"
    const modules = ['onlineInvestments', 'sukuk', 'mutualFunds', 'fixedDeposits'] as const;
    const times = modules.map(m => {
      try {
        const raw = localStorage.getItem(`syncmeta_${m}`);
        return raw ? JSON.parse(raw).lastSyncedAt ?? 0 : 0;
      } catch { return 0; }
    });
    return Math.min(...times.filter(t => t > 0)) || 0;
  });
  const [syncSummary, setSyncSummary] = useState<string>('');

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
    setSyncSummary('');

    try {
      const { pushed, pulled, failed } = await syncAllModules(
       () => ({
  onlineInvestments: state.onlineInvestments,
  sukuk:             state.sukuks,
  mutualFunds:       state.mutualFunds,
  fixedDeposits:     state.fdrs,

  incomeExpense: [
    ...state.incomeExpenseTransactions.map(t => ({
      id: t.id,
      type: 'transaction',
      data: t
    })),

    ...state.incomeExpenseAccounts.map(a => ({
      id: a.id,
      type: 'account',
      data: a
    })),

    ...state.incomeExpenseCategories.map(c => ({
      id: c.id,
      type: 'category',
      data: c
    })),

    {
      id: '__INCOME_EXPENSE_SETTINGS__',
      type: 'settings',
      data: {
        conversionRates: state.conversionRates
      }
    }
  ]
}),
        (module, data) => {
          if (data === null) return;
          // Cloud was newer — apply to local state
          if (module === 'onlineInvestments') updateState(s => ({ ...s, onlineInvestments: data }), []);
          if (module === 'sukuk')             updateState(s => ({ ...s, sukuks: data }), []);
          if (module === 'mutualFunds')       updateState(s => ({ ...s, mutualFunds: data }), []);
          if (module === 'fixedDeposits')     updateState(s => ({ ...s, fdrs: data }), []);
          if (module === 'incomeExpense') {

  const rows = data || [];

  const transactions = rows
    .filter((r: any) => r.type === 'transaction')
    .map((r: any) => r.data);

  const accounts = rows
    .filter((r: any) => r.type === 'account')
    .map((r: any) => r.data);

  const categories = rows
    .filter((r: any) => r.type === 'category')
    .map((r: any) => r.data);

  const settingsRow = rows.find((r: any) => r.type === 'settings');

  updateState(s => ({
    ...s,
    incomeExpenseTransactions: transactions,
    incomeExpenseAccounts: accounts,
    incomeExpenseCategories: categories,
    conversionRates:
      settingsRow?.data?.conversionRates || s.conversionRates
  }), []);
}
        }
      );

      const now = Date.now();
      setLastSyncedAt(now);

      const parts: string[] = [];
      if (pushed.length) parts.push(`↑ Pushed ${pushed.length}`);
      if (pulled.length) parts.push(`↓ Pulled ${pulled.length}`);
      if (failed.length) parts.push(`✗ Failed ${failed.length}`);
      setSyncSummary(parts.join('  ·  ') || 'Already up to date');
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
  }, [syncStatus, state, updateState]);

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
    });
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

  // App.tsx around line 125
// ─── On startup: resolve local vs cloud conflict for each module ───────────
  useEffect(() => {
    const demoIds = ['1', '2', 'completed-1', 'completed-2', 'mf1', 'mf2', 'fdr1', 'fdr2', 's1', 's2', 'ol1', 'ol2'];

    const moduleConfig: { module: ModuleKey; stateKey: keyof AppState }[] = [
      { module: 'fixedDeposits',     stateKey: 'fdrs'              },
      { module: 'mutualFunds',       stateKey: 'mutualFunds'       },
      { module: 'onlineInvestments', stateKey: 'onlineInvestments' },
      { module: 'sukuk',             stateKey: 'sukuks'            },
    ];

resolveOnStartup(
  'incomeExpense',

  [
    ...state.incomeExpenseTransactions.map(t => ({
      id: t.id,
      type: 'transaction',
      data: t
    })),

    ...state.incomeExpenseAccounts.map(a => ({
      id: a.id,
      type: 'account',
      data: a
    })),

    ...state.incomeExpenseCategories.map(c => ({
      id: c.id,
      type: 'category',
      data: c
    })),

    {
      id: '__INCOME_EXPENSE_SETTINGS__',
      type: 'settings',
      data: {
        conversionRates: state.conversionRates
      }
    }
  ],

  () => [
    ...latestStateRef.current.incomeExpenseTransactions.map(t => ({
      id: t.id,
      type: 'transaction',
      data: t
    })),

    ...latestStateRef.current.incomeExpenseAccounts.map(a => ({
      id: a.id,
      type: 'account',
      data: a
    })),

    ...latestStateRef.current.incomeExpenseCategories.map(c => ({
      id: c.id,
      type: 'category',
      data: c
    })),

    {
      id: '__INCOME_EXPENSE_SETTINGS__',
      type: 'settings',
      data: {
        conversionRates: latestStateRef.current.conversionRates
      }
    }
  ]

).then(cloudData => {

  if (!cloudData) return;

  const transactions = cloudData
    .filter((r: any) => r.type === 'transaction')
    .map((r: any) => r.data);

  const accounts = cloudData
    .filter((r: any) => r.type === 'account')
    .map((r: any) => r.data);

  const categories = cloudData
    .filter((r: any) => r.type === 'category')
    .map((r: any) => r.data);

  const settingsRow = cloudData.find((r: any) => r.type === 'settings');

  updateState(s => ({
    ...s,
    incomeExpenseTransactions: transactions,
    incomeExpenseAccounts: accounts,
    incomeExpenseCategories: categories,
    conversionRates:
      settingsRow?.data?.conversionRates || s.conversionRates
  }), []);
});

    moduleConfig.forEach(({ module, stateKey }) => {
    resolveOnStartup(
      module,
      (state as any)[stateKey] ?? [],
      () => ((state as any)[stateKey] ?? []).filter((item: any) => !demoIds.includes(item.id))
    ).then(cloudData => {
      if (cloudData !== null) {
        const filtered = cloudData.filter((item: any) => !demoIds.includes(item.id));
        const localData = (state as any)[stateKey] ?? [];
        if (filtered.length > 0 || localData.length === 0) {
          updateState(s => ({ ...s, [stateKey]: filtered }), []);
        }
      }
    });
  });

    // Register beforeunload flush
    flushOnUnload(() => ({
  fixedDeposits:     state.fdrs,
  mutualFunds:       state.mutualFunds,
  onlineInvestments: state.onlineInvestments,
  sukuk:             state.sukuks,

  incomeExpense: [

    ...state.incomeExpenseTransactions.map(t => ({
      id: t.id,
      type: 'transaction',
      data: t
    })),

    ...state.incomeExpenseAccounts.map(a => ({
      id: a.id,
      type: 'account',
      data: a
    })),

    ...state.incomeExpenseCategories.map(c => ({
      id: c.id,
      type: 'category',
      data: c
    })),

    {
      id: '__INCOME_EXPENSE_SETTINGS__',
      type: 'settings',
      data: {
        conversionRates: state.conversionRates
      }
    }
  ]
}));


        // Start auto sync using ALWAYS-LATEST state
    startAutoSync(() => ({
  fixedDeposits: latestStateRef.current.fdrs,
  mutualFunds: latestStateRef.current.mutualFunds,
  onlineInvestments: latestStateRef.current.onlineInvestments,
  sukuk: latestStateRef.current.sukuks,

  incomeExpense: [

    ...latestStateRef.current.incomeExpenseTransactions.map(t => ({
      id: t.id,
      type: 'transaction',
      data: t
    })),

    ...latestStateRef.current.incomeExpenseAccounts.map(a => ({
      id: a.id,
      type: 'account',
      data: a
    })),

    ...latestStateRef.current.incomeExpenseCategories.map(c => ({
      id: c.id,
      type: 'category',
      data: c
    })),

    {
      id: '__INCOME_EXPENSE_SETTINGS__',
      type: 'settings',
      data: {
        conversionRates: latestStateRef.current.conversionRates
      }
    }
  ]
}));
    
  }, []);





  // ─── When historyRange switches to 'custom', reset to current month ────────
  const handleRangeChange = (newRange: 'all' | 'last12m' | 'fiscal' | 'custom') => {
    if (newRange === 'custom') {
      const now = new Date();
      const navMonth = { year: now.getFullYear(), month: now.getMonth() };
      setCustomNavMonth(navMonth);
      setHistoryCustomDates({
        start: getFirstOfMonth(now),
        end: getTodayStr(),
      });
    }
    setHistoryRange(newRange);
    setIsRangeMenuOpen(false);
  };

  // ─── Navigate custom range by month ───────────────────────────────────────
  const navigateCustomMonth = (direction: -1 | 1) => {
    const now = new Date();
    const newMonth = customNavMonth.month + direction;
    let year = customNavMonth.year;
    let month = newMonth;
    if (month < 0) { month = 11; year -= 1; }
    if (month > 11) { month = 0; year += 1; }
    const navDate = new Date(year, month, 1);
    const isCurrentMonth = year === now.getFullYear() && month === now.getMonth();
    setCustomNavMonth({ year, month });
    setHistoryCustomDates({
      start: getFirstOfMonth(navDate),
      end: isCurrentMonth ? getTodayStr() : getLastOfMonth(navDate),
    });
  };

  const onlineInvestments = state.onlineInvestments;
  const sukuks = state.sukuks;

  // ─── Shared date range resolver ────────────────────────────────────────────
  const { startStr, endStr } = useMemo(() => {
    const now = new Date();

    if (historyRange === 'all') {
      return { startStr: '0000-01-01', endStr: '9999-12-31' };
    }
    if (historyRange === 'last12m') {
      return {
        startStr: toDateStr(new Date(now.getFullYear(), now.getMonth() - 11, 1)),
        endStr: toDateStr(new Date(now.getFullYear(), now.getMonth() + 1, 0)),
      };
    }
    if (historyRange === 'fiscal') {
      const sy = now.getMonth() >= 6 ? now.getFullYear() - 1 : now.getFullYear() - 2;
      return { startStr: `${sy}-07-01`, endStr: `${sy + 1}-06-30` };
    }
    // custom
    return {
      startStr: historyCustomDates.start || '0000-01-01',
      endStr: historyCustomDates.end || '9999-12-31',
    };
  }, [historyRange, historyCustomDates]);

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
      const currentHolding = cumulativeDeposit + cumulativeDividends - cumulativeSellAmount - cumulativeWithdrawn;
      const activeInvestment = cumulativeDeposit - cumulativeWithdrawn - cumulativeSellAmount;
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
      const rate = f.currency === 'USD' ? (f.exchangeRate || 110) : 1;
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
        fdrHoldingTotal += bal * rate;
        fdrProfitRangeTotal += (fProfitRange - fChargeRange) * rate;
      }
    });

    const fdrCurrentHolding = fdrHoldingTotal;
    const fdrActiveInvested = fdrActiveTotal;
    const fdrTotalInvested  = fdrInvestedTotal;
    const fdrTotalProfit    = fdrProfitRangeTotal;

    // ── ONLINE INVESTMENTS ──
    const olStats = state.onlineInvestments.map(inv => {
      const rate = inv.currency === 'USD' ? 110 : 1;
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
      let isActive = false;
      let holding = 0;
      if (beforeEnd(inv.issueDate)) {
        const maturityDate = new Date(inv.issueDate);
        maturityDate.setFullYear(maturityDate.getFullYear() + inv.durationYears);
        const maturityStr = toStr2(maturityDate);
        const closingStr = (inv as any).closingDate;
        const effectiveClosingStr = closingStr || maturityStr;
        if (effectiveClosingStr >= effectiveEndStr) {
          isActive = true;
          holding = inv.principalAmount;
        } else {
          const withdrawBalance = (inv as any).withdrawBalance;
          if (withdrawBalance === undefined || withdrawBalance < inv.principalAmount) {
            isActive = true;
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
        active: isActive ? inv.principalAmount : 0,
        investedRange: inRange(inv.issueDate) ? inv.principalAmount : 0,
        profitRange
      };
    });

    const skCurrentHolding = skStats.reduce((sum, s) => sum + s.holding, 0);
    const skActiveInvested = skStats.reduce((sum, s) => sum + s.active, 0);
    const skTotalInvested  = skStats.reduce((sum, s) => sum + s.investedRange, 0);
    const skTotalProfit    = skStats.reduce((sum, s) => sum + s.profitRange, 0);

    // ── DSE ──
    const typePriority: Record<string, number> = { Deposit: 1, Buy: 2, Dividend: 3, Sell: 4, Charge: 5 };
    const dseSorted = [...dseTxns].sort((a, b) => {
      const dateDiff = new Date(a.date).getTime() - new Date(b.date).getTime();
      if (dateDiff !== 0) return dateDiff;
      return (typePriority[a.type] || 99) - (typePriority[b.type] || 99);
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
          const costOfSold = t.qty * avgCost;
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

    const totalActiveInvestment = mfActiveInvested + fdrActiveInvested + olActiveInvested + skActiveInvested + dseActiveInvested;
    const totalNetWorth = mfCurrentHolding + fdrCurrentHolding + olCurrentHolding + skCurrentHolding + dseCurrentHolding;
    const totalInvested = mfTotalInvested  + fdrTotalInvested  + olTotalInvested  + skTotalInvested  + dseTotalInvested;
    const totalProfit   = mfTotalProfit    + fdrTotalProfit    + olTotalProfit    + skTotalProfit    + dseTotalProfit;

    return { totalNetWorth, totalActiveInvestment, totalInvested, totalProfit };
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

    // ── Last year stats ──
    const lastYear = computeStats(lyStartStr, lyEndStr, dseTransactions);

    // ── YoY % change ──
    const yoyPct = (curr: number, prev: number): number | null => {
      if (prev === 0) return null;
      return ((curr - prev) / Math.abs(prev)) * 100;
    };

    const netWorthYoy    = yoyPct(current.totalNetWorth, lastYear.totalNetWorth);
    const activeInvYoy   = yoyPct(current.totalActiveInvestment, lastYear.totalActiveInvestment);
    const totalInvYoy    = yoyPct(current.totalInvested, lastYear.totalInvested);
    const totalProfitYoy = yoyPct(current.totalProfit, lastYear.totalProfit);

    // ── Pie chart ──
    const inRange2 = (date: string) => date >= startStr && date <= endStr;
    const beforeEnd2 = (date: string) => date <= endStr;

    const typePriority: Record<string, number> = { Deposit: 1, Buy: 2, Dividend: 3, Sell: 4, Charge: 5 };
    const dseSorted = [...dseTransactions].sort((a, b) => {
      const dateDiff = new Date(a.date).getTime() - new Date(b.date).getTime();
      if (dateDiff !== 0) return dateDiff;
      return (typePriority[a.type] || 99) - (typePriority[b.type] || 99);
    });

    // Re-derive holdings for pie chart (current period)
    const dHoldings2: Record<string, { qty: number; totalCost: number }> = {};
    let dT_Dep2 = 0, dT_Wd2 = 0, dT_Div2 = 0, dT_Chg2 = 0, dT_PnL2 = 0;
    dseSorted.forEach(t => {
      const key = `${t.portfolio}|${t.ticker}`;
      if (t.type === 'Deposit') { if (beforeEnd2(t.date)) dT_Dep2 += Math.abs(t.total); }
      else if (t.type === 'Withdrawal') { if (beforeEnd2(t.date)) dT_Wd2 += Math.abs(t.total); }
      else if (t.type === 'Charge') { if (beforeEnd2(t.date)) dT_Chg2 += Math.abs(t.total); }
      else if (t.type === 'Dividend') { if (beforeEnd2(t.date)) dT_Div2 += Math.abs(t.total); }
      else if (t.type === 'Buy') {
        if (!beforeEnd2(t.date)) return;
        if (!dHoldings2[key]) dHoldings2[key] = { qty: 0, totalCost: 0 };
        dHoldings2[key].qty += t.qty;
        dHoldings2[key].totalCost += t.total;
      } else if (t.type === 'Sell') {
        if (!beforeEnd2(t.date)) return;
        if (!dHoldings2[key]) dHoldings2[key] = { qty: 0, totalCost: 0 };
        const h = dHoldings2[key];
        if (h.qty > 0) {
          const avgCost = h.totalCost / h.qty;
          const costOfSold = t.qty * avgCost;
          dT_PnL2 += t.total - costOfSold;
          h.qty = Math.max(0, h.qty - t.qty);
          h.totalCost = Math.max(0, h.totalCost - costOfSold);
        }
      }
    });
    const dStockHoldingCost2 = Object.values(dHoldings2).reduce((sum, h) => sum + h.totalCost, 0);
    const dCashBalance2 = dT_Dep2 - dT_Wd2 - dStockHoldingCost2 + dT_PnL2 + dT_Div2 - dT_Chg2;
    const dseCurrentHolding2 = dStockHoldingCost2 + dCashBalance2;

    const mfCurrentHolding2 = state.mutualFunds.reduce((sum, f) => {
      return sum + f.transactions.reduce((tSum, t) => {
        if (!beforeEnd2(t.date)) return tSum;
        if ((t.type === 'Buy' || t.type === 'Dividend') && !t.isDividend) return tSum + (t.sipAmount || 0);
        if (t.isDividend) return tSum + (t.sipAmount || 0);
        if (t.type === 'Sell' || t.type === 'Withdrawal') return tSum - t.amount;
        return tSum;
      }, 0);
    }, 0);

    const fdrCurrentHolding2 = state.fdrs.reduce((sum, f) => {
      if (!beforeEnd2(f.investmentDate)) return sum;
      const rate = f.currency === 'USD' ? (f.exchangeRate || 110) : 1;
      const isClosedByEnd = f.status === 'Closed' && f.closingDate && f.closingDate <= endStr;
      const addedProfit = f.transactions.filter(t => t.type === 'Profit' && t.handling === 'Added' && beforeEnd2(t.date)).reduce((p, t) => p + t.amount, 0);
      const charges = f.transactions.filter(t => t.type === 'Charge' && beforeEnd2(t.date)).reduce((c, t) => c + t.amount, 0);
      const computedBal = f.principal + addedProfit - charges;
      const bal = (isClosedByEnd && f.withdrawBalance != null) ? Math.max(0, computedBal - f.withdrawBalance) : computedBal;
      return sum + (bal * rate);
    }, 0);

    const olCurrentHolding2 = state.onlineInvestments.reduce((sum, inv) => {
      const rate = inv.currency === 'USD' ? 110 : 1;
      if (!beforeEnd2(inv.investmentDate)) return sum;
      const status = getOnlineStatus(inv);
      if (status !== 'Completed') return sum + inv.amount * rate;
      const lastInst = inv.installments?.[inv.installments.length - 1];
      const completionDate = lastInst ? (lastInst.actualDate || lastInst.date) : inv.maturityDate;
      if (completionDate > endStr) return sum + inv.amount * rate;
      return sum;
    }, 0);

    const skCurrentHolding2 = state.sukuks.reduce((sum, s) => {
      if (!beforeEnd2(s.issueDate)) return sum;
      const maturityDate = new Date(s.issueDate);
      maturityDate.setFullYear(maturityDate.getFullYear() + s.durationYears);
      const maturityStr = toDateStr(maturityDate);
      const closingStr = (s as any).closingDate as string | undefined;
      const effectiveClosingStr = closingStr || maturityStr;
      if (effectiveClosingStr >= endStr) return sum + s.principalAmount;
      const withdrawBalance = (s as any).withdrawBalance as number | undefined;
      if (withdrawBalance !== undefined && withdrawBalance >= s.principalAmount) return sum;
      return sum + s.principalAmount - (withdrawBalance || 0);
    }, 0);

    const pieData = [
      { name: 'DSE', value: dseCurrentHolding2 },
      { name: 'Mutual Funds', value: mfCurrentHolding2 },
      { name: 'Fixed Deposits', value: fdrCurrentHolding2 },
      { name: 'Online', value: olCurrentHolding2 },
      { name: 'Sukuk', value: skCurrentHolding2 },
    ].filter(item => item.value > 0);

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
        const rate = f.currency === 'USD' ? (f.exchangeRate || 110) : 1;
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
        const rate = f.currency === 'USD' ? (f.exchangeRate || 110) : 1;
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
        const rate = inv.currency === 'USD' ? 110 : 1;
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
        income += state.fdrs.reduce((sum, f) => sum + f.transactions.filter(t => t.type === 'Profit' && t.handling !== 'Added' && t.date >= monthStart && t.date <= monthEnd).reduce((s, t) => s + t.amount, 0), 0);
        income += state.sukuks.reduce((sum, s) => sum + (s.installments || []).filter(inst => inst.isPaid && (inst.actualDate || inst.date) >= monthStart && (inst.actualDate || inst.date) <= monthEnd).reduce((s2, inst) => s2 + (inst.actualAmount || inst.amount), 0), 0);
        income += state.onlineInvestments.reduce((sum, inv) => sum + (inv.installments || []).filter(inst => inst.isPaid && (inst.actualDate || inst.date) >= monthStart && (inst.actualDate || inst.date) <= monthEnd).reduce((s2, inst) => s2 + (inst.actualAmount || inst.amount), 0), 0);
        income += dseTransactions.filter((t: any) => (t.type === 'Dividend') && t.date >= monthStart && t.date <= monthEnd).reduce((s: number, t: any) => s + Math.abs(t.total), 0);

        let expense = 0;
        expense += state.mutualFunds.reduce((sum, f) => sum + f.transactions.filter(t => t.type === 'Buy' && !t.isDividend && t.date >= monthStart && t.date <= monthEnd).reduce((s, t) => s + (t.sipAmount || t.amount || 0), 0), 0);
        expense += state.onlineInvestments.filter(inv => inv.investmentDate >= monthStart && inv.investmentDate <= monthEnd).reduce((sum, inv) => sum + inv.amount, 0);
        expense += state.sukuks.filter(s => s.investmentDate >= monthStart && s.investmentDate <= monthEnd).reduce((sum, s) => sum + s.principalAmount, 0);
        expense += state.fdrs.filter(f => f.investmentDate >= monthStart && f.investmentDate <= monthEnd).reduce((sum, f) => {
          const rate = f.currency === 'USD' ? (f.exchangeRate || 110) : 1;
          return sum + (f.principal * rate);
        }, 0);
        expense += dseTransactions.filter((t: any) => t.type === 'Buy' && t.date >= monthStart && t.date <= monthEnd).reduce((s: number, t: any) => s + Math.abs(t.total), 0);

        points.push({ name: label, income, expense });
      }
      return points;
    })();

    return {
      totalNetWorth: current.totalNetWorth,
      totalActiveInvestment: current.totalActiveInvestment,
      totalInvested: current.totalInvested,
      totalProfit: current.totalProfit,
      netWorthYoy,
      activeInvYoy,
      totalInvYoy,
      totalProfitYoy,
      pieData,
      netWorthTrend: netWorthTrends,
      barData,
      seriesMetadata,
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
              {historyRange === 'all' ? 'Overall' : historyRange === 'last12m' ? 'Last 12M' : historyRange === 'fiscal' ? 'Fiscal' : 'Custom'}
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
              <div className="fixed inset-0 z-40" onClick={() => setIsRangeMenuOpen(false)} />
              <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95 backdrop-blur-md">
                {['all','last12m','fiscal','custom'].map((id) => (
                  <button
                    key={id}
                    onClick={() => handleRangeChange(id as any)}
                    className={cn(
                      "w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase",
                      historyRange === id
                        ? "bg-teal-400 text-slate-950"
                        : "text-slate-300 hover:bg-slate-800"
                    )}
                  >
                    {id === 'all' ? 'Overall' : id === 'last12m' ? 'Last 12M' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
                  </button>
                ))}
              </div>
            </>
          )}
        </div>

        {/* Desktop Tabs */}
        <div className="hidden sm:flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1">
          {['all','last12m','fiscal','custom'].map(id => (
            <button
              key={id}
              onClick={() => handleRangeChange(id as any)}
              className={cn(
                "px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all border",
                historyRange === id
                  ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20"
                  : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white"
              )}
            >
              {id === 'all' ? 'Overall' : id === 'last12m' ? 'Last 12M' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
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

    {/* LINE 2 (Mobile): Custom Dates */}
    {historyRange === 'custom' && (
      <div className="flex items-center justify-center sm:justify-start gap-1.5 px-1 animate-in fade-in slide-in-from-top-2 sm:slide-in-from-left-2 duration-300 w-full sm:w-auto">
        
        <button onClick={() => navigateCustomMonth(-1)} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400">
          <ChevronLeft size={14} />
        </button>

        <input
          type="date"
          value={historyCustomDates.start}
          onChange={(e) => setHistoryCustomDates(prev => ({ ...prev, start: e.target.value }))}
          className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none focus:border-teal-400/50 uppercase flex-1 sm:flex-none"
        />

        <span className="text-slate-700 text-[10px] font-bold">–</span>

        <input
          type="date"
          value={historyCustomDates.end}
          onChange={(e) => setHistoryCustomDates(prev => ({ ...prev, end: e.target.value }))}
          className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none focus:border-teal-400/50 uppercase flex-1 sm:flex-none"
        />

        <button onClick={() => navigateCustomMonth(1)} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400">
          <ChevronRight size={14} />
        </button>
      </div>
    )}
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
                subtitle="Total Current Asset"
                value={stats.totalNetWorth}
                trend={formatYoy(stats.netWorthYoy)}
                trendLabel="vs Last Year"
                icon={DollarSign}
                color="teal"
                hideFooter
              />
              <SummaryCard
                className="col-span-1"
                title="Active Investment"
                subtitle="Total Active Investment"
                value={stats.totalActiveInvestment}
                trend={formatYoy(stats.activeInvYoy)}
                trendLabel="vs Last Year"
                icon={Wallet}
                color="blue"
                hideFooter
              />
              <SummaryCard
                className="col-span-1"
                title="Total Invested"
                subtitle="Total Investment"
                value={stats.totalInvested}
                trend={formatYoy(stats.totalInvYoy)}
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
              lineDataMap={stats.netWorthTrend}
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
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            inheritedData={inheritedAddData}
            onClearInheritedData={() => setInheritedAddData(null)}
            onTitleChange={setOnlineInvestmentsTitle}
            onNavigateToModule={handleNavigateToModule}
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
            onClearInheritedData={() => setInheritedAddData(null)}
            onTitleChange={setMutualFundsTitle}
            onNavigateToModule={handleNavigateToModule}

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
            triggerAdd={triggerAdd}
            setTriggerAdd={setTriggerAdd}
            inheritedData={inheritedAddData}
            onClearInheritedData={() => setInheritedAddData(null)}
            onTitleChange={setFixedDepositsTitle}
            onNavigateToModule={handleNavigateToModule}
            
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
            onClearInheritedData={() => setInheritedAddData(null)}
            onTitleChange={setSukukTitle}
            onNavigateToModule={handleNavigateToModule}
            
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
            onClearInheritedData={() => setInheritedAddData(null)}
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
                DASHBOARD {historyRange !== 'all' && <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">/ {new Date(startStr).toLocaleString('default', { month: 'short', year: 'numeric' })} - {new Date(endStr).toLocaleString('default', { month: 'short', year: 'numeric' })}</span>}
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