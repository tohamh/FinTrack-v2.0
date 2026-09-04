/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Card, Button, Modal, Input, Checkbox, Select } from '../ui/BaseComponents';
import { ConfirmDialog } from '../ui/ConfirmDialog';
import { MutualFund, MutualFundTransaction, OnlineInvestmentStatus, InvestmentFrequency } from '../../types';
import { formatBDT, formatDate, cn, formatNumber } from '../../utils/formatters';
import * as XLSX from 'xlsx';
import { Plus, MoreVertical, Calendar, TrendingUp, Search, CheckCircle2, Briefcase, Settings, Download, Upload, FileSpreadsheet, XCircle, DollarSign, Coins, Percent, ArrowRight, Edit2, Trash2, ArrowUpDown, ArrowUp, ArrowDown, Filter, Clock, ChevronLeft, ChevronRight, ChevronDown, Wallet, Info } from 'lucide-react';

interface MutualFundsModuleProps {
  investments: MutualFund[];
  onAdd: (investment: Omit<MutualFund, 'id'>) => void;
  onUpdate: (id: string, updates: Partial<MutualFund>) => void;
  onDelete: (id: string) => void;
  onBatchDelete: (ids: string[]) => void;
  onBatchAdd?: (investments: Omit<MutualFund, 'id'>[]) => void;
  onReplaceAll?: (investments: Omit<MutualFund, 'id'>[]) => void;
  triggerAdd?: boolean;
  setTriggerAdd?: (val: boolean) => void;
  inheritedData?: { date: string; amount: number; description?: string; linkedTxId?: string; returnModule?: string; [key: string]: any } | null;
  onClearInheritedData?: () => void;
  onTitleChange?: (title: React.ReactNode) => void;
  onNavigateToModule?: (module: string, initialAddData?: any, openModal?: boolean) => void;
  onDeleteLinkedTransfer?: (groupId: string) => void;
  onUpdateLinkedTransfer?: (groupId: string, updates: { date: string; amount: number; description?: string }) => void;
}

// Tooltip component matching DSE Tracker style
const InfoTooltip: React.FC<{ text: string }> = ({ text }) => {
  const [visible, setVisible] = useState(false);
  const [pos, setPos] = useState({ top: 0, left: 0 });
  const btnRef = React.useRef<HTMLButtonElement>(null);

  const show = () => {
    if (btnRef.current) {
      const rect = btnRef.current.getBoundingClientRect();
      setPos({
        top: rect.top + window.scrollY - 8,
        left: rect.left + rect.width / 2,
      });
    }
    setVisible(true);
  };

  return (
    <div className="relative inline-flex items-center">
      <button
        ref={btnRef}
        onMouseEnter={show}
        onMouseLeave={() => setVisible(false)}
        onFocus={show}
        onBlur={() => setVisible(false)}
        className="p-0.5 text-slate-600 hover:text-slate-400 transition-colors rounded"
        aria-label="More info"
        type="button"
      >
        <Info size={11} />
      </button>
      {visible && (
        <div
          style={{
            position: 'fixed',
            top: pos.top,
            left: pos.left,
            transform: 'translate(-50%, -100%)',
            zIndex: 9999,
          }}
          className="w-48 bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-[10px] font-bold text-slate-200 shadow-xl pointer-events-none whitespace-normal text-left leading-relaxed animate-in fade-in zoom-in-95 duration-150"
        >
          {text}
          <div className="absolute left-1/2 -translate-x-1/2 top-full w-0 h-0 border-l-4 border-r-4 border-t-4 border-l-transparent border-r-transparent border-t-slate-700" />
        </div>
      )}
    </div>
  );
};

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

interface SettingsButtonProps {
  isOpen: boolean;
  setIsOpen: (open: boolean) => void;
  onExport: () => void;
  onImport: (e: React.ChangeEvent<HTMLInputElement>) => void;
  onTemplate: () => void;
  onAdd: () => void;
}

const SettingsButton: React.FC<SettingsButtonProps> = ({ isOpen, setIsOpen, onExport, onImport, onTemplate, onAdd }) => {
  return (
    <div className="relative">
      <button 
        onClick={() => setIsOpen(!isOpen)}
        className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/10 hover:bg-teal-300 hover:shadow-teal-400/20"
      >
        <Settings size={14} />
        <span className="hidden sm:inline">Settings</span>
        <ChevronDown size={14} className={cn("opacity-50 transition-transform", isOpen ? "rotate-180" : "")} />
      </button>

      {isOpen && (
        <>
          <div className="fixed inset-0 z-40" onClick={() => setIsOpen(false)} />
          <div className="absolute right-0 mt-2 w-48 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-2 animate-in fade-in zoom-in-95 backdrop-blur-xl transition-all">
            <button 
              onClick={() => { onExport(); setIsOpen(false); }}
              className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
            >
              <Download size={14} className="text-teal-400" />
              EXPORT
            </button>
            <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
              <Upload size={14} className="text-teal-400" />
              IMPORT
              <input type="file" className="hidden" accept=".xlsx,.xls" onChange={onImport} />
            </label>
            <button 
              onClick={() => { onTemplate(); setIsOpen(false); }}
              className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
            >
              <FileSpreadsheet size={14} className="text-teal-400" />
              TEMPLATE
            </button>

            <div className="h-px bg-slate-800/60 my-2" />

            <button 
              onClick={() => { onAdd(); setIsOpen(false); }}
              className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-teal-400 hover:bg-teal-400/10 rounded-lg transition-colors uppercase"
            >
              <Plus size={14} />
              Add Mutual Fund
            </button>
          </div>
        </>
      )}
    </div>
  );
};

export const MutualFundsModule: React.FC<MutualFundsModuleProps> = ({ 
  investments, 
  onAdd, 
  onUpdate,
  onDelete,
  onBatchDelete,
  onBatchAdd,
  onReplaceAll,
  triggerAdd,
  setTriggerAdd,
  inheritedData,
  onClearInheritedData,
  onTitleChange,
  onNavigateToModule,
  onDeleteLinkedTransfer,
  onUpdateLinkedTransfer
}) => {
    const uniqueInvestments = useMemo(() => {
    const seen = new Set();
    return (investments || []).filter(inv => {
      if (!inv || !inv.id) return false;
      if (seen.has(inv.id)) return false;
      seen.add(inv.id);
      return true;
    });
  }, [investments]);

  const isFromExternalModuleRef = useRef(false);
  const returnModuleRef = useRef<string | undefined>(undefined);
  const [inheritedIeGroupId, setInheritedIeGroupId] = useState<string | undefined>(undefined);

  const checkAndReturnToModule = () => {
    if (returnModuleRef.current && onNavigateToModule) {
      const ret = returnModuleRef.current;
      returnModuleRef.current = undefined;
      onNavigateToModule(ret, undefined, false);
    }
  };

  const [sortBy, setSortBy] = useState<'date' | 'amount' | 'amc' | 'fundName' | 'sipAmount' | 'units'>('date');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [searchQuery, setSearchQuery] = useState('');
  const [editingFund, setEditingFund] = useState<MutualFund | null>(null);
  const [editingTransaction, setEditingTransaction] = useState<{ fundId: string, transaction: any } | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isFundModalOpen, setIsFundModalOpen] = useState(false);
  const [activeMenuId, setActiveMenuId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isSettingsMenuOpen, setIsSettingsMenuOpen] = useState(false);
  const [isSortMenuOpen, setIsSortMenuOpen] = useState(false);
  const [selectedTypes, setSelectedTypes] = useState<string[]>([]);
  const [isFilterMenuOpen, setIsFilterMenuOpen] = useState(false);
  const [notification, setNotification] = useState<{ type: 'success' | 'error', message: string } | null>(null);
  const [confirmState, setConfirmState] = useState<{ 
    isOpen: boolean; 
    title: string; 
    message: string; 
    onConfirm: () => void; 
    onCancel?: () => void;
    variant?: 'danger' | 'warning' | 'info' | 'severe' | 'critical';
    confirmLabel?: string;
    cancelLabel?: string;
    details?: (string | null | undefined)[];
  } | null>(null);

  const closeConfirm = () => setConfirmState(prev => prev ? { ...prev, isOpen: false } : null);
  // ─── Date Range Configuration (same range selector as Income & Expense) ───
  // Mutual Funds intentionally opens on Custom by default.
  const [historyRange, setHistoryRange] = useState<'this' | 'fiscal' | 'custom'>('custom');
  const [historyCustomDates, setHistoryCustomDates] = useState(() => {
    return {
      start: '2025-01-01',
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
        start: '2025-01-01',
        end: getTodayStr(),
      });
    }
    setHistoryRange(newRange);
  };

  const [isRangeMenuOpen, setIsRangeMenuOpen] = useState(false);
  const [isTimeMenuOpen, setIsTimeMenuOpen] = useState(false);

  // Helper used by the selector to display the selected month.
  const formatHistoryThisMonthLabel = (d: Date) => {
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return `${months[d.getMonth()]} ${d.getFullYear()}`;
  };

  // Derive active date range for dashboard stats.
  // This uses the same This Month / Fiscal / Custom semantics as Income & Expense.
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

  useEffect(() => {
    if (onTitleChange) {
      const { start, end } = rangeDates;
      const startStr = start.toLocaleString('default', { month: 'short', year: 'numeric' });
      const endStr = end.toLocaleString('default', { month: 'short', year: 'numeric' });

      onTitleChange(
        <span className="flex items-center gap-2">
          MUTUAL FUNDS
          <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">
            / {startStr} - {endStr}
          </span>
        </span>
      );
    }
  }, [rangeDates, onTitleChange, historyRange]);

  // Time-based filtering state for transaction table
  const [timeFilterMode, setTimeFilterMode] = useState<'6months' | '1year' | 'custom'>('6months');
  const [timeOffset, setTimeOffset] = useState(0);
  const [customRange, setCustomRange] = useState({ start: '', end: '' });

  useEffect(() => {
    if (triggerAdd) {
      const isExternalTrigger = !!inheritedData;
      isFromExternalModuleRef.current = isExternalTrigger;
      returnModuleRef.current = inheritedData?.returnModule;
      setInheritedIeGroupId(isExternalTrigger ? inheritedData?.linkedTxId : undefined);
      const inheritedDate = inheritedData?.date || new Date().toISOString().split('T')[0];
      const inheritedAmt = (inheritedData?.amount !== undefined && inheritedData.amount !== null && inheritedData.amount !== 0) ? String(inheritedData.amount) : '0';
      const inheritedDesc = inheritedData?.description || '';

      // Match fund by description if provided
      let matchedFundId = '';
      if (inheritedDesc) {
        const descLower = inheritedDesc.toLowerCase();
        const matched = uniqueInvestments.find(f => 
          (f.fullName && descLower.includes(f.fullName.toLowerCase())) ||
          (f.amc && descLower.includes(f.amc.toLowerCase())) ||
          (f.name && descLower.includes(f.name.toLowerCase()))
        );
        if (matched) matchedFundId = matched.id;
      }

      // Check if there is an existing linked transaction (by linkedTxId / linkedIeGroupId or matching fund/date)
      let existingLinkedFund: MutualFund | undefined;
      let existingLinkedTx: MutualFundTransaction | undefined;

      if (inheritedData?.linkedTxId) {
        for (const fund of uniqueInvestments) {
          const t = fund.transactions.find(tx => tx.linkedIeGroupId === inheritedData.linkedTxId);
          if (t) {
            existingLinkedFund = fund;
            existingLinkedTx = t;
            break;
          }
        }
      }

      if (!existingLinkedTx && matchedFundId && inheritedData?.linkedTxId) {
        const targetFund = uniqueInvestments.find(f => f.id === matchedFundId);
        if (targetFund) {
          const t = targetFund.transactions.find(tx => tx.linkedIeGroupId === inheritedData.linkedTxId || tx.date === inheritedDate);
          if (t) {
            existingLinkedFund = targetFund;
            existingLinkedTx = t;
          }
        }
      }

      if (existingLinkedFund && existingLinkedTx) {
        // Edit mode for existing investment: preserve Units, NAV, fund and only update with the new amount
        setEditingTransaction({ fundId: existingLinkedFund.id, transaction: existingLinkedTx });
        
        let sMode: 'Deposit' | 'Dividend' | 'Surrender' = 'Deposit';
        let aMode: 'Buy' | 'Withdrawal' = 'Buy';

        if (existingLinkedTx.isWithdrawal) aMode = 'Withdrawal';
        else if (existingLinkedTx.type === 'Sell') sMode = 'Surrender';
        else if (existingLinkedTx.isDividend) sMode = 'Dividend';
        else if ((existingLinkedTx.sipAmount || 0) > 0) sMode = 'Deposit';

        setSourceMode(sMode);
        setActionMode(aMode);

        const newSipOrSourceAmt = (inheritedAmt !== '0') ? inheritedAmt : (existingLinkedTx.sipAmount?.toString() || '');

        setTransactionFormData({
          fundId: existingLinkedFund.id,
          sourceDate: inheritedDate || existingLinkedTx.pullingDate || existingLinkedTx.date,
          sourceAmount: newSipOrSourceAmt,
          sourceUnits: existingLinkedTx.type === 'Sell' ? existingLinkedTx.units.toString() : '',
          sourceNav: existingLinkedTx.type === 'Sell' ? existingLinkedTx.nav.toString() : '',
          actionDate: inheritedDate || existingLinkedTx.date,
          actionAmount: existingLinkedTx.isWithdrawal ? (inheritedAmt !== '0' ? inheritedAmt : existingLinkedTx.amount.toString()) : '',
          actionUnits: (existingLinkedTx.type === 'Buy' || existingLinkedTx.type === 'Dividend') ? existingLinkedTx.units.toString() : '',
          actionNav: (existingLinkedTx.type === 'Buy' || existingLinkedTx.type === 'Dividend') ? existingLinkedTx.nav.toString() : '',
          date: inheritedDate || existingLinkedTx.date,
          type: existingLinkedTx.type,
          units: existingLinkedTx.units.toString(),
          nav: existingLinkedTx.nav.toString(),
          amount: existingLinkedTx.amount ? existingLinkedTx.amount.toString() : inheritedAmt,
          sipAmount: newSipOrSourceAmt,
          pullingDate: (existingLinkedTx.pullingDate || inheritedDate || '').toString()
        });
      } else {
        // New investment creation
        setEditingTransaction(null);
        const fundObj = uniqueInvestments.find(f => f.id === matchedFundId);
        const isWithdrawal = inheritedData?.targetModule === 'mutual-funds-withdrawal';
        const isDividend = inheritedData?.targetModule === 'mutual-funds-dividend';

        setTransactionFormData({
          fundId: matchedFundId,
          sourceDate: inheritedDate,
          sourceAmount: (inheritedAmt !== '0') ? inheritedAmt : (fundObj ? fundObj.sipAmount.toString() : '0'),
          sourceUnits: '0',
          sourceNav: '0',
          actionDate: inheritedDate,
          actionAmount: isWithdrawal ? ((inheritedAmt !== '0') ? inheritedAmt : '0') : '0',
          actionUnits: '0',
          actionNav: '0',
          date: inheritedDate,
          sipAmount: (inheritedAmt !== '0') ? inheritedAmt : (fundObj ? fundObj.sipAmount.toString() : '0'),
          pullingDate: inheritedDate,
          units: '0',
          nav: '0',
          amount: inheritedAmt,
          type: isDividend ? 'Dividend' : (isWithdrawal ? 'Withdrawal' : 'Buy')
        });
        setSourceMode(isDividend ? 'Dividend' : 'Deposit');
        setActionMode(isWithdrawal ? 'Withdrawal' : 'Buy');
      }

      setIsModalOpen(true);
      setTriggerAdd?.(false);
      onClearInheritedData?.();
    }
  }, [triggerAdd, setTriggerAdd, inheritedData, onClearInheritedData, uniqueInvestments]);

  const handleSort = (key: any) => {
    if (sortBy === key) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortBy(key);
      setSortOrder('desc');
    }
  };

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 5000);
  };

  const exportData = () => {
    const rows: any[] = [];
    uniqueInvestments.forEach(inv => {
      if (inv.transactions.length === 0) {
        rows.push({
          'AMC': inv.amc,
          'Full Fund Name': inv.fullName,
          'Starting Month': inv.startingMonth,
          'Preferred Pulling Day': inv.pullingDate,
          'Base SIP Amount': inv.sipAmount,
          'Actual Pulling Date': '',
          'SIP/Dividend Amount': '',
          'Addition Type': '',
          'Transaction Date': '',
          'Transaction Type': '',
          'Units': '',
          'NAV': '',
          'Amount': ''
        });
      } else {
        inv.transactions.forEach(t => {
          rows.push({
            'AMC': inv.amc,
            'Full Fund Name': inv.fullName,
            'Starting Month': inv.startingMonth,
            'Preferred Pulling Day': inv.pullingDate,
            'Base SIP Amount': inv.sipAmount,
            'Actual Pulling Date': t.pullingDate || '',
            'SIP/Dividend Amount': t.sipAmount || '',
            'Addition Type': t.isDividend ? 'Dividend' : (t.type === 'Withdrawal' ? '' : 'Deposit'),
            'Transaction Date': t.date,
            'Transaction Type': t.type,
            'Units': t.units || '',
            'NAV': t.nav || '',
            'Amount': t.amount || ''
          });
        });
      }
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Mutual Funds");
    XLSX.writeFile(wb, `MF_Export_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('success', 'Export successful!');
  };

  const importData = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setConfirmState({
      isOpen: true,
      title: 'Confirm Replace',
      message: 'Are you sure you want to replace all existing Mutual Fund data with the imported data? This action cannot be undone.',
      confirmLabel: 'Replace All',
      variant: 'warning',
      onConfirm: () => {
        closeConfirm();
        const reader = new FileReader();
        reader.onload = (event) => {
          try {
            const data = new Uint8Array(event.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];
            
            const newFundsMap = new Map<string, MutualFund>();
            
            jsonData.forEach(row => {
              const amcValue = String(row['AMC'] || '').trim();
              if (!amcValue) return;

              const fundMetadata = {
                fullName: row['Full Fund Name'] ? String(row['Full Fund Name']) : undefined,
                startingMonth: row['Starting Month'] ? String(row['Starting Month']) : undefined,
                pullingDate: row['Preferred Pulling Day'] ? Number(row['Preferred Pulling Day']) : undefined,
                sipAmount: row['Base SIP Amount'] ? Number(row['Base SIP Amount']) : undefined,
              };

              if (!newFundsMap.has(amcValue)) {
                newFundsMap.set(amcValue, {
                  id: Math.random().toString(36).substr(2, 9),
                  name: amcValue,
                  amc: amcValue,
                  fullName: fundMetadata.fullName || amcValue,
                  startingMonth: fundMetadata.startingMonth || new Date().toISOString().slice(0, 7),
                  investmentDate: (fundMetadata.startingMonth || new Date().toISOString().slice(0, 7)) + '-01',
                  amount: 0,
                  currency: 'BDT' as const,
                  pullingDate: fundMetadata.pullingDate || 5,
                  sipAmount: fundMetadata.sipAmount || 0,
                  transactions: []
                });
              } else {
                const nf = newFundsMap.get(amcValue)!;
                if (fundMetadata.fullName) nf.fullName = fundMetadata.fullName;
                if (fundMetadata.startingMonth) nf.startingMonth = fundMetadata.startingMonth;
                if (fundMetadata.pullingDate) nf.pullingDate = fundMetadata.pullingDate;
                if (fundMetadata.sipAmount) nf.sipAmount = fundMetadata.sipAmount;
              }
              
              const targetTransactions = newFundsMap.get(amcValue)!.transactions;
              
              if (row['Transaction Date']) {
                const type = row['Transaction Type'] || 'Buy';
                const units = Number(row['Units'] || 0);
                const nav = Number(row['NAV'] || 0);
                let amount = Number(row['Amount'] || 0);
                
                if (amount === 0 && units > 0 && nav > 0) {
                  amount = units * nav;
                }

                const transactionId = Math.random().toString(36).substr(2, 9);
                targetTransactions.push({
                  id: transactionId,
                  date: row['Transaction Date'],
                  type: type as any,
                  units: units,
                  nav: nav,
                  amount: amount,
                  isDividend: row['Addition Type'] === 'Dividend',
                  isWithdrawal: type === 'Withdrawal',
                  pullingDate: row['Actual Pulling Date'] ? String(row['Actual Pulling Date']) : undefined,
                  sipAmount: row['SIP/Dividend Amount'] ? Number(row['SIP/Dividend Amount']) : undefined
                });
              }
            });
            
            const importedFunds = Array.from(newFundsMap.values());
            if (onReplaceAll && importedFunds.length > 0) {
              onReplaceAll(importedFunds);
              showNotification('success', `Successfully replaced data with ${importedFunds.length} mutual funds!`);
            } else if (onBatchAdd && importedFunds.length > 0) {
              onBatchAdd(importedFunds);
              showNotification('success', `Imported data for ${importedFunds.length} mutual funds!`);
            } else {
              showNotification('error', 'No valid data found in file.');
            }
          } catch (err) {
            console.error('Import failed', err);
            showNotification('error', 'Failed to import file.');
          }
        };
        reader.readAsArrayBuffer(file);
      },
      onCancel: () => {
        closeConfirm();
        e.target.value = '';
      }
    });
  };

  const [sourceMode, setSourceMode] = useState<'Deposit' | 'Dividend' | 'Surrender'>('Deposit');
  const [actionMode, setActionMode] = useState<'Buy' | 'Withdrawal'>('Buy');
  
  const [fundFormData, setFundFormData] = useState({
    amc: '',
    fullName: '',
    startingMonth: new Date().toISOString().slice(0, 7),
    pullingDate: '5',
    sipAmount: '',
  });

  const [transactionFormData, setTransactionFormData] = useState({
    fundId: '',
    sourceDate: new Date().toISOString().split('T')[0],
    sourceAmount: '0',
    sourceUnits: '0',
    sourceNav: '0',
    actionDate: new Date().toISOString().split('T')[0],
    actionAmount: '0',
    actionUnits: '0',
    actionNav: '0',
    date: new Date().toISOString().split('T')[0],
    sipAmount: '0',
    pullingDate: '',
    units: '0',
    nav: '0',
    amount: '0',
    type: 'Buy' as 'Buy' | 'Sell' | 'Dividend' | 'Withdrawal'
  });

  // Flatten transactions for the list view
  const allTransactions = useMemo(() => {
    const flat: any[] = [];
    uniqueInvestments.forEach(inv => {
      let balance = 0;
      const sortedTransactions = [...inv.transactions].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
      
      sortedTransactions.forEach(t => {
        if (t.type === 'Buy' || t.type === 'Dividend') balance += t.units;
        else if (t.type === 'Sell') balance -= t.units;
        
        flat.push({
          ...t,
          fundId: inv.id,
          fundName: inv.amc,
          amc: inv.amc,
          sipAmount: t.sipAmount || inv.sipAmount,
          pullingDate: t.pullingDate || inv.pullingDate,
          balance: balance,
          uniqueId: `${inv.id}-${t.id}`
        });
      });
    });
    return flat;
  }, [investments]);

  const filtered = useMemo(() => {
    let result = [...allTransactions];

    if (searchQuery.trim()) {
      const search = searchQuery.toLowerCase();
      result = result.filter(t => 
        t.fundName.toLowerCase().includes(search) || 
        t.amc.toLowerCase().includes(search)
      );
    }

    if (selectedTypes.length > 0) {
      result = result.filter(t => {
        return selectedTypes.some(type => {
          if (type === 'Deposit') return t.type === 'Buy' && !t.isDividend;
          if (type === 'Dividend') return t.isDividend || t.type === 'Dividend';
          if (type === 'Buy') return t.type === 'Buy';
          if (type === 'Sell') return t.type === 'Sell';
          if (type === 'Withdrawal') return t.type === 'Withdrawal';
          return false;
        });
      });
    }

    if (timeFilterMode === '6months' || timeFilterMode === '1year') {
      const now = new Date();
      const monthsCount = timeFilterMode === '6months' ? 6 : 12;
      const currentOffsetMonths = timeOffset * monthsCount;
      
      const endMonth = now.getMonth() - currentOffsetMonths;
      const rangeEnd = new Date(now.getFullYear(), endMonth + 1, 0);
      rangeEnd.setHours(23, 59, 59, 999);
      
      const startMonth = endMonth - (monthsCount - 1);
      const rangeStart = new Date(now.getFullYear(), startMonth, 1);
      rangeStart.setHours(0, 0, 0, 0);

      result = result.filter(t => {
        const d = new Date(t.date);
        return d >= rangeStart && d <= rangeEnd;
      });
    } else if (timeFilterMode === 'custom') {
      if (customRange.start) {
        const start = new Date(customRange.start);
        start.setHours(0, 0, 0, 0);
        result = result.filter(t => new Date(t.date) >= start);
      }
      if (customRange.end) {
        const end = new Date(customRange.end);
        end.setHours(23, 59, 59, 999);
        result = result.filter(t => new Date(t.date) <= end);
      }
    }

    return result.sort((a, b) => {
      let comparison = 0;
      if (sortBy === 'date') comparison = new Date(b.date).getTime() - new Date(a.date).getTime();
      else if (sortBy === 'amount') comparison = b.amount - a.amount;
      else if (sortBy === 'amc') comparison = a.amc.localeCompare(b.amc);
      else if (sortBy === 'fundName') comparison = a.fundName.localeCompare(b.fundName);
      else if (sortBy === 'sipAmount') comparison = b.sipAmount - a.sipAmount;
      else if (sortBy === 'units') comparison = b.units - a.units;
      
      return sortOrder === 'desc' ? comparison : -comparison;
    });
  }, [allTransactions, sortBy, sortOrder, searchQuery, selectedTypes, timeFilterMode, timeOffset, customRange]);

  const fundStats = useMemo(() => {
    const toDateString = (d: Date) => {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    };

    const startStr = toDateString(rangeDates.start);
    const endStr = toDateString(rangeDates.end);

    return uniqueInvestments.map(f => {
      const totalDeposit = f.transactions.reduce((tSum, t) => {
        if (t.date < startStr || t.date > endStr) return tSum;
        if ((t.type === 'Buy' || t.type === 'Dividend') && !t.isDividend) return tSum + (t.sipAmount || 0);
        return tSum;
      }, 0);

      const cumulativeDeposit = f.transactions.reduce((tSum, t) => {
        if (t.date > endStr) return tSum;
        if ((t.type === 'Buy' || t.type === 'Dividend') && !t.isDividend) return tSum + (t.sipAmount || 0);
        return tSum;
      }, 0);

      const cumulativeCostBasis = f.transactions.reduce((tSum, t) => {
        if (t.date > endStr) return tSum;
        if (t.type === 'Buy' || t.type === 'Dividend') return tSum + t.amount;
        return tSum;
      }, 0);

      const rangeCostBasis = f.transactions.reduce((tSum, t) => {
        if (t.date < startStr || t.date > endStr) return tSum;
        if (t.type === 'Buy' || t.type === 'Dividend') return tSum + t.amount;
        return tSum;
      }, 0);

      const rangeUnitsPurchased = f.transactions.reduce((uSum, t) => {
          if (t.date < startStr || t.date > endStr) return uSum;
          if (t.type === 'Buy' || t.type === 'Dividend') return uSum + t.units;
          return uSum;
      }, 0);

      const totalUnits = f.transactions.reduce((uSum, t) => {
        if (t.date > endStr) return uSum;
        if (t.type === 'Buy' || t.type === 'Dividend') return uSum + t.units;
        if (t.type === 'Sell' || t.type === 'Withdrawal') return uSum - t.units;
        return uSum;
      }, 0);

      const relevantTransactions = f.transactions.filter(t => t.date <= endStr);
      const latestNAV = relevantTransactions.length > 0
        ? relevantTransactions[relevantTransactions.length - 1].nav
        : 0;

      const totalDividends = f.transactions.reduce((dSum, t) => {
        if (t.date < startStr || t.date > endStr) return dSum;
        if (t.isDividend || t.type === 'Dividend') return dSum + (t.sipAmount || 0);
        return dSum;
      }, 0);

      const cumulativeDividends = f.transactions.reduce((dSum, t) => {
        if (t.date > endStr) return dSum;
        if (t.isDividend || t.type === 'Dividend') return dSum + (t.sipAmount || 0);
        return dSum;
      }, 0);

      const totalSellAmount = f.transactions.reduce((sSum, t) => {
        if (t.date < startStr || t.date > endStr) return sSum;
        if (t.type === 'Sell') return sSum + t.amount;
        return sSum;
      }, 0);

      const cumulativeSellAmount = f.transactions.reduce((sSum, t) => {
        if (t.date > endStr) return sSum;
        if (t.type === 'Sell') return sSum + t.amount;
        return sSum;
      }, 0);

      const totalWithdrawn = f.transactions.reduce((wSum, t) => {
        if (t.date < startStr || t.date > endStr) return wSum;
        if (t.type === 'Withdrawal') return wSum + t.amount;
        return wSum;
      }, 0);

      const cumulativeWithdrawn = f.transactions.reduce((wSum, t) => {
        if (t.date > endStr) return wSum;
        if (t.type === 'Withdrawal') return wSum + t.amount;
        return wSum;
      }, 0);

      const cashBalance = cumulativeDeposit + cumulativeDividends + cumulativeSellAmount - cumulativeCostBasis - cumulativeWithdrawn;

      const dividendReturnPercentage = totalDeposit > 0 ? (totalDividends / totalDeposit) * 100 : 0;
      const averageNav = totalUnits > 0 ? cumulativeCostBasis / totalUnits : 0;

      return {
        id: f.id,
        fullName: f.fullName,
        amc: f.amc,
        totalDeposit,
        cumulativeDeposit,
        rangeCostBasis,
        cumulativeCostBasis,
        currentHolding: cumulativeCostBasis,
        totalDividends,
        cumulativeDividends,
        totalWithdrawn,
        cumulativeWithdrawn,
        totalSellAmount,
        cumulativeSellAmount,
        cashBalance,
        totalUnits,
        rangeUnitsPurchased,
        averageNav,
        dividendReturnPercentage,
      };
    });
  }, [uniqueInvestments, rangeDates]);

  const summaryStats = useMemo(() => {
    const totalDeposit = fundStats.reduce((sum, f) => sum + f.totalDeposit, 0);
    const totalDividends = fundStats.reduce((sum, f) => sum + f.totalDividends, 0);

    const totalCumulativeCostBasis = fundStats.reduce((sum, f) => sum + f.cumulativeCostBasis, 0);
    const currentHolding = fundStats.reduce((sum, f) => sum + f.cumulativeCostBasis + f.cashBalance, 0);

    const dividendReturnPercentage = totalDeposit > 0 ? (totalDividends / totalDeposit) * 100 : 0;

    const totalPortfolioUnits = fundStats.reduce((sum, f) => sum + f.totalUnits, 0);
    const totalRangeUnitsPurchased = fundStats.reduce((sum, f) => sum + f.rangeUnitsPurchased, 0);

    const activeCount = uniqueInvestments.length;

    const endStr = rangeDates.end.toISOString().split('T')[0];
    const currentYear = rangeDates.end.getFullYear();
    const currentMonth = rangeDates.end.getMonth();
    const monthStartStr = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}-01`;

    const thisMonthDeposit = uniqueInvestments.reduce((sum, f) => {
      return sum + f.transactions.reduce((tSum, t) => {
        if (t.date < monthStartStr || t.date > endStr) return tSum;
        if ((t.type === 'Buy' || t.type === 'Dividend') && !t.isDividend) return tSum + (t.sipAmount || 0);
        return tSum;
      }, 0);
    }, 0);

    const totalPortfolioCashBalance = fundStats.reduce((sum, f) => sum + f.cashBalance, 0);
    const totalCumulativeDeposit = fundStats.reduce((sum, f) => sum + f.cumulativeDeposit, 0);
    const totalCumulativeWithdrawn = fundStats.reduce((sum, f) => sum + f.cumulativeWithdrawn, 0);
    const totalPortfolioSellAmount = fundStats.reduce((sum, f) => sum + f.totalSellAmount, 0);
    const totalCumulativeSellAmount = fundStats.reduce((sum, f) => sum + f.cumulativeSellAmount, 0);

    const activeInvestment = totalCumulativeDeposit - totalCumulativeWithdrawn;
    
    return {
      totalDeposit,
      totalCumulativeCostBasis,
      currentHolding,
      totalDividends,
      activeCount,
      dividendReturnPercentage,
      totalUnits: totalPortfolioUnits,
      rangeUnitsPurchased: totalRangeUnitsPurchased,
      activeInvestment,
      thisMonthDeposit,
      cashBalance: totalPortfolioCashBalance,
      totalWithdrawn: totalCumulativeWithdrawn,
      totalSellAmount: totalPortfolioSellAmount,
      totalCumulativeSellAmount,
    };
  }, [fundStats, uniqueInvestments, rangeDates]);

  const handleEditFund = (fund: MutualFund) => {
    setEditingFund(fund);
    setFundFormData({
      amc: fund.amc,
      fullName: fund.fullName,
      startingMonth: fund.startingMonth,
      pullingDate: fund.pullingDate.toString(),
      sipAmount: fund.sipAmount.toString(),
    });
    setIsFundModalOpen(true);
  };

  const handleFundSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const payload = {
      amc: fundFormData.amc,
      fullName: fundFormData.fullName,
      startingMonth: fundFormData.startingMonth,
      pullingDate: parseInt(fundFormData.pullingDate) || 5,
      sipAmount: parseFloat(fundFormData.sipAmount) || 0,
    };

    if (editingFund) {
      onUpdate(editingFund.id, payload);
      showNotification('success', 'Mutual Fund updated successfully!');
    } else {
      onAdd({
        ...payload,
        name: fundFormData.fullName,
        investmentDate: fundFormData.startingMonth + '-01',
        currency: 'BDT' as const,
        transactions: [],
      });
      showNotification('success', 'Mutual Fund added successfully!');

      const fundAmount = parseFloat(fundFormData.sipAmount) || 0;
      if (fundAmount > 0 && onNavigateToModule) {
        const pullingDay = String(parseInt(fundFormData.pullingDate) || 5).padStart(2, '0');
        const fundDate = fundFormData.startingMonth ? `${fundFormData.startingMonth}-${pullingDay}` : new Date().toISOString().split('T')[0];
        const fundDesc = fundFormData.fullName || fundFormData.amc || 'Mutual Fund';

        onNavigateToModule('income-expense', {
          date: fundDate,
          amount: fundAmount,
          type: 'Transfer',
          targetModule: 'mutual-funds',
          description: fundDesc
        });
      }
    }
    
    setIsFundModalOpen(false);
    setEditingFund(null);
    checkAndReturnToModule();
    setFundFormData({
      amc: '',
      fullName: '',
      startingMonth: new Date().toISOString().slice(0, 7),
      pullingDate: '5',
      sipAmount: '',
    });
  };

    const handleTransactionSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const fund = uniqueInvestments.find(f => f.id === transactionFormData.fundId);
    if (!fund) return;

    const newTransactions: MutualFundTransaction[] = [];
    const isExternalCreate = !editingTransaction && isFromExternalModuleRef.current;
    const linkGroupId: string | undefined = editingTransaction
      ? (editingTransaction.transaction.linkedIeGroupId || inheritedIeGroupId)
      : (isExternalCreate ? inheritedIeGroupId : crypto.randomUUID());

    if (sourceMode === 'Surrender') {
      const units = parseFloat(transactionFormData.sourceUnits) || 0;
      const nav = parseFloat(transactionFormData.sourceNav) || 0;
      if (units > 0) {
        newTransactions.push({
          id: editingTransaction ? editingTransaction.transaction.id : Math.random().toString(36).substr(2, 9),
          date: transactionFormData.sourceDate,
          type: 'Sell',
          units,
          nav,
          amount: units * nav,
          linkedIeGroupId: linkGroupId,
        });
      }
    }

    if (actionMode === 'Buy') {
      const units = parseFloat(transactionFormData.actionUnits) || 0;
      const nav = parseFloat(transactionFormData.actionNav) || 0;
      const sipAmount = (sourceMode === 'Deposit' || sourceMode === 'Dividend') ? (parseFloat(transactionFormData.sourceAmount) || 0) : 0;
      
      if (units > 0 || sipAmount > 0) {
        newTransactions.push({
          id: editingTransaction ? editingTransaction.transaction.id : Math.random().toString(36).substr(2, 9),
          date: transactionFormData.actionDate || transactionFormData.sourceDate,
          type: sourceMode === 'Dividend' ? 'Dividend' : 'Buy',
          units,
          nav,
          amount: (units * nav > 0) ? (units * nav) : sipAmount,
          isDividend: sourceMode === 'Dividend',
          pullingDate: (sourceMode === 'Deposit' || sourceMode === 'Dividend') ? transactionFormData.sourceDate : undefined,
          sipAmount: (sourceMode === 'Deposit' || sourceMode === 'Dividend') ? sipAmount : undefined,
          linkedIeGroupId: linkGroupId,
        });
      }
    } else if (actionMode === 'Withdrawal') {
      const amount = parseFloat(transactionFormData.actionAmount) || 0;
      const sipAmount = (sourceMode === 'Deposit' || sourceMode === 'Dividend') ? (parseFloat(transactionFormData.sourceAmount) || 0) : 0;
      
      if (amount > 0 || sipAmount > 0) {
        newTransactions.push({
          id: editingTransaction ? editingTransaction.transaction.id : Math.random().toString(36).substr(2, 9),
          date: transactionFormData.actionDate,
          type: 'Withdrawal',
          units: 0,
          nav: 0,
          amount: amount > 0 ? amount : sipAmount,
          isWithdrawal: true,
          pullingDate: (sourceMode === 'Deposit' || sourceMode === 'Dividend') ? transactionFormData.sourceDate : undefined,
          sipAmount: (sourceMode === 'Deposit' || sourceMode === 'Dividend') ? sipAmount : undefined,
          linkedIeGroupId: linkGroupId,
        });
      }
    }

    if (newTransactions.length === 0) return;

    if (editingTransaction) {
      if (editingTransaction.fundId !== fund.id) {
        const oldFund = investments.find(f => f.id === editingTransaction.fundId);
        if (oldFund) {
          onUpdate(oldFund.id, {
            transactions: oldFund.transactions.filter(t => t.id !== editingTransaction.transaction.id)
          });
        }
        onUpdate(fund.id, {
          transactions: [...fund.transactions, newTransactions[0]]
        });
      } else {
        const updatedTransactions = fund.transactions.map(t => 
          t.id === editingTransaction.transaction.id ? newTransactions[0] : t
        );
        onUpdate(fund.id, { transactions: updatedTransactions });
      }
      setEditingTransaction(null);

      if (linkGroupId && onUpdateLinkedTransfer) {
        const updatedAmt = (newTransactions[0].sipAmount && newTransactions[0].sipAmount > 0 && !newTransactions[0].isWithdrawal)
          ? newTransactions[0].sipAmount
          : newTransactions[0].amount;
        onUpdateLinkedTransfer(linkGroupId, {
          date: newTransactions[0].date,
          amount: updatedAmt,
          description: fund.fullName || fund.name || fund.amc
        });
      }
    } else {
      onUpdate(fund.id, {
        transactions: [...fund.transactions, ...newTransactions]
      });

      if (isExternalCreate) {
        isFromExternalModuleRef.current = false;
        setInheritedIeGroupId(undefined);
      } else if (onNavigateToModule && newTransactions.length > 0) {
        const trans = newTransactions[0];
        const pullingDateVal = transactionFormData.sourceDate || (trans.pullingDate ? (typeof trans.pullingDate === 'string' && trans.pullingDate.includes('-') ? trans.pullingDate : trans.date) : undefined);
        const txDate = pullingDateVal || trans.date || transactionFormData.actionDate || new Date().toISOString().split('T')[0];
        const fundShortName = fund.amc || fund.fullName || fund.name || 'Mutual Fund';

        if (actionMode === 'Withdrawal') {
          const withdrawAmount = (trans.amount && trans.amount > 0)
            ? trans.amount
            : (parseFloat(transactionFormData.actionAmount) || 0);
          onNavigateToModule('income-expense', {
            date: transactionFormData.actionDate || txDate,
            amount: withdrawAmount,
            type: 'Transfer',
            targetModule: 'mutual-funds-withdrawal',
            description: `Withdraw: ${fundShortName}`,
            linkedTxId: linkGroupId
          });
        } else if (sourceMode === 'Dividend') {
          const dividendAmount = (trans.sipAmount && trans.sipAmount > 0)
            ? trans.sipAmount
            : (trans.amount && trans.amount > 0)
              ? trans.amount
              : (parseFloat(transactionFormData.sourceAmount) || 0);
          onNavigateToModule('income-expense', {
            date: transactionFormData.sourceDate || txDate,
            amount: dividendAmount,
            type: 'Income',
            category: 'Finance Income',
            subCategory: 'Dividend',
            targetModule: 'mutual-funds-dividend',
            description: `Dividend: ${fundShortName}`,
            linkedTxId: linkGroupId
          });
        } else {
          const depositAmount = (trans.sipAmount && trans.sipAmount > 0)
            ? trans.sipAmount
            : (trans.amount && trans.amount > 0)
              ? trans.amount
              : (parseFloat(transactionFormData.sourceAmount) || 0);
          onNavigateToModule('income-expense', {
            date: transactionFormData.sourceDate || txDate,
            amount: depositAmount,
            type: 'Transfer',
            targetModule: 'mutual-funds',
            description: `${fundShortName}`,
            linkedTxId: linkGroupId
          });
        }
      }
    }

    setIsModalOpen(false);
    showNotification('success', editingTransaction ? 'Investment updated successfully!' : 'Investment(s) recorded successfully!');
    checkAndReturnToModule();
  };

  const handleDeleteTransaction = (fundId: string, transactionId: string) => {
    const fund = investments.find(f => f.id === fundId);
    if (!fund) return;
    const trans = fund.transactions.find(t => t.id === transactionId);
    
    setConfirmState({
      isOpen: true,
      title: 'Confirm Delete',
      message: `Delete ${trans?.type || 'transaction'} for ${fund.amc} on ${trans?.date || 'N/A'}?`,
      onConfirm: () => {
        closeConfirm();
        const updatedTransactions = fund.transactions.filter(t => t.id !== transactionId);
        onUpdate(fundId, { transactions: updatedTransactions });
        if (trans?.linkedIeGroupId && onDeleteLinkedTransfer) {
          onDeleteLinkedTransfer(trans.linkedIeGroupId);
        }
        showNotification('success', 'Transaction deleted successfully!');
      },
      onCancel: closeConfirm
    });
  };

  const handleBatchDelete = () => {
    if (selectedIds.length === 0) return;
    
    setConfirmState({
      isOpen: true,
      title: 'Confirm Batch Delete',
      message: `Are you sure you want to delete ${selectedIds.length} selected transactions?`,
      onConfirm: () => {
        closeConfirm();
        const deletions: { [key: string]: string[] } = {};
        selectedIds.forEach(uid => {
          const [fundId, transactionId] = uid.split('-');
          if (!deletions[fundId]) deletions[fundId] = [];
          deletions[fundId].push(transactionId);
        });

        const linkedGroupIdsToDelete: string[] = [];
        Object.entries(deletions).forEach(([fundId, transactionIds]) => {
          const fund = investments.find(f => f.id === fundId);
          if (fund) {
            fund.transactions.forEach(t => {
              if (transactionIds.includes(t.id) && t.linkedIeGroupId) {
                linkedGroupIdsToDelete.push(t.linkedIeGroupId);
              }
            });
            const updatedTransactions = fund.transactions.filter(t => !transactionIds.includes(t.id));
            onUpdate(fundId, { transactions: updatedTransactions });
          }
        });

        if (onDeleteLinkedTransfer) {
          linkedGroupIdsToDelete.forEach(gid => onDeleteLinkedTransfer(gid));
        }

        setSelectedIds([]);
        showNotification('success', `Successfully deleted ${selectedIds.length} transactions.`);
      },
      onCancel: closeConfirm
    });
  };

  const [lastSelectedId, setLastSelectedId] = useState<string | null>(null);

  const toggleSelectAll = () => {
    if (selectedIds.length === filtered.length && filtered.length > 0) {
      setSelectedIds([]);
      setLastSelectedId(null);
    } else {
      setSelectedIds(filtered.map(t => t.uniqueId));
    }
  };

  const toggleSelect = (id: string, e?: React.MouseEvent) => {
    if (e?.shiftKey && lastSelectedId && lastSelectedId !== id) {
      const lastIdx = filtered.findIndex(t => t.uniqueId === lastSelectedId);
      const currIdx = filtered.findIndex(t => t.uniqueId === id);
      if (lastIdx !== -1 && currIdx !== -1) {
        const start = Math.min(lastIdx, currIdx);
        const end = Math.max(lastIdx, currIdx);
        const rangeIds = filtered.slice(start, end + 1).map(t => t.uniqueId);
        setSelectedIds(prev => Array.from(new Set([...prev, ...rangeIds])));
        setLastSelectedId(id);
        return;
      }
    }

    setLastSelectedId(id);
    if (selectedIds.includes(id)) {
      setSelectedIds(selectedIds.filter(i => i !== id));
    } else {
      setSelectedIds([...selectedIds, id]);
    }
  };

  const handleEditTransaction = (t: any) => {
    setEditingTransaction({ fundId: t.fundId, transaction: t });
    
    let sMode: 'Deposit' | 'Dividend' | 'Surrender' = 'Deposit';
    let aMode: 'Buy' | 'Withdrawal' = 'Buy';

    if (t.isWithdrawal) aMode = 'Withdrawal';
    else if (t.type === 'Sell') sMode = 'Surrender';
    else if (t.isDividend) sMode = 'Dividend';
    else if (t.sipAmount > 0) sMode = 'Deposit';
    
    setSourceMode(sMode);
    setActionMode(aMode);

    setTransactionFormData({
      fundId: t.fundId,
      sourceDate: t.pullingDate || t.date,
      sourceAmount: ((sMode === 'Dividend' ? (t.sipAmount || t.amount) : t.sipAmount) || '').toString(),
      sourceUnits: t.type === 'Sell' ? t.units.toString() : '',
      sourceNav: t.type === 'Sell' ? t.nav.toString() : '',
      actionDate: t.date,
      actionAmount: t.isWithdrawal ? (t.amount || t.sipAmount || '').toString() : '',
      actionUnits: t.type === 'Buy' || t.type === 'Dividend' ? t.units.toString() : '',
      actionNav: t.type === 'Buy' || t.type === 'Dividend' ? t.nav.toString() : '',
      date: t.date,
      type: t.type,
      units: t.units.toString(),
      nav: t.nav.toString(),
      amount: t.amount.toString(),
      sipAmount: (t.sipAmount || '').toString(),
      pullingDate: (t.pullingDate || '').toString()
    });
    setIsModalOpen(true);
  };

  return (
    <div className="w-full">
      <div className="space-y-8">
      {/* Date Range Selector — same UI/behavior as Income & Expense */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 relative">
        <div className="flex flex-col sm:flex-row sm:items-center gap-3 w-full sm:w-auto">
          <div className="flex items-center justify-between w-full sm:w-auto gap-2">
            <div className="relative flex-1 sm:flex-none">
              {/* Mobile View: Dropdown */}
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
                    <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
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

              {/* Desktop View: Tabs */}
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

            {/* Settings Button - Mobile Only */}
            <div className="block sm:hidden">
              <SettingsButton
                isOpen={isSettingsMenuOpen}
                setIsOpen={setIsSettingsMenuOpen}
                onExport={exportData}
                onImport={importData}
                onTemplate={() => {
                  const templateData = [
                    ['AMC', 'Full Fund Name', 'Starting Month', 'Preferred Pulling Day', 'Base SIP Amount', 'Actual Pulling Date', 'SIP/Dividend Amount', 'Addition Type', 'Transaction Date', 'Transaction Type', 'Units', 'NAV', 'Amount'],
                    ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '2025-06-15', 25000, 'Deposit', '2025-06-17', 'Buy', 2434, 10.27, ''],
                    ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '2025-07-21', 25000, 'Deposit', '2025-07-23', 'Buy', 2356, 10.61, ''],
                    ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '2025-07-25', 5000, 'Dividend', '2025-07-26', 'Buy', 446, 10.7, ''],
                    ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '', '', '', '2025-07-26', 'Withdrawal', '', '', 1000]
                  ];
                  const ws = XLSX.utils.aoa_to_sheet(templateData);
                  const wb = XLSX.utils.book_new();
                  XLSX.utils.book_append_sheet(wb, ws, "Template");
                  XLSX.writeFile(wb, "Mutual_Fund_Import_Template.xlsx");
                }}
                onAdd={() => {
                  setEditingFund(null);
                  setIsFundModalOpen(true);
                }}
              />
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

        {/* Settings Button - Desktop Only */}
        <div className="hidden sm:block">
          <SettingsButton
            isOpen={isSettingsMenuOpen}
            setIsOpen={setIsSettingsMenuOpen}
            onExport={exportData}
            onImport={importData}
            onTemplate={() => {
              const templateData = [
                ['AMC', 'Full Fund Name', 'Starting Month', 'Preferred Pulling Day', 'Base SIP Amount', 'Actual Pulling Date', 'SIP/Dividend Amount', 'Addition Type', 'Transaction Date', 'Transaction Type', 'Units', 'NAV', 'Amount'],
                ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '2025-06-15', 25000, 'Deposit', '2025-06-17', 'Buy', 2434, 10.27, ''],
                ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '2025-07-21', 25000, 'Deposit', '2025-07-23', 'Buy', 2356, 10.61, ''],
                ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '2025-07-25', 5000, 'Dividend', '2025-07-26', 'Buy', 446, 10.7, ''],
                ['IDLC', 'IDLC Balanced Fund', '2025-06', 5, 25000, '', '', '', '2025-07-26', 'Withdrawal', '', '', 1000]
              ];
              const ws = XLSX.utils.aoa_to_sheet(templateData);
              const wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, ws, "Template");
              XLSX.writeFile(wb, "Mutual_Fund_Import_Template.xlsx");
            }}
            onAdd={() => {
              setEditingFund(null);
              setIsFundModalOpen(true);
            }}
          />
        </div>
      </div>

      {confirmState && (
        <ConfirmDialog 
          {...confirmState}
          onCancel={confirmState.onCancel || closeConfirm}
        />
      )}

      {notification && (
        <div className={cn(
          "flex items-center gap-3 p-4 rounded-xl border animate-in fade-in slide-in-from-top-4 duration-300",
          notification.type === 'success' ? "bg-teal-400/10 border-teal-400/20 text-teal-400" : "bg-rose-500/10 border-rose-500/20 text-rose-500"
        )}>
          {notification.type === 'success' ? <CheckCircle2 size={20} /> : <XCircle size={20} />}
          <p className="text-body font-bold uppercase">{notification.message}</p>
        </div>
      )}

      {/* ── 4 Main Summary Cards ── */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 items-stretch">

        {/* Card 1: Current Holding */}
        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-teal-400/10 flex items-center justify-center mb-4 text-teal-400 group-hover:scale-110 transition-transform">
              <DollarSign size={20} />
            </div>
            <div className="mb-2">
              <div className="flex items-center gap-1">
                <p className="text-body-sm font-bold text-white uppercase tracking-wider">Current Holding</p>
                <InfoTooltip text="Buying Cost for Units + Cash Balance" />
              </div>
              <p className="text-label font-bold text-slate-500 uppercase">Cost + Cash</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.currentHolding)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Active Funds</p>
            <p className="text-body font-bold text-teal-400 tabular-nums">
              {summaryStats.activeCount}
            </p>
          </div>
        </Card>

        {/* Card 2: Active Investment */}
        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-blue-400/50 hover:shadow-[0_0_20px_rgba(96,165,250,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-blue-400/10 flex items-center justify-center mb-4 text-blue-400 group-hover:scale-110 transition-transform">
              <Wallet size={20} />
            </div>
            <div className="mb-2">
              <div className="flex items-center gap-1">
                <p className="text-body-sm font-bold text-white uppercase tracking-wider">Active Investment</p>
                <InfoTooltip text="Total Deposit - Total Withdrawal" />
              </div>
              <p className="text-label font-bold text-slate-500 uppercase">Deposit - Withdrawal</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.activeInvestment)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Deposit</p>
            <p className="text-body font-bold text-blue-400 tabular-nums">
              {formatBDT(summaryStats.totalDeposit)}
            </p>
          </div>
        </Card>

        {/* Card 3: Total Invested */}
        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-purple-400/50 hover:shadow-[0_0_20px_rgba(167,139,250,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-purple-400/10 flex items-center justify-center mb-4 text-purple-400 group-hover:scale-110 transition-transform">
              <Briefcase size={20} />
            </div>
            <div className="mb-2">
              <div className="flex items-center gap-1">
                <p className="text-body-sm font-bold text-white uppercase tracking-wider">Total Invested</p>
                <InfoTooltip text="Total Deposit During the Selected Period" />
              </div>
              <p className="text-label font-bold text-slate-500 uppercase">Total Deposit</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.totalDeposit)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">This Month</p>
            <p className="text-body font-bold text-purple-400 tabular-nums">
              {formatBDT(summaryStats.thisMonthDeposit)}
            </p>
          </div>
        </Card>

        {/* Card 4: Total Profit */}
        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-emerald-400/50 hover:shadow-[0_0_20px_rgba(52,211,153,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-emerald-400/10 flex items-center justify-center mb-4 text-emerald-400 group-hover:scale-110 transition-transform">
              <TrendingUp size={20} />
            </div>
            <div className="mb-2">
              <div className="flex items-center gap-1">
                <p className="text-body-sm font-bold text-white uppercase tracking-wider">Total Profit</p>
                <InfoTooltip text="Total Dividend from All Funds" />
              </div>
              <p className="text-label font-bold text-slate-500 uppercase">Total Dividend</p>
            </div>
            <h3 className="text-heading font-bold text-emerald-400 mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.totalDividends)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Portfolio Yield</p>
            <p className={cn("text-body font-bold tabular-nums", summaryStats.dividendReturnPercentage >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {summaryStats.dividendReturnPercentage.toFixed(2)}%
            </p>
          </div>
        </Card>

      </div>

      {/* ── 4 Secondary Cards ── matching DSE Tracker style */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">

        {/* Units Cost */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Units Cost</p>
            <InfoTooltip text="Buying cost of All Units" />
          </div>
          <p className="text-body font-bold text-white tabular-nums">{formatBDT(summaryStats.totalCumulativeCostBasis)}</p>
        </Card>

        {/* Cash Balance */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Cash Balance</p>
            <InfoTooltip text="Total Cash in All Funds" />
          </div>
          <p className="text-body font-bold text-white tabular-nums">{formatBDT(summaryStats.cashBalance)}</p>
        </Card>

        {/* Surrender */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Surrender</p>
            <InfoTooltip text="Selling Cost of All Units" />
          </div>
          <p className="text-body font-bold text-rose-400 tabular-nums">{formatBDT(summaryStats.totalCumulativeSellAmount)}</p>
        </Card>

        {/* Withdrawal */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Withdrawal</p>
            <InfoTooltip text="Total Money withdrawn from All Funds" />
          </div>
          <p className="text-body font-bold text-rose-400 tabular-nums">{formatBDT(summaryStats.totalWithdrawn)}</p>
        </Card>

      </div>

      {/* ── Per-Fund Cards ── */}
      {fundStats.map(fund => {
        const fundData = uniqueInvestments.find(f => f.id === fund.id);
        return (
          <Card key={fund.id} className="bg-slate-900 border-slate-800">
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-2">
                <Briefcase className="text-teal-400 w-5 h-5" />
                <h3 className="text-subheading font-bold text-white uppercase">{fund.fullName}</h3>
              </div>
              <div className="flex items-center gap-1">
                <button 
                  onClick={() => fundData && handleEditFund(fundData)}
                  className="p-1.5 text-slate-500 hover:text-teal-400 hover:bg-teal-400/10 rounded-lg transition-all"
                  title="Edit Fund"
                >
                  <Edit2 size={14} />
                </button>
                <button 
                  onClick={() => {
                    const txCount = fund.transactions?.length || 0;
                    setConfirmState({
                      isOpen: true,
                      title: 'Severe Warning: Deleting Main Mutual Fund & All Transactions',
                      message: `You are about to delete the main Mutual Fund "${fund.fullName || fund.amc}". This will permanently delete the fund and ALL ${txCount} dependent buy, sell, SIP, and dividend transactions as well as historical performance records.`,
                      variant: 'severe',
                      confirmLabel: 'Delete Fund & All Transactions',
                      details: [
                        `Main Mutual Fund: "${fund.fullName || fund.amc}" (${fund.amc})`,
                        `${txCount} Dependent Buy / Sell / SIP Transactions`,
                        'All NAV, units, holding valuation, and return records for this fund'
                      ],
                      onConfirm: () => {
                        closeConfirm();
                        onDelete(fund.id);
                        showNotification('success', 'Mutual Fund and all transactions deleted successfully!');
                      },
                      onCancel: closeConfirm
                    });
                  }}
                  className="p-1.5 text-slate-500 hover:text-rose-500 hover:bg-rose-500/10 rounded-lg transition-all"
                  title="Delete Fund"
                >
                  <Trash2 size={14} />
                </button>
              </div>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {/* Card 1: Deposit & Holding */}
              <div className="bg-slate-950/50 border border-slate-800/50 rounded-xl p-4 flex flex-col justify-between transition-all hover:border-teal-400/30">
                <div>
                  <p className="text-label font-bold text-slate-300 uppercase mb-1">Total Deposit</p>
                  <p className="text-body sm:text-subheading font-bold text-white tabular-nums">{formatBDT(fund.totalDeposit)}</p>
                </div>
                <div className="mt-4 pt-4 border-t border-slate-800/50">
                  <p className="text-[10px] font-bold text-slate-300 uppercase mb-0.5">Cost of Units</p>
                  <p className="text-label font-bold text-slate-300 tabular-nums">{formatBDT(fund.rangeCostBasis)}</p>
                </div>
              </div>

              {/* Card 2: Dividend & Return */}
              <div className="bg-slate-950/50 border border-slate-800/50 rounded-xl p-4 flex flex-col justify-between transition-all hover:border-emerald-400/30">
                <div>
                  <p className="text-label font-bold text-slate-300 uppercase mb-1">Total Dividend</p>
                  <p className={cn("text-body sm:text-subheading font-bold tabular-nums", fund.totalDividends >= 0 ? "text-emerald-400" : "text-rose-500")}>
                    {formatBDT(fund.totalDividends)}
                  </p>
                </div>
                <div className="mt-4 pt-4 border-t border-slate-800/50">
                  <p className="text-[10px] font-bold text-slate-300 uppercase mb-0.5">Dividend Return (%)</p>
                  <p className={cn(
                    "text-label font-bold tabular-nums",
                    fund.dividendReturnPercentage >= 0 ? "text-emerald-400" : "text-rose-500"
                  )}>
                    {fund.dividendReturnPercentage.toFixed(2)}%
                  </p>
                </div>
              </div>

              {/* Card 3: Units Purchased & NAV */}
              <div className="bg-slate-950/50 border border-slate-800/50 rounded-xl p-4 flex flex-col justify-between transition-all hover:border-purple-400/30">
                <div>
                  <p className="text-label font-bold text-slate-300 uppercase mb-1">Units Purchased</p>
                  <p className="text-body sm:text-subheading font-bold text-white tabular-nums">
                    {formatNumber(fund.rangeUnitsPurchased)}
                  </p>
                </div>
                <div className="mt-4 pt-4 border-t border-slate-800/50">
                  <p className="text-[10px] font-bold text-slate-300 uppercase mb-0.5">Average NAV</p>
                  <p className="text-label font-bold text-slate-300 tabular-nums">{formatNumber(fund.averageNav)}</p>
                </div>
              </div>

              {/* Card 4: Cash Balance */}
              <div className="bg-slate-950/50 border border-slate-800/50 rounded-xl p-4 flex flex-col justify-between transition-all hover:border-blue-400/30">
                <div>
                  <p className="text-label font-bold text-slate-300 uppercase mb-1">Cash Balance</p>
                  <p className="text-body sm:text-subheading font-bold text-white tabular-nums">{formatBDT(fund.cashBalance)}</p>
                </div>
                <div className="mt-4 pt-4 border-t border-slate-800/50">
                  <p className="text-[10px] font-bold text-slate-300 uppercase mb-0.5">Total Withdrawn</p>
                  <p className="text-label font-bold text-slate-300 tabular-nums">{formatBDT(fund.totalWithdrawn)}</p>
                </div>
              </div>
            </div>
          </Card>
        );
      })}
      
      {/* Portfolio Header & Filters */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-6">
        <div className="flex items-center gap-2 sm:gap-3">
          <Briefcase className="text-teal-400 w-5 h-5 lg:w-6 lg:h-6" />
          <h2 className="text-body-sm sm:text-subheading lg:text-heading font-bold text-white font-display uppercase whitespace-nowrap">Mutual Funds Portfolio</h2>
        </div>
        <div className="flex items-center gap-2 flex-1 sm:flex-none justify-end">
          <div className="relative flex-1 sm:flex-none">
            <div 
              className="flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-lg p-1 px-2 h-9 hover:border-slate-700 focus-within:border-teal-400/50 transition-colors group cursor-text"
              onClick={() => {
                const input = document.getElementById('mf-search-input');
                if (input) input.focus();
              }}
            >
              <Search size={14} className="text-slate-500 group-focus-within:text-teal-400 transition-colors shrink-0" />
              <input 
                id="mf-search-input"
                type="text"
                placeholder="Search"
                className="flex-1 bg-transparent border-none focus:ring-0 text-[10px] font-bold text-white placeholder:text-slate-600 uppercase outline-none min-w-0"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
            </div>
          </div>

          <div className="relative shrink-0">
            <button 
              onClick={() => setIsFilterMenuOpen(!isFilterMenuOpen)}
              className={cn(
                "flex items-center gap-2 px-3 h-9 rounded-lg transition-all text-[10px] font-bold uppercase shadow-lg",
                selectedTypes.length > 0 
                  ? "bg-teal-400 text-slate-950 shadow-teal-400/20" 
                  : "bg-slate-900 border border-slate-800 text-slate-400 hover:border-slate-700"
              )}
            >
              <Filter size={14} />
              <span className="hidden sm:inline">Type</span>
            </button>

            {isFilterMenuOpen && (
              <>
                <div className="fixed inset-0 z-40" onClick={() => setIsFilterMenuOpen(false)} />
                <div className="absolute right-0 mt-2 w-48 bg-slate-900 border border-slate-800 rounded-lg shadow-2xl z-50 p-2 animate-in fade-in zoom-in-95">
                  {['Deposit', 'Dividend', 'Buy', 'Sell', 'Withdrawal'].map((type) => (
                    <div
                      key={type}
                      className="flex items-center gap-2 px-3 py-1 rounded hover:bg-slate-800 transition-colors cursor-pointer"
                      onClick={() => {
                        setSelectedTypes(prev => 
                          prev.includes(type) ? prev.filter(t => t !== type) : [...prev, type]
                        );
                      }}
                    >
                      <Checkbox 
                        label={type}
                        checked={selectedTypes.includes(type)} 
                        onChange={() => {}}
                      />
                    </div>
                  ))}
                  {selectedTypes.length > 0 && (
                    <button 
                      onClick={() => setSelectedTypes([])}
                      className="w-full mt-2 pt-2 border-t border-slate-800 text-center text-[10px] font-bold text-rose-400 hover:text-rose-300 uppercase transition-colors"
                    >
                      Clear All
                    </button>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Time Filtering Section */}
      <div className="flex items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 mb-3">
        {/* Mobile View: Dropdown */}
        <div className="relative flex-1 sm:hidden">
          <button 
            onClick={() => setIsTimeMenuOpen(!isTimeMenuOpen)}
            className="flex items-center justify-between gap-2 bg-slate-950 border border-slate-800 rounded-lg px-3 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase w-full"
          >
            <div className="flex items-center gap-1.5 min-w-0">
              <Clock size={14} className="text-teal-400 shrink-0" />
              <span className="truncate">
                {timeFilterMode === '6months' ? '6 Months' : 
                 timeFilterMode === '1year' ? '1 Year' : 'Custom'}
              </span>
            </div>
            <ChevronDown size={14} className={cn("text-slate-500 transition-transform duration-200 shrink-0", isTimeMenuOpen ? "rotate-180 text-teal-400" : "")} />
          </button>

          {isTimeMenuOpen && (
            <>
              <div className="fixed inset-0 z-40" onClick={() => setIsTimeMenuOpen(false)} />
              <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95 backdrop-blur-md divide-y divide-slate-800/30">
                {[
                  { id: '6months', label: '6 Months' },
                  { id: '1year', label: '1 Year' },
                  { id: 'custom', label: 'Custom' }
                ].map((opt) => (
                  <div key={opt.id} className="py-0.5 first:pt-0 last:pb-0">
                    <button 
                      onClick={() => { setTimeFilterMode(opt.id as any); setTimeOffset(0); setIsTimeMenuOpen(false); }}
                      className={cn(
                        "w-full text-left px-3 py-2.5 text-[10px] font-bold rounded-lg transition-colors uppercase",
                        timeFilterMode === opt.id ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                      )}
                    >
                      {opt.label}
                    </button>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>

        {/* Desktop View: Buttons */}
        <div className="hidden sm:flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1">
          <button
            onClick={() => { setTimeFilterMode('6months'); setTimeOffset(0); }}
            className={cn(
              "px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap border",
              timeFilterMode === '6months' 
                ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20 border-teal-400/30" 
                : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white hover:bg-slate-800/60 hover:border-slate-700/50"
            )}
          >
            6 Months
          </button>
          <button
            onClick={() => { setTimeFilterMode('1year'); setTimeOffset(0); }}
            className={cn(
              "px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap border",
              timeFilterMode === '1year' 
                ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20 border-teal-400/30" 
                : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white hover:bg-slate-800/60 hover:border-slate-700/50"
            )}
          >
            1 Year
          </button>
          <button
            onClick={() => setTimeFilterMode('custom')}
            className={cn(
              "px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap border",
              timeFilterMode === 'custom' 
                ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20 border-teal-400/30" 
                : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white hover:bg-slate-800/60 hover:border-slate-700/50"
            )}
          >
            Custom
          </button>
        </div>

        {(timeFilterMode === '6months' || timeFilterMode === '1year') && (
          <div className="flex items-center gap-1 bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 flex-1 sm:flex-none justify-between sm:justify-center max-w-[180px] sm:max-w-none">
            <button 
              onClick={() => setTimeOffset(prev => prev + 1)}
              className="p-1 px-1.5 hover:bg-slate-800 rounded-md text-slate-400 hover:text-teal-400 transition-all shrink-0"
              title={timeFilterMode === '6months' ? "Previous 6 Months" : "Previous Year"}
            >
              <ChevronLeft size={16} />
            </button>
            <div className="px-1 min-w-0 text-center overflow-hidden">
              <span className="text-[9px] sm:text-[10px] font-bold text-white uppercase tracking-wider whitespace-nowrap block truncate">
                {(() => {
                  const now = new Date();
                  const monthsCount = timeFilterMode === '6months' ? 6 : 12;
                  const currentOffsetMonths = timeOffset * monthsCount;
                  
                  const endMonth = now.getMonth() - currentOffsetMonths;
                  const end = new Date(now.getFullYear(), endMonth, 1);
                  
                  const startMonth = endMonth - (monthsCount - 1);
                  const start = new Date(now.getFullYear(), startMonth, 1);
                  
                  const startStr = start.toLocaleString('default', { month: 'short', year: 'numeric' });
                  const endStr = end.toLocaleString('default', { month: 'short', year: 'numeric' });
                  return `${startStr} - ${endStr}`;
                })()}
              </span>
            </div>
            <button 
              onClick={() => setTimeOffset(prev => Math.max(0, prev - 1))}
              disabled={timeOffset === 0}
              className={cn(
                "p-1.5 rounded-md transition-all",
                timeOffset === 0 
                  ? "text-slate-700 cursor-not-allowed" 
                  : "hover:bg-slate-800 text-slate-400 hover:text-teal-400"
              )}
              title={timeFilterMode === '6months' ? "Next 6 Months" : "Next Year"}
            >
              <ChevronRight size={16} />
            </button>
          </div>
        )}

        {timeFilterMode === 'custom' && (
          <div className="flex items-center gap-2 animate-in fade-in slide-in-from-right-2 duration-300">
            <div className="flex items-center gap-2">
              <input 
                type="date"
                className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1 text-[10px] font-bold text-white uppercase outline-none focus:border-teal-400/50"
                value={customRange.start}
                onChange={(e) => setCustomRange(prev => ({ ...prev, start: e.target.value }))}
              />
              <span className="text-slate-600 text-[10px] font-bold uppercase">To</span>
              <input 
                type="date"
                className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1 text-[10px] font-bold text-white uppercase outline-none focus:border-teal-400/50"
                value={customRange.end}
                onChange={(e) => setCustomRange(prev => ({ ...prev, end: e.target.value }))}
              />
            </div>
            {(customRange.start || customRange.end) && (
              <button 
                onClick={() => setCustomRange({ start: '', end: '' })}
                className="p-1.5 hover:bg-rose-500/10 rounded-md text-rose-500 transition-all"
                title="Clear Range"
              >
                <XCircle size={14} />
              </button>
            )}
          </div>
        )}
      </div>

      <div className="flex flex-col border border-slate-800 rounded-xl overflow-hidden bg-slate-900/50 shadow-2xl shadow-slate-950/50">
        <div className="overflow-x-auto custom-scrollbar">
          <div className="min-w-[900px] grid grid-cols-[48px_minmax(0,1.5fr)_minmax(0,1fr)_minmax(0,1.2fr)_minmax(0,1.5fr)_minmax(0,1.2fr)_96px]">
            {/* Header */}
            <div className="flex items-center justify-center border-r border-b border-slate-800 bg-slate-900/80 py-3 sticky top-0 z-10 backdrop-blur-sm">
              <Checkbox 
                checked={selectedIds.length === filtered.length && filtered.length > 0}
                onChange={toggleSelectAll}
              />
            </div>
            <div className="flex items-center gap-2 cursor-pointer group px-4 py-3 border-r border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm" onClick={() => handleSort('fundName')}>
              {sortBy === 'fundName' ? (sortOrder === 'asc' ? <ArrowUp size={12} className="text-teal-400" /> : <ArrowDown size={12} className="text-teal-400" />) : <ArrowUpDown size={12} className="text-slate-600 group-hover:text-slate-400" />}
              <span className={cn("text-label font-bold uppercase tracking-wider transition-colors", sortBy === 'fundName' ? "text-teal-400" : "text-slate-500 group-hover:text-slate-300")}>Mutual Fund</span>
            </div>
            <div className="flex items-center gap-2 cursor-pointer group px-4 py-3 border-r border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm" onClick={() => handleSort('date')}>
              {sortBy === 'date' ? (sortOrder === 'asc' ? <ArrowUp size={12} className="text-teal-400" /> : <ArrowDown size={12} className="text-teal-400" />) : <ArrowUpDown size={12} className="text-slate-600 group-hover:text-slate-400" />}
              <span className={cn("text-label font-bold uppercase tracking-wider transition-colors", sortBy === 'date' ? "text-teal-400" : "text-slate-500 group-hover:text-slate-300")}>Date</span>
            </div>
            <div className="flex items-center gap-2 cursor-pointer group px-4 py-3 border-r border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm" onClick={() => handleSort('sipAmount')}>
              {sortBy === 'sipAmount' ? (sortOrder === 'asc' ? <ArrowUp size={12} className="text-teal-400" /> : <ArrowDown size={12} className="text-teal-400" />) : <ArrowUpDown size={12} className="text-slate-600 group-hover:text-slate-400" />}
              <span className={cn("text-label font-bold uppercase tracking-wider transition-colors", sortBy === 'sipAmount' ? "text-teal-400" : "text-slate-500 group-hover:text-slate-300")}>Addition</span>
            </div>
            <div className="flex items-center gap-2 cursor-pointer group px-4 py-3 border-r border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm" onClick={() => handleSort('units')}>
              {sortBy === 'units' ? (sortOrder === 'asc' ? <ArrowUp size={12} className="text-teal-400" /> : <ArrowDown size={12} className="text-teal-400" />) : <ArrowUpDown size={12} className="text-slate-600 group-hover:text-slate-400" />}
              <span className={cn("text-label font-bold uppercase tracking-wider transition-colors", sortBy === 'units' ? "text-teal-400" : "text-slate-500 group-hover:text-slate-300")}>Transaction</span>
            </div>
            <div className="flex items-center gap-2 cursor-pointer group px-4 py-3 border-r border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm" onClick={() => handleSort('amount')}>
              {sortBy === 'amount' ? (sortOrder === 'asc' ? <ArrowUp size={12} className="text-teal-400" /> : <ArrowDown size={12} className="text-teal-400" />) : <ArrowUpDown size={12} className="text-slate-600 group-hover:text-slate-400" />}
              <span className={cn("text-label font-bold uppercase tracking-wider transition-colors", sortBy === 'amount' ? "text-teal-400" : "text-slate-500 group-hover:text-slate-300")}>Total Cost</span>
            </div>
            <div className="border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm" />

            {selectedIds.length > 0 && (
              <div className="col-span-full flex items-center justify-between bg-rose-500/10 border-b border-rose-500/20 px-4 py-2 animate-in fade-in slide-in-from-top-2">
                <span className="text-label font-bold text-rose-500 uppercase">
                  {selectedIds.length} Transactions Selected
                </span>
                <Button 
                  variant="danger" 
                  size="sm" 
                  onClick={handleBatchDelete}
                  className="h-8 px-4 text-label font-bold uppercase"
                >
                  Delete Selected
                </Button>
              </div>
            )}

            {filtered.map((t) => {
              const isSelected = selectedIds.includes(t.uniqueId);
              const cellBaseClass = cn(
                "px-4 py-3 border-r border-b border-slate-800/50 transition-colors group-hover:bg-slate-800/40 flex flex-col justify-center",
                isSelected ? "bg-teal-400/10" : "bg-slate-900/40"
              );

              return (
                <React.Fragment key={t.uniqueId}>
                  <div className={cn("flex items-center justify-center border-r border-b border-slate-800/50 transition-colors group-hover:bg-slate-800/40", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <Checkbox 
                      checked={isSelected}
                      onChange={(_, e) => toggleSelect(t.uniqueId, e)}
                    />
                  </div>
                  
                  <div className={cellBaseClass}>
                    <span className="text-body-sm font-bold text-white uppercase truncate mb-1">{t.fundName}</span>
                  </div>

                  <div className={cellBaseClass}>
                    <span className="text-body-sm font-bold text-white uppercase tabular-nums mb-1">
                      {new Date(t.date).toLocaleString('default', { month: 'short', year: 'numeric' }).toUpperCase()}
                    </span>
                  </div>

                  <div className={cellBaseClass}>
                    <span className="text-body-sm font-bold text-white tabular-nums mb-1">{t.isWithdrawal ? '' : formatBDT(t.sipAmount)}</span>
                    <span className={cn(
                      "text-[7px] sm:text-label font-bold uppercase",
                      t.isDividend ? "text-blue-400" : (t.isWithdrawal ? "text-rose-400" : "text-teal-400")
                    )}>
                      {t.isDividend ? 'Dividend' : (t.isWithdrawal ? '' : 'Deposit')}
                    </span>
                  </div>

                  <div className={cellBaseClass}>
                    <span className="text-body-sm font-bold text-white tabular-nums mb-1">
                      {t.isWithdrawal ? 'WITHDRAWAL' : `${formatNumber(t.units)} x ${formatNumber(t.nav)}`}
                    </span>
                    <span className={cn(
                      "text-[7px] sm:text-label font-bold uppercase",
                      t.type === 'Buy' || t.type === 'Dividend' ? "text-emerald-400" : "text-rose-500"
                    )}>
                      {t.type === 'Dividend' ? 'Buy' : t.type}
                    </span>
                  </div>

                  <div className={cellBaseClass}>
                    <span className="text-body-sm font-bold text-white tabular-nums mb-1">{formatBDT(t.amount)}</span>
                  </div>

                  <div className={cn("flex items-center gap-1 px-4 py-3 border-b border-slate-800/50 transition-colors group-hover:bg-slate-800/40", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <button 
                      onClick={() => handleEditTransaction(t)}
                      className="p-2 text-slate-500 hover:text-teal-400 hover:bg-teal-400/10 rounded-lg transition-all"
                      title="Edit Transaction"
                    >
                      <Edit2 size={16} />
                    </button>
                    <button 
                      onClick={() => handleDeleteTransaction(t.fundId, t.id)}
                      className="p-2 text-slate-500 hover:text-rose-500 hover:bg-rose-500/10 rounded-lg transition-all"
                      title="Delete Transaction"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                </React.Fragment>
              );
            })}
          </div>
        </div>

        {filtered.length === 0 && (
          <div className="py-12 text-center bg-slate-900/30 border border-dashed border-slate-800 rounded-b-xl">
            <p className="text-body text-slate-500 italic">No transactions matching your range found</p>
          </div>
        )}
      </div>

      {/* New Mutual Fund Modal */}
      <Modal 
        isOpen={isFundModalOpen} 
        onClose={() => {
          setIsFundModalOpen(false);
          setEditingFund(null);
          checkAndReturnToModule();
        }} 
        title={editingFund ? "Edit Mutual Fund" : "Add Mutual Fund"}
      >
        <form onSubmit={handleFundSubmit} className="space-y-4">
          <div className="grid grid-cols-2 gap-3">
            <Input 
              label="Mutual Fund (AMC)" 
              placeholder="e.g. IDLC, EDGE"
              value={fundFormData.amc}
              onChange={(e) => setFundFormData({ ...fundFormData, amc: e.target.value })}
              required
            />
            <Input 
              label="Full Fund Name" 
              placeholder="e.g. IDLC Balanced Fund"
              value={fundFormData.fullName}
              onChange={(e) => setFundFormData({ ...fundFormData, fullName: e.target.value })}
              required
            />
          </div>
          <div className="grid grid-cols-3 gap-3">
            <Input 
              label="Starting Month" 
              type="month"
              value={fundFormData.startingMonth}
              onChange={(e) => setFundFormData({ ...fundFormData, startingMonth: e.target.value })}
              required
            />
            <Input 
              label="Pulling Date" 
              type="number"
              min="1"
              max="31"
              value={fundFormData.pullingDate}
              onChange={(e) => setFundFormData({ ...fundFormData, pullingDate: e.target.value })}
              required
            />
            <Input 
              label="SIP Amount" 
              type="number"
              placeholder="0.00"
              value={fundFormData.sipAmount}
              onChange={(e) => setFundFormData({ ...fundFormData, sipAmount: e.target.value })}
              required
            />
          </div>
          <div className="pt-2 flex gap-3">
            <Button type="button" variant="secondary" className="flex-1 py-2" onClick={() => {
              setIsFundModalOpen(false);
              setEditingFund(null);
              checkAndReturnToModule();
            }}>
              Cancel
            </Button>
            <Button type="submit" className="flex-1 py-2">
              {editingFund ? "Update Fund" : "Create Fund"}
            </Button>
          </div>
        </form>
      </Modal>

      {/* New / Edit Investment Modal */}
      <Modal 
        isOpen={isModalOpen} 
        onClose={() => { setIsModalOpen(false); setEditingTransaction(null); isFromExternalModuleRef.current = false; setInheritedIeGroupId(undefined); checkAndReturnToModule(); }} 
        title={editingTransaction ? "Edit Investment" : "Add Investment"}
      >
        <form onSubmit={handleTransactionSubmit} className="space-y-4">
          <Select 
            label="Mutual Fund" 
            value={transactionFormData.fundId}
            options={uniqueInvestments.map(f => ({ label: f.amc, value: f.id }))}
            placeholder="Select a fund"
            onChange={(val) => {
              const fundId = val;
              const fund = investments.find(f => f.id === fundId);
              if (fund) {
                const today = new Date();
                const year = today.getFullYear();
                const month = String(today.getMonth() + 1).padStart(2, '0');
                const day = String(fund.pullingDate).padStart(2, '0');
                const defaultDate = `${year}-${month}-${day}`;
                
                setTransactionFormData(prev => ({ 
                   ...prev, 
                   fundId, 
                   sourceAmount: (prev.sourceAmount && prev.sourceAmount !== '0') ? prev.sourceAmount : (sourceMode === 'Deposit' ? fund.sipAmount.toString() : '0'),
                   sourceDate: prev.sourceDate || defaultDate,
                   actionDate: prev.actionDate || defaultDate,
                   sipAmount: (prev.sipAmount && prev.sipAmount !== '0') ? prev.sipAmount : fund.sipAmount.toString(),
                   pullingDate: prev.pullingDate || defaultDate,
                   date: prev.date || defaultDate
                }));
              } else {
                setTransactionFormData(prev => ({ ...prev, fundId }));
              }
            }}
            required
          />

          <label className="text-label font-bold text-slate-300 uppercase">Source (Add to Balance)</label>
          <div className="flex gap-2 mb-4">
            {['Deposit', 'Dividend', 'Surrender'].map((mode) => (
              <Button 
                key={mode}
                type="button"
                variant={sourceMode === mode ? 'primary' : 'secondary'}
                className="flex-1 h-9 text-[10px] font-bold uppercase tracking-wider"
                onClick={() => {
                  setSourceMode(mode as any);
                  if (mode === 'Dividend' || mode === 'Surrender') {
                    setTransactionFormData(prev => ({ ...prev, sourceAmount: '0', sourceUnits: '0', sourceNav: '0' }));
                  } else if (mode === 'Deposit') {
                    const fund = investments.find(f => f.id === transactionFormData.fundId);
                    if (fund) {
                      setTransactionFormData(prev => ({ ...prev, sourceAmount: fund.sipAmount.toString() }));
                    }
                  }
                }}
              >
                {mode}
              </Button>
            ))}
          </div>

          <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
            {sourceMode === 'Surrender' ? (
              <>
                <Input 
                  label="Surrender Date" 
                  type="date"
                  value={transactionFormData.sourceDate}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, sourceDate: e.target.value })}
                  required
                />
                <Input 
                  label="Units" 
                  type="number"
                  step="0.01"
                  value={transactionFormData.sourceUnits}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, sourceUnits: e.target.value })}
                  required
                />
                <Input 
                  label="NAV" 
                  type="number"
                  step="0.0001"
                  value={transactionFormData.sourceNav}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, sourceNav: e.target.value })}
                  required
                />
              </>
            ) : (
              <>
                <Input 
                  label={sourceMode === 'Dividend' ? "Dividend Date" : "Pulling Date"} 
                  type="date"
                  value={transactionFormData.sourceDate}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, sourceDate: e.target.value })}
                  required
                />
                <Input 
                  label={sourceMode === 'Dividend' ? "Dividend Amount" : "SIP Amount"} 
                  type="number"
                  placeholder="0.00"
                  value={transactionFormData.sourceAmount}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, sourceAmount: e.target.value })}
                  required
                />
              </>
            )}
          </div>

          <div className="border-t border-slate-800/50 my-6" />

          <label className="text-label font-bold text-slate-300 uppercase">Action (Deduct from Balance)</label>
          <div className="flex gap-2 mb-4">
            {['Buy', 'Withdrawal'].map((mode) => (
              <Button 
                key={mode}
                type="button"
                variant={actionMode === mode ? 'primary' : 'secondary'}
                className="flex-1 h-9 text-[10px] font-bold uppercase tracking-wider"
                onClick={() => {
                  setActionMode(mode as any);
                  setTransactionFormData(prev => ({ 
                    ...prev, 
                    actionAmount: '0', 
                    actionUnits: '0', 
                    actionNav: '0' 
                  }));
                }}
              >
                {mode === 'Withdrawal' ? 'Withdraw' : mode}
              </Button>
            ))}
          </div>

          <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
            {actionMode === 'Buy' ? (
              <>
                <Input 
                  label="Buy Date" 
                  type="date"
                  value={transactionFormData.actionDate}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, actionDate: e.target.value })}
                  required
                />
                <Input 
                  label="Units" 
                  type="number"
                  step="0.01"
                  value={transactionFormData.actionUnits}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, actionUnits: e.target.value })}
                  required
                />
                <Input 
                  label="NAV" 
                  type="number"
                  step="0.0001"
                  value={transactionFormData.actionNav}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, actionNav: e.target.value })}
                  required
                />
              </>
            ) : (
              <>
                <Input 
                  label="Withdraw Date" 
                  type="date"
                  value={transactionFormData.actionDate}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, actionDate: e.target.value })}
                  required
                />
                <Input 
                  label="Withdraw Amount" 
                  type="number"
                  placeholder="0.00"
                  value={transactionFormData.actionAmount}
                  onChange={(e) => setTransactionFormData({ ...transactionFormData, actionAmount: e.target.value })}
                  required
                />
              </>
            )}
          </div>

          <div className="pt-2 flex gap-3">
            <Button type="button" variant="secondary" className="flex-1 py-2" onClick={() => { setIsModalOpen(false); setEditingTransaction(null); isFromExternalModuleRef.current = false; setInheritedIeGroupId(undefined); checkAndReturnToModule(); }}>
              Cancel
            </Button>
            <Button type="submit" className="flex-1 py-2">
              {editingTransaction ? "Update Investment" : "Record Investment"}
            </Button>
          </div>
        </form>
      </Modal>
      {confirmState && (
        <ConfirmDialog 
          {...confirmState}
          onCancel={confirmState.onCancel || closeConfirm}
        />
      )}
    </div>
  </div>
);
};

