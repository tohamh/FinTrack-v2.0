/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import { Card, Button, Modal, Input, Checkbox, Select } from '../ui/BaseComponents';
import { ConfirmDialog } from '../ui/ConfirmDialog';
import { FDR, FDRTransaction, Currency } from '../../types';
import { formatBDT, formatDate, cn, formatNumber, formatCurrency, toDateStr, getTodayStr, getFirstOfMonth, getLastOfMonth } from '../../utils/formatters';
import * as XLSX from 'xlsx';
import { Plus, MoreVertical, Calendar, TrendingUp, Search, CheckCircle2, Briefcase, Settings, Download, Upload, FileSpreadsheet, XCircle, DollarSign, Coins, Percent, ArrowRight, Edit2, Trash2, ArrowUpDown, ArrowUp, ArrowDown, Filter, Clock, ChevronLeft, ChevronRight, ChevronDown, Landmark, Wallet, Info } from 'lucide-react';

interface FixedDepositsModuleProps {
  investments: FDR[];
  onAdd: (investment: Omit<FDR, 'id'>) => void;
  onUpdate: (id: string, updates: Partial<FDR>) => void;
  onDelete: (id: string) => void;
  onBatchDelete: (ids: string[]) => void;
  onBatchAdd?: (investments: Omit<FDR, 'id'>[]) => void;
  onReplaceAll?: (investments: Omit<FDR, 'id'>[]) => void;
  triggerAdd?: boolean;
  setTriggerAdd?: (val: boolean) => void;
  onTitleChange?: (title: React.ReactNode) => void;
  inheritedData?: { date: string; amount: number; type?: 'Transfer' | 'Income' | 'Expense' | 'Loan'; targetModule?: string; description?: string } | null;
  onClearInheritedData?: () => void;
  onNavigateToModule?: (module: string, initialAddData?: { date: string; amount: number; type?: 'Transfer' | 'Income' | 'Expense' | 'Loan'; targetModule?: string; description?: string }) => void;
}

// Tooltip component matching MF module style
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
              Add Fixed Deposit
            </button>
          </div>
        </>
      )}
    </div>
  );
};

export const FixedDepositsModule: React.FC<FixedDepositsModuleProps> = ({ 
  investments, 
  onAdd, 
  onUpdate,
  onDelete,
  onBatchDelete,
  onBatchAdd,
  onReplaceAll,
  triggerAdd,
  setTriggerAdd,
  onTitleChange,
  inheritedData,
  onClearInheritedData,
  onNavigateToModule
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

  const [sortBy, setSortBy] = useState<'date' | 'amount' | 'bankName' | 'accountNo' | 'type' | 'balance'>('date');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [searchQuery, setSearchQuery] = useState('');
  const [editingFDR, setEditingFDR] = useState<FDR | null>(null);
  const [editingTransaction, setEditingTransaction] = useState<{ fdrId: string, transaction: any } | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isFDRModalOpen, setIsFDRModalOpen] = useState(false);
  const [activeMenuId, setActiveMenuId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isSettingsMenuOpen, setIsSettingsMenuOpen] = useState(false);
  const [selectedTypes, setSelectedTypes] = useState<string[]>([]);
  const [isFilterMenuOpen, setIsFilterMenuOpen] = useState(false);
  const [notification, setNotification] = useState<{ type: 'success' | 'error', message: string } | null>(null);
  const [confirmState, setConfirmState] = useState<{ 
    isOpen: boolean; 
    title: string; 
    message: string; 
    onConfirm: () => void; 
    variant?: 'danger' | 'warning' | 'info';
    confirmLabel?: string;
  } | null>(null);

  const closeConfirm = () => setConfirmState(prev => prev ? { ...prev, isOpen: false } : null);
  const [historyRange, setHistoryRange] = useState<'all' | 'last12m' | 'fiscal' | 'custom'>('all');
  const [historyCustomDates, setHistoryCustomDates] = useState(() => {
    const now = new Date();
    return {
      start: getFirstOfMonth(now),
      end: getTodayStr(),
    };
  });
  const [customNavMonth, setCustomNavMonth] = useState<{ year: number; month: number }>(() => {
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() };
  });

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
  };
  const [isRangeMenuOpen, setIsRangeMenuOpen] = useState(false);
  const [isTimeMenuOpen, setIsTimeMenuOpen] = useState(false);

  // Time-based filtering state for transaction table
  const [timeFilterMode, setTimeFilterMode] = useState<'6months' | '1year' | 'custom'>('6months');
  const [timeOffset, setTimeOffset] = useState(0);
  const [customRange, setCustomRange] = useState({ start: '', end: '' });

  // Derive active date range for dashboard stats
  const rangeDates = useMemo(() => {
    const now = new Date();
    if (historyRange === 'all') return { start: new Date(0), end: new Date(3000, 11, 31, 23, 59, 59, 999) };
    
    if (historyRange === 'last12m') {
      const start = new Date(now.getFullYear(), now.getMonth() - 11, 1);
      const end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
      return { start, end };
    }
    
    if (historyRange === 'fiscal') {
      const currentYear = now.getFullYear();
      const startYear = now.getMonth() >= 6 ? currentYear - 1 : currentYear - 2;
      return { 
        start: new Date(startYear, 6, 1), 
        end: new Date(startYear + 1, 5, 30, 23, 59, 59, 999) 
      };
    }
    
    const start = historyCustomDates.start ? new Date(historyCustomDates.start) : new Date(0);
    const end = historyCustomDates.end ? new Date(historyCustomDates.end) : new Date();
    end.setHours(23, 59, 59, 999);
    return { start, end };
  }, [historyRange, historyCustomDates]);

  useEffect(() => {
    if (onTitleChange) {
      if (historyRange === 'all') {
        onTitleChange('Fixed Deposits');
      } else {
        const { start, end } = rangeDates;
        const startStr = start.toLocaleString('default', { month: 'short', year: 'numeric' });
        const endStr = end.toLocaleString('default', { month: 'short', year: 'numeric' });
        onTitleChange(
          <span className="flex items-center gap-2">
            FIXED DEPOSITS <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">/ {startStr} - {endStr}</span>
          </span>
        );
      }
    }
  }, [rangeDates, onTitleChange, historyRange]);

  useEffect(() => {
    if (triggerAdd) {
      setEditingTransaction(null);
      setEditingFDR(null);
      if (inheritedData) {
        const inheritedDate = inheritedData.date || new Date().toISOString().split('T')[0];
        const inheritedAmt = (inheritedData.amount !== undefined && inheritedData.amount !== null && inheritedData.amount !== 0) ? String(inheritedData.amount) : '';
        const inheritedDesc = inheritedData.description || '';

        // Match existing FDR by name, bankName, or accountNo
        const matchedFDR = uniqueInvestments.find(f => 
          (f.bankName && inheritedDesc.toLowerCase().includes(f.bankName.toLowerCase())) ||
          (f.accountNo && inheritedDesc.toLowerCase().includes(f.accountNo.toLowerCase())) ||
          (f.name && inheritedDesc.toLowerCase().includes(f.name.toLowerCase()))
        );

        if (inheritedData.type === 'Income') {
          setTransactionFormData({
            fdrId: matchedFDR?.id || (uniqueInvestments[0]?.id || ''),
            date: inheritedDate,
            type: 'Profit',
            amount: inheritedAmt,
            handling: matchedFDR ? matchedFDR.interestHandling : 'Withdrawn'
          });
          setIsModalOpen(true);
        } else {
          // Parse bank name / account number from description if possible
          // e.g. "Bank Asia (12345678)" -> bankName: "Bank Asia", accountNo: "12345678"
          let parsedBank = '';
          let parsedAcc = '';
          const match = inheritedDesc.match(/^(.*?)(?:\s*\((.*?)\)|\s*-\s*(.*?))?$/);
          if (match) {
            parsedBank = (match[1] || '').trim();
            parsedAcc = (match[2] || match[3] || '').trim();
          }

          setFdrFormData({
            bankName: parsedBank || matchedFDR?.bankName || '',
            accountNo: parsedAcc || matchedFDR?.accountNo || '',
            investmentDate: inheritedDate,
            principal: inheritedAmt,
            currency: matchedFDR?.currency || 'BDT',
            interestFrequency: matchedFDR?.interestFrequency || 'Monthly',
            interestHandling: matchedFDR?.interestHandling || 'Withdrawn',
            status: 'Active',
            exchangeRate: '',
            closingDate: '',
            withdrawBalance: '',
          });
          setIsFDRModalOpen(true);
        }
        onClearInheritedData?.();
      } else {
        setIsModalOpen(true);
      }
      setTriggerAdd?.(false);
    }
  }, [triggerAdd, setTriggerAdd, inheritedData, onClearInheritedData, uniqueInvestments]);

  const handleSort = (key: 'date' | 'amount' | 'bankName' | 'accountNo' | 'type' | 'balance') => {
    if (sortBy === key) setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    else { setSortBy(key); setSortOrder('desc'); }
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
          'Bank Name': inv.bankName,
          'Account No': inv.accountNo,
          'Opening Date': inv.investmentDate,
          'Principal': inv.principal,
          'Currency': inv.currency,
          'Exchange Rate': inv.exchangeRate || 1,
          'Profit': fdrStats.find(f => f.id === inv.id)?.netProfit || 0,
          'Current Balance': fdrStats.find(f => f.id === inv.id)?.currentBalance || 0,
          'Interest Frequency': inv.interestFrequency,
          'Interest Handling': inv.interestHandling,
          'Status': inv.status,
          'Transaction Date': '',
          'Transaction Type': '',
          'Amount': ''
        });
      } else {
        inv.transactions.forEach(t => {
          rows.push({
            'Bank Name': inv.bankName,
            'Account No': inv.accountNo,
            'Opening Date': inv.investmentDate,
            'Principal': inv.principal,
            'Currency': inv.currency,
            'Exchange Rate': inv.exchangeRate || 1,
            'Profit': fdrStats.find(f => f.id === inv.id)?.netProfit || 0,
            'Current Balance': fdrStats.find(f => f.id === inv.id)?.currentBalance || 0,
            'Interest Frequency': inv.interestFrequency,
            'Interest Handling': inv.interestHandling,
            'Status': inv.status,
            'Transaction Date': t.date,
            'Transaction Type': t.type,
            'Profit Handling': t.handling || '',
            'Amount': t.amount
          });
        });
      }
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Fixed Deposits");
    XLSX.writeFile(wb, `FD_Export_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('success', 'Export successful!');
  };

  const importData = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setConfirmState({
      isOpen: true,
      title: 'Confirm Replace',
      message: 'Are you sure you want to replace all existing Fixed Deposit data with the imported data? This action cannot be undone.',
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
            
            const newInvsMap = new Map<string, FDR>();
            
            jsonData.forEach(row => {
              const bankName = String(row['Bank Name'] || '').trim();
              const accountNo = String(row['Account No'] || '').trim();
              if (!bankName || !accountNo) return;

              const key = `${bankName}-${accountNo}`;
              
              const metadata = {
                bankName,
                accountNo,
                investmentDate: row['Opening Date'] ? String(row['Opening Date']) : undefined,
                principal: row['Principal'] ? Number(row['Principal']) : undefined,
                currency: (row['Currency'] || 'BDT') as any,
                exchangeRate: row['Exchange Rate'] ? Number(row['Exchange Rate']) : undefined,
                interestFrequency: (row['Interest Frequency'] || 'Monthly') as any,
                interestHandling: (row['Interest Handling'] || 'Withdrawn') as any,
                status: row['Status'] ? String(row['Status']) as any : 'Active',
              };

              if (!newInvsMap.has(key)) {
                newInvsMap.set(key, {
                  id: Math.random().toString(36).substr(2, 9),
                  name: `${bankName} (${accountNo})`,
                  bankName: bankName,
                  accountNo: accountNo,
                  investmentDate: metadata.investmentDate || new Date().toISOString().split('T')[0],
                  principal: metadata.principal || 0,
                  currency: metadata.currency || 'BDT',
                  exchangeRate: metadata.exchangeRate,
                  interestFrequency: metadata.interestFrequency || 'Monthly',
                  interestHandling: metadata.interestHandling || 'Withdrawn',
                  status: metadata.status || 'Active',
                  amount: 0,
                  transactions: []
                });
              }
              
              const targetTransactions = newInvsMap.get(key)!.transactions;
              
              if (row['Transaction Date']) {
                const type = row['Transaction Type'] || 'Profit';
                const handling = row['Profit Handling'] || 'Withdrawn';
                const amount = Number(row['Amount'] || 0);
                if (amount === 0) return;

                targetTransactions.push({
                  id: Math.random().toString(36).substr(2, 9),
                  date: String(row['Transaction Date']),
                  type: type as any,
                  handling: type === 'Profit' ? (handling as any) : undefined,
                  amount: amount
                });
              }
            });
            
            const importedInvs = Array.from(newInvsMap.values());
            if (onReplaceAll && importedInvs.length > 0) {
              onReplaceAll(importedInvs);
              showNotification('success', `Successfully replaced data with ${importedInvs.length} deposits!`);
            } else if (onBatchAdd && importedInvs.length > 0) {
              onBatchAdd(importedInvs);
              showNotification('success', `Imported data for ${importedInvs.length} deposits!`);
            } else {
              showNotification('error', 'No valid data found in file.');
            }
          } catch (err) {
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

  const [fdrFormData, setFdrFormData] = useState({
    bankName: '',
    accountNo: '',
    investmentDate: new Date().toISOString().split('T')[0],
    principal: '',
    currency: 'BDT' as 'BDT' | 'USD',
    interestFrequency: 'Monthly' as 'Monthly' | 'Yearly',
    interestHandling: 'Withdrawn' as 'Added' | 'Withdrawn',
    status: 'Active' as 'Active' | 'Closed',
    exchangeRate: '',
    closingDate: '',
    withdrawBalance: '',
  });

  const [transactionFormData, setTransactionFormData] = useState({
    fdrId: '',
    date: new Date().toISOString().split('T')[0],
    type: 'Profit' as any,
    amount: '',
    handling: 'Withdrawn' as 'Added' | 'Withdrawn'
  });

  const allTransactions = useMemo(() => {
    const flat: any[] = [];
    uniqueInvestments.forEach(inv => {
      let balance = inv.principal;
      const sortedTransactions = [...inv.transactions].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
      
      sortedTransactions.forEach(t => {
        if (t.type === 'Profit') {
          if (t.handling === 'Added') balance += t.amount;
        } else {
          balance -= t.amount;
        }
        
        flat.push({
          ...t,
          fdrId: inv.id,
          bankName: inv.bankName,
          accountNo: inv.accountNo,
          currency: inv.currency,
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
      result = result.filter(t => t.bankName.toLowerCase().includes(search) || t.accountNo.toLowerCase().includes(search));
    }

    if (selectedTypes.length > 0) {
      result = result.filter(t => selectedTypes.includes(t.type));
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
      else if (sortBy === 'bankName') {
        comparison = a.bankName.localeCompare(b.bankName);
        if (comparison === 0) comparison = a.accountNo.localeCompare(b.accountNo);
      }
      else if (sortBy === 'accountNo') comparison = a.accountNo.localeCompare(b.accountNo);
      else if (sortBy === 'type') comparison = a.type.localeCompare(b.type);
      else if (sortBy === 'balance') comparison = a.balance - b.balance;
      return sortOrder === 'desc' ? comparison : -comparison;
    });
  }, [allTransactions, sortBy, sortOrder, searchQuery, selectedTypes, timeFilterMode, timeOffset, customRange]);

  const fdrStats = useMemo(() => {
    const startStr = rangeDates.start.toISOString().split('T')[0];
    const endStr = rangeDates.end.toISOString().split('T')[0];
    return uniqueInvestments.map(f => {
      let totalProfitAll = 0;
      let addedProfitAll = 0;
      let chargeAll = 0;
      let totalProfitRange = 0;
      let chargeRange = 0;
      
      f.transactions.forEach(t => {
        if (t.date > endStr) return;
        if (t.type === 'Profit') {
          totalProfitAll += t.amount;
          if (t.handling === 'Added') addedProfitAll += t.amount;
          if (t.date >= startStr) totalProfitRange += t.amount;
        } else if (t.type === 'Charge') {
          chargeAll += t.amount;
          if (t.date >= startStr) chargeRange += t.amount;
        }
      });
      const computedBalance = f.principal + addedProfitAll - chargeAll;
      const currentBalance = f.status === 'Closed' && f.withdrawBalance != null
        ? computedBalance - f.withdrawBalance
        : computedBalance;
      return {
        id: f.id,
        bankName: f.bankName,
        accountNo: f.accountNo,
        principal: f.principal,
        currentBalance,
        currency: f.currency,
        netProfit: totalProfitRange - chargeRange,
        roi: f.principal > 0 ? ((totalProfitRange - chargeRange) / f.principal) * 100 : 0
      };
    });
  }, [uniqueInvestments, rangeDates]);

  const summaryStats = useMemo(() => {
    const startStr = rangeDates.start.toISOString().split('T')[0];
    const endStr = rangeDates.end.toISOString().split('T')[0];

    let totalInvestedBDT = 0;
    let totalInvestedUSD = 0;
    let currentBalanceBDT = 0;
    let currentBalanceUSD = 0;
    let totalProfitBDT = 0;
    let totalProfitUSD = 0;
    let totalChargeBDT = 0;
    let totalChargeUSD = 0;
    let activeInvestedBDT = 0;
    let activeInvestedUSD = 0;

    // Secondary card stats
    let profitAddedBDT = 0;
    let profitAddedUSD = 0;
    let profitWithdrawnBDT = 0;
    let profitWithdrawnUSD = 0;
    let chargesBDT = 0;
    let chargesUSD = 0;
    let holdingBDT = 0;
    let holdingUSD = 0;

    // This month
    const now = new Date();
    const monthStartStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-01`;

    let thisMonthDepositBDT = 0;
    let thisMonthDepositUSD = 0;

    uniqueInvestments.forEach(f => {
      const rate = f.currency === 'USD' ? (f.exchangeRate || 1) : 1;

      if (f.investmentDate <= endStr && f.investmentDate >= startStr) {
        if (f.currency === 'BDT') totalInvestedBDT += f.principal;
        else totalInvestedUSD += f.principal;
      }

      // This month deposits
      if (f.investmentDate >= monthStartStr && f.investmentDate <= endStr) {
        if (f.currency === 'BDT') thisMonthDepositBDT += f.principal;
        else thisMonthDepositUSD += f.principal;
      }

      // Active Investment
      if (f.investmentDate <= endStr) {
        const isClosedByEnd = f.status === 'Closed' && f.closingDate && f.closingDate <= endStr;
        if (!isClosedByEnd) {
          if (f.currency === 'BDT') activeInvestedBDT += f.principal;
          else activeInvestedUSD += f.principal;
        }
      }

      let fAddedProfitAsOfEnd = 0;
      let fChargeAsOfEnd = 0;
      let fTotalProfitRange = 0;
      let fChargeRange = 0;
      let fProfitAddedRange = 0;
      let fProfitWithdrawnRange = 0;

      f.transactions.forEach(t => {
        if (t.date > endStr) return;
        if (t.type === 'Profit') {
          if (t.handling === 'Added') fAddedProfitAsOfEnd += t.amount;
          if (t.date >= startStr) {
            fTotalProfitRange += t.amount;
            if (t.handling === 'Added') fProfitAddedRange += t.amount;
            else fProfitWithdrawnRange += t.amount;
          }
        } else if (t.type === 'Charge') {
          fChargeAsOfEnd += t.amount;
          if (t.date >= startStr) fChargeRange += t.amount;
        }
      });

      if (f.currency === 'BDT') {
        totalProfitBDT += fTotalProfitRange;
        totalChargeBDT += fChargeRange;
        profitAddedBDT += fProfitAddedRange;
        profitWithdrawnBDT += fProfitWithdrawnRange;
        chargesBDT += fChargeRange;
      } else {
        totalProfitUSD += fTotalProfitRange;
        totalChargeUSD += fChargeRange;
        profitAddedUSD += fProfitAddedRange;
        profitWithdrawnUSD += fProfitWithdrawnRange;
        chargesUSD += fChargeRange;
      }

      // Current Holding per currency
      if (f.investmentDate <= endStr) {
        const computedBal = f.principal + fAddedProfitAsOfEnd - fChargeAsOfEnd;
        const wasClosedByEnd = f.status === 'Closed' && f.closingDate && f.closingDate <= endStr;
        const bal = wasClosedByEnd && f.withdrawBalance != null
          ? computedBal - f.withdrawBalance
          : computedBal;

        if (bal > 0) {
          if (f.currency === 'BDT') {
            currentBalanceBDT += bal;
            holdingBDT += bal;
          } else {
            currentBalanceUSD += bal;
            holdingUSD += bal;
          }
        }
      }
    });

    const activeUSDRate = uniqueInvestments.find(f => f.currency === 'USD' && f.exchangeRate)?.exchangeRate || 110;
    const netProfitBDT = totalProfitBDT - totalChargeBDT;
    const netProfitUSD = totalProfitUSD - totalChargeUSD;

    const activeCount = uniqueInvestments.filter(f => f.status === 'Active').length;
    const totalDeposit = (totalInvestedBDT + (totalInvestedUSD * activeUSDRate));
    const dividendReturnPercentage = totalDeposit > 0
      ? ((netProfitBDT + (netProfitUSD * activeUSDRate)) / totalDeposit) * 100
      : 0;

    return {
      totalInvested: {
        total: totalInvestedBDT + (totalInvestedUSD * activeUSDRate),
        bdt: totalInvestedBDT,
        usd: totalInvestedUSD
      },
      activeInvestment: {
        total: activeInvestedBDT + (activeInvestedUSD * activeUSDRate),
        bdt: activeInvestedBDT,
        usd: activeInvestedUSD
      },
      currentBalance: {
        total: currentBalanceBDT + (currentBalanceUSD * activeUSDRate),
        bdt: currentBalanceBDT,
        usd: currentBalanceUSD
      },
      netProfit: {
        total: netProfitBDT + (netProfitUSD * activeUSDRate),
        bdt: netProfitBDT,
        usd: netProfitUSD
      },
      roi: dividendReturnPercentage,
      activeCount,
      totalDepositAmount: totalDeposit,
      thisMonthDeposit: thisMonthDepositBDT + (thisMonthDepositUSD * activeUSDRate),
      // Secondary card values
      holdingBDT,
      holdingUSD,
      profitAddedBDT,
      profitAddedUSD,
      profitWithdrawnBDT,
      profitWithdrawnUSD,
      chargesBDT,
      chargesUSD,
    };
  }, [uniqueInvestments, rangeDates]);

  const handleEditInv = (inv: FDR) => {
    setEditingFDR(inv);
    const statForEdit = fdrStats.find(f => f.id === inv.id);
    setFdrFormData({
      bankName: inv.bankName,
      accountNo: inv.accountNo,
      investmentDate: inv.investmentDate,
      principal: inv.principal?.toString() || '',
      currency: inv.currency as any || 'BDT',
      interestFrequency: inv.interestFrequency || 'Monthly',
      interestHandling: inv.interestHandling || 'Withdrawn',
      status: inv.status,
      exchangeRate: inv.exchangeRate?.toString() || '',
      closingDate: inv.closingDate || '',
      withdrawBalance: inv.withdrawBalance?.toString() || (inv.status === 'Closed' ? statForEdit?.currentBalance?.toString() || '' : ''),
    });
    setIsFDRModalOpen(true);
  };

  const handleFDRSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const payload = {
      bankName: fdrFormData.bankName,
      accountNo: fdrFormData.accountNo,
      investmentDate: fdrFormData.investmentDate,
      principal: parseFloat(fdrFormData.principal) || 0,
      currency: fdrFormData.currency,
      interestFrequency: fdrFormData.interestFrequency,
      interestHandling: fdrFormData.interestHandling,
      status: fdrFormData.status,
      exchangeRate: fdrFormData.currency === 'USD' ? (parseFloat(fdrFormData.exchangeRate) || undefined) : undefined,
      closingDate: fdrFormData.status === 'Closed' ? fdrFormData.closingDate || undefined : undefined,
      withdrawBalance: fdrFormData.status === 'Closed' ? (parseFloat(fdrFormData.withdrawBalance) || undefined) : undefined,
    };

    if (editingFDR) {
      onUpdate(editingFDR.id, payload);
      showNotification('success', 'Fixed Deposit updated successfully!');
    } else {
      const fdNameOnly = `${payload.bankName}${payload.accountNo ? ` (${payload.accountNo})` : ''}`.trim();
      onAdd({
        ...payload,
        name: fdNameOnly,
        amount: payload.principal,
        transactions: [],
      });
      showNotification('success', 'Fixed Deposit added successfully!');

      if (payload.principal > 0 && onNavigateToModule) {
        onNavigateToModule('income-expense', {
          date: payload.investmentDate,
          amount: payload.principal,
          type: 'Transfer',
          targetModule: 'fdrs',
          description: fdNameOnly
        });
      }
    }
    setIsFDRModalOpen(false);
    setEditingFDR(null);
  };

  const handleTransactionSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const fdr = uniqueInvestments.find(f => f.id === transactionFormData.fdrId);
    if (!fdr) return;

    const newTransaction: FDRTransaction = {
      id: editingTransaction ? editingTransaction.transaction.id : Math.random().toString(36).substr(2, 9),
      date: transactionFormData.date,
      type: transactionFormData.type,
      amount: parseFloat(transactionFormData.amount) || 0,
      handling: transactionFormData.type === 'Profit' ? transactionFormData.handling : undefined,
    };

    if (editingTransaction) {
      const updated = fdr.transactions.map(t => t.id === editingTransaction.transaction.id ? newTransaction : t);
      onUpdate(fdr.id, { transactions: updated });
      setEditingTransaction(null);
    } else {
      onUpdate(fdr.id, { transactions: [...fdr.transactions, newTransaction] });

      if (onNavigateToModule && newTransaction.amount > 0) {
        const fdNameOnly = `${fdr.bankName}${fdr.accountNo ? ` (${fdr.accountNo})` : ''}`.trim();
        if (newTransaction.type === 'Profit' && newTransaction.handling === 'Withdrawn') {
          onNavigateToModule('income-expense', {
            date: newTransaction.date,
            amount: newTransaction.amount,
            type: 'Income',
            targetModule: 'fdrs',
            description: `${fdNameOnly} Profit`
          });
        }
      }
    }
    setIsModalOpen(false);
    showNotification('success', 'Transaction recorded successfully!');
  };

  const handleDeleteTransaction = (fdrId: string, tId: string) => {
    const fdr = investments.find(f => f.id === fdrId);
    if (!fdr) return;
    const trans = fdr.transactions.find(t => t.id === tId);

    setConfirmState({
      isOpen: true,
      title: 'Confirm Delete',
      message: `Delete ${trans?.type || 'transaction'} for ${fdr.bankName} (${fdr.accountNo})?`,
      onConfirm: () => {
        closeConfirm();
        onUpdate(fdrId, { transactions: fdr.transactions.filter(t => t.id !== tId) });
        showNotification('success', 'Transaction deleted!');
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
          const [fdrId, tId] = uid.split('-');
          if (!deletions[fdrId]) deletions[fdrId] = [];
          deletions[fdrId].push(tId);
        });
        Object.entries(deletions).forEach(([fdrId, tIds]) => {
          const fdr = investments.find(f => f.id === fdrId);
          if (fdr) onUpdate(fdrId, { transactions: fdr.transactions.filter(t => !tIds.includes(t.id)) });
        });
        setSelectedIds([]);
        showNotification('success', `Successfully deleted ${selectedIds.length} transactions.`);
      },
      onCancel: closeConfirm
    });
  };

  const handleEditTransaction = (t: any) => {
    setEditingTransaction({ fdrId: t.fdrId, transaction: t });
    setTransactionFormData({
      fdrId: t.fdrId,
      date: t.date,
      type: t.type,
      amount: t.amount.toString(),
      handling: t.handling || 'Withdrawn'
    });
    setIsModalOpen(true);
  };

  return (
    <div className="space-y-8">

      {/* Date Range Selector */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 relative">
        <div className="flex flex-col sm:flex-row sm:items-center gap-3 w-full sm:w-auto">
          {/* LINE 1 (Mobile): Range Selector + Settings */}
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
                    {historyRange === 'all' ? 'Overall' : 
                     historyRange === 'last12m' ? 'Last 12M' : 
                     historyRange === 'fiscal' ? 'Fiscal' : 'Custom'}
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
                      {['all', 'last12m', 'fiscal', 'custom'].map((id) => (
                        <button 
                          key={id}
                          onClick={() => {
                            setHistoryRange(id as any);
                            setIsRangeMenuOpen(false);
                          }}
                          className={cn(
                            "w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase",
                            historyRange === id 
                              ? "bg-teal-400 text-slate-950" 
                              : "text-slate-300 hover:bg-slate-800"
                          )}
                        >
                          {id === 'all' ? 'Overall' : 
                           id === 'last12m' ? 'Last 12M' : 
                           id === 'fiscal' ? 'Fiscal' : 'Custom'}
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>

              {/* Desktop Tabs */}
              <div className="hidden sm:flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1">
                {['all', 'last12m', 'fiscal', 'custom'].map((id) => (
                  <button
                    key={id}
                    onClick={() => setHistoryRange(id as any)}
                    className={cn(
                      "px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all border",
                      historyRange === id 
                        ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20" 
                        : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white"
                    )}
                  >
                    {id === 'all' ? 'Overall' : 
                     id === 'last12m' ? 'Last 12M' : 
                     id === 'fiscal' ? 'Fiscal' : 'Custom'}
                  </button>
                ))}
              </div>
            </div>

            {/* Settings Button - Mobile */}
            <div className="block sm:hidden">
              <SettingsButton 
                isOpen={isSettingsMenuOpen} 
                setIsOpen={setIsSettingsMenuOpen} 
                onExport={exportData}
                onImport={importData}
                onTemplate={() => {
                  const templateData = [
                    ['Bank Name', 'Account No', 'Opening Date', 'Principal', 'Currency', 'Exchange Rate', 'Interest Frequency', 'Interest Handling', 'Status', 'Transaction Date', 'Transaction Type', 'Profit Handling', 'Amount'],
                    ['Bank Asia', '12345678', '2025-06-01', 500000, 'BDT', 1, 'Monthly', 'Withdrawn', 'Active', '2025-06-15', 'Profit', 'Withdrawn', 2500],
                  ];
                  const ws = XLSX.utils.aoa_to_sheet(templateData);
                  const wb = XLSX.utils.book_new();
                  XLSX.utils.book_append_sheet(wb, ws, "Template");
                  XLSX.writeFile(wb, "Fixed_Deposit_Import_Template.xlsx");
                }}
                onAdd={() => { setEditingFDR(null); setIsFDRModalOpen(true); }}
              />
            </div>
          </div>

          {/* LINE 2 (Mobile): Custom Date Controls */}
          {historyRange === 'custom' && (
            <div className="flex items-center justify-center sm:justify-start gap-1.5 px-1 animate-in fade-in slide-in-from-top-2 sm:slide-in-from-left-2 duration-300 w-full sm:w-auto">
              <button 
                onClick={() => navigateCustomMonth(-1)} 
                className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400"
              >
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

              <button 
                onClick={() => navigateCustomMonth(1)} 
                className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400"
              >
                <ChevronRight size={14} />
              </button>
            </div>
          )}
        </div>

        {/* Settings Button - Desktop */}
        <div className="hidden sm:block">
          <SettingsButton 
            isOpen={isSettingsMenuOpen} 
            setIsOpen={setIsSettingsMenuOpen} 
            onExport={exportData}
            onImport={importData}
            onTemplate={() => {
              const templateData = [
                ['Bank Name', 'Account No', 'Opening Date', 'Principal', 'Currency', 'Exchange Rate', 'Interest Frequency', 'Interest Handling', 'Status', 'Transaction Date', 'Transaction Type', 'Profit Handling', 'Amount'],
                ['Bank Asia', '12345678', '2025-06-01', 500000, 'BDT', 1, 'Monthly', 'Withdrawn', 'Active', '2025-06-15', 'Profit', 'Withdrawn', 2500],
              ];
              const ws = XLSX.utils.aoa_to_sheet(templateData);
              const wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, ws, "Template");
              XLSX.writeFile(wb, "Fixed_Deposit_Import_Template.xlsx");
            }}
            onAdd={() => { setEditingFDR(null); setIsFDRModalOpen(true); }}
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
        <div className={cn("flex items-center gap-3 p-4 rounded-xl border animate-in fade-in slide-in-from-top-4", notification.type === 'success' ? "bg-teal-400/10 border-teal-400/20 text-teal-400" : "bg-rose-500/10 border-rose-500/20 text-rose-500")}>
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
                <InfoTooltip text="Total Current Balance in All Accounts" />
              </div>
              <p className="text-label font-bold text-slate-500 uppercase">Current Balance</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.currentBalance.total)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Active FD</p>
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
                <InfoTooltip text="Total Principal in All Accounts" />
              </div>
              <p className="text-label font-bold text-slate-500 uppercase">Active Principal</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.activeInvestment.total)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Deposit</p>
            <p className="text-body font-bold text-blue-400 tabular-nums">
              {formatBDT(summaryStats.totalInvested.total)}
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
              <p className="text-label font-bold text-slate-500 uppercase">Total Principal</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.totalInvested.total)}
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
                <InfoTooltip text="Total Profit Received from All Accounts" />
              </div>
              <p className="text-label font-bold text-slate-500 uppercase">Added + Withdrawn</p>
            </div>
            <h3 className="text-heading font-bold text-emerald-400 mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(summaryStats.netProfit.total)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Portfolio Yield</p>
            <p className={cn("text-body font-bold tabular-nums", summaryStats.roi >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {summaryStats.roi.toFixed(2)}%
            </p>
          </div>
        </Card>

      </div>

      {/* ── 8 Secondary Cards ── matching MF module style */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">

        {/* Holding (BDT) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Holding (BDT)</p>
            <InfoTooltip text="Total Current Holding (BDT Portion)" />
          </div>
          <p className="text-body font-bold text-white tabular-nums">{formatBDT(summaryStats.holdingBDT)}</p>
        </Card>

        {/* Holding (USD) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Holding (USD)</p>
            <InfoTooltip text="Total Current Holding (USD Portion)" />
          </div>
          <p className="text-body font-bold text-white tabular-nums">{formatCurrency(summaryStats.holdingUSD, 'USD')}</p>
        </Card>

        {/* Profit Added (BDT) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Profit Added (BDT)</p>
            <InfoTooltip text="Total Profit Added (BDT Portion)" />
          </div>
          <p className="text-body font-bold text-emerald-400 tabular-nums">{formatBDT(summaryStats.profitAddedBDT)}</p>
        </Card>

        {/* Profit Added (USD) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Profit Added (USD)</p>
            <InfoTooltip text="Total Profit Added (USD Portion)" />
          </div>
          <p className="text-body font-bold text-emerald-400 tabular-nums">{formatCurrency(summaryStats.profitAddedUSD, 'USD')}</p>
        </Card>

        {/* Profit Withdrawn (BDT) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Profit Withdrawn (BDT)</p>
            <InfoTooltip text="Total Profit Withdrawn (BDT Portion)" />
          </div>
          <p className="text-body font-bold text-emerald-400 tabular-nums">{formatBDT(summaryStats.profitWithdrawnBDT)}</p>
        </Card>

        {/* Profit Withdrawn (USD) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Profit Withdrawn (USD)</p>
            <InfoTooltip text="Total Profit Withdrawn (USD Portion)" />
          </div>
          <p className="text-body font-bold text-emerald-400 tabular-nums">{formatCurrency(summaryStats.profitWithdrawnUSD, 'USD')}</p>
        </Card>

        {/* Charges (BDT) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Charges (BDT)</p>
            <InfoTooltip text="Total Charges Deducted (BDT Portion)" />
          </div>
          <p className="text-body font-bold text-rose-400 tabular-nums">{formatBDT(summaryStats.chargesBDT)}</p>
        </Card>

        {/* Charges (USD) */}
        <Card className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group" style={{ padding: '12px' }}>
          <div className="flex items-center justify-between mb-1">
            <p className="text-label font-bold text-slate-300 uppercase">Charges (USD)</p>
            <InfoTooltip text="Total Charges Deducted (USD Portion)" />
          </div>
          <p className="text-body font-bold text-rose-400 tabular-nums">{formatCurrency(summaryStats.chargesUSD, 'USD')}</p>
        </Card>

      </div>

      {/* FD Cards */}
      {fdrStats.map(fdr => (
        <Card key={fdr.id} className="bg-slate-900 border-slate-800">
          <div className="flex items-center justify-between mb-6">
            <div className="flex items-center gap-2">
              <Landmark className="text-teal-400 w-5 h-5" />
              <h3 className="text-subheading font-bold text-white uppercase">{fdr.bankName} {fdr.accountNo} ({fdr.currency})</h3>
            </div>
            <div className="flex items-center gap-1">
              <button 
                onClick={() => { const d = uniqueInvestments.find(i => i.id === fdr.id); if (d) handleEditInv(d); }} 
                className="p-1.5 text-slate-500 hover:text-teal-400 hover:bg-teal-400/10 rounded-lg"
              >
                <Edit2 size={14} />
              </button>
              <button 
                onClick={() => {
                  setConfirmState({
                    isOpen: true,
                    title: 'Confirm Delete',
                    message: 'Are you sure you want to delete this Fixed Deposit and all its transactions?',
                    onConfirm: () => {
                      closeConfirm();
                      onDelete(fdr.id);
                      showNotification('success', 'Fixed Deposit deleted successfully!');
                    },
                    onCancel: closeConfirm
                  });
                }} 
                className="p-1.5 text-slate-500 hover:text-rose-500 hover:bg-rose-500/10 rounded-lg"
              >
                <Trash2 size={14} />
              </button>
            </div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
            <div className="bg-slate-950/50 border border-slate-800/50 rounded-xl p-4 transition-all hover:border-teal-400/30">
              <p className="text-label font-bold text-slate-300 uppercase mb-1">Current Balance</p>
              <p className="text-body font-bold text-white tabular-nums">{formatCurrency(fdr.currentBalance, fdr.currency)}</p>
            </div>
            <div className={cn("bg-slate-950/50 border border-slate-800/50 rounded-xl p-4 transition-all", fdr.netProfit >= 0 ? "hover:border-emerald-400/30" : "hover:border-rose-500/30")}>
              <p className="text-label font-bold text-slate-300 uppercase mb-1">Total Profit</p>
              <p className={cn("text-body font-bold tabular-nums", fdr.netProfit >= 0 ? "text-emerald-400" : "text-rose-500")}>
                {formatCurrency(fdr.netProfit, fdr.currency)}
              </p>
            </div>
            <div className="col-span-2 md:col-span-1 bg-slate-950/50 border border-slate-800/50 rounded-xl p-4 transition-all hover:border-amber-400/30">
              <p className="text-label font-bold text-slate-300 uppercase mb-1">Profit (%)</p>
              <p className={cn("text-body font-bold tabular-nums", fdr.roi >= 0 ? "text-emerald-400" : "text-rose-500")}>{fdr.roi.toFixed(2)}%</p>
            </div>
          </div>
        </Card>
      ))}

      {/* Portfolio Table */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-6">
        <div className="flex items-center gap-2 sm:gap-3">
          <Landmark className="text-teal-400 w-5 h-5 lg:w-6 lg:h-6" />
          <h2 className="text-body-sm sm:text-subheading lg:text-heading font-bold text-white font-display uppercase whitespace-nowrap">Fixed Deposit Profits</h2>
        </div>
        <div className="flex items-center gap-2 flex-1 sm:flex-none justify-end">
          <div className="relative bg-slate-900 border border-slate-800 rounded-lg px-3 h-9 flex items-center gap-2 group focus-within:border-teal-400/50 flex-1 sm:flex-none">
            <Search size={14} className="text-slate-500 group-focus-within:text-teal-400 shrink-0" />
            <input 
              type="text" 
              placeholder="Search" 
              className="bg-transparent border-none focus:ring-0 text-[10px] font-bold text-white uppercase outline-none w-full sm:w-32 md:w-48" 
              value={searchQuery} 
              onChange={(e) => setSearchQuery(e.target.value)} 
            />
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
              <span className="text-slate-500 text-[10px] uppercase font-bold">to</span>
              <input 
                type="date"
                className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1 text-[10px] font-bold text-white uppercase outline-none focus:border-teal-400/50"
                value={customRange.end}
                onChange={(e) => setCustomRange(prev => ({ ...prev, end: e.target.value }))}
              />
            </div>
          </div>
        )}
      </div>

      {/* Transaction Table */}
      <div className="flex flex-col border border-slate-800 rounded-xl overflow-hidden bg-slate-900/50 shadow-2xl">
        <div className="overflow-x-auto custom-scrollbar">
          <div className="min-w-[900px] grid grid-cols-[48px_1.3fr_0.8fr_1.3fr_2fr_1.1fr_1.1fr_80px]">
            <div className="flex items-center justify-center border-r border-b border-slate-800 bg-slate-900/80 py-3 sticky top-0 z-10"><Checkbox checked={selectedIds.length === filtered.length && filtered.length > 0} onChange={() => setSelectedIds(selectedIds.length === filtered.length ? [] : filtered.map(t => t.uniqueId))} /></div>
            {[
              { label: 'Bank Account', id: 'bankName' }, 
              { label: 'Month', id: 'date' }, 
              { label: 'Date', id: 'date' }, 
              { label: 'Type', id: 'type' }, 
              { label: 'Amount', id: 'amount' }, 
              { label: 'Balance', id: 'balance' }, 
              { label: '', id: null }
            ].map((h, i) => (
              <div 
                key={i} 
                className={cn(
                  "px-4 py-3 border-r border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 text-label font-bold text-slate-300 uppercase tracking-widest flex items-center gap-1",
                  h.id ? "cursor-pointer hover:text-white" : ""
                )}
                onClick={() => h.id && handleSort(h.id as any)}
              >
                {h.label}
                {h.id && (
                  <ArrowUpDown 
                    size={12} 
                    className={cn(
                      "transition-transform", 
                      sortBy === h.id ? "text-teal-400" : "text-slate-600 opacity-50",
                      sortBy === h.id && sortOrder === 'asc' ? 'rotate-180' : ''
                    )} 
                  />
                )}
              </div>
            ))}

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
              const monthName = new Date(t.date).toLocaleString('default', { month: 'short', year: 'numeric' });
              return (
                <React.Fragment key={t.uniqueId}>
                  <div className={cn("flex items-center justify-center border-r border-b border-slate-800/50 py-3 transition-colors", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <Checkbox checked={isSelected} onChange={() => setSelectedIds(prev => prev.includes(t.uniqueId) ? prev.filter(id => id !== t.uniqueId) : [...prev, t.uniqueId])} />
                  </div>
                  <div className={cn("px-4 py-3 border-r border-b border-slate-800/50 transition-colors flex flex-col justify-center", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <span className="text-body-sm font-bold text-slate-300 uppercase truncate">{t.bankName} {t.accountNo}</span>
                    <span className="text-label text-teal-400 font-bold uppercase tabular-nums">{t.currency} ({t.currency === 'USD' ? '$' : '৳'})</span>
                  </div>
                  <div className={cn("px-4 py-3 border-r border-b border-slate-800/50 transition-colors flex items-center", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <span className="text-body-sm font-bold text-slate-300 tabular-nums uppercase">{monthName}</span>
                  </div>
                  <div className={cn("px-4 py-3 border-r border-b border-slate-800/50 transition-colors flex items-center", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <span className="text-body-sm font-bold text-slate-300 tabular-nums uppercase">{formatDate(t.date)}</span>
                  </div>
                  <div className={cn("px-4 py-3 border-r border-b border-slate-800/50 transition-colors flex flex-row flex-wrap items-center gap-1.5", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <div className={cn("flex items-center px-2 py-0.5 rounded-md w-fit border", t.type === 'Profit' ? "bg-emerald-400/5 text-emerald-400 border-emerald-400/20" : "bg-rose-500/5 text-rose-500 border-rose-500/20")}>
                      <span className="text-[9px] font-bold uppercase tracking-widest">{t.type}</span>
                    </div>
                    {t.type === 'Profit' && (
                      <div className={cn("flex items-center px-2 py-0.5 rounded-md w-fit border", t.handling === 'Added' ? "bg-teal-400/5 text-teal-400 border-teal-400/20" : "bg-slate-800/50 text-slate-400 border-slate-700")}>
                        <span className="text-[8px] font-bold uppercase tracking-wider">{t.handling === 'Added' ? 'Added to capital' : 'Withdrawn'}</span>
                      </div>
                    )}
                  </div>
                  <div className={cn(
                    "px-4 py-3 border-r border-b border-slate-800/50 transition-colors flex items-center tabular-nums font-bold text-body-sm tracking-tight", 
                    isSelected ? "bg-teal-400/10" : "bg-slate-900/40",
                    t.type === 'Profit' ? "text-emerald-400" : "text-rose-500"
                  )}>
                    {formatCurrency(t.amount, t.currency)}
                  </div>
                  <div className={cn("px-4 py-3 border-r border-b border-slate-800/50 transition-colors flex items-center tabular-nums text-body-sm text-slate-300 font-bold", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>{formatCurrency(t.balance, t.currency)}</div>
                  <div className={cn("flex items-center gap-1 px-4 py-3 border-b border-slate-800/50 transition-colors", isSelected ? "bg-teal-400/10" : "bg-slate-900/40")}>
                    <button onClick={() => handleEditTransaction(t)} className="p-2 text-slate-500 hover:text-teal-400 rounded-lg"><Edit2 size={16} /></button>
                    <button onClick={() => handleDeleteTransaction(t.fdrId, t.id)} className="p-2 text-slate-500 hover:text-rose-500 rounded-lg"><Trash2 size={16} /></button>
                  </div>
                </React.Fragment>
              );
            })}
          </div>
        </div>
        {filtered.length === 0 && <div className="py-12 text-center text-slate-500 italic uppercase bg-slate-900/30">No transactions found</div>}
      </div>

      {/* FDR Modal */}
      <Modal isOpen={isFDRModalOpen} onClose={() => setIsFDRModalOpen(false)} title={editingFDR ? "Edit Fixed Deposit" : "Add New FD"}>
        <form onSubmit={handleFDRSubmit} className="space-y-4">
          <Input label="Bank Name" value={fdrFormData.bankName} onChange={e => setFdrFormData({...fdrFormData, bankName: e.target.value})} required />
          <Input label="Account Number" value={fdrFormData.accountNo} onChange={e => setFdrFormData({...fdrFormData, accountNo: e.target.value})} required />
          <div className="grid grid-cols-2 gap-3">
            <Input label="Opening Date" type="date" value={fdrFormData.investmentDate} onChange={e => setFdrFormData({...fdrFormData, investmentDate: e.target.value})} required />
            <Select label="Currency" value={fdrFormData.currency} options={[{label:'BDT',value:'BDT'},{label:'USD',value:'USD'}]} onChange={val => setFdrFormData({...fdrFormData, currency: val as any})} />
          </div>
          <Input label="Principal Amount" type="number" step="0.01" value={fdrFormData.principal} onChange={e => setFdrFormData({...fdrFormData, principal: e.target.value})} required />
          <div className="grid grid-cols-2 gap-3">
            <Select label="Interest Frequency" value={fdrFormData.interestFrequency} options={[{label:'Monthly',value:'Monthly'},{label:'Yearly',value:'Yearly'}]} onChange={val => setFdrFormData({...fdrFormData, interestFrequency: val as any})} />
            <Select label="Interest Handling" value={fdrFormData.interestHandling} options={[{label:'Interest Added',value:'Added'},{label:'Interest Withdrawn',value:'Withdrawn'}]} onChange={val => setFdrFormData({...fdrFormData, interestHandling: val as any})} />
          </div>
          <Select
            label="Status"
            value={fdrFormData.status}
            options={[{label:'Active',value:'Active'},{label:'Closed',value:'Closed'}]}
            onChange={val => {
              const newStatus = val as 'Active' | 'Closed';
              if (newStatus === 'Closed') {
                const statForThisFdr = editingFDR ? fdrStats.find(f => f.id === editingFDR.id) : null;
                const defaultWithdraw = statForThisFdr ? statForThisFdr.currentBalance.toString() : '';
                setFdrFormData({
                  ...fdrFormData,
                  status: newStatus,
                  withdrawBalance: fdrFormData.withdrawBalance || defaultWithdraw,
                });
              } else {
                setFdrFormData({ ...fdrFormData, status: newStatus, closingDate: '', withdrawBalance: '' });
              }
            }}
          />

          {fdrFormData.status === 'Closed' && (
            <div className="grid grid-cols-2 gap-3 animate-in fade-in slide-in-from-top-2 duration-200">
              <Input
                label="Closing Date"
                type="date"
                value={fdrFormData.closingDate}
                onChange={e => setFdrFormData({...fdrFormData, closingDate: e.target.value})}
              />
              <Input
                label="Withdraw Balance"
                type="number"
                step="0.01"
                value={fdrFormData.withdrawBalance}
                onChange={e => setFdrFormData({...fdrFormData, withdrawBalance: e.target.value})}
                placeholder="Current balance"
              />
            </div>
          )}

          {fdrFormData.currency === 'USD' && (
            <Input label="Exchange Rate (1 USD = ? BDT)" type="number" step="0.01" value={fdrFormData.exchangeRate} onChange={e => setFdrFormData({...fdrFormData, exchangeRate: e.target.value})} required />
          )}
          <div className="flex justify-end gap-3 pt-4">
            <Button variant="secondary" onClick={() => setIsFDRModalOpen(false)}>Cancel</Button>
            <Button type="submit">{editingFDR ? "Update" : "Save"}</Button>
          </div>
        </form>
      </Modal>

      {/* Transaction Modal */}
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title="Record Transaction">
        <form onSubmit={handleTransactionSubmit} className="space-y-4">
          <Select 
            label="FD Account" 
            value={transactionFormData.fdrId} 
            options={uniqueInvestments.map(f => ({ label: `${f.bankName} ${f.accountNo}`, value: f.id }))} 
            placeholder="Select Account" 
            onChange={val => {
              const fdr = uniqueInvestments.find(f => f.id === val);
              setTransactionFormData({
                ...transactionFormData, 
                fdrId: val,
                handling: fdr ? fdr.interestHandling : transactionFormData.handling
              });
            }} 
            required 
          />
          <div className="grid grid-cols-2 gap-3">
            <Input label="Date" type="date" value={transactionFormData.date} onChange={e => setTransactionFormData({...transactionFormData, date: e.target.value})} required />
            <Select label="Type" value={transactionFormData.type} options={[{label:'Profit',value:'Profit'},{label:'Charge',value:'Charge'}]} onChange={val => setTransactionFormData({...transactionFormData, type: val as any})} required />
          </div>
          {transactionFormData.type === 'Profit' && (
            <Select 
              label="Profit Handling" 
              value={transactionFormData.handling} 
              options={[{label:'Added to FDR',value:'Added'},{label:'Withdrawn/Cash',value:'Withdrawn'}]} 
              onChange={val => setTransactionFormData({...transactionFormData, handling: val as any})} 
            />
          )}
          <Input label="Amount" placeholder="0.00"
 type="number" step="0.01" value={transactionFormData.amount} onChange={e => setTransactionFormData({...transactionFormData, amount: e.target.value})} required />
          <div className="flex justify-end gap-3 pt-4">
            <Button variant="secondary" onClick={() => setIsModalOpen(false)}>Cancel</Button>
            <Button type="submit">Record</Button>
          </div>
        </form>
      </Modal>
    </div>
  );
};