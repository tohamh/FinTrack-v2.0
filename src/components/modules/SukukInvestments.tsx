/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { Card, Button, Modal, Input, Checkbox, Select } from '../ui/BaseComponents';
import { ConfirmDialog } from '../ui/ConfirmDialog';
import { Sukuk, OnlineInvestmentStatus, InvestmentFrequency } from '../../types';
import { formatBDT, formatDate, cn, toDateStr, getTodayStr, getFirstOfMonth, getLastOfMonth } from '../../utils/formatters';
import * as XLSX from 'xlsx';
import { Plus, MoreVertical, Calendar, TrendingUp, Search, Clock, CheckCircle2, AlertCircle, Edit2, X, Briefcase, Settings, Download, Upload, FileSpreadsheet, XCircle, DollarSign, Coins, ChevronDown, Wallet, ChevronLeft, ChevronRight } from 'lucide-react';

interface SukukInvestmentsProps {
  investments: Sukuk[];
  onAdd: (investment: Omit<Sukuk, 'id'>) => void;
  onUpdate: (id: string, updates: Partial<Sukuk>) => void;
  onDelete: (id: string) => void;
  onBatchDelete: (ids: string[]) => void;
  onBatchAdd?: (investments: Omit<Sukuk, 'id'>[]) => void;
  onReplaceAll?: (investments: Omit<Sukuk, 'id'>[]) => void;
  onTitleChange?: (title: React.ReactNode) => void;
  onActiveInvestmentChange?: (amount: number) => void; // [CHANGE 3] new callback
  triggerAdd?: boolean;
  setTriggerAdd?: (val: boolean) => void;
  inheritedData?: { date: string; amount: number } | null;
  onClearInheritedData?: () => void;
  onNavigateToModule?: (module: string, initialAddData?: { date: string; amount: number; type?: 'Transfer' | 'Income' | 'Expense' | 'Loan'; targetModule?: string; description?: string }) => void;
}

export const SukukInvestments: React.FC<SukukInvestmentsProps> = ({ 
  investments, 
  onAdd, 
  onUpdate,
  onDelete,
  onBatchDelete,
  onBatchAdd,
  onReplaceAll,
  onTitleChange,
  onActiveInvestmentChange, // [CHANGE 3]
  triggerAdd,
  setTriggerAdd,
  inheritedData,
  onClearInheritedData,
  onNavigateToModule
}) => {
  const [selectedProfiles, setSelectedProfiles] = useState<OnlineInvestmentStatus[]>(['Active', 'Delayed', 'Matured']);
  const [sortBy, setSortBy] = useState<'date' | 'amount' | 'roi' | 'issuer'>('date');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [searchQuery, setSearchQuery] = useState('');
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingActualId, setEditingActualId] = useState<string | null>(null);
  const [activeMenuId, setActiveMenuId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isProfileMenuOpen, setIsProfileMenuOpen] = useState(false);
  const [isSortMenuOpen, setIsSortMenuOpen] = useState(false);
  const [isSettingsMenuOpen, setIsSettingsMenuOpen] = useState(false);
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
  const [customRepaymentIdx, setCustomRepaymentIdx] = useState<number | null>(null);
  const [customRepaymentData, setCustomRepaymentData] = useState({ date: '', amount: '' });
  const [activeWarning, setActiveWarning] = useState<{ id: string, text: string } | null>(null);
  const [formError, setFormError] = useState<string | null>(null);
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
  const [isHistoryMenuOpen, setIsHistoryMenuOpen] = useState(false);
  
  const [formData, setFormData] = useState({
    name: '',
    instrumentNo: '',
    issuer: 'Bangladesh Bank',
    rentRate: '',
    tds: '10',
    frequency: 'Semi-annual' as InvestmentFrequency,
    principalAmount: '',
    issueDate: new Date().toISOString().split('T')[0],
    durationYears: '',
    status: 'Active' as OnlineInvestmentStatus,
    closingDate: '',
    withdrawBalance: '',
  });

  // Handle external trigger for adding
  React.useEffect(() => {
    if (triggerAdd) {
      setEditingId(null);
      const inheritedDate = inheritedData?.date || new Date().toISOString().split('T')[0];
      const inheritedAmt = (inheritedData?.amount !== undefined && inheritedData.amount !== null && inheritedData.amount !== 0) ? String(inheritedData.amount) : '';
      setFormData({
        name: '',
        instrumentNo: '',
        issuer: 'Bangladesh Bank',
        rentRate: '',
        tds: '10',
        frequency: 'Semi-annual',
        principalAmount: inheritedAmt,
        issueDate: inheritedDate,
        durationYears: '',
        status: 'Active',
        closingDate: '',
        withdrawBalance: '',
      });
      setIsModalOpen(true);
      setTriggerAdd?.(false);
      onClearInheritedData?.();
    }
  }, [triggerAdd, setTriggerAdd, inheritedData, onClearInheritedData]);

  const getEffectiveStatus = (inv: Sukuk): OnlineInvestmentStatus => {
    // If explicitly marked as Matured in the DB, respect that first
    if (inv.status === 'Matured') return 'Matured';

    const today = new Date().toISOString().split('T')[0];
    const allPaid = inv.installments.every(i => i.isPaid);
    if (allPaid) return 'Matured' as OnlineInvestmentStatus;
    
    const hasDelayed = inv.installments.some(i => !i.isPaid && i.date < today);
    if (hasDelayed) return 'Delayed';
    
    return 'Active';
  };

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 5000);
  };

  const exportData = () => {
    const rows: any[] = [];
    investments.forEach(inv => {
      const baseData = {
        'Sukuk ID': inv.id,
        'Sukuk Name': inv.name,
        'Instrument No': inv.instrumentNo,
        'Issuer': inv.issuer,
        'Principal Amount': inv.principalAmount,
        'Rent Rate (%)': inv.rentRate,
        'TDS (%)': inv.tds,
        'Frequency': inv.frequency,
        'Issue Date': inv.issueDate,
        'Duration (Years)': inv.durationYears,
        'Status': getEffectiveStatus(inv)
      };

      inv.installments.forEach((inst, i) => {
        rows.push({
          ...baseData,
          'Installment No': i + 1,
          'Scheduled Date': inst.date,
          'Scheduled Amount': inst.amount,
          'Actual Paid Date': inst.isPaid ? (inst.actualDate || inst.date) : '',
          'Actual Paid Amount': inst.isPaid ? (inst.actualAmount || inst.amount) : '',
          'Payment Status': inst.isPaid ? 'Paid' : 'Unpaid'
        });
      });
    });

    const headers = [
      'Sukuk ID', 'Sukuk Name', 'Instrument No', 'Issuer', 'Principal Amount', 
      'Rent Rate (%)', 'TDS (%)', 'Frequency', 'Issue Date', 'Duration (Years)', 
      'Status', 'Installment No', 'Scheduled Date', 'Scheduled Amount', 
      'Actual Paid Date', 'Actual Paid Amount', 'Payment Status'
    ];

    const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sukuk");
    XLSX.writeFile(wb, `Sukuk_Export_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('success', 'Export successful!');
  };

  const importData = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setConfirmState({
      isOpen: true,
      title: 'Confirm Replace',
      message: 'Are you sure you want to replace all existing Sukuk data with the imported data? This action cannot be undone.',
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

            const groups: { [key: string]: any[] } = {};
            jsonData.forEach(row => {
              const id = row['Sukuk ID'] || row['SukukID'] || 'UNKNOWN';
              if (!groups[id]) groups[id] = [];
              groups[id].push(row);
            });

            const parseExcelDate = (val: any) => {
              if (!val) return '';
              if (typeof val === 'number') {
                const date = new Date((val - 25569) * 86400 * 1000);
                return date.toISOString().split('T')[0];
              }
              const str = val.toString().trim().replace(/\//g, '-');
              if (!str) return '';
              const d = new Date(str);
              return isNaN(d.getTime()) ? '' : d.toISOString().split('T')[0];
            };

            const newInvestments: Omit<Sukuk, 'id'>[] = [];
            Object.values(groups).forEach(rows => {
              const firstRow = rows[0];
              const name = firstRow['Sukuk Name'] || firstRow['name'] || '';
              const instrumentNo = firstRow['Instrument No'] || firstRow['instrumentNo'] || '';
              const issuer = firstRow['Issuer'] || firstRow['issuer'] || '';
              const principalAmount = parseFloat(firstRow['Principal Amount'] || firstRow['amount'] || '0');
              const rentRate = parseFloat(firstRow['Rent Rate (%)'] || firstRow['rentRate'] || '0');
              const tds = parseFloat(firstRow['TDS (%)'] || firstRow['tds'] || '10');
              const frequency = (firstRow['Frequency'] || firstRow['frequency'] || 'Semi-annual') as InvestmentFrequency;
              const issueDate = parseExcelDate(firstRow['Issue Date'] || firstRow['issueDate']);
              const durationYears = parseInt(firstRow['Duration (Years)'] || firstRow['durationYears'] || '0');
              const status = firstRow['Status'] || firstRow['status'] || 'Active';

              if (!name || isNaN(principalAmount) || !issueDate) return;

              const installments = rows.map((row, idx) => {
                const scheduledDate = parseExcelDate(row['Scheduled Date'] || row['date']);
                const scheduledAmount = parseFloat(row['Scheduled Amount'] || row['amount'] || '0');
                const actualDate = parseExcelDate(row['Actual Paid Date'] || row['actualDate']);
                const actualAmount = parseFloat(row['Actual Paid Amount'] || row['actualAmount'] || '0');
                const isPaid = row['Payment Status'] === 'Paid' || !!actualDate;

                return {
                  date: scheduledDate,
                  amount: scheduledAmount,
                  isPaid,
                  actualDate: actualDate || undefined,
                  actualAmount: actualAmount || undefined,
                  installmentNo: idx + 1
                };
              });

              newInvestments.push({
                name,
                investmentDate: issueDate,
                amount: principalAmount,
                currency: 'BDT',
                instrumentNo,
                issuer,
                rentRate,
                tds,
                frequency,
                principalAmount,
                issueDate,
                durationYears,
                status: status as any,
                installments,
                totalRepaid: installments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount || inst.amount) : sum, 0)
              });
            });

            if (onReplaceAll && newInvestments.length > 0) {
              onReplaceAll(newInvestments);
              showNotification('success', `Successfully replaced all data with ${newInvestments.length} Sukuk investments!`);
            } else if (newInvestments.length > 0) {
              // Fallback if onReplaceAll is not provided (shouldn't happen with new App.tsx changes)
              newInvestments.forEach(inv => onAdd(inv));
              showNotification('success', `Successfully imported ${newInvestments.length} Sukuk investments!`);
            }
          } catch (err) {
            console.error('Import failed', err);
            showNotification('error', 'Failed to import file. Please check the format.');
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

  const filtered = useMemo(() => {
    let result = selectedProfiles.length === 0 
      ? [...investments] 
      : investments.filter(i => selectedProfiles.includes(getEffectiveStatus(i)));

    if (searchQuery.trim()) {
      const search = searchQuery.toLowerCase();
      result = result.filter(i => 
        i.name.toLowerCase().includes(search) || 
        i.instrumentNo.toLowerCase().includes(search) ||
        i.issuer.toLowerCase().includes(search)
      );
    }

    return result.sort((a, b) => {
      let comparison = 0;
      if (sortBy === 'date') comparison = new Date(b.issueDate).getTime() - new Date(a.issueDate).getTime();
      else if (sortBy === 'amount') comparison = b.principalAmount - a.principalAmount;
      else if (sortBy === 'roi') comparison = b.rentRate - a.rentRate;
      else if (sortBy === 'issuer') comparison = a.issuer.localeCompare(b.issuer);
      
      return sortOrder === 'desc' ? comparison : -comparison;
    });
  }, [investments, selectedProfiles, sortBy, sortOrder, searchQuery]);

  const upcomingInstallments = useMemo(() => {
    const today = new Date();
    const todayStr = today.toISOString().split('T')[0];

    // AFTER
const all = investments
  .filter(inv => getEffectiveStatus(inv) !== 'Matured')  // exclude matured funds
  .flatMap(inv => {
    return inv.installments.map(inst => ({
      ...inst,
      sukukName: inv.name,
      issuer: inv.issuer,
      isDelayed: !inst.isPaid && inst.date < todayStr
    }));
  });

    const unpaid = all
      .filter(inst => !inst.isPaid)
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

    return unpaid.slice(0, 4);
  }, [investments]);

  const allIssuers = useMemo(() => {
    const issuers = new Set(investments.map(inv => inv.issuer).filter(Boolean));
    return Array.from(issuers).sort();
  }, [investments]);

  const historyStats = useMemo(() => {
    let startDate: Date;
    let endDate = new Date();
    const now = new Date();

    if (historyRange === 'all') {
      startDate = new Date(0);
    } else if (historyRange === 'last12m') {
      startDate = new Date();
      startDate.setFullYear(now.getFullYear() - 1);
    } else if (historyRange === 'fiscal') {
      const currentYear = now.getFullYear();
      if (now.getMonth() >= 6) {
        startDate = new Date(currentYear - 1, 6, 1);
        endDate = new Date(currentYear, 5, 30);
      } else {
        startDate = new Date(currentYear - 2, 6, 1);
        endDate = new Date(currentYear - 1, 5, 30);
      }
    } else {
      startDate = historyCustomDates.start ? new Date(historyCustomDates.start) : new Date(0);
      endDate = historyCustomDates.end ? new Date(historyCustomDates.end) : new Date();
    }

    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];

    const invested = investments
      .filter(inv => inv.issueDate >= startStr && inv.issueDate <= endStr)
      .reduce((sum, inv) => sum + inv.principalAmount, 0);

   
const activeSukuks = investments.filter(inv => {
  // Must have been issued on or before the range end
  if (inv.issueDate > endStr) return false;

  // Compute maturity date from issueDate + durationYears
  const maturityDate = new Date(inv.issueDate);
  maturityDate.setFullYear(maturityDate.getFullYear() + inv.durationYears);
  const maturityStr = maturityDate.toISOString().split('T')[0];

  // Use closingDate if explicitly set (for manually-matured entries)
  const closingStr = (inv as any).closingDate as string | undefined;
  const effectiveClosingStr = closingStr || maturityStr;

  // Active "as of endStr" means: not yet closed/matured by endStr
  if (effectiveClosingStr < endStr) {
    // It had already matured before the range end —
    // only count it if there's remaining principal not yet withdrawn
    const withdrawBalance = (inv as any).withdrawBalance as number | undefined;
    const hasRemainingBalance = withdrawBalance === undefined || withdrawBalance < inv.principalAmount;
    return hasRemainingBalance;
  }

  return true;
});

const active = activeSukuks.reduce((sum, inv) => {
  const maturityDate = new Date(inv.issueDate);
  maturityDate.setFullYear(maturityDate.getFullYear() + inv.durationYears);
  const maturityStr = maturityDate.toISOString().split('T')[0];
  const closingStr = (inv as any).closingDate as string | undefined;
  const effectiveClosingStr = closingStr || maturityStr;

  if (effectiveClosingStr < endStr) {
    // Already matured within or before the range — net out any withdrawal
    const withdrawBalance = (inv as any).withdrawBalance as number | undefined;
    const netPrincipal = withdrawBalance !== undefined
      ? inv.principalAmount - withdrawBalance
      : inv.principalAmount;
    return sum + netPrincipal;
  }
  return sum + inv.principalAmount;
}, 0);
const activeCount = activeSukuks.length;

    const yearlyProfit = activeSukuks.reduce((sum, i) => {
      if (getEffectiveStatus(i) === 'Matured') return sum;
      return sum + (i.principalAmount * (i.rentRate / 100) * (1 - i.tds / 100));
    }, 0);

    const activeWeightedROI = activeSukuks.reduce((sum, i) => {
      if (getEffectiveStatus(i) === 'Matured') return sum;
      return sum + (i.rentRate * (1 - i.tds / 100) * i.principalAmount);
    }, 0);
    const avgROIAfterTax = active > 0 ? (activeWeightedROI / active) : 0;

    const returned = investments.reduce((sum, inv) => {
      const periodReturn = (inv.installments || []).reduce((iSum, inst) => {
        const payDate = inst.actualDate || inst.date;
        if (inst.isPaid && payDate >= startStr && payDate <= endStr) {
          return iSum + (inst.actualAmount || inst.amount);
        }
        return iSum;
      }, 0);
      return sum + periodReturn;
    }, 0);

    const profit = investments.reduce((sum, inv) => {
      const periodProfit = (inv.installments || []).reduce((iSum, inst) => {
        const payDate = inst.actualDate || inst.date;
        if (inst.isPaid && payDate >= startStr && payDate <= endStr) {
          return iSum + (inst.actualAmount || inst.amount);
        }
        return iSum;
      }, 0);
      return sum + periodProfit;
    }, 0);

    const totalWeightedROI = investments.reduce((sum, i) => {
      return sum + (i.rentRate * (1 - i.tds / 100) * i.principalAmount);
    }, 0);
    const totalPrincipal = investments.reduce((sum, i) => sum + i.principalAmount, 0);
    const totalROIAfterTax = totalPrincipal > 0 ? (totalWeightedROI / totalPrincipal) : 0;

    const currentYearInvested = investments
      .filter(i => new Date(i.issueDate).getFullYear() === now.getFullYear())
      .reduce((sum, i) => sum + i.principalAmount, 0);

    return { 
      invested, 
      active, 
      activeCount, 
      yearlyProfit, 
      avgROIAfterTax, 
      returned, 
      profit, 
      totalROIAfterTax, 
      currentYearInvested, 
      startStr, 
      endStr 
    };
  }, [investments, historyRange, historyCustomDates]);

  // [CHANGE 3] Notify parent whenever active investment changes
  React.useEffect(() => {
    onActiveInvestmentChange?.(historyStats.active);
  }, [historyStats.active, onActiveInvestmentChange]);

  // Update Title with Range
  React.useEffect(() => {
    if (onTitleChange) {
      if (historyRange === 'all') {
        onTitleChange('Sukuk Funds');
      } else {
        const start = new Date(historyStats.startStr);
        const end = new Date(historyStats.endStr);
        const startStrFormatted = start.toLocaleString('default', { month: 'short', year: 'numeric' });
        const endStrFormatted = end.toLocaleString('default', { month: 'short', year: 'numeric' });
        onTitleChange(
          <span className="flex items-center gap-2">
            SUKUK FUNDS <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">/ {startStrFormatted} - {endStrFormatted}</span>
          </span>
        );
      }
    }
  }, [historyStats, onTitleChange, historyRange]);

  const handleToggleInstallment = (invId: string, index: number) => {
    const inv = investments.find(i => i.id === invId);
    if (!inv) return;

    const newInstallments = [...inv.installments];
    const isPaid = !newInstallments[index].isPaid;
    
    newInstallments[index] = { 
      ...newInstallments[index], 
      isPaid,
      actualDate: isPaid ? (newInstallments[index].actualDate || newInstallments[index].date) : undefined,
      actualAmount: isPaid ? (newInstallments[index].actualAmount || newInstallments[index].amount) : undefined,
      isAutoMarked: isPaid,
      isManuallyEdited: false
    };
    
    const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount || inst.amount) : sum, 0);
    const newStatus = getEffectiveStatus({ ...inv, installments: newInstallments });
    
    onUpdate(invId, { 
      installments: newInstallments,
      totalRepaid,
      status: newStatus
    });
  };

  const handleSaveCustomRepayment = (invId: string, index: number) => {
    const inv = investments.find(i => i.id === invId);
    if (!inv) return;

    const actualAmount = parseFloat(customRepaymentData.amount) || 0;
    const actualDate = customRepaymentData.date;

    const newInstallments = [...inv.installments];
    newInstallments[index] = { 
      ...newInstallments[index], 
      isPaid: true,
      actualDate,
      actualAmount,
      isAutoMarked: false,
      isManuallyEdited: true
    };
    
    const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount || inst.amount) : sum, 0);
    const newStatus = getEffectiveStatus({ ...inv, installments: newInstallments });
    
    onUpdate(invId, { 
      installments: newInstallments,
      totalRepaid,
      status: newStatus
    });
    setCustomRepaymentIdx(null);
  };

  const handleEdit = (inv: Sukuk) => {
    setEditingId(inv.id);
    setFormData({
      name: inv.name,
      instrumentNo: inv.instrumentNo,
      issuer: inv.issuer,
      rentRate: inv.rentRate.toString(),
      tds: inv.tds.toString(),
      frequency: inv.frequency,
      principalAmount: inv.principalAmount.toString(),
      issueDate: inv.issueDate,
      durationYears: inv.durationYears.toString(),
      status: inv.status,
      closingDate: (inv as any).closingDate || '',
      withdrawBalance: (inv as any).withdrawBalance !== undefined
        ? (inv as any).withdrawBalance.toString()
        : inv.principalAmount.toString(),
    });
    setIsModalOpen(true);
    setActiveMenuId(null);
  };

  const handleDelete = (id: string) => {
    const inv = investments.find(i => i.id === id);
    if (!inv) return;

    setConfirmState({
      isOpen: true,
      title: 'Confirm Delete',
      message: `Are you sure you want to delete the Sukuk investment "${inv.name}"?`,
      onConfirm: () => {
        closeConfirm();
        onDelete(id);
        setActiveMenuId(null);
        showNotification('success', 'Investment deleted successfully');
      },
      onCancel: closeConfirm,
    });
  };

const handleBatchDelete = () => {
    if (selectedIds.length === 0) return;

    setConfirmState({
      isOpen: true,
      title: 'Confirm Batch Delete',
      message: `Are you sure you want to delete ${selectedIds.length} selected Sukuk investments?`,
      onConfirm: () => {
        closeConfirm();
        onBatchDelete(selectedIds);
        setSelectedIds([]);
        showNotification('success', `Successfully deleted ${selectedIds.length} investments.`);
      },
      onCancel: closeConfirm,
    });
  };

  const toggleSelectAll = () => {
  if (selectedIds.length === filtered.length && filtered.length > 0) {
    setSelectedIds([]);
  } else {
    setSelectedIds(filtered.map(i => i.id));
  }
};

const toggleSelect = (id: string) => {
  if (selectedIds.includes(id)) {
    setSelectedIds(selectedIds.filter(i => i !== id));
  } else {
    setSelectedIds([...selectedIds, id]);
  }
};

  // Compute the default withdraw balance when principalAmount changes and status is Matured
  const handleFormChange = (field: string, value: string) => {
    setFormData(prev => {
      const updated = { ...prev, [field]: value };
      if (field === 'principalAmount' && updated.status === 'Matured') {
        if (prev.withdrawBalance === '' || prev.withdrawBalance === prev.principalAmount) {
          updated.withdrawBalance = value;
        }
      }
      if (field === 'status' && value === 'Matured' && updated.withdrawBalance === '') {
        updated.withdrawBalance = updated.principalAmount;
      }
      return updated;
    });
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setFormError(null);

    const principalAmount = parseFloat(formData.principalAmount) || 0;
    const rentRate = parseFloat(formData.rentRate) || 0;
    const tds = parseFloat(formData.tds) || 0;
    const durationYears = parseFloat(formData.durationYears) || 0;
    
    if (principalAmount <= 0) {
      setFormError("Please enter a valid principal amount");
      return;
    }

    if (formData.status === 'Matured' && !formData.closingDate) {
      setFormError("Please enter a closing date for the matured Sukuk");
      return;
    }

    // Generate installments based on frequency
    const installments = [];
    const issueDate = new Date(formData.issueDate);
    let intervalMonths = 12;
    if (formData.frequency === 'Semi-annual') intervalMonths = 6;
    else if (formData.frequency === 'Quarterly') intervalMonths = 3;

    const totalInstallments = Math.floor((durationYears * 12) / intervalMonths);
    const rentPerInstallment = (principalAmount * (rentRate / 100) * (intervalMonths / 12)) * (1 - tds / 100);

    for (let i = 1; i <= totalInstallments; i++) {
      const instDate = new Date(issueDate);
      instDate.setMonth(issueDate.getMonth() + i * intervalMonths);
      installments.push({
        date: instDate.toISOString().split('T')[0],
        amount: rentPerInstallment,
        isPaid: false,
        installmentNo: i
      });
    }

    // AFTER
// When editing, preserve existing installments and totalRepaid so paid
// history and Total Profit are never wiped. Only regenerate installments
// for brand-new entries.
const existingInv = editingId ? investments.find(i => i.id === editingId) : null;

const finalInstallments = existingInv
  ? existingInv.installments   // keep all paid/unpaid state intact
  : installments;              // fresh generation only for new entries

const finalTotalRepaid = existingInv
  ? existingInv.totalRepaid    // never recalculate from scratch on edit
  : 0;

const payload: any = {
  name: formData.name,
  investmentDate: formData.issueDate,
  amount: principalAmount,
  currency: 'BDT' as const,
  instrumentNo: formData.instrumentNo,
  issuer: formData.issuer,
  rentRate,
  tds,
  frequency: formData.frequency,
  principalAmount,
  issueDate: formData.issueDate,
  durationYears,
  status: formData.status,
  installments: finalInstallments,
  totalRepaid: finalTotalRepaid,
};

    if (formData.status === 'Matured') {
      payload.closingDate = formData.closingDate;
      const parsedWithdraw = parseFloat(formData.withdrawBalance);
      payload.withdrawBalance = isNaN(parsedWithdraw) ? principalAmount : parsedWithdraw;
    }

    if (editingId) {
      onUpdate(editingId, payload);
    } else {
      onAdd(payload);
      if (onNavigateToModule) {
        onNavigateToModule('income-expense', {
          date: payload.issueDate,
          amount: payload.principalAmount,
          type: 'Transfer',
          targetModule: 'sukuk',
          description: `Sukuk: ${payload.name || payload.issuer}`.trim()
        });
      }
    }

    setIsModalOpen(false);
    setEditingId(null);
    setFormData({
      name: '',
      instrumentNo: '',
      issuer: 'Bangladesh Bank',
      rentRate: '',
      tds: '10',
      frequency: 'Semi-annual',
      principalAmount: '',
      issueDate: new Date().toISOString().split('T')[0],
      durationYears: '',
      status: 'Active',
      closingDate: '',
      withdrawBalance: '',
    });
  };

  return (
    <div className="space-y-8">


{/* Date Range Selector */}
<div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 relative">
  
  <div className="flex flex-col sm:flex-row sm:items-center gap-3 w-full sm:w-auto">
    
    {/* LINE 1 (Mobile): Range + Settings */}
    <div className="flex items-center justify-between w-full sm:w-auto gap-2">

      {/* Range Selection */}
      <div className="relative flex-1 sm:flex-none">

        {/* Mobile Dropdown */}
        <div className="block sm:hidden">
          <button 
            onClick={() => setIsHistoryMenuOpen(!isHistoryMenuOpen)}
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
                isHistoryMenuOpen ? "rotate-180 text-teal-400" : ""
              )} 
            />
          </button>

          {isHistoryMenuOpen && (
            <>
              <div className="fixed inset-0 z-40" onClick={() => setIsHistoryMenuOpen(false)} />
              <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95 backdrop-blur-md">
                {['all','last12m','fiscal','custom'].map((id) => (
                  <button
                    key={id}
                    onClick={() => {
                      setHistoryRange(id as any);
                      setIsHistoryMenuOpen(false);
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
              {id === 'all' ? 'Overall' : 
               id === 'last12m' ? 'Last 12M' : 
               id === 'fiscal' ? 'Fiscal' : 'Custom'}
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
                  onClick={() => { exportData(); setIsSettingsMenuOpen(false); }}
                  className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                >
                  <Download size={14} className="text-teal-400" />
                  EXPORT
                </button>
                <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
                  <Upload size={14} className="text-teal-400" />
                  IMPORT
                  <input type="file" className="hidden" accept=".xlsx,.xls" onChange={importData} />
                </label>
                <button 
                  onClick={() => { alert('Template not implemented'); setIsSettingsMenuOpen(false); }}
                  className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                >
                  <FileSpreadsheet size={14} className="text-teal-400" />
                  TEMPLATE
                </button>
                <div className="h-px bg-slate-800/60 my-2" />
                <button 
                  onClick={() => { setIsModalOpen(true); setIsSettingsMenuOpen(false); }}
                  className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-teal-400 hover:bg-teal-400/10 rounded-lg transition-colors uppercase"
                >
                  <Plus size={14} />
                  Add Sukuk
                </button>
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
              onClick={() => { exportData(); setIsSettingsMenuOpen(false); }}
              className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
            >
              <Download size={14} className="text-teal-400" />
              EXPORT
            </button>
            <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
              <Upload size={14} className="text-teal-400" />
              IMPORT
              <input type="file" className="hidden" accept=".xlsx,.xls" onChange={importData} />
            </label>
            <button 
              onClick={() => { alert('Template not implemented'); setIsSettingsMenuOpen(false); }}
              className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
            >
              <FileSpreadsheet size={14} className="text-teal-400" />
              TEMPLATE
            </button>

            <div className="h-px bg-slate-800/60 my-2" />

            <button 
              onClick={() => { setIsModalOpen(true); setIsSettingsMenuOpen(false); }}
              className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-teal-400 hover:bg-teal-400/10 rounded-lg transition-colors uppercase"
            >
              <Plus size={14} />
              Add Sukuk
            </button>

        </div>
      </>
    )}
  </div>
</div>

      {notification && (
        <div className={cn(
          "flex items-center gap-3 p-4 rounded-xl border animate-in fade-in slide-in-from-top-4 duration-300",
          notification.type === 'success' ? "bg-teal-400/10 border-teal-400/20 text-teal-400" : "bg-rose-500/10 border-rose-500/20 text-rose-500"
        )}>
          {notification.type === 'success' ? <CheckCircle2 size={20} /> : <XCircle size={20} />}
          <p className="text-body font-bold uppercase">{notification.message}</p>
        </div>
      )}
      {/* Top Summary Panel */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 items-stretch">
        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-teal-400/10 flex items-center justify-center mb-4 text-teal-400 group-hover:scale-110 transition-transform">
              <DollarSign size={20} />
            </div>
            <div className="mb-2">
              <p className="text-body-sm font-bold text-white uppercase tracking-wider">Current Holding</p>
              <p className="text-label font-bold text-slate-500 uppercase">Active Investment</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(historyStats.active)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Active Funds</p>
            <p className="text-body font-bold text-teal-400 tabular-nums">
              {String(historyStats.activeCount).padStart(2, '0')}
            </p>
          </div>
        </Card>

        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-emerald-400/50 hover:shadow-[0_0_20px_rgba(52,211,153,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-emerald-400/10 flex items-center justify-center mb-4 text-emerald-400 group-hover:scale-110 transition-transform">
              <TrendingUp size={20} />
            </div>
            <div className="mb-2">
              <p className="text-body-sm font-bold text-white uppercase tracking-wider">Yearly Profit</p>
              <p className="text-label font-bold text-slate-500 uppercase">Profit After Tax</p>
            </div>
            <h3 className={cn("text-heading font-bold text-white tabular-nums", historyStats.yearlyProfit >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {formatBDT(historyStats.yearlyProfit)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">ROI After TAX</p>
            <p className={cn("text-body font-bold tabular-nums", historyStats.avgROIAfterTax >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {historyStats.avgROIAfterTax.toFixed(1)}%
            </p>
          </div>
        </Card>

        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-purple-400/50 hover:shadow-[0_0_20px_rgba(167,139,250,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-purple-400/10 flex items-center justify-center mb-4 text-purple-400 group-hover:scale-110 transition-transform">
              <Briefcase size={20} />
            </div>
            <div className="mb-2">
              <p className="text-body-sm font-bold text-white uppercase tracking-wider">Total Invested</p>
              <p className="text-label font-bold text-slate-500 uppercase">Total Investment</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(historyStats.invested)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">This Year</p>
            <p className={cn("text-body font-bold text-purple-400 tabular-nums", historyStats.currentYearInvested >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {formatBDT(historyStats.currentYearInvested)}
            </p>
          </div>
        </Card>

        <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-emerald-400/50 hover:shadow-[0_0_20px_rgba(52,211,153,0.1)] group">
          <div className="flex-1">
            <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-emerald-400/10 flex items-center justify-center mb-4 text-emerald-400 group-hover:scale-110 transition-transform">
              <TrendingUp size={20} />
            </div>
            <div className="mb-2">
              <p className="text-body-sm font-bold text-white uppercase tracking-wider">Total Profit</p>
              <p className="text-label font-bold text-slate-500 uppercase">Total Profit Received</p>
            </div>
            <h3 className={cn("text-heading font-bold mb-4 tracking-tight font-display tabular-nums", historyStats.profit >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {formatBDT(historyStats.profit)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Portfolio Yield</p>
            <p className={cn("text-body font-bold tabular-nums", historyStats.totalROIAfterTax >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {historyStats.totalROIAfterTax.toFixed(2)}%
            </p>
          </div>
        </Card>
      </div>

      {/* Bottom Analytics Panel */}
      <div className="grid grid-cols-1 gap-4">
        <Card className="bg-slate-900 border-slate-800">
          <div className="flex items-center gap-2 mb-6">
            <Calendar className="text-teal-400 w-5 h-5" />
            <h3 className="text-subheading font-bold text-white uppercase">Upcoming Installments</h3>
          </div>

          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            {upcomingInstallments.length > 0 ? upcomingInstallments.map((item, i) => (
              <div 
                key={i} 
                className={cn(
                  "p-1.5 sm:p-2 rounded-xl border transition-all flex flex-col group hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)]",
                  i === 0 
                    ? "bg-teal-400/5 border-teal-400/30 ring-1 ring-teal-400/20" 
                    : "bg-slate-950 border-slate-800"
                )}
              >
                <p className="text-[8px] sm:text-label font-bold mb-1 sm:mb-1.5 tabular-nums text-slate-300 flex items-center justify-between whitespace-nowrap overflow-hidden">
                  <span>{formatDate(item.date)}</span>
                  <span className="mx-1 text-slate-800">|</span>
                  <span className={cn("truncate", item.isDelayed ? "text-rose-500" : "text-blue-400")}>
                    {item.isDelayed ? 'Delayed' : 'On Schedule'}
                  </span>
                </p>
                
                <div className="pt-1 sm:pt-1.5 border-t border-slate-800 mb-1 sm:mb-1.5">
                  <p className="text-body sm:text-subheading font-bold text-white tracking-tight tabular-nums truncate">{formatBDT(item.amount)}</p>
                </div>
                
                <div className="mt-auto">
                  <p className="text-[7px] sm:text-label font-bold truncate">
                    <span className={i === 0 ? "text-emerald-400" : "text-slate-300"}>{item.sukukName}</span>
                    <span className="mx-1 text-slate-800">|</span>
                    <span className="text-slate-300">Installment #{item.installmentNo}</span>
                  </p>
                </div>
              </div>
            )) : (
              <div className="col-span-full py-6 text-center bg-slate-950/30 border border-dashed border-slate-800 rounded-xl">
                <p className="text-xs text-slate-600 italic">No upcoming installments scheduled</p>
              </div>
            )}
          </div>
        </Card>
      </div>

      {/* Portfolio Header & Filters */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-6">
        <div className="flex items-center gap-2 sm:gap-3">
          <Briefcase className="text-teal-400 w-5 h-5 lg:w-6 lg:h-6" />
          <h2 className="text-body-sm sm:text-subheading lg:text-heading font-bold text-white font-display uppercase whitespace-nowrap">Sukuk Portfolio</h2>
        </div>
        <div className="flex-1 sm:flex-none flex justify-end">
          <div className="relative flex-1 sm:flex-none">
            <div 
              className="flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-lg p-1 px-2 h-9 hover:border-slate-700 focus-within:border-teal-400/50 transition-colors group cursor-text"
              onClick={() => {
                const input = document.getElementById('sukuk-search-input');
                if (input) input.focus();
              }}
            >
              <Search size={14} className="text-slate-500 group-focus-within:text-teal-400 transition-colors shrink-0" />
              <input 
                id="sukuk-search-input"
                type="text"
                placeholder="SEARCH..."
                className="flex-1 bg-transparent border-none focus:ring-0 text-[10px] font-bold text-white placeholder:text-slate-600 uppercase outline-none min-w-0"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
            </div>
          </div>
        </div>
      </div>

      <div className="space-y-3">
        <div className="flex items-center justify-between bg-slate-900/50 border border-slate-800 rounded-xl px-3 py-1.5 h-10">
          <div className="flex items-center gap-2">
            <Checkbox 
              checked={selectedIds.length === filtered.length && filtered.length > 0}
              onChange={toggleSelectAll}
            />
            <span className="text-[10px] font-bold text-slate-300 uppercase tracking-wider">Select All</span>
          </div>
          <span className="text-[10px] font-bold text-slate-500 uppercase tabular-nums">
            {selectedIds.length} of {filtered.length} Selected
          </span>
        </div>

        {selectedIds.length > 0 && (
          <div className="flex items-center justify-between bg-rose-500/10 border border-rose-500/20 rounded-xl px-4 py-2 animate-in fade-in slide-in-from-top-2">
            <span className="text-label font-bold text-rose-500 uppercase">
              {selectedIds.length} Investments Selected
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

        {filtered.map((inv) => {
          const status = getEffectiveStatus(inv);
          const isMatured = status === 'Matured'; // [CHANGE 1 & 2] helper flag

          // [CHANGE 2] For matured sukuks with a withdrawal, show net holding
          const withdrawBalance = (inv as any).withdrawBalance as number | undefined;
          const displayPrincipal = isMatured && withdrawBalance !== undefined
            ? inv.principalAmount - withdrawBalance
            : inv.principalAmount;

          return (
            <div 
              key={inv.id} 
              className={cn(
                "bg-slate-900 border rounded-xl overflow-hidden transition-all group",
                selectedIds.includes(inv.id) ? "border-teal-400 ring-1 ring-teal-400/20 shadow-[0_0_15px_rgba(45,212,191,0.1)]" : "border-slate-800 hover:border-teal-400/50"
              )}
            >
              <div className="bg-slate-950 border-b border-slate-800 px-4 py-2 flex justify-between items-center relative">
                <div className="flex items-center gap-3">
                  <Checkbox 
                    checked={selectedIds.includes(inv.id)}
                    onChange={() => toggleSelect(inv.id)}
                  />
                  <div className={cn("w-2 h-2 rounded-full animate-pulse", isMatured ? 'bg-slate-500' : 'bg-teal-400')} />
                  {/* [CHANGE 1] Title bar: sukuk name + "(Matured)" label in red if matured */}
                  <span className="text-[9px] sm:text-body-sm font-bold uppercase tracking-tight tabular-nums truncate flex items-center gap-1.5">
                    <span className="text-slate-300">{inv.issuer}</span>
                    <span className="text-slate-800 mx-1 sm:mx-2">|</span>
                    <span className="text-teal-400">{inv.name}</span>
                    {isMatured && (
                      <span className="text-rose-500 font-bold text-[9px] sm:text-body-sm whitespace-nowrap">
                        (Matured)
                      </span>
                    )}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => setEditingActualId(editingActualId === inv.id ? null : inv.id)}
                    className={cn(
                      "p-1.5 rounded-lg transition-colors",
                      editingActualId === inv.id ? "bg-teal-400 text-white" : "text-slate-500 hover:text-white hover:bg-slate-800"
                    )}
                  >
                    <Clock size={14} />
                  </button>
                  
                  <div className="relative">
                    <button 
                      onClick={() => setActiveMenuId(activeMenuId === inv.id ? null : inv.id)}
                      className="p-1.5 text-slate-500 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
                    >
                      <MoreVertical size={14} />
                    </button>

                    {activeMenuId === inv.id && (
                      <div className="absolute right-0 mt-2 w-32 bg-slate-900 border border-slate-800 rounded-lg shadow-xl z-50 overflow-hidden animate-in fade-in zoom-in-95">
                        <button 
                          onClick={() => handleEdit(inv)}
                          className="w-full flex items-center gap-2 px-3 py-2 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white transition-colors"
                        >
                          <Edit2 size={12} className="text-teal-400" />
                          EDIT
                        </button>
                        <button 
                          onClick={() => handleDelete(inv.id)}
                          className="w-full flex items-center gap-2 px-3 py-2 text-label font-bold text-rose-500 hover:bg-rose-500/10 transition-colors"
                        >
                          <Plus size={12} className="rotate-45" />
                          DELETE
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="p-3 flex flex-row overflow-x-auto lg:grid lg:grid-cols-5 gap-3 sm:gap-4 no-scrollbar">
                {/* [CHANGE 2] Principal column: show net amount (principal - withdraw) for matured */}
                <div className="space-y-1 min-w-[120px] shrink-0">
                  <p className="text-label text-slate-300 uppercase mb-1">Principal</p>
                  <p className="text-body-sm font-bold text-white tabular-nums">{formatBDT(displayPrincipal)}</p>
                  <p className="text-label text-slate-500 font-medium tabular-nums">{formatDate(inv.issueDate)}</p>
                </div>

                <div className="space-y-1 min-w-[100px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">Duration</p>
                  <p className="text-body-sm font-bold text-white tabular-nums">{inv.durationYears} Years</p>
                  <p className="text-label text-slate-500 font-medium tabular-nums whitespace-nowrap">
                    {(() => {
                      const issueDate = new Date(inv.issueDate);
                      const maturityDate = new Date(issueDate);
                      maturityDate.setFullYear(issueDate.getFullYear() + inv.durationYears);
                      return formatDate(maturityDate.toISOString().split('T')[0]);
                    })()}
                  </p>
                </div>

                <div className="space-y-1 min-w-[120px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">Frequency</p>
                  <p className="text-body-sm font-bold text-white uppercase tabular-nums">{inv.frequency}</p>
                  <p className="text-label text-slate-500 font-medium tabular-nums whitespace-nowrap">
                    {inv.installments.slice(0, 2).map(inst => {
                      const d = new Date(inst.date);
                      return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' });
                    }).join(' | ')}
                  </p>
                </div>

                <div className="space-y-1 min-w-[120px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">Rental Rate</p>
                  <p className="text-body-sm font-bold text-white tabular-nums">{inv.rentRate.toFixed(2)}%</p>
                  <p className="text-label text-slate-500 font-medium tabular-nums">
                    {formatBDT(inv.principalAmount * (inv.rentRate / 100) * (1 - inv.tds / 100))}/Year
                  </p>
                </div>

                <div className="space-y-1 min-w-[120px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">PROFIT RECEIVED</p>
                  <p className={cn("text-body-sm font-bold tabular-nums", inv.totalRepaid >= 0 ? "text-emerald-400" : "text-rose-500")}>{formatBDT(inv.totalRepaid)}</p>
                  <p className="text-label text-slate-500 font-medium tabular-nums">
                    {inv.installments.filter(i => i.isPaid).length}/{inv.installments.length} Installments
                  </p>
                </div>
              </div>

              {editingActualId === inv.id && (
                <div className="w-full px-4 pb-4 animate-in fade-in slide-in-from-top-2">
                  <div className="pt-4 border-t border-slate-800">
                    <div className="flex flex-wrap gap-3">
                      {inv.installments.map((inst, idx) => (
                        <div key={idx} className="flex items-center gap-2">
                          {customRepaymentIdx === idx ? (
                            <div className="flex items-center gap-2 bg-slate-950 p-2 rounded-lg border border-teal-400/50 animate-in zoom-in-95">
                              <input 
                                type="date" 
                                className="bg-transparent text-label text-white border-none focus:ring-0 p-0 w-24 tabular-nums"
                                value={customRepaymentData.date}
                                onChange={(e) => setCustomRepaymentData({ ...customRepaymentData, date: e.target.value })}
                              />
                              <div className="w-px h-4 bg-slate-800" />
                              <input 
                                type="number" 
                                className="bg-transparent text-label text-emerald-400 border-none focus:ring-0 p-0 w-16 font-bold tabular-nums"
                                value={customRepaymentData.amount}
                                onChange={(e) => setCustomRepaymentData({ ...customRepaymentData, amount: e.target.value })}
                              />
                              <button 
                                onClick={() => handleSaveCustomRepayment(inv.id, idx)}
                                className="p-1 text-emerald-400 hover:bg-emerald-400/10 rounded"
                              >
                                <CheckCircle2 size={14} />
                              </button>
                              <button 
                                onClick={() => setCustomRepaymentIdx(null)}
                                className="p-1 text-slate-500 hover:bg-slate-800 rounded"
                              >
                                <X size={14} />
                              </button>
                            </div>
                          ) : (
                            <div className="relative group/inst">
                              <div
                                role="button"
                                tabIndex={0}
                                onClick={() => handleToggleInstallment(inv.id, idx)}
                                className={cn(
                                  "flex items-center gap-2 px-3 py-2 rounded-lg border transition-all cursor-pointer",
                                  inst.isPaid 
                                    ? "bg-emerald-400/10 border-emerald-400/30 text-emerald-400"
                                    : "bg-slate-950 border-slate-800 text-slate-500 hover:border-slate-600"
                                )}
                              >
                                <CheckCircle2 size={14} />
                                <div className="text-left">
                                  <p className="text-label font-bold tabular-nums">
                                    {formatDate(inst.actualDate || inst.date)}
                                  </p>
                                  <p className="text-label opacity-70 tabular-nums">
                                    {formatBDT(inst.actualAmount !== undefined ? inst.actualAmount : inst.amount)}
                                  </p>
                                </div>
                              </div>
                              <button 
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setCustomRepaymentIdx(idx);
                                  setCustomRepaymentData({ 
                                    date: inst.actualDate || inst.date, 
                                    amount: (inst.actualAmount !== undefined ? inst.actualAmount : inst.amount).toString() 
                                  });
                                }}
                                className="absolute -top-2 -right-2 p-1 bg-slate-800 border border-slate-700 rounded-full text-slate-400 hover:text-white opacity-0 group-hover/inst:opacity-100 transition-opacity"
                              >
                                <Edit2 size={10} />
                              </button>
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>

      <Modal 
        isOpen={isModalOpen} 
        onClose={() => {
          setIsModalOpen(false);
          setEditingId(null);
        }} 
        title={editingId ? "Edit Sukuk" : "Add New Sukuk"}
      >
        <form onSubmit={handleSubmit} className="space-y-4">
          {formError && (
            <div className="p-3 bg-rose-500/10 border border-rose-500/20 rounded-lg flex items-center gap-2 text-rose-500 text-label font-bold animate-in fade-in slide-in-from-top-1">
              <AlertCircle size={14} />
              <span className="uppercase">{formError}</span>
            </div>
          )}
          <div className="grid grid-cols-2 gap-3">
            <Input 
              label="Sukuk Name" 
              placeholder="e.g. Beximco Sukuk"
              value={formData.name}
              onChange={(e) => setFormData({ ...formData, name: e.target.value })}
              required
            />
            <Input 
              label="Instrument No." 
              placeholder="e.g. SUKUK-001"
              value={formData.instrumentNo}
              onChange={(e) => setFormData({ ...formData, instrumentNo: e.target.value })}
              required
            />
          </div>

          <div className="grid grid-cols-3 gap-3">
            <Input 
              label="Issuer" 
              placeholder="e.g. Bangladesh Bank"
              value={formData.issuer}
              onChange={(e) => setFormData({ ...formData, issuer: e.target.value })}
              required
            />
            <Input 
              label="Rent Rate (%)" 
              type="number" 
              placeholder="0.00"
              value={formData.rentRate}
              onChange={(e) => setFormData({ ...formData, rentRate: e.target.value })}
              required
            />
            <Input 
              label="TDS (%)" 
              type="number" 
              placeholder="10"
              value={formData.tds}
              onChange={(e) => setFormData({ ...formData, tds: e.target.value })}
              required
            />
          </div>

          <div className="grid grid-cols-2 gap-3">
            <Select 
              label="Payment Frequency" 
              value={formData.frequency}
              options={[
                { label: 'Quarterly', value: 'Quarterly' },
                { label: 'Semi-annual', value: 'Semi-annual' },
                { label: 'Annual', value: 'Annual' }
              ]}
              onChange={(val) => setFormData({ ...formData, frequency: val as any })}
              required
            />
            <Input 
              label="Principal Amount" 
              type="number" 
              placeholder="0.00"
              value={formData.principalAmount}
              onChange={(e) => handleFormChange('principalAmount', e.target.value)}
              required
            />
          </div>

          <div className="grid grid-cols-2 gap-3">
            <Input 
              label="Issue Date" 
              type="date"
              value={formData.issueDate}
              onChange={(e) => setFormData({ ...formData, issueDate: e.target.value })}
              required
            />
            <Input 
              label="Duration (Years)" 
              type="number" 
              placeholder="0"
              value={formData.durationYears}
              onChange={(e) => setFormData({ ...formData, durationYears: e.target.value })}
              required
            />
          </div>

          <Select 
            label="Status" 
            value={formData.status}
            options={[
              { label: 'Active', value: 'Active' },
              { label: 'Matured', value: 'Matured' },
              { label: 'Delayed', value: 'Delayed' }
            ]}
            onChange={(val) => handleFormChange('status', val)}
            required
          />

          {/* Matured-only fields */}
          {formData.status === 'Matured' && (
            <div className="grid grid-cols-2 gap-3 animate-in fade-in slide-in-from-top-2 duration-200">
              <div className="relative">
                <Input 
                  label="Closing Date" 
                  type="date"
                  value={formData.closingDate}
                  onChange={(e) => setFormData({ ...formData, closingDate: e.target.value })}
                  required
                />
              </div>
              <div className="relative">
                <Input 
                  label="Withdraw Balance" 
                  type="number" 
                  placeholder={formData.principalAmount || '0.00'}
                  value={formData.withdrawBalance}
                  onChange={(e) => setFormData({ ...formData, withdrawBalance: e.target.value })}
                />
                {formData.withdrawBalance !== '' && formData.withdrawBalance !== formData.principalAmount && (
                  <button
                    type="button"
                    onClick={() => setFormData({ ...formData, withdrawBalance: formData.principalAmount })}
                    className="absolute right-2 top-7 text-[9px] font-bold text-teal-400 hover:text-teal-300 uppercase transition-colors"
                    title="Reset to principal amount"
                  >
                    Reset
                  </button>
                )}
              </div>
            </div>
          )}

          <div className="pt-2 flex gap-3">
            <Button type="button" variant="secondary" className="flex-1 py-2" onClick={() => setIsModalOpen(false)}>
              Cancel
            </Button>
            <Button type="submit" className="flex-1 py-2">
              {editingId ? "Update Sukuk" : "Save Sukuk"}
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
  );
};
