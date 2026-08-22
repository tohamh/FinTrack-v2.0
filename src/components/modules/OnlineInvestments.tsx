/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { Card, Button, Modal, Input, Checkbox, Select } from '../ui/BaseComponents';
import { ConfirmDialog } from '../ui/ConfirmDialog';
import { OnlineInvestment, OnlineInvestmentStatus } from '../../types';
import { formatBDT, formatDate, cn, toDateStr, getTodayStr, getFirstOfMonth, getLastOfMonth } from '../../utils/formatters';
import * as XLSX from 'xlsx';
import { Plus, MoreVertical, Calendar, TrendingUp, Search, Globe, Building2, Clock, ArrowRight, CheckCircle2, AlertCircle, Edit2, Save, X, Download, Upload, FileSpreadsheet, Briefcase, Settings, Sparkles, XCircle, DollarSign, Coins, ChevronDown, Wallet, ChevronLeft, ChevronRight } from 'lucide-react';

interface OnlineInvestmentsProps {
  investments: OnlineInvestment[];
  onAdd: (investment: Omit<OnlineInvestment, 'id'>) => void;
  onUpdate: (id: string, updates: Partial<OnlineInvestment>) => void;
  onDelete: (id: string) => void;
  onBatchDelete: (ids: string[]) => void;
  onBatchAdd?: (investments: Omit<OnlineInvestment, 'id'>[]) => void;
  onReplaceAll?: (investments: Omit<OnlineInvestment, 'id'>[]) => void;
  triggerAdd?: boolean;
  setTriggerAdd?: (val: boolean) => void;
  inheritedData?: { date: string; amount: number } | null;
  onClearInheritedData?: () => void;
  onTitleChange?: (title: React.ReactNode) => void;
  onNavigateToModule?: (module: string, initialAddData?: { date: string; amount: number; type?: 'Transfer' | 'Income' | 'Expense' | 'Loan'; targetModule?: string; description?: string }) => void;
}

export const OnlineInvestments: React.FC<OnlineInvestmentsProps> = ({ 
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
  onNavigateToModule
}) => {
  const [selectedProfiles, setSelectedProfiles] = useState<OnlineInvestmentStatus[]>(['Active', 'Delayed']);
  const [sortBy, setSortBy] = useState<'date' | 'amount' | 'roi' | 'platform'>('date');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [companySearch, setCompanySearch] = useState('');
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingActualId, setEditingActualId] = useState<string | null>(null);
  const [isEditingSchedule, setIsEditingSchedule] = useState(false);
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
  const [customPlatforms, setCustomPlatforms] = useState<string[]>(() => {
    const saved = localStorage.getItem('fintrack_custom_platforms');
    return saved ? JSON.parse(saved) : [];
  });
  const [newPlatformName, setNewPlatformName] = useState('');
  const [isAddingPlatform, setIsAddingPlatform] = useState(false);
  const [customRepaymentIdx, setCustomRepaymentIdx] = useState<number | null>(null);
  const [customRepaymentData, setCustomRepaymentData] = useState({ date: '', amount: '' });
  const [activeWarning, setActiveWarning] = useState<{ id: string, text: string } | null>(null);
  const [formError, setFormError] = useState<string | null>(null);
  const [insightSearch, setInsightSearch] = useState('');
  const [selectedInsightCompany, setSelectedInsightCompany] = useState<string | null>(null);
  const [isInsightSearchOpen, setIsInsightSearchOpen] = useState(false);
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
    platform: 'Biniyog',
    companyName: '',
    projectName: '',
    amount: '',
    estimatedReturn: '',
    investmentDate: new Date().toISOString().split('T')[0],
    maturityDate: '',
    status: 'Active' as OnlineInvestmentStatus,
    actualProfit: '' as string | number,
    hasRepaymentSchedule: false,
    confirmNegativeProfit: false,
    manualInstallments: [] as { date: string; amount: string }[],
  });

  // Calculate duration in months rounded to nearest 0.5
  const calculatedDuration = useMemo(() => {
    if (!formData.investmentDate || !formData.maturityDate) return 0;
    const start = new Date(formData.investmentDate);
    const end = new Date(formData.maturityDate);
    const diffTime = end.getTime() - start.getTime();
    if (isNaN(diffTime) || diffTime <= 0) return 0;
    const months = diffTime / (1000 * 60 * 60 * 24 * 30.44);
    // Round to 1 decimal place instead of 0.5 to allow shorter durations
    return Math.max(0.1, Math.round(months * 10) / 10);
  }, [formData.investmentDate, formData.maturityDate]);

  // Handle external trigger for adding
  React.useEffect(() => {
    if (triggerAdd) {
      setEditingId(null);
      const inheritedDate = inheritedData?.date || new Date().toISOString().split('T')[0];
      const inheritedAmt = (inheritedData?.amount !== undefined && inheritedData.amount !== null && inheritedData.amount !== 0) ? String(inheritedData.amount) : '';
      setFormData({
        platform: 'Biniyog',
        companyName: '',
        projectName: '',
        amount: inheritedAmt,
        estimatedReturn: '',
        investmentDate: inheritedDate,
        maturityDate: '',
        status: 'Active' as OnlineInvestmentStatus,
        actualProfit: '' as string | number,
        hasRepaymentSchedule: false,
        confirmNegativeProfit: false,
        manualInstallments: [] as { date: string; amount: string }[],
      });
      setIsModalOpen(true);
      setTriggerAdd?.(false);
      onClearInheritedData?.();
    }
  }, [triggerAdd, setTriggerAdd, inheritedData, onClearInheritedData]);

  const getEffectiveStatus = (inv: OnlineInvestment): OnlineInvestmentStatus => {
    const today = new Date().toISOString().split('T')[0];
    
    if (inv.installments && inv.installments.length > 0) {
      const allPaid = inv.installments.every(i => i.isPaid);
      if (allPaid) return 'Completed';
      
      const hasDelayed = inv.installments.some(i => !i.isPaid && i.date < today);
      if (hasDelayed) return 'Delayed';
      
      return 'Active';
    } else {
      // Fallback for manual tracking
      const expectedTotal = inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100);
      if (inv.totalRepaid >= expectedTotal * 0.99) return 'Completed'; // 99% threshold for floating point
      
      if (inv.maturityDate < today) return 'Delayed';
      
      return 'Active';
    }
  };

  const filtered = useMemo(() => {
    let result = selectedProfiles.length === 0 
      ? [...investments] 
      : investments.filter(i => selectedProfiles.includes(getEffectiveStatus(i)));

    if (companySearch.trim()) {
      const search = companySearch.toLowerCase();
      result = result.filter(i => 
        (i.companyName || '').toLowerCase().includes(search) || 
        i.projectName.toLowerCase().includes(search)
      );
    }

    return result.sort((a, b) => {
      let comparison = 0;
      if (sortBy === 'date') comparison = new Date(b.investmentDate).getTime() - new Date(a.investmentDate).getTime();
      else if (sortBy === 'amount') comparison = b.amount - a.amount;
      else if (sortBy === 'roi') comparison = b.expectedROE - a.expectedROE;
      else if (sortBy === 'platform') comparison = a.platform.localeCompare(b.platform);
      
      return sortOrder === 'desc' ? comparison : -comparison;
    });
  }, [investments, selectedProfiles, sortBy, sortOrder, companySearch]);

  const upcomingInstallments = useMemo(() => {
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();

    const all = investments.flatMap(inv => {
      const status = getEffectiveStatus(inv);
      const expectedTotal = inv.estimatedReturn || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100));
      
      if (inv.installments && inv.installments.length > 0) {
        return inv.installments.map(inst => ({
          ...inst,
          projectName: inv.projectName,
          companyName: inv.companyName,
          platform: inv.platform,
          isDelayed: !inst.isPaid && inst.date < today.toISOString().split('T')[0]
        }));
      } else {
        // Use maturity date as a single installment if no schedule
        return [{
          date: inv.maturityDate,
          amount: expectedTotal,
          isPaid: status === 'Completed',
          projectName: inv.projectName,
          companyName: inv.companyName,
          platform: inv.platform,
          isDelayed: status === 'Delayed'
        }];
      }
    });

    const unpaid = all
      .filter(inst => !inst.isPaid)
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

    const thisMonth = unpaid.filter(inst => {
      const d = new Date(inst.date);
      return d.getMonth() === currentMonth && d.getFullYear() === currentYear;
    });

    // Show exactly 4 upcomings for a 2x2 grid
    return unpaid.slice(0, 4);
  }, [investments]);

  const insights = useMemo(() => {
    if (investments.length === 0) return null;

    const companyStats = investments.reduce((acc, inv) => {
      const company = inv.companyName || 'Unknown';
      if (!acc[company]) {
        acc[company] = { totalAmount: 0, totalRepaid: 0, totalExpected: 0, count: 0, delayedCount: 0, weightedAnnualizedROI: 0 };
      }
      const expectedTotal = inv.estimatedReturn || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100));
      const totalROI = (expectedTotal - inv.amount) / inv.amount;
      const annualizedROI = totalROI * (12 / Math.max(0.1, inv.durationMonths)) * 100;

      acc[company].totalAmount += inv.amount;
      acc[company].totalRepaid += inv.totalRepaid;
      acc[company].totalExpected += expectedTotal;
      acc[company].count += 1;
      acc[company].weightedAnnualizedROI += (annualizedROI * inv.amount);
      
      const status = getEffectiveStatus(inv);
      if (status === 'Delayed') acc[company].delayedCount += 1;
      
      return acc;
    }, {} as Record<string, { totalAmount: number; totalRepaid: number; totalExpected: number; count: number; delayedCount: number; weightedAnnualizedROI: number }>);

    const performance = Object.entries(companyStats)
      .filter(([name]) => name !== 'Unknown')
      .map(([name, s]: [string, any]) => {
        const avgAnnualizedROI = s.weightedAnnualizedROI / s.totalAmount;
        const repaymentRate = (s.totalRepaid / s.totalExpected) * 100;
        const reliability = 100 - (s.delayedCount / s.count) * 100;
        return { name, avgROI: avgAnnualizedROI, repaymentRate, reliability, score: avgAnnualizedROI * 0.4 + repaymentRate * 0.3 + reliability * 0.3 };
      });

    if (performance.length === 0) return null;
    
    // If a company is selected, find it
    if (selectedInsightCompany) {
      const selected = performance.find(p => p.name === selectedInsightCompany);
      if (selected) return { ...selected, isBest: selected.name === performance.sort((a, b) => b.score - a.score)[0].name };
    }

    const best = performance.sort((a, b) => b.score - a.score)[0];
    return { ...best, isBest: true };
  }, [investments, selectedInsightCompany]);

  const allCompanies = useMemo(() => {
    const companies = new Set(investments.map(inv => inv.companyName).filter(Boolean));
    return Array.from(companies).sort();
  }, [investments]);

  const historyStats = useMemo(() => {
    let startDate: Date;
    let endDate = new Date();
    const now = new Date();

    if (historyRange === 'all') {
      startDate = new Date(0);
      endDate = new Date(3000, 11, 31, 23, 59, 59, 999);
    } else if (historyRange === 'last12m') {
      startDate = new Date(now.getFullYear(), now.getMonth() - 11, 1);
      endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
    } else if (historyRange === 'fiscal') {
      const currentYear = now.getFullYear();
      const startYear = now.getMonth() >= 6 ? currentYear - 1 : currentYear - 2;
      startDate = new Date(startYear, 6, 1);
      endDate = new Date(startYear + 1, 5, 30, 23, 59, 59, 999);
    } else {
      startDate = historyCustomDates.start ? new Date(historyCustomDates.start) : new Date(0);
      endDate = historyCustomDates.end ? new Date(historyCustomDates.end) : new Date();
      endDate.setHours(23, 59, 59, 999);
    }

    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];

    const periodInvestments = investments.filter(inv => inv.investmentDate >= startStr && inv.investmentDate <= endStr);
    const totalInvested = periodInvestments.reduce((sum, inv) => sum + inv.amount, 0);

    const activeInvestments = investments.filter(inv => {
      const status = getEffectiveStatus(inv);
      if (inv.investmentDate > endStr) return false;
      if (status === 'Completed') {
        const lastInst = inv.installments?.[inv.installments.length - 1];
        const completionDate = lastInst ? (lastInst.actualDate || lastInst.date) : inv.maturityDate;
        if (completionDate <= endStr) return false;
      }
      return true;
    });
    const activeInvestmentAmount = activeInvestments.reduce((sum, inv) => sum + inv.amount, 0);
    const activeInvestmentCount = activeInvestments.length;

    const projectedProfit = activeInvestments.reduce((sum, i) => {
  const profit = i.estimatedReturn 
    ? (i.estimatedReturn - i.amount) 
    : (i.amount * i.expectedROE * (i.durationMonths / 12) / 100);
  return sum + profit;
}, 0);
    const avgProjectedROI = activeInvestmentAmount > 0 
      ? activeInvestments.reduce((sum, i) => sum + (i.expectedROE * i.amount), 0) / activeInvestmentAmount 
      : 0;

    const completedInPeriod = investments.filter(inv => {
      if (getEffectiveStatus(inv) !== 'Completed') return false;
      const lastInst = inv.installments?.[inv.installments.length - 1];
      const completionDate = lastInst ? (lastInst.actualDate || lastInst.date) : inv.maturityDate;
      return completionDate >= startStr && completionDate <= endStr;
    });
    const totalProfitReceived = completedInPeriod.reduce((sum, inv) => {
      const rate = inv.currency === 'USD' ? 110 : 1;
      const profit = inv.actualProfit !== undefined ? inv.actualProfit : (inv.totalRepaid - inv.amount);
      return sum + (profit * rate);
    }, 0);
    const completedPrincipal = completedInPeriod.reduce((sum, i) => sum + i.amount, 0);
    const avgRealizedROI = completedPrincipal > 0
      ? completedInPeriod.reduce((sum, i) => {
          const actualROE = i.totalRepaid > 0 ? ((i.totalRepaid - i.amount) / i.amount) * (12 / i.durationMonths) * 100 : 0;
          return sum + (actualROE * i.amount);
        }, 0) / completedPrincipal
      : 0;

    const currentYearInvested = investments
      .filter(i => new Date(i.investmentDate).getFullYear() === now.getFullYear())
      .reduce((sum, i) => sum + i.amount, 0);

    return { 
  invested: totalInvested, 
  active: activeInvestmentAmount,
  activeInvestment: activeInvestmentAmount,  // alias: for Online Invests, activeInvestment === Current Holding at all times
  activeCount: activeInvestmentCount,
  projectedProfit,
  avgProjectedROI,
  profit: totalProfitReceived, 
  avgRealizedROI,
  currentYearInvested,
  startStr, 
  endStr 
};
  }, [investments, historyRange, historyCustomDates]);

  React.useEffect(() => {
    if (onTitleChange) {
      if (historyRange === 'all') {
        onTitleChange('Online Invests');
      } else {
        const start = new Date(historyStats.startStr);
        const end = new Date(historyStats.endStr);
        const startStr = start.toLocaleString('default', { month: 'short', year: 'numeric' });
        const endStr = end.toLocaleString('default', { month: 'short', year: 'numeric' });
        onTitleChange(
          <span className="flex items-center gap-2">
            ONLINE INVESTS <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">/ {startStr} - {endStr}</span>
          </span>
        );
      }
    }
  }, [historyStats, onTitleChange, historyRange]);

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 5000);
  };

  const exportToExcel = () => {
    const rows: any[] = [];
    const today = new Date().toISOString().split('T')[0];
    
    investments.forEach((inv, index) => {
      const investmentId = `INV-${(index + 1).toString().padStart(3, '0')}`;
      const baseData = {
        'Investment ID': investmentId,
        'Platform': inv.platform,
        'Company': inv.companyName || '',
        'Project': inv.projectName,
        'Invested Amount': inv.amount,
        'Estimated Return': inv.estimatedReturn,
        'ROI (%)': inv.expectedROE,
        'Investment Date': inv.investmentDate,
        'Maturity Date': inv.maturityDate,
        'Duration (Months)': inv.durationMonths,
        'Investment Status': getEffectiveStatus(inv),
      };

      const installments = (inv.installments && inv.installments.length > 0) 
        ? inv.installments 
        : [{ date: inv.maturityDate, amount: inv.estimatedReturn, isPaid: inv.status === 'Completed' }];

      installments.forEach((inst, i) => {
        let paymentStatus = 'Pending';
        if (inst.isPaid) {
          if (inst.actualAmount !== undefined && inst.actualAmount < inst.amount) {
            paymentStatus = 'Partial';
          } else {
            paymentStatus = 'Paid';
          }
        } else {
          if (inst.date < today) {
            paymentStatus = 'Delayed';
          }
        }

        rows.push({
          ...baseData,
          'Installment No': i + 1,
          'Scheduled Date': inst.date,
          'Scheduled Amount': inst.amount,
          'Actual Paid Date': inst.isPaid ? (inst.actualDate || inst.date) : '',
          'Actual Paid Amount': inst.isPaid ? (inst.actualAmount || inst.amount) : '',
          'Payment Status': paymentStatus
        });
      });
    });

    const headers = [
      'Investment ID', 'Platform', 'Company', 'Project', 'Invested Amount', 
      'Estimated Return', 'ROI (%)', 'Investment Date', 'Maturity Date', 
      'Duration (Months)', 'Investment Status', 'Installment No', 
      'Scheduled Date', 'Scheduled Amount', 'Actual Paid Date', 
      'Actual Paid Amount', 'Payment Status'
    ];

    const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Investments");
    XLSX.writeFile(wb, `Online_Export_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('success', 'Export successful!');
  };

  const importData = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setConfirmState({
      isOpen: true,
      title: 'Confirm Replace',
      message: 'Are you sure you want to replace all existing Online Investment data with the imported data? This action cannot be undone.',
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

            // Group rows by Investment ID
            const groups: { [key: string]: any[] } = {};
            jsonData.forEach(row => {
              const id = row['Investment ID'] || row['InvestmentID'] || 'UNKNOWN';
              if (!groups[id]) groups[id] = [];
              groups[id].push(row);
            });

            const parseExcelDate = (val: any) => {
              if (!val) return '';
              if (typeof val === 'number') {
                // Excel serial date
                const date = new Date((val - 25569) * 86400 * 1000);
                return date.toISOString().split('T')[0];
              }
              const str = val.toString().trim().replace(/\//g, '-');
              if (!str) return '';
              const d = new Date(str);
              return isNaN(d.getTime()) ? '' : d.toISOString().split('T')[0];
            };

            const newInvestments: Omit<OnlineInvestment, 'id'>[] = [];
            Object.values(groups).forEach(rows => {
              const firstRow = rows[0];
              const platform = firstRow['Platform'] || firstRow['platform'];
              const companyName = firstRow['Company'] || firstRow['Company Name'] || firstRow['companyName'] || '';
              const projectName = firstRow['Project'] || firstRow['Project Name'] || firstRow['projectName'];
              const amount = parseFloat(firstRow['Invested Amount'] || firstRow['Amount'] || firstRow['amount'] || firstRow['Amount (BDT)']);
              const estimatedReturn = parseFloat(firstRow['Estimated Return'] || firstRow['estimatedReturn'] || firstRow['Est. Return'] || '0');
              const expectedROE = parseFloat(firstRow['ROI (%)'] || firstRow['Expected ROE'] || firstRow['expectedROE'] || firstRow['Expected ROE (%)']);
              
              const investmentDate = parseExcelDate(firstRow['Investment Date'] || firstRow['Start Date'] || firstRow['investmentDate']);
              const maturityDate = parseExcelDate(firstRow['Maturity Date'] || firstRow['maturityDate']);
              
              const status = firstRow['Investment Status'] || firstRow['Status'] || firstRow['status'] || 'Active';
              const notes = firstRow['Notes'] || firstRow['notes'] || '';
              
              if (!platform || !projectName || isNaN(amount) || !investmentDate || !maturityDate) return;

              let duration = parseFloat(firstRow['Duration (Months)'] || firstRow['Duration'] || firstRow['duration']);
              if (isNaN(duration)) {
                const start = new Date(investmentDate);
                const end = new Date(maturityDate);
                const diffTime = end.getTime() - start.getTime();
                const months = diffTime / (1000 * 60 * 60 * 24 * 30.44);
                duration = Math.max(0.1, Math.round(months * 10) / 10);
              }
              if (isNaN(duration)) duration = 1;

              const installments = rows.map(row => {
                const scheduledDate = parseExcelDate(row['Scheduled Date'] || row['date']) || maturityDate;
                const scheduledAmount = parseFloat(row['Scheduled Amount'] || row['amount'] || '0');
                const actualDate = parseExcelDate(row['Actual Paid Date'] || row['actualDate']);
                const actualAmount = parseFloat(row['Actual Paid Amount'] || row['actualAmount'] || '0');
                const paymentStatus = row['Payment Status'] || '';
                const isPaid = paymentStatus === 'Paid' || paymentStatus === 'Partial' || !!actualDate;

                return {
                  date: scheduledDate,
                  amount: scheduledAmount || estimatedReturn,
                  isPaid,
                  actualDate: actualDate || undefined,
                  actualAmount: actualAmount || undefined,
                  isAutoMarked: isPaid && !actualAmount,
                  isManuallyEdited: !!actualAmount
                };
              });

              const totalRepaid = installments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount || inst.amount) : sum, 0);
              const finalEstimatedReturn = estimatedReturn || (amount + (amount * expectedROE * (duration / 12) / 100));
              const profit = finalEstimatedReturn - amount;

              newInvestments.push({
                platform: platform as any,
                companyName,
                projectName,
                amount,
                estimatedReturn: finalEstimatedReturn,
                estimatedProfit: profit,
                expectedROE: expectedROE || (profit / amount) * (12 / duration) * 100,
                investmentDate,
                maturityDate,
                durationMonths: duration,
                status: status as any,
                currency: 'BDT',
                hasRepaymentSchedule: installments.length > 1 || (installments.length === 1 && installments[0].date !== maturityDate),
                isDefaultSchedule: installments.length <= 1 && installments[0]?.date === maturityDate,
                installments,
                totalRepaid,
                notes,
                name: projectName
              });
            });

            if (onReplaceAll && newInvestments.length > 0) {
              onReplaceAll(newInvestments);
              showNotification('success', `Successfully replaced all data with ${newInvestments.length} investments!`);
            } else if (newInvestments.length > 0) {
              newInvestments.forEach(inv => onAdd(inv));
              showNotification('success', `Successfully imported ${newInvestments.length} investments!`);
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
    
    // Calculate new status
    const tempInv = { ...inv, installments: newInstallments, totalRepaid };
    const newStatus = getEffectiveStatus(tempInv);
    
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
    const expectedTotal = inv.estimatedReturn || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100));

    if (index === -1) {
      // No schedule case
      onUpdate(invId, { 
        totalRepaid: actualAmount,
        actualMaturityDate: actualDate,
        status: actualAmount >= expectedTotal * 0.99 ? 'Completed' : 'Active',
        installments: [{
          date: inv.maturityDate,
          amount: expectedTotal,
          isPaid: actualAmount >= expectedTotal * 0.99,
          actualDate,
          actualAmount,
          isAutoMarked: false,
          isManuallyEdited: true
        }]
      });
    } else {
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
      
      // Calculate new status
      const tempInv = { ...inv, installments: newInstallments, totalRepaid };
      const newStatus = getEffectiveStatus(tempInv);
      
      onUpdate(invId, { 
        installments: newInstallments,
        totalRepaid,
        status: newStatus
      });
    }
    setCustomRepaymentIdx(null);
  };

  const getInstallmentWarning = (inst: any) => {
    if (!inst.isPaid) return null;
    const reasons: string[] = [];
    if (inst.actualDate && inst.actualDate > inst.date) {
      reasons.push(`Paid late on ${formatDate(inst.actualDate)} (Scheduled: ${formatDate(inst.date)})`);
    }
    if (inst.actualAmount !== undefined && inst.actualAmount < inst.amount - 0.01) {
      reasons.push(`Paid ${formatBDT(inst.actualAmount)} (Scheduled: ${formatBDT(inst.amount)})`);
    }
    return reasons.length > 0 ? reasons.join('. ') : null;
  };

  const getInvestmentWarning = (inv: OnlineInvestment) => {
    if (inv.installments.length > 0 && !inv.isDefaultSchedule) {
      const warnings = inv.installments.map(getInstallmentWarning).filter(Boolean);
      return warnings.length > 0 ? `Issues in ${warnings.length} installment(s): ${warnings.join('; ')}` : null;
    } else {
      const expectedTotal = inv.estimatedReturn || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100));
      const reasons: string[] = [];
      if (inv.actualMaturityDate && inv.actualMaturityDate > inv.maturityDate) {
        reasons.push(`Paid late on ${formatDate(inv.actualMaturityDate)} (Scheduled: ${formatDate(inv.maturityDate)})`);
      }
      if (inv.totalRepaid > 0 && inv.totalRepaid < expectedTotal - 0.01) {
        reasons.push(`Paid ${formatBDT(inv.totalRepaid)} (Expected: ${formatBDT(expectedTotal)})`);
      }
      return reasons.length > 0 ? reasons.join('. ') : null;
    }
  };

  const handleUpdateInstallmentValue = (invId: string, index: number, field: 'date' | 'amount', value: string) => {
    const inv = investments.find(i => i.id === invId);
    if (!inv) return;

    const newInstallments = [...inv.installments];
    const val = field === 'amount' ? parseFloat(value) || 0 : value;
    newInstallments[index] = { ...newInstallments[index], [field]: val };
    
    const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount || inst.amount) : sum, 0);
    
    // Calculate new status
    const tempInv = { ...inv, installments: newInstallments, totalRepaid };
    const newStatus = getEffectiveStatus(tempInv);
    
    onUpdate(invId, { 
      installments: newInstallments,
      totalRepaid,
      status: newStatus
    });
  };

  const handleEdit = (inv: OnlineInvestment) => {
    setEditingId(inv.id);
    setFormData({
      platform: inv.platform,
      companyName: inv.companyName || '',
      projectName: inv.projectName,
      amount: inv.amount.toString(),
      estimatedReturn: inv.estimatedReturn?.toString() || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100)).toString(),
      investmentDate: inv.investmentDate,
      maturityDate: inv.maturityDate,
      status: inv.status,
      actualProfit: inv.actualProfit?.toString() || '',
      hasRepaymentSchedule: inv.hasRepaymentSchedule || false,
      confirmNegativeProfit: (inv.estimatedReturn || 0) < inv.amount,
      manualInstallments: (inv.installments || []).map(i => ({ date: i.date, amount: i.amount.toString() })),
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
      message: `Are you sure you want to delete the investment in "${inv.projectName}"?`,
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
      message: `Are you sure you want to delete ${selectedIds.length} selected online investments?`,
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

  const handleAddPlatform = () => {
    if (!newPlatformName.trim()) return;
    const updated = [...customPlatforms, newPlatformName.trim()];
    setCustomPlatforms(updated);
    localStorage.setItem('fintrack_custom_platforms', JSON.stringify(updated));
    setNewPlatformName('');
    setIsAddingPlatform(false);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setFormError(null);

    const amount = parseFloat(formData.amount) || 0;
    const estimatedReturn = parseFloat(formData.estimatedReturn) || 0;
    const actualProfit = formData.status === 'Completed' && formData.actualProfit ? parseFloat(formData.actualProfit.toString()) : undefined;
    const profit = estimatedReturn - amount;
    const duration = calculatedDuration;
    
    if (amount <= 0) {
      setFormError("Please enter a valid investment amount");
      return;
    }

    if (estimatedReturn <= 0) {
      setFormError("Please enter a valid estimated return");
      return;
    }

    if (duration <= 0) {
      setFormError("Maturity date must be after investment date");
      return;
    }

    // Warning for potential confusion between Return and Profit
    if (estimatedReturn < amount && !formData.confirmNegativeProfit) {
      setFormError(`Estimated return is less than investment amount. If this is correct, please check the "Confirm Loss/Negative Profit" box.`);
      return;
    }

    const roe = amount > 0 ? (profit / amount) * (12 / duration) * 100 : 0;

    let installments = [];
    let isDefaultSchedule = false;

    if (formData.hasRepaymentSchedule && formData.manualInstallments.length > 0) {
      installments = formData.manualInstallments.map(inst => ({
        date: inst.date,
        amount: parseFloat(inst.amount) || 0,
        isPaid: false
      }));
    } else {
      // Default installment logic
      isDefaultSchedule = true;
      installments = [{
        date: formData.maturityDate,
        amount: estimatedReturn,
        isPaid: false
      }];
    }

    const payload = {
      name: formData.projectName,
      platform: formData.platform as any,
      companyName: formData.companyName,
      projectName: formData.projectName,
      amount,
      estimatedReturn,
      estimatedProfit: profit,
      expectedROE: parseFloat(roe.toFixed(2)),
      investmentDate: formData.investmentDate,
      durationMonths: duration,
      maturityDate: formData.maturityDate,
      status: formData.status,
      actualProfit: actualProfit,
      currency: 'BDT' as const,
      hasRepaymentSchedule: formData.hasRepaymentSchedule,
      isDefaultSchedule,
      installments,
      totalRepaid: editingId ? (investments.find(i => i.id === editingId)?.totalRepaid || 0) : 0,
    };

    if (editingId) {
      onUpdate(editingId, payload);
    } else {
      onAdd(payload);
      if (onNavigateToModule) {
        onNavigateToModule('income-expense', {
          date: payload.investmentDate,
          amount: payload.amount,
          type: 'Transfer',
          targetModule: 'online',
          description: `Online Invest: ${payload.platform}${payload.projectName ? ` - ${payload.projectName}` : payload.companyName ? ` - ${payload.companyName}` : ''}`.trim()
        });
      }
    }

    setIsModalOpen(false);
    setEditingId(null);
    setFormError(null);
    setFormData({
      platform: 'Biniyog',
      companyName: '',
      projectName: '',
      amount: '',
      estimatedReturn: '',
      investmentDate: new Date().toISOString().split('T')[0],
      maturityDate: '',
      status: 'Active',
      hasRepaymentSchedule: false,
      confirmNegativeProfit: false,
      manualInstallments: [],
    });
  };

  const addInstallment = () => {
    setFormData({
      ...formData,
      manualInstallments: [...formData.manualInstallments, { date: '', amount: '' }]
    });
  };

  const removeInstallment = (index: number) => {
    const newInst = [...formData.manualInstallments];
    newInst.splice(index, 1);
    setFormData({ ...formData, manualInstallments: newInst });
  };

  const updateInstallment = (index: number, field: 'date' | 'amount', value: string) => {
    const newInst = [...formData.manualInstallments];
    newInst[index] = { ...newInst[index], [field]: value };
    setFormData({ ...formData, manualInstallments: newInst });
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
                  onClick={() => { exportToExcel(); setIsSettingsMenuOpen(false); }}
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
                <div className="px-3 pb-2">
                  <input 
                    placeholder="Platform Name"
                    value={newPlatformName}
                    onChange={(e) => setNewPlatformName(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-[10px] text-white focus:outline-none focus:border-teal-400"
                  />
                  <button 
                    onClick={handleAddPlatform}
                    className="w-full mt-2 flex items-center justify-center gap-2 py-1.5 bg-teal-400 hover:bg-teal-300 text-slate-950 rounded-lg font-bold text-[10px] uppercase transition-colors"
                  >
                    Add Platform
                  </button>
                </div>
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
              onClick={() => { exportToExcel(); setIsSettingsMenuOpen(false); }}
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

            <div className="px-3 pb-2">
              <input 
                placeholder="Platform Name"
                value={newPlatformName}
                onChange={(e) => setNewPlatformName(e.target.value)}
                className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-[10px] text-white focus:outline-none focus:border-teal-400"
              />
              <button 
                onClick={handleAddPlatform}
                className="w-full mt-2 flex items-center justify-center gap-2 py-1.5 bg-teal-400 hover:bg-teal-300 text-slate-950 rounded-lg font-bold text-[10px] uppercase transition-colors"
              >
                Add Platform
              </button>
            </div>
          
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
            <p className="text-label font-medium text-slate-300 uppercase">Active Projects</p>
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
              <p className="text-body-sm font-bold text-white uppercase tracking-wider">Projected Profit</p>
              <p className="text-label font-bold text-slate-500 uppercase">On Active Investment</p>
            </div>
            <h3 className="text-heading font-bold text-white mb-3 tracking-tight font-display tabular-nums">
              {formatBDT(historyStats.projectedProfit)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Projected ROI</p>
            <p className={cn("text-body font-bold tabular-nums", historyStats.avgProjectedROI >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {historyStats.avgProjectedROI.toFixed(1)}%
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
            <h3 className="text-heading font-bold text-emerald-400 mb-4 tracking-tight font-display tabular-nums">
              {formatBDT(historyStats.profit)}
            </h3>
          </div>
          <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
            <p className="text-label font-medium text-slate-300 uppercase">Portfolio Yield</p>
            <p className={cn("text-body font-bold tabular-nums", historyStats.avgRealizedROI >= 0 ? "text-emerald-400" : "text-rose-500")}>
              {historyStats.avgRealizedROI.toFixed(1)}%
            </p>
          </div>
        </Card>
      </div>

      {/* Bottom Analytics Panel */}
      <div className="grid grid-cols-1 lg:grid-cols-[6fr_4fr] gap-4">
        <Card className="bg-slate-900 border-slate-800">
          <div className="flex items-center gap-2 mb-6">
            <Calendar className="text-teal-400 w-5 h-5" />
            <h3 className="text-subheading font-bold text-white uppercase">Upcoming Installments</h3>
          </div>

          <div className="grid grid-cols-2 gap-3">
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
                    <span className={i === 0 ? "text-emerald-400" : "text-slate-300"}>{item.platform}</span>
                    <span className="mx-1 text-slate-800">|</span>
                    <span className="text-slate-300">{item.companyName}</span>
                  </p>
                </div>
              </div>
            )) : (
              <div className="col-span-2 py-6 text-center bg-slate-950/30 border border-dashed border-slate-800 rounded-xl">
                <p className="text-xs text-slate-600 italic">No upcoming installments scheduled</p>
              </div>
            )}
          </div>
        </Card>

        {/* AI Insights Panel */}
        <Card className="bg-slate-900 border-slate-800 flex flex-col">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
            <div className="flex items-center gap-2">
              <Sparkles className="text-teal-400 w-5 h-5" />
              <h3 className="text-subheading font-bold text-white uppercase">AI Insights</h3>
            </div>

            <div className="relative">
              <div className="flex items-center gap-2 bg-slate-950 border border-slate-800 rounded-lg px-3 h-9 focus-within:border-teal-400/50 transition-all group">
                <Search size={14} className="text-slate-500 group-focus-within:text-teal-400" />
                <input 
                  type="text"
                  placeholder="SEARCH COMPANY..."
                  className="bg-transparent border-none focus:ring-0 text-[10px] font-bold text-white placeholder:text-slate-600 w-32 uppercase outline-none"
                  value={insightSearch}
                  onChange={(e) => {
                    setInsightSearch(e.target.value);
                    setIsInsightSearchOpen(true);
                  }}
                  onFocus={() => setIsInsightSearchOpen(true)}
                />
                {insightSearch && (
                  <button onClick={() => { setInsightSearch(''); setSelectedInsightCompany(null); }} className="text-slate-500 hover:text-white">
                    <X size={12} />
                  </button>
                )}
              </div>

              {isInsightSearchOpen && insightSearch && (
                <>
                  <div className="fixed inset-0 z-40" onClick={() => setIsInsightSearchOpen(false)} />
                  <div className="absolute right-0 mt-2 w-full min-w-[200px] bg-slate-900 border border-slate-800 rounded-lg shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
                    {allCompanies
                      .filter(c => c.toLowerCase().includes(insightSearch.toLowerCase()))
                      .slice(0, 5)
                      .map(company => (
                        <button
                          key={company}
                          onClick={() => {
                            setSelectedInsightCompany(company);
                            setInsightSearch(company);
                            setIsInsightSearchOpen(false);
                          }}
                          className="w-full text-left px-3 py-2 text-[10px] font-bold text-slate-300 hover:bg-slate-800 hover:text-teal-400 rounded transition-colors uppercase"
                        >
                          {company}
                        </button>
                      ))}
                    {allCompanies.filter(c => c.toLowerCase().includes(insightSearch.toLowerCase())).length === 0 && (
                      <div className="px-3 py-2 text-[10px] font-bold text-slate-600 uppercase">No results found</div>
                    )}
                  </div>
                </>
              )}
            </div>
          </div>
          
          <div className="flex-1 flex flex-col justify-center">
            {insights ? (
              <div className="space-y-4">
                <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 shadow-inner">
                  <div className="flex items-center justify-between mb-3">
                    <p className="text-label font-bold text-slate-500 uppercase tracking-wider">
                      {insights.isBest ? 'Top Performing Company' : 'Company Performance'}
                    </p>
                    {insights.isBest && (
                      <div className="px-2 py-0.5 bg-emerald-400/10 rounded text-[9px] font-bold text-emerald-400 uppercase tracking-widest border border-emerald-400/20">Recommended</div>
                    )}
                  </div>
                  <p className="text-subheading font-bold text-white uppercase mb-3 tracking-tight">{insights.name}</p>
                  <div className="grid grid-cols-2 gap-3">
                    <div className="p-2 bg-slate-900/50 rounded-lg border border-slate-800/50">
                      <p className="text-[9px] font-bold text-slate-500 uppercase mb-1">Annualized Yield</p>
                      <p className="text-body-sm font-bold text-emerald-400">{insights.avgROI.toFixed(1)}%</p>
                    </div>
                    <div className="p-2 bg-slate-900/50 rounded-lg border border-slate-800/50">
                      <p className="text-[9px] font-bold text-slate-500 uppercase mb-1">Reliability</p>
                      <p className="text-body-sm font-bold text-blue-400">{insights.reliability.toFixed(0)}%</p>
                    </div>
                  </div>
                </div>
                <div className="flex gap-3 items-start p-3 bg-teal-400/5 rounded-xl border border-teal-400/10">
                  <Sparkles size={16} className="text-teal-400 shrink-0 mt-0.5" />
                  <p className="text-label text-slate-400 leading-relaxed italic">
                    {insights.isBest ? (
                      <>Based on historical data, <span className="text-teal-400 font-bold">{insights.name}</span> has demonstrated the most consistent repayment schedule and superior <span className="text-emerald-400">annualized yield</span> performance across your portfolio.</>
                    ) : (
                      <>Analysis for <span className="text-teal-400 font-bold">{insights.name}</span> shows an <span className="text-emerald-400">annualized yield</span> of <span className="text-emerald-400">{insights.avgROI.toFixed(1)}%</span> with a reliability score of <span className="text-blue-400">{insights.reliability.toFixed(0)}%</span> based on your investment history.</>
                    )}
                  </p>
                </div>
              </div>
            ) : (
              <div className="py-12 text-center bg-slate-950/30 border border-dashed border-slate-800 rounded-xl">
                <p className="text-label text-slate-500 italic uppercase tracking-widest">Not enough data to generate insights</p>
              </div>
            )}
          </div>
        </Card>
      </div>

      {/* Portfolio Header & Filters */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-6">
        <div className="flex items-center gap-2 sm:gap-3">
          <Briefcase className="text-teal-400 w-5 h-5 lg:w-6 lg:h-6" />
          <h2 className="text-body-sm sm:text-subheading lg:text-heading font-bold text-white font-display uppercase whitespace-nowrap">Investment Portfolio</h2>
        </div>
        <div className="flex flex-wrap items-center gap-2 justify-end flex-1 sm:flex-none">
          <div className="relative">
            <div 
              className="flex items-center gap-1.5 bg-slate-900 border border-slate-800 rounded-lg p-1 px-2 h-9 cursor-pointer hover:border-slate-700 transition-colors"
              onClick={() => setIsProfileMenuOpen(!isProfileMenuOpen)}
            >
              <span className="text-[10px] font-bold text-slate-300 uppercase shrink-0">Profile:</span>
              <span className="text-[10px] font-bold uppercase text-blue-400 tabular-nums">
                {selectedProfiles.length === 0 ? 'All' : selectedProfiles.length === 3 ? 'All' : selectedProfiles.length === 1 ? selectedProfiles[0] : `${selectedProfiles.length}`}
              </span>
              <MoreVertical size={12} className="text-slate-500 shrink-0" />
            </div>
            
            {isProfileMenuOpen && (
              <>
                <div className="fixed inset-0 z-40" onClick={() => setIsProfileMenuOpen(false)} />
                <div className="absolute right-0 mt-2 w-40 bg-slate-900 border border-slate-800 rounded-lg shadow-xl z-50 p-2 animate-in fade-in zoom-in-95">
                  {['Active', 'Completed', 'Delayed'].map((p) => (
                    <label key={p} className="flex items-center gap-2 px-2 py-1.5 hover:bg-slate-800 rounded cursor-pointer transition-colors">
                      <input 
                        type="checkbox" 
                        className="w-3 h-3 rounded border-slate-700 bg-slate-950 text-teal-400 focus:ring-teal-400/20"
                        checked={selectedProfiles.includes(p as any)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedProfiles([...selectedProfiles, p as any]);
                          } else {
                            setSelectedProfiles(selectedProfiles.filter(sp => sp !== p));
                          }
                        }}
                      />
                      <span className="text-label font-bold uppercase text-slate-300">{p}</span>
                    </label>
                  ))}
                  <div className="border-t border-slate-800 mt-2 pt-2">
                    <button 
                      onClick={() => setSelectedProfiles([])}
                      className="w-full text-left px-2 py-1 text-label font-bold text-slate-500 hover:text-white uppercase"
                    >
                      Clear All
                    </button>
                  </div>
                </div>
              </>
            )}
          </div>

          <div className="relative flex items-center gap-1">
            <div 
              className="flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-lg p-1 px-2 sm:px-3 h-8 sm:h-9 cursor-pointer hover:border-slate-700 transition-colors"
              onClick={() => setIsSortMenuOpen(!isSortMenuOpen)}
            >
              <span className="text-[10px] sm:text-label font-bold text-slate-300 uppercase">Sort By:</span>
              <span className="text-[10px] sm:text-label font-bold uppercase text-teal-400">
                {sortBy}
              </span>
              <MoreVertical size={12} className="text-slate-500" />
            </div>

            <button 
              onClick={() => setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc')}
              className="flex items-center justify-center w-8 h-8 sm:w-9 sm:h-9 bg-slate-900 border border-slate-800 rounded-lg text-teal-400 hover:border-teal-400/50 transition-all"
              title={sortOrder === 'asc' ? 'Sort Ascending' : 'Sort Descending'}
            >
              {sortOrder === 'asc' ? <ArrowRight size={14} className="-rotate-90" /> : <ArrowRight size={14} className="rotate-90" />}
            </button>
            
            {isSortMenuOpen && (
              <>
                <div className="fixed inset-0 z-40" onClick={() => setIsSortMenuOpen(false)} />
                <div className="absolute right-0 mt-2 w-32 bg-slate-900 border border-slate-800 rounded-lg shadow-xl z-50 p-1 animate-in fade-in zoom-in-95">
                  {(['amount', 'date', 'platform', 'roi'] as const).map((s) => (
                    <button
                      key={s}
                      onClick={() => {
                        setSortBy(s);
                        setIsSortMenuOpen(false);
                      }}
                      className={cn(
                        "w-full text-left px-2 py-1.5 text-[10px] font-bold rounded transition-colors uppercase",
                        sortBy === s ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                      )}
                    >
                      {s}
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>

          <div className="relative">
            <div 
              className="flex items-center gap-1.5 bg-slate-900 border border-slate-800 rounded-lg p-1 px-2 h-9 hover:border-slate-700 focus-within:border-teal-400/50 transition-colors group cursor-text"
              onClick={() => {
                const input = document.getElementById('company-search-input');
                if (input) input.focus();
              }}
            >
              <Search size={14} className="text-slate-500 group-focus-within:text-teal-400 transition-colors shrink-0" />
              <input 
                id="company-search-input"
                type="text"
                placeholder="Search"
                className="flex-1 bg-transparent border-none focus:ring-0 text-[10px] font-bold text-white placeholder:text-slate-600 uppercase caret-teal-400 outline-none min-w-0 w-24 sm:w-32"
                value={companySearch}
                onChange={(e) => setCompanySearch(e.target.value)}
              />
              {companySearch && (
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    setCompanySearch('');
                  }} 
                  className="text-slate-500 hover:text-white"
                >
                  <X size={12} />
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className="space-y-3">
        <div className="flex items-center justify-between bg-slate-900/50 border border-slate-800 rounded-xl px-3 py-1.5 h-10">
          <div className="flex items-center gap-2">
            <Checkbox 
              label={<span className="text-[10px] font-bold uppercase text-slate-300">Select All</span>}
              checked={selectedIds.length === filtered.length && filtered.length > 0} 
              onChange={toggleSelectAll} 
            />
          </div>
          <span className="text-[10px] font-bold text-slate-500 uppercase tabular-nums">
            {selectedIds.length} of {filtered.length} selected
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
          const expectedTotal = inv.estimatedReturn || (inv.amount + (inv.amount * inv.expectedROE * (inv.durationMonths / 12) / 100));
          const timeProgress = Math.round(Math.min(100, (new Date().getTime() - new Date(inv.investmentDate).getTime()) / (new Date(inv.maturityDate).getTime() - new Date(inv.investmentDate).getTime()) * 100));
          const repaymentProgress = Math.min(100, (inv.totalRepaid / expectedTotal) * 100);
          const actualROE = inv.totalRepaid > 0 ? ((inv.totalRepaid - inv.amount) / inv.amount) * (12 / inv.durationMonths) * 100 : 0;

          return (
            <div 
              key={inv.id} 
              className={cn(
                "bg-slate-900 border rounded-xl overflow-hidden transition-all group",
                selectedIds.includes(inv.id) ? "border-teal-400 ring-1 ring-teal-400/20 shadow-[0_0_15px_rgba(45,212,191,0.1)]" : "border-slate-800 hover:border-teal-400/50"
              )}
            >
              {/* Title Bar */}
              <div className="bg-slate-950 border-b border-slate-800 px-4 py-2 flex justify-between items-center relative">
                <div className="flex items-center gap-3">
                  <Checkbox 
                    checked={selectedIds.includes(inv.id)}
                    onChange={() => toggleSelect(inv.id)}
                  />
                  <div className={cn("w-2 h-2 rounded-full animate-pulse", status === 'Completed' ? 'bg-slate-500' : 'bg-teal-400')} />
                  <span className="text-[9px] sm:text-body-sm font-bold uppercase tracking-tight tabular-nums truncate max-w-[180px] sm:max-w-none">
                    <span className="text-slate-300">{inv.platform}</span>
                    <span className="text-slate-800 mx-1 sm:mx-2">|</span>
                    <span className="text-slate-300">{inv.companyName}</span>
                    <span className="text-slate-800 mx-1 sm:mx-2">|</span>
                    <span className="text-teal-400">{inv.projectName}</span>
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => {
                      setEditingActualId(editingActualId === inv.id ? null : inv.id);
                      setIsEditingSchedule(false);
                    }}
                    className={cn(
                      "p-1.5 rounded-lg transition-colors",
                      editingActualId === inv.id ? "bg-teal-400 text-white" : "text-slate-500 hover:text-white hover:bg-slate-800"
                    )}
                    title="Repayment Tracking"
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
                          EDIT INFO
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
                {/* Card 1: Investment Info */}
                <div className="space-y-1 min-w-[120px] shrink-0">
                  <p className="text-label text-slate-300 uppercase mb-1">Investment Info</p>
                  <p className="text-body-sm font-bold text-white uppercase tabular-nums">{formatBDT(inv.amount)}</p>
                  <p className="text-label text-slate-500 font-medium tabular-nums">{formatDate(inv.investmentDate)}</p>
                </div>

                {/* Card 2: Duration */}
                <div className="space-y-1 min-w-[100px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">Duration</p>
                  <p className="text-body-sm font-bold text-white uppercase tabular-nums">{inv.durationMonths} Months</p>
                  <p className="text-label text-slate-500 font-medium tabular-nums">{formatDate(inv.maturityDate)}</p>
                </div>

                {/* Card 3: Expected Returns */}
                <div className="space-y-1 min-w-[120px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">Expected Returns</p>
                  <p className="text-body-sm font-bold text-white uppercase tabular-nums">{formatBDT(expectedTotal)}</p>
                  <p className={cn("text-label font-bold uppercase tabular-nums", inv.expectedROE >= 0 ? "text-emerald-400" : "text-rose-500")}>ROI: {inv.expectedROE.toFixed(1)}%</p>
                </div>

                {/* Card 4: Actual Performance */}
                <div className="space-y-1 min-w-[120px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">Actual Performance</p>
                  <p className={cn("text-body-sm font-bold uppercase tabular-nums", status === 'Active' ? "text-blue-400" : (inv.totalRepaid >= inv.amount ? "text-emerald-400" : "text-rose-500"))}>
                    {formatBDT(inv.totalRepaid)}
                  </p>
                  <div className="flex items-center gap-1 h-4">
                    {status === 'Completed' ? (
                      <span className={cn("text-label font-bold tabular-nums", actualROE >= 0 ? "text-emerald-400" : "text-rose-500")}>ROI: {actualROE.toFixed(1)}%</span>
                    ) : (
                      <div className="flex gap-0.5">
                        {inv.installments.length > 0 ? (
                          inv.installments.map((inst, i) => (
                            <CheckCircle2 
                              key={i} 
                              size={10} 
                              className={inst.isPaid ? "text-emerald-400" : "text-slate-800"} 
                            />
                          ))
                        ) : (
                          <CheckCircle2 
                            size={10} 
                            className={inv.totalRepaid >= expectedTotal ? "text-emerald-400" : "text-slate-800"} 
                          />
                        )}
                      </div>
                    )}
                  </div>
                </div>

                {/* Card 5: Status */}
                <div className="space-y-1 min-w-[120px] shrink-0 border-l border-slate-800 pl-4 lg:pl-6">
                  <p className="text-label text-slate-300 uppercase mb-1">Status</p>
                  <div className="flex items-center gap-2 h-5">
                    {status === 'Completed' ? (
                      <div className="flex items-center gap-1.5 text-body-sm font-bold text-emerald-400 uppercase">
                        <CheckCircle2 size={12} />
                        <span>Completed</span>
                        {getInvestmentWarning(inv) && (
                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              setActiveWarning({ id: inv.id, text: getInvestmentWarning(inv)! });
                            }}
                            className="inline-flex items-center justify-center w-3 h-3 rounded-full bg-amber-500 text-slate-950 text-[8px] font-black hover:bg-amber-400 transition-colors cursor-pointer"
                          >
                            !
                          </button>
                        )}
                      </div>
                    ) : status === 'Delayed' ? (
                      <div className="flex items-center gap-1.5 text-body-sm font-bold text-rose-500 uppercase">
                        <AlertCircle size={12} />
                        <span>Delayed</span>
                      </div>
                    ) : (
                      <div className="flex items-center gap-1.5 text-body-sm font-bold text-blue-400 uppercase">
                        <Clock size={12} />
                        <span>Active</span>
                      </div>
                    )}
                  </div>
                  <div className="flex items-center justify-between gap-2 mt-1">
                    <div className="flex-1 bg-slate-950 h-1.5 rounded-full overflow-hidden border border-slate-800">
                      <div 
                        className="bg-emerald-400 h-full rounded-full shadow-[0_0_10px_rgba(52,211,153,0.3)]" 
                        style={{ width: `${repaymentProgress}%` }} 
                      />
                    </div>
                    <span className="text-label font-bold text-emerald-400 min-w-[32px] tabular-nums text-right">{Math.round(repaymentProgress)}%</span>
                  </div>
                </div>
              </div>

              {/* Inline Editor for Actual Performance */}
              {editingActualId === inv.id && (
                <div className="w-full px-4 pb-4 animate-in fade-in slide-in-from-top-2">
                  <div className="pt-4 border-t border-slate-800">
                    <div className="flex justify-between items-center mb-4">
                      <h5 className="text-label font-bold text-slate-300 uppercase">Repayment Tracking</h5>
                      <button 
                        onClick={() => setIsEditingSchedule(!isEditingSchedule)}
                        className={cn(
                          "text-label font-bold uppercase px-3 py-1 rounded-full border transition-all",
                          isEditingSchedule 
                            ? "bg-teal-400 border-teal-400 text-white" 
                            : "bg-slate-800 border-slate-700 text-slate-300 hover:text-white"
                        )}
                      >
                        {isEditingSchedule ? 'Finish Editing' : 'Edit Schedule'}
                      </button>
                    </div>

                    <div className="flex flex-wrap gap-3">
                      {inv.installments.length > 0 ? (
                        inv.installments.map((inst, idx) => (
                          <div key={idx} className="flex items-center gap-2">
                            {isEditingSchedule ? (
                              <div className="flex items-center gap-2 bg-slate-950 p-2 rounded-lg border border-slate-800">
                                <input 
                                  type="date" 
                                  className="bg-transparent text-label text-white border-none focus:ring-0 p-0 w-24 tabular-nums"
                                  value={inst.date}
                                  onChange={(e) => handleUpdateInstallmentValue(inv.id, idx, 'date', e.target.value)}
                                />
                                <div className="w-px h-4 bg-slate-800" />
                                <input 
                                  type="number" 
                                  className="bg-transparent text-label text-emerald-400 border-none focus:ring-0 p-0 w-16 font-bold tabular-nums"
                                  value={inst.amount}
                                  onChange={(e) => handleUpdateInstallmentValue(inv.id, idx, 'amount', e.target.value)}
                                />
                              </div>
                            ) : customRepaymentIdx === idx ? (
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
                                {(() => {
                                  const warning = getInstallmentWarning(inst);
                                  return (
                                    <div
                                      role="button"
                                      tabIndex={0}
                                      onClick={() => handleToggleInstallment(inv.id, idx)}
                                      onKeyDown={(e) => e.key === 'Enter' && handleToggleInstallment(inv.id, idx)}
                                      className={cn(
                                        "flex items-center gap-2 px-3 py-2 rounded-lg border transition-all cursor-pointer",
                                        inst.isPaid 
                                          ? (warning ? "bg-amber-400/10 border-amber-400/30 text-amber-400" : "bg-emerald-400/10 border-emerald-400/30 text-emerald-400")
                                          : "bg-slate-950 border-slate-800 text-slate-500 hover:border-slate-600"
                                      )}
                                    >
                                      <CheckCircle2 size={14} />
                                      <div className="text-left">
                                        <p className="text-label font-bold tabular-nums flex items-center gap-1.5">
                                          {formatDate(inst.actualDate || inst.date)}
                                          {warning && (
                                            <button 
                                              onClick={(e) => {
                                                e.stopPropagation();
                                                setActiveWarning({ id: `${inv.id}-${idx}`, text: warning });
                                              }}
                                              className="inline-flex items-center justify-center w-2.5 h-2.5 rounded-full bg-amber-500 text-slate-950 text-[7px] font-black hover:bg-amber-400 transition-colors cursor-pointer"
                                            >
                                              !
                                            </button>
                                          )}
                                        </p>
                                        <p className="text-label opacity-70 tabular-nums">
                                          {formatBDT(inst.actualAmount !== undefined ? inst.actualAmount : inst.amount)}
                                        </p>
                                      </div>
                                    </div>
                                  );
                                })()}
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
                                  title="Custom Repayment"
                                >
                                  <Edit2 size={10} />
                                </button>
                              </div>
                            )}
                          </div>
                        ))
                      ) : (
                        <div className="flex items-center gap-4 w-full">
                          {customRepaymentIdx === -1 ? (
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
                                onClick={() => handleSaveCustomRepayment(inv.id, -1)}
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
                              {(() => {
                                const maturityWarning = getInvestmentWarning(inv);
                                return (
                                  <div
                                    role="button"
                                    tabIndex={0}
                                    onClick={() => {
                                      const isPaid = inv.totalRepaid >= expectedTotal;
                                      onUpdate(inv.id, { 
                                        totalRepaid: isPaid ? 0 : expectedTotal,
                                        status: isPaid ? 'Active' : 'Completed'
                                      });
                                    }}
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter') {
                                        const isPaid = inv.totalRepaid >= expectedTotal;
                                        onUpdate(inv.id, { 
                                          totalRepaid: isPaid ? 0 : expectedTotal,
                                          status: isPaid ? 'Active' : 'Completed'
                                        });
                                      }
                                    }}
                                    className={cn(
                                      "flex items-center gap-2 px-3 py-2 rounded-lg border transition-all cursor-pointer",
                                      inv.totalRepaid >= expectedTotal 
                                        ? (maturityWarning ? "bg-amber-400/10 border-amber-400/30 text-amber-400" : "bg-emerald-400/10 border-emerald-400/30 text-emerald-400")
                                        : "bg-slate-950 border-slate-800 text-slate-500 hover:border-slate-600"
                                    )}
                                  >
                                    <CheckCircle2 size={14} />
                                    <div className="text-left">
                                      <p className="text-label font-bold tabular-nums flex items-center gap-1.5">
                                        {formatDate(inv.actualMaturityDate || inv.maturityDate)}
                                        {maturityWarning && (
                                          <button 
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              setActiveWarning({ id: `${inv.id}-maturity`, text: maturityWarning });
                                            }}
                                            className="inline-flex items-center justify-center w-2.5 h-2.5 rounded-full bg-amber-500 text-slate-950 text-[7px] font-black hover:bg-amber-400 transition-colors cursor-pointer"
                                          >
                                            !
                                          </button>
                                        )}
                                      </p>
                                      <p className="text-label opacity-70 tabular-nums">{formatBDT(inv.totalRepaid || expectedTotal)}</p>
                                    </div>
                                  </div>
                                );
                              })()}
                              <button 
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setCustomRepaymentIdx(-1);
                                  setCustomRepaymentData({ 
                                    date: inv.actualMaturityDate || inv.maturityDate, 
                                    amount: (inv.totalRepaid || expectedTotal).toString() 
                                  });
                                }}
                                className="absolute -top-2 -right-2 p-1 bg-slate-800 border border-slate-700 rounded-full text-slate-400 hover:text-white opacity-0 group-hover/inst:opacity-100 transition-opacity"
                                title="Custom Repayment"
                              >
                                <Edit2 size={10} />
                              </button>
                            </div>
                          )}
                          <p className="text-label text-slate-500 italic">No schedule set. Single repayment at maturity.</p>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              )}
            </div>
          );
        })}

        {filtered.length === 0 && (
          <div className="py-20 flex flex-col items-center justify-center text-slate-500 border-2 border-dashed border-slate-800 rounded-xl bg-slate-900/20">
            <Globe size={48} className="mb-4 opacity-20" />
            <p className="text-subheading font-bold text-slate-300 uppercase">No investments found</p>
            <p className="text-body-sm text-slate-500">Try changing your filters or add a new one.</p>
          </div>
        )}
      </div>

      <Modal 
        isOpen={isModalOpen} 
        onClose={() => {
          setIsModalOpen(false);
          setEditingId(null);
        }} 
        title={editingId ? "Edit Investment" : "Add New Investment"}
      >
        <form onSubmit={handleSubmit} className="space-y-4">
          {formError && (
            <div className="p-3 bg-rose-500/10 border border-rose-500/20 rounded-lg flex items-center gap-2 text-rose-500 text-label font-bold animate-in fade-in slide-in-from-top-1">
              <AlertCircle size={14} />
              <span className="uppercase">{formError}</span>
            </div>
          )}
          <div className="grid grid-cols-3 gap-3">
            <Select 
              label="Platform" 
              value={formData.platform}
              options={[
                { label: 'Biniyog', value: 'Biniyog' },
                { label: 'iFarmer', value: 'iFarmer' },
                { label: 'Wegro', value: 'Wegro' },
                ...customPlatforms.map(p => ({ label: p, value: p })),
                { label: 'Other', value: 'Other' }
              ]}
              onChange={(val) => setFormData({ ...formData, platform: val })}
              required
            />
            <Input 
              label="Company Name" 
              placeholder="e.g. Agro Ltd"
              value={formData.companyName}
              onChange={(e) => setFormData({ ...formData, companyName: e.target.value })}
              required
            />
            <Input 
              label="Project Name" 
              placeholder="e.g. Cattle Farm 1"
              value={formData.projectName}
              onChange={(e) => setFormData({ ...formData, projectName: e.target.value })}
              required
            />
          </div>

          <div className="grid grid-cols-2 gap-3">
            <Input 
              label="Amount (BDT)" 
              type="number" 
              placeholder="0.00"
              value={formData.amount}
              onChange={(e) => setFormData({ ...formData, amount: e.target.value })}
              required
              className="tabular-nums"
            />
            <Input 
              label="Est. Return (BDT)" 
              type="number" 
              placeholder="0.00"
              value={formData.estimatedReturn}
              onChange={(e) => setFormData({ ...formData, estimatedReturn: e.target.value })}
              required
              className="tabular-nums"
            />
          </div>

          <div className="grid grid-cols-2 gap-3">
            <Input 
              label="Investment Date" 
              type="date"
              value={formData.investmentDate}
              onChange={(e) => setFormData({ ...formData, investmentDate: e.target.value })}
              required
              className="tabular-nums"
            />
            <Input 
              label="Maturity Date" 
              type="date"
              value={formData.maturityDate}
              onChange={(e) => setFormData({ ...formData, maturityDate: e.target.value })}
              required
              className="tabular-nums"
            />
          </div>
          
          {/* ROE & Duration Preview */}
          {(formData.amount && formData.estimatedReturn && calculatedDuration > 0) && (
            <div className="space-y-3">
              <div className="p-2 bg-emerald-400/10 border border-emerald-400/20 rounded-lg flex flex-col gap-2 text-label">
                <div className="flex justify-between items-center">
                  <span className="font-bold text-slate-300 uppercase">Estimated Profit:</span>
                  <span className={cn("font-bold tabular-nums", (parseFloat(formData.estimatedReturn) - parseFloat(formData.amount)) < 0 ? "text-rose-500" : "text-emerald-400")}>
                    {formatBDT(parseFloat(formData.estimatedReturn) - parseFloat(formData.amount))}
                  </span>
                </div>
                <div className="flex gap-4 border-t border-emerald-400/10 pt-2">
                  <span className="font-bold text-emerald-400 uppercase">ROE: <span className="text-emerald-400 font-bold tabular-nums">{(((parseFloat(formData.estimatedReturn) - parseFloat(formData.amount)) / parseFloat(formData.amount)) * (12 / calculatedDuration) * 100).toFixed(2)}%</span></span>
                  <span className="font-bold text-slate-300 uppercase">Duration: <span className="text-white tabular-nums">{calculatedDuration} Months</span></span>
                </div>
              </div>

              {(parseFloat(formData.estimatedReturn) < parseFloat(formData.amount)) && (
                <div className="p-3 bg-rose-500/10 border border-rose-500/20 rounded-xl">
                  <Checkbox 
                    label="Confirm Loss/Negative Profit" 
                    checked={formData.confirmNegativeProfit}
                    onChange={(checked) => setFormData({ ...formData, confirmNegativeProfit: checked })}
                  />
                  <p className="text-[10px] text-rose-400 mt-1 uppercase font-bold">Warning: Estimated return is less than investment amount. Did you mean to enter the profit instead?</p>
                </div>
              )}
            </div>
          )}

          <div className="p-3 bg-slate-950 border border-slate-800 rounded-xl space-y-3">
            <Checkbox 
              label="Enable Repayment Schedule" 
              checked={formData.hasRepaymentSchedule}
              onChange={(checked) => setFormData({ ...formData, hasRepaymentSchedule: checked })}
            />
            {formData.hasRepaymentSchedule && (
              <div className="space-y-2 animate-in fade-in slide-in-from-top-2">
                <div className="flex justify-between items-center">
                  <span className="text-label font-bold text-slate-300 uppercase">Installments</span>
                  <button 
                    type="button" 
                    onClick={addInstallment}
                    className="text-label font-bold text-teal-400 hover:text-teal-300 flex items-center gap-1 uppercase"
                  >
                    <Plus size={10} /> Add Row
                  </button>
                </div>
                <div className="max-h-32 overflow-y-auto space-y-2 pr-1">
                  {formData.manualInstallments.map((inst, idx) => (
                    <div key={idx} className="grid grid-cols-[1fr_1fr_auto] gap-2 items-end">
                      <Input 
                        label="Date" 
                        type="date" 
                        className="h-8 text-label py-1 tabular-nums"
                        value={inst.date}
                        onChange={(e) => updateInstallment(idx, 'date', e.target.value)}
                      />
                      <Input 
                        label="Amount" 
                        type="number" 
                        className="h-8 text-label py-1 tabular-nums"
                        placeholder="0.00"
                        value={inst.amount}
                        onChange={(e) => updateInstallment(idx, 'amount', e.target.value)}
                      />
                      <button 
                        type="button" 
                        onClick={() => removeInstallment(idx)}
                        className="p-1.5 text-rose-500 hover:bg-rose-500/10 rounded-lg mb-0.5"
                      >
                        <X size={14} />
                      </button>
                    </div>
                  ))}
                  {formData.manualInstallments.length === 0 && (
                    <p className="text-label text-slate-500 text-center py-2 italic uppercase">No installments added yet.</p>
                  )}
                </div>
              </div>
            )}
          </div>

          <Select 
            label="Status" 
            value={formData.status}
            options={[
              { label: 'Active', value: 'Active' },
              { label: 'Completed', value: 'Completed' },
              { label: 'Delayed', value: 'Delayed' }
            ]}
            onChange={(val) => setFormData({ ...formData, status: val as any })}
            required
          />

          {formData.status === 'Completed' && (
            <Input 
              label="Actual Profit Received" 
              placeholder="Defaults to calculated"
              type="number"
              value={formData.actualProfit}
              onChange={(e) => setFormData({ ...formData, actualProfit: e.target.value })}
              className="tabular-nums"
            />
          )}

          <div className="pt-2 flex gap-3">
            <Button type="button" variant="secondary" className="flex-1 py-2" onClick={() => setIsModalOpen(false)}>
              Cancel
            </Button>
            <Button type="submit" className="flex-1 py-2">
              {editingId ? "Update Investment" : "Save Investment"}
            </Button>
          </div>
        </form>
      </Modal>

      {activeWarning && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-sm animate-in fade-in">
          <div className="bg-slate-900 border border-slate-800 p-4 rounded-xl max-w-sm w-full shadow-2xl animate-in zoom-in-95">
            <div className="flex items-center gap-2 mb-3 text-amber-400">
              <AlertCircle size={18} />
              <h4 className="text-subheading font-bold uppercase">Performance Issue</h4>
            </div>
            <p className="text-body text-slate-300 mb-4 leading-relaxed">{activeWarning.text}</p>
            <Button 
              variant="secondary" 
              className="w-full bg-slate-800 hover:bg-slate-700 text-white border-slate-700"
              onClick={() => setActiveWarning(null)}
            >
              Close
            </Button>
          </div>
        </div>
      )}
      {confirmState && (
        <ConfirmDialog 
          {...confirmState}
          onCancel={confirmState.onCancel || closeConfirm}
        />
      )}
    </div>
  );
};
