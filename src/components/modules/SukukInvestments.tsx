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
  inheritedData?: { date: string; amount: number; linkedTxId?: string; description?: string; returnModule?: string; [key: string]: any } | null;
  onClearInheritedData?: () => void;
  onNavigateToModule?: (module: string, initialAddData?: any, openModal?: boolean) => void;
  onDeleteLinkedTransfer?: (groupId: string) => void;
  onUpdateLinkedTransfer?: (groupId: string, updates: { date: string; amount: number; description?: string }) => void;
  onDeleteLinkedIncomeTx?: (txId: string) => void;
  onUpdateLinkedIncomeTx?: (txId: string, updates: { date: string; amount: number; description?: string }) => void;
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
  onNavigateToModule,
  onDeleteLinkedTransfer,
  onUpdateLinkedTransfer,
  onDeleteLinkedIncomeTx,
  onUpdateLinkedIncomeTx
}) => {
  const isFromExternalModuleRef = React.useRef(false);
  const returnModuleRef = React.useRef<string | undefined>(undefined);
  const [inheritedIeGroupId, setInheritedIeGroupId] = useState<string | undefined>(undefined);

  const checkAndReturnToModule = () => {
    if (returnModuleRef.current && onNavigateToModule) {
      const ret = returnModuleRef.current;
      returnModuleRef.current = undefined;
      onNavigateToModule(ret, undefined, false);
    }
  };
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
    onCancel?: () => void;
    variant?: 'danger' | 'warning' | 'info' | 'severe' | 'critical';
    confirmLabel?: string;
    cancelLabel?: string;
    details?: (string | null | undefined)[];
  } | null>(null);

  const closeConfirm = () => setConfirmState(prev => prev ? { ...prev, isOpen: false } : null);
  const [customRepaymentIdx, setCustomRepaymentIdx] = useState<number | null>(null);
  const [customRepaymentData, setCustomRepaymentData] = useState({ date: '', amount: '' });
  const [activeWarning, setActiveWarning] = useState<{ id: string, text: string } | null>(null);
  const [formError, setFormError] = useState<string | null>(null);

  // ─── Date Range Configuration (same range selector as Mutual Funds) ───
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

  const formatHistoryThisMonthLabel = (d: Date) => {
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return `${months[d.getMonth()]} ${d.getFullYear()}`;
  };

  // Derive active date range for dashboard stats.
  // This uses the same This Month / Fiscal / Custom semantics as Mutual Funds.
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

  const [isRangeMenuOpen, setIsRangeMenuOpen] = useState(false);
  
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

  const [isRentModalOpen, setIsRentModalOpen] = useState(false);
  const [selectedRentSukukId, setSelectedRentSukukId] = useState<string>('');
  const [selectedRentInstIdx, setSelectedRentInstIdx] = useState<number>(0);
  const [rentFormData, setRentFormData] = useState({
    date: '',
    amount: '',
  });
  const [inheritedIncomeTxId, setInheritedIncomeTxId] = useState<string | undefined>(undefined);

  // Handle external trigger for adding or editing
  React.useEffect(() => {
    if (triggerAdd) {
      const isExternalTrigger = !!inheritedData;
      isFromExternalModuleRef.current = isExternalTrigger;
      returnModuleRef.current = inheritedData?.returnModule;

      const inheritedDate = inheritedData?.date || new Date().toISOString().split('T')[0];
      const inheritedAmt = (inheritedData?.amount !== undefined && inheritedData.amount !== null && inheritedData.amount !== 0) ? String(inheritedData.amount) : '';

      if (inheritedData?.type === 'Income' || inheritedData?.targetModule === 'sukuk-rent') {
        setInheritedIncomeTxId(inheritedData?.linkedTxId);
        const descLower = (inheritedData?.description || '').toLowerCase();

        // Find best matching Sukuk
        let matchedSukuk = investments.find(inv =>
          (inv.name && descLower.includes(inv.name.toLowerCase())) ||
          (inv.issuer && descLower.includes(inv.issuer.toLowerCase())) ||
          (inv.instrumentNo && descLower.includes(inv.instrumentNo.toLowerCase())) ||
          inv.installments?.some(i => i.linkedIncomeTxId === inheritedData.linkedTxId)
        );

        if (!matchedSukuk) {
          matchedSukuk = investments.find(s => s.installments?.some(i => !i.isPaid)) || investments[0];
        }

        if (matchedSukuk) {
          const sukukId = matchedSukuk.id;
          let instIdx = matchedSukuk.installments?.findIndex(i => i.linkedIncomeTxId === inheritedData.linkedTxId);
          if (instIdx === undefined || instIdx < 0) {
            instIdx = matchedSukuk.installments?.findIndex(i => !i.isPaid && (i.date === inheritedDate || i.date.substring(0, 7) === inheritedDate.substring(0, 7)));
          }
          if (instIdx === undefined || instIdx < 0) {
            instIdx = matchedSukuk.installments?.findIndex(i => !i.isPaid);
          }
          if (instIdx === undefined || instIdx < 0) {
            instIdx = 0;
          }

          setSelectedRentSukukId(sukukId);
          setSelectedRentInstIdx(instIdx);
          setRentFormData({
            date: inheritedDate,
            amount: inheritedAmt || (matchedSukuk.installments?.[instIdx] ? String(matchedSukuk.installments[instIdx].amount) : ''),
          });
          setIsRentModalOpen(true);
        }
      } else {
        setInheritedIeGroupId(isExternalTrigger ? inheritedData?.linkedTxId : undefined);

        // Check if there is an existing linked Sukuk (by linkedTxId)
        let existingLinkedSukuk: Sukuk | undefined;
        if (inheritedData?.linkedTxId) {
          existingLinkedSukuk = investments.find(inv => inv.linkedIeGroupId === inheritedData.linkedTxId);
        }

        if (existingLinkedSukuk) {
          // Edit mode: preserve existing fields and update with new amount/date
          setEditingId(existingLinkedSukuk.id);
          setFormData({
            name: existingLinkedSukuk.name,
            instrumentNo: existingLinkedSukuk.instrumentNo,
            issuer: existingLinkedSukuk.issuer,
            rentRate: existingLinkedSukuk.rentRate !== undefined ? existingLinkedSukuk.rentRate.toString() : '',
            tds: existingLinkedSukuk.tds !== undefined ? existingLinkedSukuk.tds.toString() : '10',
            frequency: existingLinkedSukuk.frequency || 'Semi-annual',
            principalAmount: (inheritedAmt && inheritedAmt !== '0') ? inheritedAmt : (existingLinkedSukuk.principalAmount !== undefined ? existingLinkedSukuk.principalAmount.toString() : ''),
            issueDate: inheritedDate || existingLinkedSukuk.issueDate,
            durationYears: existingLinkedSukuk.durationYears !== undefined ? existingLinkedSukuk.durationYears.toString() : '',
            status: existingLinkedSukuk.status || 'Active',
            closingDate: existingLinkedSukuk.closingDate || '',
            withdrawBalance: existingLinkedSukuk.withdrawBalance !== undefined
              ? existingLinkedSukuk.withdrawBalance.toString()
              : (existingLinkedSukuk.principalAmount !== undefined ? existingLinkedSukuk.principalAmount.toString() : ''),
          });
        } else {
          // New Sukuk creation
          setEditingId(null);
          setFormData({
            name: inheritedData?.description ? (inheritedData.description.startsWith('Sukuk:') ? inheritedData.description.replace(/^Sukuk:\s*/, '') : inheritedData.description) : '',
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
        }

        setIsModalOpen(true);
      }

      setTriggerAdd?.(false);
      onClearInheritedData?.();
    }
  }, [triggerAdd, setTriggerAdd, inheritedData, onClearInheritedData, investments]);

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
    const now = new Date();
    const { start: startDate, end: endDate } = rangeDates;

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
  }, [investments, rangeDates]);

  // [CHANGE 3] Notify parent whenever active investment changes
  React.useEffect(() => {
    onActiveInvestmentChange?.(historyStats.active);
  }, [historyStats.active, onActiveInvestmentChange]);

  // Update Title with Range
  React.useEffect(() => {
    if (onTitleChange) {
      const { start, end } = rangeDates;
      const startStrFormatted = start.toLocaleString('default', { month: 'short', year: 'numeric' });
      const endStrFormatted = end.toLocaleString('default', { month: 'short', year: 'numeric' });
      onTitleChange(
        <span className="flex items-center gap-2">
          SUKUK FUNDS <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">/ {startStrFormatted} - {endStrFormatted}</span>
        </span>
      );
    }
  }, [rangeDates, onTitleChange]);

  const handleToggleInstallment = (invId: string, index: number) => {
    const inv = investments.find(i => i.id === invId);
    if (!inv) return;

    const newInstallments = [...inv.installments];
    const currentlyPaid = !!newInstallments[index].isPaid;

    if (!currentlyPaid) {
      // Payment received: open Income & Expense New Ledger Entry modal with inherited date, category, amount, and narration
      const actualDate = newInstallments[index].actualDate || newInstallments[index].date;
      const actualAmount = newInstallments[index].actualAmount !== undefined && newInstallments[index].actualAmount !== null
        ? newInstallments[index].actualAmount
        : newInstallments[index].amount;
      const linkedIncomeTxId = newInstallments[index].linkedIncomeTxId || crypto.randomUUID();

      newInstallments[index] = { 
        ...newInstallments[index], 
        isPaid: true,
        actualDate,
        actualAmount,
        linkedIncomeTxId,
        isAutoMarked: false,
        isManuallyEdited: true
      };
      
      const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount !== undefined ? inst.actualAmount : inst.amount) : sum, 0);
      const newStatus = getEffectiveStatus({ ...inv, installments: newInstallments });
      
      onUpdate(invId, { 
        installments: newInstallments,
        totalRepaid,
        status: newStatus
      });

      const sukukLabel = inv.name || inv.issuer || 'Sukuk';
      const narration = `Sukuk Rent: ${sukukLabel} (Inst #${index + 1})`;

      if (onNavigateToModule) {
        onNavigateToModule('income-expense', {
          date: actualDate,
          amount: actualAmount,
          type: 'Income',
          category: 'Finance Income',
          subCategory: 'Sukuk Rent',
          description: narration,
          linkedTxId: linkedIncomeTxId,
          targetModule: 'sukuk-rent'
        });
      }
    } else {
      // Toggle unpaid & clean up linked income transaction
      const prevLinkedTxId = newInstallments[index].linkedIncomeTxId;
      newInstallments[index] = { 
        ...newInstallments[index], 
        isPaid: false,
        actualDate: undefined,
        actualAmount: undefined,
        linkedIncomeTxId: undefined,
        isAutoMarked: false,
        isManuallyEdited: false
      };
      
      const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount !== undefined ? inst.actualAmount : inst.amount) : sum, 0);
      const newStatus = getEffectiveStatus({ ...inv, installments: newInstallments });
      
      onUpdate(invId, { 
        installments: newInstallments,
        totalRepaid,
        status: newStatus
      });

      if (prevLinkedTxId && onDeleteLinkedIncomeTx) {
        onDeleteLinkedIncomeTx(prevLinkedTxId);
      }
    }
  };

  const handleSaveCustomRepayment = (invId: string, index: number) => {
    const inv = investments.find(i => i.id === invId);
    if (!inv) return;

    const actualAmount = parseFloat(customRepaymentData.amount) || 0;
    const actualDate = customRepaymentData.date;
    const linkedIncomeTxId = inv.installments[index]?.linkedIncomeTxId || crypto.randomUUID();

    const newInstallments = [...inv.installments];
    newInstallments[index] = { 
      ...newInstallments[index], 
      isPaid: true,
      actualDate,
      actualAmount,
      linkedIncomeTxId,
      isAutoMarked: false,
      isManuallyEdited: true
    };
    
    const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount !== undefined ? inst.actualAmount : inst.amount) : sum, 0);
    const newStatus = getEffectiveStatus({ ...inv, installments: newInstallments });
    
    onUpdate(invId, { 
      installments: newInstallments,
      totalRepaid,
      status: newStatus
    });

    const sukukLabel = inv.name || inv.issuer || 'Sukuk';
    const narration = `Sukuk Rent: ${sukukLabel} (Inst #${index + 1})`;

    if (onUpdateLinkedIncomeTx && linkedIncomeTxId) {
      onUpdateLinkedIncomeTx(linkedIncomeTxId, {
        date: actualDate,
        amount: actualAmount,
        description: narration
      });
    }

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

    const instCount = inv.installments?.length || 0;
    const paidCount = inv.installments?.filter(i => i.isPaid).length || 0;
    const hasDependents = instCount > 0 || !!inv.linkedIeGroupId;

    setConfirmState({
      isOpen: true,
      title: hasDependents ? 'Severe Warning: Deleting Main Investment & All Dependents' : 'Confirm Delete',
      message: hasDependents
        ? `You are about to delete the main Sukuk investment "${inv.name}". This will permanently delete the investment and ALL ${instCount} dependent rental installments (${paidCount} paid), historical profit records, and linked ledger entries.`
        : `Are you sure you want to delete the Sukuk investment "${inv.name}"?`,
      variant: hasDependents ? 'severe' : 'danger',
      confirmLabel: hasDependents ? 'Delete Investment & All Dependents' : 'Delete',
      details: hasDependents ? [
        `Main Sukuk Investment Record: "${inv.name}" (${inv.instrumentNo || 'N/A'})`,
        `${instCount} Scheduled / Paid Rental Installment Records (${paidCount} completed payments)`,
        inv.linkedIeGroupId ? 'Linked initial investment transfer in Income & Expense ledger' : null,
        inv.installments?.some(i => i.linkedIncomeTxId) ? 'Linked profit / rent income transactions in Income & Expense ledger' : null
      ] : undefined,
      onConfirm: () => {
        closeConfirm();
        onDelete(id);
        if (inv.linkedIeGroupId && onDeleteLinkedTransfer) {
          onDeleteLinkedTransfer(inv.linkedIeGroupId);
        }
        inv.installments?.forEach(inst => {
          if (inst.linkedIncomeTxId && onDeleteLinkedIncomeTx) {
            onDeleteLinkedIncomeTx(inst.linkedIncomeTxId);
          }
        });
        setActiveMenuId(null);
        showNotification('success', 'Investment and all dependent entries deleted successfully');
      },
      onCancel: closeConfirm,
    });
  };

  const handleBatchDelete = () => {
    if (selectedIds.length === 0) return;

    const selectedInvs = investments.filter(i => selectedIds.includes(i.id));
    const totalInstallments = selectedInvs.reduce((acc, i) => acc + (i.installments?.length || 0), 0);
    const totalPaid = selectedInvs.reduce((acc, i) => acc + (i.installments?.filter(inst => inst.isPaid).length || 0), 0);
    const hasDependents = totalInstallments > 0 || selectedInvs.some(i => i.linkedIeGroupId);

    setConfirmState({
      isOpen: true,
      title: hasDependents ? 'Severe Warning: Batch Deleting Main Investments & All Dependents' : 'Confirm Batch Delete',
      message: hasDependents
        ? `Your selection contains ${selectedIds.length} main Sukuk investment(s). Deleting will permanently delete all selected investments and ALL ${totalInstallments} dependent rental installments (${totalPaid} paid), repayment records, and linked ledger entries.`
        : `Are you sure you want to delete ${selectedIds.length} selected Sukuk investments?`,
      variant: hasDependents ? 'severe' : 'danger',
      confirmLabel: hasDependents ? `Delete All (${selectedIds.length} Items & Dependents)` : 'Delete Selected',
      details: hasDependents ? [
        `${selectedIds.length} Main Sukuk Investment Entries (${selectedInvs.map(i => i.name).slice(0, 3).join(', ')}${selectedInvs.length > 3 ? '...' : ''})`,
        `${totalInstallments} Dependent Rental Installment Records (${totalPaid} completed payouts)`,
        'All associated linked transfers and rent income entries in Income & Expense ledger'
      ] : undefined,
      onConfirm: () => {
        closeConfirm();
        investments.filter(i => selectedIds.includes(i.id)).forEach(i => {
          if (i.linkedIeGroupId && onDeleteLinkedTransfer) {
            onDeleteLinkedTransfer(i.linkedIeGroupId);
          }
          i.installments?.forEach(inst => {
            if (inst.linkedIncomeTxId && onDeleteLinkedIncomeTx) {
              onDeleteLinkedIncomeTx(inst.linkedIncomeTxId);
            }
          });
        });
        onBatchDelete(selectedIds);
        setSelectedIds([]);
        showNotification('success', `Successfully deleted ${selectedIds.length} investments and all dependent entries.`);
      },
      onCancel: closeConfirm,
    });
  };

  const [lastSelectedId, setLastSelectedId] = useState<string | null>(null);

  const toggleSelectAll = () => {
    if (selectedIds.length === filtered.length && filtered.length > 0) {
      setSelectedIds([]);
      setLastSelectedId(null);
    } else {
      setSelectedIds(filtered.map(i => i.id));
    }
  };

  const toggleSelect = (id: string, e?: React.MouseEvent) => {
    if (e?.shiftKey && lastSelectedId && lastSelectedId !== id) {
      const lastIdx = filtered.findIndex(i => i.id === lastSelectedId);
      const currIdx = filtered.findIndex(i => i.id === id);
      if (lastIdx !== -1 && currIdx !== -1) {
        const start = Math.min(lastIdx, currIdx);
        const end = Math.max(lastIdx, currIdx);
        const rangeIds = filtered.slice(start, end + 1).map(i => i.id);
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

    // When editing, preserve existing installments and totalRepaid so paid
    // history and Total Profit are never wiped. Only regenerate installments
    // for brand-new entries.
    const isExternalCreate = !editingId && isFromExternalModuleRef.current;
    const existingInv = editingId ? investments.find(i => i.id === editingId) : null;
    const linkGroupId: string | undefined = editingId
      ? (existingInv?.linkedIeGroupId || inheritedIeGroupId)
      : (isExternalCreate ? inheritedIeGroupId : crypto.randomUUID());

    const finalInstallments = existingInv
      ? existingInv.installments   // keep all paid/unpaid state intact
      : installments;              // fresh generation only for new entries

    const finalTotalRepaid = existingInv
      ? existingInv.totalRepaid    // never recalculate from scratch on edit
      : 0;

    const sukukDisplayName = formData.name ? formData.name : (formData.issuer ? `Sukuk: ${formData.issuer}` : 'Sukuk');

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
      linkedIeGroupId: linkGroupId,
    };

    if (formData.status === 'Matured') {
      payload.closingDate = formData.closingDate;
      const parsedWithdraw = parseFloat(formData.withdrawBalance);
      payload.withdrawBalance = isNaN(parsedWithdraw) ? principalAmount : parsedWithdraw;
    }

    if (editingId) {
      onUpdate(editingId, payload);
      if (linkGroupId && onUpdateLinkedTransfer) {
        onUpdateLinkedTransfer(linkGroupId, {
          date: payload.issueDate,
          amount: payload.principalAmount,
          description: sukukDisplayName
        });
      }
    } else {
      onAdd(payload);
      if (isExternalCreate) {
        isFromExternalModuleRef.current = false;
        setInheritedIeGroupId(undefined);
      } else if (onNavigateToModule) {
        onNavigateToModule('income-expense', {
          date: payload.issueDate,
          amount: payload.principalAmount,
          type: 'Transfer',
          targetModule: 'sukuk',
          description: sukukDisplayName,
          linkedTxId: linkGroupId
        });
      }
    }

    setIsModalOpen(false);
    setEditingId(null);
    checkAndReturnToModule();
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

  const handleRentSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const inv = investments.find(i => i.id === selectedRentSukukId);
    if (!inv) return;

    const actualAmount = parseFloat(rentFormData.amount) || 0;
    const actualDate = rentFormData.date;
    const linkedTxId = inheritedIncomeTxId || inv.installments[selectedRentInstIdx]?.linkedIncomeTxId || crypto.randomUUID();

    const newInstallments = [...inv.installments];
    if (newInstallments[selectedRentInstIdx]) {
      newInstallments[selectedRentInstIdx] = {
        ...newInstallments[selectedRentInstIdx],
        isPaid: true,
        actualDate,
        actualAmount,
        linkedIncomeTxId: linkedTxId,
        isAutoMarked: false,
        isManuallyEdited: true
      };
    }

    const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount !== undefined ? inst.actualAmount : inst.amount) : sum, 0);
    const newStatus = getEffectiveStatus({ ...inv, installments: newInstallments });

    onUpdate(inv.id, {
      installments: newInstallments,
      totalRepaid,
      status: newStatus
    });

    const sukukLabel = inv.name || inv.issuer || 'Sukuk';
    const narration = `Sukuk Rent: ${sukukLabel} (Inst #${selectedRentInstIdx + 1})`;

    if (isFromExternalModuleRef.current) {
      isFromExternalModuleRef.current = false;
      setInheritedIncomeTxId(undefined);
    } else {
      if (onNavigateToModule) {
        onNavigateToModule('income-expense', {
          date: actualDate,
          amount: actualAmount,
          type: 'Income',
          category: 'Finance Income',
          subCategory: 'Sukuk Rent',
          description: narration,
          linkedTxId,
          targetModule: 'sukuk-rent'
        });
      }
    }

    setIsRentModalOpen(false);
    showNotification('success', 'Sukuk rent recorded successfully!');
    checkAndReturnToModule();
  };

  return (
    <div className="space-y-8">

      {/* Date Range Selector — same UI/behavior as Mutual Funds */}
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
                    onChange={(_, e) => toggleSelect(inv.id, e)}
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
          isFromExternalModuleRef.current = false;
          setInheritedIeGroupId(undefined);
          checkAndReturnToModule();
        }} 
        title={editingId ? "Edit Sukuk Fund" : "Add Sukuk Fund"}
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
            <Button
              type="button"
              variant="secondary"
              className="flex-1 py-2"
              onClick={() => {
                setIsModalOpen(false);
                setEditingId(null);
                isFromExternalModuleRef.current = false;
                setInheritedIeGroupId(undefined);
                checkAndReturnToModule();
              }}
            >
              Cancel
            </Button>
            <Button type="submit" className="flex-1 py-2">
              {editingId ? "Update Sukuk" : "Create Sukuk"}
            </Button>
          </div>
        </form>
      </Modal>

      {/* Sukuk Rent Payment Confirmation Modal */}
      <Modal
        isOpen={isRentModalOpen}
        onClose={() => {
          setIsRentModalOpen(false);
          isFromExternalModuleRef.current = false;
          setInheritedIncomeTxId(undefined);
          checkAndReturnToModule();
        }}
        title="Confirm Sukuk Rent Payment"
      >
        <form onSubmit={handleRentSubmit} className="space-y-4">
          <Select
            label="Sukuk Investment"
            value={selectedRentSukukId}
            options={investments.map(inv => ({
              label: `${inv.name || inv.issuer || 'Sukuk'} (${inv.instrumentNo || 'No Inst. No'})`,
              value: inv.id
            }))}
            onChange={(val) => {
              setSelectedRentSukukId(val);
              const matched = investments.find(i => i.id === val);
              if (matched) {
                const firstUnpaid = matched.installments.findIndex(i => !i.isPaid);
                const nextIdx = firstUnpaid >= 0 ? firstUnpaid : 0;
                setSelectedRentInstIdx(nextIdx);
                if (matched.installments[nextIdx] && !rentFormData.amount) {
                  setRentFormData(prev => ({ ...prev, amount: String(matched.installments[nextIdx].amount) }));
                }
              }
            }}
            required
          />

          {selectedRentSukukId && (
            <Select
              label="Installment"
              value={String(selectedRentInstIdx)}
              options={(investments.find(i => i.id === selectedRentSukukId)?.installments || []).map((inst, idx) => ({
                label: `Inst #${idx + 1} - Due ${inst.date} (${inst.amount.toLocaleString()} BDT)${inst.isPaid ? ' [Paid]' : ''}`,
                value: String(idx)
              }))}
              onChange={(val) => {
                const idx = parseInt(val, 10);
                setSelectedRentInstIdx(idx);
                const matched = investments.find(i => i.id === selectedRentSukukId);
                if (matched?.installments[idx]) {
                  setRentFormData(prev => ({
                    ...prev,
                    amount: String(matched.installments[idx].actualAmount || matched.installments[idx].amount),
                    date: prev.date || matched.installments[idx].date
                  }));
                }
              }}
              required
            />
          )}

          <div className="grid grid-cols-2 gap-3">
            <Input
              label="Payment Date"
              type="date"
              value={rentFormData.date}
              onChange={(e) => setRentFormData({ ...rentFormData, date: e.target.value })}
              required
            />
            <Input
              label="Rent Amount (BDT)"
              type="number"
              step="0.01"
              value={rentFormData.amount}
              onChange={(e) => setRentFormData({ ...rentFormData, amount: e.target.value })}
              required
            />
          </div>

          <div className="flex justify-end gap-3 pt-4 border-t border-slate-800">
            <Button
              type="button"
              variant="secondary"
              onClick={() => {
                setIsRentModalOpen(false);
                isFromExternalModuleRef.current = false;
                setInheritedIncomeTxId(undefined);
                checkAndReturnToModule();
              }}
            >
              Cancel
            </Button>
            <Button type="submit">Confirm Rent</Button>
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
