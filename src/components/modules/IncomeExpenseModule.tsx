/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ComposedChart, Line, LabelList
} from 'recharts';
import { Card, Button, Modal, Input, Select, Checkbox } from '../ui/BaseComponents';
import { ConfirmDialog } from '../ui/ConfirmDialog';
import { BankAccount, IncomeExpenseTransaction, AppState, AppCategory, FDRTransaction } from '../../types';
import { formatBDT, formatDate, cn, formatCompactBDT } from '../../utils/formatters';
import { pushToSheet, markDirty } from '../../utils/sheetSync';
import { 
  Wallet, TrendingUp, PieChart, Landmark, ArrowUpRight, ArrowDownRight, 
  Trash2, Edit2, Plus, Calendar, ChevronLeft, ChevronRight, ChevronDown, 
  SlidersHorizontal, Download, Upload, Info, Search, Filter, GripVertical,
  Tag, CreditCard, X, Check, ArrowLeftRight, HandCoins,
  History, Sparkles, CornerDownLeft, ShoppingBag, Utensils, Car, Home,
  Zap, GraduationCap, Gift, Receipt, Briefcase, HeartPulse,
  Lightbulb, BookOpen, BadgePercent, Wrench, Scale, BarChart2,
  Settings, FileSpreadsheet
} from 'lucide-react';
import * as XLSX from 'xlsx';

// ── BO Account & Investment Module Helpers ─────────────────────────────────
export const isBoAccount = (acc?: BankAccount | null): boolean => {
  if (!acc) return false;
  if (acc.id === 'bdt-bo-account') return true;
  const name = (acc.name || '').trim().toLowerCase();
  return name === 'bo account' || name === 'bo' || name === 'bo a/c' || name.includes('bo account') || name.includes('bo a/c') || /\bbo\b/i.test(name);
};

export const isMutualFundAccount = (acc?: BankAccount | null): boolean => {
  if (!acc) return false;
  if (acc.id === 'bdt-mutual-fund') return true;
  const name = (acc.name || '').trim().toLowerCase();
  return name.includes('mutual fund') || name.includes('mutualfund') || name.includes('mutual-fund') || /\bmf\b/i.test(name);
};

export const isOnlineInvestmentAccount = (acc?: BankAccount | null): boolean => {
  if (!acc) return false;
  if (acc.id === 'bdt-online-investment') return true;
  const name = (acc.name || '').trim().toLowerCase();
  return name.includes('online investment') || name.includes('online invest') || name.includes('online-investment') || name.includes('onlineinvest');
};

export const isSukukAccount = (acc?: BankAccount | null): boolean => {
  if (!acc) return false;
  if (acc.id === 'bdt-sukuk') return true;
  const name = (acc.name || '').trim().toLowerCase();
  return name.includes('sukuk');
};

export const isFdrAccount = (acc?: BankAccount | null): boolean => {
  if (!acc) return false;
  if (acc.id === 'bdt-ibbl-fdr' || acc.id === 'usd-ibbl-fdr' || acc.id === 'bdt-fdr' || acc.id === 'bdt-fixed-deposit') return true;
  const name = (acc.name || '').trim().toLowerCase();
  return name.includes('fdr') || name.includes('fixed deposit') || name.includes('fixed-deposit') || /\bfd\b/i.test(name);
};

export const isInvestmentAccount = (acc?: BankAccount | null, allAccounts: BankAccount[] = []): boolean => {
  if (!acc) return false;
  if (isBoAccount(acc) || isMutualFundAccount(acc) || isOnlineInvestmentAccount(acc) || isSukukAccount(acc) || isFdrAccount(acc)) return true;
  const id = (acc.id || '').toLowerCase();
  const parentId = (acc.parentId || '').toLowerCase();
  const name = (acc.name || '').trim().toLowerCase();

  if (acc.parentId && allAccounts.length > 0) {
    const parent = allAccounts.find(a => a.id === acc.parentId);
    if (parent) {
      const pName = (parent.name || '').toLowerCase();
      const pId = (parent.id || '').toLowerCase();
      if (pName.includes('invest') || pId.includes('invest')) return true;
    }
  }

  return (
    id === 'bdt-investments' ||
    parentId === 'bdt-investments' ||
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
    name.includes('sukuk')
  );
};

export async function syncDseDepositTransaction({
  id,
  date,
  amount,
  notes,
  linkedIeGroupId,
  transactionType = 'Deposit',
  ticker = '',
  companyName = '',
  isDelete = false
}: {
  id: string;
  date: string;
  amount: number;
  notes: string;
  linkedIeGroupId?: string;
  transactionType?: 'Deposit' | 'Withdrawal' | 'Dividend' | 'Sell' | 'Buy' | 'Charge';
  ticker?: string;
  companyName?: string;
  isDelete?: boolean;
}) {
  try {
    const raw = localStorage.getItem('sheet_cache_dse');
    let currentDse: any[] = [];
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) currentDse = parsed;
    }

    let updatedDse: any[];
    const existingIndex = currentDse.findIndex((t: any) => t.id === id);

    if (isDelete) {
      if (existingIndex >= 0) {
        updatedDse = currentDse.filter((t: any) => t.id !== id);
        localStorage.setItem('sheet_cache_dse', JSON.stringify(updatedDse));
        window.dispatchEvent(new Event('dse-updated'));
        window.dispatchEvent(new Event('storage'));
        markDirty('dse');
        await pushToSheet('dse', 'delete', { id });
      }
      return;
    }

    const existingItem = existingIndex >= 0 ? currentDse[existingIndex] : null;

    const defaultNotes = transactionType === 'Withdrawal'
      ? 'Transfer from BO Account'
      : transactionType === 'Dividend'
        ? 'Cash Dividend'
        : transactionType === 'Sell'
          ? 'Stocks P&L'
          : transactionType === 'Charge'
            ? 'BO Account Charge'
            : 'Transfer to BO Account';

    const qty = (existingItem?.qty !== undefined && existingItem?.qty > 0) ? existingItem.qty : 1;

    let price: number;
    if (existingItem?.type === 'Sell' || existingItem?.type === 'Buy') {
      price = existingItem.price || amount;
    } else if (existingItem && existingItem.price !== undefined && existingItem.price > 0 && Math.abs((existingItem.total ?? (existingItem.qty * existingItem.price)) - amount) < 0.001) {
      price = existingItem.price;
    } else if (existingItem && existingItem.price !== undefined && existingItem.price > 0 && existingItem.qty > 0) {
      price = Number((amount / existingItem.qty).toFixed(2));
    } else {
      price = qty > 0 ? Number((amount / qty).toFixed(2)) : amount;
    }

    const dseTx = {
      id,
      date,
      type: transactionType,
      portfolio: existingItem?.portfolio || 'Investment',
      ticker: ticker || existingItem?.ticker || '',
      companyName: companyName || existingItem?.companyName || '',
      qty,
      price,
      commission: existingItem?.commission || 0,
      // For a Sell, "amount" arriving from Income & Expense is the realized P&L,
      // not the sale proceeds — never let it overwrite the DSE transaction's real total.
      total: (transactionType === 'Sell' && existingItem) ? existingItem.total : amount,
      notes: notes || existingItem?.notes || defaultNotes,
      linkedIeGroupId,
    };

    if (existingIndex >= 0) {
      updatedDse = currentDse.map((t: any, idx: number) => idx === existingIndex ? { ...t, ...dseTx } : t);
    } else {
      updatedDse = [...currentDse, dseTx];
    }

    localStorage.setItem('sheet_cache_dse', JSON.stringify(updatedDse));
    window.dispatchEvent(new Event('dse-updated'));
    window.dispatchEvent(new Event('storage'));
    markDirty('dse');
    await pushToSheet('dse', 'update', { data: dseTx });
  } catch (err) {
    console.error('Failed to sync transaction to DSE Tracker:', err);
  }
}

// ── Date Formatting for Groups ──────────────────────────────────────────
export function formatGroupDate(dateStr: string) {
  try {
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    const [y, m, d] = parts.map(Number);
    const dateObj = new Date(y, m - 1, d);
    const currentYear = new Date().getFullYear();
    
    return dateObj.toLocaleDateString('en-US', {
      weekday: 'long',
      month: 'long',
      day: 'numeric',
      ...(y !== currentYear ? { year: 'numeric' } : {})
    });
  } catch {
    return dateStr;
  }
}

// ── Visual Category Icon Resolver (Strictly Category-Based Only) ──────────
export function getCategoryVisual(category?: string, type?: string) {
  const cat = (category || '').trim().toLowerCase();

  // Transfers
  if (type === 'Transfer' || cat.includes('transfer')) {
    return {
      Icon: ArrowLeftRight,
      bg: 'bg-sky-500/20 text-sky-300 border-sky-500/30 ring-1 ring-sky-500/20'
    };
  }

  // Loans
  if (type === 'Loan' || cat.includes('loan') || cat.includes('borrow') || cat.includes('lend')) {
    return {
      Icon: HandCoins,
      bg: 'bg-amber-500/20 text-amber-300 border-amber-500/30 ring-1 ring-amber-500/20'
    };
  }

  // Initial Balance
  if (cat.includes('initial') || cat.includes('starting balance')) {
    return {
      Icon: Wallet,
      bg: 'bg-teal-500/20 text-teal-300 border-teal-500/30 ring-1 ring-teal-500/20'
    };
  }

  // Other Income (Book Royalty, Podcast Royalty, Other Income) - Single distinct Sparkles badge
  if (cat === 'other income' || cat.includes('other income') || cat.includes('royalty') || (type === 'Income' && cat.includes('other'))) {
    return {
      Icon: Sparkles,
      bg: 'bg-purple-500/20 text-purple-300 border-purple-500/30 ring-1 ring-purple-500/20'
    };
  }

  // Remittance Income (Salary, Incentive, Facebook Revenue, Freelance)
  if (cat.includes('remittance') || cat.includes('salary') || cat.includes('revenue') || cat.includes('freelance')) {
    return {
      Icon: Briefcase,
      bg: 'bg-teal-500/20 text-teal-300 border-teal-500/30 ring-1 ring-teal-500/20'
    };
  }

  // Finance Income (FDR Interest, Provisional Profit, Dividend, Sukuk Rent, Investments)
  if (cat.includes('finance') || cat.includes('dividend') || cat.includes('profit') || cat.includes('interest') || cat.includes('investment') || cat.includes('sukuk') || cat.includes('fdr')) {
    return {
      Icon: Landmark,
      bg: 'bg-cyan-500/20 text-cyan-300 border-cyan-500/30 ring-1 ring-cyan-500/20'
    };
  }

  // Accommodation (House / Rent / Housing)
  if (cat.includes('accommodation') || cat.includes('rent') || cat.includes('housing')) {
    return {
      Icon: Home,
      bg: 'bg-blue-500/20 text-blue-300 border-blue-500/30 ring-1 ring-blue-500/20'
    };
  }

  // Household & Utility / House Repair
  if (
    cat.includes('household') || cat.includes('utility') || cat.includes('repair') || cat.includes('maintenance')
  ) {
    return {
      Icon: Wrench,
      bg: 'bg-yellow-500/20 text-yellow-300 border-yellow-500/30 ring-1 ring-yellow-500/20'
    };
  }

  // Transportation (Vehicle, Fuel, Ride, Transport, Flight)
  if (
    cat.includes('transport') || cat.includes('travel') || cat.includes('fuel') || cat.includes('vehicle')
  ) {
    return {
      Icon: Car,
      bg: 'bg-orange-500/20 text-orange-300 border-orange-500/30 ring-1 ring-orange-500/20'
    };
  }

  // Food, Clothing & Essentials (Groceries, Dining, Market, Kitchen, Clothes, Apparel, Essentials) - Single consistent ShoppingBag icon
  if (
    cat.includes('food') || cat.includes('cloth') || cat.includes('grocer') || cat.includes('essential') ||
    cat.includes('dining') || cat.includes('restaurant') || cat.includes('bazaar') || cat.includes('bazar') || cat.includes('kitchen')
  ) {
    return {
      Icon: ShoppingBag,
      bg: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30 ring-1 ring-emerald-500/20'
    };
  }

  // Education (Tuition, Courses, School, Exam)
  if (cat.includes('education') || cat.includes('tuition') || cat.includes('school') || cat.includes('academic')) {
    return {
      Icon: GraduationCap,
      bg: 'bg-indigo-500/20 text-indigo-300 border-indigo-500/30 ring-1 ring-indigo-500/20'
    };
  }

  // Festival, Events & Special (Ceremony, Gifts, Tour, Holiday, Philanthropy, Sadakah)
  if (cat.includes('festival') || cat.includes('event') || cat.includes('special') || cat.includes('gift') || cat.includes('tour') || cat.includes('holiday') || cat.includes('philanthropy') || cat.includes('sadakah') || cat.includes('zakat')) {
    return {
      Icon: Gift,
      bg: 'bg-pink-500/20 text-pink-300 border-pink-500/30 ring-1 ring-pink-500/20'
    };
  }

  // Health / Medical / Healthcare
  if (cat.includes('health') || cat.includes('medical') || cat.includes('doctor') || cat.includes('hospital')) {
    return {
      Icon: HeartPulse,
      bg: 'bg-red-500/20 text-red-300 border-red-500/30 ring-1 ring-red-500/20'
    };
  }

  // Tax, Charges & TDS (TDS on Profit, Bank charges, Maintenance fees, Govt Tax, VAT, Service Fees)
  if (
    cat.includes('tax') || cat.includes('charge') || cat.includes('fee') || cat.includes('tds') ||
    cat.includes('vat') || cat.includes('duty')
  ) {
    return {
      Icon: Scale,
      bg: 'bg-rose-500/20 text-rose-300 border-rose-500/30 ring-1 ring-rose-500/20'
    };
  }

  // Other Expense (Miscellaneous, General & Uncategorized Expenses)
  if (
    cat === 'other expense' || cat.includes('other expense') || (type === 'Expense' && cat.includes('other')) || cat.includes('misc') || cat.includes('expense')
  ) {
    return {
      Icon: Receipt,
      bg: 'bg-fuchsia-500/20 text-fuchsia-300 border-fuchsia-500/30 ring-1 ring-fuchsia-500/20'
    };
  }

  // Fallbacks
  if (type === 'Income') {
    return {
      Icon: TrendingUp,
      bg: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30 ring-1 ring-emerald-500/20'
    };
  }

  if (type === 'Expense') {
    return {
      Icon: Tag,
      bg: 'bg-rose-500/20 text-rose-300 border-rose-500/30 ring-1 ring-rose-500/20'
    };
  }

  return {
    Icon: Tag,
    bg: 'bg-slate-800 text-slate-300 border-slate-700'
  };
}

// ── Account Icon & Color Resolver ───────────────────────────────────────
export function getAccountVisual(acc?: BankAccount | { id?: string; name?: string; currency?: string; parentId?: string; isParent?: boolean } | null) {
  if (!acc) {
    return {
      letter: '?',
      badgeClass: 'bg-slate-800 text-slate-300 border-slate-700 ring-1 ring-slate-700/50'
    };
  }

  const name = (acc.name || '').trim();
  const match = name.match(/[a-zA-Z0-9]/);
  const letter = (match ? match[0] : (name.charAt(0) || '?')).toUpperCase();
  const id = (acc.id || '').toLowerCase();
  const parentId = (acc.parentId || '').toLowerCase();
  const currency = acc.currency || 'BDT';
  const lowerName = name.toLowerCase();

  // 1. Investment accounts -> Pinkish color
  const isInvestment = id === 'bdt-investments' || parentId === 'bdt-investments' || 
    id.includes('investment') || id.includes('bo-account') || id.includes('mutual-fund') || id.includes('sukuk') ||
    lowerName === 'investments' || lowerName.includes('bo account') || lowerName.includes('online investment') || 
    lowerName.includes('mutual fund') || lowerName.includes('sukuk');

  if (isInvestment) {
    return {
      letter,
      badgeClass: 'bg-pink-500/20 text-pink-300 border-pink-500/35 ring-1 ring-pink-500/20'
    };
  }

  // 2. Wallet accounts -> Deep Earthy Coffee / Leather Brown (distinct from LYD)
  const isWallet = id === 'bdt-wallets' || parentId === 'bdt-wallets' || 
    id.includes('wallet') || id.includes('bkash') || id.includes('cellfin') || id.includes('mcash') ||
    lowerName === 'wallets' || lowerName.includes('bkash') || lowerName.includes('cellfin') || 
    lowerName.includes('mcash') || lowerName.includes('nagad') || lowerName.includes('rocket') || lowerName.includes('upay');

  if (isWallet) {
    return {
      letter,
      badgeClass: 'bg-[#3e2312]/90 text-[#fbd8b5] border-[#8b4f26] ring-1 ring-[#8b4f26]/30'
    };
  }

  // 3. Currency based colors:
  // USD -> Bluish
  if (currency === 'USD') {
    return {
      letter,
      badgeClass: 'bg-sky-500/20 text-sky-300 border-sky-500/35 ring-1 ring-sky-500/20'
    };
  }

  // LYD -> Vibrant Golden Amber (distinct from Coffee Brown of Wallets)
  if (currency === 'LYD') {
    return {
      letter,
      badgeClass: 'bg-amber-400/25 text-amber-300 border-amber-400/50 ring-1 ring-amber-400/25'
    };
  }

  // BDT (Default) -> Teal / Green
  return {
    letter,
    badgeClass: 'bg-teal-500/20 text-teal-300 border-teal-500/35 ring-1 ring-teal-500/20'
  };
}

export function AccountBadgeIcon({
  account,
  size = 'md',
  className = ''
}: {
  account?: BankAccount | { id?: string; name?: string; currency?: string; parentId?: string; isParent?: boolean } | null;
  size?: 'xs' | 'sm' | 'md' | 'lg';
  className?: string;
}) {
  const visual = getAccountVisual(account);

  const sizeClasses = {
    xs: 'w-4 h-4 text-[9px] rounded font-extrabold',
    sm: 'w-5 h-5 text-[10px] rounded-md font-extrabold',
    md: 'w-6 h-6 text-xs rounded-lg font-black',
    lg: 'w-8 h-8 text-sm rounded-lg font-black'
  }[size];

  return (
    <span
      className={cn(
        "inline-flex items-center justify-center shrink-0 border select-none leading-none shadow-sm",
        visual.badgeClass,
        sizeClasses,
        className
      )}
      title={account?.name}
    >
      {visual.letter}
    </span>
  );
}

// ── Search Highlight Helper ─────────────────────────────────────────────
function highlightMatch(text: string, query: string) {
  if (!query) return text;
  const lowerText = text.toLowerCase();
  const lowerQuery = query.toLowerCase();
  const index = lowerText.indexOf(lowerQuery);
  if (index === -1) return text;

  const before = text.slice(0, index);
  const match = text.slice(index, index + query.length);
  const after = text.slice(index + query.length);

  return (
    <>
      <span>{before}</span>
      <span className="text-teal-300 font-bold bg-teal-400/20 px-0.5 rounded">{match}</span>
      <span>{after}</span>
    </>
  );
}

// ── Categories & Subcategories Definitions ──────────────────────────────
export const DEFAULT_CATEGORIES: AppCategory[] = [
  // Income Categories
  {
    id: 'cat-remittance',
    name: 'Remittance Income',
    type: 'Income',
    subCategories: ['Salary', 'Incentive on Salary', 'Facebook Revenue']
  },
  {
    id: 'cat-finance-income',
    name: 'Finance Income',
    type: 'Income',
    subCategories: ['FDR Interest', 'Provisional Profit', 'Dividend', 'Sukuk Rent', 'Other Finance Income']
  },
  {
    id: 'cat-other-income',
    name: 'Other Income',
    type: 'Income',
    subCategories: ['Book Royalty', 'Podcast Royalty', 'Investment Interest', 'Other Income']
  },
  {
    id: 'cat-capital-gain',
    name: 'Capital Gain',
    type: 'Income',
    subCategories: ['Stocks Capital Gain']
  },
  {
    id: 'cat-loan-income',
    name: 'Loan',
    type: 'Income',
    subCategories: ['Borrowed', 'Received Lent Money']
  },
  // Expense Categories
  {
    id: 'cat-food-clothing',
    name: 'Food, Clothing & Essentials',
    type: 'Expense',
    subCategories: []
  },
  {
    id: 'cat-accommodation',
    name: 'Accommodation',
    type: 'Expense',
    subCategories: []
  },
  {
    id: 'cat-transportation',
    name: 'Transportation',
    type: 'Expense',
    subCategories: []
  },
  {
    id: 'cat-household-utility',
    name: 'Household & Utility',
    type: 'Expense',
    subCategories: []
  },
  {
    id: 'cat-education',
    name: 'Education',
    type: 'Expense',
    subCategories: []
  },
  {
    id: 'cat-festival-special',
    name: 'Festival, Events & Special',
    type: 'Expense',
    subCategories: ['Festival, Ceremony & Events', 'Tour & Holiday', 'Philanthropy', 'Other Special Expenses']
  },
  {
    id: 'cat-other-expense',
    name: 'Other Expense',
    type: 'Expense',
    subCategories: ['TDS on Provisional Profit', 'Account Maintenance Fees', 'Other Bank Charges', 'Account Related Fees']
  },
  {
    id: 'cat-loan-expense',
    name: 'Loan',
    type: 'Expense',
    subCategories: ['Lent', 'Repaid Borrowed Money']
  }
];

// ── Palette (DSE Tracker Theme) ──────────────────────────────────────────
const HOLDING_COLORS = [
  '#2dd4bf','#3b82f6','#a855f7','#f59e0b','#ef4444',
  '#ec4899','#06b6d4','#84cc16','#f97316','#8b5cf6',
  '#14b8a6','#eab308','#6366f1','#10b981','#fb923c',
  '#e879f9','#38bdf8','#a3e635','#fb7185','#34d399',
  '#60a5fa','#c084fc','#fbbf24','#f472b6','#4ade80',
];

interface AnalyticsDonutSlice {
  id: string;
  name: string;
  value: number;
  pct: number;
  color: string;
}

// ── SVG donut helpers (pure SVG matching DSE Tracker) ──────────────────────────
function toXY(cx: number, cy: number, r: number, angleDeg: number) {
  const rad = (angleDeg * Math.PI) / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

function buildArcPath(
  cx: number, cy: number,
  outerR: number, innerR: number,
  startDeg: number, sweepDeg: number,
): string {
  const sw = Math.min(Math.max(sweepDeg, 0.01), 359.9999);
  const end = startDeg - sw;
  const lg = sw > 180 ? 1 : 0;

  const os = toXY(cx, cy, outerR, startDeg);
  const oe = toXY(cx, cy, outerR, end);
  const ie = toXY(cx, cy, innerR, end);
  const is = toXY(cx, cy, innerR, startDeg);

  return (
    `M ${os.x.toFixed(3)} ${os.y.toFixed(3)} ` +
    `A ${outerR} ${outerR} 0 ${lg} 0 ${oe.x.toFixed(3)} ${oe.y.toFixed(3)} ` +
    `L ${ie.x.toFixed(3)} ${ie.y.toFixed(3)} ` +
    `A ${innerR} ${innerR} 0 ${lg} 1 ${is.x.toFixed(3)} ${is.y.toFixed(3)} Z`
  );
}

const AnalyticsDonutChart: React.FC<{
  slices: AnalyticsDonutSlice[];
  excluded: Set<string>;
  size?: number;
  centerLabel?: string;
  itemNoun?: string;
  totalCount?: number;
}> = ({ slices, excluded, size = 260, centerLabel = 'TOTAL', itemNoun = 'ITEMS', totalCount }) => {
  const [hov, setHov] = React.useState<number | null>(null);
  const CX = 150, CY = 150, OUTER = 126, INNER = 68, GAP = 1.4;

  const visibleSlices = React.useMemo(() => {
    const total = slices.filter(s => !excluded.has(s.id)).reduce((sum, s) => sum + s.value, 0);
    if (total === 0) return [];
    return slices
      .filter(s => !excluded.has(s.id))
      .map(s => ({ ...s, adjPct: s.value / total }));
  }, [slices, excluded]);

  const arcs = React.useMemo(() => {
    let cursor = 0;
    return visibleSlices.map(s => {
      const totalSweep = s.adjPct * 360;
      const start = cursor;
      const sweep = Math.max(0, totalSweep - GAP);
      cursor -= totalSweep;
      return { ...s, start, sweep };
    });
  }, [visibleSlices]);

  const hovSlice = hov !== null ? arcs[hov] : null;

  return (
    <svg
      viewBox="0 0 300 300"
      style={{ width: size, height: size, flexShrink: 0, overflow: 'visible' }}
    >
      {arcs.map((arc, i) => {
        if (arc.sweep <= 0) return null;
        const isHov = hov === i;
        const dimmed = hov !== null && !isHov;
        return (
          <path
            key={arc.id}
            d={buildArcPath(CX, CY, OUTER, INNER, arc.start, arc.sweep)}
            fill={arc.color}
            opacity={dimmed ? 0.28 : 1}
            style={{
              transition: 'opacity 0.5s ease, transform 0.5s ease',
              cursor: 'pointer',
              transformOrigin: `${CX}px ${CY}px`,
              transform: isHov ? 'scale(1.06)' : 'scale(1)',
            }}
            onMouseEnter={() => setHov(i)}
            onMouseLeave={() => setHov(null)}
          />
        );
      })}

      {arcs.length === 0 && (
        <circle cx={CX} cy={CY} r={OUTER} fill="none" stroke="#1e293b" strokeWidth={OUTER - INNER} />
      )}

      {hovSlice ? (
        <>
          <text x={CX} y={CY - 14} textAnchor="middle" fill="#e2e8f0"
            fontSize={12} fontWeight={700} letterSpacing={0.5}>
            {hovSlice.name.length > 18 ? hovSlice.name.slice(0, 17) + '…' : hovSlice.name}
          </text>
          <text x={CX} y={CY + 10} textAnchor="middle" fill={hovSlice.color}
            fontSize={22} fontWeight={800}>{(hovSlice.adjPct * 100).toFixed(1)}%</text>
          <text x={CX} y={CY + 28} textAnchor="middle" fill="#475569"
            fontSize={9} fontWeight={600}>{formatBDT(hovSlice.value)}</text>
        </>
      ) : (
        <>
          <text x={CX} y={CY - 8} textAnchor="middle" fill="#64748b"
            fontSize={10} fontWeight={700} letterSpacing={0.5}>{centerLabel}</text>
          <text x={CX} y={CY + 14} textAnchor="middle" fill="#e2e8f0"
            fontSize={16} fontWeight={800}>{totalCount ?? visibleSlices.length} {itemNoun}</text>
        </>
      )}
    </svg>
  );
};

const AnalyticsPieCard: React.FC<{
  title: string;
  icon?: React.ReactNode;
  emptyText: string;
  slices: AnalyticsDonutSlice[];
  excluded: Set<string>;
  onToggleExcluded: (id: string) => void;
  centerLabel: string;
  itemNoun: string;
  legendPosition?: 'side' | 'bottom';
}> = ({ title, icon, emptyText, slices, excluded, onToggleExcluded, centerLabel, itemNoun, legendPosition = 'side' }) => {
  // If more than maxItems, show first few in columns and group the rest into "Others"
  // For bottom legend (Asset by Accounts), allow 8 items + Others (up to 9 items for 3 rows of 3)
  const { displaySlices, otherSliceIds, isOthersExcluded } = React.useMemo(() => {
    const maxItems = legendPosition === 'bottom' ? 8 : 5;
    if (slices.length <= maxItems + 1) {
      return {
        displaySlices: slices,
        otherSliceIds: [] as string[],
        isOthersExcluded: false,
      };
    }

    const firstFew = slices.slice(0, maxItems);
    const remaining = slices.slice(maxItems);
    const otherIds = remaining.map(s => s.id);
    const allOthersExcluded = remaining.length > 0 && remaining.every(s => excluded.has(s.id));

    const othersVal = remaining.reduce((sum, s) => sum + s.value, 0);
    const othersPct = remaining.reduce((sum, s) => sum + s.pct, 0);

    const othersSlice: AnalyticsDonutSlice = {
      id: '__others__',
      name: 'Others',
      value: othersVal,
      pct: othersPct,
      color: '#94a3b8',
    };

    return {
      displaySlices: [...firstFew, othersSlice],
      otherSliceIds: otherIds,
      isOthersExcluded: allOthersExcluded,
    };
  }, [slices, excluded, legendPosition]);

  // Handle toggling "Others" or regular slices
  const handleToggle = (id: string) => {
    if (id === '__others__') {
      if (isOthersExcluded) {
        otherSliceIds.forEach(remId => {
          if (excluded.has(remId)) {
            onToggleExcluded(remId);
          }
        });
      } else {
        otherSliceIds.forEach(remId => {
          if (!excluded.has(remId)) {
            onToggleExcluded(remId);
          }
        });
      }
    } else {
      onToggleExcluded(id);
    }
  };

  // Excluded set to pass to Donut Chart
  const effectiveExcluded = React.useMemo(() => {
    const set = new Set(excluded);
    if (isOthersExcluded) {
      set.add('__others__');
    }
    return set;
  }, [excluded, isOthersExcluded]);

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 sm:p-6 shadow-xl h-full flex flex-col">
      <div className="flex flex-col flex-1 min-h-0">
        <div className="flex items-center gap-2 mb-5 pb-3 border-b border-slate-800">
          {icon || <PieChart size={20} className="text-teal-400" />}
          <h2 className="text-subheading font-bold text-white uppercase tracking-wider">{title}</h2>
        </div>

        {slices.length === 0 ? (
          <div className="flex items-center justify-center py-20 text-slate-600 text-[11px] font-bold uppercase tracking-wider">
            {emptyText}
          </div>
        ) : legendPosition === 'bottom' ? (
          <div className="flex flex-col items-center w-full flex-1 justify-center">
            {/* Centered Donut Chart */}
            <div className="flex justify-center shrink-0 mb-5">
              <AnalyticsDonutChart
                slices={displaySlices}
                excluded={effectiveExcluded}
                size={260}
                centerLabel={centerLabel}
                itemNoun={itemNoun}
                totalCount={slices.length}
              />
            </div>

            {/* Legends below the pie - Horizontally in 2 rows */}
            <div className="w-full pt-4 border-t border-slate-800/60">
              <div className={cn(
                "grid gap-x-4 gap-y-2.5 w-full",
                displaySlices.length <= 4 ? "grid-cols-2" : "grid-cols-2 sm:grid-cols-3"
              )}>
                {displaySlices.map(s => {
                  const isOthers = s.id === '__others__';
                  const isExcluded = isOthers ? isOthersExcluded : excluded.has(s.id);
                  return (
                    <button
                      key={s.id}
                      type="button"
                      onClick={() => handleToggle(s.id)}
                      title={`${s.name}: ${formatBDT(s.value)}`}
                      className="flex items-center justify-between gap-2 py-1.5 px-2.5 rounded-lg transition-all hover:bg-slate-800/60 active:scale-[0.98] text-left w-full group bg-slate-950/40 border border-slate-800/60"
                      style={{ opacity: isExcluded ? 0.35 : 1 }}
                    >
                      <div className="flex items-center gap-2 min-w-0 flex-1">
                        <div 
                          className="w-2.5 h-2.5 rounded-sm shrink-0" 
                          style={{ 
                            backgroundColor: s.color, 
                            boxShadow: isExcluded ? 'none' : `0 0 8px ${s.color}55` 
                          }} 
                        />
                        <span className={cn(
                          'text-[11px] font-bold text-slate-300 truncate tracking-wide transition-all group-hover:text-white',
                          isExcluded && 'line-through text-slate-600'
                        )}>
                          {s.name}&nbsp;
                          <span style={{ color: isExcluded ? undefined : s.color }}>
                            ({(s.pct * 100).toFixed(1)}%)
                          </span>
                        </span>
                      </div>
                      <span className="text-white text-[11px] font-extrabold tabular-nums shrink-0 ml-1">
                        {formatCompactBDT(s.value)}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        ) : (
          <div className="flex flex-col sm:flex-row items-center gap-6 my-auto">
            {/* Donut Chart */}
            <div className="flex justify-center shrink-0">
              <AnalyticsDonutChart
                slices={displaySlices}
                excluded={effectiveExcluded}
                size={200}
                centerLabel={centerLabel}
                itemNoun={itemNoun}
                totalCount={slices.length}
              />
            </div>

            {/* Legend - Single Column */}
            <div className="w-full flex-1 min-w-0">
              <div className="flex flex-col space-y-1.5 w-full">
                {displaySlices.map(s => {
                  const isOthers = s.id === '__others__';
                  const isExcluded = isOthers ? isOthersExcluded : excluded.has(s.id);
                  return (
                    <button
                      key={s.id}
                      type="button"
                      onClick={() => handleToggle(s.id)}
                      title={`${s.name}: ${formatBDT(s.value)}`}
                      className="flex items-center justify-between gap-2 py-1 px-2 rounded-lg transition-all hover:bg-slate-800/60 active:scale-[0.98] text-left w-full group"
                      style={{ opacity: isExcluded ? 0.35 : 1 }}
                    >
                      <div className="flex items-center gap-2 min-w-0 flex-1">
                        <div 
                          className="w-2.5 h-2.5 rounded-sm shrink-0" 
                          style={{ 
                            backgroundColor: s.color, 
                            boxShadow: isExcluded ? 'none' : `0 0 8px ${s.color}55` 
                          }} 
                        />
                        <span className={cn(
                          'text-[11px] font-bold text-slate-300 truncate tracking-wide transition-all group-hover:text-white',
                          isExcluded && 'line-through text-slate-600'
                        )}>
                          {s.name}&nbsp;
                          <span style={{ color: isExcluded ? undefined : s.color }}>
                            ({(s.pct * 100).toFixed(1)}%)
                          </span>
                        </span>
                      </div>
                      <span className="text-white text-[11px] font-extrabold tabular-nums shrink-0 ml-2">
                        {formatCompactBDT(s.value)}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

interface IncomeExpenseModuleProps {
  state: AppState;
  updateState: (updater: (prev: AppState) => AppState, dirtyModules?: any[]) => void;
  activeTab: 'summary' | 'transactions' | 'analytics' | 'settings';
  setActiveTab: (tab: 'summary' | 'transactions' | 'analytics' | 'settings') => void;
  triggerAdd?: boolean;
  setTriggerAdd?: (val: boolean) => void;
  onNavigateToModule?: (module: string, initialAddData?: { date: string; amount: number; type?: 'Transfer' | 'Income' | 'Expense' | 'Loan'; category?: string; subCategory?: string; targetModule?: string; description?: string; linkedTxId?: string; accountId?: string; returnModule?: string; [key: string]: any }, openModal?: boolean) => void;
  inheritedData?: { date: string; amount: number; type?: 'Transfer' | 'Income' | 'Expense' | 'Loan'; category?: string; subCategory?: string; targetModule?: string; description?: string; linkedTxId?: string; accountId?: string; returnModule?: string; [key: string]: any } | null;
  onClearInheritedData?: () => void;
}

// ── Category Picker Modal ────────────────────────────────────────────────
interface CategoryPickerModalProps {
  isOpen: boolean;
  onClose: () => void;
  categories: AppCategory[];
  txType: 'Income' | 'Expense' | 'Transfer' | 'Loan';
  selectedCategory: string;
  selectedSubCategory: string;
  onSelect: (category: string, subCategory: string) => void;
}

function CategoryPickerModal({
  isOpen,
  onClose,
  categories,
  txType,
  selectedCategory,
  selectedSubCategory,
  onSelect
}: CategoryPickerModalProps) {
  const [expandedCats, setExpandedCats] = useState<Record<string, boolean>>({});
  const [hoveredCat, setHoveredCat] = useState<string | null>(null);

  // Determine which categories to show
  const relevantType = (txType === 'Transfer' || txType === 'Loan') ? null : txType;
  const filteredCats = (relevantType 
    ? categories.filter(c => c.type === relevantType)
    : categories).filter(c => c.name !== 'Loan');

  const incomeCats = filteredCats.filter(c => c.type === 'Income');
  const expenseCats = filteredCats.filter(c => c.type === 'Expense');

  const handleSelectCat = (catName: string, subCat: string = '') => {
    onSelect(catName, subCat);
    onClose();
  };

  const renderCatGroup = (cats: AppCategory[], groupType: 'Income' | 'Expense') => {
    if (cats.length === 0) return null;
    return (
      <div className="mb-4">
        <div className={cn(
          "text-[9px] font-extrabold uppercase tracking-widest px-1 mb-2",
          groupType === 'Income' ? "text-emerald-500" : "text-rose-500"
        )}>
          {groupType}
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
          {cats.map(cat => {
            const isSelected = selectedCategory === cat.name;
            const isExpanded = !!expandedCats[cat.id];
            const hasSubcats = cat.subCategories.length > 0;
            const catVisual = getCategoryVisual(cat.name, cat.type);

            return (
              <div key={cat.id} className="flex flex-col">
                <button
                  type="button"
                  onClick={() => {
                    if (hasSubcats) {
                      setExpandedCats(prev => ({ ...prev, [cat.id]: !prev[cat.id] }));
                    } else {
                      handleSelectCat(cat.name, '');
                    }
                  }}
                  className={cn(
                    "w-full text-left p-2.5 rounded-lg border transition-all flex flex-col gap-1 min-h-[54px]",
                    isSelected
                      ? groupType === 'Income'
                        ? "bg-emerald-400/15 border-emerald-400/50 shadow-[0_0_10px_rgba(52,211,153,0.1)]"
                        : "bg-rose-400/15 border-rose-400/50 shadow-[0_0_10px_rgba(251,113,133,0.1)]"
                      : "bg-slate-950 border-slate-800 hover:bg-slate-900",
                    groupType === 'Income'
                      ? "hover:border-emerald-400/30"
                      : "hover:border-rose-400/30"
                  )}
                >
                  <div className="flex items-start justify-between gap-1">
                    <div className="flex items-center gap-1.5 min-w-0">
                      <div className={cn("w-4 h-4 rounded flex items-center justify-center shrink-0 border", catVisual.bg)}>
                        <catVisual.Icon size={10} />
                      </div>
                      <span className={cn(
                        "font-semibold text-[11px] leading-tight truncate",
                        isSelected
                          ? groupType === 'Income' ? "text-emerald-300" : "text-rose-300"
                          : "text-slate-200"
                      )}>
                        {cat.name}
                      </span>
                    </div>
                    {isSelected && !hasSubcats && (
                      <Check size={10} className={groupType === 'Income' ? "text-emerald-400 shrink-0 mt-0.5" : "text-rose-400 shrink-0 mt-0.5"} />
                    )}
                    {hasSubcats && (
                      <ChevronDown
                        size={10}
                        className={cn(
                          "text-slate-500 shrink-0 mt-0.5 transition-transform",
                          isExpanded ? "rotate-180" : ""
                        )}
                      />
                    )}
                  </div>
                  {hasSubcats && (
                    <span className="text-[9px] text-slate-500 font-semibold pl-5.5">
                      {cat.subCategories.length} sub-ledgers
                    </span>
                  )}
                </button>

                {/* Sub-category expanded pills */}
                {hasSubcats && isExpanded && (
                  <div className="mt-1.5 pl-2 flex flex-col gap-1">
                    {cat.subCategories.map(sub => {
                      const isSubSelected = selectedCategory === cat.name && selectedSubCategory === sub;
                      return (
                        <button
                          key={sub}
                          type="button"
                          onClick={() => handleSelectCat(cat.name, sub)}
                          className={cn(
                            "text-left px-2.5 py-1.5 rounded-md border text-[10px] font-semibold transition-all flex items-center justify-between gap-2",
                            isSubSelected
                              ? groupType === 'Income'
                                ? "bg-emerald-400/15 border-emerald-400/40 text-emerald-300"
                                : "bg-rose-400/15 border-rose-400/40 text-rose-300"
                              : "bg-slate-950/60 border-slate-800/60 text-slate-400 hover:text-slate-200 hover:border-slate-700"
                          )}
                        >
                          <span className="truncate">{sub}</span>
                          {isSubSelected && <Check size={9} className="shrink-0" />}
                        </button>
                      );
                    })}
                    {/* Also allow selecting parent category directly */}
                    <button
                      type="button"
                      onClick={() => handleSelectCat(cat.name, '')}
                      className={cn(
                        "text-left px-2.5 py-1.5 rounded-md border text-[10px] font-semibold transition-all flex items-center justify-between gap-2",
                        selectedCategory === cat.name && selectedSubCategory === ''
                          ? groupType === 'Income'
                            ? "bg-emerald-400/15 border-emerald-400/40 text-emerald-300"
                            : "bg-rose-400/15 border-rose-400/40 text-rose-300"
                          : "bg-slate-950/60 border-slate-800/60 text-slate-500 hover:text-slate-300 hover:border-slate-700 italic"
                      )}
                    >
                      <span>{cat.name} (general)</span>
                      {selectedCategory === cat.name && selectedSubCategory === '' && <Check size={9} className="shrink-0" />}
                    </button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  return (
    <Modal isOpen={isOpen} onClose={onClose} title="Select Category">
      <div className="max-h-[65vh] overflow-y-auto pr-1 space-y-2">
        {txType === 'Loan' ? (
          <div className="mb-4">
            <div className="text-[9px] font-extrabold uppercase tracking-widest px-1 mb-2 text-amber-500">
              Loan
            </div>
            <div className="grid grid-cols-1 gap-2">
              <div className="flex flex-col">
                <button
                  type="button"
                  disabled
                  className="w-full text-left p-2.5 rounded-lg border transition-all flex flex-col gap-1 min-h-[54px] bg-amber-400/10 border-amber-500/25 cursor-default"
                >
                  <div className="flex items-start justify-between gap-1 w-full">
                    <span className="font-bold text-[11px] leading-tight text-amber-400">
                      Loan
                    </span>
                    <span className="text-[9px] text-amber-500 font-semibold">
                      4 sub-ledgers
                    </span>
                  </div>
                </button>

                <div className="mt-1.5 pl-2 flex flex-col gap-1">
                  {[
                    { name: 'Borrowed', isGreen: true },
                    { name: 'Lent', isGreen: false },
                    { name: 'Received Lent Money', isGreen: true },
                    { name: 'Repaid Borrowed Money', isGreen: false }
                  ].map(sub => {
                    const isSubSelected = selectedCategory === 'Loan' && selectedSubCategory === sub.name;
                    return (
                      <button
                        key={sub.name}
                        type="button"
                        onClick={() => handleSelectCat('Loan', sub.name)}
                        className={cn(
                          "text-left px-2.5 py-1.5 rounded-md border text-[10px] font-semibold transition-all flex items-center justify-between gap-2",
                          isSubSelected
                            ? "bg-amber-400/15 border-amber-400/40 text-amber-300 shadow-[0_0_8px_rgba(245,158,11,0.12)]"
                            : "bg-slate-950/60 border-slate-800/60 text-slate-400 hover:text-slate-200 hover:border-slate-700"
                        )}
                      >
                        <span className="flex items-center gap-1.5">
                          <span className={cn(
                            "w-1.5 h-1.5 rounded-full",
                            sub.isGreen ? "bg-emerald-400" : "bg-rose-400"
                          )} />
                          {sub.name}
                        </span>
                        {isSubSelected && <Check size={9} className="shrink-0 text-amber-400" />}
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        ) : (
          <>
            {renderCatGroup(incomeCats, 'Income')}
            {renderCatGroup(expenseCats, 'Expense')}
            {filteredCats.length === 0 && (
              <div className="py-12 text-center text-slate-500 italic uppercase text-label">
                No categories available
              </div>
            )}
          </>
        )}
      </div>
    </Modal>
  );
}

// ── Account Picker Modal ────────────────────────────────────────────────
interface AccountPickerModalProps {
  isOpen: boolean;
  onClose: () => void;
  accounts: BankAccount[];
  accountBalances: Record<string, number>;
  selectedAccountId: string;
  onSelect: (accountId: string) => void;
  formatCurrency: (amount: number, currency: string) => string;
}

function AccountPickerModal({
  isOpen,
  onClose,
  accounts,
  accountBalances,
  selectedAccountId,
  onSelect,
  formatCurrency
}: AccountPickerModalProps) {
  const [expandedGroups, setExpandedGroups] = useState<Record<string, boolean>>({});

  const leafAccounts = accounts.filter(a => !a.isParent);
  const parentAccounts = accounts.filter(a => a.isParent);
  const standaloneAccounts = leafAccounts.filter(a => !a.parentId);

  const handleSelect = (accountId: string) => {
    onSelect(accountId);
    onClose();
  };

  const renderAccountCard = (acc: BankAccount, isChild: boolean = false) => {
    const balance = accountBalances[acc.id] || 0;
    const isSelected = selectedAccountId === acc.id;

    return (
      <button
        key={acc.id}
        type="button"
        onClick={() => handleSelect(acc.id)}
        className={cn(
          "w-full text-left p-2.5 rounded-lg border transition-all flex flex-col justify-between min-h-[54px]",
          isSelected
            ? acc.currency === 'BDT'
              ? "bg-teal-400/15 border-teal-400/50 shadow-[0_0_10px_rgba(45,212,191,0.1)]"
              : acc.currency === 'LYD'
                ? "bg-amber-400/15 border-amber-400/50 shadow-[0_0_10px_rgba(245,158,11,0.1)]"
                : "bg-sky-400/15 border-sky-400/50 shadow-[0_0_10px_rgba(56,189,248,0.1)]"
            : isChild
              ? "bg-slate-950/60 border-slate-800/60 hover:bg-slate-900 hover:border-slate-700"
              : "bg-slate-950 border-slate-800 hover:bg-slate-900",
          acc.currency === 'BDT' ? "hover:border-teal-400/30" : acc.currency === 'LYD' ? "hover:border-amber-400/30" : "hover:border-sky-400/30"
        )}
      >
        <div className="flex items-start justify-between gap-1">
          <div className="flex items-center gap-1.5 min-w-0 pr-1">
            <AccountBadgeIcon account={acc} size="xs" />
            <span className={cn(
              "font-semibold text-[11px] leading-tight truncate",
              isSelected
                ? acc.currency === 'BDT' ? "text-teal-300" : acc.currency === 'LYD' ? "text-amber-300" : "text-sky-300"
                : "text-slate-200"
            )}>
              {acc.name}
            </span>
          </div>
          <div className="flex items-center gap-1 shrink-0">
            <span className={cn(
              "text-[8px] font-extrabold px-1.5 py-0.5 rounded border",
              acc.currency === 'BDT' ? "bg-teal-400/10 text-teal-400 border-teal-400/20" : acc.currency === 'LYD' ? "bg-amber-400/10 text-amber-400 border-amber-400/20" : "bg-sky-400/10 text-sky-400 border-sky-400/20"
            )}>
              {acc.currency}
            </span>
            {isSelected && <Check size={10} className={acc.currency === 'BDT' ? "text-teal-400" : acc.currency === 'LYD' ? "text-amber-400" : "text-sky-400"} />}
          </div>
        </div>
        <div className={cn(
          "font-bold text-[11px] tabular-nums mt-1",
          acc.currency === 'BDT' ? "text-teal-400" : acc.currency === 'USD' ? "text-sky-400" : "text-amber-400"
        )}>
          {formatCurrency(balance, acc.currency)}
        </div>
      </button>
    );
  };

  return (
    <Modal isOpen={isOpen} onClose={onClose} title="Select Account">
      <div className="max-h-[65vh] overflow-y-auto pr-1 space-y-4">

        {/* Standalone accounts */}
        {standaloneAccounts.length > 0 && (
          <div>
            <div className="text-[9px] font-extrabold uppercase tracking-widest text-slate-500 px-1 mb-2">
              Accounts
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
              {standaloneAccounts.map(acc => renderAccountCard(acc))}
            </div>
          </div>
        )}

        {/* Parent groups */}
        {parentAccounts.map(parent => {
          const children = accounts.filter(a => a.parentId === parent.id);
          const isExpanded = !!expandedGroups[parent.id];
          const parentBalance = accountBalances[parent.id] || 0;

          return (
            <div key={parent.id}>
              <button
                type="button"
                onClick={() => setExpandedGroups(prev => ({ ...prev, [parent.id]: !prev[parent.id] }))}
                className="w-full flex items-center justify-between px-1 mb-2 group"
              >
                <div className="text-[9px] font-extrabold uppercase tracking-widest text-slate-400 group-hover:text-slate-200 transition-colors flex items-center gap-2">
                  <AccountBadgeIcon account={parent} size="xs" />
                  <span>{parent.name}</span>
                  <span className={cn(
                    "text-[8px] font-bold px-1.5 py-0.5 rounded border",
                    parent.currency === 'BDT' ? "bg-teal-400/10 text-teal-400 border-teal-400/20" : parent.currency === 'LYD' ? "bg-amber-400/10 text-amber-400 border-amber-400/20" : "bg-sky-400/10 text-sky-400 border-sky-400/20"
                  )}>
                    {formatCurrency(parentBalance, parent.currency)}
                  </span>
                </div>
                <ChevronDown size={11} className={cn("text-slate-500 transition-transform", isExpanded ? "rotate-180" : "")} />
              </button>

              {isExpanded && (
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 pl-2">
                  {children.map(child => renderAccountCard(child, true))}
                  {children.length === 0 && (
                    <div className="col-span-full py-4 text-center text-slate-600 italic text-[10px] uppercase">
                      No sub-accounts
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}

        {accounts.filter(a => !a.isParent).length === 0 && (
          <div className="py-12 text-center text-slate-500 italic uppercase text-label">
            No accounts configured
          </div>
        )}
      </div>
    </Modal>
  );
}

export function IncomeExpenseModule({
  state,
  updateState,
  activeTab,
  setActiveTab,
  triggerAdd,
  setTriggerAdd,
  onNavigateToModule,
  inheritedData,
  onClearInheritedData
}: IncomeExpenseModuleProps) {
  
  const isFromExternalModuleRef = useRef(false);

  // ─── Extract Data ──────────────────────────────────────────────────────────
  const accounts: BankAccount[] = useMemo(() => state.incomeExpenseAccounts || [], [state.incomeExpenseAccounts]);
  const categories: AppCategory[] = useMemo(() => state.incomeExpenseCategories || DEFAULT_CATEGORIES, [state.incomeExpenseCategories]);

  const INCOME_CATEGORIES: Record<string, string[]> = useMemo(() => {
    const map: Record<string, string[]> = {};
    categories.filter(c => c.type === 'Income').forEach(c => {
      map[c.name] = c.subCategories;
    });
    return map;
  }, [categories]);

  const EXPENSE_CATEGORIES: Record<string, string[]> = useMemo(() => {
    const map: Record<string, string[]> = {};
    categories.filter(c => c.type === 'Expense').forEach(c => {
      map[c.name] = c.subCategories;
    });
    return map;
  }, [categories]);

  const transactions: IncomeExpenseTransaction[] = useMemo(() => state.incomeExpenseTransactions || [], [state.incomeExpenseTransactions]);
  const conversionRates = useMemo(() => state.conversionRates || { USD_to_BDT: 120, LYD_to_BDT: 20 }, [state.conversionRates]);

  const USD_rate = conversionRates.USD_to_BDT;
  const LYD_rate = conversionRates.LYD_to_BDT;

  // ─── Date Range Configuration (This Month, Fiscal, Custom) ──
  const [rangeType, setRangeType] = useState<'this' | 'fiscal' | 'custom'>('fiscal');
  const [customDates, setCustomDates] = useState(() => {
    const now = new Date();
    return {
      start: '2025-01-01',
      end: now.toISOString().split('T')[0]
    };
  });

  const [thisMonthDate, setThisMonthDate] = useState(() => {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1);
  });

  const [fiscalStartYear, setFiscalStartYear] = useState(() => {
    const now = new Date();
    return now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
  });

  const [isRangeMenuOpen, setIsRangeMenuOpen] = useState(false);
  const [isSettingsMenuOpen, setIsSettingsMenuOpen] = useState(false);

  // Excluded items for analytics pie charts (interactive toggle matching DSE Tracker)
  const [excludedExpenseCats, setExcludedExpenseCats] = useState<Set<string>>(new Set());
  const [excludedIncomeCats, setExcludedIncomeCats] = useState<Set<string>>(new Set());
  const [excludedAssetAccounts, setExcludedAssetAccounts] = useState<Set<string>>(new Set());

  const toggleExcludedExpense = (id: string) => {
    setExcludedExpenseCats(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleExcludedIncome = (id: string) => {
    setExcludedIncomeCats(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleExcludedAsset = (id: string) => {
    setExcludedAssetAccounts(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  // Interactive series visibility for Income vs Expense Flow bar chart (matching DSE Tracker)
  const [flowView, setFlowView] = useState<'monthly' | 'cumulative'>('monthly');
  const [hiddenFlowSeries, setHiddenFlowSeries] = useState<Set<string>>(new Set());
  const toggleFlowSeries = (key: string) => {
    setHiddenFlowSeries(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  // Range state for Income vs Expense Flow bar chart (matching DSE Tracker Transaction History)
  const [flowRange, setFlowRange] = useState<'last6m' | 'last12m' | 'fiscal' | 'custom'>('last6m');
  const [flowMonthOffset, setFlowMonthOffset] = useState(0);
  const [flowFiscalOffset, setFlowFiscalOffset] = useState(0);
  const [flowCustomDates, setFlowCustomDates] = useState(() => {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth() - 5, 1);
    const toStr = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    return {
      start: toStr(start),
      end: toStr(now)
    };
  });
  const [isFlowRangeOpen, setIsFlowRangeOpen] = useState(false);

  const { flowStartStr, flowEndStr } = useMemo(() => {
    const now = new Date();
    const toStr = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const todayStr = toStr(now);

    if (flowRange === 'last6m') {
      const targetEnd = new Date(now.getFullYear(), now.getMonth() + flowMonthOffset + 1, 0);
      const targetStart = new Date(now.getFullYear(), now.getMonth() + flowMonthOffset - 5, 1);
      const endStr = flowMonthOffset === 0 ? todayStr : toStr(targetEnd);
      return { flowStartStr: toStr(targetStart), flowEndStr: endStr };
    }
    if (flowRange === 'last12m') {
      const targetEnd = new Date(now.getFullYear(), now.getMonth() + flowMonthOffset + 1, 0);
      const targetStart = new Date(now.getFullYear(), now.getMonth() + flowMonthOffset - 11, 1);
      const endStr = flowMonthOffset === 0 ? todayStr : toStr(targetEnd);
      return { flowStartStr: toStr(targetStart), flowEndStr: endStr };
    }
    if (flowRange === 'fiscal') {
      const currentYear = now.getFullYear();
      const isPostJuly = now.getMonth() >= 6;
      const baseStartYear = (isPostJuly ? currentYear : currentYear - 1) + flowFiscalOffset;
      const start = `${baseStartYear}-07-01`;
      const end = flowFiscalOffset === 0 ? todayStr : `${baseStartYear + 1}-06-30`;
      return { flowStartStr: start, flowEndStr: end };
    }
    // For custom, include whole month of start
    const startVal = flowCustomDates.start || '2020-01-01';
    const parts = startVal.split('-');
    const cy = Number(parts[0]) || 2020;
    const cm = Number(parts[1]) || 1;
    const startOfCustomMonth = `${cy}-${String(cm).padStart(2, '0')}-01`;
    return {
      flowStartStr: startOfCustomMonth,
      flowEndStr: flowCustomDates.end || '9999-12-31'
    };
  }, [flowRange, flowMonthOffset, flowFiscalOffset, flowCustomDates]);

  const navigateFlowDate = (dir: -1 | 1) => {
    if (flowRange === 'last6m' || flowRange === 'last12m') {
      setFlowMonthOffset(prev => prev + dir);
    } else if (flowRange === 'fiscal') {
      setFlowFiscalOffset(prev => prev + dir);
    } else if (flowRange === 'custom') {
      setFlowCustomDates(prev => {
        const s = new Date(prev.start || '2020-01-01');
        const e = new Date(prev.end || '2020-01-01');
        const newS = new Date(s.getFullYear(), s.getMonth() + dir, 1);
        const newE = new Date(e.getFullYear(), e.getMonth() + dir + 1, 0);
        const toStr = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        return { start: toStr(newS), end: toStr(newE) };
      });
    }
  };

  // Helper date generators
  const getFirstOfMonth = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
  const getLastOfMonth = (d: Date) => {
    const last = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const y = last.getFullYear();
    const m = String(last.getMonth() + 1).padStart(2, '0');
    const day = String(last.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };

  const calculatedDates = useMemo(() => {
    if (rangeType === 'this') {
      return {
        start: getFirstOfMonth(thisMonthDate),
        end: getLastOfMonth(thisMonthDate)
      };
    }
    if (rangeType === 'fiscal') {
      return {
        start: `${fiscalStartYear}-07-01`,
        end: `${fiscalStartYear + 1}-06-30`
      };
    }
    return {
      start: customDates.start,
      end: customDates.end
    };
  }, [rangeType, customDates, thisMonthDate, fiscalStartYear]);

  // Navigate custom range months
  const navigateCustomMonth = (offset: number) => {
    const currentStart = new Date(customDates.start);
    const nextMonth = new Date(currentStart.getFullYear(), currentStart.getMonth() + offset, 1);
    setCustomDates({
      start: getFirstOfMonth(nextMonth),
      end: getLastOfMonth(nextMonth)
    });
  };

  // Navigate this month view
  const navigateThisMonth = (offset: number) => {
    setThisMonthDate(prev => new Date(prev.getFullYear(), prev.getMonth() + offset, 1));
  };

  const formatThisMonthLabel = (d: Date) => {
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return `${months[d.getMonth()]} ${d.getFullYear()}`;
  };

  const selectedPeriodLabel = useMemo(() => {
    if (rangeType === 'this') {
      return formatThisMonthLabel(thisMonthDate);
    }
    if (rangeType === 'fiscal') {
      return `July ${fiscalStartYear} - June ${fiscalStartYear + 1}`;
    }
    return `${formatDate(customDates.start)} - ${formatDate(customDates.end)}`;
  }, [rangeType, thisMonthDate, fiscalStartYear, customDates]);

  // Drag states and handlers for reordering accounts in settings
  const [draggableRowId, setDraggableRowId] = useState<string | null>(null);
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);

  const handleDragStart = (e: React.DragEvent<HTMLTableRowElement>, index: number) => {
    e.dataTransfer.effectAllowed = 'move';
    setDraggedIndex(index);
  };

  const handleDragOver = (e: React.DragEvent<HTMLTableRowElement>, index: number) => {
    e.preventDefault();
    if (draggedIndex === null || draggedIndex === index) return;

    const updatedAccounts = [...accounts];
    const draggedItem = updatedAccounts[draggedIndex];
    updatedAccounts.splice(draggedIndex, 1);
    updatedAccounts.splice(index, 0, draggedItem);

    setDraggedIndex(index);
    updateState(prev => ({
      ...prev,
      incomeExpenseAccounts: updatedAccounts
    }));
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
    setDraggableRowId(null);
  };

  // Drag states and handlers for reordering categories in settings
  const [draggableCatId, setDraggableCatId] = useState<string | null>(null);
  const [draggedCatIndex, setDraggedCatIndex] = useState<number | null>(null);

  // Drag states and handlers for movable sub-categories
  const [draggedSubCat, setDraggedSubCat] = useState<{
    subCategory: string;
    sourceCatId: string;
    index: number;
  } | null>(null);
  const [dragOverCatId, setDragOverCatId] = useState<string | null>(null);
  const [dragOverSubIndex, setDragOverSubIndex] = useState<number | null>(null);

  const handleCatDragStart = (e: React.DragEvent<HTMLTableRowElement>, index: number) => {
    e.dataTransfer.effectAllowed = 'move';
    setDraggedCatIndex(index);
  };

  const handleCatDragOver = (e: React.DragEvent<HTMLTableRowElement>, index: number) => {
    e.preventDefault();
    if (draggedCatIndex === null || draggedCatIndex === index) return;

    const updatedCategories = [...categories];
    const draggedItem = updatedCategories[draggedCatIndex];
    updatedCategories.splice(draggedCatIndex, 1);
    updatedCategories.splice(index, 0, draggedItem);

    setDraggedCatIndex(index);
    updateState(prev => ({
      ...prev,
      incomeExpenseCategories: updatedCategories
    }));
  };

  const handleCatDragEnd = () => {
    setDraggedCatIndex(null);
    setDraggableCatId(null);
  };

  const handleSubCatDragStart = (e: React.DragEvent, subCategory: string, sourceCatId: string, index: number) => {
    e.stopPropagation();
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', subCategory);
    setDraggedSubCat({ subCategory, sourceCatId, index });
  };

  const handleSubCatDragEnd = () => {
    setDraggedSubCat(null);
    setDragOverCatId(null);
    setDragOverSubIndex(null);
  };

  const handleMoveSubCategory = (sourceCatId: string, targetCatId: string, subCategory: string, targetIndex?: number) => {
    const sourceCat = categories.find(c => c.id === sourceCatId);
    const targetCat = categories.find(c => c.id === targetCatId);
    if (!sourceCat || !targetCat) return;

    // Reorder within the same category
    if (sourceCatId === targetCatId) {
      const subCats = [...sourceCat.subCategories];
      const curIdx = subCats.indexOf(subCategory);
      if (curIdx !== -1 && targetIndex !== undefined && targetIndex !== curIdx) {
        subCats.splice(curIdx, 1);
        const insertIdx = targetIndex > curIdx ? targetIndex - 1 : targetIndex;
        subCats.splice(insertIdx, 0, subCategory);
        
        const updatedCategories = categories.map(c => c.id === sourceCatId ? { ...c, subCategories: subCats } : c);
        updateState(prev => ({
          ...prev,
          incomeExpenseCategories: updatedCategories
        }));
      }
      return;
    }

    // Moving between different categories
    const newSourceSubCats = sourceCat.subCategories.filter(s => s !== subCategory);
    const newTargetSubCats = [...targetCat.subCategories.filter(s => s !== subCategory)];
    if (targetIndex !== undefined && targetIndex >= 0 && targetIndex <= newTargetSubCats.length) {
      newTargetSubCats.splice(targetIndex, 0, subCategory);
    } else {
      newTargetSubCats.push(subCategory);
    }

    const updatedCategories = categories.map(c => {
      if (c.id === sourceCatId) return { ...c, subCategories: newSourceSubCats };
      if (c.id === targetCatId) return { ...c, subCategories: newTargetSubCats };
      return c;
    });

    // Update existing transactions so they link to the new category name
    const updatedTransactions = transactions.map(t => {
      if (t.category === sourceCat.name && t.subCategory === subCategory) {
        return {
          ...t,
          category: targetCat.name,
          type: targetCat.type === 'Income' ? ('Income' as const) : ('Expense' as const)
        };
      }
      return t;
    });

    updateState(prev => ({
      ...prev,
      incomeExpenseCategories: updatedCategories,
      incomeExpenseTransactions: updatedTransactions
    }));
  };

  // Currency formats
  const formatCurrency = (amount: number, currency: string) => {
    if (currency === 'USD') {
      return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(amount);
    }
    if (currency === 'LYD') {
      return `LD ${amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    return formatBDT(amount);
  };

  const getCurrencyColorClass = (currency?: string) => {
    if (currency === 'BDT') return 'text-teal-400';
    if (currency === 'USD') return 'text-sky-400';
    if (currency === 'LYD') return 'text-amber-400';
    return 'text-white';
  };

  const getTxColorAndPrefix = (t: IncomeExpenseTransaction) => {
    const isIncome = t.type === 'Income' || (t.type === 'Loan' && (t.subCategory === 'Borrowed' || t.subCategory === 'Received Lent Money')) || (t.type === 'Transfer' && t.transferType === 'to');
    const isExpense = t.type === 'Expense' || (t.type === 'Loan' && (t.subCategory === 'Lent' || t.subCategory === 'Repaid Borrowed Money')) || (t.type === 'Transfer' && t.transferType === 'from');

    // A Sell that resulted in a realized loss is still type "Income" (Capital Gain /
    // Stocks Capital Gain) but its amount is negative — show it in red instead of green.
    if (isIncome && t.amount < 0) {
      return { className: 'text-rose-400', prefix: '-' };
    }
    if (isIncome) {
      return { className: 'text-emerald-400', prefix: '+' };
    }
    if (isExpense) {
      return { className: 'text-rose-400', prefix: '-' };
    }
    return { className: 'text-white', prefix: '' };
  };

  // Synthesize initial balance transactions for accounts that have an initial balance tied to their initial date
  const initialBalanceTransactions: IncomeExpenseTransaction[] = useMemo(() => {
    return accounts
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
  }, [accounts]);

  // Combined ledger including initial balances on their respective initial dates
  const allTransactions = useMemo(() => {
    const combined = [...transactions, ...initialBalanceTransactions];
    return combined.sort((a, b) => {
      if (a.date !== b.date) {
        return b.date.localeCompare(a.date);
      }
      return 0;
    });
  }, [transactions, initialBalanceTransactions]);

  // Filtered transactions for the selected range
  const filteredTransactions = useMemo(() => {
    return allTransactions.filter(t => t.date >= calculatedDates.start && t.date <= calculatedDates.end);
  }, [allTransactions, calculatedDates]);

  // Upper limit date for account balances
  const balanceEndDate = useMemo(() => {
    const now = new Date();
    const yStr = now.getFullYear();
    const mStr = String(now.getMonth() + 1).padStart(2, '0');
    const dStr = String(now.getDate()).padStart(2, '0');
    const todayStr = `${yStr}-${mStr}-${dStr}`;

    if (rangeType === 'this') {
      const isCurrentMonth = thisMonthDate.getFullYear() === now.getFullYear() && 
                             thisMonthDate.getMonth() === now.getMonth();
      if (isCurrentMonth) {
        return todayStr;
      }
      return getLastOfMonth(thisMonthDate);
    }
    if (rangeType === 'fiscal') {
      return `${fiscalStartYear + 1}-06-30`;
    }
    return customDates.end;
  }, [rangeType, thisMonthDate, fiscalStartYear, customDates]);

  // ─── Balance Calculations ──────────────────────────────────────────────────
  const accountBalances = useMemo(() => {
    const balances: Record<string, number> = {};
    accounts.forEach(acc => {
      const txs = allTransactions.filter(t => t.date <= balanceEndDate && (t.accountId === acc.id || t.toAccountId === acc.id));
      const income = txs.filter(t => t.type === 'Income' && t.accountId === acc.id).reduce((sum, t) => sum + t.amount, 0);
      const expense = txs.filter(t => t.type === 'Expense' && t.accountId === acc.id).reduce((sum, t) => sum + t.amount, 0);
      
      let transOut = 0;
      let transIn = 0;
      txs.forEach(t => {
        if (t.type === 'Transfer') {
          if (t.transferType) {
            // New two-entry transfer model
            if (t.accountId === acc.id) {
              if (t.transferType === 'from') {
                transOut += t.amount;
              } else if (t.transferType === 'to') {
                transIn += t.amount;
              }
            }
          } else {
            // Legacy single-entry transfer model
            if (t.accountId === acc.id) {
              transOut += t.amount;
            }
            if (t.toAccountId === acc.id) {
              transIn += (t.toAmount !== undefined ? t.toAmount : t.amount);
            }
          }
        }
      });

      const loanInflow = txs.filter(t => (t.type === 'Loan' || t.category === 'Loan') && t.accountId === acc.id && (t.subCategory === 'Borrowed' || t.subCategory === 'Received Lent Money')).reduce((sum, t) => sum + t.amount, 0);
      const loanOutflow = txs.filter(t => (t.type === 'Loan' || t.category === 'Loan') && t.accountId === acc.id && (t.subCategory === 'Lent' || t.subCategory === 'Repaid Borrowed Money')).reduce((sum, t) => sum + t.amount, 0);

      balances[acc.id] = income - expense - transOut + transIn + loanInflow - loanOutflow;
    });
    accounts.forEach(acc => {
      if (acc.isParent) {
        const children = accounts.filter(a => a.parentId === acc.id);
        const childrenSum = children.reduce((sum, child) => sum + (balances[child.id] || 0), 0);
        balances[acc.id] = (balances[acc.id] || 0) + childrenSum;
      }
    });
    return balances;
  }, [accounts, allTransactions, balanceEndDate]);

  const totalBDT = useMemo(() => {
    return accounts
      .filter(a => a.currency === 'BDT' && !a.parentId)
      .reduce((sum, a) => sum + (accountBalances[a.id] || 0), 0);
  }, [accounts, accountBalances]);

  const totalLYD = useMemo(() => {
    return accounts
      .filter(a => a.currency === 'LYD' && !a.parentId)
      .reduce((sum, a) => sum + (accountBalances[a.id] || 0), 0);
  }, [accounts, accountBalances]);

  const totalUSD = useMemo(() => {
    return accounts
      .filter(a => a.currency === 'USD' && !a.parentId)
      .reduce((sum, a) => sum + (accountBalances[a.id] || 0), 0);
  }, [accounts, accountBalances]);

  const loanStatsByCurrency = useMemo(() => {
    const stats = {
      BDT: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
      LYD: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
      USD: { borrowed: 0, repaid: 0, lent: 0, received: 0 },
    };

    transactions.forEach(t => {
      if (t.date > balanceEndDate) return;
      if (t.type !== 'Loan' && t.category !== 'Loan') return;

      const acc = accounts.find(a => a.id === t.accountId);
      if (!acc) return;
      const currency = acc.currency as 'BDT' | 'LYD' | 'USD';
      if (!stats[currency]) return;

      if (t.subCategory === 'Borrowed') {
        stats[currency].borrowed += t.amount;
      } else if (t.subCategory === 'Repaid Borrowed Money') {
        stats[currency].repaid += t.amount;
      } else if (t.subCategory === 'Lent') {
        stats[currency].lent += t.amount;
      } else if (t.subCategory === 'Received Lent Money') {
        stats[currency].received += t.amount;
      }
    });

    return {
      BDT: {
        borrowed: Math.max(0, stats.BDT.borrowed - stats.BDT.repaid),
        lent: Math.max(0, stats.BDT.lent - stats.BDT.received)
      },
      LYD: {
        borrowed: Math.max(0, stats.LYD.borrowed - stats.LYD.repaid),
        lent: Math.max(0, stats.LYD.lent - stats.LYD.received)
      },
      USD: {
        borrowed: Math.max(0, stats.USD.borrowed - stats.USD.repaid),
        lent: Math.max(0, stats.USD.lent - stats.USD.received)
      }
    };
  }, [transactions, accounts, balanceEndDate]);

  const { totalCurrentBorrowedBDT, totalCurrentLentBDT } = useMemo(() => {
    const { BDT, LYD, USD } = loanStatsByCurrency;
    return {
      totalCurrentBorrowedBDT: BDT.borrowed + (LYD.borrowed * LYD_rate) + (USD.borrowed * USD_rate),
      totalCurrentLentBDT: BDT.lent + (LYD.lent * LYD_rate) + (USD.lent * USD_rate)
    };
  }, [loanStatsByCurrency, LYD_rate, USD_rate]);

  const totalConvertedBDT = useMemo(() => {
    const rawSum = totalBDT + (totalLYD * LYD_rate) + (totalUSD * USD_rate);
    return rawSum - totalCurrentBorrowedBDT + totalCurrentLentBDT;
  }, [totalBDT, totalLYD, totalUSD, LYD_rate, USD_rate, totalCurrentBorrowedBDT, totalCurrentLentBDT]);


  // ─── Form / Modal State ────────────────────────────────────────────────────
  const [isTxModalOpen, setIsTxModalOpen] = useState(false);
  const [editingTx, setEditingTx] = useState<IncomeExpenseTransaction | null>(null);
  
  // Sub-picker modal states
  const [isCategoryPickerOpen, setIsCategoryPickerOpen] = useState(false);
  const [isAccountPickerOpen, setIsAccountPickerOpen] = useState(false);
  const [isToAccountPickerOpen, setIsToAccountPickerOpen] = useState(false);

  // Custom non-blocking Alert & Confirm dialog states
  const [customAlert, setCustomAlert] = useState<{ message: string; title: string } | null>(null);
  const [customConfirm, setCustomConfirm] = useState<{ 
    message: string; 
    title: string; 
    onConfirm: () => void;
    variant?: 'danger' | 'warning' | 'info' | 'severe' | 'critical';
    confirmLabel?: string;
    cancelLabel?: string;
    details?: (string | null | undefined)[];
  } | null>(null);

  const runAlert = (message: string, title: string = "Notice") => {
    setCustomAlert({ message, title });
  };

  const runConfirm = (
    message: string, 
    onConfirm: () => void, 
    title: string = "Confirm Action",
    options?: {
      variant?: 'danger' | 'warning' | 'info' | 'severe' | 'critical';
      confirmLabel?: string;
      cancelLabel?: string;
      details?: (string | null | undefined)[];
    }
  ) => {
    setCustomConfirm({ 
      message, 
      title, 
      onConfirm,
      variant: options?.variant || 'danger',
      confirmLabel: options?.confirmLabel,
      cancelLabel: options?.cancelLabel,
      details: options?.details
    });
  };
  
      const [syncToDse, setSyncToDse] = useState<boolean>(true);
  const [inheritedDseTxId, setInheritedDseTxId] = useState<string | undefined>(undefined);
  const [inheritedMfGroupId, setInheritedMfGroupId] = useState<string | undefined>(undefined);
  const [inheritedSukukGroupId, setInheritedSukukGroupId] = useState<string | undefined>(undefined);
  const [inheritedSukukRentTxId, setInheritedSukukRentTxId] = useState<string | undefined>(undefined);
  const [inheritedFdrProfitTxId, setInheritedFdrProfitTxId] = useState<string | undefined>(undefined);
  const [inheritedOnlineGroupId, setInheritedOnlineGroupId] = useState<string | undefined>(undefined);
  const [inheritedOnlineInstallmentTxId, setInheritedOnlineInstallmentTxId] = useState<string | undefined>(undefined);

  const [txForm, setTxForm] = useState({
    date: new Date().toISOString().split('T')[0],
    type: 'Expense' as 'Income' | 'Expense' | 'Transfer' | 'Loan',
    accountId: accounts[0]?.id || '',
    toAccountId: accounts[1]?.id || accounts[0]?.id || '',
    category: '',
    subCategory: '',
    amount: '',
    toAmount: '',
    description: ''
  });

  // Handle setting primary account when list updates
  useEffect(() => {
    if (accounts.length > 0 && !txForm.accountId) {
      const initialToAcc = accounts[1] || accounts[0];
      setTxForm(prev => ({ 
        ...prev, 
        accountId: accounts[0].id,
        toAccountId: initialToAcc.id
      }));
      setSyncToDse(isBoAccount(initialToAcc));
    }
  }, [accounts]);

  // Watch for external "New Transaction" action from navbar
  useEffect(() => {
    if (triggerAdd) {
      setEditingTx(null);
      if (inheritedData) {
        isFromExternalModuleRef.current = true;
        const targetType = inheritedData.type || 'Transfer';
        const targetDate = inheritedData.date || new Date().toISOString().split('T')[0];
        const targetAmount = inheritedData.amount ? String(inheritedData.amount) : '';
        const targetDesc = inheritedData.description || '';
        setInheritedDseTxId(inheritedData.targetModule === 'sukuk-rent' ? undefined : inheritedData.linkedTxId);
        setInheritedSukukRentTxId(inheritedData.targetModule === 'sukuk-rent' ? inheritedData.linkedTxId : undefined);
        setInheritedFdrProfitTxId((inheritedData.targetModule === 'fdr-profit' || inheritedData.targetModule === 'fdr-charge') ? inheritedData.linkedTxId : undefined);
        setInheritedOnlineInstallmentTxId((inheritedData.targetModule === 'online-installment' || inheritedData.targetModule === 'online-interest' || inheritedData.targetModule === 'online-return') ? inheritedData.linkedTxId : undefined);

        // Find appropriate toAccountId based on target module
        let targetToAcc = accounts.find(a => !a.isParent && (a.id === 'bdt-mutual-fund' || a.id === 'bdt-online-investment' || a.id === 'bdt-sukuk' || a.id === 'bdt-ibbl-fdr' || a.id === 'bdt-fdr'));
        if (inheritedData.targetModule === 'mutual-funds') {
          targetToAcc = accounts.find(a => isMutualFundAccount(a)) || targetToAcc;
        } else if (inheritedData.targetModule === 'online') {
          targetToAcc = accounts.find(a => isOnlineInvestmentAccount(a)) || targetToAcc;
        } else if (inheritedData.targetModule === 'sukuk') {
          targetToAcc = accounts.find(a => isSukukAccount(a)) || targetToAcc;
        } else if (inheritedData.targetModule === 'fdrs' || inheritedData.targetModule === 'fixed-deposits') {
          targetToAcc = accounts.find(a => isFdrAccount(a)) || targetToAcc;
        } else if (inheritedData.targetModule === 'dse' || inheritedData.targetModule === 'dse-dividend') {
          targetToAcc = accounts.find(a => isBoAccount(a)) || targetToAcc;
        } else if (inheritedData.targetModule === 'dse-withdrawal') {
          // DSE Withdrawal = Transfer FROM BO Account TO a normal account.
          targetToAcc = accounts.find(a =>
            !a.isParent &&
            a.currency === 'BDT' &&
            !isBoAccount(a) &&
            !isMutualFundAccount(a) &&
            !isOnlineInvestmentAccount(a) &&
            !isSukukAccount(a) &&
            !isFdrAccount(a)
          ) || accounts.find(a => !a.isParent && !isBoAccount(a)) || targetToAcc;
        } else if (inheritedData.targetModule === 'online-return' || inheritedData.targetModule === 'online-interest') {
          // Online Investment repayment / interest: Transfer To / receiving Account = City Islamic
          targetToAcc = accounts.find(a => !a.isParent && a.name.toLowerCase().includes('city islamic')) || accounts.find(a => !a.isParent && a.name.toLowerCase().includes('city')) || targetToAcc;
        } else if (inheritedData.targetModule === 'mutual-funds-withdrawal') {
          // Mutual Fund Withdrawal = Transfer FROM Mutual Fund Account TO City Islamic Account.
          targetToAcc = accounts.find(a => !a.isParent && a.name.toLowerCase().includes('city islamic')) || accounts.find(a => !a.isParent && a.name.toLowerCase().includes('city')) || targetToAcc;
        }
        if (!targetToAcc) {
          targetToAcc = accounts.filter(a => !a.isParent)[1] || accounts[0];
        }

        // Pick source account (non-parent account, preferably matching specific context or default BDT bank)
        const nonParentAccounts = accounts.filter(a => !a.isParent);
        
        let explicitAcc: BankAccount | undefined = undefined;
        if (inheritedData.accountId) {
          explicitAcc = nonParentAccounts.find(a => a.id === inheritedData.accountId);
        }

        // For FDR Profit / Charge: in case of 1409 default to IBBL Savings, for 0216 default to IBBL FDR USD
        if (!explicitAcc && (inheritedData.targetModule === 'fdr-profit' || inheritedData.targetModule === 'fdr-charge' || targetDesc.toLowerCase().includes('fixed deposit') || targetDesc.toLowerCase().includes('fdr'))) {
          const descLower = targetDesc.toLowerCase();
          if (descLower.includes('1409')) {
            explicitAcc = nonParentAccounts.find(a => a.id === 'bdt-ibbl-savings' || a.name.toLowerCase().includes('ibbl savings'));
          } else if (descLower.includes('0216')) {
            explicitAcc = nonParentAccounts.find(a => a.id === 'usd-ibbl-fdr' || (a.name.toLowerCase().includes('ibbl') && a.name.toLowerCase().includes('fdr') && a.currency === 'USD') || a.name.toLowerCase().includes('ibbl fdr usd'));
          }
        }

        const onlineAcc = nonParentAccounts.find(a => isOnlineInvestmentAccount(a)) || accounts.find(a => a.id === 'bdt-online-investment');
        const mutualFundAcc = nonParentAccounts.find(a => isMutualFundAccount(a)) || accounts.find(a => a.id === 'bdt-mutual-fund');
        const cityIslamicAcc = nonParentAccounts.find(a => a.name.toLowerCase().includes('city islamic')) || nonParentAccounts.find(a => a.name.toLowerCase().includes('city'));

        const bankAcc = explicitAcc
          || (inheritedData.targetModule === 'online-interest' ? cityIslamicAcc : undefined)
          || (inheritedData.targetModule === 'dse-dividend' ? nonParentAccounts.find(a => isBoAccount(a)) : undefined)
          || nonParentAccounts.find(a => a.currency === 'BDT' && !isMutualFundAccount(a) && !isOnlineInvestmentAccount(a) && !isSukukAccount(a) && !isFdrAccount(a) && !isBoAccount(a))
          || nonParentAccounts[0]
          || accounts[0];

        const sourceAcc = inheritedData.targetModule === 'dse-withdrawal'
          ? (explicitAcc || nonParentAccounts.find(a => isBoAccount(a)) || bankAcc)
          : inheritedData.targetModule === 'dse-dividend'
            ? (explicitAcc || nonParentAccounts.find(a => isBoAccount(a)) || bankAcc)
            : inheritedData.targetModule === 'dse-charge'
              ? (explicitAcc || nonParentAccounts.find(a => isBoAccount(a)) || bankAcc)
              : inheritedData.targetModule === 'online-return'
            ? (explicitAcc || onlineAcc || bankAcc)
            : inheritedData.targetModule === 'online-interest'
              ? (targetType === 'Income' ? (cityIslamicAcc || bankAcc) : (onlineAcc || bankAcc))
              : inheritedData.targetModule === 'mutual-funds-withdrawal'
                ? (explicitAcc || mutualFundAcc || bankAcc)
                : inheritedData.targetModule === 'mutual-funds-dividend'
                  ? (explicitAcc || mutualFundAcc || bankAcc)
                  : inheritedData.targetModule === 'fdr-charge'
                    ? (explicitAcc || bankAcc)
                    : targetType === 'Income'
                    ? (explicitAcc || bankAcc)
                    : (nonParentAccounts.find(a => a.id !== targetToAcc?.id && a.currency === 'BDT' && !isMutualFundAccount(a) && !isOnlineInvestmentAccount(a) && !isSukukAccount(a) && !isFdrAccount(a) && !isBoAccount(a))
                       || nonParentAccounts.find(a => a.id !== targetToAcc?.id)
                       || accounts[0]);

        if (inheritedData.targetModule === 'online-interest' && cityIslamicAcc) {
          targetToAcc = cityIslamicAcc;
        }

        let defaultCat = targetType === 'Transfer' ? 'Transfer' : '';
        let defaultSubCat = '';
        if (inheritedData.category) {
          defaultCat = inheritedData.category;
          defaultSubCat = inheritedData.subCategory || '';
        } else if (inheritedData.targetModule === 'online-interest') {
          defaultCat = 'Other Income';
          defaultSubCat = 'Investment Interest';
        } else if (inheritedData.targetModule === 'dse-dividend') {
          const finCat = categories.find(c => c.name.toLowerCase().includes('finance') || c.name.toLowerCase().includes('income') || c.name.toLowerCase().includes('investment'));
          defaultCat = finCat ? finCat.name : 'Finance Income';
          const sub = finCat?.subCategories?.find(s => s.toLowerCase().includes('dividend'));
          defaultSubCat = sub || 'Dividend';
        } else if (targetType === 'Income') {
          const finCat = categories.find(c => c.name.toLowerCase().includes('finance') || c.name.toLowerCase().includes('income') || c.name.toLowerCase().includes('investment'));
          if (finCat) {
            defaultCat = finCat.name;
            const sub = finCat.subCategories?.find(s => s.toLowerCase().includes('online') || s.toLowerCase().includes('interest') || s.toLowerCase().includes('profit') || s.toLowerCase().includes('fdr'));
            if (sub) defaultSubCat = sub;
          }
        }

        const isTargetModuleDse = inheritedData.targetModule === 'dse' ||
          inheritedData.targetModule === 'dse-dividend' ||
          inheritedData.targetModule === 'dse-withdrawal' ||
          inheritedData.targetModule === 'dse-sell' ||
          inheritedData.targetModule === 'dse-charge';
        const isTargetModuleNonDse = inheritedData.targetModule === 'online' ||
          inheritedData.targetModule === 'online-interest' ||
          inheritedData.targetModule === 'online-return' ||
          inheritedData.targetModule === 'online-installment' ||
          inheritedData.targetModule === 'mutual-funds' ||
          inheritedData.targetModule === 'mutual-funds-dividend' ||
          inheritedData.targetModule === 'mutual-funds-withdrawal' ||
          inheritedData.targetModule === 'sukuk' ||
          inheritedData.targetModule === 'sukuk-rent' ||
          inheritedData.targetModule === 'fdr' ||
          inheritedData.targetModule === 'fdr-profit' ||
          inheritedData.targetModule === 'fdr-charge';

        setSyncToDse(
          !isTargetModuleNonDse && (
            isTargetModuleDse ||
            (targetType === 'Transfer' && (isBoAccount(targetToAcc) || isBoAccount(sourceAcc))) ||
            (targetType === 'Income' && (isBoAccount(sourceAcc) || isTargetModuleDse)) ||
            (targetType === 'Expense' && (isBoAccount(sourceAcc) || isTargetModuleDse))
          )
        );
        setInheritedMfGroupId(
          (inheritedData.targetModule === 'mutual-funds' ||
           inheritedData.targetModule === 'mutual-funds-withdrawal' ||
           inheritedData.targetModule === 'mutual-funds-dividend')
            ? inheritedData.linkedTxId
            : undefined
        );
        setInheritedSukukGroupId(inheritedData.targetModule === 'sukuk' ? inheritedData.linkedTxId : undefined);
        setInheritedOnlineGroupId(inheritedData.targetModule === 'online' ? inheritedData.linkedTxId : undefined);
        setInheritedOnlineInstallmentTxId((inheritedData.targetModule === 'online-installment' || inheritedData.targetModule === 'online-interest' || inheritedData.targetModule === 'online-return') ? inheritedData.linkedTxId : undefined);
        setTxForm({
          date: targetDate,
          type: targetType,
          accountId: sourceAcc?.id || '',
          toAccountId: targetToAcc?.id || '',
          category: defaultCat,
          subCategory: defaultSubCat,
          amount: targetAmount,
          toAmount: targetAmount,
          description: targetDesc
        });

        setTimeout(() => {
          if (onClearInheritedData) onClearInheritedData();
        }, 0);
      } else {
        isFromExternalModuleRef.current = false;
        setInheritedDseTxId(undefined);
        setInheritedMfGroupId(undefined);
        setInheritedSukukGroupId(undefined);
        setInheritedSukukRentTxId(undefined);
        setInheritedFdrProfitTxId(undefined);
        setInheritedOnlineGroupId(undefined);
        setInheritedOnlineInstallmentTxId(undefined);
        const initialToAcc = accounts[1] || accounts[0];
        setSyncToDse(isBoAccount(initialToAcc));
        setTxForm({
          date: new Date().toISOString().split('T')[0],
          type: 'Expense',
          accountId: accounts[0]?.id || '',
          toAccountId: initialToAcc?.id || accounts[0]?.id || '',
          category: '',
          subCategory: '',
          amount: '',
          toAmount: '',
          description: ''
        });
      }
      setIsTxModalOpen(true);
      setTimeout(() => {
        if (setTriggerAdd) setTriggerAdd(false);
      }, 0);
    }
  }, [triggerAdd, setTriggerAdd, accounts, inheritedData, onClearInheritedData]);

  // Derived info for selected account
  const selectedAccount = useMemo(() => {
    return accounts.find(a => a.id === txForm.accountId) || null;
  }, [accounts, txForm.accountId]);

  const selectedCategoryObj = useMemo(() => {
    return categories.find(c => c.name === txForm.category) || null;
  }, [categories, txForm.category]);

  // Auto pre-select the "Sync" checkbox when creating a brand-new Dividend
  // income entry against a BO Account (DSE Tracker).
  useEffect(() => {
    if (editingTx) return;
    if (txForm.type !== 'Income') return;
    const isDividendSub = (txForm.subCategory || '').toLowerCase().includes('dividend');
    if (!isDividendSub) return;
    if (isBoAccount(selectedAccount)) {
      setSyncToDse(true);
    } else {
      setSyncToDse(false);
    }
  }, [editingTx, txForm.type, txForm.subCategory, selectedAccount]);

  // Clean form when type switches
  const handleTypeChange = (type: 'Income' | 'Expense' | 'Transfer' | 'Loan') => {
    if (type === 'Transfer') {
      const targetAcc = accounts.find(a => a.id === txForm.toAccountId);
      const sourceAcc = accounts.find(a => a.id === txForm.accountId);
      if (isBoAccount(targetAcc) || isBoAccount(sourceAcc)) {
        setSyncToDse(true);
      }
    }
    setTxForm(prev => ({
      ...prev,
      type,
      category: type === 'Transfer' ? 'Transfer' : '',
      subCategory: ''
    }));
  };

  // ─── NARRATION & SUB-CATEGORY AUTO-COMPLETE & SUGGESTIONS ─────────────
  const [isNarrationFocused, setIsNarrationFocused] = useState(false);
  const [highlightedSuggestionIdx, setHighlightedSuggestionIdx] = useState<number>(-1);
  const [autoFillNotice, setAutoFillNotice] = useState<{
    show: boolean;
    type: string;
    category: string;
    subCategory?: string;
    accountName: string;
    source: 'subcategory' | 'history';
  } | null>(null);

  const narrationContainerRef = useRef<HTMLDivElement>(null);

  // Search previous database entries AND configured sub-categories
  const narrationSuggestions = useMemo(() => {
    const rawSearch = txForm.description.trim().toLowerCase();
    
    // Sort transactions from newest to oldest for recency analysis
    const sortedTxs = [...transactions].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());

    // Helper to find the most appropriate account for a category + subCategory combination
    const resolveAccountForSubCategory = (categoryName: string, subCategoryName: string) => {
      // 0. FDR specific routing for 1409 (IBBL Savings) and 0216 (IBBL FDR USD)
      if (subCategoryName.toLowerCase().includes('fdr') || subCategoryName.toLowerCase().includes('provisional profit') || subCategoryName.toLowerCase().includes('fixed deposit')) {
        const queryText = (rawSearch + ' ' + txForm.description).toLowerCase();
        if (queryText.includes('1409')) {
          const ibblSav = accounts.find(a => !a.isParent && (a.id === 'bdt-ibbl-savings' || a.name.toLowerCase().includes('ibbl savings')));
          if (ibblSav) {
            return {
              accountId: ibblSav.id,
              accountName: ibblSav.name,
              accountCurrency: ibblSav.currency,
              frequency: 99,
              lastDate: undefined
            };
          }
        }
        if (queryText.includes('0216')) {
          const ibblFdrUsd = accounts.find(a => !a.isParent && (a.id === 'usd-ibbl-fdr' || a.name.toLowerCase().includes('ibbl fdr usd') || (a.name.toLowerCase().includes('ibbl') && a.currency === 'USD')));
          if (ibblFdrUsd) {
            return {
              accountId: ibblFdrUsd.id,
              accountName: ibblFdrUsd.name,
              accountCurrency: ibblFdrUsd.currency,
              frequency: 99,
              lastDate: undefined
            };
          }
        }
      }

      // 1. Check history for exact category & subcategory match
      const exactMatch = sortedTxs.find(t => 
        t.category === categoryName && 
        (t.subCategory || '').toLowerCase() === subCategoryName.toLowerCase() &&
        accounts.some(a => a.id === t.accountId)
      );
      if (exactMatch) {
        const acc = accounts.find(a => a.id === exactMatch.accountId);
        return {
          accountId: exactMatch.accountId,
          accountName: acc?.name || 'Primary Account',
          accountCurrency: acc?.currency || 'BDT',
          frequency: sortedTxs.filter(t => t.category === categoryName && (t.subCategory || '').toLowerCase() === subCategoryName.toLowerCase()).length,
          lastDate: exactMatch.date
        };
      }

      // 2. Check history for category match
      const catMatch = sortedTxs.find(t => 
        t.category === categoryName && 
        accounts.some(a => a.id === t.accountId)
      );
      if (catMatch) {
        const acc = accounts.find(a => a.id === catMatch.accountId);
        return {
          accountId: catMatch.accountId,
          accountName: acc?.name || 'Primary Account',
          accountCurrency: acc?.currency || 'BDT',
          frequency: 1,
          lastDate: catMatch.date
        };
      }

      // 3. Fallback to current form account or first account
      const currentAcc = accounts.find(a => a.id === txForm.accountId) || accounts[0];
      return {
        accountId: currentAcc?.id || '',
        accountName: currentAcc?.name || 'Primary Account',
        accountCurrency: currentAcc?.currency || 'BDT',
        frequency: 0,
        lastDate: undefined
      };
    };

    interface SuggestionItem {
      key: string;
      source: 'subcategory' | 'history';
      description: string;
      type: 'Income' | 'Expense' | 'Transfer' | 'Loan';
      category: string;
      subCategory: string;
      accountId: string;
      accountName: string;
      accountCurrency: string;
      toAccountId?: string;
      toAccountName?: string;
      amount?: number;
      lastDate?: string;
      frequency: number;
    }

    const suggestionsMap = new Map<string, SuggestionItem>();

    // 1. Scan all Sub-Categories in categories catalog
    for (const cat of categories) {
      if (!cat.subCategories || cat.subCategories.length === 0) continue;

      for (const subCat of cat.subCategories) {
        const subCatTrimmed = subCat.trim();
        const subCatLower = subCatTrimmed.toLowerCase();
        const catNameLower = cat.name.toLowerCase();

        // Match if search query matches subcategory name or parent category name
        const isSubCatMatch = rawSearch ? subCatLower.includes(rawSearch) : false;
        const isCatMatch = rawSearch ? catNameLower.includes(rawSearch) : false;

        if (!rawSearch || isSubCatMatch || isCatMatch) {
          const accInfo = resolveAccountForSubCategory(cat.name, subCatTrimmed);
          const key = `subcat::${subCatLower}::${cat.name}::${cat.type}`;

          suggestionsMap.set(key, {
            key,
            source: 'subcategory',
            description: subCatTrimmed,
            type: cat.type as 'Income' | 'Expense' | 'Transfer' | 'Loan',
            category: cat.name,
            subCategory: subCatTrimmed,
            accountId: accInfo.accountId,
            accountName: accInfo.accountName,
            accountCurrency: accInfo.accountCurrency,
            lastDate: accInfo.lastDate,
            frequency: accInfo.frequency
          });
        }
      }
    }

    // 2. Scan all previous transactions in ledger history
    for (const tx of sortedTxs) {
      if (!tx.description || !tx.description.trim() || tx.id.startsWith('init-bal-')) continue;
      if (tx.type === 'Transfer' && tx.transferType === 'to') continue;

      const desc = tx.description.trim();
      const descLower = desc.toLowerCase();
      const subCatLower = (tx.subCategory || '').toLowerCase();

      // Match query against narration or transaction subcategory
      if (rawSearch && !descLower.includes(rawSearch) && !subCatLower.includes(rawSearch)) {
        continue;
      }

      const acc = accounts.find(a => a.id === tx.accountId);
      const toAcc = tx.toAccountId ? accounts.find(a => a.id === tx.toAccountId) : undefined;
      const accountName = acc ? acc.name : 'Unknown Account';
      const accountCurrency = acc ? acc.currency : 'BDT';
      const toAccountName = toAcc ? toAcc.name : undefined;

      const key = `tx::${descLower}::${tx.type}::${tx.category}::${tx.subCategory || ''}::${tx.accountId}`;

      if (!suggestionsMap.has(key)) {
        suggestionsMap.set(key, {
          key,
          source: 'history',
          description: desc,
          type: tx.type,
          category: tx.category,
          subCategory: tx.subCategory || '',
          accountId: tx.accountId,
          accountName,
          accountCurrency,
          toAccountId: tx.toAccountId,
          toAccountName,
          amount: tx.amount,
          lastDate: tx.date,
          frequency: 1
        });
      } else {
        const existing = suggestionsMap.get(key)!;
        existing.frequency += 1;
      }
    }

    const list = Array.from(suggestionsMap.values());

    if (!rawSearch) {
      // If no query typed yet, return top 6 most recent/frequently used entries
      return list
        .sort((a, b) => (b.frequency || 0) - (a.frequency || 0))
        .slice(0, 6);
    }

    // Rank list based on match relevance and subcategory priority
    return list.sort((a, b) => {
      const aSubLower = a.subCategory.toLowerCase();
      const bSubLower = b.subCategory.toLowerCase();
      const aDescLower = a.description.toLowerCase();
      const bDescLower = b.description.toLowerCase();

      // Exact subcategory match gets highest priority
      const aSubExact = aSubLower === rawSearch;
      const bSubExact = bSubLower === rawSearch;
      if (aSubExact && !bSubExact) return -1;
      if (!aSubExact && bSubExact) return 1;

      // Exact description match
      const aDescExact = aDescLower === rawSearch;
      const bDescExact = bDescLower === rawSearch;
      if (aDescExact && !bDescExact) return -1;
      if (!aDescExact && bDescExact) return 1;

      // Starts with query on subcategory
      const aSubStarts = aSubLower.startsWith(rawSearch);
      const bSubStarts = bSubLower.startsWith(rawSearch);
      if (aSubStarts && !bSubStarts) return -1;
      if (!aSubStarts && bSubStarts) return 1;

      // Starts with query on narration
      const aDescStarts = aDescLower.startsWith(rawSearch);
      const bDescStarts = bDescLower.startsWith(rawSearch);
      if (aDescStarts && !bDescStarts) return -1;
      if (!aDescStarts && bDescStarts) return 1;

      // Subcategory source priority
      if (a.source === 'subcategory' && b.source !== 'subcategory' && (aSubStarts || aSubLower.includes(rawSearch))) {
        return -1;
      }
      if (b.source === 'subcategory' && a.source !== 'subcategory' && (bSubStarts || bSubLower.includes(rawSearch))) {
        return 1;
      }

      // Frequency
      if (b.frequency !== a.frequency) {
        return b.frequency - a.frequency;
      }

      return 0;
    }).slice(0, 8);
  }, [categories, transactions, accounts, txForm.description, txForm.accountId]);

  // Close suggestions when clicking outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (narrationContainerRef.current && !narrationContainerRef.current.contains(e.target as Node)) {
        setIsNarrationFocused(false);
        setHighlightedSuggestionIdx(-1);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Auto-fill fields based on selected narration / subcategory entry
  const handleSelectNarrationSuggestion = (sug: {
    source?: 'subcategory' | 'history';
    description: string;
    type: 'Income' | 'Expense' | 'Transfer' | 'Loan';
    category: string;
    subCategory: string;
    accountId: string;
    accountName: string;
    toAccountId?: string;
  }) => {
    // Ensure account exists in current accounts list or fallback
    const targetAccountId = accounts.some(a => a.id === sug.accountId) 
      ? sug.accountId 
      : (accounts[0]?.id || txForm.accountId);

    setTxForm(prev => ({
      ...prev,
      description: sug.description,
      type: sug.type,
      category: sug.category,
      subCategory: sug.subCategory || '',
      accountId: targetAccountId,
      toAccountId: sug.toAccountId && accounts.some(a => a.id === sug.toAccountId)
        ? sug.toAccountId
        : prev.toAccountId
    }));

    setIsNarrationFocused(false);
    setHighlightedSuggestionIdx(-1);

    setAutoFillNotice({
      show: true,
      type: sug.type,
      category: sug.category,
      subCategory: sug.subCategory,
      accountName: sug.accountName,
      source: sug.source || 'history'
    });

    setTimeout(() => {
      setAutoFillNotice(null);
    }, 4500);
  };

  const handleNarrationKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!isNarrationFocused || narrationSuggestions.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setHighlightedSuggestionIdx(prev => 
        prev < narrationSuggestions.length - 1 ? prev + 1 : 0
      );
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setHighlightedSuggestionIdx(prev => 
        prev > 0 ? prev - 1 : narrationSuggestions.length - 1
      );
    } else if (e.key === 'Enter' && highlightedSuggestionIdx >= 0) {
      e.preventDefault();
      handleSelectNarrationSuggestion(narrationSuggestions[highlightedSuggestionIdx]);
    } else if (e.key === 'Escape') {
      setIsNarrationFocused(false);
      setHighlightedSuggestionIdx(-1);
    }
  };

  // ─── Mutual Fund <-> Income & Expense linked sync helper ───────────────────
  const syncMutualFundLinkedTransfer = (groupId: string, updates: { date: string; amount: number } | null) => {
    updateState(prev => {
      const mutualFunds = prev.mutualFunds || [];
      let changed = false;
      const updatedFunds = mutualFunds.map(fund => {
        const hasLink = fund.transactions.some(t => t.linkedIeGroupId === groupId);
        if (!hasLink) return fund;
        changed = true;
        if (updates === null) {
          return { ...fund, transactions: fund.transactions.filter(t => t.linkedIeGroupId !== groupId) };
        }
        return {
          ...fund,
          transactions: fund.transactions.map(t => {
            if (t.linkedIeGroupId !== groupId) return t;
            const isSip = (t.sipAmount || 0) > 0 && !t.isWithdrawal;
            return {
              ...t,
              date: updates.date,
              ...(isSip ? { sipAmount: updates.amount, pullingDate: updates.date } : { amount: updates.amount }),
            };
          })
        };
      });
      return changed ? { ...prev, mutualFunds: updatedFunds } : prev;
    });
  };

  const syncMutualFundDividendIncome = (
    txId: string,
    updates: { date: string; amount: number; description?: string; type?: string; category?: string; subCategory?: string } | null
  ) => {
    updateState(prev => {
      const mutualFunds = prev.mutualFunds || [];
      let changed = false;
      const updatedFunds = mutualFunds.map(fund => {
        const hasLink = fund.transactions.some(t => t.linkedIeGroupId === txId || t.id === txId);
        if (!hasLink) return fund;
        changed = true;
        if (updates === null) {
          return { ...fund, transactions: fund.transactions.filter(t => t.linkedIeGroupId !== txId && t.id !== txId) };
        }
        return {
          ...fund,
          transactions: fund.transactions.map(t => {
            if (t.linkedIeGroupId !== txId && t.id !== txId) return t;
            const isSip = (t.sipAmount || 0) > 0 && !t.isWithdrawal;
            return {
              ...t,
              date: updates.date,
              ...(isSip ? { sipAmount: updates.amount, pullingDate: updates.date } : { amount: updates.amount }),
            };
          })
        };
      });
      return changed ? { ...prev, mutualFunds: updatedFunds } : prev;
    });
  };

  // ─── Sukuk <-> Income & Expense linked sync helper ────────────────────────
  const syncSukukLinkedTransfer = (groupId: string, updates: { date: string; amount: number } | null) => {
    updateState(prev => {
      const sukuks = prev.sukuks || [];
      let changed = false;
      const updatedSukuks = sukuks.map(inv => {
        if (inv.linkedIeGroupId !== groupId) return inv;
        changed = true;
        if (updates === null) {
          return null;
        }
        return {
          ...inv,
          investmentDate: updates.date,
          issueDate: updates.date,
          amount: updates.amount,
          principalAmount: updates.amount,
        };
      }).filter(Boolean) as any[];
      return changed ? { ...prev, sukuks: updatedSukuks } : prev;
    });
  };

  // ─── Sukuk Rent <-> Income & Expense linked sync helper ─────────────────────
  const syncSukukRentIncome = (
    txId: string,
    updates: { date: string; amount: number; description?: string; type?: string; category?: string; subCategory?: string } | null
  ) => {
    updateState(prev => {
      const sukuks = prev.sukuks || [];
      let changed = false;

      // Case 1: An installment is already linked to this txId
      let foundLinked = false;
      const updatedSukuks = sukuks.map(inv => {
        const instIdx = inv.installments?.findIndex(inst => inst.linkedIncomeTxId === txId);
        if (instIdx !== undefined && instIdx >= 0) {
          foundLinked = true;
          changed = true;
          const newInstallments = [...inv.installments];
          const isStillSukukRent = updates && updates.type === 'Income' && (
            updates.subCategory?.toLowerCase() === 'sukuk rent' ||
            (updates.category?.toLowerCase().includes('finance') && updates.subCategory?.toLowerCase().includes('sukuk'))
          );

          if (!isStillSukukRent || updates === null) {
            newInstallments[instIdx] = {
              ...newInstallments[instIdx],
              isPaid: false,
              actualDate: undefined,
              actualAmount: undefined,
              linkedIncomeTxId: undefined,
              isAutoMarked: false,
              isManuallyEdited: false
            };
          } else {
            newInstallments[instIdx] = {
              ...newInstallments[instIdx],
              isPaid: true,
              actualDate: updates.date,
              actualAmount: updates.amount,
              linkedIncomeTxId: txId,
              isAutoMarked: false,
              isManuallyEdited: true
            };
          }

          const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount !== undefined ? inst.actualAmount : inst.amount) : sum, 0);
          const today = new Date().toISOString().split('T')[0];
          let status = inv.status;
          if (newInstallments.every(i => i.isPaid)) status = 'Completed';
          else if (newInstallments.some(i => !i.isPaid && i.date < today)) status = 'Delayed';
          else status = 'Active';

          return {
            ...inv,
            installments: newInstallments,
            totalRepaid,
            status
          };
        }
        return inv;
      });

      if (foundLinked) {
        return changed ? { ...prev, sukuks: updatedSukuks } : prev;
      }

      // Do not auto-guess unlinked Sukuk investments
      return prev;
    });
  };

  // ─── Fixed Deposit Profit <-> Income & Expense linked sync helper ──────────
  const syncFdrProfitIncome = (
    txId: string,
    updates: { date: string; amount: number; description?: string; type?: string; category?: string; subCategory?: string } | null
  ) => {
    updateState(prev => {
      const fdrs = prev.fdrs || [];
      let changed = false;

      // Case 1: A transaction is already linked to this txId
      let foundLinked = false;
      const updatedFdrs = fdrs.map(fdr => {
        const transIdx = fdr.transactions?.findIndex(t => t.linkedIncomeTxId === txId || t.id === txId);
        if (transIdx !== undefined && transIdx >= 0) {
          foundLinked = true;
          changed = true;
          const isStillFdrProfit = updates && updates.type === 'Income' && (
            updates.subCategory?.toLowerCase().includes('fdr') ||
            updates.subCategory?.toLowerCase().includes('provisional profit') ||
            updates.subCategory?.toLowerCase().includes('fixed deposit') ||
            (updates.category?.toLowerCase().includes('finance') && (updates.subCategory?.toLowerCase().includes('profit') || updates.subCategory?.toLowerCase().includes('interest')))
          );
          const isStillFdrCharge = updates && updates.type === 'Expense' && updates.category === 'Other Expense' && (
            updates.subCategory === 'Account Related Fees' ||
            updates.subCategory === 'Other Bank Charges' ||
            updates.subCategory?.toLowerCase().includes('tds')
          );
          const isStillFdrLinked = isStillFdrProfit || isStillFdrCharge;

          if (!isStillFdrLinked || updates === null) {
            return {
              ...fdr,
              transactions: fdr.transactions.filter((_, idx) => idx !== transIdx)
            };
          } else {
            const newTransactions = [...fdr.transactions];
            newTransactions[transIdx] = {
              ...newTransactions[transIdx],
              date: updates.date,
              amount: updates.amount,
              type: isStillFdrProfit ? 'Profit' : 'Charge',
              linkedIncomeTxId: txId
            };
            return {
              ...fdr,
              transactions: newTransactions
            };
          }
        }
        return fdr;
      });

      if (foundLinked) {
        return changed ? { ...prev, fdrs: updatedFdrs } : prev;
      }

      // Do not auto-guess unlinked FDR accounts
      return prev;
    });
  };

  // ─── Online Investment <-> Income & Expense linked sync helper ──────────
  const syncOnlineLinkedTransfer = (groupId: string, updates: { date: string; amount: number } | null) => {
    updateState(prev => {
      const onlineList = prev.onlineInvestments || [];
      let changed = false;
      const updatedOnline = onlineList.map(inv => {
        if (inv.linkedIeGroupId !== groupId && inv.id !== groupId) return inv;
        changed = true;
        if (updates === null) {
          return null;
        }
        return {
          ...inv,
          investmentDate: updates.date,
          amount: updates.amount,
        };
      }).filter(Boolean) as any[];
      return changed ? { ...prev, onlineInvestments: updatedOnline } : prev;
    });
  };

  // ─── Online Investment Installment <-> Income & Expense linked sync helper ──
  const syncOnlineInstallmentIncome = (
    txId: string,
    updates: { date: string; amount: number; description?: string; type?: string; category?: string; subCategory?: string } | null
  ) => {
    updateState(prev => {
      const onlineList = prev.onlineInvestments || [];
      let changed = false;

      // Case 1: An installment is already linked to this txId
      let foundLinked = false;
      const updatedOnline = onlineList.map(inv => {
        const instIdx = inv.installments?.findIndex(inst => inst.linkedIncomeTxId === txId || (txId && inst.linkedIncomeTxId && inst.linkedIncomeTxId === txId));
        if (instIdx !== undefined && instIdx >= 0) {
          foundLinked = true;
          changed = true;
          const newInstallments = [...inv.installments];
          const isStillValid = updates && (updates.type === 'Income' || updates.type === 'Transfer');

          if (!isStillValid || updates === null) {
            newInstallments[instIdx] = {
              ...newInstallments[instIdx],
              isPaid: false,
              actualDate: undefined,
              actualAmount: undefined,
              linkedIncomeTxId: undefined,
              isAutoMarked: false,
              isManuallyEdited: false
            };
          } else {
            newInstallments[instIdx] = {
              ...newInstallments[instIdx],
              isPaid: true,
              actualDate: updates.date,
              actualAmount: updates.amount,
              linkedIncomeTxId: txId,
              isAutoMarked: false,
              isManuallyEdited: true
            };
          }

          const totalRepaid = newInstallments.reduce((sum, inst) => inst.isPaid ? sum + (inst.actualAmount !== undefined ? inst.actualAmount : inst.amount) : sum, 0);
          const today = new Date().toISOString().split('T')[0];
          let status = inv.status;
          if (newInstallments.length > 0 && newInstallments.every(i => i.isPaid)) status = 'Completed';
          else if (newInstallments.some(i => !i.isPaid && i.date < today)) status = 'Delayed';
          else status = 'Active';

          return {
            ...inv,
            installments: newInstallments,
            totalRepaid,
            status
          };
        }
        return inv;
      });

      if (foundLinked) {
        return changed ? { ...prev, onlineInvestments: updatedOnline } : prev;
      }

      return prev;
    });
  };

  // Save Transaction Handler
  const handleSaveTx = (e: React.FormEvent) => {
    e.preventDefault();
    if (!txForm.accountId || !txForm.amount) return;

    const parsedAmount = parseFloat(txForm.amount);
    if (isNaN(parsedAmount) || (txForm.type === 'Transfer' ? parsedAmount <= 0 : parsedAmount === 0)) return;

    let cascadeGroupId: string | undefined;

    if (txForm.type === 'Transfer') {
      if (!txForm.toAccountId) {
        runAlert("Please select a target 'Transfer To' account.", "Validation Error");
        return;
      }
      if (txForm.accountId === txForm.toAccountId) {
        runAlert("'Transfer From' and 'Transfer To' accounts cannot be the same.", "Validation Error");
        return;
      }
      const parsedToAmount = parseFloat(txForm.toAmount || txForm.amount);
      if (isNaN(parsedToAmount) || parsedToAmount <= 0) {
        runAlert("Please enter a valid amount for the destination account.", "Validation Error");
        return;
      }

      const fromAccountObj = accounts.find(a => a.id === txForm.accountId);
      const toAccountObj = accounts.find(a => a.id === txForm.toAccountId);
      const defaultDesc = (fromAccountObj && toAccountObj) ? `${fromAccountObj.name} : ${toAccountObj.name}` : 'Transfer';
      const finalDescription = txForm.description.trim() || defaultDesc;

      const isTargetMutualFund = isMutualFundAccount(toAccountObj);
      const isSourceMutualFund = isMutualFundAccount(fromAccountObj);
      const isMutualFundTransfer = isTargetMutualFund || isSourceMutualFund;
      const isTargetSukuk = isSukukAccount(toAccountObj);
      const isTargetOnline = isOnlineInvestmentAccount(toAccountObj);
      const isSourceOnline = isOnlineInvestmentAccount(fromAccountObj);
      const groupUUID = (editingTx && editingTx.transferGroupId)
        ? editingTx.transferGroupId
        : (isMutualFundTransfer && inheritedMfGroupId
            ? inheritedMfGroupId
            : (isTargetSukuk && inheritedSukukGroupId
                ? inheritedSukukGroupId
                : (isTargetOnline && inheritedOnlineGroupId
                    ? inheritedOnlineGroupId
                    : (inheritedOnlineInstallmentTxId || inheritedOnlineGroupId || crypto.randomUUID()))));
      cascadeGroupId = groupUUID;

      const isTargetBo = isBoAccount(toAccountObj);
      const isSourceBo = isBoAccount(fromAccountObj);
      const isDseTransfer = (isTargetBo || isSourceBo) && syncToDse;
      const dseTransactionType: 'Deposit' | 'Withdrawal' = isSourceBo ? 'Withdrawal' : 'Deposit';
      const dseAmount = isSourceBo ? parsedAmount : parsedToAmount;

      let dseTxId = (editingTx && editingTx.dseTxId)
        ? editingTx.dseTxId
        : (isDseTransfer ? (inheritedDseTxId || crypto.randomUUID()) : undefined);

      if (isDseTransfer) {
        if (!dseTxId) dseTxId = inheritedDseTxId || crypto.randomUUID();
        const otherName = isSourceBo
          ? (toAccountObj?.name || 'Account')
          : (fromAccountObj?.name || 'Account');
        const defaultDseNote = isSourceBo
          ? `Transfer to ${otherName}`
          : `Transfer from ${otherName}`;
        const dseNote = txForm.description.trim() || defaultDseNote;

        syncDseDepositTransaction({
          id: dseTxId,
          date: txForm.date,
          amount: dseAmount,
          notes: dseNote,
          linkedIeGroupId: groupUUID,
          transactionType: dseTransactionType,
          isDelete: false
        });
      } else if (editingTx?.dseTxId) {
        syncDseDepositTransaction({
          id: editingTx.dseTxId,
          date: '',
          amount: 0,
          notes: '',
          isDelete: true
        });
        dseTxId = undefined;
      }

      if (isTargetMutualFund || isSourceMutualFund) {
        syncMutualFundLinkedTransfer(groupUUID, { date: txForm.date, amount: isSourceMutualFund ? parsedAmount : parsedToAmount });
      }
      if (isTargetSukuk) {
        syncSukukLinkedTransfer(groupUUID, { date: txForm.date, amount: parsedToAmount });
      }
      if (isTargetOnline || inheritedOnlineGroupId || state.onlineInvestments?.some(o => o.linkedIeGroupId === groupUUID || (editingTx?.transferGroupId && o.linkedIeGroupId === editingTx.transferGroupId))) {
        syncOnlineLinkedTransfer(groupUUID, { date: txForm.date, amount: isSourceOnline ? parsedAmount : parsedToAmount });
      }
      const linkedOnlineInstTxId = inheritedOnlineInstallmentTxId || editingTx?.onlineTxId || (state.onlineInvestments?.some(o => o.installments?.some(i => i.linkedIncomeTxId === groupUUID || i.linkedIncomeTxId === editingTx?.id || (editingTx?.transferGroupId && i.linkedIncomeTxId === editingTx.transferGroupId))) ? (inheritedOnlineInstallmentTxId || editingTx?.onlineTxId || editingTx?.id || groupUUID) : undefined);
      if (linkedOnlineInstTxId) {
        syncOnlineInstallmentIncome(linkedOnlineInstTxId, {
          date: txForm.date,
          amount: isSourceOnline ? parsedAmount : parsedToAmount,
          description: finalDescription,
          type: 'Transfer',
          category: txForm.category,
          subCategory: txForm.subCategory
        });
      }

      // We generate two separate entries
      const transactionFrom: IncomeExpenseTransaction = {
        id: (editingTx && editingTx.transferType === 'from') ? editingTx.id : crypto.randomUUID(),
        date: txForm.date,
        amount: parsedAmount,
        type: 'Transfer',
        accountId: txForm.accountId,
        toAccountId: txForm.toAccountId,
        toAmount: parsedToAmount,
        transferGroupId: groupUUID,
        transferType: 'from',
        category: 'Transfer',
        subCategory: '',
        description: finalDescription,
        dseTxId: isDseTransfer ? dseTxId : undefined,
        autoSyncDse: (isTargetBo || isSourceBo) ? syncToDse : undefined,
        onlineTxId: linkedOnlineInstTxId
      };

      const transactionTo: IncomeExpenseTransaction = {
        id: (editingTx && editingTx.transferType === 'to') ? editingTx.id : crypto.randomUUID(),
        date: txForm.date,
        amount: parsedToAmount,
        type: 'Transfer',
        accountId: txForm.toAccountId,
        toAccountId: txForm.accountId,
        toAmount: parsedAmount,
        transferGroupId: groupUUID,
        transferType: 'to',
        category: 'Transfer',
        subCategory: '',
        description: finalDescription,
        dseTxId: isDseTransfer ? dseTxId : undefined,
        autoSyncDse: (isTargetBo || isSourceBo) ? syncToDse : undefined,
        onlineTxId: linkedOnlineInstTxId
      };

      updateState(prev => {
        const currentList = prev.incomeExpenseTransactions || [];
        let filteredList = currentList;
        if (editingTx) {
          if (editingTx.transferGroupId) {
            filteredList = currentList.filter(t => t.transferGroupId !== editingTx.transferGroupId);
          } else {
            filteredList = currentList.filter(t => t.id !== editingTx.id);
          }
        }
        return {
          ...prev,
          incomeExpenseTransactions: [transactionFrom, transactionTo, ...filteredList]
        };
      });
    } else {
      let targetMfDividendNavTxId: string | undefined = undefined;
      let targetSukukRentNavTxId: string | undefined = undefined;
      let targetFdrProfitNavTxId: string | undefined = undefined;
      let targetFdrChargeNavTxId: string | undefined = undefined;

      const sourceAccObj = accounts.find(a => a.id === txForm.accountId);
      const isSourceBo = isBoAccount(sourceAccObj);
      const isSourceMf = isMutualFundAccount(sourceAccObj);
      const linkedRentId = inheritedSukukRentTxId || (editingTx?.sukukTxId) || (state.sukuks?.some(s => s.installments?.some(i => i.linkedIncomeTxId === editingTx?.id)) ? editingTx?.id : undefined);
      const linkedFdrProfitId = inheritedFdrProfitTxId || (editingTx?.fdrTxId) || (state.fdrs?.some(f => f.transactions?.some(t => t.linkedIncomeTxId === editingTx?.id || t.id === editingTx?.id)) ? editingTx?.id : undefined);
      const linkedOnlineInstId = inheritedOnlineInstallmentTxId || (editingTx?.onlineTxId) || (state.onlineInvestments?.some(o => o.installments?.some(i => i.linkedIncomeTxId === editingTx?.id)) ? editingTx?.id : undefined);
      const linkedMfDividendId = inheritedMfGroupId || (editingTx?.mfTxId) || (state.mutualFunds?.some(f => f.transactions?.some(t => t.linkedIeGroupId === editingTx?.id || t.id === editingTx?.id)) ? editingTx?.id : undefined);
      const isMfDividendForm = txForm.type === 'Income' && (
        isSourceMf ||
        !!linkedMfDividendId ||
        (txForm.category?.toLowerCase().includes('finance') && txForm.subCategory?.toLowerCase().includes('dividend') && isSourceMf)
      );

      const isCapitalGainForm = txForm.category === 'Capital Gain' || txForm.subCategory === 'Stocks Capital Gain';
      const isDseDividendForm = txForm.type === 'Income' && !isCapitalGainForm && !isMfDividendForm && !isSourceMf && !linkedRentId && !linkedFdrProfitId && !linkedOnlineInstId && (
        isSourceBo ||
        !!inheritedDseTxId ||
        !!editingTx?.dseTxId
      );
      const isDseSellForm = txForm.type === 'Income' && isCapitalGainForm && !isMfDividendForm && (isSourceBo || !!inheritedDseTxId || !!editingTx?.dseTxId);
      const isDseChargeForm = txForm.type === 'Expense' && !isSourceMf && !linkedMfDividendId && (
        isSourceBo ||
        !!inheritedDseTxId ||
        !!editingTx?.dseTxId
      );
      const isDseSyncCandidate = (isDseDividendForm || isDseSellForm || isDseChargeForm) && !isMfDividendForm && !isSourceMf;
      const linkedDseId = inheritedDseTxId || editingTx?.dseTxId || (isDseSyncCandidate && syncToDse ? crypto.randomUUID() : undefined);
      const isMfAccountDividendForm = isMfDividendForm && isMutualFundAccount(accounts.find(a => a.id === txForm.accountId));
      const txId = editingTx ? editingTx.id : (linkedMfDividendId || linkedDseId || linkedRentId || linkedFdrProfitId || linkedOnlineInstId || crypto.randomUUID());
      const isSukukRentForm = txForm.type === 'Income' && (
        txForm.subCategory?.toLowerCase() === 'sukuk rent' ||
        (txForm.category?.toLowerCase().includes('finance') && txForm.subCategory?.toLowerCase().includes('sukuk')) ||
        !!linkedRentId
      );
      const isFdrProfitForm = txForm.type === 'Income' && (
        txForm.subCategory?.toLowerCase().includes('fdr') ||
        txForm.subCategory?.toLowerCase().includes('provisional profit') ||
        txForm.subCategory?.toLowerCase().includes('fixed deposit') ||
        (txForm.category?.toLowerCase().includes('finance') && (txForm.subCategory?.toLowerCase().includes('profit') || txForm.subCategory?.toLowerCase().includes('interest'))) ||
        !!linkedFdrProfitId
      );
      const isFdrChargeForm = txForm.type === 'Expense' && (
        !!linkedFdrProfitId ||
        (
          isFdrAccount(accounts.find(a => a.id === txForm.accountId)) &&
          txForm.category === 'Other Expense' &&
          (txForm.subCategory === 'Account Related Fees' || txForm.subCategory === 'Other Bank Charges' || txForm.subCategory?.toLowerCase().includes('tds'))
        )
      );
      const isOnlineInstallmentForm = txForm.type === 'Income' && (
        !!linkedOnlineInstId
      );

      if (!editingTx) {
        if (isMfAccountDividendForm) targetMfDividendNavTxId = txId;
        else if (isSukukRentForm) targetSukukRentNavTxId = txId;
        else if (isFdrProfitForm) targetFdrProfitNavTxId = txId;
        else if (isFdrChargeForm) targetFdrChargeNavTxId = txId;
      }

      const transactionData: IncomeExpenseTransaction = {
        id: txId,
        date: txForm.date,
        amount: parsedAmount,
        type: txForm.type as any,
        accountId: txForm.accountId,
        category: txForm.category,
        subCategory: txForm.subCategory,
        description: txForm.description,
        mfTxId: isMfDividendForm ? (linkedMfDividendId || txId) : (editingTx?.mfTxId || undefined),
        dseTxId: (isDseSyncCandidate && syncToDse) ? linkedDseId : (editingTx?.dseTxId || undefined),
        autoSyncDse: isDseSyncCandidate ? syncToDse : undefined,
        sukukTxId: linkedRentId || undefined,
        fdrTxId: linkedFdrProfitId || undefined,
        onlineTxId: isOnlineInstallmentForm ? (linkedOnlineInstId || txId) : undefined
      };

      if (isDseSyncCandidate && syncToDse && linkedDseId) {
        syncDseDepositTransaction({
          id: linkedDseId,
          date: txForm.date,
          amount: parsedAmount,
          notes: txForm.description || (isDseSellForm ? 'Stocks P&L' : isDseChargeForm ? 'BO Account Charge' : 'Cash Dividend'),
          transactionType: isDseSellForm ? 'Sell' : isDseChargeForm ? 'Charge' : 'Dividend',
          isDelete: false
        });
      } else if (editingTx?.dseTxId && (!isDseSyncCandidate || !syncToDse)) {
        syncDseDepositTransaction({
          id: editingTx.dseTxId,
          date: '',
          amount: 0,
          notes: '',
          isDelete: true
        });
      }

      if (isMfDividendForm) {
        syncMutualFundDividendIncome(linkedMfDividendId || txId, {
          date: txForm.date,
          amount: parsedAmount,
          description: txForm.description,
          type: txForm.type,
          category: txForm.category,
          subCategory: txForm.subCategory
        });
      } else if (editingTx && (editingTx.mfTxId || state.mutualFunds?.some(f => f.transactions?.some(t => t.linkedIeGroupId === editingTx.id || t.id === editingTx.id)))) {
        syncMutualFundDividendIncome(editingTx.mfTxId || editingTx.id, null);
      }

      if (isSukukRentForm) {
        syncSukukRentIncome(linkedRentId || txId, {
          date: txForm.date,
          amount: parsedAmount,
          description: txForm.description,
          type: txForm.type,
          category: txForm.category,
          subCategory: txForm.subCategory
        });
      } else if (editingTx && (editingTx.subCategory?.toLowerCase() === 'sukuk rent' || editingTx.sukukTxId || state.sukuks?.some(s => s.installments?.some(i => i.linkedIncomeTxId === editingTx.id)))) {
        syncSukukRentIncome(editingTx.sukukTxId || editingTx.id, null);
      }

      if (isFdrProfitForm || isFdrChargeForm) {
        syncFdrProfitIncome(linkedFdrProfitId || txId, {
          date: txForm.date,
          amount: parsedAmount,
          description: txForm.description,
          type: txForm.type,
          category: txForm.category,
          subCategory: txForm.subCategory
        });
      } else if (editingTx && (editingTx.subCategory?.toLowerCase().includes('fdr') || editingTx.subCategory?.toLowerCase().includes('provisional profit') || editingTx.subCategory === 'Account Related Fees' || editingTx.subCategory === 'Other Bank Charges' || editingTx.fdrTxId || state.fdrs?.some(f => f.transactions?.some(t => t.linkedIncomeTxId === editingTx.id || t.id === editingTx.id)))) {
        syncFdrProfitIncome(editingTx.fdrTxId || editingTx.id, null);
      }

      if (isOnlineInstallmentForm) {
        syncOnlineInstallmentIncome(linkedOnlineInstId || txId, {
          date: txForm.date,
          amount: parsedAmount,
          description: txForm.description,
          type: txForm.type,
          category: txForm.category,
          subCategory: txForm.subCategory
        });
      } else if (editingTx && (editingTx.onlineTxId || state.onlineInvestments?.some(o => o.installments?.some(i => i.linkedIncomeTxId === editingTx.id || (editingTx.transferGroupId && i.linkedIncomeTxId === editingTx.transferGroupId))))) {
        syncOnlineInstallmentIncome(editingTx.onlineTxId || editingTx.transferGroupId || editingTx.id, null);
      }

      updateState(prev => {
        const currentList = prev.incomeExpenseTransactions || [];
        let filteredList = currentList;
        if (editingTx) {
          if (editingTx.transferGroupId) {
            filteredList = currentList.filter(t => t.transferGroupId !== editingTx.transferGroupId);
          } else {
            filteredList = currentList.filter(t => t.id !== editingTx.id);
          }
        }
        return {
          ...prev,
          incomeExpenseTransactions: [transactionData, ...filteredList]
        };
      });

      const returnModule = activeTab ? ('income-expense-' + activeTab) : 'income-expense-transactions';

      if (txForm.type === 'Income' && onNavigateToModule) {
        if (isFromExternalModuleRef.current) {
          isFromExternalModuleRef.current = false;
        } else if (!editingTx) {
          if (targetMfDividendNavTxId) {
            onNavigateToModule('mutual-funds', {
              date: txForm.date,
              amount: parsedAmount,
              description: txForm.description,
              targetModule: 'mutual-funds-dividend',
              linkedTxId: targetMfDividendNavTxId,
              returnModule
            });
          } else if (targetSukukRentNavTxId) {
            onNavigateToModule('sukuk', {
              date: txForm.date,
              amount: parsedAmount,
              type: 'Income',
              category: txForm.category,
              subCategory: txForm.subCategory,
              targetModule: 'sukuk-rent',
              description: txForm.description,
              linkedTxId: targetSukukRentNavTxId,
              returnModule
            });
          } else if (targetFdrProfitNavTxId) {
            onNavigateToModule('fdrs', {
              date: txForm.date,
              amount: parsedAmount,
              type: 'Income',
              category: txForm.category,
              subCategory: txForm.subCategory,
              targetModule: 'fdr-profit',
              description: txForm.description,
              linkedTxId: targetFdrProfitNavTxId,
              returnModule
            });
          }
        } else if (isMfDividendForm && linkedMfDividendId) {
          onNavigateToModule('mutual-funds', {
            date: txForm.date,
            amount: parsedAmount,
            description: txForm.description,
            targetModule: 'mutual-funds-dividend',
            linkedTxId: linkedMfDividendId,
            returnModule
          });
        }
      } else if (txForm.type === 'Expense' && onNavigateToModule) {
        if (isFromExternalModuleRef.current) {
          isFromExternalModuleRef.current = false;
        } else if (!editingTx && targetFdrChargeNavTxId) {
          onNavigateToModule('fdrs', {
            date: txForm.date,
            amount: parsedAmount,
            type: 'Expense',
            category: txForm.category,
            subCategory: txForm.subCategory,
            targetModule: 'fdr-charge',
            description: txForm.description,
            linkedTxId: targetFdrChargeNavTxId,
            returnModule
          });
        }
      }
    }

    setIsTxModalOpen(false);
    setEditingTx(null);
    setInheritedSukukRentTxId(undefined);
    setInheritedFdrProfitTxId(undefined);
    setInheritedOnlineInstallmentTxId(undefined);

    // Auto-open target module modal if transfer was to Mutual Fund, Online Investment, or Sukuk
    if (txForm.type === 'Transfer' && onNavigateToModule) {
      if (isFromExternalModuleRef.current) {
        isFromExternalModuleRef.current = false;
      } else {
        const returnModule = activeTab ? ('income-expense-' + activeTab) : 'income-expense-transactions';
        const toAccountObj = accounts.find(a => a.id === txForm.toAccountId);
        const transferDate = txForm.date;
        const parsedToAmountVal = parseFloat(txForm.toAmount || txForm.amount);
        const transferAmount = (!isNaN(parsedToAmountVal) && parsedToAmountVal > 0) ? parsedToAmountVal : parsedAmount;

        if (isMutualFundAccount(toAccountObj)) {
          onNavigateToModule('mutual-funds', { date: transferDate, amount: transferAmount, description: txForm.description, linkedTxId: cascadeGroupId, returnModule });
        } else if (isOnlineInvestmentAccount(toAccountObj)) {
          onNavigateToModule('online', { date: transferDate, amount: transferAmount, description: txForm.description, linkedTxId: cascadeGroupId, returnModule });
        } else if (isSukukAccount(toAccountObj)) {
          onNavigateToModule('sukuk', { date: transferDate, amount: transferAmount, description: txForm.description, linkedTxId: cascadeGroupId, returnModule });
        } else if (isFdrAccount(toAccountObj)) {
          onNavigateToModule('fdrs', { date: transferDate, amount: transferAmount, description: txForm.description, returnModule });
        }
      }
    } else {
      if (isFromExternalModuleRef.current) {
        isFromExternalModuleRef.current = false;
      }
    }
  };

  // Edit Transaction Trigger
  const handleEditTx = (tx: IncomeExpenseTransaction) => {
    setInheritedDseTxId(undefined);
    setInheritedSukukRentTxId(tx.sukukTxId || (state.sukuks?.some(s => s.installments?.some(i => i.linkedIncomeTxId === tx.id)) ? tx.id : undefined));
    setInheritedFdrProfitTxId(tx.fdrTxId || (state.fdrs?.some(f => f.transactions?.some(t => t.linkedIncomeTxId === tx.id || t.id === tx.id)) ? tx.id : undefined));
    setInheritedOnlineInstallmentTxId(
      tx.onlineTxId ||
      (state.onlineInvestments?.some(o => o.installments?.some(i => i.linkedIncomeTxId === tx.id || (tx.transferGroupId && i.linkedIncomeTxId === tx.transferGroupId)))
        ? (tx.onlineTxId || tx.id || tx.transferGroupId)
        : undefined)
    );
    setInheritedOnlineGroupId(
      (tx.transferGroupId && state.onlineInvestments?.some(o => o.linkedIeGroupId === tx.transferGroupId))
        ? tx.transferGroupId
        : (state.onlineInvestments?.some(o => o.linkedIeGroupId === tx.id || o.id === tx.id) ? (tx.transferGroupId || tx.id) : undefined)
    );
    setInheritedMfGroupId(
      tx.mfTxId ||
      ((tx.transferGroupId && state.mutualFunds?.some(f => f.transactions?.some(t => t.linkedIeGroupId === tx.transferGroupId)))
        ? tx.transferGroupId
        : (state.mutualFunds?.some(f => f.transactions?.some(t => t.linkedIeGroupId === tx.id || t.id === tx.id)) ? tx.id : undefined))
    );
    if (tx.id.startsWith('init-bal-')) {
      const acc = accounts.find(a => a.id === tx.accountId);
      if (acc) {
        handleEditAcc(acc);
      }
      return;
    }
    setEditingTx(tx);
    if (tx.type === 'Transfer' && tx.transferType === 'to') {
      const targetAcc = accounts.find(a => a.id === tx.accountId);
      if (isBoAccount(targetAcc)) {
        setSyncToDse(tx.autoSyncDse !== undefined ? tx.autoSyncDse : true);
      } else {
        setSyncToDse(true);
      }
      setTxForm({
        date: tx.date,
        type: 'Transfer',
        accountId: tx.toAccountId || '',
        toAccountId: tx.accountId,
        category: 'Transfer',
        subCategory: '',
        amount: tx.toAmount !== undefined ? String(tx.toAmount) : String(tx.amount),
        toAmount: String(tx.amount),
        description: tx.description
      });
    } else {
      const targetAcc = accounts.find(a => a.id === (tx.toAccountId || accounts.find(a => a.id !== tx.accountId)?.id || tx.accountId));
      const sourceAcc = accounts.find(a => a.id === tx.accountId);
      const isMfAccountOrTx = isMutualFundAccount(sourceAcc) || isMutualFundAccount(targetAcc) || !!tx.mfTxId || (tx.transferGroupId && state.mutualFunds?.some(f => f.transactions?.some(t => t.linkedIeGroupId === tx.transferGroupId)));
      const isBoTransfer = !isMfAccountOrTx && (isBoAccount(targetAcc) || isBoAccount(sourceAcc));
      const isBoIncome = !isMfAccountOrTx && tx.type === 'Income' && (isBoAccount(sourceAcc) || !!tx.dseTxId);
      if (tx.type === 'Transfer' && isBoTransfer) {
        setSyncToDse(tx.autoSyncDse !== undefined ? tx.autoSyncDse : true);
      } else if (isBoIncome) {
        setSyncToDse(tx.autoSyncDse !== undefined ? tx.autoSyncDse : true);
      } else {
        setSyncToDse(false);
      }
      if (tx.dseTxId && !isMfAccountOrTx) {
        setInheritedDseTxId(tx.dseTxId);
      }
      setTxForm({
        date: tx.date,
        type: tx.type as any,
        accountId: tx.accountId,
        toAccountId: tx.toAccountId || accounts.find(a => a.id !== tx.accountId)?.id || tx.accountId,
        category: tx.category,
        subCategory: tx.subCategory || '',
        amount: String(tx.amount),
        toAmount: tx.toAmount !== undefined ? String(tx.toAmount) : String(tx.amount),
        description: tx.description
      });
    }
    setIsTxModalOpen(true);
  };

  // Delete Transaction Handler
  const handleDeleteTx = (id: string) => {
    if (id.startsWith('init-bal-')) {
      const accId = id.replace('init-bal-', '');
      const acc = accounts.find(a => a.id === accId);
      runConfirm(
        `Reset the starting initial balance of "${acc?.name || 'this account'}" to 0?`,
        () => {
          updateState(prev => ({
            ...prev,
            incomeExpenseAccounts: (prev.incomeExpenseAccounts || []).map(a => 
              a.id === accId ? { ...a, initialBalance: 0 } : a
            )
          }));
        },
        'Reset Initial Balance'
      );
      return;
    }

    const tx = transactions.find(t => t.id === id);
    const isLinkedTransfer = tx?.transferGroupId;
    const linkedTx = tx?.transferGroupId 
      ? transactions.find(t => t.transferGroupId === tx.transferGroupId && t.dseTxId)
      : null;
    const targetDseId = tx?.dseTxId || linkedTx?.dseTxId;
    const isMfDividendLink = !!tx && !!(
      tx.mfTxId ||
      (tx.type === 'Income' && tx.subCategory?.toLowerCase().includes('dividend')) ||
      state.mutualFunds?.some(f => f.transactions?.some(t2 => t2.linkedIeGroupId === tx.id || t2.id === tx.id || (tx.mfTxId && t2.linkedIeGroupId === tx.mfTxId)))
    );
    const isSevereDelete = !!isLinkedTransfer || isMfDividendLink;

    const msg = isLinkedTransfer 
      ? 'You are about to delete a linked transfer transaction. This will permanently delete BOTH the source and destination ledger entries across your accounts.'
      : isMfDividendLink
        ? 'You are about to delete a Mutual Fund dividend entry. This will permanently delete this ledger entry AND the linked dividend transaction in the Mutual Funds module.'
        : 'Delete this transaction? This action is immediate and permanent.';

    runConfirm(
      msg,
      () => {
        if (targetDseId) {
          syncDseDepositTransaction({
            id: targetDseId,
            date: '',
            amount: 0,
            notes: '',
            isDelete: true
          });
        }
        if (isLinkedTransfer && tx?.transferGroupId) {
          syncMutualFundLinkedTransfer(tx.transferGroupId, null);
          syncSukukLinkedTransfer(tx.transferGroupId, null);
          syncOnlineLinkedTransfer(tx.transferGroupId, null);
        }
        if (tx && (tx.mfTxId || (tx.type === 'Income' && tx.subCategory?.toLowerCase().includes('dividend')) || state.mutualFunds?.some(f => f.transactions?.some(t => t.linkedIeGroupId === tx.id || t.id === tx.id || (tx.mfTxId && t.linkedIeGroupId === tx.mfTxId))))) {
          syncMutualFundDividendIncome(tx.mfTxId || tx.id, null);
        }
        if (tx && (tx.sukukTxId || tx.subCategory?.toLowerCase() === 'sukuk rent' || state.sukuks?.some(s => s.installments?.some(i => i.linkedIncomeTxId === tx.id)))) {
          syncSukukRentIncome(tx.sukukTxId || tx.id, null);
        }
        if (tx && (tx.fdrTxId || tx.subCategory?.toLowerCase().includes('fdr') || tx.subCategory?.toLowerCase().includes('provisional profit') || state.fdrs?.some(f => f.transactions?.some(t => t.linkedIncomeTxId === tx.id || t.id === tx.id)))) {
          syncFdrProfitIncome(tx.fdrTxId || tx.id, null);
        }
        if (tx && (tx.onlineTxId || state.onlineInvestments?.some(o => o.installments?.some(i => i.linkedIncomeTxId === tx.id || i.linkedIncomeTxId === tx.onlineTxId || (tx.transferGroupId && i.linkedIncomeTxId === tx.transferGroupId))))) {
          syncOnlineInstallmentIncome(tx.onlineTxId || tx.id || tx.transferGroupId!, null);
        }
        updateState(prev => {
          const currentList = prev.incomeExpenseTransactions || [];
          const updatedList = isLinkedTransfer
            ? currentList.filter(t => t.transferGroupId !== tx.transferGroupId)
            : currentList.filter(t => t.id !== id);
          return {
            ...prev,
            incomeExpenseTransactions: updatedList
          };
        });
      },
      isLinkedTransfer
        ? 'Severe Warning: Deleting Linked Transfer'
        : isMfDividendLink
          ? 'Severe Warning: Deleting Linked Mutual Fund Dividend'
          : 'Confirm Delete',
      {
        variant: isSevereDelete ? 'severe' : 'danger',
        confirmLabel: isLinkedTransfer ? 'Delete Both Transfer Records' : (isMfDividendLink ? 'Delete Dividend & Linked Record' : 'Delete'),
        details: isLinkedTransfer ? [
          `Source account entry (${accounts.find(a => a.id === tx?.accountId)?.name || 'Source'})`,
          `Destination account entry (${accounts.find(a => a.id === tx?.toAccountId)?.name || 'Destination'})`,
          'All linked external investment and DSE ledger synchronization records'
        ] : isMfDividendLink ? [
          'This Income & Expense ledger entry',
          'The linked dividend transaction in the Mutual Funds module'
        ] : undefined
      }
    );
  };
  // Batch Delete Transactions
  const handleBatchDelete = () => {
    if (selectedTxIds.length === 0) return;

    const currentList = state.incomeExpenseTransactions || [];
    const linkedCount = currentList.filter(
      t => selectedTxIds.includes(t.id) && t.transferGroupId
    ).length;
    const hasInitBal = selectedTxIds.some(id => id.startsWith('init-bal-'));
    const isSevere = linkedCount > 0 || hasInitBal;

    const msg = isSevere
      ? `Your selection contains ${linkedCount > 0 ? `${linkedCount} linked transfer record(s)` : ''}${hasInitBal ? (linkedCount > 0 ? ' and ' : '') + 'account initial balance(s)' : ''}. Deleting will permanently remove both source and destination records for all linked transfers and reset account balances.`
      : `Are you sure you want to delete ${selectedTxIds.length} selected transactions? This action is immediate and permanent.`;

    runConfirm(
      msg,
      () => {
        const initialBalAccIds = selectedTxIds
          .filter(id => id.startsWith('init-bal-'))
          .map(id => id.replace('init-bal-', ''));

        const nonInitSelectedIdsForMf = selectedTxIds.filter(id => !id.startsWith('init-bal-'));
        const selectedGroupIdsForMf = transactions
          .filter(t => nonInitSelectedIdsForMf.includes(t.id) && t.transferGroupId)
          .map(t => t.transferGroupId as string);

        updateState(prev => {
          let accs = prev.incomeExpenseAccounts || [];
          if (initialBalAccIds.length > 0) {
            accs = accs.map(a => initialBalAccIds.includes(a.id) ? { ...a, initialBalance: 0 } : a);
          }

          const transactionsList = prev.incomeExpenseTransactions || [];
          const nonInitSelectedIds = selectedTxIds.filter(id => !id.startsWith('init-bal-'));
          const selectedGroupIds = transactionsList
            .filter(t => nonInitSelectedIds.includes(t.id) && t.transferGroupId)
            .map(t => t.transferGroupId as string);

          const dseIdsToClean = transactionsList
            .filter(t => (nonInitSelectedIds.includes(t.id) || (t.transferGroupId && selectedGroupIds.includes(t.transferGroupId))) && t.dseTxId)
            .map(t => t.dseTxId as string);

          dseIdsToClean.forEach(dseId => {
            syncDseDepositTransaction({ id: dseId, date: '', amount: 0, notes: '', isDelete: true });
          });

          nonInitSelectedIds.forEach(selectedId => {
            const t = transactionsList.find(tx => tx.id === selectedId);
            if (t && (t.mfTxId || (t.type === 'Income' && t.subCategory?.toLowerCase().includes('dividend')) || prev.mutualFunds?.some(f => f.transactions?.some(inst => inst.linkedIeGroupId === t.id || inst.id === t.id || (t.mfTxId && inst.linkedIeGroupId === t.mfTxId))))) {
              syncMutualFundDividendIncome(t.mfTxId || t.id, null);
            }
            if (t && (t.sukukTxId || t.subCategory?.toLowerCase() === 'sukuk rent' || prev.sukuks?.some(s => s.installments?.some(i => i.linkedIncomeTxId === t.id)))) {
              syncSukukRentIncome(t.sukukTxId || t.id, null);
            }
            if (t && (t.fdrTxId || t.subCategory?.toLowerCase().includes('fdr') || t.subCategory?.toLowerCase().includes('provisional profit') || prev.fdrs?.some(f => f.transactions?.some(inst => inst.linkedIncomeTxId === t.id || inst.id === t.id)))) {
              syncFdrProfitIncome(t.fdrTxId || t.id, null);
            }
            if (t && (t.onlineTxId || prev.onlineInvestments?.some(o => o.installments?.some(i => i.linkedIncomeTxId === t.id || i.linkedIncomeTxId === t.onlineTxId || (t.transferGroupId && i.linkedIncomeTxId === t.transferGroupId))))) {
              syncOnlineInstallmentIncome(t.onlineTxId || t.id || t.transferGroupId!, null);
            }
          });

          const updatedList = transactionsList.filter(t => {
            if (nonInitSelectedIds.includes(t.id)) return false;
            if (t.transferGroupId && selectedGroupIds.includes(t.transferGroupId)) return false;
            return true;
          });

          return {
            ...prev,
            incomeExpenseAccounts: accs,
            incomeExpenseTransactions: updatedList
          };
        });

        selectedGroupIdsForMf.forEach(groupId => {
          syncMutualFundLinkedTransfer(groupId, null);
          syncSukukLinkedTransfer(groupId, null);
          syncOnlineLinkedTransfer(groupId, null);
        });

        setSelectedTxIds([]);
      },
      isSevere ? 'Severe Warning: Batch Delete Contains Linked Transfers' : 'Confirm Batch Delete',
      {
        variant: isSevere ? 'severe' : 'danger',
        confirmLabel: isSevere ? `Delete All (${selectedTxIds.length} Records & Linked Legs)` : 'Delete Selected',
        details: isSevere ? [
          `${selectedTxIds.length} Total Selected Records`,
          linkedCount > 0 ? `${linkedCount} Linked Transfer Leg(s) (will remove corresponding pair in other accounts)` : null,
          hasInitBal ? 'Starting initial balance reset for selected account(s)' : null,
          'All associated balance reconciliations across accounts'
        ] : undefined
      }
    );
  };

  // ─── Settings - Accounts Configuration state ──────────────────────────────
  const [isAccModalOpen, setIsAccModalOpen] = useState(false);
  const [editingAcc, setEditingAcc] = useState<BankAccount | null>(null);
  const [expandedGroups, setExpandedGroups] = useState<Record<string, boolean>>({});
  const [expandedCategories, setExpandedCategories] = useState<Record<string, boolean>>({});
  const [accForm, setAccForm] = useState({
    name: '',
    currency: 'BDT' as 'BDT' | 'LYD' | 'USD',
    initialBalance: '',
    initialDate: new Date().toISOString().split('T')[0],
    isParent: false,
    parentId: ''
  });

  const handleSaveAcc = (e: React.FormEvent) => {
    e.preventDefault();
    if (!accForm.name) return;

    const parsedBalance = parseFloat(accForm.initialBalance) || 0;
    const accData: BankAccount = {
      id: editingAcc ? editingAcc.id : 'acc-' + crypto.randomUUID().slice(0, 8),
      name: accForm.name,
      currency: accForm.currency,
      initialBalance: parsedBalance,
      initialDate: accForm.initialDate || new Date().toISOString().split('T')[0],
      isParent: accForm.isParent,
      parentId: accForm.parentId || undefined
    };

    updateState(prev => {
      const currentList = prev.incomeExpenseAccounts || [];
      const updatedList = editingAcc
        ? currentList.map(a => a.id === editingAcc.id ? accData : a)
        : [...currentList, accData];
      return {
        ...prev,
        incomeExpenseAccounts: updatedList
      };
    });

    setIsAccModalOpen(false);
    setEditingAcc(null);
  };

  const handleEditAcc = (acc: BankAccount) => {
    setEditingAcc(acc);
    setAccForm({
      name: acc.name,
      currency: acc.currency,
      initialBalance: String(acc.initialBalance),
      initialDate: acc.initialDate || new Date().toISOString().split('T')[0],
      isParent: !!acc.isParent,
      parentId: acc.parentId || ''
    });
    setIsAccModalOpen(true);
  };

  const handleDeleteAcc = (id: string) => {
    const associatedTxs = transactions.filter(t => t.accountId === id);
    if (associatedTxs.length > 0) {
      runAlert(
        `Cannot delete this account. It has ${associatedTxs.length} associated transactions. Delete those transactions first.`,
        'Cannot Delete Account'
      );
      return;
    }
    runConfirm(
      'Delete this account from your profile?',
      () => {
        updateState(prev => ({
          ...prev,
          incomeExpenseAccounts: (prev.incomeExpenseAccounts || []).filter(a => a.id !== id)
        }));
      },
      'Delete Account'
    );
  };

  // ─── Settings - Categories Configuration state ────────────────────────────
  const [isCatModalOpen, setIsCatModalOpen] = useState(false);
  const [editingCat, setEditingCat] = useState<AppCategory | null>(null);
  const [catForm, setCatForm] = useState({
    name: '',
    type: 'Expense' as 'Income' | 'Expense',
    subCategories: ''
  });

  const handleSaveCat = (e: React.FormEvent) => {
    e.preventDefault();
    if (!catForm.name.trim()) {
      runAlert('Please enter a category name.', 'Validation Error');
      return;
    }

    const subCatsArray = catForm.subCategories
      .split(',')
      .map(s => s.trim())
      .filter(s => s.length > 0);

    const updatedCategories = [...categories];

    if (editingCat) {
      const idx = updatedCategories.findIndex(c => c.id === editingCat.id);
      if (idx !== -1) {
        updatedCategories[idx] = {
          ...editingCat,
          name: catForm.name.trim(),
          type: catForm.type,
          subCategories: subCatsArray
        };
      }
    } else {
      const newCat: AppCategory = {
        id: `cat-${Date.now()}`,
        name: catForm.name.trim(),
        type: catForm.type,
        subCategories: subCatsArray
      };
      updatedCategories.push(newCat);
    }

    updateState(prev => ({
      ...prev,
      incomeExpenseCategories: updatedCategories
    }));

    setIsCatModalOpen(false);
    setEditingCat(null);
  };

  const handleEditCat = (cat: AppCategory) => {
    setEditingCat(cat);
    setCatForm({
      name: cat.name,
      type: cat.type,
      subCategories: cat.subCategories.join(', ')
    });
    setIsCatModalOpen(true);
  };

  const handleDeleteCat = (id: string) => {
    const catToDelete = categories.find(c => c.id === id);
    if (!catToDelete) return;

    const associatedTxs = transactions.filter(t => t.category === catToDelete.name);
    if (associatedTxs.length > 0) {
      runAlert(
        `Cannot delete this category. It has ${associatedTxs.length} associated transactions. Delete those transactions first.`,
        'Cannot Delete Category'
      );
      return;
    }

    runConfirm(
      `Delete the "${catToDelete.name}" category?`,
      () => {
        updateState(prev => ({
          ...prev,
          incomeExpenseCategories: (prev.incomeExpenseCategories || []).filter(c => c.id !== id)
        }));
      },
      'Delete Category'
    );
  };

  // ─── Export / Import ───────────────────────────────────────────────────────
  const handleExportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();
      
      const txRows = transactions.map(t => {
        const acc = accounts.find(a => a.id === t.accountId);
        return {
          ID: t.id,
          Date: t.date,
          Type: t.type,
          Account: acc ? acc.name : 'N/A',
          Currency: acc ? acc.currency : 'BDT',
          Category: t.category,
          SubCategory: t.subCategory || '',
          Amount: t.amount,
          Description: t.description
        };
      });
      const wsTx = XLSX.utils.json_to_sheet(txRows);
      XLSX.utils.book_append_sheet(wb, wsTx, 'Transactions');
      
      const accRows = accounts.map(a => ({
        ID: a.id,
        Name: a.name,
        Currency: a.currency,
        InitialDate: a.initialDate || '',
        InitialBalance: a.initialBalance,
        CurrentBalance: accountBalances[a.id] || 0
      }));
      const wsAcc = XLSX.utils.json_to_sheet(accRows);
      XLSX.utils.book_append_sheet(wb, wsAcc, 'Accounts');
      
      XLSX.writeFile(wb, `FinTrackPro_IncomeExpense_${new Date().toISOString().split('T')[0]}.xlsx`);
    } catch (e) {
      runAlert('Failed to export. Please try again.', 'Export Failed');
    }
  };

  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const binary = evt.target?.result as string;
        const workbook = XLSX.read(binary, { type: 'binary' });
        
        let importedTxs: IncomeExpenseTransaction[] = [];
        let importedAccs: BankAccount[] = [];

        const getRowVal = (row: any, ...keys: string[]): any => {
          const lowercaseKeys = keys.map(k => k.toLowerCase().trim());
          const rowKeys = Object.keys(row);
          for (const rk of rowKeys) {
            const cleanRk = rk.toLowerCase().trim();
            if (lowercaseKeys.includes(cleanRk)) {
              return row[rk];
            }
          }
          return undefined;
        };

        const parseExcelDate = (val: any): string => {
          if (val === undefined || val === null || val === '') {
            return new Date().toISOString().split('T')[0];
          }

          // 1. If it's already a JS Date object
          if (val instanceof Date) {
            const y = val.getFullYear();
            const m = String(val.getMonth() + 1).padStart(2, '0');
            const d = String(val.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
          }

          // 2. If it's a number (Excel date serial)
          const num = Number(val);
          if (!isNaN(num) && num > 10000 && num < 100000) {
            const date = new Date((num - 25569) * 86400 * 1000);
            const y = date.getUTCFullYear();
            const m = String(date.getUTCMonth() + 1).padStart(2, '0');
            const d = String(date.getUTCDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
          }

          // 3. String date parsing
          const str = String(val).trim();
          const yyyymmdd = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
          if (yyyymmdd) {
            const y = yyyymmdd[1];
            const m = yyyymmdd[2].padStart(2, '0');
            const d = yyyymmdd[3].padStart(2, '0');
            return `${y}-${m}-${d}`;
          }

          const dateParts = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
          if (dateParts) {
            const part1 = parseInt(dateParts[1], 10);
            const part2 = parseInt(dateParts[2], 10);
            const year = dateParts[3];

            let month = part1;
            let day = part2;

            if (part1 > 12) {
              month = part2;
              day = part1;
            }

            const mStr = String(month).padStart(2, '0');
            const dStr = String(day).padStart(2, '0');
            return `${year}-${mStr}-${dStr}`;
          }

          try {
            const parsed = new Date(str);
            if (!isNaN(parsed.getTime())) {
              const y = parsed.getFullYear();
              const m = String(parsed.getMonth() + 1).padStart(2, '0');
              const d = String(parsed.getDate()).padStart(2, '0');
              return `${y}-${m}-${d}`;
            }
          } catch {}

          return new Date().toISOString().split('T')[0];
        };

        if (workbook.SheetNames.includes('Accounts')) {
          const sheet = workbook.Sheets['Accounts'];
          const rows = XLSX.utils.sheet_to_json<any>(sheet);
          importedAccs = rows.map(r => ({
            id: String(getRowVal(r, 'id', 'ID') || 'acc-' + Math.random().toString(36).substr(2, 9)),
            name: String(getRowVal(r, 'name', 'Name', 'Account Name') || 'Imported Account'),
            currency: (getRowVal(r, 'currency', 'Currency') || 'BDT') as any,
            initialDate: parseExcelDate(getRowVal(r, 'initialdate', 'Initial Date', 'initial date', 'InitialDate', 'StartDate', 'Date')),
            initialBalance: parseFloat(String(getRowVal(r, 'initialbalance', 'Initial Balance', 'initial balance', 'InitialBalance') || 0))
          }));
        }

        const mapRowToTx = (r: any): IncomeExpenseTransaction => {
          const accName = String(getRowVal(r, 'account', 'AccountId', 'accountId', 'Account Name') || '').trim();
          const accFound = accounts.find(
            a => a.name.toLowerCase() === accName.toLowerCase() || a.id.toLowerCase() === accName.toLowerCase()
          ) || importedAccs.find(
            a => a.name.toLowerCase() === accName.toLowerCase() || a.id.toLowerCase() === accName.toLowerCase()
          );

          const destName = String(getRowVal(r, 'toaccount', 'ToAccount', 'toaccountid', 'ToAccountId', 'to account') || '').trim();
          const destFound = accounts.find(
            a => a.name.toLowerCase() === destName.toLowerCase() || a.id.toLowerCase() === destName.toLowerCase()
          ) || importedAccs.find(
            a => a.name.toLowerCase() === destName.toLowerCase() || a.id.toLowerCase() === destName.toLowerCase()
          );

          // Boolean income check
          const incomeValRaw = getRowVal(r, 'income');
          const isIncomePresent = incomeValRaw !== undefined;
          let isIncomeVal = false;
          if (typeof incomeValRaw === 'boolean') {
            isIncomeVal = incomeValRaw;
          } else if (incomeValRaw !== undefined && incomeValRaw !== null) {
            isIncomeVal = String(incomeValRaw).trim().toLowerCase() === 'true';
          }

          let finalType: 'Income' | 'Expense' | 'Transfer' | 'Loan';
          if (isIncomePresent) {
            finalType = isIncomeVal ? 'Income' : 'Expense';
          } else {
            const rawType = String(getRowVal(r, 'type') || '').trim();
            if (rawType.toLowerCase() === 'income') finalType = 'Income';
            else if (rawType.toLowerCase() === 'transfer') finalType = 'Transfer';
            else if (rawType.toLowerCase() === 'loan') finalType = 'Loan';
            else finalType = 'Expense';
          }

          // Combined display of title and note under Narrations / Remarks (description)
          const titlePart = String(getRowVal(r, 'title') || '').trim();
          const notePart = String(getRowVal(r, 'note') || '').trim();
          const descPart = String(getRowVal(r, 'description') || '').trim();
          
          let finalDescription = '';
          if (titlePart || notePart) {
            finalDescription = [titlePart, notePart].filter(Boolean).join(', ');
          } else {
            finalDescription = descPart;
          }

          // Absolute amount used inside ledger
          const rawAmtInput = getRowVal(r, 'amount');
          const parsedAmt = parseFloat(String(rawAmtInput !== undefined ? rawAmtInput : 0));
          const finalAmount = Math.abs(parsedAmt);

          const rawDate = getRowVal(r, 'date');
          const finalDate = parseExcelDate(rawDate);

          const rawCategory = String(getRowVal(r, 'category name', 'category', 'Category Name') || 'Other Expense').trim();
          const rawSubCategory = String(getRowVal(r, 'subcategory name', 'sub category', 'subcategory', 'Subcategory Name') || '').trim();

          const toAmountRaw = getRowVal(r, 'toamount', 'to amount');

          return {
            id: String(getRowVal(r, 'id') || crypto.randomUUID()),
            date: finalDate,
            type: finalType,
            accountId: accFound ? accFound.id : (accounts[0]?.id || ''),
            toAccountId: destFound ? destFound.id : undefined,
            toAmount: toAmountRaw !== undefined ? parseFloat(String(toAmountRaw)) : undefined,
            category: rawCategory,
            subCategory: rawSubCategory,
            amount: finalAmount,
            description: finalDescription
          };
        };

        if (workbook.SheetNames.includes('Transactions')) {
          const sheet = workbook.Sheets['Transactions'];
          const rows = XLSX.utils.sheet_to_json<any>(sheet);
          importedTxs = rows.map(mapRowToTx);
        } else if (workbook.SheetNames.length > 0) {
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json<any>(firstSheet);
          importedTxs = rows.map(mapRowToTx);
        }

        if (importedTxs.length === 0 && importedAccs.length === 0) {
          throw new Error('No valid accounts or transactions found in template sheets.');
        }

        updateState(prev => {
          const finalAccs = importedAccs.length > 0 ? importedAccs : prev.incomeExpenseAccounts || [];
          const finalTxs = importedTxs.length > 0 ? [...importedTxs, ...(prev.incomeExpenseTransactions || [])] : prev.incomeExpenseTransactions || [];
          return {
            ...prev,
            incomeExpenseAccounts: finalAccs,
            incomeExpenseTransactions: finalTxs
          };
        });

        // Set date view to the month of the imported transactions so they are visible immediately
        if (importedTxs.length > 0) {
          const sampleDateStr = importedTxs[0].date;
          const parts = sampleDateStr.split('-');
          if (parts.length === 3) {
            const yr = parseInt(parts[0], 10);
            const mo = parseInt(parts[1], 10) - 1;
            if (!isNaN(yr) && !isNaN(mo) && mo >= 0 && mo < 12) {
              setThisMonthDate(new Date(yr, mo, 1));
              setRangeType('this');
            }
          }
        }

        runAlert(`Loaded successfully! Imported ${importedAccs.length} accounts and ${importedTxs.length} transactions.`, 'Import Succeeded');
      } catch (err: any) {
        runAlert(`Failed to parse file: ${err?.message || 'Check template design'}`, 'Import Failed');
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  const handleDownloadTemplate = () => {
    try {
      const wb = XLSX.utils.book_new();

      const sampleTxs = [
        {
          ID: 'tx-sample1',
          Date: '2026-05-01',
          Type: 'Income',
          Account: 'City Bank Amex Savings',
          Category: 'Remittance Income',
          SubCategory: 'Salary',
          Amount: 85000,
          Description: 'Sample Monthly Salary Inward'
        },
        {
          ID: 'tx-sample2',
          Date: '2026-05-02',
          Type: 'Expense',
          Account: 'Cash (Wallet)',
          Category: 'Food & Clothing',
          SubCategory: '',
          Amount: 1200,
          Description: 'Weekly household kitchen shopping'
        }
      ];

      const sampleAccs = [
        { ID: 'bdt-wallet', Name: 'Cash (Wallet)', Currency: 'BDT', InitialBalance: 12000 },
        { ID: 'bdt-bank', Name: 'City Bank Amex Savings', Currency: 'BDT', InitialBalance: 150000 },
        { ID: 'lyd-cash', Name: 'Tripoli Wallet Cash', Currency: 'LYD', InitialBalance: 750 },
        { ID: 'usd-payoneer', Name: 'Payoneer Business Account', Currency: 'USD', InitialBalance: 2400 }
      ];

      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sampleTxs), 'Transactions');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sampleAccs), 'Accounts');
      XLSX.writeFile(wb, 'FinTrackPro_IncomeExpense_Template.xlsx');
    } catch {
      runAlert('Fail to download template', 'Download Failed');
    }
  };


  // ─── Transactions List Filters & Pagination ──────────────────────────────
  const [txSearch, setTxSearch] = useState('');
  const [txTypeFilter, setTxTypeFilter] = useState<'All' | 'Income' | 'Expense' | 'Transfer' | 'Loan'>('All');
  const [txAccFilter, setTxAccFilter] = useState<string>('All');
  const [txCatFilter, setTxCatFilter] = useState<string>('All');
  const [activeTxDropdown, setActiveTxDropdown] = useState<'type' | 'acc' | 'cat' | null>(null);
  const [selectedTxIds, setSelectedTxIds] = useState<string[]>([]);
  const [lastSelectedTxId, setLastSelectedTxId] = useState<string | null>(null);

  useEffect(() => {
    setSelectedTxIds([]);
    setLastSelectedTxId(null);
  }, [txSearch, txTypeFilter, txAccFilter, txCatFilter, activeTab]);

  const filteredTabTransactions = useMemo(() => {
    const query = txSearch.trim().toLowerCase();
    const cleanNumQuery = query.replace(/[৳$,\s+]/g, '').replace(/^ld\s*/i, '').replace(/^bdt\s*/i, '').replace(/^tk\.?\s*/i, '');

    return filteredTransactions.filter(t => {
      const acc = accounts.find(a => a.id === t.accountId);
      const toAcc = t.toAccountId ? accounts.find(a => a.id === t.toAccountId) : null;
      const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
      const toMult = toAcc?.currency === 'USD' ? USD_rate : toAcc?.currency === 'LYD' ? LYD_rate : 1;

      // Text fields match
      const textMatches = !query || 
        t.description?.toLowerCase().includes(query) || 
        t.category?.toLowerCase().includes(query) ||
        t.subCategory?.toLowerCase().includes(query) ||
        acc?.name.toLowerCase().includes(query) ||
        (toAcc?.name ? toAcc.name.toLowerCase().includes(query) : false);

      // Amount match
      let amountMatches = false;
      if (query) {
        const absAmt = Math.abs(t.amount);
        const bdtAmt = Math.abs(t.amount * mult);
        const formattedRaw = formatCurrency(t.amount, acc?.currency || 'BDT').toLowerCase();
        const formattedAbs = formatCurrency(absAmt, acc?.currency || 'BDT').toLowerCase();
        const formattedBdt = formatBDT(bdtAmt).toLowerCase();
        const localeStr = absAmt.toLocaleString('en-US').toLowerCase();
        const fixed2Str = absAmt.toFixed(2);
        const rawStr = t.amount.toString();
        const absStr = absAmt.toString();

        if (
          rawStr.includes(query) ||
          absStr.includes(query) ||
          formattedRaw.includes(query) ||
          formattedAbs.includes(query) ||
          formattedBdt.includes(query) ||
          localeStr.includes(query) ||
          fixed2Str.includes(query)
        ) {
          amountMatches = true;
        }

        if (!amountMatches && cleanNumQuery) {
          const bdtRoundedStr = Math.round(bdtAmt).toString();
          const bdtFixed2Str = bdtAmt.toFixed(2);

          if (
            rawStr.includes(cleanNumQuery) ||
            absStr.includes(cleanNumQuery) ||
            fixed2Str.includes(cleanNumQuery) ||
            bdtRoundedStr.includes(cleanNumQuery) ||
            bdtFixed2Str.includes(cleanNumQuery)
          ) {
            amountMatches = true;
          }

          // Also check transfer toAmount
          if (!amountMatches && t.toAmount !== undefined && t.toAmount !== null) {
            const absToAmt = Math.abs(t.toAmount);
            const toBdtAmt = Math.abs(t.toAmount * toMult);
            if (
              absToAmt.toString().includes(cleanNumQuery) ||
              absToAmt.toFixed(2).includes(cleanNumQuery) ||
              Math.round(toBdtAmt).toString().includes(cleanNumQuery) ||
              toBdtAmt.toFixed(2).includes(cleanNumQuery)
            ) {
              amountMatches = true;
            }
          }
        }
      }

      const matchesSearch = textMatches || amountMatches;
      
      const matchesType = txTypeFilter === 'All' || t.type === txTypeFilter;
      const matchesAcc = txAccFilter === 'All' || t.accountId === txAccFilter || (t.type === 'Transfer' && t.toAccountId === txAccFilter);
      const matchesCat = txCatFilter === 'All' || t.category === txCatFilter;

      return matchesSearch && matchesType && matchesAcc && matchesCat;
    });
  }, [filteredTransactions, txSearch, txTypeFilter, txAccFilter, txCatFilter, accounts, USD_rate, LYD_rate]);

  // ─── Group Transactions by Date for Mobile & Web List ─────────────────────
  const groupedTabTransactions = useMemo(() => {
    const map = new Map<string, {
      date: string;
      formattedDate: string;
      totalNetBDT: number;
      totalIncomeBDT: number;
      totalExpenseBDT: number;
      transactions: IncomeExpenseTransaction[];
    }>();

    for (const t of filteredTabTransactions) {
      if (!map.has(t.date)) {
        map.set(t.date, {
          date: t.date,
          formattedDate: formatGroupDate(t.date),
          totalNetBDT: 0,
          totalIncomeBDT: 0,
          totalExpenseBDT: 0,
          transactions: []
        });
      }
      const g = map.get(t.date)!;
      g.transactions.push(t);

      const acc = accounts.find(a => a.id === t.accountId);
      const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
      
      if (t.type === 'Income' || (t.type === 'Loan' && (t.subCategory === 'Borrowed' || t.subCategory === 'Received Lent Money'))) {
        g.totalNetBDT += (t.amount * mult);
        g.totalIncomeBDT += (t.amount * mult);
      } else if (t.type === 'Expense' || (t.type === 'Loan' && (t.subCategory === 'Lent' || t.subCategory === 'Repaid Borrowed Money'))) {
        g.totalNetBDT -= (t.amount * mult);
        g.totalExpenseBDT += (t.amount * mult);
      }
    }

    const sortedDates = Array.from(map.keys()).sort((a, b) => b.localeCompare(a));
    return sortedDates.map(d => map.get(d)!);
  }, [filteredTabTransactions, accounts, USD_rate, LYD_rate]);

  const displayedTabTransactions = useMemo(() => {
    return groupedTabTransactions.flatMap(g => g.transactions);
  }, [groupedTabTransactions]);

  const toggleSelectTx = (id: string, e?: React.MouseEvent) => {
    if (e?.shiftKey && lastSelectedTxId && lastSelectedTxId !== id) {
      const lastIdx = displayedTabTransactions.findIndex(t => t.id === lastSelectedTxId);
      const currIdx = displayedTabTransactions.findIndex(t => t.id === id);
      if (lastIdx !== -1 && currIdx !== -1) {
        const start = Math.min(lastIdx, currIdx);
        const end = Math.max(lastIdx, currIdx);
        const rangeIds = displayedTabTransactions.slice(start, end + 1).map(t => t.id);
        setSelectedTxIds(prev => Array.from(new Set([...prev, ...rangeIds])));
        setLastSelectedTxId(id);
        return;
      }
    }

    setLastSelectedTxId(id);
    if (selectedTxIds.includes(id)) {
      setSelectedTxIds(prev => prev.filter(i => i !== id));
    } else {
      setSelectedTxIds(prev => [...prev, id]);
    }
  };

  const { runningBalances, dateBalances } = useMemo(() => {
    if (txAccFilter === 'All') return { runningBalances: {}, dateBalances: {} };
    const selectedAcc = accounts.find(a => a.id === txAccFilter);
    if (!selectedAcc) return { runningBalances: {}, dateBalances: {} };

    const indexedTxs = allTransactions.map((t, idx) => ({ t, idx }));
    // Sort ascending: date comparison first, index descending to preserve oldest array elements first on tie
    indexedTxs.sort((a, b) => {
      if (a.t.date !== b.t.date) {
        return a.t.date.localeCompare(b.t.date);
      }
      if (a.t.id.startsWith('init-bal-')) return -1;
      if (b.t.id.startsWith('init-bal-')) return 1;
      return b.idx - a.idx;
    });

    const balancesMap: Record<string, number> = {};
    const dateBalancesMap: Record<string, number> = {};
    let currentBalance = 0;

    indexedTxs.forEach(({ t }) => {
      if (t.accountId !== selectedAcc.id && t.toAccountId !== selectedAcc.id) {
        return;
      }

      if (t.type === 'Income') {
        if (t.accountId === selectedAcc.id) {
          currentBalance += t.amount;
        }
      } else if (t.type === 'Expense') {
        if (t.accountId === selectedAcc.id) {
          currentBalance -= t.amount;
        }
      } else if (t.type === 'Transfer') {
        if (t.transferType) {
          if (t.accountId === selectedAcc.id) {
            if (t.transferType === 'from') {
              currentBalance -= t.amount;
            } else if (t.transferType === 'to') {
              currentBalance += t.amount;
            }
          }
        } else {
          if (t.accountId === selectedAcc.id) {
            currentBalance -= t.amount;
          }
          if (t.toAccountId === selectedAcc.id) {
            currentBalance += (t.toAmount !== undefined ? t.toAmount : t.amount);
          }
        }
      } else if (t.type === 'Loan' || t.category === 'Loan') {
        if (t.accountId === selectedAcc.id) {
          if (t.subCategory === 'Borrowed' || t.subCategory === 'Received Lent Money') {
            currentBalance += t.amount;
          } else if (t.subCategory === 'Lent' || t.subCategory === 'Repaid Borrowed Money') {
            currentBalance -= t.amount;
          }
        }
      }

      balancesMap[t.id] = currentBalance;
      dateBalancesMap[t.date] = currentBalance;
    });

    return { runningBalances: balancesMap, dateBalances: dateBalancesMap };
  }, [allTransactions, accounts, txAccFilter]);

  // Analytics Metrics
  const analyticsData = useMemo(() => {
    // 1. Generate all consecutive months from flowStartStr to flowEndStr
    const [sY, sM] = flowStartStr.split('-').map(Number);
    const [eY, eM] = flowEndStr.split('-').map(Number);
    const monthsMap: Record<string, { month: string; label: string; date: Date; income: number; expense: number; investment: number }> = {};

    let curY = sY;
    let curM = sM;
    while (curY < eY || (curY === eY && curM <= eM)) {
      const monthKey = `${curY}-${String(curM).padStart(2, '0')}`;
      const d = new Date(curY, curM - 1, 1);
      const label = d.toLocaleString('default', { month: 'short', year: '2-digit' });
      monthsMap[monthKey] = {
        month: monthKey,
        label,
        date: d,
        income: 0,
        expense: 0,
        investment: 0
      };
      curM++;
      if (curM > 12) {
        curM = 1;
        curY++;
      }
    }

    // 2. Filter transactions strictly within [flowStartStr, flowEndStr]
    allTransactions.forEach(t => {
      if (!t.date || t.date < flowStartStr || t.date > flowEndStr) return;
      // Never treat initial balances as monthly income
      if (t.category === 'Initial Balance' || t.id.startsWith('init-bal-')) return;
      // Never treat loans as income/expense
      if (t.type === 'Loan' || t.category === 'Loan') return;

      const dateParts = t.date.split('-');
      if (dateParts.length < 2) return;
      const monthKey = `${dateParts[0]}-${dateParts[1].padStart(2, '0')}`;

      if (!monthsMap[monthKey]) {
        const [yNum, mNum] = [Number(dateParts[0]), Number(dateParts[1])];
        const d = new Date(yNum, mNum - 1, 1);
        monthsMap[monthKey] = {
          month: monthKey,
          label: d.toLocaleString('default', { month: 'short', year: '2-digit' }),
          date: d,
          income: 0,
          expense: 0,
          investment: 0
        };
      }

      const acc = accounts.find(a => a.id === t.accountId);
      const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
      const bdtAmount = (t.amount || 0) * mult;

      if (t.type === 'Income') {
        monthsMap[monthKey].income += bdtAmount;
      } else if (t.type === 'Expense') {
        if (t.category === 'Investment' || t.subCategory === 'Investment') {
          monthsMap[monthKey].investment += bdtAmount;
        } else {
          monthsMap[monthKey].expense += bdtAmount;
        }
      } else if (t.type === 'Transfer') {
        if (t.transferType) {
          // Two-entry transfer model: evaluate the incoming 'to' transaction
          if (t.transferType === 'to') {
            const destAcc = accounts.find(a => a.id === t.accountId);
            if (isInvestmentAccount(destAcc, accounts)) {
              const fromTx = t.transferGroupId
                ? allTransactions.find(other => other.transferGroupId === t.transferGroupId && other.transferType === 'from')
                : null;
              const sourceAcc = fromTx ? accounts.find(a => a.id === fromTx.accountId) : null;
              if (!isInvestmentAccount(sourceAcc, accounts)) {
                monthsMap[monthKey].investment += bdtAmount;
              }
            }
          }
        } else {
          // Legacy single-entry model
          const destAcc = accounts.find(a => a.id === t.toAccountId);
          const sourceAcc = accounts.find(a => a.id === t.accountId);
          if (isInvestmentAccount(destAcc, accounts) && !isInvestmentAccount(sourceAcc, accounts)) {
            const toMult = destAcc?.currency === 'USD' ? USD_rate : destAcc?.currency === 'LYD' ? LYD_rate : 1;
            const targetBdt = (t.toAmount !== undefined ? t.toAmount : t.amount) * toMult;
            monthsMap[monthKey].investment += targetBdt;
          }
        }
      }
    });

    return Object.values(monthsMap).sort((a, b) => a.month.localeCompare(b.month));
  }, [allTransactions, accounts, USD_rate, LYD_rate, flowStartStr, flowEndStr]);

  const expenseSlices = useMemo<AnalyticsDonutSlice[]>(() => {
    const catsMap: Record<string, number> = {};
    filteredTransactions.filter(t => t.type === 'Expense').forEach(t => {
      const acc = accounts.find(a => a.id === t.accountId);
      const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
      catsMap[t.category] = (catsMap[t.category] || 0) + (t.amount * mult);
    });

    const entries = Object.entries(catsMap)
      .filter(([_, value]) => value > 0)
      .map(([name, value]) => ({ id: name, name, value }))
      .sort((a, b) => b.value - a.value);

    const total = entries.reduce((sum, e) => sum + e.value, 0);
    return entries.map((e, index) => ({
      ...e,
      pct: total > 0 ? e.value / total : 0,
      color: HOLDING_COLORS[index % HOLDING_COLORS.length],
    }));
  }, [filteredTransactions, accounts, USD_rate, LYD_rate]);

  const incomeSlices = useMemo<AnalyticsDonutSlice[]>(() => {
    const catsMap: Record<string, number> = {};
    filteredTransactions.filter(t => t.type === 'Income').forEach(t => {
      const acc = accounts.find(a => a.id === t.accountId);
      const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
      catsMap[t.category] = (catsMap[t.category] || 0) + (t.amount * mult);
    });

    const entries = Object.entries(catsMap)
      .filter(([_, value]) => value > 0)
      .map(([name, value]) => ({ id: name, name, value }))
      .sort((a, b) => b.value - a.value);

    const total = entries.reduce((sum, e) => sum + e.value, 0);
    return entries.map((e, index) => ({
      ...e,
      pct: total > 0 ? e.value / total : 0,
      color: HOLDING_COLORS[index % HOLDING_COLORS.length],
    }));
  }, [filteredTransactions, accounts, USD_rate, LYD_rate]);

  const assetSlices = useMemo<AnalyticsDonutSlice[]>(() => {
    // Accounts holding actual assets (either leaf accounts or top-level accounts without children)
    const targetAccounts = accounts.filter(a => !a.isParent || !accounts.some(c => c.parentId === a.id));

    const entries = targetAccounts
      .map(acc => {
        const bal = accountBalances[acc.id] || 0;
        const mult = acc.currency === 'USD' ? USD_rate : acc.currency === 'LYD' ? LYD_rate : 1;
        const bdtValue = bal * mult;
        return {
          id: acc.id,
          name: acc.name,
          value: bdtValue,
        };
      })
      .filter(e => e.value > 0)
      .sort((a, b) => b.value - a.value);

    const total = entries.reduce((sum, e) => sum + e.value, 0);
    return entries.map((e, index) => ({
      ...e,
      pct: total > 0 ? e.value / total : 0,
      color: HOLDING_COLORS[index % HOLDING_COLORS.length],
    }));
  }, [accounts, accountBalances, USD_rate, LYD_rate]);

  const categoryActivity = useMemo(() => {
    const results: Array<{
      id: string;
      name: string;
      type: 'Income' | 'Expense' | 'Loan';
      value: number;
      subCategories: Array<{ name: string; value: number }>;
    }> = [];

    Object.entries(INCOME_CATEGORIES).forEach(([catName, subCats]) => {
      if (catName === 'Loan') return;
      const txsForCat = filteredTransactions.filter(t => t.type === 'Income' && t.category === catName);
      let parentValue = 0;
      const subCatMap: Record<string, number> = {};

      subCats.forEach(sc => {
        subCatMap[sc] = 0;
      });

      txsForCat.forEach(t => {
        const acc = accounts.find(a => a.id === t.accountId);
        const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
        const bdtAmt = t.amount * mult;
        parentValue += bdtAmt;

        const scName = t.subCategory || 'Other';
        subCatMap[scName] = (subCatMap[scName] || 0) + bdtAmt;
      });

      const processedSubs = Object.entries(subCatMap).map(([name, value]) => ({
        name,
        value
      })).sort((a, b) => b.value - a.value);

      results.push({
        id: `income-${catName}`,
        name: catName,
        type: 'Income',
        value: parentValue,
        subCategories: processedSubs
      });
    });

    Object.entries(EXPENSE_CATEGORIES).forEach(([catName, subCats]) => {
      if (catName === 'Loan') return;
      const txsForCat = filteredTransactions.filter(t => t.type === 'Expense' && t.category === catName);
      let parentValue = 0;
      const subCatMap: Record<string, number> = {};

      subCats.forEach(sc => {
        subCatMap[sc] = 0;
      });

      txsForCat.forEach(t => {
        const acc = accounts.find(a => a.id === t.accountId);
        const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
        const bdtAmt = t.amount * mult;
        parentValue += bdtAmt;

        const scName = t.subCategory || 'Other';
        subCatMap[scName] = (subCatMap[scName] || 0) + bdtAmt;
      });

      const processedSubs = Object.entries(subCatMap).map(([name, value]) => ({
        name,
        value
      })).sort((a, b) => b.value - a.value);

      results.push({
        id: `expense-${catName}`,
        name: catName,
        type: 'Expense',
        value: parentValue,
        subCategories: processedSubs
      });
    });

    // Unified single 'Loan' category calculation
    const loanTransactions = filteredTransactions.filter(t => t.type === 'Loan' || t.category === 'Loan');
    const loanSubCatMap: Record<string, number> = {
      'Borrowed': 0,
      'Lent': 0,
      'Received Lent Money': 0,
      'Repaid Borrowed Money': 0
    };

    let loanInflows = 0;
    let loanOutflows = 0;

    loanTransactions.forEach(t => {
      const acc = accounts.find(a => a.id === t.accountId);
      const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
      const bdtAmt = t.amount * mult;

      const scName = t.subCategory || 'Other';
      if (scName in loanSubCatMap) {
        loanSubCatMap[scName] += bdtAmt;
      } else {
        loanSubCatMap[scName] = (loanSubCatMap[scName] || 0) + bdtAmt;
      }

      if (scName === 'Borrowed' || scName === 'Received Lent Money') {
        loanInflows += bdtAmt;
      } else if (scName === 'Lent' || scName === 'Repaid Borrowed Money') {
        loanOutflows += bdtAmt;
      } else {
        // Fallback
        loanInflows += bdtAmt;
      }
    });

    // Net value of loan transactions: borrowed - lent
    const loanNetValue = loanInflows - loanOutflows;

    results.push({
      id: 'loan-unified',
      name: 'Loan',
      type: 'Loan',
      value: loanNetValue,
      subCategories: Object.entries(loanSubCatMap).map(([name, value]) => ({
        name,
        value
      }))
    });

    return results;
  }, [filteredTransactions, accounts, INCOME_CATEGORIES, EXPENSE_CATEGORIES, USD_rate, LYD_rate]);

  const sortedCategories = useMemo(() => {
    const incomes = categoryActivity.filter(c => c.type === 'Income').sort((a, b) => b.value - a.value);
    const expenses = categoryActivity.filter(c => c.type === 'Expense').sort((a, b) => b.value - a.value);
    const loans = categoryActivity.filter(c => c.type === 'Loan');
    return [...incomes, ...expenses, ...loans];
  }, [categoryActivity]);

  // ─── TYPE CONFIG ─────────────────────────────────────────────────────────
  const TX_TYPES = [
    {
      value: 'Expense',
      label: 'Expense',
      color: 'rose',
      activeClasses: 'bg-rose-400/15 border-rose-400/50 text-rose-300 shadow-[0_0_10px_rgba(251,113,133,0.12)]',
      inactiveClasses: 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700 hover:text-slate-200',
      dotColor: 'bg-rose-400'
    },
    {
      value: 'Income',
      label: 'Income',
      color: 'emerald',
      activeClasses: 'bg-emerald-400/15 border-emerald-400/50 text-emerald-300 shadow-[0_0_10px_rgba(52,211,153,0.12)]',
      inactiveClasses: 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700 hover:text-slate-200',
      dotColor: 'bg-emerald-400'
    },
    {
      value: 'Transfer',
      label: 'Transfer',
      color: 'sky',
      activeClasses: 'bg-sky-400/15 border-sky-400/50 text-sky-300 shadow-[0_0_10px_rgba(56,189,248,0.12)]',
      inactiveClasses: 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700 hover:text-slate-200',
      dotColor: 'bg-sky-400'
    },
    {
      value: 'Loan',
      label: 'Loan',
      color: 'amber',
      activeClasses: 'bg-amber-400/15 border-amber-400/50 text-amber-300 shadow-[0_0_10px_rgba(245,158,11,0.12)]',
      inactiveClasses: 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700 hover:text-slate-200',
      dotColor: 'bg-amber-400'
    }
  ] as const;

  // Render Submodule Contents
  const renderSubmodule = () => {
    switch (activeTab) {
      
      // ─── SUMMARY SUBMODULE ───────────────────────────────────────────────
      case 'summary': {
        const periodIncomeBdt = filteredTransactions.filter(t => t.type === 'Income').reduce((sum, t) => {
          const acc = accounts.find(a => a.id === t.accountId);
          const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
          return sum + (t.amount * mult);
        }, 0);

        const periodExpenseBdt = filteredTransactions.filter(t => t.type === 'Expense').reduce((sum, t) => {
          const acc = accounts.find(a => a.id === t.accountId);
          const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
          return sum + (t.amount * mult);
        }, 0);

        const bdtIn = filteredTransactions.filter(t => t.type === 'Income' && accounts.find(a => a.id === t.accountId)?.currency === 'BDT').reduce((sum, t) => sum + t.amount, 0);
        const bdtOut = filteredTransactions.filter(t => t.type === 'Expense' && accounts.find(a => a.id === t.accountId)?.currency === 'BDT').reduce((sum, t) => sum + t.amount, 0);

        const lydIn = filteredTransactions.filter(t => t.type === 'Income' && accounts.find(a => a.id === t.accountId)?.currency === 'LYD').reduce((sum, t) => sum + t.amount, 0);
        const lydOut = filteredTransactions.filter(t => t.type === 'Expense' && accounts.find(a => a.id === t.accountId)?.currency === 'LYD').reduce((sum, t) => sum + t.amount, 0);

        const usdIn = filteredTransactions.filter(t => t.type === 'Income' && accounts.find(a => a.id === t.accountId)?.currency === 'USD').reduce((sum, t) => sum + t.amount, 0);
        const usdOut = filteredTransactions.filter(t => t.type === 'Expense' && accounts.find(a => a.id === t.accountId)?.currency === 'USD').reduce((sum, t) => sum + t.amount, 0);

        return (
          <div className="space-y-6 sm:space-y-8">
            
            {/* 4 Cards Container */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4 items-stretch">
              
              {/* Card 1: CONVERTED BDT */}
              <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-2.5 sm:p-4 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group">
                <div className="flex-1">
                  <div className="flex justify-between items-start mb-2 sm:mb-4">
                    <div className="w-7 h-7 sm:w-10 sm:h-10 rounded-lg sm:rounded-xl bg-teal-400/10 flex items-center justify-center text-teal-400 group-hover:scale-110 transition-transform">
                      <span className="font-extrabold text-xs sm:text-base font-display">৳</span>
                    </div>
                  </div>
                  <div className="mb-1 sm:mb-2">
                    <p className="text-[10px] sm:text-body-sm font-bold text-white uppercase tracking-wider truncate">Total Balance in BDT</p>
                  </div>
                  <h3 className="text-subheading sm:text-heading font-bold text-teal-400 mb-2 sm:mb-4 tracking-tight font-display tabular-nums truncate">
                    {formatBDT(totalBDT + (totalLYD * LYD_rate) + (totalUSD * USD_rate))}
                  </h3>
                  {(totalCurrentBorrowedBDT > 0 || totalCurrentLentBDT > 0) && (
                    <div className="flex flex-wrap items-center gap-1.5 sm:gap-2 mb-2 sm:mb-3 mt-[-4px] sm:mt-[-6px]">
                      {totalCurrentBorrowedBDT > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-rose-500/10 text-rose-400 px-1 sm:px-1.5 py-0.5 rounded border border-rose-500/20 font-bold uppercase tracking-tight truncate">
                          Borrowed: -{formatBDT(totalCurrentBorrowedBDT)}
                        </span>
                      )}
                      {totalCurrentLentBDT > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-emerald-500/10 text-emerald-400 px-1 sm:px-1.5 py-0.5 rounded border border-emerald-500/20 font-bold uppercase tracking-tight truncate">
                          Lent: +{formatBDT(totalCurrentLentBDT)}
                        </span>
                      )}
                    </div>
                  )}
                </div>
                <div className="flex flex-col justify-center gap-1 sm:gap-1.5 pt-2 sm:pt-3 border-t border-slate-800 min-h-[44px] sm:h-14">
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowUpRight size={9} className="text-emerald-400 shrink-0" /> Inflow</span>
                    <span className="text-emerald-400 font-bold tabular-nums truncate ml-1">{formatBDT(periodIncomeBdt)}</span>
                  </div>
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowDownRight size={9} className="text-rose-400 shrink-0" /> Outflow</span>
                    <span className="text-rose-400 font-bold tabular-nums truncate ml-1">{formatBDT(periodExpenseBdt)}</span>
                  </div>
                </div>
              </Card>

              {/* Card 2: BDT accounts balance */}
              <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-2.5 sm:p-4 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group">
                <div className="flex-1">
                  <div className="flex justify-between items-start mb-2 sm:mb-4">
                    <div className="w-7 h-7 sm:w-10 sm:h-10 rounded-lg sm:rounded-xl bg-teal-400/10 flex items-center justify-center text-teal-400 group-hover:scale-110 transition-transform">
                      <span className="font-extrabold text-xs sm:text-base font-display">৳</span>
                    </div>
                  </div>
                  <div className="mb-1 sm:mb-2">
                    <p className="text-[10px] sm:text-body-sm font-bold text-white uppercase tracking-wider truncate">Balance in BDT</p>
                  </div>
                  <h3 className="text-subheading sm:text-heading font-bold text-teal-400 mb-2 sm:mb-4 tracking-tight font-display tabular-nums truncate">
                    {formatBDT(totalBDT)}
                  </h3>
                  {(loanStatsByCurrency.BDT.borrowed > 0 || loanStatsByCurrency.BDT.lent > 0) && (
                    <div className="flex flex-wrap items-center gap-1.5 sm:gap-2 mb-2 sm:mb-3 mt-[-4px] sm:mt-[-6px]">
                      {loanStatsByCurrency.BDT.borrowed > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-rose-500/10 text-rose-400 px-1 sm:px-1.5 py-0.5 rounded border border-rose-500/20 font-bold uppercase tracking-tight truncate">
                          Borrowed: -{formatBDT(loanStatsByCurrency.BDT.borrowed)}
                        </span>
                      )}
                      {loanStatsByCurrency.BDT.lent > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-emerald-500/10 text-emerald-400 px-1 sm:px-1.5 py-0.5 rounded border border-emerald-500/20 font-bold uppercase tracking-tight truncate">
                          Lent: +{formatBDT(loanStatsByCurrency.BDT.lent)}
                        </span>
                      )}
                    </div>
                  )}
                </div>
                <div className="flex flex-col justify-center gap-1 sm:gap-1.5 pt-2 sm:pt-3 border-t border-slate-800 min-h-[44px] sm:h-14">
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowUpRight size={9} className="text-emerald-400 shrink-0" /> Inflow</span>
                    <span className="text-emerald-400 font-bold tabular-nums truncate ml-1">{formatBDT(bdtIn)}</span>
                  </div>
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowDownRight size={9} className="text-rose-400 shrink-0" /> Outflow</span>
                    <span className="text-rose-400 font-bold tabular-nums truncate ml-1">{formatBDT(bdtOut)}</span>
                  </div>
                </div>
              </Card>

              {/* Card 3: LYD accounts balance */}
              <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-2.5 sm:p-4 transition-all hover:border-amber-400/50 hover:shadow-[0_0_20px_rgba(245,158,11,0.1)] group">
                <div className="flex-1">
                  <div className="flex justify-between items-start mb-2 sm:mb-4">
                    <div className="w-7 h-7 sm:w-10 sm:h-10 rounded-lg sm:rounded-xl bg-amber-400/10 flex items-center justify-center text-amber-400 group-hover:scale-110 transition-transform font-display">
                      <span className="font-extrabold text-[10px] sm:text-sm font-display">LD</span>
                    </div>
                  </div>
                  <div className="mb-1 sm:mb-2">
                    <p className="text-[10px] sm:text-body-sm font-bold text-white uppercase tracking-wider truncate">Balance in LYD</p>
                  </div>
                  <h3 className="text-subheading sm:text-heading font-bold text-amber-400 mb-2 sm:mb-4 tracking-tight font-display tabular-nums truncate">
                    {formatCurrency(totalLYD, 'LYD')}
                  </h3>
                  {(loanStatsByCurrency.LYD.borrowed > 0 || loanStatsByCurrency.LYD.lent > 0) && (
                    <div className="flex flex-wrap items-center gap-1.5 sm:gap-2 mb-2 sm:mb-3 mt-[-4px] sm:mt-[-6px]">
                      {loanStatsByCurrency.LYD.borrowed > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-rose-500/10 text-rose-400 px-1 sm:px-1.5 py-0.5 rounded border border-rose-500/20 font-bold uppercase tracking-tight truncate">
                          Borrowed: -{formatCurrency(loanStatsByCurrency.LYD.borrowed, 'LYD')}
                        </span>
                      )}
                      {loanStatsByCurrency.LYD.lent > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-emerald-500/10 text-emerald-400 px-1 sm:px-1.5 py-0.5 rounded border border-emerald-500/20 font-bold uppercase tracking-tight truncate font-sans">
                          Lent: +{formatCurrency(loanStatsByCurrency.LYD.lent, 'LYD')}
                        </span>
                      )}
                    </div>
                  )}
                </div>
                <div className="flex flex-col justify-center gap-1 sm:gap-1.5 pt-2 sm:pt-3 border-t border-slate-800 min-h-[44px] sm:h-14">
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowUpRight size={9} className="text-emerald-400 shrink-0" /> Inflow</span>
                    <span className="text-emerald-400 font-bold tabular-nums truncate ml-1">{formatCurrency(lydIn, 'LYD')}</span>
                  </div>
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowDownRight size={9} className="text-rose-400 shrink-0" /> Outflow</span>
                    <span className="text-rose-400 font-bold tabular-nums truncate ml-1">{formatCurrency(lydOut, 'LYD')}</span>
                  </div>
                </div>
              </Card>

              {/* Card 4: USD accounts balance */}
              <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-2.5 sm:p-4 transition-all hover:border-sky-400/50 hover:shadow-[0_0_20px_rgba(56,189,248,0.1)] group">
                <div className="flex-1">
                  <div className="flex justify-between items-start mb-2 sm:mb-4">
                    <div className="w-7 h-7 sm:w-10 sm:h-10 rounded-lg sm:rounded-xl bg-sky-400/10 flex items-center justify-center text-sky-400 group-hover:scale-110 transition-transform font-display">
                      <span className="font-extrabold text-xs sm:text-base font-display">$</span>
                    </div>
                  </div>
                  <div className="mb-1 sm:mb-2">
                    <p className="text-[10px] sm:text-body-sm font-bold text-white uppercase tracking-wider truncate">Balance in USD</p>
                  </div>
                  <h3 className="text-subheading sm:text-heading font-bold text-sky-400 mb-2 sm:mb-4 tracking-tight font-display tabular-nums truncate">
                    {formatCurrency(totalUSD, 'USD')}
                  </h3>
                  {(loanStatsByCurrency.USD.borrowed > 0 || loanStatsByCurrency.USD.lent > 0) && (
                    <div className="flex flex-wrap items-center gap-1.5 sm:gap-2 mb-2 sm:mb-3 mt-[-4px] sm:mt-[-6px]">
                      {loanStatsByCurrency.USD.borrowed > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-rose-500/10 text-rose-400 px-1 sm:px-1.5 py-0.5 rounded border border-rose-500/20 font-bold uppercase tracking-tight truncate font-sans">
                          Borrowed: -{formatCurrency(loanStatsByCurrency.USD.borrowed, 'USD')}
                        </span>
                      )}
                      {loanStatsByCurrency.USD.lent > 0 && (
                        <span className="text-[8px] sm:text-[9px] bg-emerald-500/10 text-emerald-400 px-1 sm:px-1.5 py-0.5 rounded border border-emerald-500/20 font-bold uppercase tracking-tight truncate font-sans">
                          Lent: +{formatCurrency(loanStatsByCurrency.USD.lent, 'USD')}
                        </span>
                      )}
                    </div>
                  )}
                </div>
                <div className="flex flex-col justify-center gap-1 sm:gap-1.5 pt-2 sm:pt-3 border-t border-slate-800 min-h-[44px] sm:h-14">
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowUpRight size={9} className="text-emerald-400 shrink-0" /> Inflow</span>
                    <span className="text-emerald-400 font-bold tabular-nums truncate ml-1">{formatCurrency(usdIn, 'USD')}</span>
                  </div>
                  <div className="flex justify-between items-center text-[9px] sm:text-[10px] font-semibold text-slate-400">
                    <span className="flex items-center gap-0.5 sm:gap-1"><ArrowDownRight size={9} className="text-rose-400 shrink-0" /> Outflow</span>
                    <span className="text-rose-400 font-bold tabular-nums truncate ml-1">{formatCurrency(usdOut, 'USD')}</span>
                  </div>
                </div>
              </Card>

            </div>

            {/* Middle Section: Profiles list and Quick Stats (60% Accounts, 40% Period Activity Summary) */}
            <div className="grid grid-cols-1 lg:grid-cols-5 gap-6 lg:gap-8">
              
              {/* Account Profiles (60% width) */}
              <div className="lg:col-span-3 bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-6">
                  <h4 className="text-subheading font-bold text-white uppercase tracking-tight">Accounts</h4>
                  
                  <div className="flex items-center gap-1.5 px-1">
                    <button 
                      onClick={() => {
                        if (rangeType === 'this') navigateThisMonth(-1);
                        else if (rangeType === 'fiscal') setFiscalStartYear(prev => prev - 1);
                        else navigateCustomMonth(-1);
                      }} 
                      className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-950 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
                    >
                      <ChevronLeft size={14} />
                    </button>

                    {rangeType === 'custom' ? (
                      <div className="flex items-center gap-1.5">
                        <input type="date" value={customDates.start} onChange={(e) => setCustomDates(prev => ({ ...prev, start: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer w-[110px]" />
                        <span className="text-slate-700 font-bold text-[10px]">–</span>
                        <input type="date" value={customDates.end} onChange={(e) => setCustomDates(prev => ({ ...prev, end: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer w-[110px]" />
                      </div>
                    ) : rangeType === 'fiscal' ? (
                      <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-white uppercase select-none tracking-wider whitespace-nowrap min-w-[150px] text-center">
                        July {fiscalStartYear} - June {fiscalStartYear + 1}
                      </div>
                    ) : (
                      <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-teal-400 uppercase select-none tracking-wider whitespace-nowrap min-w-[110px] text-center font-sans">
                        {formatThisMonthLabel(thisMonthDate)}
                      </div>
                    )}

                    <button 
                      onClick={() => {
                        if (rangeType === 'this') navigateThisMonth(1);
                        else if (rangeType === 'fiscal') setFiscalStartYear(prev => prev + 1);
                        else navigateCustomMonth(1);
                      }} 
                      className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-950 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
                    >
                      <ChevronRight size={14} />
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-2 xs:grid-cols-2 sm:grid-cols-3 lg:grid-cols-3 gap-2.5 max-h-[500px] overflow-y-auto pr-1">
                  {accounts.filter(acc => !acc.parentId).map(acc => {
                    const balance = accountBalances[acc.id] || 0;
                    const children = accounts.filter(child => child.parentId === acc.id);
                    const isParent = acc.isParent;

                    if (isParent) {
                      const isExpanded = !!expandedGroups[acc.id];
                      return (
                        <div 
                          key={acc.id} 
                          className={cn(
                            "p-2.5 bg-slate-950 border rounded-lg flex flex-col justify-between hover:bg-slate-900 transition-all group relative min-h-[54px] cursor-pointer select-none",
                            isExpanded 
                              ? cn("col-span-full bg-slate-950/25 p-3 space-y-3 shadow-md", acc.currency === 'BDT' ? "border-teal-500/30" : acc.currency === 'LYD' ? "border-amber-500/30" : "border-sky-500/30")
                              : acc.currency === 'BDT' ? "border-teal-500/13 hover:border-teal-400/30" : acc.currency === 'LYD' ? "border-amber-500/13 hover:border-amber-400/30" : "border-sky-500/13 hover:border-sky-400/30"
                          )}
                          onClick={() => setExpandedGroups(prev => ({ ...prev, [acc.id]: !prev[acc.id] }))}
                        >
                          <div className="flex flex-col justify-between w-full h-full min-h-[35px]">
                            <div className="flex justify-between items-start gap-1">
                              <div className="flex items-center gap-1.5 min-w-0 pr-1">
                                <AccountBadgeIcon account={acc} size="xs" />
                                <span className="font-semibold text-[11px] text-slate-200 truncate" title={acc.name}>{acc.name}</span>
                                <ChevronDown size={11} className={cn("text-slate-400 transition-transform duration-200 shrink-0 ml-auto", isExpanded ? "rotate-180" : "")} />
                              </div>
                            </div>
                            <div className="flex justify-between items-end mt-1">
                              <div className={cn("font-bold text-[11px] tabular-nums", acc.currency === 'BDT' ? "text-teal-400" : acc.currency === 'USD' ? "text-sky-400" : "text-amber-400")}>
                                {formatCurrency(balance, acc.currency)}
                              </div>
                            </div>
                          </div>
                          {isExpanded && (
                            <div className="w-full pt-1.5 border-t border-slate-800/60" onClick={(e) => e.stopPropagation()}>
                              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                                {children.map(child => {
                                  const childBalance = accountBalances[child.id] || 0;
                                  return (
                                    <div key={child.id} className={cn("p-2 bg-slate-950/60 border rounded-lg flex flex-col justify-between hover:bg-slate-900 transition-all min-h-[50px]", child.currency === 'BDT' ? "border-teal-500/10 hover:border-teal-400/30" : child.currency === 'LYD' ? "border-amber-500/10 hover:border-amber-400/30" : "border-sky-500/10 hover:border-sky-400/30")}>
                                      <div className="flex items-center gap-1.5 min-w-0">
                                        <AccountBadgeIcon account={child} size="xs" />
                                        <span className="font-semibold text-[11px] text-slate-300 truncate" title={child.name}>{child.name}</span>
                                      </div>
                                      <div className={cn("font-bold text-[11px] tabular-nums mt-1", child.currency === 'BDT' ? "text-teal-400" : child.currency === 'USD' ? "text-sky-400" : "text-amber-400")}>
                                        {formatCurrency(childBalance, child.currency)}
                                      </div>
                                    </div>
                                  );
                                })}
                                {children.length === 0 && (
                                  <div className="col-span-full py-2 text-center text-[10px] text-slate-500 italic uppercase">No sub-accounts</div>
                                )}
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    }

                    return (
                      <div key={acc.id} className={cn("p-2.5 bg-slate-950 border rounded-lg flex flex-col justify-between hover:bg-slate-900 transition-all group relative min-h-[54px]", acc.currency === 'BDT' ? "border-teal-500/13 hover:border-teal-400/30" : acc.currency === 'LYD' ? "border-amber-500/13 hover:border-amber-400/30" : "border-sky-500/13 hover:border-sky-400/30")}>
                        <div className="flex items-center gap-1.5 min-w-0 pr-1">
                          <AccountBadgeIcon account={acc} size="xs" />
                          <div className="font-semibold text-[11px] text-slate-200 truncate" title={acc.name}>{acc.name}</div>
                        </div>
                        <div className={cn("font-bold text-[11px] tabular-nums mt-1", acc.currency === 'BDT' ? "text-teal-400" : acc.currency === 'USD' ? "text-sky-400" : "text-amber-400")}>
                          {formatCurrency(balance, acc.currency)}
                        </div>
                      </div>
                    );
                  })}
                  {accounts.length === 0 && (
                    <div className="text-center text-slate-500 text-label py-12 uppercase italic col-span-full">No accounts configured</div>
                  )}
                </div>
              </div>

              {/* Net Period Savings (40% width) */}
              <div className="lg:col-span-2 bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl flex flex-col justify-between">
                <div>
                  <h4 className="text-subheading font-bold text-white uppercase tracking-tight mb-6">Period Activity summary</h4>
                  
                  <div className="grid grid-cols-2 gap-4 mb-6">
                    <div className="bg-emerald-500/10 border border-emerald-500/20 p-4 rounded-xl">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[10px] text-emerald-400 font-bold uppercase tracking-wider">Gross Income (BDT)</span>
                        <ArrowUpRight size={16} className="text-emerald-400" />
                      </div>
                      <div className="text-lg sm:text-xl font-bold text-white tabular-nums truncate">{formatBDT(periodIncomeBdt)}</div>
                    </div>
                    <div className="bg-rose-500/10 border border-rose-500/20 p-4 rounded-xl">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[10px] text-rose-400 font-bold uppercase tracking-wider">Total Expense (BDT)</span>
                        <ArrowDownRight size={16} className="text-rose-400" />
                      </div>
                      <div className="text-lg sm:text-xl font-bold text-white tabular-nums truncate">{formatBDT(periodExpenseBdt)}</div>
                    </div>
                  </div>

                  <div className="p-4 bg-slate-950 border border-slate-800/80 rounded-xl">
                    <div className="flex justify-between items-center mb-1">
                      <span className="text-label text-slate-400 font-bold uppercase">Net Savings Flow</span>
                      <span className={cn("text-body font-bold uppercase", (periodIncomeBdt - periodExpenseBdt) >= 0 ? "text-emerald-400" : "text-rose-400")}>
                        {(periodIncomeBdt - periodExpenseBdt) >= 0 ? 'Surplus' : 'Deficit'}
                      </span>
                    </div>
                    <div className="text-2xl font-extrabold text-white tabular-nums tracking-tight">{formatBDT(periodIncomeBdt - periodExpenseBdt)}</div>
                    {periodIncomeBdt > 0 && (
                      <div className="mt-2 w-full bg-slate-900 h-2 rounded-full overflow-hidden">
                        <div className="bg-teal-400 h-full rounded-full" style={{ width: `${Math.max(0, Math.min(100, ((periodIncomeBdt - periodExpenseBdt) / periodIncomeBdt) * 100))}%` }} />
                      </div>
                    )}
                  </div>
                </div>
                
                <p className="text-[10px] text-slate-500 font-bold uppercase mt-4 text-center">
                  Calculated from matching transactions between {formatDate(calculatedDates.start)} to {formatDate(calculatedDates.end)}
                </p>
              </div>

            </div>

            {/* Connected Category Summary Panel */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h4 className="text-subheading font-bold text-white uppercase tracking-tight">Activity by Category</h4>
                  <p className="text-[10px] text-slate-500 font-bold uppercase mt-1">Activity and distributions by categories and sub-ledgers (in BDT)</p>
                </div>
                <div className="flex items-center gap-1.5 px-1 self-start sm:self-auto shrink-0">
                  <button onClick={() => { if (rangeType === 'this') navigateThisMonth(-1); else if (rangeType === 'fiscal') setFiscalStartYear(prev => prev - 1); else navigateCustomMonth(-1); }} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-950 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"><ChevronLeft size={14} /></button>
                  {rangeType === 'custom' ? (
                    <div className="flex items-center gap-1.5">
                      <input type="date" value={customDates.start} onChange={(e) => setCustomDates(prev => ({ ...prev, start: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer w-[110px]" />
                      <span className="text-slate-700 font-bold text-[10px]">–</span>
                      <input type="date" value={customDates.end} onChange={(e) => setCustomDates(prev => ({ ...prev, end: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer w-[110px]" />
                    </div>
                  ) : rangeType === 'fiscal' ? (
                    <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-white uppercase select-none tracking-wider whitespace-nowrap min-w-[150px] text-center">July {fiscalStartYear} - June {fiscalStartYear + 1}</div>
                  ) : (
                    <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-teal-400 uppercase select-none tracking-wider whitespace-nowrap min-w-[110px] text-center font-sans">{formatThisMonthLabel(thisMonthDate)}</div>
                  )}
                  <button onClick={() => { if (rangeType === 'this') navigateThisMonth(1); else if (rangeType === 'fiscal') setFiscalStartYear(prev => prev + 1); else navigateCustomMonth(1); }} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-950 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"><ChevronRight size={14} /></button>
                </div>
              </div>

              <div className="grid grid-cols-2 xs:grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-2.5 max-h-[500px] overflow-y-auto pr-1 text-label">
                {sortedCategories.map(cat => {
                  const hasSubcats = cat.subCategories.filter(sub => sub.value > 0).length > 0;
                  const isExpanded = !!expandedCategories[cat.id] && hasSubcats;
                  const catVisual = getCategoryVisual(cat.name, cat.type);
                  return (
                    <div
                      key={cat.id}
                      className={cn(
                        "p-2.5 bg-slate-950 border rounded-lg flex flex-col justify-between hover:bg-slate-900 transition-all group relative min-h-[54px] select-none",
                        hasSubcats ? "cursor-pointer" : "cursor-default",
                        isExpanded
                          ? cn("col-span-full bg-slate-950/25 p-3 space-y-3 shadow-md",
                              cat.type === 'Loan' ? "border-amber-500/30" :
                              cat.type === 'Income' ? "border-emerald-500/30" : "border-rose-500/30")
                          : cat.type === 'Loan' ? "border-amber-500/15 hover:border-amber-400/30" :
                            cat.type === 'Income' ? "border-emerald-500/13 hover:border-emerald-400/30" : "border-rose-500/13 hover:border-rose-400/30"
                      )}
                      onClick={() => {
                        if (hasSubcats) {
                          setExpandedCategories(prev => ({ ...prev, [cat.id]: !prev[cat.id] }));
                        }
                      }}
                    >
                      <div className="flex flex-col justify-between w-full h-full min-h-[35px]">
                        <div className="flex justify-between items-start gap-1">
                          <div className="flex items-center gap-1.5 min-w-0">
                            <div className={cn("w-5 h-5 rounded-md flex items-center justify-center shrink-0 border", catVisual.bg)}>
                              <catVisual.Icon size={11} />
                            </div>
                            <span className="font-semibold text-[11px] text-slate-200 truncate pr-1" title={cat.name}>{cat.name}</span>
                            {hasSubcats && <ChevronDown size={11} className={cn("text-slate-400 transition-transform duration-200 shrink-0", isExpanded ? "rotate-180" : "")} />}
                          </div>
                        </div>
                        <div className="flex justify-between items-end mt-1">
                          <div className={cn("font-bold text-[11px] tabular-nums",
                            cat.type === 'Loan' ? "text-amber-400" :
                            cat.type === 'Income' ? "text-emerald-400" : "text-rose-400"
                          )}>
                            {formatBDT(cat.value)}
                          </div>
                        </div>
                      </div>
                      {isExpanded && hasSubcats && (
                        <div className="w-full pt-1.5 border-t border-slate-800/60" onClick={(e) => e.stopPropagation()}>
                          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2">
                            {cat.subCategories.filter(sub => sub.value > 0).map(sub => {
                              const isLoanGreen = cat.type === 'Loan' && (sub.name === 'Borrowed' || sub.name === 'Received Lent Money');
                              const isLoanRed = cat.type === 'Loan' && (sub.name === 'Lent' || sub.name === 'Repaid Borrowed Money');

                              const subColorClass = isLoanGreen
                                ? "text-emerald-400"
                                : isLoanRed
                                  ? "text-rose-400"
                                  : cat.type === 'Income' ? "text-emerald-400" : "text-rose-400";

                              const borderHoverClass = isLoanGreen
                                ? "border-emerald-500/10 hover:border-emerald-400/30"
                                : isLoanRed
                                  ? "border-rose-500/10 hover:border-rose-400/30"
                                  : cat.type === 'Income'
                                    ? "border-emerald-500/10 hover:border-emerald-400/30"
                                    : "border-rose-500/10 hover:border-rose-400/30";

                              return (
                                <div key={sub.name} className={cn("p-2 bg-slate-950/60 border rounded-lg flex flex-col justify-between hover:bg-slate-900 transition-all min-h-[50px]", borderHoverClass)}>
                                  <span className="font-semibold text-[11px] text-slate-300 truncate" title={sub.name}>{sub.name}</span>
                                  <div className={cn("font-bold text-[11px] tabular-nums mt-1", subColorClass)}>{formatBDT(sub.value)}</div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Recents list */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl">
              <div className="flex justify-between items-center mb-6">
                <h4 className="text-subheading font-bold text-white uppercase tracking-tight">Recent ledger records</h4>
                <Button variant="secondary" size="sm" onClick={() => setActiveTab('transactions')}>Detailed Ledger →</Button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left">
                  <thead>
                    <tr className="border-b border-slate-800 text-[10px] text-slate-500 font-extrabold tracking-widest uppercase">
                      <th className="pb-3 w-28">DATE</th>
                      <th className="pb-3 w-28">TYPE</th>
                      <th className="pb-3 max-w-[120px]">ACCOUNT</th>
                      <th className="pb-3">CATEGORY</th>
                      <th className="pb-3 max-w-[200px]">DESCRIPTION</th>
                      <th className="pb-3 text-right">AMOUNT (NATIVE)</th>
                      <th className="pb-3 text-right w-40">AMOUNT (BDT)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {filteredTransactions.slice(0, 5).map(t => {
                      const acc = accounts.find(a => a.id === t.accountId);
                      const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
                      const toAcc = t.toAccountId ? accounts.find(a => a.id === t.toAccountId) : null;
                      const toMult = toAcc?.currency === 'USD' ? USD_rate : toAcc?.currency === 'LYD' ? LYD_rate : 1;
                      const txDetail = getTxColorAndPrefix(t);

                      return (
                        <tr key={t.id} className="text-label group hover:bg-slate-800/10">
                          <td className="py-3 font-medium tabular-nums text-slate-300">{formatDate(t.date)}</td>
                          <td className="py-3">
                            <span className={cn("px-2 py-0.5 rounded text-[8px] font-extrabold uppercase",
                              t.type === 'Income' ? "bg-emerald-400/10 text-emerald-400 border border-emerald-400/20" :
                              t.type === 'Transfer' ? "bg-sky-400/10 text-sky-400 border border-sky-400/20" :
                              t.type === 'Loan' ? "bg-amber-400/10 text-amber-400 border border-amber-400/20" :
                              "bg-rose-400/10 text-rose-400 border border-rose-400/20"
                            )}>
                              {t.type}
                            </span>
                          </td>
                          <td className="py-3 font-semibold max-w-[140px] truncate">
                            <div className="flex items-center gap-1.5 min-w-0">
                              <AccountBadgeIcon account={acc} size="xs" />
                              <span className={cn("truncate", getCurrencyColorClass(acc?.currency))}>
                                {acc ? acc.name : 'Unknown'}
                              </span>
                            </div>
                          </td>
                          <td className="py-3 text-slate-300">
                            {t.type === 'Transfer' ? (
                              t.transferType ? (
                                t.transferType === 'from' ? (
                                  <span className="font-bold text-rose-400">Transfer Out</span>
                                ) : (
                                  <span className="font-bold text-emerald-400">Transfer In</span>
                                )
                              ) : (
                                <span className="font-bold text-sky-400">Funds Transfer</span>
                              )
                            ) : (
                              <div className="flex flex-col">
                                <span className="font-bold text-white leading-tight">{t.category || '—'}</span>
                                {t.subCategory && <span className="text-[11px] text-slate-400 font-medium leading-tight mt-0.5">{t.subCategory}</span>}
                              </div>
                            )}
                          </td>
                          <td className="py-3 text-slate-400 truncate max-w-[200px] font-medium" title={t.description}>{t.description || '—'}</td>
                          <td className="py-3 text-right font-extrabold tabular-nums">
                            {t.type === 'Transfer' ? (
                              t.transferType ? (
                                t.transferType === 'from' ? (
                                  <span className="text-rose-400">-{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                ) : (
                                  <span className="text-emerald-400 text-right">+{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                )
                              ) : (
                                <div className="flex flex-col items-end gap-0.5">
                                  <span className="text-rose-400 text-[10px] tabular-nums">-{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                  <span className="text-emerald-400 text-[10px] tabular-nums">+{formatCurrency(t.toAmount ?? t.amount, toAcc?.currency || 'BDT')}</span>
                                </div>
                              )
                            ) : (
                              <span className={txDetail.className}>
                                {txDetail.prefix}{formatCurrency(Math.abs(t.amount), acc?.currency || 'BDT')}
                              </span>
                            )}
                          </td>
                          <td className={cn("py-3 text-right font-extrabold tabular-nums", txDetail.className)}>
                            {t.type === 'Transfer' ? (
                              t.transferType ? (
                                t.transferType === 'from' ? (
                                  <span className="text-rose-400/80">-{formatBDT(t.amount * mult)}</span>
                                ) : (
                                  <span className="text-emerald-400/80">+{formatBDT(t.amount * mult)}</span>
                                )
                              ) : (
                                <div className="flex flex-col items-end gap-0.5">
                                  <span className="text-rose-400/80 text-[10px] tabular-nums">-{formatBDT(t.amount * mult)}</span>
                                  <span className="text-emerald-400/80 text-[10px] tabular-nums">+{formatBDT((t.toAmount ?? t.amount) * toMult)}</span>
                                </div>
                              )
                            ) : (
                              <span>{txDetail.prefix}{formatBDT(Math.abs(t.amount) * mult)}</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                    {filteredTransactions.length === 0 && (
                      <tr><td colSpan={7} className="py-12 text-center text-slate-500 italic uppercase">No ledger records in selected range</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        );
      }

      // ─── TRANSACTION LEDGER MODULE ────────────────────────────────────────
      case 'transactions': {
        const selectedAcc = txAccFilter !== 'All' ? accounts.find(a => a.id === txAccFilter) : null;
        const isTypeOrCatFiltered = txTypeFilter !== 'All' || txCatFilter !== 'All';

        const periodTotalFiltered = filteredTabTransactions.reduce((sum, t) => {
          if (txAccFilter !== 'All' && selectedAcc) {
            return sum + Math.abs(t.amount);
          }
          const acc = accounts.find(a => a.id === t.accountId);
          const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
          return sum + (Math.abs(t.amount) * mult);
        }, 0);

        return (
          <div className="space-y-6 sm:space-y-8">
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 gap-4 flex flex-col md:flex-row items-center justify-between">
              <div className="relative w-full md:w-72">
                <Search size={14} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-500" />
                <input type="text" placeholder="SEARCH RECENT LEDGER..." value={txSearch} onChange={(e) => setTxSearch(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-10 pr-4 py-2 text-label font-bold text-white placeholder-slate-600 outline-none focus:border-teal-400 uppercase tracking-wider" />
              </div>
              <div className="grid grid-cols-3 sm:flex sm:items-center gap-1.5 sm:gap-2.5 w-full md:w-auto">
                {/* 1. Types Dropdown */}
                <div className="relative min-w-0 sm:w-36">
                  <button
                    type="button"
                    onClick={() => setActiveTxDropdown(prev => prev === 'type' ? null : 'type')}
                    className={cn(
                      "w-full flex items-center justify-between gap-1 bg-slate-950 border rounded-lg px-2 sm:px-3 py-2 text-[10px] font-extrabold uppercase transition-colors outline-none cursor-pointer",
                      txTypeFilter !== 'All' ? "border-teal-400/50 text-teal-400" : "border-slate-800 text-slate-300 hover:border-slate-700"
                    )}
                  >
                    <span className="truncate">
                      {txTypeFilter === 'All' ? 'All Types' : txTypeFilter}
                    </span>
                    <ChevronDown size={12} className={cn("text-slate-500 shrink-0 transition-transform duration-200", activeTxDropdown === 'type' ? "rotate-180 text-teal-400" : "")} />
                  </button>

                  {activeTxDropdown === 'type' && (
                    <>
                      <div className="fixed inset-0 z-40" onClick={() => setActiveTxDropdown(null)} />
                      <div className="absolute left-0 mt-1.5 w-36 sm:w-40 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95 backdrop-blur-md divide-y divide-slate-800/30">
                        {(['All', 'Income', 'Expense'] as const).map(type => (
                          <div key={type} className="py-0.5 first:pt-0 last:pb-0">
                            <button
                              type="button"
                              onClick={() => {
                                setTxTypeFilter(type);
                                setActiveTxDropdown(null);
                              }}
                              className={cn(
                                "w-full text-left px-2.5 py-1.5 text-[10px] font-bold rounded-lg transition-colors uppercase truncate cursor-pointer",
                                txTypeFilter === type ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                              )}
                            >
                              {type === 'All' ? 'All Types' : type}
                            </button>
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                </div>

                {/* 2. Accounts Dropdown */}
                <div className="relative min-w-0 sm:w-44">
                  <button
                    type="button"
                    onClick={() => setActiveTxDropdown(prev => prev === 'acc' ? null : 'acc')}
                    className={cn(
                      "w-full flex items-center justify-between gap-1 bg-slate-950 border rounded-lg px-2 sm:px-3 py-2 text-[10px] font-extrabold uppercase transition-colors outline-none cursor-pointer",
                      txAccFilter !== 'All' ? "border-teal-400/50 text-teal-400" : "border-slate-800 text-slate-300 hover:border-slate-700"
                    )}
                  >
                    <span className="truncate">
                      {txAccFilter === 'All' ? 'All Accounts' : (accounts.find(a => a.id === txAccFilter)?.name || 'All Accounts')}
                    </span>
                    <ChevronDown size={12} className={cn("text-slate-500 shrink-0 transition-transform duration-200", activeTxDropdown === 'acc' ? "rotate-180 text-teal-400" : "")} />
                  </button>

                  {activeTxDropdown === 'acc' && (
                    <>
                      <div className="fixed inset-0 z-40" onClick={() => setActiveTxDropdown(null)} />
                      <div className="absolute left-0 mt-1.5 w-44 sm:w-52 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95 backdrop-blur-md divide-y divide-slate-800/30 max-h-60 overflow-y-auto custom-scrollbar">
                        <div className="py-0.5 first:pt-0">
                          <button
                            type="button"
                            onClick={() => {
                              setTxAccFilter('All');
                              setActiveTxDropdown(null);
                            }}
                            className={cn(
                              "w-full text-left px-2.5 py-1.5 text-[10px] font-bold rounded-lg transition-colors uppercase truncate cursor-pointer",
                              txAccFilter === 'All' ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                            )}
                          >
                            All Accounts
                          </button>
                        </div>
                        {accounts.map(a => (
                          <div key={a.id} className="py-0.5">
                            <button
                              type="button"
                              onClick={() => {
                                setTxAccFilter(a.id);
                                setActiveTxDropdown(null);
                              }}
                              className={cn(
                                "w-full text-left px-2.5 py-1.5 text-[10px] font-bold rounded-lg transition-colors uppercase truncate flex items-center justify-between gap-1 cursor-pointer",
                                txAccFilter === a.id ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                              )}
                            >
                              <span className="truncate">{a.name}</span>
                              <span className="text-[8px] opacity-70 shrink-0">{a.currency}</span>
                            </button>
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                </div>

                {/* 3. Categories Dropdown */}
                <div className="relative min-w-0 sm:w-44">
                  <button
                    type="button"
                    onClick={() => setActiveTxDropdown(prev => prev === 'cat' ? null : 'cat')}
                    className={cn(
                      "w-full flex items-center justify-between gap-1 bg-slate-950 border rounded-lg px-2 sm:px-3 py-2 text-[10px] font-extrabold uppercase transition-colors outline-none cursor-pointer",
                      txCatFilter !== 'All' ? "border-teal-400/50 text-teal-400" : "border-slate-800 text-slate-300 hover:border-slate-700"
                    )}
                  >
                    <span className="truncate">
                      {txCatFilter === 'All' ? 'All Categories' : txCatFilter}
                    </span>
                    <ChevronDown size={12} className={cn("text-slate-500 shrink-0 transition-transform duration-200", activeTxDropdown === 'cat' ? "rotate-180 text-teal-400" : "")} />
                  </button>

                  {activeTxDropdown === 'cat' && (
                    <>
                      <div className="fixed inset-0 z-40" onClick={() => setActiveTxDropdown(null)} />
                      <div className="absolute right-0 sm:left-0 mt-1.5 w-48 sm:w-56 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95 backdrop-blur-md divide-y divide-slate-800/30 max-h-64 overflow-y-auto custom-scrollbar">
                        <div className="py-0.5 first:pt-0">
                          <button
                            type="button"
                            onClick={() => {
                              setTxCatFilter('All');
                              setActiveTxDropdown(null);
                            }}
                            className={cn(
                              "w-full text-left px-2.5 py-1.5 text-[10px] font-bold rounded-lg transition-colors uppercase truncate cursor-pointer",
                              txCatFilter === 'All' ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                            )}
                          >
                            All Categories
                          </button>
                        </div>
                        <div className="py-0.5">
                          <button
                            type="button"
                            onClick={() => {
                              setTxCatFilter('Initial Balance');
                              setActiveTxDropdown(null);
                            }}
                            className={cn(
                              "w-full text-left px-2.5 py-1.5 text-[10px] font-bold rounded-lg transition-colors uppercase truncate cursor-pointer",
                              txCatFilter === 'Initial Balance' ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                            )}
                          >
                            Initial Balance
                          </button>
                        </div>
                        {Object.keys(INCOME_CATEGORIES).map(c => (
                          <div key={c} className="py-0.5">
                            <button
                              type="button"
                              onClick={() => {
                                setTxCatFilter(c);
                                setActiveTxDropdown(null);
                              }}
                              className={cn(
                                "w-full text-left px-2.5 py-1.5 text-[10px] font-bold rounded-lg transition-colors uppercase truncate cursor-pointer",
                                txCatFilter === c ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                              )}
                            >
                              {c}
                            </button>
                          </div>
                        ))}
                        {Object.keys(EXPENSE_CATEGORIES).map(c => (
                          <div key={c} className="py-0.5">
                            <button
                              type="button"
                              onClick={() => {
                                setTxCatFilter(c);
                                setActiveTxDropdown(null);
                              }}
                              className={cn(
                                "w-full text-left px-2.5 py-1.5 text-[10px] font-bold rounded-lg transition-colors uppercase truncate cursor-pointer",
                                txCatFilter === c ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                              )}
                            >
                              {c}
                            </button>
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-2xl shadow-xl overflow-hidden">
              <div className="px-5 sm:px-6 py-4 border-b border-slate-800 flex justify-between items-center bg-slate-950/30">
                <div className="flex flex-wrap items-center gap-3">
                  <h4 className="text-subheading font-bold text-white uppercase tracking-tight">Transactions ({filteredTabTransactions.length})</h4>
                  {txAccFilter !== 'All' && selectedAcc && (
                    <span className="hidden sm:inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[10px] font-extrabold bg-slate-800 text-white border border-slate-700 shadow-sm">
                      <AccountBadgeIcon account={selectedAcc} size="xs" />
                      <span>{selectedAcc.name}</span>
                    </span>
                  )}
                  {isTypeOrCatFiltered && (
                    <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[10px] font-extrabold bg-teal-500/10 text-teal-400 border border-teal-500/20 shadow-sm">
                      <span className="text-slate-400 uppercase text-[9px]">Total:</span>
                      <span className="tabular-nums">
                        {txAccFilter !== 'All' && selectedAcc
                          ? formatCurrency(periodTotalFiltered, selectedAcc.currency)
                          : formatBDT(periodTotalFiltered)}
                      </span>
                    </span>
                  )}
                </div>
                <Button size="sm" onClick={() => {
                  setEditingTx(null);
                  setInheritedDseTxId(undefined);
                  const initialToAcc = accounts[1] || accounts[0];
                  setSyncToDse(isBoAccount(initialToAcc));
                  setTxForm({ 
                    date: new Date().toISOString().split('T')[0], 
                    type: 'Expense', 
                    accountId: accounts[0]?.id || '', 
                    toAccountId: initialToAcc?.id || accounts[0]?.id || '',
                    category: '', 
                    subCategory: '', 
                    amount: '', 
                    toAmount: '',
                    description: '' 
                  });
                  setIsTxModalOpen(true);
                }}>
                  <Plus size={14} /> Add Transaction
                </Button>
              </div>

              {/* Batch selection action bar */}
              {selectedTxIds.length > 0 && (
                <div className="bg-rose-500/10 border-b border-rose-500/20 px-5 sm:px-6 py-2.5 flex items-center justify-between">
                  <span className="text-label font-bold text-rose-500 uppercase">
                    {selectedTxIds.length} Transactions Selected
                  </span>
                  <Button 
                    variant="danger" 
                    size="sm" 
                    onClick={handleBatchDelete}
                    className="h-8 px-4 text-label font-bold uppercase animate-pulse hover:animate-none"
                  >
                    Delete Selected
                  </Button>
                </div>
              )}

              {/* Web Table Header (Hidden on Mobile) */}
              <div className="hidden md:grid grid-cols-12 gap-4 px-6 py-3 border-b border-slate-800 text-[10px] text-slate-400 font-extrabold tracking-widest uppercase bg-slate-950/60 select-none">
                <div className="col-span-4 flex items-center gap-3">
                  <Checkbox 
                    label="" 
                    checked={selectedTxIds.length === displayedTabTransactions.length && displayedTabTransactions.length > 0} 
                    onChange={(checked) => {
                      setSelectedTxIds(checked ? displayedTabTransactions.map(t => t.id) : []);
                      setLastSelectedTxId(null);
                    }} 
                  />
                  <span>CATEGORY / SUB</span>
                </div>
                <div className="col-span-3 flex items-center">DESCRIPTION</div>
                <div className="col-span-2 flex items-center">ACCOUNT</div>
                <div className="col-span-2 flex items-center justify-end text-right">
                  AMOUNT
                </div>
                <div className="col-span-1 flex items-center justify-center text-center">ACTIONS</div>
              </div>

              {/* Date Grouped Transactions */}
              <div className="divide-y divide-slate-800/60">
                {groupedTabTransactions.map(group => {
                  const groupTotal = group.transactions.reduce((sum, t) => {
                    if (txAccFilter !== 'All' && selectedAcc) {
                      return sum + Math.abs(t.amount);
                    }
                    const acc = accounts.find(a => a.id === t.accountId);
                    const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
                    return sum + (Math.abs(t.amount) * mult);
                  }, 0);

                  return (
                    <div key={group.date} className="w-full">
                      {/* Date Group Header */}
                      <div className="px-4 sm:px-6 py-2.5 bg-slate-950/70 border-b border-slate-800/80 flex items-center justify-between sticky top-0 z-10 backdrop-blur-md">
                        <div className="flex items-center gap-2">
                          <span className="text-xs sm:text-sm font-bold text-slate-300 tracking-tight">
                            {group.formattedDate}
                          </span>
                          <span className="text-[10px] font-semibold text-slate-500">
                            ({group.transactions.length})
                          </span>
                        </div>
                        <div className="text-right">
                          {isTypeOrCatFiltered ? (
                            <div className="flex items-center gap-1.5 justify-end">
                              <span className="text-[10px] sm:text-xs font-bold text-slate-500 uppercase tracking-wider">Total:</span>
                              <span className={cn("text-xs sm:text-sm font-extrabold tabular-nums", txAccFilter !== 'All' && selectedAcc ? getCurrencyColorClass(selectedAcc.currency) : "text-teal-400")}>
                                {txAccFilter !== 'All' && selectedAcc
                                  ? formatCurrency(groupTotal, selectedAcc.currency)
                                  : formatBDT(groupTotal)}
                              </span>
                            </div>
                          ) : txAccFilter !== 'All' && selectedAcc ? (
                            <div className="flex items-center gap-1.5 justify-end">
                              <span className="text-[10px] sm:text-xs font-bold text-slate-500 uppercase tracking-wider">Bal:</span>
                              <span className={cn("text-xs sm:text-sm font-extrabold tabular-nums", getCurrencyColorClass(selectedAcc.currency))}>
                                {dateBalances[group.date] !== undefined 
                                  ? formatCurrency(dateBalances[group.date], selectedAcc.currency) 
                                  : '—'}
                              </span>
                            </div>
                          ) : (
                            <span className="text-xs sm:text-sm font-extrabold tabular-nums text-slate-400">
                              {formatBDT(Math.abs(group.totalNetBDT))}
                            </span>
                          )}
                        </div>
                      </div>

                      {/* Transaction Rows in this Date Group */}
                      <div className="divide-y divide-slate-800/30">
                        {group.transactions.map((t, idx) => {
                          const acc = accounts.find(a => a.id === t.accountId);
                          const mult = acc?.currency === 'USD' ? USD_rate : acc?.currency === 'LYD' ? LYD_rate : 1;
                          const toAcc = t.toAccountId ? accounts.find(a => a.id === t.toAccountId) : null;
                          const toMult = toAcc?.currency === 'USD' ? USD_rate : toAcc?.currency === 'LYD' ? LYD_rate : 1;
                          const txDetail = getTxColorAndPrefix(t);
                          const isSelected = selectedTxIds.includes(t.id);
                          const visual = getCategoryVisual(t.category, t.type);

                          const prevTx = idx > 0 ? group.transactions[idx - 1] : null;
                          const nextTx = idx < group.transactions.length - 1 ? group.transactions[idx + 1] : null;

                          const isTransfer = t.type === 'Transfer';
                          const isTransferFrom = isTransfer && t.transferType === 'from';
                          const isTransferTo = isTransfer && t.transferType === 'to';

                          const isPairFrom = isTransferFrom && nextTx && nextTx.type === 'Transfer' && nextTx.transferType === 'to';
                          const isPairTo = isTransferTo && prevTx && prevTx.type === 'Transfer' && prevTx.transferType === 'from';
                          const prevWasPairTo = prevTx && prevTx.type === 'Transfer' && prevTx.transferType === 'to';

                          // Category & Sub Category label construction
                          let categoryTitle = t.category || '—';
                          let subCategoryLabel = t.subCategory || '';

                          if (t.type === 'Transfer') {
                            if (t.transferType === 'from') {
                              categoryTitle = 'Transfer Out';
                              subCategoryLabel = acc ? `From: ${acc.name}` : 'Funds Transfer';
                            } else if (t.transferType === 'to') {
                              categoryTitle = 'Transfer In';
                              subCategoryLabel = acc ? `To: ${acc.name}` : 'Funds Transfer';
                            } else {
                              categoryTitle = acc && toAcc ? `${acc.name} → ${toAcc.name}` : 'Funds Transfer';
                              subCategoryLabel = 'Account Transfer';
                            }
                          } else if (t.category === 'Initial Balance') {
                            categoryTitle = 'Initial Balance';
                            subCategoryLabel = acc ? acc.name : 'Starting Balance';
                          }

                          // Border logic:
                          // - Transfer Out in pair: top border, no bottom border
                          // - Transfer In in pair: no top border (no dividing line between pair), bottom border
                          // - Standalone transfer: top and bottom borders
                          // - Row after a transfer pair: clear top border from divide-y to prevent double border
                          const borderClass = isPairFrom
                            ? "!border-t-2 !border-sky-500/40 border-b-0"
                            : isPairTo
                            ? "!border-t-0 !border-b-2 !border-sky-500/40"
                            : isTransfer
                            ? "!border-t-2 !border-b-2 !border-sky-500/40"
                            : prevWasPairTo
                            ? "!border-t-0"
                            : "";

                          return (
                            <div 
                              key={t.id} 
                              onClick={() => handleEditTx(t)}
                              className={cn(
                                "group px-4 sm:px-6 py-3.5 flex flex-col md:grid md:grid-cols-12 md:gap-4 md:items-center hover:bg-slate-800/25 transition-colors cursor-pointer select-none",
                                isSelected
                                  ? "bg-teal-400/10 hover:bg-teal-400/15"
                                  : isTransfer && "bg-sky-500/[0.07] hover:bg-sky-500/[0.12]",
                                borderClass
                              )}
                            >
                              {/* 1st Column: Icon + Category + Sub Category (Web & Mobile) */}
                              <div className="md:col-span-4 flex items-center justify-between md:justify-start gap-3 min-w-0">
                                <div className="flex items-center gap-3 min-w-0 flex-1">
                                  <div 
                                    onClick={(e) => e.stopPropagation()} 
                                    className="flex items-center shrink-0"
                                  >
                                    <Checkbox 
                                      label="" 
                                      checked={isSelected} 
                                      onChange={(_, e) => toggleSelectTx(t.id, e)} 
                                    />
                                  </div>

                                  {/* Category Icon Badge */}
                                  <div className={cn("w-10 h-10 rounded-full flex items-center justify-center border shrink-0 shadow-sm", visual.bg)}>
                                    <visual.Icon size={18} />
                                  </div>

                                  {/* Category & Sub Category Text with Account Badge in Subtitle */}
                                  <div className="min-w-0 flex-1 flex flex-col justify-center">
                                    <span className="font-extrabold text-white text-sm leading-snug truncate">
                                      {categoryTitle}
                                    </span>
                                    <div className="flex items-center gap-1.5 min-w-0 mt-0.5">
                                      <AccountBadgeIcon account={acc} size="xs" />
                                      <span className="text-[11px] font-medium text-slate-400 leading-snug truncate">
                                        {subCategoryLabel || '—'}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                {/* Mobile Only: Amount and Account Column on the Right Side */}
                                <div className="flex md:hidden flex-col items-end justify-center shrink-0 text-right min-w-[100px] ml-2">
                                  <div className={cn("text-sm sm:text-base font-extrabold tabular-nums leading-tight", txDetail.className)}>
                                    {t.type === 'Transfer' ? (
                                      t.transferType === 'from' ? (
                                        <span>-{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                      ) : t.transferType === 'to' ? (
                                        <span>+{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                      ) : (
                                        <span>{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                      )
                                    ) : (
                                      <span>{txDetail.prefix}{formatCurrency(Math.abs(t.amount), acc?.currency || 'BDT')}</span>
                                    )}
                                  </div>
                                  {/* Mobile: Show Account Name */}
                                  <div className="text-[11px] text-slate-400 font-medium leading-tight mt-0.5 truncate max-w-[120px]">
                                    {acc ? acc.name : 'Unknown Account'}
                                  </div>
                                </div>
                              </div>

                              {/* 2nd Column (Web Only): Description */}
                              <div className="hidden md:flex md:col-span-3 items-center min-w-0 text-xs font-semibold text-slate-400 truncate pr-2" title={t.description}>
                                {t.description ? (
                                  <span className="truncate text-slate-300 font-medium">{t.description}</span>
                                ) : (
                                  <span className="text-slate-600 font-normal">—</span>
                                )}
                              </div>

                              {/* 3rd Column (Web Only): Account Name */}
                              <div className="hidden md:flex md:col-span-2 items-center min-w-0 text-xs font-bold truncate">
                                <span className={cn("truncate font-bold", getCurrencyColorClass(acc?.currency))}>
                                  {acc ? acc.name : 'Unknown Account'}
                                </span>
                              </div>

                              {/* 4th Column (Web Only): Amount & Amount (BDT) */}
                              <div className="hidden md:flex md:col-span-2 flex-col items-end justify-center text-right pr-1">
                                <div className={cn("text-sm font-extrabold tabular-nums leading-snug", txDetail.className)}>
                                  {t.type === 'Transfer' ? (
                                    t.transferType === 'from' ? (
                                      <span>-{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                    ) : t.transferType === 'to' ? (
                                      <span>+{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                    ) : (
                                      <div className="flex flex-col items-end gap-0.5">
                                        <span className="text-rose-400 text-[10px] tabular-nums">-{formatCurrency(t.amount, acc?.currency || 'BDT')}</span>
                                        <span className="text-emerald-400 text-[10px] tabular-nums">+{formatCurrency(t.toAmount ?? t.amount, toAcc?.currency || 'BDT')}</span>
                                      </div>
                                    )
                                  ) : (
                                    <span>{txDetail.prefix}{formatCurrency(Math.abs(t.amount), acc?.currency || 'BDT')}</span>
                                  )}
                                </div>

                                {/* Web: Show 2nd line (BDT conversion) ONLY when currency is not BDT */}
                                {(acc?.currency !== 'BDT' || (toAcc && toAcc.currency !== 'BDT')) && (
                                  <div className="text-[11px] text-slate-500 font-medium tabular-nums leading-snug mt-0.5">
                                    {t.type === 'Transfer' ? (
                                      t.transferType === 'from' ? (
                                        <span className="text-rose-400/80">-{formatBDT(t.amount * mult)}</span>
                                      ) : t.transferType === 'to' ? (
                                        <span className="text-emerald-400/80">+{formatBDT(t.amount * mult)}</span>
                                      ) : (
                                        <span className="text-slate-500">{formatBDT(t.amount * mult)} BDT</span>
                                      )
                            ) : (
                              <span>{txDetail.prefix}{formatBDT(Math.abs(t.amount) * mult)}</span>
                            )}
                                  </div>
                                )}
                              </div>

                              {/* Actions Column (Web & Mobile Hover/Tap) */}
                              <div 
                                onClick={(e) => e.stopPropagation()} 
                                className="hidden md:flex md:col-span-1 items-center justify-center gap-1 opacity-40 group-hover:opacity-100 transition-opacity"
                              >
                                <button 
                                  onClick={() => handleEditTx(t)} 
                                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-teal-400 transition-colors"
                                  title="Edit transaction"
                                >
                                  <Edit2 size={13} />
                                </button>
                                <button 
                                  onClick={() => handleDeleteTx(t.id)} 
                                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-rose-500 transition-colors"
                                  title="Delete transaction"
                                >
                                  <Trash2 size={13} />
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}

                {filteredTabTransactions.length === 0 && (
                  <div className="py-16 px-4 text-center text-slate-500 italic uppercase">
                    No transactions matched your filtering criteria
                  </div>
                )}
              </div>

              {/* Total During Selected Period Footer */}
              {isTypeOrCatFiltered && (
                <div className="px-5 sm:px-6 py-4 bg-slate-950/70 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <span className="text-xs sm:text-sm font-bold text-slate-300">
                      Total During {selectedPeriodLabel}:
                    </span>
                  </div>
                  <div className={cn("text-sm sm:text-base font-extrabold tabular-nums", txAccFilter !== 'All' && selectedAcc ? getCurrencyColorClass(selectedAcc.currency) : "text-teal-400")}>
                    {txAccFilter !== 'All' && selectedAcc
                      ? formatCurrency(periodTotalFiltered, selectedAcc.currency)
                      : formatBDT(periodTotalFiltered)}
                  </div>
                </div>
              )}
            </div>
          </div>
        );
      }

      // ─── ANALYTICS MODULE ────────────────────────────────────────────────
      case 'analytics': {
        const FLOW_TEAL = '#2dd4bf';
        const FLOW_ROSE = '#f43f5e';
        const FLOW_PINK = '#f472b6';

        const ANALYTICS_FLOW_LEGENDS = [
          { key: 'income', label: 'Income', color: FLOW_TEAL },
          { key: 'expense', label: 'Expense', color: FLOW_ROSE },
          { key: 'investment', label: 'Investment', color: FLOW_PINK },
        ];

        return (
          <div className="space-y-6 sm:space-y-8">
            {/* ── Row 1: Income vs Expense flow (50%) + Asset by Accounts (50%) ── */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 sm:gap-8 w-full items-stretch">
              {/* ── Income vs Expense flow (BDT value) Chart ── */}
              {(() => {
                const isMobile = typeof window !== 'undefined' && window.innerWidth < 640;

                // Apply monthly or cumulative aggregation
                const baseFlowData = (() => {
                  if (flowView === 'cumulative') {
                    let runInc = 0;
                    let runExp = 0;
                    let runInv = 0;
                    return analyticsData.map(g => {
                      runInc += g.income;
                      runExp += g.expense;
                      runInv += g.investment;
                      return {
                        ...g,
                        income: runInc,
                        expense: runExp,
                        investment: runInv,
                      };
                    });
                  }
                  return analyticsData;
                })();

                // Build visible chart data with hidden series zeroed out
                const visibleChartData = baseFlowData.map(g => ({
                  ...g,
                  income: hiddenFlowSeries.has('income') ? undefined : g.income,
                  expense: hiddenFlowSeries.has('expense') ? undefined : g.expense,
                  investment: hiddenFlowSeries.has('investment') ? undefined : g.investment,
                  _income: g.income,
                  _expense: g.expense,
                  _investment: g.investment,
                }));

                const allVals = baseFlowData.flatMap(g => [
                  hiddenFlowSeries.has('income') ? 0 : g.income,
                  (hiddenFlowSeries.has('expense') ? 0 : g.expense) + (hiddenFlowSeries.has('investment') ? 0 : g.investment),
                ]);
                const dataMax = Math.max(...allVals, 0);
                const step = dataMax > 1000000 ? 200000 : dataMax > 500000 ? 100000 : dataMax > 200000 ? 50000 : dataMax > 50000 ? 20000 : 10000;
                const maxTick = Math.ceil(dataMax / step) * step + step;
                const analyticsTicks: number[] = [];
                for (let v = 0; v <= maxTick; v += step) {
                  analyticsTicks.push(v);
                }

                const fmtAxis = (val: number) => {
                  const abs = Math.abs(val);
                  const sign = val < 0 ? '-' : '';
                  if (abs === 0) return '0';
                  if (abs >= 1000000) {
                    const m = abs / 1000000;
                    return `${sign}${m % 1 === 0 ? m.toFixed(0) : m.toFixed(1)}M`;
                  }
                  if (abs >= 1000) {
                    const k = abs / 1000;
                    return `${sign}${k % 1 === 0 ? k.toFixed(0) : k.toFixed(1)}K`;
                  }
                  return `${sign}${abs.toFixed(0)}`;
                };

                return (
                  <Card className="p-4 sm:p-6 overflow-hidden bg-slate-900 border-slate-800 rounded-2xl shadow-xl h-full flex flex-col justify-between">
                    {/* Title Header matching DSE Tracker */}
                    <div className="mb-5 pb-3 border-b border-slate-800 flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <BarChart2 className="text-teal-400" size={20} />
                        <h3 className="text-subheading font-bold text-white uppercase tracking-wider">
                          Income vs Expense (BDT)
                        </h3>
                      </div>
                    </div>

                  {/* View Selector: Monthly/Cumulative & Range Selector */}
                  <div className="flex flex-wrap items-center justify-center gap-3 mb-4">
                    <div className="flex items-center h-9 bg-slate-950 rounded-xl border border-slate-800 overflow-hidden">
                      <button
                        onClick={() => setFlowView('monthly')}
                        className={cn(
                          "flex items-center gap-2 px-4 h-full text-[10px] font-bold uppercase transition-all tracking-wider border-r border-slate-800",
                          flowView === 'monthly' ? "bg-teal-400 text-slate-950" : "text-slate-400 hover:text-white hover:bg-slate-800/60"
                        )}
                      >
                        <span className={cn("w-2 h-2 rounded-full shrink-0", flowView === 'monthly' ? "bg-slate-950/60" : "bg-teal-400")} />
                        Monthly
                      </button>
                      <button
                        onClick={() => setFlowView('cumulative')}
                        className={cn(
                          "flex items-center gap-2 px-4 h-full text-[10px] font-bold uppercase transition-all tracking-wider",
                          flowView === 'cumulative' ? "bg-teal-400 text-slate-950" : "text-slate-400 hover:text-white hover:bg-slate-800/60"
                        )}
                      >
                        <span className={cn("w-2 h-2 rounded-full shrink-0", flowView === 'cumulative' ? "bg-slate-950/60" : "bg-teal-400")} />
                        Cumulative
                      </button>
                    </div>

                    {/* Range Selector with < and > */}
                    <div className="flex items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() => navigateFlowDate(-1)}
                        className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-950 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
                        title="Previous period"
                      >
                        <ChevronLeft size={14} />
                      </button>

                      <div className="relative">
                        <button
                          onClick={() => setIsFlowRangeOpen(!isFlowRangeOpen)}
                          className="flex items-center gap-3 bg-slate-950 border border-slate-800 rounded-xl px-4 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase tracking-widest"
                        >
                          <Calendar size={14} className="text-teal-400" />
                          {flowRange === 'last6m' ? 'Last 6M' : flowRange === 'last12m' ? 'Last 12M' : flowRange === 'fiscal' ? 'Fiscal' : 'Custom'}
                          <ChevronDown size={14} className={cn("transition-transform", isFlowRangeOpen ? "rotate-180" : "")} />
                        </button>
                        {isFlowRangeOpen && (
                          <>
                            <div className="fixed inset-0 z-40" onClick={() => setIsFlowRangeOpen(false)} />
                            <div className="absolute left-0 mt-2 w-48 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
                              {['last6m', 'last12m', 'fiscal', 'custom'].map(id => (
                                <button
                                  key={id}
                                  onClick={() => {
                                    setFlowRange(id as any);
                                    setFlowMonthOffset(0);
                                    setFlowFiscalOffset(0);
                                    setIsFlowRangeOpen(false);
                                  }}
                                  className={cn(
                                    "w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase tracking-widest",
                                    flowRange === id ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800"
                                  )}
                                >
                                  {id === 'last6m' ? 'Last 6M' : id === 'last12m' ? 'Last 12M' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
                                </button>
                              ))}
                            </div>
                          </>
                        )}
                      </div>

                      <button
                        type="button"
                        onClick={() => navigateFlowDate(1)}
                        className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-950 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
                        title="Next period"
                      >
                        <ChevronRight size={14} />
                      </button>
                    </div>

                    {flowRange === 'custom' && (
                      <div className="flex items-center gap-2 animate-in fade-in slide-in-from-left-2">
                        <input
                          type="date"
                          value={flowCustomDates.start}
                          onChange={e => setFlowCustomDates(p => ({ ...p, start: e.target.value }))}
                          className="bg-slate-950 border border-slate-800 rounded-xl px-3 h-8 text-[10px] font-bold text-white outline-none focus:border-teal-400/50"
                        />
                        <span className="text-slate-600 text-[10px] font-bold">TO</span>
                        <input
                          type="date"
                          value={flowCustomDates.end}
                          onChange={e => setFlowCustomDates(p => ({ ...p, end: e.target.value }))}
                          className="bg-slate-950 border border-slate-800 rounded-xl px-3 h-8 text-[10px] font-bold text-white outline-none focus:border-teal-400/50"
                        />
                      </div>
                    )}
                  </div>

                  {visibleChartData.length === 0 ? (
                    <div className="py-20 text-center text-slate-500 uppercase italic text-[11px] font-bold tracking-wider">
                      Insufficient ledger history in selected date range
                    </div>
                  ) : (
                    <>
                      <div className="h-[360px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart
                            data={visibleChartData}
                            margin={{ top: 38, right: isMobile ? 14 : 20, left: isMobile ? -10 : 8, bottom: 18 }}
                            barGap={2}
                            barCategoryGap="25%"
                          >
                            <CartesianGrid stroke="#334155" strokeOpacity={0.4} vertical={true} horizontal={true} />
                            <XAxis
                              dataKey="label"
                              stroke="#94a3b8"
                              fontSize={11}
                              fontWeight={700}
                              tickLine={false}
                              axisLine={false}
                              tickFormatter={(val) => String(val).toUpperCase()}
                              tick={{
                                dy: 10,
                                fill: '#94a3b8',
                                fontSize: 11,
                                fontWeight: 700,
                                style: { letterSpacing: '0.05em' }
                              }}
                            />
                            <YAxis
                              stroke="#475569"
                              fontSize={10}
                              fontWeight={700}
                              tickLine={false}
                              axisLine={false}
                              ticks={analyticsTicks}
                              tickFormatter={fmtAxis}
                              tick={{ fill: '#475569', fontSize: 10, fontWeight: 700 }}
                              width={52}
                            />
                            <Tooltip
                              cursor={{ fill: '#ffffff08' }}
                              content={({ active, payload, label }) => {
                                if (!active || !payload?.length) return null;
                                const dData = payload[0]?.payload;
                                const inc = dData?._income ?? 0;
                                const exp = dData?._expense ?? 0;
                                const inv = dData?._investment ?? 0;
                                const totalOutflow = exp + inv;
                                const netSavings = inc - totalOutflow;

                                return (
                                  <div className="bg-[#0f172a]/95 backdrop-blur border border-slate-700/60 p-4 rounded-xl shadow-2xl min-w-[240px]">
                                    <p className="text-[12px] font-bold text-white mb-3 border-b border-slate-700/60 pb-2">
                                      {label}{flowView === 'cumulative' ? ' (Cumulative)' : ''}
                                    </p>
                                    <div className="space-y-2">
                                      {!hiddenFlowSeries.has('income') && (
                                        <div className="flex items-center justify-between gap-8">
                                          <div className="flex items-center gap-2">
                                            <div className="w-2 h-2 rounded-full" style={{ backgroundColor: FLOW_TEAL }} />
                                            <span className="text-[11px] text-slate-300">
                                              {flowView === 'cumulative' ? 'Total Income' : 'Income'}
                                            </span>
                                          </div>
                                          <span className="text-[11px] font-bold text-white tabular-nums">
                                            ৳{Math.round(inc).toLocaleString('en-IN')}
                                          </span>
                                        </div>
                                      )}
                                      {!hiddenFlowSeries.has('expense') && (
                                        <div className="flex items-center justify-between gap-8">
                                          <div className="flex items-center gap-2">
                                            <div className="w-2 h-2 rounded-full" style={{ backgroundColor: FLOW_ROSE }} />
                                            <span className="text-[11px] text-slate-300">
                                              {flowView === 'cumulative' ? 'Total Expense' : 'Expense'}
                                            </span>
                                          </div>
                                          <span className="text-[11px] font-bold text-rose-400 tabular-nums">
                                            ৳{Math.round(exp).toLocaleString('en-IN')}
                                          </span>
                                        </div>
                                      )}
                                      {!hiddenFlowSeries.has('investment') && (
                                        <div className="flex items-center justify-between gap-8">
                                          <div className="flex items-center gap-2">
                                            <div className="w-2 h-2 rounded-full" style={{ backgroundColor: FLOW_PINK }} />
                                            <span className="text-[11px] text-slate-300">
                                              {flowView === 'cumulative' ? 'Total Investment' : 'Investment'}
                                            </span>
                                          </div>
                                          <span className="text-[11px] font-bold text-pink-300 tabular-nums">
                                            ৳{Math.round(inv).toLocaleString('en-IN')}
                                          </span>
                                        </div>
                                      )}
                                      <div className="flex items-center justify-between gap-8 pt-2 border-t border-slate-700/40">
                                        <div className="flex items-center gap-2">
                                          <span className="text-[11px] font-bold text-slate-200">
                                            {flowView === 'cumulative' ? 'Cumulative Net Savings' : 'Net Savings'}
                                          </span>
                                        </div>
                                        <span className="text-[11px] font-bold tabular-nums" style={{ color: netSavings >= 0 ? FLOW_TEAL : '#f87171' }}>
                                          {netSavings >= 0 ? '+' : '-'}৳{Math.abs(Math.round(netSavings)).toLocaleString('en-IN')}
                                        </span>
                                      </div>
                                    </div>
                                  </div>
                                );
                              }}
                            />

                            {/* Left bar: Income */}
                            <Bar
                              dataKey="income"
                              stackId="incomeStack"
                              fill={FLOW_TEAL}
                              radius={[4, 4, 0, 0]}
                              barSize={22}
                              isAnimationActive={false}
                              legendType="none"
                            >
                              <LabelList
                                dataKey="income"
                                position="top"
                                content={(props: any) => {
                                  const { x, y, width, value } = props;
                                  if (!value || hiddenFlowSeries.has('income')) return null;
                                  return (
                                    <g transform={`translate(${x + width / 2},${y - 10})`}>
                                      <text transform="rotate(-90)" fill="#64748b" textAnchor="start" fontSize={10} fontWeight={700} className="font-display">
                                        {fmtAxis(value)}
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>

                            {/* Right bar: Expense (Bottom stack) */}
                            <Bar
                              dataKey="expense"
                              stackId="outflowStack"
                              fill={FLOW_ROSE}
                              radius={[0, 0, 0, 0]}
                              barSize={22}
                              isAnimationActive={false}
                              legendType="none"
                            >
                              <LabelList
                                dataKey="expense"
                                position="top"
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  if (hiddenFlowSeries.has('expense')) return null;
                                  const gData = visibleChartData[index];
                                  const invVal = gData?.investment ?? 0;
                                  // If investment is showing on top, do not duplicate label at expense top
                                  if (invVal > 0 && !hiddenFlowSeries.has('investment')) return null;
                                  if (!value) return null;
                                  return (
                                    <g transform={`translate(${x + width / 2},${y - 10})`}>
                                      <text transform="rotate(-90)" fill="#64748b" textAnchor="start" fontSize={10} fontWeight={700} className="font-display">
                                        {fmtAxis(value)}
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>

                            {/* Right bar: Investment (Top stack) */}
                            <Bar
                              dataKey="investment"
                              stackId="outflowStack"
                              fill={FLOW_PINK}
                              radius={[4, 4, 0, 0]}
                              barSize={22}
                              isAnimationActive={false}
                              legendType="none"
                            >
                              <LabelList
                                dataKey="investment"
                                position="top"
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  if (hiddenFlowSeries.has('investment') && hiddenFlowSeries.has('expense')) return null;
                                  const gData = visibleChartData[index];
                                  const expVal = gData?.expense ?? 0;
                                  const stackTotal = expVal + (value ?? 0);
                                  if (stackTotal <= 0 && !value) return null;
                                  return (
                                    <g transform={`translate(${x + width / 2},${y - 10})`}>
                                      <text transform="rotate(-90)" fill="#64748b" textAnchor="start" fontSize={10} fontWeight={700} className="font-display">
                                        {fmtAxis(stackTotal)}
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>

                      {/* Clickable Legend matching DSE Tracker */}
                      <div className="mt-6 pt-6 border-t border-slate-800/50 flex flex-wrap justify-center gap-x-6 gap-y-3">
                        {ANALYTICS_FLOW_LEGENDS.map(({ key, label, color }) => {
                          const hidden = hiddenFlowSeries.has(key);
                          return (
                            <button
                              key={key}
                              type="button"
                              onClick={() => toggleFlowSeries(key)}
                              className="flex items-center gap-2 transition-opacity hover:opacity-80 active:scale-95"
                              style={{ opacity: hidden ? 0.35 : 1 }}
                            >
                              <div className="w-3 h-3 rounded-sm shrink-0" style={{ backgroundColor: color }} />
                              <span
                                className="text-[12px] font-bold text-slate-300 uppercase tracking-tight whitespace-nowrap"
                                style={{ textDecoration: hidden ? 'line-through' : 'none' }}
                              >
                                {label}
                              </span>
                            </button>
                          );
                        })}
                      </div>
                    </>
                  )}
                </Card>
              );
            })()}

            {/* ── Asset by Accounts (BDT) (50%) ── */}
            <AnalyticsPieCard
              title="Asset by Accounts (BDT)"
              icon={<PieChart size={20} className="text-sky-400" />}
              emptyText="No active account assets available"
              slices={assetSlices}
              excluded={excludedAssetAccounts}
              onToggleExcluded={toggleExcludedAsset}
              centerLabel="TOTAL ASSETS"
              itemNoun="ACCOUNTS"
              legendPosition="bottom"
            />
          </div>

          {/* ── Row 2: Incomes by Category (50%) + Expenses by Category (50%) ── */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 sm:gap-8 w-full items-stretch">
            {/* 1. Incomes by Category (BDT) (50%) */}
            <AnalyticsPieCard
              title="Incomes by Category (BDT)"
              icon={<PieChart size={20} className="text-teal-400" />}
              emptyText="No income logged in selected period"
              slices={incomeSlices}
              excluded={excludedIncomeCats}
              onToggleExcluded={toggleExcludedIncome}
              centerLabel="INCOMES"
              itemNoun="CATEGORIES"
            />

            {/* 2. Expenses by Category (BDT) (50%) */}
            <AnalyticsPieCard
              title="Expenses by Category (BDT)"
              icon={<PieChart size={20} className="text-rose-400" />}
              emptyText="No expenses logged in selected period"
              slices={expenseSlices}
              excluded={excludedExpenseCats}
              onToggleExcluded={toggleExcludedExpense}
              centerLabel="EXPENSES"
              itemNoun="CATEGORIES"
            />
          </div>
        </div>
      );
      }

      // ─── SETTINGS SUBMODULE ──────────────────────────────────────────────
      case 'settings': {
        return (
          <div className="space-y-6 sm:space-y-8">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl">
              <div className="flex justify-between items-center mb-6">
                <div>
                  <h4 className="text-subheading font-bold text-white uppercase tracking-tight">Financial Account Profiles</h4>
                  <p className="text-label text-slate-500 font-semibold uppercase">Manage multi-currency pocket profiles and digital ledgers</p>
                </div>
                <Button size="sm" onClick={() => { setEditingAcc(null); setAccForm({ name: '', currency: 'BDT', initialBalance: '', initialDate: new Date().toISOString().split('T')[0], isParent: false, parentId: '' }); setIsAccModalOpen(true); }}>
                  <Plus size={14} /> Add Profile
                </Button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left">
                  <thead>
                    <tr className="border-b border-slate-800 text-[10px] text-slate-500 font-extrabold tracking-widest uppercase bg-slate-950/20">
                      <th className="px-4 py-3 w-10 text-center"></th>
                      <th className="py-3">NAME</th>
                      <th className="py-3">CURRENCY</th>
                      <th className="py-3">INITIAL DATE</th>
                      <th className="py-3 text-right">START BALANCE</th>
                      <th className="py-3 text-right">CURRENT BALANCE</th>
                      <th className="py-3 text-right">CURRENT IN BDT</th>
                      <th className="py-3 text-center w-32">ACTIONS</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/40">
                    {accounts.map((acc, index) => {
                      const bal = accountBalances[acc.id] || 0;
                      const mult = acc.currency === 'USD' ? USD_rate : acc.currency === 'LYD' ? LYD_rate : 1;
                      return (
                        <tr key={acc.id} draggable={draggableRowId === acc.id} onDragStart={(e) => handleDragStart(e, index)} onDragOver={(e) => handleDragOver(e, index)} onDragEnd={handleDragEnd} className={cn("text-label group hover:bg-slate-800/10 transition-colors duration-150", draggedIndex === index ? "opacity-30 bg-slate-800/40" : "")}>
                          <td className="px-4 py-4 text-center">
                            <button type="button" onMouseDown={() => setDraggableRowId(acc.id)} onMouseLeave={() => { if (draggableRowId === acc.id) setDraggableRowId(null); }} onMouseUp={() => setDraggableRowId(null)} className="cursor-grab active:cursor-grabbing text-slate-600 hover:text-teal-400 p-1 flex items-center justify-center rounded hover:bg-slate-800/30 transition-colors focus:outline-none">
                              <GripVertical size={14} />
                            </button>
                          </td>
                          <td className="py-4 font-bold text-white">
                            <div className="flex items-center gap-2.5">
                              <AccountBadgeIcon account={acc} size="sm" />
                              <span>{acc.name}</span>
                            </div>
                          </td>
                          <td className="py-4">
                            <span className={cn("px-2 py-0.5 rounded text-[8px] font-extrabold uppercase", acc.currency === 'BDT' ? "bg-teal-400/10 text-teal-400 border border-teal-400/20" : acc.currency === 'USD' ? "bg-sky-400/10 text-sky-400 border border-sky-400/20" : "bg-amber-400/10 text-amber-400 border border-amber-400/20")}>{acc.currency}</span>
                          </td>
                          <td className="py-4 text-slate-400 font-semibold tabular-nums text-xs">
                            {acc.initialDate ? formatDate(acc.initialDate) : '—'}
                          </td>
                          <td className="py-4 text-right font-bold text-slate-300 tabular-nums">{formatCurrency(acc.initialBalance, acc.currency)}</td>
                          <td className="py-4 text-right font-extrabold text-white tabular-nums">{formatCurrency(bal, acc.currency)}</td>
                          <td className="py-4 text-right font-extrabold text-teal-400 tabular-nums">{formatBDT(bal * mult)}</td>
                          <td className="py-4 text-center">
                            <div className="flex justify-center gap-2 opacity-60 group-hover:opacity-100 transition-opacity">
                              <button onClick={() => handleEditAcc(acc)} className="p-1 hover:text-teal-400 transition-colors"><Edit2 size={13} /></button>
                              <button onClick={() => handleDeleteAcc(acc.id)} className="p-1 hover:text-rose-500 transition-colors"><Trash2 size={13} /></button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {accounts.length === 0 && (
                      <tr><td colSpan={8} className="py-12 text-center text-slate-500 italic uppercase">No accounts found. Create some above!</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                <div>
                  <div className="flex items-center gap-2.5">
                    <h4 className="text-subheading font-bold text-white uppercase tracking-tight font-display">Ledger Categories</h4>
                    <span className="px-2 py-0.5 rounded text-[9px] font-bold bg-teal-400/10 text-teal-400 border border-teal-400/20 uppercase">
                      Drag & Drop enabled
                    </span>
                  </div>
                  <p className="text-label text-slate-500 font-semibold uppercase mt-1">
                    Drag sub-categories to move them across categories or reorder within a category
                  </p>
                </div>
                <Button size="sm" onClick={() => { setEditingCat(null); setCatForm({ name: '', type: 'Expense', subCategories: '' }); setIsCatModalOpen(true); }}>
                  <Plus size={14} /> Add Category
                </Button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left">
                  <thead>
                    <tr className="border-b border-slate-800 text-[10px] text-slate-500 font-extrabold tracking-widest uppercase bg-slate-950/20">
                      <th className="px-4 py-3 w-10 text-center"></th>
                      <th className="py-3 w-48">CATEGORY NAME</th>
                      <th className="py-3 w-28">TYPE</th>
                      <th className="py-3">SUB-CATEGORIES / SUB-LEDGERS</th>
                      <th className="py-3 text-center w-32">ACTIONS</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/40">
                    {categories.map((cat, index) => {
                      const isCatBeingHoveredForDrop = draggedSubCat && dragOverCatId === cat.id && draggedSubCat.sourceCatId !== cat.id;
                      const catVisual = getCategoryVisual(cat.name, cat.type);

                      return (
                        <tr
                          key={cat.id}
                          draggable={draggableCatId === cat.id}
                          onDragStart={(e) => handleCatDragStart(e, index)}
                          onDragOver={(e) => {
                            if (draggableCatId) {
                              handleCatDragOver(e, index);
                            } else if (draggedSubCat) {
                              e.preventDefault();
                              setDragOverCatId(cat.id);
                            }
                          }}
                          onDragLeave={() => {
                            if (dragOverCatId === cat.id) setDragOverCatId(null);
                          }}
                          onDrop={(e) => {
                            if (draggedSubCat) {
                              e.preventDefault();
                              e.stopPropagation();
                              handleMoveSubCategory(draggedSubCat.sourceCatId, cat.id, draggedSubCat.subCategory, dragOverSubIndex ?? undefined);
                              handleSubCatDragEnd();
                            }
                          }}
                          onDragEnd={handleCatDragEnd}
                          className={cn(
                            "text-label group transition-all duration-150",
                            draggedCatIndex === index ? "opacity-30 bg-slate-800/40" : "",
                            isCatBeingHoveredForDrop ? "bg-teal-950/30 ring-1 ring-inset ring-teal-500/40" : "hover:bg-slate-800/10"
                          )}
                        >
                          <td className="px-4 py-4 text-center">
                            <button
                              type="button"
                              onMouseDown={() => setDraggableCatId(cat.id)}
                              onMouseLeave={() => { if (draggableCatId === cat.id) setDraggableCatId(null); }}
                              onMouseUp={() => setDraggableCatId(null)}
                              className="cursor-grab active:cursor-grabbing text-slate-600 hover:text-teal-400 p-1 flex items-center justify-center rounded hover:bg-slate-800/30 transition-colors focus:outline-none"
                              title="Drag to reorder category"
                            >
                              <GripVertical size={14} />
                            </button>
                          </td>
                          <td className="py-4 font-bold text-white">
                            <div className="flex items-center gap-2">
                              <div className={cn("w-5.5 h-5.5 rounded-md flex items-center justify-center shrink-0 border", catVisual.bg)}>
                                <catVisual.Icon size={12} />
                              </div>
                              <span>{cat.name}</span>
                              {isCatBeingHoveredForDrop && (
                                <span className="text-[9px] font-bold text-teal-400 bg-teal-400/10 px-1.5 py-0.5 rounded animate-pulse">
                                  Drop Target
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="py-4">
                            <span className={cn("px-2 py-0.5 rounded text-[8px] font-extrabold uppercase", cat.type === 'Income' ? "bg-emerald-500/10 text-emerald-400 border border-emerald-400/20" : "bg-rose-500/10 text-rose-400 border border-rose-400/20")}>{cat.type}</span>
                          </td>
                          <td
                            className="py-4 text-slate-400 font-semibold text-xs leading-normal"
                            onDragOver={(e) => {
                              if (draggedSubCat) {
                                e.preventDefault();
                                e.stopPropagation();
                                setDragOverCatId(cat.id);
                              }
                            }}
                            onDrop={(e) => {
                              if (draggedSubCat) {
                                e.preventDefault();
                                e.stopPropagation();
                                handleMoveSubCategory(draggedSubCat.sourceCatId, cat.id, draggedSubCat.subCategory, dragOverSubIndex ?? undefined);
                                handleSubCatDragEnd();
                              }
                            }}
                          >
                            <div className="flex flex-wrap items-center gap-1.5 min-h-[28px]">
                              {cat.subCategories.length > 0 ? (
                                cat.subCategories.map((sub, subIdx) => {
                                  const isDraggingThis = draggedSubCat?.subCategory === sub && draggedSubCat?.sourceCatId === cat.id;
                                  const isDropIndicatorHere = draggedSubCat && dragOverCatId === cat.id && dragOverSubIndex === subIdx && !isDraggingThis;

                                  return (
                                    <div key={sub} className="flex items-center">
                                      {isDropIndicatorHere && (
                                        <span className="w-1 h-5 bg-teal-400 rounded-full mr-1 animate-pulse" />
                                      )}
                                      <div
                                        draggable={true}
                                        onDragStart={(e) => handleSubCatDragStart(e, sub, cat.id, subIdx)}
                                        onDragEnd={handleSubCatDragEnd}
                                        onDragOver={(e) => {
                                          if (draggedSubCat) {
                                            e.preventDefault();
                                            e.stopPropagation();
                                            setDragOverCatId(cat.id);
                                            setDragOverSubIndex(subIdx);
                                          }
                                        }}
                                        onDrop={(e) => {
                                          if (draggedSubCat) {
                                            e.preventDefault();
                                            e.stopPropagation();
                                            handleMoveSubCategory(draggedSubCat.sourceCatId, cat.id, draggedSubCat.subCategory, subIdx);
                                            handleSubCatDragEnd();
                                          }
                                        }}
                                        className={cn(
                                          "inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md border text-[10px] font-medium select-none cursor-grab active:cursor-grabbing transition-all duration-150 group/chip",
                                          isDraggingThis
                                            ? "opacity-30 border-dashed border-teal-400 bg-slate-900"
                                            : "bg-slate-950 border-slate-800/80 text-slate-300 hover:border-teal-500/50 hover:bg-slate-900 hover:text-teal-300 shadow-sm"
                                        )}
                                        title="Drag to move this sub-category to another category or reorder"
                                      >
                                        <GripVertical size={10} className="text-slate-600 group-hover/chip:text-teal-400 shrink-0 transition-colors" />
                                        <span>{sub}</span>
                                      </div>
                                    </div>
                                  );
                                })
                              ) : (
                                <span className={cn("text-[11px] italic transition-colors", isCatBeingHoveredForDrop ? "text-teal-400 font-semibold" : "text-slate-600")}>
                                  {isCatBeingHoveredForDrop ? "Drop sub-category here to add" : "No sub-ledger indices defined (Drag & drop items here)"}
                                </span>
                              )}

                              {/* Show drop target placeholder if dragging from another category */}
                              {isCatBeingHoveredForDrop && dragOverSubIndex === null && (
                                <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-teal-400/15 rounded-md border border-dashed border-teal-400/60 text-[10px] font-bold text-teal-300 animate-pulse">
                                  + Drop "{draggedSubCat.subCategory}" here
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="py-4 text-center">
                            <div className="flex justify-center gap-2 opacity-60 group-hover:opacity-100 transition-opacity">
                              <button onClick={() => handleEditCat(cat)} className="p-1 hover:text-teal-400 transition-colors" title="Edit category"><Edit2 size={13} /></button>
                              <button onClick={() => handleDeleteCat(cat.id)} className="p-1 hover:text-rose-500 transition-colors" title="Delete category"><Trash2 size={13} /></button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {categories.length === 0 && (
                      <tr><td colSpan={5} className="py-12 text-center text-slate-500 italic uppercase">No categories found. Create some above!</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        );
      }
    }
  };

  function formatNumber(num: number, decimals: number = 2): string {
    return new Intl.NumberFormat('en-US', { minimumFractionDigits: 0, maximumFractionDigits: decimals }).format(num);
  }

  // ─── Derived display values for transaction modal ──────────────────────
  const activeCurrency = selectedAccount?.currency || 'BDT';
  const currencySymbol = activeCurrency === 'BDT' ? '৳' : activeCurrency === 'USD' ? '$' : 'LD';
  const currencyColorClass = activeCurrency === 'BDT' ? 'text-teal-400' : activeCurrency === 'USD' ? 'text-sky-400' : 'text-amber-400';
  const currencyBgClass = activeCurrency === 'BDT' ? 'bg-teal-400/10 border-teal-400/20' : activeCurrency === 'USD' ? 'bg-sky-400/10 border-sky-400/20' : 'bg-amber-400/10 border-amber-400/20';

  // Amount is locked (read-only) when editing a transaction synced from the DSE
  // Tracker as a Dividend or a Capital Gain (Stocks Sell P&L) — the real value lives
  // in the DSE Tracker and must be edited there to keep both ledgers in sync.
  const isDseAmountLocked = !!editingTx && !!(inheritedDseTxId || editingTx?.dseTxId) && (
    txForm.category === 'Capital Gain' ||
    txForm.subCategory === 'Stocks Capital Gain' ||
    (txForm.type === 'Income' && (txForm.subCategory || '').toLowerCase().includes('dividend'))
  );

  return (
    <div className="space-y-6 sm:space-y-8">
      {/* Dynamic Header Range & Action Row */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 relative">
        <div className="flex flex-col sm:flex-row sm:items-center gap-3 w-full sm:w-auto">
          <div className="flex items-center justify-between w-full sm:w-auto gap-2">
            <div className="relative flex-1 sm:flex-none">
              <div className="block sm:hidden">
                <button onClick={() => setIsRangeMenuOpen(!isRangeMenuOpen)} className="flex items-center justify-between gap-4 bg-slate-950 border border-slate-800 rounded-lg px-4 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase w-full">
                  <div className="flex items-center gap-2"><Calendar size={14} className="text-teal-400" />{rangeType === 'this' ? 'This month' : rangeType === 'fiscal' ? 'Fiscal' : 'Custom'}</div>
                  <ChevronDown size={14} className={cn("text-slate-500 transition-transform", isRangeMenuOpen ? "rotate-180 text-teal-400" : "")} />
                </button>
                {isRangeMenuOpen && (
                  <>
                    <div className="fixed inset-0 z-40" onClick={() => setIsRangeMenuOpen(false)} />
                    <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
                      {(['this', 'fiscal', 'custom'] as const).map((id) => (
                        <button key={id} onClick={() => { setRangeType(id); setIsRangeMenuOpen(false); }} className={cn("w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase", rangeType === id ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800")}>
                          {id === 'this' ? 'This month' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
              <div className="hidden sm:flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1 font-sans">
                {(['this', 'fiscal', 'custom'] as const).map(id => (
                  <button key={id} onClick={() => setRangeType(id)} className={cn("px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all border", rangeType === id ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20" : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white")}>
                    {id === 'this' ? 'This month' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
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
                        onClick={() => { handleExportExcel(); setIsSettingsMenuOpen(false); }}
                        className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                      >
                        <Download size={14} className="text-teal-400" />
                        EXPORT
                      </button>
                      <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
                        <Upload size={14} className="text-teal-400" />
                        IMPORT
                        <input type="file" className="hidden" accept=".xlsx,.xls,.csv" onChange={(e) => { setIsSettingsMenuOpen(false); handleImportExcel(e); }} />
                      </label>
                      <button 
                        onClick={() => { handleDownloadTemplate(); setIsSettingsMenuOpen(false); }}
                        className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                      >
                        <FileSpreadsheet size={14} className="text-teal-400" />
                        TEMPLATE
                      </button>
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>

          <div className="flex items-center gap-1.5 px-1 animate-in fade-in duration-300">
            <button onClick={() => { if (rangeType === 'this') navigateThisMonth(-1); else if (rangeType === 'fiscal') setFiscalStartYear(prev => prev - 1); else navigateCustomMonth(-1); }} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"><ChevronLeft size={14} /></button>
            {rangeType === 'custom' ? (
              <div className="flex items-center gap-1.5">
                <input type="date" value={customDates.start} onChange={(e) => setCustomDates(prev => ({ ...prev, start: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer" />
                <span className="text-slate-700 font-bold text-[10px]">–</span>
                <input type="date" value={customDates.end} onChange={(e) => setCustomDates(prev => ({ ...prev, end: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none uppercase cursor-pointer" />
              </div>
            ) : rangeType === 'fiscal' ? (
              <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-white uppercase select-none tracking-wider whitespace-nowrap min-w-[160px] text-center">July {fiscalStartYear} - June {fiscalStartYear + 1}</div>
            ) : (
              <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-[10px] font-bold text-teal-400 uppercase select-none tracking-wider whitespace-nowrap min-w-[120px] text-center">{formatThisMonthLabel(thisMonthDate)}</div>
            )}
            <button onClick={() => { if (rangeType === 'this') navigateThisMonth(1); else if (rangeType === 'fiscal') setFiscalStartYear(prev => prev + 1); else navigateCustomMonth(1); }} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"><ChevronRight size={14} /></button>
          </div>
        </div>

        {/* Settings Button - Desktop Only */}
        <div className="hidden sm:block">
          <div className="relative">
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
                    onClick={() => { handleExportExcel(); setIsSettingsMenuOpen(false); }}
                    className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                  >
                    <Download size={14} className="text-teal-400" />
                    EXPORT
                  </button>
                  <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
                    <Upload size={14} className="text-teal-400" />
                    IMPORT
                    <input type="file" className="hidden" accept=".xlsx,.xls,.csv" onChange={(e) => { setIsSettingsMenuOpen(false); handleImportExcel(e); }} />
                  </label>
                  <button 
                    onClick={() => { handleDownloadTemplate(); setIsSettingsMenuOpen(false); }}
                    className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase"
                  >
                    <FileSpreadsheet size={14} className="text-teal-400" />
                    TEMPLATE
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Main Switch Sub-view Dashboard */}
      {renderSubmodule()}

      {/* ─── REDESIGNED TRANSACTION MODAL ─────────────────────────────────── */}
      <Modal
        isOpen={isTxModalOpen}
        onClose={() => { setIsTxModalOpen(false); setEditingTx(null); isFromExternalModuleRef.current = false; }}
        title={editingTx ? "Edit Transaction" : "New Transaction"}
      >
        <form onSubmit={handleSaveTx} className="space-y-5">

          {/* Row 1: Date */}
          <Input
            label="Transaction Date"
            type="date"
            required
            value={txForm.date}
            onChange={(e) => setTxForm(prev => ({ ...prev, date: e.target.value }))}
          />

          {/* Row 2: Type — 4-way toggle */}
          <div>
            <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
              Transaction Type
            </label>
            <div className="grid grid-cols-4 gap-1.5">
              {TX_TYPES.map(typeOpt => (
                <button
                  key={typeOpt.value}
                  type="button"
                  onClick={() => handleTypeChange(typeOpt.value as any)}
                  className={cn(
                    "relative flex flex-col items-center justify-center gap-1.5 py-2.5 px-1 rounded-xl border transition-all duration-150 font-bold text-[10px] uppercase tracking-wide",
                    txForm.type === typeOpt.value
                      ? typeOpt.activeClasses
                      : typeOpt.inactiveClasses
                  )}
                >
                  <span className={cn("w-1.5 h-1.5 rounded-full", txForm.type === typeOpt.value ? typeOpt.dotColor : "bg-slate-600")} />
                  <span>{typeOpt.label}</span>
                  {txForm.type === typeOpt.value && (
                    <span className="absolute top-1.5 right-1.5">
                      <Check size={9} />
                    </span>
                  )}
                </button>
              ))}
            </div>
          </div>

          {/* Row 3 & 4: Dual Account and Amount Pickers for Transfer vs. Single Picker for Income/Expense */}
          {txForm.type === 'Transfer' ? (
            <div className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                {/* Column 1: Transfer From Account & Amount */}
                <div className="space-y-4">
                  {/* Account Picker Card ("Transfer From") */}
                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
                      Transfer From Account
                    </label>
                    <button
                      type="button"
                      onClick={() => setIsAccountPickerOpen(true)}
                      className={cn(
                        "w-full flex items-center gap-3 p-3 rounded-xl border transition-all duration-150 text-left group",
                        txForm.accountId
                          ? "bg-slate-900 border-slate-700 hover:border-slate-600"
                          : "bg-slate-950 border-slate-800 border-dashed hover:border-slate-600"
                      )}
                    >
                      {txForm.accountId && selectedAccount ? (
                        <AccountBadgeIcon account={selectedAccount} size="lg" />
                      ) : (
                        <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0 transition-all bg-slate-800 text-slate-500 group-hover:bg-slate-700">
                          <CreditCard size={14} />
                        </div>
                      )}
                      <div className="flex-1 min-w-0">
                        {txForm.accountId && selectedAccount ? (
                          <>
                            <div className={cn(
                              "font-bold text-[11px] truncate",
                              selectedAccount.currency === 'BDT' ? "text-teal-300" : selectedAccount.currency === 'LYD' ? "text-amber-300" : "text-sky-300"
                            )}>
                              {selectedAccount.name}
                            </div>
                            <div className={cn("text-[9px] font-extrabold uppercase mt-0.5", selectedAccount.currency === 'BDT' ? "text-teal-500" : selectedAccount.currency === 'LYD' ? "text-amber-500" : "text-sky-500")}>
                              {selectedAccount.currency}
                            </div>
                          </>
                        ) : (
                          <span className="text-[11px] text-slate-500 font-semibold">Select Account</span>
                        )}
                      </div>
                      <ChevronDown size={12} className="text-slate-600 group-hover:text-slate-400 shrink-0 transition-colors" />
                    </button>
                  </div>

                  {/* Amount input box ("Transfer From Amount") */}
                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
                      Transfer From Amount ({selectedAccount?.currency || 'BDT'})
                    </label>
                    <div className="relative">
                      <div className={cn(
                        "absolute left-3 top-1/2 -translate-y-1/2 flex items-center gap-1 px-1.5 py-0.5 rounded border text-[9px] font-extrabold uppercase select-none",
                        selectedAccount?.currency === 'BDT' ? "bg-teal-400/10 text-teal-400 border-teal-500/20" :
                        selectedAccount?.currency === 'LYD' ? "bg-amber-400/10 text-amber-400 border-amber-500/20" :
                        "bg-sky-400/10 text-sky-400 border-sky-500/20"
                      )}>
                        <span>{selectedAccount?.currency === 'BDT' ? '৳' : selectedAccount?.currency === 'LYD' ? 'LD' : '$'}</span>
                        <span>{selectedAccount?.currency || 'BDT'}</span>
                      </div>
                      <input
                        type="number"
                        step="any"
                        required
                        placeholder="0.00"
                        value={txForm.amount}
                        onChange={(e) => {
                          const val = e.target.value;
                          setTxForm(prev => ({ 
                            ...prev, 
                            amount: val, 
                            toAmount: prev.toAmount === prev.amount || !prev.toAmount ? val : prev.toAmount 
                          }));
                        }}
                        className={cn(
                          "w-full bg-slate-950 border border-slate-800 rounded-xl pl-20 pr-4 py-2.5 text-sm font-extrabold text-white placeholder-slate-700 outline-none tabular-nums transition-all",
                          "focus:border-slate-600 focus:ring-1 focus:ring-slate-600/40",
                          selectedAccount?.currency === 'BDT' ? "focus:border-teal-500/40 focus:ring-teal-500/10" :
                          selectedAccount?.currency === 'LYD' ? "focus:border-amber-500/40 focus:ring-amber-500/10" :
                          "focus:border-sky-500/40 focus:ring-sky-500/10"
                        )}
                      />
                    </div>
                  </div>
                </div>

                {/* Column 2: Transfer To Account & Amount */}
                <div className="space-y-4">
                  {/* Account Picker Card ("Transfer To") */}
                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
                      Transfer To Account
                    </label>
                    <button
                      type="button"
                      onClick={() => setIsToAccountPickerOpen(true)}
                      className={cn(
                        "w-full flex items-center gap-3 p-3 rounded-xl border transition-all duration-150 text-left group",
                        txForm.toAccountId
                          ? "bg-slate-900 border-slate-700 hover:border-slate-600"
                          : "bg-slate-950 border-slate-800 border-dashed hover:border-slate-600"
                      )}
                    >
                      {txForm.toAccountId && accounts.find(a => a.id === txForm.toAccountId) ? (
                        <AccountBadgeIcon account={accounts.find(a => a.id === txForm.toAccountId)} size="lg" />
                      ) : (
                        <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0 transition-all bg-slate-800 text-slate-500 group-hover:bg-slate-700">
                          <CreditCard size={14} />
                        </div>
                      )}
                      <div className="flex-1 min-w-0">
                        {txForm.toAccountId && accounts.find(a => a.id === txForm.toAccountId) ? (
                          <>
                            <div className={cn(
                              "font-bold text-[11px] truncate",
                              accounts.find(a => a.id === txForm.toAccountId)?.currency === 'BDT' ? "text-teal-300" : accounts.find(a => a.id === txForm.toAccountId)?.currency === 'LYD' ? "text-amber-300" : "text-sky-300"
                            )}>
                              {accounts.find(a => a.id === txForm.toAccountId)?.name}
                            </div>
                            <div className={cn("text-[9px] font-extrabold uppercase mt-0.5", accounts.find(a => a.id === txForm.toAccountId)?.currency === 'BDT' ? "text-teal-500" : accounts.find(a => a.id === txForm.toAccountId)?.currency === 'LYD' ? "text-amber-500" : "text-sky-500")}>
                              {accounts.find(a => a.id === txForm.toAccountId)?.currency}
                            </div>
                          </>
                        ) : (
                          <span className="text-[11px] text-slate-500 font-semibold">Select Account</span>
                        )}
                      </div>
                      <ChevronDown size={12} className="text-slate-600 group-hover:text-slate-400 shrink-0 transition-colors" />
                    </button>
                  </div>

                  {/* Amount input box ("Transfer To Amount") */}
                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
                      Transfer To Amount ({accounts.find(a => a.id === txForm.toAccountId)?.currency || 'BDT'})
                    </label>
                    <div className="relative">
                      <div className={cn(
                        "absolute left-3 top-1/2 -translate-y-1/2 flex items-center gap-1 px-1.5 py-0.5 rounded border text-[9px] font-extrabold uppercase select-none",
                        accounts.find(a => a.id === txForm.toAccountId)?.currency === 'BDT' ? "bg-teal-400/10 text-teal-400 border-teal-500/20" :
                        accounts.find(a => a.id === txForm.toAccountId)?.currency === 'LYD' ? "bg-amber-400/10 text-amber-400 border-amber-500/20" :
                        "bg-sky-400/10 text-sky-400 border-sky-500/20"
                      )}>
                        <span>{accounts.find(a => a.id === txForm.toAccountId)?.currency === 'BDT' ? '৳' : accounts.find(a => a.id === txForm.toAccountId)?.currency === 'LYD' ? 'LD' : '$'}</span>
                        <span>{accounts.find(a => a.id === txForm.toAccountId)?.currency || 'BDT'}</span>
                      </div>
                      <input
                        type="number"
                        step="any"
                        required
                        placeholder="0.00"
                        value={txForm.toAmount || txForm.amount}
                        onChange={(e) => setTxForm(prev => ({ ...prev, toAmount: e.target.value }))}
                        className={cn(
                          "w-full bg-slate-950 border border-slate-800 rounded-xl pl-20 pr-4 py-2.5 text-sm font-extrabold text-white placeholder-slate-700 outline-none tabular-nums transition-all",
                          "focus:border-slate-600 focus:ring-1 focus:ring-slate-600/40",
                          accounts.find(a => a.id === txForm.toAccountId)?.currency === 'BDT' ? "focus:border-teal-500/40 focus:ring-teal-500/10" :
                          accounts.find(a => a.id === txForm.toAccountId)?.currency === 'LYD' ? "focus:border-amber-500/40 focus:ring-amber-500/10" :
                          "focus:border-sky-500/40 focus:ring-sky-500/10"
                        )}
                      />
                    </div>
                  </div>
                </div>
              </div>

            </div>
          ) : (
            <>
              {/* Row 3: Category + Account picker cards */}
              <div className="grid grid-cols-2 gap-3">
                {/* Category Picker Card */}
                <div>
                  <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
                    Category
                  </label>
                  <button
                    type="button"
                    onClick={() => setIsCategoryPickerOpen(true)}
                    className={cn(
                      "w-full flex items-center gap-3 p-3 rounded-xl border transition-all duration-150 text-left group",
                      txForm.category
                        ? "bg-slate-900 border-slate-700 hover:border-slate-600"
                        : "bg-slate-950 border-slate-800 border-dashed hover:border-slate-600"
                    )}
                  >
                    <div className={cn(
                      "w-8 h-8 rounded-lg flex items-center justify-center shrink-0 transition-all",
                      txForm.category
                        ? (txForm.type === 'Loan' || txForm.category === 'Loan')
                          ? "bg-amber-400/15 text-amber-400"
                          : selectedCategoryObj?.type === 'Income'
                            ? "bg-emerald-400/15 text-emerald-400"
                            : "bg-rose-400/15 text-rose-400"
                        : "bg-slate-800 text-slate-500 group-hover:bg-slate-700"
                    )}>
                      <Tag size={14} />
                    </div>
                    <div className="flex-1 min-w-0">
                      {txForm.category ? (
                        <>
                          <div className={cn(
                            "font-bold text-[11px] truncate",
                            (txForm.type === 'Loan' || txForm.category === 'Loan')
                              ? "text-amber-300"
                              : selectedCategoryObj?.type === 'Income' ? "text-emerald-300" : "text-rose-300"
                          )}>
                            {txForm.category}
                          </div>
                          {txForm.subCategory && (
                            <div className="text-[9px] text-slate-500 font-semibold truncate mt-0.5">
                              {txForm.subCategory}
                            </div>
                          )}
                        </>
                      ) : (
                        <span className="text-[11px] text-slate-500 font-semibold">Select Category</span>
                      )}
                    </div>
                    <ChevronDown size={12} className="text-slate-600 group-hover:text-slate-400 shrink-0 transition-colors" />
                  </button>
                </div>

                {/* Account Picker Card */}
                <div>
                  <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
                    Account
                  </label>
                  <button
                    type="button"
                    onClick={() => setIsAccountPickerOpen(true)}
                    className={cn(
                      "w-full flex items-center gap-3 p-3 rounded-xl border transition-all duration-150 text-left group",
                      txForm.accountId
                        ? "bg-slate-900 border-slate-700 hover:border-slate-600"
                        : "bg-slate-950 border-slate-800 border-dashed hover:border-slate-600"
                    )}
                  >
                    {txForm.accountId && selectedAccount ? (
                      <AccountBadgeIcon account={selectedAccount} size="lg" />
                    ) : (
                      <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0 transition-all bg-slate-800 text-slate-500 group-hover:bg-slate-700">
                        <CreditCard size={14} />
                      </div>
                    )}
                    <div className="flex-1 min-w-0">
                      {txForm.accountId && selectedAccount ? (
                        <>
                          <div className={cn(
                            "font-bold text-[11px] truncate",
                            selectedAccount.currency === 'BDT' ? "text-teal-300" : selectedAccount.currency === 'LYD' ? "text-amber-300" : "text-sky-300"
                          )}>
                            {selectedAccount.name}
                          </div>
                          <div className={cn("text-[9px] font-extrabold uppercase mt-0.5", selectedAccount.currency === 'BDT' ? "text-teal-500" : selectedAccount.currency === 'LYD' ? "text-amber-500" : "text-sky-500")}>
                            {selectedAccount.currency}
                          </div>
                        </>
                      ) : (
                        <span className="text-[11px] text-slate-500 font-semibold">Select Account</span>
                      )}
                    </div>
                    <ChevronDown size={12} className="text-slate-600 group-hover:text-slate-400 shrink-0 transition-colors" />
                  </button>
                </div>
              </div>

              {/* Row 4: Amount with dynamic currency label */}
              <div>
                <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-2">
                  Amount ({activeCurrency})
                </label>
                <div className="relative">
                  {/* Currency badge on the left */}
                  <div className={cn(
                    "absolute left-3 top-1/2 -translate-y-1/2 flex items-center gap-1.5 px-2 py-1 rounded-md border text-[10px] font-extrabold uppercase select-none",
                    currencyBgClass,
                    currencyColorClass
                  )}>
                    <span>{currencySymbol}</span>
                    <span>{activeCurrency}</span>
                  </div>
                  <input
                    type="number"
                    step="any"
                    required
                    placeholder="0.00"
                    value={txForm.amount}
                    readOnly={isDseAmountLocked}
                    onChange={(e) => {
                      if (isDseAmountLocked) return;
                      setTxForm(prev => ({ ...prev, amount: e.target.value }));
                    }}
                    className={cn(
                      "w-full bg-slate-950 border border-slate-800 rounded-xl pl-24 pr-4 py-3 text-base font-extrabold text-white placeholder-slate-700 outline-none tabular-nums transition-all",
                      isDseAmountLocked
                        ? "opacity-60 cursor-not-allowed text-slate-400"
                        : cn(
                            "focus:border-slate-600 focus:ring-1 focus:ring-slate-600/40",
                            activeCurrency === 'BDT' ? "focus:border-teal-500/40 focus:ring-teal-500/10" :
                            activeCurrency === 'LYD' ? "focus:border-amber-500/40 focus:ring-amber-500/10" :
                            "focus:border-sky-500/40 focus:ring-sky-500/10"
                          )
                    )}
                  />
                  {/* BDT equivalent hint if not BDT */}
                  {txForm.amount && activeCurrency !== 'BDT' && (
                    <div className="absolute right-3 top-1/2 -translate-y-1/2 text-[10px] text-slate-500 font-bold tabular-nums">
                      ≈ {formatBDT(parseFloat(txForm.amount || '0') * (activeCurrency === 'USD' ? USD_rate : LYD_rate))}
                    </div>
                  )}
                </div>
                {isDseAmountLocked && (
                  <p className="mt-1.5 text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                    Locked — edit this amount from the DSE Tracker instead
                  </p>
                )}
              </div>
            </>
          )}

          {/* Sync with DSE Tracker option ONLY for transactions related to DSE Tracker module */}
          {(() => {
            const isNonDseLinked = !!inheritedOnlineInstallmentTxId || !!inheritedOnlineGroupId || !!editingTx?.onlineTxId ||
              !!inheritedSukukGroupId || !!inheritedSukukRentTxId || !!editingTx?.sukukTxId ||
              !!inheritedFdrProfitTxId || !!editingTx?.fdrTxId ||
              !!inheritedMfGroupId || !!editingTx?.mfTxId ||
              isMutualFundAccount(accounts.find(a => a.id === txForm.accountId)) ||
              isMutualFundAccount(accounts.find(a => a.id === txForm.toAccountId)) ||
              isOnlineInvestmentAccount(accounts.find(a => a.id === txForm.accountId)) ||
              isOnlineInvestmentAccount(accounts.find(a => a.id === txForm.toAccountId)) ||
              isSukukAccount(accounts.find(a => a.id === txForm.accountId)) ||
              isSukukAccount(accounts.find(a => a.id === txForm.toAccountId)) ||
              isFdrAccount(accounts.find(a => a.id === txForm.accountId)) ||
              isFdrAccount(accounts.find(a => a.id === txForm.toAccountId));

            if (isNonDseLinked) return false;

            const fromAcc = accounts.find(a => a.id === txForm.accountId);
            const toAcc = accounts.find(a => a.id === txForm.toAccountId);

            if (txForm.type === 'Transfer') {
              return isBoAccount(fromAcc) || isBoAccount(toAcc);
            }

            if (txForm.type === 'Income') {
              return !!inheritedDseTxId || !!editingTx?.dseTxId || isBoAccount(fromAcc) || txForm.category === 'Capital Gain' || txForm.subCategory === 'Stocks Capital Gain';
            }

            if (txForm.type === 'Expense') {
              return !!inheritedDseTxId || !!editingTx?.dseTxId || isBoAccount(fromAcc);
            }

            return false;
          })() && (
            <div
              id="dse-auto-update-option"
              onClick={() => setSyncToDse(prev => !prev)}
              className={cn(
                "w-full p-3 rounded-xl border transition-all cursor-pointer flex items-center justify-between gap-3 select-none",
                syncToDse
                  ? "bg-teal-950/40 border-teal-500/40 text-teal-300 shadow-[0_0_15px_rgba(45,212,191,0.08)]"
                  : "bg-slate-950/60 border-slate-800 text-slate-400 hover:border-slate-700"
              )}
            >
              <div className="min-w-0 flex-1">
                <span className="text-[11px] font-bold text-white uppercase tracking-tight whitespace-nowrap">
                  Sync with DSE Tracker {
                    txForm.type === 'Income' && (txForm.category === 'Capital Gain' || txForm.subCategory === 'Stocks Capital Gain')
                      ? '(Stocks P&L)'
                      : txForm.type === 'Income'
                        ? '(Dividend)'
                        : txForm.type === 'Expense'
                          ? '(Charge)'
                          : ''
                  }
                </span>
              </div>
              <div className="shrink-0 pl-2" onClick={(e) => e.stopPropagation()}>
                <Checkbox
                  id="sync-to-dse-checkbox"
                  checked={syncToDse}
                  onChange={(val) => setSyncToDse(val)}
                />
              </div>
            </div>
          )}

          {/* Row 5: Narrations with Autocomplete Suggestions */}
          <div ref={narrationContainerRef} className="relative space-y-1.5">
            <div className="flex items-center justify-between">
              <label className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest flex items-center gap-1.5">
                <span>Narrations / Remarks</span>
                {txForm.description.trim() && narrationSuggestions.length > 0 && (
                  <span className="text-[9px] text-teal-400 font-bold lowercase tracking-normal flex items-center gap-1 bg-teal-400/10 px-1.5 py-0.5 rounded">
                    <Sparkles size={10} /> {narrationSuggestions.length} matching {narrationSuggestions.length === 1 ? 'entry' : 'entries'}
                  </span>
                )}
              </label>
              {txForm.description && (
                <button
                  type="button"
                  onClick={() => {
                    setTxForm(prev => ({ ...prev, description: '' }));
                    setIsNarrationFocused(false);
                  }}
                  className="text-[10px] text-slate-500 hover:text-slate-300 transition-colors uppercase tracking-wider font-bold"
                >
                  Clear
                </button>
              )}
            </div>

            <div className="relative">
              <input
                type="text"
                placeholder="Agora Shopping, Payout received, taxi, etc."
                value={txForm.description}
                onChange={(e) => {
                  setTxForm(prev => ({ ...prev, description: e.target.value }));
                  setIsNarrationFocused(true);
                  setHighlightedSuggestionIdx(-1);
                }}
                onFocus={() => {
                  setIsNarrationFocused(true);
                }}
                onKeyDown={handleNarrationKeyDown}
                className={cn(
                  "w-full bg-slate-950 border border-slate-800 rounded-xl pl-4 pr-10 py-2.5 text-sm font-medium text-white placeholder:text-slate-600 outline-none transition-all",
                  "focus:border-teal-500/50 focus:ring-1 focus:ring-teal-500/20",
                  isNarrationFocused && narrationSuggestions.length > 0 && "border-slate-700 shadow-lg"
                )}
              />
              <div className="absolute right-3 top-1/2 -translate-y-1/2 flex items-center gap-1 text-slate-600 pointer-events-none">
                {txForm.description.trim() ? (
                  <Search size={14} className="text-slate-500" />
                ) : (
                  <History size={14} className="text-slate-600" />
                )}
              </div>
            </div>

            {/* Suggestions Dropdown */}
            {isNarrationFocused && narrationSuggestions.length > 0 && (
              <div className="absolute left-0 right-0 top-full mt-1.5 z-50 bg-slate-900 border border-slate-700/80 rounded-xl shadow-2xl overflow-hidden backdrop-blur-md">
                <div className="px-3 py-2 bg-slate-950/90 border-b border-slate-800 flex items-center justify-between text-[10px] text-slate-400 uppercase tracking-wider font-extrabold">
                  <span className="flex items-center gap-1.5 text-teal-400 font-bold">
                    <Sparkles size={12} />
                    {txForm.description.trim() ? 'Matching Sub-Categories & Entries' : 'Recent Entries & Sub-Categories'}
                  </span>
                  <span className="text-slate-500 normal-case font-medium text-[10px]">
                    Select to auto-fill Type, Category & Account
                  </span>
                </div>

                <div className="max-h-64 overflow-y-auto divide-y divide-slate-800/60">
                  {narrationSuggestions.map((sug, idx) => {
                    const isHighlighted = idx === highlightedSuggestionIdx;
                    const isSubCategorySource = sug.source === 'subcategory';

                    return (
                      <button
                        key={sug.key}
                        type="button"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          handleSelectNarrationSuggestion(sug);
                        }}
                        onMouseEnter={() => setHighlightedSuggestionIdx(idx)}
                        className={cn(
                          "w-full text-left p-3 transition-colors flex flex-col gap-1.5 group cursor-pointer",
                          isHighlighted ? "bg-slate-800/90 text-white" : "hover:bg-slate-800/50 text-slate-300"
                        )}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-2 min-w-0">
                            {isSubCategorySource ? (
                              <span className="shrink-0 px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider bg-teal-500/20 text-teal-300 border border-teal-500/30 flex items-center gap-1">
                                <Tag size={8} /> Sub-Category
                              </span>
                            ) : (
                              <span className="shrink-0 px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider bg-slate-800 text-slate-400 border border-slate-700 flex items-center gap-1">
                                <History size={8} /> Previous Entry
                              </span>
                            )}
                            <span className="font-bold text-xs text-white group-hover:text-teal-300 transition-colors truncate">
                              {highlightMatch(sug.description, txForm.description.trim())}
                            </span>
                          </div>

                          <span className="text-[9px] text-slate-500 font-semibold shrink-0">
                            {sug.frequency > 1 ? `${sug.frequency}x used` : sug.lastDate ? formatDate(sug.lastDate) : isSubCategorySource ? 'Auto-map' : ''}
                          </span>
                        </div>

                        {/* Metadata badges for Type, Category, and Account */}
                        <div className="flex flex-wrap items-center gap-1.5 text-[10px]">
                          {/* Transaction Type */}
                          <span className={cn(
                            "px-1.5 py-0.5 rounded text-[8px] font-extrabold uppercase border",
                            sug.type === 'Income' ? "bg-emerald-500/15 text-emerald-400 border-emerald-500/30" :
                            sug.type === 'Expense' ? "bg-rose-500/15 text-rose-400 border-rose-500/30" :
                            sug.type === 'Loan' ? "bg-amber-500/15 text-amber-400 border-amber-500/30" :
                            "bg-sky-500/15 text-sky-400 border-sky-500/30"
                          )}>
                            {sug.type}
                          </span>

                          {/* Category Badge */}
                          <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-slate-950 border border-slate-800 text-slate-300 text-[9px] font-medium">
                            <Tag size={9} className="text-slate-500 shrink-0" />
                            <span>{sug.category || 'No Category'}</span>
                            {sug.subCategory && (
                              <span className="text-teal-400 font-semibold">› {sug.subCategory}</span>
                            )}
                          </span>

                          {/* Account Badge */}
                          <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-slate-950 border border-slate-800 text-slate-300 text-[9px] font-medium">
                            <CreditCard size={9} className="text-slate-500 shrink-0" />
                            <span>{sug.accountName}</span>
                            <span className="text-slate-500 font-bold uppercase text-[8px]">({sug.accountCurrency})</span>
                          </span>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Auto-fill notification feedback banner */}
            {autoFillNotice && (
              <div className="flex items-center justify-between gap-2 px-3 py-2 bg-teal-950/40 border border-teal-500/30 rounded-xl text-teal-300 text-[11px] font-medium animate-in fade-in slide-in-from-top-1 duration-200">
                <div className="flex items-center gap-1.5 truncate">
                  <Sparkles size={13} className="text-teal-400 shrink-0" />
                  <span className="truncate">
                    Auto-filled: <strong className="text-white">{autoFillNotice.type}</strong> • <strong className="text-white">{autoFillNotice.category}{autoFillNotice.subCategory ? ` (${autoFillNotice.subCategory})` : ''}</strong> • <strong className="text-white">{autoFillNotice.accountName}</strong>
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => setAutoFillNotice(null)}
                  className="text-slate-400 hover:text-white shrink-0 p-0.5"
                >
                  <X size={12} />
                </button>
              </div>
            )}
          </div>

          {/* Actions */}
          <div className="pt-2 flex gap-3">
            <Button type="button" variant="secondary" className="flex-1" onClick={() => { setIsTxModalOpen(false); setEditingTx(null); isFromExternalModuleRef.current = false; }}>
              Cancel
            </Button>
            <Button type="submit" className="flex-1">
              {editingTx ? "Update Transaction" : "Record Transaction"}
            </Button>
          </div>
        </form>
      </Modal>

      {/* ─── CATEGORY PICKER MODAL ────────────────────────────────────────── */}
      <CategoryPickerModal
        isOpen={isCategoryPickerOpen}
        onClose={() => setIsCategoryPickerOpen(false)}
        categories={categories}
        txType={txForm.type}
        selectedCategory={txForm.category}
        selectedSubCategory={txForm.subCategory}
        onSelect={(category, subCategory) => {
          setTxForm(prev => ({ ...prev, category, subCategory }));
        }}
      />

      {/* ─── ACCOUNT PICKER MODAL (From / Primary) ─────────────────────────── */}
      <AccountPickerModal
        isOpen={isAccountPickerOpen}
        onClose={() => setIsAccountPickerOpen(false)}
        accounts={accounts}
        accountBalances={accountBalances}
        selectedAccountId={txForm.accountId}
        onSelect={(accountId) => {
          setTxForm(prev => ({ ...prev, accountId }));
        }}
        formatCurrency={formatCurrency}
      />

      {/* ─── ACCOUNT PICKER MODAL (To / Secondary) ───────────────────────────── */}
      <AccountPickerModal
        isOpen={isToAccountPickerOpen}
        onClose={() => setIsToAccountPickerOpen(false)}
        accounts={accounts}
        accountBalances={accountBalances}
        selectedAccountId={txForm.toAccountId}
        onSelect={(toAccountId) => {
          const targetAcc = accounts.find(a => a.id === toAccountId);
          if (isBoAccount(targetAcc)) {
            setSyncToDse(true);
          }
          setTxForm(prev => ({ ...prev, toAccountId }));
        }}
        formatCurrency={formatCurrency}
      />

      {/* ─── MODAL ADD/EDIT PROFILE ACCOUNT ─────────────────────────────────── */}
      <Modal isOpen={isAccModalOpen} onClose={() => { setIsAccModalOpen(false); setEditingAcc(null); }} title={editingAcc ? "Edit Account Profile" : "Add Account Profile"}>
        <form onSubmit={handleSaveAcc} className="space-y-4">
          {/* Row 1: Account Name and Currency */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <Input label="Account Name" type="text" required placeholder="e.g. City Bank Savings" value={accForm.name} onChange={(e) => setAccForm(prev => ({ ...prev, name: e.target.value }))} />
            <Select label="Currency" value={accForm.currency} onChange={(v) => setAccForm(prev => ({ ...prev, currency: v as any }))} options={[{ label: 'BDT (৳)', value: 'BDT' }, { label: 'LYD (LD)', value: 'LYD' }, { label: 'USD ($)', value: 'USD' }]} />
          </div>

          {/* Row 2: Initial Date and Initial Balance */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <Input label="Initial Date" type="date" required value={accForm.initialDate} onChange={(e) => setAccForm(prev => ({ ...prev, initialDate: e.target.value }))} />
            <Input label="Initial Balance" type="number" step="any" placeholder="e.g. 15000" value={accForm.initialBalance} onChange={(e) => setAccForm(prev => ({ ...prev, initialBalance: e.target.value }))} />
          </div>

          <div className="flex items-center gap-2 py-1">
            <Checkbox id="acc-is-parent-checkbox" label="Is Group Account (Contains sub-accounts)" checked={accForm.isParent} onChange={(checked) => setAccForm(prev => ({ ...prev, isParent: checked, parentId: checked ? '' : prev.parentId }))} />
          </div>
          {!accForm.isParent && (
            <Select label="Assign to Parent Group (Optional)" value={accForm.parentId} onChange={(v) => setAccForm(prev => ({ ...prev, parentId: v }))} options={[{ label: 'None (Standalone)', value: '' }, ...accounts.filter(a => a.isParent && a.id !== editingAcc?.id).map(a => ({ label: a.name, value: a.id }))]} />
          )}
          <div className="pt-4 flex gap-3">
            <Button type="button" variant="secondary" className="flex-1" onClick={() => { setIsAccModalOpen(false); setEditingAcc(null); }}>Cancel</Button>
            <Button type="submit" className="flex-1">Save Profile</Button>
          </div>
        </form>
      </Modal>

      {/* Custom Category Modal */}
      <Modal isOpen={isCatModalOpen} onClose={() => { setIsCatModalOpen(false); setEditingCat(null); }} title={editingCat ? 'Edit Category' : 'Add Category'}>
        <form onSubmit={handleSaveCat} className="space-y-4">
          <Input label="Category Name" placeholder="e.g. Remittance Income" value={catForm.name} onChange={(e) => setCatForm(prev => ({ ...prev, name: e.target.value }))} required />
          <Select label="Type" value={catForm.type} onChange={(v) => setCatForm(prev => ({ ...prev, type: v as 'Income' | 'Expense' }))} options={[{ label: 'Income', value: 'Income' }, { label: 'Expense', value: 'Expense' }]} />
          <Input label="Sub-categories (comma-separated)" placeholder="e.g. Salary, Incentive, Facebook" value={catForm.subCategories} onChange={(e) => setCatForm(prev => ({ ...prev, subCategories: e.target.value }))} helperText="Write sub-categories separated by commas. Leave empty for no subcategories." />
          <div className="flex gap-3 justify-end pt-5 border-t border-slate-800">
            <Button type="button" variant="secondary" onClick={() => { setIsCatModalOpen(false); setEditingCat(null); }} className="flex-1">Cancel</Button>
            <Button type="submit" className="flex-1">Save Category</Button>
          </div>
        </form>
      </Modal>

      {/* Custom Non-blocking Alert Modal */}
      {customAlert && (
        <Modal isOpen={!!customAlert} onClose={() => setCustomAlert(null)} title={customAlert.title}>
          <div className="space-y-5 py-2">
            <p className="text-body font-semibold text-slate-300 leading-relaxed">{customAlert.message}</p>
            <div className="pt-3 flex justify-end">
              <Button type="button" onClick={() => setCustomAlert(null)} className="px-6">OK</Button>
            </div>
          </div>
        </Modal>
      )}

      {/* Custom Non-blocking Confirm Dialog */}
      {customConfirm && (
        <ConfirmDialog
          isOpen={!!customConfirm}
          title={customConfirm.title}
          message={customConfirm.message}
          variant={customConfirm.variant}
          confirmLabel={customConfirm.confirmLabel}
          cancelLabel={customConfirm.cancelLabel}
          details={customConfirm.details}
          onConfirm={() => {
            customConfirm.onConfirm();
            setCustomConfirm(null);
          }}
          onCancel={() => setCustomConfirm(null)}
        />
      )}

    </div>
  );
}