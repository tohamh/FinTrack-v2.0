/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { format } from 'date-fns';

export const toDateStr = (d: Date): string => {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
};

export const getTodayStr = (): string => toDateStr(new Date());

export const getFirstOfMonth = (date: Date): string => {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-01`;
};

export const getLastOfMonth = (date: Date): string => {
  const last = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  return toDateStr(last);
};

/**
 * Formats a number as BDT currency: ##,##,###.##
 */
export const formatBDT = (amount: number): string => {
  const formatter = new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'BDT',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return formatter.format(amount).replace('BDT', '৳');
};

export const formatCurrency = (amount: number, currency: string = 'BDT'): string => {
  if (currency === 'USD') {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(amount);
  }
  return formatBDT(amount);
};

/**
 * Formats a number with commas: ##,###.##
 */
export const formatNumber = (num: number, decimals: number = 2): string => {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  }).format(num);
};

export const formatDate = (date: string | Date): string => {
  if (!date) return 'N/A';
  const d = new Date(date);
  if (isNaN(d.getTime())) return 'Invalid Date';
  return format(d, 'dd MMM yyyy');
};

export const cn = (...classes: (string | undefined | boolean | null)[]) => {
  return classes.filter(Boolean).join(' ');
};
