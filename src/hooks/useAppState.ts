/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect } from 'react';
import { AppState } from '../types';
import { markDirty, pushModuleData, serializeIncomeExpense } from '../utils/sheetSync';
import type { ModuleKey } from '../utils/sheetSync';

const STORAGE_KEY = 'fintrack_pro_data';

// Map from AppState key → ModuleKey for dirty-marking
const STATE_TO_MODULE: Partial<Record<keyof AppState, ModuleKey>> = {
  onlineInvestments:         'onlineInvestments',
  sukuks:                    'sukuk',
  mutualFunds:               'mutualFunds',
  fdrs:                      'fixedDeposits',
  incomeExpenseTransactions: 'incomeExpense',
  incomeExpenseAccounts:     'incomeExpense',
  incomeExpenseCategories:   'incomeExpense',
  conversionRates:           'incomeExpense',
};

const initialData: AppState = {
  dseHoldings: [],
  mutualFunds: [],
  onlineInvestments: [],
  sukuks: [],
  fdrs: [],
  transactions: [],
  cashBalance: 0,
  pin: null,
  isLocked: true,
  incomeExpenseAccounts: [
    // Standalones
    { id: 'bdt-city-islamic', name: 'City Islamic', currency: 'BDT', initialBalance: 670568.95, initialDate: '2026-05-01' },
    { id: 'bdt-ibbl-savings', name: 'IBBL Savings', currency: 'BDT', initialBalance: 119405.89, initialDate: '2026-05-01' },
    { id: 'bdt-antora', name: 'Antora BDT', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01' },
    { id: 'bdt-toha', name: 'Toha BDT', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01' },
    { id: 'bdt-ibbl-fdr', name: 'IBBL FDR BDT', currency: 'BDT', initialBalance: 3000000, initialDate: '2026-05-01' },
    
    // Parents
    { id: 'bdt-investments', name: 'Investments', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', isParent: true },
    { id: 'bdt-wallets', name: 'Wallets', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', isParent: true },

    // Children under Investments
    { id: 'bdt-bo-account', name: 'BO Account', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', parentId: 'bdt-investments' },
    { id: 'bdt-online-investment', name: 'Online Investment', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', parentId: 'bdt-investments' },
    { id: 'bdt-mutual-fund', name: 'Mutual Fund', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', parentId: 'bdt-investments' },
    { id: 'bdt-sukuk', name: 'Sukuk', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', parentId: 'bdt-investments' },

    // Children under Wallets
    { id: 'bdt-ebl-acc', name: 'EBL Savings', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', parentId: 'bdt-wallets' },
    { id: 'bdt-ebl-visa', name: 'EBL VISA BDT', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', parentId: 'bdt-wallets' },
    { id: 'bdt-bkash', name: 'bKash', currency: 'BDT', initialBalance: 0, initialDate: '2026-05-01', parentId: 'bdt-wallets' },
    { id: 'bdt-cellfin-mcash', name: 'Cellfin/MCash', currency: 'BDT', initialBalance: 3194, initialDate: '2026-05-01', parentId: 'bdt-wallets' },

    // LYD
    { id: 'lyd-ammu', name: 'Ammu LYD', currency: 'LYD', initialBalance: 2500, initialDate: '2026-05-01' },
    { id: 'lyd-toha', name: 'Toha LYD', currency: 'LYD', initialBalance: 1200, initialDate: '2026-05-01' },

    // USD
    { id: 'usd-cash', name: 'Cash USD', currency: 'USD', initialBalance: 0, initialDate: '2026-05-01' },
    { id: 'usd-bank-group', name: 'Bank USD', currency: 'USD', initialBalance: 0, initialDate: '2026-05-01', isParent: true },
    { id: 'usd-city-fcy', name: 'City FCY', currency: 'USD', initialBalance: 73205.94, initialDate: '2026-05-01', parentId: 'usd-bank-group' },
    { id: 'usd-ibbl-fcy', name: 'IBBL FCY', currency: 'USD', initialBalance: 0, initialDate: '2026-05-01', parentId: 'usd-bank-group' },
    { id: 'usd-ibbl-fdr', name: 'IBBL FDR USD', currency: 'USD', initialBalance: 0, initialDate: '2026-05-01', parentId: 'usd-bank-group' },
    { id: 'usd-ebl-usd', name: 'EBL VISA USD', currency: 'USD', initialBalance: 0, initialDate: '2026-05-01' }
  ],
  incomeExpenseTransactions: [
    { id: 'tx-1', date: '2026-05-01', amount: 5000, type: 'Income', accountId: 'bdt-city-islamic', category: 'Remittance Income', subCategory: 'Salary', description: 'Freelance work payout' },
    { id: 'tx-2', date: '2026-05-04', amount: 500, type: 'Expense', accountId: 'bdt-bkash', category: 'Food & Clothing', description: 'Unimart grocery shopping' },
    { id: 'tx-3', date: '2026-05-08', amount: 50, type: 'Expense', accountId: 'lyd-toha', category: 'Transportation', description: 'Taxi fare' },
    { id: 'tx-4', date: '2026-05-12', amount: 100, type: 'Income', accountId: 'usd-cash', category: 'Other Income', subCategory: 'Other Income', description: 'Consulting project bonus' },
    { id: 'tx-5', date: '2026-05-15', amount: 1500, type: 'Expense', accountId: 'bdt-city-islamic', category: 'Household & utility', description: 'Electricity and Internet bill' }
  ],
  conversionRates: {
    USD_to_BDT: 120,
    LYD_to_BDT: 20
  },
  incomeExpenseCategories: [
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
      id: 'cat-loan-income',
      name: 'Loan',
      type: 'Income',
      subCategories: ['Borrowed', 'Received Lent Money']
    },
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
      subCategories: ['TDS on Provisional Profit', 'Account Maintenance Fees', 'Other Bank Charges']
    },
    {
      id: 'cat-loan-expense',
      name: 'Loan',
      type: 'Expense',
      subCategories: ['Lent', 'Repaid Borrowed Money']
    }
  ]
};

export function useAppState() {
  const [state, setState] = useState<AppState>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        return { 
          ...initialData, 
          ...parsed,
          incomeExpenseAccounts: Array.isArray(parsed.incomeExpenseAccounts) && parsed.incomeExpenseAccounts.length > 0
            ? parsed.incomeExpenseAccounts
            : initialData.incomeExpenseAccounts,
          incomeExpenseTransactions: Array.isArray(parsed.incomeExpenseTransactions)
            ? parsed.incomeExpenseTransactions
            : initialData.incomeExpenseTransactions,
          incomeExpenseCategories: (() => {
            const rawCats = Array.isArray(parsed.incomeExpenseCategories) && parsed.incomeExpenseCategories.length > 0
              ? parsed.incomeExpenseCategories
              : initialData.incomeExpenseCategories;
            return rawCats.map((c: any) => {
              if (c.name === 'Other Income') {
                const subCats = c.subCategories || [];
                if (!subCats.includes('Investment Interest')) {
                  const updatedSubs = [...subCats];
                  const otherIdx = updatedSubs.indexOf('Other Income');
                  if (otherIdx !== -1) {
                    updatedSubs.splice(otherIdx, 0, 'Investment Interest');
                  } else {
                    updatedSubs.push('Investment Interest');
                  }
                  return { ...c, subCategories: updatedSubs };
                }
              }
              if (c.name === 'Finance Income') {
                return {
                  ...c,
                  subCategories: (c.subCategories || []).filter((s: string) => s !== 'Investment Interest')
                };
              }
              return c;
            });
          })(),
          conversionRates: parsed.conversionRates || initialData.conversionRates,
          isLocked: true 
        };
      }
    } catch (e) {
      console.error('Failed to load state from localStorage', e);
    }
    return initialData;
  });

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      console.error('Failed to save state to localStorage', e);
      if (e instanceof Error && e.name === 'QuotaExceededError') {
        alert('Storage quota exceeded. Please delete some data or backups.');
      }
    }
  }, [state]);

  const updateState = (updater: (prev: AppState) => AppState, dirtyModules?: ModuleKey[]) => {
    setState(prev => {
      const newState = updater(prev);

      // Auto-detect which module arrays changed and mark them dirty
      const detectedModules = (Object.entries(STATE_TO_MODULE) as [keyof AppState, ModuleKey][])
        .filter(([stateKey]) => newState[stateKey] !== prev[stateKey])
        .map(([, moduleKey]) => moduleKey);

      const modulesToMark = dirtyModules ?? Array.from(new Set(detectedModules));

      // Execute side-effects asynchronously outside the React state calculation
      setTimeout(() => {
        modulesToMark.forEach(m => {
          markDirty(m);
          // Instant sync with Google Sheets without any delay
          if (m === 'fixedDeposits') {
            pushModuleData('fixedDeposits', newState.fdrs);
          } else if (m === 'mutualFunds') {
            pushModuleData('mutualFunds', newState.mutualFunds);
          } else if (m === 'sukuk') {
            pushModuleData('sukuk', newState.sukuks);
          } else if (m === 'onlineInvestments') {
            pushModuleData('onlineInvestments', newState.onlineInvestments);
          } else if (m === 'incomeExpense') {
            pushModuleData('incomeExpense', serializeIncomeExpense(newState));
          }
        });
      }, 0);

      return newState;
    });
  };

  return { state, updateState };
}