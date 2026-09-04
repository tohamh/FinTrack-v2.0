/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type Currency = 'BDT' | 'USD';

export interface BaseInvestment {
  id: string;
  name: string;
  investmentDate: string;
  amount: number;
  currency: Currency;
  notes?: string;
}

export interface DSEHolding extends BaseInvestment {
  symbol: string;
  units: number;
  avgCost: number;
  currentPrice: number;
  realizedProfit: number;
  dividendReceived: number;
}

export interface MutualFundTransaction {
  id: string;
  date: string;
  type: 'Buy' | 'Sell' | 'Dividend' | 'Withdrawal';
  units: number;
  nav: number;
  amount: number;
  isDividend?: boolean;
  isWithdrawal?: boolean;
  pullingDate?: string;
  sipAmount?: number;
  linkedIeGroupId?: string;
}

export interface MutualFund extends BaseInvestment {
  amc: string;
  fullName: string;
  startingMonth: string;
  pullingDate: number;
  sipAmount: number;
  transactions: MutualFundTransaction[];
}

export type OnlineInvestmentStatus = 'Active' | 'Completed' | 'Delayed' | 'Matured';

export type OnlineInstallmentType = 'Repayment' | 'Interest';

export interface OnlineInvestment extends BaseInvestment {
  platform: string;
  companyName: string;
  projectName: string;
  estimatedReturn: number;
  estimatedProfit: number;
  expectedROE: number; // % per annum
  durationMonths: number;
  maturityDate: string;
  actualMaturityDate?: string;
  actualProfit?: number;
  actualROE?: number;
  closingDate?: string;
  withdrawBalance?: number;
  status: OnlineInvestmentStatus;
  hasRepaymentSchedule: boolean;
  isDefaultSchedule: boolean;
  linkedIeGroupId?: string;
  installments: { 
    date: string; 
    amount: number; 
    isPaid: boolean; 
    actualDate?: string; 
    actualAmount?: number;
    isAutoMarked?: boolean;
    isManuallyEdited?: boolean;
    linkedIncomeTxId?: string;
    type?: OnlineInstallmentType;
  }[];
  totalRepaid: number;
}

export type InvestmentFrequency = 'Quarterly' | 'Semi-annual' | 'Annual';

export interface Sukuk extends BaseInvestment {
  instrumentNo: string;
  issuer: string;
  rentRate: number; // %
  tds: number; // %
  frequency: InvestmentFrequency;
  principalAmount: number;
  issueDate: string;
  durationYears: number;
  totalRepaid: number;
  status: OnlineInvestmentStatus;
  closingDate?: string;
  withdrawBalance?: number;
  linkedIeGroupId?: string;
  installments: { 
    date: string; 
    amount: number; 
    isPaid: boolean; 
    actualDate?: string; 
    actualAmount?: number;
    isAutoMarked?: boolean;
    isManuallyEdited?: boolean;
    installmentNo: number;
    linkedIncomeTxId?: string;
  }[];
}

export interface FDRTransaction {
  id: string;
  date: string;
  type: 'Profit' | 'Charge';
  amount: number;
  handling?: 'Added' | 'Withdrawn';
  linkedIncomeTxId?: string;
}

export interface FDR extends BaseInvestment {
  bankName: string;
  accountNo: string;
  principal: number;
  interestFrequency: 'Monthly' | 'Yearly';
  interestHandling: 'Added' | 'Withdrawn';
  status: 'Active' | 'Matured' | 'Closed';
  exchangeRate?: number;
  closingDate?: string;
  withdrawBalance?: number;
  transactions: FDRTransaction[];
}

export type TransactionType = 'Income' | 'Expense';

export interface Transaction {
  id: string;
  date: string;
  amount: number;
  category: string;
  type: TransactionType;
  description: string;
}

export interface BankAccount {
  id: string;
  name: string;
  currency: 'BDT' | 'LYD' | 'USD';
  initialBalance: number;
  initialDate?: string;
  isParent?: boolean;
  parentId?: string;
}

export interface IncomeExpenseTransaction {
  id: string;
  date: string;
  amount: number;
  type: 'Income' | 'Expense' | 'Transfer' | 'Loan';
  accountId: string;
  toAccountId?: string;
  toAmount?: number;
  transferGroupId?: string;
  transferType?: 'from' | 'to';
  category: string;
  subCategory?: string;
  description: string;
  dseTxId?: string;
  autoSyncDse?: boolean;
  mfTxId?: string;
  sukukTxId?: string;
  sukukInstallmentNo?: number;
  fdrTxId?: string;
  onlineTxId?: string;
  onlineInstallmentNo?: number;
}

export interface ConversionRates {
  USD_to_BDT: number;
  LYD_to_BDT: number;
}

export interface AppCategory {
  id: string;
  name: string;
  type: 'Income' | 'Expense';
  subCategories: string[];
}

export interface AppState {
  dseHoldings: DSEHolding[];
  mutualFunds: MutualFund[];
  onlineInvestments: OnlineInvestment[];
  sukuks: Sukuk[];
  fdrs: FDR[];
  transactions: Transaction[];
  cashBalance: number;
  pin: string | null;
  isLocked: boolean;
  incomeExpenseAccounts?: BankAccount[];
  incomeExpenseTransactions?: IncomeExpenseTransaction[];
  conversionRates?: ConversionRates;
  incomeExpenseCategories?: AppCategory[];
}
