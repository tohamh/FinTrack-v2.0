/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { 
  fetchFromSheet, 
  pushToSheet, 
  getCachedData,
  pushModuleData,
  markDirty
} from '../../utils/sheetSync';

import { ConfirmDialog } from '../ui/ConfirmDialog';
import * as XLSX from 'xlsx';
import React, { useState, useMemo, useEffect, useCallback } from 'react';
import { 
  PieChart as RechartsPieChart, Pie, Cell, ResponsiveContainer, Legend,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ReferenceLine, LabelList,
  ComposedChart, Line
} from 'recharts';
import { Card, Checkbox, Button, Modal, Input, Select } from '../ui/BaseComponents';
import { DSEHolding } from '../../types';
import { formatBDT, cn } from '../../utils/formatters';
import {
  Settings, Plus, DollarSign, Briefcase,
  Calendar, ChevronDown, RefreshCw, Wallet, TrendingUp, Info, ChevronLeft, ChevronRight,
  Search, Filter, ArrowUpDown, ArrowUp, ArrowDown, Trash2, Edit2, AlertTriangle,
  Download, Upload, FileSpreadsheet, BarChart2, PieChart, Layers, SlidersHorizontal, Percent
} from 'lucide-react';



interface DseTransaction {
  id: string;
  date: string;
  type: 'Buy' | 'Sell' | 'Deposit' | 'Withdrawal' | 'Dividend' | 'Charge';
  portfolio: 'Global' | 'Investment' | 'Trading';
  ticker: string;
  companyName: string;
  qty: number;
  price: number;
  commission: number;
  total: number;
  notes: string;
}

interface DseTrackerModuleProps {
  holdings: DSEHolding[];
  onAdd: (holding: Omit<DSEHolding, 'id'>) => void;
  onUpdate: (id: string, updates: Partial<DSEHolding>) => void;
  onDelete: (id: string) => void;
  onTitleChange?: (title: React.ReactNode) => void;
  triggerAdd?: boolean;
  setTriggerAdd?: (val: boolean) => void;
}

// ── Tab Types ──────────────────────────────────────────────────────────────

type DseTab = 'summary' | 'transactions' | 'holdings' | 'analytics' | 'settings';

// ── Color maps ─────────────────────────────────────────────────────────────

const TYPE_BG: Record<string, string> = {
  Buy:        'bg-emerald-400/15 border-emerald-400/25 text-emerald-400',
  Sell:       'bg-rose-400/15    border-rose-400/25    text-rose-400',
  Deposit:    'bg-teal-400/15    border-teal-400/25    text-teal-400',
  Withdrawal: 'bg-orange-400/15  border-orange-400/25  text-orange-400',
  Dividend:   'bg-blue-400/15    border-blue-400/25    text-blue-400',
  Charge:     'bg-slate-400/15   border-slate-400/25   text-slate-400',
};

const TYPE_TEXT: Record<string, string> = {
  Buy:        'text-white',
  Sell:       'text-rose-400',
  Deposit:    'text-teal-400',
  Withdrawal: 'text-orange-400',
  Dividend:   'text-blue-400',
  Charge:     'text-slate-400',
};

const PORTFOLIO_BG: Record<string, string> = {
  Global:     'bg-purple-400/15  border-purple-400/25  text-white',
  Investment: 'bg-blue-400/15    border-blue-400/25    text-white',
  Trading:    'bg-amber-400/15   border-amber-400/25   text-white',
};

const PORTFOLIO_BTN: Record<string, string> = {
  Global:     'bg-purple-500/20 border border-purple-400/40 text-purple-300 hover:bg-purple-500/30',
  Investment: 'bg-blue-500/20   border border-blue-400/40   text-blue-300   hover:bg-blue-500/30',
  Trading:    'bg-amber-500/20  border border-amber-400/40  text-amber-300  hover:bg-amber-500/30',
};

const PORTFOLIO_DOT: Record<string, string> = {
  Global: 'bg-purple-400', Investment: 'bg-blue-400', Trading: 'bg-amber-400',
};

// ── Helpers ────────────────────────────────────────────────────────────────

const toDateStr = (d: Date): string => {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
};
const getTodayStr = () => toDateStr(new Date());
const getFirstOfMonth = (date: Date) =>
  `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-01`;
const getLastOfMonth = (date: Date) =>
  toDateStr(new Date(date.getFullYear(), date.getMonth() + 1, 0));

const formatDDMMYYYY = (dateStr: string): string => {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-');
  return `${d}/${m}/${y}`;
};

const typeLabel = (t: string) => t === 'Withdrawal' ? 'Withdraw' : t;
const portfolioLabel = (p: string) => p === 'Investment' ? 'Invest' : p;

// ── Default blank form ─────────────────────────────────────────────────────

const blankForm = (date: string): Partial<DseTransaction> => ({
  id: Math.random().toString(36).substr(2, 9),
  date,
  type: 'Buy',
  portfolio: 'Investment',
  ticker: '',
  companyName: '',
  qty: 0,
  price: 0,
  commission: 0,
  total: 0,
  notes: '',
});

// ── InfoTooltip ────────────────────────────────────────────────────────────

const InfoTooltip: React.FC<{ text: string }> = ({ text }) => {
  const [visible, setVisible] = useState(false);
  const [pos, setPos] = useState({ top: 0, left: 0 });
  const btnRef = React.useRef<HTMLButtonElement>(null);
  const show = () => {
    if (btnRef.current) {
      const rect = btnRef.current.getBoundingClientRect();
      setPos({ top: rect.top + window.scrollY - 8, left: rect.left + rect.width / 2 });
    }
    setVisible(true);
  };
  return (
    <div className="relative inline-flex items-center">
      <button ref={btnRef} onMouseEnter={show} onMouseLeave={() => setVisible(false)}
        onFocus={show} onBlur={() => setVisible(false)}
        className="p-0.5 text-slate-600 hover:text-slate-400 transition-colors rounded" aria-label="More info" type="button">
        <Info size={11} />
      </button>
      {visible && (
        <div style={{ position: 'fixed', top: pos.top, left: pos.left, transform: 'translate(-50%, -100%)', zIndex: 9999 }}
          className="w-48 bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-[10px] font-bold text-slate-200 shadow-xl pointer-events-none whitespace-normal text-left leading-relaxed animate-in fade-in zoom-in-95 duration-150">
          {text}
          <div className="absolute left-1/2 -translate-x-1/2 top-full w-0 h-0 border-l-4 border-r-4 border-t-4 border-l-transparent border-r-transparent border-t-slate-700" />
        </div>
      )}
    </div>
  );
};

// ── Settings Button ────────────────────────────────────────────────────────

interface SettingsButtonProps { 
  isOpen: boolean; 
  setIsOpen: (v: boolean) => void; 
  onExport: () => void;
  onImport: (e: React.ChangeEvent<HTMLInputElement>) => void;
  onTemplate: () => void;
  onAdd: () => void;
  newTicker: string;
  setNewTicker: (v: string) => void;
  newCompany: string;
  setNewCompany: (v: string) => void;
  onAddStock: () => void;
}
const SettingsButton: React.FC<SettingsButtonProps> = ({ isOpen, setIsOpen, onExport, onImport, onTemplate, onAdd, newTicker, setNewTicker, newCompany, setNewCompany, onAddStock }) => (
  <div className="relative">
    <button onClick={() => setIsOpen(!isOpen)}
      className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/10 hover:bg-teal-300">
      <Settings size={14} />
      <span className="hidden sm:inline">Settings</span>
      <ChevronDown size={14} className={cn("opacity-50 transition-transform", isOpen ? "rotate-180" : "")} />
    </button>
    {isOpen && (
      <>
        <div className="fixed inset-0 z-40" onClick={() => setIsOpen(false)} />
        <div className="absolute right-0 mt-2 w-64 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-3 animate-in fade-in zoom-in-95">
          <button onClick={() => { onExport(); setIsOpen(false); }}
            className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase">
            <Download size={14} className="text-teal-400" />EXPORT
          </button>
          <label className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase cursor-pointer">
            <Upload size={14} className="text-teal-400" />IMPORT
            <input type="file" className="hidden" accept=".xlsx,.xls" onChange={onImport} />
          </label>
          <button onClick={() => { onTemplate(); setIsOpen(false); }}
            className="w-full flex items-center gap-2 px-3 py-1.5 text-label font-bold text-slate-300 hover:bg-slate-800 hover:text-white rounded-lg transition-colors uppercase">
            <FileSpreadsheet size={14} className="text-teal-400" />TEMPLATE
          </button>
          <div className="h-px bg-slate-800/60 my-2" />
          <div className="space-y-2 p-1">
            <input 
              value={newTicker} 
              onChange={(e) => setNewTicker(e.target.value)} 
              placeholder="Ticker"
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white"
            />
            <input 
              value={newCompany} 
              onChange={(e) => setNewCompany(e.target.value)} 
              placeholder="Company"
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white"
            />
            <button onClick={onAddStock} className="w-full bg-teal-400 text-slate-950 font-bold py-1.5 rounded-lg text-xs hover:bg-teal-300">
              Add Stock
            </button>
          </div>
        </div>
      </>
    )}
  </div>
);

// ── Transaction Modal ──────────────────────────────────────────────────────

interface TransactionModalProps {
  isOpen: boolean;
  editingTransaction: DseTransaction | null;
  onClose: () => void;
  onSave: (t: DseTransaction, keepOpen?: boolean) => Promise<void>;
}

interface DseSettings {
  commissionRate: number;
  customStocks: { ticker: string; name: string }[];
}

const TransactionModal: React.FC<TransactionModalProps & { 
  commissionRate: number; 
  customStocks: { ticker: string; name: string }[];
  existingStocks: { ticker: string; name: string }[];
}> = ({
  isOpen,
  editingTransaction,
  onClose,
  onSave,
  commissionRate,
  customStocks,
  existingStocks
}) => {
  const [formData, setFormData] = useState<Partial<DseTransaction>>(
    editingTransaction ?? blankForm(getTodayStr())
  );

  const [commMode, setCommMode] = useState<'Auto' | 'Manual'>('Auto');
  const [suggestions, setSuggestions] = useState<any[]>([]);
  const [error, setError] = useState<string>('');

  useEffect(() => {
    if (isOpen) {
      let init = editingTransaction ?? blankForm(getTodayStr());
      if (['Deposit', 'Withdrawal', 'Dividend', 'Charge'].includes(init.type || '')) {
        if (!init.qty || init.qty === 0) {
          init = {
            ...init,
            qty: 1,
            price: init.price || init.total || 0,
            total: init.total || init.price || 0,
          };
        }
      }
      setFormData(init);
      setCommMode('Auto');
      setSuggestions([]);
      setError('');
    }
  }, [isOpen, editingTransaction]);

  const set = (patch: Partial<DseTransaction>) =>
    setFormData(prev => ({ ...prev, ...patch }));

  // ── STOCK LIST (existing holdings + custom) ──
const STOCK_LIST = useMemo(() => {
  const map = new Map<string, { ticker: string; name: string }>();

  existingStocks.forEach(stock => {
    if (stock.ticker) {
      map.set(stock.ticker.toUpperCase(), {
        ticker: stock.ticker.toUpperCase(),
        name: stock.name || stock.ticker
      });
    }
  });

  customStocks.forEach(stock => {
    if (stock.ticker) {
      map.set(stock.ticker.toUpperCase(), {
        ticker: stock.ticker.toUpperCase(),
        name: stock.name || stock.ticker
      });
    }
  });

  return Array.from(map.values()).sort((a, b) =>
    a.ticker.localeCompare(b.ticker)
  );
}, [existingStocks, customStocks]);

  // ── SUGGESTIONS ──
  useEffect(() => {
    // If ticker matches exactly, or ticker input is disabled, hide suggestions
    if (STOCK_LIST.some(s => s.ticker === formData.ticker) || ['Deposit', 'Withdrawal', 'Charge'].includes(formData.type || '')) {
      setSuggestions([]);
      return;
    }

    if (!formData.ticker) {
      setSuggestions([]);
      return;
    }

    const q = formData.ticker.toLowerCase();

    setSuggestions(
      STOCK_LIST.filter(
        s =>
          s.ticker.toLowerCase().includes(q) ||
          s.name.toLowerCase().includes(q)
      )
    );
  }, [formData.ticker, STOCK_LIST]);

  // ── CALCULATION ──
  useEffect(() => {
    const qty = formData.qty || 0;
    const price = formData.price || 0;
    const subtotal = qty * price;

    setFormData(prev => {
      let commission = prev.commission || 0;

      if (commMode === 'Auto') {
        commission =
          (prev.type === 'Buy' || prev.type === 'Sell')
            ? subtotal * (commissionRate / 100)
            : 0;
      }

      let total = 0;
      if (prev.type === 'Buy') total = subtotal + commission;
      else if (prev.type === 'Sell') total = subtotal - commission;
      else total = subtotal;

      return { ...prev, commission, total };
    });
  }, [formData.qty, formData.price, formData.type, commMode, commissionRate]);

const handleNext = async () => {
    const snapshot = formData as DseTransaction;
    const keepDate = snapshot.date || getTodayStr();

    setFormData({
      ...blankForm(keepDate),
      date: keepDate,
      portfolio: snapshot.portfolio,
    });
    setSuggestions([]);
    setCommMode('Auto');
    setError('');

    onSave(snapshot, true);
  };

  const inputCls =
    "w-full h-10 bg-black border border-slate-700 rounded-lg px-3 text-[11px] font-bold text-white outline-none focus:border-teal-400/60";

  const labelCls =
    "text-[11px] font-bold text-white uppercase mb-1 block";

  if (!isOpen) return null;

  return (
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      title={editingTransaction ? "Edit Transaction" : "New Transaction"}
    >
      <div className="space-y-4">

        {/* Date + Type */}
        <div className="flex gap-3 items-end">
          <div className="flex-1">
            <label className={labelCls}>Date</label>
            <input
              type="date"
              className={inputCls}
              value={formData.date || getTodayStr()}
              onChange={e => set({ date: e.target.value })}
            />
          </div>

          <div className="flex-1">
            <label className={labelCls}>Type</label>
            <select
              className={inputCls}
              value={formData.type}
              onChange={e => {
                const newType = e.target.value as DseTransaction['type'];
                if (['Deposit', 'Withdrawal', 'Dividend', 'Charge'].includes(newType)) {
                  set({
                    type: newType,
                    qty: (formData.qty && formData.qty > 0) ? formData.qty : 1,
                    price: formData.price || formData.total || 0,
                  });
                } else {
                  set({ type: newType });
                }
              }}
            >
              {['Buy','Sell','Deposit','Withdrawal','Dividend','Charge'].map(t => (
                <option key={t}>{t}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Portfolio */}
        <div>
          <label className={labelCls}>Portfolio</label>
          <div className="flex gap-2">
            {[
              { id: 'Investment' },
              { id: 'Trading' }
            ].map(p => (
              <button
                key={p.id}
                type="button"
                onClick={() =>
                  set({ portfolio: p.id as DseTransaction['portfolio'] })
                }
                className={cn(
  "flex-1 px-4 py-2 rounded-lg text-label font-bold uppercase transition-colors",
  formData.portfolio === p.id
    ? 'bg-teal-400 text-slate-950 hover:bg-teal-300'
    : 'bg-slate-800 border border-slate-700 text-slate-300 hover:bg-slate-700'
)}
              >
                {p.id}
              </button>
            ))}
          </div>
        </div>

        {/* Ticker */}
        <div className="relative">
          <label className={labelCls}>Ticker / Stock</label> 
          <input placeholder="Ticker / Stock Name"
            disabled={['Deposit', 'Withdrawal', 'Charge'].includes(formData.type || '')}
            className={cn(inputCls, ['Deposit', 'Withdrawal', 'Charge'].includes(formData.type || '') && "opacity-50 cursor-not-allowed")}
            value={formData.ticker || ''}
            onChange={e => set({ ticker: e.target.value })}
          />

          {suggestions.length > 0 && (
            <div className="absolute z-50 w-full mt-1 bg-slate-900 border border-slate-700 rounded-lg overflow-hidden">
              {suggestions.map(s => (
                <div
                  key={s.ticker}
                  onClick={() => {
                    set({
                      ticker: s.ticker,
                      companyName: s.name
                    });
                    setSuggestions([]);
                  }}
                  className="px-3 py-2 hover:bg-slate-800 cursor-pointer"
                >
                  <div className="text-[12px] font-bold text-white">
                    {s.ticker}
                  </div>
                  <div className="text-[10px] text-slate-400">
                    {s.name}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Qty + Price */}
        <div className="flex gap-3">
          <div className="flex-1">
            <label className={labelCls}>Quantity</label>
            <input
              type="number"
              className={inputCls}
              value={formData.qty || ''}
              placeholder="0"
              onChange={e => {
                  const val = e.target.value === '' ? 0 : +e.target.value;
                  set({ qty: val });
                }
              }
            />
          </div>

          <div className="flex-1">
            <label className={labelCls}>Price (৳)</label>
            <input
              type="number"
              className={inputCls}
              value={formData.price || ''}
              placeholder="0.00"
              onChange={e => {
                  const val = e.target.value === '' ? 0 : +e.target.value;
                  set({ price: val });
                }
              }
            />
          </div>
        </div>

        {error && <p className="text-[11px] font-bold text-rose-500 uppercase">{error}</p>}

        {/* Comm + Total */}
        <div className="flex gap-3 items-end">
          <div className="flex-1">
            <div className="flex items-center justify-between h-[20px] mb-1">
              <span className="text-[11px] font-bold text-white uppercase leading-none">
                Comm (৳)
              </span>
              <button
                type="button"
                onClick={() => setCommMode(c => (c === 'Auto' ? 'Manual' : 'Auto'))}
                className="text-[10px] font-bold text-teal-400 leading-none hover:text-teal-300 transition-colors uppercase"
              >
                [{commMode}]
              </button>
            </div>

            <input
              type="number"
              value={formData.commission || 0}
              onChange={e => set({ commission: +e.target.value })}
              readOnly={commMode === 'Auto'}
              className={cn(
                "w-full h-10 rounded-lg px-3 text-[11px] font-bold outline-none border",
                commMode === 'Manual'
                  ? "bg-black border-slate-700 text-white"
                  : "bg-slate-800 border-slate-700 text-slate-400"
              )}
            />
          </div>

          <div className="flex-1">
            <div className="flex items-center h-[20px] mb-1">
              <span className="text-[11px] font-bold text-white uppercase leading-none">
                Total (৳)
              </span>
            </div>

            <input
              type="number"
              readOnly
              value={formData.total || 0}
              className="w-full h-10 bg-black border border-teal-400/30 rounded-lg px-3 text-[11px] font-bold text-teal-300"
            />
          </div>
        </div>

        {/* Buttons */}
        <div className="pt-2 flex gap-3">
          <button
            type="button"
            onClick={onClose}
            className="flex-1 px-4 py-2 rounded-lg bg-slate-800 border border-slate-700 text-slate-300 text-label font-bold uppercase hover:bg-slate-700 transition-colors"
          >
            Cancel
          </button>

          <button
            type="button"
            onClick={async () => {
              if ((formData.qty || 0) <= 0 || (formData.price || 0) <= 0) {
                 setError('Please enter a valid quantity and price');
                 return;
              }
              await onSave(formData as DseTransaction);
              onClose();
            }}
            className="flex-1 px-4 py-2 rounded-lg bg-teal-400 text-slate-950 text-label font-bold uppercase hover:bg-teal-300 transition-colors"
          >
            Save & Close
          </button>

          {!editingTransaction && (
            <button
              type="button"
              onClick={async () => {
                if ((formData.qty || 0) <= 0 || (formData.price || 0) <= 0) {
                   setError('Please enter a valid quantity and price');
                   return;
                }
                await handleNext();
              }}
              className="flex-1 px-4 py-2 rounded-lg bg-teal-400 text-slate-950 text-label font-bold uppercase hover:bg-teal-300 transition-colors"
            >
              Save & Add Another
            </button>
          )}
        </div>

      </div>
    </Modal>
  );
};


// ── Under Construction Placeholder ────────────────────────────────────────

interface UnderConstructionProps {
  icon: React.ReactNode;
  label: string;
}
const UnderConstruction: React.FC<UnderConstructionProps> = ({ icon, label }) => (
  <div className="flex flex-col items-center justify-center py-24 gap-5">
    <div className="w-16 h-16 rounded-2xl bg-slate-800 border border-slate-700 flex items-center justify-center text-slate-500">
      {icon}
    </div>
    <div className="text-center">
      <p className="text-body-sm font-bold text-white uppercase tracking-wider mb-1">{label}</p>
      <p className="text-label font-bold text-slate-500 uppercase">Under Construction</p>
    </div>
    <div className="flex items-center gap-2 px-4 py-2 rounded-xl bg-slate-900 border border-slate-800">
      <div className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse" />
      <span className="text-[10px] font-bold text-slate-400 uppercase">Coming Soon</span>
    </div>
  </div>
);

// ── Portfolio Allocation ──

// ══ START ══════════════════════════════════════════════════════════════════════
 
// ── Palette ───────────────────────────────────────────────────────────────────
const HOLDING_COLORS = [
  '#2dd4bf','#3b82f6','#a855f7','#f59e0b','#ef4444',
  '#ec4899','#06b6d4','#84cc16','#f97316','#8b5cf6',
  '#14b8a6','#eab308','#6366f1','#10b981','#fb923c',
  '#e879f9','#38bdf8','#a3e635','#fb7185','#34d399',
  '#60a5fa','#c084fc','#fbbf24','#f472b6','#4ade80',
];
 
// ── Types ─────────────────────────────────────────────────────────────────────
type HoldingsPortfolio = 'Investment' | 'Trading' | 'Global';
 
interface HoldingRow {
  ticker: string;
  companyName: string;
  portfolio: HoldingsPortfolio;
  totalBuyQty: number;
  totalBuyValue: number;
  totalSellQty: number;
  totalSellValue: number;
  currentQty: number;
  avgCost: number;
  currentPrice: number;
  realizedPnL: number;
  totalDividends: number;
  totalCharges: number;
  totalCostOfSoldShares: number;
}
 
// ── SVG donut helpers ─────────────────────────────────────────────────────────
// NOTE: angles start at top (12 o'clock) and go clockwise.
// We rotate -90° so that 0° corresponds to the top of the circle.
function toXY(cx: number, cy: number, r: number, angleDeg: number) {
  const rad = (angleDeg * Math.PI) / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}
 
function buildArcPath(
  cx: number, cy: number,
  outerR: number, innerR: number,
  startDeg: number, sweepDeg: number,
): string {
  // Clamp sweep
  const sw  = Math.min(Math.max(sweepDeg, 0.01), 359.9999);
  // For counter-clockwise, the end angle is start - sw
  const end = startDeg - sw;
  const lg  = sw > 180 ? 1 : 0;
 
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
 
// ── Tooltip ───────────────────────────────────────────────────────────────────
const HoldingsTooltip: React.FC<{ text: string }> = ({ text }) => {
  const [vis, setVis] = React.useState(false);
  const ref = React.useRef<HTMLButtonElement>(null);
  const [pos, setPos] = React.useState({ top: 0, left: 0 });
  const show = () => {
    if (ref.current) {
      const r = ref.current.getBoundingClientRect();
      setPos({ top: r.top + window.scrollY - 8, left: r.left + r.width / 2 });
    }
    setVis(true);
  };
  return (
    <div className="relative inline-flex items-center">
      <button ref={ref} type="button"
        onMouseEnter={show} onMouseLeave={() => setVis(false)}
        onFocus={show} onBlur={() => setVis(false)}
        className="p-0.5 text-slate-600 hover:text-slate-400 transition-colors rounded">
        <Info size={11} />
      </button>
      {vis && (
        <div
          style={{ position:'fixed', top:pos.top, left:pos.left, transform:'translate(-50%,-100%)', zIndex:9999 }}
          className="w-48 bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-[10px] font-bold text-slate-200 shadow-xl pointer-events-none whitespace-normal leading-relaxed animate-in fade-in zoom-in-95 duration-150"
        >
          {text}
          <div className="absolute left-1/2 -translate-x-1/2 top-full w-0 h-0 border-l-4 border-r-4 border-t-4 border-l-transparent border-r-transparent border-t-slate-700" />
        </div>
      )}
    </div>
  );
};
 
// ── Donut Chart (pure SVG — no recharts) ─────────────────────────────────────
interface DonutSlice { ticker: string; pct: number; color: string; }
 
const HoldingsDonutChart: React.FC<{
  slices: DonutSlice[];
  excluded: Set<string>;
  size?: number;
}> = ({ slices, excluded, size = 280 }) => {
  const [hov, setHov] = React.useState<number | null>(null);
  const CX = 150, CY = 150, OUTER = 126, INNER = 68, GAP = 1.4;

  const visibleSlices = React.useMemo(() => {
    const total = slices.filter(s => !excluded.has(s.ticker)).reduce((sum, s) => sum + s.pct, 0);
    if (total === 0) return [];
    return slices
      .filter(s => !excluded.has(s.ticker))
      .map(s => ({ ...s, adjPct: s.pct / total }));
  }, [slices, excluded]);

  const arcs = React.useMemo(() => {
    let cursor = 0;
    return visibleSlices.map(s => {
      const totalSweep = s.adjPct * 360;
      // Start is exactly at cursor, for clockwise or counter-clock
      // To have the first slice start at 0, cursor starts at 0.
      const start      = cursor;
      const sweep      = Math.max(0, totalSweep - GAP);
      cursor          -= totalSweep;
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
        const isHov  = hov === i;
        const dimmed = hov !== null && !isHov;
        return (
          <path
            key={arc.ticker}
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
            fontSize={12} fontWeight={700} letterSpacing={1}>{hovSlice.ticker}</text>
          <text x={CX} y={CY + 10} textAnchor="middle" fill={hovSlice.color}
            fontSize={22} fontWeight={800}>{(hovSlice.adjPct * 100).toFixed(1)}%</text>
          <text x={CX} y={CY + 28} textAnchor="middle" fill="#475569"
            fontSize={9} fontWeight={600}>OF PORTFOLIO</text>
        </>
      ) : (
        <>
          <text x={CX} y={CY - 8} textAnchor="middle" fill="#64748b"
            fontSize={10} fontWeight={700} letterSpacing={0.5}>PORTFOLIO</text>
          <text x={CX} y={CY + 14} textAnchor="middle" fill="#e2e8f0"
            fontSize={16} fontWeight={800}>{visibleSlices.length} STOCKS</text>
        </>
      )}
    </svg>
  );
};
 
// ── Metric box ────────────────────────────────────────────────────────────────
const MBox: React.FC<{
  label: string; value: string; color?: string;
  badge?: number | null; tip?: string;
}> = ({ label, value, color, badge, tip }) => (
  <div className="bg-slate-800/60 border border-slate-700/50 rounded-lg px-3 py-2.5 flex flex-col gap-1 min-w-0">
    <div className="flex items-center justify-between gap-1">
      <span className="text-[9px] font-bold text-slate-500 uppercase tracking-wider truncate">{label}</span>
      <div className="flex items-center gap-1 shrink-0">
        {badge != null && <span className="text-[9px] font-bold text-teal-400 tabular-nums">{badge}</span>}
        {tip && <HoldingsTooltip text={tip} />}
      </div>
    </div>
    <span className="text-[13px] font-bold tabular-nums truncate" style={{ color: color ?? '#f1f5f9' }}>
      {value}
    </span>
  </div>
);

interface SummaryBarItemProps {
  label: string; 
  value: number; 
  color: string; 
  barPct: number; 
  fmtShort: (n: number) => string;
}

const SummaryBarItem: React.FC<SummaryBarItemProps> = ({ 
  label, 
  value, 
  color, 
  barPct, 
  fmtShort 
}) => {
  const [hovered, setHovered] = React.useState(false);
  const fmtFull = (n: number) => {
    const abs = Math.abs(n);
    const sign = n < 0 ? '-৳ ' : '৳ ';
    return sign + abs.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };
  
  return (
    <div className="flex items-center gap-3 group relative"
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}>
      {/* Label */}
      <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider w-20 shrink-0 text-right">
        {label}
      </span>

      {/* Bar track */}
      <div className="flex-1 relative h-6 rounded overflow-visible">
        <div
          className="h-full rounded-r transition-all duration-700 ease-out relative"
          style={{
            width: `${barPct}%`,
            backgroundColor: color,
            border: hovered ? `1.5px solid white` : 'none',
            minWidth: barPct > 0 ? 2 : 0,
            filter: hovered ? 'brightness(1.1)' : 'brightness(1)',
            transition: 'all 0.2s ease',
          }}
        >
          {/* Inline value label — positioned just after bar end */}
          {barPct > 0 && (
            <span
              className="absolute left-full top-1/2 -translate-y-1/2 ml-2 text-[10px] font-bold tabular-nums whitespace-nowrap pointer-events-none"
              style={{ color: value >= 0 ? '#f1f5f9' : '#f87171' }}
            >
              {fmtShort(value)}
            </span>
          )}
        </div>

        {/* Tooltip on hover */}
        {hovered && (
          <div className="absolute left-0 -top-9 z-50 px-3 py-1.5 rounded-lg border text-[11px] font-bold whitespace-nowrap pointer-events-none animate-in fade-in zoom-in-95 duration-150"
            style={{
              backgroundColor: '#0f172a',
              borderColor: `${color}55`,
              color: color,
              boxShadow: `0 4px 20px ${color}33`,
              left: `${Math.min(barPct, 60)}%`,
              transform: 'translateX(-50%)',
            }}>
            {label}: {fmtFull(value)}
            <div className="absolute left-1/2 -translate-x-1/2 top-full w-0 h-0 border-l-4 border-r-4 border-t-4 border-l-transparent border-r-transparent"
              style={{ borderTopColor: `${color}55` }} />
          </div>
        )}
      </div>
    </div>
  );
};
 
// ── Holding Card ──────────────────────────────────────────────────────────────
const HoldingCard: React.FC<{ row: HoldingRow; color: string }> = ({ row, color }) => {
  const unrealPnL = row.currentQty > 0 ? (row.currentPrice - row.avgCost) * row.currentQty : 0;
  
  // Formulas requested by user:
  // p&l% = realized p&l / total cost of sold shares
  const pnlPct = row.totalCostOfSoldShares > 0 
    ? (row.realizedPnL / row.totalCostOfSoldShares) * 100 : 0;
    
  // Net return = realized p&l + dividend - charges
  const netReturn = row.realizedPnL + row.totalDividends - row.totalCharges;
  
  // return % = net return / current holding cost %
  // Interpretation: current holding cost = currentQty * avgCost
  const curHoldingCost = row.currentQty * row.avgCost;
  const returnPct = curHoldingCost > 0 ? (netReturn / curHoldingCost) * 100 : 0;
 
  const G = '#4ade80', R = '#f87171', W = '#f1f5f9';
  const pnlC  = unrealPnL === 0 ? W : (unrealPnL > 0 ? G : R);
  const realPctC = pnlPct === 0 ? W : (pnlPct > 0 ? G : R);
  const netC  = netReturn === 0 ? W : (netReturn > 0 ? G : R);
  const retC  = returnPct === 0 ? W : (returnPct > 0 ? G : R);
  const realC = row.realizedPnL === 0 ? W : (row.realizedPnL > 0 ? G : R);
 
  const isPartial = row.currentQty > 0 && row.totalSellQty > 0;
  const isClosed  = row.currentQty === 0 && row.totalSellQty > 0;
  const badge = isPartial
    ? { label:'PARTIAL', cls:'bg-amber-400/15 border-amber-400/30 text-amber-400' }
    : isClosed
    ? { label:'CLOSED',  cls:'bg-slate-500/15 border-slate-500/30 text-slate-400' }
    : { label:'HOLDING', cls:'bg-teal-400/15  border-teal-400/30  text-teal-400' };
 
  const fmtBD = (n: number): string => {
    const abs = Math.abs(Math.round(n * 100) / 100);
    const fixed = abs.toFixed(2);
    const [intPart, decPart] = fixed.split('.');
    const digits = intPart.replace(/^0+(?=\d)/, '');
    let result = '';
    if (digits.length <= 3) {
      result = digits;
    } else {
      const last3 = digits.slice(-3);
      const rest = digits.slice(0, -3);
      const groups: string[] = [];
      let remaining = rest;
      while (remaining.length > 2) { groups.unshift(remaining.slice(-2)); remaining = remaining.slice(0, -2); }
      if (remaining.length) groups.unshift(remaining);
      result = groups.join(',') + ',' + last3;
    }
    return `৳ ${result}.${decPart}`;
  };
  const bd = fmtBD;
  const bq = (n: number): string => {
    const abs = Math.floor(Math.abs(n));
    const s = String(abs);
    if (s.length <= 3) return s;
    const last3 = s.slice(-3);
    const rest = s.slice(0, -3);
    const groups: string[] = [];
    let remaining = rest;
    while (remaining.length > 2) { groups.unshift(remaining.slice(-2)); remaining = remaining.slice(0, -2); }
    if (remaining.length) groups.unshift(remaining);
    return groups.join(',') + ',' + last3;
  };
  const sgn = (n: number) => `${n >= 0 ? '+' : '-'}${bd(n)}`;
 
  return (
    <div className="bg-slate-900/70 border border-slate-800 rounded-xl p-4 hover:border-slate-700 transition-all duration-200 hover:shadow-[0_0_20px_rgba(0,0,0,0.4)]">
      {/* Header */}
      <div className="flex items-start justify-between gap-3 mb-4">
        <div className="min-w-0">
          <div className="flex items-center gap-2 mb-0.5">
            <div className="w-2 h-2 rounded-full shrink-0"
              style={{ backgroundColor: color, boxShadow: `0 0 6px ${color}55` }} />
            <span className="text-[15px] font-bold text-teal-400 tracking-wide leading-none">{row.ticker}</span>
          </div>
          {row.companyName && (
            <p className="text-[11px] text-slate-500 font-medium truncate ml-4">{row.companyName}</p>
          )}
        </div>
        <span className={cn('text-[9px] font-bold border rounded px-2 py-0.5 uppercase tracking-wider shrink-0', badge.cls)}>
          {badge.label}
        </span>
      </div>
 
      {/* Row 1 */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-2">
        <MBox label="AVG COST"
          value={row.avgCost > 0 ? bd(row.avgCost) : '৳ 0.00'}
          tip="Weighted average price per share for currently held stocks" />
        <MBox label="CURRENT HOLDING"
          value={bd(row.currentQty * (row.currentPrice || row.avgCost))}
          badge={row.currentQty > 0 ? row.currentQty : 0}
          color={row.currentQty > 0 ? '#2dd4bf' : '#475569'}
          tip="Actual buying cost and quantity of shares currently held" />
        <MBox label="TOTAL BUY"
          value={bd(row.totalBuyValue)}
          badge={row.totalBuyQty ? Number(bq(row.totalBuyQty).replace(/,/g, '')) : 0}
          tip="Total buying cost ever and total quantity of shares ever bought" />
        <MBox label="TOTAL SELL"
          value={bd(row.totalSellValue)}
          badge={row.totalSellQty ? Number(bq(row.totalSellQty).replace(/,/g, '')) : undefined}
          tip="Total amount received from selling and total quantity of shares ever sold" />
      </div>
 
      {/* Row 2 */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <MBox label="REALIZED P&L"  value={sgn(row.realizedPnL)} color={realC}
          tip="Profit or loss from completed trades (sold shares)" />
        <MBox label="P&L %"
          value={`${pnlPct >= 0 ? '+' : ''}${pnlPct.toFixed(2)}%`}
          color={realPctC}
          tip="(Realized P&L / Total Cost of Sold Shares) × 100" />
        <MBox label="NET RETURN"    value={sgn(netReturn)}       color={netC}
          tip="Realized P&L + Dividends - Charges for this stock" />
        <MBox label="RETURN %"
          value={`${returnPct >= 0 ? '+' : ''}${returnPct.toFixed(2)}%`}
          color={retC}
          tip="(Net Return / Current Holding Cost) × 100" />
      </div>
    </div>
  );
};
 
// ── Main ──────────────────────────────────────────────────────────────────────
interface PortfolioAllocationProps {
  holdings: DSEHolding[];
  transactions: DseTransaction[];
}
 
const PortfolioAllocation: React.FC<PortfolioAllocationProps> = ({ holdings, transactions }) => {
  const [searchQuery, setSearchQuery] = React.useState('');
  const [excludedTickers, setExcludedTickers] = React.useState<Set<string>>(new Set());

  const toggleExcluded = React.useCallback((ticker: string) => {
    setExcludedTickers(prev => {
      const next = new Set(prev);
      if (next.has(ticker)) next.delete(ticker); else next.add(ticker);
      return next;
    });
  }, []);

  const [sortBy, setSortBy]           = React.useState<'ticker' | 'value' | 'pnl' | 'return'>('ticker');
  const [sortDir, setSortDir]         = React.useState<'asc' | 'desc'>('asc');
  const [isSortMenuOpen, setIsSortMenuOpen] = React.useState(false);
  const [holdingsPortfolioFilter, setHoldingsPortfolioFilter] = React.useState<HoldingsPortfolio>('Investment');
  const [isHoldingsPortfolioMenuOpen, setIsHoldingsPortfolioMenuOpen] = React.useState(false);
 
  // ── Build rows from transactions ──────────────────────────────────────────
  const allRows = React.useMemo<HoldingRow[]>(() => {
    type E = {
      ticker: string; companyName: string; portfolio: HoldingsPortfolio;
      buyQty: number; buyValue: number;
      sellQty: number; sellValue: number;
      costBasis: number; realizedPnL: number;
      dividends: number; charges: number;
      costOfSoldShares: number;
    };
    const map: Record<string, E> = {};
 
    [...transactions]
      .filter(t => ['Buy', 'Sell', 'Dividend', 'Charge'].includes(t.type))
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
      .forEach(t => {
        if (!t.ticker) return;
        const key = `${t.portfolio}|${t.ticker}`;
        if (!map[key]) map[key] = {
          ticker: t.ticker, companyName: t.companyName || '',
          portfolio: t.portfolio as HoldingsPortfolio,
          buyQty: 0, buyValue: 0, sellQty: 0, sellValue: 0,
          costBasis: 0, realizedPnL: 0,
          dividends: 0, charges: 0, costOfSoldShares: 0,
        };
        const e = map[key];
        
        if (t.type === 'Buy') {
          e.buyQty   += t.qty;
          e.buyValue += t.total;
          e.costBasis += t.total;
        } else if (t.type === 'Sell') {
          const held = e.buyQty - e.sellQty;
          if (held > 0) {
            const avg  = e.costBasis / held;
            const cost = avg * Math.min(t.qty, held);
            e.realizedPnL += t.total - cost;
            e.costBasis    = Math.max(0, e.costBasis - cost);
            e.costOfSoldShares += cost;
          }
          e.sellQty   += t.qty;
          e.sellValue += t.total;
        } else if (t.type === 'Dividend') {
          e.dividends += t.total;
        } else if (t.type === 'Charge') {
          e.charges += t.total;
        }
      });
 
    const priceMap: Record<string, number> = {};
    holdings.forEach(h => { priceMap[h.symbol] = h.currentPrice; });
 
    return Object.values(map).map(e => {
      const currentQty = Math.max(0, e.buyQty - e.sellQty);
      const avgCost    = currentQty > 0
        ? e.costBasis / currentQty
        : e.buyQty > 0 ? e.buyValue / e.buyQty : 0;
      return {
        ticker: e.ticker, companyName: e.companyName, portfolio: e.portfolio,
        totalBuyQty: e.buyQty, totalBuyValue: e.buyValue,
        totalSellQty: e.sellQty, totalSellValue: e.sellValue,
        currentQty, avgCost,
        currentPrice: priceMap[e.ticker] ?? avgCost,
        realizedPnL: e.realizedPnL,
        totalDividends: e.dividends,
        totalCharges: e.charges,
        totalCostOfSoldShares: e.costOfSoldShares,
      };
    });
  }, [transactions, holdings]);
 
  // ── Tab filter & Global Consolidation ────────────────────────────────────
  const tabRows = React.useMemo(() => {
    if (holdingsPortfolioFilter !== 'Global') {
      return allRows.filter(r => r.portfolio === holdingsPortfolioFilter);
    }
    
    // Aggregate by ticker for Global view
    const aggregated: Record<string, HoldingRow & { _costBasis: number }> = {};
    const priceMap: Record<string, number> = {};
    holdings.forEach(h => { priceMap[h.symbol] = h.currentPrice; });

    allRows.forEach(row => {
      // Each row's cost basis for currently held shares = currentQty * avgCost
      const rowCostBasis = row.currentQty * row.avgCost;
      if (!aggregated[row.ticker]) {
        aggregated[row.ticker] = { ...row, portfolio: 'Global', _costBasis: rowCostBasis };
      } else {
        const target = aggregated[row.ticker];
        target.totalBuyQty += row.totalBuyQty;
        target.totalBuyValue += row.totalBuyValue;
        target.totalSellQty += row.totalSellQty;
        target.totalSellValue += row.totalSellValue;
        target.currentQty += row.currentQty;
        target.realizedPnL += row.realizedPnL;
        target.totalDividends += row.totalDividends;
        target.totalCharges += row.totalCharges;
        target.totalCostOfSoldShares += row.totalCostOfSoldShares;
        target._costBasis += rowCostBasis;

        // avgCost = cost of currently held shares / qty currently held
        target.avgCost = target.currentQty > 0 ? target._costBasis / target.currentQty : 0;
        target.currentPrice = priceMap[row.ticker] ?? target.avgCost;
      }
    });
    return Object.values(aggregated).map(({ _costBasis, ...row }) => row);
  }, [allRows, holdingsPortfolioFilter, holdings]);
 
  const activeRows   = React.useMemo(() => tabRows.filter(r => r.currentQty > 0),  [tabRows]);
  const inactiveRows = React.useMemo(() => tabRows.filter(r => r.currentQty === 0), [tabRows]);
 
  // ── Donut — only stocks currently held ───────────────────────────────────
  const donutSlices = React.useMemo<DonutSlice[]>(() => {
    const items = activeRows
      .map(r => ({ ticker: r.ticker, value: r.currentQty * (r.currentPrice || r.avgCost) }))
      .filter(r => r.value > 0)
      .sort((a, b) => b.value - a.value);
    const total = items.reduce((s, i) => s + i.value, 0);
    if (total === 0) return [];
    return items.map((item, idx) => ({
      ticker: item.ticker,
      pct:   item.value / total,
      color: HOLDING_COLORS[idx % HOLDING_COLORS.length],
    }));
  }, [activeRows]);
 
  const colorMap = React.useMemo(() => {
    const m: Record<string, string> = {};
    donutSlices.forEach(s => { m[s.ticker] = s.color; });
    return m;
  }, [donutSlices]);
 
  const totalCurrentValue = tabRows.reduce((s, r) => s + (r.currentQty * r.currentPrice), 0);
 
  // ── Sort ──────────────────────────────────────────────────────────────────
  const sortRows = (rows: HoldingRow[]) =>
    [...rows].sort((a, b) => {
      let cmp = 0;
      if (sortBy === 'ticker') {
        cmp = a.ticker.localeCompare(b.ticker);
      } else if (sortBy === 'value') {
        cmp = (b.currentQty * b.currentPrice) - (a.currentQty * a.currentPrice);
      } else if (sortBy === 'pnl') {
        const pA = (a.currentPrice - a.avgCost) * a.currentQty + a.realizedPnL;
        const pB = (b.currentPrice - b.avgCost) * b.currentQty + b.realizedPnL;
        cmp = pB - pA;
      } else {
        const rA = a.totalBuyValue > 0
          ? (a.totalSellValue + a.currentQty * a.currentPrice - a.totalBuyValue) / a.totalBuyValue : 0;
        const rB = b.totalBuyValue > 0
          ? (b.totalSellValue + b.currentQty * b.currentPrice - b.totalBuyValue) / b.totalBuyValue : 0;
        cmp = rB - rA;
      }
      return sortDir === 'asc' ? cmp : -cmp;
    });
 
  const handleSort = (key: typeof sortBy) => {
    if (sortBy === key) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortBy(key); setSortDir('desc'); }
  };
 
  const filterSearch = (rows: HoldingRow[]) => {
    if (!searchQuery.trim()) return rows;
    const q = searchQuery.toLowerCase();
    return rows.filter(r =>
      r.ticker.toLowerCase().includes(q) || r.companyName.toLowerCase().includes(q));
  };
 
  const visActive   = sortRows(filterSearch(activeRows));
  const visInactive = sortRows(filterSearch(inactiveRows));
 
  // ── Sort button ───────────────────────────────────────────────────────────
  const SORT_OPTIONS: { col: typeof sortBy; label: string }[] = [
    { col: 'ticker', label: 'Ticker' },
    { col: 'value',  label: 'Holding'  },
    { col: 'pnl',    label: 'P&L'    },
    { col: 'return', label: 'Return' },
  ];
 
  const TAB_ACTIVE: Record<HoldingsPortfolio, string> = {
    Investment: 'bg-blue-500   text-white border-blue-400/30   shadow-lg shadow-blue-500/20',
    Trading:    'bg-amber-500  text-white border-amber-400/30  shadow-lg shadow-amber-500/20',
    Global:     'bg-purple-500 text-white border-purple-400/30 shadow-lg shadow-purple-500/20',
  };
  const TAB_IDLE: Record<HoldingsPortfolio, string> = {
    Investment: 'bg-blue-500/10   border-blue-400/20   text-blue-300   hover:bg-blue-500/20',
    Trading:    'bg-amber-500/10  border-amber-400/20  text-amber-300  hover:bg-amber-500/20',
    Global:     'bg-purple-500/10 border-purple-400/20 text-purple-300 hover:bg-purple-500/20',
  };
 
  return (
    <div className="space-y-6">
 
      {/* ── Portfolio Allocation Header ────────────────────────────────────── */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div className="flex flex-wrap items-center gap-4 sm:gap-8">
          <div className="flex items-center gap-2">
            <PieChart size={22} className="text-teal-400" />
            <h2 className="text-subheading font-bold text-white uppercase tracking-wider">Portfolio Allocation</h2>
          </div>

          {/* Portfolio tabs - same line as title, linked with dropdown */}
          <div className="flex flex-nowrap items-center gap-1 bg-slate-950/50 rounded-xl p-0.5 border border-slate-800/50">
            {(['Investment', 'Trading', 'Global'] as const).map(tab => (
              <button key={tab} 
                onClick={() => { setHoldingsPortfolioFilter(tab); setIsHoldingsPortfolioMenuOpen(false); }}
                className={cn(
                  'flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[10px] sm:text-[12px] font-bold uppercase transition-all border outline-none active:scale-95 hover:opacity-90',
                  holdingsPortfolioFilter === tab
                    ? tab === 'Global'
                      ? 'bg-purple-500 text-slate-950 border-purple-400/40 shadow-lg shadow-purple-500/20'
                      : tab === 'Trading'
                      ? 'bg-amber-500 text-slate-950 border-amber-400/40 shadow-lg shadow-amber-500/20'
                      : 'bg-blue-500 text-slate-950 border-blue-400/40 shadow-lg shadow-blue-500/20'
                    : 'bg-slate-900/40 border-slate-800/40 text-slate-500 hover:text-white hover:bg-slate-800/60'
                )}>
                <span className={cn('w-2 h-2 rounded-full shrink-0',
                  holdingsPortfolioFilter === tab 
                    ? 'bg-slate-950/60' 
                    : (tab === 'Global' ? 'bg-purple-400' : tab === 'Trading' ? 'bg-amber-400' : 'bg-blue-400')
                )} />
                {tab}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* ── Portfolio Chart Area Card ─────────────────────────────────────── */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-8 lg:p-10 shadow-xl" style={{ marginTop: 16 }}>
        {donutSlices.length === 0 ? (
          <div className="flex items-center justify-center py-20 text-slate-600 text-[11px] font-bold uppercase tracking-wider">
            No active holdings for this portfolio
          </div>
        ) : (
          <div className="flex flex-col lg:flex-row items-center gap-10 lg:gap-16">
            {/* Chart - 60% width */}
            <div className="w-full lg:w-[60%] flex justify-center">
              <HoldingsDonutChart
                slices={donutSlices}
                excluded={excludedTickers}
                size={340}
              />
            </div>
 
            {/* Legend - 40% width, clickable */}
            <div className="w-full lg:w-[40%]">
              <div className={cn(
                "columns-1 gap-x-8 space-y-0.5",
                donutSlices.length > 8 && "lg:columns-2"
              )}>
                {donutSlices.map(s => {
                  const isExcluded = excludedTickers.has(s.ticker);
                  return (
                    <button
                      key={s.ticker}
                      onClick={() => toggleExcluded(s.ticker)}
                      className="flex items-center gap-2.5 min-w-0 break-inside-avoid py-0.5 transition-all hover:opacity-80 active:scale-95 text-left w-full"
                      style={{ opacity: isExcluded ? 0.35 : 1 }}
                    >
                      <div 
                        className="w-2 h-2 rounded-sm shrink-0" 
                        style={{ 
                          backgroundColor: s.color, 
                          boxShadow: isExcluded ? 'none' : `0 0 8px ${s.color}55` 
                        }} 
                      />
                      <span className={cn(
                        'text-[11px] font-bold text-slate-400 truncate tracking-wide transition-all',
                        isExcluded && 'line-through text-slate-600'
                      )}>
                        {s.ticker}&nbsp;<span style={{ color: isExcluded ? undefined : s.color }}>({(s.pct * 100).toFixed(1)}%)</span>
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>
 
      {/* ── Current Holdings ─────────────────────────────────────────────── */}
      <div className="space-y-4">
        {/* Header */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <Briefcase size={20} className="text-teal-400" />
            <h2 className="text-subheading font-bold text-white uppercase tracking-wider">Current Holdings</h2>
          </div>
          <div className="flex items-center gap-2 flex-wrap justify-end">
            {/* Search */}
            <div className="flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-lg px-3 h-9 hover:border-slate-700 focus-within:border-teal-400/50 transition-colors cursor-text"
              onClick={() => document.getElementById('hld-search')?.focus()}>
              <Search size={13} className="text-slate-500 shrink-0" />
              <input id="hld-search" type="text" placeholder="Search ticker..."
                value={searchQuery} onChange={e => setSearchQuery(e.target.value)}
                className="bg-transparent border-none focus:ring-0 text-[10px] font-bold text-white placeholder:text-slate-600 uppercase outline-none w-32" />
            </div>

            {/* Portfolio filter */}
            <div className="relative shrink-0">
              <button
                onClick={() => { setIsHoldingsPortfolioMenuOpen(p => !p); setIsSortMenuOpen(false); }}
                className={cn(
                  'flex items-center gap-2 px-3 h-9 rounded-lg transition-all text-[10px] font-bold uppercase border',
                  holdingsPortfolioFilter === 'Investment'
                    ? 'bg-blue-500/20 border-blue-400/40 text-blue-300 hover:bg-blue-500/30'
                    : holdingsPortfolioFilter === 'Trading'
                    ? 'bg-amber-500/20 border-amber-400/40 text-amber-300 hover:bg-amber-500/30'
                    : 'bg-purple-500/20 border-purple-400/40 text-purple-300 hover:bg-purple-500/30'
                )}>
                <Filter size={13} />
                {holdingsPortfolioFilter}
              </button>
              {isHoldingsPortfolioMenuOpen && (
                <>
                  <div className="fixed inset-0 z-40" onClick={() => setIsHoldingsPortfolioMenuOpen(false)} />
                  <div className="absolute right-0 mt-2 w-44 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-2 animate-in fade-in zoom-in-95">
                    {(['Investment', 'Trading', 'Global'] as const).map(p => (
                      <button key={p} onClick={() => { setHoldingsPortfolioFilter(p); setIsHoldingsPortfolioMenuOpen(false); }}
                        className={cn(
                          'w-full flex items-center gap-2.5 px-3 py-1.5 rounded-lg text-[10px] font-bold uppercase transition-colors',
                          holdingsPortfolioFilter === p ? 'bg-teal-400/10 text-teal-400' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                        )}>
                        <span className={cn('w-2 h-2 rounded-full shrink-0', PORTFOLIO_DOT[p])} />
                        {p}
                      </button>
                    ))}
                  </div>
                </>
              )}
            </div>

            {/* Sort dropdown */}
            <div className="relative shrink-0 flex items-center bg-slate-900 border border-slate-800 rounded-lg h-9">
              <button
                onClick={(e) => { 
                  e.stopPropagation();
                  setIsSortMenuOpen(!isSortMenuOpen); 
                  setIsHoldingsPortfolioMenuOpen(false); 
                }}
                className={cn(
                  'flex items-center gap-2 px-3 h-full transition-all text-[10px] font-bold uppercase border-r border-slate-800 hover:bg-slate-800/60',
                  isSortMenuOpen ? 'text-teal-400 bg-slate-800' : 'text-slate-400 hover:text-white'
                )}>
                <ArrowUpDown size={13} />
                Sort: {SORT_OPTIONS.find(o => o.col === sortBy)?.label}
              </button>
              
              {/* Direct toggle direction buttons */}
              <div className="flex items-center h-full">
                <button 
                  onClick={(e) => { e.stopPropagation(); setSortDir('asc'); }}
                  className={cn(
                    'px-2.5 h-full transition-colors flex items-center justify-center border-r border-slate-800 hover:bg-teal-400/10',
                    sortDir === 'asc' ? 'text-teal-400 bg-teal-400/5' : 'text-slate-600 hover:text-slate-400'
                  )}
                >
                  <ArrowUp size={12} className={cn('transition-transform', sortDir === 'asc' && 'scale-110')} />
                </button>
                <button 
                  onClick={(e) => { e.stopPropagation(); setSortDir('desc'); }}
                  className={cn(
                    'px-2.5 h-full transition-colors flex items-center justify-center hover:bg-teal-400/10',
                    sortDir === 'desc' ? 'text-teal-400 bg-teal-400/5' : 'text-slate-600 hover:text-slate-400'
                  )}
                >
                  <ArrowDown size={12} className={cn('transition-transform', sortDir === 'desc' && 'scale-110')} />
                </button>
              </div>

              {isSortMenuOpen && (
                <>
                  <div className="fixed inset-0 z-[60]" onClick={() => setIsSortMenuOpen(false)} />
                  <div className="absolute right-0 top-full mt-2 w-44 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-[70] p-2 animate-in fade-in slide-in-from-top-1">
                    <div className="px-3 py-1.5 text-[9px] font-bold text-slate-500 uppercase tracking-widest border-b border-slate-800 mb-1">Sort Columns</div>
                    {SORT_OPTIONS.map(({ col, label }) => (
                      <button key={col} 
                        onClick={(e) => { e.stopPropagation(); setSortBy(col); setIsSortMenuOpen(false); }}
                        className={cn(
                          'w-full flex items-center justify-between px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-colors',
                          sortBy === col ? 'bg-teal-400/10 text-teal-400' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                        )}>
                        {label}
                        {sortBy === col && (sortDir === 'asc' ? <ArrowUp size={10} /> : <ArrowDown size={10} />)}
                      </button>
                    ))}
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
 
        {/* Empty state */}
        {visActive.length === 0 && visInactive.length === 0 ? (
          <div className="py-16 text-center bg-slate-900/30 border border-dashed border-slate-800 rounded-xl">
            <p className="text-[11px] font-bold text-slate-500 uppercase italic">
              {searchQuery ? 'No stocks match your search' : 'No transactions found for this portfolio'}
            </p>
          </div>
        ) : (
          <div className="space-y-3">
            {[...visActive, ...visInactive].map(row => (
              <HoldingCard
                key={`${row.portfolio}|${row.ticker}`}
                row={row}
                color={colorMap[row.ticker] ?? '#2dd4bf'}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
 
// ══ END ════════════════════════════════════════════════════════════════════════


// ── Main Component ─────────────────────────────────────────────────────────

export const DseTrackerModule: React.FC<DseTrackerModuleProps & { activeTab?: DseTab }> = ({
  holdings, onAdd, onUpdate, onDelete, onTitleChange, triggerAdd, setTriggerAdd, activeTab = 'summary'
}) => {
  const [isSettingsMenuOpen, setIsSettingsMenuOpen] = useState(false);
  const [isMobile, setIsMobile] = useState(false);
  useEffect(() => {
    const handleResize = () => setIsMobile(window.innerWidth < 768);
    handleResize();
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);
  const [isRangeMenuOpen, setIsRangeMenuOpen] = useState(false);
  const [historyRange, setHistoryRange] = useState<'all' | 'last12m' | 'fiscal' | 'custom'>('all');
  const [historyCustomDates, setHistoryCustomDates] = useState(() => {
    const now = new Date();
    return { start: getFirstOfMonth(now), end: getTodayStr() };
  });

  // ── Analytics Filters ──
  const [analyticsView, setAnalyticsView] = useState<'monthly' | 'cumulative'>('monthly');
  const [analyticsRange, setAnalyticsRange] = useState<'last6m' | 'last12m' | 'fiscal' | 'custom'>(() => 
    typeof window !== 'undefined' && window.innerWidth < 768 ? 'last6m' : 'last12m'
  );
  const [analyticsCustomDates, setAnalyticsCustomDates] = useState(() => {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth() - 5, 1);
    return { start: toDateStr(start), end: getTodayStr() };
  });
  const [isAnalyticsRangeOpen, setIsAnalyticsRangeOpen] = useState(false);
  const [hiddenSeries, setHiddenSeries] = useState<Set<string>>(new Set());
  const [returnHistoryView, setReturnHistoryView] = useState<'monthly' | 'cumulative'>('monthly');
  const [returnHistoryRange, setReturnHistoryRange] = useState<'last6m' | 'last12m' | 'fiscal' | 'custom'>(() => 
    typeof window !== 'undefined' && window.innerWidth < 768 ? 'last6m' : 'last12m'
  );
  const [returnHistoryCustomDates, setReturnHistoryCustomDates] = useState(() => {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth() - 5, 1);
    return { start: toDateStr(start), end: getTodayStr() };
  });
  const [isReturnHistoryRangeOpen, setIsReturnHistoryRangeOpen] = useState(false);
  const [returnByStockType, setReturnByStockType] = useState<'Dividend' | 'Realized P&L' | 'Net Return'>('Realized P&L');
  const [customNavMonth, setCustomNavMonth] = useState(() => {
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() };
  });

  const [tableMode, setTableMode] = useState<'1m' | '1y' | 'custom'>('1m');
  const [tableOffset, setTableOffset] = useState(0);
  const [tableCustomDates, setTableCustomDates] = useState(() => {
    const now = new Date();
    return { start: getFirstOfMonth(now), end: getTodayStr() };
  });
  const [isTableTimeMenuOpen, setIsTableTimeMenuOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedPortfolio, setSelectedPortfolio] = useState<'Global' | 'Investment' | 'Trading'>('Global');
  const [isPortfolioMenuOpen, setIsPortfolioMenuOpen] = useState(false);
  const [selectedTypes, setSelectedTypes] = useState<string[]>([]);
  const [isTypeMenuOpen, setIsTypeMenuOpen] = useState(false);
  const [sortBy, setSortBy] = useState<'date' | 'total' | 'ticker' | 'type'>('date');
const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
const [selectedIds, setSelectedIds] = useState<string[]>([]);

  const [confirmDialog, setConfirmDialog] = useState<{
    isOpen: boolean; 
    title: string;
    message: string; 
    onConfirm: () => void;
    confirmLabel?: string;
    variant?: 'danger' | 'warning' | 'info';
  }>({ isOpen: false, title: '', message: '', onConfirm: () => {} });

  const [transactions, setTransactions] = useState<DseTransaction[]>(() => getCachedData('dse'));

  useEffect(() => {
    const handleDseUpdate = () => {
      const fresh = getCachedData('dse');
      if (Array.isArray(fresh)) {
        setTransactions(fresh.filter((r: any) => r.type !== '__SETTINGS__'));
      }
    };
    window.addEventListener('storage', handleDseUpdate);
    window.addEventListener('dse-updated', handleDseUpdate);
    return () => {
      window.removeEventListener('storage', handleDseUpdate);
      window.removeEventListener('dse-updated', handleDseUpdate);
    };
  }, []);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isError, setIsError] = useState(false);
  const [lastSynced, setLastSynced] = useState<Date | null>(() => {
  try { const t = localStorage.getItem('sheet_cache_time_dse'); return t ? new Date(parseInt(t)) : null; } catch { return null; }
});
  const [apiStatus, setApiStatus] = useState<{ message: string; isError: boolean } | null>(null);
  const [newTicker, setNewTicker] = useState('');
  const [newCompany, setNewCompany] = useState('');

  const [dseSettings, setDseSettings] = useState<DseSettings>(() => {
    try {
      const saved = localStorage.getItem('dse_settings');
      return saved ? JSON.parse(saved) : { commissionRate: 0.4, customStocks: [] };
    } catch {
      return { commissionRate: 0.4, customStocks: [] };
    }
  });

  const onAddStock = () => {
    if (!newTicker || !newCompany) return;
    const newSettings = { 
      ...dseSettings, 
      customStocks: [...dseSettings.customStocks, { ticker: newTicker, name: newCompany }] 
    };
    saveSettings(newSettings);
    setNewTicker('');
    setNewCompany('');
  };

  const saveSettings = async (newSettings: DseSettings) => {
    setDseSettings(newSettings);
    localStorage.setItem('dse_settings', JSON.stringify(newSettings));

    try {
      const settingsRow = {
        id: '__DSE_SETTINGS__',
        type: '__SETTINGS__',
        settings: newSettings
      };

      await pushToSheet('dse', 'update', {
        data: settingsRow
      });
    } catch (err) {
      console.error('Failed to sync settings', err);
    }
  };


  const exportData = () => {
    const rows = transactions.map(t => ({
      'Date': t.date,
      'Type': t.type,
      'Portfolio': t.portfolio,
      'Ticker': t.ticker,
      'Company Name': t.companyName,
      'Quantity': t.qty,
      'Price': t.price,
      'Commission': t.commission,
      'Total': t.total,
      'Notes': t.notes
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DSE Transactions");
    XLSX.writeFile(wb, `DSE_Export_${new Date().toISOString().split('T')[0]}.xlsx`);
    setApiStatus({ message: 'Export successful!', isError: false });
    setTimeout(() => setApiStatus(null), 3000);
  };

  const importData = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setConfirmDialog({
      isOpen: true,
      title: 'Confirm Replace',
      message: 'Are you sure you want to replace all existing DSE transactions with the imported data? This action cannot be undone.',
      confirmLabel: 'Replace All',
      variant: 'warning',
      onConfirm: async () => {
        setConfirmDialog(prev => ({ ...prev, isOpen: false }));
        const reader = new FileReader();
        reader.onload = async (event) => {
          try {
            const data = new Uint8Array(event.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];

            const imported = jsonData.map(row => ({
              id: crypto.randomUUID(),
              date: String(row['Date'] || getTodayStr()),
              type: (row['Type'] || 'Buy') as any,
              portfolio: (row['Portfolio'] || 'Investment') as any,
              ticker: String(row['Ticker'] || ''),
              companyName: String(row['Company Name'] || ''),
              qty: Number(row['Quantity'] || 0),
              price: Number(row['Price'] || 0),
              commission: Number(row['Commission'] || 0),
              total: Number(row['Total'] || 0),
              notes: String(row['Notes'] || '')
            }));

            if (imported.length > 0) {
              // For DSE, we've decided to replace all data on import instead of appending.
              const merged = imported;

              // Update UI immediately
              setTransactions(merged);

              // Save local cache
              localStorage.setItem('sheet_cache_dse', JSON.stringify(merged));

              // IMPORTANT:
              // Mark local data newer than cloud
              markDirty('dse');

              // Push FULL restored dataset to Google Sheets
              const ok = await pushModuleData('dse', merged);

              setApiStatus({
                message: ok
                  ? `Imported & synced ${imported.length} transactions!`
                  : `Imported locally, but cloud sync failed.`,
                isError: !ok
              });
            } else {
              setApiStatus({ message: 'No valid data found.', isError: true });
            }
          } catch (err) {
            console.error('Import failed', err);
            setApiStatus({ message: 'Import failed.', isError: true });
          }
          setTimeout(() => setApiStatus(null), 3000);
        };
        reader.readAsArrayBuffer(file);
      }
    });

    e.target.value = '';
  };

  const generateTemplate = () => {
    const headers = [
      {
        'Date': getTodayStr(),
        'Type': 'Buy',
        'Portfolio': 'Investment',
        'Ticker': 'GP',
        'Company Name': 'Grameenphone Ltd',
        'Quantity': 100,
        'Price': 250,
        'Commission': 100,
        'Total': 25100,
        'Notes': 'Sample Row'
      }
    ];
    const ws = XLSX.utils.json_to_sheet(headers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "DSE_Import_Template.xlsx");
  };

  // ── Fetch ──
  
const fetchFromSheets = useCallback(async (force = false) => {
  setIsSyncing(true); setIsError(false);
  const result = await fetchFromSheet('dse');
  if (result) {
    // Only replace local data if cloud returned non-empty, 
    // OR if local is also empty (avoid wiping local data with empty cloud)
    if (result.data.length > 0) {

  const settingsRow = result.data.find(
    (r: any) => r.type === '__SETTINGS__'
  );

  if (settingsRow?.settings) {
    setDseSettings(settingsRow.settings);

    localStorage.setItem(
      'dse_settings',
      JSON.stringify(settingsRow.settings)
    );
  }

  const cleanTransactions = result.data.filter(
    (r: any) => r.type !== '__SETTINGS__'
  );

  setTransactions(cleanTransactions);

  try {
    localStorage.setItem(
      'sheet_cache_dse',
      JSON.stringify(cleanTransactions)
    );
  } catch {}
}
    setLastSynced(new Date());
    setIsError(false);
  } else {
    setIsError(true);
    // On failure, re-read from localStorage as fallback
    try {
      const cached = localStorage.getItem('sheet_cache_dse');
      if (cached) {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed) && parsed.length > 0) setTransactions(parsed);
      }
    } catch {}
  }
  setIsSyncing(false);
}, []);

  useEffect(() => { fetchFromSheets(); }, [fetchFromSheets]);

  // ── Push to Sheets ──
  // BEFORE: the whole pushToSheets useCallback

// AFTER:
const pushToSheets = useCallback(async (action: 'update' | 'delete', transaction: DseTransaction) => {
  setApiStatus(null);
  const payload = action === 'update'
    ? { data: transaction }
    : { id: transaction.id };
  const ok = await pushToSheet('dse', action, payload);
  setApiStatus({
    message: ok
      ? (action === 'delete' ? 'Deleted from Google Sheets.' : 'Saved to Google Sheets.')
      : 'Google Sheets sync failed.',
    isError: !ok,
  });
  setTimeout(() => setApiStatus(null), 4000);
}, []);

  // ── Delete with confirmation ──
  const handleDeleteTransaction = (t: DseTransaction) => {
    setConfirmDialog({
      isOpen: true,
      title: 'Confirm Delete',
      message: `Delete ${typeLabel(t.type)} transaction${t.ticker ? ` for ${t.ticker}` : ''} on ${formatDDMMYYYY(t.date)}?`,
      confirmLabel: 'Delete',
      variant: 'danger',
      onConfirm: async () => {
        setConfirmDialog(prev => ({ ...prev, isOpen: false }));
        setTransactions(prev => {
          const updated = prev.filter(tx => tx.id !== t.id);
          localStorage.setItem('sheet_cache_dse', JSON.stringify(updated));
          return updated;
        });
        await pushToSheets('delete', t);
      },
    });
  };

  const handleBatchDelete = () => {
    if (selectedIds.length === 0) return;
    setConfirmDialog({
      isOpen: true,
      title: 'Confirm Delete',
      message: `Delete ${selectedIds.length} selected transactions?`,
      confirmLabel: 'Delete All',
      variant: 'danger',
      onConfirm: async () => {
        setConfirmDialog(prev => ({ ...prev, isOpen: false }));
        const toDelete = transactions.filter(t => selectedIds.includes(t.id));
        setTransactions(prev => {
          const updated = prev.filter(t => !selectedIds.includes(t.id));
          localStorage.setItem('sheet_cache_dse', JSON.stringify(updated));
          return updated;
        });
        setSelectedIds([]);
        // Sync each deletion (ideally batch, but existing pushToSheets handles one by one)
        for (const t of toDelete) {
          await pushToSheets('delete', t);
        }
      }
    });
  };

  // Transaction Modal State
  const [isTransactionModalOpen, setIsTransactionModalOpen] = useState(false);
  const [editingTransaction, setEditingTransaction] = useState<DseTransaction | null>(null);

  const openTransactionModal = (t?: DseTransaction) => {
    setEditingTransaction(t || null);
    setIsTransactionModalOpen(true);
  };

  const handleEditTransaction = (t: DseTransaction) => openTransactionModal(t);

  useEffect(() => {
    if (triggerAdd && setTriggerAdd) {
      openTransactionModal();
      setTriggerAdd(false);
    }
  }, [triggerAdd, setTriggerAdd]);

  // Save handler (called by both Add and Next)
  const handleSaveTransaction = async (t: DseTransaction) => {
    if (editingTransaction) {
      setTransactions(prev => {
        const updated = prev.map(tx => tx.id === t.id ? t : tx);
        localStorage.setItem('sheet_cache_dse', JSON.stringify(updated));
        return updated;
      });
      setIsTransactionModalOpen(false);
    } else {
      setTransactions(prev => {
        const updated = [...prev, t];
        localStorage.setItem('sheet_cache_dse', JSON.stringify(updated));
        return updated;
      });
      // For "Add" (not Next), close is handled by caller; for "Next" modal stays open
      setIsTransactionModalOpen(false);
    }
    await pushToSheets('update', t);
  };

  // Save and keep open (Next)
  const handleSaveAndNext = async (t: DseTransaction) => {
    setTransactions(prev => {
      const updated = [...prev, t];
      localStorage.setItem('sheet_cache_dse', JSON.stringify(updated));
      return updated;
    });
    await pushToSheets('update', t);
    // Modal stays open — TransactionModal handles its own reset via resetForNext()
  };

  // ── Stats range navigation ──
  const navigateCustomMonth = (direction: -1 | 1) => {
    const now = new Date();
    let { year, month } = customNavMonth;
    month += direction;
    if (month < 0) { month = 11; year--; }
    if (month > 11) { month = 0; year++; }
    const navDate = new Date(year, month, 1);
    const isCurrent = year === now.getFullYear() && month === now.getMonth();
    setCustomNavMonth({ year, month });
    setHistoryCustomDates({ start: getFirstOfMonth(navDate), end: isCurrent ? getTodayStr() : getLastOfMonth(navDate) });
  };

  const handleRangeChange = (newRange: 'all' | 'last12m' | 'fiscal' | 'custom') => {
    if (newRange === 'custom') {
      const now = new Date();
      setCustomNavMonth({ year: now.getFullYear(), month: now.getMonth() });
      setHistoryCustomDates({ start: getFirstOfMonth(now), end: getTodayStr() });
    }
    setHistoryRange(newRange);
    setIsRangeMenuOpen(false);
  };

  const { startStr, endStr } = useMemo(() => {
    const now = new Date();
    const toStr = (d: Date) =>
      `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    if (historyRange === 'all') return { startStr: '0000-01-01', endStr: '9999-12-31' };
    if (historyRange === 'last12m') return {
      startStr: toStr(new Date(now.getFullYear(), now.getMonth() - 11, 1)),
      endStr: toStr(new Date(now.getFullYear(), now.getMonth() + 1, 0)),
    };
    if (historyRange === 'fiscal') {
      const fy = now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
      return { startStr: `${fy}-07-01`, endStr: `${fy + 1}-06-30` };
    }
    return { startStr: historyCustomDates.start || '0000-01-01', endStr: historyCustomDates.end || '9999-12-31' };
  }, [historyRange, historyCustomDates]);

  // ── Title ──
  useEffect(() => {
    if (!onTitleChange) return;
    if (historyRange === 'all') { onTitleChange('DSE Tracker'); return; }
    const now = new Date();
    let startDate: Date, endDate: Date;
    if (historyRange === 'last12m') {
      startDate = new Date(now.getFullYear(), now.getMonth() - 11, 1); endDate = now;
    } else if (historyRange === 'fiscal') {
      const fy = now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
      startDate = new Date(fy, 6, 1); endDate = new Date(fy + 1, 5, 30);
    } else {
      const [sy, sm] = (historyCustomDates.start || '').split('-').map(Number);
      const [ey, em] = (historyCustomDates.end || '').split('-').map(Number);
      startDate = sy ? new Date(sy, sm - 1, 1) : now;
      endDate = ey ? new Date(ey, em - 1, 1) : now;
    }
    onTitleChange(
      <span className="flex items-center gap-2">
        DSE TRACKER{' '}
        <span className="text-teal-400 font-display text-sm font-bold opacity-100 tracking-wider leading-none">
          / {startDate.toLocaleString('default', { month: 'short', year: 'numeric' })} - {endDate.toLocaleString('default', { month: 'short', year: 'numeric' })}
        </span>
      </span>
    );
    return () => { onTitleChange('DSE Tracker'); };
  }, [historyRange, historyCustomDates, onTitleChange]);

  // ── Stats computation ──
  const stats = useMemo(() => {
    const inRange = (d: string) => d >= startStr && d <= endStr;
    const beforeEnd = (d: string) => d <= endStr;
    const typePriority: Record<string, number> = { Deposit: 1, Buy: 2, Dividend: 3, Sell: 4, Charge: 5 };
    const sorted = [...transactions].sort((a, b) => {
      const dd = new Date(a.date).getTime() - new Date(b.date).getTime();
      return dd !== 0 ? dd : (typePriority[a.type] || 99) - (typePriority[b.type] || 99);
    });
    let totalDeposits = 0, totalWithdrawals = 0, totalDividends = 0, totalCharges = 0, totalRealizedPnL = 0;
    let totalDepositsRange = 0, totalRealizedPnLRange = 0, totalDividendsRange = 0, totalChargesRange = 0;
    let totalDepositsThisMonth = 0;
    const monthStart = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}-01`;
    const hMap: Record<string, { qty: number; totalCost: number; portfolio: string }> = {};

    sorted.forEach(t => {
      const key = t.ticker ? `${t.portfolio}|${t.ticker}` : null;
      if (t.type === 'Deposit') {
        if (beforeEnd(t.date)) totalDeposits += Math.abs(t.total);
        if (inRange(t.date)) totalDepositsRange += Math.abs(t.total);
        if (t.date >= monthStart) totalDepositsThisMonth += Math.abs(t.total);
      } else if (t.type === 'Withdrawal') {
        if (beforeEnd(t.date)) totalWithdrawals += Math.abs(t.total);
      } else if (t.type === 'Charge') {
        if (beforeEnd(t.date)) totalCharges += Math.abs(t.total);
        if (inRange(t.date)) totalChargesRange += Math.abs(t.total);
      } else if (t.type === 'Dividend') {
        if (beforeEnd(t.date)) totalDividends += Math.abs(t.total);
        if (inRange(t.date)) totalDividendsRange += Math.abs(t.total);
      } else if (t.type === 'Buy') {
        if (!key || !beforeEnd(t.date)) return;
        if (!hMap[key]) hMap[key] = { qty: 0, totalCost: 0, portfolio: t.portfolio };
        hMap[key].qty += t.qty; hMap[key].totalCost += t.total;
      } else if (t.type === 'Sell') {
        if (!key || !beforeEnd(t.date)) return;
        if (!hMap[key]) hMap[key] = { qty: 0, totalCost: 0, portfolio: t.portfolio };
        const h = hMap[key];
        if (h.qty > 0) {
          const avg = h.totalCost / h.qty;
          const pnl = t.total - t.qty * avg;
          if (beforeEnd(t.date)) totalRealizedPnL += pnl;
          if (inRange(t.date)) totalRealizedPnLRange += pnl;
          h.qty = Math.max(0, h.qty - t.qty);
          h.totalCost = Math.max(0, h.totalCost - t.qty * avg);
        }
      }
    });

    const stockCost = Object.values(hMap).reduce((s, h) => s + h.totalCost, 0);
    const cashBalance = totalDeposits - totalWithdrawals - stockCost + totalRealizedPnL + totalDividends - totalCharges;
    const investmentHolding = Object.values(hMap).filter(h => h.portfolio === 'Investment').reduce((s, h) => s + h.totalCost, 0);
    const tradingHolding = Object.values(hMap).filter(h => h.portfolio === 'Trading').reduce((s, h) => s + h.totalCost, 0);
    const totalProfit = totalRealizedPnLRange + totalDividendsRange - totalChargesRange;

    return {
      currentHolding: stockCost + cashBalance,
      stockCost, investmentHolding, tradingHolding, cashBalance,
      totalInvested: totalDeposits, totalInvestedRange: totalDepositsRange, totalWithdrawals,
      totalProfit,
      profitPercentage: stockCost > 0 ? (totalProfit / stockCost) * 100 : 0,
      realizedPnL: totalRealizedPnL,
      dividendReceived: totalDividends,
      activeInvestmentSnapshot: totalDeposits - totalWithdrawals,
      totalDepositsSnapshot: totalDeposits,
      totalDepositsThisMonth,
      totalCharges: totalChargesRange,
    };
  }, [transactions, startStr, endStr]);

  // ── Table date range ──
  const { tableStartStr, tableEndStr, tablePeriodLabel } = useMemo(() => {
    const now = new Date();
    if (tableMode === '1m') {
      const endMonth = now.getMonth() - tableOffset;
      const s = new Date(now.getFullYear(), endMonth, 1);
      const e = new Date(now.getFullYear(), endMonth + 1, 0);
      return {
        tableStartStr: toDateStr(s),
        tableEndStr: tableOffset === 0 ? getTodayStr() : toDateStr(e),
        tablePeriodLabel: s.toLocaleString('default', { month: 'short', year: 'numeric' }).toUpperCase(),
      };
    }
    if (tableMode === '1y') {
      const year = now.getFullYear() - tableOffset;
      return {
        tableStartStr: `${year}-01-01`,
        tableEndStr: tableOffset === 0 ? getTodayStr() : `${year}-12-31`,
        tablePeriodLabel: `${year}`,
      };
    }
    return {
      tableStartStr: tableCustomDates.start || '0000-01-01',
      tableEndStr: tableCustomDates.end || '9999-12-31',
      tablePeriodLabel: 'CUSTOM',
    };
  }, [tableMode, tableOffset, tableCustomDates]);

  // ── Filtered rows ──
  const filteredTransactions = useMemo(() => {
    let r = transactions.filter(t => {
      if (t.date < tableStartStr || t.date > tableEndStr) return false;
      if (selectedPortfolio !== 'Global' && t.portfolio !== selectedPortfolio) return false;
      if (selectedTypes.length > 0 && !selectedTypes.includes(t.type)) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        if (!t.ticker?.toLowerCase().includes(q) && !t.companyName?.toLowerCase().includes(q)) return false;
      }
      return true;
    });
    r.sort((a, b) => {
      let cmp = 0;
      if (sortBy === 'date') cmp = new Date(b.date).getTime() - new Date(a.date).getTime();
      else if (sortBy === 'total') cmp = b.total - a.total;
      else if (sortBy === 'ticker') cmp = (a.ticker || '').localeCompare(b.ticker || '');
      else if (sortBy === 'type') cmp = a.type.localeCompare(b.type);
      return sortOrder === 'desc' ? cmp : -cmp;
    });
    return r;
  }, [transactions, tableStartStr, tableEndStr, selectedPortfolio, selectedTypes, searchQuery, sortBy, sortOrder]);

  const handleSort = (key: typeof sortBy) => {
    if (sortBy === key) setSortOrder(o => o === 'asc' ? 'desc' : 'asc');
    else { setSortBy(key); setSortOrder('desc'); }
  };

  const toggleSelectAll = () => {
    if (selectedIds.length === filteredTransactions.length && filteredTransactions.length > 0) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filteredTransactions.map(t => t.id));
    }
  };

  const toggleSelect = (id: string) => {
    if (selectedIds.includes(id)) {
      setSelectedIds(selectedIds.filter(i => i !== id));
    } else {
      setSelectedIds([...selectedIds, id]);
    }
  };

  const SortIcon = ({ col }: { col: typeof sortBy }) =>
    sortBy === col
      ? (sortOrder === 'asc' ? <ArrowUp size={12} className="text-teal-400" /> : <ArrowDown size={12} className="text-teal-400" />)
      : <ArrowUpDown size={12} className="text-slate-600 group-hover:text-slate-400" />;

  const typeFilterLabel = useMemo(() => {
    if (selectedTypes.length === 0) return 'Type';
    if (selectedTypes.length === 1) return typeLabel(selectedTypes[0]);
    return `${selectedTypes.length} Types`;
  }, [selectedTypes]);

  const syncLabel = useMemo(() => {
    if (isSyncing) return 'Syncing with Google Sheets...';
    if (isError) return 'Could not reach Google Sheets. Showing cached data.';
    if (lastSynced) {
      const mins = Math.floor((Date.now() - lastSynced.getTime()) / 60000);
      if (mins < 1) return 'Synced just now';
      if (mins < 60) return `Synced ${mins}m ago`;
      return `Synced at ${lastSynced.toLocaleTimeString()}`;
    }
    return 'Not yet synced';
  }, [isSyncing, isError, lastSynced]);

  const { analyticsStartStr, analyticsEndStr } = useMemo(() => {
    const now = new Date();
    if (analyticsRange === 'last6m') {
      const start = new Date(now.getFullYear(), now.getMonth() - 5, 1);
      return { analyticsStartStr: toDateStr(start), analyticsEndStr: getTodayStr() };
    }
    if (analyticsRange === 'last12m') {
      const start = new Date(now.getFullYear(), now.getMonth() - 11, 1);
      return { analyticsStartStr: toDateStr(start), analyticsEndStr: getTodayStr() };
    }
    if (analyticsRange === 'fiscal') {
      const currentYear = now.getFullYear();
      const isPostJuly = now.getMonth() >= 6;
      const startYear = isPostJuly ? currentYear : currentYear - 1;
      const start = `${startYear}-07-01`;
      return { analyticsStartStr: start, analyticsEndStr: getTodayStr() };
    }
    // For custom, we want to include the whole month if the start date is within it
    const customStart = analyticsCustomDates.start ? new Date(analyticsCustomDates.start) : new Date(0);
    const startOfCustomMonth = new Date(customStart.getFullYear(), customStart.getMonth(), 1);
    return { 
      analyticsStartStr: toDateStr(startOfCustomMonth), 
      analyticsEndStr: analyticsCustomDates.end || '9999-12-31' 
    };
  }, [analyticsRange, analyticsCustomDates]);

  const { returnHistoryStartStr, returnHistoryEndStr } = useMemo(() => {
    const now = new Date();
    if (returnHistoryRange === 'last6m') {
      const start = new Date(now.getFullYear(), now.getMonth() - 5, 1);
      return { returnHistoryStartStr: toDateStr(start), returnHistoryEndStr: getTodayStr() };
    }
    if (returnHistoryRange === 'last12m') {
      const start = new Date(now.getFullYear(), now.getMonth() - 11, 1);
      return { returnHistoryStartStr: toDateStr(start), returnHistoryEndStr: getTodayStr() };
    }
    if (returnHistoryRange === 'fiscal') {
      const fy = now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
      return { returnHistoryStartStr: `${fy}-07-01`, returnHistoryEndStr: getTodayStr() };
    }
    const customStart = returnHistoryCustomDates.start ? new Date(returnHistoryCustomDates.start) : new Date(0);
    const startOfCustomMonth = new Date(customStart.getFullYear(), customStart.getMonth(), 1);
    return { returnHistoryStartStr: toDateStr(startOfCustomMonth), returnHistoryEndStr: returnHistoryCustomDates.end || '9999-12-31' };
  }, [returnHistoryRange, returnHistoryCustomDates]);

  const chartData = useMemo(() => {
    const relevant = transactions.filter(t => t.date <= analyticsEndStr);
    const groups: Record<string, { label: string; date: Date; deposit: number; trading: number; investment: number }> = {};

    transactions.forEach(t => {
      const d = new Date(t.date);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      const label = d.toLocaleString('default', { month: 'short', year: '2-digit' });

      if (!groups[key]) {
        groups[key] = { label, date: new Date(d.getFullYear(), d.getMonth(), 1), deposit: 0, trading: 0, investment: 0 };
      }

      const amount = t.total || 0;
      if (t.type === 'Deposit') groups[key].deposit += amount;
      else if (t.portfolio === 'Investment' && t.type === 'Buy') groups[key].investment += amount;
      else if (t.portfolio === 'Trading') {
        if (t.type === 'Buy') groups[key].trading += amount;
        else if (t.type === 'Sell') groups[key].trading -= amount;
      }
    });

    const sortedGroups = Object.values(groups).sort((a, b) => a.date.getTime() - b.date.getTime());

    if (analyticsView === 'cumulative') {
      let runDep = 0, runTrad = 0, runInv = 0;
      sortedGroups.forEach(g => {
        runDep += g.deposit;
        runTrad += g.trading;
        runInv += g.investment;
        g.deposit = runDep;
        g.trading = runTrad;
        g.investment = runInv;
      });
    }

    return sortedGroups.filter(g => toDateStr(g.date) >= analyticsStartStr && toDateStr(g.date) <= analyticsEndStr);
  }, [transactions, analyticsStartStr, analyticsEndStr, analyticsView]);
// ── Return History Data ──
// ── Return History Data ──
  const returnHistoryData = useMemo(() => {
    const sortedT = [...transactions].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const holdings: Record<string, { qty: number; cost: number }> = {};
    const monthMap: Record<string, any> = {};

    sortedT.forEach(t => {
      const d = new Date(t.date);
      const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      if (!monthMap[monthKey]) {
        monthMap[monthKey] = {
          monthKey,
          label: d.toLocaleString('en-US', { month: 'short', year: '2-digit' }),
          pnl: 0, dividend: 0, charge: 0,
          date: new Date(d.getFullYear(), d.getMonth(), 1),
        };
      }
      const key = `${t.portfolio}|${t.ticker}`;
      if (t.type === 'Buy') {
        if (!holdings[key]) holdings[key] = { qty: 0, cost: 0 };
        holdings[key].qty += t.qty;
        holdings[key].cost += t.total;
      } else if (t.type === 'Sell') {
        if (holdings[key] && holdings[key].qty > 0) {
          const avg = holdings[key].cost / holdings[key].qty;
          const cost = avg * t.qty;
          monthMap[monthKey].pnl += t.total - cost;
          holdings[key].qty -= t.qty;
          holdings[key].cost -= cost;
        }
      } else if (t.type === 'Dividend') {
        monthMap[monthKey].dividend += t.total;
      } else if (t.type === 'Charge') {
        monthMap[monthKey].charge += t.total;
      }
    });

    let allMonths = Object.values(monthMap).sort((a: any, b: any) => a.monthKey.localeCompare(b.monthKey));

    if (returnHistoryView === 'cumulative') {
      let rPnl = 0, rDiv = 0, rCharge = 0;
      allMonths = allMonths.map((m: any) => {
        rPnl += m.pnl; rDiv += m.dividend; rCharge += m.charge;
        return { ...m, pnl: rPnl, dividend: rDiv, charge: rCharge,
          netReturn: rPnl + rDiv - Math.abs(rCharge),
          labelY: Math.max(0, rPnl) + rDiv };
      });
    } else {
      allMonths = allMonths.map((m: any) => ({
        ...m, netReturn: m.pnl + m.dividend - Math.abs(m.charge),
        labelY: Math.max(0, m.pnl) + m.dividend,
      }));
    }

    let data = allMonths.filter((m: any) =>
      toDateStr(m.date) >= returnHistoryStartStr && toDateStr(m.date) <= returnHistoryEndStr
    ).map((m: any) => {
      const pnl = m.pnl || 0;
      const dividend = m.dividend || 0;
      const charge = m.charge || 0;
      
      const showProfit = !hiddenSeries.has('Realized Profit');
      const showLoss = !hiddenSeries.has('Realized Loss');
      const showDividend = !hiddenSeries.has('Dividend');

      const displayPnl = (pnl >= 0) ? (showProfit ? pnl : 0) : (showLoss ? pnl : 0);
      const displayDiv = showDividend ? dividend : 0;
      
      return {
        ...m,
        pnl: displayPnl,
        dividend: displayDiv,
        netReturn: displayPnl + displayDiv - Math.abs(charge),
        labelY: Math.max(0, displayPnl) + displayDiv
      };
    });
    const firstIdx = data.findIndex((m: any) => m.pnl !== 0 || m.dividend !== 0);
    if (firstIdx > 0) data = data.slice(firstIdx);
    return data;
  }, [transactions, returnHistoryView, returnHistoryStartStr, returnHistoryEndStr, hiddenSeries]);

  // ── Return by Stock Data ──
  const returnByStockData = useMemo(() => {
    const holdingsMap: Record<string, { qty: number; cost: number }> = {};
    const stockMap: Record<string, { ticker: string; realizedPnL: number; dividends: number }> = {};

    [...transactions]
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
      .forEach(t => {
        if (!t.ticker) return;
        const key = `${t.portfolio}|${t.ticker}`;
        if (!stockMap[t.ticker]) stockMap[t.ticker] = { ticker: t.ticker, realizedPnL: 0, dividends: 0 };
        if (t.type === 'Buy') {
          if (!holdingsMap[key]) holdingsMap[key] = { qty: 0, cost: 0 };
          holdingsMap[key].qty += t.qty;
          holdingsMap[key].cost += t.total;
        } else if (t.type === 'Sell') {
          if (holdingsMap[key] && holdingsMap[key].qty > 0) {
            const avg = holdingsMap[key].cost / holdingsMap[key].qty;
            const cost = avg * t.qty;
            stockMap[t.ticker].realizedPnL += t.total - cost;
            holdingsMap[key].qty -= t.qty;
            holdingsMap[key].cost -= cost;
          }
        } else if (t.type === 'Dividend') {
          stockMap[t.ticker].dividends += t.total;
        }
      });

    return Object.values(stockMap).sort((a, b) => b.realizedPnL - a.realizedPnL);
  }, [transactions]);
  const yAxisTicksReturn = useMemo(() => {
    if (returnHistoryData.length === 0) return [0];
    const values = returnHistoryData.flatMap((d: any) => [d.pnl, d.dividend, d.netReturn]);
    const min = Math.min(...values, 0);
    const max = Math.max(...values, 2000);
    const interval = 2000;
    const ticks = [];
    for (let i = Math.floor(min / interval) * interval; i <= Math.ceil(max / interval) * interval; i += interval) ticks.push(i);
    return ticks;
  }, [returnHistoryData]);
  const maxVal = Math.max(...chartData.map(g => Math.max(g.deposit, (g.investment + (g.trading > 0 ? g.trading : 0))), 100000));
  const yTicks = useMemo(() => {
    const ticks = [];
    for (let i = 0; i <= maxVal + 50000; i += 50000) ticks.push(i);
    return ticks;
  }, [maxVal]);

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className="space-y-8">
      {/* Transaction Modal */}
      <TransactionModal
  isOpen={isTransactionModalOpen}
  editingTransaction={editingTransaction}
  onClose={() => setIsTransactionModalOpen(false)}
  commissionRate={dseSettings.commissionRate}
  customStocks={dseSettings.customStocks}
  existingStocks={[
    ...new Map(
      transactions
        .filter(t => t.ticker)
        .map(t => [
          t.ticker.toUpperCase(),
          {
            ticker: t.ticker.toUpperCase(),
            name: t.companyName || t.ticker
          }
        ])
    ).values()
  ]}
        onSave={async (t, keepOpen?: boolean) => {
          if (editingTransaction) {
            setTransactions(prev => {
              const updated = prev.map(tx => tx.id === t.id ? t : tx);
              localStorage.setItem('sheet_cache_dse', JSON.stringify(updated));
              return updated;
            });
            setIsTransactionModalOpen(false);
          } else {
            setTransactions(prev => {
              const updated = [...prev, t];
              localStorage.setItem('sheet_cache_dse', JSON.stringify(updated));
              return updated;
            });
            if (!keepOpen) setIsTransactionModalOpen(false);
          }
          await pushToSheets('update', t);
        }}
      />

      {/* Confirm dialog */}
      <ConfirmDialog
        isOpen={confirmDialog.isOpen}
        message={confirmDialog.message}
        onConfirm={confirmDialog.onConfirm}
        onCancel={() => setConfirmDialog(prev => ({ ...prev, isOpen: false }))}
      />

      {/* API status toast */}
      {apiStatus && (
        <div className={cn(
          "flex items-center gap-3 px-4 py-3 rounded-xl border text-[10px] font-bold uppercase animate-in fade-in slide-in-from-top-4 duration-300",
          apiStatus.isError
            ? "bg-rose-500/10 border-rose-500/20 text-rose-400"
            : "bg-teal-400/10 border-teal-400/20 text-teal-400"
        )}>
          {apiStatus.isError ? <AlertTriangle size={14} /> : <RefreshCw size={14} />}
          {apiStatus.message}
        </div>
      )}

      {/* ══════════════════════════════════════════════════
          SUMMARY TAB
      ══════════════════════════════════════════════════ */}
      {activeTab === 'summary' && (
        <div className="space-y-8">
          {/* ── Stats Date Range Selector ── */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 relative">
            <div className="flex flex-col sm:flex-row sm:items-center gap-3 w-full sm:w-auto">
              <div className="flex items-center justify-between w-full sm:w-auto gap-2">
                <div className="relative flex-1 sm:flex-none">
                  {/* Mobile */}
                  <div className="block sm:hidden">
                    <button onClick={() => setIsRangeMenuOpen(!isRangeMenuOpen)}
                      className="flex items-center justify-between gap-4 bg-slate-950 border border-slate-800 rounded-lg px-4 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase w-full">
                      <div className="flex items-center gap-2">
                        <Calendar size={14} className="text-teal-400" />
                        {historyRange === 'all' ? 'Overall' : historyRange === 'last12m' ? 'Last 12M' : historyRange === 'fiscal' ? 'Fiscal' : 'Custom'}
                      </div>
                      <ChevronDown size={14} className={cn("text-slate-500 transition-transform", isRangeMenuOpen ? "rotate-180 text-teal-400" : "")} />
                    </button>
                    {isRangeMenuOpen && (
                      <>
                        <div className="fixed inset-0 z-40" onClick={() => setIsRangeMenuOpen(false)} />
                        <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
                          {['all', 'last12m', 'fiscal', 'custom'].map(id => (
                            <button key={id} onClick={() => handleRangeChange(id as any)}
                              className={cn("w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase",
                                historyRange === id ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800")}>
                              {id === 'all' ? 'Overall' : id === 'last12m' ? 'Last 12M' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                  {/* Desktop */}
                  <div className="hidden sm:flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1">
                    {['all', 'last12m', 'fiscal', 'custom'].map(id => (
                      <button key={id} onClick={() => handleRangeChange(id as any)}
                        className={cn("px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all border",
                          historyRange === id ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20" : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white")}>
                        {id === 'all' ? 'Overall' : id === 'last12m' ? 'Last 12M' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="block sm:hidden">
                  <SettingsButton 
                    isOpen={isSettingsMenuOpen} 
                    setIsOpen={setIsSettingsMenuOpen} 
                    onExport={exportData}
                    onImport={importData}
                    onTemplate={generateTemplate}
                    onAdd={() => openTransactionModal()}
                    newTicker={newTicker}
                    setNewTicker={setNewTicker}
                    newCompany={newCompany}
                    setNewCompany={setNewCompany}
                    onAddStock={onAddStock}
                  />
                </div>
              </div>
              {historyRange === 'custom' && (
                <div className="flex items-center justify-center sm:justify-start gap-1.5 px-1 animate-in fade-in slide-in-from-top-2 duration-300 w-full sm:w-auto">
                  <button onClick={() => navigateCustomMonth(-1)} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400"><ChevronLeft size={14} /></button>
                  <input type="date" value={historyCustomDates.start} onChange={e => setHistoryCustomDates(p => ({ ...p, start: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none focus:border-teal-400/50 flex-1 sm:flex-none" />
                  <span className="text-slate-700 text-[10px] font-bold">–</span>
                  <input type="date" value={historyCustomDates.end} onChange={e => setHistoryCustomDates(p => ({ ...p, end: e.target.value }))} className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-bold text-white outline-none focus:border-teal-400/50 flex-1 sm:flex-none" />
                  <button onClick={() => navigateCustomMonth(1)} className="flex items-center justify-center w-8 h-8 rounded-md bg-slate-900 border border-slate-800 text-slate-400"><ChevronRight size={14} /></button>
                </div>
              )}
            </div>
            <div className="hidden sm:block">
              <SettingsButton 
                isOpen={isSettingsMenuOpen} 
                setIsOpen={setIsSettingsMenuOpen} 
                onExport={exportData}
                onImport={importData}
                onTemplate={generateTemplate}
                onAdd={() => openTransactionModal()}
                newTicker={newTicker}
                setNewTicker={setNewTicker}
                newCompany={newCompany}
                setNewCompany={setNewCompany}
                onAddStock={onAddStock}
              />
            </div>
          </div>

          {/* ── Sync Status Bar ── */}
          <div className={cn(
            "flex items-center justify-between gap-3 px-4 py-2 rounded-xl border text-[10px] font-bold uppercase transition-colors",
            isSyncing ? "bg-slate-900/50 border-slate-800 text-slate-400" :
            isError ? "bg-rose-500/10 border-rose-500/20 text-rose-400" :
            "bg-teal-400/5 border-teal-400/10 text-teal-500"
          )}>
            <div className="flex items-center gap-2">
              <RefreshCw size={12} className={cn(isSyncing ? "animate-spin" : "opacity-50")} />
              <span>{syncLabel}</span>
            </div>
            <button onClick={() => fetchFromSheets(true)} disabled={isSyncing}
              className={cn("flex items-center gap-1.5 px-3 py-1 rounded-lg border transition-all",
                isSyncing ? "border-slate-700 text-slate-600 cursor-not-allowed" :
                isError ? "border-rose-500/30 text-rose-400 hover:bg-rose-500/10" :
                "border-teal-400/30 text-teal-400 hover:bg-teal-400/10")}>
              <RefreshCw size={10} className={isSyncing ? "animate-spin" : ""} />SYNC NOW
            </button>
          </div>

          {/* ── 4 Main Summary Cards ── */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 items-stretch">
            <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)] group">
              <div className="flex-1">
                <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-teal-400/10 flex items-center justify-center mb-4 text-teal-400 group-hover:scale-110 transition-transform"><DollarSign size={20} /></div>
                <div className="mb-2">
                  <div className="flex items-center gap-1"><p className="text-body-sm font-bold text-white uppercase tracking-wider">Current Holding</p><InfoTooltip text="Buying cost for Stocks + Cash Balance" /></div>
                  <p className="text-label font-bold text-slate-500 uppercase">Cash + Stocks</p>
                </div>
                <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">{formatBDT(stats.currentHolding)}</h3>
              </div>
              <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
                <p className="text-label font-medium text-slate-300 uppercase">Cash Balance</p>
                <p className="text-body font-bold text-teal-400 tabular-nums">{formatBDT(stats.cashBalance)}</p>
              </div>
            </Card>

            <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-blue-400/50 hover:shadow-[0_0_20px_rgba(37,99,235,0.1)] group">
              <div className="flex-1">
                <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-blue-400/10 flex items-center justify-center mb-4 text-blue-400 group-hover:scale-110 transition-transform"><Wallet size={20} /></div>
                <div className="mb-2">
                  <div className="flex items-center gap-1"><p className="text-body-sm font-bold text-white uppercase tracking-wider">Active Investment</p><InfoTooltip text="Total Deposit - Total Withdrawal" /></div>
                  <p className="text-label font-bold text-slate-500 uppercase">Deposit - Withdrawal</p>
                </div>
                <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">{formatBDT(stats.activeInvestmentSnapshot)}</h3>
              </div>
              <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
                <p className="text-label font-medium text-slate-300 uppercase">Deposit</p>
                <p className="text-body font-bold text-blue-400 tabular-nums">{formatBDT(stats.totalDepositsSnapshot)}</p>
              </div>
            </Card>

            <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-purple-400/50 hover:shadow-[0_0_20px_rgba(167,139,250,0.1)] group">
              <div className="flex-1">
                <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-purple-400/10 flex items-center justify-center mb-4 text-purple-400 group-hover:scale-110 transition-transform"><Briefcase size={20} /></div>
                <div className="mb-2">
                  <div className="flex items-center gap-1"><p className="text-body-sm font-bold text-white uppercase tracking-wider">Total Invested</p><InfoTooltip text="Total Deposit during the selected period" /></div>
                  <p className="text-label font-bold text-slate-500 uppercase">Total Deposit</p>
                </div>
                <h3 className="text-heading font-bold text-white mb-4 tracking-tight font-display tabular-nums">{formatBDT(stats.totalInvestedRange)}</h3>
              </div>
              <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
                <p className="text-label font-medium text-slate-300 uppercase">This Month</p>
                <p className="text-body font-bold text-purple-400 tabular-nums">{formatBDT(stats.totalDepositsThisMonth)}</p>
              </div>
            </Card>

            <Card className="bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all hover:border-emerald-400/50 hover:shadow-[0_0_20px_rgba(52,211,153,0.1)] group">
              <div className="flex-1">
                <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl bg-emerald-400/10 flex items-center justify-center mb-4 text-emerald-400 group-hover:scale-110 transition-transform"><TrendingUp size={20} /></div>
                <div className="mb-2">
                  <div className="flex items-center gap-1"><p className="text-body-sm font-bold text-white uppercase tracking-wider">Total Profit</p><InfoTooltip text="Dividend + Realized P&L from trading" /></div>
                  <p className="text-label font-bold text-slate-500 uppercase">Dividend + P&L</p>
                </div>
                <h3 className={cn("text-heading font-bold mb-4 tracking-tight font-display tabular-nums", stats.totalProfit >= 0 ? "text-emerald-400" : "text-rose-500")}>{formatBDT(stats.totalProfit)}</h3>
              </div>
              <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
                <p className="text-label font-medium text-slate-300 uppercase">Portfolio Yield</p>
                <p className={cn("text-body font-bold tabular-nums", stats.profitPercentage >= 0 ? "text-emerald-400" : "text-rose-500")}>{stats.profitPercentage.toFixed(2)}%</p>
              </div>
            </Card>
          </div>

          {/* ── 6 Secondary Cards ── */}
          <div className="grid grid-cols-2 lg:grid-cols-3 gap-4">
            {[
              { label: 'Investment', tooltip: 'Buying cost for investment purpose', value: stats.investmentHolding, color: 'text-blue-400' },
              { label: 'Trading', tooltip: 'Buying cost for trading purpose', value: stats.tradingHolding, color: 'text-amber-400' },
              { label: 'Stocks', tooltip: 'Buying cost for Stocks (Investment + Trading)', value: stats.investmentHolding + stats.tradingHolding, color: 'text-white' },
              { label: 'Dividend', tooltip: 'Total dividend received', value: stats.dividendReceived, color: 'text-emerald-400' },
              { label: 'Realized P&L', tooltip: 'Realized Profit or Loss from trading', value: stats.realizedPnL, color: stats.realizedPnL >= 0 ? 'text-emerald-400' : 'text-rose-500' },
              { label: 'Withdrawal', tooltip: 'Money withdrawn from BO account', value: stats.totalWithdrawals, color: 'text-rose-400' },
            ].map(({ label, tooltip, value, color }) => (
              <Card key={label} className="bg-slate-900 border-slate-800 transition-all hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)]" style={{ padding: '12px' }}>
                <div className="flex items-center justify-between mb-1">
                  <p className="text-label font-bold text-slate-300 uppercase">{label}</p>
                  <InfoTooltip text={tooltip} />
                </div>
                <p className={cn("text-body font-bold tabular-nums", color)}>{formatBDT(value)}</p>
              </Card>
            ))}
          </div>

          {/* ── Transaction Summary Bar Chart ── */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 lg:p-8 shadow-xl">
            <div className="flex items-center gap-3 mb-8 pb-4 border-b border-slate-800">
              <BarChart2 className="text-teal-400" size={20} />
              <h3 className="text-body font-bold text-white uppercase tracking-widest">Transaction Summary</h3>
            </div>

            {(() => {
              const chartData = [
                { label: 'Balance',    value: stats.cashBalance,         color: '#2dd4bf' },
                { label: 'Return',     value: stats.totalProfit,         color: stats.totalProfit >= 0 ? '#10b981' : '#f43f5e' },
                { label: 'Trading',    value: stats.tradingHolding,      color: '#f59e0b' },
                { label: 'Investment', value: stats.investmentHolding,   color: '#3b82f6' },
                { label: 'Holding',    value: stats.stockCost,           color: '#a855f7' },
                { label: 'Deposit',    value: stats.totalInvestedRange,  color: '#06b6d4' }
              ];
              // Buffer maxVal by 15% to prevent bars/labels from overflowing
              const rawMax = Math.max(...chartData.map(d => Math.abs(d.value)), 1);
              const maxVal = rawMax * 1.15;

const fmtShort = (n: number) => {
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 10000000) return `${sign}${(abs / 10000000).toFixed(1)}Cr`;
  if (abs >= 100000)   return `${sign}${(abs / 100000).toFixed(1)}L`;
  if (abs >= 1000)     return `${sign}${(abs / 1000).toFixed(1)}K`;
  return `${sign}${abs.toFixed(0)}`;
};

              return (
                <div className="space-y-6">
  {/* Vertical grid + bars */}
  <div className="relative">
    {/* Vertical grid lines */}
<div
  className="absolute pointer-events-none"
  style={{
    left: 92,
    right: 0,
    top: -10,     // extend 10px above new top line
    bottom: 0,
  }}
>
  {(() => {
    const step = 200000; // Fixed 2L interval
    const count = Math.ceil(maxVal / step);

    return Array.from({ length: count + 1 }, (_, i) => {
      const pct = ((step * i) / maxVal) * 100;
      if (pct > 100) return null;

      return (
        <div
          key={i}
          className="absolute border-l border-slate-700/40"
          style={{
            left: `${pct}%`,
            top: 0,
            bottom: 0,
          }}
        />
      );
    });
  })()}

  {/* NEW TOP HORIZONTAL GRID */}
  <div
    className="absolute left-[-20px] right-0 border-t border-slate-700/40"
    style={{
      top: 22, // same spacing above Balance bar as bottom line below Deposit
    }}
  />
</div>

    <div className="space-y-3 pt-8">
      {chartData.map((d) => (
        <SummaryBarItem 
          key={d.label}
          label={d.label}
          value={d.value}
          color={d.color}
          barPct={(Math.abs(d.value) / maxVal) * 100}
          fmtShort={fmtShort}
        />
      ))}
    </div>

    {/* Axis labels below grid */}
    <div className="relative mt-5 h-5 px-1" style={{ marginLeft: 72, marginRight: 0 }}>
      {/* Horizontal Line - extending 20px before 0 and to the right */}
      <div className="absolute top-0 left-0 right-0 border-t border-slate-800" />
      
      {/* Labels - positioned relative to the 0 line at 92px (72 + 20) */}
      <div className="absolute top-0 left-[20px] right-0 h-full">
        {(() => {
          const step = 200000; // Fixed 2L interval
          const count = Math.ceil(maxVal / step);
          return Array.from({ length: count + 1 }, (_, i) => {
            const val = step * i;
            const pct = (val / maxVal) * 100;
            if (pct > 104) return null;
            const label = val >= 10000000 ? `${(val/10000000).toFixed(val%10000000===0?0:1)}Cr` : `${(val/100000).toFixed(0)}L`;
            return (
              <span key={i} className="absolute top-6 text-[9px] font-bold text-slate-500 -translate-x-1/2 whitespace-nowrap"
                style={{ left: `${pct}%` }}>{label}</span>
            );
          });
        })()}
      </div>
    </div>
  </div>

  {/* Legends */}
  <div className="mt-10 pt-10 border-t border-slate-800/50 flex flex-wrap justify-center gap-x-6 gap-y-3">
    {chartData.map((d) => (
      <div key={d.label} className="flex items-center gap-2">
        <div className="w-3 h-3 rounded-sm shrink-0" style={{ backgroundColor: d.color }} />
        <span className="text-[12px] font-bold text-slate-300 uppercase tracking-tight whitespace-nowrap">{d.label}</span>
      </div>
    ))}
  </div>
</div>
              );
            })()}
          </div>
        </div>
      )}

      {/* ══════════════════════════════════════════════════
          TRANSACTIONS TAB
      ══════════════════════════════════════════════════ */}
      {activeTab === 'transactions' && (
        <div className="space-y-4">
          {/* Table header row */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-3">
            <div className="flex items-center gap-2 sm:gap-3">
              <Briefcase className="text-teal-400 w-5 h-5 lg:w-6 lg:h-6" />
              <h2 className="text-body-sm sm:text-subheading lg:text-heading font-bold text-white font-display uppercase whitespace-nowrap">DSE Transactions</h2>
            </div>

            <div className="flex items-center gap-2 flex-1 sm:flex-none justify-end">
              <div className="relative flex-1 sm:flex-none sm:w-56">
                <div className="flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-lg p-1 px-2 h-9 hover:border-slate-700 focus-within:border-teal-400/50 transition-colors group cursor-text"
                  onClick={() => document.getElementById('dse-search-input')?.focus()}>
                  <Search size={14} className="text-slate-500 group-focus-within:text-teal-400 transition-colors shrink-0" />
                  <input id="dse-search-input" type="text" placeholder="Search Ticker/Company"
                    className="flex-1 bg-transparent border-none focus:ring-0 text-[10px] font-bold text-white placeholder:text-slate-600 uppercase outline-none min-w-0"
                    value={searchQuery} onChange={e => setSearchQuery(e.target.value)} />
                </div>
              </div>

              <div className="relative shrink-0">
                <button
                  onClick={() => { setIsPortfolioMenuOpen(!isPortfolioMenuOpen); setIsTypeMenuOpen(false); }}
                  className={cn(
                    "flex items-center gap-2 px-3 h-9 rounded-lg transition-all text-[10px] font-bold uppercase",
                    PORTFOLIO_BTN[selectedPortfolio]
                  )}
                >
                  <Filter size={14} />
                  <span className="hidden sm:inline">{portfolioLabel(selectedPortfolio)}</span>
                </button>
                {isPortfolioMenuOpen && (
                  <>
                    <div className="fixed inset-0 z-40" onClick={() => setIsPortfolioMenuOpen(false)} />
                    <div className="absolute right-0 mt-2 w-44 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-2 animate-in fade-in zoom-in-95">
                      {(['Global', 'Investment', 'Trading'] as const).map(p => (
                        <button key={p} onClick={() => { setSelectedPortfolio(p); setIsPortfolioMenuOpen(false); }}
                          className={cn(
                            "w-full flex items-center gap-2.5 px-3 py-1.5 rounded-lg text-[10px] font-bold uppercase transition-colors",
                            selectedPortfolio === p ? "bg-teal-400/10 text-teal-400" : "text-slate-300 hover:bg-slate-800 hover:text-white"
                          )}>
                          <span className={cn("w-2 h-2 rounded-full shrink-0", PORTFOLIO_DOT[p])} />
                          {portfolioLabel(p)}
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>

              <div className="relative shrink-0">
                <button
                  onClick={() => { setIsTypeMenuOpen(!isTypeMenuOpen); setIsPortfolioMenuOpen(false); }}
                  className={cn(
                    "flex items-center gap-2 px-3 h-9 rounded-lg transition-all text-[10px] font-bold uppercase",
                    selectedTypes.length > 0
                      ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20"
                      : "bg-slate-900 border border-slate-800 text-slate-400 hover:border-slate-700"
                  )}
                >
                  <Filter size={14} />
                  <span className="hidden sm:inline">{typeFilterLabel}</span>
                </button>
                {isTypeMenuOpen && (
                  <>
                    <div className="fixed inset-0 z-40" onClick={() => setIsTypeMenuOpen(false)} />
                    <div className="absolute right-0 mt-2 w-44 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-2 animate-in fade-in zoom-in-95">
                      {(['Buy', 'Sell', 'Deposit', 'Withdrawal', 'Charge', 'Dividend'] as const).map(type => (
                        <div key={type}
                          className="flex items-center gap-2 px-3 py-1 rounded hover:bg-slate-800 transition-colors cursor-pointer"
                          onClick={() => setSelectedTypes(prev => prev.includes(type) ? prev.filter(t => t !== type) : [...prev, type])}>
                          <Checkbox label={typeLabel(type)} checked={selectedTypes.includes(type)} onChange={() => {}} />
                        </div>
                      ))}
                      {selectedTypes.length > 0 && (
                        <button onClick={() => setSelectedTypes([])}
                          className="w-full mt-2 pt-2 border-t border-slate-800 text-center text-[10px] font-bold text-rose-400 hover:text-rose-300 uppercase transition-colors">
                          Clear All
                        </button>
                      )}
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>

          {/* Table Time Filter Bar */}
          <div className="flex items-center justify-between gap-3 bg-slate-900/50 border border-slate-800 rounded-xl p-2 mb-3">
            <div className="relative flex-1 sm:hidden">
              <button onClick={() => setIsTableTimeMenuOpen(!isTableTimeMenuOpen)}
                className="flex items-center justify-between gap-2 bg-slate-950 border border-slate-800 rounded-lg px-3 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase w-full">
                <div className="flex items-center gap-1.5">
                  <Calendar size={14} className="text-teal-400 shrink-0" />
                  <span>{tableMode === '1m' ? 'Last 1M' : tableMode === '1y' ? 'Last 1Y' : 'Custom'}</span>
                </div>
                <ChevronDown size={14} className={cn("text-slate-500 transition-transform", isTableTimeMenuOpen ? "rotate-180 text-teal-400" : "")} />
              </button>
              {isTableTimeMenuOpen && (
                <>
                  <div className="fixed inset-0 z-40" onClick={() => setIsTableTimeMenuOpen(false)} />
                  <div className="absolute left-0 mt-2 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
                    {[['1m', 'Last 1M'], ['1y', 'Last 1Y'], ['custom', 'Custom']].map(([id, label]) => (
                      <button key={id} onClick={() => { setTableMode(id as any); setTableOffset(0); setIsTableTimeMenuOpen(false); }}
                        className={cn("w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase",
                          tableMode === id ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800")}>
                        {label}
                      </button>
                    ))}
                  </div>
                </>
              )}
            </div>

            <div className="hidden sm:flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1">
              {[['1m', 'Last 1M'], ['1y', 'Last 1Y'], ['custom', 'Custom']].map(([id, label]) => (
                <button key={id} onClick={() => { setTableMode(id as any); setTableOffset(0); }}
                  className={cn("px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all whitespace-nowrap border",
                    tableMode === id ? "bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20 border-teal-400/30" : "bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white hover:bg-slate-800/60")}>
                  {label}
                </button>
              ))}
            </div>

            {(tableMode === '1m' || tableMode === '1y') && (
              <div className="flex items-center gap-1 bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 flex-1 sm:flex-none justify-between sm:justify-center max-w-[200px] sm:max-w-none">
                <button onClick={() => setTableOffset(p => p + 1)}
                  className="p-1 px-1.5 hover:bg-slate-800 rounded-md text-slate-400 hover:text-teal-400 transition-all shrink-0">
                  <ChevronLeft size={16} />
                </button>
                <span className="text-[9px] sm:text-[10px] font-bold text-white uppercase tracking-wider whitespace-nowrap px-1">
                  {tablePeriodLabel}
                </span>
                <button onClick={() => setTableOffset(p => Math.max(0, p - 1))} disabled={tableOffset === 0}
                  className={cn("p-1.5 rounded-md transition-all", tableOffset === 0 ? "text-slate-700 cursor-not-allowed" : "hover:bg-slate-800 text-slate-400 hover:text-teal-400")}>
                  <ChevronRight size={16} />
                </button>
              </div>
            )}

            {tableMode === 'custom' && (
              <div className="flex items-center gap-2 animate-in fade-in slide-in-from-right-2 duration-300">
                <input type="date" value={tableCustomDates.start} onChange={e => setTableCustomDates(p => ({ ...p, start: e.target.value }))}
                  className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1 text-[10px] font-bold text-white outline-none focus:border-teal-400/50" />
                <span className="text-slate-600 text-[10px] font-bold uppercase">To</span>
                <input type="date" value={tableCustomDates.end} onChange={e => setTableCustomDates(p => ({ ...p, end: e.target.value }))}
                  className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1 text-[10px] font-bold text-white outline-none focus:border-teal-400/50" />
              </div>
            )}
          </div>

          {/* ── The Table ── */}
          <div className="flex flex-col border border-slate-800 rounded-xl overflow-hidden bg-slate-900/50 shadow-2xl shadow-slate-950/50">
            <div className="overflow-x-auto custom-scrollbar">
              <div className="min-w-[800px] grid grid-cols-[48px_minmax(0,0.9fr)_minmax(0,1.2fr)_minmax(0,1.5fr)_minmax(0,1.5fr)_minmax(0,1fr)_80px]">

                {/* Header */}
                <div className="flex items-center justify-center border-r border-b border-slate-800 bg-slate-900/80 py-3 sticky top-0 z-10 backdrop-blur-sm">
                  <Checkbox
                    checked={selectedIds.length === filteredTransactions.length && filteredTransactions.length > 0}
                    onChange={toggleSelectAll}
                  />
                </div>
                {[
                  { label: 'Date',       col: 'date'   as const },
                  { label: 'Type',       col: 'type'   as const },
                  { label: 'Stock',      col: 'ticker' as const },
                  { label: 'Qty & Price', col: null },
                  { label: 'Total',      col: 'total'  as const },
                ].map(({ label, col }) => (
                  <div key={label}
                    onClick={col ? () => handleSort(col) : undefined}
                    className={cn(
                      "flex items-center gap-2 px-4 py-3 border-r border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm",
                      col ? "cursor-pointer group" : ""
                    )}>
                    {col ? <SortIcon col={col} /> : null}
                    <span className={cn("text-label font-bold uppercase tracking-wider transition-colors",
                      col && sortBy === col ? "text-teal-400" : "text-slate-500 group-hover:text-slate-300")}>
                      {label}
                    </span>
                  </div>
                ))}
                <div className="border-b border-slate-800 bg-slate-900/80 sticky top-0 z-10 backdrop-blur-sm" />

                {selectedIds.length > 0 && (
                  <div className="col-span-full flex items-center justify-between bg-rose-500/10 border-b border-rose-500/20 px-4 py-2 animate-in fade-in slide-in-from-top-2">
                    <span className="text-label font-bold text-rose-500 uppercase">
                      {selectedIds.length} Transaction{selectedIds.length > 1 ? 's' : ''} Selected
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

                {/* Data rows */}
                {filteredTransactions.map((t, i) => {
                  const isSelected = selectedIds.includes(t.id);
                  const cellBase = cn(
                    "px-4 py-3 border-r border-b border-slate-800/50 transition-colors flex flex-col justify-center",
                    isSelected ? "bg-teal-400/10" : "bg-slate-900/40 hover:bg-slate-800/40"
                  );
                  const qtyPriceColor = TYPE_TEXT[t.type] || 'text-white';
                  const hasQtyPrice = t.qty > 0 || t.price > 0;

                  return (
                    <React.Fragment key={t.id || i}>
                      <div className={cn("flex items-center justify-center border-r border-b border-slate-800/50 transition-colors", isSelected ? "bg-teal-400/10" : "bg-slate-900/40 hover:bg-slate-800/40")}>
                        <Checkbox
                          checked={isSelected}
                          onChange={() => toggleSelect(t.id)}
                        />
                      </div>
                      <div className={cellBase}>
                        <span className="text-body-sm font-normal text-white tabular-nums">{formatDDMMYYYY(t.date)}</span>
                      </div>

                      <div className={cellBase}>
                        <span className={cn("inline-flex items-center px-2 py-0.5 rounded text-[9px] font-bold border uppercase w-fit mb-1", TYPE_BG[t.type] || 'bg-slate-400/15 border-slate-400/25 text-white')}>
                          {typeLabel(t.type)}
                        </span>
                        <span className={cn("inline-flex items-center px-2 py-0.5 rounded text-[9px] font-bold border uppercase w-fit", PORTFOLIO_BG[t.portfolio] || 'bg-slate-400/15 border-slate-400/25 text-white')}>
                          {portfolioLabel(t.portfolio)}
                        </span>
                      </div>

                      <div className={cellBase}>
                        {t.ticker ? (
                          <>
                            <span className="text-body-sm font-bold text-white uppercase">{t.ticker}</span>
                            {t.companyName && <span className="text-[10px] text-slate-500 font-medium truncate">{t.companyName}</span>}
                          </>
                        ) : (
                          <span className="text-[10px] text-slate-500 italic">—</span>
                        )}
                      </div>

                      <div className={cellBase}>
                        {hasQtyPrice ? (
                          <>
                            <span className={cn("text-body-sm font-bold tabular-nums", qtyPriceColor)}>
                              {t.qty > 0 ? t.qty.toLocaleString() : '—'} × {t.price > 0
                                ? t.price.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })
                                : '—'}
                            </span>
                            {t.commission > 0 && (
                              <span className="text-[10px] font-bold text-slate-500 tabular-nums">
                                Comm {t.commission.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </span>
                            )}
                          </>
                        ) : (
                          <span className="text-[10px] text-slate-500 italic">—</span>
                        )}
                      </div>

                      <div className={cellBase}>
                        <span className={cn("text-body-sm font-bold tabular-nums", TYPE_TEXT[t.type] || 'text-white')}>
                          {formatBDT(Math.abs(t.total))}
                        </span>
                      </div>

                      <div className={cn("flex items-center gap-1 px-3 py-3 border-b border-slate-800/50 transition-colors", isSelected ? "bg-teal-400/10" : "bg-slate-900/40 hover:bg-slate-800/40")}>
                        <button onClick={() => handleEditTransaction(t)} className="p-1.5 text-slate-500 hover:text-teal-400 hover:bg-teal-400/10 rounded-lg transition-all" title="Edit">
                          <Edit2 size={14} />
                        </button>
                        <button onClick={() => handleDeleteTransaction(t)} className="p-1.5 text-slate-500 hover:text-rose-500 hover:bg-rose-500/10 rounded-lg transition-all" title="Delete">
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </React.Fragment>
                  );
                })}
              </div>
            </div>

            {filteredTransactions.length === 0 && (
              <div className="py-12 text-center bg-slate-900/30 border-t border-dashed border-slate-800 rounded-b-xl">
                <p className="text-body text-slate-500 italic">No transactions found for this period</p>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ══════════════════════════════════════════════════
          HOLDINGS TAB
      ══════════════════════════════════════════════════ */}
      {activeTab === 'holdings' && (
        <PortfolioAllocation holdings={holdings} transactions={transactions} />
      )}

      {/* ══════════════════════════════════════════════════
          ANALYTICS TAB
      ══════════════════════════════════════════════════ */}
      {activeTab === 'analytics' && (
        <div className="space-y-6">
          
          
          {/* ── Transaction History Chart ── */}
          {(() => {
            const toggleSeries = (key: string) => setHiddenSeries(prev => {
              const next = new Set(prev);
              if (next.has(key)) next.delete(key); else next.add(key);
              return next;
            });

            const TEAL = '#2dd4bf';
            const AMBER = '#f59e0b';
            const BLUE = '#3b82f6';

            // For negative trading: we split into positive/negative stacks
            // so the bar always shows amber color regardless of sign.
            // We use a "tradingPos" and "tradingNeg" approach with a single stackId trick:
            // Instead, we use a Cell-based fill on a single bar with no stackId issues.
            // Simplest correct approach: use two separate Bar entries for positive/negative trading.

            const ANALYTICS_LEGENDS = [
              { key: 'deposit',     label: 'Deposit',        color: TEAL  },
              { key: 'trading',     label: 'Net Trading',    color: AMBER },
              { key: 'investment',  label: 'Investment Buy', color: BLUE  },
            ];

            // Generate Y-axis ticks at every 50k
            const allVals = chartData.flatMap(g => [
              hiddenSeries.has('deposit')    ? 0 : g.deposit,
              hiddenSeries.has('trading')    ? 0 : g.trading,
              hiddenSeries.has('investment') ? 0 : g.investment,
            ]);
            const dataMax = Math.max(...allVals, 0);
            const dataMin = Math.min(...allVals, 0);
            const STEP = 50000;
            const tickMax = Math.ceil(dataMax / STEP) * STEP + STEP;
            const tickMin = Math.floor(dataMin / STEP) * STEP;
            const analyticsTicks: number[] = [];
            for (let v = tickMin; v <= tickMax; v += STEP) analyticsTicks.push(v);

            // Build chart data with hidden series zeroed out
            const visibleChartData = chartData.map(g => ({
              ...g,
              deposit:    hiddenSeries.has('deposit')    ? undefined : g.deposit,
              investment: hiddenSeries.has('investment') ? undefined : g.investment,
              tradingPos: hiddenSeries.has('trading')    ? undefined : (g.trading >= 0 ? g.trading : 0),
              tradingNeg: hiddenSeries.has('trading')    ? undefined : (g.trading < 0  ? g.trading : 0),
              // keep original for tooltip
              _trading:   g.trading,
              _deposit:   g.deposit,
              _investment: g.investment,
            }));

            const fmtAxis = (val: number) => {
              const abs = Math.abs(val);
              const sign = val < 0 ? '-' : '';
              if (abs === 0) return '0';
              if (abs >= 1000) {
                const k = abs / 1000;
                return `${sign}${k % 1 === 0 ? k.toFixed(0) : k.toFixed(1)}K`;
              }
              return `${sign}${abs.toFixed(0)}`;
            };

            return (
              <Card className="p-2 sm:p-6 lg:p-8 overflow-hidden bg-slate-900 border-slate-800/60 shadow-2xl">
                {/* Title — same style as Transaction Summary */}
                <div className="mb-6 pb-6 border-b border-slate-800">
  <div className="flex items-center gap-3">
    <BarChart2 className="text-teal-400" size={20} />
    <h3 className="text-body font-bold text-white uppercase tracking-widest">Transaction History</h3>
  </div>
</div>
<div className="flex flex-wrap items-center justify-center gap-3 mb-2">
                  {/* View Selector: Monthly/Cumulative */}
                  <div className="flex items-center h-9 bg-slate-950 rounded-xl border border-slate-800 overflow-hidden">
                    <button
                      onClick={() => setAnalyticsView('monthly')}
                      className={cn(
                        "flex items-center gap-2 px-4 h-full text-[10px] font-bold uppercase transition-all tracking-wider border-r border-slate-800",
                        analyticsView === 'monthly' ? "bg-teal-400 text-slate-950" : "text-slate-400 hover:text-white hover:bg-slate-800/60"
                      )}
                    >
                      <span className={cn("w-2 h-2 rounded-full shrink-0", analyticsView === 'monthly' ? "bg-slate-950/60" : "bg-teal-400")} />
                      Monthly
                    </button>
                    <button
                      onClick={() => setAnalyticsView('cumulative')}
                      className={cn(
                        "flex items-center gap-2 px-4 h-full text-[10px] font-bold uppercase transition-all tracking-wider",
                        analyticsView === 'cumulative' ? "bg-teal-400 text-slate-950" : "text-slate-400 hover:text-white hover:bg-slate-800/60"
                      )}
                    >
                      <span className={cn("w-2 h-2 rounded-full shrink-0", analyticsView === 'cumulative' ? "bg-slate-950/60" : "bg-teal-400")} />
                      Cumulative
                    </button>
                  </div>
                  {/* Range Selector */}
                  <div className="relative">
                    <button
                      onClick={() => setIsAnalyticsRangeOpen(!isAnalyticsRangeOpen)}
                      className="flex items-center gap-3 bg-slate-950 border border-slate-800 rounded-xl px-4 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase tracking-widest"
                    >
                      <Calendar size={14} className="text-teal-400" />
                      {analyticsRange === 'last6m' ? 'Last 6M' : analyticsRange === 'last12m' ? 'Last 12M' : analyticsRange === 'fiscal' ? 'Fiscal' : 'Custom'}
                      <ChevronDown size={14} className={cn("transition-transform", isAnalyticsRangeOpen ? "rotate-180" : "")} />
                    </button>
                    {isAnalyticsRangeOpen && (
                      <>
                        <div className="fixed inset-0 z-40" onClick={() => setIsAnalyticsRangeOpen(false)} />
                        <div className="absolute left-0 mt-2 w-48 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
                          {['last6m', 'last12m', 'fiscal', 'custom'].map(id => (
                            <button
                              key={id}
                              onClick={() => { setAnalyticsRange(id as any); setIsAnalyticsRangeOpen(false); }}
                              className={cn(
                                "w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase tracking-widest",
                                analyticsRange === id ? "bg-teal-400 text-slate-950" : "text-slate-300 hover:bg-slate-800"
                              )}
                            >
                              {id === 'last6m' ? 'Last 6M' : id === 'last12m' ? 'Last 12M' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                  {analyticsRange === 'custom' && (
                    <div className="flex items-center gap-2 animate-in fade-in slide-in-from-left-2">
                      <input
                        type="date"
                        value={analyticsCustomDates.start}
                        onChange={e => setAnalyticsCustomDates(p => ({ ...p, start: e.target.value }))}
                        className="bg-slate-950 border border-slate-800 rounded-xl px-3 h-8 text-[10px] font-bold text-white outline-none focus:border-teal-400/50"
                      />
                      <span className="text-slate-600 text-[10px] font-bold">TO</span>
                      <input
                        type="date"
                        value={analyticsCustomDates.end}
                        onChange={e => setAnalyticsCustomDates(p => ({ ...p, end: e.target.value }))}
                        className="bg-slate-950 border border-slate-800 rounded-xl px-3 h-8 text-[10px] font-bold text-white outline-none focus:border-teal-400/50"
                      />
                    </div>
                  )}
                </div>

                <div className="h-[450px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={visibleChartData}
                      margin={{ top: 38, right: isMobile ? 13 : 38, left: isMobile ? -12 : 18, bottom: 18 }}
                      barGap={1}
                      barCategoryGap="25%"
                      stackOffset="sign"
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
                        width={48}
                      />
                      <Tooltip
                        cursor={{ fill: '#ffffff08' }}
                        content={({ active, payload, label }) => {
                          if (!active || !payload?.length) return null;
                          const dData = payload[0]?.payload;
                          const dep  = dData?._deposit    ?? 0;
                          const trad = dData?._trading    ?? 0;
                          const inv  = dData?._investment ?? 0;
                          // Reconstruct trading buy/sell from net trading stored in original transactions
                          // We store net trading = buy - sell, so we show it split
                          const tradPos = trad >= 0 ? trad : 0;
                          const tradNeg = trad < 0 ? Math.abs(trad) : 0;
                          const netTrad = trad;
                          return (
                            <div className="bg-[#0f172a]/95 backdrop-blur border border-slate-700/60 p-4 rounded-xl shadow-2xl min-w-[240px]">
                              <p className="text-[12px] font-bold text-white mb-3 border-b border-slate-700/60 pb-2">
                                {label}{analyticsView === 'cumulative' ? ' (Cumulative)' : ''}
                              </p>
                              <div className="space-y-2">
                                {!hiddenSeries.has('deposit') && (
                                  <div className="flex items-center justify-between gap-8">
                                    <div className="flex items-center gap-2">
                                      <div className="w-2 h-2 rounded-full" style={{ backgroundColor: TEAL }} />
                                      <span className="text-[11px] text-slate-300">
                                        {analyticsView === 'cumulative' ? 'Total Deposit' : 'Monthly Deposit'}
                                      </span>
                                    </div>
                                    <span className="text-[11px] font-bold text-white tabular-nums">
                                      ৳{dep.toLocaleString('en-IN')}
                                    </span>
                                  </div>
                                )}
                                {!hiddenSeries.has('investment') && (
                                  <div className="flex items-center justify-between gap-8">
                                    <div className="flex items-center gap-2">
                                      <div className="w-2 h-2 rounded-full" style={{ backgroundColor: BLUE }} />
                                      <span className="text-[11px] text-slate-300">Investment Buy</span>
                                    </div>
                                    <span className="text-[11px] font-bold text-white tabular-nums">
                                      ৳{inv.toLocaleString('en-IN')}
                                    </span>
                                  </div>
                                )}
                                {!hiddenSeries.has('trading') && (
                                  <>
                                    <div className="flex items-center justify-between gap-8">
                                      <div className="flex items-center gap-2">
                                        <div className="w-2 h-2 rounded-full" style={{ backgroundColor: AMBER }} />
                                        <span className="text-[11px] text-slate-300">Trading Buy</span>
                                      </div>
                                      <span className="text-[11px] font-bold text-white tabular-nums">
                                        ৳{tradPos.toLocaleString('en-IN')}
                                      </span>
                                    </div>
                                    <div className="flex items-center justify-between gap-8">
                                      <div className="flex items-center gap-2">
                                        <div className="w-2 h-2 rounded-full bg-rose-400" />
                                        <span className="text-[11px] text-slate-300">Trading Sell</span>
                                      </div>
                                      <span className="text-[11px] font-bold text-rose-400 tabular-nums">
                                        ৳{tradNeg.toLocaleString('en-IN')}
                                      </span>
                                    </div>
                                    <div className="flex items-center justify-between gap-8 pt-1 border-t border-slate-700/40">
                                      <div className="flex items-center gap-2">
                                        <div className="w-2 h-2 rounded-full" style={{ backgroundColor: AMBER }} />
                                        <span className="text-[11px] font-bold text-slate-200">Net Trading</span>
                                      </div>
                                      <span className="text-[11px] font-bold tabular-nums" style={{ color: netTrad < 0 ? '#f87171' : '#f59e0b' }}>
                                        ৳{netTrad.toLocaleString('en-IN')}
                                      </span>
                                    </div>
                                  </>
                                )}
                              </div>
                            </div>
                          );
                        }}
                      />
                      <ReferenceLine y={0} stroke="#334155" strokeWidth={1} />

                      {/* Deposit bar — LEFT of each group */}
                      <Bar dataKey="deposit" stackId="deposit" fill={TEAL} radius={[4, 4, 0, 0]} barSize={22} isAnimationActive={false} legendType="none">
                        <LabelList
                          dataKey="deposit"
                          position="top"
                          content={(props: any) => {
                            const { x, y, width, value } = props;
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

                      {/* Stacked Portfolio bars — RIGHT of each group */}
                      {/* Negative trading (goes below zero) */}
                      <Bar dataKey="tradingNeg" stackId="portfolio" fill={AMBER} radius={[0, 0, 0, 0]} barSize={22} isAnimationActive={false} legendType="none">

                        <LabelList
                          dataKey="tradingNeg"
                          position="bottom"
                          content={(props: any) => {
                            const { x, y, width, value } = props;
                            if (!value) return null;
                            return (
                              <g transform={`translate(${x + width / 2},${y + 10})`}>
                                <text transform="rotate(-90)" fill="#64748b" textAnchor="end" fontSize={10} fontWeight={700} className="font-display">
                                  {fmtAxis(value)}
                                </text>
                              </g>
                            );
                          }}
                        />
                      </Bar>
                      {/* Positive trading */}
                      <Bar dataKey="tradingPos" stackId="portfolio" fill={AMBER} barSize={22} isAnimationActive={false} legendType="none" />
                      {/* Investment stacked on top of positive trading */}
                      <Bar dataKey="investment" stackId="portfolio" fill={BLUE} radius={[4, 4, 0, 0]} barSize={22} isAnimationActive={false} legendType="none">
                        <LabelList
                          dataKey="investment"
                          position="top"
                          content={(props: any) => {
                            const { x, y, width, value, index } = props;
                            if (hiddenSeries.has('investment') && hiddenSeries.has('trading')) return null;
                            const gData = visibleChartData[index];
                            const stackTotal = (gData?.tradingPos ?? 0) + (value ?? 0);
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

                {/* Legends — same style as Transaction Summary, clickable */}
                <div className="mt-6 pt-6 border-t border-slate-800/50 flex flex-wrap justify-center gap-x-6 gap-y-3">
  {ANALYTICS_LEGENDS.map(({ key, label, color }) => {
                    const hidden = hiddenSeries.has(key);
                    return (
                      <button
                        key={key}
                        type="button"
                        onClick={() => toggleSeries(key)}
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
              </Card>
            );
          })()}

          {/* ── Return History Chart ── */}
          {(() => {
            const fmtAxis = (val: number) => {
              const abs = Math.abs(val);
              const sign = val < 0 ? '-' : '';
              if (abs === 0) return '0';
              if (abs >= 1000) return `${sign}${(abs/1000).toFixed(abs % 1000 === 0 ? 0 : 1)}K`;
              return `${sign}${abs.toFixed(0)}`;
            };
            const fmtFull = (n: number) => {
              const abs = Math.abs(n);
              const sign = n < 0 ? '-৳ ' : '৳ ';
              return sign + abs.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            };

            return (
              <Card className="p-2 sm:p-6 lg:p-8 overflow-hidden bg-slate-900 border-slate-800/60 shadow-2xl">
                  {/* Header */}
                  <div className="mb-6 pb-6 border-b border-slate-800">
  <div className="flex items-center gap-3">
    <BarChart2 className="text-teal-400" size={20} />
    <h3 className="text-body font-bold text-white uppercase tracking-widest">Return History</h3>
  </div>
</div>
<div className="flex flex-wrap items-center justify-center gap-3 mb-2">
                    {/* Monthly / Cumulative toggle */}
                  <div className="flex items-center h-9 bg-slate-950 rounded-xl border border-slate-800 overflow-hidden">
                    <button
                      onClick={() => setReturnHistoryView('monthly')}
                      className={cn(
                        'flex items-center gap-2 px-4 h-full text-[10px] font-bold uppercase transition-all tracking-wider border-r border-slate-800',
                        returnHistoryView === 'monthly' ? 'bg-teal-400 text-slate-950' : 'text-slate-400 hover:text-white hover:bg-slate-800/60'
                      )}
                    >
                      <span className={cn('w-2 h-2 rounded-full shrink-0', returnHistoryView === 'monthly' ? 'bg-slate-950/60' : 'bg-teal-400')} />
                      Monthly
                    </button>
                    <button
                      onClick={() => setReturnHistoryView('cumulative')}
                      className={cn(
                        'flex items-center gap-2 px-4 h-full text-[10px] font-bold uppercase transition-all tracking-wider',
                        returnHistoryView === 'cumulative' ? 'bg-teal-400 text-slate-950' : 'text-slate-400 hover:text-white hover:bg-slate-800/60'
                      )}
                    >
                      <span className={cn('w-2 h-2 rounded-full shrink-0', returnHistoryView === 'cumulative' ? 'bg-slate-950/60' : 'bg-teal-400')} />
                      Cumulative
                    </button>
                  </div>

                    {/* Range selector */}
                    <div className="relative">
                      <button
                        onClick={() => setIsReturnHistoryRangeOpen(!isReturnHistoryRangeOpen)}
                        className="flex items-center gap-3 bg-slate-950 border border-slate-800 rounded-xl px-4 h-9 text-[10px] font-bold text-slate-300 hover:text-white transition-all uppercase tracking-widest"
                      >
                        <Calendar size={14} className="text-teal-400" />
                        {returnHistoryRange === 'last6m' ? 'Last 6M' : returnHistoryRange === 'last12m' ? 'Last 12M' : returnHistoryRange === 'fiscal' ? 'Fiscal' : 'Custom'}
                        <ChevronDown size={14} className={cn('transition-transform', isReturnHistoryRangeOpen ? 'rotate-180' : '')} />
                      </button>
                      {isReturnHistoryRangeOpen && (
                        <>
                          <div className="fixed inset-0 z-40" onClick={() => setIsReturnHistoryRangeOpen(false)} />
                          <div className="absolute left-0 mt-2 w-48 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-50 p-1 animate-in fade-in zoom-in-95">
                            {['last6m', 'last12m', 'fiscal', 'custom'].map(id => (
                              <button key={id}
                                onClick={() => { setReturnHistoryRange(id as any); setIsReturnHistoryRangeOpen(false); }}
                                className={cn(
                                  'w-full text-left px-3 py-2 text-[10px] font-bold rounded-lg transition-colors uppercase tracking-widest',
                                  returnHistoryRange === id ? 'bg-teal-400 text-slate-950' : 'text-slate-300 hover:bg-slate-800'
                                )}
                              >
                                {id === 'last6m' ? 'Last 6M' : id === 'last12m' ? 'Last 12M' : id === 'fiscal' ? 'Fiscal' : 'Custom'}
                              </button>
                            ))}
                          </div>
                        </>
                      )}
                    </div>

                    {/* Custom date inputs */}
                    {returnHistoryRange === 'custom' && (
                      <div className="flex items-center gap-2 animate-in fade-in slide-in-from-left-2">
                        <input type="date" value={returnHistoryCustomDates.start}
                          onChange={e => setReturnHistoryCustomDates(p => ({ ...p, start: e.target.value }))}
                          className="bg-slate-950 border border-slate-800 rounded-xl px-3 h-8 text-[10px] font-bold text-white outline-none focus:border-teal-400/50" />
                        <span className="text-slate-600 text-[10px] font-bold">TO</span>
                        <input type="date" value={returnHistoryCustomDates.end}
                          onChange={e => setReturnHistoryCustomDates(p => ({ ...p, end: e.target.value }))}
                          className="bg-slate-950 border border-slate-800 rounded-xl px-3 h-8 text-[10px] font-bold text-white outline-none focus:border-teal-400/50" />
                      </div>
                    )}
                  </div>

                <div className="h-[320px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart
                      data={returnHistoryData}
                      margin={{ top: 38, right: isMobile ? 13 : 38, left: isMobile ? -12 : 18, bottom: 18 }}
                      stackOffset="sign"
                    >
                      <CartesianGrid stroke="#334155" strokeOpacity={0.4} vertical={true} horizontal={true} />
                      <XAxis dataKey="label" stroke="#94a3b8" fontSize={11} fontWeight={700}
                        tickLine={false} axisLine={false}
                        tickFormatter={(val) => String(val).toUpperCase()}
                        tick={{ 
                          dy: 10, 
                          fill: '#94a3b8', 
                          fontSize: 11, 
                          fontWeight: 700,
                          style: { letterSpacing: '0.05em' }
                        }} />
                      <YAxis stroke="#475569" fontSize={10} fontWeight={700}
                        tickLine={false} axisLine={false}
                        ticks={yAxisTicksReturn} tickFormatter={fmtAxis}
                        tick={{ fill: '#475569', fontSize: 10, fontWeight: 700 }} width={48} />
                      <Tooltip
                        cursor={{ fill: '#ffffff08' }}
                        content={({ active, payload, label }) => {
                          if (!active || !payload?.length) return null;
                          const d = payload[0]?.payload;
                          const netReturn = (d?.pnl ?? 0) + (d?.dividend ?? 0) - Math.abs(d?.charge ?? 0);
                          return (
                            <div className="bg-[#0f172a]/95 backdrop-blur border border-slate-700/60 p-4 rounded-xl shadow-2xl min-w-[220px]">
                              <p className="text-[12px] font-bold text-white mb-3 border-b border-slate-700/60 pb-2">
                                {label}{returnHistoryView === 'cumulative' ? ' (Cumulative)' : ''}
                              </p>
                              <div className="space-y-2">
                                <div className="flex items-center justify-between gap-8">
                                  <div className="flex items-center gap-2">
                                    <div className="w-2 h-2 rounded-full bg-purple-500" />
                                    <span className="text-[11px] text-slate-300">Dividend</span>
                                  </div>
                                  <span className="text-[11px] font-bold text-white tabular-nums">{fmtFull(d?.dividend ?? 0)}</span>
                                </div>
                                <div className="flex items-center justify-between gap-8">
                                  <div className="flex items-center gap-2">
                                    <div className="w-2 h-2 rounded-full" style={{ backgroundColor: (d?.pnl ?? 0) >= 0 ? '#4ade80' : '#f87171' }} />
                                    <span className="text-[11px] text-slate-300">Realized P&L</span>
                                  </div>
                                  <span className={`text-[11px] font-bold tabular-nums ${(d?.pnl ?? 0) >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>{fmtFull(d?.pnl ?? 0)}</span>
                                </div>
                                {(d?.charge ?? 0) > 0 && (
                                  <div className="flex items-center justify-between gap-8">
                                    <div className="flex items-center gap-2">
                                      <div className="w-2 h-2 rounded-full bg-rose-500" />
                                      <span className="text-[11px] text-slate-300">Charge</span>
                                    </div>
                                    <span className="text-[11px] font-bold text-rose-400 tabular-nums">{fmtFull(d?.charge ?? 0)}</span>
                                  </div>
                                )}
                                <div className="flex items-center justify-between gap-8 pt-1 border-t border-slate-700/40">
                                  <span className="text-[11px] font-bold text-slate-200">Net Return</span>
                                  <span className={`text-[11px] font-bold tabular-nums ${netReturn >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>{fmtFull(netReturn)}</span>
                                </div>
                              </div>
                            </div>
                          );
                        }}
                      />
                      <ReferenceLine y={0} stroke="#334155" strokeWidth={1} />

                      <Bar dataKey="pnl" name="Realized P&L" stackId="a" radius={[3, 3, 0, 0]} barSize={22} isAnimationActive={false} legendType="none">
                        {returnHistoryData.map((entry: any, index: number) => (
                          <Cell key={`cell-pnl-${index}`} fill={entry.pnl >= 0 ? '#4ade80' : '#f87171'} />
                        ))}
                      </Bar>
                      <Bar dataKey="dividend" name="Dividend" stackId="a" fill="#a855f7" radius={[3, 3, 0, 0]} barSize={22} isAnimationActive={false} legendType="none" />
                      <Line dataKey="labelY" stroke="none" dot={false} activeDot={false} isAnimationActive={false} legendType="none">
                        <LabelList dataKey="netReturn" position="top" offset={10}
                          formatter={(val: any) => val === 0 ? '' : fmtAxis(val)}
                          style={{ fill: '#475569', fontSize: 10, fontWeight: 700 }} />
                      </Line>
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>

                <div className="mt-6 pt-6 border-t border-slate-800/50 flex flex-wrap justify-center gap-x-6 gap-y-3">
  {[
    { color: '#4ade80', label: 'Realized Profit' },
                    { color: '#f87171', label: 'Realized Loss' },
                    { color: '#a855f7', label: 'Dividend' },
                  ].map(({ color, label }) => {
                    const isHidden = hiddenSeries.has(label);
                    return (
                      <button 
                        key={label} 
                        className="flex items-center gap-2 group transition-opacity hover:opacity-80"
                        onClick={() => {
                          setHiddenSeries(prev => {
                            const next = new Set(prev);
                            if (next.has(label)) next.delete(label);
                            else next.add(label);
                            return next;
                          });
                        }}
                      >
                        <div 
                          className="w-3 h-3 rounded-sm shrink-0 transition-all" 
                          style={{ 
                            backgroundColor: isHidden ? '#1e293b' : color,
                            boxShadow: isHidden ? 'none' : `0 0 8px ${color}40`,
                            border: isHidden ? '1px solid #334155' : 'none'
                          }} 
                        />
                        <span className={cn(
                          "text-[12px] font-bold uppercase tracking-tight transition-all",
                          isHidden ? "text-slate-600 line-through decoration-slate-500" : "text-slate-300 group-hover:text-slate-200"
                        )}>
                          {label}
                        </span>
                      </button>
                    );
                  })}
                </div>
              </Card>
            );
          })()}

          {/* ── Return by Stock Chart ── */}

          {(() => {
            const stockChartData = returnByStockType === 'Dividend'
              ? returnByStockData.filter((s: any) => s.dividends !== 0)
                  .map((s: any) => ({ name: s.ticker, value: s.dividends }))
                  .sort((a: any, b: any) => b.value - a.value)
              : returnByStockType === 'Net Return'
                ? returnByStockData.map((s: any) => ({ name: s.ticker, value: s.realizedPnL + s.dividends }))
                    .filter((s: any) => s.value !== 0)
                    .sort((a: any, b: any) => b.value - a.value)
                : returnByStockData.filter((s: any) => s.realizedPnL !== 0)
                    .map((s: any) => ({ name: s.ticker, value: s.realizedPnL }))
                    .sort((a: any, b: any) => b.value - a.value);

            const fmtAxis = (val: number) => {
              const abs = Math.abs(val);
              const sign = val < 0 ? '-' : '';
              if (abs === 0) return '0';
              if (abs >= 10000000) return `${sign}${(abs/10000000).toFixed(abs%10000000===0?0:1)}Cr`;
              if (abs >= 100000) return `${sign}${(abs/100000).toFixed(0)}L`;
              if (abs >= 1000) return `${sign}${(abs/1000).toFixed(abs%1000===0?0:1)}K`;
              return `${sign}${abs.toFixed(0)}`;
            };

           const xTicks = (() => {
              const values = stockChartData.map((d: any) => d.value);
              if (values.length === 0) return [0, 500];
              const minVal = Math.min(...values, 0);
              const maxVal = Math.max(...values, 0);
              const step = 500;
              const ticks = [];
              const start = Math.floor(minVal / step) * step;
              const end = Math.ceil(maxVal / step) * step;
              for (let i = start; i <= end; i += step) ticks.push(i);
              return ticks;
            })();

            const BAR_SIZE = 16;
const BAR_GAP = 10;
const chartHeight = Math.max(200, stockChartData.length * (BAR_SIZE + BAR_GAP) + 40);

            return (
              <Card className="p-2 sm:p-6 lg:p-8 overflow-hidden bg-slate-900 border-slate-800/60 shadow-2xl">
                <div className="mb-6 pb-6 border-b border-slate-800">
  <div className="flex items-center gap-3">
    <BarChart2 className="text-teal-400" size={20} />
    <h3 className="text-body font-bold text-white uppercase tracking-widest">Return by Stock</h3>
  </div>
</div>

{/* Toggle */}
<div className="flex justify-center mb-6">
                  <div className="flex items-center bg-slate-950/50 rounded-lg p-1 border border-slate-800/50 gap-1">
                    {(['Realized P&L', 'Dividend', 'Net Return'] as const).map(type => (
                      <button key={type} onClick={() => setReturnByStockType(type)}
                        className={cn(
                          'px-4 py-1.5 rounded-md text-[10px] font-bold uppercase transition-all border whitespace-nowrap',
                          returnByStockType === type
                            ? 'bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20 border-teal-400/30'
                            : 'bg-slate-900/40 border-slate-800/40 text-slate-300 hover:text-white hover:bg-slate-800/60'
                        )}>
                        {type}
                      </button>
                    ))}
                  </div>
                </div>

                {stockChartData.length === 0 ? (
                  <div className="flex items-center justify-center py-16 text-slate-600 text-[11px] font-bold uppercase tracking-wider">
                    No {returnByStockType} data found
                  </div>
                ) : (
                  <div className="w-full relative" style={{ height: chartHeight }}>
                    {/* Custom vertical grid lines + top/bottom borders */}
                    {(() => {
                      const leftOffset = isMobile ? 52 : 128;
                      const rightOffset = isMobile ? 40 : 80;
                      const LABEL_AREA = 28; // px below bottom line for x-axis labels
                      return (
                        <div
                          className="absolute pointer-events-none"
                          style={{
                            left: leftOffset,
                            right: rightOffset,
                            top: -10,
                            bottom: 0,
                          }}
                        >
                          {/* Vertical grid lines + custom labels */}
                          {xTicks.map((tick, i) => {
                            const minTick = xTicks[0];
                            const maxTick = xTicks[xTicks.length - 1];
                            const range = maxTick - minTick;
                            if (range === 0) return null;
                            const pct = ((tick - minTick) / range) * 100;
                            if (isMobile && i % 2 !== 0) return null;
                            return (
                              <div key={i} className="absolute" style={{ left: `${pct}%`, top: 0, bottom: 0 }}>
                                {/* Vertical line — stops at bottom axis line */}
                                <div className="absolute border-l border-slate-700/40" style={{ top: 0, bottom: 20, left: 0 }} />
                                {/* Label below bottom horizontal line */}
                                <span
                                  className="absolute text-[9px] font-bold text-slate-500 -translate-x-1/2 whitespace-nowrap"
                                  style={{ bottom: 0 }}
                                >
                                  {fmtAxis(tick)}
                                </span>
                              </div>
                            );
                          })}
                          {/* Top horizontal line */}
                          <div
                            className="absolute border-t border-slate-700/40"
                            style={{ top: 20, left: -20, right: -20 }}
                          />
                          {/* Bottom horizontal line — above the label area */}
                          <div
                            className="absolute border-t border-slate-700/40"
                            style={{ bottom: 40, left: -20, right: -20 }}
                          />
                        </div>
                      );
                    })()}
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart
                        layout="vertical"
                        data={stockChartData}
                        margin={{ top: 20, right: isMobile ? 12 : 80, left: isMobile ? -6 : 48, bottom: 20 }}
                        barSize={16}
                        barCategoryGap="38%"
                      >
                        {/* custom grid handled by wrapper div below */}
                        <XAxis
                          type="number"
                          axisLine={false}
                          tickLine={false}
                          tick={false}
                          ticks={xTicks}
                          height={28}
                                                />
                        <YAxis
  dataKey="name"
  type="category"
  axisLine={false}
  tickLine={false}
  tick={{ fill: '#94a3b8', fontSize: isMobile ? 8 : 11, fontWeight: 700 }}
  width={isMobile ? 72 : 80}
  interval={0}
  className="uppercase tracking-wider font-bold"
/>
                        <Tooltip
                          cursor={{ fill: '#ffffff08' }}
                          content={({ active, payload }) => {
                            if (!active || !payload?.length) return null;
                            const d = payload[0]?.payload;
                            return (
                              <div className="bg-[#0f172a]/95 backdrop-blur border border-slate-700/60 p-3 rounded-xl shadow-2xl">
                                <p className="text-[12px] font-bold text-white mb-1">{d.name}</p>
                                <p className={`text-[12px] font-bold tabular-nums ${d.value >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                                  {returnByStockType}: ৳{d.value.toLocaleString('en-IN')}
                                </p>
                              </div>
                            );
                          }}
                        />
                        <ReferenceLine x={0} stroke="#1e293b" strokeWidth={2} />
                        <Bar dataKey="value" radius={[0, 4, 4, 0]} isAnimationActive={false}>
                          {stockChartData.map((entry: any, index: number) => (
                            <Cell
                              key={`cell-stock-${index}`}
                              fill={entry.value >= 0
                                ? (returnByStockType === 'Dividend' ? '#a855f7' : returnByStockType === 'Net Return' ? '#38bdf8' : '#4ade80')
                                : '#f87171'}
                            />
                          ))}
                          <LabelList
  dataKey="value"
  content={(props: any) => {
    const { x, y, width, height, value } = props;
    if (value === 0 || value == null) return null;
    const abs = Math.abs(value);
    const sign = value < 0 ? '-' : '';
    const short = abs >= 10000000
      ? `${sign}${(abs / 10000000).toFixed(1)}Cr`
      : abs >= 100000
      ? `${sign}${(abs / 100000).toFixed(1)}L`
      : abs >= 1000
      ? `${sign}${(abs / 1000).toFixed(1)}K`
      : `${sign}${abs.toFixed(0)}`;

    if (value >= 0) {
      // Positive: label to the RIGHT of the bar end
      return (
        <text
          x={x + width + 6}
          y={y + height / 2}
          fill="#94a3b8"
          textAnchor="start"
          dominantBaseline="middle"
          fontSize={10}
          fontWeight={700}
        >
          {`৳${short}`}
        </text>
      );
    } else {
      // Negative: label in the MIDDLE of the bar
      return (
        <text
          x={x + width / 2}
          y={y + height / 2}
          fill="white"
          textAnchor="middle"
          dominantBaseline="middle"
          fontSize={10}
          fontWeight={700}
        >
          {`৳${short}`}
        </text>
      );
    }
  }}
/>
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                )}

                {/* Legends — same style as Transaction Summary */}
                {stockChartData.length > 0 && (
                 <div className="mt-6 pt-6 border-t border-slate-800/50 flex flex-wrap justify-center gap-x-6 gap-y-3">
  <div className="flex items-center gap-2">
    <div className="w-3 h-3 rounded-sm shrink-0" 
      style={{ backgroundColor: returnByStockType === 'Dividend' ? '#a855f7' : returnByStockType === 'Net Return' ? '#38bdf8' : '#4ade80' }} />
    <span className="text-[12px] font-bold text-slate-300 uppercase tracking-tight whitespace-nowrap">
      {returnByStockType === 'Dividend' ? 'Dividend' : 
       returnByStockType === 'Realized P&L' ? 'Realized Profit' : 
       `Positive ${returnByStockType}`}
    </span>
  </div>
                    {returnByStockType !== 'Dividend' && (
                      <div className="flex items-center gap-2">
                        <div className="w-3 h-3 rounded-sm shrink-0" style={{ backgroundColor: '#f87171' }} />
                        <span className="text-[12px] font-bold text-slate-300 uppercase tracking-tight whitespace-nowrap">
                          {returnByStockType === 'Realized P&L' ? 'Realized Loss' : 
                           `Negative ${returnByStockType}`}
                        </span>
                      </div>
                    )}
                  </div>
                )}
              </Card>
            );
          })()}

          
        </div>
      )}

      {/* ══════════════════════════════════════════════════
          SETTINGS TAB
      ══════════════════════════════════════════════════ */}
      {activeTab === 'settings' && (
        <div className="space-y-4 sm:space-y-6 max-w-4xl mx-auto pb-8 px-1 sm:px-0">
          {/* Commission Settings */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 sm:p-6 shadow-xl">
            <div className="flex items-center gap-3 mb-6">
              <Percent size={18} className="text-teal-400 sm:size-5" />
              <h3 className="text-[11px] sm:text-body font-bold text-white uppercase tracking-widest">Commission Settings</h3>
            </div>
            
            <div className="flex flex-col sm:flex-row items-end gap-3 sm:gap-4">
              <div className="w-full sm:flex-1">
                <label className="text-[10px] font-bold text-slate-400 uppercase mb-2 block tracking-wider">Default Rate (%)</label>
                <div className="relative">
                  <input 
                    type="number" 
                    step="0.01"
                    className="w-full h-9 bg-slate-950 border border-slate-800 rounded-xl px-4 text-xs font-bold text-white focus:border-teal-400/50 outline-none transition-all tabular-nums"
                    value={dseSettings.commissionRate}
                    onChange={(e) => setDseSettings(prev => ({ ...prev, commissionRate: parseFloat(e.target.value) || 0 }))}
                  />
                  <span className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-600 text-xs font-bold">%</span>
                </div>
              </div>
              <button 
                onClick={() => {
                  saveSettings(dseSettings);
                  setApiStatus({ message: 'Commission rate saved successfully', isError: false });
                  setTimeout(() => setApiStatus(null), 3000);
                }}
                className="w-full sm:w-auto h-9 px-6 bg-teal-400 hover:bg-teal-300 text-slate-950 text-xs font-bold rounded-xl flex items-center justify-center gap-2 transition-all active:scale-95 shadow-lg shadow-teal-400/10"
              >
                <RefreshCw size={14} className="sm:size-4" /> Update
              </button>
            </div>
            <p className="mt-3 text-[9px] sm:text-[10px] font-medium text-slate-500 italic uppercase tracking-tight">Used for auto-calculating commission for new transactions.</p>
          </div>

          {/* Stock Database */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 sm:p-6 shadow-xl">
            <div className="flex items-center gap-3 mb-2">
              <Layers size={18} className="text-teal-400 sm:size-5" />
              <h3 className="text-[11px] sm:text-body font-bold text-white uppercase tracking-widest">Stock Database</h3>
            </div>
            <p className="text-[9px] sm:text-[10px] font-medium text-slate-500 uppercase tracking-tight mb-6 leading-relaxed">Add new stocks or edit existing company names. Custom entries will override default DSE stock names.</p>

            <div className="grid grid-cols-1 sm:grid-cols-[1fr_2fr_auto] gap-3 sm:gap-4 items-end mb-6 sm:mb-8">
              <div>
                <label className="text-[10px] font-bold text-slate-400 uppercase mb-2 block tracking-wider">Ticker</label>
                <input 
                  type="text" 
                  placeholder="e.g. MYSTOCK"
                  className="w-full h-9 bg-slate-950 border border-slate-800 rounded-xl px-4 text-xs font-bold text-white focus:border-teal-400/50 outline-none transition-all uppercase"
                  id="new-stock-ticker"
                />
              </div>
              <div>
                <label className="text-[10px] font-bold text-slate-400 uppercase mb-2 block tracking-wider">Company Name</label>
                <input 
                  type="text" 
                  placeholder="e.g. My Custom Company"
                  className="w-full h-9 bg-slate-950 border border-slate-800 rounded-xl px-4 text-xs font-bold text-white focus:border-teal-400/50 outline-none transition-all"
                  id="new-stock-name"
                />
              </div>
              <button 
                onClick={async () => {
                  const tickerInput = document.getElementById('new-stock-ticker') as HTMLInputElement;
                  const nameInput = document.getElementById('new-stock-name') as HTMLInputElement;
                  const ticker = tickerInput.value.trim().toUpperCase();
                  const name = nameInput.value.trim();
                  
                  if (ticker && name) {
                    const existsInTransactions = transactions.some(t => t.ticker?.toUpperCase() === ticker);
                    
                    if (existsInTransactions) {
                      const updatedTransactions = transactions.map(t => 
                        t.ticker?.toUpperCase() === ticker ? { ...t, companyName: name } : t
                      );
                      setTransactions(updatedTransactions);
                      localStorage.setItem('sheet_cache_dse', JSON.stringify(updatedTransactions));
                      await pushModuleData('dse', updatedTransactions);
                    }

                    const existsInCustom = dseSettings.customStocks.some(s => s.ticker === ticker);
                    let newCustom;
                    if (existsInCustom) {
                      newCustom = dseSettings.customStocks.map(s => s.ticker === ticker ? { ticker, name } : s);
                    } else {
                      newCustom = [...dseSettings.customStocks, { ticker, name }];
                    }
                    saveSettings({ ...dseSettings, customStocks: newCustom });
                    
                    tickerInput.value = '';
                    nameInput.value = '';
                    
                    const statusMsg = existsInTransactions 
                      ? `Stock ${ticker} updated in custom database and all transaction history`
                      : `Stock ${ticker} ${existsInCustom ? 'updated' : 'added'} successfully`;
                    setApiStatus({ message: statusMsg, isError: false });
                    setTimeout(() => setApiStatus(null), 3500);
                  }
                }}
                className="w-full sm:w-auto h-9 px-6 bg-teal-400 hover:bg-teal-300 text-slate-950 text-xs font-bold rounded-xl flex items-center justify-center gap-2 transition-all active:scale-95 shadow-lg shadow-teal-400/10"
              >
                <Plus size={14} className="sm:size-4" /> Add/Update
              </button>
            </div>

            {/* Custom stocks list */}
            {dseSettings.customStocks.length > 0 && (
              <div className="space-y-2 mt-4 max-h-60 overflow-y-auto custom-scrollbar pr-2">
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2 block">Custom/Edited Stocks</label>
                {dseSettings.customStocks.map((stock) => (
                  <div key={stock.ticker} className="flex items-center justify-between p-3 bg-slate-950/50 border border-slate-800 rounded-xl group hover:border-slate-700 transition-all">
                    <div>
                      <p className="text-xs font-bold text-white tracking-widest">{stock.ticker}</p>
                      <p className="text-[10px] text-slate-500 font-medium">{stock.name}</p>
                    </div>
                    <button 
                      onClick={() => {
                        setConfirmDialog({
                          isOpen: true,
                          title: 'Remove Stock',
                          message: `Are you sure you want to remove ${stock.ticker} from your custom stock list?`,
                          confirmLabel: 'Remove',
                          variant: 'warning',
                          onConfirm: () => {
                            setConfirmDialog(prev => ({ ...prev, isOpen: false }));
                            const newCustom = dseSettings.customStocks.filter(s => s.ticker !== stock.ticker);
                            saveSettings({ ...dseSettings, customStocks: newCustom });
                          }
                        });
                      }}
                      className="p-2 text-slate-600 hover:text-rose-400 transition-colors"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

        </div>
      )}

      {confirmDialog.isOpen && (
        <ConfirmDialog 
          {...confirmDialog}
          onCancel={() => setConfirmDialog(prev => ({ ...prev, isOpen: false }))}
        />
      )}
    </div>
  );
};