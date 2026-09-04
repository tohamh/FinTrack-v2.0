/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { 
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LabelList, Line, ComposedChart, ReferenceLine
} from 'recharts';
import { Card } from '../ui/BaseComponents';
import { formatBDT, cn } from '../../utils/formatters';
import { 
  ArrowUpRight, ArrowDownRight, Wallet, TrendingUp, Briefcase, 
  PieChart as PieIcon, ChevronLeft, ChevronRight, ChevronDown, Landmark 
} from 'lucide-react';

const COLORS = ['#00f5d4', '#3a86ff', '#8338ec', '#ff0a54', '#ffd60a'];

const COLOR_CONFIG = {
  teal:    { icon: 'bg-teal-400/10 text-teal-400',    value: 'text-teal-400',    hover: 'hover:border-teal-400/50 hover:shadow-[0_0_20px_rgba(45,212,191,0.1)]',    footer: 'text-teal-400'    },
  amber:   { icon: 'bg-amber-400/10 text-amber-400',  value: 'text-amber-400',   hover: 'hover:border-amber-400/50 hover:shadow-[0_0_20px_rgba(251,191,36,0.1)]',    footer: 'text-amber-400'   },
  purple:  { icon: 'bg-purple-400/10 text-purple-400',value: 'text-purple-400',  hover: 'hover:border-purple-400/50 hover:shadow-[0_0_20px_rgba(167,139,250,0.1)]',  footer: 'text-purple-400'  },
  emerald: { icon: 'bg-emerald-400/10 text-emerald-400',value:'text-emerald-400',hover: 'hover:border-emerald-400/50 hover:shadow-[0_0_20px_rgba(52,211,153,0.1)]',  footer: 'text-emerald-400' },
  rose:    { icon: 'bg-rose-500/10 text-rose-500',    value: 'text-rose-500',    hover: 'hover:border-rose-500/50 hover:shadow-[0_0_20px_rgba(244,63,94,0.1)]',      footer: 'text-rose-500'    },
  blue:    { icon: 'bg-blue-400/10 text-blue-400',    value: 'text-blue-400',    hover: 'hover:border-blue-400/50 hover:shadow-[0_0_20px_rgba(96,165,250,0.1)]',     footer: 'text-blue-400'    },
  indigo:  { icon: 'bg-indigo-400/10 text-indigo-400',value:'text-indigo-400',   hover: 'hover:border-indigo-400/50 hover:shadow-[0_0_20px_rgba(129,140,248,0.1)]',  footer: 'text-indigo-400'  },
};

export const SummaryCard: React.FC<{ 
  title: string; 
  subtitle?: string;
  value: number; 
  trend?: number | null;
  trendLabel?: string;
  footerLabel?: string;
  footerValue?: number | null;
  icon: React.ElementType;
  color?: keyof typeof COLOR_CONFIG;
  className?: string;
  hideFooter?: boolean;
  borrowedValue?: number;
  lentValue?: number;
}> = ({ title, subtitle, value, trend, trendLabel, footerLabel, footerValue, icon: Icon, color = 'teal', className, hideFooter, borrowedValue, lentValue }) => {
  const cfg = COLOR_CONFIG[color] ?? COLOR_CONFIG.teal;

  const isProfit = title.toLowerCase().includes('profit');
  const mainValueColor = isProfit
    ? (value >= 0 ? 'text-emerald-400' : 'text-rose-500')
    : cfg.value;

  const fv = footerValue ?? 0;
  const footerValueColor = fv >= 0 ? 'text-emerald-400' : 'text-rose-500';
  const resolvedTrendLabel = trendLabel ?? 'vs Last Year';

  return (
    <Card className={cn(
      "bg-slate-900 border-slate-800 h-full flex flex-col p-3 sm:p-4 transition-all group",
      cfg.hover,
      className
    )}>
      <div className="flex-1">
        <div className={cn(
          "w-8 h-8 sm:w-10 sm:h-10 rounded-xl flex items-center justify-center mb-4 group-hover:scale-110 transition-transform",
          cfg.icon
        )}>
          <Icon size={20} />
        </div>
        <div className="mb-2">
          <div className="flex items-center gap-1">
            <p className="text-body-sm font-bold text-white uppercase tracking-wider">{title}</p>
          </div>
          {subtitle && <p className="text-label font-bold text-slate-500 uppercase">{subtitle}</p>}
        </div>
        <h3 className={cn(
          "text-heading font-bold mb-4 tracking-tight font-display tabular-nums",
          mainValueColor
        )}>
          {formatBDT(value)}
        </h3>
        {((borrowedValue ?? 0) > 0 || (lentValue ?? 0) > 0) && (
          <div className="flex flex-wrap items-center gap-2 mb-3 mt-[-6px]">
            {(borrowedValue ?? 0) > 0 && (
              <span className="text-[9px] bg-rose-500/10 text-rose-400 px-1.5 py-0.5 rounded border border-rose-500/20 font-bold uppercase tracking-tight">
                Borrowed: -{formatBDT(borrowedValue ?? 0)}
              </span>
            )}
            {(lentValue ?? 0) > 0 && (
              <span className="text-[9px] bg-emerald-500/10 text-emerald-400 px-1.5 py-0.5 rounded border border-emerald-500/20 font-bold uppercase tracking-tight">
                Lent: +{formatBDT(lentValue ?? 0)}
              </span>
            )}
          </div>
        )}
      </div>

      {!hideFooter && (
        <div className="flex flex-col justify-center gap-0.5 pt-4 border-t border-slate-800 h-14">
          {trend !== undefined && trend !== null ? (
            <>
              <p className="text-label font-medium text-slate-300 uppercase">{resolvedTrendLabel}</p>
              <p className={cn(
                "text-body font-bold tabular-nums flex items-center gap-0.5",
                trend >= 0 ? 'text-emerald-400' : 'text-rose-500'
              )}>
                {trend >= 0 ? <ArrowUpRight size={13} /> : <ArrowDownRight size={13} />}
                {Math.abs(trend).toFixed(2)}%
              </p>
            </>
          ) : footerLabel !== undefined ? (
            <>
              <p className="text-label font-medium text-slate-300 uppercase">{footerLabel}</p>
              <p className={cn("text-body font-bold tabular-nums", footerValueColor)}>
                {footerValue !== null && footerValue !== undefined ? formatBDT(footerValue) : '—'}
              </p>
            </>
          ) : (
            <div className="h-full" />
          )}
        </div>
      )}
    </Card>
  );
};

// ─── Left axis config (area / net worth) ──────────────────────────────────────
function getAxisConfig(data: { value: number }[]): {
  domain: [number, number];
  ticks: number[];
  tickFormatter: (v: number) => string;
} {
  const fallback = {
    domain: [0, 100] as [number, number],
    ticks: [0, 25, 50, 75, 100],
    tickFormatter: (v: number) => `৳${v}`,
  };

  if (!data || data.length === 0) return fallback;
  const values = data.map(d => d.value).filter(v => typeof v === 'number' && isFinite(v));
  if (values.length === 0) return fallback;

  const rawMin = Math.min(...values);
  const rawMax = Math.max(...values);
  const span = rawMax - rawMin || rawMax || 1;

  const L = 100_000;
  const candidates = [
    0.1 * L, 0.25 * L, 0.5 * L, 1 * L, 2 * L, 2.5 * L,
    5 * L, 10 * L, 25 * L, 50 * L, 100 * L, 250 * L,
    500 * L, 1000 * L,
  ];
  let step = candidates[candidates.length - 1];
  for (const c of candidates) {
    if (span / c <= 8) { step = c; break; }
  }

  const niceMin = Math.floor(rawMin / step) * step;
  const niceMax = Math.ceil(rawMax  / step) * step;
  const niceSpan = niceMax - niceMin || step;

  const pad = niceSpan * 0.20;
  const paddedMin = Math.floor((niceMin - pad) / step) * step;
  // Top: just one step above the data ceiling for minimal breathing room
  const paddedMax = niceMax + step;

  const ticks: number[] = [];
  const eps = step * 0.001;
  for (let t = paddedMin; t <= paddedMax + eps; t = Math.round((t + step) / step) * step) {
    ticks.push(Math.round(t / step) * step);
  }

  const absMax = Math.max(Math.abs(paddedMin), Math.abs(paddedMax));
  let tickFormatter: (v: number) => string;
  if (absMax >= 100 * L) {
    tickFormatter = (v: number) => `${(v / (L * 100)).toFixed(v % (L * 100) === 0 ? 0 : 1)}Cr`;
  } else {
    tickFormatter = (v: number) => {
      const inL = v / L;
      return `${Number.isInteger(inL) ? inL.toFixed(0) : inL.toFixed(1)}L`;
    };
  }

  return { domain: [paddedMin, paddedMax], ticks, tickFormatter };
}

// ─── Accounts Trend Axis Config (2 decimal places, K unit for IBBL Savings and City FCY) ──
function getAccountsAxisConfig(data: { value: number; investment?: number; expense?: number }[], tab: string): {
  domain: [number, number];
  ticks: number[];
  tickFormatter: (v: number) => string;
} {
  const isIbbl = tab === 'IBBL Savings';
  const isCityFcy = tab === 'City FCY';
  const isKScale = isIbbl || isCityFcy;
  const K = 1_000;
  const L = 100_000;

  const fallback = {
    domain: [0, 100] as [number, number],
    ticks: [0, 25, 50, 75, 100],
    tickFormatter: (v: number) => isCityFcy ? `$${v.toFixed(2)}` : `${v.toFixed(2)}`,
  };

  if (!data || data.length === 0) return fallback;
  const values: number[] = [];
  data.forEach(d => {
    if (typeof d.value === 'number' && isFinite(d.value)) values.push(d.value);
    if (tab === 'Total (BDT)' && typeof d.investment === 'number' && isFinite(d.investment)) values.push(d.investment);
    if ((tab === 'Cash Flow (BDT)' || tab === 'Savings (BDT)') && typeof d.expense === 'number' && isFinite(d.expense)) values.push(d.expense);
  });
  if (values.length === 0) return fallback;

  const rawMin = Math.min(...values);
  const rawMax = Math.max(...values);
  const span = rawMax - rawMin || rawMax || 1;

  const candidates = isKScale
    ? [
        0.5 * K, 1 * K, 2 * K, 2.5 * K, 5 * K, 10 * K, 20 * K, 25 * K, 50 * K,
        100 * K, 200 * K, 250 * K, 500 * K, 1000 * K, 2500 * K, 5000 * K,
      ]
    : [
        0.05 * L, 0.1 * L, 0.25 * L, 0.5 * L, 1 * L, 2 * L, 2.5 * L,
        5 * L, 10 * L, 25 * L, 50 * L, 100 * L, 250 * L,
        500 * L, 1000 * L,
      ];

  let step = candidates[candidates.length - 1];
  for (const c of candidates) {
    if (span / c <= 8) { step = c; break; }
  }

  // Calculate starting (domainMin) and ending (domainMax) values so that
  // the data range between lowest and highest value occupies ~95% of graph height
  // for Total (BDT), and ~80% for other tabs.
  // NOTE: for Total (BDT) we do NOT snap the padded bounds to the nearest step —
  // that snapping was adding far more than 5% of extra headroom.
  const pad = tab === 'Total (BDT)' ? span * 0.05 : span * 0.10;
  let domainMin: number;
  let domainMax: number;

  if (tab === 'Total (BDT)') {
    domainMin = rawMin - pad;
    domainMax = rawMax + pad;
  } else {
    domainMin = Math.floor((rawMin - pad) / step) * step;
    domainMax = Math.ceil((rawMax + pad) / step) * step;
  }

  // Keep non-negative if data has no negative values
  if (rawMin >= 0 && domainMin < 0) {
    domainMin = 0;
  }

  if (domainMax <= domainMin) {
    domainMax = domainMin + step;
  }

  const ticks: number[] = [];
  const eps = step * 0.001;
  for (let t = domainMin; t <= domainMax + eps; t = Math.round((t + step) / step) * step) {
    ticks.push(Math.round(t / step) * step);
  }

  const absMax = Math.max(Math.abs(domainMin), Math.abs(domainMax));

  const tickFormatter = (v: number) => {
    if (isCityFcy) {
      if (absMax >= 1_000_000) {
        return `$${(v / 1_000_000).toFixed(2)} M`;
      } else if (absMax >= K) {
        return `$${(v / K).toFixed(2)} K`;
      }
      return `$${v.toFixed(2)}`;
    }

    if (isIbbl) {
      if (absMax >= 100 * L) {
        return `${(v / (100 * L)).toFixed(2)} Cr`;
      }
      return `${(v / K).toFixed(2)} K`;
    }

    if (absMax >= 100 * L) {
      return `${(v / (100 * L)).toFixed(2)} Cr`;
    } else if (absMax >= L) {
      return `${(v / L).toFixed(2)} L`;
    } else if (absMax >= K) {
      return `${(v / K).toFixed(2)} K`;
    } else {
      return `${v.toFixed(2)}`;
    }
  };

  return { domain: [domainMin, domainMax], ticks, tickFormatter };
}

// ─── Right axis config (profit — fully independent domain & yAxisId) ──────────
// The profit Line uses yAxisId="right" with its OWN domain, completely separate
// from the left axis.  To ensure the profit line always renders BELOW the area
// curve, we inflate the right-axis domain ceiling:
//
//   domainMax = paddedProfitMax / VISUAL_SCALE   (VISUAL_SCALE = 0.82)
//
// This makes Recharts believe there is more headroom above the data than there
// really is, so the profit peak only reaches 82 % of chart height.
interface ProfitAxisConfig {
  domain: [number, number];
  ticks: number[];
  tickFormatter: (v: number) => string;
}

function getProfitAxisConfig(profitData: { profit: number }[]): ProfitAxisConfig {
  const fallback: ProfitAxisConfig = {
    domain: [0, 1] as [number, number],
    ticks: [],
    tickFormatter: () => '',
  };

  if (!profitData || profitData.length === 0) return fallback;
  const profits = profitData.map(d => d.profit).filter(v => isFinite(v));
  if (profits.length === 0) return fallback;

  const rawProfitMax = Math.max(...profits);
  const rawProfitMin = Math.min(...profits);
  if (rawProfitMax === 0 && rawProfitMin === 0) return fallback;

  const profitSpan = rawProfitMax - rawProfitMin || rawProfitMax || 1;

  // Step candidates: start from 1K so small profit values get K-scale ticks
  const K = 1_000;
  const L = 100_000;
  const candidates = [
    1*K, 2*K, 5*K, 10*K, 20*K, 25*K, 50*K,
    100*K, 250*K, 500*K,
    1*L, 2*L, 5*L, 10*L, 25*L, 50*L, 100*L,
    250*L, 500*L, 1000*L,
  ];
  let step = candidates[candidates.length - 1];
  for (const c of candidates) {
    if (profitSpan / c <= 6) { step = c; break; }
  }

  const niceProfitMin = Math.floor(rawProfitMin / step) * step;
  const niceProfitMax = Math.ceil(rawProfitMax  / step) * step;
  const niceProfitSpan = niceProfitMax - niceProfitMin || step;

  // 20% bottom padding, one-step top headroom (mirrors left axis logic)
  const pad = niceProfitSpan * 0.20;
  const paddedProfitMin = Math.floor((niceProfitMin - pad) / step) * step;
  const paddedProfitMax = niceProfitMax + step;

  // Inflate the domain ceiling so the profit peak sits at ≤82 % of chart height
  const VISUAL_SCALE = 1.0;
  const domainMax = paddedProfitMax / VISUAL_SCALE;
  const domainMin = paddedProfitMin;

  // Ticks only go up to paddedProfitMax (real data region), not domainMax
  const ticks: number[] = [];
  const eps = step * 0.001;
  for (let t = paddedProfitMin; t <= paddedProfitMax + eps; t = Math.round((t + step) / step) * step) {
    ticks.push(Math.round(t / step) * step);
  }

  // Formatter: K / L / Cr based on magnitude
  const absMax = Math.max(Math.abs(domainMin), Math.abs(paddedProfitMax));
  let tickFormatter: (v: number) => string;
  if (absMax >= 100 * L) {
    tickFormatter = (v: number) =>
      `${(v / (L * 100)).toFixed(v % (L * 100) === 0 ? 0 : 1)}Cr`;
  } else if (absMax >= L) {
    tickFormatter = (v: number) => {
      const inL = v / L;
      return `${Number.isInteger(inL) ? inL.toFixed(0) : inL.toFixed(1)}L`;
    };
  } else {
    tickFormatter = (v: number) => {
      const inK = v / K;
      return `${Number.isInteger(inK) ? inK.toFixed(0) : inK.toFixed(1)}K`;
    };
  }

  return { domain: [domainMin, domainMax], ticks, tickFormatter };
}

// ─── Types ────────────────────────────────────────────────────────────────────
interface TrendPoint {
  name: string;
  value: number;
  profit?: number;
  rawValue?: number;
  investment?: number;
}

// ── SVG arc helpers ───────────────────────────────────────────────────────────
function dashToXY(cx: number, cy: number, r: number, angleDeg: number) {
  const rad = (angleDeg * Math.PI) / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

// ─── Grid Aspect Ratio Hook (Main grid length:width ratio 3:1) ──────────────
function useChartGridDimensions({
  leftYAxisWidth = 56,
  rightYAxisWidth = 0,
  marginTop = 20,
  marginBottom = 8,
  marginLeft = 12,
  marginRight = 12,
  xAxisHeight = 28,
  ratio = 3,
}: {
  leftYAxisWidth?: number;
  rightYAxisWidth?: number;
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  xAxisHeight?: number;
  ratio?: number;
}) {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [containerWidth, setContainerWidth] = React.useState<number>(800);

  React.useEffect(() => {
    if (!containerRef.current) return;
    const el = containerRef.current;
    const update = () => {
      if (el) {
        const w = el.clientWidth;
        if (w > 0) setContainerWidth(w);
      }
    };
    update();
    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        const w = entry.contentRect.width;
        if (w > 0) setContainerWidth(w);
      }
    });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  const gridWidth = Math.max(
    80,
    containerWidth - (leftYAxisWidth + rightYAxisWidth + marginLeft + marginRight)
  );
  const gridHeight = gridWidth / ratio;
  const totalHeight = Math.round(gridHeight + marginTop + marginBottom + xAxisHeight);

  return { containerRef, totalHeight, gridWidth, gridHeight };
}

const DashboardTooltip = ({ active, payload, label, prefix = '৳' }: any) => {
  if (!active || !payload?.length) return null;

  const isSavingsTab = payload.some((p: any) => p.name === 'Income') && payload.some((p: any) => p.name === 'Expense');
  const incomeVal = payload.find((p: any) => p.name === 'Income')?.value || 0;
  const expenseVal = payload.find((p: any) => p.name === 'Expense')?.value || 0;
  const netSavings = incomeVal - expenseVal;

  // De-duplicate entries sharing the same name (an Area + Line pair for the
  // same series otherwise both show up in the tooltip).
  const seenNames = new Set<string>();
  const uniquePayload = payload.filter((item: any) => {
    const key = item.name || item.dataKey;
    if (seenNames.has(key)) return false;
    seenNames.add(key);
    return true;
  });

  return (
    <div className="bg-[#0f172a]/95 backdrop-blur-md border border-slate-700/60 p-4 rounded-xl shadow-2xl min-w-[220px]">
      <p className="text-[12px] font-bold text-white mb-3 border-b border-slate-700/60 pb-2 uppercase tracking-widest">
        {label}
      </p>
      <div className="space-y-2.5">
        {uniquePayload.map((item: any, index: number) => (
          <div key={index} className="flex items-center justify-between gap-8">
            <div className="flex items-center gap-2">
              <div 
                className="w-2 h-2 rounded-full" 
                style={{ backgroundColor: item.color || item.stroke || item.fill || '#2dd4bf' }} 
              />
              <span className="text-[11px] font-bold text-slate-400 uppercase tracking-tight">
                {item.name || item.dataKey}
              </span>
            </div>
            <span className="text-[11px] font-black text-white tabular-nums">
              {prefix}{item.value?.toLocaleString('en-IN')}
            </span>
          </div>
        ))}

        {isSavingsTab && (
          <div className="flex items-center justify-between gap-8 pt-2 border-t border-slate-700/40">
            <div className="flex items-center gap-2">
              <span className="text-[11px] font-bold text-slate-300 uppercase tracking-tight">
                Net Savings
              </span>
            </div>
            <span className={cn("text-[11px] font-black tabular-nums", netSavings >= 0 ? "text-teal-400" : "text-rose-400")}>
              {netSavings >= 0 ? '+' : '-'}{prefix}{Math.abs(Math.round(netSavings)).toLocaleString('en-IN')}
            </span>
          </div>
        )}
      </div>
    </div>
  );
};

function dashBuildArcPath(
  cx: number, cy: number,
  outerR: number, innerR: number,
  startDeg: number, sweepDeg: number,
): string {
  const sw  = Math.min(Math.max(sweepDeg, 0.01), 359.9999);
  const end = startDeg - sw;
  const lg  = sw > 180 ? 1 : 0;
  const os = dashToXY(cx, cy, outerR, startDeg);
  const oe = dashToXY(cx, cy, outerR, end);
  const ie = dashToXY(cx, cy, innerR, end);
  const is = dashToXY(cx, cy, innerR, startDeg);
  return (
    `M ${os.x.toFixed(3)} ${os.y.toFixed(3)} ` +
    `A ${outerR} ${outerR} 0 ${lg} 0 ${oe.x.toFixed(3)} ${oe.y.toFixed(3)} ` +
    `L ${ie.x.toFixed(3)} ${ie.y.toFixed(3)} ` +
    `A ${innerR} ${innerR} 0 ${lg} 1 ${is.x.toFixed(3)} ${is.y.toFixed(3)} Z`
  );
}

const DashboardDonutChart: React.FC<{
  slices: { name: string; value: number; color: string }[];
  centerLabel?: string;
  subLabel?: string;
  size?: number;
}> = ({ slices, centerLabel = 'PORTFOLIO', subLabel = 'OF PORTFOLIO', size = 260 }) => {
  const [hov, setHov] = React.useState<number | null>(null);
  const CX = 150, CY = 150, OUTER = 126, INNER = 68, GAP = 1.4;

  const total = slices.reduce((s, i) => s + i.value, 0);

  const arcs = React.useMemo(() => {
    let cursor = 0;
    return slices.map(s => {
      const pct = total > 0 ? s.value / total : 0;
      const totalSweep = pct * 360;
      const start = cursor;
      const sweep = Math.max(0, totalSweep - GAP);
      cursor -= totalSweep;
      return { ...s, pct, start, sweep };
    });
  }, [slices, total]);

  const hovSlice = hov !== null ? arcs[hov] : null;

  const fmtBDT = (n: number) => {
    const abs = Math.abs(Math.round(n * 100) / 100);
    const fixed = abs.toFixed(0);
    const digits = fixed.replace(/^0+(?=\d)/, '');
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
    return `৳${result}`;
  };

  return (
    <svg
      viewBox="0 0 300 300"
      style={{ width: '100%', maxWidth: size, height: 'auto', aspectRatio: '1/1', flexShrink: 0, overflow: 'visible' }}
    >
      {arcs.map((arc, i) => {
        if (arc.sweep <= 0) return null;
        const isHov  = hov === i;
        const dimmed = hov !== null && !isHov;
        return (
          <path
            key={arc.name}
            d={dashBuildArcPath(CX, CY, OUTER, INNER, arc.start, arc.sweep)}
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
          <text x={CX} y={CY - 22} textAnchor="middle" fill="#e2e8f0"
            fontSize={12} fontWeight={700} letterSpacing={1}>{hovSlice.name}</text>
          <text x={CX} y={CY + 4} textAnchor="middle" fill={hovSlice.color}
            fontSize={22} fontWeight={800}>{(hovSlice.pct * 100).toFixed(1)}%</text>
          <text x={CX} y={CY + 22} textAnchor="middle" fill="#94a3b8"
            fontSize={10} fontWeight={600}>{fmtBDT(hovSlice.value)}</text>
          <text x={CX} y={CY + 36} textAnchor="middle" fill="#475569"
            fontSize={9} fontWeight={600}>{subLabel}</text>
        </>
      ) : (
        <>
          <text x={CX} y={CY - 8} textAnchor="middle" fill="#64748b"
            fontSize={10} fontWeight={700} letterSpacing={0.5}>{centerLabel}</text>
          <text x={CX} y={CY + 14} textAnchor="middle" fill="#e2e8f0"
            fontSize={16} fontWeight={800}>{slices.length} {centerLabel === 'PORTFOLIO' ? 'ASSETS' : 'ITEMS'}</text>
        </>
      )}
    </svg>
  );
};

export const DashboardCharts: React.FC<{
  pieData?: { name: string; value: number }[];
  assetAllocationData?: { name: string; value: number; color?: string }[];
  lineDataMap?: Record<string, TrendPoint[]>;
  accountsLineDataMap?: Record<string, TrendPoint[]>;
  barData?: { name: string; income: number; expense: number }[];
  seriesMetadata?: Record<string, { min: number; max: number; range: number }>;
}> = ({
  pieData = [],
  assetAllocationData = [],
  lineDataMap = {},
  accountsLineDataMap = {},
  barData = [],
  seriesMetadata
}) => {
  // ── Accounts Trend Tab & Pagination ────────────────────────────────────────
  const ACCOUNTS_TAB_DISPLAY = ['Cash Flow (BDT)', 'Total (BDT)', 'City Islamic', 'IBBL Savings', 'City FCY'];
  const [accountsActiveTab, setAccountsActiveTab] = React.useState('Cash Flow (BDT)');
  const [accountsCurrentPage, setAccountsCurrentPage] = React.useState(0);

  // ── Investment Trend (formerly Total Trend) Tab & Pagination ───────────────
  const TAB_DISPLAY  = ['Total', 'DSE', 'Mutual', 'FD', 'Sukuk', 'Online'];
  const TAB_DATA_KEYS = ['Total', 'DSE Tracker', 'Mutual Funds', 'Fixed Deposits', 'Sukuk Funds', 'Online Invests'];

  const [activeTab, setActiveTab]       = React.useState('Total');
  const [currentPage, setCurrentPage]   = React.useState(0);
  const [isMobile, setIsMobile]         = React.useState(window.innerWidth < 640);

  React.useEffect(() => {
    const handleResize = () => setIsMobile(window.innerWidth < 640);
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const itemsPerPage = isMobile ? 6 : 12;
  const totalValue   = pieData.reduce((sum, item) => sum + item.value, 0);

  // ── Asset Allocation Slice Preparation ────────────────────────────────────
  const ASSET_COLORS: Record<string, string> = {
    Expense: '#ef4444',
    Saving: '#2dd4bf',
    Investment: '#3b82f6',
  };

  const displayAssetSlices = (assetAllocationData.length > 0 ? assetAllocationData : [
    { name: 'Expense', value: 0, color: '#ef4444' },
    { name: 'Saving', value: 0, color: '#2dd4bf' },
    { name: 'Investment', value: 0, color: '#3b82f6' },
  ]).map((item, idx) => ({
    name: item.name,
    value: item.value,
    color: item.color || ASSET_COLORS[item.name] || COLORS[idx % COLORS.length],
  }));

  const totalAssetValue = displayAssetSlices.reduce((s, i) => s + i.value, 0);

  // ── Accounts Trend Data calculations ───────────────────────────────────────
  const accountsLineData = accountsLineDataMap[accountsActiveTab] || (accountsActiveTab === 'Cash Flow (BDT)' ? accountsLineDataMap['Savings (BDT)'] : []) || [];
  const accountsTotalPoints = accountsLineData.length;
  const accountsMaxPages = Math.max(1, Math.ceil(accountsTotalPoints / itemsPerPage));

  const displayAccountsData = React.useMemo(() => {
    const end = accountsTotalPoints - (accountsCurrentPage * itemsPerPage);
    const start = Math.max(0, end - itemsPerPage);
    return accountsLineData.slice(start, end);
  }, [accountsLineData, accountsCurrentPage, accountsTotalPoints, itemsPerPage]);

  const accountsAxisConfig = React.useMemo(
    () => getAccountsAxisConfig(displayAccountsData, accountsActiveTab),
    [displayAccountsData, accountsActiveTab]
  );

  const accountsChartTitle = accountsActiveTab === 'Total (BDT)'
    ? 'Accounts Trend'
    : accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)'
    ? 'Cash Flow (BDT) Trend'
    : `${accountsActiveTab} Trend`;

  // ── Investment Trend Data calculations ─────────────────────────────────────
  const activeDataKey = TAB_DATA_KEYS[TAB_DISPLAY.indexOf(activeTab)] ?? activeTab;
  const lineData      = lineDataMap[activeDataKey] || [];

  const totalPoints = lineData.length;
  const maxPages    = Math.max(1, Math.ceil(totalPoints / itemsPerPage));

  const displayData = React.useMemo(() => {
    const end   = totalPoints - (currentPage * itemsPerPage);
    const start = Math.max(0, end - itemsPerPage);
    return lineData.slice(start, end);
  }, [lineData, currentPage, totalPoints, itemsPerPage]);

  // ── Left axis (area / value) ───────────────────────────────────────────────
  const axisConfig = React.useMemo(() => getAxisConfig(displayData), [displayData]);

  // ── Right axis (profit — independent domain) ─────────────────────────────
  const profitAxisConfig = React.useMemo(
    () => getProfitAxisConfig(displayData.map(d => ({ profit: d.profit ?? 0 }))),
    [displayData],
  );

  // profit field used directly by the Line via yAxisId="right"
  const chartData = displayData;

  const chartTitle = activeTab === 'Total'
    ? 'Investment Trend'
    : `${activeTab} Trend`;

  const hasProfitData = displayData.some(d => (d.profit ?? 0) !== 0);

  // ── 3:1 Grid Aspect Ratio Setup for Accounts Trend ────────────────────────
  const accountsLeftYAxisWidth = isMobile ? 46 : 56;
  const accountsXAxisHeight = isMobile ? 24 : 28;
  const accountsMargin = React.useMemo(() => ({
    top: isMobile ? 12 : 20,
    right: isMobile ? 8 : 12,
    left: isMobile ? 8 : 12,
    bottom: isMobile ? 6 : 8,
  }), [isMobile]);

  const accountsRatio = accountsActiveTab === 'Total (BDT)' ? 2.5 : 3;
  const { containerRef: accountsContainerRef, totalHeight: accountsChartHeight } = useChartGridDimensions({
    leftYAxisWidth: accountsLeftYAxisWidth,
    rightYAxisWidth: 0,
    marginTop: accountsMargin.top,
    marginBottom: accountsMargin.bottom,
    marginLeft: accountsMargin.left,
    marginRight: accountsMargin.right,
    xAxisHeight: accountsXAxisHeight,
    ratio: accountsRatio,
  });

  // ── 3:1 Grid Aspect Ratio Setup for Investment Trend ──────────────────────
  const investmentLeftYAxisWidth = isMobile ? 46 : 56;
  const investmentRightYAxisWidth = hasProfitData ? (isMobile ? 46 : 56) : 0;
  const investmentXAxisHeight = isMobile ? 24 : 28;
  const investmentMargin = React.useMemo(() => ({
    top: isMobile ? 12 : 20,
    right: isMobile ? 8 : 12,
    left: isMobile ? 8 : 12,
    bottom: isMobile ? 6 : 8,
  }), [isMobile]);

  const { containerRef: investmentContainerRef, totalHeight: investmentChartHeight } = useChartGridDimensions({
    leftYAxisWidth: investmentLeftYAxisWidth,
    rightYAxisWidth: investmentRightYAxisWidth,
    marginTop: investmentMargin.top,
    marginBottom: investmentMargin.bottom,
    marginLeft: investmentMargin.left,
    marginRight: investmentMargin.right,
    xAxisHeight: investmentXAxisHeight,
    ratio: 4,
  });

  return (
    <div className="flex flex-col gap-6 sm:gap-8 w-full">

      {/* ── Row 1: Accounts Trend (70%) + Asset Allocation (30%) ─────────────── */}
      <div className="flex flex-col lg:flex-row gap-6 sm:gap-8 w-full items-stretch">
        <div className="w-full lg:w-[70%] flex flex-col">
          {/* ── Accounts Trend Chart (Single Left Axis, No Secondary Axis) ─────── */}
          <Card className="w-full h-full p-1 sm:p-6 border-slate-800 bg-slate-900 shadow-2xl relative group overflow-hidden flex flex-col justify-between">
        <div className="flex flex-col mb-2 sm:mb-4 relative z-10 gap-3 sm:gap-4 ml-1 mr-1">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 px-1 sm:px-0 pb-4 border-b border-slate-800">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-teal-400/10 flex items-center justify-center text-teal-400 shadow-[0_0_20px_rgba(45,212,191,0.1)] group-hover:scale-110 transition-transform duration-500">
                <Landmark size={20} />
              </div>
              <div className="min-w-0">
                <h3 className="text-[10px] font-bold text-slate-500 uppercase tracking-[0em] mb-0.5 truncate">Accounts Balance Trend</h3>
                <h2 className="text-base sm:text-lg font-bold text-white tracking-tight uppercase truncate">{accountsChartTitle}</h2>
              </div>
            </div>

            <div className="flex items-center justify-between sm:justify-end w-full sm:w-auto gap-3">
              {/* Mobile: Pagination + Legend */}
              <div className="flex sm:hidden items-center gap-3">
                {/* Pagination */}
                <div className="flex items-center bg-slate-950 border border-slate-800 rounded-lg p-0.5 shrink-0">
                  <button
                    onClick={() => setAccountsCurrentPage(prev => Math.min(prev + 1, accountsMaxPages - 1))}
                    disabled={accountsCurrentPage >= accountsMaxPages - 1}
                    className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
                    title="Previous"
                  >
                    <ChevronLeft size={16} />
                  </button>

                  <div className="w-px h-3 bg-slate-800 mx-0.5" />

                  <button
                    onClick={() => setAccountsCurrentPage(prev => Math.max(prev - 1, 0))}
                    disabled={accountsCurrentPage <= 0}
                    className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
                    title="Next"
                  >
                    <ChevronRight size={16} />
                  </button>
                </div>

                {/* Mobile Legend */}
                <div className="flex items-center gap-3">
                  <div className="flex items-center gap-1.5">
                    <div className="w-3 h-0.5 bg-teal-400 rounded-full" />
                    <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
                      {accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)' ? 'INCOME' : 'BALANCE'}
                    </span>
                  </div>

                  {(accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)') && (
                    <div className="flex items-center gap-1.5">
                      <div className="w-3 h-0.5 bg-rose-500 rounded-full" />
                      <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
                        EXPENSE
                      </span>
                    </div>
                  )}

                  {accountsActiveTab === 'Total (BDT)' && (
                    <div className="flex items-center gap-1.5">
                      <div className="w-3 h-0.5 bg-[#be185d] rounded-full shadow-[0_0_8px_rgba(190,24,93,0.6)]" />
                      <span className="text-[9px] font-bold text-pink-500 uppercase tracking-widest">
                        INVESTMENT
                      </span>
                    </div>
                  )}
                </div>
              </div>

              {/* Desktop Legend */}
              <div className="hidden sm:flex items-center gap-6">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-0.5 bg-teal-400 rounded-full" />
                  <span className="text-[12px] font-bold text-slate-300 uppercase tracking-tight">
                    {accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)' ? 'INCOME' : 'BALANCE'}
                  </span>
                </div>

                {(accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)') && (
                  <div className="flex items-center gap-2">
                    <div className="w-3 h-0.5 bg-rose-500 rounded-full" />
                    <span className="text-[12px] font-bold text-slate-300 uppercase tracking-tight">
                      EXPENSE
                    </span>
                  </div>
                )}

                {accountsActiveTab === 'Total (BDT)' && (
                  <div className="flex items-center gap-2">
                    <div className="w-3 h-0.5 bg-[#be185d] rounded-full shadow-[0_0_8px_rgba(190,24,93,0.6)]" />
                    <span className="text-[12px] font-bold text-pink-500 uppercase tracking-tight">
                      INVESTMENT
                    </span>
                  </div>
                )}
              </div>

              {/* Desktop Pagination */}
              <div className="hidden sm:flex items-center gap-2">
                <div className="flex items-center bg-slate-950 border border-slate-800 rounded-lg p-0.5">
                  <button
                    onClick={() => setAccountsCurrentPage(prev => Math.min(prev + 1, accountsMaxPages - 1))}
                    disabled={accountsCurrentPage >= accountsMaxPages - 1}
                    className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
                    title="Previous"
                  >
                    <ChevronLeft size={16} />
                  </button>

                  <div className="w-px h-3 bg-slate-800 mx-0.5" />

                  <button
                    onClick={() => setAccountsCurrentPage(prev => Math.max(prev - 1, 0))}
                    disabled={accountsCurrentPage <= 0}
                    className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
                    title="Next"
                  >
                    <ChevronRight size={16} />
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* Accounts Tabs */}
          <div className="flex justify-center w-full mt-2 sm:mt-0">
            <div className="flex items-center h-10 bg-slate-950 rounded-xl border border-slate-800 overflow-x-auto sm:overflow-hidden no-scrollbar shadow-inner shadow-black/20 max-w-[calc(100vw-2rem)] sm:max-w-none">
              {ACCOUNTS_TAB_DISPLAY.map((tab, idx) => (
                <button
                  key={tab}
                  onClick={() => { setAccountsActiveTab(tab); setAccountsCurrentPage(0); }}
                  className={cn(
                    "flex items-center gap-2 px-4 sm:px-5 h-full text-[10px] font-bold uppercase transition-all tracking-wider whitespace-nowrap",
                    idx < ACCOUNTS_TAB_DISPLAY.length - 1 ? "border-r border-slate-800" : "",
                    accountsActiveTab === tab
                      ? "bg-teal-400 text-slate-950 shadow-[inset_0_1px_1px_rgba(255,255,255,0.4)]"
                      : "text-slate-400 hover:text-white hover:bg-slate-800/60"
                  )}
                >
                  <span className={cn(
                    "w-2 h-2 rounded-full shrink-0 transition-all duration-300",
                    accountsActiveTab === tab ? "bg-slate-950/60 scale-110" : "bg-teal-400/50"
                  )} />
                  {tab}
                </button>
              ))}
            </div>
          </div>
        </div>

        <div ref={accountsContainerRef} className="w-full relative z-10" style={{ height: accountsChartHeight }}>
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart 
              data={displayAccountsData} 
              margin={accountsMargin}
            >
              <defs>
                <linearGradient id="colorAccountValue" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#2dd4bf" stopOpacity={0.8} />
                  <stop offset="95%" stopColor="#2dd4bf" stopOpacity={0} />
                </linearGradient>
                <linearGradient id="colorSavingsExpense" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#ef4444" stopOpacity={0.85} />
                  <stop offset="95%" stopColor="#881337" stopOpacity={0.5} />
                </linearGradient>
                <linearGradient id="colorTotalInvestment" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#be185d" stopOpacity={0.6} />
                  <stop offset="95%" stopColor="#be185d" stopOpacity={0.05} />
                </linearGradient>
              </defs>

              <CartesianGrid stroke="#334155" strokeOpacity={0.4} vertical={true} horizontal={true} strokeDasharray="0" />

              <XAxis
                dataKey="name"
                height={accountsXAxisHeight}
                stroke="#94a3b8"
                fontSize={11}
                tickLine={false}
                axisLine={false}
                dy={8}
                padding={{ left: 0, right: 0 }}
                className="font-bold uppercase tracking-widest"
                tick={{ fill: '#94a3b8', fontSize: 11, fontWeight: 700, letterSpacing: '0.05em' }}
              />

              {/* Single LEFT axis only - No Right Axis */}
              <YAxis
                yAxisId="left"
                width={accountsLeftYAxisWidth}
                stroke="#475569"
                fontSize={10}
                fontWeight={700}
                tickLine={false}
                axisLine={false}
                dx={-4}
                domain={accountsAxisConfig.domain}
                ticks={accountsAxisConfig.ticks}
                tickFormatter={accountsAxisConfig.tickFormatter}
                tick={{ fill: '#475569', fontSize: 10, fontWeight: 700 }}
              />

              {accountsAxisConfig.ticks.map((tick) => (
                <ReferenceLine
                  key={tick}
                  y={tick}
                  yAxisId="left"
                  stroke="#334155"
                  strokeOpacity={0.4}
                  strokeWidth={1}
                  ifOverflow="visible"
                />
              ))}

              <Tooltip
                cursor={{ stroke: '#334155', strokeWidth: 1, strokeDasharray: '4 4' }}
                content={<DashboardTooltip prefix={accountsActiveTab === 'City FCY' ? '$' : '৳'} />}
              />

              {/* Area — balance or income */}
              <Area
                yAxisId="left"
                type="monotone"
                name={accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)' ? 'Income' : 'Balance'}
                dataKey="value"
                stroke="#2dd4bf"
                strokeWidth={4}
                fillOpacity={1}
                fill="url(#colorAccountValue)"
                isAnimationActive={false}
                dot={{ r: 3.5, fill: '#2dd4bf', stroke: '#0f172a', strokeWidth: 1.5 }}
                activeDot={{ r: 5.5, fill: '#2dd4bf', stroke: '#0f172a', strokeWidth: 2 }}
              >
                {!isMobile && (
                  <LabelList
                    dataKey="value"
                    position="top"
                    offset={15}
                    content={(props: any) => {
                      const { x, y, value } = props;
                      if (value === undefined || value === null) return null;
                      return (
                        <text x={x} y={y - 5} fill="#fff" fontSize={10} fontWeight="bold" textAnchor="middle" className="tabular-nums">
                          {accountsAxisConfig.tickFormatter(value)}
                        </text>
                      );
                    }}
                  />
                )}
              </Area>

              {/* Red Area & Line for Expense when Cash Flow (BDT) is selected */}
              {(accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)') && (
                <Area
                  yAxisId="left"
                  type="monotone"
                  name="Expense"
                  dataKey="expense"
                  stroke="#ef4444"
                  strokeWidth={3}
                  fillOpacity={1}
                  fill="url(#colorSavingsExpense)"
                  isAnimationActive={false}
                  dot={{ r: 3.5, fill: '#ef4444', stroke: '#0f172a', strokeWidth: 1.5 }}
                  activeDot={{ r: 5.5, fill: '#ef4444', stroke: '#0f172a', strokeWidth: 2 }}
                >
                  {!isMobile && (
                    <LabelList
                      dataKey="expense"
                      position="top"
                      offset={15}
                      content={(props: any) => {
                        const { x, y, value } = props;
                        if (value === undefined || value === null) return null;
                        return (
                          <text x={x} y={y - 5} fill="#f87171" fontSize={10} fontWeight="bold" textAnchor="middle" className="tabular-nums">
                            {accountsAxisConfig.tickFormatter(value)}
                          </text>
                        );
                      }}
                    />
                  )}
                </Area>
              )}

              {/* Crisp top Income Line when Cash Flow (BDT) */}
              {(accountsActiveTab === 'Cash Flow (BDT)' || accountsActiveTab === 'Savings (BDT)') && (
                <Line
                  yAxisId="left"
                  type="monotone"
                  name="Income"
                  dataKey="value"
                  stroke="#2dd4bf"
                  strokeWidth={4}
                  dot={{ r: 3.5, fill: '#2dd4bf', stroke: '#0f172a', strokeWidth: 1.5 }}
                  activeDot={{ r: 5.5, fill: '#2dd4bf', stroke: '#0f172a', strokeWidth: 2 }}
                  tooltipType="none"
                  legendType="none"
                  isAnimationActive={false}
                />
              )}

              {/* Dark Pink Area and Dark Pink line for Investment when Total (BDT) is selected */}
              {accountsActiveTab === 'Total (BDT)' && (
                <Area
                  yAxisId="left"
                  type="monotone"
                  name="Investment"
                  dataKey="investment"
                  stroke="#be185d"
                  strokeWidth={2.5}
                  fillOpacity={1}
                  fill="url(#colorTotalInvestment)"
                  dot={{ r: 3.5, fill: '#be185d', stroke: '#0f172a', strokeWidth: 1.5 }}
                  activeDot={{ r: 5.5, fill: '#be185d', stroke: '#0f172a', strokeWidth: 2 }}
                  isAnimationActive={false}
                >
                  {!isMobile && (
                    <LabelList
                      dataKey="investment"
                      position="top"
                      offset={15}
                      content={(props: any) => {
                        const { x, y, value } = props;
                        if (value === undefined || value === null) return null;
                        return (
                          <text x={x} y={y - 5} fill="#f472b6" fontSize={10} fontWeight="bold" textAnchor="middle" className="tabular-nums">
                            {accountsAxisConfig.tickFormatter(value)}
                          </text>
                        );
                      }}
                    />
                  )}
                </Area>
              )}

              {accountsActiveTab === 'Total (BDT)' && (
                <Line
                  yAxisId="left"
                  type="monotone"
                  name="Investment"
                  dataKey="investment"
                  stroke="#be185d"
                  strokeWidth={2.5}
                  dot={{ r: 3.5, fill: '#be185d', stroke: '#0f172a', strokeWidth: 1.5 }}
                  activeDot={{ r: 5.5, fill: '#be185d', stroke: '#0f172a', strokeWidth: 2 }}
                  tooltipType="none"
                  legendType="none"
                  isAnimationActive={false}
                />
              )}
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </Card>
    </div>

    {/* ── Asset Allocation (30%) ──────────────────────────────────────── */}
    <div className="w-full lg:w-[30%] flex flex-col">
      <Card className="w-full h-full p-4 sm:p-6 border-slate-800 bg-slate-900 shadow-2xl relative group overflow-hidden flex flex-col justify-between">
        <div className="flex items-center gap-3 mb-6 pb-4 border-b border-slate-800">
          <div className="w-10 h-10 rounded-xl bg-teal-400/10 flex items-center justify-center text-teal-400 group-hover:scale-110 transition-transform duration-500 shadow-[0_0_20px_rgba(45,212,191,0.1)]">
            <PieIcon size={20} />
          </div>
          <div>
            <h3 className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] mb-0.5">Asset Breakdown</h3>
            <h2 className="text-lg font-bold text-white tracking-tight uppercase">Asset Allocation</h2>
          </div>
        </div>

        <div className="flex flex-col sm:flex-row lg:flex-col items-center gap-5 sm:gap-6 relative z-10 my-auto">
          <div className="w-full sm:w-1/2 lg:w-full flex justify-center shrink-0">
            <DashboardDonutChart
              slices={displayAssetSlices}
              centerLabel="ALLOCATION"
              subLabel="OF ASSETS"
              size={180}
            />
          </div>

          <div className="w-full sm:w-1/2 lg:w-full space-y-3">
            {displayAssetSlices.map((item) => {
              const percent = totalAssetValue > 0 ? (item.value / totalAssetValue) * 100 : 0;
              return (
                <div key={item.name} className="group/item cursor-default">
                  <div className="flex justify-between items-center mb-1">
                    <div className="flex items-center gap-2">
                      <div className="w-2 h-2 rounded-sm shrink-0" style={{ backgroundColor: item.color, boxShadow: `0 0 8px ${item.color}55` }} />
                      <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest group-hover/item:text-white transition-colors">{item.name}</span>
                    </div>
                    <span className="text-[10px] font-black text-white tabular-nums">{percent.toFixed(1)}%</span>
                  </div>
                  <div className="h-1.5 w-full bg-slate-800 rounded-full overflow-hidden">
                    <div
                      className="h-full transition-all duration-1000 ease-out rounded-full shadow-[0_0_10px_rgba(255,255,255,0.1)]"
                      style={{ width: `${percent}%`, backgroundColor: item.color }}
                    />
                  </div>
                  <div className="flex justify-end mt-1">
                    <span className="text-[9px] font-bold text-slate-500 tabular-nums uppercase">{formatBDT(item.value)}</span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </Card>
    </div>
  </div>

  {/* ── Row 2: Investment Trend (70%) + Investment Portfolios (30%) ──────── */}
  <div className="flex flex-col lg:flex-row gap-6 sm:gap-8 w-full items-stretch">
    <div className="w-full lg:w-[70%] flex flex-col">
      {/* ── Investment Trend (formerly Net Worth Trend) Chart ─────────────── */}
      <Card className="w-full h-full p-1 sm:p-6 border-slate-800 bg-slate-900 shadow-2xl relative group overflow-hidden flex flex-col">
        <div className="flex flex-col mb-2 sm:mb-4 relative z-10 gap-3 sm:gap-4 ml-1 mr-1">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 px-1 sm:px-0 pb-4 border-b border-slate-800">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-teal-400/10 flex items-center justify-center text-teal-400 shadow-[0_0_20px_rgba(45,212,191,0.1)] group-hover:scale-110 transition-transform duration-500">
                <TrendingUp size={20} />
              </div>
              <div className="min-w-0">
                <h3 className="text-[10px] font-bold text-slate-500 uppercase tracking-[0em] mb-0.5 truncate">Net Worth vs Profit Trend</h3>
                <h2 className="text-base sm:text-lg font-bold text-white tracking-tight uppercase truncate">{chartTitle}</h2>
              </div>
            </div>

            <div className="flex items-center justify-between sm:justify-end w-full sm:w-auto gap-3">

  {/* Mobile: Pagination + Legend in same line */}
  <div className="flex sm:hidden items-center gap-3">
    
    {/* Pagination */}
    <div className="flex items-center bg-slate-950 border border-slate-800 rounded-lg p-0.5 shrink-0">
      <button
        onClick={() => setCurrentPage(prev => Math.min(prev + 1, maxPages - 1))}
        disabled={currentPage >= maxPages - 1}
        className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
        title="Previous"
      >
        <ChevronLeft size={16} />
      </button>

      <div className="w-px h-3 bg-slate-800 mx-0.5" />

      <button
        onClick={() => setCurrentPage(prev => Math.max(prev - 1, 0))}
        disabled={currentPage <= 0}
        className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
        title="Next"
      >
        <ChevronRight size={16} />
      </button>
    </div>

    {/* Mobile Legend */}
    <div className="flex items-center gap-4">
      <div className="flex items-center gap-1.5">
        <div className="w-3 h-0.5 bg-teal-400 rounded-full" />
        <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
          Net Worth
        </span>
      </div>

      {hasProfitData && (
        <div className="flex items-center gap-1.5">
          <div className="w-3 border-t-2 border-dashed border-amber-400" />
          <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
            Profit
          </span>
        </div>
      )}
    </div>
  </div>

  {/* Desktop Legend */}
  <div className="hidden sm:flex items-center gap-6">
    <div className="flex items-center gap-2">
      <div className="w-3 h-0.5 bg-teal-400 rounded-full" />
      <span className="text-[12px] font-bold text-slate-300 uppercase tracking-tight">
        Net Worth
      </span>
    </div>

    {hasProfitData && (
      <div className="flex items-center gap-2">
        <div className="w-3 border-t-2 border-dashed border-amber-400" />
        <span className="text-[12px] font-bold text-slate-300 uppercase tracking-tight">
          Profit
        </span>
      </div>
    )}
  </div>

  {/* Desktop Pagination */}
  <div className="hidden sm:flex items-center gap-2">
    <div className="flex items-center bg-slate-950 border border-slate-800 rounded-lg p-0.5">
      <button
        onClick={() => setCurrentPage(prev => Math.min(prev + 1, maxPages - 1))}
        disabled={currentPage >= maxPages - 1}
        className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
        title="Previous"
      >
        <ChevronLeft size={16} />
      </button>

      <div className="w-px h-3 bg-slate-800 mx-0.5" />

      <button
        onClick={() => setCurrentPage(prev => Math.max(prev - 1, 0))}
        disabled={currentPage <= 0}
        className="p-1 hover:bg-slate-800 rounded text-slate-400 hover:text-white disabled:opacity-20 disabled:cursor-not-allowed transition-all"
        title="Next"
      >
        <ChevronRight size={16} />
      </button>
    </div>
  </div>

</div>
          </div>

          {/* Category tabs */}
          <div className="flex justify-center w-full mt-2 sm:mt-0">
            <div className="flex items-center h-10 bg-slate-950 rounded-xl border border-slate-800 overflow-x-auto sm:overflow-hidden no-scrollbar shadow-inner shadow-black/20 max-w-[calc(100vw-2rem)] sm:max-w-none">
              {TAB_DISPLAY.map((tab, idx) => (
                <button
                  key={tab}
                  onClick={() => { setActiveTab(tab); setCurrentPage(0); }}
                  className={cn(
                    "flex items-center gap-2 px-4 sm:px-5 h-full text-[10px] font-bold uppercase transition-all tracking-wider whitespace-nowrap",
                    idx < TAB_DISPLAY.length - 1 ? "border-r border-slate-800" : "",
                    activeTab === tab
                      ? "bg-teal-400 text-slate-950 shadow-[inset_0_1px_1px_rgba(255,255,255,0.4)]"
                      : "text-slate-400 hover:text-white hover:bg-slate-800/60"
                  )}
                >
                  <span className={cn(
                    "w-2 h-2 rounded-full shrink-0 transition-all duration-300",
                    activeTab === tab ? "bg-slate-950/60 scale-110" : "bg-teal-400/50"
                  )} />
                  {tab}
                </button>
              ))}
            </div>
          </div>
        </div>

        <div ref={investmentContainerRef} className="w-full relative z-10 flex-1 min-h-[300px]" style={{ minHeight: investmentChartHeight, height: '100%' }}>
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart 
              data={chartData} 
              margin={investmentMargin}
            >
              <defs>
                <linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%"  stopColor="#2dd4bf" stopOpacity={0.8} />
                  <stop offset="95%" stopColor="#2dd4bf" stopOpacity={0}   />
                </linearGradient>
              </defs>

              <CartesianGrid stroke="#334155" strokeOpacity={0.4} vertical={true} horizontal={true} strokeDasharray="0" />

              <XAxis
                dataKey="name"
                height={investmentXAxisHeight}
                stroke="#94a3b8"
                fontSize={11}
                tickLine={false}
                axisLine={false}
                dy={8}
                padding={{ left: 0, right: 0 }}
                className="font-bold uppercase tracking-widest"
                tick={{ fill: '#94a3b8', fontSize: 11, fontWeight: 700, letterSpacing: '0.05em' }}
              />

              {/* LEFT axis — net worth */}
              <YAxis
                yAxisId="left"
                width={investmentLeftYAxisWidth}
                stroke="#475569"
                fontSize={10}
                fontWeight={700}
                tickLine={false}
                axisLine={false}
                dx={-4}
                domain={axisConfig.domain}
                ticks={axisConfig.ticks}
                tickFormatter={axisConfig.tickFormatter}
                tick={{ fill: '#475569', fontSize: 10, fontWeight: 700 }}
              />

              {/* RIGHT axis — profit, independent scale */}
              {hasProfitData && (
                <YAxis
                  yAxisId="right"
                  orientation="right"
                  width={investmentRightYAxisWidth}
                  stroke="#fbbf24"
                  fontSize={10}
                  fontWeight={700}
                  tickLine={false}
                  axisLine={false}
                  dx={4}
                  domain={profitAxisConfig.domain}
                  ticks={profitAxisConfig.ticks}
                  tickFormatter={profitAxisConfig.tickFormatter}
                  tick={{ fill: '#fbbf24', fontSize: 10, fontWeight: 700 }}
                />
              )}

              {axisConfig.ticks.map((tick) => (
                <ReferenceLine
                  key={tick}
                  y={tick}
                  yAxisId="left"
                  stroke="#334155"
                  strokeOpacity={0.4}
                  strokeWidth={1}
                  ifOverflow="visible"
                />
              ))}
              <Tooltip
                cursor={{ stroke: '#334155', strokeWidth: 1, strokeDasharray: '4 4' }}
                content={<DashboardTooltip />}
              />

              {/* Area — net worth */}
              <Area
                yAxisId="left"
                type="monotone"
                dataKey="value"
                stroke="#2dd4bf"
                strokeWidth={4}
                fillOpacity={1}
                fill="url(#colorValue)"
                isAnimationActive={false}
                dot={{ r: 3.5, fill: '#2dd4bf', stroke: '#0f172a', strokeWidth: 1.5 }}
                activeDot={{ r: 5.5, fill: '#2dd4bf', stroke: '#0f172a', strokeWidth: 2 }}
              >
                {!isMobile && (
                  <LabelList
                    dataKey="value"
                    position="top"
                    offset={15}
                    content={(props: any) => {
                      const { x, y, value } = props;
                      return (
                        <text x={x} y={y - 5} fill="#fff" fontSize={10} fontWeight="bold" textAnchor="middle" className="tabular-nums">
                          {axisConfig.tickFormatter(value)}
                        </text>
                      );
                    }}
                  />
                )}
              </Area>

              {/* Profit line — independent right axis */}
              {hasProfitData && (
                <Line
                  yAxisId="right"
                  type="monotone"
                  dataKey="profit"
                  stroke="#fbbf24"
                  strokeWidth={2}
                  strokeDasharray="5 3"
                  dot={{ r: 3, fill: '#fbbf24', stroke: '#0f172a', strokeWidth: 1.5 }}
                  activeDot={{ r: 5, fill: '#fbbf24', stroke: '#0f172a', strokeWidth: 2 }}
                  isAnimationActive={false}
                />
              )}
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </Card>
    </div>

    {/* ── Investment Portfolios (30%) ─────────────────────────────────── */}
    <div className="w-full lg:w-[30%] flex flex-col">
      <Card className="w-full h-full p-4 sm:p-6 border-slate-800 bg-slate-900 shadow-2xl relative group overflow-hidden flex flex-col justify-between">
        <div className="flex items-center gap-3 mb-6 pb-4 border-b border-slate-800">
          <div className="w-10 h-10 rounded-xl bg-purple-400/10 flex items-center justify-center text-purple-400 group-hover:scale-110 transition-transform duration-500 shadow-[0_0_20px_rgba(168,85,247,0.1)]">
            <PieIcon size={20} />
          </div>
          <div>
            <h3 className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] mb-0.5">Portfolio Strategy</h3>
            <h2 className="text-lg font-bold text-white tracking-tight uppercase">Investment Portfolios</h2>
          </div>
        </div>

        <div className="flex flex-col sm:flex-row lg:flex-col items-center gap-5 sm:gap-6 relative z-10 my-auto">
          <div className="w-full sm:w-1/2 lg:w-full flex justify-center shrink-0">
            <DashboardDonutChart
              slices={pieData.map((item, idx) => ({ name: item.name, value: item.value, color: COLORS[idx % COLORS.length] }))}
              centerLabel="PORTFOLIO"
              subLabel="OF PORTFOLIO"
              size={180}
            />
          </div>

          <div className="w-full sm:w-1/2 lg:w-full space-y-1.5">
            {pieData.map((item, idx) => {
              const percent = totalValue > 0 ? (item.value / totalValue) * 100 : 0;
              const color = COLORS[idx % COLORS.length];
              return (
                <div key={item.name} className="group/item cursor-default">
                  <div className="flex justify-between items-center mb-0.5">
                    <div className="flex items-center gap-2">
                      <div className="w-2 h-2 rounded-sm shrink-0" style={{ backgroundColor: color, boxShadow: `0 0 8px ${color}55` }} />
                      <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest group-hover/item:text-white transition-colors">{item.name}</span>
                    </div>
                    <span className="text-[10px] font-black text-white tabular-nums">{percent.toFixed(1)}%</span>
                  </div>
                  <div className="h-1.5 w-full bg-slate-800 rounded-full overflow-hidden">
                    <div
                      className="h-full transition-all duration-1000 ease-out rounded-full shadow-[0_0_10px_rgba(255,255,255,0.1)]"
                      style={{ width: `${percent}%`, backgroundColor: color }}
                    />
                  </div>
                  <div className="flex justify-end mt-0.5">
                    <span className="text-[9px] font-bold text-slate-500 tabular-nums uppercase">{formatBDT(item.value)}</span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </Card>
    </div>
  </div>

</div>
  );
};