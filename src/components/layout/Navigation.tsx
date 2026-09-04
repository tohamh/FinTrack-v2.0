/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { 
  LayoutDashboard, 
  TrendingUp, 
  PieChart, 
  Globe, 
  ShieldCheck, 
  CreditCard, 
  ArrowUpRight, 
  ArrowDownRight, 
  BarChart3, 
  Download, 
  Settings,
  ChevronLeft,
  ChevronRight,
  Menu,
  Plus,
  DollarSign,
  Layers,
  Briefcase,
  BarChart2,
  SlidersHorizontal
} from 'lucide-react';
import { cn } from '../../utils/formatters';

function formatLastSynced(ts: number): string {
  if (!ts) return '';
  const diff = Math.floor((Date.now() - ts) / 1000);
  if (diff < 60)   return 'just now';
  if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
  if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
  return new Date(ts).toLocaleDateString();
}

interface SidebarProps {
  activeModule: string;
  setActiveModule: (module: string) => void;
  isCollapsed: boolean;
  setIsCollapsed: (collapsed: boolean) => void;
  isMobileOpen: boolean;
  setIsMobileOpen: (open: boolean) => void;
}

const menuItems = [
  { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { 
    id: 'dse', 
    label: 'DSE Tracker', 
    icon: TrendingUp,
    children: [
      { id: 'dse-summary', label: 'Summary', icon: DollarSign },
      { id: 'dse-transactions', label: 'Transactions', icon: Layers },
      { id: 'dse-holdings', label: 'Holdings', icon: Briefcase },
      { id: 'dse-analytics', label: 'Analytics', icon: BarChart2 },
      { id: 'dse-settings', label: 'Settings', icon: SlidersHorizontal },
    ]
  },
  { id: 'mutual-funds', label: 'Mutual Funds', icon: PieChart },
  { id: 'fdrs', label: 'Fixed Deposits', icon: CreditCard },
  { id: 'sukuk', label: 'Sukuk Funds', icon: ShieldCheck },
  { id: 'online', label: 'Online Invests', icon: Globe },
  { 
    id: 'income-expense', 
    label: 'Income & Expense', 
    icon: BarChart3,
    children: [
      { id: 'income-expense-summary', label: 'Summary', icon: DollarSign },
      { id: 'income-expense-transactions', label: 'Transactions', icon: Layers },
      { id: 'income-expense-analytics', label: 'Analytics', icon: BarChart2 },
      { id: 'income-expense-settings', label: 'Settings', icon: SlidersHorizontal },
    ]
  },
  { id: 'settings', label: 'Settings', icon: Settings },
];

export const Sidebar: React.FC<SidebarProps> = ({ 
  activeModule, 
  setActiveModule, 
  isCollapsed, 
  setIsCollapsed,
  isMobileOpen,
  setIsMobileOpen
}) => {
  const [expanded, setExpanded] = React.useState<string | null>(null);

  return (
    <>
      {/* Mobile Overlay */}
      {isMobileOpen && (
        <div 
          className="fixed inset-0 bg-[#0b1121]/80 backdrop-blur-sm z-40 lg:hidden"
          onClick={() => setIsMobileOpen(false)}
        />
      )}

      <aside 
        className={cn(
          'bg-[#0b1121] border-r border-slate-800 transition-all duration-300 flex flex-col h-screen sticky top-0 z-50',
          isCollapsed ? 'lg:w-20' : 'lg:w-52',
          isMobileOpen 
            ? 'fixed inset-0 w-52 translate-x-0' 
            : 'hidden lg:flex -translate-x-full lg:translate-x-0'
        )}
      >
        <div className="h-16 p-4 flex items-center justify-between border-b border-slate-800">
          <div className={cn("flex items-center gap-2", isCollapsed && !isMobileOpen && "hidden")}>
              <div className="w-8 h-8 bg-teal-400 flex items-center justify-center">
                <TrendingUp className="text-slate-950 w-5 h-5" />
              </div>
              <span className="text-subheading font-bold text-white tracking-tight font-display">FinTrack<span className="text-teal-400">Pro</span></span>
          </div>
          
          <button 
            onClick={() => {
              if (window.innerWidth < 1024) {
                setIsMobileOpen(false);
              } else {
                setIsCollapsed(!isCollapsed);
              }
            }}
            className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-900 rounded-md transition-colors"
          >
            {isMobileOpen ? <ChevronLeft size={20} /> : <Menu size={20} />}
          </button>
        </div>

        <nav className="flex-1 px-3 space-y-0.5 overflow-y-auto mt-4">
          {menuItems.map((item) => (
            <React.Fragment key={item.id}>
              <button
                onClick={() => {
                  if (item.id !== 'dse') {
                    setExpanded(null);
                  }
                  if (item.children) {
                    setExpanded(expanded === item.id ? null : item.id);
                    setActiveModule(item.id + '-summary');
                  } else {
                    setActiveModule(item.id);
                  }
                  if (window.innerWidth < 1024 && !item.children) setIsMobileOpen(false);
                }}
                className={cn(
                  'w-full flex items-center gap-3 px-3 py-2 transition-all group relative ' + (isCollapsed ? 'rounded-none' : 'rounded-lg'),
                  (activeModule === item.id || (item.children && activeModule.startsWith(item.id + '-')))
                    ? 'bg-teal-400 text-slate-950 shadow-lg shadow-teal-400/20' 
                    : 'text-slate-300 hover:bg-slate-900 hover:text-white'
                )}
              >
                <item.icon className={cn('w-5 h-5 shrink-0', (activeModule === item.id || (item.children && activeModule.startsWith(item.id + '-'))) ? 'text-slate-950' : 'group-hover:text-white')} />
                {(!isCollapsed || isMobileOpen) && <span className="font-bold text-label uppercase">{item.label}</span>}
                {isCollapsed && !isMobileOpen && activeModule === item.id && (
                  <div className="absolute left-0 w-1 h-6 bg-slate-950 rounded-r-full" />
                )}
              </button>
              
              {item.children && expanded === item.id && (!isCollapsed || isMobileOpen) && (
                <div className="pl-6 space-y-0.5 mt-1">
                  {item.children.map(child => (
                    <button
                      key={child.id}
                      onClick={() => {
                        setActiveModule(child.id);
                        if (window.innerWidth < 1024) setIsMobileOpen(false);
                      }}
                      className={cn(
                        'w-full flex items-center gap-2 text-left px-3 py-1.5 rounded-lg text-label font-bold text-slate-400 hover:text-white hover:bg-slate-900',
                         activeModule === child.id && 'text-teal-400 bg-slate-900'
                      )}
                    >
                      <child.icon className="w-4 h-4 shrink-0" />
                      {child.label}
                    </button>
                  ))}
                </div>
              )}
            </React.Fragment>
          ))}
        </nav>

        <div className="p-4 border-t border-slate-800 text-label font-bold text-slate-300 text-center uppercase">
          Version: v1.0
        </div>
      </aside>
    </>
  );
};

type SyncStatus = 'idle' | 'syncing' | 'success' | 'error';

export const Navbar: React.FC<{
  title: React.ReactNode;
  onAdd?: () => void;
  addLabel?: string;
  onToggleSidebar: () => void;
  syncStatus?: SyncStatus;
  syncSummary?: string;
  lastSyncedAt?: number;
  onSync?: () => void;
}> = ({ title, onAdd, addLabel, onToggleSidebar, syncStatus = 'idle', syncSummary = '', lastSyncedAt = 0, onSync }) => {
  return (
    <header className="h-16 border-b border-slate-800 bg-[#0b1121] flex items-center justify-between px-4 lg:px-8 sticky top-0 z-10">
      <div className="flex items-center gap-3">
        <button
          onClick={onToggleSidebar}
          className="p-2 text-slate-400 hover:text-white hover:bg-slate-900 rounded-lg lg:hidden"
        >
          <Menu size={24} />
        </button>
        <Globe className="text-teal-400 w-5 h-5 lg:w-6 lg:h-6" />
        <h2 className="text-subheading lg:text-heading font-extrabold text-white font-display tracking-tight uppercase truncate max-w-[120px] sm:max-w-none">{title}</h2>
      </div>

      <div className="flex items-center gap-2 lg:gap-3">

        {/* ── Sync button ── */}
        {onSync && (
          <div className="flex items-center gap-2">
            {/* Last synced / summary label */}
            {syncStatus === 'idle' && lastSyncedAt > 0 && (
              <span className="hidden md:block text-[10px] text-slate-500 whitespace-nowrap">
                {formatLastSynced(lastSyncedAt)}
              </span>
            )}
            {(syncStatus === 'success' || syncStatus === 'error') && syncSummary && (
              <span className={`hidden md:block text-[10px] font-semibold whitespace-nowrap ${syncStatus === 'success' ? 'text-teal-400' : 'text-red-400'}`}>
                {syncSummary}
              </span>
            )}

            {/* The button itself */}
            <button
              onClick={onSync}
              disabled={syncStatus === 'syncing'}
              title={lastSyncedAt > 0 ? `Last synced: ${new Date(lastSyncedAt).toLocaleString()}` : 'Sync with Google Sheets'}
              className={cn(
                'flex items-center gap-1.5 px-2.5 h-8 sm:h-9 rounded-full text-[10px] lg:text-label font-bold uppercase border transition-all duration-200',
                syncStatus === 'idle'    && 'bg-slate-900 border-slate-700 text-slate-300 hover:bg-slate-800 hover:text-white hover:border-slate-500',
                syncStatus === 'syncing' && 'bg-slate-800 border-slate-700 text-slate-500 cursor-not-allowed',
                syncStatus === 'success' && 'bg-teal-900/40 border-teal-600 text-teal-300',
                syncStatus === 'error'   && 'bg-red-900/40 border-red-700 text-red-300',
              )}
            >
              {/* Spinner */}
              {syncStatus === 'syncing' && (
                <svg className="animate-spin h-3 w-3 shrink-0" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
                </svg>
              )}
              {/* Checkmark */}
              {syncStatus === 'success' && (
                <svg className="h-3 w-3 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="3">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                </svg>
              )}
              {/* Error X */}
              {syncStatus === 'error' && (
                <svg className="h-3 w-3 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="3">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                </svg>
              )}
              {/* Sync arrows */}
              {syncStatus === 'idle' && (
                <svg className="h-3 w-3 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2.5">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                </svg>
              )}
              <span className="hidden sm:inline">
                {syncStatus === 'syncing' ? 'Syncing…' : syncStatus === 'success' ? 'Synced!' : syncStatus === 'error' ? 'Retry' : 'Sync'}
              </span>
            </button>
          </div>
        )}

        {/* ── Add button ── */}
        {onAdd && (
          <button
            onClick={onAdd}
            className="flex items-center gap-2 px-3 lg:px-4 h-8 sm:h-9 bg-teal-400 hover:bg-teal-300 text-slate-950 rounded-full transition-all text-[10px] lg:text-label font-bold uppercase shadow-lg shadow-teal-400/20"
          >
            <Plus size={14} />
            <span className="hidden sm:inline">{addLabel || "New Investment"}</span>
          </button>
        )}

      </div>
    </header>
  );
};
