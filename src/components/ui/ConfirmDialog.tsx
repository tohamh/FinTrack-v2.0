import React from 'react';
import { AlertTriangle, AlertOctagon, ShieldAlert, Layers } from 'lucide-react';
import { cn } from '../../utils/formatters';

export interface ConfirmDialogProps {
  isOpen: boolean;
  title: string;
  message: string | React.ReactNode;
  confirmLabel?: string;
  cancelLabel?: string;
  onConfirm: () => void;
  onCancel: () => void;
  variant?: 'danger' | 'warning' | 'info' | 'severe' | 'critical';
  details?: (string | null | undefined)[];
}

export const ConfirmDialog: React.FC<ConfirmDialogProps> = ({ 
  isOpen, 
  title, 
  message, 
  confirmLabel = 'Delete', 
  cancelLabel = 'Cancel', 
  onConfirm, 
  onCancel,
  variant = 'danger',
  details
}) => {
  if (!isOpen) return null;

  const isSevere = variant === 'severe' || variant === 'critical';

  const variantStyles = {
    danger: {
      iconBg: 'bg-rose-500/10',
      iconColor: 'text-rose-400',
      confirmBg: 'bg-rose-500 hover:bg-rose-400 shadow-rose-500/20',
      dialogBorder: 'border-slate-700/50',
      dialogShadow: 'shadow-2xl',
      Icon: AlertTriangle,
    },
    severe: {
      iconBg: 'bg-rose-600/25 ring-2 ring-rose-500/50 animate-pulse',
      iconColor: 'text-rose-400',
      confirmBg: 'bg-rose-600 hover:bg-rose-500 shadow-lg shadow-rose-600/40 ring-1 ring-rose-400/50',
      dialogBorder: 'border-rose-500/60 ring-2 ring-rose-500/30',
      dialogShadow: 'shadow-[0_0_60px_rgba(225,29,72,0.35)]',
      Icon: AlertOctagon,
    },
    critical: {
      iconBg: 'bg-rose-600/25 ring-2 ring-rose-500/50 animate-pulse',
      iconColor: 'text-rose-400',
      confirmBg: 'bg-rose-600 hover:bg-rose-500 shadow-lg shadow-rose-600/40 ring-1 ring-rose-400/50',
      dialogBorder: 'border-rose-500/60 ring-2 ring-rose-500/30',
      dialogShadow: 'shadow-[0_0_60px_rgba(225,29,72,0.35)]',
      Icon: ShieldAlert,
    },
    warning: {
      iconBg: 'bg-amber-500/10',
      iconColor: 'text-amber-400',
      confirmBg: 'bg-amber-500 hover:bg-amber-400 shadow-amber-500/20',
      dialogBorder: 'border-slate-700/50',
      dialogShadow: 'shadow-2xl',
      Icon: AlertTriangle,
    },
    info: {
      iconBg: 'bg-teal-500/10',
      iconColor: 'text-teal-400',
      confirmBg: 'bg-teal-500 hover:bg-teal-400 shadow-teal-500/20',
      dialogBorder: 'border-slate-700/50',
      dialogShadow: 'shadow-2xl',
      Icon: AlertTriangle,
    }
  };

  const style = variantStyles[variant] || variantStyles.danger;
  const DialogIcon = style.Icon;
  const validDetails = details?.filter(Boolean) as string[];

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
      <div 
        className={cn(
          "absolute inset-0 backdrop-blur-sm transition-opacity",
          isSevere ? "bg-slate-950/90" : "bg-slate-950/80"
        )} 
        onClick={onCancel} 
      />
      <div className={cn(
        "relative bg-slate-900 border rounded-2xl p-6 max-w-lg w-full animate-in fade-in zoom-in-95 duration-200",
        style.dialogBorder,
        style.dialogShadow
      )}>
        {/* Severe Header Banner */}
        {isSevere && (
          <div className="mb-4 -mt-2 -mx-2 px-3 py-1.5 rounded-xl bg-rose-500/15 border border-rose-500/30 flex items-center justify-between">
            <div className="flex items-center gap-2 text-[10px] font-black uppercase tracking-widest text-rose-300">
              <span className="w-2 h-2 rounded-full bg-rose-400 animate-ping" />
              <span>Severe Warning: Cascading Permanent Deletion</span>
            </div>
            <span className="text-[9px] font-bold text-rose-400/80 uppercase tracking-tight">Irreversible Action</span>
          </div>
        )}

        <div className="flex items-start gap-4 mb-4">
          <div className={cn("w-12 h-12 rounded-2xl flex items-center justify-center shrink-0", style.iconBg)}>
            <DialogIcon className={cn("w-6 h-6", style.iconColor)} />
          </div>
          <div className="flex-1">
            <p className={cn(
              "font-bold uppercase tracking-tight leading-none mb-1 text-base",
              isSevere ? "text-rose-200" : "text-white"
            )}>
              {title}
            </p>
            <div className="text-sm text-slate-300 leading-relaxed font-medium">
              {message}
            </div>
          </div>
        </div>

        {/* Detailed cascading impact breakdown */}
        {validDetails && validDetails.length > 0 && (
          <div className={cn(
            "mb-5 p-3.5 rounded-xl border text-xs space-y-2",
            isSevere 
              ? "bg-rose-950/40 border-rose-800/40 text-rose-200" 
              : "bg-slate-950/50 border-slate-800 text-slate-300"
          )}>
            <div className="flex items-center gap-1.5 font-bold uppercase tracking-wider text-[10px] text-slate-400">
              <Layers size={12} className={isSevere ? "text-rose-400" : "text-slate-400"} />
              <span>Entries and dependencies that will be permanently erased:</span>
            </div>
            <ul className="space-y-1 pl-1">
              {validDetails.map((item, idx) => (
                <li key={idx} className="flex items-start gap-2 text-xs">
                  <span className={cn("mt-1 w-1.5 h-1.5 rounded-full shrink-0", isSevere ? "bg-rose-400" : "bg-slate-400")} />
                  <span className="leading-snug">{item}</span>
                </li>
              ))}
            </ul>
          </div>
        )}

        <div className="flex gap-3 justify-end pt-2 border-t border-slate-800/60">
          <button 
            type="button"
            onClick={onCancel}
            className="px-5 py-2.5 rounded-xl text-[11px] font-bold uppercase tracking-wider text-slate-300 bg-slate-800 hover:bg-slate-700 hover:text-white transition-all active:scale-95 cursor-pointer"
          >
            {cancelLabel}
          </button>
          <button 
            type="button"
            onClick={onConfirm}
            className={cn(
              "px-6 py-2.5 rounded-xl text-[11px] font-black uppercase tracking-wider text-white transition-all active:scale-95 cursor-pointer shadow-md",
              style.confirmBg
            )}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};
