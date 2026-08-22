import React from 'react';
import { AlertTriangle } from 'lucide-react';
import { cn } from '../../utils/formatters';

interface ConfirmDialogProps {
  isOpen: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  cancelLabel?: string;
  onConfirm: () => void;
  onCancel: () => void;
  variant?: 'danger' | 'warning' | 'info';
}

export const ConfirmDialog: React.FC<ConfirmDialogProps> = ({ 
  isOpen, 
  title, 
  message, 
  confirmLabel = 'Delete', 
  cancelLabel = 'Cancel', 
  onConfirm, 
  onCancel,
  variant = 'danger'
}) => {
  if (!isOpen) return null;

  const variantStyles = {
    danger: {
      iconBg: 'bg-rose-500/10',
      iconColor: 'text-rose-400',
      confirmBg: 'bg-rose-500 hover:bg-rose-400',
    },
    warning: {
      iconBg: 'bg-amber-500/10',
      iconColor: 'text-amber-400',
      confirmBg: 'bg-amber-500 hover:bg-amber-400',
    },
    info: {
      iconBg: 'bg-teal-500/10',
      iconColor: 'text-teal-400',
      confirmBg: 'bg-teal-500 hover:bg-teal-400',
    }
  };

  const style = variantStyles[variant];

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
      <div 
        className="absolute inset-0 bg-slate-950/80 backdrop-blur-sm" 
        onClick={onCancel} 
      />
      <div className="relative bg-slate-900 border border-slate-700/50 rounded-2xl shadow-2xl p-6 max-w-md w-full animate-in fade-in zoom-in-95 duration-200">
        <div className="flex items-start gap-4 mb-6">
          <div className={cn("w-12 h-12 rounded-2xl flex items-center justify-center shrink-0", style.iconBg)}>
            <AlertTriangle className={cn("w-6 h-6", style.iconColor)} />
          </div>
          <div className="flex-1">
            <p className="text-body font-bold text-white uppercase tracking-tight leading-none mb-1 text-base">{title}</p>
            <p className="text-sm text-slate-400 leading-relaxed font-medium">{message}</p>
          </div>
        </div>
        <div className="flex gap-3 justify-end">
          <button 
            onClick={onCancel}
            className="px-6 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-wider text-slate-300 bg-slate-800/80 hover:bg-slate-700 transition-all active:scale-95"
          >
            {cancelLabel}
          </button>
          <button 
            onClick={onConfirm}
            className={cn(
              "px-6 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-wider text-white transition-all active:scale-95",
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
