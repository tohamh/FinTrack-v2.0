/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { cn } from '../../utils/formatters';
import { ChevronDown } from 'lucide-react';

interface CardProps {
  children: React.ReactNode;
  className?: string;
  title?: string;
  subtitle?: string;
}

export const Card: React.FC<CardProps> = ({ children, className, title, subtitle }) => {
  return (
    <div className={cn('bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl flex flex-col', className)}>
      {(title || subtitle) && (
        <div className="px-6 py-4 border-b border-slate-800">
          {title && <h3 className="text-heading font-bold text-white uppercase tracking-tight">{title}</h3>}
          {subtitle && <p className="text-label font-medium text-slate-300 uppercase">{subtitle}</p>}
        </div>
      )}
      <div className="p-6 flex-1 flex flex-col">{children}</div>
    </div>
  );
};

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'primary' | 'secondary' | 'ghost' | 'danger';
  size?: 'sm' | 'md' | 'lg';
}

export const Button: React.FC<ButtonProps> = ({ 
  children, 
  className, 
  variant = 'primary', 
  size = 'md', 
  ...props 
}) => {
  const variants = {
    primary:  "px-4 py-2 rounded-lg bg-teal-400 text-slate-950 text-label font-bold uppercase hover:bg-teal-300 transition-colors",
    secondary: "px-4 py-2 rounded-lg bg-slate-800 border border-slate-700 text-slate-300 text-label font-bold uppercase hover:bg-slate-700 transition-colors",
    ghost:  "px-4 py-2 rounded-lg bg-slate-800 border border-slate-700 text-slate-300 text-label font-bold uppercase hover:bg-slate-700 transition-colors",
    danger:  'px-4 py-2 rounded-lg bg-rose-500 text-white text-label font-bold uppercase hover:bg-rose-600 transition-colors',
  };

  const sizes = {
    sm: 'px-3 py-1.5 text-label uppercase tracking-wider',
    md: 'px-4 py-2 text-body',
    lg: 'px-6 py-3 text-subheading',
  };

  return (
    <button 
      className={cn(
        'rounded-lg font-bold transition-all active:scale-95 disabled:opacity-50 disabled:pointer-events-none flex items-center justify-center gap-2',
        variants[variant],
        sizes[size],
        className
      )}
      {...props}
    >
      {children}
    </button>
  );
};

interface ModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  children: React.ReactNode;
}

export const Modal: React.FC<ModalProps> = ({ isOpen, onClose, title, children }) => {
  const modalRef = useRef<HTMLDivElement>(null);
  
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    if (isOpen) {
      document.addEventListener('keydown', handleKeyDown);
    }
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, onClose]);

  useEffect(() => {
    if (isOpen) {
      // Auto-focus first input field only once when modal opens
      const timer = setTimeout(() => {
        const firstInput = modalRef.current?.querySelector('input, select, textarea, button:not([title="Close"])') as HTMLElement;
        firstInput?.focus();
      }, 100);
      return () => clearTimeout(timer);
    }
  }, [isOpen]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#020617]/80 backdrop-blur-sm">
      <div ref={modalRef} className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-800 flex justify-between items-center">
          <h3 className="text-heading font-bold text-white uppercase tracking-tight">{title}</h3>
          <button onClick={onClose} title="Close" className="text-slate-300 hover:text-white transition-colors">
            <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
        <div className="p-6 max-h-[80vh] overflow-y-auto">
          {children}
        </div>
      </div>
    </div>
  );
};

interface InputProps extends React.InputHTMLAttributes<HTMLInputElement | HTMLSelectElement> {
  label: string;
  as?: 'input' | 'select';
  helperText?: string;
}

export const Input: React.FC<InputProps> = ({ label, className, as = 'input', helperText, ...props }) => {
  const Component = as as any;
  return (
    <div className="space-y-1.5">
      <label className="text-label font-bold text-slate-300 uppercase">{label}</label>
      <Component
        className={cn(
          'w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-2 text-body text-slate-300 placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-teal-400/50 focus:border-teal-400 transition-all tabular-nums appearance-none',
          className
        )}
        {...props}
      />
      {helperText && (
        <p className="text-[10px] text-slate-500 font-semibold tracking-wide uppercase mt-1">
          {helperText}
        </p>
      )}
    </div>
  );
};

export const Checkbox: React.FC<{ label?: string; checked: boolean; onChange: (checked: boolean) => void; id?: string }> = ({ label, checked, onChange, id }) => {
  return (
    <label id={id} className="flex items-center gap-3 cursor-pointer group">
      <div 
        onClick={() => onChange(!checked)}
        className={cn(
          'w-5 h-5 rounded border flex items-center justify-center transition-all shrink-0',
          checked ? 'bg-teal-400 border-teal-400' : 'bg-slate-950 border-slate-700 group-hover:border-slate-500'
        )}
      >
        {checked && <svg className="w-3.5 h-3.5 text-slate-950 stroke-[3]" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" /></svg>}
      </div>
      {label && <span className="text-label font-bold text-slate-300 uppercase group-hover:text-white transition-colors">{label}</span>}
    </label>
  );
};

interface SelectProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: { label: string; value: string }[];
  placeholder?: string;
  className?: string;
  required?: boolean;
}

export const Select: React.FC<SelectProps> = ({ 
  label, 
  value, 
  onChange, 
  options, 
  placeholder = 'Select an option',
  className,
  required
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const [highlightedIndex, setHighlightedIndex] = useState(-1);

  const handleKeyDown = (e: KeyboardEvent) => {
    if (!isOpen) return;
    
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setHighlightedIndex(prev => Math.min(prev + 1, options.length - 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setHighlightedIndex(prev => Math.max(prev - 1, 0));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (highlightedIndex >= 0) {
        onChange(options[highlightedIndex].value);
        setIsOpen(false);
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setIsOpen(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      document.addEventListener('keydown', handleKeyDown);
      setHighlightedIndex(options.findIndex(o => o.value === value));
    }
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, options, value]);

  const selectedOption = options.find(opt => opt.value === value);

  return (
    <div className={cn("space-y-1.5 relative", className)} ref={containerRef}>
      <label className="text-label font-bold text-slate-300 uppercase">
        {label} {required && <span className="text-rose-500">*</span>}
      </label>
      
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className={cn(
          "w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-2 text-left transition-all flex items-center justify-between group",
          isOpen ? "border-teal-400 ring-2 ring-teal-400/20" : "hover:border-slate-700"
        )}
      >
        <span className={cn(
          "text-body truncate uppercase font-bold",
          selectedOption ? "text-white" : "text-slate-600"
        )}>
          {selectedOption ? selectedOption.label : placeholder}
        </span>
        <ChevronDown 
          size={16} 
          className={cn(
            "text-slate-500 transition-transform duration-200",
            isOpen ? "rotate-180 text-teal-400" : "group-hover:text-slate-400"
          )} 
        />
      </button>

      {isOpen && (
        <div className="absolute left-0 right-0 mt-2 bg-slate-900 border border-slate-800 rounded-xl shadow-2xl z-[100] py-1 max-h-60 overflow-y-auto animate-in fade-in zoom-in-95">
          {options.length === 0 ? (
            <div className="px-4 py-3 text-label text-slate-500 italic text-center uppercase">
              No options available
            </div>
          ) : (
            options.map((option, index) => (
              <button
                key={option.value}
                type="button"
                onClick={() => {
                  onChange(option.value);
                  setIsOpen(false);
                }}
                className={cn(
                  "w-full text-left px-4 py-2.5 text-label font-bold uppercase transition-colors",
                  option.value === value || highlightedIndex === index
                    ? "bg-teal-400 text-slate-950" 
                    : "text-slate-300 hover:bg-slate-800 hover:text-white"
                )}
              >
                {option.label}
              </button>
            ))
          )}
        </div>
      )}
    </div>
  );
};
