/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useCallback } from 'react';
import { Lock, Delete, ChevronRight } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';

interface PinLoginProps {
  onSuccess: () => void;
  correctPin: string | null;
  setPin: (pin: string) => void;
}

export const PinLogin: React.FC<PinLoginProps> = ({ onSuccess, correctPin, setPin }) => {
  const [input, setInput] = useState('');
  const [error, setError] = useState(false);

  const handleNumber = useCallback((num: string) => {
    if (error) return;
    if (input.length < 4) {
      const newInput = input + num;
      setInput(newInput);
      
      if (newInput.length === 4) {
        if (!correctPin) {
          // First time setup
          setPin(newInput);
          onSuccess();
        } else if (newInput === correctPin) {
          onSuccess();
        } else {
          setError(true);
          setTimeout(() => {
            setError(false);
            setInput('');
          }, 1000);
        }
      }
    }
  }, [input, error, correctPin, setPin, onSuccess]);

  const handleDelete = useCallback(() => {
    if (error) return;
    setInput(prev => prev.slice(0, -1));
  }, [error]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key >= '0' && e.key <= '9') {
        handleNumber(e.key);
      } else if (e.key === 'Backspace') {
        handleDelete();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleNumber, handleDelete]);

  return (
    <div className="min-h-screen bg-[#020617] flex flex-col items-center justify-center p-6">
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="w-full max-w-sm text-center"
      >
        <div className="w-20 h-20 bg-teal-400 rounded-3xl flex items-center justify-center mx-auto mb-8 shadow-2xl shadow-teal-400/20">
          <Lock className="text-white w-10 h-10" />
        </div>
        
        <h1 className="text-heading font-bold text-white mb-2 uppercase">
          {!correctPin ? 'Set Security PIN' : 'Welcome Back'}
        </h1>
        <p className="text-slate-300 mb-12 text-label uppercase font-bold">
          {!correctPin ? 'Create a 4-digit PIN to secure your data' : 'Enter your PIN to unlock FinTrack Pro'}
        </p>

        <div className="flex justify-center gap-4 mb-16">
          {[0, 1, 2, 3].map((i) => (
            <motion.div
              key={i}
              animate={error ? { x: [0, -10, 10, -10, 10, 0] } : {}}
              className={`w-4 h-4 rounded-full border-2 transition-all duration-300 ${
                input.length > i 
                  ? 'bg-teal-400 border-teal-400 scale-125 shadow-[0_0_15px_rgba(45,212,191,0.5)]' 
                  : 'border-slate-700'
              }`}
            />
          ))}
        </div>

        <div className="grid grid-cols-3 gap-3">
          {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((num) => (
            <button
              key={num}
              onClick={() => handleNumber(num.toString())}
              className="w-16 h-16 rounded-xl bg-slate-900 border border-slate-800 text-xl font-bold text-white hover:bg-slate-800 active:scale-90 transition-all flex items-center justify-center mx-auto"
            >
              {num}
            </button>
          ))}
          <div />
          <button
            onClick={() => handleNumber('0')}
            className="w-16 h-16 rounded-xl bg-slate-900 border border-slate-800 text-xl font-bold text-white hover:bg-slate-800 active:scale-90 transition-all flex items-center justify-center mx-auto"
          >
            0
          </button>
          <button
            onClick={handleDelete}
            className="w-16 h-16 rounded-xl bg-slate-900/50 text-slate-400 hover:text-white active:scale-90 transition-all flex items-center justify-center mx-auto"
          >
            <Delete size={24} />
          </button>
        </div>
      </motion.div>
    </div>
  );
};
