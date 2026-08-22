/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useRef, useState } from 'react';
import { Card, Button, Modal } from '../ui/BaseComponents';
import { AppState } from '../../types';
import { cn } from '../../utils/formatters';
import { Download, Upload, Shield, Database, AlertCircle, CheckCircle2, XCircle } from 'lucide-react';

interface SettingsModuleProps {
  state: AppState;
  updateState: (updater: (prev: AppState) => AppState) => void;
}

export const SettingsModule: React.FC<SettingsModuleProps> = ({ state, updateState }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [pendingRestore, setPendingRestore] = useState<AppState | null>(null);
  const [notification, setNotification] = useState<{ type: 'success' | 'error', message: string } | null>(null);
  const [isClearModalOpen, setIsClearModalOpen] = useState(false);
  const [clearConfirmText, setClearConfirmText] = useState('');

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 5000);
  };

  const handleBackup = () => {
    try {
      const dataStr = JSON.stringify(state, null, 2);
      const dataBlob = new Blob([dataStr], { type: 'application/json' });
      const url = URL.createObjectURL(dataBlob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `fintrack_pro_backup_${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      showNotification('success', 'Backup created and downloaded successfully!');
    } catch (error) {
      console.error('Backup failed:', error);
      showNotification('error', 'Failed to create backup. Please try again.');
    }
  };

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const importedState = JSON.parse(content);

        if (!importedState || typeof importedState !== 'object') {
          throw new Error('Invalid backup file format.');
        }

        const coreKeys = ['onlineInvestments', 'sukuks', 'cashBalance'];
        const hasCoreKeys = coreKeys.some(key => key in importedState);
        
        if (!hasCoreKeys) {
          throw new Error('This file does not appear to be a valid FinTrack Pro backup.');
        }

        setPendingRestore(importedState);
      } catch (error) {
        console.error('Restore failed:', error);
        showNotification('error', `Failed to restore data: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    };
    reader.onerror = () => {
      showNotification('error', 'Failed to read the file. Please try again.');
    };
    reader.readAsText(file);
    event.target.value = '';
  };

  const confirmRestore = () => {
    if (pendingRestore) {
      updateState(prev => ({
        ...prev,
        ...pendingRestore,
        isLocked: true
      }));
      setPendingRestore(null);
      showNotification('success', 'Data restored successfully! The application is now locked for security.');
    }
  };

  const handleClearData = () => {
    if (clearConfirmText === 'DELETE') {
      localStorage.clear();
      window.location.reload();
    } else {
      showNotification('error', 'Invalid confirmation text.');
    }
  };

  return (
    <div className="space-y-6 max-w-4xl mx-auto">
      <div className="flex items-center gap-3 mb-6">
        <Shield className="text-teal-400 w-6 h-6" />
        <h2 className="text-heading font-bold text-white font-display uppercase tracking-tight">Security & Data Settings</h2>
      </div>

      {notification && (
        <div className={cn(
          "flex items-center gap-3 p-4 rounded-xl border animate-in fade-in slide-in-from-top-4 duration-300",
          notification.type === 'success' ? "bg-teal-400/10 border-teal-400/20 text-teal-400" : "bg-rose-500/10 border-rose-500/20 text-rose-500"
        )}>
          {notification.type === 'success' ? <CheckCircle2 size={20} /> : <XCircle size={20} />}
          <p className="text-body font-bold uppercase">{notification.message}</p>
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Data Management Card */}
        <Card title="Data Management" subtitle="Backup and restore your entire portfolio">
          <div className="space-y-6">
            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800">
              <div className="flex items-start gap-3">
                <Database className="text-blue-400 w-5 h-5 mt-1 shrink-0" />
                <div>
                  <p className="text-body font-bold text-white mb-1">Local Data Storage</p>
                  <p className="text-label text-slate-400 leading-relaxed">
                    Your data is stored locally in your browser. We recommend taking regular backups to prevent data loss if you clear your browser cache.
                  </p>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <Button 
                onClick={handleBackup}
                variant="secondary"
                className="flex flex-col items-center justify-center p-6 h-auto gap-3 border-dashed hover:border-teal-400/50 hover:bg-teal-400/5 group transition-all"
              >
                <div className="w-12 h-12 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center group-hover:scale-110 transition-transform">
                  <Download className="text-teal-400 w-6 h-6" />
                </div>
                <div className="text-center">
                  <p className="text-body font-bold text-white">Backup Data</p>
                  <p className="text-[10px] text-slate-500 uppercase font-bold">Export to JSON</p>
                </div>
              </Button>

              <Button 
                onClick={() => fileInputRef.current?.click()}
                variant="secondary"
                className="flex flex-col items-center justify-center p-6 h-auto gap-3 border-dashed hover:border-blue-400/50 hover:bg-blue-400/5 group transition-all"
              >
                <div className="w-12 h-12 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center group-hover:scale-110 transition-transform">
                  <Upload className="text-blue-400 w-6 h-6" />
                </div>
                <div className="text-center">
                  <p className="text-body font-bold text-white">Restore Data</p>
                  <p className="text-[10px] text-slate-500 uppercase font-bold">Import from JSON</p>
                </div>
              </Button>
              <input 
                type="file" 
                ref={fileInputRef}
                onChange={handleFileSelect}
                accept=".json"
                className="hidden"
              />
            </div>

            <div className="flex items-center gap-2 p-3 bg-amber-500/10 border border-amber-500/20 rounded-lg">
              <AlertCircle className="text-amber-500 w-4 h-4 shrink-0" />
              <p className="text-[10px] font-bold text-amber-500 uppercase">
                Restoring data will overwrite all current information.
              </p>
            </div>

            <div className="pt-4 border-t border-slate-800">
              <button 
                onClick={() => setIsClearModalOpen(true)}
                className="text-[10px] font-bold text-rose-500/50 hover:text-rose-500 uppercase transition-colors"
              >
                Clear All Application Data
              </button>
            </div>
          </div>
        </Card>

        {/* Security Settings Card */}
        <Card title="Security" subtitle="Manage your access control">
          <div className="space-y-6">
            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800">
              <p className="text-body font-bold text-white mb-2">PIN Protection</p>
              <p className="text-label text-slate-400 mb-4">
                {state.pin ? 'Your application is protected by a PIN.' : 'Set a PIN to protect your financial data from unauthorized access.'}
              </p>
              <Button variant="primary" className="w-full" disabled>
                {state.pin ? 'Change PIN' : 'Set PIN'}
              </Button>
              <p className="text-[10px] text-slate-600 mt-2 italic text-center uppercase font-bold">PIN Management coming in next update</p>
            </div>

            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800">
              <p className="text-body font-bold text-white mb-2">Auto-Lock</p>
              <p className="text-label text-slate-400 mb-4">
                Automatically lock the application when you close the tab.
              </p>
              <div className="flex items-center justify-between p-3 bg-slate-900 rounded-lg border border-slate-800">
                <span className="text-label font-bold text-slate-300 uppercase">Enabled</span>
                <div className="w-10 h-5 bg-teal-400 rounded-full relative">
                  <div className="absolute right-1 top-1 w-3 h-3 bg-white rounded-full" />
                </div>
              </div>
            </div>
          </div>
        </Card>
      </div>

      <Card title="About FinTrack Pro" className="border-teal-400/20">
        <div className="space-y-4">
          <p className="text-body text-slate-300 leading-relaxed">
            FinTrack Pro is a comprehensive investment and expense tracking solution designed for personal use.
          </p>
          <div className="space-y-1 pt-2">
            <p className="text-body">
              <span className="font-bold text-slate-400">Version: </span>
              <span className="font-bold text-white">v1.0</span>
            </p>
            <p className="text-body">
              <span className="font-bold text-slate-400">Last Update: </span>
              <span className="font-bold text-white">May 2026</span>
            </p>
          </div>
        </div>
      </Card>

      {/* Restore Confirmation Modal */}
      <Modal 
        isOpen={!!pendingRestore} 
        onClose={() => setPendingRestore(null)} 
        title="Confirm Restore"
      >
        <div className="space-y-6">
          <div className="flex items-start gap-4 p-4 bg-amber-500/10 border border-amber-500/20 rounded-xl">
            <AlertCircle className="text-amber-500 w-6 h-6 shrink-0 mt-1" />
            <div>
              <p className="text-body font-bold text-white mb-1 uppercase">Overwrite Existing Data?</p>
              <p className="text-label text-slate-400 leading-relaxed">
                Restoring will replace all your current investments, transactions, and settings with the data from the backup file. This action cannot be undone.
              </p>
            </div>
          </div>

          <div className="flex gap-3">
            <Button variant="ghost" className="flex-1" onClick={() => setPendingRestore(null)}>
              Cancel
            </Button>
            <Button variant="primary" className="flex-1" onClick={confirmRestore}>
              Confirm Restore
            </Button>
          </div>
        </div>
      </Modal>

      {/* Clear Data Modal */}
      <Modal 
        isOpen={isClearModalOpen} 
        onClose={() => setIsClearModalOpen(false)} 
        title="Critical: Clear All Data"
      >
        <div className="space-y-6">
          <div className="flex items-start gap-4 p-4 bg-rose-500/10 border border-rose-500/20 rounded-xl">
            <AlertCircle className="text-rose-500 w-6 h-6 shrink-0 mt-1" />
            <div>
              <p className="text-body font-bold text-white mb-1 uppercase">Permanent Deletion</p>
              <p className="text-label text-slate-400 leading-relaxed">
                This will permanently delete ALL your data and reset the application to its original state.
              </p>
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-label font-bold text-slate-300 uppercase">Type "DELETE" to confirm</label>
            <input 
              type="text"
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-2 text-white focus:outline-none focus:ring-2 focus:ring-rose-500/50"
              value={clearConfirmText}
              onChange={(e) => setClearConfirmText(e.target.value)}
              placeholder="DELETE"
            />
          </div>

          <div className="flex gap-3">
            <Button variant="ghost" className="flex-1" onClick={() => setIsClearModalOpen(false)}>
              Cancel
            </Button>
            <Button 
              variant="danger" 
              className="flex-1" 
              onClick={handleClearData}
              disabled={clearConfirmText !== 'DELETE'}
            >
              Clear Everything
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
};
