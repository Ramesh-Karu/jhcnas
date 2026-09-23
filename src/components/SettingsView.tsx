import React, { useState } from 'react';
import { 
  Settings, 
  Cpu, 
  Clock, 
  ShieldCheck, 
  Database, 
  Save, 
  Check, 
  Server, 
  CloudRain, 
  HardDrive 
} from 'lucide-react';
import { SyncSettings } from '../types';

interface SettingsViewProps {
  settings: SyncSettings;
  onSaveSettings: (s: SyncSettings) => void;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  settings,
  onSaveSettings
}) => {
  const [formData, setFormData] = useState<SyncSettings>(settings);
  const [savedNotice, setSavedNotice] = useState<string | null>(null);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    onSaveSettings(formData);
    setSavedNotice('Worker scheduling and storage settings saved.');
    setTimeout(() => setSavedNotice(null), 3000);
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-slate-100 text-slate-800">
            <Settings className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">System & Worker Settings</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Configure the Coolify background worker cadence, Supabase Storage archive policies, and synchronization rules.
            </p>
          </div>
        </div>
      </div>

      {savedNotice && (
        <div className="p-3.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600" />
          <span>{savedNotice}</span>
        </div>
      )}

      <form onSubmit={handleSave} className="space-y-6">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Worker Scheduling (Section 5) */}
          <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-4">
            <h2 className="text-base font-semibold text-slate-900 flex items-center space-x-2">
              <Cpu className="w-4 h-4 text-blue-600" />
              <span>Coolify Background Worker</span>
            </h2>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                Synchronization Schedule Interval
              </label>
              <select
                value={formData.syncInterval}
                onChange={(e) => setFormData({ ...formData, syncInterval: e.target.value as any })}
                className="w-full px-3 py-2 text-sm rounded-lg border border-slate-300 font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500"
              >
                <option value="5m">Every 5 minutes (Rapid ingest)</option>
                <option value="15m">Every 15 minutes (Standard recommendation)</option>
                <option value="30m">Every 30 minutes</option>
                <option value="1h">Hourly</option>
                <option value="1d">Daily</option>
              </select>
              <p className="text-[11px] text-slate-500 mt-1">
                The worker wakes up on this schedule, scans Nextcloud WebDAV, checks SHA-256 hashes, and processes modified files.
              </p>
            </div>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                Worker API Endpoint URL
              </label>
              <input
                type="text"
                value={formData.workerUrl}
                onChange={(e) => setFormData({ ...formData, workerUrl: e.target.value })}
                placeholder="https://worker.yourdomain.com or http://worker:8000"
                className="w-full px-3 py-2 text-sm rounded-lg border border-slate-300 font-mono focus:outline-none focus:ring-2 focus:ring-emerald-500"
              />
              <p className="text-[11px] text-slate-500 mt-1">
                If using a Coolify public domain/FQDN with Traefik (e.g. <code>https://worker.yourdomain.com</code>), <strong>do not include port 8000</strong>. Traefik routes standard port 443 directly to container port 8000.
              </p>
            </div>

            <div className="pt-2">
              <label className="flex items-center space-x-3 cursor-pointer">
                <input
                  type="checkbox"
                  checked={formData.autoSyncEnabled}
                  onChange={(e) => setFormData({ ...formData, autoSyncEnabled: e.target.checked })}
                  className="w-4 h-4 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500"
                />
                <span className="text-xs font-medium text-slate-800">
                  Enable automated background polling
                </span>
              </label>
            </div>
          </div>

          {/* Supabase Storage Archiving (Section 15) */}
          <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-4">
            <h2 className="text-base font-semibold text-slate-900 flex items-center space-x-2">
              <HardDrive className="w-4 h-4 text-teal-600" />
              <span>Storage & Backup Archival</span>
            </h2>

            <div className="pt-1">
              <label className="flex items-center space-x-3 cursor-pointer">
                <input
                  type="checkbox"
                  checked={formData.backupToStorage}
                  onChange={(e) => setFormData({ ...formData, backupToStorage: e.target.checked })}
                  className="w-4 h-4 text-teal-600 rounded border-slate-300 focus:ring-teal-500"
                />
                <span className="text-xs font-medium text-slate-800">
                  Archive copy of imported Excel files to Supabase Storage
                </span>
              </label>
              <p className="text-[11px] text-slate-500 mt-1 pl-7">
                Creates an immutable versioned snapshot in cloud storage for regulatory compliance and disaster recovery.
              </p>
            </div>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                Supabase Storage Bucket Name
              </label>
              <input
                type="text"
                value={formData.storageBucket}
                onChange={(e) => setFormData({ ...formData, storageBucket: e.target.value })}
                className="w-full px-3 py-2 text-sm rounded-lg border border-slate-300 font-mono focus:outline-none focus:ring-2 focus:ring-teal-500"
              />
            </div>

            <div className="p-3.5 rounded-lg bg-slate-50 border border-slate-200 text-xs text-slate-600 space-y-1">
              <div className="font-semibold text-slate-900">Non-Destructive File Principle</div>
              <p className="leading-relaxed">
                The sync pipeline treats Nextcloud Excel files as read-only assets. Original files on TrueNAS are never overwritten or deleted by the synchronization engine.
              </p>
            </div>
          </div>
        </div>

        <div className="flex justify-end">
          <button
            type="submit"
            id="btn-save-settings"
            className="inline-flex items-center space-x-2 px-5 py-2.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-semibold text-white shadow-2xs transition-colors"
          >
            <Save className="w-4 h-4" />
            <span>Save System Configuration</span>
          </button>
        </div>
      </form>
    </div>
  );
};
