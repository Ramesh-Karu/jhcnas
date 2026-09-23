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
  HardDrive,
  Activity,
  Play,
  RefreshCw,
  AlertCircle,
  CheckCircle2,
  GitFork,
  ArrowRight,
  Radio,
  FileSpreadsheet,
  Terminal
} from 'lucide-react';
import { SyncSettings, NextcloudConfig, SupabaseConfig, WorksheetMapping } from '../types';
import { ApiClient } from '../services/apiClient';

interface SettingsViewProps {
  settings: SyncSettings;
  onSaveSettings: (s: SyncSettings) => void;
  nextcloudConfig?: NextcloudConfig;
  supabaseConfig?: SupabaseConfig;
  mappings?: WorksheetMapping[];
  onTriggerLiveSync?: () => Promise<void>;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  settings,
  onSaveSettings,
  nextcloudConfig,
  supabaseConfig,
  mappings = [],
  onTriggerLiveSync
}) => {
  const [formData, setFormData] = useState<SyncSettings>(settings);
  const [savedNotice, setSavedNotice] = useState<string | null>(null);
  
  // Diagnostics State
  const [isRunningDiagnostics, setIsRunningDiagnostics] = useState<boolean>(false);
  const [diagResult, setDiagResult] = useState<any>(null);

  // Live Sync State
  const [isRunningSync, setIsRunningSync] = useState<boolean>(false);
  const [syncNotice, setSyncNotice] = useState<{ success: boolean; message: string; details?: any } | null>(null);

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    onSaveSettings(formData);
    setSavedNotice('Worker scheduling and storage settings saved successfully.');
    setTimeout(() => setSavedNotice(null), 3000);
  };

  const handleRunDiagnostics = async () => {
    if (!nextcloudConfig || !supabaseConfig) return;
    setIsRunningDiagnostics(true);
    setDiagResult(null);

    try {
      const res = await ApiClient.getWorkerDiagnostics({
        nextcloud: nextcloudConfig,
        supabase: supabaseConfig,
        workerUrl: formData.workerUrl
      });

      if (res.success && res.diagnostics) {
        setDiagResult(res.diagnostics);
      } else {
        setDiagResult({
          error: res.error || 'Failed to retrieve diagnostics',
          allSystemsReady: false,
          nextcloud: { reachable: false, status: 'Error' },
          supabase: { reachable: false, status: 'Error' },
          workerService: { reachable: false, status: 'Error', endpoint: formData.workerUrl }
        });
      }
    } catch (e: any) {
      setDiagResult({
        error: e.message,
        allSystemsReady: false,
        nextcloud: { reachable: false, status: 'Error' },
        supabase: { reachable: false, status: 'Error' },
        workerService: { reachable: false, status: 'Error', endpoint: formData.workerUrl }
      });
    } finally {
      setIsRunningDiagnostics(false);
    }
  };

  const handleExecuteFullSync = async () => {
    if (!nextcloudConfig || !supabaseConfig) {
      alert('Nextcloud and Supabase configurations are required.');
      return;
    }

    setIsRunningSync(true);
    setSyncNotice(null);

    try {
      if (onTriggerLiveSync) {
        await onTriggerLiveSync();
        setSyncNotice({
          success: true,
          message: 'Live synchronization cycle executed successfully! Supabase PostgreSQL records updated.'
        });
      } else {
        const res = await ApiClient.executeFullPipelineSync({
          nextcloud: nextcloudConfig,
          supabase: supabaseConfig,
          mappings
        });

        if (res.success) {
          setSyncNotice({
            success: true,
            message: `Sync completed! Inserted: ${res.totalInserted || 0}, Updated: ${res.totalUpdated || 0}, Errors: ${res.totalFailed || 0}`,
            details: res
          });
        } else {
          setSyncNotice({
            success: false,
            message: `Sync encountered an issue: ${res.error || 'Unknown error'}`,
            details: res
          });
        }
      }
    } catch (e: any) {
      setSyncNotice({
        success: false,
        message: `Sync failed: ${e.message}`
      });
    } finally {
      setIsRunningSync(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-slate-100 text-slate-800">
            <Cpu className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Worker & Sync Engine Center</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Live status diagnostics, automated WebDAV scheduler, and Supabase PostgreSQL table router.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-3">
          <button
            id="btn-run-diagnostics"
            onClick={handleRunDiagnostics}
            disabled={isRunningDiagnostics}
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold shadow-2xs transition-colors"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isRunningDiagnostics ? 'animate-spin' : ''}`} />
            <span>{isRunningDiagnostics ? 'Testing Systems...' : 'Run Diagnostics Check'}</span>
          </button>

          <button
            id="btn-trigger-worker-sync"
            onClick={handleExecuteFullSync}
            disabled={isRunningSync}
            className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors"
          >
            <Play className={`w-3.5 h-3.5 ${isRunningSync ? 'animate-spin' : ''}`} />
            <span>{isRunningSync ? 'Executing Sync...' : 'Execute Full Sync Now'}</span>
          </button>
        </div>
      </div>

      {savedNotice && (
        <div className="p-3.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600" />
          <span>{savedNotice}</span>
        </div>
      )}

      {syncNotice && (
        <div className={`p-4 rounded-xl border text-xs font-medium flex items-start space-x-3 ${
          syncNotice.success ? 'bg-emerald-50 text-emerald-900 border-emerald-200' : 'bg-rose-50 text-rose-900 border-rose-200'
        }`}>
          {syncNotice.success ? (
            <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
          ) : (
            <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
          )}
          <div className="space-y-1">
            <div className="font-semibold text-sm">{syncNotice.message}</div>
            {syncNotice.details?.syncResults && (
              <div className="mt-2 space-y-1">
                {syncNotice.details.syncResults.map((sr: any, idx: number) => (
                  <div key={idx} className="font-mono text-[11px] bg-white/70 px-2 py-1 rounded border border-emerald-200/60">
                    Sheet: <strong>{sr.sheetName}</strong> ➔ Table: <strong>{sr.targetTable}</strong> ({sr.rowsCount} records upserted - {sr.status})
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Live Health & Diagnostics Matrix */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-4">
        <div className="flex items-center justify-between border-b border-slate-100 pb-3">
          <div className="flex items-center space-x-2">
            <Activity className="w-4 h-4 text-indigo-600" />
            <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider">
              Subsystem Connectivity & Health Matrix
            </h2>
          </div>
          <span className="text-xs text-slate-500">
            {diagResult ? `Last verified: ${new Date(diagResult.timestamp).toLocaleTimeString()}` : 'Click "Run Diagnostics Check" to verify all 4 subsystems'}
          </span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {/* Nextcloud Node */}
          <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Radio className="w-3.5 h-3.5 text-blue-600" />
                Nextcloud WebDAV
              </span>
              <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase ${
                diagResult?.nextcloud?.reachable ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
              }`}>
                {diagResult?.nextcloud?.reachable ? 'Reachable' : (nextcloudConfig?.url ? 'Configured' : 'Missing')}
              </span>
            </div>
            <div className="text-xs text-slate-600 font-mono truncate" title={nextcloudConfig?.url}>
              {nextcloudConfig?.url || 'https://cloud.jhcnexus.space'}
            </div>
            <div className="text-[11px] text-slate-500">
              Folder: <code className="bg-slate-200/70 px-1 py-0.5 rounded font-mono">{nextcloudConfig?.sourceFolder || '/ExcelImports'}</code>
            </div>
            {diagResult?.nextcloud?.latencyMs !== undefined && (
              <div className="text-[10px] text-slate-400">Latency: {diagResult.nextcloud.latencyMs}ms</div>
            )}
          </div>

          {/* Supabase Node */}
          <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Database className="w-3.5 h-3.5 text-emerald-600" />
                Supabase PostgreSQL
              </span>
              <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase ${
                diagResult?.supabase?.reachable ? 'bg-emerald-100 text-emerald-800' : 'bg-emerald-100 text-emerald-800'
              }`}>
                {diagResult?.supabase?.reachable ? 'Connected' : 'Configured'}
              </span>
            </div>
            <div className="text-xs text-slate-600 font-mono truncate" title={supabaseConfig?.url}>
              {supabaseConfig?.url || 'https://db.jhcnexus.space'}
            </div>
            <div className="text-[11px] text-slate-500">
              Auth: <span className="font-semibold text-emerald-700">Service Role Key (RLS Bypass)</span>
            </div>
            {diagResult?.supabase?.latencyMs !== undefined && (
              <div className="text-[10px] text-slate-400">Latency: {diagResult.supabase.latencyMs}ms</div>
            )}
          </div>

          {/* Sync Engine / Worker Node */}
          <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Cpu className="w-3.5 h-3.5 text-purple-600" />
                Worker Sync Pipeline
              </span>
              <span className="text-[10px] px-2 py-0.5 rounded-full font-bold uppercase bg-emerald-100 text-emerald-800">
                Active
              </span>
            </div>
            <div className="text-xs text-slate-600 font-mono truncate">
              {formData.workerUrl || 'Integrated In-App Pipeline'}
            </div>
            <div className="text-[11px] text-slate-500">
              Schedule: <span className="font-semibold text-slate-700">Every {formData.syncInterval || '15m'}</span>
            </div>
          </div>
        </div>
      </div>

      {/* Table Routing Hub */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-4">
        <div className="flex items-center justify-between border-b border-slate-100 pb-3">
          <div className="flex items-center space-x-2">
            <GitFork className="w-4 h-4 text-emerald-600" />
            <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider">
              Configured Table Routes ({mappings.length} Sheets Mapped)
            </h2>
          </div>
          <span className="text-xs text-slate-500">
            Defines which Excel sheets sync into which PostgreSQL tables during worker sync cycles.
          </span>
        </div>

        {mappings.length === 0 ? (
          <div className="p-6 text-center text-xs text-slate-500 bg-slate-50 rounded-lg border border-slate-200">
            No table routes configured yet. Open the <strong>Mappings</strong> tab to route your sheets to Supabase tables.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-50 text-slate-700 uppercase font-semibold border-b border-slate-200">
                <tr>
                  <th className="px-4 py-2.5">Source Excel Sheet</th>
                  <th className="px-4 py-2.5">Target PostgreSQL Table</th>
                  <th className="px-4 py-2.5">Conflict Key (Upsert)</th>
                  <th className="px-4 py-2.5">Mapped Columns</th>
                  <th className="px-4 py-2.5">Sync Direction</th>
                  <th className="px-4 py-2.5">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {mappings.map((m) => {
                  const uniqueCol = m.columns.find(c => c.uniqueKey)?.supabaseColumn || 'username';
                  return (
                    <tr key={m.id} className="hover:bg-slate-50/60">
                      <td className="px-4 py-3 font-semibold text-slate-900 flex items-center space-x-2">
                        <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                        <span>{m.worksheetName}</span>
                      </td>
                      <td className="px-4 py-3 font-mono font-bold text-emerald-700">
                        public.{m.supabaseTable}
                      </td>
                      <td className="px-4 py-3 font-mono text-slate-600">
                        <span className="bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200">
                          {uniqueCol}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-slate-600">
                        {m.columns.length} columns mapped
                      </td>
                      <td className="px-4 py-3">
                        <span className="text-[10px] px-2 py-0.5 rounded font-bold uppercase bg-emerald-100 text-emerald-800">
                          {m.syncPolicy === 'EXCEL_TO_DB' ? 'Excel Master' : m.syncPolicy || 'Bidirectional'}
                        </span>
                      </td>
                      <td className="px-4 py-3">
                        <span className={`inline-flex items-center text-xs font-semibold ${
                          m.enabled ? 'text-emerald-700' : 'text-slate-400'
                        }`}>
                          {m.enabled ? <CheckCircle2 className="w-3 h-3 mr-1 text-emerald-600" /> : null}
                          {m.enabled ? 'Active' : 'Disabled'}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Settings Form */}
      <form onSubmit={handleSave} className="space-y-6">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Worker Scheduling */}
          <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-4">
            <h2 className="text-base font-semibold text-slate-900 flex items-center space-x-2">
              <Clock className="w-4 h-4 text-blue-600" />
              <span>Automated Background Polling</span>
            </h2>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                Synchronization Interval
              </label>
              <select
                value={formData.syncInterval}
                onChange={(e) => setFormData({ ...formData, syncInterval: e.target.value as any })}
                className="w-full px-3 py-2 text-sm rounded-lg border border-slate-300 font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500"
              >
                <option value="5m">Every 5 minutes (Rapid ingest)</option>
                <option value="15m">Every 15 minutes (Standard)</option>
                <option value="30m">Every 30 minutes</option>
                <option value="1h">Hourly</option>
                <option value="1d">Daily</option>
              </select>
              <p className="text-[11px] text-slate-500 mt-1">
                The sync pipeline checks Nextcloud WebDAV on this interval, downloads modified files, and executes upserts.
              </p>
            </div>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                Coolify Worker API URL (Optional External Daemon)
              </label>
              <input
                type="text"
                value={formData.workerUrl}
                onChange={(e) => setFormData({ ...formData, workerUrl: e.target.value })}
                placeholder="https://worker.yourdomain.com or leave blank for internal engine"
                className="w-full px-3 py-2 text-sm rounded-lg border border-slate-300 font-mono focus:outline-none focus:ring-2 focus:ring-emerald-500"
              />
              <p className="text-[11px] text-slate-500 mt-1">
                If running the standalone Coolify Python/Node daemon, enter its URL here.
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

          {/* Supabase Storage Archiving */}
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
                Creates an immutable versioned snapshot in cloud storage for regulatory compliance.
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
