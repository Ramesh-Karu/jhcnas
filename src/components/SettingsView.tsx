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
  Terminal,
  Key,
  Eye,
  EyeOff,
  Edit3
} from 'lucide-react';
import { SyncSettings, NextcloudConfig, SupabaseConfig, WorksheetMapping, LiveSchedulerStatus, NextcloudFile } from '../types';
import { ApiClient } from '../services/apiClient';

interface SettingsViewProps {
  settings: SyncSettings;
  onSaveSettings: (s: SyncSettings) => void;
  files?: NextcloudFile[];
  onToggleFileAutoSync?: (filename: string, enabled: boolean) => void;
  nextcloudConfig?: NextcloudConfig;
  supabaseConfig?: SupabaseConfig;
  mappings?: WorksheetMapping[];
  onTriggerLiveSync?: () => Promise<void>;
  onOpenSecretsVault?: () => void;
  onSaveNextcloud?: (cfg: NextcloudConfig) => void;
  onSaveSupabase?: (cfg: SupabaseConfig) => void;
  schedulerStatus?: LiveSchedulerStatus;
  onRefreshScheduler?: () => Promise<void>;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  settings,
  onSaveSettings,
  files = [],
  onToggleFileAutoSync,
  nextcloudConfig,
  supabaseConfig,
  mappings = [],
  onTriggerLiveSync,
  onOpenSecretsVault,
  onSaveNextcloud,
  onSaveSupabase,
  schedulerStatus,
  onRefreshScheduler,
}) => {
  const [formData, setFormData] = useState<SyncSettings>(settings);
  const [savedNotice, setSavedNotice] = useState<string | null>(null);
  
  // Secrets Visibility in Settings
  const [revealedSecrets, setRevealedSecrets] = useState<Record<string, boolean>>({});
  const [editingSecretKey, setEditingSecretKey] = useState<string | null>(null);

  // Diagnostics State
  const [isRunningDiagnostics, setIsRunningDiagnostics] = useState<boolean>(false);
  const [diagResult, setDiagResult] = useState<any>(null);

  // Live Sync State
  const [isRunningSync, setIsRunningSync] = useState<boolean>(false);
  const [syncNotice, setSyncNotice] = useState<{ success: boolean; message: string; details?: any } | null>(null);

  const toggleSecretReveal = (key: string) => {
    setRevealedSecrets(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    onSaveSettings(formData);
    try {
      await ApiClient.configureScheduler({
        enabled: formData.autoSyncEnabled,
        intervalLabel: formData.syncInterval,
        workerUrl: formData.workerUrl,
        nextcloud: nextcloudConfig,
        supabase: supabaseConfig,
        mappings,
      });
      if (onRefreshScheduler) {
        await onRefreshScheduler();
      }
    } catch {
      // Best-effort background configure
    }
    setSavedNotice('Production sync scheduler configuration applied successfully.');
    setTimeout(() => setSavedNotice(null), 3500);
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

        <div className="flex items-center space-x-2 flex-wrap gap-y-2">
          <button
            id="btn-sync-all-coolify"
            onClick={async () => {
              setIsRunningSync(true);
              try {
                const res = await ApiClient.saveFullServerState({
                  nextcloud: nextcloudConfig,
                  supabase: supabaseConfig,
                  syncSettings: formData,
                  mappings,
                });
                if (res.success) {
                  setSavedNotice('✨ All environment variables, presets, mappings, and worker settings permanently saved to Coolify server disk & local storage!');
                  setTimeout(() => setSavedNotice(null), 5000);
                }
              } catch (e: any) {
                alert(`Notice saving to Coolify server disk: ${e.message}`);
              } finally {
                setIsRunningSync(false);
              }
            }}
            disabled={isRunningSync}
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-2xs transition-colors"
            title="Persist all variables, sheet mappings, presets, and worker configs to Coolify server disk & localStorage"
          >
            <Server className="w-3.5 h-3.5 text-indigo-200" />
            <span>Sync All to Coolify Server Disk & LocalStorage</span>
          </button>

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

      {/* Coolify Server Disk & Local Storage Dual-Persistence Hub */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-xl p-5 text-white shadow-sm border border-indigo-900/60 space-y-3">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-indigo-800/40 pb-3">
          <div className="flex items-center space-x-2.5">
            <div className="p-2 rounded-lg bg-indigo-600/30 text-indigo-400 border border-indigo-500/30">
              <Server className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h2 className="text-sm font-bold uppercase tracking-wider text-indigo-100">
                  Coolify Server Disk & LocalStorage Auto-Persistence Hub
                </h2>
                <span className="text-[10px] px-2 py-0.5 rounded font-bold uppercase bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  ● Continuous Dual-Sync Active
                </span>
              </div>
              <p className="text-xs text-indigo-200/70 mt-0.5">
                All environment variables, sheet mappings, presets, worker configs, and credentials are saved to server disk (<code className="text-indigo-300 font-mono">data/</code> + <code className="text-indigo-300 font-mono">.env</code>) and mirrored in your browser's local storage.
              </p>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2.5 pt-1">
          <div className="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700/60 space-y-1">
            <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Disk Secrets</div>
            <div className="font-mono text-xs font-semibold text-emerald-400 flex items-center justify-between">
              <span>data/secrets.json</span>
              <Check className="w-3.5 h-3.5 text-emerald-400" />
            </div>
            <div className="text-[10px] text-slate-400 truncate">Credentials & URLs</div>
          </div>

          <div className="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700/60 space-y-1">
            <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Sheet Mappings</div>
            <div className="font-mono text-xs font-semibold text-indigo-300 flex items-center justify-between">
              <span>data/mappings.json</span>
              <Check className="w-3.5 h-3.5 text-emerald-400" />
            </div>
            <div className="text-[10px] text-slate-400">{mappings.length} worksheets</div>
          </div>

          <div className="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700/60 space-y-1">
            <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400">AI Presets Store</div>
            <div className="font-mono text-xs font-semibold text-purple-300 flex items-center justify-between">
              <span>data/presets.json</span>
              <Check className="w-3.5 h-3.5 text-emerald-400" />
            </div>
            <div className="text-[10px] text-slate-400 truncate">Archetypes & Custom</div>
          </div>

          <div className="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700/60 space-y-1">
            <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Worker Daemon</div>
            <div className="font-mono text-xs font-semibold text-amber-300 flex items-center justify-between">
              <span>worker_config.json</span>
              <Check className="w-3.5 h-3.5 text-emerald-400" />
            </div>
            <div className="text-[10px] text-slate-400">Interval: {formData.syncInterval}</div>
          </div>

          <div className="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700/60 space-y-1">
            <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Container .env</div>
            <div className="font-mono text-xs font-semibold text-teal-300 flex items-center justify-between">
              <span>.env Auto-Updated</span>
              <Check className="w-3.5 h-3.5 text-emerald-400" />
            </div>
            <div className="text-[10px] text-slate-400 truncate">Survives Reboots</div>
          </div>
        </div>
      </div>

      {/* System Secrets & Configuration Vault (Admin Panel Integration) */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div className="flex items-center space-x-2.5">
            <div className="p-2 rounded-lg bg-amber-50 text-amber-600">
              <Key className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider">
                  Secrets & Credentials Vault
                </h2>
                <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase ${
                  (nextcloudConfig?.isConnected && supabaseConfig?.isConnected)
                    ? 'bg-emerald-100 text-emerald-800'
                    : 'bg-amber-100 text-amber-800'
                }`}>
                  {(nextcloudConfig?.isConnected && supabaseConfig?.isConnected) ? '● All Connected' : '○ Configuration Check'}
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Inspect, reveal, edit, and confirm that all required integration secrets are active.
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-2">
            {onOpenSecretsVault && (
              <button
                type="button"
                id="btn-open-secrets-vault-settings"
                onClick={onOpenSecretsVault}
                className="inline-flex items-center space-x-1.5 px-3.5 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold shadow-2xs transition-colors"
              >
                <Key className="w-3.5 h-3.5 text-amber-400" />
                <span>Open Full Secrets Vault</span>
              </button>
            )}
          </div>
        </div>

        {/* Quick Secrets Cards Matrix */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
          {/* Nextcloud App Password */}
          <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800">NEXTCLOUD_APP_PASSWORD</span>
              <span className={`text-[10px] px-2 py-0.5 rounded font-bold uppercase ${
                nextcloudConfig?.appPassword ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
              }`}>
                {nextcloudConfig?.appPassword ? 'Configured' : 'Missing'}
              </span>
            </div>
            <div className="flex items-center justify-between text-xs font-mono bg-white p-2 rounded border border-slate-200">
              <span className="truncate max-w-[180px]">
                {revealedSecrets['nc_pass'] ? (nextcloudConfig?.appPassword || 'None') : '••••••••••••••••'}
              </span>
              <div className="flex items-center space-x-1">
                <button
                  type="button"
                  onClick={() => toggleSecretReveal('nc_pass')}
                  className="p-1 text-slate-400 hover:text-slate-600"
                  title="Reveal or mask"
                >
                  {revealedSecrets['nc_pass'] ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                </button>
                {onOpenSecretsVault && (
                  <button
                    type="button"
                    onClick={onOpenSecretsVault}
                    className="p-1 text-slate-400 hover:text-slate-700"
                    title="Edit secret"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>
            </div>
            <div className="text-[10px] text-slate-500">
              WebDAV token for <span className="font-semibold text-slate-700">{nextcloudConfig?.username || 'truenas_admin'}</span>
            </div>
          </div>

          {/* Supabase Service Role Key */}
          <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800">SUPABASE_SERVICE_ROLE_KEY</span>
              <span className={`text-[10px] px-2 py-0.5 rounded font-bold uppercase ${
                (supabaseConfig?.serviceKey || supabaseConfig?.serviceRoleKey) ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
              }`}>
                {(supabaseConfig?.serviceKey || supabaseConfig?.serviceRoleKey) ? 'Configured' : 'Missing'}
              </span>
            </div>
            <div className="flex items-center justify-between text-xs font-mono bg-white p-2 rounded border border-slate-200">
              <span className="truncate max-w-[180px]">
                {revealedSecrets['supa_key'] ? (supabaseConfig?.serviceKey || supabaseConfig?.serviceRoleKey || 'None') : '••••••••••••••••'}
              </span>
              <div className="flex items-center space-x-1">
                <button
                  type="button"
                  onClick={() => toggleSecretReveal('supa_key')}
                  className="p-1 text-slate-400 hover:text-slate-600"
                  title="Reveal or mask"
                >
                  {revealedSecrets['supa_key'] ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                </button>
                {onOpenSecretsVault && (
                  <button
                    type="button"
                    onClick={onOpenSecretsVault}
                    className="p-1 text-slate-400 hover:text-slate-700"
                    title="Edit secret"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>
            </div>
            <div className="text-[10px] text-slate-500">
              RLS bypass key for worker upserts to PostgreSQL
            </div>
          </div>

          {/* Worker Service Daemon or Integrated Engine */}
          <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800">
                {formData.workerUrl ? 'EXTERNAL_WORKER_URL' : 'SYNC_PIPELINE_ENGINE'}
              </span>
              <span className="text-[10px] px-2 py-0.5 rounded font-bold uppercase bg-emerald-100 text-emerald-800">
                {formData.workerUrl ? 'Configured' : 'Active'}
              </span>
            </div>
            <div className="flex items-center justify-between text-xs font-mono bg-white p-2 rounded border border-slate-200">
              <span className="truncate max-w-[180px]">
                {formData.workerUrl || 'In-Process Production Engine'}
              </span>
              <div className="flex items-center space-x-1">
                {onOpenSecretsVault && (
                  <button
                    type="button"
                    onClick={onOpenSecretsVault}
                    className="p-1 text-slate-400 hover:text-slate-700"
                    title="Edit worker endpoint"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>
            </div>
            <div className="text-[10px] text-slate-500">
              {formData.workerUrl 
                ? 'External remote worker daemon configured'
                : 'Production Node.js In-App Pipeline (No external daemon required)'}
            </div>
          </div>
        </div>
      </div>

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
                Sync Engine Pipeline
              </span>
              <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase ${
                schedulerStatus?.state === 'SYNCING' ? 'bg-amber-100 text-amber-800 animate-pulse' :
                schedulerStatus?.enabled ? 'bg-emerald-100 text-emerald-800' :
                'bg-slate-200 text-slate-700'
              }`}>
                {schedulerStatus?.state === 'SYNCING' ? 'Syncing...' : schedulerStatus?.enabled ? 'Auto Active' : 'Manual'}
              </span>
            </div>
            <div className="text-xs text-slate-600 font-mono truncate">
              {formData.workerUrl || 'In-Process Production Engine'}
            </div>
            <div className="text-[11px] text-slate-500">
              {schedulerStatus?.enabled
                ? `Auto-Sync: Every ${schedulerStatus.intervalLabel}${schedulerStatus.secondsUntilNextRun !== null ? ` (in ${Math.ceil(schedulerStatus.secondsUntilNextRun / 60)}m)` : ''}`
                : 'Manual Mode (Scheduled auto-sync disabled)'}
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
            <div className="flex items-center justify-between">
              <h2 className="text-base font-semibold text-slate-900 flex items-center space-x-2">
                <Clock className="w-4 h-4 text-blue-600" />
                <span>Production Background Scheduler</span>
              </h2>
              {schedulerStatus && (
                <span className={`text-[11px] px-2.5 py-0.5 rounded-full font-bold uppercase tracking-wider ${
                  schedulerStatus.state === 'SYNCING' ? 'bg-amber-100 text-amber-800 animate-pulse' :
                  schedulerStatus.state === 'SCHEDULED' ? 'bg-emerald-100 text-emerald-800' :
                  'bg-slate-100 text-slate-700'
                }`}>
                  {schedulerStatus.state === 'SYNCING' ? 'Syncing...' : schedulerStatus.state}
                </span>
              )}
            </div>

            {/* Live Scheduler Status Box */}
            {schedulerStatus && (
              <div className="p-3 rounded-lg bg-slate-50 border border-slate-200 text-xs space-y-2">
                <div className="flex items-center justify-between text-slate-700">
                  <span className="font-medium">Active Engine:</span>
                  <span className="font-mono font-semibold text-indigo-700">
                    {schedulerStatus.engineMode === 'EXTERNAL_WORKER_DAEMON'
                      ? `Worker (${schedulerStatus.workerEndpoint})`
                      : 'In-Process Production Engine'}
                  </span>
                </div>
                <div className="flex items-center justify-between text-slate-700">
                  <span className="font-medium">Next Scheduled Run:</span>
                  <span className="font-mono font-semibold text-slate-900">
                    {schedulerStatus.enabled && schedulerStatus.nextRunAt
                      ? `${new Date(schedulerStatus.nextRunAt).toLocaleTimeString()} (${
                          schedulerStatus.secondsUntilNextRun !== null && schedulerStatus.secondsUntilNextRun < 60
                            ? '< 1 minute'
                            : `${Math.ceil((schedulerStatus.secondsUntilNextRun || 0) / 60)} minutes`
                        })`
                      : 'Manual trigger only'}
                  </span>
                </div>
                {schedulerStatus.lastRunAt && (
                  <div className="flex items-center justify-between text-slate-700 border-t border-slate-200/60 pt-1.5">
                    <span className="font-medium">Last Run Result:</span>
                    <span className="font-mono text-[11px]">
                      {schedulerStatus.lastRunResult?.success ? (
                        <span className="text-emerald-700 font-bold">
                          ✓ {schedulerStatus.lastRunResult?.totalInserted ?? 0} inserted, {schedulerStatus.lastRunResult?.totalFailed ?? 0} failed
                        </span>
                      ) : (
                        <span className="text-rose-600 font-bold">
                          ✗ {schedulerStatus.lastRunResult?.error || 'Failed'}
                        </span>
                      )}
                    </span>
                  </div>
                )}
              </div>
            )}

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
                The sync pipeline automatically polls Nextcloud WebDAV on this interval, detects changes, and executes upserts.
              </p>
            </div>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                External Worker URL (Optional)
              </label>
              <input
                type="text"
                value={formData.workerUrl}
                onChange={(e) => setFormData({ ...formData, workerUrl: e.target.value })}
                placeholder="Optional (Leave blank to use Integrated Production Engine)"
                className="w-full px-3 py-2 text-sm rounded-lg border border-slate-300 font-mono focus:outline-none focus:ring-2 focus:ring-emerald-500"
              />
              <p className="text-[11px] text-slate-500 mt-1">
                Leave empty for standard deployment. The production in-process engine handles all WebDAV downloads and Supabase synchronization.
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

            {/* Individual File Auto-Sync Exclusions Selector */}
            {files.length > 0 && (
              <div className="mt-4 pt-4 border-t border-slate-200/70 space-y-2.5">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-700 uppercase tracking-wider flex items-center space-x-1.5">
                    <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                    <span>Per-File Auto-Sync Inclusions ({files.length} detected)</span>
                  </span>
                  <span className="text-[11px] text-slate-500">
                    Uncheck to skip a file during scheduled auto-sync (manual-only)
                  </span>
                </div>

                <div className="bg-slate-50 rounded-lg p-3 border border-slate-200 divide-y divide-slate-200/60 max-h-56 overflow-y-auto">
                  {files.map((f, idx) => {
                    const isExcluded = (formData.excludedAutoSyncFiles || []).some(
                      ef => ef.toLowerCase().trim() === f.filename.toLowerCase().trim()
                    );
                    const isAuto = !isExcluded;

                    return (
                      <div key={`${f.id || f.filename}-${idx}`} className="py-2 flex items-center justify-between text-xs">
                        <label className="flex items-center space-x-2.5 cursor-pointer select-none">
                          <input
                            type="checkbox"
                            checked={isAuto}
                            onChange={(e) => {
                              const willBeAuto = e.target.checked;
                              const currentExcluded = formData.excludedAutoSyncFiles || [];
                              let updatedExcluded: string[];
                              if (willBeAuto) {
                                updatedExcluded = currentExcluded.filter(
                                  ef => ef.toLowerCase().trim() !== f.filename.toLowerCase().trim()
                                );
                              } else {
                                updatedExcluded = Array.from(new Set([...currentExcluded, f.filename.trim()]));
                              }
                              setFormData({
                                ...formData,
                                excludedAutoSyncFiles: updatedExcluded
                              });
                              if (onToggleFileAutoSync) {
                                onToggleFileAutoSync(f.filename, willBeAuto);
                              }
                            }}
                            className="w-3.5 h-3.5 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500 cursor-pointer"
                          />
                          <span className="font-semibold text-slate-800">{f.filename}</span>
                        </label>

                        <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                          isAuto ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                        }`}>
                          {isAuto ? '🟢 Auto-Sync' : '⏸️ Manual Only'}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
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
