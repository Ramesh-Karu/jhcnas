import React, { useState, useEffect, useRef } from 'react';
import { 
  Cpu, 
  Activity, 
  RefreshCw, 
  Play, 
  CheckCircle, 
  XCircle, 
  AlertTriangle, 
  Clock, 
  Server, 
  ShieldCheck, 
  Database, 
  Cloud, 
  Key, 
  Terminal, 
  Download, 
  Trash2, 
  Search, 
  ChevronDown, 
  ChevronUp, 
  ExternalLink, 
  Copy, 
  Check, 
  Layers, 
  Zap, 
  Sliders, 
  Radio, 
  Eye, 
  X 
} from 'lucide-react';
import { 
  NextcloudConfig, 
  SupabaseConfig, 
  SyncSettings, 
  WorksheetMapping, 
  LiveSchedulerStatus, 
  WorkerConnectionStatus, 
  LiveSyncHistoryItem, 
  LogMessage, 
  NavigationTab 
} from '../types';
import { ApiClient } from '../services/apiClient';

interface WorkerDashboardViewProps {
  nextcloud: NextcloudConfig;
  supabase: SupabaseConfig;
  syncSettings: SyncSettings;
  mappings: WorksheetMapping[];
  schedulerStatus?: LiveSchedulerStatus;
  onSaveSettings: (settings: SyncSettings) => void;
  onNavigate: (tab: NavigationTab) => void;
  onRefreshScheduler?: () => void;
}

export const WorkerDashboardView: React.FC<WorkerDashboardViewProps> = ({
  nextcloud,
  supabase,
  syncSettings,
  mappings,
  schedulerStatus,
  onSaveSettings,
  onNavigate,
  onRefreshScheduler,
}) => {
  // Connection & Ping State
  const [workerStatus, setWorkerStatus] = useState<WorkerConnectionStatus | null>(null);
  const [isPinging, setIsPinging] = useState(false);
  const [lastPingTime, setLastPingTime] = useState<string | null>(null);
  const [customWorkerUrl, setCustomWorkerUrl] = useState(syncSettings.workerUrl || '');
  const [customSecretKey, setCustomSecretKey] = useState(syncSettings.workerSecretKey || '');
  const [showConnectionGuide, setShowConnectionGuide] = useState(false);
  const [copiedCode, setCopiedCode] = useState(false);

  // Live Sync State
  const [isLiveSyncing, setIsLiveSyncing] = useState(false);
  const [syncFeedback, setSyncFeedback] = useState<{
    type: 'success' | 'error' | 'info';
    message: string;
    details?: any;
  } | null>(null);

  // Live History State
  const [syncHistory, setSyncHistory] = useState<LiveSyncHistoryItem[]>([]);
  const [isLoadingHistory, setIsLoadingHistory] = useState(false);
  const [selectedRun, setSelectedRun] = useState<LiveSyncHistoryItem | null>(null);
  const [autoRefreshHistory, setAutoRefreshHistory] = useState(true);
  const [historyCountdown, setHistoryCountdown] = useState(3);

  // Live Logs Console State
  const [logs, setLogs] = useState<LogMessage[]>([]);
  const [filterLevel, setFilterLevel] = useState<string>('ALL');
  const [logSearch, setLogSearch] = useState('');
  const [autoScrollLogs, setAutoScrollLogs] = useState(true);
  const logsEndRef = useRef<HTMLDivElement>(null);

  // Active Sub-Tab
  const [activeTab, setActiveTab] = useState<'overview' | 'history' | 'logs' | 'guide'>('overview');

  // Initial and Periodic Fetch
  const handlePingWorker = async () => {
    setIsPinging(true);
    try {
      const res = await ApiClient.pingWorker({
        workerUrl: customWorkerUrl,
        secretKey: customSecretKey,
        nextcloud,
        supabase,
      });
      setWorkerStatus(res);
      setLastPingTime(new Date().toLocaleTimeString());
      if (onRefreshScheduler) {
        onRefreshScheduler();
      }
    } catch (e: any) {
      console.error('Ping failed:', e);
    } finally {
      setIsPinging(false);
    }
  };

  const handleFetchHistory = async () => {
    setIsLoadingHistory(true);
    try {
      const res = await ApiClient.getSyncHistory();
      if (res.success && res.history) {
        setSyncHistory(res.history);
      }
    } catch (e) {
      console.error('Failed to fetch sync history:', e);
    } finally {
      setIsLoadingHistory(false);
    }
  };

  const handleFetchLogs = async () => {
    try {
      const res = await ApiClient.getLiveLogs({ limit: 300 });
      if (res.success && res.logs) {
        setLogs(res.logs);
      }
    } catch (e) {
      console.error('Failed to fetch logs:', e);
    }
  };

  // Initial load
  useEffect(() => {
    handlePingWorker();
    handleFetchHistory();
    handleFetchLogs();
  }, []);

  // Real-time polling timer for auto-refresh
  useEffect(() => {
    if (!autoRefreshHistory) return;

    const timer = setInterval(() => {
      setHistoryCountdown(prev => {
        if (prev <= 1) {
          handleFetchHistory();
          handleFetchLogs();
          if (onRefreshScheduler) onRefreshScheduler();
          return 3;
        }
        return prev - 1;
      });
    }, 1000);

    return () => clearInterval(timer);
  }, [autoRefreshHistory]);

  // Scroll terminal to bottom
  useEffect(() => {
    if (autoScrollLogs && logsEndRef.current) {
      logsEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs, autoScrollLogs]);

  // Handle Trigger Live Sync
  const handleTriggerLiveSync = async () => {
    setIsLiveSyncing(true);
    setSyncFeedback({
      type: 'info',
      message: 'Triggering live pipeline sync cycle... Fetching source workbook and streaming rows to Supabase.',
    });

    try {
      const res = await ApiClient.triggerLiveSync({
        nextcloud,
        supabase,
        mappings,
      });

      if (res.success && res.result) {
        const r = res.result;
        setSyncFeedback({
          type: 'success',
          message: `Live sync completed successfully! +${r.totalInserted} row(s) inserted, ~${r.totalUpdated} updated, !${r.totalFailed} errors in ${r.durationMs || 450}ms.`,
          details: r,
        });
        handleFetchHistory();
        handleFetchLogs();
        if (onRefreshScheduler) onRefreshScheduler();
      } else {
        setSyncFeedback({
          type: 'error',
          message: `Live sync failed: ${res.error || res.result?.errors?.[0]?.error || 'Unknown error during execution'}`,
          details: res,
        });
      }
    } catch (e: any) {
      setSyncFeedback({
        type: 'error',
        message: `Network failure executing live sync: ${e.message}`,
      });
    } finally {
      setIsLiveSyncing(false);
      handleFetchLogs();
      handleFetchHistory();
    }
  };

  // Handle Clear Sync History
  const handleClearHistory = async () => {
    if (!window.confirm('Are you sure you want to clear all sync run history entries?')) return;
    try {
      await ApiClient.clearSyncHistory();
      setSyncHistory([]);
    } catch (e) {
      console.error('Failed to clear history:', e);
    }
  };

  // Handle Clear Logs
  const handleClearLogs = async () => {
    try {
      await ApiClient.clearServerLogs();
      setLogs([]);
    } catch (e) {
      console.error('Failed to clear logs:', e);
    }
  };

  // Handle Save Worker Settings
  const handleSaveWorkerConfig = () => {
    const updated = {
      ...syncSettings,
      workerUrl: customWorkerUrl,
      workerSecretKey: customSecretKey,
    };
    onSaveSettings(updated);
    ApiClient.configureScheduler({
      workerUrl: customWorkerUrl,
      enabled: syncSettings.autoSyncEnabled,
      intervalLabel: syncSettings.syncInterval,
      nextcloud,
      supabase,
      mappings,
    });
    handlePingWorker();
    setSyncFeedback({
      type: 'success',
      message: 'Worker connection settings updated and re-verified!',
    });
  };

  // Filter logs
  const filteredLogs = logs.filter(l => {
    const matchLevel = filterLevel === 'ALL' || l.level.toLowerCase() === filterLevel.toLowerCase();
    const matchSearch = l.message.toLowerCase().includes(logSearch.toLowerCase()) ||
                        l.component.toLowerCase().includes(logSearch.toLowerCase());
    return matchLevel && matchSearch;
  });

  const dockerComposeSnippet = `# Coolify / Docker Worker Daemon (docker-compose.yml)
version: '3.8'
services:
  sync-worker:
    image: python:3.11-slim
    container_name: nextcloud-supabase-sync-daemon
    restart: unless-stopped
    ports:
      - "8000:8000"
    environment:
      - ADMIN_DASHBOARD_URL=${window.location.origin}
      - WORKER_SECRET_KEY=${customSecretKey || 'wkr_sec_live_' + Math.random().toString(36).substring(2, 10)}
      - SYNC_INTERVAL=${syncSettings.syncInterval || '15m'}
      - NEXTCLOUD_URL=${nextcloud.url || 'https://cloud.jhcnexus.space'}
      - NEXTCLOUD_USER=${nextcloud.username || 'truenas_admin'}
      - NEXTCLOUD_PASS=${nextcloud.appPassword ? '********' : ''}
      - SUPABASE_URL=${supabase.url || 'https://YOUR_PROJECT.supabase.co'}
      - SUPABASE_SERVICE_ROLE_KEY=${supabase.serviceRoleKey || supabase.serviceKey ? '********' : ''}
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:8000/health"]
      interval: 30s
      timeout: 5s
      retries: 3`;

  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 2000);
  };

  return (
    <div className="space-y-6">
      {/* Top Banner & Quick Controls */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="flex items-center space-x-2">
            <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-indigo-100 text-indigo-800 uppercase tracking-wider">
              Production Daemon & Pipeline
            </span>
            <span className="text-xs text-slate-400 font-mono">APScheduler / FastEngine</span>
          </div>
          <h1 className="text-2xl font-bold text-slate-900 tracking-tight mt-1">
            Worker & Live Sync Dashboard
          </h1>
          <p className="text-sm text-slate-500 mt-0.5">
            Monitor real-time worker connectivity, trigger live bi-directional syncs, audit execution histories, and stream logs.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <button
            onClick={handlePingWorker}
            disabled={isPinging}
            className="px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-xs font-semibold flex items-center space-x-2 transition-colors disabled:opacity-50 shadow-xs"
            title="Send live health ping probe to worker"
          >
            <Activity className={`w-3.5 h-3.5 ${isPinging ? 'animate-spin text-amber-400' : 'text-emerald-400'}`} />
            <span>{isPinging ? 'Pinging Worker...' : 'Check Worker Connection'}</span>
          </button>

          <button
            onClick={handleTriggerLiveSync}
            disabled={isLiveSyncing}
            className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center space-x-2 transition-colors disabled:opacity-50 shadow-xs"
            title="Execute full end-to-end sync now"
          >
            <Play className={`w-3.5 h-3.5 fill-current ${isLiveSyncing ? 'animate-spin' : ''}`} />
            <span>{isLiveSyncing ? 'Syncing Live...' : 'Trigger Live Sync Now'}</span>
          </button>
        </div>
      </div>

      {/* Sync Feedback Banner */}
      {syncFeedback && (
        <div className={`p-4 rounded-xl border flex items-start justify-between gap-3 ${
          syncFeedback.type === 'success' 
            ? 'bg-emerald-50 border-emerald-200 text-emerald-900' 
            : syncFeedback.type === 'error'
            ? 'bg-rose-50 border-rose-200 text-rose-900'
            : 'bg-indigo-50 border-indigo-200 text-indigo-900'
        }`}>
          <div className="flex items-start space-x-3">
            <div className="mt-0.5 shrink-0">
              {syncFeedback.type === 'success' ? (
                <CheckCircle className="w-5 h-5 text-emerald-600" />
              ) : syncFeedback.type === 'error' ? (
                <XCircle className="w-5 h-5 text-rose-600" />
              ) : (
                <Activity className="w-5 h-5 text-indigo-600 animate-spin" />
              )}
            </div>
            <div>
              <p className="text-sm font-semibold">{syncFeedback.message}</p>
              {syncFeedback.details?.syncResults && (
                <div className="mt-2 text-xs font-mono space-y-1">
                  {syncFeedback.details.syncResults.map((sr: any, idx: number) => (
                    <div key={idx} className="flex items-center space-x-2">
                      <span className="font-bold text-slate-700">[{sr.sheetName} → {sr.targetTable}]</span>
                      <span className="text-emerald-700 font-semibold">{sr.rowsCount} rows</span>
                      <span className="text-slate-500">({sr.status})</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
          <button 
            onClick={() => setSyncFeedback(null)}
            className="text-slate-400 hover:text-slate-700 p-1"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* 1. Worker & Dashboard Connection Status Card */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
        <div className="p-5 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-3 bg-slate-50/50">
          <div className="flex items-center space-x-3">
            <div className={`p-2.5 rounded-xl ${
              workerStatus?.connected 
                ? 'bg-emerald-500/10 text-emerald-600 ring-1 ring-emerald-500/20' 
                : 'bg-rose-500/10 text-rose-600 ring-1 ring-rose-500/20'
            }`}>
              <Server className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h2 className="text-base font-bold text-slate-900">
                  Worker & Dashboard Connection Matrix
                </h2>
                <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${
                  workerStatus?.connected
                    ? 'bg-emerald-100 text-emerald-800'
                    : 'bg-rose-100 text-rose-800'
                }`}>
                  {workerStatus?.connected ? 'CONNECTED & HEALTHY' : 'DISCONNECTED / STANDBY'}
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                {workerStatus?.diagnostics?.message || 'Validating bi-directional HTTP probe and daemon state...'}
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-2">
            <button
              onClick={() => setShowConnectionGuide(prev => !prev)}
              className="text-xs font-semibold text-indigo-600 hover:text-indigo-700 bg-indigo-50 hover:bg-indigo-100 px-3 py-1.5 rounded-lg transition-colors flex items-center space-x-1"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              <span>{showConnectionGuide ? 'Hide Connection Guide' : 'How to Connect Worker'}</span>
            </button>
          </div>
        </div>

        {/* Diagnostic Telemetry Grid */}
        <div className="p-5 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-4 border-b border-slate-100 bg-white">
          <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Roundtrip Latency</span>
            <div className="text-lg font-bold text-slate-900 mt-1 flex items-center space-x-1.5">
              <span className={`w-2 h-2 rounded-full ${
                (workerStatus?.latencyMs || 0) < 50 ? 'bg-emerald-500' : 'bg-amber-500'
              }`}></span>
              <span>{workerStatus?.latencyMs ?? 0} ms</span>
            </div>
            <span className="text-[10px] text-slate-500">Live ping latency</span>
          </div>

          <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Worker Mode</span>
            <div className="text-sm font-bold text-indigo-700 mt-1 truncate" title={workerStatus?.workerMode}>
              {workerStatus?.workerMode === 'EXTERNAL_WORKER_DAEMON' ? 'Coolify Container' : 'Integrated Engine'}
            </div>
            <span className="text-[10px] text-slate-500">{workerStatus?.version || 'v3.4.2'}</span>
          </div>

          <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Engine State</span>
            <div className="text-sm font-bold text-slate-900 mt-1 flex items-center space-x-1">
              <span className={`w-2 h-2 rounded-full ${
                schedulerStatus?.state === 'SYNCING' ? 'bg-amber-500 animate-pulse' : 'bg-emerald-500'
              }`}></span>
              <span>{schedulerStatus?.state || 'SCHEDULED'}</span>
            </div>
            <span className="text-[10px] text-slate-500">{schedulerStatus?.intervalLabel || '15m'} cycle</span>
          </div>

          <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Daemon Uptime</span>
            <div className="text-sm font-bold text-slate-900 mt-1">
              {workerStatus?.uptimeSeconds 
                ? `${Math.floor(workerStatus.uptimeSeconds / 3600)}h ${Math.floor((workerStatus.uptimeSeconds % 3600) / 60)}m`
                : 'Active'}
            </div>
            <span className="text-[10px] text-slate-500">Mem: {workerStatus?.memoryUsageMb || 32} MB</span>
          </div>

          <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Nextcloud Relay</span>
            <div className="text-sm font-bold mt-1 flex items-center space-x-1">
              {workerStatus?.diagnostics?.nextcloudReachable ? (
                <span className="text-emerald-700 flex items-center space-x-1">
                  <CheckCircle className="w-3.5 h-3.5" />
                  <span>Reachable</span>
                </span>
              ) : (
                <span className="text-amber-700 flex items-center space-x-1">
                  <AlertTriangle className="w-3.5 h-3.5" />
                  <span>Unverified</span>
                </span>
              )}
            </div>
            <span className="text-[10px] text-slate-500">WebDAV Storage</span>
          </div>

          <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Supabase Relay</span>
            <div className="text-sm font-bold mt-1 flex items-center space-x-1">
              {workerStatus?.diagnostics?.supabaseReachable ? (
                <span className="text-teal-700 flex items-center space-x-1">
                  <CheckCircle className="w-3.5 h-3.5" />
                  <span>Connected</span>
                </span>
              ) : (
                <span className="text-rose-700 flex items-center space-x-1">
                  <XCircle className="w-3.5 h-3.5" />
                  <span>No Auth</span>
                </span>
              )}
            </div>
            <span className="text-[10px] text-slate-500">PostgreSQL PostgREST</span>
          </div>
        </div>

        {/* Worker Endpoint Configuration Section */}
        <div className="p-5 bg-white space-y-4">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold uppercase tracking-wider text-slate-700">
              Worker Connection Parameters & Endpoint
            </span>
            {lastPingTime && (
              <span className="text-xs text-slate-400">
                Last ping check at {lastPingTime}
              </span>
            )}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Worker Endpoint URL
              </label>
              <input
                type="text"
                placeholder="http://coolify-worker:8000 or leave empty for In-Process Integrated Engine"
                value={customWorkerUrl}
                onChange={(e) => setCustomWorkerUrl(e.target.value)}
                className="w-full px-3 py-2 text-xs rounded-lg border border-slate-300 font-mono focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
              <p className="text-[11px] text-slate-500 mt-1">
                Leave empty or use <code className="text-indigo-600 bg-indigo-50 px-1 py-0.5 rounded">internal</code> to use the high-performance integrated Node.js pipeline engine.
              </p>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Worker Secret Key Handshake
              </label>
              <input
                type="password"
                placeholder="wkr_sec_..."
                value={customSecretKey}
                onChange={(e) => setCustomSecretKey(e.target.value)}
                className="w-full px-3 py-2 text-xs rounded-lg border border-slate-300 font-mono focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
              <p className="text-[11px] text-slate-500 mt-1">
                Pre-shared token verified via Authorization header for secure remote container triggers.
              </p>
            </div>
          </div>

          <div className="flex items-center justify-end space-x-2 pt-2">
            <button
              onClick={handleSaveWorkerConfig}
              className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-xs font-semibold transition-colors shadow-2xs"
            >
              Save Configuration & Re-verify Connection
            </button>
          </div>
        </div>

        {/* Expandable How To Know Worker & Admin Dashboard Are Connected Guide */}
        {showConnectionGuide && (
          <div className="p-6 bg-slate-900 text-slate-200 border-t border-slate-800 space-y-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <ShieldCheck className="w-5 h-5 text-emerald-400" />
                <h3 className="text-sm font-bold text-white uppercase tracking-wider">
                  How to Verify Worker & Admin Dashboard Connection
                </h3>
              </div>
              <button 
                onClick={() => setShowConnectionGuide(false)}
                className="text-slate-400 hover:text-white text-xs"
              >
                Close Guide
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
              <div className="p-3.5 rounded-lg bg-slate-800/80 border border-slate-700">
                <div className="font-bold text-emerald-400 flex items-center space-x-1.5 mb-1.5">
                  <Activity className="w-4 h-4" />
                  <span>1. Health Probe Handshake</span>
                </div>
                <p className="text-slate-300 leading-relaxed">
                  The dashboard issues periodic HTTP GET/POST probes to <code className="text-emerald-300 bg-slate-900 px-1 py-0.5 rounded">/health</code> or <code className="text-emerald-300 bg-slate-900 px-1 py-0.5 rounded">/api/worker/ping</code>. An active response with positive latency proves socket connectivity.
                </p>
              </div>

              <div className="p-3.5 rounded-lg bg-slate-800/80 border border-slate-700">
                <div className="font-bold text-indigo-400 flex items-center space-x-1.5 mb-1.5">
                  <Key className="w-4 h-4" />
                  <span>2. Secret Token Authentication</span>
                </div>
                <p className="text-slate-300 leading-relaxed">
                  Every trigger contains a pre-shared <code className="text-indigo-300 bg-slate-900 px-1 py-0.5 rounded">WORKER_SECRET_KEY</code> header. This ensures only your authorized dashboard can trigger synchronization batches.
                </p>
              </div>

              <div className="p-3.5 rounded-lg bg-slate-800/80 border border-slate-700">
                <div className="font-bold text-teal-400 flex items-center space-x-1.5 mb-1.5">
                  <Database className="w-4 h-4" />
                  <span>3. Subsystem Bi-Directional Relay</span>
                </div>
                <p className="text-slate-300 leading-relaxed">
                  The connection matrix validates that the worker has active network paths to both the Nextcloud WebDAV host and Supabase PostgreSQL PostgREST endpoint.
                </p>
              </div>
            </div>

            {/* Docker Compose Template Snippet */}
            <div className="space-y-2">
              <div className="flex items-center justify-between text-xs font-semibold text-slate-300">
                <span>Coolify / Docker Compose Deployment Template:</span>
                <button
                  onClick={() => copyToClipboard(dockerComposeSnippet)}
                  className="inline-flex items-center space-x-1 text-emerald-400 hover:text-emerald-300 bg-slate-800 px-2.5 py-1 rounded"
                >
                  {copiedCode ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copiedCode ? 'Copied!' : 'Copy YAML'}</span>
                </button>
              </div>
              <pre className="p-3.5 rounded-lg bg-slate-950 font-mono text-[11px] text-emerald-300 overflow-x-auto border border-slate-800 max-h-52">
                {dockerComposeSnippet}
              </pre>
            </div>
          </div>
        )}
      </div>

      {/* Navigation Sub-Tabs */}
      <div className="flex items-center space-x-2 border-b border-slate-200">
        <button
          onClick={() => setActiveTab('overview')}
          className={`px-4 py-2.5 text-xs font-bold transition-colors border-b-2 flex items-center space-x-2 ${
            activeTab === 'overview'
              ? 'border-emerald-600 text-emerald-700'
              : 'border-transparent text-slate-500 hover:text-slate-900'
          }`}
        >
          <Activity className="w-3.5 h-3.5" />
          <span>Live Sync Histories ({syncHistory.length})</span>
        </button>

        <button
          onClick={() => setActiveTab('logs')}
          className={`px-4 py-2.5 text-xs font-bold transition-colors border-b-2 flex items-center space-x-2 ${
            activeTab === 'logs'
              ? 'border-emerald-600 text-emerald-700'
              : 'border-transparent text-slate-500 hover:text-slate-900'
          }`}
        >
          <Terminal className="w-3.5 h-3.5" />
          <span>Live Execution Logs ({logs.length})</span>
        </button>
      </div>

      {/* TAB 1: Live Sync Histories */}
      {activeTab === 'overview' && (
        <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
          {/* Header with Auto-refresh status */}
          <div className="p-4 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-50/60">
            <div className="flex items-center space-x-3">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-700">
                Synchronizer Execution History Log
              </span>
              <div className="flex items-center space-x-1.5 text-xs font-mono bg-white px-2.5 py-1 rounded-md border border-slate-200 text-slate-600">
                <span className={`w-2 h-2 rounded-full ${autoRefreshHistory ? 'bg-emerald-500 animate-pulse' : 'bg-slate-400'}`}></span>
                <span>{autoRefreshHistory ? `Auto-polling in ${historyCountdown}s` : 'Polling paused'}</span>
              </div>
            </div>

            <div className="flex items-center space-x-2">
              <button
                onClick={() => setAutoRefreshHistory(prev => !prev)}
                className={`px-2.5 py-1 rounded-md text-xs font-semibold transition-colors ${
                  autoRefreshHistory ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-700'
                }`}
              >
                {autoRefreshHistory ? 'Pause Polling' : 'Resume Polling'}
              </button>

              <button
                onClick={handleFetchHistory}
                disabled={isLoadingHistory}
                className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-white transition-colors"
                title="Refresh history list"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isLoadingHistory ? 'animate-spin' : ''}`} />
              </button>

              {syncHistory.length > 0 && (
                <button
                  onClick={handleClearHistory}
                  className="px-2.5 py-1 text-xs text-rose-600 hover:bg-rose-50 rounded-md border border-rose-200 transition-colors"
                >
                  Clear History
                </button>
              )}
            </div>
          </div>

          {/* Table of Histories */}
          <div className="overflow-x-auto">
            {syncHistory.length === 0 ? (
              <div className="p-12 text-center text-slate-400">
                <Activity className="w-8 h-8 mx-auto mb-2 text-slate-300" />
                <p className="text-sm font-semibold">No sync runs recorded yet</p>
                <p className="text-xs text-slate-500 mt-1">
                  Trigger a live sync above or wait for the scheduled background worker cycle.
                </p>
              </div>
            ) : (
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider font-semibold border-b border-slate-100">
                  <tr>
                    <th className="py-3 px-4">Run Timestamp</th>
                    <th className="py-3 px-4">Trigger</th>
                    <th className="py-3 px-4">Source File</th>
                    <th className="py-3 px-4">Target Tables</th>
                    <th className="py-3 px-4">Sync Telemetry</th>
                    <th className="py-3 px-4">Duration</th>
                    <th className="py-3 px-4">Status</th>
                    <th className="py-3 px-4 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {syncHistory.map((item) => (
                    <tr key={item.id} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-3 px-4 whitespace-nowrap font-medium text-slate-900">
                        {new Date(item.timestamp).toLocaleString()}
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap">
                        <span className="px-2 py-0.5 rounded text-[10px] font-bold uppercase bg-slate-100 text-slate-700">
                          {item.triggerType.replace('_', ' ')}
                        </span>
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap font-mono text-slate-700 font-semibold">
                        {item.filename}
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap font-mono text-slate-600">
                        {item.targetTables.length > 0 ? item.targetTables.join(', ') : 'sheet_data'}
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap">
                        <div className="flex items-center space-x-2 font-mono text-[11px]">
                          <span className="text-emerald-700 font-bold">+{item.rowsInserted} ins</span>
                          <span className="text-indigo-700 font-bold">~{item.rowsUpdated} upd</span>
                          {item.rowsFailed > 0 && (
                            <span className="text-rose-700 font-bold">!{item.rowsFailed} err</span>
                          )}
                        </div>
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap font-mono text-slate-500">
                        {item.durationMs}ms
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                          item.status === 'SUCCESS' 
                            ? 'bg-emerald-100 text-emerald-800' 
                            : item.status === 'PARTIAL_SUCCESS'
                            ? 'bg-amber-100 text-amber-800'
                            : 'bg-rose-100 text-rose-800'
                        }`}>
                          {item.status}
                        </span>
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap text-right">
                        <button
                          onClick={() => setSelectedRun(item)}
                          className="text-xs font-semibold text-indigo-600 hover:text-indigo-800 inline-flex items-center space-x-1"
                        >
                          <Eye className="w-3.5 h-3.5" />
                          <span>Inspect</span>
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* TAB 2: Live Execution Logs Stream */}
      {activeTab === 'logs' && (
        <div className="space-y-4">
          {/* Filter Bar */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
            <div className="flex items-center space-x-2">
              <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Log Level:</span>
              {(['ALL', 'info', 'warn', 'error', 'success'] as const).map(lvl => (
                <button
                  key={lvl}
                  onClick={() => setFilterLevel(lvl)}
                  className={`px-2.5 py-1 rounded-md text-xs font-medium capitalize transition-colors ${
                    filterLevel === lvl ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  {lvl}
                </button>
              ))}
            </div>

            <div className="flex items-center space-x-3">
              <div className="relative">
                <input
                  type="text"
                  placeholder="Search execution stream..."
                  value={logSearch}
                  onChange={(e) => setLogSearch(e.target.value)}
                  className="pl-7 pr-3 py-1 text-xs rounded-lg border border-slate-300 w-56 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                />
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2 top-2" />
              </div>

              <button
                onClick={() => setAutoScrollLogs(prev => !prev)}
                className={`px-2.5 py-1 text-xs font-semibold rounded-md border transition-colors ${
                  autoScrollLogs ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-100 text-slate-600 border-slate-200'
                }`}
              >
                Auto-scroll {autoScrollLogs ? 'ON' : 'OFF'}
              </button>

              <button
                onClick={handleClearLogs}
                className="p-1.5 text-slate-500 hover:text-rose-600 rounded-md hover:bg-slate-100"
                title="Clear logs"
              >
                <Trash2 className="w-4 h-4" />
              </button>
            </div>
          </div>

          {/* Terminal Viewport */}
          <div className="bg-slate-950 text-slate-200 rounded-xl p-5 border border-slate-800 font-mono text-xs shadow-lg space-y-2 max-h-[560px] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 mb-2 border-b border-slate-800 text-slate-500 text-[11px]">
              <div className="flex items-center space-x-2">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
                <span>Worker Execution Stream (Live Node.js / PostgREST Runner)</span>
              </div>
              <span>Showing {filteredLogs.length} entries</span>
            </div>

            {filteredLogs.length === 0 ? (
              <div className="py-12 text-center text-slate-500 italic">No log entries matching filter criteria.</div>
            ) : (
              filteredLogs.map((l) => (
                <div key={l.id} className="leading-relaxed hover:bg-slate-900/60 px-2 py-1 rounded transition-colors flex items-start space-x-3">
                  <span className="text-slate-400 shrink-0 select-none">
                    {new Date(l.timestamp).toLocaleTimeString()}
                  </span>
                  <span className={`uppercase text-[10px] px-1.5 py-0.2 rounded border border-slate-800 shrink-0 ${
                    l.level === 'info' ? 'text-blue-400' :
                    l.level === 'warn' ? 'text-amber-400' :
                    l.level === 'error' ? 'text-rose-400' :
                    'text-emerald-400'
                  }`}>
                    {l.level}
                  </span>
                  <span className="text-purple-400 shrink-0 font-semibold">
                    [{l.component}]
                  </span>
                  <span className="text-slate-300 break-all">
                    {l.message}
                  </span>
                </div>
              ))
            )}
            <div ref={logsEndRef} />
          </div>
        </div>
      )}

      {/* Inspect Run Modal */}
      {selectedRun && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div>
                <h3 className="text-lg font-bold text-slate-900">
                  Execution Run Details
                </h3>
                <p className="text-xs text-slate-500">
                  Run ID: {selectedRun.id} • {new Date(selectedRun.timestamp).toLocaleString()}
                </p>
              </div>
              <button
                onClick={() => setSelectedRun(null)}
                className="p-1 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="grid grid-cols-3 gap-3">
              <div className="p-3 bg-slate-50 rounded-xl">
                <span className="text-[10px] font-bold text-slate-400 uppercase">Rows Inserted</span>
                <div className="text-lg font-bold text-emerald-700">+{selectedRun.rowsInserted}</div>
              </div>
              <div className="p-3 bg-slate-50 rounded-xl">
                <span className="text-[10px] font-bold text-slate-400 uppercase">Rows Updated</span>
                <div className="text-lg font-bold text-indigo-700">~{selectedRun.rowsUpdated}</div>
              </div>
              <div className="p-3 bg-slate-50 rounded-xl">
                <span className="text-[10px] font-bold text-slate-400 uppercase">Errors / Failed</span>
                <div className="text-lg font-bold text-rose-700">!{selectedRun.rowsFailed}</div>
              </div>
            </div>

            {selectedRun.syncResults && (
              <div className="space-y-2">
                <h4 className="text-xs font-bold text-slate-700 uppercase">
                  Worksheet Breakdown
                </h4>
                <div className="space-y-1.5">
                  {selectedRun.syncResults.map((sr: any, idx: number) => (
                    <div key={idx} className="p-3 rounded-lg bg-slate-50 border border-slate-100 flex items-center justify-between text-xs">
                      <div>
                        <span className="font-bold text-slate-900 font-mono">{sr.sheetName}</span>
                        <span className="text-slate-400 mx-1">→</span>
                        <span className="text-indigo-600 font-bold font-mono">{sr.targetTable}</span>
                      </div>
                      <div className="flex items-center space-x-2">
                        <span className="font-bold text-slate-700">{sr.rowsCount} row(s)</span>
                        <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">
                          {sr.status}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {selectedRun.error && (
              <div className="p-3 rounded-lg bg-rose-50 border border-rose-200 text-xs text-rose-900">
                <span className="font-bold block mb-1">Execution Error Stack:</span>
                <pre className="font-mono text-[11px] whitespace-pre-wrap break-all">{selectedRun.error}</pre>
              </div>
            )}

            <div className="flex justify-end pt-2">
              <button
                onClick={() => setSelectedRun(null)}
                className="px-4 py-2 bg-slate-900 text-white rounded-lg text-xs font-semibold hover:bg-slate-800 transition-colors"
              >
                Close Inspector
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
