import React from 'react';
import { 
  Cloud, 
  Database, 
  Clock, 
  CheckCircle2, 
  XCircle, 
  Layers, 
  RefreshCw, 
  ArrowUpRight, 
  Play, 
  FileSpreadsheet, 
  ArrowRight,
  ArrowLeftRight,
  ShieldCheck,
  Zap,
  Activity,
  Key
} from 'lucide-react';
import { 
  NextcloudConfig, 
  SupabaseConfig, 
  SyncSettings, 
  NextcloudFile, 
  ImportLog, 
  NavigationTab,
  LiveSchedulerStatus 
} from '../types';

interface DashboardViewProps {
  nextcloud: NextcloudConfig;
  supabase: SupabaseConfig;
  syncSettings: SyncSettings;
  files: NextcloudFile[];
  importLogs: ImportLog[];
  onNavigate: (tab: NavigationTab) => void;
  onTriggerSync: () => void;
  isSyncing: boolean;
  onOpenSecretsVault?: () => void;
  schedulerStatus?: LiveSchedulerStatus;
}

export const DashboardView: React.FC<DashboardViewProps> = ({
  nextcloud,
  supabase,
  syncSettings,
  files,
  importLogs,
  onNavigate,
  onTriggerSync,
  isSyncing,
  onOpenSecretsVault,
  schedulerStatus,
}) => {
  // Compute Dashboard Metrics
  const filesWaiting = files.filter(f => f.status === 'New' || f.status === 'Modified').length;
  
  const successfulImports = importLogs.filter(l => l.status === 'Success').length;
  const failedImports = importLogs.filter(l => l.status === 'Failed' || l.status === 'Partial Success').length;
  
  const totalRecordsImported = importLogs.reduce((acc, l) => acc + (l.rowsInserted || 0), 0);
  const totalRecordsUpdated = importLogs.reduce((acc, l) => acc + (l.rowsUpdated || 0), 0);

  const lastImport = importLogs.length > 0 ? importLogs[0] : null;

  return (
    <div className="space-y-6">
      {/* Top Banner / System Title */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 tracking-tight">Sync Platform Overview</h1>
          <p className="text-sm text-slate-500 mt-1">
            Automated Nextcloud WebDAV document reader, merged-cell transformer, and Supabase PostgreSQL synchronizer.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2.5">
          <button
            id="btn-quick-worker"
            onClick={() => onNavigate('worker')}
            className="inline-flex items-center space-x-2 px-3.5 py-2 rounded-lg font-bold text-sm text-white bg-slate-900 hover:bg-slate-800 shadow-2xs transition-colors"
          >
            <Activity className="w-4 h-4 text-emerald-400" />
            <span>Worker Dashboard</span>
          </button>
          <button
            id="btn-quick-sync"
            onClick={onTriggerSync}
            disabled={isSyncing}
            className={`inline-flex items-center space-x-2 px-4 py-2 rounded-lg font-medium text-sm text-white shadow-xs transition-colors ${
              isSyncing ? 'bg-slate-400 cursor-not-allowed' : 'bg-emerald-600 hover:bg-emerald-700'
            }`}
          >
            <RefreshCw className={`w-4 h-4 ${isSyncing ? 'animate-spin' : ''}`} />
            <span>{isSyncing ? 'Syncing...' : 'Trigger Sync Cycle'}</span>
          </button>
          <button
            id="btn-quick-dryrun"
            onClick={() => onNavigate('import')}
            className="inline-flex items-center space-x-2 px-3.5 py-2 rounded-lg font-medium text-sm text-slate-700 bg-slate-100 hover:bg-slate-200 transition-colors"
          >
            <Play className="w-4 h-4 text-slate-600" />
            <span>Dry Run</span>
          </button>
          <button
            id="btn-quick-twoway"
            onClick={() => onNavigate('twoway')}
            className="inline-flex items-center space-x-2 px-3.5 py-2 rounded-lg font-semibold text-sm text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 transition-colors"
          >
            <ArrowLeftRight className="w-4 h-4 text-emerald-600" />
            <span>Two-Way Sync Hub</span>
          </button>
          {onOpenSecretsVault && (
            <button
              id="btn-quick-secrets"
              onClick={onOpenSecretsVault}
              className="inline-flex items-center space-x-2 px-3.5 py-2 rounded-lg font-semibold text-sm text-slate-800 bg-white hover:bg-slate-50 border border-slate-300 shadow-2xs transition-colors"
            >
              <Key className="w-4 h-4 text-amber-500" />
              <span>Secrets Vault</span>
            </button>
          )}
        </div>
      </div>

      {/* 8 Required Dashboard Metric Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Nextcloud Status */}
        <div 
          onClick={() => onNavigate('nextcloud')}
          className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors cursor-pointer"
          title="Click to manage Nextcloud WebDAV configuration"
        >
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Nextcloud Status</span>
            <div className={`p-2 rounded-lg ${nextcloud.isConnected ? 'bg-emerald-50 text-emerald-600' : 'bg-amber-50 text-amber-600'}`}>
              <Cloud className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className={`text-xl font-bold ${nextcloud.isConnected ? 'text-emerald-700' : 'text-amber-700'}`}>
              {nextcloud.isConnected ? 'Connected' : 'Unverified'}
            </div>
            <p className="text-xs text-slate-500 mt-1 truncate" title={nextcloud.webdavUrl}>
              {nextcloud.isConnected 
                ? `TrueNAS SCALE • ${nextcloud.sourceFolder}` 
                : 'Click to test live WebDAV connection'}
            </p>
          </div>
        </div>

        {/* Card 2: Supabase Status */}
        <div 
          onClick={() => onNavigate('supabase')}
          className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors cursor-pointer"
          title="Click to configure Supabase credentials"
        >
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Supabase Status</span>
            <div className={`p-2 rounded-lg ${
              supabase.isConnected 
                ? 'bg-teal-50 text-teal-600' 
                : (!supabase.url || (!supabase.anonKey && !supabase.serviceKey && !supabase.serviceRoleKey))
                  ? 'bg-slate-100 text-slate-400'
                  : 'bg-rose-50 text-rose-600'
            }`}>
              <Database className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className={`text-xl font-bold ${
              supabase.isConnected 
                ? 'text-teal-700' 
                : (!supabase.url || (!supabase.anonKey && !supabase.serviceKey && !supabase.serviceRoleKey))
                  ? 'text-slate-600'
                  : 'text-rose-600'
            }`}>
              {supabase.isConnected 
                ? 'Connected' 
                : (!supabase.url || (!supabase.anonKey && !supabase.serviceKey && !supabase.serviceRoleKey))
                  ? 'Not Configured'
                  : 'Disconnected'}
            </div>
            <p className="text-xs text-slate-500 mt-1 truncate">
              {supabase.isConnected 
                ? 'PostgreSQL Live Verified' 
                : (!supabase.url || (!supabase.anonKey && !supabase.serviceKey && !supabase.serviceRoleKey))
                  ? 'Click to enter project URL & API key'
                  : 'Connection test unverified or failed'}
            </p>
          </div>
        </div>

        {/* Card 3: Files Waiting */}
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Files Waiting</span>
            <div className="p-2 rounded-lg bg-amber-50 text-amber-600">
              <Clock className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className="text-2xl font-bold text-slate-900">{filesWaiting}</div>
            <p className="text-xs text-slate-500 mt-1">
              {filesWaiting === 0 ? 'All Excel files up to date' : 'Requires worker sync cycle'}
            </p>
          </div>
        </div>

        {/* Card 4: Last Import */}
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Last Import</span>
            <div className="p-2 rounded-lg bg-blue-50 text-blue-600">
              <Activity className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className="text-lg font-bold text-slate-900 truncate" title={lastImport?.filename || 'None'}>
              {lastImport ? lastImport.filename : 'No imports yet'}
            </div>
            <p className="text-xs text-slate-500 mt-1">
              {lastImport ? `${lastImport.status} • ${lastImport.rowsProcessed} rows` : 'Idle'}
            </p>
          </div>
        </div>

        {/* Card 5: Successful Imports */}
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Successful Imports</span>
            <div className="p-2 rounded-lg bg-emerald-50 text-emerald-600">
              <CheckCircle2 className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className="text-2xl font-bold text-emerald-700">{successfulImports}</div>
            <p className="text-xs text-slate-500 mt-1">100% idempotent upserts</p>
          </div>
        </div>

        {/* Card 6: Failed Imports */}
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Failed Imports</span>
            <div className="p-2 rounded-lg bg-rose-50 text-rose-600">
              <XCircle className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className="text-2xl font-bold text-rose-700">{failedImports}</div>
            <p className="text-xs text-slate-500 mt-1">Isolated in error audit log</p>
          </div>
        </div>

        {/* Card 7: Records Imported */}
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Records Inserted</span>
            <div className="p-2 rounded-lg bg-indigo-50 text-indigo-600">
              <Layers className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className="text-2xl font-bold text-slate-900">{totalRecordsImported.toLocaleString()}</div>
            <p className="text-xs text-slate-500 mt-1">New rows created in database</p>
          </div>
        </div>

        {/* Card 8: Records Updated */}
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs hover:border-slate-300 transition-colors">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Records Updated</span>
            <div className="p-2 rounded-lg bg-purple-50 text-purple-600">
              <Zap className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-3">
            <div className="text-2xl font-bold text-slate-900">{totalRecordsUpdated.toLocaleString()}</div>
            <p className="text-xs text-slate-500 mt-1">Modified via unique key matches</p>
          </div>
        </div>
      </div>

      {/* Synchronizer Architecture Pipeline Flow */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
        <h2 className="text-base font-semibold text-slate-900 mb-4 flex items-center space-x-2">
          <ShieldCheck className="w-5 h-5 text-emerald-600" />
          <span>Active Architecture Topology & Pipeline</span>
        </h2>
        
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 relative">
          {/* Node 1: Nextcloud */}
          <div className="p-4 rounded-lg bg-slate-50 border border-slate-200 relative">
            <div className="text-xs font-bold text-slate-500 uppercase tracking-wider">Source</div>
            <div className="font-semibold text-slate-900 mt-1">Nextcloud TrueNAS</div>
            <p className="text-xs text-slate-600 mt-1">WebDAV HTTPS directory watcher</p>
            <div className="mt-3 inline-flex items-center text-xs text-emerald-700 bg-emerald-100/60 px-2 py-0.5 rounded font-mono">
              /ExcelImports
            </div>
          </div>

          {/* Node 2: Production Engine / Scheduler */}
          <div className="p-4 rounded-lg bg-indigo-50/50 border border-indigo-200 relative">
            <div className="text-xs font-bold text-indigo-600 uppercase tracking-wider">Engine</div>
            <div className="font-semibold text-slate-900 mt-1">
              {schedulerStatus?.engineMode === 'EXTERNAL_WORKER_DAEMON' ? 'Worker Daemon' : 'Sync Pipeline Engine'}
            </div>
            <p className="text-xs text-slate-600 mt-1">
              {schedulerStatus?.engineMode === 'EXTERNAL_WORKER_DAEMON'
                ? `FastAPI remote daemon at ${schedulerStatus.workerEndpoint?.replace(/^https?:\/\//, '')}`
                : 'Node.js production pipeline & XLSX parsing'}
            </p>
            <div className="mt-3 inline-flex items-center text-xs text-indigo-700 bg-indigo-100/60 px-2 py-0.5 rounded font-mono">
              {schedulerStatus?.enabled
                ? `Auto: ${schedulerStatus.intervalLabel}${schedulerStatus.secondsUntilNextRun !== null ? ` • in ${Math.ceil(schedulerStatus.secondsUntilNextRun / 60)}m` : ''}`
                : 'Manual Trigger'}
            </div>
          </div>

          {/* Node 3: Supabase */}
          <div className="p-4 rounded-lg bg-teal-50/50 border border-teal-200 relative">
            <div className="text-xs font-bold text-teal-600 uppercase tracking-wider">Destination</div>
            <div className="font-semibold text-slate-900 mt-1">Supabase PostgreSQL</div>
            <p className="text-xs text-slate-600 mt-1">Idempotent upsert & audit logs</p>
            <div className="mt-3 inline-flex items-center text-xs text-teal-700 bg-teal-100/60 px-2 py-0.5 rounded font-mono">
              5 target tables
            </div>
          </div>

          {/* Node 4: Vercel Admin */}
          <div className="p-4 rounded-lg bg-purple-50/50 border border-purple-200 relative">
            <div className="text-xs font-bold text-purple-600 uppercase tracking-wider">Control</div>
            <div className="font-semibold text-slate-900 mt-1">Vercel Dashboard</div>
            <p className="text-xs text-slate-600 mt-1">Mapping config & dry run engine</p>
            <div className="mt-3 inline-flex items-center text-xs text-purple-700 bg-purple-100/60 px-2 py-0.5 rounded font-mono">
              Web Admin UI
            </div>
          </div>
        </div>
      </div>

      {/* Quick Access Action Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Nextcloud Detected Files Box */}
        <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-semibold text-slate-900 flex items-center space-x-2">
              <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
              <span>Monitored Excel Workbooks</span>
            </h3>
            <button
              onClick={() => onNavigate('files')}
              className="text-xs font-semibold text-emerald-700 hover:text-emerald-800 flex items-center space-x-1"
            >
              <span>View All Files</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="divide-y divide-slate-100">
            {files.slice(0, 3).map((f) => (
              <div key={f.id} className="py-3 flex items-center justify-between">
                <div>
                  <div className="text-sm font-medium text-slate-900">{f.filename}</div>
                  <div className="text-xs text-slate-500">
                    {f.fileSizeFormatted} • {f.worksheetsCount || 1} worksheets
                  </div>
                </div>
                <div className="flex items-center space-x-2">
                  <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${
                    f.status === 'Synced' ? 'bg-emerald-100 text-emerald-800' :
                    f.status === 'Modified' ? 'bg-amber-100 text-amber-800' :
                    'bg-slate-100 text-slate-700'
                  }`}>
                    {f.status}
                  </span>
                  <button
                    onClick={() => onNavigate('analyzer')}
                    className="p-1 rounded text-slate-400 hover:text-slate-600"
                    title="Analyze workbook"
                  >
                    <ArrowUpRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Recent Import Activity Box */}
        <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-semibold text-slate-900 flex items-center space-x-2">
              <Activity className="w-4 h-4 text-blue-600" />
              <span>Recent Import Logs</span>
            </h3>
            <button
              onClick={() => onNavigate('history')}
              className="text-xs font-semibold text-blue-700 hover:text-blue-800 flex items-center space-x-1"
            >
              <span>Full Audit History</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="divide-y divide-slate-100">
            {importLogs.slice(0, 3).map((log) => (
              <div key={log.id} className="py-3 flex items-center justify-between">
                <div>
                  <div className="text-sm font-medium text-slate-900">{log.filename}</div>
                  <div className="text-xs text-slate-500">
                    +{log.rowsInserted} inserted • {log.rowsUpdated} updated • {log.rowsFailed} failed
                  </div>
                </div>
                <div className="flex items-center space-x-2">
                  <span className={`text-xs px-2.5 py-0.5 rounded-full font-semibold ${
                    log.status === 'Success' ? 'bg-emerald-100 text-emerald-800' :
                    log.status === 'Partial Success' ? 'bg-amber-100 text-amber-800' :
                    log.status === 'Skipped' ? 'bg-slate-100 text-slate-700' :
                    'bg-rose-100 text-rose-800'
                  }`}>
                    {log.status}
                  </span>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
