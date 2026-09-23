import React, { useState, useEffect } from 'react';
import { 
  X, 
  Key, 
  ShieldCheck, 
  Eye, 
  EyeOff, 
  Check, 
  Edit3, 
  RefreshCw, 
  AlertTriangle, 
  CheckCircle2, 
  Cloud, 
  Database, 
  Cpu, 
  Sparkles,
  ExternalLink,
  Copy
} from 'lucide-react';
import { NextcloudConfig, SupabaseConfig, SyncSettings } from '../types';
import { ApiClient } from '../services/apiClient';

interface SecretsVaultModalProps {
  isOpen: boolean;
  onClose: () => void;
  nextcloud: NextcloudConfig;
  supabase: SupabaseConfig;
  syncSettings: SyncSettings;
  onSaveNextcloud: (cfg: NextcloudConfig) => void;
  onSaveSupabase: (cfg: SupabaseConfig) => void;
  onSaveSyncSettings: (s: SyncSettings) => void;
}

interface SecretField {
  id: string;
  category: 'Nextcloud WebDAV' | 'Supabase PostgreSQL' | 'Worker Array' | 'AI Assistant';
  key: string;
  label: string;
  description: string;
  value: string;
  isSecret: boolean;
  isConfigured: boolean;
  isConnected?: boolean;
  statusMessage?: string;
  latencyMs?: number;
  placeholder: string;
}

export const SecretsVaultModal: React.FC<SecretsVaultModalProps> = ({
  isOpen,
  onClose,
  nextcloud,
  supabase,
  syncSettings,
  onSaveNextcloud,
  onSaveSupabase,
  onSaveSyncSettings,
}) => {
  // Local editable draft state
  const [formData, setFormData] = useState({
    nextcloudUrl: nextcloud.url || '',
    nextcloudWebdavUrl: nextcloud.webdavUrl || '',
    nextcloudUsername: nextcloud.username || '',
    nextcloudAppPassword: nextcloud.appPassword || '',
    nextcloudFolder: nextcloud.sourceFolder || '/ExcelImports',
    supabaseUrl: supabase.url || '',
    supabaseServiceKey: supabase.serviceKey || supabase.serviceRoleKey || '',
    supabaseAnonKey: supabase.anonKey || '',
    workerUrl: syncSettings.workerUrl || '',
    syncInterval: syncSettings.syncInterval || '15m',
    geminiApiKey: typeof window !== 'undefined' ? (localStorage.getItem('gemini_api_key') || '') : '',
  });

  const [revealedKeys, setRevealedKeys] = useState<Record<string, boolean>>({});
  const [revealAll, setRevealAll] = useState(false);
  const [editingKey, setEditingKey] = useState<string | null>(null);
  const [isTesting, setIsTesting] = useState(false);
  const [testReport, setTestReport] = useState<any>(null);
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  // Sync draft when props change
  useEffect(() => {
    setFormData({
      nextcloudUrl: nextcloud.url || '',
      nextcloudWebdavUrl: nextcloud.webdavUrl || '',
      nextcloudUsername: nextcloud.username || '',
      nextcloudAppPassword: nextcloud.appPassword || '',
      nextcloudFolder: nextcloud.sourceFolder || '/ExcelImports',
      supabaseUrl: supabase.url || '',
      supabaseServiceKey: supabase.serviceKey || supabase.serviceRoleKey || '',
      supabaseAnonKey: supabase.anonKey || '',
      workerUrl: syncSettings.workerUrl || '',
      syncInterval: syncSettings.syncInterval || '15m',
      geminiApiKey: typeof window !== 'undefined' ? (localStorage.getItem('gemini_api_key') || '') : '',
    });
  }, [nextcloud, supabase, syncSettings, isOpen]);

  if (!isOpen) return null;

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3500);
  };

  const toggleReveal = (key: string) => {
    setRevealedKeys(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const handleToggleRevealAll = () => {
    const nextState = !revealAll;
    setRevealAll(nextState);
    const updated: Record<string, boolean> = {};
    secretsList.forEach(s => {
      updated[s.key] = nextState;
    });
    setRevealedKeys(updated);
  };

  const handleSaveField = (key: string, val: string) => {
    const updated = { ...formData, [key]: val };
    setFormData(updated);

    // Persist to underlying store
    if (key.startsWith('nextcloud')) {
      const updatedNc: NextcloudConfig = {
        ...nextcloud,
        url: updated.nextcloudUrl,
        webdavUrl: updated.nextcloudWebdavUrl || `${updated.nextcloudUrl.replace(/\/+$/, '')}/remote.php/dav/files/${encodeURIComponent(updated.nextcloudUsername)}/`,
        username: updated.nextcloudUsername,
        appPassword: updated.nextcloudAppPassword,
        sourceFolder: updated.nextcloudFolder,
      };
      onSaveNextcloud(updatedNc);
    } else if (key.startsWith('supabase')) {
      const updatedSb: SupabaseConfig = {
        ...supabase,
        url: updated.supabaseUrl,
        serviceKey: updated.supabaseServiceKey,
        serviceRoleKey: updated.supabaseServiceKey,
        anonKey: updated.supabaseAnonKey,
      };
      onSaveSupabase(updatedSb);
    } else if (key === 'workerUrl' || key === 'syncInterval') {
      const updatedSet: SyncSettings = {
        ...syncSettings,
        workerUrl: updated.workerUrl,
        syncInterval: updated.syncInterval as any,
      };
      onSaveSyncSettings(updatedSet);
    } else if (key === 'geminiApiKey') {
      if (typeof window !== 'undefined') {
        localStorage.setItem('gemini_api_key', val);
      }
    }

    // Also forward update to worker array on port 8000
    ApiClient.updateWorkerSecrets({
      nextcloud_url: updated.nextcloudUrl,
      nextcloud_webdav_url: updated.nextcloudWebdavUrl,
      nextcloud_username: updated.nextcloudUsername,
      nextcloud_app_password: updated.nextcloudAppPassword,
      nextcloud_folder: updated.nextcloudFolder,
      supabase_url: updated.supabaseUrl,
      supabase_key: updated.supabaseServiceKey,
      sync_interval: updated.syncInterval,
      workerUrl: updated.workerUrl,
    });

    setEditingKey(null);
    showToast(`Updated and saved ${key}`);
  };

  const handleRunTestAll = async () => {
    setIsTesting(true);
    setTestReport(null);

    try {
      const res = await ApiClient.testAllSecrets({
        nextcloud: {
          url: formData.nextcloudUrl,
          webdavUrl: formData.nextcloudWebdavUrl,
          username: formData.nextcloudUsername,
          appPassword: formData.nextcloudAppPassword,
          sourceFolder: formData.nextcloudFolder,
          isConnected: false,
        },
        supabase: {
          url: formData.supabaseUrl,
          anonKey: formData.supabaseAnonKey,
          serviceKey: formData.supabaseServiceKey,
          serviceRoleKey: formData.supabaseServiceKey,
          isConnected: false,
        },
        workerUrl: formData.workerUrl,
      });

      if (res.success && res.report) {
        setTestReport(res.report);
        if (res.report.allConnected) {
          showToast('All secrets verified and connected successfully!');
        } else {
          showToast('Verification completed. Check failing secrets below.');
        }
      } else {
        setTestReport({
          allConnected: false,
          error: res.error || 'Connection verification failed',
        });
      }
    } catch (e: any) {
      setTestReport({
        allConnected: false,
        error: e.message,
      });
    } finally {
      setIsTesting(false);
    }
  };

  // Build list of secrets for tabular inspection
  const secretsList: SecretField[] = [
    {
      id: 'nc_url',
      category: 'Nextcloud WebDAV',
      key: 'NEXTCLOUD_URL',
      label: 'TrueNAS Nextcloud Host URL',
      description: 'Base HTTPS domain of your Nextcloud server (e.g. https://cloud.jhcnexus.space)',
      value: formData.nextcloudUrl,
      isSecret: false,
      isConfigured: Boolean(formData.nextcloudUrl),
      isConnected: testReport?.nextcloud?.checks?.hostReachability ?? nextcloud.isConnected,
      latencyMs: testReport?.nextcloud?.latencyMs,
      placeholder: 'https://cloud.jhcnexus.space',
    },
    {
      id: 'nc_webdav',
      category: 'Nextcloud WebDAV',
      key: 'NEXTCLOUD_WEBDAV_URL',
      label: 'Full WebDAV Endpoint',
      description: 'WebDAV remote path (e.g. /remote.php/dav/files/truenas_admin/)',
      value: formData.nextcloudWebdavUrl,
      isSecret: false,
      isConfigured: Boolean(formData.nextcloudWebdavUrl),
      isConnected: testReport?.nextcloud?.isConnected ?? nextcloud.isConnected,
      placeholder: 'https://cloud.jhcnexus.space/remote.php/dav/files/truenas_admin/',
    },
    {
      id: 'nc_user',
      category: 'Nextcloud WebDAV',
      key: 'NEXTCLOUD_USERNAME',
      label: 'WebDAV Service User',
      description: 'Dedicated WebDAV authorized username on TrueNAS SCALE',
      value: formData.nextcloudUsername,
      isSecret: false,
      isConfigured: Boolean(formData.nextcloudUsername),
      isConnected: testReport?.nextcloud?.checks?.authValid,
      placeholder: 'truenas_admin',
    },
    {
      id: 'nc_pass',
      category: 'Nextcloud WebDAV',
      key: 'NEXTCLOUD_APP_PASSWORD',
      label: 'Nextcloud App Password',
      description: 'Granular device token generated under Nextcloud Security Settings',
      value: formData.nextcloudAppPassword,
      isSecret: true,
      isConfigured: Boolean(formData.nextcloudAppPassword),
      isConnected: testReport?.nextcloud?.checks?.authValid,
      placeholder: 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT',
    },
    {
      id: 'nc_folder',
      category: 'Nextcloud WebDAV',
      key: 'NEXTCLOUD_FOLDER',
      label: 'Source Folder Path',
      description: 'Directory where Excel files are deposited for synchronization',
      value: formData.nextcloudFolder,
      isSecret: false,
      isConfigured: Boolean(formData.nextcloudFolder),
      isConnected: testReport?.nextcloud?.checks?.folderAccessible,
      placeholder: '/ExcelImports',
    },
    {
      id: 'supa_url',
      category: 'Supabase PostgreSQL',
      key: 'SUPABASE_URL',
      label: 'Supabase Project URL',
      description: 'PostgreSQL PostgREST API base URL (e.g. https://db.jhcnexus.space)',
      value: formData.supabaseUrl,
      isSecret: false,
      isConfigured: Boolean(formData.supabaseUrl),
      isConnected: testReport?.supabase?.checks?.hostReachability ?? supabase.isConnected,
      latencyMs: testReport?.supabase?.latencyMs,
      placeholder: 'https://db.jhcnexus.space',
    },
    {
      id: 'supa_service_key',
      category: 'Supabase PostgreSQL',
      key: 'SUPABASE_SERVICE_ROLE_KEY',
      label: 'Supabase Service Role Key',
      description: 'Privileged JWT token used to bypass Row Level Security (RLS) for worker upserts',
      value: formData.supabaseServiceKey,
      isSecret: true,
      isConfigured: Boolean(formData.supabaseServiceKey),
      isConnected: testReport?.supabase?.isConnected ?? supabase.isConnected,
      placeholder: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...',
    },
    {
      id: 'supa_anon_key',
      category: 'Supabase PostgreSQL',
      key: 'SUPABASE_ANON_KEY',
      label: 'Supabase Anon Public Key',
      description: 'Client-side anonymous key for public table discovery',
      value: formData.supabaseAnonKey,
      isSecret: true,
      isConfigured: Boolean(formData.supabaseAnonKey),
      placeholder: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...',
    },
    {
      id: 'worker_url',
      category: 'Worker Array',
      key: 'WORKER_SERVICE_URL',
      label: 'External Worker Daemon Endpoint (Optional)',
      description: 'Optional remote FastAPI worker daemon. Leave blank to run the Integrated Production In-Process Engine.',
      value: formData.workerUrl,
      isSecret: false,
      isConfigured: true,
      isConnected: testReport?.worker?.isConnected ?? true,
      latencyMs: testReport?.worker?.latencyMs,
      placeholder: 'None (Using Integrated Production Engine)',
    },
    {
      id: 'sync_interval',
      category: 'Worker Array',
      key: 'SYNC_INTERVAL',
      label: 'Automated Polling Interval',
      description: 'Frequency of automated scheduled background WebDAV sync cycles (5m, 15m, 30m, 1h, daily)',
      value: formData.syncInterval,
      isSecret: false,
      isConfigured: Boolean(formData.syncInterval),
      placeholder: '15m',
    },
    {
      id: 'gemini_key',
      category: 'AI Assistant',
      key: 'GEMINI_API_KEY',
      label: 'Gemini AI Assistant Key',
      description: 'Google GenAI token for automated Excel column matching & schema detection',
      value: formData.geminiApiKey,
      isSecret: true,
      isConfigured: Boolean(formData.geminiApiKey),
      placeholder: 'AIzaSy...',
    },
  ];

  const totalConfigured = secretsList.filter(s => s.isConfigured).length;
  const missingSecrets = secretsList.filter(s => !s.isConfigured);
  const allConnected = testReport ? testReport.allConnected : (nextcloud.isConnected && supabase.isConnected);

  return (
    <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-3 sm:p-6 animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-4xl w-full flex flex-col max-h-[90vh] overflow-hidden">
        
        {/* Header */}
        <div className="px-6 py-5 border-b border-slate-200 flex items-center justify-between bg-slate-50/80">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 flex items-center justify-center text-white shadow-xs">
              <Key className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h2 className="text-lg font-bold text-slate-900 tracking-tight">
                  Unified Secrets & Environment Vault
                </h2>
                <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold ${
                  allConnected ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                }`}>
                  <span className={`w-1.5 h-1.5 rounded-full mr-1.5 ${allConnected ? 'bg-emerald-600' : 'bg-amber-600'}`} />
                  {allConnected ? 'All Systems Connected' : `${missingSecrets.length} Checks Needed`}
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Inspect, reveal, edit, and test live connectivity for TrueNAS WebDAV, Coolify Worker Array (Port 8000), and Supabase.
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-2">
            <button
              onClick={onClose}
              className="p-2 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-200/60 transition-colors"
              title="Close modal"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Global Action Toolbar */}
        <div className="px-6 py-3.5 bg-slate-100/70 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center space-x-2">
            <span className="font-semibold text-slate-700">Configured:</span>
            <span className="font-mono bg-white px-2 py-0.5 rounded border border-slate-200 text-slate-800 font-bold">
              {totalConfigured} / {secretsList.length}
            </span>
            <span className="text-slate-400">|</span>
            <span className="text-slate-600">
              Worker Array: <code className="font-mono bg-white px-1.5 py-0.5 rounded border border-slate-200">Port 8000</code>
            </span>
          </div>

          <div className="flex items-center space-x-2.5">
            <button
              onClick={handleToggleRevealAll}
              className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 font-medium transition-colors shadow-2xs"
            >
              {revealAll ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
              <span>{revealAll ? 'Mask All Secrets' : 'Reveal All Secrets'}</span>
            </button>

            <button
              onClick={handleRunTestAll}
              disabled={isTesting}
              className="inline-flex items-center space-x-1.5 px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold transition-colors shadow-2xs"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isTesting ? 'animate-spin' : ''}`} />
              <span>{isTesting ? 'Testing Connectivity...' : 'Test All Connections'}</span>
            </button>
          </div>
        </div>

        {/* Alert Banner if missing */}
        {missingSecrets.length > 0 && (
          <div className="mx-6 mt-4 p-3.5 rounded-xl bg-amber-50 border border-amber-200 text-xs text-amber-800 flex items-start space-x-2.5">
            <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
            <div>
              <span className="font-bold">Missing Credentials Detected: </span>
              {missingSecrets.map(m => m.label).join(', ')} are not yet configured. Click <strong>Edit</strong> beside any secret to configure it.
            </div>
          </div>
        )}

        {/* Live Test Results Alert */}
        {testReport && (
          <div className={`mx-6 mt-4 p-4 rounded-xl border text-xs space-y-1.5 ${
            testReport.allConnected ? 'bg-emerald-50 text-emerald-900 border-emerald-200' : 'bg-slate-50 text-slate-800 border-slate-200'
          }`}>
            <div className="font-bold flex items-center space-x-2">
              <ShieldCheck className={`w-4 h-4 ${testReport.allConnected ? 'text-emerald-600' : 'text-amber-600'}`} />
              <span>Live Diagnostic Connectivity Report ({new Date(testReport.timestamp).toLocaleTimeString()})</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-1 font-mono text-[11px]">
              <div className="p-2 rounded bg-white border border-slate-200">
                <strong>Nextcloud WebDAV:</strong> {testReport.nextcloud?.statusText || 'Unverified'} ({testReport.nextcloud?.latencyMs || 0}ms)
              </div>
              <div className="p-2 rounded bg-white border border-slate-200">
                <strong>Supabase Database:</strong> {testReport.supabase?.statusText || 'Unverified'} ({testReport.supabase?.latencyMs || 0}ms)
              </div>
              <div className="p-2 rounded bg-white border border-slate-200">
                <strong>Worker Array:</strong> {testReport.worker?.statusText || 'Active'} ({testReport.worker?.latencyMs || 0}ms)
              </div>
            </div>
          </div>
        )}

        {/* Secrets Table Viewport */}
        <div className="p-6 overflow-y-auto flex-1 space-y-6">
          {['Nextcloud WebDAV', 'Supabase PostgreSQL', 'Worker Array', 'AI Assistant'].map((category) => {
            const items = secretsList.filter(s => s.category === category);
            return (
              <div key={category} className="space-y-3">
                <div className="flex items-center space-x-2 text-xs font-bold text-slate-500 uppercase tracking-wider">
                  {category === 'Nextcloud WebDAV' && <Cloud className="w-3.5 h-3.5 text-blue-600" />}
                  {category === 'Supabase PostgreSQL' && <Database className="w-3.5 h-3.5 text-teal-600" />}
                  {category === 'Worker Array' && <Cpu className="w-3.5 h-3.5 text-purple-600" />}
                  {category === 'AI Assistant' && <Sparkles className="w-3.5 h-3.5 text-amber-600" />}
                  <span>{category}</span>
                </div>

                <div className="border border-slate-200 rounded-xl divide-y divide-slate-100 bg-white shadow-2xs overflow-hidden">
                  {items.map((sec) => {
                    const isRevealed = revealAll || revealedKeys[sec.key];
                    const isEditing = editingKey === sec.id;

                    return (
                      <div key={sec.id} className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 hover:bg-slate-50/50 transition-colors">
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center space-x-2">
                            <span className="font-mono text-xs font-bold text-slate-800">{sec.key}</span>
                            <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase ${
                              sec.isConfigured ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                            }`}>
                              {sec.isConfigured ? 'Configured' : 'Missing'}
                            </span>
                            {sec.latencyMs !== undefined && (
                              <span className="text-[10px] text-slate-400 font-mono">
                                {sec.latencyMs}ms
                              </span>
                            )}
                          </div>
                          <p className="text-xs text-slate-500 mt-0.5">{sec.description}</p>
                          
                          {/* Value Display / Editor */}
                          <div className="mt-2">
                            {isEditing ? (
                              <div className="flex items-center space-x-2 max-w-lg">
                                <input
                                  type={sec.isSecret && !isRevealed ? 'password' : 'text'}
                                  defaultValue={sec.value}
                                  id={`input-${sec.id}`}
                                  placeholder={sec.placeholder}
                                  className="w-full text-xs font-mono px-3 py-1.5 rounded-lg border border-emerald-500 focus:outline-none focus:ring-1 focus:ring-emerald-500 bg-white"
                                  autoFocus
                                />
                                <button
                                  onClick={() => {
                                    const input = document.getElementById(`input-${sec.id}`) as HTMLInputElement;
                                    if (input) {
                                      // map id to draft field name
                                      const keyMap: Record<string, string> = {
                                        nc_url: 'nextcloudUrl',
                                        nc_webdav: 'nextcloudWebdavUrl',
                                        nc_user: 'nextcloudUsername',
                                        nc_pass: 'nextcloudAppPassword',
                                        nc_folder: 'nextcloudFolder',
                                        supa_url: 'supabaseUrl',
                                        supa_service_key: 'supabaseServiceKey',
                                        supa_anon_key: 'supabaseAnonKey',
                                        worker_url: 'workerUrl',
                                        sync_interval: 'syncInterval',
                                        gemini_key: 'geminiApiKey',
                                      };
                                      handleSaveField(keyMap[sec.id] || sec.id, input.value);
                                    }
                                  }}
                                  className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-medium text-xs shadow-2xs"
                                >
                                  Save
                                </button>
                                <button
                                  onClick={() => setEditingKey(null)}
                                  className="px-2.5 py-1.5 rounded-lg border border-slate-300 hover:bg-slate-100 text-slate-600 text-xs"
                                >
                                  Cancel
                                </button>
                              </div>
                            ) : (
                              <div className="flex items-center space-x-2">
                                <span className="font-mono text-xs bg-slate-100 px-2.5 py-1 rounded-md text-slate-700 truncate max-w-md select-all">
                                  {sec.value ? (
                                    isRevealed || !sec.isSecret ? sec.value : '••••••••••••••••••••••••'
                                  ) : (
                                    <span className="text-slate-400 italic">Not configured</span>
                                  )}
                                </span>
                              </div>
                            )}
                          </div>
                        </div>

                        {/* Action buttons */}
                        {!isEditing && (
                          <div className="flex items-center space-x-1.5 shrink-0 self-start sm:self-center">
                            {sec.isSecret && (
                              <button
                                onClick={() => toggleReveal(sec.key)}
                                className="p-1.5 rounded-lg text-slate-500 hover:text-slate-700 hover:bg-slate-100 transition-colors"
                                title={isRevealed ? 'Mask secret' : 'Reveal secret'}
                              >
                                {isRevealed ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                              </button>
                            )}

                            <button
                              onClick={() => setEditingKey(sec.id)}
                              className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-lg border border-slate-300 hover:bg-slate-50 text-slate-700 font-medium text-xs transition-colors shadow-2xs"
                            >
                              <Edit3 className="w-3.5 h-3.5 text-slate-500" />
                              <span>Edit</span>
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
          <div className="flex items-center space-x-2">
            <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block" />
            <span>Changes persist automatically to runtime configuration & worker state.</span>
          </div>

          <button
            onClick={onClose}
            className="px-4 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white font-semibold transition-colors shadow-2xs"
          >
            Close Vault
          </button>
        </div>
      </div>

      {/* Floating Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-60 bg-slate-900 text-white px-4 py-2.5 rounded-xl shadow-lg border border-slate-700 text-xs font-medium flex items-center space-x-2 animate-in slide-in-from-bottom-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMessage}</span>
        </div>
      )}
    </div>
  );
};
