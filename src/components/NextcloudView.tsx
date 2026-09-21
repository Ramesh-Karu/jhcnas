import React, { useState } from 'react';
import { 
  Cloud, 
  CheckCircle2, 
  XCircle, 
  AlertCircle, 
  RefreshCw, 
  Folder, 
  Lock, 
  ShieldCheck, 
  ExternalLink,
  Save,
  Server,
  Eye,
  EyeOff
} from 'lucide-react';
import { NextcloudConfig, NextcloudFile } from '../types';
import { ApiClient } from '../services/apiClient';

interface NextcloudViewProps {
  config: NextcloudConfig;
  files: NextcloudFile[];
  onSaveConfig: (cfg: NextcloudConfig) => void;
}

export const NextcloudView: React.FC<NextcloudViewProps> = ({
  config,
  files,
  onSaveConfig
}) => {
  const [formData, setFormData] = useState<NextcloudConfig>(config);
  const [isTesting, setIsTesting] = useState(false);
  const [testResult, setTestResult] = useState<{
    tested: boolean;
    success: boolean;
    urlReachable: boolean;
    authSuccess: boolean;
    folderAccessible: boolean;
    filesListed: number;
    message: string;
  } | null>(null);
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const [showPassword, setShowPassword] = useState(false);

  const handleTestConnection = async () => {
    setIsTesting(true);
    setTestResult(null);

    // Call real backend WebDAV proxy
    const realResult = await ApiClient.testNextcloudConnection(formData);
    const filesResult = await ApiClient.listNextcloudFiles(formData);

    const isHealthy = realResult.success && realResult.checks.hostReachability && realResult.checks.authValid;
    const realFilesCount = filesResult.files ? filesResult.files.length : files.length;

    setTestResult({
      tested: true,
      success: isHealthy,
      urlReachable: realResult.checks.hostReachability,
      authSuccess: realResult.checks.authValid,
      folderAccessible: realResult.checks.folderExists,
      filesListed: realFilesCount,
      message: realResult.message || realResult.error || (isHealthy ? 'WebDAV connection verified!' : 'Connection test failed.')
    });

    if (isHealthy) {
      const updated = { 
        ...formData, 
        isConnected: true, 
        lastChecked: new Date().toISOString(), 
        statusMessage: `Connected to Nextcloud (${realResult.checks.details?.versionstring || 'v34'}) at ${new URL(formData.url).hostname}` 
      };
      setFormData(updated);
      onSaveConfig(updated);
    }

    setIsTesting(false);
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    onSaveConfig(formData);
    setSaveMessage('Nextcloud configuration updated successfully.');
    setTimeout(() => setSaveMessage(null), 3000);
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-emerald-50 text-emerald-600">
            <Cloud className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Nextcloud WebDAV Integration</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Connect to your Nextcloud instance running on TrueNAS SCALE via dedicated WebDAV App Password.
            </p>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Form Configuration */}
        <div className="lg:col-span-2 bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
          <h2 className="text-base font-semibold text-slate-900 mb-4 flex items-center space-x-2">
            <Server className="w-4 h-4 text-emerald-600" />
            <span>Connection Credentials</span>
          </h2>

          <form onSubmit={handleSave} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                Nextcloud Host URL
              </label>
              <input
                type="text"
                value={formData.url}
                onChange={(e) => setFormData({ ...formData, url: e.target.value })}
                placeholder="https://cloud.example.com"
                className="w-full px-3.5 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 font-mono"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                WebDAV Endpoint URL
              </label>
              <input
                type="text"
                value={formData.webdavUrl}
                onChange={(e) => setFormData({ ...formData, webdavUrl: e.target.value })}
                placeholder="https://cloud.example.com/remote.php/dav/files/excel-sync/"
                className="w-full px-3.5 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 font-mono"
              />
              <p className="text-[11px] text-slate-500 mt-1">
                Typically follows format: <code className="bg-slate-100 px-1 py-0.5 rounded">/remote.php/dav/files/USERNAME/</code>
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                  Username
                </label>
                <input
                  type="text"
                  value={formData.username}
                  onChange={(e) => {
                    const newUsername = e.target.value;
                    let newWebdav = formData.webdavUrl;
                    if (newWebdav.includes('/remote.php/dav/files/')) {
                      const base = newWebdav.substring(0, newWebdav.indexOf('/remote.php/dav/files/') + '/remote.php/dav/files/'.length);
                      newWebdav = `${base}${newUsername}/`;
                    }
                    setFormData({ ...formData, username: newUsername, webdavUrl: newWebdav });
                  }}
                  placeholder="truenas_admin"
                  className="w-full px-3.5 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                  App Password
                </label>
                <div className="relative">
                  <input
                    type={showPassword ? 'text' : 'password'}
                    value={formData.appPassword}
                    onChange={(e) => setFormData({ ...formData, appPassword: e.target.value })}
                    placeholder="••••••••••••••••••••••••"
                    className="w-full pl-3.5 pr-10 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 font-mono"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute right-2.5 top-2.5 text-slate-400 hover:text-slate-600 focus:outline-none"
                    title={showPassword ? 'Hide password' : 'Show password'}
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
              </div>
            </div>

            <div>
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                Target Source Folder
              </label>
              <div className="relative">
                <input
                  type="text"
                  value={formData.sourceFolder}
                  onChange={(e) => setFormData({ ...formData, sourceFolder: e.target.value })}
                  placeholder="/ExcelImports"
                  className="w-full pl-9 pr-3.5 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 font-mono"
                />
                <Folder className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
              </div>
              <p className="text-[11px] text-slate-500 mt-1">
                The worker strictly confines read operations to this folder. Original Excel files are never modified.
              </p>
            </div>

            {saveMessage && (
              <div className="p-3 rounded-lg bg-emerald-50 text-emerald-700 text-xs font-medium border border-emerald-200">
                {saveMessage}
              </div>
            )}

            <div className="flex items-center justify-between pt-2">
              <button
                type="button"
                id="btn-test-nextcloud"
                onClick={handleTestConnection}
                disabled={isTesting}
                className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 hover:bg-slate-50 shadow-2xs transition-colors"
              >
                <RefreshCw className={`w-4 h-4 text-slate-500 ${isTesting ? 'animate-spin' : ''}`} />
                <span>{isTesting ? 'Testing Connection...' : 'Test Nextcloud Connection'}</span>
              </button>

              <button
                type="submit"
                id="btn-save-nextcloud"
                className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-sm font-medium text-white shadow-xs transition-colors"
              >
                <Save className="w-4 h-4" />
                <span>Save Settings</span>
              </button>
            </div>
          </form>
        </div>

        {/* Right Column: Connection Diagnostics & Security Info */}
        <div className="space-y-6">
          {/* Test Connection Results Card */}
          <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
            <h3 className="text-sm font-semibold text-slate-900 mb-3 flex items-center space-x-2">
              <ShieldCheck className="w-4 h-4 text-emerald-600" />
              <span>Connection Verification Matrix</span>
            </h3>

            {testResult ? (
              <div className="space-y-3">
                <div className={`p-3 rounded-lg text-xs font-medium ${
                  testResult.success ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-rose-50 text-rose-800 border border-rose-200'
                }`}>
                  {testResult.message}
                </div>

                <div className="space-y-2 text-xs">
                  <div className="flex items-center justify-between py-1 border-b border-slate-100">
                    <span className="text-slate-600">1. URL Reachable:</span>
                    <span className="flex items-center space-x-1 font-medium">
                      {testResult.urlReachable ? (
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                      ) : (
                        <XCircle className="w-3.5 h-3.5 text-rose-600" />
                      )}
                      <span>{testResult.urlReachable ? 'HTTP 200/207 OK' : 'Unreachable'}</span>
                    </span>
                  </div>

                  <div className="flex items-center justify-between py-1 border-b border-slate-100">
                    <span className="text-slate-600">2. Authentication:</span>
                    <span className="flex items-center space-x-1 font-medium">
                      {testResult.authSuccess ? (
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                      ) : (
                        <XCircle className="w-3.5 h-3.5 text-rose-600" />
                      )}
                      <span>{testResult.authSuccess ? 'App Password Verified' : 'Auth Failed'}</span>
                    </span>
                  </div>

                  <div className="flex items-center justify-between py-1 border-b border-slate-100">
                    <span className="text-slate-600">3. Folder Accessible:</span>
                    <span className="flex items-center space-x-1 font-medium">
                      {testResult.folderAccessible ? (
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                      ) : (
                        <XCircle className="w-3.5 h-3.5 text-rose-600" />
                      )}
                      <span>{testResult.folderAccessible ? 'PROPFIND Successful' : 'Denied'}</span>
                    </span>
                  </div>

                  <div className="flex items-center justify-between py-1">
                    <span className="text-slate-600">4. Files Found:</span>
                    <span className="font-semibold text-slate-900">{testResult.filesListed} workbooks</span>
                  </div>
                </div>
              </div>
            ) : (
              <div className="text-center py-6 text-slate-500 text-xs">
                Click <span className="font-semibold text-slate-700">Test Nextcloud Connection</span> to verify TrueNAS WebDAV connectivity.
              </div>
            )}
          </div>

          {/* Security Best Practice Notice */}
          <div className="bg-slate-50 rounded-xl p-5 border border-slate-200 text-xs text-slate-600 space-y-2">
            <div className="flex items-center space-x-2 font-semibold text-slate-900">
              <Lock className="w-4 h-4 text-slate-700" />
              <span>TrueNAS Security Architecture</span>
            </div>
            <p className="leading-relaxed">
              Nextcloud runs containerized on TrueNAS SCALE. The Coolify background worker communicates directly via HTTPS WebDAV using a dedicated App Password.
            </p>
            <p className="text-[11px] text-slate-500">
              The browser dashboard never directly connects to Nextcloud with raw credentials; all sync tasks are dispatched to the server-side worker.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};
