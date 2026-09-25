import React, { useState, useEffect } from 'react';
import { 
  Cloud, 
  RefreshCw, 
  CheckCircle2, 
  AlertTriangle, 
  Folder, 
  FileSpreadsheet, 
  Clock, 
  Settings, 
  Play, 
  Pause, 
  X, 
  Check, 
  DownloadCloud, 
  Layers, 
  Server, 
  History, 
  Sliders, 
  CheckSquare, 
  Square,
  HardDrive,
  ExternalLink,
  ShieldCheck,
  ChevronRight,
  Info
} from 'lucide-react';
import { NextcloudGradeSyncConfig, NextcloudConfig } from '../types';
import { ApiClient } from '../services/apiClient';
import { StorageService } from '../services/storage';

interface NextcloudGradeSyncModalProps {
  isOpen: boolean;
  onClose: () => void;
  nextcloudConfig: NextcloudConfig;
  onSyncComplete: (result: {
    files: any[];
    records: any[];
    referenceColumns: string[];
    summary: any;
    message: string;
  }) => void;
}

export const NextcloudGradeSyncModal: React.FC<NextcloudGradeSyncModalProps> = ({
  isOpen,
  onClose,
  nextcloudConfig: initialNcConfig,
  onSyncComplete
}) => {
  const [activeTab, setActiveTab] = useState<'pull' | 'interval' | 'connection' | 'history'>('pull');

  // Nextcloud sync configuration state
  const [syncConfig, setSyncConfig] = useState<NextcloudGradeSyncConfig>(() => 
    StorageService.getNextcloudGradeSyncConfig()
  );

  // Connection fields
  const [url, setUrl] = useState(syncConfig.url || initialNcConfig.url || 'https://cloud.jhcnexus.space');
  const [username, setUsername] = useState(syncConfig.username || initialNcConfig.username || 'truenas_admin');
  const [appPassword, setAppPassword] = useState(syncConfig.appPassword || initialNcConfig.appPassword || 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT');
  const [sourceFolder, setSourceFolder] = useState(syncConfig.sourceFolder || initialNcConfig.sourceFolder || '/ExcelImports');

  // Browsed files state
  const [browsedFiles, setBrowsedFiles] = useState<{
    filename: string;
    path: string;
    fileSize: number;
    fileSizeFormatted: string;
    lastModified: string;
    etag: string;
    isDirectory: boolean;
    isSpreadsheet: boolean;
  }[]>([]);
  const [selectedFileNames, setSelectedFileNames] = useState<string[]>(syncConfig.selectedFiles || []);
  const [isLoadingFiles, setIsLoadingFiles] = useState(false);
  const [browseError, setBrowseError] = useState<string | null>(null);

  // Sync execution state
  const [isPulling, setIsPulling] = useState(false);
  const [pullProgressMessage, setPullProgressMessage] = useState('');
  const [pullError, setPullError] = useState<string | null>(null);

  // Interval sync state
  const [autoSyncEnabled, setAutoSyncEnabled] = useState(syncConfig.autoSyncEnabled);
  const [syncIntervalMinutes, setSyncIntervalMinutes] = useState(syncConfig.syncIntervalMinutes || 15);
  const [customMinutesInput, setCustomMinutesInput] = useState(String(syncConfig.syncIntervalMinutes || 15));

  // Connection test state
  const [isTestingConnection, setIsTestingConnection] = useState(false);
  const [testResult, setTestResult] = useState<{
    success: boolean;
    message: string;
    checks?: {
      hostReachability: boolean;
      webdavHandshake: boolean;
      authValid: boolean;
      folderExists: boolean;
    };
  } | null>(null);

  // Fetch server sync config on open
  useEffect(() => {
    if (!isOpen) return;

    let isMounted = true;
    async function loadServerConfig() {
      try {
        const res = await ApiClient.getNextcloudGradeSyncConfig();
        if (isMounted && res.success && res.config) {
          setSyncConfig(res.config);
          setUrl(res.config.url || url);
          setUsername(res.config.username || username);
          setAppPassword(res.config.appPassword || appPassword);
          setSourceFolder(res.config.sourceFolder || sourceFolder);
          setSelectedFileNames(res.config.selectedFiles || []);
          setAutoSyncEnabled(res.config.autoSyncEnabled);
          setSyncIntervalMinutes(res.config.syncIntervalMinutes || 15);
          setCustomMinutesInput(String(res.config.syncIntervalMinutes || 15));
        }
      } catch (e) {
        console.warn('Could not load server Nextcloud grade sync config:', e);
      }
    }
    loadServerConfig();

    // Auto browse files in folder
    handleBrowseFiles();

    return () => { isMounted = false; };
  }, [isOpen]);

  // Browse files in the folder
  const handleBrowseFiles = async (customFolder?: string) => {
    setIsLoadingFiles(true);
    setBrowseError(null);
    try {
      const res = await ApiClient.browseNextcloudGradeFiles({
        url,
        username,
        appPassword,
        sourceFolder: customFolder || sourceFolder
      });

      if (res.success && Array.isArray(res.files)) {
        setBrowsedFiles(res.files);
        // If no files were previously selected, select all spreadsheet files by default
        if (selectedFileNames.length === 0) {
          const spreadsheets = res.files.filter(f => f.isSpreadsheet).map(f => f.filename);
          setSelectedFileNames(spreadsheets);
        }
      } else {
        setBrowseError(res.error || 'Failed to list files from Nextcloud folder');
      }
    } catch (e: any) {
      setBrowseError(e.message || 'Error connecting to Nextcloud WebDAV');
    } finally {
      setIsLoadingFiles(false);
    }
  };

  // Test Nextcloud Connection
  const handleTestConnection = async () => {
    setIsTestingConnection(true);
    setTestResult(null);
    try {
      const webdavUrl = `${url.replace(/\/+$/, '')}/remote.php/dav/files/${encodeURIComponent(username)}/`;
      const res = await ApiClient.testNextcloudConnection({
        url,
        webdavUrl,
        username,
        appPassword,
        sourceFolder,
        isConnected: false
      });
      setTestResult({
        success: res.success,
        message: res.message || (res.success ? 'Connection verified successfully!' : res.error || 'Connection failed'),
        checks: res.checks
      });
    } catch (e: any) {
      setTestResult({
        success: false,
        message: e.message || 'Network error calling Nextcloud WebDAV'
      });
    } finally {
      setIsTestingConnection(false);
    }
  };

  // Toggle selection of a file
  const toggleFileSelection = (filename: string) => {
    setSelectedFileNames(prev => 
      prev.includes(filename) 
        ? prev.filter(f => f !== filename) 
        : [...prev, filename]
    );
  };

  // Select all spreadsheets
  const handleSelectAllSpreadsheets = () => {
    const spreadsheets = browsedFiles.filter(f => f.isSpreadsheet).map(f => f.filename);
    setSelectedFileNames(spreadsheets);
  };

  // Clear all selections
  const handleDeselectAll = () => {
    setSelectedFileNames([]);
  };

  // Save Sync Configuration & Update Server Worker
  const handleSaveConfig = async (overrideAutoSync?: boolean) => {
    const isAuto = overrideAutoSync !== undefined ? overrideAutoSync : autoSyncEnabled;
    const interval = Number(syncIntervalMinutes) || 15;

    const updated: NextcloudGradeSyncConfig = {
      ...syncConfig,
      url,
      username,
      appPassword,
      sourceFolder,
      selectedFiles: selectedFileNames,
      autoSyncEnabled: isAuto,
      syncIntervalMinutes: interval
    };

    setSyncConfig(updated);
    StorageService.saveNextcloudGradeSyncConfig(updated);

    try {
      await ApiClient.saveNextcloudGradeSyncConfig(updated);
    } catch (e) {
      console.warn('Error saving sync config to backend:', e);
    }
  };

  // Pull Selected Files from Nextcloud WebDAV
  const handleExecutePull = async () => {
    setIsPulling(true);
    setPullError(null);
    setPullProgressMessage('Connecting to Nextcloud WebDAV, pulling student class workbooks...');

    try {
      // Save active config first
      await handleSaveConfig();

      const res = await ApiClient.pullNextcloudGradeWorkbooks({
        url,
        username,
        appPassword,
        sourceFolder,
        selectedFiles: selectedFileNames.length > 0 ? selectedFileNames : undefined
      });

      if (!res.success) {
        throw new Error(res.error || res.message || 'Failed to pull grade files from Nextcloud');
      }

      setPullProgressMessage(res.message || `Successfully pulled ${res.filesPulled} workbooks (${res.recordsCount} student records)`);
      
      // Notify parent component
      if (res.files && res.records && res.referenceColumns) {
        onSyncComplete({
          files: res.files,
          records: res.records,
          referenceColumns: res.referenceColumns,
          summary: res.summary,
          message: res.message || 'Successfully pulled student workbooks from Nextcloud.'
        });
      }

      setTimeout(() => {
        setIsPulling(false);
        onClose();
      }, 1200);
    } catch (e: any) {
      console.error('Nextcloud pull error:', e);
      setPullError(e.message || 'An error occurred while pulling workbooks from Nextcloud.');
      setIsPulling(false);
    }
  };

  if (!isOpen) return null;

  const spreadsheetFiles = browsedFiles.filter(f => f.isSpreadsheet);

  return (
    <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-4xl w-full overflow-hidden flex flex-col max-h-[92vh] animate-in fade-in zoom-in-95 duration-150">
        
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-sky-500/10 text-sky-600 rounded-xl border border-sky-200">
              <Cloud className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                Nextcloud Student Workbooks & Time-Interval Sync
                <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-sky-100 text-sky-800">
                  WebDAV Live
                </span>
              </h2>
              <p className="text-xs text-slate-500">
                Pull class grade Excel sheets directly from Nextcloud and synchronize them automatically on a schedule.
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-slate-600 rounded-lg hover:bg-slate-100 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Navigation Tabs */}
        <div className="px-6 pt-3 border-b border-slate-100 flex gap-2 bg-slate-50/30">
          <button
            onClick={() => setActiveTab('pull')}
            className={`px-4 py-2.5 text-xs font-semibold rounded-t-lg transition-colors border-b-2 flex items-center gap-1.5 ${
              activeTab === 'pull'
                ? 'border-sky-600 text-sky-600 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            <DownloadCloud className="w-4 h-4" />
            Browse & Pull Workbooks
            {selectedFileNames.length > 0 && (
              <span className="ml-1 px-1.5 py-0.2 text-[10px] bg-sky-100 text-sky-700 rounded-full font-bold">
                {selectedFileNames.length}
              </span>
            )}
          </button>

          <button
            onClick={() => setActiveTab('interval')}
            className={`px-4 py-2.5 text-xs font-semibold rounded-t-lg transition-colors border-b-2 flex items-center gap-1.5 ${
              activeTab === 'interval'
                ? 'border-sky-600 text-sky-600 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            <Clock className="w-4 h-4" />
            Time-Interval Auto-Sync
            {autoSyncEnabled && (
              <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
            )}
          </button>

          <button
            onClick={() => setActiveTab('connection')}
            className={`px-4 py-2.5 text-xs font-semibold rounded-t-lg transition-colors border-b-2 flex items-center gap-1.5 ${
              activeTab === 'connection'
                ? 'border-sky-600 text-sky-600 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            <Sliders className="w-4 h-4" />
            Server & Credentials
          </button>

          <button
            onClick={() => setActiveTab('history')}
            className={`px-4 py-2.5 text-xs font-semibold rounded-t-lg transition-colors border-b-2 flex items-center gap-1.5 ${
              activeTab === 'history'
                ? 'border-sky-600 text-sky-600 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            <History className="w-4 h-4" />
            Sync Logs & History
            {syncConfig.history && syncConfig.history.length > 0 && (
              <span className="ml-1 px-1.5 py-0.2 text-[10px] bg-slate-200 text-slate-700 rounded-full font-bold">
                {syncConfig.history.length}
              </span>
            )}
          </button>
        </div>

        {/* Content Body */}
        <div className="p-6 overflow-y-auto flex-1 space-y-6">
          
          {/* TAB 1: BROWSE & PULL WORKBOOKS */}
          {activeTab === 'pull' && (
            <div className="space-y-5">
              {/* Folder Selector & Refresh */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-white rounded-lg border border-slate-200 text-sky-600">
                    <Folder className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
                      Remote Nextcloud Folder
                    </div>
                    <div className="text-sm font-bold text-slate-900 flex items-center gap-2">
                      <span>{sourceFolder}</span>
                      <span className="text-xs text-slate-400 font-normal">({url})</span>
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => handleBrowseFiles()}
                    disabled={isLoadingFiles}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-white text-slate-700 hover:bg-slate-100 border border-slate-200 transition-colors shadow-sm"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isLoadingFiles ? 'animate-spin text-sky-600' : ''}`} />
                    Refresh Nextcloud Files
                  </button>
                </div>
              </div>

              {/* Browse Error Banner */}
              {browseError && (
                <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-800 flex items-start gap-2.5">
                  <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                  <div>
                    <div className="font-semibold">Unable to list files from Nextcloud</div>
                    <div>{browseError}</div>
                    <button 
                      onClick={() => setActiveTab('connection')}
                      className="mt-1.5 text-xs text-rose-700 underline font-medium hover:text-rose-900"
                    >
                      Check Nextcloud URL and App Password settings &rarr;
                    </button>
                  </div>
                </div>
              )}

              {/* Selection Bar */}
              <div className="flex items-center justify-between text-xs text-slate-600 pt-1">
                <div className="flex items-center gap-2">
                  <span className="font-semibold text-slate-800">
                    {spreadsheetFiles.length} Spreadsheet(s) Found
                  </span>
                  <span>&bull;</span>
                  <span className="text-sky-700 font-medium">
                    {selectedFileNames.length} selected for consolidation
                  </span>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={handleSelectAllSpreadsheets}
                    className="text-xs text-sky-600 hover:text-sky-800 font-medium hover:underline"
                  >
                    Select All Spreadsheets
                  </button>
                  <span className="text-slate-300">|</span>
                  <button
                    onClick={handleDeselectAll}
                    className="text-xs text-slate-500 hover:text-slate-700 font-medium hover:underline"
                  >
                    Deselect All
                  </button>
                </div>
              </div>

              {/* Files Table / List */}
              <div className="border border-slate-200 rounded-xl overflow-hidden bg-white shadow-sm">
                {isLoadingFiles ? (
                  <div className="py-12 text-center text-slate-500 flex flex-col items-center gap-2">
                    <RefreshCw className="w-6 h-6 animate-spin text-sky-600" />
                    <span className="text-sm font-medium">Connecting to Nextcloud and scanning folder...</span>
                  </div>
                ) : spreadsheetFiles.length === 0 ? (
                  <div className="py-12 text-center text-slate-500 px-4">
                    <FileSpreadsheet className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                    <p className="text-sm font-medium text-slate-700">No Excel spreadsheets found in {sourceFolder}</p>
                    <p className="text-xs text-slate-400 mt-1 max-w-md mx-auto">
                      Place your grade files (e.g. Grade 6.xlsx, Grade 7.xlsx, Grade 8.xlsx) in the Nextcloud folder <code className="px-1 py-0.5 bg-slate-100 rounded text-slate-600">{sourceFolder}</code> and click Refresh.
                    </p>
                  </div>
                ) : (
                  <div className="max-h-72 overflow-y-auto divide-y divide-slate-100">
                    {spreadsheetFiles.map(file => {
                      const isSelected = selectedFileNames.includes(file.filename);
                      return (
                        <div
                          key={file.path || file.filename}
                          onClick={() => toggleFileSelection(file.filename)}
                          className={`p-3.5 flex items-center justify-between gap-3 hover:bg-slate-50 cursor-pointer transition-colors ${
                            isSelected ? 'bg-sky-50/50' : ''
                          }`}
                        >
                          <div className="flex items-center gap-3 min-w-0">
                            <div className="text-sky-600">
                              {isSelected ? (
                                <CheckSquare className="w-5 h-5 text-sky-600" />
                              ) : (
                                <Square className="w-5 h-5 text-slate-300" />
                              )}
                            </div>
                            <div className="p-2 bg-emerald-50 text-emerald-700 rounded-lg border border-emerald-100 shrink-0">
                              <FileSpreadsheet className="w-4 h-4" />
                            </div>
                            <div className="min-w-0">
                              <div className="text-sm font-semibold text-slate-900 truncate">
                                {file.filename}
                              </div>
                              <div className="text-xs text-slate-500 flex items-center gap-2">
                                <span>{file.fileSizeFormatted}</span>
                                <span>&bull;</span>
                                <span>Modified: {new Date(file.lastModified).toLocaleDateString()}</span>
                              </div>
                            </div>
                          </div>

                          <div className="flex items-center gap-2 shrink-0">
                            {isSelected ? (
                              <span className="text-[11px] font-semibold px-2 py-0.5 rounded-full bg-sky-100 text-sky-800 border border-sky-200">
                                Ready to Pull
                              </span>
                            ) : (
                              <span className="text-[11px] font-medium px-2 py-0.5 rounded-full bg-slate-100 text-slate-600">
                                Excluded
                              </span>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Status / Error Message */}
              {pullError && (
                <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-800 flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
                  <span>{pullError}</span>
                </div>
              )}

              {pullProgressMessage && (
                <div className="p-3.5 bg-sky-50 border border-sky-200 rounded-xl text-xs text-sky-800 flex items-center gap-2 font-medium">
                  <RefreshCw className="w-4 h-4 text-sky-600 animate-spin shrink-0" />
                  <span>{pullProgressMessage}</span>
                </div>
              )}

              {/* Action Buttons */}
              <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-3">
                <div className="text-xs text-slate-500">
                  {selectedFileNames.length} of {spreadsheetFiles.length} files selected
                </div>

                <div className="flex items-center gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={onClose}
                    className="px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleExecutePull}
                    disabled={isPulling || selectedFileNames.length === 0}
                    className="flex-1 sm:flex-none inline-flex items-center justify-center gap-2 px-5 py-2.5 bg-sky-600 hover:bg-sky-700 text-white text-xs font-semibold rounded-lg shadow-sm transition-all disabled:opacity-50"
                  >
                    {isPulling ? (
                      <>
                        <RefreshCw className="w-4 h-4 animate-spin" />
                        Pulling & Merging Workbooks...
                      </>
                    ) : (
                      <>
                        <DownloadCloud className="w-4 h-4" />
                        Pull Selected Sheets ({selectedFileNames.length})
                      </>
                    )}
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: TIME-INTERVAL AUTO-SYNC */}
          {activeTab === 'interval' && (
            <div className="space-y-6">
              {/* Auto Sync Switch Card */}
              <div className="p-5 bg-gradient-to-br from-sky-50 to-indigo-50/50 rounded-2xl border border-sky-100">
                <div className="flex items-start justify-between gap-4">
                  <div className="space-y-1">
                    <div className="flex items-center gap-2">
                      <Clock className="w-5 h-5 text-sky-600" />
                      <h3 className="text-sm font-bold text-slate-900">
                        Periodic Background Auto-Sync Engine
                      </h3>
                    </div>
                    <p className="text-xs text-slate-600 max-w-xl">
                      When enabled, the server will periodically query Nextcloud in the background, download updated Excel grade files, detect division sheets (A, B, C, D...), merge student records, and persist them safely in server storage.
                    </p>
                  </div>

                  <label className="relative inline-flex items-center cursor-pointer shrink-0 mt-1">
                    <input
                      type="checkbox"
                      checked={autoSyncEnabled}
                      onChange={(e) => {
                        const val = e.target.checked;
                        setAutoSyncEnabled(val);
                        handleSaveConfig(val);
                      }}
                      className="sr-only peer"
                    />
                    <div className="w-11 h-6 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-sky-600"></div>
                  </label>
                </div>

                {autoSyncEnabled && (
                  <div className="mt-4 pt-3 border-t border-sky-200/60 flex items-center justify-between text-xs">
                    <span className="flex items-center gap-2 font-medium text-emerald-800">
                      <span className="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span>
                      Auto-sync is currently ACTIVE on server
                    </span>
                    <span className="text-slate-500 font-medium">
                      Interval: Every {syncIntervalMinutes} min(s)
                    </span>
                  </div>
                )}
              </div>

              {/* Sync Interval Picker */}
              <div className="space-y-3">
                <label className="text-xs font-bold text-slate-800 uppercase tracking-wider">
                  Select Synchronization Interval
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                  {[
                    { label: '1 Minute (Testing)', value: 1 },
                    { label: '5 Minutes', value: 5 },
                    { label: '15 Minutes (Default)', value: 15 },
                    { label: '30 Minutes', value: 30 },
                    { label: '1 Hour', value: 60 },
                    { label: '6 Hours', value: 360 },
                    { label: '12 Hours', value: 720 },
                    { label: '24 Hours (Daily)', value: 1440 },
                  ].map(option => (
                    <button
                      key={option.value}
                      onClick={() => {
                        setSyncIntervalMinutes(option.value);
                        setCustomMinutesInput(String(option.value));
                      }}
                      className={`p-3 rounded-xl border text-left transition-all ${
                        syncIntervalMinutes === option.value
                          ? 'border-sky-500 bg-sky-50 text-sky-900 font-bold ring-2 ring-sky-200'
                          : 'border-slate-200 bg-white text-slate-700 hover:border-slate-300'
                      }`}
                    >
                      <div className="text-xs font-semibold">{option.label}</div>
                      <div className="text-[10px] text-slate-400 mt-0.5">Every {option.value} min(s)</div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Custom Interval Option */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3">
                <div>
                  <div className="text-xs font-semibold text-slate-800">Custom Sync Interval</div>
                  <div className="text-[11px] text-slate-500">Specify any custom interval in minutes</div>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    min="1"
                    max="10080"
                    value={customMinutesInput}
                    onChange={(e) => {
                      setCustomMinutesInput(e.target.value);
                      const num = parseInt(e.target.value, 10);
                      if (!isNaN(num) && num > 0) {
                        setSyncIntervalMinutes(num);
                      }
                    }}
                    className="w-24 px-3 py-1.5 text-xs rounded-lg border border-slate-300 bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-sky-500"
                    placeholder="Minutes"
                  />
                  <span className="text-xs text-slate-600 font-medium">Minutes</span>
                </div>
              </div>

              {/* Save & Apply Button */}
              <div className="pt-2 flex justify-end gap-2.5">
                <button
                  onClick={async () => {
                    await handleSaveConfig();
                    alert(`Nextcloud sync settings saved! Auto-sync is ${autoSyncEnabled ? `active every ${syncIntervalMinutes} minute(s)` : 'paused'}.`);
                  }}
                  className="inline-flex items-center gap-2 px-5 py-2.5 bg-sky-600 hover:bg-sky-700 text-white text-xs font-semibold rounded-lg shadow-sm transition-all"
                >
                  <Check className="w-4 h-4" />
                  Save & Apply Sync Schedule
                </button>
              </div>
            </div>
          )}

          {/* TAB 3: CONNECTION & CREDENTIALS */}
          {activeTab === 'connection' && (
            <div className="space-y-5">
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider mb-1">
                  Nextcloud WebDAV Connection Details
                </h3>
                <p className="text-xs text-slate-500">
                  Connect securely via WebDAV API. Use your Nextcloud username and dedicated App Password.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-semibold text-slate-700">Nextcloud Server Host URL</label>
                  <input
                    type="text"
                    value={url}
                    onChange={(e) => setUrl(e.target.value)}
                    placeholder="https://cloud.jhcnexus.space"
                    className="w-full px-3.5 py-2 text-xs rounded-lg border border-slate-300 focus:ring-2 focus:ring-sky-500 outline-none"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-semibold text-slate-700">WebDAV Source Folder</label>
                  <input
                    type="text"
                    value={sourceFolder}
                    onChange={(e) => setSourceFolder(e.target.value)}
                    placeholder="/ExcelImports or /StudentClasses"
                    className="w-full px-3.5 py-2 text-xs rounded-lg border border-slate-300 focus:ring-2 focus:ring-sky-500 outline-none"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-semibold text-slate-700">Username / Account</label>
                  <input
                    type="text"
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder="truenas_admin"
                    className="w-full px-3.5 py-2 text-xs rounded-lg border border-slate-300 focus:ring-2 focus:ring-sky-500 outline-none"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-semibold text-slate-700">App Password / Token</label>
                  <input
                    type="password"
                    value={appPassword}
                    onChange={(e) => setAppPassword(e.target.value)}
                    placeholder="••••••••••••••••••••"
                    className="w-full px-3.5 py-2 text-xs rounded-lg border border-slate-300 focus:ring-2 focus:ring-sky-500 outline-none"
                  />
                </div>
              </div>

              {/* Test Result Display */}
              {testResult && (
                <div className={`p-4 rounded-xl border text-xs ${
                  testResult.success 
                    ? 'bg-emerald-50 border-emerald-200 text-emerald-900' 
                    : 'bg-rose-50 border-rose-200 text-rose-900'
                }`}>
                  <div className="flex items-center gap-2 font-bold mb-2">
                    {testResult.success ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                    ) : (
                      <AlertTriangle className="w-4 h-4 text-rose-600" />
                    )}
                    <span>{testResult.message}</span>
                  </div>

                  {testResult.checks && (
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 pt-2 border-t border-slate-200/50">
                      <div className="flex items-center gap-1.5">
                        <span className={`w-2 h-2 rounded-full ${testResult.checks.hostReachability ? 'bg-emerald-500' : 'bg-rose-500'}`}></span>
                        <span>Host Reachable</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <span className={`w-2 h-2 rounded-full ${testResult.checks.webdavHandshake ? 'bg-emerald-500' : 'bg-rose-500'}`}></span>
                        <span>WebDAV Handshake</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <span className={`w-2 h-2 rounded-full ${testResult.checks.authValid ? 'bg-emerald-500' : 'bg-rose-500'}`}></span>
                        <span>Auth Valid</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <span className={`w-2 h-2 rounded-full ${testResult.checks.folderExists ? 'bg-emerald-500' : 'bg-rose-500'}`}></span>
                        <span>Folder Found</span>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* Actions */}
              <div className="pt-2 flex items-center justify-between">
                <button
                  onClick={handleTestConnection}
                  disabled={isTestingConnection}
                  className="inline-flex items-center gap-1.5 px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg transition-colors border border-slate-300"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isTestingConnection ? 'animate-spin' : ''}`} />
                  {isTestingConnection ? 'Testing...' : 'Test Nextcloud Connection'}
                </button>

                <button
                  onClick={async () => {
                    await handleSaveConfig();
                    alert('Nextcloud connection credentials saved successfully.');
                  }}
                  className="inline-flex items-center gap-1.5 px-5 py-2 bg-sky-600 hover:bg-sky-700 text-white text-xs font-semibold rounded-lg shadow-sm transition-all"
                >
                  <Check className="w-4 h-4" />
                  Save Credentials
                </button>
              </div>
            </div>
          )}

          {/* TAB 4: SYNC LOGS & HISTORY */}
          {activeTab === 'history' && (
            <div className="space-y-4">
              <div className="flex items-center justify-between text-xs text-slate-600">
                <span className="font-semibold text-slate-800">
                  Recent Nextcloud Synchronization Runs
                </span>
                <span className="text-slate-400">
                  Last 50 automated & manual sync cycles
                </span>
              </div>

              {(!syncConfig.history || syncConfig.history.length === 0) ? (
                <div className="py-12 text-center text-slate-500 border border-slate-200 rounded-xl bg-slate-50">
                  <History className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                  <p className="text-sm font-medium text-slate-700">No sync runs recorded yet</p>
                  <p className="text-xs text-slate-400 mt-1">
                    Execute a pull or enable auto-sync to start logging synchronization activity.
                  </p>
                </div>
              ) : (
                <div className="border border-slate-200 rounded-xl overflow-hidden divide-y divide-slate-100 bg-white">
                  {syncConfig.history.map((logItem) => (
                    <div key={logItem.id} className="p-3.5 flex items-start justify-between gap-3 text-xs">
                      <div className="flex items-start gap-3">
                        <div className="mt-0.5">
                          {logItem.status === 'success' ? (
                            <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                          ) : (
                            <AlertTriangle className="w-4 h-4 text-rose-600" />
                          )}
                        </div>
                        <div>
                          <div className="font-semibold text-slate-900">
                            {logItem.message}
                          </div>
                          <div className="text-slate-400 text-[11px] mt-0.5 flex items-center gap-2">
                            <span>{new Date(logItem.timestamp).toLocaleString()}</span>
                            <span>&bull;</span>
                            <span>{logItem.filesCount} file(s)</span>
                            <span>&bull;</span>
                            <span>{logItem.recordsCount} record(s)</span>
                            <span>&bull;</span>
                            <span>{(logItem.durationMs / 1000).toFixed(1)}s duration</span>
                          </div>
                        </div>
                      </div>

                      <span className={`text-[10px] font-semibold px-2 py-0.5 rounded-full ${
                        logItem.status === 'success'
                          ? 'bg-emerald-100 text-emerald-800'
                          : 'bg-rose-100 text-rose-800'
                      }`}>
                        {logItem.status.toUpperCase()}
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

        </div>

      </div>
    </div>
  );
};
