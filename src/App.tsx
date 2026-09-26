/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { 
  NavigationTab, 
  NextcloudConfig, 
  SupabaseConfig, 
  SyncSettings, 
  NextcloudFile, 
  WorkbookAnalysis, 
  WorksheetMapping, 
  ImportLog, 
  ImportAuditError, 
  LogMessage, 
  SheetAnalysis,
  SyncConflictRecord,
  LiveSchedulerStatus
} from './types';
import { StorageService } from './services/storage';
import { ApiClient } from './services/apiClient';
import { TwoWaySyncEngine } from './services/twoWaySyncEngine';
import { SchemaGenerator } from './services/schemaGenerator';
import { Navbar } from './components/Navbar';
import { Sidebar } from './components/Sidebar';
import { DashboardView } from './components/DashboardView';
import { NextcloudView } from './components/NextcloudView';
import { SupabaseView } from './components/SupabaseView';
import { ExcelFilesView } from './components/ExcelFilesView';
import { WorkbookAnalyzerView } from './components/WorkbookAnalyzerView';
import { MappingsView } from './components/MappingsView';
import { ImportDryRunView } from './components/ImportDryRunView';
import { TwoWaySyncView } from './components/TwoWaySyncView';
import { WorkerDashboardView } from './components/WorkerDashboardView';
import { ImportHistoryView } from './components/ImportHistoryView';
import { ErrorsView } from './components/ErrorsView';
import { LogsView } from './components/LogsView';
import { SettingsView } from './components/SettingsView';
import { ServedSheetsMatrixView } from './components/ServedSheetsMatrixView';
import { MultiGradeMergerView } from './components/MultiGradeMergerView';
import { AiAssistantModal } from './components/AiAssistantModal';
import { SampleWorkbookModal } from './components/SampleWorkbookModal';
import { SecretsVaultModal } from './components/SecretsVaultModal';

export default function App() {
  const handleFetchNextcloudFiles = async () => {
    try {
      const res = await ApiClient.listNextcloudFiles(nextcloud);
      if (res.success && Array.isArray(res.files)) {
        setFiles(res.files);
        StorageService.saveFiles(res.files);

        const liveFileNames = new Set(res.files.map(f => f.filename.toLowerCase().trim()));

        // Reconcile mappings with server reconciled mappings or live files
        if (res.mappings && Array.isArray(res.mappings)) {
          const cleanMappings = res.mappings.filter((m: any) => {
            const wb = (m.workbookName || '').toLowerCase().trim();
            return !wb || liveFileNames.has(wb);
          });
          setMappings(cleanMappings);
          StorageService.saveMappings(cleanMappings);
        } else {
          setMappings(prev => {
            const clean = prev.filter(m => {
              const wb = (m.workbookName || '').toLowerCase().trim();
              return !wb || liveFileNames.has(wb);
            });
            StorageService.saveMappings(clean);
            return clean;
          });
        }

        // If current analysis is null, or of a fake/deleted file, switch to first live file
        if (!currentAnalysis?.filename || currentAnalysis.filename === 'students_complex.xlsx' || currentAnalysis.filename === 'students.xlsx' || !liveFileNames.has(currentAnalysis.filename.toLowerCase().trim())) {
          if (res.files.length > 0) {
            handleSelectFileForAnalysis(res.files[0], currentTab);
          } else {
            setCurrentAnalysis(null);
            StorageService.saveCurrentAnalysis(null as any);
          }
        }
      }
    } catch (e) {
      console.error('Failed to fetch files from Nextcloud:', e);
    }
  };

  // Application State
  const [currentTab, setCurrentTab] = useState<NavigationTab>('dashboard');
  const [nextcloud, setNextcloud] = useState<NextcloudConfig>(() => StorageService.getNextcloudConfig());
  const [supabase, setSupabase] = useState<SupabaseConfig>(() => StorageService.getSupabaseConfig());
  const [syncSettings, setSyncSettings] = useState<SyncSettings>(() => StorageService.getSyncSettings());
  const [files, setFiles] = useState<NextcloudFile[]>(() => StorageService.getFiles());
  const [currentAnalysis, setCurrentAnalysis] = useState<WorkbookAnalysis | null>(() => StorageService.getCurrentAnalysis());
  const [mappings, setMappings] = useState<WorksheetMapping[]>(() => StorageService.getMappings());
  const [importLogs, setImportLogs] = useState<ImportLog[]>(() => StorageService.getImportLogs());
  const [importErrors, setImportErrors] = useState<ImportAuditError[]>(() => StorageService.getImportErrors());
  const [workerLogs, setWorkerLogs] = useState<LogMessage[]>(() => StorageService.getWorkerLogs());
  const [databaseState, setDatabaseState] = useState<Record<string, any[]>>(() => StorageService.getDatabaseState());
  const [conflicts, setConflicts] = useState<SyncConflictRecord[]>(() => StorageService.getConflicts());
  const [schedulerStatus, setSchedulerStatus] = useState<LiveSchedulerStatus | null>(null);

  const handleFetchSchedulerStatus = async () => {
    try {
      const res = await ApiClient.getSchedulerStatus();
      if (res.success && res.status) {
        setSchedulerStatus(res.status);
      }
    } catch (e) {
      console.error('Failed to poll scheduler status:', e);
    }
  };

  // Coolify server disk sync status
  const [serverSyncTime, setServerSyncTime] = useState<string | null>(null);
  const [isSyncingServer, setIsSyncingServer] = useState(false);

  // Load unified Coolify server state & sync with browser localStorage
  const syncServerAndLocalStorage = async () => {
    setIsSyncingServer(true);
    try {
      const serverState = await ApiClient.getFullServerState();
      if (serverState.success) {
        // 1. If server has environment variables or disk credentials from Coolify, hydrate local state
        if (serverState.nextcloud?.url) {
          setNextcloud(prev => {
            const updated = { ...prev, ...serverState.nextcloud };
            StorageService.saveNextcloudConfig(updated);
            return updated;
          });
        }
        if (serverState.supabase?.url) {
          setSupabase(prev => {
            const updated = { ...prev, ...serverState.supabase };
            StorageService.saveSupabaseConfig(updated);
            return updated;
          });
        }
        if (serverState.syncSettings) {
          setSyncSettings(prev => {
            const updated = { ...prev, ...serverState.syncSettings };
            StorageService.saveSyncSettings(updated);
            return updated;
          });
        }
        if (serverState.mappings && serverState.mappings.length > 0) {
          setMappings(serverState.mappings);
          StorageService.saveMappings(serverState.mappings);
        } else {
          // Also try pulling directly from Supabase & server disk
          ApiClient.loadPermanentMappings().then(res => {
            if (res.success && res.mappings && res.mappings.length > 0) {
              setMappings(res.mappings);
              StorageService.saveMappings(res.mappings);
            }
          }).catch(() => {});
        }
        if (serverState.workerStatus) {
          setSchedulerStatus(serverState.workerStatus);
        }
        setServerSyncTime(serverState.savedAt || new Date().toISOString());

        // 2. Push whatever client has in localStorage (mappings, presets, etc.) to the server disk
        const currentMappings = StorageService.getMappings();
        const currentPresets = StorageService.getAiPresets();
        if (currentMappings.length > 0 || currentPresets.length > 0) {
          ApiClient.saveFullServerState({
            nextcloud: serverState.nextcloud,
            supabase: serverState.supabase,
            syncSettings: serverState.syncSettings,
            mappings: currentMappings.length > 0 ? currentMappings : serverState.mappings,
            presets: currentPresets,
            workbookInfo: currentAnalysis ? {
              filename: currentAnalysis.filename,
              fileHash: currentAnalysis.fileHash,
              totalWorksheets: currentAnalysis.totalWorksheets
            } : undefined,
          }).catch(() => {});
        }
      }
    } catch (e) {
      console.warn('Auto-sync with Coolify server state notice:', e);
    } finally {
      setIsSyncingServer(false);
    }
  };

  // Initialize sample data on first visit, load real files from Nextcloud & poll production scheduler
  useEffect(() => {
    StorageService.initSampleIfNeeded();
    handleFetchNextcloudFiles();
    handleFetchSchedulerStatus();
    syncServerAndLocalStorage();

    // Poll live production scheduler state every 10 seconds
    const interval = setInterval(() => {
      handleFetchSchedulerStatus();
    }, 10000);

    // Initial scheduler configuration check
    ApiClient.configureScheduler({
      enabled: syncSettings.autoSyncEnabled,
      intervalLabel: syncSettings.syncInterval,
      workerUrl: syncSettings.workerUrl,
      nextcloud,
      supabase,
      mappings,
    }).then(res => {
      if (res.success && res.status) {
        setSchedulerStatus(res.status);
      }
    }).catch(() => {});

    // Load stored conflicts if available
    const stored = StorageService.getConflicts();
    if (stored.length > 0) {
      setConflicts(stored);
    }

    return () => clearInterval(interval);
  }, []);
  const [isAiModalOpen, setIsAiModalOpen] = useState(false);
  const [isSampleModalOpen, setIsSampleModalOpen] = useState(false);
  const [isSecretsModalOpen, setIsSecretsModalOpen] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [aiTargetSheet, setAiTargetSheet] = useState<SheetAnalysis | null>(null);

  // Sync Cycle State
  const [isSyncing, setIsSyncing] = useState(false);

  // Update handlers with dual persistence (Browser LocalStorage + Coolify Server Disk)
  const handleSaveNextcloud = (cfg: NextcloudConfig) => {
    setNextcloud(cfg);
    StorageService.saveNextcloudConfig(cfg);
    ApiClient.saveFullServerState({
      nextcloud: cfg,
      supabase,
      syncSettings,
      mappings,
    }).then(() => setServerSyncTime(new Date().toISOString())).catch(() => {});
  };

  const handleSaveSupabase = (cfg: SupabaseConfig) => {
    setSupabase(cfg);
    StorageService.saveSupabaseConfig(cfg);
    ApiClient.saveFullServerState({
      nextcloud,
      supabase: cfg,
      syncSettings,
      mappings,
    }).then(() => setServerSyncTime(new Date().toISOString())).catch(() => {});
  };

  const handleSaveSettings = (s: SyncSettings) => {
    setSyncSettings(s);
    StorageService.saveSyncSettings(s);
    ApiClient.saveFullServerState({
      nextcloud,
      supabase,
      syncSettings: s,
      mappings,
    }).then(() => setServerSyncTime(new Date().toISOString())).catch(() => {});
  };

  const handleSaveMappings = (newMappings: WorksheetMapping[]) => {
    const currentWb = currentAnalysis?.filename;
    const stampedNewMappings = newMappings.map(m => {
      if (!m.workbookName && currentWb) {
        return { ...m, workbookName: currentWb };
      }
      return m;
    });

    // Non-destructive merge with existing mappings by composite key (workbookName::worksheetName)
    const getMappingKey = (m: WorksheetMapping) => {
      const wb = String(m.workbookName || '').toLowerCase().trim();
      const ws = String(m.worksheetName || '').toLowerCase().trim();
      if (wb && ws) return `${wb}::${ws}`;
      if (m.id) return String(m.id);
      if (ws) return `unspecified::${ws}`;
      return `wm-${Math.random()}`;
    };

    const mapByKey = new Map<string, WorksheetMapping>();
    mappings.forEach(m => {
      mapByKey.set(getMappingKey(m), m);
    });
    stampedNewMappings.forEach(m => {
      const key = getMappingKey(m);
      const prev = mapByKey.get(key);
      mapByKey.set(key, { ...prev, ...m });
    });
    const merged = Array.from(mapByKey.values());

    setMappings(merged);
    StorageService.saveMappings(merged);
    ApiClient.savePermanentMappings({
      mappings: merged,
      workbookInfo: currentAnalysis ? {
        filename: currentAnalysis.filename,
        fileHash: currentAnalysis.fileHash,
        totalWorksheets: merged.length
      } : undefined,
      supabase,
    }).catch(() => {});

    ApiClient.saveFullServerState({
      nextcloud,
      supabase,
      syncSettings,
      mappings: merged,
      workbookInfo: currentAnalysis ? {
        filename: currentAnalysis.filename,
        fileHash: currentAnalysis.fileHash,
        totalWorksheets: merged.length
      } : undefined,
    }).then(() => setServerSyncTime(new Date().toISOString())).catch(() => {});
  };

  const handleUpdateAnalysis = (analysis: WorkbookAnalysis) => {
    setCurrentAnalysis(analysis);
    StorageService.saveCurrentAnalysis(analysis);
    if (analysis.base64Data) {
      ApiClient.cacheActiveWorkbook({
        base64Data: analysis.base64Data,
        filename: analysis.filename
      }).catch(console.warn);
    }
  };

  const handleUpdateDatabase = (tableName: string, newRecords: any[]) => {
    const updated = { ...databaseState, [tableName]: newRecords };
    setDatabaseState(updated);
    StorageService.saveDatabaseState(updated);
  };

  const handleAddImportLog = (newLog: ImportLog, newErrors: any[]) => {
    const updatedLogs = [newLog, ...importLogs];
    setImportLogs(updatedLogs);
    StorageService.saveImportLogs(updatedLogs);

    if (newErrors.length > 0) {
      const formattedErrors: ImportAuditError[] = newErrors.map((err, i) => ({
        id: `err-${Date.now()}-${i}`,
        importLogId: newLog.id,
        filename: newLog.filename,
        worksheetName: err.worksheetName,
        rowNumber: err.rowNumber,
        excelColumn: err.excelColumn,
        columnName: err.columnName,
        rawValue: err.rawValue,
        errorMessage: err.errorMessage,
        errorType: err.errorType,
        createdAt: new Date().toISOString()
      }));
      const updatedErrors = [...formattedErrors, ...importErrors];
      setImportErrors(updatedErrors);
      StorageService.saveImportErrors(updatedErrors);
    }
  };

  const handleClearErrors = () => {
    setImportErrors([]);
    StorageService.saveImportErrors([]);
  };

  const handleClearLogs = () => {
    setWorkerLogs([]);
    StorageService.saveWorkerLogs([]);
    ApiClient.clearServerLogs().catch(() => {});
  };

  const handleRefreshAll = () => {
    setNextcloud(StorageService.getNextcloudConfig());
    setSupabase(StorageService.getSupabaseConfig());
    setSyncSettings(StorageService.getSyncSettings());
    setFiles(StorageService.getFiles());
    setImportLogs(StorageService.getImportLogs());
    setDatabaseState(StorageService.getDatabaseState());
    setConflicts(StorageService.getConflicts());
    handleFetchSchedulerStatus();
    ApiClient.getLiveLogs({ limit: 100 }).then(res => {
      if (res.success && res.logs && res.logs.length > 0) {
        setWorkerLogs(res.logs);
      }
    }).catch(() => {});
  };

  const handleUpdateConflicts = (newConflicts: SyncConflictRecord[]) => {
    setConflicts(newConflicts);
    StorageService.saveConflicts(newConflicts);
  };

  const handleTriggerSync = async () => {
    setIsSyncing(true);
    const now = new Date();
    const newLogMsg: LogMessage = {
      id: `l-${Date.now()}`,
      timestamp: now.toISOString(),
      level: 'info',
      component: 'Worker',
      message: `Full pipeline synchronization initiated. Polling Nextcloud WebDAV at ${nextcloud.url}...`
    };

    let updatedWorkerLogs = [newLogMsg, ...workerLogs];
    setWorkerLogs(updatedWorkerLogs);
    StorageService.saveWorkerLogs(updatedWorkerLogs);

    try {
      // 1. Refresh file list from Nextcloud
      const filesResult = await ApiClient.listNextcloudFiles(nextcloud);
      if (filesResult.success && filesResult.files) {
        setFiles(filesResult.files);
        StorageService.saveFiles(filesResult.files);
      }

      // 2. Execute end-to-end sync pipeline across ALL served files (download -> dynamic parse -> transform -> Supabase upsert)
      const pipelineRes = await ApiClient.executeFullPipelineSync({
        nextcloud,
        supabase,
        mappings,
        syncAllFiles: true,
      });

      if (pipelineRes.success) {
        const fileNamesDisplay = pipelineRes.filesSynced && pipelineRes.filesSynced.length > 0
          ? pipelineRes.filesSynced.join(', ')
          : (pipelineRes.filename || 'All served files');

        const completeMsg: LogMessage = {
          id: `l-done-${Date.now()}`,
          timestamp: new Date().toISOString(),
          level: 'success',
          component: 'Worker',
          message: `Multi-file pipeline sync completed successfully for [${fileNamesDisplay}]! Inserted: ${pipelineRes.totalInserted || 0}, Updated: ${pipelineRes.totalUpdated || 0}, Errors: ${pipelineRes.totalFailed || 0} across ${pipelineRes.syncResults?.length || 0} table mappings.`
        };

        updatedWorkerLogs = [completeMsg, ...updatedWorkerLogs];
        setWorkerLogs(updatedWorkerLogs);
        StorageService.saveWorkerLogs(updatedWorkerLogs);

        // Add to import history
        const logId = `log-${Date.now()}`;
        const newImportLog: ImportLog = {
          id: logId,
          filename: pipelineRes.filename || 'All served files',
          filePath: nextcloud.sourceFolder + '/' + (pipelineRes.filesSynced?.[0] || 'all_served'),
          fileHash: pipelineRes.fileHash || 'synced-hash',
          status: (pipelineRes.totalFailed && pipelineRes.totalFailed > 0) ? 'Partial Success' : 'Success',
          isDryRun: false,
          numberOfWorksheets: pipelineRes.syncResults?.length || 1,
          rowsProcessed: (pipelineRes.totalInserted || 0) + (pipelineRes.totalUpdated || 0) + (pipelineRes.totalFailed || 0),
          rowsInserted: pipelineRes.totalInserted || 0,
          rowsUpdated: pipelineRes.totalUpdated || 0,
          rowsFailed: pipelineRes.totalFailed || 0,
          startedAt: now.toISOString(),
          completedAt: new Date().toISOString(),
          durationMs: 850,
          errorSummary: (pipelineRes.totalFailed && pipelineRes.totalFailed > 0) ? `${pipelineRes.totalFailed} row errors` : undefined,
          details: { syncResults: pipelineRes.syncResults, filesSynced: pipelineRes.filesSynced }
        };

        const updatedHistory = [newImportLog, ...importLogs];
        setImportLogs(updatedHistory);
        StorageService.saveImportLogs(updatedHistory);

        // If there were any errors, save to audit errors
        if (pipelineRes.errors && pipelineRes.errors.length > 0) {
          const formattedErrors: ImportAuditError[] = pipelineRes.errors.map((err: any, idx: number) => ({
            id: `err-${Date.now()}-${idx}`,
            importLogId: logId,
            filename: pipelineRes.filename || 'students.xlsx',
            worksheetName: err.sheetName || 'Sheet1',
            rowNumber: err.rowNumber || 1,
            excelColumn: 'A',
            columnName: 'record',
            rawValue: '',
            errorMessage: err.error || 'Upsert error',
            errorType: 'foreign_key' as const,
            createdAt: new Date().toISOString()
          }));
          const updatedErrors = [...formattedErrors, ...importErrors];
          setImportErrors(updatedErrors);
          StorageService.saveImportErrors(updatedErrors);
        }
      } else {
        const errorMsg: LogMessage = {
          id: `l-err-${Date.now()}`,
          timestamp: new Date().toISOString(),
          level: 'error',
          component: 'Worker',
          message: `Pipeline sync failed: ${pipelineRes.error || 'Unknown error'}`
        };
        updatedWorkerLogs = [errorMsg, ...updatedWorkerLogs];
        setWorkerLogs(updatedWorkerLogs);
        StorageService.saveWorkerLogs(updatedWorkerLogs);
      }
    } catch (e: any) {
      const errorMsg: LogMessage = {
        id: `l-err-${Date.now()}`,
        timestamp: new Date().toISOString(),
        level: 'error',
        component: 'Worker',
        message: `Sync cycle failed: ${e.message}`
      };
      const finalLogs = [errorMsg, ...updatedWorkerLogs];
      setWorkerLogs(finalLogs);
      StorageService.saveWorkerLogs(finalLogs);
    } finally {
      setIsSyncing(false);
      handleFetchSchedulerStatus();
    }
  };

  const handleOpenAiAssistant = (sheet?: SheetAnalysis) => {
    if (sheet) {
      setAiTargetSheet(sheet);
    } else if (currentAnalysis?.worksheets?.[0]) {
      setAiTargetSheet(currentAnalysis.worksheets[0]);
    }
    setIsAiModalOpen(true);
  };

  const handleApplyAiSuggestions = (suggestedMapping: WorksheetMapping) => {
    const targetWb = (suggestedMapping.workbookName || currentAnalysis?.filename || '').toLowerCase().trim();
    const existingIndex = mappings.findIndex(m => 
      (!targetWb || (m.workbookName || '').toLowerCase().trim() === targetWb) &&
      m.worksheetName.toLowerCase().trim() === suggestedMapping.worksheetName.toLowerCase().trim()
    );
    let updated: WorksheetMapping[];
    if (existingIndex >= 0) {
      updated = [...mappings];
      updated[existingIndex] = { ...suggestedMapping, workbookName: suggestedMapping.workbookName || currentAnalysis?.filename || 'Workbook.xlsx' };
    } else {
      updated = [...mappings, { ...suggestedMapping, workbookName: suggestedMapping.workbookName || currentAnalysis?.filename || 'Workbook.xlsx' }];
    }
    handleSaveMappings(updated);
  };

  const handleUpdateMappingHeaderDataRow = (sheetName: string, headerRow: number, dataStartRow: number, targetWbName?: string) => {
    const wb = (targetWbName || currentAnalysis?.filename || '').toLowerCase().trim();
    const updated = mappings.map(m => {
      const matchWb = !wb || (m.workbookName || '').toLowerCase().trim() === wb;
      if (matchWb && m.worksheetName.toLowerCase().trim() === sheetName.toLowerCase().trim()) {
        return { ...m, headerRow, dataStartRow };
      }
      return m;
    });
    handleSaveMappings(updated);
  };

  const handleSelectFileForAnalysis = async (file: NextcloudFile, targetTab?: NavigationTab) => {
    if (targetTab) {
      setCurrentTab(targetTab);
    } else {
      setCurrentTab('analyzer');
    }
    // Ensure analyzer knows which file was clicked immediately
    if (!currentAnalysis || currentAnalysis.filename !== file.filename) {
      setCurrentAnalysis({
        filename: file.filename,
        fileHash: file.fileHash,
        fileSize: file.fileSize,
        fileSizeFormatted: file.fileSizeFormatted,
        totalWorksheets: 0,
        worksheets: [],
        analyzedAt: new Date().toISOString(),
      });
    }
    try {
      const res = await ApiClient.fetchAndParseWorkbook(nextcloud, file.path, file.filename);
      if (res.success && res.analysis) {
        handleUpdateAnalysis(res.analysis);

        // Auto-generate initial mappings if this file has no mappings yet
        const fileWb = file.filename.toLowerCase().trim();
        const existingForThisFile = mappings.filter(m => (m.workbookName || '').toLowerCase().trim() === fileWb);
        if (existingForThisFile.length === 0 && res.analysis.worksheets && res.analysis.worksheets.length > 0) {
          const autoMappings = SchemaGenerator.generateMappingsFromPlans(
            res.analysis,
            SchemaGenerator.generateSchemas(res.analysis, 'SEPARATE_TABLES', {}, []),
            'SEPARATE_TABLES',
            {}
          );
          if (autoMappings.length > 0) {
            const updated = [...mappings, ...autoMappings];
            handleSaveMappings(updated);
          }
        }
        return;
      }
    } catch (e) {
      console.error('Error fetching file for analysis:', e);
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans text-slate-900 antialiased">
      {/* Top Navbar */}
      <Navbar
        nextcloud={nextcloud}
        supabase={supabase}
        syncSettings={syncSettings}
        onRefreshAll={handleRefreshAll}
        onOpenSampleModal={() => setIsSampleModalOpen(true)}
        onOpenAiAssistant={() => handleOpenAiAssistant()}
        onOpenSecretsVault={() => setIsSecretsModalOpen(true)}
        onToggleMobileMenu={() => setIsMobileMenuOpen(prev => !prev)}
        allSecretsConnected={nextcloud.isConnected && supabase.isConnected}
        schedulerStatus={schedulerStatus ?? undefined}
      />

      {/* Main Layout Body */}
      <div className="flex-1 flex flex-col md:flex-row relative">
        {/* Left Sidebar */}
        <Sidebar
          currentTab={currentTab}
          onSelectTab={setCurrentTab}
          errorCount={importErrors.length}
          filesWaitingCount={files.filter(f => f.status === 'New' || f.status === 'Modified').length}
          conflictCount={conflicts.filter(c => c.state === 'CONFLICT').length}
          isOpenMobile={isMobileMenuOpen}
          onCloseMobile={() => setIsMobileMenuOpen(false)}
          onOpenSecretsVault={() => setIsSecretsModalOpen(true)}
          allSecretsConnected={nextcloud.isConnected && supabase.isConnected}
          schedulerStatus={schedulerStatus ?? undefined}
        />

        {/* Content Viewport */}
        <main className="flex-1 p-3 sm:p-6 lg:p-8 max-w-7xl mx-auto w-full min-w-0">
          {currentTab === 'dashboard' && (
            <DashboardView
              nextcloud={nextcloud}
              supabase={supabase}
              syncSettings={syncSettings}
              files={files}
              importLogs={importLogs}
              onNavigate={setCurrentTab}
              onTriggerSync={handleTriggerSync}
              isSyncing={isSyncing}
              onOpenSecretsVault={() => setIsSecretsModalOpen(true)}
              schedulerStatus={schedulerStatus ?? undefined}
            />
          )}

          {currentTab === 'grade_merger' && (
            <MultiGradeMergerView
              supabaseConfig={supabase}
              nextcloudConfig={nextcloud}
              files={files}
              onNavigate={setCurrentTab}
            />
          )}

          {currentTab === 'worker' && (
            <WorkerDashboardView
              nextcloud={nextcloud}
              supabase={supabase}
              syncSettings={syncSettings}
              mappings={mappings}
              schedulerStatus={schedulerStatus ?? undefined}
              onSaveSettings={handleSaveSettings}
              onNavigate={setCurrentTab}
              onRefreshScheduler={handleFetchSchedulerStatus}
            />
          )}

          {currentTab === 'nextcloud' && (
            <NextcloudView
              config={nextcloud}
              files={files}
              onSaveConfig={handleSaveNextcloud}
            />
          )}

          {currentTab === 'supabase' && (
            <SupabaseView
              config={supabase}
              databaseState={databaseState}
              mappings={mappings}
              onSaveConfig={handleSaveSupabase}
            />
          )}

          {currentTab === 'files' && (
            <ExcelFilesView
              files={files}
              onSelectFileForAnalysis={handleSelectFileForAnalysis}
              onNavigate={setCurrentTab}
              onRefreshFiles={handleFetchNextcloudFiles}
            />
          )}

          {currentTab === 'analyzer' && (
            <WorkbookAnalyzerView
              currentAnalysis={currentAnalysis}
              onAnalysisUpdate={handleUpdateAnalysis}
              onNavigate={setCurrentTab}
              onOpenAiAssistant={handleOpenAiAssistant}
              activeMappings={mappings}
              onUpdateMappingHeaderDataRow={handleUpdateMappingHeaderDataRow}
              nextcloudConfig={nextcloud}
              supabaseConfig={supabase}
              onSaveMappings={handleSaveMappings}
              files={files}
              onSelectFileForAnalysis={handleSelectFileForAnalysis}
            />
          )}

          {currentTab === 'mappings' && (
            <MappingsView
              mappings={mappings}
              onSaveMappings={handleSaveMappings}
              onNavigate={setCurrentTab}
              onOpenAiAssistant={() => handleOpenAiAssistant()}
              currentAnalysis={currentAnalysis}
              supabaseConfig={supabase}
              files={files}
              onRefreshFiles={handleFetchNextcloudFiles}
            />
          )}

          {currentTab === 'served_mappings' && (
            <ServedSheetsMatrixView
              mappings={mappings}
              currentAnalysis={currentAnalysis}
              files={files}
              importLogs={importLogs}
              supabaseConfig={supabase}
              nextcloudConfig={nextcloud}
              onNavigate={setCurrentTab}
              onSaveMappings={handleSaveMappings}
              onTriggerSync={handleTriggerSync}
            />
          )}

          {currentTab === 'import' && (
            <ImportDryRunView
              currentAnalysis={currentAnalysis}
              mappings={mappings}
              databaseState={databaseState}
              onUpdateDatabase={handleUpdateDatabase}
              onAddImportLog={handleAddImportLog}
              onNavigate={setCurrentTab}
              supabaseConfig={supabase}
              onSaveMappings={handleSaveMappings}
              files={files}
              nextcloudConfig={nextcloud}
              onSelectFileForAnalysis={handleSelectFileForAnalysis}
              onAnalysisUpdate={handleUpdateAnalysis}
            />
          )}

          {currentTab === 'twoway' && (
            <TwoWaySyncView
              currentAnalysis={currentAnalysis}
              mappings={mappings}
              databaseState={databaseState}
              nextcloudConfig={nextcloud}
              supabaseConfig={supabase}
              onUpdateDatabase={handleUpdateDatabase}
              onNavigate={setCurrentTab}
              conflicts={conflicts}
              onUpdateConflicts={handleUpdateConflicts}
              schedulerStatus={schedulerStatus ?? undefined}
              onRefreshScheduler={handleFetchSchedulerStatus}
              onUpdateAnalysis={handleUpdateAnalysis}
            />
          )}

          {currentTab === 'history' && (
            <ImportHistoryView
              logs={importLogs}
              onNavigate={setCurrentTab}
            />
          )}

          {currentTab === 'errors' && (
            <ErrorsView
              errors={importErrors}
              onClearErrors={handleClearErrors}
            />
          )}

          {currentTab === 'logs' && (
            <LogsView
              logs={workerLogs}
              onClearLogs={handleClearLogs}
              onRefreshLogs={handleRefreshAll}
            />
          )}

          {currentTab === 'settings' && (
            <SettingsView
              settings={syncSettings}
              onSaveSettings={handleSaveSettings}
              nextcloudConfig={nextcloud}
              supabaseConfig={supabase}
              mappings={mappings}
              onTriggerLiveSync={handleTriggerSync}
              onOpenSecretsVault={() => setIsSecretsModalOpen(true)}
              onSaveNextcloud={handleSaveNextcloud}
              onSaveSupabase={handleSaveSupabase}
              schedulerStatus={schedulerStatus ?? undefined}
              onRefreshScheduler={handleFetchSchedulerStatus}
            />
          )}
        </main>
      </div>

      {/* Unified Secrets Vault Modal */}
      <SecretsVaultModal
        isOpen={isSecretsModalOpen}
        onClose={() => setIsSecretsModalOpen(false)}
        nextcloud={nextcloud}
        supabase={supabase}
        syncSettings={syncSettings}
        onSaveNextcloud={handleSaveNextcloud}
        onSaveSupabase={handleSaveSupabase}
        onSaveSyncSettings={handleSaveSettings}
      />

      {/* AI Assistant Modal (Section 29) */}
      <AiAssistantModal
        isOpen={isAiModalOpen}
        onClose={() => setIsAiModalOpen(false)}
        currentSheet={aiTargetSheet}
        allSheets={currentAnalysis?.worksheets || []}
        workbookName={currentAnalysis?.filename || files?.[0]?.filename || 'Workbook.xlsx'}
        onApplyAiSuggestions={handleApplyAiSuggestions}
      />

      {/* Sample Complex Workbook Modal */}
      <SampleWorkbookModal
        isOpen={isSampleModalOpen}
        onClose={() => setIsSampleModalOpen(false)}
        onLoadIntoAnalyzer={(analysis) => {
          handleUpdateAnalysis(analysis);
          setCurrentTab('analyzer');
        }}
      />
    </div>
  );
}
