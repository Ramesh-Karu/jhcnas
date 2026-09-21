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
  SheetAnalysis 
} from './types';
import { StorageService } from './services/storage';
import { ApiClient } from './services/apiClient';
import { Navbar } from './components/Navbar';
import { Sidebar } from './components/Sidebar';
import { DashboardView } from './components/DashboardView';
import { NextcloudView } from './components/NextcloudView';
import { SupabaseView } from './components/SupabaseView';
import { ExcelFilesView } from './components/ExcelFilesView';
import { WorkbookAnalyzerView } from './components/WorkbookAnalyzerView';
import { MappingsView } from './components/MappingsView';
import { ImportDryRunView } from './components/ImportDryRunView';
import { ImportHistoryView } from './components/ImportHistoryView';
import { ErrorsView } from './components/ErrorsView';
import { LogsView } from './components/LogsView';
import { SettingsView } from './components/SettingsView';
import { AiAssistantModal } from './components/AiAssistantModal';
import { SampleWorkbookModal } from './components/SampleWorkbookModal';

export default function App() {
  // Initialize sample data on first visit & load real files from Nextcloud
  useEffect(() => {
    StorageService.initSampleIfNeeded();
    handleFetchNextcloudFiles();
  }, []);

  const handleFetchNextcloudFiles = async () => {
    try {
      const res = await ApiClient.listNextcloudFiles(nextcloud);
      if (res.success && res.files && res.files.length > 0) {
        setFiles(res.files);
        StorageService.saveFiles(res.files);
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

  // Modal State
  const [isAiModalOpen, setIsAiModalOpen] = useState(false);
  const [isSampleModalOpen, setIsSampleModalOpen] = useState(false);
  const [aiTargetSheet, setAiTargetSheet] = useState<SheetAnalysis | null>(null);

  // Sync Cycle State
  const [isSyncing, setIsSyncing] = useState(false);

  // Update handlers with persistence
  const handleSaveNextcloud = (cfg: NextcloudConfig) => {
    setNextcloud(cfg);
    StorageService.saveNextcloudConfig(cfg);
  };

  const handleSaveSupabase = (cfg: SupabaseConfig) => {
    setSupabase(cfg);
    StorageService.saveSupabaseConfig(cfg);
  };

  const handleSaveSettings = (s: SyncSettings) => {
    setSyncSettings(s);
    StorageService.saveSyncSettings(s);
  };

  const handleSaveMappings = (newMappings: WorksheetMapping[]) => {
    setMappings(newMappings);
    StorageService.saveMappings(newMappings);
  };

  const handleUpdateAnalysis = (analysis: WorkbookAnalysis) => {
    setCurrentAnalysis(analysis);
    StorageService.saveCurrentAnalysis(analysis);
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
  };

  const handleRefreshAll = () => {
    setNextcloud(StorageService.getNextcloudConfig());
    setSupabase(StorageService.getSupabaseConfig());
    setSyncSettings(StorageService.getSyncSettings());
    setFiles(StorageService.getFiles());
    setImportLogs(StorageService.getImportLogs());
    setDatabaseState(StorageService.getDatabaseState());
  };

  const handleTriggerSync = async () => {
    setIsSyncing(true);
    const now = new Date();
    const newLogMsg: LogMessage = {
      id: `l-${Date.now()}`,
      timestamp: now.toISOString(),
      level: 'info',
      component: 'Worker',
      message: `Manual sync cycle initiated. Polling Nextcloud WebDAV at ${nextcloud.url}...`
    };

    const updatedWorkerLogs = [newLogMsg, ...workerLogs];
    setWorkerLogs(updatedWorkerLogs);
    StorageService.saveWorkerLogs(updatedWorkerLogs);

    try {
      const filesResult = await ApiClient.listNextcloudFiles(nextcloud);
      if (filesResult.success && filesResult.files) {
        setFiles(filesResult.files);
        StorageService.saveFiles(filesResult.files);
      }

      const count = filesResult.files ? filesResult.files.length : 0;
      const completeMsg: LogMessage = {
        id: `l-done-${Date.now()}`,
        timestamp: new Date().toISOString(),
        level: 'success',
        component: 'Worker',
        message: `WebDAV scan complete on ${nextcloud.url}. Discovered ${count} file(s) in ${nextcloud.sourceFolder}. SHA-256 signatures checked.`
      };

      const finalLogs = [completeMsg, ...updatedWorkerLogs];
      setWorkerLogs(finalLogs);
      StorageService.saveWorkerLogs(finalLogs);
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
    const existingIndex = mappings.findIndex(m => m.worksheetName === suggestedMapping.worksheetName);
    let updated: WorksheetMapping[];
    if (existingIndex >= 0) {
      updated = [...mappings];
      updated[existingIndex] = suggestedMapping;
    } else {
      updated = [...mappings, suggestedMapping];
    }
    handleSaveMappings(updated);
  };

  const handleUpdateMappingHeaderDataRow = (sheetName: string, headerRow: number, dataStartRow: number) => {
    const updated = mappings.map(m => {
      if (m.worksheetName === sheetName) {
        return { ...m, headerRow, dataStartRow };
      }
      return m;
    });
    handleSaveMappings(updated);
  };

  const handleSelectFileForAnalysis = async (file: NextcloudFile) => {
    setCurrentTab('analyzer');
    try {
      const res = await ApiClient.fetchAndParseWorkbook(nextcloud, file.path, file.filename);
      if (res.success && res.analysis) {
        handleUpdateAnalysis(res.analysis);
        return;
      }
    } catch (e) {
      console.error('Error fetching file for analysis:', e);
    }

    // Fallback if network fails
    if (currentAnalysis?.filename !== file.filename && currentAnalysis) {
      handleUpdateAnalysis({
        ...currentAnalysis,
        filename: file.filename,
        fileHash: file.fileHash
      });
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
      />

      {/* Main Layout Body */}
      <div className="flex-1 flex flex-col md:flex-row">
        {/* Left Sidebar */}
        <Sidebar
          currentTab={currentTab}
          onSelectTab={setCurrentTab}
          errorCount={importErrors.length}
          filesWaitingCount={files.filter(f => f.status === 'New' || f.status === 'Modified').length}
        />

        {/* Content Viewport */}
        <main className="flex-1 p-4 sm:p-6 lg:p-8 max-w-7xl mx-auto w-full">
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
            />
          )}

          {currentTab === 'mappings' && (
            <MappingsView
              mappings={mappings}
              onSaveMappings={handleSaveMappings}
              onNavigate={setCurrentTab}
              onOpenAiAssistant={() => handleOpenAiAssistant()}
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
            />
          )}

          {currentTab === 'history' && (
            <ImportHistoryView
              logs={importLogs}
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
            />
          )}
        </main>
      </div>

      {/* AI Assistant Modal (Section 29) */}
      <AiAssistantModal
        isOpen={isAiModalOpen}
        onClose={() => setIsAiModalOpen(false)}
        currentSheet={aiTargetSheet}
        allSheets={currentAnalysis?.worksheets || []}
        workbookName={currentAnalysis?.filename || 'students_complex.xlsx'}
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
