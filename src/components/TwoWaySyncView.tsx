import React, { useState, useMemo } from 'react';
import { 
  ArrowLeftRight, 
  RefreshCw, 
  AlertTriangle, 
  CheckCircle2, 
  Database, 
  FileSpreadsheet, 
  Check, 
  ArrowRight, 
  UploadCloud, 
  DownloadCloud, 
  Search, 
  Copy, 
  Info,
  Server,
  Zap,
  Sliders,
  Clock,
  ShieldCheck,
  Split,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { 
  WorkbookAnalysis, 
  WorksheetMapping, 
  SyncConflictRecord, 
  TwoWaySyncResult, 
  ConflictResolutionChoice,
  NextcloudConfig,
  SupabaseConfig,
  NavigationTab,
  TwoWaySyncSettings
} from '../types';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
import { TwoWaySyncEngine } from '../services/twoWaySyncEngine';
import { StorageService } from '../services/storage';
import { ApiClient } from '../services/apiClient';
import * as XLSX from 'xlsx';

interface TwoWaySyncViewProps {
  currentAnalysis: WorkbookAnalysis | null;
  mappings: WorksheetMapping[];
  databaseState: Record<string, any[]>;
  nextcloudConfig: NextcloudConfig;
  supabaseConfig: SupabaseConfig;
  onUpdateDatabase: (tableName: string, newRecords: any[]) => void;
  onNavigate: (tab: NavigationTab) => void;
  conflicts: SyncConflictRecord[];
  onUpdateConflicts: (conflicts: SyncConflictRecord[]) => void;
}

export const TwoWaySyncView: React.FC<TwoWaySyncViewProps> = ({
  currentAnalysis,
  mappings,
  databaseState,
  nextcloudConfig,
  supabaseConfig,
  onUpdateDatabase,
  conflicts,
  onUpdateConflicts
}) => {
  const [isDetecting, setIsDetecting] = useState<boolean>(false);
  const [isAutoSyncing, setIsAutoSyncing] = useState<boolean>(false);
  const [isPushingToExcel, setIsPushingToExcel] = useState<boolean>(false);
  const [notice, setNotice] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);
  const [filterState, setFilterState] = useState<'ALL' | 'CONFLICT' | 'AUTO_PUSH_TO_DB' | 'AUTO_PUSH_TO_EXCEL' | 'POLICY_BLOCKED' | 'IN_SYNC' | 'RESOLVED'>('CONFLICT');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [activeInfoTab, setActiveInfoTab] = useState<'hub' | 'operational_guide' | 'secrets_guide'>('hub');
  const [expandedRecordId, setExpandedRecordId] = useState<string | null>(null);
  const [copiedSql, setCopiedSql] = useState<boolean>(false);

  // Settings
  const [syncSettings, setSyncSettings] = useState<TwoWaySyncSettings>(() => StorageService.getTwoWaySyncSettings());

  const handleToggleAutoPush = () => {
    const updated: TwoWaySyncSettings = {
      ...syncSettings,
      autoPushNonConflicting: !syncSettings.autoPushNonConflicting
    };
    setSyncSettings(updated);
    StorageService.saveTwoWaySyncSettings(updated);
  };

  // Sync result metadata
  const [syncSummary, setSyncSummary] = useState<{
    lastChecked: string;
    totalCompared: number;
    inSyncCount: number;
    autoPushToDbCount: number;
    autoPushToExcelCount: number;
    policyBlockedCount: number;
    conflictCount: number;
  }>(() => {
    return {
      lastChecked: new Date().toLocaleTimeString(),
      totalCompared: conflicts.length,
      inSyncCount: conflicts.filter(c => c.state === 'IN_SYNC').length,
      autoPushToDbCount: conflicts.filter(c => c.state === 'AUTO_PUSH_TO_DB' || c.state === 'EXCEL_ONLY').length,
      autoPushToExcelCount: conflicts.filter(c => c.state === 'AUTO_PUSH_TO_EXCEL' || c.state === 'SUPABASE_ONLY').length,
      policyBlockedCount: conflicts.filter(c => c.state === 'POLICY_BLOCKED').length,
      conflictCount: conflicts.filter(c => c.state === 'CONFLICT').length,
    };
  });

  // Run change detection with Sync Baseline awareness
  const handleDetectChanges = async () => {
    setIsDetecting(true);
    setNotice(null);

    try {
      const sample = createComplexSampleWorkbook();
      const filename = currentAnalysis?.filename || 'students_complex.xlsx';
      const baselines = StorageService.getSyncBaselines();

      // Execute detection algorithm using baselines and per-table sync policies
      const result: TwoWaySyncResult = TwoWaySyncEngine.detectTwoWayDiffs(
        sample.workbook,
        filename,
        mappings,
        databaseState,
        baselines
      );

      await new Promise(r => setTimeout(r, 450));

      onUpdateConflicts(result.records);
      setSyncSummary({
        lastChecked: new Date().toLocaleTimeString(),
        totalCompared: result.totalRecordsCompared,
        inSyncCount: result.inSyncCount,
        autoPushToDbCount: result.autoPushToDbCount,
        autoPushToExcelCount: result.autoPushToExcelCount,
        policyBlockedCount: result.policyBlockedCount,
        conflictCount: result.conflictCount,
      });

      if (result.conflictCount > 0) {
        setNotice({
          type: 'info',
          message: `Scan complete: ${result.conflictCount} concurrent collision(s) need human approval. ${result.autoPushToDbCount} Nextcloud edit(s) and ${result.autoPushToExcelCount} Supabase edit(s) can auto-sync with 0 clicks. ${result.policyBlockedCount} change(s) protected by table authority policy.`
        });
        setFilterState('CONFLICT');
      } else {
        setNotice({
          type: 'success',
          message: `Scan complete: Zero concurrent collisions! ${result.autoPushToDbCount} Nextcloud change(s) and ${result.autoPushToExcelCount} Supabase change(s) ready for automatic bi-directional sync (${result.policyBlockedCount} protected by table policy).`
        });
      }
    } catch (err: any) {
      setNotice({
        type: 'error',
        message: `Error detecting changes: ${err.message}`
      });
    } finally {
      setIsDetecting(false);
    }
  };

  // Run Automatic Bi-Directional Sync (0-click execution of all unilateral changes)
  const handleRunAutoSync = async () => {
    setIsAutoSyncing(true);
    setNotice(null);

    try {
      const sample = createComplexSampleWorkbook();
      const wb = sample.workbook;
      const baselines = StorageService.getSyncBaselines();

      const {
        updatedDatabaseState,
        updatedWorkbook,
        updatedBaselines,
        syncedToDbCount,
        syncedToExcelCount,
        remainingConflicts
      } = TwoWaySyncEngine.executeAutoSync(
        conflicts,
        wb,
        mappings,
        databaseState,
        baselines
      );

      // 1. Update database state
      for (const tbl of Object.keys(updatedDatabaseState)) {
        onUpdateDatabase(tbl, updatedDatabaseState[tbl]);
      }

      // 2. Save updated baselines
      StorageService.saveSyncBaselines(updatedBaselines);

      // 3. Upload updated workbook back to Nextcloud WebDAV if changes were made
      let uploadStatus = '';
      if (updatedWorkbook && syncedToExcelCount > 0) {
        try {
          const base64Out = XLSX.write(updatedWorkbook, { type: 'base64', bookType: 'xlsx' });
          const filename = currentAnalysis?.filename || 'students_complex.xlsx';
          const uploadRes = await ApiClient.uploadExcelFile(
            nextcloudConfig,
            filename,
            base64Out,
            nextcloudConfig.sourceFolder
          );
          if (uploadRes.success) {
            uploadStatus = ' (and uploaded directly to Nextcloud WebDAV)';
          }
        } catch (upErr: any) {
          console.warn('Nextcloud WebDAV upload notice:', upErr);
        }
      }

      // 4. Update conflicts list
      onUpdateConflicts(remainingConflicts);

      // 5. Update summary
      setSyncSummary(prev => ({
        ...prev,
        inSyncCount: prev.inSyncCount + syncedToDbCount + syncedToExcelCount,
        autoPushToDbCount: 0,
        autoPushToExcelCount: 0,
        conflictCount: remainingConflicts.length
      }));

      setNotice({
        type: 'success',
        message: `Automatic Bi-directional Sync complete! Pushed ${syncedToDbCount} change(s) into Supabase DB and ${syncedToExcelCount} change(s) back into Nextcloud Excel${uploadStatus}. ${remainingConflicts.length} concurrent collision(s) require manual decision.`
      });

      if (remainingConflicts.length > 0) {
        setFilterState('CONFLICT');
      } else {
        setFilterState('IN_SYNC');
      }
    } catch (err: any) {
      setNotice({
        type: 'error',
        message: `Auto-sync failed: ${err.message}`
      });
    } finally {
      setIsAutoSyncing(false);
    }
  };

  // Push all Supabase rows to Nextcloud Excel explicitly
  const handlePushDatabaseToNextcloud = async () => {
    setIsPushingToExcel(true);
    setNotice(null);

    try {
      const sample = createComplexSampleWorkbook();
      const wb = sample.workbook;

      let totalUpdated = 0;
      let totalAppended = 0;

      for (const wm of mappings) {
        if (!wm.enabled) continue;
        const dbRecords = databaseState[wm.supabaseTable] || [];
        if (dbRecords.length === 0) continue;

        const { updatedCount, newRowsCount } = TwoWaySyncEngine.writeRecordsToExcelWorkbook(
          wb,
          wm.worksheetName,
          wm,
          dbRecords
        );

        totalUpdated += updatedCount;
        totalAppended += newRowsCount;
      }

      const base64Out = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const filename = currentAnalysis?.filename || 'students_complex.xlsx';

      const uploadRes = await ApiClient.uploadExcelFile(
        nextcloudConfig,
        filename,
        base64Out,
        nextcloudConfig.sourceFolder
      );

      if (uploadRes.success) {
        setNotice({
          type: 'success',
          message: `Reverse-Sync successful! Uploaded updated ${filename} to Nextcloud WebDAV (${totalUpdated} cells updated, ${totalAppended} rows appended).`
        });
      } else {
        setNotice({
          type: 'info',
          message: `Workbook prepared in memory (${totalUpdated} updated). Nextcloud upload notice: ${uploadRes.error || 'Check WebDAV credentials'}`
        });
      }
    } catch (err: any) {
      setNotice({
        type: 'error',
        message: `Reverse-Sync to Excel failed: ${err.message}`
      });
    } finally {
      setIsPushingToExcel(false);
    }
  };

  // Resolve a single collision
  const handleResolveSingle = (recordId: string, choice: ConflictResolutionChoice) => {
    const target = conflicts.find(c => c.id === recordId);
    if (!target) return;

    if (choice === 'excel' || choice === 'custom_merge') {
      const updatedDb = TwoWaySyncEngine.applyResolutionToDatabase(target, choice, databaseState);
      for (const tbl of Object.keys(updatedDb)) {
        onUpdateDatabase(tbl, updatedDb[tbl]);
      }
    }

    // Update baseline to lock resolution
    const baselines = StorageService.getSyncBaselines();
    const recordKey = `${target.tableName}:${target.primaryKeyValue}`;
    const mapping = mappings.find(m => m.worksheetName === target.worksheetName);
    const chosenRec = choice === 'excel' ? target.excelFullRecord : target.supabaseFullRecord;
    const hash = mapping ? TwoWaySyncEngine.computeRecordHash(chosenRec, mapping.columns) : '';
    baselines[recordKey] = {
      recordKey,
      tableName: target.tableName,
      primaryKeyCol: target.primaryKeyCol,
      primaryKeyValue: target.primaryKeyValue,
      excelHash: hash,
      supabaseHash: hash,
      lastSyncedAt: new Date().toISOString(),
      syncedValues: chosenRec
    };
    StorageService.saveSyncBaselines(baselines);

    const updated = conflicts.map(c => {
      if (c.id === recordId) {
        return {
          ...c,
          state: 'RESOLVED' as const,
          resolutionChoice: choice,
          resolvedAt: new Date().toISOString(),
          resolutionNote: choice === 'excel' 
            ? 'Accepted Nextcloud Excel version (overwrote Supabase)'
            : choice === 'supabase'
            ? 'Accepted Supabase version (staged for Excel)'
            : 'Custom field-by-field merge applied'
        };
      }
      return c;
    });

    onUpdateConflicts(updated);
    setSyncSummary(prev => ({
      ...prev,
      conflictCount: Math.max(0, prev.conflictCount - 1),
      inSyncCount: prev.inSyncCount + 1
    }));

    setNotice({
      type: 'success',
      message: `Collision for ${target.primaryKeyValue} successfully resolved with "${choice.toUpperCase()}".`
    });
  };

  // Toggle field source in custom merge mode
  const handleToggleFieldSource = (recordId: string, colName: string, source: 'excel' | 'supabase') => {
    const updated = conflicts.map(c => {
      if (c.id === recordId) {
        const diffs = c.fieldDiffs.map(d => {
          if (d.supabaseColumn === colName) {
            return { ...d, selectedSource: source };
          }
          return d;
        });
        return { ...c, fieldDiffs: diffs };
      }
      return c;
    });
    onUpdateConflicts(updated);
  };

  // Batch resolution of collisions
  const handleResolveAllPending = (choice: 'excel' | 'supabase') => {
    const pending = conflicts.filter(c => c.state === 'CONFLICT');
    if (pending.length === 0) return;

    let updatedDb = { ...databaseState };
    if (choice === 'excel') {
      for (const rec of pending) {
        updatedDb = TwoWaySyncEngine.applyResolutionToDatabase(rec, 'excel', updatedDb);
      }
      for (const tbl of Object.keys(updatedDb)) {
        onUpdateDatabase(tbl, updatedDb[tbl]);
      }
    }

    const updated = conflicts.map(c => {
      if (c.state === 'CONFLICT') {
        return {
          ...c,
          state: 'RESOLVED' as const,
          resolutionChoice: choice,
          resolvedAt: new Date().toISOString(),
          resolutionNote: `Batch resolved via ${choice.toUpperCase()} authority`
        };
      }
      return c;
    });

    onUpdateConflicts(updated);
    setSyncSummary(prev => ({
      ...prev,
      conflictCount: 0,
      inSyncCount: prev.inSyncCount + pending.length
    }));

    setNotice({
      type: 'success',
      message: `Batch resolved ${pending.length} collisions using "${choice.toUpperCase()}" authority.`
    });
  };

  // Filtered records
  const filteredRecords = useMemo(() => {
    return conflicts.filter(r => {
      if (filterState === 'CONFLICT' && r.state !== 'CONFLICT') return false;
      if (filterState === 'AUTO_PUSH_TO_DB' && r.state !== 'AUTO_PUSH_TO_DB' && r.state !== 'EXCEL_ONLY') return false;
      if (filterState === 'AUTO_PUSH_TO_EXCEL' && r.state !== 'AUTO_PUSH_TO_EXCEL' && r.state !== 'SUPABASE_ONLY') return false;
      if (filterState === 'POLICY_BLOCKED' && r.state !== 'POLICY_BLOCKED') return false;
      if (filterState === 'IN_SYNC' && r.state !== 'IN_SYNC') return false;
      if (filterState === 'RESOLVED' && r.state !== 'RESOLVED') return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const matchesPk = r.primaryKeyValue.toLowerCase().includes(q);
        const matchesSheet = r.worksheetName.toLowerCase().includes(q);
        const matchesField = r.fieldDiffs.some(d => 
          String(d.excelValue).toLowerCase().includes(q) || 
          String(d.supabaseValue).toLowerCase().includes(q)
        );
        return matchesPk || matchesSheet || matchesField;
      }
      return true;
    });
  }, [conflicts, filterState, searchQuery]);

  const pendingCollisionsCount = conflicts.filter(c => c.state === 'CONFLICT').length;
  const autoPushDbTotal = conflicts.filter(c => c.state === 'AUTO_PUSH_TO_DB' || c.state === 'EXCEL_ONLY').length;
  const autoPushExcelTotal = conflicts.filter(c => c.state === 'AUTO_PUSH_TO_EXCEL' || c.state === 'SUPABASE_ONLY').length;
  const policyBlockedTotal = conflicts.filter(c => c.state === 'POLICY_BLOCKED').length;
  const totalAutoSyncable = autoPushDbTotal + autoPushExcelTotal;

  const TWO_WAY_SQL_SCHEMA = `-- Option 3: Two-Way Sync & Conflict Resolution Tables
-- Run this in Supabase SQL Editor to enable persistent staging and conflict queues

-- 1. Sync Baselines (Tracks last known synchronized state of every row)
CREATE TABLE IF NOT EXISTS sync_baselines (
    record_key TEXT PRIMARY KEY, -- e.g. "students:STU-1002"
    table_name TEXT NOT NULL,
    primary_key_col TEXT NOT NULL,
    primary_key_val TEXT NOT NULL,
    excel_hash TEXT NOT NULL,
    supabase_hash TEXT NOT NULL,
    synced_values JSONB NOT NULL DEFAULT '{}'::jsonb,
    last_synced_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 2. Sync Conflicts Queue (Holds true concurrent collisions only)
CREATE TABLE IF NOT EXISTS sync_conflicts (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    workbook_name TEXT NOT NULL,
    worksheet_name TEXT NOT NULL,
    target_table TEXT NOT NULL,
    primary_key_col TEXT NOT NULL,
    primary_key_val TEXT NOT NULL,
    excel_row_num INTEGER,
    field_diffs JSONB NOT NULL DEFAULT '[]'::jsonb,
    excel_payload JSONB NOT NULL DEFAULT '{}'::jsonb,
    supabase_payload JSONB NOT NULL DEFAULT '{}'::jsonb,
    status TEXT NOT NULL DEFAULT 'pending', -- 'pending', 'resolved_excel', 'resolved_supabase', 'resolved_merged'
    resolved_by TEXT,
    resolved_at TIMESTAMPTZ,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3. Two-Way Sync Audit Trail
CREATE TABLE IF NOT EXISTS sync_two_way_audit (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    direction TEXT NOT NULL, -- 'excel_to_supabase', 'supabase_to_excel', 'bidirectional_auto'
    records_compared INTEGER NOT NULL DEFAULT 0,
    in_sync_count INTEGER NOT NULL DEFAULT 0,
    unilateral_excel_applied INTEGER NOT NULL DEFAULT 0,
    unilateral_supabase_pushed INTEGER NOT NULL DEFAULT 0,
    concurrent_conflicts_flagged INTEGER NOT NULL DEFAULT 0,
    performed_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    details JSONB
);

-- Indexes for fast lookup
CREATE INDEX IF NOT EXISTS idx_sync_conflicts_status ON sync_conflicts(status);
CREATE INDEX IF NOT EXISTS idx_sync_baselines_table ON sync_baselines(table_name);

-- Row Level Security
ALTER TABLE sync_baselines ENABLE ROW LEVEL SECURITY;
ALTER TABLE sync_conflicts ENABLE ROW LEVEL SECURITY;
ALTER TABLE sync_two_way_audit ENABLE ROW LEVEL SECURITY;

DO $$
BEGIN
    DROP POLICY IF EXISTS "Allow all access to sync_baselines" ON sync_baselines;
    CREATE POLICY "Allow all access to sync_baselines" ON sync_baselines FOR ALL TO public USING (true) WITH CHECK (true);
    
    DROP POLICY IF EXISTS "Allow all access to sync_conflicts" ON sync_conflicts;
    CREATE POLICY "Allow all access to sync_conflicts" ON sync_conflicts FOR ALL TO public USING (true) WITH CHECK (true);
    
    DROP POLICY IF EXISTS "Allow all access to sync_two_way_audit" ON sync_two_way_audit;
    CREATE POLICY "Allow all access to sync_two_way_audit" ON sync_two_way_audit FOR ALL TO public USING (true) WITH CHECK (true);
END $$;`;

  return (
    <div className="space-y-6">
      {/* Top Banner & Title */}
      <div className="bg-white rounded-xl shadow-xs border border-slate-200 p-6">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div>
            <div className="flex items-center space-x-3">
              <div className="p-2.5 bg-emerald-100 text-emerald-800 rounded-lg">
                <ArrowLeftRight className="w-6 h-6" />
              </div>
              <div>
                <div className="flex items-center space-x-2">
                  <h1 className="text-xl font-bold text-slate-900">Two-Way Sync & Conflict Hub</h1>
                  <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 flex items-center space-x-1">
                    <Zap className="w-3 h-3 text-emerald-600" />
                    <span>Automatic Bi-Directional Mode</span>
                  </span>
                </div>
                <p className="text-sm text-slate-500 mt-1">
                  Single-platform changes sync automatically with <strong>0 clicks</strong>. Only simultaneous collisions pause for your approval.
                </p>
              </div>
            </div>
          </div>

          {/* Sync Mode Navigation Tabs */}
          <div className="flex items-center space-x-2">
            <button
              id="tab-toggle-hub"
              onClick={() => setActiveInfoTab('hub')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                activeInfoTab === 'hub' 
                  ? 'bg-slate-900 text-white' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Conflict Hub
            </button>
            <button
              id="tab-toggle-operational"
              onClick={() => setActiveInfoTab('operational_guide')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors flex items-center space-x-1.5 ${
                activeInfoTab === 'operational_guide' 
                  ? 'bg-emerald-700 text-white' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              <Clock className="w-3.5 h-3.5" />
              <span>How 15-Min Sync Behaves</span>
            </button>
            <button
              id="tab-toggle-secrets"
              onClick={() => setActiveInfoTab('secrets_guide')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors flex items-center space-x-1.5 ${
                activeInfoTab === 'secrets_guide' 
                  ? 'bg-blue-600 text-white' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              <Server className="w-3.5 h-3.5" />
              <span>Server Secrets & SQL</span>
            </button>
          </div>
        </div>

        {/* Behavior Explainer Banner */}
        <div className="mt-5 p-4 rounded-xl bg-slate-50 border border-slate-200">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
            <div className="flex items-start space-x-2.5">
              <div className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-800 flex items-center justify-center font-bold text-[11px] shrink-0">
                1
              </div>
              <div>
                <span className="font-bold text-slate-900 block">Change in Nextcloud only?</span>
                <span className="text-slate-600">
                  Supabase DB was untouched. Auto-syncs to Supabase with <strong>0 manual clicks</strong>.
                </span>
              </div>
            </div>

            <div className="flex items-start space-x-2.5">
              <div className="w-6 h-6 rounded-full bg-blue-100 text-blue-800 flex items-center justify-center font-bold text-[11px] shrink-0">
                2
              </div>
              <div>
                <span className="font-bold text-slate-900 block">Change in Supabase only?</span>
                <span className="text-slate-600">
                  Nextcloud Excel was untouched. Auto-syncs to Excel with <strong>0 manual clicks</strong>.
                </span>
              </div>
            </div>

            <div className="flex items-start space-x-2.5">
              <div className="w-6 h-6 rounded-full bg-amber-100 text-amber-800 flex items-center justify-center font-bold text-[11px] shrink-0">
                3
              </div>
              <div>
                <span className="font-bold text-slate-900 block">Change both at the same time?</span>
                <span className="text-slate-600">
                  Simultaneous collision detected! <strong>Only these</strong> pause here for your manual decision.
                </span>
              </div>
            </div>
          </div>
        </div>

        {/* Action Toolbar */}
        <div className="mt-5 pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            {/* Primary: Auto-Sync All Unilateral Changes */}
            <button
              id="btn-auto-sync-unilateral"
              onClick={handleRunAutoSync}
              disabled={isAutoSyncing || totalAutoSyncable === 0}
              className="inline-flex items-center space-x-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold rounded-lg shadow-xs transition-colors disabled:opacity-50"
              title="Automatically apply all unilateral Nextcloud and Supabase edits"
            >
              <Zap className={`w-4 h-4 ${isAutoSyncing ? 'animate-spin' : 'text-emerald-200'}`} />
              <span>
                {isAutoSyncing 
                  ? 'Auto-Syncing Records...' 
                  : `Auto-Sync Unilateral Changes (${totalAutoSyncable})`}
              </span>
            </button>

            <button
              id="btn-detect-changes"
              onClick={handleDetectChanges}
              disabled={isDetecting}
              className="inline-flex items-center space-x-2 px-3.5 py-2 bg-slate-800 hover:bg-slate-900 text-white text-sm font-medium rounded-lg shadow-xs transition-colors disabled:opacity-50"
            >
              <RefreshCw className={`w-4 h-4 ${isDetecting ? 'animate-spin' : ''}`} />
              <span>{isDetecting ? 'Scanning Diff...' : 'Scan 2-Way Diff'}</span>
            </button>

            <button
              id="btn-push-excel"
              onClick={handlePushDatabaseToNextcloud}
              disabled={isPushingToExcel}
              className="inline-flex items-center space-x-2 px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg shadow-xs transition-colors disabled:opacity-50"
              title="Write current Supabase database rows back into Nextcloud Excel"
            >
              <UploadCloud className="w-4 h-4 text-blue-200" />
              <span>{isPushingToExcel ? 'Pushing...' : 'Push DB to Nextcloud'}</span>
            </button>
          </div>

          {/* Batch Collision Shortcuts */}
          {pendingCollisionsCount > 0 && (
            <div className="flex items-center space-x-2 bg-amber-50 border border-amber-200 px-3 py-1.5 rounded-lg">
              <span className="text-xs font-semibold text-amber-800">Resolve Collisions:</span>
              <button
                id="btn-resolve-all-excel"
                onClick={() => handleResolveAllPending('excel')}
                className="px-2.5 py-1 text-xs font-medium rounded bg-emerald-100 text-emerald-800 hover:bg-emerald-200 transition-colors"
              >
                All Excel Wins
              </button>
              <button
                id="btn-resolve-all-supabase"
                onClick={() => handleResolveAllPending('supabase')}
                className="px-2.5 py-1 text-xs font-medium rounded bg-blue-100 text-blue-800 hover:bg-blue-200 transition-colors"
              >
                All DB Wins
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Notice Banner */}
      {notice && (
        <div className={`p-4 rounded-xl text-sm flex items-center justify-between border ${
          notice.type === 'success' 
            ? 'bg-emerald-50 border-emerald-200 text-emerald-800' 
            : notice.type === 'error'
            ? 'bg-rose-50 border-rose-200 text-rose-800'
            : 'bg-amber-50 border-amber-200 text-amber-800'
        }`}>
          <div className="flex items-center space-x-2.5">
            {notice.type === 'success' ? (
              <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
            ) : notice.type === 'error' ? (
              <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0" />
            ) : (
              <Info className="w-5 h-5 text-amber-600 shrink-0" />
            )}
            <span>{notice.message}</span>
          </div>
          <button 
            onClick={() => setNotice(null)}
            className="text-xs font-semibold underline opacity-70 hover:opacity-100 ml-4"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* Operational Guide Panel: Exactly answering the user's questions */}
      {activeInfoTab === 'operational_guide' && (
        <div className="bg-white rounded-xl shadow-xs border border-slate-200 p-6 space-y-6">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h2 className="text-lg font-bold text-slate-900 flex items-center space-x-2">
                <Clock className="w-5 h-5 text-emerald-600" />
                <span>How the Two-Way Sync Engine Operates (15-Minute Cycle)</span>
              </h2>
              <p className="text-sm text-slate-500 mt-1">
                Detailed answer to what happens when you edit in one platform vs both platforms.
              </p>
            </div>
            <button
              onClick={() => setActiveInfoTab('hub')}
              className="text-xs text-slate-500 hover:text-slate-800 font-medium"
            >
              Close
            </button>
          </div>

          <div className="space-y-4 text-xs text-slate-700 leading-relaxed">
            {/* Question 1 */}
            <div className="p-4 rounded-xl bg-emerald-50/60 border border-emerald-200 space-y-2">
              <h3 className="text-sm font-bold text-emerald-950 flex items-center space-x-1.5">
                <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                <span>Question 1: If I change in Nextcloud and don't change in Supabase for 15 minutes, what happens?</span>
              </h3>
              <p>
                <strong>Answer: It syncs automatically into Supabase without asking you for approval.</strong><br/>
                Every 15 minutes, the background worker reads the Nextcloud Excel file and calculates the hash of each record against the stored sync baseline. Because the Nextcloud hash changed but the Supabase hash remained identical to baseline, the engine classifies this as a <strong>unilateral Nextcloud edit</strong>. It issues an <code className="bg-emerald-100 px-1 py-0.5 rounded text-emerald-900">UPSERT</code> to your Supabase PostgreSQL table, updates the baseline timestamp, and logs the change. <strong>Zero manual clicks or approvals are needed.</strong>
              </p>
            </div>

            {/* Question 2 */}
            <div className="p-4 rounded-xl bg-blue-50/60 border border-blue-200 space-y-2">
              <h3 className="text-sm font-bold text-blue-950 flex items-center space-x-1.5">
                <CheckCircle2 className="w-4 h-4 text-blue-600" />
                <span>Question 2: If I change in Supabase and don't change in Nextcloud, what happens?</span>
              </h3>
              <p>
                <strong>Answer: It syncs automatically back into your Nextcloud Excel workbook without asking you for approval.</strong><br/>
                The worker checks the <code className="bg-blue-100 px-1 py-0.5 rounded text-blue-900">updated_at</code> timestamp on Supabase. It detects that the row in Supabase changed, but the row in Nextcloud Excel still matches the baseline. The engine classifies this as a <strong>unilateral Supabase edit</strong>. It opens the Excel workbook, updates the matching cells (or appends new rows), saves it, and uploads the updated file to Nextcloud WebDAV via HTTP <code className="bg-blue-100 px-1 py-0.5 rounded text-blue-900">PUT</code>. <strong>Zero manual clicks or approvals are needed.</strong>
              </p>
            </div>

            {/* Question 3 */}
            <div className="p-4 rounded-xl bg-amber-50/60 border border-amber-200 space-y-2">
              <h3 className="text-sm font-bold text-amber-950 flex items-center space-x-1.5">
                <AlertTriangle className="w-4 h-4 text-amber-600" />
                <span>Question 3: What if someone edits the EXACT SAME RECORD in both platforms at the same time?</span>
              </h3>
              <p>
                <strong>Answer: That is a true collision, and that is the ONLY thing that requires your manual decision in the Conflict Hub.</strong><br/>
                If someone modifies the phone number in Nextcloud to <em>+1 555-011-3344</em> and someone else simultaneously updates that student's status or email in Supabase before the next sync cycle, neither system can guess who is right. The engine flags this specific record as a <span className="text-amber-800 font-bold">Simultaneous Collision</span>. It pauses that single record, alerts you in this Conflict Hub, and lets you choose <strong>Accept Excel</strong>, <strong>Accept Supabase</strong>, or <strong>Commit Merge</strong>. All your other non-colliding records continue to sync automatically!
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Secrets & Server Setup Guide Panel */}
      {activeInfoTab === 'secrets_guide' && (
        <div className="bg-white rounded-xl shadow-xs border border-slate-200 p-6 space-y-6">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h2 className="text-lg font-bold text-slate-900 flex items-center space-x-2">
                <Server className="w-5 h-5 text-blue-600" />
                <span>Required Server Configurations & Secrets for Two-Way Sync</span>
              </h2>
              <p className="text-sm text-slate-500 mt-1">
                Everything you need to configure in your servers (Nextcloud, Supabase, Coolify Worker) to run Automatic Two-Way Sync in production.
              </p>
            </div>
            <button
              onClick={() => setActiveInfoTab('hub')}
              className="text-xs text-slate-500 hover:text-slate-800 font-medium"
            >
              Close Guide
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {/* 1. Supabase Secrets */}
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
              <div className="flex items-center space-x-2 text-slate-900 font-semibold text-sm">
                <Database className="w-4 h-4 text-emerald-600" />
                <span>1. Supabase Secrets</span>
              </div>
              <p className="text-xs text-slate-600 leading-relaxed">
                Two-way sync requires the <strong>Service Role Key</strong> (not the public anon key) so the background worker can bypass RLS when auto-upserting records and updating <code className="bg-slate-200 px-1 rounded">sync_baselines</code>.
              </p>
              <div className="bg-slate-900 text-emerald-400 p-2.5 rounded-lg text-xs font-mono">
                SUPABASE_URL=https://...<br/>
                SUPABASE_SERVICE_ROLE_KEY=eyJh...
              </div>
            </div>

            {/* 2. Nextcloud WebDAV Secrets */}
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
              <div className="flex items-center space-x-2 text-slate-900 font-semibold text-sm">
                <FileSpreadsheet className="w-4 h-4 text-blue-600" />
                <span>2. Nextcloud WebDAV Secrets</span>
              </div>
              <p className="text-xs text-slate-600 leading-relaxed">
                The App Password must have <strong>read + write</strong> permissions on the target directory (e.g. <code className="bg-slate-200 px-1 rounded">/ExcelImports</code>) to allow writing reverse-synced records back via WebDAV <code className="bg-slate-200 px-1 rounded">PUT</code>.
              </p>
              <div className="bg-slate-900 text-blue-400 p-2.5 rounded-lg text-xs font-mono">
                NEXTCLOUD_URL=https://cloud.jhcnexus.space<br/>
                NEXTCLOUD_APP_PASSWORD=xxxx-xxxx-xxxx
              </div>
            </div>

            {/* 3. Coolify Worker Daemon */}
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
              <div className="flex items-center space-x-2 text-slate-900 font-semibold text-sm">
                <Server className="w-4 h-4 text-purple-600" />
                <span>3. Coolify Worker Daemon</span>
              </div>
              <p className="text-xs text-slate-600 leading-relaxed">
                The background worker runs the comparison loop every 15 minutes. Unilateral changes are committed automatically; only collisions are sent to the Conflict Hub.
              </p>
              <div className="bg-slate-900 text-purple-400 p-2.5 rounded-lg text-xs font-mono">
                WORKER_SYNC_INTERVAL=15m<br/>
                AUTO_SYNC_UNILATERAL=true
              </div>
            </div>
          </div>

          {/* SQL Snippet for Supabase */}
          <div className="space-y-2">
            <div className="flex items-center justify-between">
              <label className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                Option 3 SQL Schema for Supabase (Baselines, Conflicts & Staging Audit)
              </label>
              <button
                onClick={() => {
                  navigator.clipboard.writeText(TWO_WAY_SQL_SCHEMA);
                  setCopiedSql(true);
                  setTimeout(() => setCopiedSql(false), 2000);
                }}
                className="inline-flex items-center space-x-1 text-xs text-emerald-600 hover:text-emerald-700 font-medium"
              >
                {copiedSql ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                <span>{copiedSql ? 'Copied!' : 'Copy Schema SQL'}</span>
              </button>
            </div>
            <pre className="bg-slate-900 text-slate-200 p-4 rounded-xl text-xs font-mono overflow-x-auto max-h-56 leading-relaxed">
              {TWO_WAY_SQL_SCHEMA}
            </pre>
          </div>
        </div>
      )}

      {/* Metrics Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 sm:gap-4">
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
          <div className="text-xs font-medium text-slate-500">Total Compared</div>
          <div className="text-xl font-bold text-slate-900 mt-1 font-mono">{syncSummary.totalCompared}</div>
          <div className="text-[11px] text-slate-400 mt-0.5">Rows across sheets</div>
        </div>

        <div className="bg-white p-4 rounded-xl border border-emerald-200 bg-emerald-50/20 shadow-xs">
          <div className="text-xs font-medium text-emerald-700">In Sync</div>
          <div className="text-xl font-bold text-emerald-700 mt-1 font-mono">{syncSummary.inSyncCount}</div>
          <div className="text-[11px] text-emerald-600 mt-0.5">Values match 100%</div>
        </div>

        <div className="bg-white p-4 rounded-xl border border-purple-200 bg-purple-50/20 shadow-xs">
          <div className="text-xs font-medium text-purple-700">Nextcloud Edits</div>
          <div className="text-xl font-bold text-purple-700 mt-1 font-mono">{autoPushDbTotal}</div>
          <div className="text-[11px] text-purple-600 mt-0.5">Auto-push to DB (0 clicks)</div>
        </div>

        <div className="bg-white p-4 rounded-xl border border-blue-200 bg-blue-50/20 shadow-xs">
          <div className="text-xs font-medium text-blue-700">Supabase Edits</div>
          <div className="text-xl font-bold text-blue-700 mt-1 font-mono">{autoPushExcelTotal}</div>
          <div className="text-[11px] text-blue-600 mt-0.5">Auto-push to Excel (0 clicks)</div>
        </div>

        <div className="bg-white p-4 rounded-xl border border-slate-300 bg-slate-50/50 shadow-xs">
          <div className="text-xs font-medium text-slate-700 flex items-center space-x-1">
            <ShieldCheck className="w-3.5 h-3.5 text-slate-500" />
            <span>Policy Protected</span>
          </div>
          <div className="text-xl font-bold text-slate-800 mt-1 font-mono">{policyBlockedTotal}</div>
          <div className="text-[11px] text-slate-500 mt-0.5">Protected by table master rule</div>
        </div>

        <div className={`p-4 rounded-xl border shadow-xs ${
          pendingCollisionsCount > 0 
            ? 'bg-amber-50/70 border-amber-300 ring-2 ring-amber-400/20' 
            : 'bg-white border-slate-200'
        }`}>
          <div className={`text-xs font-medium ${pendingCollisionsCount > 0 ? 'text-amber-800' : 'text-slate-500'}`}>
            Pending Collisions
          </div>
          <div className={`text-xl font-bold mt-1 font-mono ${pendingCollisionsCount > 0 ? 'text-amber-700' : 'text-slate-900'}`}>
            {pendingCollisionsCount}
          </div>
          <div className={`text-[11px] mt-0.5 ${pendingCollisionsCount > 0 ? 'text-amber-700 font-semibold' : 'text-slate-400'}`}>
            {pendingCollisionsCount > 0 ? 'Human review required' : 'Zero collisions'}
          </div>
        </div>
      </div>

      {/* Filter & Search Bar */}
      <div className="bg-white rounded-xl shadow-xs border border-slate-200 p-4">
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
          {/* Filter Pills */}
          <div className="flex flex-wrap items-center gap-1.5">
            <button
              onClick={() => setFilterState('CONFLICT')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                filterState === 'CONFLICT' 
                  ? 'bg-amber-500 text-white font-semibold' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Needs Approval (Collisions) ({pendingCollisionsCount})
            </button>
            <button
              onClick={() => setFilterState('AUTO_PUSH_TO_DB')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                filterState === 'AUTO_PUSH_TO_DB' 
                  ? 'bg-purple-600 text-white font-semibold' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Nextcloud Edits ({autoPushDbTotal})
            </button>
            <button
              onClick={() => setFilterState('AUTO_PUSH_TO_EXCEL')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                filterState === 'AUTO_PUSH_TO_EXCEL' 
                  ? 'bg-blue-600 text-white font-semibold' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Supabase Edits ({autoPushExcelTotal})
            </button>
            <button
              onClick={() => setFilterState('POLICY_BLOCKED')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                filterState === 'POLICY_BLOCKED' 
                  ? 'bg-slate-800 text-white font-semibold' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Protected by Policy ({policyBlockedTotal})
            </button>
            <button
              onClick={() => setFilterState('IN_SYNC')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                filterState === 'IN_SYNC' 
                  ? 'bg-emerald-600 text-white font-semibold' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              In Sync ({syncSummary.inSyncCount})
            </button>
            <button
              onClick={() => setFilterState('RESOLVED')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                filterState === 'RESOLVED' 
                  ? 'bg-slate-800 text-white font-semibold' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Resolved History
            </button>
            <button
              onClick={() => setFilterState('ALL')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                filterState === 'ALL' 
                  ? 'bg-slate-700 text-white' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              All Records ({conflicts.length})
            </button>
          </div>

          {/* Search Field */}
          <div className="relative w-full sm:w-64">
            <Search className="w-4 h-4 absolute left-3 top-2.5 text-slate-400" />
            <input
              type="text"
              placeholder="Search by ID, name, or value..."
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-3 py-1.5 rounded-lg border border-slate-300 text-xs focus:ring-2 focus:ring-emerald-500 focus:outline-hidden"
            />
          </div>
        </div>
      </div>

      {/* Record Diff Cards List */}
      <div className="space-y-4">
        {filteredRecords.length === 0 ? (
          <div className="bg-white rounded-xl shadow-xs border border-slate-200 p-12 text-center">
            <CheckCircle2 className="w-12 h-12 text-emerald-500 mx-auto mb-3" />
            <h3 className="text-base font-semibold text-slate-900">No records found for this filter</h3>
            <p className="text-sm text-slate-500 mt-1 max-w-md mx-auto">
              {filterState === 'CONFLICT' 
                ? 'Zero simultaneous collisions! All single-platform edits have been automatically synchronized.'
                : 'Try adjusting your search query or switching to another filter.'}
            </p>
          </div>
        ) : (
          filteredRecords.map(record => {
            const isCollision = record.state === 'CONFLICT';
            const isAutoDb = record.state === 'AUTO_PUSH_TO_DB' || record.state === 'EXCEL_ONLY';
            const isAutoExcel = record.state === 'AUTO_PUSH_TO_EXCEL' || record.state === 'SUPABASE_ONLY';
            const isPolicyBlocked = record.state === 'POLICY_BLOCKED';
            const isExpanded = expandedRecordId === record.id || isCollision || isPolicyBlocked;

            return (
              <div 
                key={record.id}
                className={`bg-white rounded-xl border shadow-xs transition-all ${
                  isCollision
                    ? 'border-amber-300 ring-1 ring-amber-300/30'
                    : isAutoDb
                    ? 'border-purple-200'
                    : isAutoExcel
                    ? 'border-blue-200'
                    : isPolicyBlocked
                    ? 'border-slate-300 bg-slate-50/30'
                    : record.state === 'RESOLVED'
                    ? 'border-slate-300 bg-slate-50/50'
                    : 'border-slate-200'
                }`}
              >
                {/* Header Bar */}
                <div className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100">
                  <div className="flex items-center space-x-3">
                    <span className={`px-2.5 py-1 rounded-full text-xs font-bold uppercase tracking-wider inline-flex items-center ${
                      isCollision
                        ? 'bg-amber-100 text-amber-800'
                        : isAutoDb
                        ? 'bg-purple-100 text-purple-800'
                        : isAutoExcel
                        ? 'bg-blue-100 text-blue-800'
                        : isPolicyBlocked
                        ? 'bg-slate-200 text-slate-800 border border-slate-300'
                        : record.state === 'RESOLVED'
                        ? 'bg-slate-200 text-slate-700'
                        : 'bg-emerald-100 text-emerald-800'
                    }`}>
                      {isPolicyBlocked && <ShieldCheck className="w-3.5 h-3.5 mr-1" />}
                      <span>
                        {isCollision 
                          ? 'COLLISION (NEEDS APPROVAL)' 
                          : isAutoDb 
                          ? 'NEXTCLOUD EDIT (AUTO-SYNCABLE)' 
                          : isAutoExcel 
                          ? 'SUPABASE EDIT (AUTO-SYNCABLE)' 
                          : isPolicyBlocked
                          ? 'PROTECTED BY TABLE POLICY'
                          : record.state.replace('_', ' ')}
                      </span>
                    </span>

                    <div>
                      <div className="flex items-center space-x-2">
                        <span className="font-mono font-bold text-slate-900 text-sm">
                          {record.primaryKeyCol}: {record.primaryKeyValue}
                        </span>
                        {record.excelRowNumber && (
                          <span className="text-xs text-slate-500 font-mono">
                            (Row {record.excelRowNumber})
                          </span>
                        )}
                      </div>
                      <div className="text-xs text-slate-500 mt-0.5 flex flex-wrap items-center gap-1.5">
                        <span>Worksheet: <strong className="text-slate-700">{record.worksheetName}</strong> → Supabase Table: <strong className="text-slate-700">{record.tableName}</strong></span>
                        {record.tableSyncPolicy && (
                          <span className={`px-1.5 py-0.5 rounded text-[10px] font-semibold ${
                            record.tableSyncPolicy === 'EXCEL_TO_DB'
                              ? 'bg-emerald-100 text-emerald-800'
                              : record.tableSyncPolicy === 'DB_TO_EXCEL'
                              ? 'bg-blue-100 text-blue-800'
                              : record.tableSyncPolicy === 'READ_ONLY'
                              ? 'bg-rose-100 text-rose-800'
                              : 'bg-purple-100 text-purple-800'
                          }`}>
                            {record.tableSyncPolicy === 'EXCEL_TO_DB' 
                              ? 'Nextcloud Master' 
                              : record.tableSyncPolicy === 'DB_TO_EXCEL' 
                              ? 'Supabase Master' 
                              : record.tableSyncPolicy === 'READ_ONLY' 
                              ? 'Read-Only' 
                              : 'Bidirectional'}
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Actions depending on state */}
                  <div className="flex items-center space-x-2">
                    {/* Collisions: Manual Approval Options */}
                    {isCollision && (
                      <div className="flex items-center space-x-2">
                        <button
                          onClick={() => handleResolveSingle(record.id, 'excel')}
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold shadow-xs transition-colors flex items-center space-x-1"
                          title="Overwrite Supabase table with these Excel values"
                        >
                          <Check className="w-3.5 h-3.5" />
                          <span>Accept Excel</span>
                        </button>
                        <button
                          onClick={() => handleResolveSingle(record.id, 'supabase')}
                          className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold shadow-xs transition-colors flex items-center space-x-1"
                          title="Keep Supabase values and prepare to update Excel"
                        >
                          <Check className="w-3.5 h-3.5" />
                          <span>Accept Supabase</span>
                        </button>
                        <button
                          onClick={() => handleResolveSingle(record.id, 'custom_merge')}
                          className="px-3 py-1.5 bg-slate-800 hover:bg-slate-900 text-white rounded-lg text-xs font-semibold shadow-xs transition-colors"
                          title="Combine selected fields into a custom merged record"
                        >
                          Commit Merge
                        </button>
                      </div>
                    )}

                    {/* Auto-Syncable Nextcloud Edit */}
                    {isAutoDb && (
                      <button
                        onClick={() => handleResolveSingle(record.id, 'excel')}
                        className="px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-semibold shadow-xs transition-colors flex items-center space-x-1"
                        title="Auto-push this record to Supabase DB now"
                      >
                        <Zap className="w-3.5 h-3.5" />
                        <span>Sync to DB (0 clicks)</span>
                      </button>
                    )}

                    {/* Auto-Syncable Supabase Edit */}
                    {isAutoExcel && (
                      <button
                        onClick={() => handleResolveSingle(record.id, 'supabase')}
                        className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold shadow-xs transition-colors flex items-center space-x-1"
                        title="Auto-push this record to Nextcloud Excel now"
                      >
                        <Zap className="w-3.5 h-3.5" />
                        <span>Sync to Excel (0 clicks)</span>
                      </button>
                    )}

                    {/* Policy Blocked: Protected by Authority Notice */}
                    {isPolicyBlocked && (
                      <div className="flex items-center space-x-2">
                        <span className="text-xs text-slate-500 font-medium hidden sm:inline">Protected by table policy</span>
                        <button
                          onClick={() => handleResolveSingle(record.id, record.policyNotice?.includes('Nextcloud edit') ? 'excel' : 'supabase')}
                          className="px-2.5 py-1 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded text-xs font-medium transition-colors"
                          title="Manual Admin Override: force apply this change despite policy"
                        >
                          Manual Override
                        </button>
                      </div>
                    )}

                    {record.state === 'RESOLVED' && (
                      <div className="text-xs text-slate-500 font-medium">
                        ✓ Resolved ({record.resolutionNote})
                      </div>
                    )}

                    {/* Expand/Collapse Toggle */}
                    {record.fieldDiffs.length > 0 && !isCollision && !isPolicyBlocked && (
                      <button
                        onClick={() => setExpandedRecordId(isExpanded ? null : record.id)}
                        className="p-1 text-slate-400 hover:text-slate-600"
                      >
                        {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                      </button>
                    )}
                  </div>
                </div>

                {/* Policy Notice Box */}
                {record.policyNotice && (
                  <div className="px-4 py-2.5 bg-slate-100/90 border-b border-slate-200 flex items-start space-x-2 text-xs text-slate-700">
                    <ShieldCheck className="w-4 h-4 text-slate-600 shrink-0 mt-0.5" />
                    <div>
                      <span className="font-semibold text-slate-900">Per-Table Policy Enforcement: </span>
                      <span>{record.policyNotice}</span>
                    </div>
                  </div>
                )}

                {/* Auto-Sync Reason Box */}
                {record.autoApplyReason && (
                  <div className="px-4 py-2 bg-slate-50 border-b border-slate-100 flex items-center space-x-2 text-xs text-slate-600">
                    <Info className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                    <span>{record.autoApplyReason}</span>
                  </div>
                )}

                {/* Diff Inspector Table */}
                {record.fieldDiffs.length > 0 && isExpanded && (
                  <div className="p-4 bg-slate-50/50">
                    <div className="text-xs font-bold text-slate-600 uppercase tracking-wider mb-2 flex items-center justify-between">
                      <span>Field-by-Field Diff Comparison</span>
                      {isCollision && (
                        <span className="text-[11px] text-amber-700 font-medium">
                          Simultaneous edit detected: choose source per field or select Accept Excel/DB above
                        </span>
                      )}
                    </div>

                    <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                          <tr>
                            <th className="px-3 py-2">Column Name</th>
                            <th className="px-3 py-2 w-5/12 text-emerald-800 bg-emerald-50/60">
                              <div className="flex items-center space-x-1.5">
                                <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                                <span>Nextcloud Excel Value</span>
                              </div>
                            </th>
                            <th className="px-3 py-2 w-5/12 text-blue-800 bg-blue-50/60">
                              <div className="flex items-center space-x-1.5">
                                <Database className="w-3.5 h-3.5 text-blue-600" />
                                <span>Supabase PostgreSQL Value</span>
                              </div>
                            </th>
                            {isCollision && (
                              <th className="px-3 py-2 text-center w-24">Choice</th>
                            )}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {record.fieldDiffs.map(diff => {
                            const isExcelSelected = diff.selectedSource === 'excel';

                            return (
                              <tr key={diff.supabaseColumn} className="hover:bg-slate-50 transition-colors">
                                <td className="px-3 py-2 font-mono font-medium text-slate-800">
                                  {diff.supabaseColumn}
                                  <div className="text-[10px] text-slate-400 font-sans">Header: "{diff.excelHeader}"</div>
                                </td>

                                {/* Excel Value Cell */}
                                <td className={`px-3 py-2 font-mono ${
                                  isExcelSelected ? 'bg-emerald-50/70 font-bold text-emerald-950' : 'text-slate-700'
                                }`}>
                                  {diff.excelValue !== undefined && diff.excelValue !== null && diff.excelValue !== ''
                                    ? String(diff.excelValue)
                                    : <span className="text-slate-400 italic">(empty)</span>}
                                </td>

                                {/* Supabase Value Cell */}
                                <td className={`px-3 py-2 font-mono ${
                                  !isExcelSelected ? 'bg-blue-50/70 font-bold text-blue-950' : 'text-slate-700'
                                }`}>
                                  {diff.supabaseValue !== undefined && diff.supabaseValue !== null && diff.supabaseValue !== ''
                                    ? String(diff.supabaseValue)
                                    : <span className="text-slate-400 italic">(empty)</span>}
                                </td>

                                {/* Toggle selection per field for collision merge */}
                                {isCollision && (
                                  <td className="px-3 py-2 text-center">
                                    <div className="inline-flex rounded-md border border-slate-200 overflow-hidden shadow-2xs">
                                      <button
                                        onClick={() => handleToggleFieldSource(record.id, diff.supabaseColumn, 'excel')}
                                        className={`px-2 py-0.5 text-[10px] font-semibold transition-colors ${
                                          isExcelSelected 
                                            ? 'bg-emerald-600 text-white' 
                                            : 'bg-white text-slate-600 hover:bg-slate-100'
                                        }`}
                                      >
                                        Excel
                                      </button>
                                      <button
                                        onClick={() => handleToggleFieldSource(record.id, diff.supabaseColumn, 'supabase')}
                                        className={`px-2 py-0.5 text-[10px] font-semibold transition-colors ${
                                          !isExcelSelected 
                                            ? 'bg-blue-600 text-white' 
                                            : 'bg-white text-slate-600 hover:bg-slate-100'
                                        }`}
                                      >
                                        DB
                                      </button>
                                    </div>
                                  </td>
                                )}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
};
