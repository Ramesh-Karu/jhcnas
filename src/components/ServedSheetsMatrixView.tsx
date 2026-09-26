import React, { useState, useEffect, useMemo } from 'react';
import { 
  Layers, 
  Database, 
  FileSpreadsheet, 
  GitFork, 
  Sparkles, 
  Clock, 
  CheckCircle2, 
  AlertTriangle, 
  XCircle, 
  Search, 
  Filter, 
  RefreshCw, 
  ArrowRight, 
  Play, 
  Eye, 
  ExternalLink, 
  Download, 
  Copy, 
  Check, 
  ArrowDownToLine, 
  Tag, 
  ShieldCheck, 
  ChevronDown, 
  ChevronRight, 
  Hash, 
  Table, 
  Calendar, 
  BookOpen,
  Activity,
  Key,
  Info,
  Folder,
  GraduationCap
} from 'lucide-react';
import { 
  WorksheetMapping, 
  WorkbookAnalysis, 
  SupabaseConfig, 
  NextcloudConfig, 
  NextcloudFile, 
  ImportLog, 
  NavigationTab,
  AiWorkbookPreset,
  SupabaseTableInfo,
  TableSyncPolicy
} from '../types';
import { ApiClient } from '../services/apiClient';
import { AiPresetsModal } from './AiPresetsModal';

interface ServedSheetsMatrixViewProps {
  mappings: WorksheetMapping[];
  currentAnalysis: WorkbookAnalysis | null;
  files?: NextcloudFile[];
  importLogs?: ImportLog[];
  supabaseConfig?: SupabaseConfig;
  nextcloudConfig?: NextcloudConfig;
  onNavigate: (tab: NavigationTab) => void;
  onSaveMappings: (mappings: WorksheetMapping[]) => void;
  onTriggerSync?: () => Promise<void>;
}

type MatrixViewMode = 'SHEETS_MATRIX' | 'MERGED_TABLES' | 'UNIFIED_HISTORY';

export const ServedSheetsMatrixView: React.FC<ServedSheetsMatrixViewProps> = ({
  mappings,
  currentAnalysis,
  files = [],
  importLogs = [],
  supabaseConfig,
  nextcloudConfig,
  onNavigate,
  onSaveMappings,
  onTriggerSync,
}) => {
  const [viewMode, setViewMode] = useState<MatrixViewMode>('SHEETS_MATRIX');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [tableFilter, setTableFilter] = useState<string>('ALL');
  const [selectedWorkbookFilter, setSelectedWorkbookFilter] = useState<string>('ALL');
  const [presetFilter, setPresetFilter] = useState<string>('ALL');
  const [policyFilter, setPolicyFilter] = useState<string>('ALL');
  const [expandedSheetId, setExpandedSheetId] = useState<string | null>(null);
  
  // Live Presets and Supabase Schema data
  const [presets, setPresets] = useState<AiWorkbookPreset[]>([]);
  const [supabaseTables, setSupabaseTables] = useState<SupabaseTableInfo[]>([]);
  const [storedServedSheets, setStoredServedSheets] = useState<any[]>([]);
  const [isLoadingMetadata, setIsLoadingMetadata] = useState<boolean>(false);
  const [permanentSavedAt, setPermanentSavedAt] = useState<string | null>(null);
  const [copyFeedback, setCopyFeedback] = useState<string | null>(null);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const [isPullingSupabase, setIsPullingSupabase] = useState<boolean>(false);
  const [syncFeedback, setSyncFeedback] = useState<{ type: 'success' | 'error'; message: string } | null>(null);

  const [isSavingPermanent, setIsSavingPermanent] = useState<boolean>(false);
  const [isRefreshingNextcloud, setIsRefreshingNextcloud] = useState<boolean>(false);
  const [showAiPresetsModal, setShowAiPresetsModal] = useState<boolean>(false);

  const handleApplyPreset = (preset: AiWorkbookPreset) => {
    if (preset.sheetMappings && preset.sheetMappings.length > 0) {
      const getMappingKey = (m: WorksheetMapping) => {
        const wb = String(m.workbookName || '').toLowerCase().trim();
        const ws = String(m.worksheetName || '').toLowerCase().trim();
        if (wb && ws) return `${wb}::${ws}`;
        if (m.id) return String(m.id);
        return `unspecified::${ws}`;
      };
      const mapByKey = new Map<string, WorksheetMapping>();
      mappings.forEach(m => mapByKey.set(getMappingKey(m), m));
      preset.sheetMappings.forEach(m => mapByKey.set(getMappingKey(m), m));
      const merged = Array.from(mapByKey.values());
      onSaveMappings(merged);
      if (preset.filenamePattern) {
        setSelectedWorkbookFilter(preset.filenamePattern);
      }
      setSyncFeedback({
        type: 'success',
        message: `✨ Applied preset '${preset.name}'! (${preset.sheetMappings.length} worksheets configured into combined table public.${preset.sheetMappings[0]?.supabaseTable}).`
      });
      setTimeout(() => setSyncFeedback(null), 5000);
    }
  };

  const handleApplyAllStudentPresets = (studentPresets: AiWorkbookPreset[]) => {
    const allSheetMappings: WorksheetMapping[] = [];
    studentPresets.forEach(p => {
      if (p.sheetMappings) {
        allSheetMappings.push(...p.sheetMappings);
      }
    });
    if (allSheetMappings.length > 0) {
      const getMappingKey = (m: WorksheetMapping) => {
        const wb = String(m.workbookName || '').toLowerCase().trim();
        const ws = String(m.worksheetName || '').toLowerCase().trim();
        if (wb && ws) return `${wb}::${ws}`;
        if (m.id) return String(m.id);
        return `unspecified::${ws}`;
      };
      const mapByKey = new Map<string, WorksheetMapping>();
      mappings.forEach(m => mapByKey.set(getMappingKey(m), m));
      allSheetMappings.forEach(m => mapByKey.set(getMappingKey(m), m));
      const merged = Array.from(mapByKey.values());
      onSaveMappings(merged);
      setSelectedWorkbookFilter('ALL');
      setSyncFeedback({
        type: 'success',
        message: `🎉 Successfully applied all 8 JHC Student Batch Presets! (${allSheetMappings.length} worksheets configured into combined tables).`
      });
      setTimeout(() => setSyncFeedback(null), 5000);
    }
  };

  // Fetch presets, permanent mapping metadata, and live Supabase tables
  const loadAllMetadata = async () => {
    setIsLoadingMetadata(true);
    try {
      const [presetsRes, permRes, schemaRes, servedRes] = await Promise.allSettled([
        ApiClient.getPresets(),
        ApiClient.loadPermanentMappings(),
        supabaseConfig?.url ? ApiClient.getSupabaseSchema(supabaseConfig) : Promise.resolve({ success: false }),
        ApiClient.loadServedSheets()
      ]);

      let pulledMappings: WorksheetMapping[] = [];

      if (presetsRes.status === 'fulfilled' && presetsRes.value.success && presetsRes.value.presets) {
        setPresets(presetsRes.value.presets);
      }
      if (permRes.status === 'fulfilled' && permRes.value.success) {
        if (permRes.value.savedAt) setPermanentSavedAt(permRes.value.savedAt);
        if (Array.isArray(permRes.value.mappings) && permRes.value.mappings.length > 0) {
          pulledMappings = permRes.value.mappings;
        }
      }
      if (servedRes.status === 'fulfilled' && servedRes.value.success) {
        if (servedRes.value.savedAt) setPermanentSavedAt(servedRes.value.savedAt);
        if (Array.isArray(servedRes.value.servedSheets)) {
          setStoredServedSheets(servedRes.value.servedSheets);
        }
        if (Array.isArray(servedRes.value.mappings) && servedRes.value.mappings.length > 0) {
          pulledMappings = [...pulledMappings, ...servedRes.value.mappings];
        }
      }
      if (schemaRes.status === 'fulfilled' && (schemaRes.value as any).success && (schemaRes.value as any).tables) {
        setSupabaseTables((schemaRes.value as any).tables);
      }

      // If remote/disk has mappings, merge non-destructively so all past mappings are available
      if (pulledMappings.length > 0) {
        const mapByKey = new Map<string, WorksheetMapping>();
        mappings.forEach(m => {
          const key = `${(m.workbookName || '').toLowerCase().trim()}::${(m.worksheetName || '').toLowerCase().trim()}`;
          mapByKey.set(key, m);
        });
        pulledMappings.forEach(m => {
          const key = `${(m.workbookName || '').toLowerCase().trim()}::${(m.worksheetName || '').toLowerCase().trim()}`;
          if (!mapByKey.has(key)) {
            mapByKey.set(key, m);
          }
        });
        const merged = Array.from(mapByKey.values());
        if (merged.length > mappings.length) {
          onSaveMappings(merged);
        }
      }
    } catch (e) {
      console.warn('Metadata loading notice:', e);
    } finally {
      setIsLoadingMetadata(false);
    }
  };

  const handlePullFromSupabase = async () => {
    setIsPullingSupabase(true);
    try {
      const res = await ApiClient.loadPermanentMappings();
      const servedRes = await ApiClient.loadServedSheets();
      
      const combined: WorksheetMapping[] = [
        ...(Array.isArray(res.mappings) ? res.mappings : []),
        ...(Array.isArray(servedRes.mappings) ? servedRes.mappings : [])
      ];

      if (combined.length > 0) {
        const mapByKey = new Map<string, WorksheetMapping>();
        mappings.forEach(m => {
          const key = `${(m.workbookName || '').toLowerCase().trim()}::${(m.worksheetName || '').toLowerCase().trim()}`;
          mapByKey.set(key, m);
        });
        combined.forEach(m => {
          const key = `${(m.workbookName || '').toLowerCase().trim()}::${(m.worksheetName || '').toLowerCase().trim()}`;
          mapByKey.set(key, m);
        });
        const merged = Array.from(mapByKey.values());
        onSaveMappings(merged);
        if (Array.isArray(servedRes.servedSheets)) {
          setStoredServedSheets(servedRes.servedSheets);
        }
        setCopyFeedback(`📥 Pulled ${combined.length} mappings from Supabase PostgreSQL & Server Disk! Total ${merged.length} active sheets.`);
      } else {
        setCopyFeedback('No additional remote mappings found in Supabase.');
      }
    } catch (err: any) {
      setSyncFeedback({ type: 'error', message: `Pull failed: ${err.message}` });
    } finally {
      setIsPullingSupabase(false);
      setTimeout(() => setCopyFeedback(null), 4000);
    }
  };

  const handleSaveServedSheetsPermanently = async () => {
    setIsSavingPermanent(true);
    try {
      const res = await ApiClient.saveServedSheets({
        servedSheets: servedSheets.map(s => ({
          sheetName: s.sheetName,
          workbookName: s.workbookName,
          targetTable: s.targetTable,
          syncPolicy: s.syncPolicy,
          primaryMergeKey: s.primaryMergeKey,
          skipMergedYearRows: s.skipMergedYearRows,
          isEnabled: s.isEnabled,
          columnsCount: s.columnsCount,
          mapping: s.mapping,
        })),
        presets,
        mappings,
        workbookInfo: currentAnalysis ? {
          filename: currentAnalysis.filename,
          fileHash: currentAnalysis.fileHash,
          totalWorksheets: currentAnalysis.worksheets?.length || 0,
        } : null,
      });

      if (supabaseConfig?.url && (supabaseConfig.serviceKey || supabaseConfig.serviceRoleKey || supabaseConfig.anonKey)) {
        await ApiClient.syncMappingsToSupabase({
          mappings,
          workbookInfo: currentAnalysis ? {
            filename: currentAnalysis.filename,
            fileHash: currentAnalysis.fileHash,
            totalWorksheets: currentAnalysis.worksheets?.length || 0,
          } : undefined,
          supabase: supabaseConfig
        }).catch(() => null);
      }

      if (res.success) {
        setPermanentSavedAt(res.savedAt || new Date().toISOString());
        setCopyFeedback('🛡️ Served sheets, mapping sheets & mapping hubs permanently secured in Supabase PostgreSQL & server storage!');
      } else {
        throw new Error(res.error || 'Failed to save served sheets');
      }
    } catch (err: any) {
      setSyncFeedback({
        type: 'error',
        message: `Permanent save error: ${err.message}`,
      });
    } finally {
      setIsSavingPermanent(false);
      setTimeout(() => setCopyFeedback(null), 4000);
    }
  };

  useEffect(() => {
    loadAllMetadata();
  }, [supabaseConfig]);

  // Aggregate Served Sheet Items combining ALL workbook mappings, stored served sheets, and current workbook analysis
  const servedSheets = useMemo(() => {
    const sheetMap = new Map<string, {
      id: string;
      sheetName: string;
      workbookName: string;
      mapping: WorksheetMapping | null;
      analysis: {
        totalRows: number;
        totalColumns: number;
        headerRow: number;
        dataStartRow: number;
        sampleCount: number;
        mergedRangesCount: number;
      } | null;
      targetTable: string;
      presetMatch: AiWorkbookPreset | null;
      primaryMergeKey: string | null;
      syncPolicy: TableSyncPolicy;
      isEnabled: boolean;
      skipMergedYearRows: boolean;
      columnsCount: number;
      lastSyncInfo: {
        timestamp: string | null;
        status: string;
        inserted: number;
        updated: number;
        failed: number;
        rowsCount: number;
      } | null;
    }>();

    // Helper to find last sync info in import logs
    const findLastSync = (wbName: string, wsName: string, tblName: string, fallbackTotalRows?: number) => {
      for (const log of importLogs) {
        const detailResults = log.details?.syncResults as any[];
        if (detailResults && Array.isArray(detailResults)) {
          const sheetRes = detailResults.find((sr: any) => 
            (sr.sheetName && sr.sheetName.toLowerCase() === wsName.toLowerCase()) ||
            (sr.targetTable && sr.targetTable.toLowerCase() === tblName.toLowerCase())
          );
          if (sheetRes) {
            return {
              timestamp: log.completedAt || log.startedAt,
              status: sheetRes.status || log.status,
              inserted: sheetRes.insertedCount ?? 0,
              updated: sheetRes.updatedCount ?? 0,
              failed: sheetRes.failedCount ?? 0,
              rowsCount: sheetRes.rowsCount ?? (fallbackTotalRows || 0),
            };
          }
        }
        if (log.filename && (log.filename.toLowerCase() === wbName.toLowerCase() || log.filename.includes(wsName.toLowerCase()))) {
          return {
            timestamp: log.completedAt || log.startedAt,
            status: log.status,
            inserted: log.rowsInserted,
            updated: log.rowsUpdated,
            failed: log.rowsFailed,
            rowsCount: log.rowsProcessed,
          };
        }
      }
      return null;
    };

    // 1. Add all mappings from state (covers ALL workbooks: past, present, and remote)
    mappings.forEach(m => {
      const wbName = m.workbookName || currentAnalysis?.filename || 'Workbook.xlsx';
      const wsName = m.worksheetName;
      if (!wsName) return;
      const key = `${wbName.toLowerCase().trim()}::${wsName.toLowerCase().trim()}`;
      const targetTable = m.supabaseTable || wsName.toLowerCase().replace(/[^a-z0-9_]/g, '_');
      const pkey = m.columns?.find(c => c.uniqueKey)?.supabaseColumn || null;
      const matchedPreset = presets.find(p => 
        (p.filenamePattern && wbName.toLowerCase().includes(p.filenamePattern.toLowerCase())) ||
        p.sheetMappings?.some(sm => sm.workbookName?.toLowerCase() === wbName.toLowerCase() && sm.worksheetName.toLowerCase() === wsName.toLowerCase())
      ) || presets.find(p =>
        p.sheetMappings?.some(sm => sm.supabaseTable === targetTable)
      ) || presets.find(p => 
        p.sheetMappings?.some(sm => sm.worksheetName.toLowerCase() === wsName.toLowerCase())
      ) || null;

      const lastSync = findLastSync(wbName, wsName, targetTable);

      sheetMap.set(key, {
        id: m.id || `wm-${key}`,
        sheetName: wsName,
        workbookName: wbName,
        mapping: m,
        analysis: null,
        targetTable,
        presetMatch: matchedPreset,
        primaryMergeKey: pkey,
        syncPolicy: m.syncPolicy || 'BIDIRECTIONAL',
        isEnabled: m.enabled !== false,
        skipMergedYearRows: m.skipMergedYearRows !== false,
        columnsCount: m.columns?.length || 0,
        lastSyncInfo: lastSync,
      });
    });

    // 2. Add or enrich with active workbook analysis if loaded
    if (currentAnalysis && currentAnalysis.worksheets) {
      const wbName = currentAnalysis.filename || 'Workbook.xlsx';
      currentAnalysis.worksheets.forEach((ws, idx) => {
        const key = `${wbName.toLowerCase().trim()}::${ws.sheetName.toLowerCase().trim()}`;
        const existing = sheetMap.get(key);

        const targetTable = existing?.targetTable || ws.sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_');
        const matchedPreset = presets.find(p => 
          (p.filenamePattern && wbName.toLowerCase().includes(p.filenamePattern.toLowerCase())) ||
          p.sheetMappings?.some(sm => sm.workbookName?.toLowerCase() === wbName.toLowerCase() && sm.worksheetName.toLowerCase() === ws.sheetName.toLowerCase())
        ) || presets.find(p =>
          p.sheetMappings?.some(sm => sm.supabaseTable === targetTable)
        ) || presets.find(p => 
          p.sheetMappings?.some(sm => sm.worksheetName.toLowerCase() === ws.sheetName.toLowerCase())
        ) || (currentAnalysis.detectedArchetype ? presets.find(p => p.archetype === currentAnalysis.detectedArchetype) : null) || existing?.presetMatch || null;

        const pkey = existing?.primaryMergeKey || null;
        const lastSync = findLastSync(wbName, ws.sheetName, targetTable, ws.totalRows) || existing?.lastSyncInfo || null;

        sheetMap.set(key, {
          id: existing?.id || `sheet-${idx}-${ws.sheetName}`,
          sheetName: ws.sheetName,
          workbookName: wbName,
          mapping: existing?.mapping || null,
          analysis: {
            totalRows: ws.totalRows,
            totalColumns: ws.totalColumns,
            headerRow: ws.detectedHeaderRow,
            dataStartRow: ws.detectedDataStartRow,
            sampleCount: ws.sampleRows?.length || 0,
            mergedRangesCount: ws.mergedRanges?.length || 0,
          },
          targetTable,
          presetMatch: matchedPreset,
          primaryMergeKey: pkey,
          syncPolicy: existing?.syncPolicy || 'BIDIRECTIONAL',
          isEnabled: existing ? existing.isEnabled : true,
          skipMergedYearRows: existing ? existing.skipMergedYearRows : true,
          columnsCount: existing?.columnsCount || ws.headers?.length || 0,
          lastSyncInfo: lastSync,
        });
      });
    }

    // 3. Add any stored served sheets from server if not already in map
    storedServedSheets.forEach(s => {
      const wbName = s.workbookName || 'Workbook.xlsx';
      const wsName = s.sheetName;
      if (!wsName) return;
      const key = `${wbName.toLowerCase().trim()}::${wsName.toLowerCase().trim()}`;
      if (!sheetMap.has(key)) {
        const targetTable = s.targetTable || wsName.toLowerCase().replace(/[^a-z0-9_]/g, '_');
        const lastSync = findLastSync(wbName, wsName, targetTable);
        sheetMap.set(key, {
          id: s.id || `stored-${key}`,
          sheetName: wsName,
          workbookName: wbName,
          mapping: s.mapping || null,
          analysis: s.analysis || null,
          targetTable,
          presetMatch: null,
          primaryMergeKey: s.primaryMergeKey || null,
          syncPolicy: s.syncPolicy || 'BIDIRECTIONAL',
          isEnabled: s.isEnabled !== false,
          skipMergedYearRows: s.skipMergedYearRows !== false,
          columnsCount: s.columnsCount || s.mapping?.columns?.length || 0,
          lastSyncInfo: lastSync,
        });
      }
    });

    return Array.from(sheetMap.values());
  }, [currentAnalysis, mappings, presets, importLogs, storedServedSheets]);

  // Unique workbooks detected across all served sheets, mappings, and files
  const uniqueWorkbooks = useMemo(() => {
    const set = new Set<string>();
    servedSheets.forEach(s => {
      if (s.workbookName && typeof s.workbookName === 'string' && s.workbookName.trim()) {
        set.add(s.workbookName.trim());
      }
    });
    mappings.forEach(m => {
      if (m.workbookName && typeof m.workbookName === 'string' && m.workbookName.trim()) {
        set.add(m.workbookName.trim());
      }
    });
    if (currentAnalysis?.filename) {
      set.add(currentAnalysis.filename.trim());
    }
    files.forEach(f => {
      const fname = f.filename || (f as any).name;
      if (fname && typeof fname === 'string' && fname.trim()) {
        set.add(fname.trim());
      }
    });
    return Array.from(set);
  }, [servedSheets, mappings, currentAnalysis, files]);

  // Group served sheets by Target Supabase Table (Merged Table Topology)
  const mergedTablesTopology = useMemo(() => {
    const tableMap: Record<string, {
      tableName: string;
      isLiveInSupabase: boolean;
      approxDbRowCount: number | null;
      supabaseColumns: string[];
      mergedSheets: typeof servedSheets;
      primaryKeys: string[];
      syncPolicies: TableSyncPolicy[];
      totalEstimatedRows: number;
      lastSyncTimestamp: string | null;
    }> = {};

    servedSheets.forEach(sheet => {
      const tbl = sheet.targetTable;
      if (!tableMap[tbl]) {
        const liveInfo = supabaseTables.find(st => st.name.toLowerCase() === tbl.toLowerCase());
        tableMap[tbl] = {
          tableName: tbl,
          isLiveInSupabase: Boolean(liveInfo),
          approxDbRowCount: liveInfo?.approximateRowCount ?? null,
          supabaseColumns: liveInfo?.columns.map(c => c.name) || [],
          mergedSheets: [],
          primaryKeys: [],
          syncPolicies: [],
          totalEstimatedRows: 0,
          lastSyncTimestamp: null,
        };
      }

      tableMap[tbl].mergedSheets.push(sheet);
      if (sheet.primaryMergeKey && !tableMap[tbl].primaryKeys.includes(sheet.primaryMergeKey)) {
        tableMap[tbl].primaryKeys.push(sheet.primaryMergeKey);
      }
      if (!tableMap[tbl].syncPolicies.includes(sheet.syncPolicy)) {
        tableMap[tbl].syncPolicies.push(sheet.syncPolicy);
      }
      tableMap[tbl].totalEstimatedRows += sheet.analysis?.totalRows || 0;
      if (sheet.lastSyncInfo?.timestamp) {
        if (!tableMap[tbl].lastSyncTimestamp || new Date(sheet.lastSyncInfo.timestamp) > new Date(tableMap[tbl].lastSyncTimestamp!)) {
          tableMap[tbl].lastSyncTimestamp = sheet.lastSyncInfo.timestamp;
        }
      }
    });

    return Object.values(tableMap);
  }, [servedSheets, supabaseTables]);

  // Unified Chronological History Rows across all sheets
  const unifiedHistoryRows = useMemo(() => {
    const rows: {
      id: string;
      timestamp: string;
      workbookName: string;
      sheetName: string;
      targetTable: string;
      status: string;
      rowsProcessed: number;
      rowsInserted: number;
      rowsUpdated: number;
      rowsFailed: number;
      duration?: number;
      errorSummary?: string;
      triggerType: string;
    }[] = [];

    importLogs.forEach(log => {
      const syncResults = log.details?.syncResults as any[];
      if (syncResults && Array.isArray(syncResults) && syncResults.length > 0) {
        syncResults.forEach((sr, sIdx) => {
          rows.push({
            id: `${log.id}-sheet-${sIdx}`,
            timestamp: log.completedAt || log.startedAt,
            workbookName: log.filename,
            sheetName: sr.sheetName || 'Sheet',
            targetTable: sr.targetTable || 'table',
            status: sr.status || log.status,
            rowsProcessed: sr.rowsCount ?? (sr.insertedCount + sr.updatedCount + sr.failedCount),
            rowsInserted: sr.insertedCount ?? 0,
            rowsUpdated: sr.updatedCount ?? 0,
            rowsFailed: sr.failedCount ?? 0,
            duration: log.durationMs,
            errorSummary: sr.error || (sr.failedCount > 0 ? log.errorSummary : undefined),
            triggerType: log.details?.triggerType || (log.isDryRun ? 'DRY_RUN' : 'PIPELINE'),
          });
        });
      } else {
        // Single overall log entry
        rows.push({
          id: log.id,
          timestamp: log.completedAt || log.startedAt,
          workbookName: log.filename,
          sheetName: 'All Configured Sheets',
          targetTable: 'Supabase Public',
          status: log.status,
          rowsProcessed: log.rowsProcessed,
          rowsInserted: log.rowsInserted,
          rowsUpdated: log.rowsUpdated,
          rowsFailed: log.rowsFailed,
          duration: log.durationMs,
          errorSummary: log.errorSummary,
          triggerType: log.details?.triggerType || (log.isDryRun ? 'DRY_RUN' : 'PIPELINE'),
        });
      }
    });

    // Sort descending by timestamp
    return rows.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
  }, [importLogs]);

  // Unique lists for filter dropdowns
  const availableTargetTables = useMemo(() => {
    const set = new Set<string>();
    servedSheets.forEach(s => set.add(s.targetTable));
    return Array.from(set);
  }, [servedSheets]);

  const availablePresets = useMemo(() => {
    const list: { id: string; name: string }[] = [];
    presets.forEach(p => list.push({ id: p.id, name: p.name }));
    return list;
  }, [presets]);

  // Filtered Sheets
  const filteredSheets = useMemo(() => {
    return servedSheets.filter(sheet => {
      const q = searchQuery.toLowerCase().trim();
      const matchesSearch = !q || 
        sheet.sheetName.toLowerCase().includes(q) || 
        sheet.targetTable.toLowerCase().includes(q) ||
        sheet.workbookName.toLowerCase().includes(q) ||
        (sheet.primaryMergeKey && sheet.primaryMergeKey.toLowerCase().includes(q)) ||
        (sheet.presetMatch && sheet.presetMatch.name.toLowerCase().includes(q));

      const matchesWb = selectedWorkbookFilter === 'ALL' || 
        sheet.workbookName.toLowerCase().trim() === selectedWorkbookFilter.toLowerCase().trim();
      const matchesTable = tableFilter === 'ALL' || sheet.targetTable.toLowerCase() === tableFilter.toLowerCase();
      const matchesPreset = presetFilter === 'ALL' || (sheet.presetMatch?.id === presetFilter);
      const matchesPolicy = policyFilter === 'ALL' || sheet.syncPolicy === policyFilter;

      return matchesWb && matchesSearch && matchesTable && matchesPreset && matchesPolicy;
    });
  }, [servedSheets, searchQuery, tableFilter, presetFilter, policyFilter, selectedWorkbookFilter]);

  // Filtered Merged Tables
  const filteredMergedTables = useMemo(() => {
    return mergedTablesTopology.filter(tbl => {
      const q = searchQuery.toLowerCase().trim();
      const matchesSearch = !q || 
        tbl.tableName.toLowerCase().includes(q) || 
        tbl.mergedSheets.some(s => s.sheetName.toLowerCase().includes(q) || s.workbookName.toLowerCase().includes(q)) ||
        tbl.primaryKeys.some(pk => pk.toLowerCase().includes(q));
      
      const matchesWb = selectedWorkbookFilter === 'ALL' || 
        tbl.mergedSheets.some(s => s.workbookName.toLowerCase().trim() === selectedWorkbookFilter.toLowerCase().trim());
      const matchesTable = tableFilter === 'ALL' || tbl.tableName.toLowerCase() === tableFilter.toLowerCase();
      return matchesWb && matchesSearch && matchesTable;
    });
  }, [mergedTablesTopology, searchQuery, tableFilter, selectedWorkbookFilter]);

  // Filtered History
  const filteredHistory = useMemo(() => {
    return unifiedHistoryRows.filter(row => {
      const q = searchQuery.toLowerCase().trim();
      const matchesSearch = !q || 
        row.sheetName.toLowerCase().includes(q) || 
        row.targetTable.toLowerCase().includes(q) ||
        row.workbookName.toLowerCase().includes(q);

      const matchesWb = selectedWorkbookFilter === 'ALL' || 
        row.workbookName.toLowerCase().trim() === selectedWorkbookFilter.toLowerCase().trim();
      const matchesTable = tableFilter === 'ALL' || row.targetTable.toLowerCase() === tableFilter.toLowerCase();
      return matchesWb && matchesSearch && matchesTable;
    });
  }, [unifiedHistoryRows, searchQuery, tableFilter, selectedWorkbookFilter]);

  const handleRefreshAllNextcloudWorkbooks = async () => {
    setIsRefreshingNextcloud(true);
    setSyncFeedback(null);
    try {
      const res = await ApiClient.refreshAllNextcloudWorkbooks({
        nextcloud: nextcloudConfig || {} as any,
        supabase: supabaseConfig || {} as any,
        mappings,
      });
      if (res.success) {
        setSyncFeedback({
          type: 'success',
          message: `🔄 ${res.message || 'Successfully re-downloaded all workbooks from Nextcloud and synchronized all sheets into Supabase!'}`
        });
        if (onTriggerSync) {
          onTriggerSync();
        }
      } else {
        setSyncFeedback({
          type: 'error',
          message: `Re-fetch error: ${res.error || 'Failed to re-fetch workbooks from Nextcloud'}`
        });
      }
    } catch (e: any) {
      setSyncFeedback({
        type: 'error',
        message: `Re-fetch error: ${e.message}`
      });
    } finally {
      setIsRefreshingNextcloud(false);
    }
  };

  // Trigger one-click synchronization
  const handleExecuteSync = async () => {
    if (!onTriggerSync) return;
    setIsSyncing(true);
    setSyncFeedback(null);
    try {
      await onTriggerSync();
      setSyncFeedback({
        type: 'success',
        message: 'Full sync pipeline completed successfully across all served Excel sheets!'
      });
      loadAllMetadata();
    } catch (e: any) {
      setSyncFeedback({
        type: 'error',
        message: e.message || 'Sync execution encountered an error'
      });
    } finally {
      setIsSyncing(false);
      setTimeout(() => setSyncFeedback(null), 5000);
    }
  };

  // Export Matrix as JSON or CSV
  const handleExportMatrixJson = () => {
    const payload = {
      exportedAt: new Date().toISOString(),
      permanentSavedAt,
      workbook: currentAnalysis?.filename || 'Workbook.xlsx',
      totalServedSheets: servedSheets.length,
      mergedTablesCount: mergedTablesTopology.length,
      sheets: servedSheets.map(s => ({
        sheetName: s.sheetName,
        targetTable: s.targetTable,
        syncPolicy: s.syncPolicy,
        primaryMergeKey: s.primaryMergeKey,
        preset: s.presetMatch ? { id: s.presetMatch.id, name: s.presetMatch.name, archetype: s.presetMatch.archetype } : null,
        skipMergedYearRows: s.skipMergedYearRows,
        columnsCount: s.columnsCount,
        lastSync: s.lastSyncInfo,
        columns: s.mapping?.columns || []
      })),
      mergedTables: mergedTablesTopology.map(t => ({
        tableName: t.tableName,
        mergedSheetNames: t.mergedSheets.map(s => s.sheetName),
        primaryKeys: t.primaryKeys,
        syncPolicies: t.syncPolicies,
        approxDbRowCount: t.approxDbRowCount,
      })),
      recentHistory: unifiedHistoryRows.slice(0, 50)
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `served_mappings_matrix_${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);

    setCopyFeedback('Downloaded comprehensive mappings & history matrix JSON!');
    setTimeout(() => setCopyFeedback(null), 4000);
  };

  const formatRelativeTime = (isoString?: string | null) => {
    if (!isoString) return 'Never';
    const date = new Date(isoString);
    if (isNaN(date.getTime())) return isoString;
    const diffMs = Date.now() - date.getTime();
    const diffSec = Math.floor(diffMs / 1000);
    if (diffSec < 60) return `${diffSec}s ago`;
    const diffMin = Math.floor(diffSec / 60);
    if (diffMin < 60) return `${diffMin}m ago`;
    const diffHours = Math.floor(diffMin / 60);
    if (diffHours < 24) return `${diffHours}h ago`;
    return date.toLocaleDateString() + ' ' + date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  };

  return (
    <div className="space-y-6">
      {/* Top Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3.5">
          <div className="p-3 rounded-xl bg-gradient-to-br from-indigo-500 to-purple-600 text-white shadow-xs">
            <Layers className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center space-x-2.5">
              <h1 className="text-xl font-bold text-slate-900 tracking-tight">Served Sheets & Mappings Hub</h1>
              <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-indigo-100 text-indigo-800">
                Single Truth Matrix
              </span>
              {permanentSavedAt && (
                <span className="hidden lg:inline-flex items-center space-x-1 px-2 py-0.5 rounded-full text-[11px] bg-emerald-50 text-emerald-700 border border-emerald-200">
                  <ShieldCheck className="w-3 h-3 text-emerald-600" />
                  <span>Saved: {formatRelativeTime(permanentSavedAt)}</span>
                </span>
              )}
            </div>
            <p className="text-xs text-slate-500 mt-1">
              One centralized place to inspect all served Excel sheets, mapped Supabase tables, active presets, primary merge keys, and unified sync history.
            </p>
          </div>
        </div>

        {/* Global Hub Actions */}
        <div className="flex items-center space-x-2 flex-wrap gap-y-2">
          <button
            onClick={loadAllMetadata}
            disabled={isLoadingMetadata}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg border border-slate-300 text-xs font-semibold text-slate-700 hover:bg-slate-50 shadow-2xs transition-colors"
            title="Refresh metadata, live Supabase schema, and permanent mappings"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isLoadingMetadata ? 'animate-spin' : ''}`} />
            <span>Refresh</span>
          </button>

          <button
            onClick={handleSaveServedSheetsPermanently}
            disabled={isSavingPermanent}
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg bg-emerald-50 border border-emerald-300 text-xs font-bold text-emerald-800 hover:bg-emerald-100 shadow-2xs transition-colors"
            title="Permanently preserve served sheets, merged table topology, and mapping hubs to server storage"
          >
            <ShieldCheck className={`w-3.5 h-3.5 text-emerald-600 ${isSavingPermanent ? 'animate-spin' : ''}`} />
            <span>{isSavingPermanent ? 'Saving...' : 'Save Permanently to Storage'}</span>
          </button>

          <button
            onClick={handleExportMatrixJson}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg border border-slate-300 text-xs font-semibold text-slate-700 hover:bg-slate-50 shadow-2xs transition-colors"
            title="Export complete served sheets, merged tables, and sync history to JSON"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>Export Matrix JSON</span>
          </button>

          <button
            onClick={() => setShowAiPresetsModal(true)}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg bg-purple-50 border border-purple-200 text-xs font-semibold text-purple-800 hover:bg-purple-100 shadow-2xs transition-colors"
            title="Browse pre-built workbook archetype presets and combined student batch tables"
          >
            <Sparkles className="w-3.5 h-3.5 text-purple-600" />
            <span>AI Presets ({presets.length})</span>
          </button>

          <button
            onClick={() => onNavigate('mappings')}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg bg-indigo-50 border border-indigo-200 text-xs font-semibold text-indigo-700 hover:bg-indigo-100 shadow-2xs transition-colors"
          >
            <GitFork className="w-3.5 h-3.5 text-indigo-600" />
            <span>Edit Mappings</span>
          </button>

          <button
            onClick={handleRefreshAllNextcloudWorkbooks}
            disabled={isRefreshingNextcloud || isSyncing}
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg bg-sky-600 hover:bg-sky-700 text-xs font-semibold text-white shadow-2xs transition-colors disabled:opacity-50"
            title="Automatically re-download all Excel files from Nextcloud WebDAV and sync all sheets into Supabase"
          >
            <ArrowDownToLine className={`w-3.5 h-3.5 ${isRefreshingNextcloud ? 'animate-bounce' : ''}`} />
            <span>{isRefreshingNextcloud ? 'Re-fetching Nextcloud Files...' : 'Re-fetch All Nextcloud Files & Sync'}</span>
          </button>

          {onTriggerSync && (
            <button
              onClick={handleExecuteSync}
              disabled={isSyncing}
              className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-semibold text-white shadow-2xs transition-colors disabled:opacity-50"
            >
              <Play className={`w-3.5 h-3.5 fill-current ${isSyncing ? 'animate-pulse' : ''}`} />
              <span>{isSyncing ? 'Syncing All Served Files...' : 'Sync All Served Files Now'}</span>
            </button>
          )}
        </div>
      </div>

      {/* Status Notices */}
      {syncFeedback && (
        <div className={`p-4 rounded-xl border text-xs font-medium flex items-center space-x-2.5 ${
          syncFeedback.type === 'success' 
            ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
            : 'bg-rose-50 text-rose-800 border-rose-200'
        }`}>
          {syncFeedback.type === 'success' ? (
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          ) : (
            <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
          )}
          <span>{syncFeedback.message}</span>
        </div>
      )}

      {copyFeedback && (
        <div className="p-3 rounded-lg bg-indigo-50 text-indigo-800 border border-indigo-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-indigo-600 shrink-0" />
          <span>{copyFeedback}</span>
        </div>
      )}

      {/* Bird's-Eye Overview KPI Cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3.5">
        {/* Card 1: Served Worksheets */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center space-x-3.5">
          <div className="p-2.5 rounded-lg bg-emerald-50 text-emerald-600 shrink-0">
            <FileSpreadsheet className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Served Sheets</div>
            <div className="text-xl font-bold text-slate-900 mt-0.5">{servedSheets.length}</div>
            <div className="text-[11px] text-slate-500 truncate">
              {servedSheets.filter(s => s.isEnabled).length} active • {currentAnalysis?.filename || 'All Workbooks'}
            </div>
          </div>
        </div>

        {/* Card 2: Merged Supabase Tables */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center space-x-3.5">
          <div className="p-2.5 rounded-lg bg-teal-50 text-teal-600 shrink-0">
            <Database className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Destination Tables</div>
            <div className="text-xl font-bold text-slate-900 mt-0.5">{mergedTablesTopology.length}</div>
            <div className="text-[11px] text-slate-500 truncate">
              {mergedTablesTopology.filter(t => t.mergedSheets.length > 1).length} merged tables (1:N)
            </div>
          </div>
        </div>

        {/* Card 3: AI Presets Active */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center space-x-3.5">
          <div className="p-2.5 rounded-lg bg-purple-50 text-purple-600 shrink-0">
            <Sparkles className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">AI Presets & Archetypes</div>
            <div className="text-xl font-bold text-slate-900 mt-0.5">
              {servedSheets.filter(s => s.presetMatch).length} / {servedSheets.length}
            </div>
            <div className="text-[11px] text-slate-500 truncate">
              {presets.length} total presets in store
            </div>
          </div>
        </div>

        {/* Card 4: Last Updated Time & Success Rate */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center space-x-3.5">
          <div className="p-2.5 rounded-lg bg-indigo-50 text-indigo-600 shrink-0">
            <Clock className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Last Sync Run</div>
            <div className="text-base font-bold text-slate-900 mt-0.5 truncate">
              {formatRelativeTime(importLogs[0]?.completedAt || importLogs[0]?.startedAt || permanentSavedAt)}
            </div>
            <div className="text-[11px] text-slate-500 truncate">
              {importLogs[0]?.status ? (
                <span className="font-semibold text-emerald-600">{importLogs[0].status}</span>
              ) : (
                'Ready for sync'
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Main Container with 3 View Modes */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
        {/* Navigation & Controls Bar */}
        <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
          {/* Mode Switcher */}
          <div className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-white text-xs font-medium shadow-2xs">
            <button
              onClick={() => setViewMode('SHEETS_MATRIX')}
              className={`px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition-all ${
                viewMode === 'SHEETS_MATRIX'
                  ? 'bg-indigo-600 text-white font-semibold shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Served Sheets Matrix</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${
                viewMode === 'SHEETS_MATRIX' ? 'bg-indigo-700 text-white' : 'bg-slate-100 text-slate-700'
              }`}>
                {servedSheets.length}
              </span>
            </button>

            <button
              onClick={() => setViewMode('MERGED_TABLES')}
              className={`px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition-all ${
                viewMode === 'MERGED_TABLES'
                  ? 'bg-indigo-600 text-white font-semibold shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Database className="w-3.5 h-3.5" />
              <span>Merged Supabase Tables</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${
                viewMode === 'MERGED_TABLES' ? 'bg-indigo-700 text-white' : 'bg-slate-100 text-slate-700'
              }`}>
                {mergedTablesTopology.length}
              </span>
            </button>

            <button
              onClick={() => setViewMode('UNIFIED_HISTORY')}
              className={`px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition-all ${
                viewMode === 'UNIFIED_HISTORY'
                  ? 'bg-indigo-600 text-white font-semibold shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Clock className="w-3.5 h-3.5" />
              <span>Unified Sheet History</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${
                viewMode === 'UNIFIED_HISTORY' ? 'bg-indigo-700 text-white' : 'bg-slate-100 text-slate-700'
              }`}>
                {unifiedHistoryRows.length}
              </span>
            </button>
          </div>

          {/* Search & Filter Controls */}
          <div className="flex items-center space-x-2 flex-wrap gap-y-2">
            <div className="relative">
              <input
                type="text"
                placeholder="Search sheet, table, key..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="pl-8 pr-3 py-1.5 rounded-lg border border-slate-300 text-xs w-48 sm:w-56 focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
              />
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
            </div>

            {/* Workbook Filter Dropdown */}
            <select
              value={selectedWorkbookFilter}
              onChange={(e) => setSelectedWorkbookFilter(e.target.value)}
              className="px-2.5 py-1.5 rounded-lg border border-slate-300 text-xs text-slate-700 bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 font-medium"
            >
              <option value="ALL">All Workbooks ({servedSheets.length} sheets)</option>
              {uniqueWorkbooks.map(wb => {
                const count = servedSheets.filter(s => s.workbookName.toLowerCase().trim() === wb.toLowerCase().trim()).length;
                return (
                  <option key={wb} value={wb}>{wb} ({count} sheets)</option>
                );
              })}
            </select>

            <select
              value={tableFilter}
              onChange={(e) => setTableFilter(e.target.value)}
              className="px-2.5 py-1.5 rounded-lg border border-slate-300 text-xs text-slate-700 bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500"
            >
              <option value="ALL">All Supabase Tables</option>
              {availableTargetTables.map(tbl => (
                <option key={tbl} value={tbl}>{tbl}</option>
              ))}
            </select>

            {viewMode === 'SHEETS_MATRIX' && (
              <select
                value={policyFilter}
                onChange={(e) => setPolicyFilter(e.target.value)}
                className="px-2.5 py-1.5 rounded-lg border border-slate-300 text-xs text-slate-700 bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500"
              >
                <option value="ALL">All Policies</option>
                <option value="BIDIRECTIONAL">Bi-Directional</option>
                <option value="EXCEL_TO_DB">Excel → Supabase</option>
                <option value="DB_TO_EXCEL">Supabase → Excel</option>
                <option value="READ_ONLY">Read Only</option>
              </select>
            )}
          </div>
        </div>

        {/* Workbook Filter Pills Bar */}
        {uniqueWorkbooks.length > 0 && (
          <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center space-x-2 overflow-x-auto text-xs">
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider shrink-0 mr-1 flex items-center space-x-1">
              <Folder className="w-3.5 h-3.5 text-indigo-500" />
              <span>Workbook Filter:</span>
            </span>
            <button
              onClick={() => setSelectedWorkbookFilter('ALL')}
              className={`px-3 py-1 rounded-full text-xs font-semibold whitespace-nowrap transition-colors shadow-2xs ${
                selectedWorkbookFilter === 'ALL'
                  ? 'bg-indigo-600 text-white'
                  : 'bg-white text-slate-700 border border-slate-200 hover:bg-slate-100'
              }`}
            >
              All Workbooks ({servedSheets.length})
            </button>
            {uniqueWorkbooks.map(wb => {
              const count = servedSheets.filter(s => s.workbookName.toLowerCase().trim() === wb.toLowerCase().trim()).length;
              return (
                <button
                  key={wb}
                  onClick={() => setSelectedWorkbookFilter(wb)}
                  className={`px-3 py-1 rounded-full text-xs whitespace-nowrap transition-colors shadow-2xs flex items-center space-x-1.5 ${
                    selectedWorkbookFilter.toLowerCase().trim() === wb.toLowerCase().trim()
                      ? 'bg-indigo-600 text-white font-semibold'
                      : 'bg-white text-slate-700 border border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  <FileSpreadsheet className="w-3 h-3 text-emerald-500 shrink-0" />
                  <span>{wb}</span>
                  <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${
                    selectedWorkbookFilter.toLowerCase().trim() === wb.toLowerCase().trim()
                      ? 'bg-indigo-700 text-white'
                      : 'bg-slate-100 text-slate-600'
                  }`}>
                    {count}
                  </span>
                </button>
              );
            })}
          </div>
        )}

        {/* ========================================================= */}
        {/* VIEW 1: SERVED SHEETS MATRIX                              */}
        {/* ========================================================= */}
        {viewMode === 'SHEETS_MATRIX' && (
          <div className="p-4 space-y-4">
            {filteredSheets.length === 0 ? (
              <div className="p-12 text-center text-slate-400">
                <FileSpreadsheet className="w-10 h-10 mx-auto mb-2 opacity-50" />
                <div className="font-semibold text-slate-700 text-sm">No served worksheets match current filters</div>
                <div className="text-xs text-slate-500 mt-1">Try resetting the search query or table filter</div>
              </div>
            ) : (
              filteredSheets.map((sheet, sIdx) => {
                const isExpanded = expandedSheetId === sheet.id;
                return (
                  <div 
                    key={sheet.id}
                    className="border border-slate-200 rounded-xl bg-white shadow-2xs hover:border-indigo-300 transition-all overflow-hidden"
                  >
                    {/* Sheet Summary Card Header */}
                    <div className="p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-3 bg-gradient-to-r from-slate-50 to-white">
                      {/* Left: Sheet Details */}
                      <div className="flex items-start space-x-3">
                        <div className="p-2.5 rounded-lg bg-emerald-100 text-emerald-800 font-mono font-bold text-xs shrink-0 mt-0.5">
                          #{sIdx + 1}
                        </div>
                        <div>
                          <div className="flex items-center space-x-2 flex-wrap gap-y-1">
                            <span className="text-sm font-bold text-slate-900 tracking-tight">
                              {sheet.sheetName}
                            </span>
                            <span className="text-[11px] font-mono px-2 py-0.5 rounded-md bg-slate-100 text-slate-600">
                              {sheet.workbookName}
                            </span>
                            {sheet.isEnabled ? (
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-emerald-100 text-emerald-800">
                                Active Sync
                              </span>
                            ) : (
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-slate-200 text-slate-600">
                                Paused
                              </span>
                            )}
                            {sheet.skipMergedYearRows && (
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200" title="Merged year and divider banners are excluded from data ingestion">
                                🛡️ Year Rows Excluded
                              </span>
                            )}
                          </div>

                          <div className="flex items-center space-x-3 text-xs text-slate-500 mt-1.5 flex-wrap gap-y-1">
                            {sheet.analysis && (
                              <span>
                                <strong className="text-slate-700">{sheet.analysis.totalRows}</strong> rows × <strong className="text-slate-700">{sheet.analysis.totalColumns}</strong> cols (Header: Row {sheet.analysis.headerRow}, Data: Row {sheet.analysis.dataStartRow})
                              </span>
                            )}
                            <span>•</span>
                            <span className="inline-flex items-center space-x-1 font-semibold text-slate-700">
                              <Tag className="w-3 h-3 text-slate-400" />
                              <span>{sheet.columnsCount} columns mapped</span>
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* Center: Mapped Supabase Table & Preset Badge */}
                      <div className="flex items-center space-x-3 flex-wrap gap-y-2">
                        {/* Target Supabase Table */}
                        <div className="p-2 rounded-lg bg-teal-50 border border-teal-200 text-xs">
                          <div className="text-[10px] font-semibold text-teal-800 uppercase tracking-wider flex items-center space-x-1">
                            <Database className="w-3 h-3" />
                            <span>Destination Table</span>
                          </div>
                          <div className="font-mono font-bold text-teal-900 mt-0.5">
                            public.{sheet.targetTable}
                          </div>
                          <div className="text-[10px] text-teal-700 mt-0.5">
                            {sheet.primaryMergeKey ? (
                              <span className="inline-flex items-center space-x-1">
                                <Key className="w-2.5 h-2.5" />
                                <span>Key: <strong>{sheet.primaryMergeKey}</strong> (Upsert)</span>
                              </span>
                            ) : (
                              <span>Direct Insert (No Unique Key)</span>
                            )}
                          </div>
                        </div>

                        {/* Associated AI Preset / Archetype */}
                        <div className="p-2 rounded-lg bg-purple-50 border border-purple-200 text-xs min-w-[140px]">
                          <div className="text-[10px] font-semibold text-purple-800 uppercase tracking-wider flex items-center space-x-1">
                            <Sparkles className="w-3 h-3" />
                            <span>Preset / Archetype</span>
                          </div>
                          <div className="font-semibold text-purple-900 mt-0.5 truncate">
                            {sheet.presetMatch?.name || 'Standard Tabular'}
                          </div>
                          <div className="text-[10px] text-purple-700 mt-0.5">
                            {sheet.presetMatch?.badge || '📊 Tabular Mapping'}
                          </div>
                        </div>

                        {/* Last Sync Info */}
                        <div className="p-2 rounded-lg bg-slate-100 text-xs min-w-[130px]">
                          <div className="text-[10px] font-semibold text-slate-600 uppercase tracking-wider flex items-center space-x-1">
                            <Clock className="w-3 h-3" />
                            <span>Last Updated</span>
                          </div>
                          <div className="font-semibold text-slate-800 mt-0.5">
                            {formatRelativeTime(sheet.lastSyncInfo?.timestamp)}
                          </div>
                          <div className="text-[10px] text-slate-600 mt-0.5">
                            {sheet.lastSyncInfo ? (
                              <span>+{sheet.lastSyncInfo.inserted} ins, ~{sheet.lastSyncInfo.updated} upd</span>
                            ) : (
                              <span>Pending first run</span>
                            )}
                          </div>
                        </div>
                      </div>

                      {/* Right: Expand & Navigation Actions */}
                      <div className="flex items-center space-x-2 shrink-0">
                        <button
                          onClick={() => {
                            onNavigate('mappings');
                          }}
                          className="px-2.5 py-1.5 rounded-lg border border-slate-300 text-xs font-semibold text-slate-700 hover:bg-slate-100 transition-colors"
                          title="Open in Table & Column Mappings view"
                        >
                          Edit Mappings
                        </button>

                        <button
                          onClick={() => setExpandedSheetId(isExpanded ? null : sheet.id)}
                          className="p-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-600 transition-colors"
                          title={isExpanded ? 'Collapse column mappings' : 'Expand column mappings'}
                        >
                          {isExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
                        </button>
                      </div>
                    </div>

                    {/* Expandable Column Mapping Table */}
                    {isExpanded && sheet.mapping && (
                      <div className="p-4 bg-slate-50 border-t border-slate-200">
                        <div className="flex items-center justify-between mb-2">
                          <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center space-x-1.5">
                            <GitFork className="w-3.5 h-3.5 text-indigo-600" />
                            <span>Mapped Columns for {sheet.sheetName} → public.{sheet.targetTable}</span>
                          </h4>
                          <span className="text-[11px] text-slate-500">
                            Sync Policy: <strong>{sheet.syncPolicy}</strong>
                          </span>
                        </div>

                        <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
                          <table className="w-full text-left text-xs">
                            <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                              <tr>
                                <th className="px-3 py-2 text-center w-12">Col</th>
                                <th className="px-3 py-2">Excel Header</th>
                                <th className="px-3 py-2 text-center">Direction</th>
                                <th className="px-3 py-2">Supabase Column</th>
                                <th className="px-3 py-2">Data Type</th>
                                <th className="px-3 py-2 text-center">Primary Key</th>
                                <th className="px-3 py-2">Transformation</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100 font-mono">
                              {sheet.mapping.columns.map((col, cIdx) => (
                                <tr key={col.id || cIdx} className="hover:bg-slate-50">
                                  <td className="px-3 py-1.5 text-center font-bold text-slate-500">{col.excelColumn}</td>
                                  <td className="px-3 py-1.5 font-sans font-medium text-slate-800">{col.excelHeader}</td>
                                  <td className="px-3 py-1.5 text-center text-slate-400">
                                    <ArrowRight className="w-3 h-3 mx-auto text-indigo-500" />
                                  </td>
                                  <td className="px-3 py-1.5 font-bold text-teal-800">{col.supabaseColumn}</td>
                                  <td className="px-3 py-1.5 font-sans">
                                    <span className="px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 text-[11px]">
                                      {col.dataType}
                                    </span>
                                  </td>
                                  <td className="px-3 py-1.5 text-center">
                                    {col.uniqueKey ? (
                                      <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 font-sans font-bold text-[10px] inline-flex items-center space-x-1">
                                        <Key className="w-2.5 h-2.5" />
                                        <span>MERGE KEY</span>
                                      </span>
                                    ) : (
                                      <span className="text-slate-300 font-sans">-</span>
                                    )}
                                  </td>
                                  <td className="px-3 py-1.5 font-sans text-slate-600 text-[11px]">
                                    {col.transformation || 'none'}
                                  </td>
                                </tr>
                              ))}
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
        )}

        {/* ========================================================= */}
        {/* VIEW 2: MERGED SUPABASE TABLES TOPOLOGY                   */}
        {/* ========================================================= */}
        {viewMode === 'MERGED_TABLES' && (
          <div className="p-4 space-y-4">
            <div className="p-3 rounded-lg bg-teal-50 border border-teal-200 text-xs text-teal-800 flex items-center space-x-2">
              <Info className="w-4 h-4 text-teal-600 shrink-0" />
              <span>
                <strong>Merged Table Topology:</strong> This view reveals which multiple Excel worksheets combine into single unified destination tables in Supabase.
              </span>
            </div>

            {filteredMergedTables.length === 0 ? (
              <div className="p-12 text-center text-slate-400">
                <Database className="w-10 h-10 mx-auto mb-2 opacity-50" />
                <div className="font-semibold text-slate-700 text-sm">No destination tables found</div>
              </div>
            ) : (
              filteredMergedTables.map((tbl, tIdx) => (
                <div 
                  key={tbl.tableName}
                  className="border border-slate-200 rounded-xl bg-white shadow-2xs overflow-hidden"
                >
                  <div className="p-4 bg-slate-900 text-white flex flex-col md:flex-row md:items-center justify-between gap-3">
                    <div className="flex items-center space-x-3">
                      <div className="p-2.5 rounded-lg bg-teal-600 text-white font-mono">
                        <Database className="w-5 h-5" />
                      </div>
                      <div>
                        <div className="flex items-center space-x-2 flex-wrap">
                          <h3 className="text-base font-bold font-mono tracking-tight text-white">
                            public.{tbl.tableName}
                          </h3>
                          {tbl.isLiveInSupabase ? (
                            <span className="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-teal-500/20 text-teal-300 border border-teal-500/30">
                              ✓ Verified in Supabase
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                              ⚠️ Pending Creation
                            </span>
                          )}
                          {tbl.mergedSheets.length > 1 && (
                            <span className="px-2 py-0.5 rounded-full text-[11px] font-bold bg-purple-500/30 text-purple-300 border border-purple-400/40">
                              🔗 {tbl.mergedSheets.length} Sheets Merged
                            </span>
                          )}
                        </div>
                        <p className="text-xs text-slate-400 mt-1">
                          Primary Conflict Merge Key:{' '}
                          <strong className="text-amber-400 font-mono">
                            {tbl.primaryKeys.join(', ') || 'No explicit unique key (direct INSERT)'}
                          </strong>
                          {tbl.approxDbRowCount !== null && (
                            <span className="ml-2 font-sans">• Approx {tbl.approxDbRowCount} live rows in DB</span>
                          )}
                        </p>
                      </div>
                    </div>

                    <div className="text-right text-xs">
                      <div className="text-slate-400 text-[11px]">Last Updated</div>
                      <div className="font-semibold text-slate-200 font-mono">
                        {formatRelativeTime(tbl.lastSyncTimestamp)}
                      </div>
                    </div>
                  </div>

                  {/* Feeding Sheets List */}
                  <div className="p-4 bg-slate-50 border-b border-slate-200">
                    <div className="text-xs font-bold text-slate-700 uppercase tracking-wider mb-2.5">
                      Excel Worksheets Feeding into public.{tbl.tableName}:
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                      {tbl.mergedSheets.map((ms) => (
                        <div 
                          key={ms.id}
                          className="p-3 rounded-lg border border-slate-200 bg-white shadow-2xs flex flex-col justify-between"
                        >
                          <div>
                            <div className="flex items-center justify-between">
                              <span className="font-bold text-slate-900 text-xs truncate">
                                {ms.sheetName}
                              </span>
                              <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-100 text-slate-600">
                                {ms.workbookName}
                              </span>
                            </div>
                            <div className="text-[11px] text-slate-500 mt-1">
                              {ms.columnsCount} columns mapped • {ms.analysis ? `${ms.analysis.totalRows} rows` : 'Custom Sheet'}
                            </div>
                            {ms.presetMatch && (
                              <div className="text-[10px] text-purple-700 font-semibold mt-1 flex items-center space-x-1">
                                <Sparkles className="w-3 h-3" />
                                <span>{ms.presetMatch.name}</span>
                              </div>
                            )}
                          </div>

                          <div className="mt-2.5 pt-2 border-t border-slate-100 flex items-center justify-between text-[11px]">
                            <span className="text-slate-500">
                              Merge Key: <strong className="text-teal-700 font-mono">{ms.primaryMergeKey || 'default'}</strong>
                            </span>
                            <span className="font-semibold text-emerald-600">
                              {ms.lastSyncInfo ? `${ms.lastSyncInfo.status}` : 'Pending'}
                            </span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>
        )}

        {/* ========================================================= */}
        {/* VIEW 3: UNIFIED SHEET SYNC & AUDIT HISTORY                */}
        {/* ========================================================= */}
        {viewMode === 'UNIFIED_HISTORY' && (
          <div className="p-4 space-y-4">
            <div className="flex items-center justify-between">
              <div className="text-xs text-slate-500">
                Chronological historical audit records of every sheet processed, row count, and database response.
              </div>
              <span className="text-xs font-mono font-bold text-slate-700">
                {filteredHistory.length} Total Execution Events
              </span>
            </div>

            {filteredHistory.length === 0 ? (
              <div className="p-12 text-center text-slate-400">
                <Clock className="w-10 h-10 mx-auto mb-2 opacity-50" />
                <div className="font-semibold text-slate-700 text-sm">No sync history recorded yet</div>
                <div className="text-xs text-slate-500 mt-1">Run a pipeline sync or dry run to generate audit logs</div>
              </div>
            ) : (
              <div className="overflow-x-auto rounded-xl border border-slate-200">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 uppercase font-semibold border-b border-slate-200">
                    <tr>
                      <th className="px-4 py-3">Timestamp</th>
                      <th className="px-4 py-3">Workbook & Sheet</th>
                      <th className="px-4 py-3">Supabase Destination</th>
                      <th className="px-3 py-3 text-center">Status</th>
                      <th className="px-3 py-3 text-center">Processed</th>
                      <th className="px-3 py-3 text-emerald-700 text-center">Inserted</th>
                      <th className="px-3 py-3 text-purple-700 text-center">Updated</th>
                      <th className="px-3 py-3 text-rose-700 text-center">Failed</th>
                      <th className="px-4 py-3">Trigger</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 bg-white">
                    {filteredHistory.map((item) => (
                      <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                        <td className="px-4 py-3 font-mono text-[11px] text-slate-600">
                          <div>{new Date(item.timestamp).toLocaleDateString()}</div>
                          <div className="text-slate-400 text-[10px]">
                            {new Date(item.timestamp).toLocaleTimeString()} ({formatRelativeTime(item.timestamp)})
                          </div>
                        </td>

                        <td className="px-4 py-3">
                          <div className="font-bold text-slate-900">{item.sheetName}</div>
                          <div className="font-mono text-[10px] text-slate-500">{item.workbookName}</div>
                          {item.errorSummary && (
                            <div className="text-[10px] text-rose-600 mt-0.5">{item.errorSummary}</div>
                          )}
                        </td>

                        <td className="px-4 py-3 font-mono font-semibold text-teal-800">
                          public.{item.targetTable}
                        </td>

                        <td className="px-3 py-3 text-center">
                          {item.status === 'Success' ? (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">
                              Success
                            </span>
                          ) : item.status === 'Partial Success' ? (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800">
                              Partial
                            </span>
                          ) : item.status === 'Failed' ? (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-100 text-rose-800">
                              Failed
                            </span>
                          ) : (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-700">
                              {item.status}
                            </span>
                          )}
                        </td>

                        <td className="px-3 py-3 text-center font-mono font-semibold text-slate-800">
                          {item.rowsProcessed}
                        </td>

                        <td className="px-3 py-3 text-center font-mono font-bold text-emerald-600">
                          +{item.rowsInserted}
                        </td>

                        <td className="px-3 py-3 text-center font-mono font-bold text-purple-600">
                          ~{item.rowsUpdated}
                        </td>

                        <td className="px-3 py-3 text-center font-mono font-bold text-rose-600">
                          {item.rowsFailed}
                        </td>

                        <td className="px-4 py-3">
                          <span className="px-2 py-0.5 rounded-md font-mono text-[10px] font-semibold bg-slate-100 text-slate-700">
                            {item.triggerType}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>

      {/* AI Presets & Inbuilt Analyser Modal */}
      <AiPresetsModal
        isOpen={showAiPresetsModal}
        onClose={() => setShowAiPresetsModal(false)}
        currentAnalysis={currentAnalysis || null}
        activeMappings={mappings}
        onApplyPreset={handleApplyPreset}
        onApplyAllPresets={handleApplyAllStudentPresets}
        onLoadPresetWorkbook={() => {
          setShowAiPresetsModal(false);
          onNavigate('analyzer');
        }}
        supabaseConfig={supabaseConfig}
        supabaseTables={supabaseTables}
      />
    </div>
  );
};
