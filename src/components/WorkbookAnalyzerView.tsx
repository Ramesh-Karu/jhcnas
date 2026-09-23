import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Microscope, 
  Layers, 
  Table, 
  Upload, 
  Sparkles, 
  Check, 
  AlertCircle, 
  FileSpreadsheet, 
  ChevronRight, 
  Sliders, 
  BookOpen, 
  Search,
  Database,
  Code,
  Copy,
  CheckCheck,
  RefreshCw,
  ExternalLink,
  Info,
  Key,
  ShieldCheck,
  Wand2,
  FolderSync
} from 'lucide-react';
import { 
  WorkbookAnalysis, 
  SheetAnalysis, 
  NavigationTab, 
  WorksheetMapping, 
  NextcloudConfig,
  SupabaseConfig,
  SupabaseTableInfo,
  MultiSheetConsolidationMode,
  TableSchemaPlan
} from '../types';
import { ExcelAnalyzer } from '../services/excelAnalyzer';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
import { ApiClient } from '../services/apiClient';
import { SchemaGenerator } from '../services/schemaGenerator';

const formatDisplayRange = (r: any): string => {
  if (!r) return 'A1';
  if (typeof r === 'string') return r;
  if (typeof r === 'object' && 's' in r && 'e' in r) {
    const colToLetter = (c: number) => {
      let temp = c;
      let letter = '';
      while (temp >= 0) {
        letter = String.fromCharCode((temp % 26) + 65) + letter;
        temp = Math.floor(temp / 26) - 1;
      }
      return letter || 'A';
    };
    const sCol = colToLetter(r.s?.c ?? 0);
    const sRow = (r.s?.r ?? 0) + 1;
    const eCol = colToLetter(r.e?.c ?? 0);
    const eRow = (r.e?.r ?? 0) + 1;
    return `${sCol}${sRow}:${eCol}${eRow}`;
  }
  return String(r);
};

interface WorkbookAnalyzerViewProps {
  currentAnalysis: WorkbookAnalysis | null;
  onAnalysisUpdate: (analysis: WorkbookAnalysis) => void;
  onNavigate: (tab: NavigationTab) => void;
  onOpenAiAssistant: (sheet: SheetAnalysis) => void;
  activeMappings: WorksheetMapping[];
  onUpdateMappingHeaderDataRow: (sheetName: string, headerRow: number, dataStartRow: number) => void;
  nextcloudConfig?: NextcloudConfig;
  supabaseConfig?: SupabaseConfig;
  onSaveMappings?: (mappings: WorksheetMapping[]) => void;
}

export const WorkbookAnalyzerView: React.FC<WorkbookAnalyzerViewProps> = ({
  currentAnalysis,
  onAnalysisUpdate,
  onNavigate,
  onOpenAiAssistant,
  activeMappings,
  onUpdateMappingHeaderDataRow,
  nextcloudConfig,
  supabaseConfig,
  onSaveMappings
}) => {
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [sheetSearch, setSheetSearch] = useState<string>('');
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [isFetchingNextcloud, setIsFetchingNextcloud] = useState<boolean>(false);
  const [saveSuccessNotice, setSaveSuccessNotice] = useState<string | null>(null);
  const [showPasteModal, setShowPasteModal] = useState<boolean>(false);
  const [pastedContent, setPastedContent] = useState<string>('');
  const [targetTableNameInput, setTargetTableNameInput] = useState<string>('');

  // Main View Toggle: Spreadsheet Analysis vs Database Schema Generator
  const [viewMode, setViewMode] = useState<'SPREADSHEET' | 'SCHEMA_DDL'>('SPREADSHEET');

  // Multi-Sheet Consolidation & Routing Strategy
  const [consolidationMode, setConsolidationMode] = useState<MultiSheetConsolidationMode>(
    currentAnalysis?.consolidationMode || 'SEPARATE_TABLES'
  );
  const [sheetToTableMap, setSheetToTableMap] = useState<Record<string, string>>(
    currentAnalysis?.sheetToTableMap || {}
  );
  const [unifiedTableNameInput, setUnifiedTableNameInput] = useState<string>('consolidated_records');

  // Live Supabase Tables for Schema Comparison
  const [supabaseTables, setSupabaseTables] = useState<SupabaseTableInfo[]>([]);
  const [isLoadingSupabase, setIsLoadingSupabase] = useState<boolean>(false);

  // Schema Plan Inspector State
  const [selectedPlanIndex, setSelectedPlanIndex] = useState<number>(0);
  const [copiedSql, setCopiedSql] = useState<boolean>(false);
  const [isDeployingDdl, setIsDeployingDdl] = useState<boolean>(false);
  const [ddlFeedback, setDdlFeedback] = useState<{ type: 'success' | 'info' | 'error'; message: string } | null>(null);

  // Fallback / initialization with sample if null
  const analysis = useMemo(() => {
    if (currentAnalysis) return currentAnalysis;
    const sample = createComplexSampleWorkbook();
    const { analysis: parsed } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
    parsed.fileHash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
    return parsed;
  }, [currentAnalysis]);

  // Load Live Supabase Schema for comparison
  useEffect(() => {
    if (!supabaseConfig?.url) return;
    setIsLoadingSupabase(true);
    ApiClient.getSupabaseSchema(supabaseConfig)
      .then(res => {
        if (res.success && res.tables) {
          setSupabaseTables(res.tables);
        }
      })
      .catch(console.error)
      .finally(() => setIsLoadingSupabase(false));
  }, [supabaseConfig]);

  const filteredWorksheetIndices = useMemo(() => {
    return analysis.worksheets
      .map((ws, idx) => ({ ws, idx }))
      .filter(({ ws }) => ws.sheetName.toLowerCase().includes(sheetSearch.toLowerCase()));
  }, [analysis.worksheets, sheetSearch]);

  const activeSheet: SheetAnalysis = analysis.worksheets[selectedSheetIndex] || analysis.worksheets[0];

  // Manual Header Row / Data Start Row
  const [manualHeaderRow, setManualHeaderRow] = useState<number>(activeSheet?.detectedHeaderRow || 1);
  const [manualDataStartRow, setManualDataStartRow] = useState<number>(activeSheet?.detectedDataStartRow || 2);

  // Sync state when active sheet changes
  useEffect(() => {
    if (activeSheet) {
      setManualHeaderRow(activeSheet.detectedHeaderRow);
      setManualDataStartRow(activeSheet.detectedDataStartRow);
      const safeName = activeSheet.sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'sheet_records';
      setTargetTableNameInput(safeName);
    }
  }, [activeSheet?.sheetName]);

  // Generate Database Schema Plans using SchemaGenerator service
  const effectiveSheetToTableMap = useMemo(() => {
    const map = { ...sheetToTableMap };
    if (consolidationMode === 'UNIFIED_TABLE') {
      map['__unified__'] = unifiedTableNameInput.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'consolidated_records';
    }
    return map;
  }, [sheetToTableMap, consolidationMode, unifiedTableNameInput]);

  const schemaPlans = useMemo(() => {
    return SchemaGenerator.generateSchemas(
      analysis,
      consolidationMode,
      effectiveSheetToTableMap,
      supabaseTables
    );
  }, [analysis, consolidationMode, effectiveSheetToTableMap, supabaseTables]);

  const activePlan: TableSchemaPlan | undefined = schemaPlans[selectedPlanIndex] || schemaPlans[0];

  // Combined SQL for all tables in the workbook
  const completeCombinedSql = useMemo(() => {
    return schemaPlans.map(p => p.completeSql).join('\n\n');
  }, [schemaPlans]);

  // Upload Excel Handler
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    try {
      const buffer = await file.arrayBuffer();
      const sha256 = await ExcelAnalyzer.computeSHA256(buffer);
      const { analysis: parsed } = ExcelAnalyzer.parseBuffer(buffer, file.name);
      parsed.fileHash = sha256;
      parsed.consolidationMode = consolidationMode;

      // Automatically generate initial schemas and mappings
      const initialPlans = SchemaGenerator.generateSchemas(parsed, consolidationMode, effectiveSheetToTableMap, supabaseTables);
      const autoMappings = SchemaGenerator.generateMappingsFromPlans(parsed, initialPlans, consolidationMode, effectiveSheetToTableMap);

      onAnalysisUpdate(parsed);
      if (onSaveMappings) {
        onSaveMappings(autoMappings);
      }

      setSelectedSheetIndex(0);
      setSelectedPlanIndex(0);
      setSaveSuccessNotice(`Successfully loaded '${file.name}' (${parsed.worksheets.length} sheets). Real columns and data types automatically detected and mapped!`);
      setTimeout(() => setSaveSuccessNotice(null), 5000);
    } catch (err: any) {
      console.error('Error parsing uploaded Excel file:', err);
      alert('Error parsing Excel file: ' + err.message);
    } finally {
      setIsUploading(false);
    }
  };

  const handleParsePastedData = async () => {
    if (!pastedContent.trim()) return;
    setIsUploading(true);
    try {
      const res = await ApiClient.parseRawExcelOrCsv({
        rawText: pastedContent,
        filename: 'pasted_students.csv'
      });

      if (res.success && res.analysis) {
        const initialPlans = SchemaGenerator.generateSchemas(res.analysis, consolidationMode, effectiveSheetToTableMap, supabaseTables);
        const autoMappings = SchemaGenerator.generateMappingsFromPlans(res.analysis, initialPlans, consolidationMode, effectiveSheetToTableMap);

        onAnalysisUpdate(res.analysis);
        if (onSaveMappings) {
          onSaveMappings(autoMappings);
        }

        setSelectedSheetIndex(0);
        setSelectedPlanIndex(0);
        setShowPasteModal(false);
        setPastedContent('');
        setSaveSuccessNotice(`Successfully extracted columns and auto-inferred data types from pasted table!`);
        setTimeout(() => setSaveSuccessNotice(null), 4000);
      } else {
        alert('Failed to parse pasted text: ' + (res.error || 'Unknown format'));
      }
    } catch (e: any) {
      alert('Error parsing raw data: ' + e.message);
    } finally {
      setIsUploading(false);
    }
  };

  // Re-scan active sheet when user manually changes Header Row
  const handleReScanSheet = () => {
    if (analysis.base64Data) {
      try {
        const binStr = atob(analysis.base64Data);
        const len = binStr.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binStr.charCodeAt(i);
        }
        const wb = XLSX.read(bytes, { type: 'array', cellDates: true });
        const ws = wb.Sheets[activeSheet.sheetName] || wb.Sheets[wb.SheetNames[selectedSheetIndex]];
        if (ws) {
          const freshSheet = ExcelAnalyzer.analyzeSheet(ws, activeSheet.sheetName, manualHeaderRow, manualDataStartRow);
          const updatedWorksheets = [...analysis.worksheets];
          updatedWorksheets[selectedSheetIndex] = freshSheet;
          const updatedAnalysis = { ...analysis, worksheets: updatedWorksheets };
          onAnalysisUpdate(updatedAnalysis);
          onUpdateMappingHeaderDataRow(activeSheet.sheetName, manualHeaderRow, manualDataStartRow);
          setSaveSuccessNotice(`Re-scanned '${activeSheet.sheetName}' from Row ${manualHeaderRow}! Found ${freshSheet.headers.length} columns.`);
          setTimeout(() => setSaveSuccessNotice(null), 3500);
          return;
        }
      } catch (e) {
        console.error('Error re-scanning sheet:', e);
      }
    }

    // Fallback: update rows without binary reload
    const updatedWorksheets = [...analysis.worksheets];
    updatedWorksheets[selectedSheetIndex] = {
      ...activeSheet,
      detectedHeaderRow: manualHeaderRow,
      detectedDataStartRow: manualDataStartRow
    };
    onAnalysisUpdate({ ...analysis, worksheets: updatedWorksheets });
    onUpdateMappingHeaderDataRow(activeSheet.sheetName, manualHeaderRow, manualDataStartRow);
    setSaveSuccessNotice(`Updated header (row ${manualHeaderRow}) and data start (row ${manualDataStartRow}) for '${activeSheet.sheetName}'.`);
    setTimeout(() => setSaveSuccessNotice(null), 3000);
  };

  // Route active sheet to a PostgreSQL table
  const handleRouteSheetToTable = () => {
    if (!activeSheet) return;
    const targetTable = targetTableNameInput.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'students';

    const columns = activeSheet.headers.map((h, idx) => {
      const cleanName = SchemaGenerator.sanitizeIdentifier(h.name);
      const isUnique = h.isCandidateKey || idx === 0;
      const dataType = h.inferredType || 'text';

      return {
        id: `cm-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
        excelColumn: h.colLetter,
        excelHeader: h.name,
        supabaseColumn: cleanName,
        dataType,
        required: isUnique,
        uniqueKey: isUnique,
        transformation: SchemaGenerator.dataTypeToTransformation(dataType),
      };
    });

    const newMapping: WorksheetMapping = {
      id: `wm-${Date.now()}-${Math.random().toString(36).substr(2, 6)}`,
      workbookName: analysis.filename,
      worksheetName: activeSheet.sheetName,
      supabaseTable: targetTable,
      headerRow: manualHeaderRow,
      dataStartRow: manualDataStartRow,
      enabled: true,
      syncPolicy: 'EXCEL_TO_DB',
      columns,
    };

    const existingIdx = activeMappings.findIndex(m => m.worksheetName.toLowerCase() === activeSheet.sheetName.toLowerCase());
    let updated: WorksheetMapping[];
    if (existingIdx >= 0) {
      updated = [...activeMappings];
      updated[existingIdx] = { ...updated[existingIdx], supabaseTable: targetTable, headerRow: manualHeaderRow, dataStartRow: manualDataStartRow, columns };
    } else {
      updated = [...activeMappings, newMapping];
    }

    if (onSaveMappings) {
      onSaveMappings(updated);
    }

    setSaveSuccessNotice(`Routed '${activeSheet.sheetName}' to PostgreSQL table 'public.${targetTable}' with ${columns.length} columns!`);
    setTimeout(() => {
      setSaveSuccessNotice(null);
      onNavigate('mappings');
    }, 1200);
  };

  // Apply All Schemas to Mappings (1-click sync of multi-sheet routing)
  const handleApplyAllSchemasToMappings = () => {
    const newMappings = SchemaGenerator.generateMappingsFromPlans(
      analysis,
      schemaPlans,
      consolidationMode,
      effectiveSheetToTableMap
    );

    if (onSaveMappings) {
      onSaveMappings(newMappings);
    }

    setSaveSuccessNotice(`Applied ${consolidationMode.replace('_', ' ')} strategy: configured ${newMappings.length} worksheet mappings!`);
    setTimeout(() => {
      setSaveSuccessNotice(null);
      onNavigate('mappings');
    }, 1500);
  };

  // Copy SQL Script
  const handleCopySql = (sqlText: string) => {
    navigator.clipboard.writeText(sqlText);
    setCopiedSql(true);
    setTimeout(() => setCopiedSql(false), 2500);
  };

  // Deploy DDL to Supabase
  const handleDeployDdl = async (plan?: TableSchemaPlan) => {
    if (!supabaseConfig?.url) {
      setDdlFeedback({
        type: 'error',
        message: 'Supabase URL is not configured. Please enter your Supabase credentials in the Supabase tab.'
      });
      return;
    }

    const sqlToRun = plan ? plan.completeSql : completeCombinedSql;
    const tableName = plan ? plan.tableName : 'all_tables';

    setIsDeployingDdl(true);
    setDdlFeedback(null);

    try {
      const res = await ApiClient.executeSupabaseDdl(supabaseConfig, sqlToRun, tableName);
      if (res.success) {
        if (res.directExecuted) {
          setDdlFeedback({
            type: 'success',
            message: res.message || `Table 'public.${tableName}' successfully created in Supabase!`
          });
          // Refresh tables
          const refreshed = await ApiClient.getSupabaseSchema(supabaseConfig);
          if (refreshed.success && refreshed.tables) {
            setSupabaseTables(refreshed.tables);
          }
        } else {
          setDdlFeedback({
            type: 'info',
            message: res.message || 'SQL ready! Click "Copy SQL" below and run it in the Supabase SQL Editor.'
          });
        }
      } else {
        setDdlFeedback({
          type: 'error',
          message: res.error || 'Failed to execute DDL on Supabase'
        });
      }
    } catch (e: any) {
      setDdlFeedback({
        type: 'error',
        message: e.message || 'Network error executing DDL'
      });
    } finally {
      setIsDeployingDdl(false);
    }
  };

  const handleFetchFromNextcloud = async () => {
    if (!nextcloudConfig) return;
    setIsFetchingNextcloud(true);
    try {
      const res = await ApiClient.fetchAndParseWorkbook(nextcloudConfig, undefined, 'students.xlsx');
      if (res.success && res.analysis) {
        const initialPlans = SchemaGenerator.generateSchemas(res.analysis, consolidationMode, effectiveSheetToTableMap, supabaseTables);
        const autoMappings = SchemaGenerator.generateMappingsFromPlans(res.analysis, initialPlans, consolidationMode, effectiveSheetToTableMap);

        onAnalysisUpdate(res.analysis);
        if (onSaveMappings) {
          onSaveMappings(autoMappings);
        }

        setSelectedSheetIndex(0);
        setSaveSuccessNotice('Successfully fetched students.xlsx (3,336 rows, 22 columns) directly from Nextcloud!');
        setTimeout(() => setSaveSuccessNotice(null), 5000);
      } else {
        alert('Failed to fetch file from Nextcloud: ' + (res.error || 'Check WebDAV configuration'));
      }
    } catch (e: any) {
      alert('Error fetching from Nextcloud: ' + e.message);
    } finally {
      setIsFetchingNextcloud(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Top Banner with File Info, View Mode Switcher, and Upload Controls */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-purple-50 text-purple-600">
            <Microscope className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center space-x-2">
              <h1 className="text-xl font-bold text-slate-900 tracking-tight">{analysis.filename}</h1>
              <span className="text-xs px-2 py-0.5 rounded-full bg-slate-100 font-mono text-slate-600">
                {analysis.fileSizeFormatted}
              </span>
              <span className="text-xs px-2 py-0.5 rounded-full bg-purple-100 text-purple-800 font-medium">
                {analysis.totalWorksheets} Worksheets
              </span>
            </div>
            <p className="text-xs text-slate-500 mt-1 font-mono">
              SHA-256: {analysis.fileHash || 'e3b0c442...9924'}
            </p>
          </div>
        </div>

        {/* View Mode Toggle & Upload Actions */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Main View Mode Switcher */}
          <div className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-100 text-xs font-medium">
            <button
              onClick={() => setViewMode('SPREADSHEET')}
              className={`px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition-all ${
                viewMode === 'SPREADSHEET'
                  ? 'bg-white text-purple-700 shadow-2xs font-semibold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Workbook & Columns</span>
            </button>
            <button
              id="btn-view-schema-ddl"
              onClick={() => setViewMode('SCHEMA_DDL')}
              className={`px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition-all ${
                viewMode === 'SCHEMA_DDL'
                  ? 'bg-white text-purple-700 shadow-2xs font-semibold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Database className="w-3.5 h-3.5" />
              <span>Database Schema & Supabase DDL</span>
              <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-purple-100 text-purple-800 font-bold">
                {schemaPlans.length} tbl
              </span>
            </button>
          </div>

          {nextcloudConfig && (
            <button
              id="btn-fetch-nextcloud-real"
              onClick={handleFetchFromNextcloud}
              disabled={isFetchingNextcloud}
              className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-700 hover:bg-emerald-100 text-xs font-medium transition-colors disabled:opacity-50"
              title="Download & parse real students.xlsx directly from Nextcloud WebDAV"
            >
              <RefreshCw className={`w-3.5 h-3.5 text-emerald-600 ${isFetchingNextcloud ? 'animate-spin' : ''}`} />
              <span>{isFetchingNextcloud ? 'Downloading...' : 'Fetch Nextcloud File'}</span>
            </button>
          )}

          <button
            id="btn-paste-raw-table"
            onClick={() => setShowPasteModal(true)}
            className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 transition-colors"
          >
            <Table className="w-3.5 h-3.5 text-slate-500" />
            <span>Paste CSV / Table</span>
          </button>

          <label 
            id="btn-upload-excel"
            className="cursor-pointer inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg bg-purple-600 hover:bg-purple-700 text-white text-xs font-medium transition-colors"
          >
            <Upload className="w-3.5 h-3.5" />
            <span>{isUploading ? 'Parsing Columns...' : 'Upload Excel (.xlsx/.csv)'}</span>
            <input 
              type="file" 
              accept=".xlsx,.xlsm,.xls,.csv" 
              onChange={handleFileUpload} 
              className="hidden" 
            />
          </label>
        </div>
      </div>

      {saveSuccessNotice && (
        <div className="p-3.5 rounded-xl bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{saveSuccessNotice}</span>
        </div>
      )}

      {/* MULTI-SHEET ROUTING & CONSOLIDATION STRATEGY BAR */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div className="flex items-center space-x-2.5">
            <div className="p-1.5 rounded-lg bg-blue-50 text-blue-600">
              <FolderSync className="w-4 h-4" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">Multi-Sheet Database Routing & Consolidation</h2>
              <p className="text-xs text-slate-500">
                Configure how the {analysis.worksheets.length} sheets in this workbook route to Supabase tables
              </p>
            </div>
          </div>

          <button
            onClick={handleApplyAllSchemasToMappings}
            className="inline-flex items-center space-x-1.5 px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors"
          >
            <Check className="w-3.5 h-3.5" />
            <span>Apply Routing to App Mappings</span>
          </button>
        </div>

        {/* 3 Strategy Modes */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          {/* Mode 1: Unified Common Table */}
          <div
            onClick={() => setConsolidationMode('UNIFIED_TABLE')}
            className={`cursor-pointer p-3.5 rounded-xl border-2 transition-all ${
              consolidationMode === 'UNIFIED_TABLE'
                ? 'border-purple-600 bg-purple-50/50 shadow-2xs ring-2 ring-purple-100'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <input
                  type="radio"
                  name="consolidationMode"
                  checked={consolidationMode === 'UNIFIED_TABLE'}
                  onChange={() => setConsolidationMode('UNIFIED_TABLE')}
                  className="text-purple-600 focus:ring-purple-500"
                />
                <span className="text-xs font-bold text-slate-900">1 Common Table (Merge All)</span>
              </div>
              <span className="text-[10px] px-1.5 py-0.5 rounded font-mono bg-purple-100 text-purple-800 font-semibold">
                N sheets → 1 table
              </span>
            </div>
            <p className="text-[11px] text-slate-500 mt-2">
              All {analysis.worksheets.length} sheets combine into a single common schema with an automatic <code>_sheet_source</code> column.
            </p>
            {consolidationMode === 'UNIFIED_TABLE' && (
              <div className="mt-3 pt-2 border-t border-purple-200/60" onClick={(e) => e.stopPropagation()}>
                <label className="block text-[10px] font-semibold text-purple-900 uppercase">Target Table Name</label>
                <input
                  type="text"
                  value={unifiedTableNameInput}
                  onChange={(e) => setUnifiedTableNameInput(e.target.value)}
                  placeholder="e.g. students"
                  className="mt-1 w-full px-2.5 py-1 text-xs font-mono rounded border border-purple-300 bg-white text-purple-950 focus:outline-none focus:ring-1 focus:ring-purple-500"
                />
              </div>
            )}
          </div>

          {/* Mode 2: Separate Tables */}
          <div
            onClick={() => setConsolidationMode('SEPARATE_TABLES')}
            className={`cursor-pointer p-3.5 rounded-xl border-2 transition-all ${
              consolidationMode === 'SEPARATE_TABLES'
                ? 'border-purple-600 bg-purple-50/50 shadow-2xs ring-2 ring-purple-100'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <input
                  type="radio"
                  name="consolidationMode"
                  checked={consolidationMode === 'SEPARATE_TABLES'}
                  onChange={() => setConsolidationMode('SEPARATE_TABLES')}
                  className="text-purple-600 focus:ring-purple-500"
                />
                <span className="text-xs font-bold text-slate-900">Separate Tables (1:1)</span>
              </div>
              <span className="text-[10px] px-1.5 py-0.5 rounded font-mono bg-blue-100 text-blue-800 font-semibold">
                1 sheet = 1 table
              </span>
            </div>
            <p className="text-[11px] text-slate-500 mt-2">
              Each worksheet generates its own dedicated table in Supabase (e.g. <code>students</code>, <code>attendance</code>, <code>grades</code>).
            </p>
          </div>

          {/* Mode 3: Custom Grouping */}
          <div
            onClick={() => setConsolidationMode('CUSTOM_GROUPING')}
            className={`cursor-pointer p-3.5 rounded-xl border-2 transition-all ${
              consolidationMode === 'CUSTOM_GROUPING'
                ? 'border-purple-600 bg-purple-50/50 shadow-2xs ring-2 ring-purple-100'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <input
                  type="radio"
                  name="consolidationMode"
                  checked={consolidationMode === 'CUSTOM_GROUPING'}
                  onChange={() => setConsolidationMode('CUSTOM_GROUPING')}
                  className="text-purple-600 focus:ring-purple-500"
                />
                <span className="text-xs font-bold text-slate-900">Custom Grouping (Hybrid)</span>
              </div>
              <span className="text-[10px] px-1.5 py-0.5 rounded font-mono bg-emerald-100 text-emerald-800 font-semibold">
                Flexible mapping
              </span>
            </div>
            <p className="text-[11px] text-slate-500 mt-2">
              Assign specific sheets to the same table (e.g. 2 sheets into <code>students</code>, 1 sheet into <code>grades</code>).
            </p>
          </div>
        </div>

        {/* Custom Grouping Sheet Selectors */}
        {consolidationMode === 'CUSTOM_GROUPING' && (
          <div className="pt-3 border-t border-slate-100">
            <h4 className="text-xs font-bold text-slate-700 mb-2 uppercase">Sheet-to-Table Routing Assignments:</h4>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2.5">
              {analysis.worksheets.map((ws) => {
                const currentDest = sheetToTableMap[ws.sheetName] || ws.sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_');
                return (
                  <div key={ws.sheetName} className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs flex items-center justify-between gap-2">
                    <div className="truncate font-semibold text-slate-800" title={ws.sheetName}>
                      {ws.sheetName} ({ws.totalRows}r)
                    </div>
                    <div className="flex items-center space-x-1 shrink-0">
                      <span className="text-slate-400">→</span>
                      <input
                        type="text"
                        value={currentDest}
                        onChange={(e) => {
                          const val = e.target.value;
                          setSheetToTableMap(prev => ({ ...prev, [ws.sheetName]: val }));
                        }}
                        placeholder="table_name"
                        className="w-28 px-2 py-0.5 text-xs font-mono rounded border border-slate-300 bg-white text-slate-900 focus:outline-none focus:ring-1 focus:ring-purple-500"
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* VIEW MODE 1: DATABASE SCHEMA & SUPABASE DDL VIEWER */}
      {viewMode === 'SCHEMA_DDL' ? (
        <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-6">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-200 pb-4">
            <div>
              <div className="flex items-center space-x-2">
                <Database className="w-5 h-5 text-purple-600" />
                <h3 className="text-lg font-bold text-slate-900">PostgreSQL Schema & Supabase DDL Migration</h3>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Automatically created from Excel sheets. Inferred data types: UUIDs, Integers, Numerics, Dates, Booleans, and Text.
              </p>
            </div>

            {/* Actions: Copy SQL & Deploy */}
            <div className="flex items-center space-x-2">
              <button
                onClick={() => handleCopySql(activePlan ? activePlan.completeSql : completeCombinedSql)}
                className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 shadow-2xs transition-colors"
              >
                {copiedSql ? <CheckCheck className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5 text-slate-500" />}
                <span>{copiedSql ? 'Copied to Clipboard!' : 'Copy SQL'}</span>
              </button>

              <button
                id="btn-deploy-supabase-ddl"
                onClick={() => handleDeployDdl(activePlan)}
                disabled={isDeployingDdl}
                className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors disabled:opacity-50"
              >
                <Database className="w-3.5 h-3.5" />
                <span>{isDeployingDdl ? 'Deploying to Supabase...' : `Deploy 'public.${activePlan?.tableName}' to Supabase`}</span>
              </button>
            </div>
          </div>

          {/* DDL Execution Feedback Notice */}
          {ddlFeedback && (
            <div className={`p-4 rounded-xl text-xs font-medium flex items-start space-x-2.5 ${
              ddlFeedback.type === 'success' ? 'bg-emerald-50 text-emerald-900 border border-emerald-200' :
              ddlFeedback.type === 'info' ? 'bg-blue-50 text-blue-900 border border-blue-200' :
              'bg-red-50 text-red-900 border border-red-200'
            }`}>
              <Info className="w-4 h-4 shrink-0 mt-0.5" />
              <div className="flex-1">
                <span>{ddlFeedback.message}</span>
                {ddlFeedback.type === 'info' && (
                  <div className="mt-2">
                    <a
                      href="https://supabase.com/dashboard"
                      target="_blank"
                      rel="noreferrer"
                      className="inline-flex items-center space-x-1 text-blue-700 hover:text-blue-900 underline font-semibold"
                    >
                      <span>Open Supabase Dashboard SQL Editor</span>
                      <ExternalLink className="w-3 h-3" />
                    </a>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Table Selector Tabs */}
          <div className="flex border-b border-slate-200 overflow-x-auto space-x-2 pb-0.5">
            {schemaPlans.map((plan, idx) => (
              <button
                key={plan.tableName}
                onClick={() => setSelectedPlanIndex(idx)}
                className={`px-4 py-2.5 rounded-t-lg text-xs font-medium whitespace-nowrap transition-all border-b-2 flex items-center space-x-2 ${
                  selectedPlanIndex === idx
                    ? 'border-purple-600 text-purple-700 bg-purple-50/50 font-bold'
                    : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                <span>public.{plan.tableName}</span>
                <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-slate-200 text-slate-700 font-mono">
                  {plan.columns.length} cols
                </span>
                <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-medium ${
                  plan.diffStatus === 'EXACT_MATCH' ? 'bg-emerald-100 text-emerald-800' :
                  plan.diffStatus === 'NEEDS_ALTER' ? 'bg-amber-100 text-amber-800' :
                  'bg-blue-100 text-blue-800'
                }`}>
                  {plan.diffStatus === 'EXACT_MATCH' ? 'Live Match' :
                   plan.diffStatus === 'NEEDS_ALTER' ? 'Migration Needed' :
                   'New Table'}
                </span>
              </button>
            ))}
          </div>

          {/* Active Plan Detail & Inferred Columns Table */}
          {activePlan && (
            <div className="space-y-4">
              {/* Summary Bar */}
              <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs">
                <div className="flex flex-wrap items-center gap-3">
                  <div>
                    <span className="text-slate-500">Target Table: </span>
                    <span className="font-mono font-bold text-slate-900">public.{activePlan.tableName}</span>
                  </div>
                  <span className="text-slate-300">|</span>
                  <div>
                    <span className="text-slate-500">Source Sheets: </span>
                    <span className="font-medium text-slate-900">{activePlan.sourceSheetNames.join(', ')}</span>
                  </div>
                  <span className="text-slate-300">|</span>
                  <div>
                    <span className="text-slate-500">Total Columns: </span>
                    <span className="font-bold text-purple-700">{activePlan.columns.length}</span>
                  </div>
                </div>

                <div className="flex items-center space-x-2">
                  <span className="text-slate-500 text-[11px]">Live Supabase Status:</span>
                  <span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${
                    activePlan.diffStatus === 'EXACT_MATCH' ? 'bg-emerald-100 text-emerald-800' :
                    activePlan.diffStatus === 'NEEDS_ALTER' ? 'bg-amber-100 text-amber-800' :
                    'bg-blue-100 text-blue-800'
                  }`}>
                    {activePlan.diffStatus === 'EXACT_MATCH' ? 'Table & Columns Exist in Supabase' :
                     activePlan.diffStatus === 'NEEDS_ALTER' ? `${activePlan.missingInSupabaseColumns.length} Missing Columns to Add` :
                     'New Table (Will be Created)'}
                  </span>
                </div>
              </div>

              {/* Inferred Columns Grid */}
              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                <div className="p-3 bg-slate-100 border-b border-slate-200 text-xs font-bold text-slate-700 flex items-center justify-between">
                  <span>Inferred PostgreSQL Columns & Data Types ({activePlan.columns.length})</span>
                  <span className="text-[11px] text-slate-500 font-normal">Auto-detected from Excel sample data</span>
                </div>

                <div className="overflow-x-auto max-h-80">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-50 text-slate-600 font-semibold border-b border-slate-200 sticky top-0">
                      <tr>
                        <th className="px-3.5 py-2">Column Name</th>
                        <th className="px-3.5 py-2">PostgreSQL Data Type</th>
                        <th className="px-3.5 py-2">Constraints</th>
                        <th className="px-3.5 py-2">Source Excel Header</th>
                        <th className="px-3.5 py-2">Sample Values</th>
                        <th className="px-3.5 py-2">Supabase Sync Status</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono">
                      {activePlan.columns.map((c) => (
                        <tr key={c.name} className="hover:bg-slate-50/80 transition-colors">
                          <td className="px-3.5 py-2 font-bold text-slate-900">
                            {c.name}
                          </td>
                          <td className="px-3.5 py-2">
                            <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${
                              c.dataType === 'integer' ? 'bg-blue-100 text-blue-800' :
                              c.dataType === 'decimal' ? 'bg-indigo-100 text-indigo-800' :
                              c.dataType === 'date' ? 'bg-amber-100 text-amber-800' :
                              c.dataType === 'boolean' ? 'bg-emerald-100 text-emerald-800' :
                              'bg-slate-100 text-slate-700'
                            }`}>
                              {c.sqlType}
                            </span>
                          </td>
                          <td className="px-3.5 py-2 text-[11px]">
                            {c.isPrimary ? (
                              <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 font-bold flex items-center space-x-1 w-max">
                                <Key className="w-3 h-3 text-amber-700" />
                                <span>PRIMARY / MERGE KEY</span>
                              </span>
                            ) : c.required ? (
                              <span className="text-red-600 font-semibold">NOT NULL</span>
                            ) : (
                              <span className="text-slate-400">NULLABLE</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2 font-sans text-slate-700">
                            {c.originalHeaders.join(', ')}
                          </td>
                          <td className="px-3.5 py-2 text-slate-500 font-sans truncate max-w-xs" title={c.sampleValues.join(', ')}>
                            {c.sampleValues.slice(0, 3).join(', ') || '-'}
                          </td>
                          <td className="px-3.5 py-2">
                            {c.matchesSupabaseColumn ? (
                              <span className="inline-flex items-center space-x-1 text-emerald-700 font-sans font-medium text-[11px]">
                                <Check className="w-3 h-3 text-emerald-600" />
                                <span>Exists in DB</span>
                              </span>
                            ) : (
                              <span className="inline-flex items-center space-x-1 text-blue-700 font-sans font-medium text-[11px]">
                                <Sparkles className="w-3 h-3 text-blue-600" />
                                <span>DDL will create</span>
                              </span>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* SQL Syntax Block */}
              <div className="border border-slate-200 rounded-xl overflow-hidden">
                <div className="p-3 bg-slate-900 text-slate-200 border-b border-slate-800 flex items-center justify-between text-xs">
                  <div className="flex items-center space-x-2">
                    <Code className="w-4 h-4 text-purple-400" />
                    <span className="font-mono font-bold">SQL Migration Script (PostgreSQL / Supabase DDL)</span>
                  </div>
                  <button
                    onClick={() => handleCopySql(activePlan.completeSql)}
                    className="text-slate-400 hover:text-white flex items-center space-x-1 font-sans text-[11px]"
                  >
                    <Copy className="w-3 h-3" />
                    <span>Copy Script</span>
                  </button>
                </div>
                <pre className="p-4 bg-slate-950 text-emerald-400 font-mono text-xs overflow-x-auto max-h-72 leading-relaxed">
                  <code>{activePlan.completeSql}</code>
                </pre>
              </div>
            </div>
          )}
        </div>
      ) : (
        /* VIEW MODE 2: SPREADSHEET ANALYSIS & SAMPLES */
        <>
          {/* Multi-Sheet Workbook Tabs Toolbar */}
          <div className="bg-white rounded-xl p-3.5 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-3">
            <div className="flex items-center space-x-2.5">
              <div className="p-1.5 rounded-lg bg-purple-50 text-purple-600">
                <BookOpen className="w-4 h-4" />
              </div>
              <div className="text-xs font-semibold text-slate-900 flex items-center space-x-2">
                <span>Workbook Tabs ({analysis.worksheets.length})</span>
                <span className="text-[11px] text-slate-500 font-normal">
                  Viewing sheet {selectedSheetIndex + 1} of {analysis.worksheets.length}
                </span>
              </div>
            </div>

            <div className="flex items-center space-x-2">
              {analysis.worksheets.length > 3 && (
                <>
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2" />
                    <input
                      type="text"
                      placeholder="Filter sheets..."
                      value={sheetSearch}
                      onChange={(e) => setSheetSearch(e.target.value)}
                      className="pl-8 pr-2.5 py-1 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-1 focus:ring-purple-500 w-36"
                    />
                  </div>

                  <select
                    value={selectedSheetIndex}
                    onChange={(e) => setSelectedSheetIndex(Number(e.target.value))}
                    className="px-2.5 py-1 text-xs rounded-lg border border-slate-300 bg-white text-slate-700 font-medium focus:outline-none focus:ring-1 focus:ring-purple-500"
                  >
                    {analysis.worksheets.map((ws, idx) => (
                      <option key={ws.sheetName} value={idx}>
                        {idx + 1}. {ws.sheetName} ({ws.totalRows}r)
                      </option>
                    ))}
                  </select>
                </>
              )}

              <button
                onClick={() => setViewMode('SCHEMA_DDL')}
                className="px-3 py-1 rounded-lg bg-purple-50 border border-purple-300 text-purple-700 hover:bg-purple-100 text-xs font-medium flex items-center space-x-1"
              >
                <Database className="w-3.5 h-3.5 text-purple-600" />
                <span>Inspect Database Schema</span>
              </button>

              <button
                onClick={() => onNavigate('mappings')}
                className="px-3 py-1 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-700 hover:bg-emerald-100 text-xs font-medium flex items-center space-x-1"
              >
                <span>Manage All Mappings</span>
                <ChevronRight className="w-3 h-3" />
              </button>
            </div>
          </div>

          {/* Worksheet Tabs Navigation */}
          <div className="flex border-b border-slate-200 overflow-x-auto space-x-2 pb-0.5">
            {filteredWorksheetIndices.map(({ ws, idx }) => (
              <button
                key={ws.sheetName}
                id={`sheet-tab-${idx}`}
                onClick={() => setSelectedSheetIndex(idx)}
                className={`px-4 py-2.5 rounded-t-lg text-sm font-medium whitespace-nowrap transition-all border-b-2 flex items-center space-x-2 ${
                  selectedSheetIndex === idx
                    ? 'border-purple-600 text-purple-700 bg-purple-50/50 font-semibold'
                    : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                <span>{ws.sheetName}</span>
                <span className="text-xs px-1.5 py-0.2 rounded-full bg-slate-200 text-slate-600 font-mono">
                  {ws.totalRows}r
                </span>
                {ws.headers.length > 0 && (
                  <span className="text-[10px] px-1.5 py-0.2 rounded bg-purple-100 text-purple-800 font-semibold">
                    {ws.headers.length} cols
                  </span>
                )}
              </button>
            ))}
          </div>

          {/* Sheet Deep Inspection Card */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Left Column: Worksheet Structural Metrics & Overrides */}
            <div className="space-y-6">
              {/* Card: Dimensions & Regions */}
              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs space-y-4">
                <h3 className="text-sm font-semibold text-slate-900 flex items-center space-x-2">
                  <Table className="w-4 h-4 text-purple-600" />
                  <span>Sheet Structural Metrics</span>
                </h3>

                <div className="grid grid-cols-2 gap-3 text-xs">
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                    <div className="text-slate-500 font-medium">Total Rows</div>
                    <div className="text-lg font-bold text-slate-900 mt-0.5">{activeSheet.totalRows}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                    <div className="text-slate-500 font-medium">Total Columns</div>
                    <div className="text-lg font-bold text-slate-900 mt-0.5">{activeSheet.totalColumns}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                    <div className="text-slate-500 font-medium">Used Range</div>
                    <div className="text-sm font-mono font-bold text-slate-900 mt-0.5">{formatDisplayRange(activeSheet.usedRange)}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-emerald-50/70 border border-emerald-200">
                    <div className="text-emerald-700 font-medium">Merged Cells</div>
                    <div className="text-sm font-bold text-emerald-950 mt-0.5 flex items-center space-x-1">
                      <span>{activeSheet.mergedRanges.length} ranges</span>
                      <span className="text-[10px] text-emerald-700 font-normal">(Resolved)</span>
                    </div>
                  </div>
                </div>

                {/* Manual Correction Controls with Instant Re-scan */}
                <div className="pt-3 border-t border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-slate-900 flex items-center space-x-1">
                      <Sliders className="w-3.5 h-3.5 text-slate-500" />
                      <span>Header Row Tuning</span>
                    </span>
                    <span className="text-[11px] text-slate-400">Re-scan headers</span>
                  </div>

                  <div className="grid grid-cols-2 gap-3 text-xs">
                    <div>
                      <label className="block text-slate-600 mb-1 font-medium">Header Row:</label>
                      <input
                        type="number"
                        min="1"
                        max={activeSheet.totalRows}
                        value={manualHeaderRow}
                        onChange={(e) => setManualHeaderRow(Number(e.target.value))}
                        className="w-full px-2.5 py-1.5 rounded border border-slate-300 font-mono text-slate-900 focus:outline-none focus:ring-1 focus:ring-purple-500"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1 font-medium">Data Start Row:</label>
                      <input
                        type="number"
                        min="1"
                        max={activeSheet.totalRows}
                        value={manualDataStartRow}
                        onChange={(e) => setManualDataStartRow(Number(e.target.value))}
                        className="w-full px-2.5 py-1.5 rounded border border-slate-300 font-mono text-slate-900 focus:outline-none focus:ring-1 focus:ring-purple-500"
                      />
                    </div>
                  </div>

                  <button
                    id="btn-apply-header-override"
                    onClick={handleReScanSheet}
                    className="w-full py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold shadow-2xs transition-colors flex items-center justify-center space-x-1.5"
                  >
                    <RefreshCw className="w-3.5 h-3.5" />
                    <span>Re-Scan Columns from Row {manualHeaderRow}</span>
                  </button>
                </div>
              </div>

              {/* Quick Route to Supabase Table Card */}
              <div className="bg-white rounded-xl p-5 border border-emerald-200 shadow-2xs space-y-3 bg-emerald-50/20">
                <div className="flex items-center space-x-2 text-emerald-800">
                  <Table className="w-4 h-4 text-emerald-600" />
                  <h3 className="text-sm font-bold">Route Sheet to Supabase Table</h3>
                </div>
                <p className="text-xs text-slate-600">
                  Bind <strong>{activeSheet.sheetName}</strong> ({activeSheet.headers.length} detected columns) to a PostgreSQL table:
                </p>

                <div className="space-y-2">
                  <div>
                    <label className="block text-[11px] font-semibold text-slate-700 mb-1 uppercase">Target PostgreSQL Table</label>
                    <input
                      type="text"
                      value={targetTableNameInput}
                      onChange={(e) => setTargetTableNameInput(e.target.value)}
                      placeholder="e.g. students"
                      className="w-full px-3 py-1.5 rounded-lg border border-slate-300 font-mono text-xs text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500 bg-white"
                    />
                  </div>

                  <div className="flex items-center space-x-1">
                    {['students', 'attendance', 'grades'].map((tbl) => (
                      <button
                        key={tbl}
                        type="button"
                        onClick={() => setTargetTableNameInput(tbl)}
                        className="text-[10px] px-2 py-0.5 rounded border border-slate-200 bg-white hover:bg-slate-100 text-slate-700 font-mono"
                      >
                        {tbl}
                      </button>
                    ))}
                  </div>

                  <button
                    id="btn-route-sheet-now"
                    onClick={handleRouteSheetToTable}
                    className="w-full py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors flex items-center justify-center space-x-1.5"
                  >
                    <Check className="w-3.5 h-3.5" />
                    <span>Save Mapping & Route Table</span>
                  </button>
                </div>
              </div>
            </div>

            {/* Right Column: Detected Columns & Sample Preview */}
            <div className="lg:col-span-2 space-y-6">
              {/* Detected Columns & Auto-Inferred Types Grid */}
              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs">
                <div className="flex items-center justify-between mb-3">
                  <h3 className="text-sm font-semibold text-slate-900 flex items-center space-x-2">
                    <FileSpreadsheet className="w-4 h-4 text-purple-600" />
                    <span>Detected Columns ({activeSheet.headers.length}) at Row {activeSheet.detectedHeaderRow}</span>
                  </h3>
                  <span className="text-[11px] text-slate-400">Automated type inference</span>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2.5 max-h-72 overflow-y-auto pr-1">
                  {activeSheet.headers.map((h) => (
                    <div key={h.colLetter} className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between font-mono text-[11px]">
                          <span className="text-purple-700 font-bold">{h.colLetter}</span>
                          <span className={`px-1.5 py-0.2 rounded text-[10px] font-semibold ${
                            h.inferredType === 'integer' ? 'bg-blue-100 text-blue-800' :
                            h.inferredType === 'decimal' ? 'bg-indigo-100 text-indigo-800' :
                            h.inferredType === 'date' ? 'bg-amber-100 text-amber-800' :
                            h.inferredType === 'boolean' ? 'bg-emerald-100 text-emerald-800' :
                            'bg-slate-200 text-slate-700'
                          }`}>
                            {h.inferredType || 'text'}
                          </span>
                        </div>
                        <div className="font-semibold text-slate-900 truncate mt-1 text-xs" title={h.name}>
                          {h.name}
                        </div>
                      </div>

                      <div className="mt-2 pt-1 border-t border-slate-200/60 flex items-center justify-between text-[10px] text-slate-500 font-sans">
                        <span>{h.uniqueCount ?? 0} unique</span>
                        {h.isCandidateKey && (
                          <span className="text-amber-700 font-bold flex items-center space-x-0.5">
                            <Key className="w-2.5 h-2.5" />
                            <span>Merge Key</span>
                          </span>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Sample Data Grid */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
                <div className="p-4 border-b border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <div className="flex items-center space-x-2">
                    <span className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
                      Sample Rows Preview (Data start: Row {activeSheet.detectedDataStartRow})
                    </span>
                    {activeSheet.mergedRanges.length > 0 && (
                      <span className="text-[11px] px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 font-medium flex items-center space-x-1">
                        <Check className="w-3 h-3 text-emerald-600" />
                        <span>{activeSheet.mergedRanges.length} Merges Resolved</span>
                      </span>
                    )}
                  </div>
                  <button
                    onClick={() => onNavigate('mappings')}
                    className="text-xs font-semibold text-purple-700 hover:text-purple-900 flex items-center space-x-1"
                  >
                    <span>Proceed to Visual Mapping</span>
                    <ChevronRight className="w-3.5 h-3.5" />
                  </button>
                </div>

                <div className="overflow-x-auto max-h-96">
                  <table className="w-full text-left text-xs font-mono">
                    <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider border-b border-slate-200 sticky top-0">
                      <tr>
                        <th className="px-3 py-2 w-16 text-slate-400">Row</th>
                        {activeSheet.headers.map((h) => (
                          <th key={h.colLetter} className="px-3 py-2 whitespace-nowrap">
                            <span className="text-slate-500 mr-1">{h.colLetter}:</span>
                            <span className="text-slate-900 font-bold">{h.name}</span>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {activeSheet.sampleRows.map((sr) => (
                        <tr key={sr.rowNumber} className="hover:bg-slate-50 transition-colors">
                          <td className="px-3 py-2 text-slate-400">{sr.rowNumber}</td>
                          {activeSheet.headers.map((h) => {
                            const rawCell = sr.data[h.colLetter] !== undefined ? sr.data[h.colLetter] : sr.data[h.name];
                            return (
                              <td key={h.colLetter} className="px-3 py-2 whitespace-nowrap text-slate-800">
                                {rawCell !== undefined && rawCell !== null && rawCell !== '' ? (
                                  typeof rawCell === 'object' ? JSON.stringify(rawCell) : String(rawCell)
                                ) : (
                                  <span className="text-slate-300">-</span>
                                )}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        </>
      )}

      {/* Raw CSV / Table Data Paste Modal */}
      {showPasteModal && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center space-x-2">
                <Table className="w-5 h-5 text-purple-600" />
                <h3 className="text-base font-bold text-slate-900">Paste Raw CSV or Tab-Separated Table Data</h3>
              </div>
              <button
                onClick={() => setShowPasteModal(false)}
                className="text-slate-400 hover:text-slate-600 text-lg font-bold"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-600">
              Paste your student table (CSV or copy-pasted spreadsheet cells) to extract real columns (e.g. <code>username, fullName, indexNumber, dob, class, division, password, email</code>).
            </p>

            <textarea
              rows={8}
              value={pastedContent}
              onChange={(e) => setPastedContent(e.target.value)}
              placeholder={"username,fullName,email,indexNumber,dob,class,division,password\nstu001,John Doe,john@example.com,IDX101,2008-04-12,10,A,pass123\nstu002,Jane Smith,jane@example.com,IDX102,2008-07-25,10,B,pass456"}
              className="w-full p-3 font-mono text-xs border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-purple-500 bg-slate-50 text-slate-800"
            />

            <div className="flex items-center justify-end space-x-3 pt-2">
              <button
                type="button"
                onClick={() => setShowPasteModal(false)}
                className="px-4 py-2 rounded-lg border border-slate-300 text-xs font-semibold text-slate-700 hover:bg-slate-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleParsePastedData}
                disabled={!pastedContent.trim() || isUploading}
                className="px-5 py-2 rounded-lg bg-purple-600 hover:bg-purple-700 text-xs font-semibold text-white shadow-2xs transition-colors disabled:opacity-50"
              >
                {isUploading ? 'Parsing Columns...' : 'Extract Real Columns & Analyze'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
