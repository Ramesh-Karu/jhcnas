import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  PlaySquare, 
  Play, 
  CheckCircle2, 
  AlertTriangle, 
  RefreshCw, 
  Layers, 
  ArrowRight, 
  Database, 
  FileSpreadsheet, 
  Download,
  AlertCircle,
  GitMerge,
  Sparkles,
  ShieldAlert,
  Wand2,
  Check,
  Copy,
  Info,
  Upload,
  Search,
  ChevronLeft,
  ChevronRight,
  Filter,
  FileCheck
} from 'lucide-react';
import { 
  WorkbookAnalysis, 
  WorksheetMapping, 
  DryRunResult, 
  RowValidationError, 
  ImportLog, 
  NavigationTab,
  SupabaseConfig,
  NextcloudFile,
  NextcloudConfig
} from '../types';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
import { DryRunEngine } from '../services/dryRunEngine';
import { ApiClient } from '../services/apiClient';
import { ExcelAnalyzer } from '../services/excelAnalyzer';

interface ImportDryRunViewProps {
  currentAnalysis: WorkbookAnalysis | null;
  mappings: WorksheetMapping[];
  databaseState: Record<string, any[]>;
  onUpdateDatabase: (tableName: string, newRecords: any[]) => void;
  onAddImportLog: (log: ImportLog, errors: RowValidationError[]) => void;
  onNavigate: (tab: NavigationTab) => void;
  supabaseConfig?: SupabaseConfig;
  onSaveMappings?: (mappings: WorksheetMapping[]) => void;
  files?: NextcloudFile[];
  nextcloudConfig?: NextcloudConfig;
  onSelectFileForAnalysis?: (file: NextcloudFile) => void;
  onAnalysisUpdate?: (analysis: WorkbookAnalysis) => void;
}

export const ImportDryRunView: React.FC<ImportDryRunViewProps> = ({
  currentAnalysis,
  mappings,
  databaseState,
  onUpdateDatabase,
  onAddImportLog,
  onNavigate,
  supabaseConfig,
  onSaveMappings,
  files = [],
  nextcloudConfig,
  onSelectFileForAnalysis,
  onAnalysisUpdate
}) => {
  const [dryRunResult, setDryRunResult] = useState<DryRunResult | null>(null);
  const [isRunningDryRun, setIsRunningDryRun] = useState<boolean>(false);
  const [isImporting, setIsImporting] = useState<boolean>(false);
  const [isLoadingFile, setIsLoadingFile] = useState<boolean>(false);
  const [importNotice, setImportNotice] = useState<{ success: boolean; message: string } | null>(null);
  const [filterAction, setFilterAction] = useState<'ALL' | 'INSERT' | 'UPDATE' | 'FAIL'>('ALL');
  const [selectedSheetFilter, setSelectedSheetFilter] = useState<string>('ALL');
  const [selectedErrorType, setSelectedErrorType] = useState<string>('ALL');
  const [copiedAlterSql, setCopiedAlterSql] = useState<boolean>(false);
  const [searchRecordQuery, setSearchRecordQuery] = useState<string>('');
  const [recordsPage, setRecordsPage] = useState<number>(1);
  const [recordsPerPage, setRecordsPerPage] = useState<number>(50);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [typeMismatchAlert, setTypeMismatchAlert] = useState<{
    tableName: string;
    columnName: string;
    offendingValue: string;
    alterSql: string;
    isFixing?: boolean;
  } | null>(null);

  const handleLocalFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsLoadingFile(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const { analysis } = ExcelAnalyzer.parseBuffer(arrayBuffer, file.name);
      analysis.fileHash = await ExcelAnalyzer.computeSHA256(arrayBuffer);
      if (onAnalysisUpdate) {
        onAnalysisUpdate(analysis);
      }
      setImportNotice({
        success: true,
        message: `Loaded '${file.name}' (${analysis.totalWorksheets} sheets, ${analysis.worksheets.reduce((acc, w) => acc + (w.sampleRows?.length || 0), 0)} rows). Running simulation...`
      });
    } catch (err: any) {
      setImportNotice({
        success: false,
        message: `Failed to parse '${file.name}': ${err.message}`
      });
    } finally {
      setIsLoadingFile(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  const handleSelectNextcloudFile = async (filename: string) => {
    if (!files || files.length === 0) return;
    const targetFile = files.find(f => f.filename === filename);
    if (!targetFile) return;

    setIsLoadingFile(true);
    try {
      if (nextcloudConfig?.url && nextcloudConfig.isConnected) {
        const res = await ApiClient.fetchAndParseWorkbook(nextcloudConfig, targetFile.path, targetFile.filename);
        if (res.success && res.analysis && onAnalysisUpdate) {
          onAnalysisUpdate(res.analysis);
          setImportNotice({
            success: true,
            message: `Loaded live Nextcloud file '${targetFile.filename}'. Running simulation...`
          });
          return;
        }
      }
      if (onSelectFileForAnalysis) {
        onSelectFileForAnalysis(targetFile);
      }
    } catch (err: any) {
      console.error('Failed to load remote file:', err);
    } finally {
      setIsLoadingFile(false);
    }
  };

  const getEffectiveWorkbook = (): XLSX.WorkBook => {
    if (currentAnalysis?.base64Data) {
      try {
        const binStr = atob(currentAnalysis.base64Data);
        const len = binStr.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binStr.charCodeAt(i);
        }
        return XLSX.read(bytes, { type: 'array', cellDates: true });
      } catch (e) {
        console.warn('Could not parse base64Data as workbook:', e);
      }
    }

    // If analysis has worksheets with parsed sample/all rows, synthesize workbook from all rows
    if (currentAnalysis?.worksheets && currentAnalysis.worksheets.length > 0) {
      const syntheticWb = XLSX.utils.book_new();
      let hasAnySheet = false;
      for (const ws of currentAnalysis.worksheets) {
        if (ws.sampleRows && ws.sampleRows.length > 0) {
          hasAnySheet = true;
          const rows: any[][] = [];
          const headerNames = ws.headers.map(h => h.name || h.colLetter);
          rows.push(headerNames);
          for (const sr of ws.sampleRows) {
            const r: any[] = [];
            for (const h of ws.headers) {
              r.push(sr.data[h.colLetter] ?? sr.data[h.name] ?? '');
            }
            rows.push(r);
          }
          const sheet = XLSX.utils.aoa_to_sheet(rows);
          XLSX.utils.book_append_sheet(syntheticWb, sheet, ws.sheetName);
        }
      }
      if (hasAnySheet) {
        return syntheticWb;
      }
    }

    const sample = createComplexSampleWorkbook();
    return sample.workbook;
  };

  const handleRunDryRun = async () => {
    setIsRunningDryRun(true);
    setImportNotice(null);

    try {
      const wb = getEffectiveWorkbook();
      const filename = currentAnalysis?.filename || mappings[0]?.workbookName || 'students.xlsx';

      // Synchronize latest existing records from Supabase in parallel if connected
      const mergedDbState: Record<string, any[]> = { ...databaseState };
      if (supabaseConfig?.url && (supabaseConfig.anonKey || supabaseConfig.serviceKey || supabaseConfig.serviceRoleKey)) {
        try {
          const uniqueTables = Array.from(new Set(mappings.map(m => m.supabaseTable)));
          await Promise.allSettled(
            uniqueTables.map(async (tbl) => {
              const res = await ApiClient.fetchSupabaseTableRows(supabaseConfig, tbl, 50);
              if (res.success && res.rows && res.rows.length > 0) {
                mergedDbState[tbl] = res.rows;
                onUpdateDatabase(tbl, res.rows);
              }
            })
          );
        } catch (e) {
          console.warn('Supabase fetch prior to dry-run notice:', e);
        }
      }

      const result = DryRunEngine.executeDryRun(
        wb,
        filename,
        mappings,
        mergedDbState,
        currentAnalysis
      );

      await new Promise(r => setTimeout(r, 150));
      setDryRunResult(result);
    } catch (err: any) {
      console.error('Dry run error:', err);
    } finally {
      setIsRunningDryRun(false);
    }
  };

  // Switch all mappings to direct INSERT mode (clear all merge keys to bypass conflict validation errors)
  const handleSwitchToInsertOnly = () => {
    if (!onSaveMappings) return;
    const updated = mappings.map(m => ({
      ...m,
      columns: m.columns.map(c => ({ ...c, uniqueKey: false }))
    }));
    onSaveMappings(updated);
    setImportNotice({
      success: true,
      message: 'All worksheets switched to Direct INSERT mode. Re-running simulation with zero key conflicts...'
    });
    setTimeout(() => {
      handleRunDryRun();
    }, 150);
  };

  // Run dry run automatically on mount so user never sees empty screen
  useEffect(() => {
    handleRunDryRun();
  }, [currentAnalysis?.filename, mappings.length]);

  const handleExecuteImport = async () => {
    setIsImporting(true);
    setImportNotice(null);

    const wb = getEffectiveWorkbook();
    const filename = currentAnalysis?.filename || mappings[0]?.workbookName || 'students.xlsx';
    const fileHash = currentAnalysis?.fileHash || 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';

    // Capture updated records locally so they can be dispatched immediately to Supabase
    const capturedDbRecords: Record<string, any[]> = {};
    const trackingUpdateDatabase = (tbl: string, recs: any[]) => {
      capturedDbRecords[tbl] = recs;
      onUpdateDatabase(tbl, recs);
    };

    const { log, errors } = DryRunEngine.executeRealImport(
      wb,
      filename,
      fileHash,
      mappings,
      databaseState,
      trackingUpdateDatabase,
      currentAnalysis
    );

    // If Supabase is connected, sanitize and write clean records directly to Supabase PostgREST
    let supabaseWriteNotice = '';
    let supabaseSuccessCount = 0;
    const supabaseErrors: string[] = [];

    if (supabaseConfig?.url && (supabaseConfig.anonKey || supabaseConfig.serviceKey || supabaseConfig.serviceRoleKey)) {
      try {
        for (const wm of mappings) {
          if (!wm.enabled) continue;
          const rawRecords = capturedDbRecords[wm.supabaseTable] || databaseState[wm.supabaseTable] || [];
          if (rawRecords.length > 0) {
            // ONLY send mapped Supabase columns to avoid PostgREST rejecting extra metadata
            const allowedCols = new Set(wm.columns.map(c => c.supabaseColumn));
            if (wm.sectionHeadingTargetCol) {
              allowedCols.add(wm.sectionHeadingTargetCol);
            }

            const cleanRecordsForSupabase = rawRecords.map(r => {
              const clean: Record<string, any> = {};
              for (const col of allowedCols) {
                if (r[col] !== undefined) {
                  clean[col] = r[col];
                }
              }
              return clean;
            }).filter(r => Object.keys(r).length > 0);

            if (cleanRecordsForSupabase.length === 0) continue;

            const uniqueKeyCol = wm.columns.find(c => c.uniqueKey)?.supabaseColumn;
            const upRes = await ApiClient.upsertSupabaseRecords(
              supabaseConfig,
              wm.supabaseTable,
              cleanRecordsForSupabase,
              uniqueKeyCol
            );

            if (upRes.success) {
              supabaseSuccessCount += cleanRecordsForSupabase.length;
              if ((upRes as any).warning) {
                supabaseWriteNotice += ` Synced ${cleanRecordsForSupabase.length} rows to '${wm.supabaseTable}'. ${(upRes as any).warning}`;
              } else {
                supabaseWriteNotice += ` Synced ${cleanRecordsForSupabase.length} rows to '${wm.supabaseTable}'.`;
              }
              // Fetch latest live rows from Supabase to update local UI databaseState
              const latest = await ApiClient.fetchSupabaseTableRows(supabaseConfig, wm.supabaseTable, 100);
              if (latest.success && latest.rows) {
                onUpdateDatabase(wm.supabaseTable, latest.rows);
              }
            } else {
              const errStr = upRes.error || '';
              supabaseErrors.push(`${wm.supabaseTable}: ${errStr || 'Check table schema'}`);

              // Detect PostgreSQL 22P02: invalid input syntax for type integer: "GRADE 07 A"
              if (errStr.includes('invalid input syntax for type integer') || errStr.includes('22P02')) {
                const intMatch = errStr.match(/invalid input syntax for type integer:\s*"([^"]+)"/i);
                const offending = intMatch ? intMatch[1] : 'GRADE 07 A';

                // Find culprit column
                let culpritCol = '';
                for (const r of cleanRecordsForSupabase) {
                  for (const [k, v] of Object.entries(r)) {
                    if (String(v).trim() === offending.trim()) {
                      culpritCol = k;
                      break;
                    }
                  }
                  if (culpritCol) break;
                }
                if (!culpritCol) {
                  culpritCol = wm.sectionHeadingTargetCol || 'class';
                }

                setTypeMismatchAlert({
                  tableName: wm.supabaseTable,
                  columnName: culpritCol,
                  offendingValue: offending,
                  alterSql: `ALTER TABLE public.${wm.supabaseTable} ALTER COLUMN ${culpritCol} TYPE text;`
                });
              }
            }
          }
        }
      } catch (err: any) {
        supabaseErrors.push(err.message);
      }
    }

    await new Promise(r => setTimeout(r, 400));

    onAddImportLog(log, errors);
    setIsImporting(false);

    if (supabaseErrors.length > 0) {
      setImportNotice({
        success: false,
        message: `Import processed locally, but Supabase reported issues: ${supabaseErrors.join(' | ')}. Please check the Supabase Schema tab or use the 1-Click fix below.`
      });
    } else {
      setImportNotice({
        success: true,
        message: `Production Import completed! ${log.rowsInserted} records inserted, ${log.rowsUpdated} records merged/updated, ${log.rowsFailed} failed.${supabaseWriteNotice || (supabaseSuccessCount > 0 ? ` (Synced ${supabaseSuccessCount} rows to Supabase)` : '')}`
      });
    }

    // Re-run dry run to show updated status
    const updatedResult = DryRunEngine.executeDryRun(wb, filename, mappings, capturedDbRecords, currentAnalysis);
    setDryRunResult(updatedResult);
  };

  // Convert column type to text in Supabase with 1-click DDL
  const handleFixTypeMismatchByAlterSql = async () => {
    if (!typeMismatchAlert || !supabaseConfig) return;
    setTypeMismatchAlert(prev => prev ? { ...prev, isFixing: true } : null);
    try {
      const ddlRes = await ApiClient.executeSupabaseDdl(
        supabaseConfig,
        typeMismatchAlert.alterSql,
        typeMismatchAlert.tableName
      );

      if (ddlRes.success && ddlRes.directExecuted) {
        setImportNotice({
          success: true,
          message: `Schema updated! Successfully executed: ${typeMismatchAlert.alterSql}. Retrying import to Supabase now...`
        });
        setTypeMismatchAlert(null);
        setTimeout(() => {
          handleExecuteImport();
        }, 500);
      } else {
        setImportNotice({
          success: false,
          message: `Execute SQL in Supabase SQL Editor: ${typeMismatchAlert.alterSql} (Copy the command below).`
        });
        setTypeMismatchAlert(prev => prev ? { ...prev, isFixing: false } : null);
      }
    } catch (e: any) {
      setTypeMismatchAlert(prev => prev ? { ...prev, isFixing: false } : null);
    }
  };

  // Auto-extract numbers only (converts "GRADE 07 A" -> 7)
  const handleFixTypeMismatchByExtractNumber = () => {
    if (!typeMismatchAlert || !onSaveMappings) return;
    const colName = typeMismatchAlert.columnName;
    const updated = mappings.map(m => {
      if (m.supabaseTable.toLowerCase() === typeMismatchAlert.tableName.toLowerCase()) {
        return {
          ...m,
          columns: m.columns.map(c => {
            if (c.supabaseColumn.toLowerCase() === colName.toLowerCase()) {
              return { ...c, dataType: 'integer' as const, transformation: 'parse_number' as const };
            }
            return c;
          })
        };
      }
      return m;
    });
    onSaveMappings(updated);
    setTypeMismatchAlert(null);
    setImportNotice({
      success: true,
      message: `Updated mapping for '${colName}' to parse numbers (e.g. extracts 7 from 'GRADE 07 A'). Retrying import now...`
    });
    setTimeout(() => {
      handleExecuteImport();
    }, 400);
  };

  // Auto-resolve non-critical validation errors by unsetting strict requirement or setting flexible transformations
  const handleAutoFixValidationErrors = () => {
    if (!onSaveMappings) return;
    const updatedMappings = mappings.map(m => ({
      ...m,
      columns: m.columns.map(c => {
        // If column had type errors and is not uniqueKey, make required=false and use flexible parsing
        if (!c.uniqueKey) {
          return {
            ...c,
            required: false,
            transformation: c.dataType === 'date' ? 'parse_date' : c.dataType === 'integer' ? 'parse_number' : c.transformation
          };
        }
        return c;
      })
    }));

    onSaveMappings(updatedMappings);
    setImportNotice({
      success: true,
      message: 'Applied relaxed type tolerance: optional fields will coerce nulls instead of throwing validation errors!'
    });
    setTimeout(() => {
      handleRunDryRun();
    }, 400);
  };

  // Available worksheets in current run
  const availableSheets = useMemo(() => {
    if (dryRunResult?.sheetSummaries && dryRunResult.sheetSummaries.length > 0) {
      return dryRunResult.sheetSummaries.map(s => s.sheetName);
    }
    if (currentAnalysis?.worksheets && currentAnalysis.worksheets.length > 0) {
      return currentAnalysis.worksheets.map(w => w.sheetName);
    }
    return [];
  }, [dryRunResult, currentAnalysis]);

  // Filtered records across action, worksheet, and search query
  const filteredRecords = useMemo(() => {
    if (!dryRunResult) return [];
    return dryRunResult.sampleTransformedRecords.filter(item => {
      // Filter by Action
      if (filterAction !== 'ALL' && item.action !== filterAction) return false;
      // Filter by Sheet
      if (selectedSheetFilter !== 'ALL' && item.sheetName !== selectedSheetFilter) return false;
      // Filter by Search
      if (searchRecordQuery.trim() !== '') {
        const query = searchRecordQuery.toLowerCase();
        const payloadStr = JSON.stringify(item.record).toLowerCase();
        const sheetStr = item.sheetName.toLowerCase();
        const tableStr = item.targetTable.toLowerCase();
        const rowStr = String(item.rowNumber);
        if (!payloadStr.includes(query) && !sheetStr.includes(query) && !tableStr.includes(query) && !rowStr.includes(query)) {
          return false;
        }
      }
      return true;
    });
  }, [dryRunResult, filterAction, selectedSheetFilter, searchRecordQuery]);

  const totalPages = Math.max(1, Math.ceil(filteredRecords.length / (recordsPerPage === -1 ? filteredRecords.length || 1 : recordsPerPage)));
  const currentPage = Math.min(recordsPage, totalPages);

  const paginatedRecords = useMemo(() => {
    if (recordsPerPage === -1) return filteredRecords;
    const start = (currentPage - 1) * recordsPerPage;
    return filteredRecords.slice(start, start + recordsPerPage);
  }, [filteredRecords, currentPage, recordsPerPage]);

  const filteredErrors = dryRunResult?.errors.filter(err => {
    if (selectedErrorType === 'ALL') return true;
    return err.errorType === selectedErrorType;
  }) || [];

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-emerald-50 text-emerald-600">
            <PlaySquare className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Interactive Dry Run Simulator</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Simulate transformations, unique key conflict resolution, and schema validation without modifying live data.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-3">
          <button
            id="btn-trigger-dryrun"
            onClick={handleRunDryRun}
            disabled={isRunningDryRun || isLoadingFile}
            className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-xs font-semibold text-slate-700 shadow-2xs transition-colors"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isRunningDryRun || isLoadingFile ? 'animate-spin' : ''}`} />
            <span>{isRunningDryRun ? 'Simulating...' : 'Re-Run Simulation'}</span>
          </button>

          <button
            id="btn-import-now"
            onClick={handleExecuteImport}
            disabled={isImporting || isLoadingFile}
            className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-semibold text-white shadow-2xs transition-colors cursor-pointer"
          >
            <Play className={`w-3.5 h-3.5 ${isImporting ? 'animate-spin' : ''}`} />
            <span>{isImporting ? 'Writing to Supabase...' : 'Import Now (Write to DB)'}</span>
          </button>
        </div>
      </div>

      {/* Source File & Monitored Workbooks Selection Bar */}
      <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex flex-wrap items-center gap-3">
          <div className="flex items-center space-x-2">
            <FileSpreadsheet className="w-5 h-5 text-emerald-600 shrink-0" />
            <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">Active File:</span>
          </div>

          {/* Select from Monitored / Nextcloud files */}
          <div className="relative">
            <select
              value={currentAnalysis?.filename || ''}
              onChange={(e) => handleSelectNextcloudFile(e.target.value)}
              disabled={isLoadingFile || isRunningDryRun}
              className="text-xs font-semibold text-slate-800 bg-slate-50 hover:bg-slate-100 border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-emerald-500 transition-colors"
            >
              {currentAnalysis && !files.some(f => f.filename === currentAnalysis.filename) && (
                <option value={currentAnalysis.filename}>{currentAnalysis.filename} (Current File)</option>
              )}
              {files.map(f => (
                <option key={f.id || f.filename} value={f.filename}>
                  {f.filename} ({f.fileSizeFormatted || `${f.fileSize} B`})
                </option>
              ))}
              {files.length === 0 && !currentAnalysis && (
                <option value="">No files detected</option>
              )}
            </select>
          </div>

          <div className="flex items-center space-x-2 text-xs text-slate-500 font-medium">
            <span className="inline-flex items-center px-2 py-0.5 rounded-full bg-slate-100 text-slate-700 border border-slate-200">
              {currentAnalysis?.totalWorksheets || dryRunResult?.sheetSummaries.length || 1} Worksheets
            </span>
            <span className="inline-flex items-center px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 font-bold">
              {dryRunResult?.totalRows || 0} Total Rows Extracted
            </span>
          </div>
        </div>

        <div className="flex items-center space-x-2">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleLocalFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
          />
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={isLoadingFile}
            className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-xs font-semibold text-slate-700 shadow-2xs transition-colors cursor-pointer"
            title="Upload any local Excel or CSV file to simulate"
          >
            <Upload className="w-3.5 h-3.5 text-slate-500" />
            <span>Upload Local File</span>
          </button>

          <button
            type="button"
            onClick={() => onNavigate('files')}
            className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold text-emerald-700 hover:text-emerald-800 hover:bg-emerald-50 transition-colors"
          >
            <span>All Files ({files.length})</span>
            <ArrowRight className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Simulation Notice Banner */}
      <div className="p-4 rounded-xl bg-sky-50 border border-sky-200 text-xs text-sky-950 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-2xs">
        <div className="flex items-start space-x-2.5">
          <Info className="w-4 h-4 text-sky-600 shrink-0 mt-0.5" />
          <div>
            <span className="font-bold text-sky-900">Dry Run is a Simulation (Zero Database Writes):</span>
            <p className="text-sky-800 mt-0.5">
              The Dry Run validates your mappings, transformations, and preview rows without modifying Supabase. To write the verified rows to your live database, click the green <strong>&quot;Import Now (Write to DB)&quot;</strong> button.
            </p>
          </div>
        </div>
        <button
          onClick={handleExecuteImport}
          disabled={isImporting || (dryRunResult ? dryRunResult.validRows === 0 : false)}
          className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-bold text-white shadow-2xs transition-colors shrink-0 disabled:opacity-50 cursor-pointer"
        >
          <Play className={`w-3.5 h-3.5 ${isImporting ? 'animate-spin' : ''}`} />
          <span>{isImporting ? 'Writing...' : 'Push to Supabase Now'}</span>
        </button>
      </div>

      {/* Duplicate / Key Conflict Auto-Fix Banner */}
      {dryRunResult && dryRunResult.errors.some(e => e.errorType === 'duplicate') && (
        <div className="p-4 rounded-xl bg-amber-50 border border-amber-300 text-amber-950 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 text-xs shadow-2xs">
          <div className="flex items-start space-x-2.5">
            <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
            <div>
              <span className="font-bold text-amber-900">Merge Key Conflict in Data:</span>
              <p className="text-amber-800 mt-0.5">
                The column marked as Merge Key contains duplicate values across rows. To import all rows directly without unique key blocks, switch to Direct INSERT.
              </p>
            </div>
          </div>
          <button
            onClick={handleSwitchToInsertOnly}
            className="px-3.5 py-1.5 rounded-lg bg-amber-700 hover:bg-amber-800 text-white font-bold text-xs shrink-0 shadow-2xs cursor-pointer"
          >
            Switch to Direct INSERT (Clear Merge Keys)
          </button>
        </div>
      )}

      {importNotice && (
        <div className={`p-4 rounded-xl text-xs font-medium flex items-center space-x-2 border shadow-2xs ${
          importNotice.success
            ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
            : 'bg-rose-50 text-rose-900 border-rose-200'
        }`}>
          {importNotice.success ? (
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          ) : (
            <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
          )}
          <span className="flex-1">{importNotice.message}</span>
        </div>
      )}

      {typeMismatchAlert && (
        <div className="p-5 rounded-xl bg-amber-50 border-2 border-amber-300 text-slate-800 shadow-xs space-y-4">
          <div className="flex items-start justify-between gap-3">
            <div className="flex items-start space-x-3">
              <div className="p-2 rounded-lg bg-amber-100 text-amber-800 shrink-0 mt-0.5">
                <Database className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-slate-900">
                  Supabase PostgreSQL Data Type Mismatch (Error 22P02)
                </h3>
                <p className="text-xs text-slate-700 mt-1 leading-relaxed">
                  Supabase column <code className="px-1.5 py-0.5 bg-amber-100 rounded font-mono font-bold text-amber-900">{typeMismatchAlert.columnName}</code> in table <code className="px-1.5 py-0.5 bg-amber-100 rounded font-mono font-bold text-amber-900">{typeMismatchAlert.tableName}</code> is defined as <strong className="font-semibold text-rose-700">INTEGER</strong> in PostgreSQL, but your source data has text: <code className="px-1.5 py-0.5 bg-white border border-amber-200 rounded font-mono font-bold text-slate-900">&quot;{typeMismatchAlert.offendingValue}&quot;</code>. PostgreSQL cannot store words inside integer columns.
                </p>
              </div>
            </div>
            <button
              onClick={() => setTypeMismatchAlert(null)}
              className="text-xs text-slate-400 hover:text-slate-600 font-semibold"
            >
              Dismiss
            </button>
          </div>

          <div className="bg-white p-3.5 rounded-lg border border-amber-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
            <div className="space-y-1">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Recommended Resolution (Store Full Text):</span>
              <div className="font-mono text-xs bg-slate-900 text-emerald-400 px-3 py-1.5 rounded flex items-center justify-between">
                <span>{typeMismatchAlert.alterSql}</span>
                <button
                  type="button"
                  onClick={() => {
                    navigator.clipboard.writeText(typeMismatchAlert.alterSql);
                    setCopiedAlterSql(true);
                    setTimeout(() => setCopiedAlterSql(false), 2000);
                  }}
                  className="ml-3 text-slate-400 hover:text-white"
                  title="Copy SQL to Clipboard"
                >
                  {copiedAlterSql ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                </button>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2 pt-2 md:pt-0">
              <button
                type="button"
                onClick={handleFixTypeMismatchByAlterSql}
                disabled={typeMismatchAlert.isFixing}
                className="px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-xs shadow-2xs transition-colors flex items-center space-x-1.5"
              >
                <Sparkles className="w-3.5 h-3.5" />
                <span>{typeMismatchAlert.isFixing ? 'Altering Column...' : '1-Click Convert to TEXT & Retry'}</span>
              </button>

              <button
                type="button"
                onClick={handleFixTypeMismatchByExtractNumber}
                className="px-3 py-2 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 font-medium text-xs shadow-2xs transition-colors"
                title="Converts 'GRADE 07 A' to number 7"
              >
                Extract Number (e.g. 7)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Dry Run Results Cards */}
      {dryRunResult ? (
        <div className="space-y-6">
          {/* Summary Metrics */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-slate-500 uppercase">Total Rows</span>
              <div className="text-2xl font-bold text-slate-900 mt-1">{dryRunResult.totalRows.toLocaleString()}</div>
              <span className="text-[11px] text-slate-500">Across {dryRunResult.sheetSummaries.length} worksheets</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-emerald-600 uppercase">Valid Rows</span>
              <div className="text-2xl font-bold text-emerald-600 mt-1">{dryRunResult.validRows.toLocaleString()}</div>
              <span className="text-[11px] text-emerald-600">Passed all validation</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-blue-600 uppercase">Proposed Inserts</span>
              <div className="text-2xl font-bold text-blue-600 mt-1">+{dryRunResult.proposedInserts.toLocaleString()}</div>
              <span className="text-[11px] text-slate-500">Brand new records</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-purple-600 uppercase">Proposed Updates (Merge)</span>
              <div className="text-2xl font-bold text-purple-600 mt-1">{dryRunResult.proposedUpdates.toLocaleString()}</div>
              <span className="text-[11px] text-slate-500">Conflict merge updates</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-rose-600 uppercase">Validation Errors</span>
              <div className="text-2xl font-bold text-rose-600 mt-1">{dryRunResult.failedRows.toLocaleString()}</div>
              <span className="text-[11px] text-slate-500">
                {dryRunResult.failedRows === 0 ? 'Zero errors!' : 'Integrity protection'}
              </span>
            </div>
          </div>

          {/* Action callout after dry-run metrics */}
          <div className="p-4 rounded-xl bg-gradient-to-r from-emerald-50 via-teal-50 to-blue-50 border border-emerald-200 flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center space-x-3">
              <div className="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0">
                <CheckCircle2 className="w-5 h-5" />
              </div>
              <div>
                <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                  Simulation Ready: {dryRunResult.validRows} Record{dryRunResult.validRows === 1 ? '' : 's'} Verified
                </h4>
                <p className="text-xs text-slate-600 mt-0.5">
                  Dry run is a simulation. Click below to commit and sync all verified records directly into your Supabase database.
                </p>
              </div>
            </div>
            <button
              onClick={handleExecuteImport}
              disabled={isImporting || dryRunResult.validRows === 0}
              className="inline-flex items-center space-x-2 px-5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-bold text-white shadow-2xs transition-all hover:scale-[1.02] shrink-0 cursor-pointer disabled:opacity-50"
            >
              <Play className={`w-3.5 h-3.5 ${isImporting ? 'animate-spin' : ''}`} />
              <span>{isImporting ? 'Writing to Supabase...' : 'Push & Apply to Supabase Now'}</span>
            </button>
          </div>

          {/* Worksheet Level Breakdown */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h3 className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
                Worksheet Simulation Breakdown
              </h3>
              <span className="text-xs text-slate-500">
                Routing {dryRunResult.sheetSummaries.length} Excel sheet(s) to PostgreSQL tables
              </span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-100 text-slate-700 uppercase font-semibold border-b border-slate-200">
                  <tr>
                    <th className="px-4 py-3">Worksheet Name</th>
                    <th className="px-4 py-3">Target Table</th>
                    <th className="px-4 py-3">Total Rows</th>
                    <th className="px-4 py-3 text-emerald-700">Inserts</th>
                    <th className="px-4 py-3 text-purple-700">Updates</th>
                    <th className="px-4 py-3 text-rose-700">Fails</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-mono">
                  {dryRunResult.sheetSummaries.map((ss) => (
                    <tr key={ss.sheetName} className="hover:bg-slate-50">
                      <td className="px-4 py-2.5 font-sans font-medium text-slate-900">{ss.sheetName}</td>
                      <td className="px-4 py-2.5 text-teal-700">public.{ss.targetTable}</td>
                      <td className="px-4 py-2.5">{ss.totalRows}</td>
                      <td className="px-4 py-2.5 text-emerald-600 font-bold">+{ss.inserts}</td>
                      <td className="px-4 py-2.5 text-purple-600 font-bold">{ss.updates}</td>
                      <td className="px-4 py-2.5 text-rose-600 font-bold">{ss.fails}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Validation Errors Detail Table (if any) */}
          {dryRunResult.errors.length > 0 && (
            <div className="bg-white rounded-xl border border-rose-200 shadow-2xs overflow-hidden">
              <div className="p-4 bg-rose-50 border-b border-rose-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div className="flex items-center space-x-2">
                  <ShieldAlert className="w-5 h-5 text-rose-600" />
                  <div>
                    <span className="text-xs font-bold text-rose-950 uppercase tracking-wider">
                      Validation Errors Detected ({dryRunResult.errors.length})
                    </span>
                    <p className="text-[11px] text-rose-800">
                      Rows with validation issues are skipped from database write to protect data integrity.
                    </p>
                  </div>
                </div>

                <div className="flex items-center space-x-2">
                  {/* Error Filter */}
                  <select
                    value={selectedErrorType}
                    onChange={(e) => setSelectedErrorType(e.target.value)}
                    className="px-2.5 py-1 text-xs rounded border border-rose-300 bg-white text-rose-900 font-medium"
                  >
                    <option value="ALL">All Errors ({dryRunResult.errors.length})</option>
                    <option value="missing_required">Missing Required</option>
                    <option value="type_mismatch">Type Mismatch</option>
                    <option value="invalid_date">Invalid Date</option>
                    <option value="duplicate">Duplicate Unique Key</option>
                  </select>

                  {onSaveMappings && (
                    <button
                      onClick={handleAutoFixValidationErrors}
                      className="px-3 py-1 rounded bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold shadow-2xs transition-colors flex items-center space-x-1"
                      title="Make optional fields nullable and apply flexible parser"
                    >
                      <Wand2 className="w-3.5 h-3.5" />
                      <span>Auto-Relax Optional Constraints</span>
                    </button>
                  )}
                </div>
              </div>

              <div className="overflow-x-auto max-h-72">
                <table className="w-full text-left text-xs font-mono">
                  <thead className="bg-rose-100/50 text-rose-900 uppercase font-semibold border-b border-rose-200 sticky top-0">
                    <tr>
                      <th className="px-4 py-2.5">Worksheet</th>
                      <th className="px-3 py-2.5">Row</th>
                      <th className="px-3 py-2.5">Col</th>
                      <th className="px-4 py-2.5">Field</th>
                      <th className="px-4 py-2.5">Raw Value</th>
                      <th className="px-4 py-2.5">Error Detail</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-rose-100">
                    {filteredErrors.slice(0, 100).map((err, idx) => (
                      <tr key={idx} className="hover:bg-rose-50/50">
                        <td className="px-4 py-2 font-sans font-medium text-slate-900">{err.worksheetName}</td>
                        <td className="px-3 py-2 font-bold text-rose-700">{err.rowNumber}</td>
                        <td className="px-3 py-2">{err.excelColumn}</td>
                        <td className="px-4 py-2 text-slate-800">{err.columnName}</td>
                        <td className="px-4 py-2 text-slate-500 italic truncate max-w-xs">&quot;{err.rawValue}&quot;</td>
                        <td className="px-4 py-2 text-rose-700 font-sans">{err.errorMessage}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Transformed Records Preview */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            <div className="p-4 bg-slate-50 border-b border-slate-200 space-y-3">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="flex items-center space-x-2">
                  <span className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
                    Transformed Records Preview
                  </span>
                  <span className="text-xs font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                    {filteredRecords.length} records {searchRecordQuery ? '(filtered)' : ''}
                  </span>
                </div>

                {/* Search in Record Payloads */}
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={searchRecordQuery}
                    onChange={(e) => {
                      setSearchRecordQuery(e.target.value);
                      setRecordsPage(1);
                    }}
                    placeholder="Search rows or values..."
                    className="pl-8 pr-3 py-1 text-xs rounded-lg border border-slate-300 bg-white text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500 w-full sm:w-64"
                  />
                </div>
              </div>

              {/* Filters Bar: Action Type & Worksheet */}
              <div className="flex flex-wrap items-center justify-between gap-3 pt-1 border-t border-slate-200/60">
                <div className="flex flex-wrap items-center gap-1.5 text-xs">
                  <span className="text-[11px] font-bold text-slate-500 uppercase mr-1">Action:</span>
                  {(['ALL', 'INSERT', 'UPDATE', 'FAIL'] as const).map(action => (
                    <button
                      key={action}
                      onClick={() => {
                        setFilterAction(action);
                        setRecordsPage(1);
                      }}
                      className={`px-2.5 py-0.5 rounded text-xs font-semibold transition-colors cursor-pointer ${
                        filterAction === action
                          ? action === 'INSERT' ? 'bg-emerald-600 text-white' :
                            action === 'UPDATE' ? 'bg-purple-600 text-white' :
                            action === 'FAIL' ? 'bg-rose-600 text-white' :
                            'bg-slate-900 text-white'
                          : 'bg-slate-200/70 text-slate-700 hover:bg-slate-300'
                      }`}
                    >
                      {action}
                    </button>
                  ))}
                </div>

                {availableSheets.length > 1 && (
                  <div className="flex flex-wrap items-center gap-1 text-xs">
                    <span className="text-[11px] font-bold text-slate-500 uppercase mr-1">Worksheet:</span>
                    <button
                      onClick={() => {
                        setSelectedSheetFilter('ALL');
                        setRecordsPage(1);
                      }}
                      className={`px-2 py-0.5 rounded text-xs font-medium transition-colors cursor-pointer ${
                        selectedSheetFilter === 'ALL'
                          ? 'bg-slate-800 text-white font-bold'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      All Sheets
                    </button>
                    {availableSheets.map(s => (
                      <button
                        key={s}
                        onClick={() => {
                          setSelectedSheetFilter(s);
                          setRecordsPage(1);
                        }}
                        className={`px-2 py-0.5 rounded text-xs font-medium transition-colors cursor-pointer ${
                          selectedSheetFilter === s
                            ? 'bg-emerald-700 text-white font-bold'
                            : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                        }`}
                      >
                        {s}
                      </button>
                    ))}
                  </div>
                )}

                {/* Rows Per Page */}
                <div className="flex items-center space-x-1.5 text-xs text-slate-500 ml-auto">
                  <span>Show:</span>
                  <select
                    value={recordsPerPage}
                    onChange={(e) => {
                      setRecordsPerPage(Number(e.target.value));
                      setRecordsPage(1);
                    }}
                    className="px-2 py-0.5 rounded border border-slate-300 bg-white text-slate-700 font-semibold text-xs"
                  >
                    <option value={25}>25 rows</option>
                    <option value={50}>50 rows</option>
                    <option value={100}>100 rows</option>
                    <option value={250}>250 rows</option>
                    <option value={-1}>All ({filteredRecords.length})</option>
                  </select>
                </div>
              </div>
            </div>

            <div className="overflow-x-auto max-h-96">
              <table className="w-full text-left text-xs font-mono">
                <thead className="bg-slate-100 text-slate-700 uppercase font-semibold border-b border-slate-200 sticky top-0">
                  <tr>
                    <th className="px-3 py-2.5 w-16">Row</th>
                    <th className="px-3 py-2.5">Worksheet</th>
                    <th className="px-3 py-2.5">Target Table</th>
                    <th className="px-3 py-2.5 text-center">Action</th>
                    <th className="px-4 py-2.5">Transformed Record Payload</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {paginatedRecords.length > 0 ? (
                    paginatedRecords.map((item, idx) => (
                      <tr key={idx} className="hover:bg-slate-50">
                        <td className="px-3 py-2 text-slate-400 font-bold">{item.rowNumber}</td>
                        <td className="px-3 py-2 text-slate-900 font-sans font-medium">{item.sheetName}</td>
                        <td className="px-3 py-2 text-teal-700 font-sans">public.{item.targetTable}</td>
                        <td className="px-3 py-2 text-center">
                          <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-bold ${
                            item.action === 'INSERT' ? 'bg-emerald-100 text-emerald-800' :
                            item.action === 'UPDATE' ? 'bg-purple-100 text-purple-800' :
                            'bg-rose-100 text-rose-800'
                          }`}>
                            {item.action}
                          </span>
                        </td>
                        <td className="px-4 py-2 text-slate-700 max-w-xl truncate font-mono text-[11px]">
                          {JSON.stringify(item.record)}
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={5} className="px-4 py-8 text-center text-slate-400 font-sans text-xs">
                        No transformed records matched the selected filters.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {totalPages > 1 && recordsPerPage !== -1 && (
              <div className="p-3 bg-slate-50 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
                <div className="text-slate-500 font-medium">
                  Showing <span className="font-bold text-slate-800">{(currentPage - 1) * recordsPerPage + 1}</span>–<span className="font-bold text-slate-800">{Math.min(currentPage * recordsPerPage, filteredRecords.length)}</span> of <span className="font-bold text-slate-800">{filteredRecords.length}</span> records
                </div>

                <div className="flex items-center space-x-1.5">
                  <button
                    onClick={() => setRecordsPage(p => Math.max(1, p - 1))}
                    disabled={currentPage === 1}
                    className="p-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                    title="Previous Page"
                  >
                    <ChevronLeft className="w-4 h-4" />
                  </button>

                  <span className="px-2.5 py-1 text-slate-700 font-semibold bg-white border border-slate-300 rounded-lg">
                    Page {currentPage} of {totalPages}
                  </span>

                  <button
                    onClick={() => setRecordsPage(p => Math.min(totalPages, p + 1))}
                    disabled={currentPage === totalPages}
                    className="p-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                    title="Next Page"
                  >
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      ) : (
        /* Empty State */
        <div className="bg-white rounded-xl p-12 border border-slate-200 shadow-2xs text-center space-y-4">
          <div className="mx-auto w-12 h-12 rounded-xl bg-slate-100 flex items-center justify-center text-slate-500">
            <PlaySquare className="w-6 h-6" />
          </div>
          <div>
            <h3 className="text-base font-semibold text-slate-900">Dry Run Not Executed Yet</h3>
            <p className="text-xs text-slate-500 max-w-md mx-auto mt-1">
              Click <span className="font-semibold text-slate-700">Execute Dry Run</span> to parse your workbook with active mappings, compute section heading propagation, and identify all proposed Inserts and Updates.
            </p>
          </div>
          <button
            onClick={handleRunDryRun}
            className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-semibold text-white shadow-2xs"
          >
            <Play className="w-4 h-4" />
            <span>Launch Simulation Now</span>
          </button>
        </div>
      )}
    </div>
  );
};
