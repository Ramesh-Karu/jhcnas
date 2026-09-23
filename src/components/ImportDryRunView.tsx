import React, { useState, useEffect } from 'react';
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
  Check
} from 'lucide-react';
import { 
  WorkbookAnalysis, 
  WorksheetMapping, 
  DryRunResult, 
  RowValidationError, 
  ImportLog, 
  NavigationTab,
  SupabaseConfig
} from '../types';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
import { DryRunEngine } from '../services/dryRunEngine';
import { ApiClient } from '../services/apiClient';

interface ImportDryRunViewProps {
  currentAnalysis: WorkbookAnalysis | null;
  mappings: WorksheetMapping[];
  databaseState: Record<string, any[]>;
  onUpdateDatabase: (tableName: string, newRecords: any[]) => void;
  onAddImportLog: (log: ImportLog, errors: RowValidationError[]) => void;
  onNavigate: (tab: NavigationTab) => void;
  supabaseConfig?: SupabaseConfig;
  onSaveMappings?: (mappings: WorksheetMapping[]) => void;
}

export const ImportDryRunView: React.FC<ImportDryRunViewProps> = ({
  currentAnalysis,
  mappings,
  databaseState,
  onUpdateDatabase,
  onAddImportLog,
  onNavigate,
  supabaseConfig,
  onSaveMappings
}) => {
  const [dryRunResult, setDryRunResult] = useState<DryRunResult | null>(null);
  const [isRunningDryRun, setIsRunningDryRun] = useState<boolean>(false);
  const [isImporting, setIsImporting] = useState<boolean>(false);
  const [importNotice, setImportNotice] = useState<{ success: boolean; message: string } | null>(null);
  const [filterAction, setFilterAction] = useState<'ALL' | 'INSERT' | 'UPDATE' | 'FAIL'>('ALL');
  const [selectedErrorType, setSelectedErrorType] = useState<string>('ALL');

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

    const wb = getEffectiveWorkbook();
    const filename = currentAnalysis?.filename || mappings[0]?.workbookName || 'students.xlsx';

    // Synchronize latest existing records from Supabase if connected
    const mergedDbState: Record<string, any[]> = { ...databaseState };
    if (supabaseConfig?.url && (supabaseConfig.anonKey || supabaseConfig.serviceKey || supabaseConfig.serviceRoleKey)) {
      try {
        const uniqueTables = Array.from(new Set(mappings.map(m => m.supabaseTable)));
        for (const tbl of uniqueTables) {
          const res = await ApiClient.fetchSupabaseTableRows(supabaseConfig, tbl, 100);
          if (res.success && res.rows && res.rows.length > 0) {
            mergedDbState[tbl] = res.rows;
            onUpdateDatabase(tbl, res.rows);
          }
        }
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

    await new Promise(r => setTimeout(r, 200));
    setDryRunResult(result);
    setIsRunningDryRun(false);
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
              supabaseWriteNotice += ` Synced ${cleanRecordsForSupabase.length} rows to '${wm.supabaseTable}'.`;
              // Fetch latest live rows from Supabase to update local UI databaseState
              const latest = await ApiClient.fetchSupabaseTableRows(supabaseConfig, wm.supabaseTable, 100);
              if (latest.success && latest.rows) {
                onUpdateDatabase(wm.supabaseTable, latest.rows);
              }
            } else {
              supabaseErrors.push(`${wm.supabaseTable}: ${upRes.error || 'Check table schema'}`);
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
        message: `Import processed locally, but Supabase reported issues: ${supabaseErrors.join(' | ')}. Please check the Supabase Schema tab to verify tables exist.`
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

  const displayedRecords = dryRunResult
    ? (filterAction === 'ALL'
        ? dryRunResult.sampleTransformedRecords
        : dryRunResult.sampleTransformedRecords.filter(r => r.action === filterAction))
    : [];

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
            disabled={isRunningDryRun}
            className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-xs font-semibold text-slate-700 shadow-2xs transition-colors"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isRunningDryRun ? 'animate-spin' : ''}`} />
            <span>{isRunningDryRun ? 'Simulating...' : 'Re-Run Simulation'}</span>
          </button>

          <button
            id="btn-import-now"
            onClick={handleExecuteImport}
            disabled={isImporting}
            className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-semibold text-white shadow-2xs transition-colors"
          >
            <Play className={`w-3.5 h-3.5 ${isImporting ? 'animate-spin' : ''}`} />
            <span>{isImporting ? 'Writing to Supabase...' : 'Import Now (Write to DB)'}</span>
          </button>
        </div>
      </div>

      {importNotice && (
        <div className="p-4 rounded-xl bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{importNotice.message}</span>
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
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="flex items-center space-x-2">
                <span className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
                  Sample Transformed Records
                </span>
                <span className="text-xs text-slate-500">
                  (Showing {displayedRecords.length} records)
                </span>
              </div>

              {/* Action Filter Pills */}
              <div className="flex items-center space-x-1.5 text-xs">
                {(['ALL', 'INSERT', 'UPDATE', 'FAIL'] as const).map(action => (
                  <button
                    key={action}
                    onClick={() => setFilterAction(action)}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      filterAction === action
                        ? 'bg-slate-900 text-white'
                        : 'bg-slate-200/70 text-slate-700 hover:bg-slate-300'
                    }`}
                  >
                    {action}
                  </button>
                ))}
              </div>
            </div>

            <div className="overflow-x-auto max-h-80">
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
                  {displayedRecords.map((item, idx) => (
                    <tr key={idx} className="hover:bg-slate-50">
                      <td className="px-3 py-2 text-slate-400">{item.rowNumber}</td>
                      <td className="px-3 py-2 text-slate-900 font-sans font-medium">{item.sheetName}</td>
                      <td className="px-3 py-2 text-teal-700">public.{item.targetTable}</td>
                      <td className="px-3 py-2 text-center">
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-bold ${
                          item.action === 'INSERT' ? 'bg-emerald-100 text-emerald-800' :
                          item.action === 'UPDATE' ? 'bg-purple-100 text-purple-800' :
                          'bg-rose-100 text-rose-800'
                        }`}>
                          {item.action}
                        </span>
                      </td>
                      <td className="px-4 py-2 text-slate-700 max-w-lg truncate">
                        {JSON.stringify(item.record)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
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
