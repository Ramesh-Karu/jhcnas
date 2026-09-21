import React, { useState } from 'react';
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
  AlertCircle
} from 'lucide-react';
import { 
  WorkbookAnalysis, 
  WorksheetMapping, 
  DryRunResult, 
  RowValidationError, 
  ImportLog, 
  NavigationTab 
} from '../types';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
import { DryRunEngine } from '../services/dryRunEngine';

interface ImportDryRunViewProps {
  currentAnalysis: WorkbookAnalysis | null;
  mappings: WorksheetMapping[];
  databaseState: Record<string, any[]>;
  onUpdateDatabase: (tableName: string, newRecords: any[]) => void;
  onAddImportLog: (log: ImportLog, errors: RowValidationError[]) => void;
  onNavigate: (tab: NavigationTab) => void;
}

export const ImportDryRunView: React.FC<ImportDryRunViewProps> = ({
  currentAnalysis,
  mappings,
  databaseState,
  onUpdateDatabase,
  onAddImportLog,
  onNavigate
}) => {
  const [dryRunResult, setDryRunResult] = useState<DryRunResult | null>(null);
  const [isRunningDryRun, setIsRunningDryRun] = useState<boolean>(false);
  const [isImporting, setIsImporting] = useState<boolean>(false);
  const [importNotice, setImportNotice] = useState<{ success: boolean; message: string } | null>(null);
  const [filterAction, setFilterAction] = useState<'ALL' | 'INSERT' | 'UPDATE' | 'FAIL'>('ALL');

  const handleRunDryRun = async () => {
    setIsRunningDryRun(true);
    setImportNotice(null);

    // Generate or fetch workbook
    const sample = createComplexSampleWorkbook();
    const result = DryRunEngine.executeDryRun(
      sample.workbook,
      currentAnalysis?.filename || 'students_complex.xlsx',
      mappings,
      databaseState
    );

    await new Promise(r => setTimeout(r, 600));
    setDryRunResult(result);
    setIsRunningDryRun(false);
  };

  const handleExecuteImport = async () => {
    setIsImporting(true);
    setImportNotice(null);

    const sample = createComplexSampleWorkbook();
    const filename = currentAnalysis?.filename || 'students_complex.xlsx';
    const fileHash = currentAnalysis?.fileHash || 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';

    const { log, errors } = DryRunEngine.executeRealImport(
      sample.workbook,
      filename,
      fileHash,
      mappings,
      databaseState,
      onUpdateDatabase
    );

    await new Promise(r => setTimeout(r, 800));

    onAddImportLog(log, errors);
    setIsImporting(false);
    setImportNotice({
      success: true,
      message: `Production Import completed! ${log.rowsInserted} records inserted, ${log.rowsUpdated} records updated, ${log.rowsFailed} failed.`
    });

    // Re-run dry run to show updated status
    const updatedResult = DryRunEngine.executeDryRun(sample.workbook, filename, mappings, databaseState);
    setDryRunResult(updatedResult);
  };

  const displayedRecords = dryRunResult
    ? (filterAction === 'ALL'
        ? dryRunResult.sampleTransformedRecords
        : dryRunResult.sampleTransformedRecords.filter(r => r.action === filterAction))
    : [];

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
              Simulate transformations, unique key conflict resolution, and schema validation without database modifications.
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
            <span>{isRunningDryRun ? 'Simulating...' : 'Execute Dry Run'}</span>
          </button>

          <button
            id="btn-import-now"
            onClick={handleExecuteImport}
            disabled={isImporting}
            className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-semibold text-white shadow-2xs transition-colors"
          >
            <Play className={`w-3.5 h-3.5 ${isImporting ? 'animate-spin' : ''}`} />
            <span>{isImporting ? 'Importing...' : 'Import Now (Write to DB)'}</span>
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
              <div className="text-2xl font-bold text-slate-900 mt-1">{dryRunResult.totalRows}</div>
              <span className="text-[11px] text-slate-500">{dryRunResult.totalWorksheets} active worksheets</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-emerald-600 uppercase">Proposed Inserts</span>
              <div className="text-2xl font-bold text-emerald-700 mt-1">+{dryRunResult.proposedInserts}</div>
              <span className="text-[11px] text-slate-500">New unique keys</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-purple-600 uppercase">Proposed Updates</span>
              <div className="text-2xl font-bold text-purple-700 mt-1">{dryRunResult.proposedUpdates}</div>
              <span className="text-[11px] text-slate-500">Idempotent key matches</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-rose-600 uppercase">Validation Failures</span>
              <div className="text-2xl font-bold text-rose-700 mt-1">{dryRunResult.failedRows}</div>
              <span className="text-[11px] text-slate-500">Rejected before write</span>
            </div>

            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <span className="text-xs font-semibold text-blue-600 uppercase">Validation Rate</span>
              <div className="text-2xl font-bold text-blue-700 mt-1">
                {dryRunResult.totalRows > 0
                  ? `${Math.round((dryRunResult.validRows / dryRunResult.totalRows) * 100)}%`
                  : '0%'}
              </div>
              <span className="text-[11px] text-slate-500">Schema compliance</span>
            </div>
          </div>

          {/* Worksheet Level Breakdown */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            <div className="p-4 bg-slate-50 border-b border-slate-200">
              <h3 className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
                Worksheet Simulation Breakdown
              </h3>
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

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs font-mono">
                <thead className="bg-slate-100 text-slate-700 uppercase font-semibold border-b border-slate-200">
                  <tr>
                    <th className="px-3 py-3 w-16">Row</th>
                    <th className="px-3 py-3">Worksheet</th>
                    <th className="px-3 py-3">Target Table</th>
                    <th className="px-3 py-3 text-center">Action</th>
                    <th className="px-4 py-3">Transformed Record Payload</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {displayedRecords.map((item, idx) => (
                    <tr key={idx} className="hover:bg-slate-50">
                      <td className="px-3 py-2.5 text-slate-400">{item.rowNumber}</td>
                      <td className="px-3 py-2.5 text-slate-900 font-sans font-medium">{item.sheetName}</td>
                      <td className="px-3 py-2.5 text-teal-700">{item.targetTable}</td>
                      <td className="px-3 py-2.5 text-center">
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-bold ${
                          item.action === 'INSERT' ? 'bg-emerald-100 text-emerald-800' :
                          item.action === 'UPDATE' ? 'bg-purple-100 text-purple-800' :
                          'bg-rose-100 text-rose-800'
                        }`}>
                          {item.action}
                        </span>
                      </td>
                      <td className="px-4 py-2.5 text-slate-700 max-w-lg truncate">
                        {JSON.stringify(item.record)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Row Validation Errors Table if any */}
          {dryRunResult.errors.length > 0 && (
            <div className="bg-white rounded-xl border border-rose-200 shadow-2xs overflow-hidden">
              <div className="p-4 bg-rose-50 border-b border-rose-200 flex items-center justify-between">
                <span className="text-xs font-semibold text-rose-900 uppercase tracking-wider flex items-center space-x-2">
                  <AlertTriangle className="w-4 h-4 text-rose-600" />
                  <span>Validation Errors Detected ({dryRunResult.errors.length})</span>
                </span>
                <span className="text-xs text-rose-700">Records will be rejected to preserve data integrity</span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-rose-100/50 text-rose-900 uppercase font-semibold border-b border-rose-200">
                    <tr>
                      <th className="px-4 py-3">Worksheet</th>
                      <th className="px-3 py-3">Row</th>
                      <th className="px-3 py-3">Col</th>
                      <th className="px-4 py-3">Field</th>
                      <th className="px-4 py-3">Raw Value</th>
                      <th className="px-4 py-3">Error Message</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-rose-100 font-mono">
                    {dryRunResult.errors.map((err, idx) => (
                      <tr key={idx} className="hover:bg-rose-50/50">
                        <td className="px-4 py-2.5 font-sans font-medium text-slate-900">{err.worksheetName}</td>
                        <td className="px-3 py-2.5 font-bold text-rose-700">{err.rowNumber}</td>
                        <td className="px-3 py-2.5">{err.excelColumn}</td>
                        <td className="px-4 py-2.5 text-slate-700">{err.columnName}</td>
                        <td className="px-4 py-2.5 text-slate-500 italic">"{err.rawValue}"</td>
                        <td className="px-4 py-2.5 text-rose-700 font-sans">{err.errorMessage}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
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
              Click <span className="font-semibold text-slate-700">Execute Dry Run</span> to parse your workbook with active mappings, compute downward section heading propagation, and identify all proposed Inserts and Updates.
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
