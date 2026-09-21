import React, { useState } from 'react';
import { 
  AlertTriangle, 
  Download, 
  Trash2, 
  Filter, 
  Search, 
  CheckCircle2, 
  Info,
  Wrench
} from 'lucide-react';
import { ImportAuditError } from '../types';

interface ErrorsViewProps {
  errors: ImportAuditError[];
  onClearErrors: () => void;
}

export const ErrorsView: React.FC<ErrorsViewProps> = ({
  errors,
  onClearErrors
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedError, setSelectedError] = useState<ImportAuditError | null>(errors[0] || null);

  const filteredErrors = errors.filter(e => 
    e.filename.toLowerCase().includes(searchTerm.toLowerCase()) ||
    e.worksheetName.toLowerCase().includes(searchTerm.toLowerCase()) ||
    e.columnName.toLowerCase().includes(searchTerm.toLowerCase()) ||
    e.errorMessage.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const getSuggestedFix = (err: ImportAuditError) => {
    switch (err.errorType) {
      case 'missing_required':
        return `Supply a non-empty value for '${err.columnName}' in Excel cell ${err.excelColumn}${err.rowNumber}, or configure a default value in Mappings.`;
      case 'invalid_date':
        return `Format the date in Excel cell ${err.excelColumn}${err.rowNumber} to standard ISO (YYYY-MM-DD), or adjust parse_date transformation.`;
      case 'type_mismatch':
        return `Ensure the value in cell ${err.excelColumn}${err.rowNumber} conforms to data type rules (e.g. valid numbers without currency symbols).`;
      case 'duplicate':
        return `Review unique key constraint on '${err.columnName}'. The identifier '${err.rawValue}' already exists in this batch or database.`;
      default:
        return 'Verify the cell formatting in the original Nextcloud Excel workbook and re-run Dry Run.';
    }
  };

  const handleExportCSV = () => {
    const headers = ['ID', 'Filename', 'Worksheet', 'Row', 'Column', 'Field', 'Raw Value', 'Error Message', 'Created At'];
    const rows = filteredErrors.map(e => [
      e.id,
      `"${e.filename}"`,
      `"${e.worksheetName}"`,
      e.rowNumber,
      e.excelColumn,
      e.columnName,
      `"${e.rawValue}"`,
      `"${e.errorMessage.replace(/"/g, '""')}"`,
      e.createdAt
    ]);
    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `import_errors_${new Date().toISOString().split('T')[0]}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-rose-50 text-rose-600">
            <AlertTriangle className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Data Integrity & Error Inspector</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Rejected rows quarantined in <code className="bg-slate-100 px-1 py-0.5 rounded font-mono">import_errors</code>. No bad data is ever written into Supabase.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-2">
          <button
            onClick={handleExportCSV}
            disabled={errors.length === 0}
            className="inline-flex items-center space-x-1 px-3 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 shadow-2xs"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Export Error CSV</span>
          </button>

          <button
            onClick={onClearErrors}
            disabled={errors.length === 0}
            className="inline-flex items-center space-x-1 px-3 py-2 rounded-lg border border-rose-200 text-xs font-medium text-rose-700 bg-rose-50 hover:bg-rose-100 shadow-2xs"
          >
            <Trash2 className="w-3.5 h-3.5" />
            <span>Clear Audit Errors</span>
          </button>
        </div>
      </div>

      {errors.length === 0 ? (
        <div className="bg-white rounded-xl p-12 border border-slate-200 shadow-2xs text-center space-y-3">
          <div className="mx-auto w-12 h-12 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-600">
            <CheckCircle2 className="w-6 h-6" />
          </div>
          <h3 className="text-base font-semibold text-slate-900">Zero Quarantine Errors</h3>
          <p className="text-xs text-slate-500 max-w-sm mx-auto">
            All synchronized Excel rows have passed schema, required field, date, and unique constraint validations.
          </p>
        </div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Errors List */}
          <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
                Quarantined Rows ({filteredErrors.length})
              </span>
              <div className="relative">
                <input
                  type="text"
                  placeholder="Filter errors..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="pl-7 pr-3 py-1 rounded-lg border border-slate-300 text-xs w-44 focus:outline-none"
                />
                <Search className="w-3 h-3 text-slate-400 absolute left-2.5 top-2" />
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-100 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="px-4 py-3">Workbook & Sheet</th>
                    <th className="px-3 py-3 text-center">Row</th>
                    <th className="px-3 py-3 text-center">Col</th>
                    <th className="px-4 py-3">Field</th>
                    <th className="px-4 py-3">Error</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredErrors.map((err) => (
                    <tr 
                      key={err.id} 
                      onClick={() => setSelectedError(err)}
                      className={`cursor-pointer transition-colors ${
                        selectedError?.id === err.id ? 'bg-rose-50/70 font-medium' : 'hover:bg-slate-50'
                      }`}
                    >
                      <td className="px-4 py-3">
                        <div className="font-semibold text-slate-900">{err.filename}</div>
                        <div className="text-[11px] text-slate-500">{err.worksheetName}</div>
                      </td>
                      <td className="px-3 py-3 text-center font-bold text-rose-700 font-mono">
                        {err.rowNumber}
                      </td>
                      <td className="px-3 py-3 text-center font-mono font-bold text-slate-700">
                        {err.excelColumn}
                      </td>
                      <td className="px-4 py-3 font-mono text-slate-800">
                        {err.columnName}
                      </td>
                      <td className="px-4 py-3 text-rose-600 truncate max-w-xs">
                        {err.errorMessage}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Error Detail & Suggested Fix Inspector */}
          <div className="space-y-4">
            {selectedError ? (
              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs space-y-4">
                <div className="flex items-center justify-between pb-3 border-b border-slate-200">
                  <span className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
                    Error Investigation
                  </span>
                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-rose-100 text-rose-800 font-bold uppercase">
                    {selectedError.errorType}
                  </span>
                </div>

                <div className="space-y-3 text-xs">
                  <div>
                    <span className="text-slate-500">Location:</span>
                    <div className="font-medium text-slate-900 mt-0.5">
                      {selectedError.filename} → Sheet: <span className="font-semibold">{selectedError.worksheetName}</span>
                    </div>
                    <div className="text-slate-600 font-mono mt-0.5">
                      Cell: <span className="font-bold text-purple-700">{selectedError.excelColumn}{selectedError.rowNumber}</span> (Row {selectedError.rowNumber})
                    </div>
                  </div>

                  <div>
                    <span className="text-slate-500">Target Field:</span>
                    <div className="font-mono font-bold text-teal-800 mt-0.5">
                      {selectedError.columnName}
                    </div>
                  </div>

                  <div>
                    <span className="text-slate-500">Raw Value Read from Excel:</span>
                    <div className="p-2.5 rounded-lg bg-slate-100 font-mono text-slate-800 mt-1 break-all">
                      {selectedError.rawValue === '' ? (
                        <span className="text-slate-400 italic">(Empty / Null)</span>
                      ) : (
                        selectedError.rawValue
                      )}
                    </div>
                  </div>

                  <div>
                    <span className="text-slate-500">Validation Failure:</span>
                    <div className="p-2.5 rounded-lg bg-rose-50 border border-rose-200 text-rose-800 font-medium mt-1">
                      {selectedError.errorMessage}
                    </div>
                  </div>

                  {/* Suggested Fix */}
                  <div className="pt-2 border-t border-slate-200">
                    <div className="flex items-center space-x-1.5 text-blue-700 font-semibold mb-1">
                      <Wrench className="w-3.5 h-3.5" />
                      <span>Suggested Resolution</span>
                    </div>
                    <p className="text-slate-600 leading-relaxed bg-blue-50/50 p-3 rounded-lg border border-blue-100">
                      {getSuggestedFix(selectedError)}
                    </p>
                  </div>
                </div>
              </div>
            ) : (
              <div className="bg-white rounded-xl p-6 border border-slate-200 text-center text-slate-500 text-xs">
                Select an error from the table to view the cell location, raw value, and remediation instructions.
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
