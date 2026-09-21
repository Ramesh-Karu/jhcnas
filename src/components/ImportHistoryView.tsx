import React, { useState } from 'react';
import { 
  History, 
  CheckCircle2, 
  XCircle, 
  AlertTriangle, 
  Clock, 
  FileSpreadsheet, 
  ArrowDownToLine, 
  Filter,
  Search
} from 'lucide-react';
import { ImportLog } from '../types';

interface ImportHistoryViewProps {
  logs: ImportLog[];
  onSelectLogForErrors?: (logId: string) => void;
}

export const ImportHistoryView: React.FC<ImportHistoryViewProps> = ({
  logs,
  onSelectLogForErrors
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('ALL');

  const filteredLogs = logs.filter(log => {
    const matchesSearch = log.filename.toLowerCase().includes(searchTerm.toLowerCase()) ||
                          log.fileHash.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesStatus = statusFilter === 'ALL' || log.status === statusFilter;
    return matchesSearch && matchesStatus;
  });

  const getStatusBadge = (status: ImportLog['status']) => {
    switch (status) {
      case 'Success':
        return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800">Success</span>;
      case 'Partial Success':
        return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-amber-100 text-amber-800">Partial Success</span>;
      case 'Skipped':
        return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-slate-100 text-slate-700">Skipped (Hash Match)</span>;
      case 'Failed':
        return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-rose-100 text-rose-800">Failed</span>;
      default:
        return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-100 text-blue-800">{status}</span>;
    }
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-blue-50 text-blue-600">
            <History className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Import Audit History</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Complete historical record of every WebDAV scan, file hash comparison, and Supabase upsert cycle.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-3">
          {/* Search bar */}
          <div className="relative">
            <input
              type="text"
              placeholder="Search filename or hash..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-8 pr-3 py-1.5 rounded-lg border border-slate-300 text-xs w-56 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
          </div>

          {/* Status Filter */}
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="px-3 py-1.5 rounded-lg border border-slate-300 text-xs text-slate-700 bg-white focus:outline-none focus:ring-1 focus:ring-blue-500"
          >
            <option value="ALL">All Statuses</option>
            <option value="Success">Success</option>
            <option value="Partial Success">Partial Success</option>
            <option value="Skipped">Skipped</option>
            <option value="Failed">Failed</option>
          </select>
        </div>
      </div>

      {/* History Table */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200">
              <tr>
                <th className="px-5 py-3.5">Filename & Hash</th>
                <th className="px-4 py-3.5">Status</th>
                <th className="px-3 py-3.5 text-center">Worksheets</th>
                <th className="px-4 py-3.5 text-center">Processed</th>
                <th className="px-4 py-3.5 text-emerald-700 text-center">Inserted</th>
                <th className="px-4 py-3.5 text-purple-700 text-center">Updated</th>
                <th className="px-4 py-3.5 text-rose-700 text-center">Failed</th>
                <th className="px-4 py-3.5">Timestamp</th>
                <th className="px-4 py-3.5">Duration</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredLogs.map((log) => (
                <tr key={log.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-5 py-3.5">
                    <div className="font-semibold text-slate-900">{log.filename}</div>
                    <div className="font-mono text-[10px] text-slate-500 mt-0.5">
                      {log.fileHash.substring(0, 12)}...{log.fileHash.substring(log.fileHash.length - 8)}
                    </div>
                    {log.errorSummary && (
                      <div className="text-[11px] text-amber-700 mt-0.5">
                        {log.errorSummary}
                      </div>
                    )}
                  </td>

                  <td className="px-4 py-3.5 whitespace-nowrap">
                    {getStatusBadge(log.status)}
                  </td>

                  <td className="px-3 py-3.5 text-center font-mono text-slate-700">
                    {log.numberOfWorksheets}
                  </td>

                  <td className="px-4 py-3.5 text-center font-mono font-semibold text-slate-800">
                    {log.rowsProcessed}
                  </td>

                  <td className="px-4 py-3.5 text-center font-mono font-bold text-emerald-600">
                    +{log.rowsInserted}
                  </td>

                  <td className="px-4 py-3.5 text-center font-mono font-bold text-purple-600">
                    {log.rowsUpdated}
                  </td>

                  <td className="px-4 py-3.5 text-center font-mono font-bold text-rose-600">
                    {log.rowsFailed}
                  </td>

                  <td className="px-4 py-3.5 text-slate-600 whitespace-nowrap">
                    {new Date(log.completedAt || log.startedAt).toLocaleString()}
                  </td>

                  <td className="px-4 py-3.5 font-mono text-slate-600 whitespace-nowrap">
                    {log.durationMs}ms
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
