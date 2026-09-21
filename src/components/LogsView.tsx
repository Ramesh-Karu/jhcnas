import React, { useState } from 'react';
import { 
  Terminal, 
  Trash2, 
  Download, 
  Filter, 
  Search, 
  Cpu, 
  RefreshCw 
} from 'lucide-react';
import { LogMessage } from '../types';

interface LogsViewProps {
  logs: LogMessage[];
  onClearLogs: () => void;
  onRefreshLogs: () => void;
}

export const LogsView: React.FC<LogsViewProps> = ({
  logs,
  onClearLogs,
  onRefreshLogs
}) => {
  const [filterLevel, setFilterLevel] = useState<string>('ALL');
  const [search, setSearch] = useState('');

  const filtered = logs.filter(l => {
    const matchLevel = filterLevel === 'ALL' || l.level === filterLevel;
    const matchSearch = l.message.toLowerCase().includes(search.toLowerCase()) ||
                        l.component.toLowerCase().includes(search.toLowerCase());
    return matchLevel && matchSearch;
  });

  const getLevelStyle = (level: LogMessage['level']) => {
    switch (level) {
      case 'info': return 'text-blue-400 font-semibold';
      case 'warn': return 'text-amber-400 font-semibold';
      case 'error': return 'text-rose-400 font-semibold';
      case 'success': return 'text-emerald-400 font-semibold';
    }
  };

  const handleDownload = () => {
    const text = filtered.map(l => `[${l.timestamp}] [${l.level.toUpperCase()}] [${l.component}] ${l.message}`).join('\n');
    const blob = new Blob([text], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `sync_worker_logs_${new Date().toISOString().split('T')[0]}.txt`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-slate-900 text-emerald-400">
            <Terminal className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Worker Synchronization Logs</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Live audit stream from the Coolify background worker executing WebDAV scans and openpyxl transformations.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-2">
          <button
            onClick={onRefreshLogs}
            className="p-2 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50 transition-colors"
            title="Refresh logs"
          >
            <RefreshCw className="w-4 h-4" />
          </button>
          <button
            onClick={handleDownload}
            className="inline-flex items-center space-x-1 px-3 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 shadow-2xs"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Download</span>
          </button>
          <button
            onClick={onClearLogs}
            className="inline-flex items-center space-x-1 px-3 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 shadow-2xs"
          >
            <Trash2 className="w-3.5 h-3.5" />
            <span>Clear</span>
          </button>
        </div>
      </div>

      {/* Filter Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
        <div className="flex items-center space-x-2">
          <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Filter Level:</span>
          {(['ALL', 'info', 'warn', 'error', 'success'] as const).map(lvl => (
            <button
              key={lvl}
              onClick={() => setFilterLevel(lvl)}
              className={`px-2.5 py-1 rounded-md text-xs font-medium capitalize transition-colors ${
                filterLevel === lvl ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
              }`}
            >
              {lvl}
            </button>
          ))}
        </div>

        <div className="relative">
          <input
            type="text"
            placeholder="Search log messages..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="pl-7 pr-3 py-1 text-xs rounded-lg border border-slate-300 w-56 focus:outline-none focus:ring-1 focus:ring-emerald-500"
          />
          <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2 top-2" />
        </div>
      </div>

      {/* Terminal Viewport */}
      <div className="bg-slate-950 text-slate-200 rounded-xl p-5 border border-slate-800 font-mono text-xs shadow-lg space-y-2 max-h-[540px] overflow-y-auto">
        <div className="flex items-center justify-between pb-3 mb-2 border-b border-slate-800 text-slate-500 text-[11px]">
          <div className="flex items-center space-x-2">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
            <span>coolify-worker:8000 (Python 3.11 / FastAPI / APScheduler)</span>
          </div>
          <span>Showing {filtered.length} entries</span>
        </div>

        {filtered.length === 0 ? (
          <div className="py-8 text-center text-slate-500 italic">No log entries matching criteria.</div>
        ) : (
          filtered.map((l) => (
            <div key={l.id} className="leading-relaxed hover:bg-slate-900/60 px-2 py-1 rounded transition-colors flex items-start space-x-3">
              <span className="text-slate-400 shrink-0 select-none">
                {new Date(l.timestamp).toLocaleTimeString()}
              </span>
              <span className={`uppercase text-[10px] px-1.5 py-0.2 rounded border border-slate-800 shrink-0 ${getLevelStyle(l.level)}`}>
                {l.level}
              </span>
              <span className="text-purple-400 shrink-0 font-semibold">
                [{l.component}]
              </span>
              <span className="text-slate-300 break-all">
                {l.message}
              </span>
            </div>
          ))
        )}
      </div>
    </div>
  );
};
