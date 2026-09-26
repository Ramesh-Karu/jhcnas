import React, { useState } from 'react';
import { 
  FileSpreadsheet, 
  Folder, 
  CheckCircle2, 
  Play, 
  Microscope, 
  GitFork, 
  RefreshCw,
  Info,
  Zap,
  Clock,
  SlidersHorizontal,
  Check,
  Ban,
  ShieldCheck
} from 'lucide-react';
import { NextcloudFile, NavigationTab } from '../types';

interface ExcelFilesViewProps {
  files: NextcloudFile[];
  excludedAutoSyncFiles?: string[];
  onToggleFileAutoSync?: (filename: string, enabled: boolean) => void;
  onSyncFileNow?: (file: NextcloudFile) => Promise<void> | void;
  onSelectFileForAnalysis: (file: NextcloudFile, targetTab?: NavigationTab) => void;
  onNavigate: (tab: NavigationTab) => void;
  onRefreshFiles: () => void;
  syncingFileId?: string | null;
}

export const ExcelFilesView: React.FC<ExcelFilesViewProps> = ({
  files,
  excludedAutoSyncFiles = [],
  onToggleFileAutoSync,
  onSyncFileNow,
  onSelectFileForAnalysis,
  onNavigate,
  onRefreshFiles,
  syncingFileId
}) => {
  const [filterMode, setFilterMode] = useState<'ALL' | 'AUTO_ONLY' | 'MANUAL_ONLY'>('ALL');
  const [searchQuery, setSearchQuery] = useState('');
  const [feedbackMessage, setFeedbackMessage] = useState<string | null>(null);

  const excludedSet = new Set(excludedAutoSyncFiles.map(f => f.toLowerCase().trim()));

  const isFileAutoSync = (filename: string) => {
    return !excludedSet.has(filename.toLowerCase().trim());
  };

  const autoCount = files.filter(f => isFileAutoSync(f.filename)).length;
  const manualCount = files.filter(f => !isFileAutoSync(f.filename)).length;

  const filteredFiles = files.filter(file => {
    const isAuto = isFileAutoSync(file.filename);
    if (filterMode === 'AUTO_ONLY' && !isAuto) return false;
    if (filterMode === 'MANUAL_ONLY' && isAuto) return false;
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      return file.filename.toLowerCase().includes(q) || file.path.toLowerCase().includes(q);
    }
    return true;
  });

  const handleToggle = (filename: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    const currentAuto = isFileAutoSync(filename);
    const nextAuto = !currentAuto;
    if (onToggleFileAutoSync) {
      onToggleFileAutoSync(filename, nextAuto);
      setFeedbackMessage(
        nextAuto 
          ? `🟢 Enabled Auto-Sync for "${filename}". It will be processed automatically on scheduled cycles.` 
          : `⏸️ Stopped Auto-Sync for "${filename}". It is now set to Manual Sync Only.`
      );
      setTimeout(() => setFeedbackMessage(null), 4500);
    }
  };

  const handleEnableAll = () => {
    files.forEach(f => {
      if (!isFileAutoSync(f.filename) && onToggleFileAutoSync) {
        onToggleFileAutoSync(f.filename, true);
      }
    });
    setFeedbackMessage('🟢 Auto-Sync enabled for all detected Excel files.');
    setTimeout(() => setFeedbackMessage(null), 4000);
  };

  const handleDisableAll = () => {
    files.forEach(f => {
      if (isFileAutoSync(f.filename) && onToggleFileAutoSync) {
        onToggleFileAutoSync(f.filename, false);
      }
    });
    setFeedbackMessage('⏸️ All Excel files switched to Manual Sync Only. Auto-sync is paused for all workbooks.');
    setTimeout(() => setFeedbackMessage(null), 4000);
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-emerald-50 text-emerald-600">
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Detected Nextcloud Excel Files</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Workbooks discovered via WebDAV in <code className="bg-slate-100 px-1 py-0.5 rounded font-mono">/ExcelImports</code>. Control which files sync automatically vs manual-only.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-2">
          <button
            onClick={onRefreshFiles}
            className="inline-flex items-center space-x-2 px-3.5 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 transition-colors shadow-2xs cursor-pointer"
          >
            <RefreshCw className="w-3.5 h-3.5 text-slate-500" />
            <span>Rescan WebDAV</span>
          </button>
        </div>
      </div>

      {/* Stats & Sync Mode Summary Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Total Discovered</span>
            <FileSpreadsheet className="w-4 h-4 text-slate-400" />
          </div>
          <div className="text-2xl font-bold text-slate-900 mt-1">{files.length}</div>
          <div className="text-[11px] text-slate-500 mt-0.5">Workbooks in WebDAV folder</div>
        </div>

        <div className="bg-white p-4 rounded-xl border border-emerald-200 shadow-2xs bg-emerald-50/20">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-emerald-700 uppercase tracking-wider flex items-center space-x-1">
              <Zap className="w-3.5 h-3.5 text-emerald-600" />
              <span>Auto-Sync Active</span>
            </span>
            <span className="text-[10px] bg-emerald-100 text-emerald-800 font-bold px-1.5 py-0.5 rounded">
              Scheduled
            </span>
          </div>
          <div className="text-2xl font-bold text-emerald-700 mt-1">{autoCount}</div>
          <div className="text-[11px] text-emerald-600/80 mt-0.5">Syncs automatically every timer cycle</div>
        </div>

        <div className="bg-white p-4 rounded-xl border border-amber-200 shadow-2xs bg-amber-50/20">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-amber-800 uppercase tracking-wider flex items-center space-x-1">
              <Clock className="w-3.5 h-3.5 text-amber-600" />
              <span>Manual Only (Excluded)</span>
            </span>
            <span className="text-[10px] bg-amber-100 text-amber-800 font-bold px-1.5 py-0.5 rounded">
              Paused Auto
            </span>
          </div>
          <div className="text-2xl font-bold text-amber-700 mt-1">{manualCount}</div>
          <div className="text-[11px] text-amber-700/80 mt-0.5">Skipped on timer; synced only on demand</div>
        </div>
      </div>

      {/* Feedback Banner */}
      {feedbackMessage && (
        <div className="p-3.5 rounded-xl bg-slate-900 text-white text-xs flex items-center justify-between animate-fadeIn shadow-md">
          <div className="flex items-center space-x-2">
            <ShieldCheck className="w-4 h-4 text-emerald-400 shrink-0" />
            <span>{feedbackMessage}</span>
          </div>
          <button 
            onClick={() => setFeedbackMessage(null)}
            className="text-slate-400 hover:text-white text-xs px-2 py-0.5 rounded cursor-pointer"
          >
            ✕
          </button>
        </div>
      )}

      {/* Filter & Batch Controls Bar */}
      <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-3">
        {/* Filter Tabs */}
        <div className="flex items-center space-x-1 bg-slate-100 p-1 rounded-lg">
          <button
            onClick={() => setFilterMode('ALL')}
            className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
              filterMode === 'ALL'
                ? 'bg-white text-slate-900 shadow-2xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            All Files ({files.length})
          </button>
          <button
            onClick={() => setFilterMode('AUTO_ONLY')}
            className={`inline-flex items-center space-x-1 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
              filterMode === 'AUTO_ONLY'
                ? 'bg-white text-emerald-700 shadow-2xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block"></span>
            <span>Auto-Sync Active ({autoCount})</span>
          </button>
          <button
            onClick={() => setFilterMode('MANUAL_ONLY')}
            className={`inline-flex items-center space-x-1 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
              filterMode === 'MANUAL_ONLY'
                ? 'bg-white text-amber-700 shadow-2xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <span className="w-2 h-2 rounded-full bg-amber-500 inline-block"></span>
            <span>Manual Only ({manualCount})</span>
          </button>
        </div>

        {/* Search & Batch Actions */}
        <div className="flex flex-wrap items-center gap-2">
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search files..."
            className="px-3 py-1.5 rounded-lg border border-slate-300 text-xs bg-white text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none w-44"
          />

          <button
            onClick={handleEnableAll}
            className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-lg border border-emerald-300 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-semibold transition-colors cursor-pointer"
            title="Enable auto-sync on all files"
          >
            <Check className="w-3.5 h-3.5" />
            <span>Enable All Auto</span>
          </button>

          <button
            onClick={handleDisableAll}
            className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-lg border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-semibold transition-colors cursor-pointer"
            title="Switch all files to manual sync only"
          >
            <Ban className="w-3.5 h-3.5" />
            <span>Set All Manual Only</span>
          </button>
        </div>
      </div>

      {/* Files Table */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100/80 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200">
              <tr>
                <th className="px-5 py-3.5">Filename & Path</th>
                <th className="px-4 py-3.5">Auto-Sync Control</th>
                <th className="px-4 py-3.5">Size</th>
                <th className="px-4 py-3.5">SHA-256 Hash</th>
                <th className="px-4 py-3.5">Last Modified</th>
                <th className="px-4 py-3.5">Sync Status</th>
                <th className="px-5 py-3.5 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredFiles.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-5 py-10 text-center text-slate-400">
                    <FileSpreadsheet className="w-8 h-8 mx-auto mb-2 text-slate-300" />
                    <p className="font-semibold text-slate-600">No Excel files match current filter</p>
                    <p className="text-xs text-slate-400 mt-0.5">Try changing your search query or filter tab</p>
                  </td>
                </tr>
              ) : (
                filteredFiles.map((file, fIdx) => {
                  const isAuto = isFileAutoSync(file.filename);
                  const isSyncingThis = syncingFileId === file.id || syncingFileId === file.filename;

                  return (
                    <tr 
                      key={`${file.id}-${fIdx}`} 
                      onClick={() => onSelectFileForAnalysis(file, 'analyzer')}
                      className={`hover:bg-slate-50/80 transition-colors cursor-pointer ${
                        !isAuto ? 'bg-slate-50/30' : ''
                      }`}
                    >
                      <td className="px-5 py-4">
                        <div className="flex items-center space-x-3">
                          <div className={`p-2 rounded-lg ${isAuto ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}`}>
                            <FileSpreadsheet className="w-5 h-5" />
                          </div>
                          <div>
                            <div className="font-semibold text-slate-900 text-sm hover:text-emerald-700 transition-colors flex items-center space-x-2">
                              <span>{file.filename}</span>
                              {!isAuto && (
                                <span className="text-[10px] font-bold px-1.5 py-0.2 rounded bg-amber-100 text-amber-800 border border-amber-200">
                                  Manual Only
                                </span>
                              )}
                            </div>
                            <div className="text-[11px] text-slate-500 font-mono flex items-center space-x-1 mt-0.5">
                              <Folder className="w-3 h-3 text-slate-400" />
                              <span>{file.path}</span>
                            </div>
                          </div>
                        </div>
                      </td>

                      {/* Auto-Sync Toggle Control */}
                      <td className="px-4 py-4 whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                        <div className="flex items-center space-x-2.5">
                          <button
                            type="button"
                            role="switch"
                            aria-checked={isAuto}
                            onClick={(e) => handleToggle(file.filename, e)}
                            className={`relative inline-flex h-5 w-9 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${
                              isAuto ? 'bg-emerald-600' : 'bg-slate-300'
                            }`}
                            title={isAuto ? 'Auto-Sync is ACTIVE (Click to stop auto-sync for this file)' : 'Auto-Sync is STOPPED (Click to enable auto-sync for this file)'}
                          >
                            <span
                              className={`pointer-events-none inline-block h-4 w-4 transform rounded-full bg-white shadow-sm ring-0 transition duration-200 ease-in-out ${
                                isAuto ? 'translate-x-4' : 'translate-x-0'
                              }`}
                            />
                          </button>

                          <span className={`text-xs font-semibold ${isAuto ? 'text-emerald-700' : 'text-slate-500'}`}>
                            {isAuto ? 'Auto (15m)' : 'Manual Only'}
                          </span>
                        </div>
                      </td>

                      <td className="px-4 py-4 text-slate-700 font-medium whitespace-nowrap">
                        {file.fileSizeFormatted}
                      </td>

                      <td className="px-4 py-4 font-mono text-[11px] text-slate-600 whitespace-nowrap">
                        <span className="bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200" title={file.fileHash}>
                          {file.fileHash ? `${file.fileHash.substring(0, 10)}...${file.fileHash.substring(file.fileHash.length - 6)}` : 'hash-pending'}
                        </span>
                      </td>

                      <td className="px-4 py-4 text-slate-600 whitespace-nowrap">
                        {file.lastModified ? new Date(file.lastModified).toLocaleString() : 'Recent'}
                      </td>

                      <td className="px-4 py-4 whitespace-nowrap">
                        <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold ${
                          file.status === 'Synced' ? 'bg-emerald-100 text-emerald-800' :
                          file.status === 'Unchanged' ? 'bg-slate-100 text-slate-700' :
                          file.status === 'Modified' ? 'bg-amber-100 text-amber-800' :
                          'bg-blue-100 text-blue-800'
                        }`}>
                          {file.status === 'Synced' && <CheckCircle2 className="w-3 h-3 mr-1 text-emerald-600" />}
                          {file.status}
                        </span>
                      </td>

                      <td className="px-5 py-4 text-right whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                        <div className="inline-flex items-center space-x-1.5">
                          {/* Manual Single-File Sync Now button */}
                          {onSyncFileNow && (
                            <button
                              id={`btn-sync-single-${file.id || fIdx}`}
                              onClick={() => onSyncFileNow(file)}
                              disabled={isSyncingThis}
                              className={`inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-md text-xs font-bold transition-all shadow-2xs cursor-pointer ${
                                isSyncingThis
                                  ? 'bg-slate-200 text-slate-500 cursor-not-allowed'
                                  : 'bg-indigo-600 hover:bg-indigo-700 text-white hover:scale-[1.02]'
                              }`}
                              title={`Execute immediate manual sync for ${file.filename} into Supabase`}
                            >
                              <Zap className={`w-3.5 h-3.5 ${isSyncingThis ? 'animate-spin' : ''}`} />
                              <span>{isSyncingThis ? 'Syncing...' : 'Sync Now'}</span>
                            </button>
                          )}

                          <button
                            id={`btn-analyze-${file.id || fIdx}`}
                            onClick={() => {
                              onSelectFileForAnalysis(file, 'analyzer');
                            }}
                            className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-md border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-medium shadow-2xs cursor-pointer"
                            title="Inspect worksheets, merged cells, and headers"
                          >
                            <Microscope className="w-3.5 h-3.5 text-slate-500" />
                            <span>Analyze</span>
                          </button>

                          <button
                            id={`btn-map-${file.id || fIdx}`}
                            onClick={() => {
                              onSelectFileForAnalysis(file, 'mappings');
                            }}
                            className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-md border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-medium shadow-2xs cursor-pointer"
                            title="Configure worksheet to table mappings"
                          >
                            <GitFork className="w-3.5 h-3.5 text-slate-500" />
                            <span>Map</span>
                          </button>

                          <button
                            id={`btn-dryrun-${file.id || fIdx}`}
                            onClick={() => {
                              onSelectFileForAnalysis(file, 'import');
                            }}
                            className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-medium shadow-2xs cursor-pointer"
                            title="Run Dry Run simulation"
                          >
                            <Play className="w-3.5 h-3.5" />
                            <span>Dry Run</span>
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

