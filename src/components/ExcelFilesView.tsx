import React from 'react';
import { 
  FileSpreadsheet, 
  Folder, 
  Hash, 
  Clock, 
  CheckCircle2, 
  AlertCircle, 
  Play, 
  Microscope, 
  GitFork, 
  RefreshCw,
  Info
} from 'lucide-react';
import { NextcloudFile, NavigationTab } from '../types';

interface ExcelFilesViewProps {
  files: NextcloudFile[];
  onSelectFileForAnalysis: (file: NextcloudFile, targetTab?: NavigationTab) => void;
  onNavigate: (tab: NavigationTab) => void;
  onRefreshFiles: () => void;
}

export const ExcelFilesView: React.FC<ExcelFilesViewProps> = ({
  files,
  onSelectFileForAnalysis,
  onNavigate,
  onRefreshFiles
}) => {
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
              Workbooks discovered via WebDAV in <code className="bg-slate-100 px-1 py-0.5 rounded font-mono">/ExcelImports</code> on TrueNAS SCALE.
            </p>
          </div>
        </div>

        <button
          onClick={onRefreshFiles}
          className="inline-flex items-center space-x-2 px-3.5 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 transition-colors shadow-2xs"
        >
          <RefreshCw className="w-3.5 h-3.5 text-slate-500" />
          <span>Rescan WebDAV Folder</span>
        </button>
      </div>

      {/* File Change Detection Explanation Card */}
      <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 text-xs text-slate-600 flex items-start space-x-3">
        <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
        <div>
          <span className="font-semibold text-slate-800">SHA-256 File Change Detection (Section 14): </span>
          The Coolify worker computes a cryptographic SHA-256 hash upon downloading each file. Unchanged files with matching hashes are skipped automatically to avoid redundant imports and prevent duplicate records.
        </div>
      </div>

      {/* Files Table */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100/80 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200">
              <tr>
                <th className="px-5 py-3.5">Filename & Path</th>
                <th className="px-4 py-3.5">Size</th>
                <th className="px-4 py-3.5">SHA-256 Hash</th>
                <th className="px-4 py-3.5">Last Modified</th>
                <th className="px-4 py-3.5">Sync Status</th>
                <th className="px-5 py-3.5 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {files.map((file) => (
                <tr 
                  key={file.id} 
                  onClick={() => onSelectFileForAnalysis(file, 'analyzer')}
                  className="hover:bg-slate-50/80 transition-colors cursor-pointer"
                >
                  <td className="px-5 py-4">
                    <div className="flex items-center space-x-3">
                      <div className="p-2 rounded-lg bg-emerald-50 text-emerald-700">
                        <FileSpreadsheet className="w-5 h-5" />
                      </div>
                      <div>
                        <div className="font-semibold text-slate-900 text-sm hover:text-emerald-700 transition-colors">{file.filename}</div>
                        <div className="text-[11px] text-slate-500 font-mono flex items-center space-x-1 mt-0.5">
                          <Folder className="w-3 h-3 text-slate-400" />
                          <span>{file.path}</span>
                        </div>
                      </div>
                    </div>
                  </td>

                  <td className="px-4 py-4 text-slate-700 font-medium whitespace-nowrap">
                    {file.fileSizeFormatted}
                  </td>

                  <td className="px-4 py-4 font-mono text-[11px] text-slate-600 whitespace-nowrap">
                    <span className="bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200" title={file.fileHash}>
                      {file.fileHash.substring(0, 10)}...{file.fileHash.substring(file.fileHash.length - 6)}
                    </span>
                  </td>

                  <td className="px-4 py-4 text-slate-600 whitespace-nowrap">
                    {new Date(file.lastModified).toLocaleString()}
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
                    <div className="inline-flex items-center space-x-2">
                      <button
                        id={`btn-analyze-${file.id}`}
                        onClick={() => {
                          onSelectFileForAnalysis(file, 'analyzer');
                        }}
                        className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-md border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-medium shadow-2xs"
                        title="Inspect worksheets, merged cells, and headers"
                      >
                        <Microscope className="w-3.5 h-3.5 text-slate-500" />
                        <span>Analyze</span>
                      </button>

                      <button
                        id={`btn-map-${file.id}`}
                        onClick={() => {
                          onSelectFileForAnalysis(file, 'mappings');
                        }}
                        className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-md border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-medium shadow-2xs"
                        title="Configure worksheet to table mappings"
                      >
                        <GitFork className="w-3.5 h-3.5 text-slate-500" />
                        <span>Map</span>
                      </button>

                      <button
                        id={`btn-dryrun-${file.id}`}
                        onClick={() => {
                          onSelectFileForAnalysis(file, 'import');
                        }}
                        className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-medium shadow-2xs"
                        title="Run Dry Run simulation"
                      >
                        <Play className="w-3.5 h-3.5" />
                        <span>Dry Run</span>
                      </button>
                    </div>
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
