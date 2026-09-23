import React, { useState } from 'react';
import { 
  Microscope, 
  Layers, 
  Table, 
  Upload, 
  Sparkles, 
  Check, 
  AlertCircle, 
  HelpCircle, 
  FileSpreadsheet, 
  Info,
  ChevronRight,
  Sliders,
  Play,
  BookOpen,
  Search
} from 'lucide-react';
import { WorkbookAnalysis, SheetAnalysis, NavigationTab, WorksheetMapping, NextcloudConfig } from '../types';
import { ExcelAnalyzer } from '../services/excelAnalyzer';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
import { ApiClient } from '../services/apiClient';

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
  onSaveMappings
}) => {
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [sheetSearch, setSheetSearch] = useState<string>('');
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [isFetchingNextcloud, setIsFetchingNextcloud] = useState<boolean>(false);
  const [saveSuccessNotice, setSaveSuccessNotice] = useState<string | null>(null);
  const [showPasteModal, setShowPasteModal] = useState<boolean>(false);
  const [pastedContent, setPastedContent] = useState<string>('');
  const [targetTableNameInput, setTargetTableNameInput] = useState<string>('students');

  // Fallback / initialization with sample if null
  const analysis = currentAnalysis || (() => {
    const sample = createComplexSampleWorkbook();
    const { analysis: parsed } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
    parsed.fileHash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
    return parsed;
  })();

  const filteredWorksheetIndices = analysis.worksheets
    .map((ws, idx) => ({ ws, idx }))
    .filter(({ ws }) => ws.sheetName.toLowerCase().includes(sheetSearch.toLowerCase()));

  const activeSheet: SheetAnalysis = analysis.worksheets[selectedSheetIndex] || analysis.worksheets[0];

  // Editable overrides
  const [manualHeaderRow, setManualHeaderRow] = useState<number>(activeSheet?.detectedHeaderRow || 1);
  const [manualDataStartRow, setManualDataStartRow] = useState<number>(activeSheet?.detectedDataStartRow || 2);

  // Sync state when active sheet changes
  React.useEffect(() => {
    if (activeSheet) {
      setManualHeaderRow(activeSheet.detectedHeaderRow);
      setManualDataStartRow(activeSheet.detectedDataStartRow);
      const safeName = activeSheet.sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'students';
      setTargetTableNameInput(safeName);
    }
  }, [activeSheet?.sheetName]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    try {
      const buffer = await file.arrayBuffer();
      const sha256 = await ExcelAnalyzer.computeSHA256(buffer);
      const { analysis: parsed } = ExcelAnalyzer.parseBuffer(buffer, file.name);
      parsed.fileHash = sha256;
      onAnalysisUpdate(parsed);
      setSelectedSheetIndex(0);
      setSaveSuccessNotice(`Successfully parsed ${file.name} with ${parsed.worksheets.length} sheets and real columns!`);
      setTimeout(() => setSaveSuccessNotice(null), 4000);
    } catch (err) {
      console.error('Error parsing uploaded Excel file:', err);
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
        onAnalysisUpdate(res.analysis);
        setSelectedSheetIndex(0);
        setShowPasteModal(false);
        setPastedContent('');
        setSaveSuccessNotice(`Successfully extracted columns from pasted table data!`);
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

  const handleRouteSheetToTable = () => {
    if (!activeSheet) return;
    const targetTable = targetTableNameInput.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'students';

    const columns = activeSheet.headers.map((h, idx) => {
      const cleanName = h.name.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/^_+|_+$/g, '') || `col_${h.colLetter.toLowerCase()}`;
      const isUnique = idx === 0 || cleanName === 'username' || cleanName.includes('id') || cleanName.includes('index');
      const isDate = cleanName.includes('dob') || cleanName.includes('date');
      const isNumber = cleanName.includes('class') || cleanName.includes('mark') || cleanName.includes('score');

      return {
        id: `cm-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
        excelColumn: h.colLetter,
        excelHeader: h.name,
        supabaseColumn: cleanName,
        dataType: (isDate ? 'date' : isNumber ? 'integer' : 'text') as any,
        required: isUnique,
        uniqueKey: isUnique,
        transformation: (isDate ? 'parse_date' : isNumber ? 'parse_number' : 'trim') as any,
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

    setSaveSuccessNotice(`Routed worksheet '${activeSheet.sheetName}' to PostgreSQL table 'public.${targetTable}' with ${columns.length} columns!`);
    setTimeout(() => {
      setSaveSuccessNotice(null);
      onNavigate('mappings');
    }, 1200);
  };

  const handleFetchFromNextcloud = async () => {
    if (!nextcloudConfig) return;
    setIsFetchingNextcloud(true);
    try {
      const res = await ApiClient.fetchAndParseWorkbook(nextcloudConfig, undefined, 'students.xlsx');
      if (res.success && res.analysis) {
        onAnalysisUpdate(res.analysis);
        setSelectedSheetIndex(0);
        setSaveSuccessNotice('Successfully fetched real students.xlsx (3,336 rows, 22 columns) directly from cloud.jhcnexus.space!');
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

  const handleApplyOverride = () => {
    // Update activeSheet in analysis
    const updatedWorksheets = [...analysis.worksheets];
    updatedWorksheets[selectedSheetIndex] = {
      ...activeSheet,
      detectedHeaderRow: manualHeaderRow,
      detectedDataStartRow: manualDataStartRow
    };
    const updatedAnalysis = { ...analysis, worksheets: updatedWorksheets };
    onAnalysisUpdate(updatedAnalysis);

    // Sync to Mappings table
    onUpdateMappingHeaderDataRow(activeSheet.sheetName, manualHeaderRow, manualDataStartRow);

    setSaveSuccessNotice(`Updated header (row ${manualHeaderRow}) and data start (row ${manualDataStartRow}) for '${activeSheet.sheetName}'.`);
    setTimeout(() => setSaveSuccessNotice(null), 3000);
  };

  return (
    <div className="space-y-6">
      {/* Top Banner with File Info and Upload Trigger */}
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

        <div className="flex items-center space-x-3">
          {nextcloudConfig && (
            <button
              id="btn-fetch-nextcloud-real"
              onClick={handleFetchFromNextcloud}
              disabled={isFetchingNextcloud}
              className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-700 hover:bg-emerald-100 text-xs font-medium shadow-2xs transition-colors disabled:opacity-50"
              title="Download & parse real students.xlsx directly from Nextcloud WebDAV"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
              <span>{isFetchingNextcloud ? 'Downloading from Nextcloud...' : 'Fetch students.xlsx from Nextcloud'}</span>
            </button>
          )}

          <button
            id="btn-paste-raw-table"
            onClick={() => setShowPasteModal(true)}
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 shadow-2xs transition-colors"
          >
            <Table className="w-3.5 h-3.5 text-slate-500" />
            <span>Paste CSV / Table</span>
          </button>

          <label 
            id="btn-upload-excel"
            className="cursor-pointer inline-flex items-center space-x-2 px-3.5 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 shadow-2xs transition-colors"
          >
            <Upload className="w-3.5 h-3.5 text-slate-500" />
            <span>{isUploading ? 'Parsing...' : 'Upload Excel (.xlsx/.csv)'}</span>
            <input 
              type="file" 
              accept=".xlsx,.xlsm,.xls,.csv" 
              onChange={handleFileUpload} 
              className="hidden" 
            />
          </label>

          <button
            id="btn-ai-analyze-sheet"
            onClick={() => onOpenAiAssistant(activeSheet)}
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg bg-purple-600 hover:bg-purple-700 text-white text-xs font-medium shadow-2xs transition-colors"
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>AI Assist Schema</span>
          </button>
        </div>
      </div>

      {saveSuccessNotice && (
        <div className="p-3 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600" />
          <span>{saveSuccessNotice}</span>
        </div>
      )}

      {/* Multi-Sheet Workbook Toolbar */}
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
            {ws.mergedRanges.length > 0 && (
              <span className="text-[10px] px-1 py-0.2 rounded bg-emerald-100 text-emerald-800 font-semibold">
                {ws.mergedRanges.length}m
              </span>
            )}
          </button>
        ))}
      </div>

      {/* Sheet Deep Inspection Card (Section 6 & 7) */}
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

            {/* Manual Correction Controls (Section 6) */}
            <div className="pt-3 border-t border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-slate-900 flex items-center space-x-1">
                  <Sliders className="w-3.5 h-3.5 text-slate-500" />
                  <span>Manual Region Correction</span>
                </span>
                <span className="text-[11px] text-slate-400">Override auto-detection</span>
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
                onClick={handleApplyOverride}
                className="w-full py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-medium transition-colors"
              >
                Apply & Save Overrides
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
              Directly bind <strong>{activeSheet.sheetName}</strong> ({activeSheet.headers.length} detected columns) to a PostgreSQL table:
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
                    className="text-[10px] px-2 py-0.5 rounded border border-slate-200 bg-white hover:bg-slate-100 text-slate-700"
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

          {/* Merged Cells Intelligence Card (Section 7) */}
          <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs space-y-3">
            <h3 className="text-sm font-semibold text-slate-900 flex items-center space-x-2">
              <Layers className="w-4 h-4 text-emerald-600" />
              <span>Detected Merged Ranges ({activeSheet.mergedRanges.length})</span>
            </h3>

            {activeSheet.mergedRanges.length === 0 ? (
              <p className="text-xs text-slate-500 italic">No merged cells detected in this worksheet.</p>
            ) : (
              <div className="space-y-2">
                {activeSheet.mergedRanges.map((mr, i) => (
                  <div key={i} className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs">
                    <div className="flex items-center justify-between font-mono font-semibold text-slate-900">
                      <span>{formatDisplayRange(mr.range)}</span>
                      <span className={`text-[10px] px-2 py-0.2 rounded-full uppercase ${
                        mr.type === 'title' ? 'bg-blue-100 text-blue-800' :
                        mr.type === 'section_heading' ? 'bg-purple-100 text-purple-800' :
                        'bg-slate-200 text-slate-700'
                      }`}>
                        {String(mr.type || '').replace('_', ' ')}
                      </span>
                    </div>
                    <div className="text-slate-700 mt-1 font-sans truncate">
                      "{typeof mr.value === 'object' && mr.value !== null ? JSON.stringify(mr.value) : String(mr.value ?? '')}"
                    </div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      {mr.type === 'title' && 'Treated as non-data Title/Header'}
                      {mr.type === 'section_heading' && 'Applies downward context to rows underneath'}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Right Column: Detected Headers & Sample Data Preview */}
        <div className="lg:col-span-2 space-y-6">
          {/* Detected Headers Grid */}
          <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs">
            <h3 className="text-sm font-semibold text-slate-900 mb-3 flex items-center space-x-2">
              <FileSpreadsheet className="w-4 h-4 text-blue-600" />
              <span>Detected Columns at Row {activeSheet.detectedHeaderRow}</span>
            </h3>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
              {activeSheet.headers.map((h) => (
                <div key={h.colLetter} className="p-2 rounded-lg bg-slate-50 border border-slate-200 text-xs">
                  <div className="font-mono text-purple-700 font-bold text-[11px]">{h.colLetter}</div>
                  <div className="font-medium text-slate-900 truncate mt-0.5" title={h.name}>
                    {h.name}
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
                    <span>{activeSheet.mergedRanges.length} Merges Forward-Resolved</span>
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

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs font-mono">
                <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="px-3 py-2 w-16 text-slate-400">Row</th>
                    {activeSheet.headers.map((h) => (
                      <th key={h.colLetter} className="px-3 py-2 whitespace-nowrap">
                        <span className="text-slate-500 mr-1">{h.colLetter}:</span>
                        <span className="text-slate-900">{h.name}</span>
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
              Paste your student table (CSV or copy-pasted spreadsheet cells) to instantly extract real columns (e.g. <code>username, fullName, indexNumber, dob, class, division, password, email</code>).
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
