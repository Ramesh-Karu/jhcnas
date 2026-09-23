import React, { useState, useEffect } from 'react';
import { 
  GitFork, 
  Plus, 
  Trash2, 
  Save, 
  RotateCcw, 
  Check, 
  HelpCircle, 
  Play, 
  Sparkles,
  Layers,
  Table,
  CheckCircle2,
  BookOpen,
  Search,
  ToggleLeft,
  ToggleRight,
  EyeOff,
  Eye,
  Info
} from 'lucide-react';
import { 
  WorksheetMapping, 
  ColumnMapping, 
  TransformationType, 
  DataType, 
  NavigationTab,
  WorkbookAnalysis,
  SheetAnalysis,
  SheetHeader,
  TableSyncPolicy
} from '../types';
import { getDefaultSampleMappings } from '../services/sampleWorkbook';

interface MappingsViewProps {
  mappings: WorksheetMapping[];
  onSaveMappings: (mappings: WorksheetMapping[]) => void;
  onNavigate: (tab: NavigationTab) => void;
  onOpenAiAssistant: () => void;
  currentAnalysis?: WorkbookAnalysis | null;
}

export const MappingsView: React.FC<MappingsViewProps> = ({
  mappings,
  onSaveMappings,
  onNavigate,
  onOpenAiAssistant,
  currentAnalysis
}) => {
  const [activeSheetId, setActiveSheetId] = useState<string>(mappings[0]?.id || '');
  const [currentMappings, setCurrentMappings] = useState<WorksheetMapping[]>(mappings);
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const [sheetSearch, setSheetSearch] = useState<string>('');

  const activeMapping = currentMappings.find(m => m.id === activeSheetId) || currentMappings[0];

  // Keep activeSheetId valid if currentMappings changes
  useEffect(() => {
    if (currentMappings.length > 0 && !currentMappings.some(m => m.id === activeSheetId)) {
      setActiveSheetId(currentMappings[0].id);
    }
  }, [currentMappings, activeSheetId]);

  const transformationOptions: { value: TransformationType; label: string }[] = [
    { value: 'none', label: 'None' },
    { value: 'trim', label: 'Trim Whitespace' },
    { value: 'uppercase', label: 'Uppercase' },
    { value: 'lowercase', label: 'Lowercase' },
    { value: 'parse_date', label: 'Parse Date (YYYY-MM-DD)' },
    { value: 'parse_number', label: 'Parse Number' },
    { value: 'yes_no_to_boolean', label: 'Yes/No to Boolean' },
    { value: 'pa_to_status', label: 'P/A/L/E to Full Status' },
    { value: 'normalize_phone', label: 'Normalize Phone Number' },
    { value: 'normalize_id', label: 'Normalize ID (Remove spaces/dashes)' }
  ];

  const dataTypeOptions: DataType[] = ['text', 'integer', 'decimal', 'boolean', 'date', 'timestamp', 'json'];

  const handleUpdateActiveMapping = (updates: Partial<WorksheetMapping>) => {
    if (!activeMapping) return;
    const updated = currentMappings.map(m => m.id === activeMapping.id ? { ...m, ...updates } : m);
    setCurrentMappings(updated);
  };

  const handleToggleActiveMappingEnabled = () => {
    if (!activeMapping) return;
    handleUpdateActiveMapping({ enabled: !activeMapping.enabled });
  };

  const handleDeleteActiveMapping = () => {
    if (!activeMapping || currentMappings.length <= 1) return;
    const remaining = currentMappings.filter(m => m.id !== activeMapping.id);
    setCurrentMappings(remaining);
    setActiveSheetId(remaining[0].id);
    onSaveMappings(remaining);
    setSaveMessage(`Removed mapping for ${activeMapping.worksheetName}`);
    setTimeout(() => setSaveMessage(null), 3000);
  };

  const handleAutoGenerateAllSheets = () => {
    if (!currentAnalysis || !currentAnalysis.worksheets.length) return;
    const existingNames = new Set(currentMappings.map(m => m.worksheetName.toLowerCase()));
    const newMappings: WorksheetMapping[] = [...currentMappings];

    let addedCount = 0;
    for (const ws of currentAnalysis.worksheets) {
      if (ws.totalRows <= 1 && ws.headers.length === 0) continue;
      if (existingNames.has(ws.sheetName.toLowerCase())) continue;

      const safeTableName = ws.sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/^_+|_+$/g, '') || 'sheet_data';
      
      const columns: ColumnMapping[] = ws.headers.map((h: SheetHeader, idx: number) => {
        const safeCol = h.name.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/^_+|_+$/g, '') || `col_${h.colLetter.toLowerCase()}`;
        const isId = idx === 0 || safeCol.includes('id') || safeCol.includes('code') || safeCol.includes('number');
        return {
          id: `cm-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
          excelColumn: h.colLetter,
          excelHeader: h.name,
          supabaseColumn: safeCol,
          dataType: 'text',
          required: isId,
          uniqueKey: isId && idx === 0,
          transformation: 'trim'
        };
      });

      newMappings.push({
        id: `wm-${Date.now()}-${Math.random().toString(36).substr(2, 6)}`,
        workbookName: currentAnalysis.filename,
        worksheetName: ws.sheetName,
        supabaseTable: safeTableName,
        headerRow: ws.detectedHeaderRow || 1,
        dataStartRow: ws.detectedDataStartRow || 2,
        enabled: true,
        columns
      });
      addedCount++;
    }

    if (addedCount > 0) {
      setCurrentMappings(newMappings);
      onSaveMappings(newMappings);
      setSaveMessage(`Successfully generated mappings for ${addedCount} additional worksheets!`);
    } else {
      setSaveMessage('All sheets from this workbook are already mapped.');
    }
    setTimeout(() => setSaveMessage(null), 4000);
  };

  const handleAddSheetFromAnalysis = (sheetName: string) => {
    if (!currentAnalysis) return;
    const ws = currentAnalysis.worksheets.find((w: SheetAnalysis) => w.sheetName === sheetName);
    if (!ws) return;

    const safeTableName = sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/^_+|_+$/g, '') || 'sheet_data';
    const columns: ColumnMapping[] = ws.headers.map((h: SheetHeader, idx: number) => {
      const safeCol = h.name.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/^_+|_+$/g, '') || `col_${h.colLetter.toLowerCase()}`;
      return {
        id: `cm-${Date.now()}-${idx}`,
        excelColumn: h.colLetter,
        excelHeader: h.name,
        supabaseColumn: safeCol,
        dataType: 'text',
        required: idx === 0,
        uniqueKey: idx === 0,
        transformation: 'trim'
      };
    });

    const newMapping: WorksheetMapping = {
      id: `wm-${Date.now()}-${Math.random().toString(36).substr(2, 6)}`,
      workbookName: currentAnalysis.filename,
      worksheetName: sheetName,
      supabaseTable: safeTableName,
      headerRow: ws.detectedHeaderRow || 1,
      dataStartRow: ws.detectedDataStartRow || 2,
      enabled: true,
      columns
    };

    const updated = [...currentMappings, newMapping];
    setCurrentMappings(updated);
    setActiveSheetId(newMapping.id);
    onSaveMappings(updated);
    setSaveMessage(`Added mapping for worksheet '${sheetName}'`);
    setTimeout(() => setSaveMessage(null), 3000);
  };

  const handleUpdateColumn = (colId: string, updates: Partial<ColumnMapping>) => {
    if (!activeMapping) return;
    const updatedCols = activeMapping.columns.map(c => c.id === colId ? { ...c, ...updates } : c);
    handleUpdateActiveMapping({ columns: updatedCols });
  };

  const handleAddColumn = () => {
    if (!activeMapping) return;
    const newCol: ColumnMapping = {
      id: `cm-${Date.now()}`,
      excelColumn: 'A',
      excelHeader: 'New Header',
      supabaseColumn: 'new_column',
      dataType: 'text',
      required: false,
      uniqueKey: false,
      transformation: 'trim'
    };
    handleUpdateActiveMapping({ columns: [...activeMapping.columns, newCol] });
  };

  const handleDeleteColumn = (colId: string) => {
    if (!activeMapping) return;
    handleUpdateActiveMapping({ columns: activeMapping.columns.filter(c => c.id !== colId) });
  };

  const handleSave = () => {
    onSaveMappings(currentMappings);
    setSaveMessage('All worksheet mappings saved to database configuration.');
    setTimeout(() => setSaveMessage(null), 3000);
  };

  const handleResetDefaults = () => {
    const defaults = getDefaultSampleMappings(activeMapping?.workbookName || 'students_complex.xlsx');
    setCurrentMappings(defaults);
    onSaveMappings(defaults);
    setSaveMessage('Reset to verified default sample mappings.');
    setTimeout(() => setSaveMessage(null), 3000);
  };

  // Find unmapped sheets from the loaded analysis
  const unmappedSheets = currentAnalysis?.worksheets.filter(
    (ws: SheetAnalysis) => !currentMappings.some(m => m.worksheetName.toLowerCase() === ws.sheetName.toLowerCase())
  ) || [];

  const filteredMappings = currentMappings.filter(m => 
    m.worksheetName.toLowerCase().includes(sheetSearch.toLowerCase()) ||
    m.supabaseTable.toLowerCase().includes(sheetSearch.toLowerCase())
  );

  if (!activeMapping) {
    return (
      <div className="p-8 text-center bg-white rounded-xl border border-slate-200 shadow-2xs space-y-3">
        <div className="text-slate-500">No active mappings defined.</div>
        {currentAnalysis && currentAnalysis.worksheets.length > 0 && (
          <button
            onClick={handleAutoGenerateAllSheets}
            className="px-4 py-2 bg-emerald-600 text-white rounded-lg text-xs font-semibold hover:bg-emerald-700"
          >
            Auto-Generate Mappings for All Sheets in {currentAnalysis.filename}
          </button>
        )}
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-emerald-50 text-emerald-600">
            <GitFork className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Worksheet to PostgreSQL Mappings</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Deterministic transformation & validation rules for automatic background synchronization.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-2">
          <button
            onClick={handleResetDefaults}
            className="inline-flex items-center space-x-1 px-3 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-600 hover:bg-slate-50 shadow-2xs"
            title="Reset to default students.xlsx mappings"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Reset Defaults</span>
          </button>

          <button
            onClick={onOpenAiAssistant}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg bg-purple-50 border border-purple-200 text-xs font-medium text-purple-700 hover:bg-purple-100 shadow-2xs"
          >
            <Sparkles className="w-3.5 h-3.5 text-purple-600" />
            <span>AI Mapping Suggester</span>
          </button>

          <button
            id="btn-save-mappings"
            onClick={handleSave}
            className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-medium text-white shadow-2xs"
          >
            <Save className="w-3.5 h-3.5" />
            <span>Save Mappings</span>
          </button>
        </div>
      </div>

      {/* Multi-sheet Workbook Management Toolbar */}
      <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center space-x-3">
          <div className="p-2 rounded-lg bg-indigo-50 text-indigo-600">
            <BookOpen className="w-5 h-5" />
          </div>
          <div>
            <div className="text-sm font-semibold text-slate-900 flex items-center space-x-2">
              <span>Multi-Sheet Workbook Controller</span>
              <span className="text-xs px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 font-mono">
                {currentMappings.length} sheets mapped
              </span>
            </div>
            <p className="text-xs text-slate-500">
              {currentAnalysis 
                ? `Workbook '${currentAnalysis.filename}' has ${currentAnalysis.totalWorksheets} total sheets (${unmappedSheets.length} unmapped).`
                : 'Configure mappings for each worksheet tab or skip unneeded sheets.'}
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-2 flex-wrap gap-y-2">
          {currentAnalysis && unmappedSheets.length > 0 && (
            <>
              <button
                onClick={handleAutoGenerateAllSheets}
                className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-300 text-xs font-medium text-emerald-700 hover:bg-emerald-100 shadow-2xs"
                title="Auto-generate PostgreSQL table mappings for all remaining unmapped sheets"
              >
                <Plus className="w-3.5 h-3.5 text-emerald-600" />
                <span>Map All {unmappedSheets.length} Remaining Sheets</span>
              </button>

              <select
                onChange={(e) => {
                  if (e.target.value) {
                    handleAddSheetFromAnalysis(e.target.value);
                    e.target.value = '';
                  }
                }}
                defaultValue=""
                className="px-2.5 py-1.5 rounded-lg border border-slate-300 text-xs bg-white text-slate-700 font-medium focus:outline-none focus:ring-1 focus:ring-emerald-500"
              >
                <option value="" disabled>+ Add Specific Sheet...</option>
                {unmappedSheets.map((ws: SheetAnalysis) => (
                  <option key={ws.sheetName} value={ws.sheetName}>
                    {ws.sheetName} ({ws.totalRows} rows)
                  </option>
                ))}
              </select>
            </>
          )}

          {currentMappings.length > 4 && (
            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
              <input
                type="text"
                placeholder="Filter sheets..."
                value={sheetSearch}
                onChange={(e) => setSheetSearch(e.target.value)}
                className="pl-8 pr-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-1 focus:ring-indigo-500 w-36"
              />
            </div>
          )}
        </div>
      </div>

      {saveMessage && (
        <div className="p-3.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600" />
          <span>{saveMessage}</span>
        </div>
      )}

      {/* Worksheet Switcher Tabs */}
      <div className="flex border-b border-slate-200 space-x-2 pb-0.5 overflow-x-auto">
        {filteredMappings.map((m) => {
          const policy = m.syncPolicy || 'BIDIRECTIONAL';
          return (
            <button
              key={m.id}
              id={`mapping-tab-${m.id}`}
              onClick={() => setActiveSheetId(m.id)}
              className={`px-4 py-2.5 rounded-t-lg text-sm font-medium whitespace-nowrap transition-all border-b-2 flex items-center space-x-2 ${
                activeSheetId === m.id
                  ? 'border-emerald-600 text-emerald-700 bg-emerald-50/50 font-semibold'
                  : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              } ${!m.enabled ? 'opacity-60 line-through' : ''}`}
            >
              <span>{m.worksheetName}</span>
              <span className="text-xs px-1.5 py-0.2 rounded-full bg-slate-200 text-slate-600 font-mono">
                → {m.supabaseTable}
              </span>
              <span className={`text-[10px] px-1.5 py-0.5 rounded font-bold uppercase tracking-wider ${
                policy === 'EXCEL_TO_DB'
                  ? 'bg-emerald-100 text-emerald-800'
                  : policy === 'DB_TO_EXCEL'
                  ? 'bg-blue-100 text-blue-800'
                  : policy === 'BIDIRECTIONAL'
                  ? 'bg-purple-100 text-purple-800'
                  : 'bg-slate-200 text-slate-700'
              }`}>
                {policy === 'EXCEL_TO_DB' 
                  ? 'Excel Master' 
                  : policy === 'DB_TO_EXCEL' 
                  ? 'DB Master' 
                  : policy === 'BIDIRECTIONAL' 
                  ? '2-Way' 
                  : 'Read-Only'}
              </span>
              {!m.enabled && (
                <span className="text-[10px] px-1 rounded bg-amber-100 text-amber-800 uppercase font-bold">
                  Ignored
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* Worksheet Configuration Settings */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div className="flex items-center space-x-2">
            <Table className="w-4 h-4 text-emerald-600" />
            <span className="text-sm font-semibold text-slate-900">Worksheet Configuration: {activeMapping.worksheetName}</span>
          </div>

          <div className="flex items-center space-x-3">
            {/* Sync Toggle */}
            <button
              onClick={handleToggleActiveMappingEnabled}
              className={`inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border transition-colors ${
                activeMapping.enabled
                  ? 'bg-emerald-50 border-emerald-300 text-emerald-700 hover:bg-emerald-100'
                  : 'bg-slate-100 border-slate-300 text-slate-600 hover:bg-slate-200'
              }`}
              title={activeMapping.enabled ? 'Click to disable syncing this worksheet' : 'Click to enable syncing this worksheet'}
            >
              {activeMapping.enabled ? (
                <>
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Sync: Active</span>
                </>
              ) : (
                <>
                  <EyeOff className="w-3.5 h-3.5 text-slate-500" />
                  <span>Sync: Disabled (Ignored)</span>
                </>
              )}
            </button>

            {currentMappings.length > 1 && (
              <button
                onClick={handleDeleteActiveMapping}
                className="inline-flex items-center space-x-1 px-2.5 py-1.5 rounded-lg text-xs font-medium text-rose-600 border border-rose-200 hover:bg-rose-50"
                title="Remove this sheet from mapping configuration"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Remove Sheet</span>
              </button>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
          <div>
            <div className="flex items-center justify-between mb-1">
              <label className="block text-slate-700 font-bold uppercase tracking-wider">Target Supabase Table</label>
              <span className="text-[10px] text-emerald-600 font-medium">PostgreSQL</span>
            </div>
            <input
              type="text"
              value={activeMapping.supabaseTable}
              onChange={(e) => handleUpdateActiveMapping({ supabaseTable: e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, '_') })}
              placeholder="e.g. students"
              className="w-full px-3 py-1.5 rounded-lg border border-slate-300 font-mono text-slate-800 font-semibold focus:ring-2 focus:ring-emerald-500 focus:outline-none"
            />
            <div className="flex items-center space-x-1.5 mt-1.5">
              <span className="text-[10px] text-slate-400">Route to:</span>
              {['students', 'attendance', 'grades', 'users'].map((tbl) => (
                <button
                  key={tbl}
                  type="button"
                  onClick={() => handleUpdateActiveMapping({ supabaseTable: tbl })}
                  className={`text-[10px] px-1.5 py-0.5 rounded border transition-colors ${
                    activeMapping.supabaseTable === tbl
                      ? 'bg-emerald-100 text-emerald-800 border-emerald-300 font-semibold'
                      : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  {tbl}
                </button>
              ))}
            </div>
          </div>

          <div>
            <label className="block text-slate-600 mb-1 font-semibold uppercase tracking-wider">Header Row</label>
            <input
              type="number"
              min="1"
              value={activeMapping.headerRow}
              onChange={(e) => handleUpdateActiveMapping({ headerRow: Number(e.target.value) })}
              className="w-full px-3 py-1.5 rounded-lg border border-slate-300 font-mono focus:ring-1 focus:ring-emerald-500 focus:outline-none"
            />
            <p className="text-[10px] text-slate-400 mt-1">Excel row containing column headers</p>
          </div>

          <div>
            <label className="block text-slate-600 mb-1 font-semibold uppercase tracking-wider">Data Start Row</label>
            <input
              type="number"
              min="1"
              value={activeMapping.dataStartRow}
              onChange={(e) => handleUpdateActiveMapping({ dataStartRow: Number(e.target.value) })}
              className="w-full px-3 py-1.5 rounded-lg border border-slate-300 font-mono focus:ring-1 focus:ring-emerald-500 focus:outline-none"
            />
            <p className="text-[10px] text-slate-400 mt-1">First row containing student/record data</p>
          </div>

          <div>
            <label className="block text-slate-600 mb-1 font-semibold uppercase tracking-wider">
              Section Heading Column (Optional)
            </label>
            <input
              type="text"
              placeholder="e.g. class"
              value={activeMapping.sectionHeadingTargetCol || ''}
              onChange={(e) => handleUpdateActiveMapping({ sectionHeadingTargetCol: e.target.value || undefined })}
              className="w-full px-3 py-1.5 rounded-lg border border-slate-300 font-mono focus:ring-1 focus:ring-emerald-500 focus:outline-none"
            />
            <p className="text-[10px] text-slate-400 mt-1">Target column for merged banner labels</p>
          </div>
        </div>

        {/* Quick Schema & Column Auto-Mapper Tools */}
        <div className="p-3 bg-slate-50 rounded-lg border border-slate-200 flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center space-x-2 text-xs text-slate-700">
            <span className="font-semibold text-slate-900">Auto-Mapping Helpers:</span>
            <span className="text-slate-500">Quickly align columns with your actual Excel headers</span>
          </div>

          <div className="flex items-center space-x-2">
            {currentAnalysis && (
              <button
                type="button"
                onClick={() => {
                  const ws = currentAnalysis.worksheets.find(w => w.sheetName.toLowerCase() === activeMapping.worksheetName.toLowerCase()) || currentAnalysis.worksheets[0];
                  if (!ws || !ws.headers.length) {
                    setSaveMessage('No headers found in current analysis for this sheet.');
                    return;
                  }

                  const newCols: ColumnMapping[] = ws.headers.map((h, idx) => {
                    const cleanName = h.name.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/^_+|_+$/g, '') || `col_${h.colLetter.toLowerCase()}`;
                    const isId = idx === 0 || cleanName.includes('username') || cleanName.includes('id') || cleanName.includes('number') || cleanName.includes('index');
                    const isDate = cleanName.includes('dob') || cleanName.includes('date');
                    const isNumber = cleanName.includes('class') || cleanName.includes('score') || cleanName.includes('total') || cleanName.includes('mark');

                    let trans: TransformationType = 'trim';
                    if (isDate) trans = 'parse_date';
                    else if (isNumber) trans = 'parse_number';
                    else if (cleanName.includes('id') || cleanName.includes('index')) trans = 'normalize_id';

                    return {
                      id: `cm-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
                      excelColumn: h.colLetter,
                      excelHeader: h.name,
                      supabaseColumn: cleanName,
                      dataType: isDate ? 'date' : isNumber ? 'integer' : 'text',
                      required: isId,
                      uniqueKey: idx === 0 || cleanName === 'username' || cleanName === 'indexnumber',
                      transformation: trans,
                    };
                  });

                  handleUpdateActiveMapping({ columns: newCols });
                  setSaveMessage(`Extracted and mapped ${newCols.length} columns directly from sheet '${ws.sheetName}'!`);
                  setTimeout(() => setSaveMessage(null), 3500);
                }}
                className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-md bg-white border border-emerald-300 text-emerald-700 hover:bg-emerald-50 text-xs font-medium shadow-2xs transition-colors"
                title="Populate columns using detected headers from analyzed Excel sheet"
              >
                <Sparkles className="w-3.5 h-3.5 text-emerald-600" />
                <span>Auto-Map from Analyzed Sheet ({currentAnalysis.filename})</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => {
                const studentPresetCols: ColumnMapping[] = [
                  { id: `cm-${Date.now()}-1`, excelColumn: 'A', excelHeader: 'username', supabaseColumn: 'username', dataType: 'text', required: true, uniqueKey: true, transformation: 'normalize_id' },
                  { id: `cm-${Date.now()}-2`, excelColumn: 'B', excelHeader: 'fullName', supabaseColumn: 'full_name', dataType: 'text', required: true, uniqueKey: false, transformation: 'trim' },
                  { id: `cm-${Date.now()}-3`, excelColumn: 'C', excelHeader: 'email', supabaseColumn: 'email', dataType: 'text', required: false, uniqueKey: false, transformation: 'trim' },
                  { id: `cm-${Date.now()}-4`, excelColumn: 'D', excelHeader: 'indexNumber', supabaseColumn: 'index_number', dataType: 'text', required: false, uniqueKey: false, transformation: 'normalize_id' },
                  { id: `cm-${Date.now()}-5`, excelColumn: 'E', excelHeader: 'dob', supabaseColumn: 'dob', dataType: 'date', required: false, uniqueKey: false, transformation: 'parse_date' },
                  { id: `cm-${Date.now()}-6`, excelColumn: 'F', excelHeader: 'class', supabaseColumn: 'class', dataType: 'integer', required: false, uniqueKey: false, transformation: 'parse_number' },
                  { id: `cm-${Date.now()}-7`, excelColumn: 'G', excelHeader: 'division', supabaseColumn: 'division', dataType: 'text', required: false, uniqueKey: false, transformation: 'trim' },
                  { id: `cm-${Date.now()}-8`, excelColumn: 'H', excelHeader: 'password', supabaseColumn: 'password', dataType: 'text', required: false, uniqueKey: false, transformation: 'trim' },
                ];
                handleUpdateActiveMapping({
                  supabaseTable: 'students',
                  columns: studentPresetCols,
                });
                setSaveMessage('Applied Standard Student Roster Schema (8 columns with username as unique key).');
                setTimeout(() => setSaveMessage(null), 3500);
              }}
              className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-md bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 text-xs font-medium shadow-2xs transition-colors"
              title="Apply username, fullName, indexNumber, dob, class, division, password, email schema"
            >
              <Table className="w-3.5 h-3.5 text-blue-600" />
              <span>Apply Student Roster Preset</span>
            </button>
          </div>
        </div>

        {/* Table-Level Sync Direction & Authority Policy Selector */}
        <div className="pt-2 border-t border-slate-100">
          <div className="flex items-center justify-between mb-2">
            <label className="block text-slate-800 font-bold uppercase tracking-wider text-xs">
              Table Sync Authority & Direction Policy
            </label>
            <span className="text-[11px] text-slate-500">
              Controls whether Nextcloud or Supabase owns this table to prevent conflicting overwrites.
            </span>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            {/* EXCEL_TO_DB */}
            <button
              type="button"
              onClick={() => handleUpdateActiveMapping({ syncPolicy: 'EXCEL_TO_DB' })}
              className={`p-3 rounded-lg border text-left transition-all relative ${
                (activeMapping.syncPolicy || 'BIDIRECTIONAL') === 'EXCEL_TO_DB'
                  ? 'border-emerald-500 bg-emerald-50/70 ring-2 ring-emerald-500/20'
                  : 'border-slate-200 hover:border-slate-300 bg-white'
              }`}
            >
              <div className="flex items-center justify-between mb-1">
                <span className="text-xs font-bold text-emerald-800 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                  Nextcloud Master
                </span>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-mono font-semibold">
                  Excel → DB
                </span>
              </div>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Nextcloud Excel is the authoritative master. Edits push to Supabase; Supabase cannot overwrite Excel.
              </p>
            </button>

            {/* DB_TO_EXCEL */}
            <button
              type="button"
              onClick={() => handleUpdateActiveMapping({ syncPolicy: 'DB_TO_EXCEL' })}
              className={`p-3 rounded-lg border text-left transition-all relative ${
                activeMapping.syncPolicy === 'DB_TO_EXCEL'
                  ? 'border-blue-500 bg-blue-50/70 ring-2 ring-blue-500/20'
                  : 'border-slate-200 hover:border-slate-300 bg-white'
              }`}
            >
              <div className="flex items-center justify-between mb-1">
                <span className="text-xs font-bold text-blue-800 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-blue-500"></span>
                  Supabase Master
                </span>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-800 font-mono font-semibold">
                  DB → Excel
                </span>
              </div>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Live Supabase database is master (e.g. attendance). Auto-mirrors to Excel; Excel cannot overwrite DB.
              </p>
            </button>

            {/* BIDIRECTIONAL */}
            <button
              type="button"
              onClick={() => handleUpdateActiveMapping({ syncPolicy: 'BIDIRECTIONAL' })}
              className={`p-3 rounded-lg border text-left transition-all relative ${
                (activeMapping.syncPolicy || 'BIDIRECTIONAL') === 'BIDIRECTIONAL'
                  ? 'border-purple-500 bg-purple-50/70 ring-2 ring-purple-500/20'
                  : 'border-slate-200 hover:border-slate-300 bg-white'
              }`}
            >
              <div className="flex items-center justify-between mb-1">
                <span className="text-xs font-bold text-purple-800 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-purple-500"></span>
                  Bidirectional
                </span>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-800 font-mono font-semibold">
                  Excel ⇄ DB
                </span>
              </div>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Collaborative 2-way sync. Unilateral edits sync automatically; concurrent collisions pause for review.
              </p>
            </button>

            {/* READ_ONLY */}
            <button
              type="button"
              onClick={() => handleUpdateActiveMapping({ syncPolicy: 'READ_ONLY' })}
              className={`p-3 rounded-lg border text-left transition-all relative ${
                activeMapping.syncPolicy === 'READ_ONLY'
                  ? 'border-slate-500 bg-slate-100 ring-2 ring-slate-400/20'
                  : 'border-slate-200 hover:border-slate-300 bg-white'
              }`}
            >
              <div className="flex items-center justify-between mb-1">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-slate-400"></span>
                  Read-Only Audit
                </span>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-200 text-slate-700 font-mono font-semibold">
                  No Auto-Write
                </span>
              </div>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Inspection & audit only. Neither Nextcloud nor Supabase is automatically modified.
              </p>
            </button>
          </div>
        </div>

        {/* Merged Cell & Downward Propagation Notice */}
        <div className="p-3 bg-blue-50/70 border border-blue-200 rounded-lg text-xs text-blue-900 flex items-start space-x-2">
          <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
          <div className="space-y-0.5">
            <span className="font-semibold">Merged Cells & Section Headings Auto-Handling:</span>
            <p className="text-blue-800">
              Vertically and horizontally merged cells are automatically forward-propagated so every row inherits the correct value. Wide banner merges (e.g. <code>A3:H3 &quot;CLASS 10A&quot;</code>) are captured and saved to the column specified above.
            </p>
          </div>
        </div>
      </div>

      {/* Column Mappings Table (Section 8 & 9) */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
        <div className="p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
          <div className="flex items-center space-x-2">
            <span className="text-xs font-semibold text-slate-900 uppercase tracking-wider">
              Column Transformation & Validation Rules ({activeMapping.columns.length})
            </span>
            <span className="text-[11px] text-slate-500">
              Unique keys define idempotent upsert conflicts.
            </span>
          </div>
          <button
            id="btn-add-column-mapping"
            onClick={handleAddColumn}
            className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-medium shadow-2xs"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Add Column</span>
          </button>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200">
              <tr>
                <th className="px-3 py-3 w-16">Col</th>
                <th className="px-4 py-3">Excel Header</th>
                <th className="px-4 py-3">Supabase Column</th>
                <th className="px-3 py-3 w-28">Type</th>
                <th className="px-4 py-3">Transformation</th>
                <th className="px-3 py-3 text-center">Req?</th>
                <th className="px-3 py-3 text-center">Key?</th>
                <th className="px-3 py-3 text-center w-12">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {activeMapping.columns.map((col) => (
                <tr key={col.id} className="hover:bg-slate-50 transition-colors">
                  {/* Excel Column Letter */}
                  <td className="px-3 py-2.5">
                    <input
                      type="text"
                      value={col.excelColumn}
                      onChange={(e) => handleUpdateColumn(col.id, { excelColumn: e.target.value.toUpperCase() })}
                      className="w-12 px-2 py-1 rounded border border-slate-300 font-mono font-bold text-center focus:outline-none focus:ring-1 focus:ring-emerald-500"
                    />
                  </td>

                  {/* Excel Header */}
                  <td className="px-4 py-2.5">
                    <input
                      type="text"
                      value={col.excelHeader}
                      onChange={(e) => handleUpdateColumn(col.id, { excelHeader: e.target.value })}
                      className="w-full px-2 py-1 rounded border border-slate-300 focus:outline-none focus:ring-1 focus:ring-emerald-500 font-medium text-slate-900"
                    />
                  </td>

                  {/* Supabase Column */}
                  <td className="px-4 py-2.5">
                    <input
                      type="text"
                      value={col.supabaseColumn}
                      onChange={(e) => handleUpdateColumn(col.id, { supabaseColumn: e.target.value })}
                      className="w-full px-2 py-1 rounded border border-slate-300 font-mono text-teal-800 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                    />
                  </td>

                  {/* Data Type */}
                  <td className="px-3 py-2.5">
                    <select
                      value={col.dataType}
                      onChange={(e) => handleUpdateColumn(col.id, { dataType: e.target.value as DataType })}
                      className="w-full px-2 py-1 rounded border border-slate-300 font-mono text-slate-700 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                    >
                      {dataTypeOptions.map(t => (
                        <option key={t} value={t}>{t}</option>
                      ))}
                    </select>
                  </td>

                  {/* Transformation */}
                  <td className="px-4 py-2.5">
                    <select
                      value={col.transformation}
                      onChange={(e) => handleUpdateColumn(col.id, { transformation: e.target.value as TransformationType })}
                      className="w-full px-2 py-1 rounded border border-slate-300 text-slate-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 font-medium"
                    >
                      {transformationOptions.map(opt => (
                        <option key={opt.value} value={opt.value}>{opt.label}</option>
                      ))}
                    </select>
                  </td>

                  {/* Required Checkbox */}
                  <td className="px-3 py-2.5 text-center">
                    <input
                      type="checkbox"
                      checked={col.required}
                      onChange={(e) => handleUpdateColumn(col.id, { required: e.target.checked })}
                      className="w-4 h-4 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500"
                    />
                  </td>

                  {/* Unique Key Checkbox */}
                  <td className="px-3 py-2.5 text-center">
                    <input
                      type="checkbox"
                      checked={col.uniqueKey}
                      onChange={(e) => handleUpdateColumn(col.id, { uniqueKey: e.target.checked })}
                      className="w-4 h-4 text-purple-600 rounded border-slate-300 focus:ring-purple-500"
                    />
                  </td>

                  {/* Delete Column Button */}
                  <td className="px-3 py-2.5 text-center">
                    <button
                      onClick={() => handleDeleteColumn(col.id)}
                      className="p-1 text-slate-400 hover:text-rose-600 rounded transition-colors"
                      title="Remove mapping"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Footer Navigation Action */}
      <div className="flex items-center justify-between p-4 bg-white rounded-xl border border-slate-200">
        <span className="text-xs text-slate-500">
          Ready to verify these mappings with your real workbook data without writing to the database?
        </span>
        <button
          onClick={() => onNavigate('import')}
          className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-xs"
        >
          <Play className="w-3.5 h-3.5" />
          <span>Launch Dry Run Simulator</span>
        </button>
      </div>
    </div>
  );
};
