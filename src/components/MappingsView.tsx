import React, { useState, useEffect, useCallback } from 'react';
import { 
  GitFork, 
  Plus, 
  Trash2, 
  Save, 
  RotateCcw, 
  Check, 
  Play, 
  Sparkles,
  Layers,
  Table,
  CheckCircle2,
  BookOpen,
  Search,
  EyeOff,
  Eye,
  Info,
  Database,
  Key,
  ArrowRight,
  ExternalLink,
  RefreshCw,
  HelpCircle,
  FileSpreadsheet,
  AlertCircle
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
  TableSyncPolicy,
  SupabaseConfig,
  SupabaseTableInfo,
  SupabaseTableColumn,
  AiWorkbookPreset
} from '../types';
import { getDefaultSampleMappings } from '../services/sampleWorkbook';
import { ApiClient } from '../services/apiClient';
import { AiPresetsModal } from './AiPresetsModal';
import { smartSanitizeIdentifier } from '../services/tamilTranslator';

interface MappingsViewProps {
  mappings: WorksheetMapping[];
  onSaveMappings: (mappings: WorksheetMapping[]) => void;
  onNavigate: (tab: NavigationTab) => void;
  onOpenAiAssistant: () => void;
  currentAnalysis?: WorkbookAnalysis | null;
  supabaseConfig?: SupabaseConfig;
}

export const MappingsView: React.FC<MappingsViewProps> = ({
  mappings,
  onSaveMappings,
  onNavigate,
  onOpenAiAssistant,
  currentAnalysis,
  supabaseConfig
}) => {
  const [activeSheetId, setActiveSheetId] = useState<string>(mappings[0]?.id || '');
  const [currentMappings, setCurrentMappings] = useState<WorksheetMapping[]>(mappings);
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const [sheetSearch, setSheetSearch] = useState<string>('');

  // Live Supabase tables state
  const [supabaseTables, setSupabaseTables] = useState<SupabaseTableInfo[]>([]);
  const [isLoadingSchema, setIsLoadingSchema] = useState<boolean>(false);
  const [schemaError, setSchemaError] = useState<string | null>(null);
  const [showSupabasePanel, setShowSupabasePanel] = useState<boolean>(true);

  // AI Presets & Archetypes Modal state
  const [showAiPresetsModal, setShowAiPresetsModal] = useState<boolean>(false);
  const [presetPrefill, setPresetPrefill] = useState<boolean>(false);

  // Active mapping reference
  const activeMapping = currentMappings.find(m => m.id === activeSheetId) || currentMappings[0];

  // Fetch Supabase schema
  const fetchSupabaseSchema = useCallback(async () => {
    if (!supabaseConfig?.url) return;
    setIsLoadingSchema(true);
    setSchemaError(null);
    try {
      const res = await ApiClient.getSupabaseSchema(supabaseConfig);
      if (res.success && res.tables) {
        setSupabaseTables(res.tables);
      } else {
        setSchemaError(res.error || 'Could not load tables from Supabase');
      }
    } catch (e: any) {
      setSchemaError(e.message);
    } finally {
      setIsLoadingSchema(false);
    }
  }, [supabaseConfig]);

  useEffect(() => {
    fetchSupabaseSchema();
  }, [fetchSupabaseSchema]);

  // Keep currentMappings in sync when mappings prop updates
  useEffect(() => {
    setCurrentMappings(mappings);
    if (mappings.length > 0 && (!activeSheetId || !mappings.some(m => m.id === activeSheetId))) {
      setActiveSheetId(mappings[0].id);
    }
  }, [mappings]);

  // Keep activeSheetId valid if currentMappings changes
  useEffect(() => {
    if (currentMappings.length > 0 && !currentMappings.some(m => m.id === activeSheetId)) {
      setActiveSheetId(currentMappings[0].id);
    }
  }, [currentMappings, activeSheetId]);

  // Find active worksheet analysis if loaded
  const activeSheetAnalysis = currentAnalysis?.worksheets.find(
    (w) => w.sheetName.toLowerCase() === activeMapping?.worksheetName?.toLowerCase()
  ) || currentAnalysis?.worksheets[0];

  // Find active target Supabase table schema
  const activeSupabaseTable = supabaseTables.find(
    (t) => t.name.toLowerCase() === activeMapping?.supabaseTable?.toLowerCase()
  );

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
    onSaveMappings(updated);
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

  // Clear unwanted/sample mappings and reset cleanly
  const handleClearStaleMappings = () => {
    if (!currentAnalysis || currentAnalysis.worksheets.length === 0) {
      setCurrentMappings([]);
      onSaveMappings([]);
      setSaveMessage('Cleared all mappings. You can now create fresh mappings.');
      setTimeout(() => setSaveMessage(null), 3000);
      return;
    }

    // Build fresh mappings ONLY for the active workbook sheets
    const newMappings: WorksheetMapping[] = currentAnalysis.worksheets.map((ws, sIdx) => {
      const cleanWsName = smartSanitizeIdentifier(ws.sheetName, 'sheet_data');
      const targetTable = supabaseTables.some(t => t.name === cleanWsName) 
        ? cleanWsName 
        : (supabaseTables[0]?.name || cleanWsName);
      const matchedTable = supabaseTables.find(t => t.name === targetTable);

      const usedSupabaseCols = new Set<string>();
      const columns: ColumnMapping[] = ws.headers.map((h, idx) => {
        const rawName = h.name.trim();
        let normName = smartSanitizeIdentifier(rawName, `col_${idx + 1}`);

        // Try to match with a column in the Supabase table
        let matchedCol = normName;
        if (matchedTable) {
          const directMatch = matchedTable.columns.find(
            c => c.name.toLowerCase() === normName || c.name.toLowerCase() === rawName.toLowerCase()
          );
          if (directMatch) matchedCol = directMatch.name;
          else if (normName.includes('fullname')) matchedCol = 'full_name';
          else if (normName.includes('indexnumber')) matchedCol = 'index_number';
        }

        // Deduplicate column name in case duplicate headers or normalization collision
        if (usedSupabaseCols.has(matchedCol.toLowerCase())) {
          let suffix = 2;
          while (usedSupabaseCols.has(`${matchedCol}_${suffix}`.toLowerCase())) {
            suffix++;
          }
          matchedCol = `${matchedCol}_${suffix}`;
        }
        usedSupabaseCols.add(matchedCol.toLowerCase());

        const isId = idx === 0 || matchedCol === 'username' || matchedCol === 'id' || matchedCol.includes('code');
        const isDate = h.inferredType === 'date' || matchedCol.includes('dob') || matchedCol.includes('date');
        const isNumber = h.inferredType === 'integer' || h.inferredType === 'decimal' || matchedCol.includes('score') || matchedCol.includes('count') || matchedCol.includes('mark') || matchedCol.includes('total');

        let trans: TransformationType = 'trim';
        if (isDate) trans = 'parse_date';
        else if (isNumber) trans = 'parse_number';
        else if (isId) trans = 'normalize_id';

        return {
          id: `cm-${Date.now()}-${sIdx}-${idx}`,
          excelColumn: h.colLetter,
          excelHeader: h.name,
          supabaseColumn: matchedCol,
          dataType: isDate ? 'date' : isNumber ? (h.inferredType === 'decimal' ? 'decimal' : 'integer') : (h.inferredType || 'text'),
          required: isId,
          uniqueKey: isId && idx === 0,
          transformation: trans
        };
      });

      return {
        id: `wm-${Date.now()}-${sIdx}`,
        workbookName: currentAnalysis.filename,
        worksheetName: ws.sheetName,
        supabaseTable: targetTable,
        headerRow: ws.detectedHeaderRow || 1,
        dataStartRow: ws.detectedDataStartRow || 2,
        enabled: true,
        columns,
        syncPolicy: 'EXCEL_TO_DB'
      };
    });

    setCurrentMappings(newMappings);
    if (newMappings.length > 0) {
      setActiveSheetId(newMappings[0].id);
    }
    onSaveMappings(newMappings);
    setSaveMessage(`Reset mappings to match only the ${newMappings.length} sheets in '${currentAnalysis.filename}'.`);
    setTimeout(() => setSaveMessage(null), 3500);
  };

  // Route active sheet to a selected Supabase table and align columns
  const handleSelectSupabaseTable = (targetTableName: string) => {
    if (!activeMapping) return;

    const tableSchema = supabaseTables.find(t => t.name.toLowerCase() === targetTableName.toLowerCase());
    const tableCols = tableSchema?.columns || [];

    // Intelligently auto-map existing sheet columns to target table columns
    const updatedCols: ColumnMapping[] = activeMapping.columns.map((col) => {
      const headerNorm = col.excelHeader.toLowerCase().replace(/[^a-z0-9]/g, '');
      const colNorm = col.excelColumn.toLowerCase();

      // Look for match in tableCols
      const matched = tableCols.find(tc => {
        const tcNorm = tc.name.toLowerCase().replace(/[^a-z0-9]/g, '');
        return tcNorm === headerNorm || 
               (headerNorm.includes('fullname') && tcNorm === 'fullname') ||
               (headerNorm.includes('index') && tcNorm.includes('index')) ||
               (headerNorm.includes('dob') && tcNorm.includes('dob')) ||
               (headerNorm === 'username' && tcNorm === 'username') ||
               (headerNorm === 'class' && tcNorm === 'class') ||
               (headerNorm === 'email' && tcNorm === 'email');
      });

      if (matched) {
        let dt: DataType = 'text';
        let trans: TransformationType = 'trim';
        if (matched.type === 'integer') { dt = 'integer'; trans = 'parse_number'; }
        else if (matched.type === 'decimal') { dt = 'decimal'; trans = 'parse_number'; }
        else if (matched.type === 'date') { dt = 'date'; trans = 'parse_date'; }
        else if (matched.type === 'boolean') { dt = 'boolean'; trans = 'yes_no_to_boolean'; }
        else if (matched.type === 'timestamp') { dt = 'timestamp'; trans = 'parse_date'; }

        return {
          ...col,
          supabaseColumn: matched.name,
          dataType: dt,
          transformation: trans,
          uniqueKey: col.uniqueKey,
          required: matched.required || col.required,
        };
      }

      return col;
    });

    handleUpdateActiveMapping({
      supabaseTable: targetTableName,
      columns: updatedCols,
    });

    setSaveMessage(`Active worksheet '${activeMapping.worksheetName}' is now routed to Supabase table '${targetTableName}'.`);
    setTimeout(() => setSaveMessage(null), 3500);
  };

  // Auto-match columns by name
  const handleAutoMatchColumns = () => {
    if (!activeMapping || !activeSupabaseTable) return;
    const tableCols = activeSupabaseTable.columns;

    const updatedCols: ColumnMapping[] = activeMapping.columns.map(col => {
      const hClean = col.excelHeader.toLowerCase().replace(/[^a-z0-9]/g, '');

      const match = tableCols.find(tc => {
        const tcClean = tc.name.toLowerCase().replace(/[^a-z0-9]/g, '');
        return tcClean === hClean || 
               (hClean.includes('fullname') && tcClean === 'fullname') ||
               (hClean.includes('index') && tcClean.includes('index')) ||
               (hClean.includes('dob') && tcClean.includes('dob')) ||
               (hClean === 'user' && tcClean === 'username') ||
               (hClean === 'pass' && tcClean === 'password');
      });

      if (match) {
        let dt: DataType = 'text';
        let trans: TransformationType = 'trim';
        if (match.type === 'integer') { dt = 'integer'; trans = 'parse_number'; }
        else if (match.type === 'decimal') { dt = 'decimal'; trans = 'parse_number'; }
        else if (match.type === 'date') { dt = 'date'; trans = 'parse_date'; }
        else if (match.type === 'boolean') { dt = 'boolean'; trans = 'yes_no_to_boolean'; }
        else if (match.type === 'timestamp') { dt = 'timestamp'; trans = 'parse_date'; }

        return {
          ...col,
          supabaseColumn: match.name,
          dataType: dt,
          transformation: trans,
          uniqueKey: col.uniqueKey,
          required: match.required || col.required,
        };
      }
      return col;
    });

    handleUpdateActiveMapping({ columns: updatedCols });
    setSaveMessage(`Auto-matched columns against Supabase table '${activeSupabaseTable.name}'!`);
    setTimeout(() => setSaveMessage(null), 3000);
  };

  const handleConsolidateAllSheets = (targetTbl: string) => {
    const cleanTbl = targetTbl.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'students';
    const updated = currentMappings.map(m => ({
      ...m,
      supabaseTable: cleanTbl
    }));
    setCurrentMappings(updated);
    onSaveMappings(updated);
    setSaveMessage(`Consolidated all ${updated.length} worksheets into Supabase table 'public.${cleanTbl}'!`);
    setTimeout(() => setSaveMessage(null), 3500);
  };

  const handleSplitAllSheetsSeparate = () => {
    const updated = currentMappings.map(m => ({
      ...m,
      supabaseTable: m.worksheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_')
    }));
    setCurrentMappings(updated);
    onSaveMappings(updated);
    setSaveMessage(`Configured separate 1:1 tables for each of the ${updated.length} worksheets.`);
    setTimeout(() => setSaveMessage(null), 3500);
  };

  // Toggle a column as the primary merge conflict key (single active key policy)
  const handleToggleMergeKey = (colId: string) => {
    if (!activeMapping) return;
    const targetCol = activeMapping.columns.find(c => c.id === colId);
    if (!targetCol) return;

    const newUnique = !targetCol.uniqueKey;
    const updatedCols = activeMapping.columns.map(c => {
      if (c.id === colId) {
        return { ...c, uniqueKey: newUnique, required: newUnique ? true : c.required };
      }
      // If activating this key, unset all other columns to avoid composite key conflicts in Supabase
      if (newUnique) {
        return { ...c, uniqueKey: false };
      }
      return c;
    });

    handleUpdateActiveMapping({ columns: updatedCols });
    setSaveMessage(newUnique 
      ? `Column '${targetCol.supabaseColumn}' set as PRIMARY MERGE KEY (ON CONFLICT DO UPDATE).`
      : `Merge key unset. Worksheets will import via direct INSERT.`
    );
    setTimeout(() => setSaveMessage(null), 3500);
  };

  // Explicitly select primary merge key from dropdown or clear it (Insert Only)
  const handleSetPrimaryMergeKey = (targetColName: string) => {
    if (!activeMapping) return;
    const updatedCols = activeMapping.columns.map(c => {
      const isMatch = c.supabaseColumn === targetColName && targetColName !== '';
      return {
        ...c,
        uniqueKey: isMatch,
        required: isMatch ? true : c.required
      };
    });

    handleUpdateActiveMapping({ columns: updatedCols });
    setSaveMessage(targetColName 
      ? `Column '${targetColName}' set as PRIMARY MERGE KEY (ON CONFLICT DO UPDATE).`
      : `Merge key cleared. All rows will import via direct INSERT.`
    );
    setTimeout(() => setSaveMessage(null), 3500);
  };

  const handleUpdateColumn = (colId: string, updates: Partial<ColumnMapping>) => {
    if (!activeMapping) return;
    const updatedCols = activeMapping.columns.map(c => c.id === colId ? { ...c, ...updates } : c);
    handleUpdateActiveMapping({ columns: updatedCols });
  };

  const handleAddColumn = () => {
    if (!activeMapping) return;
    const nextColLetter = String.fromCharCode(65 + activeMapping.columns.length);
    const newCol: ColumnMapping = {
      id: `cm-${Date.now()}`,
      excelColumn: nextColLetter,
      excelHeader: `Column ${nextColLetter}`,
      supabaseColumn: `col_${nextColLetter.toLowerCase()}`,
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

  const handleSave = async () => {
    onSaveMappings(currentMappings);
    try {
      await ApiClient.savePermanentMappings({
        mappings: currentMappings,
        workbookInfo: {
          filename: currentAnalysis?.filename || activeMapping?.workbookName || 'students.xlsx',
          fileHash: currentAnalysis?.fileHash,
          totalWorksheets: currentMappings.length,
        },
        supabase: supabaseConfig,
      });
      setSaveMessage('✨ Mappings permanently saved to server disk and Supabase PostgreSQL metadata tables!');
    } catch {
      setSaveMessage('All worksheet mappings saved to database configuration.');
    }
    setTimeout(() => setSaveMessage(null), 4000);
  };

  const handleResetDefaults = () => {
    const defaults = getDefaultSampleMappings(activeMapping?.workbookName || 'students_complex.xlsx');
    setCurrentMappings(defaults);
    onSaveMappings(defaults);
    setSaveMessage('Reset to default sample mappings.');
    setTimeout(() => setSaveMessage(null), 3000);
  };

  const handleApplyPreset = (preset: AiWorkbookPreset) => {
    if (preset.sheetMappings && preset.sheetMappings.length > 0) {
      setCurrentMappings(preset.sheetMappings);
      setActiveSheetId(preset.sheetMappings[0].id);
      onSaveMappings(preset.sheetMappings);
      setSaveMessage(`✨ Successfully applied AI Preset '${preset.name}'! (${preset.sheetMappings.length} worksheets configured, merge keys preserved).`);
      setTimeout(() => setSaveMessage(null), 4500);
    }
  };

  // Find sample values for a column from the analyzed sheet
  const getSampleValuesForCol = (colLetter: string, colHeaderName: string): string => {
    if (!activeSheetAnalysis) return '';
    const headerInfo = activeSheetAnalysis.headers.find(
      h => h.colLetter.toUpperCase() === colLetter.toUpperCase() || h.name.toLowerCase() === colHeaderName.toLowerCase()
    );
    if (headerInfo && headerInfo.sampleValues.length > 0) {
      return headerInfo.sampleValues.slice(0, 3).join(', ');
    }
    // Check in sampleRows
    if (activeSheetAnalysis.sampleRows && activeSheetAnalysis.sampleRows.length > 0) {
      const vals = activeSheetAnalysis.sampleRows
        .map(sr => sr.data[colLetter] ?? sr.data[colHeaderName])
        .filter(v => v !== undefined && v !== null && String(v).trim() !== '')
        .slice(0, 3);
      if (vals.length > 0) return vals.join(', ');
    }
    return '';
  };

  // Find unique merge keys for current mapping
  const mergeKeys = activeMapping?.columns.filter(c => c.uniqueKey).map(c => c.supabaseColumn) || [];

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-emerald-50 text-emerald-600">
            <GitFork className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Excel to Supabase Table Mappings</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Select which Excel sheet maps to which Supabase table, map columns, and define merge keys for upsert synchronization.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-2 flex-wrap gap-y-2">
          {currentAnalysis && (
            <button
              onClick={handleClearStaleMappings}
              className="inline-flex items-center space-x-1 px-3 py-2 rounded-lg border border-amber-300 bg-amber-50 text-xs font-semibold text-amber-800 hover:bg-amber-100 shadow-2xs"
              title="Clear unwanted sample mappings and map only the sheets from your uploaded file"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Map Only Current Excel</span>
            </button>
          )}

          <button
            onClick={fetchSupabaseSchema}
            disabled={isLoadingSchema}
            className="inline-flex items-center space-x-1 px-3 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 hover:bg-slate-50 shadow-2xs"
            title="Refresh tables and columns from Supabase"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isLoadingSchema ? 'animate-spin' : ''}`} />
            <span>Refresh Supabase Tables</span>
          </button>

          <button
            id="btn-open-presets-library"
            onClick={() => {
              setPresetPrefill(false);
              setShowAiPresetsModal(true);
            }}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg bg-indigo-50 border border-indigo-200 text-xs font-semibold text-indigo-800 hover:bg-indigo-100 shadow-2xs transition-colors"
            title="Browse pre-built workbook archetype presets (Timetable, Donations, Allocations)"
          >
            <BookOpen className="w-3.5 h-3.5 text-indigo-600" />
            <span>AI Presets Library</span>
          </button>

          <button
            id="btn-view-served-sheets-hub"
            onClick={() => onNavigate('served_mappings')}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-xs font-semibold text-white shadow-2xs transition-colors"
            title="Inspect unified overview of all served sheets, merged Supabase tables, and single sync history"
          >
            <Layers className="w-3.5 h-3.5" />
            <span>Served Sheets Hub</span>
          </button>

          <button
            id="btn-save-as-ai-preset"
            onClick={() => {
              setPresetPrefill(true);
              setShowAiPresetsModal(true);
            }}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg bg-purple-50 border border-purple-200 text-xs font-semibold text-purple-800 hover:bg-purple-100 shadow-2xs transition-colors"
            title="Save these current worksheet mappings as a reusable AI Preset in server disk & Supabase"
          >
            <Sparkles className="w-3.5 h-3.5 text-purple-600" />
            <span>Save as AI Preset</span>
          </button>

          <button
            onClick={onOpenAiAssistant}
            className="inline-flex items-center space-x-1.5 px-3 py-2 rounded-lg bg-slate-50 border border-slate-300 text-xs font-medium text-slate-700 hover:bg-slate-100 shadow-2xs"
          >
            <Sparkles className="w-3.5 h-3.5 text-slate-600" />
            <span>AI Suggester</span>
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

      {saveMessage && (
        <div className="p-3.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{saveMessage}</span>
        </div>
      )}

      {/* Supabase Discovered Tables Explorer Panel */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
        <div className="p-4 bg-slate-900 text-white flex items-center justify-between">
          <div className="flex items-center space-x-2.5">
            <div className="p-1.5 rounded bg-emerald-500/20 text-emerald-400">
              <Database className="w-4 h-4" />
            </div>
            <div>
              <div className="text-sm font-bold flex items-center space-x-2">
                <span>Supabase PostgreSQL Tables</span>
                <span className="text-xs px-2 py-0.5 rounded-full bg-slate-800 text-emerald-400 font-mono font-normal">
                  {supabaseTables.length} tables discovered
                </span>
              </div>
              <p className="text-[11px] text-slate-400">
                Click any table below to immediately route your active Excel worksheet to it.
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-2">
            <button
              onClick={() => setShowSupabasePanel(!showSupabasePanel)}
              className="text-xs text-slate-400 hover:text-white px-2 py-1 rounded bg-slate-800"
            >
              {showSupabasePanel ? 'Collapse' : 'Expand'}
            </button>
          </div>
        </div>

        {showSupabasePanel && (
          <div className="p-4 bg-slate-50 border-b border-slate-200">
            {isLoadingSchema ? (
              <div className="p-4 text-center text-xs text-slate-500 flex items-center justify-center space-x-2">
                <RefreshCw className="w-4 h-4 animate-spin text-emerald-600" />
                <span>Discovering tables and column schemas from Supabase...</span>
              </div>
            ) : supabaseTables.length > 0 ? (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                {supabaseTables.map((tbl, tblIdx) => {
                  const isMappedToActive = activeMapping?.supabaseTable?.toLowerCase() === tbl.name.toLowerCase();
                  return (
                    <div
                      key={`${tbl.name}-${tblIdx}`}
                      className={`p-3.5 rounded-lg border text-left transition-all ${
                        isMappedToActive
                          ? 'border-emerald-500 bg-white ring-2 ring-emerald-500/20 shadow-xs'
                          : 'border-slate-200 bg-white hover:border-slate-300'
                      }`}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <div className="flex items-center space-x-1.5 font-mono text-xs font-bold text-slate-900">
                          <Table className="w-3.5 h-3.5 text-emerald-600" />
                          <span>public.{tbl.name}</span>
                        </div>
                        <span className="text-[10px] px-1.5 py-0.5 rounded font-mono bg-slate-100 text-slate-600">
                          {tbl.columns.length} cols
                        </span>
                      </div>

                      {/* Columns preview */}
                      <div className="flex flex-wrap gap-1 mb-3">
                        {tbl.columns.slice(0, 6).map(c => (
                          <span
                            key={c.name}
                            className={`text-[10px] px-1.5 py-0.5 rounded font-mono ${
                              c.isPrimary
                                ? 'bg-amber-100 text-amber-800 font-semibold'
                                : 'bg-slate-100 text-slate-700'
                            }`}
                          >
                            {c.name}{c.isPrimary ? '🔑' : ''}
                          </span>
                        ))}
                        {tbl.columns.length > 6 && (
                          <span className="text-[10px] px-1 py-0.5 text-slate-400">
                            +{tbl.columns.length - 6} more
                          </span>
                        )}
                      </div>

                      <div className="flex items-center justify-between pt-2 border-t border-slate-100">
                        {isMappedToActive ? (
                          <span className="inline-flex items-center space-x-1 text-[11px] font-semibold text-emerald-600">
                            <CheckCircle2 className="w-3.5 h-3.5" />
                            <span>Active Destination</span>
                          </span>
                        ) : (
                          <button
                            type="button"
                            onClick={() => handleSelectSupabaseTable(tbl.name)}
                            className="inline-flex items-center space-x-1 text-xs font-semibold text-emerald-700 hover:text-emerald-800 hover:underline"
                          >
                            <span>Map Sheet ➔ {tbl.name}</span>
                            <ArrowRight className="w-3 h-3" />
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="p-3 text-center text-xs text-slate-500">
                No tables detected. Verify your Supabase URL and service key in the Supabase tab.
              </div>
            )}
          </div>
        )}
      </div>

      {/* Active Mapping Setup Card */}
      {activeMapping && (
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs space-y-5">
          {/* Multi-Sheet Routing Toolbar */}
          <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs">
            <div className="flex items-center space-x-2">
              <span className="font-bold text-slate-800">Multi-Sheet Routing Strategy:</span>
              <span className="text-slate-500">Configure how workbook tabs route to Supabase tables</span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={() => handleConsolidateAllSheets(activeMapping.supabaseTable || 'students')}
                className="px-2.5 py-1.5 rounded-lg bg-purple-50 border border-purple-300 text-purple-700 hover:bg-purple-100 font-semibold transition-colors"
                title="Route all worksheets into a single common Supabase table"
              >
                Merge All Sheets ➔ public.{activeMapping.supabaseTable || 'students'}
              </button>

              <button
                type="button"
                onClick={handleSplitAllSheetsSeparate}
                className="px-2.5 py-1.5 rounded-lg bg-blue-50 border border-blue-300 text-blue-700 hover:bg-blue-100 font-semibold transition-colors"
                title="Route each worksheet to its own separate 1:1 Supabase table"
              >
                Separate Tables (1:1)
              </button>

              <button
                type="button"
                onClick={() => onNavigate('analyzer')}
                className="px-2.5 py-1.5 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-700 hover:bg-emerald-100 font-semibold transition-colors flex items-center space-x-1"
                title="View and deploy PostgreSQL DDL migrations for this schema"
              >
                <Database className="w-3.5 h-3.5 text-emerald-600" />
                <span>View Generated SQL Schema</span>
              </button>
            </div>
          </div>

          {/* Sheet Selector Tabs */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center space-x-1.5">
                <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                <span>Worksheet Routing Tabs ({currentMappings.length})</span>
              </span>
              <span className="text-xs text-slate-500">
                Switch between Excel sheets in your workbook
              </span>
            </div>

            <div className="flex border-b border-slate-200 space-x-2 pb-0.5 overflow-x-auto">
              {currentMappings.map((m) => {
                const isActive = m.id === activeMapping.id;
                return (
                  <button
                    key={m.id}
                    onClick={() => setActiveSheetId(m.id)}
                    className={`px-4 py-2.5 rounded-t-lg text-xs font-medium whitespace-nowrap transition-all border-b-2 flex items-center space-x-2 ${
                      isActive
                        ? 'border-emerald-600 text-emerald-700 bg-emerald-50/50 font-bold'
                        : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                    }`}
                  >
                    <span>{m.worksheetName}</span>
                    <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-slate-200 text-slate-700 font-mono font-semibold">
                      ➔ {m.supabaseTable}
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Source Sheet to Target Supabase Table Router */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 p-4 rounded-xl bg-slate-50 border border-slate-200">
            {/* Left: Source Excel Sheet */}
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center space-x-1">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                  <span>1. Source Excel Worksheet</span>
                </label>
                {activeSheetAnalysis && (
                  <span className="text-[10px] text-slate-500 font-mono">
                    {activeSheetAnalysis.totalRows} rows · {activeSheetAnalysis.headers.length} headers
                  </span>
                )}
              </div>

              {currentAnalysis && currentAnalysis.worksheets.length > 0 ? (
                <select
                  value={activeMapping.worksheetName}
                  onChange={(e) => {
                    const selectedWs = currentAnalysis.worksheets.find(w => w.sheetName === e.target.value);
                    if (selectedWs) {
                      const newCols: ColumnMapping[] = selectedWs.headers.map((h, idx) => ({
                        id: `cm-${Date.now()}-${idx}`,
                        excelColumn: h.colLetter,
                        excelHeader: h.name,
                        supabaseColumn: smartSanitizeIdentifier(h.name, `col_${idx + 1}`),
                        dataType: 'text',
                        required: idx === 0,
                        uniqueKey: idx === 0,
                        transformation: 'trim'
                      }));
                      handleUpdateActiveMapping({
                        worksheetName: selectedWs.sheetName,
                        columns: newCols,
                        headerRow: selectedWs.detectedHeaderRow || 1,
                        dataStartRow: selectedWs.detectedDataStartRow || 2
                      });
                    }
                  }}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 bg-white font-medium text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                >
                  {currentAnalysis.worksheets.map((w, wIdx) => (
                    <option key={`${w.sheetName}-${wIdx}`} value={w.sheetName}>
                      {w.sheetName} ({w.totalRows} rows, {w.headers.length} detected columns)
                    </option>
                  ))}
                </select>
              ) : (
                <input
                  type="text"
                  value={activeMapping.worksheetName}
                  onChange={(e) => handleUpdateActiveMapping({ worksheetName: e.target.value })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 font-medium text-xs text-slate-800"
                />
              )}

              <div className="flex flex-wrap items-center gap-4 text-[11px] text-slate-600 pt-1">
                <label className="flex items-center space-x-1.5">
                  <span className="font-semibold text-slate-700">Header Row:</span>
                  <input
                    type="number"
                    min={1}
                    max={200}
                    value={activeMapping.headerRow || 1}
                    onChange={(e) => handleUpdateActiveMapping({ headerRow: parseInt(e.target.value) || 1 })}
                    className="w-14 px-2 py-0.5 rounded border border-slate-300 bg-white font-bold font-mono text-center text-xs focus:ring-1 focus:ring-emerald-500"
                  />
                </label>
                <label className="flex items-center space-x-1.5">
                  <span className="font-semibold text-slate-700">Data Start Row:</span>
                  <input
                    type="number"
                    min={1}
                    max={200}
                    value={activeMapping.dataStartRow || 2}
                    onChange={(e) => handleUpdateActiveMapping({ dataStartRow: parseInt(e.target.value) || 2 })}
                    className="w-14 px-2 py-0.5 rounded border border-slate-300 bg-white font-bold font-mono text-center text-xs focus:ring-1 focus:ring-emerald-500"
                  />
                </label>
              </div>
            </div>

            {/* Right: Target Supabase Table */}
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center space-x-1">
                  <Database className="w-3.5 h-3.5 text-blue-600" />
                  <span>2. Target Supabase Table (Merge Into)</span>
                </label>
                <span className="text-[10px] text-emerald-600 font-semibold">PostgreSQL</span>
              </div>

              <div className="flex items-center space-x-2">
                <select
                  value={activeMapping.supabaseTable}
                  onChange={(e) => handleSelectSupabaseTable(e.target.value)}
                  className="flex-1 px-3 py-2 rounded-lg border border-slate-300 bg-white font-bold font-mono text-xs text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                >
                  {supabaseTables.map((t, tIdx) => (
                    <option key={`${t.name}-${tIdx}`} value={t.name}>
                      public.{t.name} ({t.columns.length} columns)
                    </option>
                  ))}
                  {!supabaseTables.some(t => t.name === activeMapping.supabaseTable) && (
                    <option value={activeMapping.supabaseTable}>
                      public.{activeMapping.supabaseTable} (Custom)
                    </option>
                  )}
                </select>

                <button
                  type="button"
                  onClick={handleAutoMatchColumns}
                  className="px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-700 hover:bg-emerald-100 text-xs font-semibold whitespace-nowrap shadow-2xs"
                  title="Automatically match Excel columns with Supabase table columns by name"
                >
                  Auto-Match Columns
                </button>
              </div>

              <div className="flex flex-wrap items-center gap-2 pt-1.5">
                <span className="text-[11px] font-bold text-slate-700 uppercase tracking-wide flex items-center space-x-1">
                  <Key className="w-3.5 h-3.5 text-amber-600" />
                  <span>Unique Merge Key:</span>
                </span>
                <select
                  value={mergeKeys[0] || ''}
                  onChange={(e) => handleSetPrimaryMergeKey(e.target.value)}
                  className="px-2.5 py-1 rounded-md border border-slate-300 bg-white font-mono text-xs text-slate-900 font-bold focus:ring-1 focus:ring-emerald-500 shadow-2xs"
                >
                  <option value="">None (Insert Only — Safe, recommended)</option>
                  {activeMapping.columns.map(c => (
                    <option key={c.id} value={c.supabaseColumn}>
                      🔑 {c.supabaseColumn} (col {c.excelColumn}: {c.excelHeader || c.excelColumn})
                    </option>
                  ))}
                </select>
                {mergeKeys.length > 0 ? (
                  <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 font-mono font-bold text-[11px] border border-amber-300">
                    ON CONFLICT ({mergeKeys.join(', ')}) DO UPDATE
                  </span>
                ) : (
                  <span className="px-2 py-0.5 rounded bg-blue-50 text-blue-700 font-medium text-[11px] border border-blue-200">
                    Direct INSERT mode (no unique key required)
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* Merge Strategy Summary Banner */}
          <div className="p-3.5 rounded-lg bg-emerald-50/80 border border-emerald-200 text-xs text-emerald-950 flex items-start space-x-2.5">
            <GitFork className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
            <div className="space-y-0.5">
              <span className="font-bold text-emerald-900">Merge & Upsert Strategy:</span>
              <p className="text-emerald-800 leading-relaxed">
                Excel sheet <strong>&quot;{activeMapping.worksheetName}&quot;</strong> merges into Supabase table <strong>&quot;public.{activeMapping.supabaseTable}&quot;</strong>.
                {mergeKeys.length > 0 ? (
                  <> When an Excel row has the same <strong>{mergeKeys.join(', ')}</strong> as an existing record in Supabase, the Supabase record is <strong>MERGED &amp; UPDATED</strong>. If the column does not have a UNIQUE constraint in PostgreSQL, it automatically falls back to safe <strong>INSERT</strong>.</>
                ) : (
                  <> Operating in <strong>Direct INSERT Mode</strong>. All verified rows will be appended as new records in <code>public.{activeMapping.supabaseTable}</code>. If your table has a unique column (e.g. <code>admission_no</code>), you can select it from the dropdown above to merge duplicates.</>
                )}
              </p>
            </div>
          </div>

          {/* Column Mappings Table */}
          <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
            <div className="p-3.5 bg-slate-100 border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <span className="text-xs font-bold text-slate-900 uppercase tracking-wider">
                  Column Mappings ({activeMapping.columns.length})
                </span>
                <span className="text-[11px] text-slate-500 hidden sm:inline">
                  Map each Excel column to its corresponding Supabase field
                </span>
              </div>

              <div className="flex items-center space-x-2">
                <button
                  type="button"
                  onClick={handleAddColumn}
                  className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold shadow-2xs"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>Add Column</span>
                </button>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="min-w-[780px] w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="px-3 py-3 w-16 text-center">Col</th>
                    <th className="px-4 py-3">Excel Column Header</th>
                    <th className="px-3 py-3 text-slate-500">Excel Sample Preview</th>
                    <th className="px-2 py-3 text-center"></th>
                    <th className="px-4 py-3">Target Supabase Column</th>
                    <th className="px-3 py-3 w-28">Data Type</th>
                    <th className="px-3 py-3 text-center">Merge Key?</th>
                    <th className="px-4 py-3">Transformation</th>
                    <th className="px-3 py-3 text-center w-12">Delete</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {activeMapping.columns.map((col) => {
                    const sampleVal = getSampleValuesForCol(col.excelColumn, col.excelHeader);
                    const isSupabaseColMatched = activeSupabaseTable?.columns.some(c => c.name === col.supabaseColumn);

                    return (
                      <tr key={col.id} className="hover:bg-slate-50 transition-colors">
                        {/* Excel Col Letter */}
                        <td className="px-3 py-2.5 text-center">
                          <input
                            type="text"
                            value={col.excelColumn}
                            onChange={(e) => handleUpdateColumn(col.id, { excelColumn: e.target.value.toUpperCase() })}
                            className="w-10 px-1.5 py-1 rounded border border-slate-300 font-mono font-bold text-center focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </td>

                        {/* Excel Header */}
                        <td className="px-4 py-2.5">
                          <input
                            type="text"
                            value={col.excelHeader}
                            onChange={(e) => handleUpdateColumn(col.id, { excelHeader: e.target.value })}
                            className="w-full px-2.5 py-1 rounded border border-slate-300 font-medium text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </td>

                        {/* Excel Sample Preview */}
                        <td className="px-3 py-2.5 text-slate-500 font-mono text-[11px] truncate max-w-xs" title={sampleVal}>
                          {sampleVal ? (
                            <span className="text-slate-600">{sampleVal}</span>
                          ) : (
                            <span className="text-slate-300 italic">No sample</span>
                          )}
                        </td>

                        {/* Arrow */}
                        <td className="px-2 py-2.5 text-center text-slate-400">
                          <ArrowRight className="w-3.5 h-3.5 inline" />
                        </td>

                        {/* Supabase Column Dropdown or Input */}
                        <td className="px-4 py-2.5">
                          {activeSupabaseTable && activeSupabaseTable.columns.length > 0 ? (
                            <select
                              value={col.supabaseColumn}
                              onChange={(e) => {
                                const chosen = e.target.value;
                                const matchedCol = activeSupabaseTable.columns.find(c => c.name === chosen);
                                let dt = col.dataType;
                                let trans = col.transformation;
                                if (matchedCol) {
                                  if (matchedCol.type === 'integer') {
                                    dt = 'integer';
                                    trans = 'parse_number';
                                  } else if (matchedCol.type === 'decimal') {
                                    dt = 'decimal';
                                    trans = 'parse_number';
                                  } else if (matchedCol.type === 'date') {
                                    dt = 'date';
                                    trans = 'parse_date';
                                  } else if (matchedCol.type === 'boolean') {
                                    dt = 'boolean';
                                    trans = 'yes_no_to_boolean';
                                  } else if (matchedCol.type === 'timestamp') {
                                    dt = 'timestamp';
                                    trans = 'parse_date';
                                  } else {
                                    dt = 'text';
                                    trans = 'trim';
                                  }
                                }
                                handleUpdateColumn(col.id, {
                                  supabaseColumn: chosen,
                                  dataType: dt,
                                  transformation: trans,
                                  uniqueKey: col.uniqueKey
                                });
                              }}
                              className="w-full px-2.5 py-1 rounded border border-slate-300 font-mono font-bold text-xs text-emerald-900 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                            >
                              {activeSupabaseTable.columns.map(sc => (
                                <option key={sc.name} value={sc.name}>
                                  {sc.name} ({sc.type}{sc.isPrimary ? ', PK' : ''})
                                </option>
                              ))}
                              {!isSupabaseColMatched && col.supabaseColumn && (
                                <option value={col.supabaseColumn}>
                                  {col.supabaseColumn} (Custom)
                                </option>
                              )}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={col.supabaseColumn}
                              onChange={(e) => handleUpdateColumn(col.id, { supabaseColumn: e.target.value })}
                              className="w-full px-2.5 py-1 rounded border border-slate-300 font-mono font-bold text-xs text-emerald-900 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                            />
                          )}
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

                        {/* Merge Key Button */}
                        <td className="px-3 py-2.5 text-center">
                          <button
                            type="button"
                            onClick={() => handleToggleMergeKey(col.id)}
                            className={`px-2.5 py-1 rounded-md text-[11px] font-bold inline-flex items-center space-x-1 transition-colors ${
                              col.uniqueKey
                                ? 'bg-amber-100 text-amber-800 border border-amber-300 shadow-2xs'
                                : 'bg-slate-100 text-slate-400 hover:text-slate-700 hover:bg-slate-200 border border-transparent'
                            }`}
                            title={col.uniqueKey ? 'Active Unique Merge Key' : 'Click to make this column the primary merge key'}
                          >
                            <Key className="w-3 h-3" />
                            <span>{col.uniqueKey ? 'MERGE KEY' : 'Set Key'}</span>
                          </button>
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

                        {/* Delete Button */}
                        <td className="px-3 py-2.5 text-center">
                          <button
                            type="button"
                            onClick={() => handleDeleteColumn(col.id)}
                            className="p-1 rounded text-slate-400 hover:text-rose-600 hover:bg-rose-50"
                            title="Remove column mapping"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* AI Presets & Inbuilt Analyser Modal */}
      <AiPresetsModal
        isOpen={showAiPresetsModal}
        onClose={() => setShowAiPresetsModal(false)}
        currentAnalysis={currentAnalysis || null}
        activeMappings={currentMappings}
        onApplyPreset={handleApplyPreset}
        onLoadPresetWorkbook={() => {
          onNavigate('analyzer');
        }}
        supabaseConfig={supabaseConfig}
        supabaseTables={supabaseTables}
        prefillFromCustomMappings={presetPrefill}
      />
    </div>
  );
};
