import React, { useState } from 'react';
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
  CheckCircle2
} from 'lucide-react';
import { 
  WorksheetMapping, 
  ColumnMapping, 
  TransformationType, 
  DataType, 
  NavigationTab 
} from '../types';
import { getDefaultSampleMappings } from '../services/sampleWorkbook';

interface MappingsViewProps {
  mappings: WorksheetMapping[];
  onSaveMappings: (mappings: WorksheetMapping[]) => void;
  onNavigate: (tab: NavigationTab) => void;
  onOpenAiAssistant: () => void;
}

export const MappingsView: React.FC<MappingsViewProps> = ({
  mappings,
  onSaveMappings,
  onNavigate,
  onOpenAiAssistant
}) => {
  const [activeSheetId, setActiveSheetId] = useState<string>(mappings[0]?.id || '');
  const [currentMappings, setCurrentMappings] = useState<WorksheetMapping[]>(mappings);
  const [saveMessage, setSaveMessage] = useState<string | null>(null);

  const activeMapping = currentMappings.find(m => m.id === activeSheetId) || currentMappings[0];

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

  if (!activeMapping) {
    return <div className="p-8 text-center text-slate-500">No active mappings defined.</div>;
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

      {saveMessage && (
        <div className="p-3.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600" />
          <span>{saveMessage}</span>
        </div>
      )}

      {/* Worksheet Switcher Pills */}
      <div className="flex border-b border-slate-200 space-x-2 pb-0.5 overflow-x-auto">
        {currentMappings.map((m) => (
          <button
            key={m.id}
            id={`mapping-tab-${m.id}`}
            onClick={() => setActiveSheetId(m.id)}
            className={`px-4 py-2.5 rounded-t-lg text-sm font-medium whitespace-nowrap transition-all border-b-2 flex items-center space-x-2 ${
              activeSheetId === m.id
                ? 'border-emerald-600 text-emerald-700 bg-emerald-50/50 font-semibold'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
            }`}
          >
            <span>{m.worksheetName}</span>
            <span className="text-xs px-1.5 py-0.2 rounded-full bg-slate-200 text-slate-600 font-mono">
              → {m.supabaseTable}
            </span>
          </button>
        ))}
      </div>

      {/* Worksheet Configuration Settings */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs">
        <h3 className="text-sm font-semibold text-slate-900 mb-3 flex items-center space-x-2">
          <Table className="w-4 h-4 text-emerald-600" />
          <span>Worksheet Configuration: {activeMapping.worksheetName}</span>
        </h3>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
          <div>
            <label className="block text-slate-600 mb-1 font-semibold uppercase tracking-wider">Target Supabase Table</label>
            <input
              type="text"
              value={activeMapping.supabaseTable}
              onChange={(e) => handleUpdateActiveMapping({ supabaseTable: e.target.value })}
              className="w-full px-3 py-1.5 rounded-lg border border-slate-300 font-mono focus:ring-1 focus:ring-emerald-500 focus:outline-none"
            />
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
