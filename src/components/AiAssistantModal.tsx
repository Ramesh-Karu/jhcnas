import React, { useState } from 'react';
import { 
  Sparkles, 
  X, 
  Check, 
  Table, 
  Layers, 
  AlertCircle, 
  ArrowRight,
  Bot
} from 'lucide-react';
import { SheetAnalysis, WorksheetMapping, ColumnMapping } from '../types';
import { GeminiWorkbookService } from '../services/geminiService';

interface AiAssistantModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentSheet: SheetAnalysis | null;
  allSheets: SheetAnalysis[];
  workbookName: string;
  onApplyAiSuggestions: (suggestedMapping: WorksheetMapping) => void;
}

export const AiAssistantModal: React.FC<AiAssistantModalProps> = ({
  isOpen,
  onClose,
  currentSheet,
  allSheets,
  workbookName,
  onApplyAiSuggestions
}) => {
  const [selectedSheetName, setSelectedSheetName] = useState<string>(
    currentSheet?.sheetName || allSheets[0]?.sheetName || ''
  );
  const [isAnalyzing, setIsAnalyzing] = useState<boolean>(false);
  const [result, setResult] = useState<{
    suggestedTable: string;
    headerRow: number;
    dataStartRow: number;
    sectionHeadingTargetCol?: string;
    columns: ColumnMapping[];
    reasoning: string;
  } | null>(null);

  if (!isOpen) return null;

  const targetSheet = allSheets.find(s => s.sheetName === selectedSheetName) || currentSheet || allSheets[0];

  const handleRunAiAnalysis = async () => {
    if (!targetSheet) return;
    setIsAnalyzing(true);
    setResult(null);

    try {
      const suggestion = await GeminiWorkbookService.analyzeSheetAndSuggestMappings(targetSheet, workbookName);
      setResult(suggestion);
    } catch (err) {
      console.error('AI Analysis failed:', err);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleApply = () => {
    if (!result || !targetSheet) return;
    const newMapping: WorksheetMapping = {
      id: `wm-ai-${Date.now()}`,
      workbookName,
      worksheetName: targetSheet.sheetName,
      supabaseTable: result.suggestedTable,
      headerRow: result.headerRow,
      dataStartRow: result.dataStartRow,
      sectionHeadingTargetCol: result.sectionHeadingTargetCol,
      enabled: true,
      columns: result.columns
    };
    onApplyAiSuggestions(newMapping);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="p-6 border-b border-slate-200 flex items-center justify-between bg-gradient-to-r from-purple-50 to-indigo-50/40">
          <div className="flex items-center space-x-3">
            <div className="p-2 rounded-xl bg-purple-600 text-white shadow-xs">
              <Sparkles className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900">AI Schema & Mapping Assistant</h2>
              <p className="text-xs text-slate-500">
                Initial configuration assistant for unknown Excel structures (Section 29).
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto space-y-5">
          {/* Target Sheet Picker */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-3.5 rounded-xl bg-slate-50 border border-slate-200">
            <div>
              <label className="text-xs font-semibold uppercase tracking-wider text-slate-600 block mb-1">
                Select Worksheet to Analyze
              </label>
              <select
                value={selectedSheetName}
                onChange={(e) => {
                  setSelectedSheetName(e.target.value);
                  setResult(null);
                }}
                className="px-3 py-1.5 rounded-lg border border-slate-300 text-xs font-semibold text-slate-800 bg-white"
              >
                {allSheets.map(s => (
                  <option key={s.sheetName} value={s.sheetName}>
                    {s.sheetName} ({s.totalRows} rows, {s.totalColumns} cols)
                  </option>
                ))}
              </select>
            </div>

            <button
              onClick={handleRunAiAnalysis}
              disabled={isAnalyzing}
              className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-purple-600 hover:bg-purple-700 text-white text-xs font-semibold shadow-xs transition-colors"
            >
              <Bot className={`w-4 h-4 ${isAnalyzing ? 'animate-bounce' : ''}`} />
              <span>{isAnalyzing ? 'Analyzing Worksheet...' : 'Generate AI Mapping Suggestion'}</span>
            </button>
          </div>

          {/* AI Result View */}
          {result ? (
            <div className="space-y-4">
              {/* Reasoning Card */}
              <div className="p-4 rounded-xl bg-purple-50/70 border border-purple-200 text-xs text-purple-900 leading-relaxed">
                <span className="font-bold">AI Analysis Reasoning: </span>
                {result.reasoning}
              </div>

              {/* Proposed Structural Settings */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-slate-500 font-medium">Target Table</span>
                  <div className="font-mono font-bold text-teal-800 mt-0.5">{result.suggestedTable}</div>
                </div>
                <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-slate-500 font-medium">Header Row</span>
                  <div className="font-mono font-bold text-slate-900 mt-0.5">{result.headerRow}</div>
                </div>
                <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-slate-500 font-medium">Data Start Row</span>
                  <div className="font-mono font-bold text-slate-900 mt-0.5">{result.dataStartRow}</div>
                </div>
                <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-slate-500 font-medium">Section Propagation</span>
                  <div className="font-mono font-bold text-purple-800 mt-0.5">
                    {result.sectionHeadingTargetCol || 'None'}
                  </div>
                </div>
              </div>

              {/* Columns Table */}
              <div className="rounded-xl border border-slate-200 overflow-hidden">
                <div className="p-3 bg-slate-100 text-xs font-semibold text-slate-700 uppercase">
                  Suggested Column Transformations & Unique Keys
                </div>
                <div className="max-h-56 overflow-y-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-50 text-slate-600 uppercase border-b border-slate-200">
                      <tr>
                        <th className="px-3 py-2 w-12">Col</th>
                        <th className="px-3 py-2">Excel Header</th>
                        <th className="px-3 py-2">Supabase Field</th>
                        <th className="px-3 py-2">Type</th>
                        <th className="px-3 py-2">Transformation</th>
                        <th className="px-3 py-2 text-center">Key</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono">
                      {result.columns.map((c, i) => (
                        <tr key={i} className="hover:bg-slate-50">
                          <td className="px-3 py-2 font-bold text-purple-700">{c.excelColumn}</td>
                          <td className="px-3 py-2 font-sans text-slate-900">{c.excelHeader}</td>
                          <td className="px-3 py-2 text-teal-700">{c.supabaseColumn}</td>
                          <td className="px-3 py-2 text-slate-600">{c.dataType}</td>
                          <td className="px-3 py-2 text-slate-700 font-sans">{c.transformation}</td>
                          <td className="px-3 py-2 text-center">
                            {c.uniqueKey ? <span className="text-purple-600 font-bold">YES</span> : '-'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          ) : (
            <div className="text-center py-10 text-slate-500 text-xs space-y-2">
              <Bot className="w-8 h-8 text-purple-400 mx-auto" />
              <p>Click <span className="font-semibold text-slate-700">Generate AI Mapping Suggestion</span> to inspect header candidates and infer database schemas.</p>
              <p className="text-[11px] text-slate-400">Deterministic production sync runs do not depend on this service.</p>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 hover:bg-slate-100"
          >
            Cancel
          </button>

          {result && (
            <button
              onClick={handleApply}
              className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-xs"
            >
              <Check className="w-4 h-4" />
              <span>Apply Mappings to Pipeline</span>
            </button>
          )}
        </div>
      </div>
    </div>
  );
};
