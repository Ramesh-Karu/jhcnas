import React, { useState, useEffect } from 'react';
import { 
  Sparkles, 
  Database, 
  Server, 
  Trash2, 
  Check, 
  Copy, 
  RefreshCw, 
  BookOpen, 
  X, 
  AlertCircle, 
  CheckCircle2, 
  ExternalLink,
  Plus,
  Sliders,
  Layers,
  ArrowRight,
  ShieldCheck,
  Calendar,
  DollarSign,
  GraduationCap,
  Table,
  Filter
} from 'lucide-react';
import { 
  AiWorkbookPreset, 
  WorkbookAnalysis, 
  WorksheetMapping, 
  SupabaseConfig, 
  SupabaseTableInfo,
  WorkbookArchetype
} from '../types';
import { ApiClient } from '../services/apiClient';
import { StorageService } from '../services/storage';

interface AiPresetsModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentAnalysis: WorkbookAnalysis | null;
  activeMappings: WorksheetMapping[];
  onApplyPreset: (preset: AiWorkbookPreset) => void;
  onApplyAllPresets?: (presets: AiWorkbookPreset[]) => void;
  onLoadPresetWorkbook: (presetId: 'timetable' | 'donations' | 'teacher_allocations' | string) => void;
  supabaseConfig?: SupabaseConfig;
  supabaseTables?: SupabaseTableInfo[];
  prefillFromCustomMappings?: boolean;
}

export const AiPresetsModal: React.FC<AiPresetsModalProps> = ({
  isOpen,
  onClose,
  currentAnalysis,
  activeMappings,
  onApplyPreset,
  onApplyAllPresets,
  onLoadPresetWorkbook,
  supabaseConfig,
  supabaseTables = [],
  prefillFromCustomMappings = false,
}) => {
  const [activeTab, setActiveTab] = useState<'LIBRARY' | 'CREATE_AI'>('LIBRARY');
  const [presets, setPresets] = useState<AiWorkbookPreset[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [categoryFilter, setCategoryFilter] = useState<'ALL' | 'STUDENT_BATCHES' | 'TEMPLATES'>('ALL');

  // Create Custom Preset Form State
  const [customName, setCustomName] = useState<string>('');
  const [customDesc, setCustomDesc] = useState<string>('');
  const [customArchetype, setCustomArchetype] = useState<WorkbookArchetype>('STANDARD_TABULAR');
  const [customTags, setCustomTags] = useState<string>('Custom AI, School, Production');
  const [isGeneratingAi, setIsGeneratingAi] = useState<boolean>(false);
  const [previewPreset, setPreviewPreset] = useState<AiWorkbookPreset | null>(null);
  const [isSavingPreset, setIsSavingPreset] = useState<boolean>(false);
  const [filterQuery, setFilterQuery] = useState<string>('');

  // Load presets list from Server / Supabase & mirror to local storage
  const loadPresets = async () => {
    setIsLoading(true);
    try {
      const res = await ApiClient.getPresets();
      if (res.success && res.presets) {
        setPresets(res.presets);
        StorageService.saveAiPresets(res.presets);
      } else {
        // Fallback to local storage if network glitch
        const localPresets = StorageService.getAiPresets();
        if (localPresets.length > 0) {
          setPresets(localPresets);
        }
      }
    } catch (e: any) {
      console.error('Failed to load presets:', e);
      const localPresets = StorageService.getAiPresets();
      if (localPresets.length > 0) {
        setPresets(localPresets);
      }
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      loadPresets();
      if (prefillFromCustomMappings && activeMappings.length > 0) {
        setActiveTab('CREATE_AI');
        initFromActiveMappings();
      } else {
        setActiveTab('LIBRARY');
      }
    }
  }, [isOpen, prefillFromCustomMappings]);

  // Pre-fill creation form from active mappings
  const initFromActiveMappings = () => {
    const filename = currentAnalysis?.filename || activeMappings[0]?.workbookName || 'Custom Workbook';
    const baseName = filename.replace(/\.xlsx?$/i, '');
    setCustomName(`${baseName} Preset`);
    setCustomDesc(`Custom mapping preset for ${filename} (${activeMappings.length} mapped sheets)`);
    setCustomArchetype(currentAnalysis?.detectedArchetype || 'STANDARD_TABULAR');

    const syntheticPreset: AiWorkbookPreset = {
      id: `preset-${Date.now()}`,
      name: `${baseName} Preset`,
      description: `Custom mapping preset for ${filename}`,
      archetype: currentAnalysis?.detectedArchetype || 'STANDARD_TABULAR',
      badge: currentAnalysis?.archetypeBadge || '📊 Custom AI Preset',
      isSystem: false,
      skipMergedYearRows: true, // Industrial standard: always exclude merged year rows
      tags: ['Custom Mapping', currentAnalysis?.detectedArchetype || 'STANDARD_TABULAR'],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      sheetMappings: activeMappings.map(m => ({
        ...m,
        skipMergedYearRows: true
      }))
    };
    setPreviewPreset(syntheticPreset);
  };

  // Run Inbuilt AI Analyzer to create preset
  const handleGenerateAiPreset = async () => {
    if (!currentAnalysis) {
      setStatusMessage({
        type: 'error',
        text: 'Please upload, paste, or fetch a workbook first so the AI Analyser can inspect sheets, headers, and merged structures.'
      });
      return;
    }

    setIsGeneratingAi(true);
    setStatusMessage(null);
    try {
      const res = await ApiClient.createPresetWithAi({
        workbook: currentAnalysis,
        supabaseTables,
        customName: customName.trim() || undefined,
        customDescription: customDesc.trim() || undefined,
        tags: customTags.split(',').map(t => t.trim()).filter(Boolean),
      });

      if (res.success && res.preset) {
        setPreviewPreset(res.preset);
        setCustomName(res.preset.name);
        setCustomDesc(res.preset.description || '');
        setCustomArchetype(res.preset.archetype);
        setStatusMessage({
          type: 'success',
          text: `✨ Inbuilt AI Analyzer generated preset with ${res.preset.sheetMappings.length} sheet mappings! Merged year rows are automatically flagged to be excluded from database records.`
        });
      } else {
        setStatusMessage({
          type: 'error',
          text: res.error || 'Failed to analyze workbook and create preset.'
        });
      }
    } catch (e: any) {
      setStatusMessage({ type: 'error', text: e.message });
    } finally {
      setIsGeneratingAi(false);
    }
  };

  // Permanently save preset to Server & Supabase
  const handleSavePermanentPreset = async () => {
    if (!previewPreset) return;

    setIsSavingPreset(true);
    setStatusMessage(null);
    try {
      const presetToSave: AiWorkbookPreset = {
        ...previewPreset,
        name: customName.trim() || previewPreset.name,
        description: customDesc.trim() || previewPreset.description,
        archetype: customArchetype,
        skipMergedYearRows: true,
        tags: customTags.split(',').map(t => t.trim()).filter(Boolean),
        updatedAt: new Date().toISOString(),
      };

      const res = await ApiClient.savePreset({
        preset: presetToSave,
        supabase: supabaseConfig,
      });

      if (res.success && res.preset) {
        setStatusMessage({
          type: 'success',
          text: `🎉 Permanently saved preset '${res.preset.name}' to server disk${res.syncedToSupabase ? ' & Supabase PostgreSQL database' : ''}!`
        });
        await loadPresets();
        setTimeout(() => {
          setActiveTab('LIBRARY');
        }, 1200);
      } else {
        setStatusMessage({
          type: 'error',
          text: res.error || 'Failed to permanently save preset.'
        });
      }
    } catch (e: any) {
      setStatusMessage({ type: 'error', text: e.message });
    } finally {
      setIsSavingPreset(false);
    }
  };

  // Delete custom preset
  const handleDeletePreset = async (id: string, name: string) => {
    if (!confirm(`Are you sure you want to permanently delete custom preset '${name}'? This will remove it from both the server and Supabase.`)) {
      return;
    }

    try {
      const res = await ApiClient.deletePreset(id, supabaseConfig);
      if (res.success) {
        setStatusMessage({ type: 'info', text: `Deleted preset '${name}'` });
        loadPresets();
      } else {
        setStatusMessage({ type: 'error', text: res.error || 'Could not delete preset' });
      }
    } catch (e: any) {
      setStatusMessage({ type: 'error', text: e.message });
    }
  };

  const handleCopyPresetJson = (preset: AiWorkbookPreset) => {
    navigator.clipboard.writeText(JSON.stringify(preset, null, 2));
    setCopiedId(preset.id);
    setTimeout(() => setCopiedId(null), 2500);
  };

  if (!isOpen) return null;

  const studentBatchPresets = presets.filter(p =>
    p.id.startsWith('preset-jhc-students') || p.tags?.includes('students_') || p.name.includes('Student Roster')
  );

  const filteredPresets = presets.filter(p => {
    if (categoryFilter === 'STUDENT_BATCHES') {
      if (!p.id.startsWith('preset-jhc-students') && !p.tags?.includes('students_') && !p.name.includes('Student Roster')) return false;
    } else if (categoryFilter === 'TEMPLATES') {
      if (p.id.startsWith('preset-jhc-students') || p.tags?.includes('students_') || p.name.includes('Student Roster')) return false;
    }
    if (!filterQuery) return true;
    const q = filterQuery.toLowerCase();
    return (
      p.name.toLowerCase().includes(q) ||
      (p.description && p.description.toLowerCase().includes(q)) ||
      p.archetype.toLowerCase().includes(q) ||
      (p.tags && p.tags.some(t => t.toLowerCase().includes(q))) ||
      (p.filenamePattern && p.filenamePattern.toLowerCase().includes(q))
    );
  });

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/70 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-5xl max-h-[92vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        
        {/* Modal Header */}
        <div className="p-5 bg-linear-to-r from-slate-950 via-indigo-950 to-purple-950 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center space-x-3">
            <div className="p-2.5 rounded-xl bg-purple-500/20 text-purple-300 border border-purple-400/30">
              <Sparkles className="w-6 h-6 text-purple-300" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h3 className="font-bold text-base text-white tracking-tight">
                  Inbuilt AI Analyser & Custom Presets Library
                </h3>
                <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-purple-500/30 text-purple-200 border border-purple-400/30">
                  Industrial Standard
                </span>
              </div>
              <p className="text-xs text-slate-300 mt-0.5">
                Automatically detects archetypes, filters merged year rows, and permanently persists mapping rules to Server Disk & Supabase PostgreSQL.
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-white/10 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Switcher & Status Messages */}
        <div className="bg-slate-50 border-b border-slate-200 px-6 py-2.5 flex flex-wrap items-center justify-between gap-3">
          <div className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-200/70 text-xs font-semibold">
            <button
              type="button"
              onClick={() => setActiveTab('LIBRARY')}
              className={`px-4 py-1.5 rounded-md flex items-center space-x-2 transition-all ${
                activeTab === 'LIBRARY'
                  ? 'bg-white text-purple-900 shadow-2xs font-bold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <BookOpen className="w-3.5 h-3.5" />
              <span>Presets Library ({presets.length})</span>
            </button>
            <button
              type="button"
              onClick={() => {
                setActiveTab('CREATE_AI');
                if (!previewPreset && activeMappings.length > 0) {
                  initFromActiveMappings();
                }
              }}
              className={`px-4 py-1.5 rounded-md flex items-center space-x-2 transition-all ${
                activeTab === 'CREATE_AI'
                  ? 'bg-white text-purple-900 shadow-2xs font-bold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Sparkles className="w-3.5 h-3.5 text-purple-600" />
              <span>Create / AI-Generate Custom Preset</span>
            </button>
          </div>

          <div className="flex items-center space-x-3 text-xs text-slate-500">
            <div className="flex items-center space-x-1">
              <Server className="w-3.5 h-3.5 text-slate-600" />
              <span>Server Disk: <strong className="text-slate-700 font-mono">data/presets.json</strong></span>
            </div>
            <span>•</span>
            <div className="flex items-center space-x-1">
              <Database className="w-3.5 h-3.5 text-teal-600" />
              <span>Supabase: <strong className="text-teal-700 font-mono">public.ai_presets</strong></span>
            </div>
          </div>
        </div>

        {statusMessage && (
          <div className={`px-6 py-3 border-b text-xs flex items-center justify-between ${
            statusMessage.type === 'success' 
              ? 'bg-emerald-50 border-emerald-200 text-emerald-800' 
              : statusMessage.type === 'error'
              ? 'bg-rose-50 border-rose-200 text-rose-800'
              : 'bg-indigo-50 border-indigo-200 text-indigo-800'
          }`}>
            <div className="flex items-center space-x-2">
              {statusMessage.type === 'success' ? (
                <Check className="w-4 h-4 text-emerald-600 shrink-0" />
              ) : (
                <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
              )}
              <span>{statusMessage.text}</span>
            </div>
            <button onClick={() => setStatusMessage(null)} className="text-xs opacity-60 hover:opacity-100">
              Dismiss
            </button>
          </div>
        )}

        {/* Modal Content */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">

          {/* TAB 1: PRESETS LIBRARY */}
          {activeTab === 'LIBRARY' && (
            <div className="space-y-5">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div className="space-y-1">
                  <h4 className="font-bold text-slate-900 text-sm flex items-center space-x-2">
                    <span>Industrial Presets & Custom AI Blueprints</span>
                    <span className="text-xs font-normal text-slate-500">
                      (All presets exclude merged year rows and preserve relational schema integrity)
                    </span>
                  </h4>
                  {/* Category Pills */}
                  <div className="flex items-center space-x-1.5 pt-1">
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('ALL')}
                      className={`px-2.5 py-1 rounded-md text-xs font-semibold transition-all ${
                        categoryFilter === 'ALL'
                          ? 'bg-slate-900 text-white shadow-2xs'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      All ({presets.length})
                    </button>
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('STUDENT_BATCHES')}
                      className={`px-2.5 py-1 rounded-md text-xs font-semibold flex items-center space-x-1 transition-all ${
                        categoryFilter === 'STUDENT_BATCHES'
                          ? 'bg-purple-600 text-white shadow-2xs'
                          : 'bg-purple-50 text-purple-700 hover:bg-purple-100 border border-purple-200'
                      }`}
                    >
                      <GraduationCap className="w-3.5 h-3.5" />
                      <span>JHC Student Batches ({studentBatchPresets.length} Combined Tables)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('TEMPLATES')}
                      className={`px-2.5 py-1 rounded-md text-xs font-semibold transition-all ${
                        categoryFilter === 'TEMPLATES'
                          ? 'bg-slate-900 text-white shadow-2xs'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      System Templates ({presets.length - studentBatchPresets.length})
                    </button>
                  </div>
                </div>

                <div className="flex items-center space-x-2">
                  <input
                    type="text"
                    value={filterQuery}
                    onChange={(e) => setFilterQuery(e.target.value)}
                    placeholder="Search presets by name, archetype, tag, file..."
                    className="px-3 py-1.5 rounded-lg border border-slate-300 text-xs w-64 focus:outline-hidden focus:ring-2 focus:ring-purple-500"
                  />
                  <button
                    type="button"
                    onClick={loadPresets}
                    disabled={isLoading}
                    className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 transition-colors"
                    title="Refresh presets from server and Supabase"
                  >
                    <RefreshCw className={`w-4 h-4 ${isLoading ? 'animate-spin' : ''}`} />
                  </button>
                </div>
              </div>

              {/* JHC Student Batch Presets Hero Card */}
              {studentBatchPresets.length > 0 && (categoryFilter === 'ALL' || categoryFilter === 'STUDENT_BATCHES') && (
                <div className="p-4 rounded-xl bg-gradient-to-r from-purple-50 via-indigo-50 to-sky-50 border border-purple-200 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div className="space-y-1">
                    <div className="flex items-center space-x-2">
                      <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-purple-600 text-white shadow-2xs flex items-center space-x-1">
                        <GraduationCap className="w-3 h-3" />
                        <span>8 Combined Book Tables</span>
                      </span>
                      <h4 className="font-bold text-slate-900 text-sm">
                        Jaffna Hindu College Student Roster Presets (Batches 2026 – 2034)
                      </h4>
                    </div>
                    <p className="text-xs text-slate-600 max-w-2xl leading-relaxed">
                      Each workbook combines all its division sheets (12A-12G, Arts, Tech, Commerce, etc.) into a single unified table (<code className="text-purple-700 font-mono text-[11px] bg-purple-100/60 px-1 py-0.5 rounded">students_2026</code> through <code className="text-purple-700 font-mono text-[11px] bg-purple-100/60 px-1 py-0.5 rounded">students_2034</code>) with merged year header rows automatically excluded.
                    </p>
                  </div>
                  <div className="flex items-center space-x-2 shrink-0">
                    <button
                      type="button"
                      onClick={() => {
                        if (onApplyAllPresets) {
                          onApplyAllPresets(studentBatchPresets);
                        } else {
                          studentBatchPresets.forEach(p => onApplyPreset(p));
                        }
                        setStatusMessage({
                          type: 'success',
                          text: `Applied all ${studentBatchPresets.length} live student batch presets! Worksheets across the live workbooks are mapped to combined tables.`
                        });
                      }}
                      className="px-4 py-2.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition-colors flex items-center space-x-1.5"
                    >
                      <Sparkles className="w-4 h-4 text-purple-200" />
                      <span>Apply All {studentBatchPresets.length} Combined Tables</span>
                    </button>
                  </div>
                </div>
              )}

              {/* Grid of Presets */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {filteredPresets.map((preset) => (
                  <div 
                    key={preset.id}
                    className="p-4 rounded-xl border border-slate-200 bg-linear-to-b from-white to-slate-50 flex flex-col justify-between hover:shadow-md transition-all relative group"
                  >
                    <div className="space-y-3">
                      <div className="flex items-start justify-between gap-2">
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-50 border border-indigo-200 text-indigo-800">
                          {preset.badge || preset.archetype}
                        </span>
                        <div className="flex items-center space-x-1">
                          {preset.isSystem ? (
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-700">
                              Built-in
                            </span>
                          ) : (
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">
                              Custom AI
                            </span>
                          )}
                          {!preset.isSystem && (
                            <button
                              type="button"
                              onClick={() => handleDeletePreset(preset.id, preset.name)}
                              className="text-slate-400 hover:text-rose-600 p-1 transition-colors"
                              title="Delete custom preset from server and Supabase"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      </div>

                      <div>
                        <h5 className="font-bold text-slate-900 text-sm">{preset.name}</h5>
                        {preset.filenamePattern && (
                          <div className="text-[11px] font-mono text-indigo-700 font-medium mt-0.5 flex items-center space-x-1">
                            <span>Workbook:</span>
                            <span className="bg-indigo-50 px-1.5 py-0.2 rounded border border-indigo-100">{preset.filenamePattern}</span>
                          </div>
                        )}
                        <p className="text-xs text-slate-600 mt-1 line-clamp-3 leading-relaxed">
                          {preset.description}
                        </p>
                      </div>

                      {/* Specs */}
                      <div className="pt-2 border-t border-slate-100 text-[11px] text-slate-500 space-y-1">
                        {preset.sheetMappings && preset.sheetMappings.length > 0 && preset.sheetMappings[0].supabaseTable && (
                          <div className="flex items-center justify-between">
                            <span>Destination Table:</span>
                            <strong className="text-teal-800 font-mono font-semibold">
                              public.{preset.sheetMappings[0].supabaseTable}
                            </strong>
                          </div>
                        )}
                        <div className="flex items-center justify-between">
                          <span>Worksheet Mappings:</span>
                          <strong className="text-slate-800">{preset.sheetMappings?.length || 0} sheets (Combined)</strong>
                        </div>
                        <div className="flex items-center justify-between">
                          <span>Merged Year Rows:</span>
                          <span className="text-emerald-700 font-semibold flex items-center space-x-0.5">
                            <ShieldCheck className="w-3 h-3 text-emerald-600" />
                            <span>Auto-Excluded from Data</span>
                          </span>
                        </div>
                        {preset.unpivotConfig?.enabled && (
                          <div className="flex items-center justify-between text-amber-700 font-medium">
                            <span>Matrix Unpivot:</span>
                            <span>⚡ Active</span>
                          </div>
                        )}
                      </div>

                      {/* Tags */}
                      {preset.tags && preset.tags.length > 0 && (
                        <div className="flex flex-wrap gap-1 pt-1">
                          {preset.tags.slice(0, 3).map((t, i) => (
                            <span key={i} className="text-[9px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-600">
                              #{t}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* Actions */}
                    <div className="pt-4 mt-3 border-t border-slate-100 flex flex-col gap-2">
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            onApplyPreset(preset);
                            setStatusMessage({
                              type: 'success',
                              text: `Applied preset '${preset.name}'! Configured ${preset.sheetMappings.length} sheet mappings.`
                            });
                          }}
                          className="flex-1 py-1.5 px-2 rounded-lg bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition-colors flex items-center justify-center space-x-1"
                        >
                          <CheckCircle2 className="w-3.5 h-3.5" />
                          <span>Apply Mappings</span>
                        </button>

                        <button
                          type="button"
                          onClick={() => handleCopyPresetJson(preset)}
                          className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs transition-colors"
                          title="Copy Preset JSON to clipboard"
                        >
                          {copiedId === preset.id ? (
                            <Check className="w-3.5 h-3.5 text-emerald-600" />
                          ) : (
                            <Copy className="w-3.5 h-3.5" />
                          )}
                        </button>
                      </div>

                      {preset.isSystem && (
                        <button
                          type="button"
                          onClick={() => {
                            onLoadPresetWorkbook(preset.id as any);
                            onClose();
                          }}
                          className="w-full py-1.5 px-2 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-[11px] font-semibold transition-colors flex items-center justify-center space-x-1"
                        >
                          <BookOpen className="w-3 h-3 text-indigo-600" />
                          <span>Load Sample Workbook</span>
                        </button>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 2: CREATE / AI-GENERATE CUSTOM PRESET */}
          {activeTab === 'CREATE_AI' && (
            <div className="space-y-6">
              <div className="bg-purple-50/60 rounded-xl p-5 border border-purple-200 space-y-3">
                <div className="flex items-center space-x-2">
                  <Sparkles className="w-5 h-5 text-purple-600" />
                  <h4 className="font-bold text-slate-900 text-sm">
                    Inbuilt AI Preset Analyser Engine
                  </h4>
                </div>
                <p className="text-xs text-slate-600 leading-relaxed max-w-3xl">
                  The AI Analyser scans all sheets, cell coordinates, and merged ranges. It identifies tables, separates multi-row year banners (e.g. <em>"2023"</em>, <em>"Year 2023"</em>) from real student/transaction rows, and generates ready-to-save relational mappings for Supabase.
                </p>

                <div className="pt-2 flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    disabled={isGeneratingAi || !currentAnalysis}
                    onClick={handleGenerateAiPreset}
                    className="inline-flex items-center space-x-2 px-4 py-2 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-md transition-colors disabled:opacity-50"
                  >
                    {isGeneratingAi ? (
                      <RefreshCw className="w-4 h-4 animate-spin text-white" />
                    ) : (
                      <Sparkles className="w-4 h-4" />
                    )}
                    <span>Run AI Analyser on Current Workbook</span>
                  </button>

                  {activeMappings.length > 0 && (
                    <button
                      type="button"
                      onClick={initFromActiveMappings}
                      className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold shadow-2xs transition-colors"
                    >
                      <Layers className="w-3.5 h-3.5 text-slate-500" />
                      <span>Use Active Mappings ({activeMappings.length} sheets)</span>
                    </button>
                  )}
                </div>
              </div>

              {/* Preset Configuration Form */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700">Preset Name</label>
                  <input
                    type="text"
                    value={customName}
                    onChange={(e) => setCustomName(e.target.value)}
                    placeholder="e.g. School Donations Financial Ledger"
                    className="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs font-medium focus:ring-2 focus:ring-purple-500 focus:outline-hidden"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700">Detected / Target Archetype</label>
                  <select
                    value={customArchetype}
                    onChange={(e) => setCustomArchetype(e.target.value as WorkbookArchetype)}
                    className="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs font-medium focus:ring-2 focus:ring-purple-500 focus:outline-hidden bg-white"
                  >
                    <option value="STANDARD_TABULAR">Standard Tabular Records</option>
                    <option value="TIMETABLE_MATRIX">Timetable Matrix (2-cell slots: Subject + Teacher)</option>
                    <option value="MULTI_SHEET_LEDGER">Multi-Sheet Ledger (Financial with Year Dividers)</option>
                    <option value="PIVOT_ALLOCATION_MATRIX">Pivot Allocation Matrix (Staff Teacher Grid)</option>
                  </select>
                </div>

                <div className="space-y-1.5 md:col-span-2">
                  <label className="text-xs font-bold text-slate-700">Description</label>
                  <input
                    type="text"
                    value={customDesc}
                    onChange={(e) => setCustomDesc(e.target.value)}
                    placeholder="Describe how sheets are mapped and structured..."
                    className="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs focus:ring-2 focus:ring-purple-500 focus:outline-hidden"
                  />
                </div>

                <div className="space-y-1.5 md:col-span-2">
                  <label className="text-xs font-bold text-slate-700">Tags (comma-separated)</label>
                  <input
                    type="text"
                    value={customTags}
                    onChange={(e) => setCustomTags(e.target.value)}
                    placeholder="e.g. School, Donations, Financial, 2023"
                    className="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs focus:ring-2 focus:ring-purple-500 focus:outline-hidden"
                  />
                </div>
              </div>

              {/* Industrial Standards Checklist Banner */}
              <div className="p-4 rounded-xl border border-emerald-200 bg-emerald-50/60 space-y-2 text-xs">
                <div className="font-bold text-emerald-950 flex items-center space-x-1.5">
                  <ShieldCheck className="w-4 h-4 text-emerald-600" />
                  <span>Industrial Standard Enforcement:</span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-emerald-800 text-[11px]">
                  <div className="flex items-center space-x-1.5">
                    <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>Merged year rows flagged as dividers (NEVER inserted as data)</span>
                  </div>
                  <div className="flex items-center space-x-1.5">
                    <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>Permanent persistence in Server Disk (<code className="font-mono">presets.json</code>)</span>
                  </div>
                  <div className="flex items-center space-x-1.5">
                    <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>Synchronized to Supabase (<code className="font-mono">ai_presets</code> & <code className="font-mono">worksheet_mappings</code>)</span>
                  </div>
                </div>
              </div>

              {/* Preview of Mappings in this Preset */}
              {previewPreset && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between">
                    <h5 className="font-bold text-slate-900 text-xs">
                      Preset Sheet Mappings Preview ({previewPreset.sheetMappings.length} Worksheets)
                    </h5>
                    <span className="text-[11px] text-slate-500">
                      All sheets have <strong className="text-emerald-700">skipMergedYearRows: true</strong>
                    </span>
                  </div>

                  <div className="border border-slate-200 rounded-xl overflow-hidden max-h-56 overflow-y-auto">
                    <table className="min-w-full divide-y divide-slate-200 text-xs">
                      <thead className="bg-slate-50 text-slate-700 font-semibold sticky top-0">
                        <tr>
                          <th className="px-3 py-2 text-left">Worksheet</th>
                          <th className="px-3 py-2 text-left">Target Supabase Table</th>
                          <th className="px-3 py-2 text-center">Header / Data Row</th>
                          <th className="px-3 py-2 text-center">Columns</th>
                          <th className="px-3 py-2 text-right">Merged Year Rows</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 bg-white font-mono">
                        {previewPreset.sheetMappings.map((sm, idx) => (
                          <tr key={idx} className="hover:bg-slate-50">
                            <td className="px-3 py-2 font-medium text-slate-900">{sm.worksheetName}</td>
                            <td className="px-3 py-2 text-purple-700">{sm.supabaseTable}</td>
                            <td className="px-3 py-2 text-center text-slate-600">Row {sm.headerRow} / Row {sm.dataStartRow}</td>
                            <td className="px-3 py-2 text-center text-slate-800 font-bold">{sm.columns?.length || 0} cols</td>
                            <td className="px-3 py-2 text-right text-emerald-700 font-sans font-semibold">
                              ✓ Filtered Out
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Save Button */}
              <div className="flex items-center justify-end space-x-3 pt-2">
                <button
                  type="button"
                  onClick={() => setActiveTab('LIBRARY')}
                  className="px-4 py-2 rounded-xl border border-slate-300 text-slate-700 hover:bg-slate-100 text-xs font-semibold transition-colors"
                >
                  Cancel
                </button>

                <button
                  type="button"
                  disabled={isSavingPreset || !previewPreset}
                  onClick={handleSavePermanentPreset}
                  className="px-5 py-2.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-md transition-colors flex items-center space-x-2 disabled:opacity-50"
                >
                  {isSavingPreset ? (
                    <>
                      <RefreshCw className="w-4 h-4 animate-spin text-white" />
                      <span>Permanently Saving to Server & Supabase...</span>
                    </>
                  ) : (
                    <>
                      <Check className="w-4 h-4" />
                      <span>Permanently Save Preset (Server Disk & Supabase)</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          )}

        </div>

        {/* Modal Footer */}
        <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between shrink-0">
          <div className="text-[11px] text-slate-500 flex items-center space-x-2">
            <ShieldCheck className="w-4 h-4 text-emerald-600" />
            <span>Industrial standard compliance: zero year divider pollution in SQL inserts</span>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold transition-colors"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};
