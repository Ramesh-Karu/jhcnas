import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Microscope, 
  Layers, 
  Table, 
  Upload, 
  Sparkles, 
  Check, 
  AlertCircle, 
  FileSpreadsheet, 
  ChevronRight, 
  Sliders, 
  BookOpen, 
  Search,
  Database,
  Code,
  Copy,
  CheckCheck,
  RefreshCw,
  ExternalLink,
  Info,
  Key,
  ShieldCheck,
  Wand2,
  FolderSync,
  Bot,
  Brain,
  Zap,
  CheckCircle2,
  ArrowRight,
  HelpCircle,
  X,
  Globe,
  Calendar,
  DollarSign,
  GraduationCap,
  Users,
  Scissors,
  FileCheck
} from 'lucide-react';
import { 
  WorkbookAnalysis, 
  SheetAnalysis, 
  NavigationTab, 
  WorksheetMapping, 
  NextcloudConfig,
  NextcloudFile,
  SupabaseConfig,
  SupabaseTableInfo,
  MultiSheetConsolidationMode,
  TableSchemaPlan,
  AiWorkbookAnalysisResult,
  AiSheetMappingSolution,
  WorkbookArchetype,
  AiWorkbookPreset
} from '../types';
import { ExcelAnalyzer } from '../services/excelAnalyzer';
import { MatrixTransformer } from '../services/matrixTransformer';
import { 
  createComplexSampleWorkbook, 
  createSchoolCustomLayoutWorkbook, 
  generateLayoutStructureExport 
} from '../services/sampleWorkbook';
import { ApiClient } from '../services/apiClient';
import { SchemaGenerator } from '../services/schemaGenerator';
import { GeminiWorkbookService } from '../services/geminiService';
import { AiPresetsModal } from './AiPresetsModal';
import { smartSanitizeIdentifier } from '../services/tamilTranslator';

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
  supabaseConfig?: SupabaseConfig;
  onSaveMappings?: (mappings: WorksheetMapping[]) => void;
  files?: NextcloudFile[];
  onSelectFileForAnalysis?: (file: NextcloudFile) => void;
}

export const WorkbookAnalyzerView: React.FC<WorkbookAnalyzerViewProps> = ({
  currentAnalysis,
  onAnalysisUpdate,
  onNavigate,
  onOpenAiAssistant,
  activeMappings,
  onUpdateMappingHeaderDataRow,
  nextcloudConfig,
  supabaseConfig,
  onSaveMappings,
  files,
  onSelectFileForAnalysis
}) => {
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [sheetSearch, setSheetSearch] = useState<string>('');
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [isFetchingNextcloud, setIsFetchingNextcloud] = useState<boolean>(false);
  const [selectedNextcloudFilename, setSelectedNextcloudFilename] = useState<string>('');
  const [saveSuccessNotice, setSaveSuccessNotice] = useState<string | null>(null);
  const [showPasteModal, setShowPasteModal] = useState<boolean>(false);
  const [pastedContent, setPastedContent] = useState<string>('');
  const [targetTableNameInput, setTargetTableNameInput] = useState<string>('');

  // Main View Toggle: Spreadsheet Analysis vs Database Schema Generator
  const [viewMode, setViewMode] = useState<'SPREADSHEET' | 'SCHEMA_DDL'>('SPREADSHEET');

  // Multi-Sheet Consolidation & Routing Strategy
  const [consolidationMode, setConsolidationMode] = useState<MultiSheetConsolidationMode>(
    currentAnalysis?.consolidationMode || 'SEPARATE_TABLES'
  );
  const [sheetToTableMap, setSheetToTableMap] = useState<Record<string, string>>(
    currentAnalysis?.sheetToTableMap || {}
  );
  const [unifiedTableNameInput, setUnifiedTableNameInput] = useState<string>('consolidated_records');

  // Live Supabase Tables for Schema Comparison
  const [supabaseTables, setSupabaseTables] = useState<SupabaseTableInfo[]>([]);
  const [isLoadingSupabase, setIsLoadingSupabase] = useState<boolean>(false);

  // Schema Plan Inspector State
  const [selectedPlanIndex, setSelectedPlanIndex] = useState<number>(0);
  const [copiedSql, setCopiedSql] = useState<boolean>(false);
  const [isDeployingDdl, setIsDeployingDdl] = useState<boolean>(false);
  const [ddlFeedback, setDdlFeedback] = useState<{ type: 'success' | 'info' | 'error'; message: string } | null>(null);

  // AI Workbook Analyzer State
  const [isAiAnalyzingFullWorkbook, setIsAiAnalyzingFullWorkbook] = useState<boolean>(false);
  const [aiAnalysisResult, setAiAnalysisResult] = useState<AiWorkbookAnalysisResult | null>(null);
  const [showAiResultModal, setShowAiResultModal] = useState<boolean>(false);
  const [selectedAiSheetIdx, setSelectedAiSheetIdx] = useState<number>(0);
  const [isApplyingAi, setIsApplyingAi] = useState<boolean>(false);
  const [showLayoutExportModal, setShowLayoutExportModal] = useState<boolean>(false);
  const [hasCopiedExportText, setHasCopiedExportText] = useState<boolean>(false);

  // Google Sheets & Presets Modal State
  const [showGoogleSheetsModal, setShowGoogleSheetsModal] = useState<boolean>(false);
  const [showAiPresetsModal, setShowAiPresetsModal] = useState<boolean>(false);
  const [googleSheetUrlInput, setGoogleSheetUrlInput] = useState<string>('');
  const [isFetchingGoogleSheet, setIsFetchingGoogleSheet] = useState<boolean>(false);
  const [googleSheetError, setGoogleSheetError] = useState<string | null>(null);

  // Matrix Unpivoting & Normalization State
  const [isNormalizedView, setIsNormalizedView] = useState<boolean>(false);
  const [normalizedRecordsPreview, setNormalizedRecordsPreview] = useState<any[] | null>(null);
  const [showNormalizationModal, setShowNormalizationModal] = useState<boolean>(false);

  // Real workbook analysis (strictly prioritizes real live Nextcloud files over sample)
  const analysis = useMemo(() => {
    const liveNames = new Set((files || []).map(f => f.filename.toLowerCase().trim()));
    // 1. Real analysis with parsed worksheets
    if (currentAnalysis && currentAnalysis.worksheets && currentAnalysis.worksheets.length > 0) {
      if (currentAnalysis.filename !== 'students_complex.xlsx' || !files || files.length === 0) {
        return currentAnalysis;
      }
    }
    // 2. Real filename selected from Nextcloud
    if (currentAnalysis && currentAnalysis.filename && currentAnalysis.filename !== 'students_complex.xlsx' && currentAnalysis.filename !== 'students.xlsx') {
      return currentAnalysis;
    }
    // 3. Fallback to first live file in Nextcloud
    if (files && files.length > 0) {
      return {
        filename: files[0].filename,
        fileHash: files[0].fileHash,
        fileSize: files[0].fileSize,
        fileSizeFormatted: files[0].fileSizeFormatted,
        totalWorksheets: 0,
        worksheets: [],
        analyzedAt: new Date().toISOString()
      };
    }
    // 4. Sample fallback only if zero files exist
    const sample = createComplexSampleWorkbook();
    const { analysis: parsed } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
    parsed.fileHash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
    return parsed;
  }, [currentAnalysis, files]);

  // Keep selectedNextcloudFilename in sync with currentAnalysis
  useEffect(() => {
    if (currentAnalysis?.filename && currentAnalysis.filename !== 'students_complex.xlsx') {
      setSelectedNextcloudFilename(currentAnalysis.filename);
    }
  }, [currentAnalysis?.filename]);

  // If no workbook is analyzed yet, automatically fetch the target live Nextcloud file
  useEffect(() => {
    if ((!currentAnalysis || !currentAnalysis.worksheets || currentAnalysis.worksheets.length === 0) && files && files.length > 0 && !isFetchingNextcloud) {
      const liveNames = new Set(files.map(f => f.filename.toLowerCase().trim()));
      let target = '';
      if (currentAnalysis?.filename && currentAnalysis.filename !== 'students_complex.xlsx' && liveNames.has(currentAnalysis.filename.toLowerCase().trim())) {
        target = currentAnalysis.filename;
      } else if (selectedNextcloudFilename && selectedNextcloudFilename !== 'students_complex.xlsx' && liveNames.has(selectedNextcloudFilename.toLowerCase().trim())) {
        target = selectedNextcloudFilename;
      } else {
        target = files[0].filename;
      }
      if (target) {
        handleFetchFromNextcloud(target);
      }
    }
  }, [files, currentAnalysis?.filename]);

  // Load Live Supabase Schema for comparison
  useEffect(() => {
    if (!supabaseConfig?.url) return;
    setIsLoadingSupabase(true);
    ApiClient.getSupabaseSchema(supabaseConfig)
      .then(res => {
        if (res.success && res.tables) {
          setSupabaseTables(res.tables);
        }
      })
      .catch(console.error)
      .finally(() => setIsLoadingSupabase(false));
  }, [supabaseConfig]);

  const filteredWorksheetIndices = useMemo(() => {
    return analysis.worksheets
      .map((ws, idx) => ({ ws, idx }))
      .filter(({ ws }) => ws.sheetName.toLowerCase().includes(sheetSearch.toLowerCase()));
  }, [analysis.worksheets, sheetSearch]);

  const activeSheet: SheetAnalysis = analysis.worksheets[selectedSheetIndex] || analysis.worksheets[0];

  // Manual Header Row / Data Start Row
  const [manualHeaderRow, setManualHeaderRow] = useState<number>(activeSheet?.detectedHeaderRow || 1);
  const [manualDataStartRow, setManualDataStartRow] = useState<number>(activeSheet?.detectedDataStartRow || 2);

  // Sync state when active sheet changes
  useEffect(() => {
    if (activeSheet) {
      setManualHeaderRow(activeSheet.detectedHeaderRow);
      setManualDataStartRow(activeSheet.detectedDataStartRow);
      const safeName = smartSanitizeIdentifier(activeSheet.sheetName, 'sheet_records');
      setTargetTableNameInput(safeName);
    }
  }, [activeSheet?.sheetName]);

  // Generate Database Schema Plans using SchemaGenerator service
  const effectiveSheetToTableMap = useMemo(() => {
    const map = { ...sheetToTableMap };
    if (consolidationMode === 'UNIFIED_TABLE') {
      map['__unified__'] = smartSanitizeIdentifier(unifiedTableNameInput, 'consolidated_records');
    }
    return map;
  }, [sheetToTableMap, consolidationMode, unifiedTableNameInput]);

  // Custom Primary Key Selection State
  const [customPrimaryKeys, setCustomPrimaryKeys] = useState<Record<string, string>>({});

  const handleSetTablePrimaryKey = (tableName: string, colName: string) => {
    setCustomPrimaryKeys(prev => ({
      ...prev,
      [tableName]: colName
    }));
  };

  const schemaPlans = useMemo(() => {
    return SchemaGenerator.generateSchemas(
      analysis,
      consolidationMode,
      effectiveSheetToTableMap,
      supabaseTables,
      customPrimaryKeys
    );
  }, [analysis, consolidationMode, effectiveSheetToTableMap, supabaseTables, customPrimaryKeys]);

  const activePlan: TableSchemaPlan | undefined = schemaPlans[selectedPlanIndex] || schemaPlans[0];

  // Combined SQL for all tables in the workbook
  const completeCombinedSql = useMemo(() => {
    return schemaPlans.map(p => p.completeSql).join('\n\n');
  }, [schemaPlans]);

  // Upload Excel Handler
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    try {
      const buffer = await file.arrayBuffer();
      const sha256 = await ExcelAnalyzer.computeSHA256(buffer);
      const { analysis: parsed } = ExcelAnalyzer.parseBuffer(buffer, file.name);
      parsed.fileHash = sha256;
      parsed.consolidationMode = consolidationMode;

      // Automatically generate initial schemas and mappings
      const initialPlans = SchemaGenerator.generateSchemas(parsed, consolidationMode, effectiveSheetToTableMap, supabaseTables);
      const autoMappings = SchemaGenerator.generateMappingsFromPlans(parsed, initialPlans, consolidationMode, effectiveSheetToTableMap);

      onAnalysisUpdate(parsed);
      if (onSaveMappings) {
        onSaveMappings(autoMappings);
      }

      setSelectedSheetIndex(0);
      setSelectedPlanIndex(0);
      setSaveSuccessNotice(`Successfully loaded '${file.name}' (${parsed.worksheets.length} sheets). Real columns and data types automatically detected and mapped!`);
      setTimeout(() => setSaveSuccessNotice(null), 5000);
    } catch (err: any) {
      console.error('Error parsing uploaded Excel file:', err);
      alert('Error parsing Excel file: ' + err.message);
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
        const initialPlans = SchemaGenerator.generateSchemas(res.analysis, consolidationMode, effectiveSheetToTableMap, supabaseTables);
        const autoMappings = SchemaGenerator.generateMappingsFromPlans(res.analysis, initialPlans, consolidationMode, effectiveSheetToTableMap);

        onAnalysisUpdate(res.analysis);
        if (onSaveMappings) {
          onSaveMappings(autoMappings);
        }

        setSelectedSheetIndex(0);
        setSelectedPlanIndex(0);
        setShowPasteModal(false);
        setPastedContent('');
        setSaveSuccessNotice(`Successfully extracted columns and auto-inferred data types from pasted table!`);
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

  // Re-scan active sheet when user manually changes Header Row
  const handleReScanSheet = () => {
    if (analysis.base64Data) {
      try {
        const binStr = atob(analysis.base64Data);
        const len = binStr.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binStr.charCodeAt(i);
        }
        const wb = XLSX.read(bytes, { type: 'array', cellDates: true });
        const ws = wb.Sheets[activeSheet.sheetName] || wb.Sheets[wb.SheetNames[selectedSheetIndex]];
        if (ws) {
          const freshSheet = ExcelAnalyzer.analyzeSheet(ws, activeSheet.sheetName, manualHeaderRow, manualDataStartRow);
          const updatedWorksheets = [...analysis.worksheets];
          updatedWorksheets[selectedSheetIndex] = freshSheet;
          const updatedAnalysis = { ...analysis, worksheets: updatedWorksheets };
          onAnalysisUpdate(updatedAnalysis);
          onUpdateMappingHeaderDataRow(activeSheet.sheetName, manualHeaderRow, manualDataStartRow);
          setSaveSuccessNotice(`Re-scanned '${activeSheet.sheetName}' from Row ${manualHeaderRow}! Found ${freshSheet.headers.length} columns.`);
          setTimeout(() => setSaveSuccessNotice(null), 3500);
          return;
        }
      } catch (e) {
        console.error('Error re-scanning sheet:', e);
      }
    }

    // Fallback: update rows without binary reload
    const updatedWorksheets = [...analysis.worksheets];
    updatedWorksheets[selectedSheetIndex] = {
      ...activeSheet,
      detectedHeaderRow: manualHeaderRow,
      detectedDataStartRow: manualDataStartRow
    };
    onAnalysisUpdate({ ...analysis, worksheets: updatedWorksheets });
    onUpdateMappingHeaderDataRow(activeSheet.sheetName, manualHeaderRow, manualDataStartRow);
    setSaveSuccessNotice(`Updated header (row ${manualHeaderRow}) and data start (row ${manualDataStartRow}) for '${activeSheet.sheetName}'.`);
    setTimeout(() => setSaveSuccessNotice(null), 3000);
  };

  // Route active sheet to a PostgreSQL table
  const handleRouteSheetToTable = () => {
    if (!activeSheet) return;
    const defaultTblName = activeSheet.sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/^_+|_+$/g, '') || 'table_data';
    const targetTable = smartSanitizeIdentifier(targetTableNameInput, defaultTblName);

    const columns = activeSheet.headers.map((h, idx) => {
      const cleanName = SchemaGenerator.sanitizeIdentifier(h.name);
      const isUnique = h.isCandidateKey || idx === 0;
      const dataType = h.inferredType || 'text';

      return {
        id: `cm-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
        excelColumn: h.colLetter,
        excelHeader: h.name,
        supabaseColumn: cleanName,
        dataType,
        required: isUnique,
        uniqueKey: isUnique,
        transformation: SchemaGenerator.dataTypeToTransformation(dataType),
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

    const currentWb = (analysis.filename || '').toLowerCase().trim();
    const existingIdx = activeMappings.findIndex(m => 
      (m.workbookName || '').toLowerCase().trim() === currentWb &&
      m.worksheetName.toLowerCase().trim() === activeSheet.sheetName.toLowerCase().trim()
    );
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

    setSaveSuccessNotice(`Routed '${activeSheet.sheetName}' to PostgreSQL table 'public.${targetTable}' with ${columns.length} columns!`);
    setTimeout(() => {
      setSaveSuccessNotice(null);
      onNavigate('mappings');
    }, 1200);
  };

  // Apply All Schemas to Mappings (1-click sync of multi-sheet routing)
  const handleApplyAllSchemasToMappings = () => {
    const newMappings = SchemaGenerator.generateMappingsFromPlans(
      analysis,
      schemaPlans,
      consolidationMode,
      effectiveSheetToTableMap
    );

    if (onSaveMappings) {
      onSaveMappings(newMappings);
    }

    setSaveSuccessNotice(`Applied ${consolidationMode.replace('_', ' ')} strategy: configured ${newMappings.length} worksheet mappings!`);
    setTimeout(() => {
      setSaveSuccessNotice(null);
      onNavigate('mappings');
    }, 1500);
  };

  // Copy SQL Script
  const handleCopySql = (sqlText: string) => {
    navigator.clipboard.writeText(sqlText);
    setCopiedSql(true);
    setTimeout(() => setCopiedSql(false), 2500);
  };

  // Deploy DDL to Supabase
  const handleDeployDdl = async (plan?: TableSchemaPlan) => {
    if (!supabaseConfig?.url) {
      setDdlFeedback({
        type: 'error',
        message: 'Supabase URL is not configured. Please enter your Supabase credentials in the Supabase tab.'
      });
      return;
    }

    const sqlToRun = plan ? plan.completeSql : completeCombinedSql;
    const tableName = plan ? plan.tableName : 'all_tables';

    setIsDeployingDdl(true);
    setDdlFeedback(null);

    try {
      const res = await ApiClient.executeSupabaseDdl(supabaseConfig, sqlToRun, tableName);
      if (res.success) {
        if (res.directExecuted) {
          setDdlFeedback({
            type: 'success',
            message: res.message || `Table 'public.${tableName}' successfully created in Supabase!`
          });
          // Refresh tables
          const refreshed = await ApiClient.getSupabaseSchema(supabaseConfig);
          if (refreshed.success && refreshed.tables) {
            setSupabaseTables(refreshed.tables);
          }
        } else {
          setDdlFeedback({
            type: 'info',
            message: res.message || 'SQL ready! Click "Copy SQL" below and run it in the Supabase SQL Editor.'
          });
        }
      } else {
        setDdlFeedback({
          type: 'error',
          message: res.error || 'Failed to execute DDL on Supabase'
        });
      }
    } catch (e: any) {
      setDdlFeedback({
        type: 'error',
        message: e.message || 'Network error executing DDL'
      });
    } finally {
      setIsDeployingDdl(false);
    }
  };

  const handleFetchFromNextcloud = async (customFilename?: string) => {
    if (!nextcloudConfig) return;
    setIsFetchingNextcloud(true);
    try {
      const liveNames = new Set((files || []).map(f => f.filename.toLowerCase().trim()));
      let targetFilename = customFilename || '';
      if (!targetFilename && selectedNextcloudFilename && selectedNextcloudFilename !== 'students_complex.xlsx' && liveNames.has(selectedNextcloudFilename.toLowerCase().trim())) {
        targetFilename = selectedNextcloudFilename;
      }
      if (!targetFilename && currentAnalysis?.filename && currentAnalysis.filename !== 'students_complex.xlsx' && liveNames.has(currentAnalysis.filename.toLowerCase().trim())) {
        targetFilename = currentAnalysis.filename;
      }
      if (!targetFilename && files && files.length > 0) {
        targetFilename = files[0].filename;
      }
      if (!targetFilename) {
        setIsFetchingNextcloud(false);
        return;
      }
      setSelectedNextcloudFilename(targetFilename);
      const targetFile = files?.find(f => f.filename.toLowerCase().trim() === targetFilename.toLowerCase().trim());
      const filePath = targetFile?.path;

      const res = await ApiClient.fetchAndParseWorkbook(nextcloudConfig, filePath, targetFilename);
      if (res.success && res.analysis) {
        const initialPlans = SchemaGenerator.generateSchemas(res.analysis, consolidationMode, effectiveSheetToTableMap, supabaseTables);
        const autoMappings = SchemaGenerator.generateMappingsFromPlans(res.analysis, initialPlans, consolidationMode, effectiveSheetToTableMap);

        onAnalysisUpdate(res.analysis);
        if (onSaveMappings) {
          onSaveMappings(autoMappings);
        }

        setSelectedSheetIndex(0);
        const totalRowsAll = res.analysis.worksheets.reduce((acc, ws) => acc + (ws.totalRows || 0), 0);
        setSaveSuccessNotice(`Successfully fetched ${res.analysis.filename} (${totalRowsAll.toLocaleString()} rows across ${res.analysis.worksheets.length} sheets) directly from Nextcloud WebDAV!`);
        setTimeout(() => setSaveSuccessNotice(null), 6000);
      } else {
        setSaveSuccessNotice(`Notice fetching ${targetFilename}: ${res.error || 'Check WebDAV configuration'}`);
        setTimeout(() => setSaveSuccessNotice(null), 8000);
      }
    } catch (e: any) {
      setSaveSuccessNotice(`Error fetching from Nextcloud: ${e.message}`);
      setTimeout(() => setSaveSuccessNotice(null), 8000);
    } finally {
      setIsFetchingNextcloud(false);
    }
  };

  // Full AI Workbook Analysis Trigger
  const handleRunFullAiWorkbookAnalysis = async () => {
    setIsAiAnalyzingFullWorkbook(true);
    try {
      const result = await GeminiWorkbookService.analyzeFullWorkbook(analysis, supabaseTables);
      setAiAnalysisResult(result);
      setSelectedAiSheetIdx(0);
      setShowAiResultModal(true);
    } catch (err: any) {
      console.error('AI Workbook Analysis failed:', err);
      alert('AI Workbook Analysis error: ' + (err.message || 'Unknown error'));
    } finally {
      setIsAiAnalyzingFullWorkbook(false);
    }
  };

  // Apply ALL AI Generated Mappings across all sheets
  const handleApplyAllAiSolutions = async () => {
    if (!aiAnalysisResult) return;
    setIsApplyingAi(true);
    try {
      const newMappings = GeminiWorkbookService.convertSolutionsToMappings(aiAnalysisResult, analysis.filename);

      if (onSaveMappings) {
        onSaveMappings(newMappings);
      }

      // Persist permanently to backend disk & Supabase metadata
      await ApiClient.savePermanentMappings({
        mappings: newMappings,
        workbookInfo: {
          filename: analysis.filename,
          fileHash: analysis.fileHash,
          totalWorksheets: analysis.totalWorksheets,
        },
        supabase: supabaseConfig,
      });

      setShowAiResultModal(false);
      setSaveSuccessNotice(
        `✨ AI Solution successfully applied! Configured and permanently saved ${newMappings.length} sheet mappings with intelligent primary keys and transformations.`
      );
      setTimeout(() => setSaveSuccessNotice(null), 6000);
    } catch (e: any) {
      alert('Failed to save AI mappings: ' + e.message);
    } finally {
      setIsApplyingAi(false);
    }
  };

  // Apply Single Selected Sheet AI Solution
  const handleApplySingleAiSolution = (solution: AiSheetMappingSolution) => {
    const singleMapping: WorksheetMapping = {
      id: `wm-ai-${Date.now()}`,
      workbookName: analysis.filename,
      worksheetName: solution.worksheetName,
      supabaseTable: solution.suggestedTable,
      headerRow: solution.headerRow,
      dataStartRow: solution.dataStartRow,
      dataEndRow: solution.dataEndRow,
      sectionHeadingTargetCol: solution.sectionHeadingTargetCol,
      enabled: true,
      syncPolicy: solution.syncPolicy || 'BIDIRECTIONAL',
      columns: solution.columns.map((c, idx) => ({
        id: c.id || `col-ai-${idx}-${Date.now()}`,
        excelColumn: c.excelColumn,
        excelHeader: c.excelHeader,
        supabaseColumn: c.supabaseColumn,
        dataType: c.dataType,
        required: c.required,
        uniqueKey: c.uniqueKey,
        defaultValue: c.defaultValue,
        transformation: c.transformation,
        validationRegex: c.validationRegex,
      })),
    };

    const currentWb = (analysis.filename || '').toLowerCase().trim();
    const existingIdx = activeMappings.findIndex(
      m => (m.workbookName || '').toLowerCase().trim() === currentWb &&
           m.worksheetName.toLowerCase().trim() === solution.worksheetName.toLowerCase().trim()
    );
    let updated: WorksheetMapping[];
    if (existingIdx >= 0) {
      updated = [...activeMappings];
      updated[existingIdx] = singleMapping;
    } else {
      updated = [...activeMappings, singleMapping];
    }

    if (onSaveMappings) {
      onSaveMappings(updated);
    }

    // Persist permanently
    ApiClient.savePermanentMappings({
      mappings: updated,
      workbookInfo: {
        filename: analysis.filename,
        fileHash: analysis.fileHash,
        totalWorksheets: analysis.totalWorksheets,
      },
      supabase: supabaseConfig,
    }).catch(console.warn);

    setShowAiResultModal(false);
    setSaveSuccessNotice(
      `✨ AI Mapping for '${solution.worksheetName}' applied to table 'public.${solution.suggestedTable}'!`
    );
    setTimeout(() => setSaveSuccessNotice(null), 5000);
  };

  const handleLoadSchoolSample = () => {
    const sample = createSchoolCustomLayoutWorkbook();
    const { analysis: parsed } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
    parsed.fileHash = 'school_custom_sample_sha256';
    onAnalysisUpdate(parsed);
    setSelectedSheetIndex(0);
    setSaveSuccessNotice(
      `Loaded sample '${sample.filename}'! Includes Grade 6, 7, 8 sheets, Timetable Matrix (2-cell slots: Subject + Teacher/Room), and Year 2020.`
    );
    setTimeout(() => setSaveSuccessNotice(null), 6000);
  };

  const handleCopyLayoutBlueprint = () => {
    const exportText = generateLayoutStructureExport(analysis);
    navigator.clipboard.writeText(exportText);
    setHasCopiedExportText(true);
    setTimeout(() => setHasCopiedExportText(false), 3000);
  };

  const handleLoadPreset = async (presetId: 'timetable' | 'donations' | 'teacher_allocations' | 'jhc_inventory' | 'preset-jhc-inventory' | string) => {
    setIsFetchingGoogleSheet(true);
    setGoogleSheetError(null);
    try {
      const res = await ApiClient.loadPresetWorkbook(presetId);
      if (res.success && res.analysis) {
        onAnalysisUpdate(res.analysis);
        setSelectedSheetIndex(0);
        setShowGoogleSheetsModal(false);
        setIsNormalizedView(false);
        setNormalizedRecordsPreview(null);
        setSaveSuccessNotice(`✨ Loaded preset '${res.analysis.filename}'! Detected: ${res.analysis.archetypeTitle || res.analysis.detectedArchetype}`);
        setTimeout(() => setSaveSuccessNotice(null), 6000);
      } else {
        setGoogleSheetError(res.error || 'Failed to load preset workbook');
      }
    } catch (err: any) {
      setGoogleSheetError(err.message || 'Error loading preset');
    } finally {
      setIsFetchingGoogleSheet(false);
    }
  };

  const handleApplyPreset = (preset: AiWorkbookPreset) => {
    if (preset.sheetMappings && preset.sheetMappings.length > 0) {
      if (onSaveMappings) {
        onSaveMappings(preset.sheetMappings);
      }
      ApiClient.savePermanentMappings({
        mappings: preset.sheetMappings,
        workbookInfo: {
          filename: preset.filenamePattern || analysis.filename,
          fileHash: analysis.fileHash,
          totalWorksheets: preset.sheetMappings.length,
        },
        supabase: supabaseConfig,
      }).catch(console.warn);

      setSaveSuccessNotice(
        `✨ Applied Preset '${preset.name}'! (${preset.sheetMappings.length} worksheets configured into combined table public.${preset.sheetMappings[0]?.supabaseTable}).`
      );
      setTimeout(() => setSaveSuccessNotice(null), 5000);
    }
  };

  const handleApplyAllStudentPresets = (studentPresets: AiWorkbookPreset[]) => {
    const allSheetMappings: WorksheetMapping[] = [];
    studentPresets.forEach(p => {
      if (p.sheetMappings) {
        allSheetMappings.push(...p.sheetMappings);
      }
    });
    if (allSheetMappings.length > 0) {
      if (onSaveMappings) {
        onSaveMappings(allSheetMappings);
      }
      ApiClient.savePermanentMappings({
        mappings: allSheetMappings,
        supabase: supabaseConfig,
      }).catch(console.warn);

      setSaveSuccessNotice(
        `🎉 Successfully applied all 8 JHC Student Batch Presets! (${allSheetMappings.length} worksheets configured into combined tables).`
      );
      setTimeout(() => setSaveSuccessNotice(null), 5000);
    }
  };

  const handleFetchGoogleSheetUrl = async (customUrl?: string) => {
    const targetUrl = (customUrl || googleSheetUrlInput).trim();
    if (!targetUrl) {
      setGoogleSheetError('Please enter a Google Sheets URL or Spreadsheet ID.');
      return;
    }
    setIsFetchingGoogleSheet(true);
    setGoogleSheetError(null);
    try {
      const res = await ApiClient.fetchGoogleSheet(targetUrl);
      if (res.success && res.analysis) {
        onAnalysisUpdate(res.analysis);
        setSelectedSheetIndex(0);
        setShowGoogleSheetsModal(false);
        setIsNormalizedView(false);
        setNormalizedRecordsPreview(null);
        setSaveSuccessNotice(`✨ Successfully fetched & analyzed '${res.analysis.filename}'! Detected: ${res.analysis.archetypeTitle || res.analysis.detectedArchetype}`);
        setTimeout(() => setSaveSuccessNotice(null), 6000);
      } else {
        setGoogleSheetError(res.error || 'Failed to download Google Sheet. Verify the sheet is public ("Anyone with the link can view").');
      }
    } catch (err: any) {
      setGoogleSheetError(err.message || 'Error fetching Google Sheet');
    } finally {
      setIsFetchingGoogleSheet(false);
    }
  };

  const handleNormalizeMatrix = () => {
    if (!analysis.base64Data && !analysis.rawWorkbookBase64) return;
    try {
      const b64 = analysis.rawWorkbookBase64 || analysis.base64Data || '';
      const binary = atob(b64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const rawWb = XLSX.read(bytes, { type: 'array', cellDates: true });

      if (analysis.detectedArchetype === 'TIMETABLE_MATRIX') {
        const { analysis: normalizedAnalysis, records } = MatrixTransformer.createNormalizedTimetableAnalysis(rawWb, analysis.filename);
        const fullNormalized: WorkbookAnalysis = {
          ...analysis,
          worksheets: [normalizedAnalysis],
          totalWorksheets: 1,
          isNormalizedMatrix: true,
          rawWorkbookBase64: b64,
        };
        onAnalysisUpdate(fullNormalized);
        setSelectedSheetIndex(0);
        setNormalizedRecordsPreview(records.slice(0, 20));
        setIsNormalizedView(true);
        setSaveSuccessNotice(`✨ Timetable Matrix Unpivoted! Transformed into ${records.length} clean relational records in 'school_timetables'.`);
        setTimeout(() => setSaveSuccessNotice(null), 6000);
      } else if (analysis.detectedArchetype === 'PIVOT_ALLOCATION_MATRIX') {
        const { analysis: normalizedAnalysis, records } = MatrixTransformer.createNormalizedStaffAllocationAnalysis(rawWb, analysis.filename);
        const fullNormalized: WorkbookAnalysis = {
          ...analysis,
          worksheets: [normalizedAnalysis],
          totalWorksheets: 1,
          isNormalizedMatrix: true,
          rawWorkbookBase64: b64,
        };
        onAnalysisUpdate(fullNormalized);
        setSelectedSheetIndex(0);
        setNormalizedRecordsPreview(records.slice(0, 20));
        setIsNormalizedView(true);
        setSaveSuccessNotice(`✨ Staff Pivot Matrix Unpivoted! Transformed into ${records.length} clean relational records in 'teacher_subject_assignments'.`);
        setTimeout(() => setSaveSuccessNotice(null), 6000);
      }
    } catch (err: any) {
      console.error('Normalization failed:', err);
    }
  };

  const handleRestoreRawMatrix = () => {
    const b64 = analysis.rawWorkbookBase64 || analysis.base64Data;
    if (!b64) return;
    try {
      const binary = atob(b64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const { analysis: restored } = ExcelAnalyzer.parseBuffer(bytes, analysis.filename);
      onAnalysisUpdate(restored);
      setSelectedSheetIndex(0);
      setIsNormalizedView(false);
      setNormalizedRecordsPreview(null);
      setSaveSuccessNotice(`↩️ Restored raw multi-sheet matrix view (${restored.totalWorksheets} sheets).`);
      setTimeout(() => setSaveSuccessNotice(null), 5000);
    } catch (err: any) {
      console.error('Restore failed:', err);
    }
  };

  return (
    <div className="space-y-6">
      {/* Top Banner with File Info, View Mode Switcher, and Upload Controls */}
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

        {/* View Mode Toggle & Upload Actions */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Main View Mode Switcher */}
          <div className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-100 text-xs font-medium">
            <button
              onClick={() => setViewMode('SPREADSHEET')}
              className={`px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition-all ${
                viewMode === 'SPREADSHEET'
                  ? 'bg-white text-purple-700 shadow-2xs font-semibold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Workbook & Columns</span>
            </button>
            <button
              id="btn-view-schema-ddl"
              onClick={() => setViewMode('SCHEMA_DDL')}
              className={`px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition-all ${
                viewMode === 'SCHEMA_DDL'
                  ? 'bg-white text-purple-700 shadow-2xs font-semibold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Database className="w-3.5 h-3.5" />
              <span>Database Schema & Supabase DDL</span>
              <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-purple-100 text-purple-800 font-bold">
                {schemaPlans.length} tbl
              </span>
            </button>
          </div>

          {nextcloudConfig && (
            <div className="inline-flex items-center space-x-1.5">
              {files && files.length > 1 && (
                <select
                  value={selectedNextcloudFilename || currentAnalysis?.filename || files[0]?.filename}
                  onChange={(e) => {
                    const fn = e.target.value;
                    setSelectedNextcloudFilename(fn);
                    handleFetchFromNextcloud(fn);
                  }}
                  disabled={isFetchingNextcloud}
                  className="px-2 py-1.5 rounded-lg border border-emerald-300 bg-white text-emerald-900 text-xs font-semibold shadow-2xs focus:ring-1 focus:ring-emerald-500"
                  title="Choose Nextcloud file to fetch & analyze"
                >
                  {files.map(f => (
                    <option key={f.id} value={f.filename}>
                      {f.filename} ({f.fileSizeFormatted})
                    </option>
                  ))}
                </select>
              )}
              <button
                id="btn-fetch-nextcloud-real"
                onClick={() => handleFetchFromNextcloud()}
                disabled={isFetchingNextcloud}
                className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-700 hover:bg-emerald-100 text-xs font-medium transition-colors disabled:opacity-50"
                title={`Download & parse '${selectedNextcloudFilename || currentAnalysis?.filename || 'monitored file'}' directly from Nextcloud WebDAV`}
              >
                <RefreshCw className={`w-3.5 h-3.5 text-emerald-600 ${isFetchingNextcloud ? 'animate-spin' : ''}`} />
                <span>{isFetchingNextcloud ? 'Downloading...' : `Fetch Nextcloud (${selectedNextcloudFilename || currentAnalysis?.filename || 'File'})`}</span>
              </button>
            </div>
          )}

          <button
            id="btn-open-google-sheets-modal"
            onClick={() => setShowGoogleSheetsModal(true)}
            className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-sky-300 bg-sky-50 text-xs font-semibold text-sky-800 hover:bg-sky-100 transition-colors shadow-2xs"
            title="Load Google Sheets URL or load: Timetable Matrix, Donations Ledger, Staff Allocations"
          >
            <Globe className="w-3.5 h-3.5 text-sky-600" />
            <span>Google Sheets & Presets</span>
          </button>

          <button
            id="btn-open-ai-presets-modal"
            onClick={() => setShowAiPresetsModal(true)}
            className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-purple-300 bg-purple-50 text-xs font-semibold text-purple-900 hover:bg-purple-100 transition-colors shadow-2xs"
            title="Open AI Presets Library & Inbuilt Analyser (Permanent Server & Supabase Storage)"
          >
            <Sparkles className="w-3.5 h-3.5 text-purple-600" />
            <span>AI Presets & Analyser</span>
          </button>

          <button
            id="btn-load-school-sample"
            onClick={handleLoadSchoolSample}
            className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-purple-200 bg-purple-50 text-xs font-semibold text-purple-700 hover:bg-purple-100 transition-colors shadow-2xs"
            title="Load realistic school multi-sheet workbook with Grade 6..8, Timetable matrix (2-cell slots), and Year 2020"
          >
            <BookOpen className="w-3.5 h-3.5 text-purple-600" />
            <span>Load Grades & Timetable Sample</span>
          </button>

          <button
            id="btn-copy-layout-chat"
            onClick={() => setShowLayoutExportModal(true)}
            className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-indigo-200 bg-indigo-50 text-xs font-semibold text-indigo-700 hover:bg-indigo-100 transition-colors shadow-2xs"
            title="Open sharing helper and copy structured layout blueprint for the AI chat"
          >
            <Copy className="w-3.5 h-3.5 text-indigo-600" />
            <span>Copy Layout for Chat</span>
          </button>

          <button
            id="btn-paste-raw-table"
            onClick={() => setShowPasteModal(true)}
            className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 transition-colors"
          >
            <Table className="w-3.5 h-3.5 text-slate-500" />
            <span>Paste CSV / Table</span>
          </button>

          <label 
            id="btn-upload-excel"
            className="cursor-pointer inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg bg-purple-600 hover:bg-purple-700 text-white text-xs font-medium transition-colors"
          >
            <Upload className="w-3.5 h-3.5" />
            <span>{isUploading ? 'Parsing Columns...' : 'Upload Excel (.xlsx/.csv)'}</span>
            <input 
              type="file" 
              accept=".xlsx,.xlsm,.xls,.csv" 
              onChange={handleFileUpload} 
              className="hidden" 
            />
          </label>
        </div>
      </div>

      {saveSuccessNotice && (
        <div className="p-3.5 rounded-xl bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-medium flex items-center space-x-2">
          <Check className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{saveSuccessNotice}</span>
        </div>
      )}

      {/* ARCHETYPE INTELLIGENCE BANNER */}
      {analysis.detectedArchetype && analysis.detectedArchetype !== 'STANDARD_TABULAR' && (
        <div className="bg-linear-to-r from-slate-900 via-indigo-950 to-purple-950 rounded-2xl p-5 border border-indigo-500/30 text-white shadow-lg space-y-4">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="space-y-1">
              <div className="flex items-center space-x-2">
                <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-indigo-500/30 border border-indigo-400/40 text-indigo-200">
                  {analysis.archetypeBadge || 'Architecture Detected'}
                </span>
                <span className="text-sm font-bold text-white tracking-tight">
                  {analysis.archetypeTitle}
                </span>
              </div>
              <p className="text-xs text-slate-300 max-w-3xl leading-relaxed">
                {analysis.archetypeSummary}
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {analysis.detectedArchetype === 'TIMETABLE_MATRIX' && (
                analysis.isNormalizedMatrix ? (
                  <div className="flex items-center space-x-2">
                    <span className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-lg bg-emerald-500/20 border border-emerald-400/40 text-emerald-200 text-xs font-bold">
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                      <span>Matrix Normalized (Relational View Active)</span>
                    </span>
                    <button
                      onClick={handleRestoreRawMatrix}
                      className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 border border-slate-600 text-xs font-medium text-slate-200 transition-colors"
                    >
                      ↩️ View Raw 5-Day Matrix
                    </button>
                  </div>
                ) : (
                  <button
                    id="btn-normalize-timetable-matrix"
                    onClick={handleNormalizeMatrix}
                    className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-xl bg-linear-to-r from-amber-500 to-orange-600 hover:from-amber-600 hover:to-orange-700 text-white font-bold text-xs shadow-md transition-all animate-pulse"
                  >
                    <Zap className="w-4 h-4 text-white" />
                    <span>⚡ 1-Click Normalize Matrix (Unpivot to 2,482 Timetable Slots)</span>
                  </button>
                )
              )}

              {analysis.detectedArchetype === 'PIVOT_ALLOCATION_MATRIX' && (
                analysis.isNormalizedMatrix ? (
                  <div className="flex items-center space-x-2">
                    <span className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-lg bg-emerald-500/20 border border-emerald-400/40 text-emerald-200 text-xs font-bold">
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                      <span>Pivot Unpivoted (Consolidated Table Active)</span>
                    </span>
                    <button
                      onClick={handleRestoreRawMatrix}
                      className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 border border-slate-600 text-xs font-medium text-slate-200 transition-colors"
                    >
                      ↩️ View Raw Grade Sheets
                    </button>
                  </div>
                ) : (
                  <button
                    id="btn-normalize-staff-matrix"
                    onClick={handleNormalizeMatrix}
                    className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-xl bg-linear-to-r from-emerald-500 to-teal-600 hover:from-emerald-600 hover:to-teal-700 text-white font-bold text-xs shadow-md transition-all animate-pulse"
                  >
                    <Zap className="w-4 h-4 text-white" />
                    <span>⚡ 1-Click Normalize Matrix (Consolidate 8 Grades into 729 Records)</span>
                  </button>
                )
              )}

              {analysis.detectedArchetype === 'MULTI_SHEET_LEDGER' && (
                <div className="inline-flex items-center space-x-1 px-3 py-1.5 rounded-lg bg-emerald-500/20 border border-emerald-400/30 text-emerald-300 text-xs font-medium">
                  <CheckCheck className="w-3.5 h-3.5 text-emerald-400" />
                  <span>Section Divider Rows Filtered & Currencies Sanitized</span>
                </div>
              )}
            </div>
          </div>

          {analysis.archetypeFeatures && analysis.archetypeFeatures.length > 0 && (
            <div className="pt-2 border-t border-slate-800/80 flex flex-wrap gap-2 text-[11px] text-slate-300">
              {analysis.archetypeFeatures.map((feat, fIdx) => (
                <span key={fIdx} className="px-2.5 py-1 rounded-md bg-slate-800/60 border border-slate-700/60 flex items-center space-x-1.5">
                  <span className="w-1.5 h-1.5 rounded-full bg-indigo-400" />
                  <span>{feat}</span>
                </span>
              ))}
            </div>
          )}
        </div>
      )}

      {/* AI WORKBOOK ANALYZER HERO BANNER */}
      <div className="relative overflow-hidden bg-linear-to-r from-purple-900 via-indigo-900 to-slate-900 rounded-2xl p-6 text-white shadow-xl border border-purple-500/20">
        <div className="absolute top-0 right-0 -mt-10 -mr-10 w-64 h-64 bg-purple-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="absolute bottom-0 left-1/3 -mb-10 w-48 h-48 bg-indigo-500/10 rounded-full blur-2xl pointer-events-none" />

        <div className="relative z-10 flex flex-col lg:flex-row lg:items-center justify-between gap-6">
          <div className="space-y-2 max-w-2xl">
            <div className="inline-flex items-center space-x-2 px-3 py-1 rounded-full bg-purple-500/20 border border-purple-400/30 text-purple-200 text-xs font-medium backdrop-blur-xs">
              <Sparkles className="w-3.5 h-3.5 text-purple-300 animate-pulse" />
              <span>AI Workbook Analyzer & Auto-Mapper</span>
              <span className="bg-purple-400/20 text-purple-200 text-[10px] px-2 py-0.2 rounded-full font-mono">
                Gemini 3.8 Flash
              </span>
            </div>
            <h2 className="text-xl lg:text-2xl font-black text-white tracking-tight">
              Auto-Solve Sheet Mappings & PostgreSQL Schemas
            </h2>
            <p className="text-xs lg:text-sm text-purple-100/80 leading-relaxed">
              Scan all <strong className="text-white font-semibold">{analysis.totalWorksheets} sheets</strong> in <span className="font-mono text-purple-200">{analysis.filename}</span>. 
              The AI automatically bypasses decorative title banners, extracts merged section values (e.g. class / division), assigns primary upsert keys, infers PostgreSQL data types, and suggests clean table routes.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-3 shrink-0">
            <button
              id="btn-ai-analyze-full-workbook"
              onClick={handleRunFullAiWorkbookAnalysis}
              disabled={isAiAnalyzingFullWorkbook}
              className="inline-flex items-center space-x-2.5 px-5 py-3 rounded-xl bg-linear-to-r from-purple-500 to-indigo-600 hover:from-purple-400 hover:to-indigo-500 text-white font-bold text-xs shadow-lg shadow-purple-900/50 hover:shadow-purple-700/50 transition-all duration-200 transform hover:-translate-y-0.5 disabled:opacity-50"
            >
              <Bot className={`w-4 h-4 text-purple-200 ${isAiAnalyzingFullWorkbook ? 'animate-bounce' : ''}`} />
              <span>{isAiAnalyzingFullWorkbook ? 'Analyzing All Sheets...' : '✨ Run AI Full Workbook Analyzer'}</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>

            <button
              onClick={() => onOpenAiAssistant(activeSheet)}
              className="inline-flex items-center space-x-2 px-4 py-3 rounded-xl bg-white/10 hover:bg-white/15 border border-white/20 text-white text-xs font-semibold backdrop-blur-xs transition-colors"
              title="Analyze active sheet with AI Assistant"
            >
              <Brain className="w-4 h-4 text-purple-300" />
              <span>Analyze Active Sheet</span>
            </button>
          </div>
        </div>

        {/* AI Stats Row */}
        <div className="relative z-10 mt-6 pt-5 border-t border-white/10 grid grid-cols-2 sm:grid-cols-4 gap-4 text-xs">
          <div>
            <span className="text-purple-300/80 block text-[11px]">Worksheets Ready</span>
            <span className="font-bold text-white text-sm">{analysis.totalWorksheets} Sheets</span>
          </div>
          <div>
            <span className="text-purple-300/80 block text-[11px]">Columns Scanned</span>
            <span className="font-bold text-white text-sm">
              {analysis.worksheets.reduce((acc, ws) => acc + ws.totalColumns, 0)} Detected
            </span>
          </div>
          <div>
            <span className="text-purple-300/80 block text-[11px]">Total Data Rows</span>
            <span className="font-bold text-white text-sm">
              {analysis.worksheets.reduce((acc, ws) => acc + ws.totalRows, 0).toLocaleString()} Rows
            </span>
          </div>
          <div>
            <span className="text-purple-300/80 block text-[11px]">Active Destination Tables</span>
            <span className="font-bold text-white text-sm">{activeMappings.length} Configured</span>
          </div>
        </div>
      </div>

      {/* MULTI-SHEET ROUTING & CONSOLIDATION STRATEGY BAR */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div className="flex items-center space-x-2.5">
            <div className="p-1.5 rounded-lg bg-blue-50 text-blue-600">
              <FolderSync className="w-4 h-4" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">Multi-Sheet Database Routing & Consolidation</h2>
              <p className="text-xs text-slate-500">
                Configure how the {analysis.worksheets.length} sheets in this workbook route to Supabase tables
              </p>
            </div>
          </div>

          <button
            onClick={handleApplyAllSchemasToMappings}
            className="inline-flex items-center space-x-1.5 px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors"
          >
            <Check className="w-3.5 h-3.5" />
            <span>Apply Routing to App Mappings</span>
          </button>
        </div>

        {/* 3 Strategy Modes */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          {/* Mode 1: Unified Common Table */}
          <div
            onClick={() => setConsolidationMode('UNIFIED_TABLE')}
            className={`cursor-pointer p-3.5 rounded-xl border-2 transition-all ${
              consolidationMode === 'UNIFIED_TABLE'
                ? 'border-purple-600 bg-purple-50/50 shadow-2xs ring-2 ring-purple-100'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <input
                  type="radio"
                  name="consolidationMode"
                  checked={consolidationMode === 'UNIFIED_TABLE'}
                  onChange={() => setConsolidationMode('UNIFIED_TABLE')}
                  className="text-purple-600 focus:ring-purple-500"
                />
                <span className="text-xs font-bold text-slate-900">1 Common Table (Merge All)</span>
              </div>
              <span className="text-[10px] px-1.5 py-0.5 rounded font-mono bg-purple-100 text-purple-800 font-semibold">
                N sheets → 1 table
              </span>
            </div>
            <p className="text-[11px] text-slate-500 mt-2">
              All {analysis.worksheets.length} sheets combine into a single common schema with an automatic <code>_sheet_source</code> column.
            </p>
            {consolidationMode === 'UNIFIED_TABLE' && (
              <div className="mt-3 pt-2 border-t border-purple-200/60" onClick={(e) => e.stopPropagation()}>
                <label className="block text-[10px] font-semibold text-purple-900 uppercase">Target Table Name</label>
                <input
                  type="text"
                  value={unifiedTableNameInput}
                  onChange={(e) => setUnifiedTableNameInput(e.target.value)}
                  placeholder="e.g. students"
                  className="mt-1 w-full px-2.5 py-1 text-xs font-mono rounded border border-purple-300 bg-white text-purple-950 focus:outline-none focus:ring-1 focus:ring-purple-500"
                />
              </div>
            )}
          </div>

          {/* Mode 2: Separate Tables */}
          <div
            onClick={() => setConsolidationMode('SEPARATE_TABLES')}
            className={`cursor-pointer p-3.5 rounded-xl border-2 transition-all ${
              consolidationMode === 'SEPARATE_TABLES'
                ? 'border-purple-600 bg-purple-50/50 shadow-2xs ring-2 ring-purple-100'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <input
                  type="radio"
                  name="consolidationMode"
                  checked={consolidationMode === 'SEPARATE_TABLES'}
                  onChange={() => setConsolidationMode('SEPARATE_TABLES')}
                  className="text-purple-600 focus:ring-purple-500"
                />
                <span className="text-xs font-bold text-slate-900">Separate Tables (1:1)</span>
              </div>
              <span className="text-[10px] px-1.5 py-0.5 rounded font-mono bg-blue-100 text-blue-800 font-semibold">
                1 sheet = 1 table
              </span>
            </div>
            <p className="text-[11px] text-slate-500 mt-2">
              Each worksheet generates its own dedicated table in Supabase (e.g. <code>students</code>, <code>attendance</code>, <code>grades</code>).
            </p>
          </div>

          {/* Mode 3: Custom Grouping */}
          <div
            onClick={() => setConsolidationMode('CUSTOM_GROUPING')}
            className={`cursor-pointer p-3.5 rounded-xl border-2 transition-all ${
              consolidationMode === 'CUSTOM_GROUPING'
                ? 'border-purple-600 bg-purple-50/50 shadow-2xs ring-2 ring-purple-100'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <input
                  type="radio"
                  name="consolidationMode"
                  checked={consolidationMode === 'CUSTOM_GROUPING'}
                  onChange={() => setConsolidationMode('CUSTOM_GROUPING')}
                  className="text-purple-600 focus:ring-purple-500"
                />
                <span className="text-xs font-bold text-slate-900">Custom Grouping (Hybrid)</span>
              </div>
              <span className="text-[10px] px-1.5 py-0.5 rounded font-mono bg-emerald-100 text-emerald-800 font-semibold">
                Flexible mapping
              </span>
            </div>
            <p className="text-[11px] text-slate-500 mt-2">
              Assign specific sheets to the same table (e.g. 2 sheets into <code>students</code>, 1 sheet into <code>grades</code>).
            </p>
          </div>
        </div>

        {/* Custom Grouping Sheet Selectors */}
        {consolidationMode === 'CUSTOM_GROUPING' && (
          <div className="pt-3 border-t border-slate-100">
            <h4 className="text-xs font-bold text-slate-700 mb-2 uppercase">Sheet-to-Table Routing Assignments:</h4>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2.5">
              {analysis.worksheets.map((ws, wsIdx) => {
                const currentDest = sheetToTableMap[ws.sheetName] || smartSanitizeIdentifier(ws.sheetName, 'sheet_data');
                return (
                  <div key={`${ws.sheetName}-${wsIdx}`} className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs flex items-center justify-between gap-2">
                    <div className="truncate font-semibold text-slate-800" title={ws.sheetName}>
                      {ws.sheetName} ({ws.totalRows}r)
                    </div>
                    <div className="flex items-center space-x-1 shrink-0">
                      <span className="text-slate-400">→</span>
                      <input
                        type="text"
                        value={currentDest}
                        onChange={(e) => {
                          const val = e.target.value;
                          setSheetToTableMap(prev => ({ ...prev, [ws.sheetName]: val }));
                        }}
                        placeholder="table_name"
                        className="w-28 px-2 py-0.5 text-xs font-mono rounded border border-slate-300 bg-white text-slate-900 focus:outline-none focus:ring-1 focus:ring-purple-500"
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* VIEW MODE 1: DATABASE SCHEMA & SUPABASE DDL VIEWER */}
      {viewMode === 'SCHEMA_DDL' ? (
        <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs space-y-6">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-200 pb-4">
            <div>
              <div className="flex items-center space-x-2">
                <Database className="w-5 h-5 text-purple-600" />
                <h3 className="text-lg font-bold text-slate-900">PostgreSQL Schema & Supabase DDL Migration</h3>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Automatically created from Excel sheets. Inferred data types: UUIDs, Integers, Numerics, Dates, Booleans, and Text.
              </p>
            </div>

            {/* Actions: Copy SQL & Deploy */}
            <div className="flex items-center space-x-2">
              <button
                onClick={() => handleCopySql(activePlan ? activePlan.completeSql : completeCombinedSql)}
                className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 shadow-2xs transition-colors"
              >
                {copiedSql ? <CheckCheck className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5 text-slate-500" />}
                <span>{copiedSql ? 'Copied to Clipboard!' : 'Copy SQL'}</span>
              </button>

              <button
                id="btn-deploy-supabase-ddl"
                onClick={() => handleDeployDdl(activePlan)}
                disabled={isDeployingDdl}
                className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors disabled:opacity-50"
              >
                <Database className="w-3.5 h-3.5" />
                <span>{isDeployingDdl ? 'Deploying to Supabase...' : `Deploy 'public.${activePlan?.tableName}' to Supabase`}</span>
              </button>
            </div>
          </div>

          {/* DDL Execution Feedback Notice */}
          {ddlFeedback && (
            <div className={`p-4 rounded-xl text-xs font-medium flex items-start space-x-2.5 ${
              ddlFeedback.type === 'success' ? 'bg-emerald-50 text-emerald-900 border border-emerald-200' :
              ddlFeedback.type === 'info' ? 'bg-blue-50 text-blue-900 border border-blue-200' :
              'bg-red-50 text-red-900 border border-red-200'
            }`}>
              <Info className="w-4 h-4 shrink-0 mt-0.5" />
              <div className="flex-1">
                <span>{ddlFeedback.message}</span>
                {ddlFeedback.type === 'info' && (
                  <div className="mt-2">
                    <a
                      href="https://supabase.com/dashboard"
                      target="_blank"
                      rel="noreferrer"
                      className="inline-flex items-center space-x-1 text-blue-700 hover:text-blue-900 underline font-semibold"
                    >
                      <span>Open Supabase Dashboard SQL Editor</span>
                      <ExternalLink className="w-3 h-3" />
                    </a>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Table Selector Tabs */}
          <div className="flex border-b border-slate-200 overflow-x-auto space-x-2 pb-0.5">
            {schemaPlans.map((plan, idx) => (
              <button
                key={`${plan.tableName}-${idx}`}
                onClick={() => setSelectedPlanIndex(idx)}
                className={`px-4 py-2.5 rounded-t-lg text-xs font-medium whitespace-nowrap transition-all border-b-2 flex items-center space-x-2 ${
                  selectedPlanIndex === idx
                    ? 'border-purple-600 text-purple-700 bg-purple-50/50 font-bold'
                    : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                <span>public.{plan.tableName}</span>
                <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-slate-200 text-slate-700 font-mono">
                  {plan.columns.length} cols
                </span>
                <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-medium ${
                  plan.diffStatus === 'EXACT_MATCH' ? 'bg-emerald-100 text-emerald-800' :
                  plan.diffStatus === 'NEEDS_ALTER' ? 'bg-amber-100 text-amber-800' :
                  'bg-blue-100 text-blue-800'
                }`}>
                  {plan.diffStatus === 'EXACT_MATCH' ? 'Live Match' :
                   plan.diffStatus === 'NEEDS_ALTER' ? 'Migration Needed' :
                   'New Table'}
                </span>
              </button>
            ))}
          </div>

          {/* Active Plan Detail & Inferred Columns Table */}
          {activePlan && (
            <div className="space-y-4">
              {/* Summary Bar */}
              <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs">
                <div className="flex flex-wrap items-center gap-3">
                  <div>
                    <span className="text-slate-500">Target Table: </span>
                    <span className="font-mono font-bold text-slate-900">public.{activePlan.tableName}</span>
                  </div>
                  <span className="text-slate-300">|</span>
                  <div>
                    <span className="text-slate-500">Source Sheets: </span>
                    <span className="font-medium text-slate-900">{activePlan.sourceSheetNames.join(', ')}</span>
                  </div>
                  <span className="text-slate-300">|</span>
                  <div>
                    <span className="text-slate-500">Total Columns: </span>
                    <span className="font-bold text-purple-700">{activePlan.columns.length}</span>
                  </div>
                </div>

                <div className="flex items-center space-x-2">
                  <span className="text-slate-500 text-[11px]">Live Supabase Status:</span>
                  <span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${
                    activePlan.diffStatus === 'EXACT_MATCH' ? 'bg-emerald-100 text-emerald-800' :
                    activePlan.diffStatus === 'NEEDS_ALTER' ? 'bg-amber-100 text-amber-800' :
                    'bg-blue-100 text-blue-800'
                  }`}>
                    {activePlan.diffStatus === 'EXACT_MATCH' ? 'Table & Columns Exist in Supabase' :
                     activePlan.diffStatus === 'NEEDS_ALTER' ? `${activePlan.missingInSupabaseColumns.length} Missing Columns to Add` :
                     'New Table (Will be Created)'}
                  </span>
                </div>
              </div>

              {/* Primary Key Selection Bar */}
              <div className="bg-gradient-to-r from-amber-50 to-orange-50 rounded-xl p-3.5 border border-amber-200/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs shadow-2xs">
                <div className="flex items-center space-x-2.5">
                  <div className="p-2 rounded-lg bg-amber-500 text-white shadow-xs shrink-0">
                    <Key className="w-4 h-4" />
                  </div>
                  <div>
                    <div className="font-bold text-amber-950 flex items-center space-x-1.5">
                      <span>Choose Table Primary Key for <code className="bg-amber-100 px-1 py-0.5 rounded text-amber-900 font-mono">public.{activePlan.tableName}</code>:</span>
                    </div>
                    <p className="text-[11px] text-amber-800/90 mt-0.5">
                      Pick your unique column (e.g. <code>admission_no</code>, <code>student_id</code>, <code>roll_no</code>, <code>code</code>, <code>email</code>) rather than a forced ID column.
                    </p>
                  </div>
                </div>

                <div className="flex items-center space-x-2 shrink-0">
                  <select
                    value={activePlan.columns.find(c => c.isPrimary)?.name || '__NONE__'}
                    onChange={(e) => handleSetTablePrimaryKey(activePlan.tableName, e.target.value)}
                    className="px-3 py-1.5 rounded-lg border border-amber-300 bg-white text-xs font-semibold text-slate-900 shadow-2xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                  >
                    <option value="__NONE__">No Primary Key (Direct Insert Only)</option>
                    {activePlan.columns.map(c => (
                      <option key={c.name} value={c.name}>
                        Key: {c.name} ({c.sqlType}) {c.originalHeaders[0] ? `— from "${c.originalHeaders[0]}"` : ''}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Inferred Columns Grid */}
              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                <div className="p-3 bg-slate-100 border-b border-slate-200 text-xs font-bold text-slate-700 flex items-center justify-between">
                  <span>Inferred PostgreSQL Columns & Data Types ({activePlan.columns.length})</span>
                  <span className="text-[11px] text-slate-500 font-normal">Click any "Set as PK" button below to assign your chosen Primary Key</span>
                </div>

                <div className="overflow-x-auto max-h-80">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-50 text-slate-600 font-semibold border-b border-slate-200 sticky top-0">
                      <tr>
                        <th className="px-3.5 py-2">Column Name</th>
                        <th className="px-3.5 py-2">PostgreSQL Data Type</th>
                        <th className="px-3.5 py-2">Primary Key / Constraints</th>
                        <th className="px-3.5 py-2">Source Excel Header</th>
                        <th className="px-3.5 py-2">Sample Values</th>
                        <th className="px-3.5 py-2">Supabase Sync Status</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono">
                      {activePlan.columns.map((c) => (
                        <tr key={c.name} className="hover:bg-slate-50/80 transition-colors">
                          <td className="px-3.5 py-2 font-bold text-slate-900">
                            {c.name}
                          </td>
                          <td className="px-3.5 py-2">
                            <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${
                              c.dataType === 'integer' ? 'bg-blue-100 text-blue-800' :
                              c.dataType === 'decimal' ? 'bg-indigo-100 text-indigo-800' :
                              c.dataType === 'date' ? 'bg-amber-100 text-amber-800' :
                              c.dataType === 'boolean' ? 'bg-emerald-100 text-emerald-800' :
                              'bg-slate-100 text-slate-700'
                            }`}>
                              {c.sqlType}
                            </span>
                          </td>
                          <td className="px-3.5 py-2 text-[11px]">
                            <button
                              type="button"
                              onClick={() => handleSetTablePrimaryKey(activePlan.tableName, c.isPrimary ? '__NONE__' : c.name)}
                              className={`px-2.5 py-1 rounded text-[11px] font-bold flex items-center space-x-1.5 transition-all ${
                                c.isPrimary
                                  ? 'bg-amber-500 text-white shadow-2xs hover:bg-amber-600 ring-2 ring-amber-300'
                                  : 'bg-slate-100 text-slate-600 hover:text-amber-800 hover:bg-amber-50 border border-slate-200'
                              }`}
                              title={c.isPrimary ? 'Click to unset Primary Key' : 'Click to make this column the table PRIMARY KEY'}
                            >
                              <Key className={`w-3 h-3 ${c.isPrimary ? 'text-white' : 'text-slate-400'}`} />
                              <span>{c.isPrimary ? 'PRIMARY KEY' : 'Set as PK'}</span>
                            </button>
                          </td>
                          <td className="px-3.5 py-2 font-sans text-slate-700">
                            {c.originalHeaders.join(', ')}
                          </td>
                          <td className="px-3.5 py-2 text-slate-500 font-sans truncate max-w-xs" title={c.sampleValues.join(', ')}>
                            {c.sampleValues.slice(0, 3).join(', ') || '-'}
                          </td>
                          <td className="px-3.5 py-2">
                            {c.matchesSupabaseColumn ? (
                              <span className="inline-flex items-center space-x-1 text-emerald-700 font-sans font-medium text-[11px]">
                                <Check className="w-3 h-3 text-emerald-600" />
                                <span>Exists in DB</span>
                              </span>
                            ) : (
                              <span className="inline-flex items-center space-x-1 text-blue-700 font-sans font-medium text-[11px]">
                                <Sparkles className="w-3 h-3 text-blue-600" />
                                <span>DDL will create</span>
                              </span>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* SQL Syntax Block */}
              <div className="border border-slate-200 rounded-xl overflow-hidden">
                <div className="p-3 bg-slate-900 text-slate-200 border-b border-slate-800 flex items-center justify-between text-xs">
                  <div className="flex items-center space-x-2">
                    <Code className="w-4 h-4 text-purple-400" />
                    <span className="font-mono font-bold">SQL Migration Script (PostgreSQL / Supabase DDL)</span>
                  </div>
                  <button
                    onClick={() => handleCopySql(activePlan.completeSql)}
                    className="text-slate-400 hover:text-white flex items-center space-x-1 font-sans text-[11px]"
                  >
                    <Copy className="w-3 h-3" />
                    <span>Copy Script</span>
                  </button>
                </div>
                <pre className="p-4 bg-slate-950 text-emerald-400 font-mono text-xs overflow-x-auto max-h-72 leading-relaxed">
                  <code>{activePlan.completeSql}</code>
                </pre>
              </div>
            </div>
          )}
        </div>
      ) : (
        /* VIEW MODE 2: SPREADSHEET ANALYSIS & SAMPLES */
        <>
          {/* Multi-Sheet Workbook Tabs Toolbar */}
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
                      <option key={`${ws.sheetName}-${idx}`} value={idx}>
                        {idx + 1}. {ws.sheetName} ({ws.totalRows}r)
                      </option>
                    ))}
                  </select>
                </>
              )}

              <button
                onClick={() => setViewMode('SCHEMA_DDL')}
                className="px-3 py-1 rounded-lg bg-purple-50 border border-purple-300 text-purple-700 hover:bg-purple-100 text-xs font-medium flex items-center space-x-1"
              >
                <Database className="w-3.5 h-3.5 text-purple-600" />
                <span>Inspect Database Schema</span>
              </button>

              <button
                onClick={() => onNavigate('mappings')}
                className="px-3 py-1 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-700 hover:bg-emerald-100 text-xs font-medium flex items-center space-x-1"
              >
                <span>Manage All Mappings</span>
                <ChevronRight className="w-3 h-3" />
              </button>
            </div>
          </div>

          {/* Worksheet Tabs Navigation & Deep Inspection */}
          {analysis.worksheets.length === 0 || !activeSheet ? (
            <div className="bg-white rounded-xl p-10 border border-slate-200 shadow-2xs text-center space-y-4">
              <div className="w-14 h-14 rounded-2xl bg-purple-50 text-purple-600 flex items-center justify-center mx-auto">
                <FileSpreadsheet className={`w-7 h-7 ${isFetchingNextcloud ? 'animate-pulse text-purple-500' : ''}`} />
              </div>
              <div className="space-y-1">
                <h3 className="text-base font-bold text-slate-900">
                  {isFetchingNextcloud ? `Downloading & Inspecting '${analysis.filename}'...` : `Workbook '${analysis.filename}'`}
                </h3>
                <p className="text-xs text-slate-500 max-w-md mx-auto">
                  {isFetchingNextcloud
                    ? 'Connecting to Nextcloud WebDAV to extract worksheets, columns, merged cells, and data rows...'
                    : 'Worksheet structure is loading from Nextcloud. Click below to inspect.'}
                </p>
              </div>
              <div className="pt-2">
                <button
                  type="button"
                  onClick={() => handleFetchFromNextcloud(analysis.filename)}
                  disabled={isFetchingNextcloud}
                  className="inline-flex items-center space-x-2 px-5 py-2.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-xs transition-colors disabled:opacity-50"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isFetchingNextcloud ? 'animate-spin' : ''}`} />
                  <span>{isFetchingNextcloud ? 'Fetching WebDAV...' : `Fetch & Inspect '${analysis.filename}'`}</span>
                </button>
              </div>
            </div>
          ) : (
            <>
              <div className="flex border-b border-slate-200 overflow-x-auto space-x-2 pb-0.5">
                {filteredWorksheetIndices.map(({ ws, idx }) => (
                  <button
                    key={`${ws.sheetName}-${idx}`}
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
                    {ws.headers.length > 0 && (
                      <span className="text-[10px] px-1.5 py-0.2 rounded bg-purple-100 text-purple-800 font-semibold">
                        {ws.headers.length} cols
                      </span>
                    )}
                  </button>
                ))}
              </div>

              {/* Sheet Deep Inspection Card */}
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

                {/* Manual Correction Controls with Instant Re-scan */}
                <div className="pt-3 border-t border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-slate-900 flex items-center space-x-1">
                      <Sliders className="w-3.5 h-3.5 text-slate-500" />
                      <span>Header Row Tuning</span>
                    </span>
                    <span className="text-[11px] text-slate-400">Re-scan headers</span>
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
                    onClick={handleReScanSheet}
                    className="w-full py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold shadow-2xs transition-colors flex items-center justify-center space-x-1.5"
                  >
                    <RefreshCw className="w-3.5 h-3.5" />
                    <span>Re-Scan Columns from Row {manualHeaderRow}</span>
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
                  Bind <strong>{activeSheet.sheetName}</strong> ({activeSheet.headers.length} detected columns) to a PostgreSQL table:
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
                        className="text-[10px] px-2 py-0.5 rounded border border-slate-200 bg-white hover:bg-slate-100 text-slate-700 font-mono"
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
            </div>

            {/* Right Column: Detected Columns & Sample Preview */}
            <div className="lg:col-span-2 space-y-6">
              {/* Detected Columns & Auto-Inferred Types Grid */}
              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-2xs">
                <div className="flex items-center justify-between mb-3">
                  <h3 className="text-sm font-semibold text-slate-900 flex items-center space-x-2">
                    <FileSpreadsheet className="w-4 h-4 text-purple-600" />
                    <span>Detected Columns ({activeSheet.headers.length}) at Row {activeSheet.detectedHeaderRow}</span>
                  </h3>
                  <span className="text-[11px] text-slate-400">Automated type inference</span>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2.5 max-h-72 overflow-y-auto pr-1">
                  {activeSheet.headers.map((h, hIdx) => (
                    <div key={`${h.colLetter}-${h.name}-${hIdx}`} className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between font-mono text-[11px]">
                          <span className="text-purple-700 font-bold">{h.colLetter}</span>
                          <span className={`px-1.5 py-0.2 rounded text-[10px] font-semibold ${
                            h.inferredType === 'integer' ? 'bg-blue-100 text-blue-800' :
                            h.inferredType === 'decimal' ? 'bg-indigo-100 text-indigo-800' :
                            h.inferredType === 'date' ? 'bg-amber-100 text-amber-800' :
                            h.inferredType === 'boolean' ? 'bg-emerald-100 text-emerald-800' :
                            'bg-slate-200 text-slate-700'
                          }`}>
                            {h.inferredType || 'text'}
                          </span>
                        </div>
                        <div className="font-semibold text-slate-900 truncate mt-1 text-xs" title={h.name}>
                          {h.name}
                        </div>
                      </div>

                      <div className="mt-2 pt-1 border-t border-slate-200/60 flex items-center justify-between text-[10px] text-slate-500 font-sans">
                        <span>{h.uniqueCount ?? 0} unique</span>
                        {h.isCandidateKey && (
                          <span className="text-amber-700 font-bold flex items-center space-x-0.5">
                            <Key className="w-2.5 h-2.5" />
                            <span>Merge Key</span>
                          </span>
                        )}
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
                        <span>{activeSheet.mergedRanges.length} Merges Resolved</span>
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

                <div className="overflow-x-auto max-h-96">
                  <table className="w-full text-left text-xs font-mono">
                    <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider border-b border-slate-200 sticky top-0">
                      <tr>
                        <th className="px-3 py-2 w-16 text-slate-400">Row</th>
                        {activeSheet.headers.map((h, hIdx) => (
                          <th key={`${h.colLetter}-${h.name}-${hIdx}`} className="px-3 py-2 whitespace-nowrap">
                            <span className="text-slate-500 mr-1">{h.colLetter}:</span>
                            <span className="text-slate-900 font-bold">{h.name}</span>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {activeSheet.sampleRows.map((sr) => (
                        <tr key={sr.rowNumber} className="hover:bg-slate-50 transition-colors">
                          <td className="px-3 py-2 text-slate-400">{sr.rowNumber}</td>
                          {activeSheet.headers.map((h, hIdx) => {
                            const rawCell = sr.data[h.colLetter] !== undefined ? sr.data[h.colLetter] : sr.data[h.name];
                            return (
                              <td key={`${h.colLetter}-${hIdx}`} className="px-3 py-2 whitespace-nowrap text-slate-800">
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
            </>
          )}
        </>
      )}

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
              Paste your student table (CSV or copy-pasted spreadsheet cells) to extract real columns (e.g. <code>username, fullName, indexNumber, dob, class, division, password, email</code>).
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

      {/* AI WORKBOOK ANALYZER MULTI-SHEET SOLUTION MODAL */}
      {showAiResultModal && aiAnalysisResult && (
        <div className="fixed inset-0 bg-slate-950/70 backdrop-blur-sm flex items-center justify-center p-3 sm:p-6 z-50 overflow-y-auto animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl max-w-5xl w-full max-h-[92vh] flex flex-col shadow-2xl border border-purple-200 overflow-hidden my-auto">
            {/* Modal Header */}
            <div className="p-5 bg-linear-to-r from-purple-900 via-indigo-900 to-slate-900 text-white flex items-center justify-between shrink-0 border-b border-purple-800">
              <div className="flex items-center space-x-3">
                <div className="p-2.5 rounded-xl bg-purple-500/20 border border-purple-400/30 text-purple-300 backdrop-blur-xs">
                  <Sparkles className="w-5 h-5 text-purple-300" />
                </div>
                <div>
                  <div className="flex items-center space-x-2">
                    <h3 className="text-base sm:text-lg font-black text-white">AI Workbook Mapping Solution</h3>
                    <span className="px-2 py-0.5 rounded-full text-[10px] font-mono bg-purple-400/20 text-purple-200 border border-purple-400/30">
                      {aiAnalysisResult.modelUsed || 'gemini-3.8-flash'}
                    </span>
                    {aiAnalysisResult.aiPowered && (
                      <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-400/30">
                        Live AI Verified
                      </span>
                    )}
                  </div>
                  <p className="text-xs text-purple-200/80 font-mono mt-0.5">
                    {aiAnalysisResult.filename} • {aiAnalysisResult.sheetSolutions.length} Worksheets Analyzed
                  </p>
                </div>
              </div>

              <button
                onClick={() => setShowAiResultModal(false)}
                className="p-2 rounded-lg text-purple-300 hover:text-white hover:bg-white/10 transition-colors"
                title="Close"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-5 overflow-y-auto space-y-5 flex-1 bg-slate-50/50">
              {/* Architecture Overview Banner */}
              <div className="p-4 rounded-xl bg-purple-50/80 border border-purple-200/70 text-xs flex items-start space-x-3 text-purple-950">
                <Brain className="w-5 h-5 text-purple-600 shrink-0 mt-0.5" />
                <div className="flex-1">
                  <div className="flex items-center justify-between">
                    <h4 className="font-bold text-purple-950 text-xs">Workbook Data Architecture Summary</h4>
                    {aiAnalysisResult.fallbackActive && (
                      <span className="text-[10px] px-2 py-0.5 rounded-full font-bold bg-amber-100 text-amber-900 border border-amber-300">
                        ⚡ Heuristic Rule-Engine
                      </span>
                    )}
                  </div>
                  <p className="text-purple-900 mt-1 leading-relaxed">
                    {aiAnalysisResult.architectureSummary}
                  </p>
                  {aiAnalysisResult.fallbackNotice && (
                    <div className="mt-2.5 p-2.5 rounded-lg bg-amber-50 border border-amber-200 text-amber-900 text-xs flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                      <div className="flex items-center space-x-2">
                        <Info className="w-4 h-4 text-amber-700 shrink-0" />
                        <span>{aiAnalysisResult.fallbackNotice}</span>
                      </div>
                      <button
                        onClick={handleRunFullAiWorkbookAnalysis}
                        disabled={isAiAnalyzingFullWorkbook}
                        className="px-3 py-1 rounded-lg bg-amber-600 hover:bg-amber-700 text-white font-semibold text-xs shrink-0 transition-colors shadow-2xs disabled:opacity-50"
                      >
                        {isAiAnalyzingFullWorkbook ? 'Retrying AI...' : '⚡ Retry AI Analysis'}
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Sheet Navigation Selector */}
              <div className="space-y-2">
                <label className="text-xs font-bold text-slate-700 uppercase tracking-wider block">
                  Select Worksheet Mapping Solution ({aiAnalysisResult.sheetSolutions.length} Sheets):
                </label>
                <div className="flex flex-wrap gap-2">
                  {aiAnalysisResult.sheetSolutions.map((sol, sIdx) => {
                    const isSelected = selectedAiSheetIdx === sIdx;
                    return (
                      <button
                        key={`${sol.worksheetName}-${sIdx}`}
                        onClick={() => setSelectedAiSheetIdx(sIdx)}
                        className={`px-3.5 py-2 rounded-xl text-xs font-semibold flex items-center space-x-2 transition-all border ${
                          isSelected
                            ? 'bg-purple-600 text-white border-purple-600 shadow-md shadow-purple-600/20'
                            : 'bg-white text-slate-700 border-slate-200 hover:border-purple-300 hover:bg-purple-50/40'
                        }`}
                      >
                        <span>{sol.worksheetName}</span>
                        <span className={`text-[10px] px-1.5 py-0.5 rounded font-mono ${
                          isSelected ? 'bg-purple-700 text-purple-100' : 'bg-slate-100 text-slate-600'
                        }`}>
                          → {sol.suggestedTable}
                        </span>
                        <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-bold ${
                          isSelected ? 'bg-emerald-400 text-emerald-950' : 'bg-emerald-100 text-emerald-800'
                        }`}>
                          {sol.confidence}%
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Active Sheet Solution Details */}
              {aiAnalysisResult.sheetSolutions[selectedAiSheetIdx] && (() => {
                const currentSol = aiAnalysisResult.sheetSolutions[selectedAiSheetIdx];
                return (
                  <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs space-y-4">
                    {/* Header Spec Grid */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 text-xs">
                      <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                        <span className="text-slate-500 block text-[11px]">Target Supabase Table</span>
                        <span className="font-mono font-bold text-slate-900 text-sm">public.{currentSol.suggestedTable}</span>
                      </div>
                      <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                        <span className="text-slate-500 block text-[11px]">Header Row & Data Start</span>
                        <span className="font-mono font-bold text-slate-900 text-sm">
                          Row {currentSol.headerRow} (Data: Row {currentSol.dataStartRow})
                        </span>
                      </div>
                      <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                        <span className="text-slate-500 block text-[11px]">Primary Unique Upsert Key</span>
                        <span className="font-mono font-bold text-purple-700 text-sm flex items-center space-x-1">
                          <Key className="w-3.5 h-3.5 text-purple-600" />
                          <span>{currentSol.uniqueKeyColumn || currentSol.columns.find(c => c.uniqueKey)?.supabaseColumn || 'None'}</span>
                        </span>
                      </div>
                      <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                        <span className="text-slate-500 block text-[11px]">Section Heading Extractor</span>
                        <span className="font-mono font-semibold text-slate-800 text-xs">
                          {currentSol.sectionHeadingTargetCol ? (
                            <span className="text-indigo-700">Column: '{currentSol.sectionHeadingTargetCol}'</span>
                          ) : (
                            <span className="text-slate-400">No merged section</span>
                          )}
                        </span>
                      </div>
                    </div>

                    {/* AI Reasoning Note */}
                    <div className="p-3.5 rounded-lg bg-indigo-50/60 border border-indigo-100 text-xs text-indigo-950 flex items-start space-x-2">
                      <Info className="w-4 h-4 text-indigo-600 shrink-0 mt-0.5" />
                      <div>
                        <strong className="font-semibold text-indigo-950">AI Reasoning: </strong>
                        <span className="text-indigo-900">{currentSol.reasoning}</span>
                      </div>
                    </div>

                    {/* Columns Matrix Table */}
                    <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                      <div className="p-3 bg-slate-100 border-b border-slate-200 text-xs font-bold text-slate-700 flex items-center justify-between">
                        <span>Mapped Columns & Inferred Types ({currentSol.columns.length})</span>
                        <span className="text-[11px] text-slate-500 font-normal">Auto-mapped by AI</span>
                      </div>
                      <div className="overflow-x-auto max-h-72">
                        <table className="w-full text-left text-xs">
                          <thead className="bg-slate-50 text-slate-600 font-semibold border-b border-slate-200 sticky top-0">
                            <tr>
                              <th className="px-3.5 py-2">Col</th>
                              <th className="px-3.5 py-2">Source Excel Header</th>
                              <th className="px-3.5 py-2">Target Supabase Column</th>
                              <th className="px-3.5 py-2">Inferred Type</th>
                              <th className="px-3.5 py-2">Transformation</th>
                              <th className="px-3.5 py-2">Upsert Key / Required</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-mono">
                            {currentSol.columns.map((col, cIdx) => (
                              <tr key={col.id || cIdx} className="hover:bg-slate-50/80 transition-colors">
                                <td className="px-3.5 py-2 font-bold text-slate-500">{col.excelColumn}</td>
                                <td className="px-3.5 py-2 font-sans text-slate-800 font-medium">{col.excelHeader}</td>
                                <td className="px-3.5 py-2 font-bold text-purple-700">public.{currentSol.suggestedTable}.{col.supabaseColumn}</td>
                                <td className="px-3.5 py-2">
                                  <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                    col.dataType === 'integer' ? 'bg-blue-100 text-blue-800' :
                                    col.dataType === 'decimal' ? 'bg-indigo-100 text-indigo-800' :
                                    col.dataType === 'date' ? 'bg-amber-100 text-amber-800' :
                                    col.dataType === 'boolean' ? 'bg-emerald-100 text-emerald-800' :
                                    'bg-slate-100 text-slate-700'
                                  }`}>
                                    {col.dataType.toUpperCase()}
                                  </span>
                                </td>
                                <td className="px-3.5 py-2">
                                  <span className="px-2 py-0.5 rounded text-[10px] bg-slate-100 font-mono text-slate-700">
                                    {col.transformation}
                                  </span>
                                </td>
                                <td className="px-3.5 py-2 font-sans text-[11px]">
                                  {col.uniqueKey ? (
                                    <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 font-bold flex items-center space-x-1 w-max">
                                      <Key className="w-3 h-3 text-amber-700" />
                                      <span>UNIQUE KEY</span>
                                    </span>
                                  ) : col.required ? (
                                    <span className="text-red-600 font-semibold">Required</span>
                                  ) : (
                                    <span className="text-slate-400">Optional</span>
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>

            {/* Modal Footer Actions */}
            <div className="p-4 bg-white border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 shrink-0">
              <div className="text-xs text-slate-500">
                <span>Applying AI mappings sets live synchronization rules and saves permanently to system database.</span>
              </div>

              <div className="flex items-center space-x-2 w-full sm:w-auto justify-end">
                <button
                  type="button"
                  onClick={() => setShowAiResultModal(false)}
                  className="px-4 py-2 rounded-xl border border-slate-300 text-xs font-semibold text-slate-700 hover:bg-slate-50 transition-colors"
                >
                  Cancel
                </button>

                {aiAnalysisResult.sheetSolutions[selectedAiSheetIdx] && (
                  <button
                    type="button"
                    onClick={() => handleApplySingleAiSolution(aiAnalysisResult.sheetSolutions[selectedAiSheetIdx])}
                    className="px-4 py-2 rounded-xl bg-purple-100 hover:bg-purple-200 text-purple-900 text-xs font-bold transition-colors"
                  >
                    Apply Sheet '{aiAnalysisResult.sheetSolutions[selectedAiSheetIdx].worksheetName}'
                  </button>
                )}

                <button
                  id="btn-apply-all-ai-solutions"
                  type="button"
                  onClick={handleApplyAllAiSolutions}
                  disabled={isApplyingAi}
                  className="px-5 py-2.5 rounded-xl bg-linear-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-xs font-bold shadow-lg shadow-emerald-700/20 transition-all flex items-center space-x-2 disabled:opacity-50"
                >
                  <CheckCircle2 className="w-4 h-4 text-white" />
                  <span>{isApplyingAi ? 'Applying Mappings...' : `✨ Apply All ${aiAnalysisResult.sheetSolutions.length} Sheet Mappings`}</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* SHARING & LAYOUT BLUEPRINT EXPORT MODAL */}
      {showLayoutExportModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-xs">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-4xl max-h-[92vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            {/* Modal Header */}
            <div className="p-5 border-b border-slate-200 flex items-center justify-between bg-linear-to-r from-indigo-900 via-purple-900 to-slate-900 text-white">
              <div className="flex items-center space-x-3">
                <div className="p-2.5 rounded-xl bg-indigo-500/20 text-indigo-300 border border-indigo-400/30">
                  <Copy className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-base font-bold text-white flex items-center space-x-2">
                    <span>Share Custom Excel Layouts with AI</span>
                    <span className="px-2 py-0.5 rounded-full text-[10px] bg-indigo-400/20 text-indigo-200 font-mono">
                      {analysis.filename}
                    </span>
                  </h2>
                  <p className="text-xs text-indigo-200/80">
                    How to get your custom Excel designs (timetables, multi-grade sheets, 2-cell slots) to the AI assistant
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowLayoutExportModal(false)}
                className="p-1.5 rounded-lg text-indigo-200 hover:text-white hover:bg-white/10 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 overflow-y-auto space-y-6 text-xs text-slate-600">
              {/* 4 Practical Ways Card */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                <div className="p-4 rounded-xl border border-purple-200 bg-purple-50/60 flex flex-col justify-between">
                  <div>
                    <div className="font-bold text-purple-950 flex items-center space-x-2 text-sm">
                      <span className="w-5 h-5 rounded-full bg-purple-600 text-white flex items-center justify-center text-[11px] font-bold">1</span>
                      <span>Copy & Paste Direct from Excel into Chat</span>
                    </div>
                    <p className="mt-2 text-slate-600 leading-relaxed">
                      In Excel, highlight 10–25 rows of your sheet (including headers, merged banners, and timetable slots). Press <kbd className="px-1.5 py-0.5 rounded bg-white border border-slate-300 font-mono text-slate-800">Ctrl+C</kbd>, then paste (<kbd className="px-1.5 py-0.5 rounded bg-white border border-slate-300 font-mono text-slate-800">Ctrl+V</kbd>) directly into the AI chat. Excel automatically pastes as Tab-Separated Values (TSV) preserving all cell boundaries!
                    </p>
                  </div>
                  <div className="mt-3 text-[11px] text-purple-700 font-medium">⚡ Fastest method: no file conversion needed</div>
                </div>

                <div className="p-4 rounded-xl border border-indigo-200 bg-indigo-50/60 flex flex-col justify-between">
                  <div>
                    <div className="font-bold text-indigo-950 flex items-center space-x-2 text-sm">
                      <span className="w-5 h-5 rounded-full bg-indigo-600 text-white flex items-center justify-center text-[11px] font-bold">2</span>
                      <span>Upload Your File in this Web App</span>
                    </div>
                    <p className="mt-2 text-slate-600 leading-relaxed">
                      Click the purple <strong>"Upload Excel (.xlsx/.csv)"</strong> button in the top bar of this page. The system parses all sheets in your browser, discovers all merged coordinates, examines column types, and calculates SHA-256 hashes immediately.
                    </p>
                  </div>
                  <div className="mt-3 text-[11px] text-indigo-700 font-medium">📁 Uploads full multi-sheet .xlsx files</div>
                </div>

                <div className="p-4 rounded-xl border border-emerald-200 bg-emerald-50/60 flex flex-col justify-between">
                  <div>
                    <div className="font-bold text-emerald-950 flex items-center space-x-2 text-sm">
                      <span className="w-5 h-5 rounded-full bg-emerald-600 text-white flex items-center justify-center text-[11px] font-bold">3</span>
                      <span>1-Click Copy Structured Blueprint</span>
                    </div>
                    <p className="mt-2 text-slate-600 leading-relaxed">
                      Click the <strong>"Copy Blueprint to Clipboard"</strong> button below. It generates a compact text blueprint of all sheets, merged cell ranges (e.g. B3:C3 Period 1), and sample rows, which you can paste straight into our conversation.
                    </p>
                  </div>
                  <div className="mt-3">
                    <button
                      type="button"
                      onClick={handleCopyLayoutBlueprint}
                      className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs transition-colors"
                    >
                      {hasCopiedExportText ? (
                        <>
                          <Check className="w-3.5 h-3.5 text-white" />
                          <span>Copied to Clipboard!</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>Copy Blueprint to Clipboard</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>

                <div className="p-4 rounded-xl border border-amber-200 bg-amber-50/60 flex flex-col justify-between">
                  <div>
                    <div className="font-bold text-amber-950 flex items-center space-x-2 text-sm">
                      <span className="w-5 h-5 rounded-full bg-amber-600 text-white flex items-center justify-center text-[11px] font-bold">4</span>
                      <span>Nextcloud or Cloud Drive Link</span>
                    </div>
                    <p className="mt-2 text-slate-600 leading-relaxed">
                      Upload your spreadsheet to your Nextcloud <code className="bg-white px-1 py-0.5 rounded font-mono border border-amber-300">/ExcelImports</code> WebDAV folder, or upload to Google Drive / Dropbox / OneDrive and paste the share link into our chat.
                    </p>
                  </div>
                  <div className="mt-3 text-[11px] text-amber-800 font-medium">☁️ Ideal for enterprise or large workbooks</div>
                </div>
              </div>

              {/* Blueprint Preview Card */}
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-900 text-xs">
                    Current Loaded Layout Blueprint Preview ({analysis.totalWorksheets} Sheets)
                  </span>
                  <button
                    type="button"
                    onClick={handleCopyLayoutBlueprint}
                    className="inline-flex items-center space-x-1 text-indigo-600 hover:text-indigo-800 font-semibold"
                  >
                    <Copy className="w-3.5 h-3.5" />
                    <span>{hasCopiedExportText ? 'Copied!' : 'Copy to Clipboard'}</span>
                  </button>
                </div>

                <div className="bg-slate-900 rounded-xl p-4 text-emerald-400 font-mono text-[11px] max-h-60 overflow-y-auto leading-relaxed border border-slate-800 shadow-inner">
                  <pre className="whitespace-pre-wrap">{generateLayoutStructureExport(analysis)}</pre>
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={handleLoadSchoolSample}
                className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-xl border border-purple-300 bg-white hover:bg-purple-50 text-purple-700 text-xs font-semibold shadow-2xs transition-colors"
              >
                <BookOpen className="w-3.5 h-3.5" />
                <span>Load Sample with Grade 6..8 & Timetable Matrix</span>
              </button>

              <button
                type="button"
                onClick={() => setShowLayoutExportModal(false)}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold transition-colors"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* GOOGLE SHEETS & PRESETS MODAL */}
      {showGoogleSheetsModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/70 backdrop-blur-xs p-4">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            {/* Modal Header */}
            <div className="p-5 bg-linear-to-r from-sky-900 via-indigo-900 to-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center space-x-3">
                <div className="p-2.5 rounded-xl bg-sky-500/20 text-sky-300 border border-sky-400/30">
                  <Globe className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-base text-white tracking-tight flex items-center space-x-2">
                    <span>Google Sheets & Real-World Workbooks</span>
                    <span className="px-2 py-0.5 rounded-full text-[10px] bg-sky-500/30 text-sky-200 font-semibold">
                      Live Connect
                    </span>
                  </h3>
                  <p className="text-xs text-sky-200/80 mt-0.5">
                    Load live Google Spreadsheets directly or test real-world complex architectures (Timetables, Ledgers, Staff Pivots).
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowGoogleSheetsModal(false)}
                className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-white/10 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 overflow-y-auto space-y-6">
              {googleSheetError && (
                <div className="p-3.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-800 text-xs flex items-center space-x-2">
                  <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                  <span>{googleSheetError}</span>
                </div>
              )}

              {/* Section 1: Quick Load Real-World Workbooks */}
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold uppercase tracking-wider text-slate-500">
                    Real-World Production Workbooks (Pre-Loaded for Instant Access)
                  </span>
                  <span className="text-[11px] text-slate-400">Zero network wait time</span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  {/* Card 1: Master Timetable */}
                  <div className="p-4 rounded-xl border border-amber-200 bg-linear-to-b from-amber-50/50 to-white flex flex-col justify-between hover:shadow-md transition-all">
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="p-2 rounded-lg bg-amber-100 text-amber-700">
                          <Calendar className="w-4 h-4" />
                        </span>
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800">
                          5 Days • Matrix
                        </span>
                      </div>
                      <h4 className="font-bold text-slate-900 text-sm">Class Wise Time Table</h4>
                      <p className="text-xs text-slate-600 leading-relaxed">
                        School master timetable. Vertical grade blocks (Grades 6..13) with paired interleaved rows: Row 1 = Division & Subjects, Row 2 = Assigned Teachers.
                      </p>
                      <div className="pt-2 text-[11px] text-amber-800 font-medium">
                        ⚡ Unpivots to 2,482 relational period slots
                      </div>
                    </div>

                    <button
                      type="button"
                      disabled={isFetchingGoogleSheet}
                      onClick={() => handleLoadPreset('timetable')}
                      className="mt-4 w-full py-2 px-3 rounded-lg bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs shadow-xs transition-colors flex items-center justify-center space-x-1.5 disabled:opacity-50"
                    >
                      {isFetchingGoogleSheet ? (
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                      ) : (
                        <Sparkles className="w-3.5 h-3.5" />
                      )}
                      <span>Load Timetable Matrix</span>
                    </button>
                  </div>

                  {/* Card 2: Donations Ledger */}
                  <div className="p-4 rounded-xl border border-emerald-200 bg-linear-to-b from-emerald-50/50 to-white flex flex-col justify-between hover:shadow-md transition-all">
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="p-2 rounded-lg bg-emerald-100 text-emerald-700">
                          <DollarSign className="w-4 h-4" />
                        </span>
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">
                          8 Sheets • Ledger
                        </span>
                      </div>
                      <h4 className="font-bold text-slate-900 text-sm">JHC Donation Details</h4>
                      <p className="text-xs text-slate-600 leading-relaxed">
                        Multi-sheet financial ledger: Things Donation, Books, Projects, Cash SDC, Needy Students, and Receipts. Top banner offsets & mid-table year section divider rows.
                      </p>
                      <div className="pt-2 text-[11px] text-emerald-800 font-medium">
                        🧹 Filters Year-2023 dividers & "Rs " strings
                      </div>
                    </div>

                    <button
                      type="button"
                      disabled={isFetchingGoogleSheet}
                      onClick={() => handleLoadPreset('donations')}
                      className="mt-4 w-full py-2 px-3 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs transition-colors flex items-center justify-center space-x-1.5 disabled:opacity-50"
                    >
                      {isFetchingGoogleSheet ? (
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                      ) : (
                        <Sparkles className="w-3.5 h-3.5" />
                      )}
                      <span>Load Financial Ledger</span>
                    </button>
                  </div>

                  {/* Card 3: Staff Allocations */}
                  <div className="p-4 rounded-xl border border-indigo-200 bg-linear-to-b from-indigo-50/50 to-white flex flex-col justify-between hover:shadow-md transition-all">
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="p-2 rounded-lg bg-indigo-100 text-indigo-700">
                          <GraduationCap className="w-4 h-4" />
                        </span>
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-100 text-indigo-800">
                          8 Grades • Pivot
                        </span>
                      </div>
                      <h4 className="font-bold text-slate-900 text-sm">Subject Teacher Class Wise</h4>
                      <p className="text-xs text-slate-600 leading-relaxed">
                        Teacher allocations across Grade 6..11 and AL. 2D Pivot grid: Subjects along rows, Class Divisions A..H along columns, teacher names and co-teachers in cells.
                      </p>
                      <div className="pt-2 text-[11px] text-indigo-800 font-medium">
                        ⚡ Unpivots to 729 clean assignment slots
                      </div>
                    </div>

                    <button
                      type="button"
                      disabled={isFetchingGoogleSheet}
                      onClick={() => handleLoadPreset('teacher_allocations')}
                      className="mt-4 w-full py-2 px-3 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs shadow-xs transition-colors flex items-center justify-center space-x-1.5 disabled:opacity-50"
                    >
                      {isFetchingGoogleSheet ? (
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                      ) : (
                        <Sparkles className="w-3.5 h-3.5" />
                      )}
                      <span>Load Staff Allocations</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* Section 2: Paste Google Sheets URL */}
              <div className="p-5 rounded-xl border border-slate-200 bg-slate-50 space-y-4">
                <div className="space-y-1">
                  <h4 className="font-bold text-slate-900 text-sm flex items-center space-x-1.5">
                    <Globe className="w-4 h-4 text-sky-600" />
                    <span>Or Fetch Live from Google Sheets Link</span>
                  </h4>
                  <p className="text-xs text-slate-500">
                    Paste any public Google Sheets link or Google Drive spreadsheet link to download and analyze it live.
                  </p>
                </div>

                <div className="flex flex-col sm:flex-row gap-2">
                  <input
                    type="text"
                    value={googleSheetUrlInput}
                    onChange={(e) => setGoogleSheetUrlInput(e.target.value)}
                    placeholder="https://docs.google.com/spreadsheets/d/1Ir0ySRehyaAvVhUowspQgobCLObFViNx/edit?..."
                    className="flex-1 px-3.5 py-2.5 rounded-xl border border-slate-300 text-xs font-mono bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-500 shadow-2xs"
                  />
                  <button
                    type="button"
                    disabled={isFetchingGoogleSheet || !googleSheetUrlInput.trim()}
                    onClick={() => handleFetchGoogleSheetUrl()}
                    className="px-5 py-2.5 rounded-xl bg-sky-600 hover:bg-sky-700 text-white font-bold text-xs shadow-xs transition-colors flex items-center justify-center space-x-2 shrink-0 disabled:opacity-50"
                  >
                    {isFetchingGoogleSheet ? (
                      <>
                        <RefreshCw className="w-4 h-4 animate-spin text-white" />
                        <span>Fetching Sheet...</span>
                      </>
                    ) : (
                      <>
                        <ExternalLink className="w-4 h-4" />
                        <span>Fetch & Analyze</span>
                      </>
                    )}
                  </button>
                </div>

                {/* Quick Paste Pills */}
                <div className="space-y-1.5 pt-1">
                  <span className="text-[11px] font-semibold text-slate-500">
                    Quick-paste links from your prompt:
                  </span>
                  <div className="flex flex-wrap gap-2 text-[11px]">
                    <button
                      type="button"
                      onClick={() => {
                        setGoogleSheetUrlInput('https://docs.google.com/spreadsheets/d/1Ir0ySRehyaAvVhUowspQgobCLObFViNx/edit');
                        handleFetchGoogleSheetUrl('https://docs.google.com/spreadsheets/d/1Ir0ySRehyaAvVhUowspQgobCLObFViNx/edit');
                      }}
                      className="px-2.5 py-1 rounded-lg bg-white border border-slate-200 hover:border-sky-300 hover:bg-sky-50 text-slate-700 font-mono transition-colors"
                    >
                      🔗 Link 1 (Class Wise Time Table)
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setGoogleSheetUrlInput('https://docs.google.com/spreadsheets/d/1OEVvR6YruzJ0kpqYoA9-vDggwJozpweM/edit');
                        handleFetchGoogleSheetUrl('https://docs.google.com/spreadsheets/d/1OEVvR6YruzJ0kpqYoA9-vDggwJozpweM/edit');
                      }}
                      className="px-2.5 py-1 rounded-lg bg-white border border-slate-200 hover:border-emerald-300 hover:bg-emerald-50 text-slate-700 font-mono transition-colors"
                    >
                      🔗 Link 2 (JHC Donation Details)
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setGoogleSheetUrlInput('https://docs.google.com/spreadsheets/d/1FX9-OiRMBTa5wpXC-wOlDXe1dv6MwHLq/edit');
                        handleFetchGoogleSheetUrl('https://docs.google.com/spreadsheets/d/1FX9-OiRMBTa5wpXC-wOlDXe1dv6MwHLq/edit');
                      }}
                      className="px-2.5 py-1 rounded-lg bg-white border border-slate-200 hover:border-indigo-300 hover:bg-indigo-50 text-slate-700 font-mono transition-colors"
                    >
                      🔗 Link 3 (Subject Teacher Class Wise)
                    </button>
                  </div>
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end">
              <button
                type="button"
                onClick={() => setShowGoogleSheetsModal(false)}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold transition-colors"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* INBUILT AI ANALYSER & PRESETS LIBRARY MODAL */}
      <AiPresetsModal
        isOpen={showAiPresetsModal}
        onClose={() => setShowAiPresetsModal(false)}
        currentAnalysis={analysis}
        activeMappings={activeMappings}
        onApplyPreset={handleApplyPreset}
        onApplyAllPresets={handleApplyAllStudentPresets}
        onLoadPresetWorkbook={(presetId) => handleLoadPreset(presetId as any)}
        supabaseConfig={supabaseConfig}
        supabaseTables={supabaseTables}
      />
    </div>
  );
};
