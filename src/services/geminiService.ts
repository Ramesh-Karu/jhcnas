import { 
  SheetAnalysis, 
  WorkbookAnalysis,
  WorksheetMapping, 
  ColumnMapping, 
  DataType, 
  TransformationType,
  SupabaseTableInfo,
  AiWorkbookAnalysisResult,
  AiSheetMappingSolution
} from '../types';
import { ApiClient } from './apiClient';

export class GeminiWorkbookService {
  /**
   * Run full workbook AI analysis across all sheets simultaneously, producing a complete mapping solution.
   */
  static async analyzeFullWorkbook(
    workbook: WorkbookAnalysis,
    supabaseTables?: SupabaseTableInfo[]
  ): Promise<AiWorkbookAnalysisResult> {
    try {
      const res = await ApiClient.analyzeWorkbookWithAi({
        workbook,
        supabaseTables,
        consolidationMode: workbook.consolidationMode || 'SEPARATE_TABLES'
      });

      if (res.success && res.result) {
        if (res.fallbackActive !== undefined) {
          res.result.fallbackActive = res.fallbackActive;
        }
        if (res.fallbackNotice) {
          res.result.fallbackNotice = res.fallbackNotice;
        }
        return res.result;
      }
    } catch (e) {
      console.info('[GeminiWorkbookService] Server AI call reached fallback path:', e);
    }

    // Local client-side fallback
    const solutions: AiSheetMappingSolution[] = await Promise.all(
      workbook.worksheets.map(async (sheet) => {
        const single = await this.analyzeSheetAndSuggestMappings(sheet, workbook.filename);
        return {
          worksheetName: sheet.sheetName,
          suggestedTable: single.suggestedTable,
          headerRow: single.headerRow,
          dataStartRow: single.dataStartRow,
          dataEndRow: sheet.totalRows,
          sectionHeadingTargetCol: single.sectionHeadingTargetCol,
          uniqueKeyColumn: single.columns.find(c => c.uniqueKey)?.supabaseColumn || single.columns[0]?.supabaseColumn,
          confidence: 95,
          reasoning: single.reasoning,
          columns: single.columns,
        };
      })
    );

    return {
      filename: workbook.filename,
      totalSheets: workbook.worksheets.length,
      architectureSummary: `Auto-analyzed ${solutions.length} worksheets with entity table mappings and column type inference.`,
      recommendedConsolidationMode: workbook.consolidationMode || 'SEPARATE_TABLES',
      sheetSolutions: solutions,
      aiPowered: false,
      modelUsed: 'Local Heuristic Engine'
    };
  }

  /**
   * Convert an AI Workbook Analysis Result directly into a list of WorksheetMapping objects for immediate execution.
   */
  static convertSolutionsToMappings(
    analysisResult: AiWorkbookAnalysisResult,
    workbookName: string
  ): WorksheetMapping[] {
    return analysisResult.sheetSolutions.map((solution, idx) => ({
      id: `wm-ai-${idx}-${Date.now()}`,
      workbookName,
      worksheetName: solution.worksheetName,
      supabaseTable: solution.suggestedTable,
      headerRow: solution.headerRow,
      dataStartRow: solution.dataStartRow,
      dataEndRow: solution.dataEndRow,
      sectionHeadingTargetCol: solution.sectionHeadingTargetCol,
      enabled: true,
      syncPolicy: solution.syncPolicy || 'BIDIRECTIONAL',
      columns: solution.columns.map((c, cIdx) => ({
        id: c.id || `col-ai-${idx}-${cIdx}-${Date.now()}`,
        excelColumn: c.excelColumn,
        excelHeader: c.excelHeader,
        supabaseColumn: c.supabaseColumn,
        dataType: c.dataType,
        required: c.required,
        uniqueKey: c.uniqueKey,
        defaultValue: c.defaultValue,
        transformation: c.transformation,
        validationRegex: c.validationRegex
      }))
    }));
  }

  static async analyzeSheetAndSuggestMappings(
    sheet: SheetAnalysis,
    workbookName: string
  ): Promise<{
    suggestedTable: string;
    headerRow: number;
    dataStartRow: number;
    sectionHeadingTargetCol?: string;
    columns: ColumnMapping[];
    reasoning: string;
  }> {
    // Prepare prompt payload with sheet statistics, merged cells, and headers
    const sampleHeaders = sheet.headers.map(h => `${h.colLetter}: "${h.name}" (sample: ${h.sampleValues.slice(0, 2).join(', ')})`).join('\n');
    const sampleMerges = sheet.mergedRanges.map(m => `${m.range} (${m.type}): "${m.value}"`).join('\n');

    const prompt = `You are an expert PostgreSQL database architect and Excel data integration specialist.
Analyze this Excel worksheet metadata:
Workbook: "${workbookName}"
Sheet Name: "${sheet.sheetName}"
Total Rows: ${sheet.totalRows}, Total Columns: ${sheet.totalColumns}
Merged Ranges:
${sampleMerges || 'None'}
Detected Header Candidates: ${JSON.stringify(sheet.candidateHeaderRows)}
Detected Headers:
${sampleHeaders}

Task:
Suggest the optimal:
1. Supabase PostgreSQL table name (snake_case, lowercase, singular/plural e.g. "students", "attendance")
2. Header Row number and Data Start Row number
3. Merged section heading target column if applicable (e.g. if A3:H3 has "CLASS 10A", suggest 'class')
4. Column mappings with:
   - excelColumn (e.g. "A")
   - excelHeader
   - supabaseColumn (snake_case, e.g. "student_number")
   - dataType ('text' | 'integer' | 'decimal' | 'boolean' | 'date')
   - required (boolean)
   - uniqueKey (boolean, identify primary unique upsert key)
   - transformation ('none' | 'trim' | 'uppercase' | 'lowercase' | 'parse_date' | 'parse_number' | 'yes_no_to_boolean' | 'pa_to_status' | 'normalize_phone' | 'normalize_id')

Format your response as pure JSON matching this structure:
{
  "suggestedTable": "students",
  "headerRow": 5,
  "dataStartRow": 6,
  "sectionHeadingTargetCol": "class",
  "reasoning": "Detected merged title at A1:H1 and section header at A3:H3. Column A is unique student ID.",
  "columns": [
    {
      "excelColumn": "A",
      "excelHeader": "Student ID",
      "supabaseColumn": "student_number",
      "dataType": "text",
      "required": true,
      "uniqueKey": true,
      "transformation": "normalize_id"
    }
  ]
}`;

    try {
      const response = await fetch('/api/gemini/suggest-mapping', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt }),
      });

      if (response.ok) {
        const result = await response.json();
        if (result.success && result.data) {
          const parsed = result.data;
          return {
            suggestedTable: parsed.suggestedTable || sheet.sheetName.toLowerCase().replace(/\s+/g, '_'),
            headerRow: parsed.headerRow || sheet.detectedHeaderRow,
            dataStartRow: parsed.dataStartRow || sheet.detectedDataStartRow,
            sectionHeadingTargetCol: parsed.sectionHeadingTargetCol,
            columns: (parsed.columns || []).map((c: any, idx: number) => ({
              id: `ai-${idx}-${Date.now()}`,
              excelColumn: c.excelColumn || sheet.headers[idx]?.colLetter || 'A',
              excelHeader: c.excelHeader || sheet.headers[idx]?.name || '',
              supabaseColumn: c.supabaseColumn || (c.excelHeader || '').toLowerCase().replace(/\s+/g, '_'),
              dataType: (c.dataType as DataType) || 'text',
              required: Boolean(c.required),
              uniqueKey: Boolean(c.uniqueKey),
              transformation: (c.transformation as TransformationType) || 'trim'
            })),
            reasoning: parsed.reasoning || "AI generated mapping based on schema structure and sample values."
          };
        }
      }
    } catch (err) {
      console.info("[Gemini] Server route unavailable or returned error, using heuristic intelligence:", err);
    }

    // Heuristic Fallback Analysis
    return this.heuristicAnalysis(sheet);
  }

  private static heuristicAnalysis(sheet: SheetAnalysis) {
    const rawName = sheet.sheetName.toLowerCase().replace(/\s+/g, '_');
    const tableMap: Record<string, string> = {
      student_details: 'students',
      students: 'students',
      attendance: 'attendance',
      sports: 'sports',
      medical: 'medical',
      results: 'results'
    };
    const suggestedTable = tableMap[rawName] || rawName;

    // Detect section heading if any merged cells exist with section_heading type
    const sectionMerge = sheet.mergedRanges.find(m => m.type === 'section_heading');
    const sectionHeadingTargetCol = sectionMerge ? 'class' : undefined;

    const columns: ColumnMapping[] = sheet.headers.map((h, i) => {
      const lower = h.name.toLowerCase();
      let dataType: DataType = 'text';
      let transformation: TransformationType = 'trim';
      let uniqueKey = false;
      let required = false;

      if (lower.includes('id') || lower.includes('number') || lower.includes('code')) {
        transformation = 'normalize_id';
        if (i === 0) {
          uniqueKey = true;
          required = true;
        }
      } else if (lower.includes('date') || lower.includes('dob')) {
        dataType = 'date';
        transformation = 'parse_date';
      } else if (lower.includes('phone') || lower.includes('mobile')) {
        transformation = 'normalize_phone';
      } else if (lower.includes('status') && sheet.sheetName.toLowerCase().includes('attendance')) {
        transformation = 'pa_to_status';
        required = true;
      } else if (lower.includes('clearance') || lower.includes('active') || lower.includes('enrolled')) {
        dataType = 'boolean';
        transformation = 'yes_no_to_boolean';
      } else if (lower.includes('score') || lower.includes('math') || lower.includes('science') || lower.includes('total')) {
        dataType = 'decimal';
        transformation = 'parse_number';
      }

      const supaCol = lower
        .replace(/student\s*id/i, 'student_number')
        .replace(/student\s*name/i, 'name')
        .replace(/\s+/g, '_')
        .replace(/[^a-z0-9_]/g, '');

      return {
        id: `h-cm-${i}-${Date.now()}`,
        excelColumn: h.colLetter,
        excelHeader: h.name,
        supabaseColumn: supaCol || `col_${h.colLetter.toLowerCase()}`,
        dataType,
        required,
        uniqueKey,
        transformation
      };
    });

    return {
      suggestedTable,
      headerRow: sheet.detectedHeaderRow,
      dataStartRow: sheet.detectedDataStartRow,
      sectionHeadingTargetCol,
      columns,
      reasoning: `Intelligent analyzer identified header row at ${sheet.detectedHeaderRow} with ${columns.length} columns and ${sheet.mergedRanges.length} merged cell ranges. Mappings configured with types, upsert keys, and transformations.`
    };
  }
}
