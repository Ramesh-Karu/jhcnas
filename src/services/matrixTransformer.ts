import * as XLSX from 'xlsx';
import { SheetAnalysis, SheetHeader, DataType } from '../types';

export type WorkbookArchetype = 
  | 'TIMETABLE_MATRIX'        // Weekly/Daily matrix with periods & interleaved Subject/Teacher rows
  | 'MULTI_SHEET_LEDGER'       // Financial / inventory ledger with title banners, inline section dividers & cross-sheet receipts
  | 'PIVOT_ALLOCATION_MATRIX'  // Multi-sheet pivot with Division columns (A-H) and Subject rows
  | 'STANDARD_TABULAR';        // Normal flat rows and columns

export interface NormalizedTimetableRecord {
  day_of_week: string;
  grade: string;
  division: string;
  period: number;
  subject: string;
  primary_teacher: string;
  secondary_teacher: string | null;
  raw_teacher_cell: string;
}

export interface NormalizedTeacherAllocationRecord {
  academic_level: string;
  subject: string;
  division: string;
  primary_teacher: string;
  secondary_teacher: string | null;
  raw_cell_value: string;
}

export interface ArchetypeDetectionResult {
  archetype: WorkbookArchetype;
  confidence: number;
  title: string;
  badge: string;
  summary: string;
  features: string[];
  recommendations: string[];
  canNormalize: boolean;
  normalizedRowCount?: number;
  sampleNormalizedRecords?: any[];
}

export class MatrixTransformer {
  /**
   * Detects the specific structural archetype of a workbook
   */
  static detectArchetype(
    sheetNames: string[],
    worksheetsData: { sheetName: string; rows: any[][] }[]
  ): ArchetypeDetectionResult {
    const sheetNamesLower = sheetNames.map(s => s.toLowerCase());

    // 1. Check for School Master Timetable (Monday - Friday with periods & grades)
    const daysOfWeek = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
    const matchedDays = sheetNamesLower.filter(s => daysOfWeek.includes(s));
    
    if (matchedDays.length >= 3) {
      // Confirm if inner structure has period numbers and grade blocks
      let hasPeriodsAndDivisions = false;
      for (const ws of worksheetsData) {
        for (const row of ws.rows.slice(0, 15)) {
          const rowStr = row.map(c => String(c || '').toLowerCase()).join(' ');
          if (rowStr.includes('div') && (rowStr.includes('1') && rowStr.includes('2') && rowStr.includes('3'))) {
            hasPeriodsAndDivisions = true;
            break;
          }
        }
        if (hasPeriodsAndDivisions) break;
      }

      if (hasPeriodsAndDivisions) {
        return {
          archetype: 'TIMETABLE_MATRIX',
          confidence: 96,
          title: 'School Master Timetable (Interleaved Day Matrix)',
          badge: '📅 Timetable Matrix',
          summary: 'Detected 5-day school timetable with vertically stacked grade sections (Grade 6..13) and interleaved paired rows: Row 1 = Division & Period Subjects, Row 2 = Assigned Teachers.',
          features: [
            '5 Weekly Worksheets (Monday — Friday)',
            'Multi-grade blocks stacked down each worksheet',
            'Periods 1 to 7 represented as horizontal columns',
            'Paired interleaved rows (Subject code followed by Teacher name)',
            'Co-teacher slots separated by slashes or newlines (e.g. Mr.V.Puvaneesan / Mrs.L.Mathiyamuthan)'
          ],
          recommendations: [
            'Unpivot the 2D matrix into a normalized relational table: (day_of_week, grade, division, period, subject, primary_teacher, secondary_teacher)',
            'Enables direct SQL queries such as teacher workload, classroom allocations, and period schedules',
            'Generates over 2,400 clean relational records ready for Supabase PostgreSQL'
          ],
          canNormalize: true
        };
      }
    }

    // 2. Check for Multi-Grade Staff Subject Allocation Pivot Matrix
    const gradeSheetPattern = /^(grade[- ]?\d+|al[- ]?\d{4})/i;
    const gradeMatches = sheetNames.filter(s => gradeSheetPattern.test(s));
    if (gradeMatches.length >= 3) {
      let isStaffMatrix = false;
      for (const ws of worksheetsData) {
        const topRows = ws.rows.slice(0, 5);
        for (const row of topRows) {
          const firstCol = String(row[0] || '').toLowerCase();
          const hasDivCols = row.slice(1, 8).some(c => /^[A-H]$/i.test(String(c || '').trim()));
          if ((firstCol.includes('subject') || firstCol.includes('mathematics') || firstCol.includes('science')) && hasDivCols) {
            isStaffMatrix = true;
            break;
          }
        }
        if (isStaffMatrix) break;
      }

      if (isStaffMatrix) {
        return {
          archetype: 'PIVOT_ALLOCATION_MATRIX',
          confidence: 95,
          title: 'Teacher-Subject Allocation Matrix (Multi-Grade Pivot)',
          badge: '👨‍🏫 Staff Allocation Grid',
          summary: 'Detected homogeneous multi-grade staff allocation sheets (Grade 6..11, AL 2024/2025). Each sheet is a pivot matrix: Rows = Subjects, Columns = Class Divisions (A..H), Cells = Teachers.',
          features: [
            `${sheetNames.length} Grade-Level Worksheets with homogeneous structure`,
            'Class Divisions (A through H) formatted as column headers',
            'Subject curriculum list along the row axis',
            'Cell values contain Teacher names and co-teaching pairs (split with / or line breaks)'
          ],
          recommendations: [
            'Consolidate all grade sheets into a single unified table: teacher_subject_assignments(grade, subject, division, primary_teacher, secondary_teacher)',
            'Extracts ~700+ clean normalized teacher assignment slots',
            'Alternative: Generate separate per-grade tables if distinct grade isolation is preferred'
          ],
          canNormalize: true
        };
      }
    }

    // 3. Check for Multi-Sheet Financial Ledger with Section Dividers & Receipts
    const ledgerKeywords = ['donation', 'receipt', 'cash', 'project', 'contributions', 'sdc', 'folio'];
    const ledgerScore = sheetNamesLower.filter(s => ledgerKeywords.some(kw => s.includes(kw))).length;

    if (ledgerScore >= 2 || sheetNames.some(s => s.toLowerCase().includes('receipt'))) {
      return {
        archetype: 'MULTI_SHEET_LEDGER',
        confidence: 94,
        title: 'School Financial & Contributions Ledger',
        badge: '💰 Multi-Sheet Ledger',
        summary: 'Detected multi-sheet school ledger tracking donations, student contributions, SDC cash funds, and receipt numbers with title banners and embedded mid-table year section divider rows.',
        features: [
          '8 distinct ledger worksheets with varying schemas (Things, Books, Projects, Cash SDC, Needy Students, Receipts)',
          'Top decorative title banners offseting actual column headers to Row 2',
          'Embedded mid-table divider rows (e.g. "Year-2023" or 2020 injected into record rows)',
          'Cross-sheet relational foreign keys linking receipts ("donation receipt no" <-> "Things Donation")',
          'Multilingual UTF-8 Tamil script content and mixed currency representations ("Rs 40000" vs numeric)'
        ],
        recommendations: [
          'Enable intelligent divider-row filtering to prevent section headings from corrupting database rows',
          'Auto-strip currency prefixes ("Rs ", "$") and format strings into standard PostgreSQL numeric(12,2)',
          'Convert DD.MM.YYYY string dates into ISO date format (YYYY-MM-DD)',
          'Create foreign key constraint between donation records and master receipts'
        ],
        canNormalize: true
      };
    }

    return {
      archetype: 'STANDARD_TABULAR',
      confidence: 80,
      title: 'Standard Tabular Workbook',
      badge: '📊 Tabular Workbook',
      summary: 'Standard multi-sheet workbook with tabular rows and columns.',
      features: [`${sheetNames.length} worksheets detected`],
      recommendations: ['Map sheets to Supabase tables using standard column mapping'],
      canNormalize: false
    };
  }

  /**
   * Unpivots a 5-day school timetable workbook into normalized relational records
   */
  static unpivotTimetableWorkbook(wb: XLSX.WorkBook): {
    records: NormalizedTimetableRecord[];
    summary: { totalSlots: number; grades: string[]; days: string[]; teachersCount: number; subjectsCount: number };
  } {
    const records: NormalizedTimetableRecord[] = [];
    const teachersSet = new Set<string>();
    const subjectsSet = new Set<string>();
    const gradesSet = new Set<string>();
    const daysSet = new Set<string>();

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;

      const dayName = sheetName.trim();
      daysSet.add(dayName);

      const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      let currentGrade = '';
      let periods: { colIndex: number; periodNumber: number }[] = [];

      for (let r = 0; r < rows.length; r++) {
        const row = rows[r] || [];
        const firstCell = String(row[0] || '').trim();
        const fullRowText = row.map(c => String(c || '').trim()).join(' ');

        // Check if Grade title row (e.g. "Grade : 6", "Grade 7", "Grade : 12 Com,Arts,Technology")
        const gradeMatch = fullRowText.match(/Grade\s*[:;]?\s*([0-9A-Za-z,\s]+)/i);
        if (gradeMatch && !firstCell.toLowerCase().includes('div')) {
          currentGrade = gradeMatch[1].replace(/_+/g, ' ').trim();
          gradesSet.add(currentGrade);
          continue;
        }

        // Check if period header row (e.g. "Div", 1, 2, 3, 4, 5, 6, 7)
        if (firstCell.toLowerCase().includes('div')) {
          periods = [];
          for (let c = 1; c < row.length; c++) {
            const pVal = String(row[c] || '').trim();
            const pNum = parseInt(pVal, 10);
            if (!isNaN(pNum) && pNum >= 1 && pNum <= 12) {
              periods.push({ colIndex: c, periodNumber: pNum });
            }
          }
          continue;
        }

        // Check for Division letter row (e.g. "A", "B", "C", "D"...)
        if (firstCell && /^[A-Z0-9]$/i.test(firstCell) && r + 1 < rows.length && periods.length > 0) {
          const division = firstCell.toUpperCase();
          const subjectRow = row;
          const teacherRow = rows[r + 1] || [];

          for (const p of periods) {
            const subject = String(subjectRow[p.colIndex] || '').replace(/[\r\n]+/g, ' ').trim();
            const teacherRaw = String(teacherRow[p.colIndex] || '').trim();

            if (subject || teacherRaw) {
              // Parse potential co-teachers
              const teacherLines = teacherRaw.split(/[\r\n\/]+/).map(t => t.trim()).filter(Boolean);
              const primaryTeacher = teacherLines[0] || teacherRaw;
              const secondaryTeacher = teacherLines.length > 1 ? teacherLines.slice(1).join(' / ') : null;

              if (primaryTeacher) teachersSet.add(primaryTeacher);
              if (subject) subjectsSet.add(subject);

              records.push({
                day_of_week: dayName,
                grade: currentGrade || 'Unspecified Grade',
                division,
                period: p.periodNumber,
                subject: subject || 'Unspecified',
                primary_teacher: primaryTeacher || 'None Assigned',
                secondary_teacher: secondaryTeacher,
                raw_teacher_cell: teacherRaw
              });
            }
          }

          // Skip the paired teacher row so we don't process it as a division
          r++;
        }
      }
    }

    return {
      records,
      summary: {
        totalSlots: records.length,
        grades: Array.from(gradesSet),
        days: Array.from(daysSet),
        teachersCount: teachersSet.size,
        subjectsCount: subjectsSet.size
      }
    };
  }

  /**
   * Unpivots homogeneous multi-grade staff subject allocation sheets into normalized relational records
   */
  static unpivotStaffAllocationWorkbook(wb: XLSX.WorkBook): {
    records: NormalizedTeacherAllocationRecord[];
    summary: { totalAssignments: number; grades: string[]; subjectsCount: number; teachersCount: number };
  } {
    const records: NormalizedTeacherAllocationRecord[] = [];
    const gradesSet = new Set<string>();
    const subjectsSet = new Set<string>();
    const teachersSet = new Set<string>();

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;

      const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) continue;

      gradesSet.add(sheetName);

      // Header row: Subject, A, B, C, D, E, F, G, H...
      let headerRowIndex = 0;
      for (let r = 0; r < Math.min(5, rows.length); r++) {
        const row = rows[r];
        const hasSubject = String(row[0] || '').toLowerCase().includes('subject');
        const hasDivs = row.slice(1).some(c => /^[A-H]$/i.test(String(c || '').trim()));
        if (hasSubject || hasDivs) {
          headerRowIndex = r;
          break;
        }
      }

      const headerRow = rows[headerRowIndex] || [];
      const divisions: { colIndex: number; division: string }[] = [];

      for (let c = 1; c < headerRow.length; c++) {
        const divName = String(headerRow[c] || '').trim().toUpperCase();
        if (divName && /^[A-Z0-9]+$/i.test(divName)) {
          divisions.push({ colIndex: c, division: divName });
        }
      }

      for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        const subject = String(row[0] || '').replace(/[\r\n]+/g, ' ').trim();
        if (!subject) continue;

        subjectsSet.add(subject);

        for (const div of divisions) {
          const rawTeacher = String(row[div.colIndex] || '').trim();
          if (rawTeacher) {
            // Split co-teachers
            const teacherParts = rawTeacher.split(/[\/\r\n]+/).map(t => t.trim()).filter(Boolean);
            const primary = teacherParts[0] || rawTeacher;
            const secondary = teacherParts.length > 1 ? teacherParts.slice(1).join(' / ') : null;

            teachersSet.add(primary);

            records.push({
              academic_level: sheetName,
              subject,
              division: div.division,
              primary_teacher: primary,
              secondary_teacher: secondary,
              raw_cell_value: rawTeacher
            });
          }
        }
      }
    }

    return {
      records,
      summary: {
        totalAssignments: records.length,
        grades: Array.from(gradesSet),
        subjectsCount: subjectsSet.size,
        teachersCount: teachersSet.size
      }
    };
  }

  /**
   * Cleans dirty financial ledger worksheets:
   * 1. Strips mid-table year/section divider rows (e.g. ["Year-2023"] or [2020])
   * 2. Propagates the active year/section to subsequent records
   * 3. Sanitizes currency strings ("Rs 40000" -> 40000)
   * 4. Converts DD.MM.YYYY dates to ISO YYYY-MM-DD
   */
  static cleanLedgerRecords(
    rawRows: Record<string, any>[],
    headers: SheetHeader[]
  ): {
    cleanedRecords: Record<string, any>[];
    dividerRowsCount: number;
    detectedYears: string[];
  } {
    const cleanedRecords: Record<string, any>[] = [];
    let activeYear: string | null = null;
    let dividerCount = 0;
    const detectedYearsSet = new Set<string>();

    for (const row of rawRows) {
      const values = Object.entries(row).filter(([k, v]) => !k.startsWith('_') && v !== null && v !== undefined && String(v).trim() !== '');

      // Check if this row is an embedded section divider (e.g. single cell with "Year-2023" or 2020)
      if (values.length === 1) {
        const singleVal = String(values[0][1]).trim();
        const yearMatch = singleVal.match(/\b(19\d{2}|20\d{2})\b/);
        if (yearMatch || /year[- ]?\d{4}/i.test(singleVal)) {
          dividerCount++;
          activeYear = yearMatch ? yearMatch[1] : singleVal;
          detectedYearsSet.add(activeYear);
          continue; // Skip the divider row itself
        }
      }

      // Check if all cells except one are empty and that one cell is just a year
      if (values.length <= 2) {
        const textVal = values.map(v => String(v[1]).trim()).join(' ');
        if (/^(year\s*[-:]?\s*)?(19\d{2}|20\d{2})$/i.test(textVal)) {
          dividerCount++;
          const match = textVal.match(/\b(19\d{2}|20\d{2})\b/);
          if (match) {
            activeYear = match[1];
            detectedYearsSet.add(activeYear);
          }
          continue;
        }
      }

      // Normal data row: clean data types
      const cleanedRow: Record<string, any> = { ...row };

      if (activeYear && !cleanedRow['active_year']) {
        cleanedRow['active_year'] = activeYear;
      }

      for (const h of headers) {
        const rawVal = cleanedRow[h.colLetter] ?? cleanedRow[h.name];
        if (rawVal === undefined || rawVal === null || String(rawVal).trim() === '') continue;

        const strVal = String(rawVal).trim();

        // 1. Clean Currency strings (e.g. "Rs 40000", "Rs. 15,000", "$500")
        if (h.name.toLowerCase().includes('amount') || h.inferredType === 'decimal' || h.inferredType === 'integer') {
          if (/^Rs\.?\s*[\d,]+/i.test(strVal) || /^[$€£₹]\s*[\d,]+/i.test(strVal)) {
            const numericStr = strVal.replace(/[^0-9.]/g, '');
            const parsedNum = parseFloat(numericStr);
            if (!isNaN(parsedNum)) {
              cleanedRow[h.colLetter] = parsedNum;
              cleanedRow[h.name] = parsedNum;
            }
          }
        }

        // 2. Format DD.MM.YYYY or DD/MM/YYYY into standard ISO YYYY-MM-DD
        if (h.name.toLowerCase().includes('date') || h.inferredType === 'date') {
          const dmyMatch = strVal.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
          if (dmyMatch) {
            const day = dmyMatch[1].padStart(2, '0');
            const month = dmyMatch[2].padStart(2, '0');
            let year = dmyMatch[3];
            if (year.length === 2) year = `20${year}`;
            const iso = `${year}-${month}-${day}`;
            cleanedRow[h.colLetter] = iso;
            cleanedRow[h.name] = iso;
          }
        }
      }

      cleanedRecords.push(cleanedRow);
    }

    return {
      cleanedRecords,
      dividerRowsCount: dividerCount,
      detectedYears: Array.from(detectedYearsSet)
    };
  }

  /**
   * Synthesizes a clean normalized WorkbookAnalysis from an unpivoted dataset
   */
  static createNormalizedTimetableAnalysis(
    rawWb: XLSX.WorkBook,
    originalFilename: string
  ): {
    analysis: SheetAnalysis;
    records: NormalizedTimetableRecord[];
  } {
    const { records, summary } = this.unpivotTimetableWorkbook(rawWb);

    const headers: SheetHeader[] = [
      { colLetter: 'A', colIndex: 1, name: 'day_of_week', sampleValues: ['Monday', 'Tuesday', 'Wednesday'], inferredType: 'text' },
      { colLetter: 'B', colIndex: 2, name: 'grade', sampleValues: summary.grades.slice(0, 3), inferredType: 'text' },
      { colLetter: 'C', colIndex: 3, name: 'division', sampleValues: ['A', 'B', 'C', 'D'], inferredType: 'text' },
      { colLetter: 'D', colIndex: 4, name: 'period', sampleValues: ['1', '2', '3'], inferredType: 'integer' },
      { colLetter: 'E', colIndex: 5, name: 'subject', sampleValues: ['Sci', 'Math', 'Tam', 'Eng'], inferredType: 'text' },
      { colLetter: 'F', colIndex: 6, name: 'primary_teacher', sampleValues: ['Miss.T.Anoja', 'Mr.B.Ullasanan'], inferredType: 'text' },
      { colLetter: 'G', colIndex: 7, name: 'secondary_teacher', sampleValues: ['Mrs.N.Sasikumar'], inferredType: 'text' },
      { colLetter: 'H', colIndex: 8, name: 'raw_teacher_cell', sampleValues: ['Miss.T.Anoja'], inferredType: 'text' }
    ];

    const sampleRows = records.map((rec, idx) => ({
      rowNumber: idx + 1,
      data: {
        A: rec.day_of_week,
        day_of_week: rec.day_of_week,
        B: rec.grade,
        grade: rec.grade,
        C: rec.division,
        division: rec.division,
        D: rec.period,
        period: rec.period,
        E: rec.subject,
        subject: rec.subject,
        F: rec.primary_teacher,
        primary_teacher: rec.primary_teacher,
        G: rec.secondary_teacher || '',
        secondary_teacher: rec.secondary_teacher || '',
        H: rec.raw_teacher_cell,
        raw_teacher_cell: rec.raw_teacher_cell
      }
    }));

    const sheetAnalysis: SheetAnalysis = {
      sheetName: 'school_timetables',
      totalRows: records.length,
      totalColumns: headers.length,
      usedRange: `A1:H${records.length + 1}`,
      mergedRanges: [],
      detectedHeaderRow: 1,
      detectedDataStartRow: 2,
      candidateHeaderRows: [{ row: 1, headers: headers.map(h => h.name), confidence: 1.0 }],
      titleRows: [],
      sectionHeadings: [],
      emptyRowsCount: 0,
      repeatedHeadersCount: 0,
      headers,
      sampleRows
    };

    return {
      analysis: sheetAnalysis,
      records
    };
  }

  /**
   * Synthesizes a clean normalized WorkbookAnalysis from a staff allocation pivot matrix
   */
  static createNormalizedStaffAllocationAnalysis(
    rawWb: XLSX.WorkBook,
    originalFilename: string
  ): {
    analysis: SheetAnalysis;
    records: NormalizedTeacherAllocationRecord[];
  } {
    const { records, summary } = this.unpivotStaffAllocationWorkbook(rawWb);

    const headers: SheetHeader[] = [
      { colLetter: 'A', colIndex: 1, name: 'academic_level', sampleValues: summary.grades.slice(0, 3), inferredType: 'text' },
      { colLetter: 'B', colIndex: 2, name: 'subject', sampleValues: ['Mathematics', 'Science', 'Religion'], inferredType: 'text' },
      { colLetter: 'C', colIndex: 3, name: 'division', sampleValues: ['A', 'B', 'C', 'D'], inferredType: 'text' },
      { colLetter: 'D', colIndex: 4, name: 'primary_teacher', sampleValues: ['Mr.B.Ullasanan', 'Mrs.S.Vakeeswaran'], inferredType: 'text' },
      { colLetter: 'E', colIndex: 5, name: 'secondary_teacher', sampleValues: ['Mrs.L.Mathiyamuthan'], inferredType: 'text' },
      { colLetter: 'F', colIndex: 6, name: 'raw_cell_value', sampleValues: ['Mr.B.Ullasanan'], inferredType: 'text' }
    ];

    const sampleRows = records.map((rec, idx) => ({
      rowNumber: idx + 1,
      data: {
        A: rec.academic_level,
        academic_level: rec.academic_level,
        B: rec.subject,
        subject: rec.subject,
        C: rec.division,
        division: rec.division,
        D: rec.primary_teacher,
        primary_teacher: rec.primary_teacher,
        E: rec.secondary_teacher || '',
        secondary_teacher: rec.secondary_teacher || '',
        F: rec.raw_cell_value,
        raw_cell_value: rec.raw_cell_value
      }
    }));

    const sheetAnalysis: SheetAnalysis = {
      sheetName: 'teacher_subject_assignments',
      totalRows: records.length,
      totalColumns: headers.length,
      usedRange: `A1:F${records.length + 1}`,
      mergedRanges: [],
      detectedHeaderRow: 1,
      detectedDataStartRow: 2,
      candidateHeaderRows: [{ row: 1, headers: headers.map(h => h.name), confidence: 1.0 }],
      titleRows: [],
      sectionHeadings: [],
      emptyRowsCount: 0,
      repeatedHeadersCount: 0,
      headers,
      sampleRows
    };

    return {
      analysis: sheetAnalysis,
      records
    };
  }
}
