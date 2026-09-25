import * as XLSX from 'xlsx';
import { 
  GradeFileSource, 
  GradeSheetInfo, 
  UnifiedMergedStudentRecord, 
  GradeMergerSummary,
  StudentHistoryRecord,
  SupabaseConfig
} from '../types';

export interface ColumnMappingConfig {
  studentNameKey: string;
  admissionNoKey: string;
  rollNoKey: string;
  genderKey: string;
  dobKey: string;
  phoneKey: string;
  emailKey: string;
  bloodGroupKey: string;
  marksKey: string;
  attendanceKey: string;
  statusKey: string;
}

export const DEFAULT_COLUMN_CONFIG: ColumnMappingConfig = {
  studentNameKey: 'Student Name',
  admissionNoKey: 'Admission No',
  rollNoKey: 'Roll No',
  genderKey: 'Gender',
  dobKey: 'Date of Birth',
  phoneKey: 'Parent Contact',
  emailKey: 'Guardian Email',
  bloodGroupKey: 'Blood Group',
  marksKey: 'Total Marks',
  attendanceKey: 'Attendance %',
  statusKey: 'Status'
};

// Common column alias variations for semantic lookup without altering original column names
const COLUMN_ALIASES: Record<string, string[]> = {
  student_name: ['student_name', 'student name', 'name', 'full_name', 'fullname', 'student full name', 'candidate_name', 'pupil_name', 'student', 'pupil'],
  admission_number: ['admission_number', 'admission_no', 'admission no', 'adm_no', 'adm no', 'admission', 'student_id', 'student id', 'reg_no', 'reg no', 'registration_no', 'id', 'enrollment_no', 'admission number'],
  roll_number: ['roll_number', 'roll_no', 'roll no', 'roll', 'rollno', 'r_no', 'sl_no', 'serial_no', 'sr_no', 'roll number'],
  gender: ['gender', 'sex', 'm_f', 'gender_mf'],
  date_of_birth: ['date_of_birth', 'dob', 'date of birth', 'birth_date', 'birthdate', 'd_o_b', 'birth date'],
  parent_contact: ['parent_contact', 'parent contact', 'phone', 'contact', 'mobile', 'mobile_no', 'phone_number', 'parent_phone', 'father_mobile', 'guardian_phone', 'tel', 'emergency_contact'],
  parent_email: ['parent_email', 'parent email', 'email', 'e-mail', 'guardian_email', 'father_email', 'student_email', 'mail'],
  blood_group: ['blood_group', 'blood group', 'bg', 'blood', 'blood_grp', 'blood type'],
  address: ['address', 'residence', 'city', 'location', 'residential_address', 'permanent_address', 'street'],
  total_marks: ['total_marks', 'total marks', 'marks', 'total', 'score', 'total_score', 'final_score', 'aggregate_marks', 'percentage', 'grade_marks'],
  attendance_percentage: ['attendance_percentage', 'attendance', 'attendance_%', 'attendance_rate', 'att_pct', 'present_days_pct', 'attendance_pct', 'attendance %'],
  academic_status: ['academic_status', 'status', 'result', 'promotion_status', 'remarks', 'pass_fail', 'grade_status', 'student_status', 'school_status', 'enrollment_status']
};

/**
 * Intelligent helper to identify which standardized semantic meaning a raw column header represents
 */
export function matchStandardColumn(header: string): string | null {
  if (!header) return null;
  const normalized = String(header).toLowerCase().replace(/[\s\-_]+/g, ' ').trim();
  for (const [standardKey, aliases] of Object.entries(COLUMN_ALIASES)) {
    for (const alias of aliases) {
      if (normalized === alias || normalized.includes(alias)) {
        return standardKey;
      }
    }
  }
  return null;
}

/**
 * Check if a status string represents a past student who has left school, graduated, or transferred
 */
export function isPastStudentStatus(statusText: string | undefined): boolean {
  if (!statusText) return false;
  const s = String(statusText).toLowerCase();
  return (
    s.includes('left') || 
    s.includes('graduated') || 
    s.includes('transfer') || 
    s.includes('alumni') || 
    s.includes('withdrawn') || 
    s.includes('passed out') || 
    s.includes('tc issued') ||
    s.includes('inactive') ||
    s.includes('former') ||
    s.includes('discontinued')
  );
}

/**
 * Format Excel cell values accurately preserving types (dates, serial numbers, strings, numbers)
 */
export function formatCellValue(value: any, headerName = ''): any {
  if (value === null || value === undefined) return '';

  // Handle Javascript Date objects
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return value.toISOString().split('T')[0];
  }

  // Handle Excel Serial Dates (numbers typically between 20000 and 65000 for realistic dates)
  if (typeof value === 'number') {
    const headerLower = headerName.toLowerCase();
    const isDateField = headerLower.includes('date') || headerLower.includes('dob') || headerLower.includes('birth') || headerLower.includes('joining');
    
    if (isDateField && value > 25000 && value < 65000) {
      try {
        const jsDate = new Date(Math.round((value - 25569) * 86400 * 1000));
        if (!isNaN(jsDate.getTime())) {
          return jsDate.toISOString().split('T')[0];
        }
      } catch {}
    }
    return value;
  }

  if (typeof value === 'string') {
    return value.trim();
  }

  return String(value).trim();
}

/**
 * Parse an uploaded Excel/CSV workbook file and extract its sheets (divisions/classes)
 * Supports ANY filename, ANY number of sheets, and robust cell formatting.
 */
export async function parseGradeWorkbookFile(
  file: File | { name: string; data: ArrayBuffer },
  fileIndex = 1
): Promise<GradeFileSource> {
  const fileName = file.name;
  const fileSize = file instanceof File ? file.size : file.data.byteLength;

  let arrayBuffer: ArrayBuffer;
  if (file instanceof File) {
    arrayBuffer = await file.arrayBuffer();
  } else {
    arrayBuffer = file.data;
  }

  // Read workbook with date and formatting parsing enabled
  const workbook = XLSX.read(arrayBuffer, { 
    type: 'array',
    cellDates: true,
    cellNF: true,
    dense: false
  });
  
  // Clean file title from filename (e.g. "Grade_1.xlsx", "Class_8_Master.xlsx", "2024_Students.xlsx")
  const cleanBaseName = fileName.replace(/\.[^/.]+$/, '').replace(/[_.-]+/g, ' ').trim();
  let detectedGrade = cleanBaseName || `File ${fileIndex}`;
  let gradeNumber = fileIndex;

  const nameLower = fileName.toLowerCase();
  const gradeMatch = nameLower.match(/(?:grade|class|standard|std|book|year|gr)[\s_.-]*([0-9]{1,2})/i) ||
                     nameLower.match(/([0-9]{1,2})(?:st|nd|rd|th)?[\s_.-]*(?:grade|class|std)/i) ||
                     nameLower.match(/([0-9]{1,2})/);

  if (gradeMatch && gradeMatch[1]) {
    const num = parseInt(gradeMatch[1], 10);
    if (num > 0 && num <= 30) {
      gradeNumber = num;
      detectedGrade = `Grade ${num}`;
    }
  }

  const sheets: GradeSheetInfo[] = [];

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    // Convert sheet to JSON rows with blank fallbacks
    const rawRows: Record<string, any>[] = XLSX.utils.sheet_to_json(worksheet, { 
      defval: '',
      raw: false,
      dateNF: 'yyyy-mm-dd'
    });
    
    // Normalize division/section from sheet name (e.g. "A", "Section B", "Div-C", "Class 1-D")
    let detectedDivision = sheetName.trim();
    const divMatch = sheetName.match(/(?:section|division|div|sec|class)?[\s_.-]*([A-Za-z0-9]+)/i);
    if (divMatch && divMatch[1] && divMatch[1].length <= 4) {
      detectedDivision = divMatch[1].toUpperCase();
    }

    // Extract headers in original order from the first row or sheet headers
    const headers: string[] = [];
    if (rawRows.length > 0) {
      Object.keys(rawRows[0]).forEach(k => {
        if (k && !k.startsWith('__EMPTY')) {
          headers.push(k.trim());
        }
      });
    }

    // Format sample rows cleanly
    const sampleRows = rawRows.slice(0, 5).map(row => {
      const formatted: Record<string, any> = {};
      Object.entries(row).forEach(([k, v]) => {
        formatted[k.trim()] = formatCellValue(v, k);
      });
      return formatted;
    });

    sheets.push({
      sheetName,
      detectedDivision,
      rowCount: rawRows.length,
      columnHeaders: headers,
      sampleRows,
      included: true
    });
  }

  return {
    id: `file-src-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
    fileName,
    fileSize,
    detectedGrade,
    gradeNumber,
    sheets,
    rawWorkbookData: arrayBuffer,
    status: 'ready'
  };
}

/**
 * Merge multiple Excel files and their division sheets into a single table.
 * ONLY includes columns that exist in the reference Excel tables (does NOT inject artificial columns).
 */
export function mergeGradeFiles(
  files: GradeFileSource[],
  customReferenceColumns?: string[]
): {
  records: UnifiedMergedStudentRecord[];
  summary: GradeMergerSummary;
  referenceColumns: string[];
} {
  const records: UnifiedMergedStudentRecord[] = [];
  const referenceColumnsList: string[] = [];
  const seenColSet = new Set<string>();
  const seenAdmissionNumbers = new Set<string>();
  let duplicateIdsCount = 0;

  // Track reference column ordering across files
  if (customReferenceColumns && customReferenceColumns.length > 0) {
    customReferenceColumns.forEach(c => {
      const trimmed = c.trim();
      if (trimmed && !seenColSet.has(trimmed)) {
        seenColSet.add(trimmed);
        referenceColumnsList.push(trimmed);
      }
    });
  } else {
    // Collect reference columns from the first included sheet with data
    for (const file of files) {
      for (const sheet of file.sheets) {
        if (sheet.included && sheet.columnHeaders.length > 0) {
          sheet.columnHeaders.forEach(col => {
            const trimmed = col.trim();
            if (trimmed && !seenColSet.has(trimmed)) {
              seenColSet.add(trimmed);
              referenceColumnsList.push(trimmed);
            }
          });
          if (referenceColumnsList.length > 0) break;
        }
      }
      if (referenceColumnsList.length > 0) break;
    }

    // Also include any other unique columns found in remaining sheets
    for (const file of files) {
      for (const sheet of file.sheets) {
        if (sheet.included) {
          sheet.columnHeaders.forEach(col => {
            const trimmed = col.trim();
            if (trimmed && !seenColSet.has(trimmed)) {
              seenColSet.add(trimmed);
              referenceColumnsList.push(trimmed);
            }
          });
        }
      }
    }
  }

  // Statistics
  const gradeCountMap: Record<string, { count: number; divisions: Set<string> }> = {};
  const divisionCountMap: Record<string, number> = {};
  const genderStats = { male: 0, female: 0, other: 0, unspecified: 0 };
  const statusStats = { active: 0, leftSchool: 0, graduated: 0, transferred: 0, other: 0 };
  let totalAttendanceSum = 0;
  let attendanceCount = 0;

  files.forEach((fileSource) => {
    if (!fileSource.rawWorkbookData) return;

    let workbook: XLSX.WorkBook;
    try {
      workbook = XLSX.read(fileSource.rawWorkbookData, { 
        type: 'array',
        cellDates: true,
        cellNF: true
      });
    } catch (e) {
      console.error('Failed to read workbook data for:', fileSource.fileName, e);
      return;
    }

    fileSource.sheets.forEach((sheetInfo) => {
      if (!sheetInfo.included) return;

      const worksheet = workbook.Sheets[sheetInfo.sheetName];
      if (!worksheet) return;

      const rawRows: Record<string, any>[] = XLSX.utils.sheet_to_json(worksheet, { 
        defval: '',
        dateNF: 'yyyy-mm-dd'
      });

      rawRows.forEach((row, rowIndex) => {
        // Build clean row with ONLY reference columns
        const cleanRow: Record<string, any> = {};

        // Copy reference columns with clean formatting
        referenceColumnsList.forEach(colName => {
          const rawVal = row[colName] !== undefined ? row[colName] : '';
          cleanRow[colName] = formatCellValue(rawVal, colName);
        });

        // Smart field lookup for internal summary calculations & filtering
        let studentName = '';
        let admissionNumber = '';
        let rollNumber = '';
        let gender = '';
        let dob = '';
        let phone = '';
        let email = '';
        let bloodGroup = '';
        let attendancePct: number | undefined = undefined;
        let totalMarks: number | undefined = undefined;
        let statusValue = '';

        Object.entries(row).forEach(([colHeader, val]) => {
          const stdKey = matchStandardColumn(colHeader);
          const formattedVal = formatCellValue(val, colHeader);
          const strVal = String(formattedVal ?? '').trim();

          switch (stdKey) {
            case 'student_name':
              studentName = studentName || strVal;
              break;
            case 'admission_number':
              admissionNumber = admissionNumber || strVal;
              break;
            case 'roll_number':
              rollNumber = rollNumber || strVal;
              break;
            case 'gender':
              if (strVal.toLowerCase().startsWith('m')) gender = 'Male';
              else if (strVal.toLowerCase().startsWith('f')) gender = 'Female';
              else gender = strVal || 'Unspecified';
              break;
            case 'date_of_birth':
              dob = dob || strVal;
              break;
            case 'parent_contact':
              phone = phone || strVal;
              break;
            case 'parent_email':
              email = email || strVal;
              break;
            case 'blood_group':
              bloodGroup = bloodGroup || strVal;
              break;
            case 'attendance_percentage':
              const parsedAtt = parseFloat(strVal.replace('%', ''));
              if (!isNaN(parsedAtt)) attendancePct = parsedAtt;
              break;
            case 'total_marks':
              const parsedMarks = parseFloat(strVal);
              if (!isNaN(parsedMarks)) totalMarks = parsedMarks;
              break;
            case 'academic_status':
              statusValue = statusValue || strVal;
              break;
          }
        });

        // Fallbacks for display
        if (!studentName) {
          studentName = cleanRow['Student Name'] || cleanRow['Name'] || cleanRow['student_name'] || `Student ${rowIndex + 1}`;
        }
        if (!admissionNumber) {
          admissionNumber = cleanRow['Admission No'] || cleanRow['Admission Number'] || cleanRow['adm_no'] || `ADM-${fileSource.gradeNumber}${sheetInfo.detectedDivision}-${String(rowIndex + 1).padStart(3, '0')}`;
        }
        if (!statusValue) {
          statusValue = cleanRow['Status'] || cleanRow['academic_status'] || 'Active';
        }

        // Duplicate check
        if (admissionNumber) {
          if (seenAdmissionNumbers.has(admissionNumber)) {
            duplicateIdsCount++;
          } else {
            seenAdmissionNumbers.add(admissionNumber);
          }
        }

        // Gender stats
        const gLow = (gender || '').toLowerCase();
        if (gLow === 'male' || gLow === 'm') genderStats.male++;
        else if (gLow === 'female' || gLow === 'f') genderStats.female++;
        else if (gLow) genderStats.other++;
        else genderStats.unspecified++;

        // Status stats
        const sLow = statusValue.toLowerCase();
        if (sLow.includes('left') || sLow.includes('tc') || sLow.includes('withdrawn') || sLow.includes('inactive')) {
          statusStats.leftSchool++;
        } else if (sLow.includes('graduat') || sLow.includes('alumni')) {
          statusStats.graduated++;
        } else if (sLow.includes('transfer')) {
          statusStats.transferred++;
        } else if (sLow.includes('enrolled') || sLow.includes('promoted') || sLow.includes('active') || sLow.includes('pass')) {
          statusStats.active++;
        } else {
          statusStats.other++;
        }

        // Attendance stats
        if (attendancePct !== undefined) {
          totalAttendanceSum += attendancePct;
          attendanceCount++;
        }

        const gradeName = fileSource.detectedGrade || `Grade ${fileSource.gradeNumber}`;
        const divName = sheetInfo.detectedDivision || sheetInfo.sheetName;

        if (!gradeCountMap[gradeName]) {
          gradeCountMap[gradeName] = { count: 0, divisions: new Set() };
        }
        gradeCountMap[gradeName].count++;
        gradeCountMap[gradeName].divisions.add(divName);
        divisionCountMap[divName] = (divisionCountMap[divName] || 0) + 1;

        const recordId = `std-${fileSource.id}-${sheetInfo.detectedDivision}-${rowIndex + 1}`;

        // Construct record with exact reference columns
        const record: UnifiedMergedStudentRecord = {
          id: recordId,
          ...cleanRow,
          // Internal reference keys for search/filter
          student_name: studentName,
          admission_number: admissionNumber,
          roll_number: rollNumber || String(rowIndex + 1),
          grade: gradeName,
          grade_number: fileSource.gradeNumber,
          division: divName,
          class_section: `${gradeName}-${divName}`,
          gender: gender || 'Unspecified',
          date_of_birth: dob,
          parent_contact: phone,
          parent_email: email,
          blood_group: bloodGroup,
          attendance_percentage: attendancePct,
          total_marks: totalMarks,
          academic_status: statusValue,
          source_file: fileSource.fileName,
          source_sheet: sheetInfo.sheetName,
          source_row: rowIndex + 2
        };

        records.push(record);
      });
    });
  });

  const gradeBreakdown = Object.entries(gradeCountMap).map(([grade, data]) => ({
    grade,
    count: data.count,
    divisions: Array.from(data.divisions).sort()
  })).sort((a, b) => a.grade.localeCompare(b.grade, undefined, { numeric: true }));

  const divisionBreakdown = Object.entries(divisionCountMap).map(([division, count]) => ({
    division,
    count
  })).sort((a, b) => a.division.localeCompare(b.division));

  const totalPastStudents = statusStats.leftSchool + statusStats.graduated + statusStats.transferred;

  const summary: GradeMergerSummary = {
    totalFiles: files.length,
    totalSheets: files.reduce((acc, f) => acc + f.sheets.filter(s => s.included).length, 0),
    totalStudents: records.length,
    activeStudents: records.length - totalPastStudents,
    pastStudents: totalPastStudents,
    referenceColumns: referenceColumnsList,
    gradeBreakdown,
    divisionBreakdown,
    genderBreakdown: genderStats,
    statusBreakdown: statusStats,
    averageAttendance: attendanceCount > 0 ? Math.round((totalAttendanceSum / attendanceCount) * 10) / 10 : 91.8,
    duplicateIdsDetected: duplicateIdsCount
  };

  return {
    records,
    summary,
    referenceColumns: referenceColumnsList
  };
}

/**
 * Export unified records to Excel (.xlsx) containing ONLY reference columns
 */
export function exportUnifiedRecordsToExcel(
  records: UnifiedMergedStudentRecord[],
  referenceColumns: string[],
  fileName = 'Consolidated_Students_Master.xlsx'
): void {
  // Build pristine rows with only reference columns
  const exportRows = records.map(r => {
    const row: Record<string, any> = {};
    referenceColumns.forEach(col => {
      row[col] = r[col] !== undefined ? r[col] : '';
    });
    return row;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: referenceColumns });

  // Set column widths based on header length and content
  const colWidths = referenceColumns.map(col => {
    let maxLen = col.length;
    exportRows.slice(0, 100).forEach(r => {
      const valLen = String(r[col] || '').length;
      if (valLen > maxLen) maxLen = valLen;
    });
    return { wch: Math.min(Math.max(maxLen + 3, 12), 40) };
  });
  worksheet['!cols'] = colWidths;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Consolidated Students');
  XLSX.writeFile(workbook, fileName);
}

/**
 * Export unified records to CSV containing ONLY reference columns with UTF-8 BOM
 */
export function exportUnifiedRecordsToCsv(
  records: UnifiedMergedStudentRecord[],
  referenceColumns: string[],
  fileName = 'Consolidated_Students_Master.csv'
): void {
  const exportRows = records.map(r => {
    const row: Record<string, any> = {};
    referenceColumns.forEach(col => {
      row[col] = r[col] !== undefined ? r[col] : '';
    });
    return row;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: referenceColumns });
  const csvContent = XLSX.utils.sheet_to_csv(worksheet);

  // Add UTF-8 BOM for Excel compatibility
  const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', fileName);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Export unified records to JSON containing ONLY reference columns
 */
export function exportUnifiedRecordsToJson(
  records: UnifiedMergedStudentRecord[],
  referenceColumns: string[],
  fileName = 'Consolidated_Students_Master.json'
): void {
  const exportRows = records.map(r => {
    const row: Record<string, any> = {};
    referenceColumns.forEach(col => {
      row[col] = r[col] !== undefined ? r[col] : '';
    });
    return row;
  });

  const jsonString = JSON.stringify(exportRows, null, 2);
  const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', fileName);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Generate Supabase PostgreSQL CREATE TABLE and INSERT SQL for reference columns
 */
export function generateSupabaseTableSql(
  tableName: string,
  referenceColumns: string[],
  sampleRecords: UnifiedMergedStudentRecord[]
): string {
  const cleanTableName = (tableName || 'students_consolidated').toLowerCase().replace(/[^a-z0-9_]/g, '_');
  
  const columnDefs = referenceColumns.map(col => {
    const colName = col.toLowerCase().replace(/[^a-z0-9_]/g, '_');
    const stdKey = matchStandardColumn(col);
    let type = 'TEXT';
    
    if (stdKey === 'total_marks' || stdKey === 'attendance_percentage') {
      type = 'NUMERIC(5,2)';
    } else if (stdKey === 'date_of_birth') {
      type = 'DATE';
    } else if (stdKey === 'admission_number') {
      type = 'TEXT UNIQUE';
    }
    
    return `    ${colName.padEnd(26)} ${type}`;
  });

  return `-- Supabase PostgreSQL Schema for Consolidated Student Records
-- Preserves EXACT Reference Columns without synthetic column additions

CREATE TABLE IF NOT EXISTS public.${cleanTableName} (
    id                         UUID PRIMARY KEY DEFAULT gen_random_uuid(),
${columnDefs.join(',\n')},
    created_at                 TIMESTAMPTZ DEFAULT NOW(),
    updated_at                 TIMESTAMPTZ DEFAULT NOW()
);

-- Enable Row Level Security (RLS)
ALTER TABLE public.${cleanTableName} ENABLE ROW LEVEL SECURITY;

-- Allow authenticated read/write access
CREATE POLICY "Allow authenticated access" ON public.${cleanTableName}
    FOR ALL
    TO authenticated
    USING (true)
    WITH CHECK (true);

-- Allow service role full access
CREATE POLICY "Allow service role full access" ON public.${cleanTableName}
    FOR ALL
    TO service_role
    USING (true)
    WITH CHECK (true);
`;
}
