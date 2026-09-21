import * as XLSX from 'xlsx';
import { WorkbookAnalysis, SheetAnalysis, WorksheetMapping } from '../types';

export function createComplexSampleWorkbook(): { workbook: XLSX.WorkBook; binaryData: Uint8Array; filename: string } {
  const wb = XLSX.utils.book_new();

  // 1. Sheet: Student Details
  const studentDetailsData = [
    // Row 1 (Index 0): Merged Title
    ["ST. JUDE ACADEMY - STUDENT INFORMATION", "", "", "", "", "", "", ""],
    // Row 2 (Index 1): Academic Metadata
    ["Academic Year: 2025-2026", "Status: Official", "Source: TrueNAS Nextcloud", "", "", "", "", ""],
    // Row 3 (Index 2): Section Heading Class 10A (Merged A3:H3)
    ["CLASS 10A", "", "", "", "", "", "", ""],
    // Row 4 (Index 3): Blank row
    ["", "", "", "", "", "", "", ""],
    // Row 5 (Index 4): Header Row
    ["Student ID", "Student Name", "Gender", "DOB", "Phone", "Emergency Contact", "Status", "Remarks"],
    // Rows 6-12: Class 10A Students
    ["STU-1001", "Alice Smith", "Female", "2008-04-12", "+1 (555) 019-2831", "+1 555 019 9999", "Active", "Honors student"],
    ["STU-1002", "Bob Jones", "Male", "2008-09-21", "+1 (555) 018-4422", "+1 555 018 8888", "Active", "Locker #42"],
    ["STU-1003", "Charlie Davis", "Male", "2007-12-05", "555-017-3311", "555-017-0000", "Active", "Transfer student"],
    ["STU-1004", "Diana Prince", "Female", "2008-06-18", "+1 (555) 016-5544", "+1 555 016 7777", "Active", "Student Council"],
    ["STU-1005", "Ethan Hunt", "Male", "2008-01-30", "+1 (555) 015-6677", "+1 555 015 2222", "Active", "Track & Field"],
    ["STU-1006", "Fiona Gallagher", "Female", "2008-11-14", "+1 (555) 014-9988", "+1 555 014 3333", "Active", "Science Club"],
    ["STU-1007", "George Clark", "Male", "2008-03-08", "+1 (555) 013-1122", "+1 555 013 4444", "Active", "Math Team"],
    // Row 13 (Index 12): Section Heading Class 10B (Merged A13:H13)
    ["CLASS 10B", "", "", "", "", "", "", ""],
    // Row 14 (Index 13): Blank row
    ["", "", "", "", "", "", "", ""],
    // Rows 15-20: Class 10B Students
    ["STU-1008", "Hannah Abbott", "Female", "2008-07-22", "+1 (555) 012-7788", "+1 555 012 5555", "Active", "Choir"],
    ["STU-1009", "Ian Malcolm", "Male", "2007-10-15", "+1 (555) 011-3344", "+1 555 011 6666", "Active", "Debate Lead"],
    ["STU-1010", "Julia Roberts", "Female", "2008-02-19", "+1 (555) 010-5566", "+1 555 010 7777", "Active", "Art Club"],
    ["STU-1011", "Kevin Bacon", "Male", "2008-08-25", "+1 (555) 009-8899", "+1 555 009 8888", "Active", "Drama Club"],
    ["STU-1012", "Luna Lovegood", "Female", "2008-05-02", "+1 (555) 008-4455", "+1 555 008 9999", "Active", "Newspaper Editor"],
  ];

  const wsDetails = XLSX.utils.aoa_to_sheet(studentDetailsData);
  // Merges: A1:H1 (Title), A3:H3 (Class 10A), A13:H13 (Class 10B)
  wsDetails['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, // A1:H1
    { s: { r: 2, c: 0 }, e: { r: 2, c: 7 } }, // A3:H3
    { s: { r: 12, c: 0 }, e: { r: 12, c: 7 } } // A13:H13
  ];
  XLSX.utils.book_append_sheet(wb, wsDetails, "Student Details");

  // 2. Sheet: Attendance
  const attendanceData = [
    ["Student ID", "Date", "Status", "Remarks"],
    ["STU-1001", "2025-09-15", "P", "On time"],
    ["STU-1002", "2025-09-15", "A", "Excused medical leave"],
    ["STU-1003", "2025-09-15", "L", "Bus delay"],
    ["STU-1004", "2025-09-15", "P", "On time"],
    ["STU-1005", "2025-09-15", "P", "On time"],
    ["STU-1006", "2025-09-15", "P", "On time"],
    ["STU-1007", "2025-09-15", "A", "Unexcused absence"],
    ["STU-1008", "2025-09-15", "P", "On time"],
    ["STU-1009", "2025-09-15", "P", "On time"],
    ["STU-1010", "2025-09-15", "P", "On time"],
    ["STU-1011", "2025-09-15", "L", "Family appointment"],
    ["STU-1012", "2025-09-15", "P", "On time"]
  ];
  const wsAttendance = XLSX.utils.aoa_to_sheet(attendanceData);
  XLSX.utils.book_append_sheet(wb, wsAttendance, "Attendance");

  // 3. Sheet: Sports
  const sportsData = [
    ["Student ID", "Sport Name", "Position", "Medical Clearance"],
    ["STU-1001", "Basketball", "Point Guard", "Yes"],
    ["STU-1004", "Soccer", "Midfielder", "Yes"],
    ["STU-1005", "Track & Field", "Sprinter", "Yes"],
    ["STU-1007", "Swimming", "Freestyle 100m", "Yes"],
    ["STU-1009", "Tennis", "Singles", "No"],
    ["STU-1011", "Volleyball", "Setter", "Yes"]
  ];
  const wsSports = XLSX.utils.aoa_to_sheet(sportsData);
  XLSX.utils.book_append_sheet(wb, wsSports, "Sports");

  // 4. Sheet: Medical
  const medicalData = [
    ["Student ID", "Blood Group", "Allergies", "Doctor Contact", "Last Checkup"],
    ["STU-1001", "O+", "Peanuts", "Dr. Wilson (555-0100)", "2025-08-20"],
    ["STU-1002", "A+", "None", "Dr. Gomez (555-0200)", "2025-08-15"],
    ["STU-1003", "B+", "Penicillin", "Dr. Adams (555-0300)", "2025-07-30"],
    ["STU-1004", "AB+", "Latex", "Dr. Chen (555-0400)", "2025-08-10"],
    ["STU-1005", "O-", "None", "Dr. Vance (555-0500)", "2025-09-01"],
    ["STU-1006", "A-", "Bee Stings", "Dr. Taylor (555-0600)", "2025-08-25"]
  ];
  const wsMedical = XLSX.utils.aoa_to_sheet(medicalData);
  XLSX.utils.book_append_sheet(wb, wsMedical, "Medical");

  // 5. Sheet: Results
  const resultsData = [
    ["Student ID", "Term", "Mathematics", "Science", "English", "Total Score", "Grade"],
    ["STU-1001", "Term 1", 94.5, 88.0, 92.0, 274.5, "A"],
    ["STU-1002", "Term 1", 76.0, 82.5, 79.0, 237.5, "B"],
    ["STU-1003", "Term 1", 85.0, 90.0, 84.5, 259.5, "A"],
    ["STU-1004", "Term 1", 91.0, 95.0, 89.0, 275.0, "A"],
    ["STU-1005", "Term 1", 68.0, 72.0, 74.5, 214.5, "C"],
    ["STU-1006", "Term 1", 99.0, 97.5, 95.0, 291.5, "A+"]
  ];
  const wsResults = XLSX.utils.aoa_to_sheet(resultsData);
  XLSX.utils.book_append_sheet(wb, wsResults, "Results");

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const binaryData = new Uint8Array(wbout);

  return {
    workbook: wb,
    binaryData,
    filename: "students_complex.xlsx"
  };
}

export function getDefaultSampleMappings(workbookName: string = "students_complex.xlsx"): WorksheetMapping[] {
  return [
    {
      id: "wm-1",
      workbookName,
      worksheetName: "Student Details",
      supabaseTable: "students",
      headerRow: 5,
      dataStartRow: 6,
      sectionHeadingTargetCol: "class",
      enabled: true,
      columns: [
        { id: "cm-1", excelColumn: "A", excelHeader: "Student ID", supabaseColumn: "student_number", dataType: "text", required: true, uniqueKey: true, transformation: "normalize_id" },
        { id: "cm-2", excelColumn: "B", excelHeader: "Student Name", supabaseColumn: "name", dataType: "text", required: true, uniqueKey: false, transformation: "trim" },
        { id: "cm-3", excelColumn: "C", excelHeader: "Gender", supabaseColumn: "gender", dataType: "text", required: false, uniqueKey: false, transformation: "trim" },
        { id: "cm-4", excelColumn: "D", excelHeader: "DOB", supabaseColumn: "dob", dataType: "date", required: false, uniqueKey: false, transformation: "parse_date" },
        { id: "cm-5", excelColumn: "E", excelHeader: "Phone", supabaseColumn: "phone", dataType: "text", required: false, uniqueKey: false, transformation: "normalize_phone" },
        { id: "cm-6", excelColumn: "F", excelHeader: "Emergency Contact", supabaseColumn: "emergency_contact", dataType: "text", required: false, uniqueKey: false, transformation: "trim" },
        { id: "cm-7", excelColumn: "G", excelHeader: "Status", supabaseColumn: "status", dataType: "text", required: false, uniqueKey: false, defaultValue: "Active", transformation: "trim" }
      ]
    },
    {
      id: "wm-2",
      workbookName,
      worksheetName: "Attendance",
      supabaseTable: "attendance",
      headerRow: 1,
      dataStartRow: 2,
      enabled: true,
      columns: [
        { id: "cm-8", excelColumn: "A", excelHeader: "Student ID", supabaseColumn: "student_number", dataType: "text", required: true, uniqueKey: true, transformation: "normalize_id" },
        { id: "cm-9", excelColumn: "B", excelHeader: "Date", supabaseColumn: "attendance_date", dataType: "date", required: true, uniqueKey: true, transformation: "parse_date" },
        { id: "cm-10", excelColumn: "C", excelHeader: "Status", supabaseColumn: "status", dataType: "text", required: true, uniqueKey: false, transformation: "pa_to_status" },
        { id: "cm-11", excelColumn: "D", excelHeader: "Remarks", supabaseColumn: "remarks", dataType: "text", required: false, uniqueKey: false, transformation: "trim" }
      ]
    },
    {
      id: "wm-3",
      workbookName,
      worksheetName: "Sports",
      supabaseTable: "sports",
      headerRow: 1,
      dataStartRow: 2,
      enabled: true,
      columns: [
        { id: "cm-12", excelColumn: "A", excelHeader: "Student ID", supabaseColumn: "student_number", dataType: "text", required: true, uniqueKey: true, transformation: "normalize_id" },
        { id: "cm-13", excelColumn: "B", excelHeader: "Sport Name", supabaseColumn: "sport_name", dataType: "text", required: true, uniqueKey: true, transformation: "trim" },
        { id: "cm-14", excelColumn: "C", excelHeader: "Position", supabaseColumn: "position", dataType: "text", required: false, uniqueKey: false, transformation: "trim" },
        { id: "cm-15", excelColumn: "D", excelHeader: "Medical Clearance", supabaseColumn: "medical_clearance", dataType: "boolean", required: false, uniqueKey: false, transformation: "yes_no_to_boolean" }
      ]
    },
    {
      id: "wm-4",
      workbookName,
      worksheetName: "Medical",
      supabaseTable: "medical",
      headerRow: 1,
      dataStartRow: 2,
      enabled: true,
      columns: [
        { id: "cm-16", excelColumn: "A", excelHeader: "Student ID", supabaseColumn: "student_number", dataType: "text", required: true, uniqueKey: true, transformation: "normalize_id" },
        { id: "cm-17", excelColumn: "B", excelHeader: "Blood Group", supabaseColumn: "blood_group", dataType: "text", required: false, uniqueKey: false, transformation: "trim" },
        { id: "cm-18", excelColumn: "C", excelHeader: "Allergies", supabaseColumn: "allergies", dataType: "text", required: false, uniqueKey: false, transformation: "trim" },
        { id: "cm-19", excelColumn: "D", excelHeader: "Doctor Contact", supabaseColumn: "doctor_contact", dataType: "text", required: false, uniqueKey: false, transformation: "trim" },
        { id: "cm-20", excelColumn: "E", excelHeader: "Last Checkup", supabaseColumn: "last_checkup_date", dataType: "date", required: false, uniqueKey: false, transformation: "parse_date" }
      ]
    },
    {
      id: "wm-5",
      workbookName,
      worksheetName: "Results",
      supabaseTable: "results",
      headerRow: 1,
      dataStartRow: 2,
      enabled: true,
      columns: [
        { id: "cm-21", excelColumn: "A", excelHeader: "Student ID", supabaseColumn: "student_number", dataType: "text", required: true, uniqueKey: true, transformation: "normalize_id" },
        { id: "cm-22", excelColumn: "B", excelHeader: "Term", supabaseColumn: "term", dataType: "text", required: true, uniqueKey: true, transformation: "trim" },
        { id: "cm-23", excelColumn: "C", excelHeader: "Mathematics", supabaseColumn: "mathematics", dataType: "decimal", required: false, uniqueKey: false, transformation: "parse_number" },
        { id: "cm-24", excelColumn: "D", excelHeader: "Science", supabaseColumn: "science", dataType: "decimal", required: false, uniqueKey: false, transformation: "parse_number" },
        { id: "cm-25", excelColumn: "E", excelHeader: "English", supabaseColumn: "english", dataType: "decimal", required: false, uniqueKey: false, transformation: "parse_number" },
        { id: "cm-26", excelColumn: "F", excelHeader: "Total Score", supabaseColumn: "total_score", dataType: "decimal", required: false, uniqueKey: false, transformation: "parse_number" },
        { id: "cm-27", excelColumn: "G", excelHeader: "Grade", supabaseColumn: "grade", dataType: "text", required: false, uniqueKey: false, transformation: "uppercase" }
      ]
    }
  ];
}
