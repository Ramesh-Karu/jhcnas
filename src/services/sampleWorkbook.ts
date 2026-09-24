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
      syncPolicy: "EXCEL_TO_DB", // Nextcloud Excel Master -> Pushes to Supabase
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
      syncPolicy: "DB_TO_EXCEL", // Supabase DB Master (teachers mark in app) -> Mirrors to Excel
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
      syncPolicy: "BIDIRECTIONAL", // Collaborative: Both sides can edit
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
      syncPolicy: "EXCEL_TO_DB", // Nextcloud Excel Master (Nurse clinic records)
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
      syncPolicy: "DB_TO_EXCEL", // Supabase DB Master (Exam portal) -> Mirrors to Excel
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

export function getSampleInitialDatabaseState(): Record<string, any[]> {
  return {
    students: [
      {
        student_number: "STU-1001",
        name: "Alice Smith",
        gender: "Female",
        dob: "2008-04-12",
        phone: "+1 555-019-2831",
        emergency_contact: "+1 555 019 9999",
        status: "Active",
        class: "CLASS 10A"
      },
      {
        student_number: "STU-1002",
        name: "Bob Jones",
        gender: "Male",
        dob: "2008-09-21",
        phone: "+1 (555) 999-1122", // Different from Excel!
        emergency_contact: "+1 555 018 8888",
        status: "Active",
        class: "CLASS 10A"
      },
      {
        student_number: "STU-1003",
        name: "Charlie Davis",
        gender: "Male",
        dob: "2007-12-05",
        phone: "555-017-3311",
        emergency_contact: "555-017-0000",
        status: "Active",
        class: "CLASS 10A"
      },
      {
        student_number: "STU-1004",
        name: "Diana Prince",
        gender: "Female",
        dob: "2008-06-18",
        phone: "+1 (555) 016-5544",
        emergency_contact: "+1 555 016 7777",
        status: "Active",
        class: "CLASS 10A"
      },
      {
        student_number: "STU-1005",
        name: "Ethan Hunt",
        gender: "Male",
        dob: "2008-01-30",
        phone: "+1 (555) 015-6677",
        emergency_contact: "+1 555 015 2222",
        status: "On Medical Leave", // Different from Excel!
        class: "CLASS 10A"
      },
      {
        student_number: "STU-1006",
        name: "Fiona Gallagher",
        gender: "Female",
        dob: "2008-11-14",
        phone: "+1 (555) 014-9988",
        emergency_contact: "+1 555 014 3333",
        status: "Active",
        class: "CLASS 10A"
      },
      {
        student_number: "STU-1007",
        name: "George Clark",
        gender: "Male",
        dob: "2008-03-08",
        phone: "+1 (555) 013-1122",
        emergency_contact: "+1 555 013 4444",
        status: "Active",
        class: "CLASS 10A"
      },
      {
        student_number: "STU-1008",
        name: "Hannah Abbott",
        gender: "Female",
        dob: "2008-07-22",
        phone: "+1 (555) 012-7788",
        emergency_contact: "+1 555 012 5555",
        status: "Active",
        class: "CLASS 10B"
      },
      {
        student_number: "STU-1009",
        name: "Ian Malcolm",
        gender: "Male",
        dob: "2007-10-15",
        phone: "+1 555-888-0000", // Different from Excel!
        emergency_contact: "+1 555 011 6666",
        status: "Active",
        class: "CLASS 10B"
      },
      {
        student_number: "STU-1010",
        name: "Julia Roberts",
        gender: "Female",
        dob: "2008-02-19",
        phone: "+1 (555) 010-5566",
        emergency_contact: "+1 555 010 7777",
        status: "Active",
        class: "CLASS 10B"
      },
      {
        // Exists in Supabase, but NOT in Excel!
        student_number: "STU-2001",
        name: "Marcus Aurelius",
        gender: "Male",
        dob: "2007-04-26",
        phone: "+1 (555) 777-3322",
        emergency_contact: "+1 555 777 9999",
        status: "Active",
        class: "CLASS 10B"
      }
    ],
    attendance: [
      { student_number: "STU-1001", attendance_date: "2025-09-15", status: "Present", remarks: "On time" },
      { student_number: "STU-1002", attendance_date: "2025-09-15", status: "Absent", remarks: "Excused medical leave" }
    ],
    results: [
      { student_number: "STU-1001", term: "Term 1", mathematics: 92, science: 88, english: 95, total_score: 275, grade: "A" }
    ]
  };
}

export function createSchoolCustomLayoutWorkbook(): { workbook: XLSX.WorkBook; binaryData: Uint8Array; filename: string } {
  const wb = XLSX.utils.book_new();

  // 1. Grade 6 Sheet (Grade partitioned)
  const grade6Data = [
    ["GRADE 6 - STUDENT ACADEMIC MARKSHEET (2024-2025)", "", "", "", "", ""],
    ["Class Teacher: Mrs. Anderson", "Room: 106", "Term: Final Assessment", "", "", ""],
    ["", "", "", "", "", ""],
    ["Roll No", "Student Name", "Mathematics", "Science", "English", "Result"],
    ["G6-101", "Emma Watson", 92, 88, 95, "Passed"],
    ["G6-102", "Daniel Craig", 78, 82, 74, "Passed"],
    ["G6-103", "Rupert Grint", 65, 70, 68, "Passed"],
    ["G6-104", "Tom Felton", 88, 91, 85, "Passed"],
    ["G6-105", "Bonnie Wright", 94, 96, 92, "Passed"]
  ];
  const wsG6 = XLSX.utils.aoa_to_sheet(grade6Data);
  wsG6['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
  XLSX.utils.book_append_sheet(wb, wsG6, "Grade 6");

  // 2. Grade 7 Sheet (Grade partitioned)
  const grade7Data = [
    ["GRADE 7 - STUDENT ACADEMIC MARKSHEET (2024-2025)", "", "", "", "", ""],
    ["Class Teacher: Mr. Davis", "Room: 204", "Term: Final Assessment", "", "", ""],
    ["", "", "", "", "", ""],
    ["Roll No", "Student Name", "Mathematics", "Science", "English", "Result"],
    ["G7-101", "Peter Parker", 96, 98, 89, "Passed"],
    ["G7-102", "Gwen Stacy", 99, 95, 98, "Passed"],
    ["G7-103", "Miles Morales", 89, 92, 94, "Passed"],
    ["G7-104", "Harry Osborn", 74, 80, 71, "Passed"]
  ];
  const wsG7 = XLSX.utils.aoa_to_sheet(grade7Data);
  wsG7['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
  XLSX.utils.book_append_sheet(wb, wsG7, "Grade 7");

  // 3. Grade 8 Sheet (Grade partitioned)
  const grade8Data = [
    ["GRADE 8 - STUDENT ACADEMIC MARKSHEET (2024-2025)", "", "", "", "", ""],
    ["Class Teacher: Dr. Banner", "Room: 302", "Term: Final Assessment", "", "", ""],
    ["", "", "", "", "", ""],
    ["Roll No", "Student Name", "Mathematics", "Science", "English", "Result"],
    ["G8-101", "Tony Stark", 100, 100, 95, "Passed"],
    ["G8-102", "Steve Rogers", 85, 87, 92, "Passed"],
    ["G8-103", "Natasha Romanoff", 94, 91, 98, "Passed"],
    ["G8-104", "Bruce Wayne", 97, 95, 96, "Passed"]
  ];
  const wsG8 = XLSX.utils.aoa_to_sheet(grade8Data);
  wsG8['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
  XLSX.utils.book_append_sheet(wb, wsG8, "Grade 8");

  // 4. Timetable Sheet: Dual-Cell Format per Period (Subject + Teacher/Room)
  const timetableData = [
    ["WEEKLY TIMETABLE - GRADE 10 (ACADEMIC YEAR 2024-2025)", "", "", "", "", "", "", "", ""],
    ["Section: 10-A", "Class Teacher: Prof. McGonagall", "Room: Hall 3B", "", "", "", "", "", ""],
    // Row 3 (Index 2): Main Period Headers with 2-cell span
    ["Day", "Period 1 (08:30-09:30)", "", "Period 2 (09:30-10:30)", "", "Break", "Period 3 (11:00-12:00)", "", "Period 4 (12:00-01:00)"],
    // Row 4 (Index 3): Subheaders for the two cells
    ["Day", "Subject", "Teacher / Room", "Subject", "Teacher / Room", "Interval", "Subject", "Teacher / Room", "Subject", "Teacher / Room"],
    // Rows 5-9: Timetable rows with 2-cell data
    ["Monday", "Mathematics", "Mr. Clark / R101", "Physics", "Dr. Banner / Lab A", "Recess", "English Lit", "Ms. Watson / R102", "Chemistry", "Dr. Wells / Lab B"],
    ["Tuesday", "Biology", "Dr. Grey / Lab C", "Mathematics", "Mr. Clark / R101", "Recess", "History", "Mr. Davis / R104", "Computer Sci", "Ms. Lovelace / Lab D"],
    ["Wednesday", "Physics", "Dr. Banner / Lab A", "Physical Ed", "Coach Taylor / Gym", "Recess", "Mathematics", "Mr. Clark / R101", "Art & Craft", "Ms. Kahlo / Studio"],
    ["Thursday", "Chemistry", "Dr. Wells / Lab B", "English Lit", "Ms. Watson / R102", "Recess", "Geography", "Mr. Scott / R105", "Economics", "Mr. Keynes / R106"],
    ["Friday", "Mathematics", "Mr. Clark / R101", "Computer Sci", "Ms. Lovelace / Lab D", "Recess", "Library / Reading", "Mrs. Pince / Lib", "Club / Activities", "Faculty Team"]
  ];
  const wsTimetable = XLSX.utils.aoa_to_sheet(timetableData);
  wsTimetable['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }, // A1:J1 Banner
    { s: { r: 2, c: 1 }, e: { r: 2, c: 2 } }, // B3:C3 Period 1 (2 cells)
    { s: { r: 2, c: 3 }, e: { r: 2, c: 4 } }, // D3:E3 Period 2 (2 cells)
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, // G3:H3 Period 3 (2 cells)
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, // I3:J3 Period 4 (2 cells)
  ];
  XLSX.utils.book_append_sheet(wb, wsTimetable, "Timetable_Grade10");

  // 5. Year 2020 Sheet (Year partitioned)
  const yr2020Data = [
    ["ACADEMIC ENROLLMENTS - YEAR 2020", "", "", "", ""],
    ["Student ID", "Student Name", "Enrollment Date", "Grade Enrolled", "Tuition Status"],
    ["ENR-2020-001", "Alex Ferguson", "2020-01-15", "Grade 6", "Paid"],
    ["ENR-2020-002", "Beth Harmon", "2020-01-18", "Grade 7", "Paid"],
    ["ENR-2020-003", "Carl Sagan", "2020-01-20", "Grade 8", "Scholarship"]
  ];
  const ws2020 = XLSX.utils.aoa_to_sheet(yr2020Data);
  ws2020['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
  XLSX.utils.book_append_sheet(wb, ws2020, "Academic_2020");

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const binaryData = new Uint8Array(wbout);

  return {
    workbook: wb,
    binaryData,
    filename: "school_custom_designs_grades_timetable.xlsx"
  };
}

export function generateLayoutStructureExport(analysis: WorkbookAnalysis): string {
  let output = `========================================================\n`;
  output += `WORKBOOK ARCHITECTURE EXPORT: ${analysis.filename}\n`;
  output += `Total Worksheets: ${analysis.totalWorksheets} | Size: ${analysis.fileSizeFormatted}\n`;
  output += `========================================================\n\n`;

  analysis.worksheets.forEach((ws, idx) => {
    output += `--- [Sheet ${idx + 1}/${analysis.totalWorksheets}]: "${ws.sheetName}" ---\n`;
    output += `Rows: ${ws.totalRows} | Cols: ${ws.totalColumns} | Header Row: ${ws.detectedHeaderRow} | Data Start: ${ws.detectedDataStartRow}\n`;
    
    if (ws.mergedRanges && ws.mergedRanges.length > 0) {
      output += `Merged Cells (${ws.mergedRanges.length}):\n`;
      ws.mergedRanges.slice(0, 10).forEach(m => {
        const valStr = m.value ? ` -> "${m.value}"` : '';
        output += `  • ${m.range} (${m.type})${valStr}\n`;
      });
      if (ws.mergedRanges.length > 10) {
        output += `  ... and ${ws.mergedRanges.length - 10} more merged ranges\n`;
      }
    } else {
      output += `Merged Cells: None\n`;
    }

    output += `\nColumns Detected:\n`;
    ws.headers.forEach(h => {
      output += `  [Col ${h.colLetter}] "${h.name}" (${h.inferredType || 'text'})\n`;
    });

    if (ws.sampleRows && ws.sampleRows.length > 0) {
      output += `\nSample Grid (First ${Math.min(ws.sampleRows.length, 6)} rows):\n`;
      ws.sampleRows.slice(0, 6).forEach((row) => {
        const rowVals = Object.keys(row.data).map(k => `${k}: ${JSON.stringify(row.data[k])}`).join(' | ');
        output += `  Row ${row.rowNumber}: ${rowVals}\n`;
      });
    }
    output += `\n`;
  });

  return output;
}

