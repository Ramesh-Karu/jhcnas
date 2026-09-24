import * as XLSX from 'xlsx';

/**
 * Generates realistic binary Excel workbook for Master Timetable Matrix
 * Archetype: TIMETABLE_MATRIX (Monday - Friday, Grade 6..13, Interleaved Subject / Teacher pairs)
 */
export function generateMasterTimetableWorkbook(): Buffer {
  const wb = XLSX.utils.book_new();
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const grades = [
    { name: 'Grade : 6', divisions: ['A', 'B', 'C'] },
    { name: 'Grade : 7', divisions: ['A', 'B', 'C', 'D'] },
    { name: 'Grade : 8', divisions: ['A', 'B', 'C'] },
    { name: 'Grade : 12 Science', divisions: ['A', 'B'] }
  ];

  const subjects = ['MATH', 'SCI', 'ENG', 'TAM', 'HIST', 'GEO', 'ICT', 'REL', 'COMM'];
  const teachers = [
    'Mr.K.Nimal', 'Mrs.V.Puvana', 'Mr.S.Ramesh / Mrs.T.Latha', 
    'Mrs.M.Selvi', 'Mr.P.Kumar', 'Mrs.R.Shanthi', 'Mr.T.Balan'
  ];

  for (const day of days) {
    const sheetData: any[][] = [];
    const merges: XLSX.Range[] = [];
    let currentRow = 0;

    // Title banner
    sheetData.push([`JAFFNA HINDU COLLEGE - MASTER TIMETABLE (${day.toUpperCase()})`, '', '', '', '', '', '', '']);
    merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow, c: 7 } });
    currentRow++;

    for (const g of grades) {
      // Grade Section Header (Merged)
      sheetData.push([g.name, '', '', '', '', '', '', '']);
      merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow, c: 7 } });
      currentRow++;

      // Period header
      sheetData.push(['Div', '1', '2', '3', '4', '5', '6', '7']);
      currentRow++;

      for (let dIdx = 0; dIdx < g.divisions.length; dIdx++) {
        const div = g.divisions[dIdx];
        
        // Row 1: Subject row
        const subjRow = [div];
        for (let p = 1; p <= 7; p++) {
          const s = subjects[(p + dIdx + currentRow) % subjects.length];
          subjRow.push(s);
        }
        sheetData.push(subjRow);
        currentRow++;

        // Row 2: Teacher row (interleaved pair)
        const teachRow = [''];
        for (let p = 1; p <= 7; p++) {
          const t = teachers[(p * 2 + dIdx + currentRow) % teachers.length];
          teachRow.push(t);
        }
        sheetData.push(teachRow);
        currentRow++;
      }

      // Spacer row
      sheetData.push(['', '', '', '', '', '', '', '']);
      currentRow++;
    }

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws['!merges'] = merges;
    XLSX.utils.book_append_sheet(wb, ws, day);
  }

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return Buffer.from(wbout);
}

/**
 * Generates realistic binary Excel workbook for School Financial Ledger
 * Archetype: MULTI_SHEET_LEDGER
 * IMPORTANT: Explicitly contains merged rows defining years (e.g. "Year 2023", "Year 2024")
 * merged across columns A to G, ensuring we verify that the pipeline NEVER uses them as data records!
 */
export function generateDonationsLedgerWorkbook(): Buffer {
  const wb = XLSX.utils.book_new();

  // 1. Sheet: Things Donation
  const thingsData: any[][] = [];
  const thingsMerges: XLSX.Range[] = [];
  let r = 0;

  // Title Banner
  thingsData.push(['JAFFNA HINDU COLLEGE - CONTRIBUTIONS & DONATIONS IN KIND', '', '', '', '', '', '']);
  thingsMerges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } });
  r++;

  // Column Headers
  thingsData.push(['Receipt No', 'Date', 'Donor Name', 'Item Description', 'Quantity', 'Valuation (Rs)', 'Folio']);
  r++;

  // Merged Year Header Row: "Year-2023" (Merged A3:G3) - MUST NOT BE USED AS A DATA ROW
  thingsData.push(['Year-2023', '', '', '', '', '', '']);
  thingsMerges.push({ s: { r, c: 0 }, e: { r, c: 6 } });
  r++;

  // 2023 Records
  thingsData.push(['RCP-2023-01', '12.01.2023', 'Old Boys Association UK', 'Library English Literature Books', '150', 'Rs 75000', 'F-101']);
  r++;
  thingsData.push(['RCP-2023-02', '24.03.2023', 'Sydney Alumni Chapter', 'Chemistry Lab Glassware Sets', '25', 'Rs 120000', 'F-102']);
  r++;
  thingsData.push(['RCP-2023-03', '15.06.2023', 'Dr. K. Sivalingam', 'Desktop Computers Core i7', '5', 'Rs 450000', 'F-103']);
  r++;
  thingsData.push(['RCP-2023-04', '08.09.2023', 'Batch of 1994', 'Cricket Equipment & Kits', '12', 'Rs 85000', 'F-104']);
  r++;
  thingsData.push(['RCP-2023-05', '19.11.2023', 'Mr. V. Tharmarajah', 'Microscopes for Bio Lab', '8', 'Rs 160000', 'F-105']);
  r++;

  // Merged Year Header Row: "Year-2024" (Merged across columns) - MUST NOT BE USED AS A DATA ROW
  thingsData.push(['Year-2024', '', '', '', '', '', '']);
  thingsMerges.push({ s: { r, c: 0 }, e: { r, c: 6 } });
  r++;

  // 2024 Records
  thingsData.push(['RCP-2024-01', '10.01.2024', 'Colombo OBA Trust', 'Smart Classroom Interactive Board', '2', 'Rs 600000', 'F-201']);
  r++;
  thingsData.push(['RCP-2024-02', '18.02.2024', 'Canada Alumni Chapter', 'Robotics & Arduino Starter Kits', '20', 'Rs 240000', 'F-202']);
  r++;
  thingsData.push(['RCP-2024-03', '05.04.2024', 'Eng. T. Mahendran', 'Solar Inverter Backup Unit 5kVA', '1', 'Rs 380000', 'F-203']);
  r++;

  const wsThings = XLSX.utils.aoa_to_sheet(thingsData);
  wsThings['!merges'] = thingsMerges;
  XLSX.utils.book_append_sheet(wb, wsThings, 'Things Donation');

  // 2. Sheet: Cash SDC
  const cashData: any[][] = [];
  const cashMerges: XLSX.Range[] = [];
  let cr = 0;

  cashData.push(['SCHOOL DEVELOPMENT COMMITTEE (SDC) - CASH RECEIPTS', '', '', '', '', '']);
  cashMerges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });
  cr++;

  cashData.push(['Receipt No', 'Date', 'Contributor / Member', 'Purpose', 'Amount (Rs)', 'Approved By']);
  cr++;

  // Merged Year Header Row: "Year 2023" (Merged A3:F3) - MUST NOT BE USED AS DATA
  cashData.push(['Year 2023', '', '', '', '', '']);
  cashMerges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: 5 } });
  cr++;

  cashData.push(['SDC-23-001', '15.02.2023', 'Mr. S. Nadarajah', 'Swimming Pool Maintenance Fund', 'Rs 25000', 'Principal']);
  cr++;
  cashData.push(['SDC-23-002', '20.03.2023', 'Mrs. K. Bavani', 'Annual Prize Giving Trophy Fund', 'Rs 15000', 'Vice Principal']);
  cr++;
  cashData.push(['SDC-23-003', '10.07.2023', 'Batch of 1988', 'Hostel Renovation Project', 'Rs 100000', 'Principal']);
  cr++;

  // Merged Year Header Row: "Year 2024" (Merged A7:F7) - MUST NOT BE USED AS DATA
  cashData.push(['Year 2024', '', '', '', '', '']);
  cashMerges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: 5 } });
  cr++;

  cashData.push(['SDC-24-001', '05.01.2024', 'Dr. V. Pathmanathan', 'Medical First Aid Center Fund', 'Rs 50000', 'Principal']);
  cr++;
  cashData.push(['SDC-24-002', '12.03.2024', 'Melbourne Alumni Asscn', 'Teachers ICT Skill Training Workshop', 'Rs 75000', 'Principal']);
  cr++;

  const wsCash = XLSX.utils.aoa_to_sheet(cashData);
  wsCash['!merges'] = cashMerges;
  XLSX.utils.book_append_sheet(wb, wsCash, 'Cash SDC');

  // 3. Sheet: Project Donation
  const projData: any[][] = [];
  const projMerges: XLSX.Range[] = [];
  let pr = 0;

  projData.push(['DEVELOPMENT CAPITAL PROJECTS & SPECIAL GRANTS', '', '', '', '', '']);
  projMerges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });
  pr++;

  projData.push(['Project ID', 'Project Name', 'Sponsoring Entity', 'Allocated Budget (Rs)', 'Completion Date', 'Status']);
  pr++;

  projData.push(['PRJ-01', 'Solar Rooftop Installation Phase 1', 'Ministry of Education Grant', 'Rs 1500000', '30.06.2023', 'Completed']);
  pr++;
  projData.push(['PRJ-02', 'Auditorium Acoustic Upgrade', 'Old Boys Global Trust', 'Rs 2200000', '15.11.2023', 'Completed']);
  pr++;
  projData.push(['PRJ-03', 'Language Learning Media Lab', 'Norwegian Tamil Society', 'Rs 850000', '28.02.2024', 'Operational']);
  pr++;
  projData.push(['PRJ-04', 'Science Pavilion Extension', 'Alumni Endowment 2024', 'Rs 3000000', '15.12.2024', 'In Progress']);
  pr++;

  const wsProj = XLSX.utils.aoa_to_sheet(projData);
  wsProj['!merges'] = projMerges;
  XLSX.utils.book_append_sheet(wb, wsProj, 'Project Donation');

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return Buffer.from(wbout);
}

/**
 * Generates realistic binary Excel workbook for Teacher-Subject Staff Allocation Matrix
 * Archetype: PIVOT_ALLOCATION_MATRIX (Grade 6..11, Subjects in Rows, Divisions A..H in Columns)
 */
export function generateTeacherAllocationsWorkbook(): Buffer {
  const wb = XLSX.utils.book_new();
  const grades = ['Grade 6', 'Grade 7', 'Grade 8', 'Grade 9', 'Grade 10', 'Grade 11'];
  const subjects = [
    'Mathematics', 'Science', 'English Language', 'Tamil Language', 
    'History', 'Geography', 'Information Technology', 'Health & Physical Ed', 
    'Religion', 'Art / Music / Drama'
  ];

  const staffPool = [
    'Mr.K.Sivalingam', 'Mrs.V.Puvana', 'Mr.P.Ganeshan', 'Mrs.R.Kavitha',
    'Mr.T.Balan', 'Mrs.M.Selvi', 'Mr.S.Kumar', 'Mrs.N.Sangeetha',
    'Mr.J.Robert / Mrs.A.Mary', 'Mr.V.Puvaneesan / Mrs.L.Mathi'
  ];

  for (let gIdx = 0; gIdx < grades.length; gIdx++) {
    const gName = grades[gIdx];
    const sheetData: any[][] = [];

    // Header row: Subject, A, B, C, D, E, F
    sheetData.push(['Subject', 'A', 'B', 'C', 'D', 'E', 'F']);

    for (let sIdx = 0; sIdx < subjects.length; sIdx++) {
      const subj = subjects[sIdx];
      const row = [subj];
      for (let divIdx = 0; divIdx < 6; divIdx++) {
        const staff = staffPool[(sIdx * 2 + divIdx + gIdx) % staffPool.length];
        row.push(staff);
      }
      sheetData.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(wb, ws, gName);
  }

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return Buffer.from(wbout);
}
