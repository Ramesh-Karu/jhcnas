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

/**
 * Generates realistic binary Excel workbook for Jaffna Hindu College Inventory & Departmental Ledger
 * Archetype: MULTI_SHEET_LEDGER (Master inventory, Responsible person directory, 18 Departmental Books)
 */
export function generateJhcInventoryWorkbook(): Buffer {
  // If running in Node and downloaded_sheet.xlsx exists, serve authentic copy
  try {
    const fs = (globalThis as any).process?.versions?.node ? require('fs') : null;
    const path = (globalThis as any).process?.versions?.node ? require('path') : null;
    if (fs && path) {
      const candidates = [
        path.join(process.cwd(), 'downloaded_sheet.xlsx'),
        path.join(process.cwd(), 'data', 'downloaded_sheet.xlsx'),
        '/data/downloaded_sheet.xlsx'
      ];
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          return fs.readFileSync(p);
        }
      }
    }
  } catch {}

  // Fallback programmatic generation
  const wb = XLSX.utils.book_new();

  // 1. Sheet: Master inventory
  const masterData: any[][] = [
    [' ', 'description', 'பிரிவும் இலக்கமும்', '', 'பேரேட்டில் காட்டியவாறான மீதி', 'உண்மையான கையிருப்பு மீதி', 'Total', 'Book 1 Office', 'Book 1 YMHA', 'Book 1 Band', 'Book 1 Scout , Redcross', 'Book 1 Library', 'Page', 'Book 2 arts,', 'Book 2 agri', 'book 2 Music', 'Book 2 Motor', 'Page', 'Book3 Sports', 'Book 3 Chess', 'Book 3 carrom', 'book 3 T.T', 'Page', 'Book4', 'Page', 'Book5', 'Page', 'Book6', 'Page', 'Book7', 'Page', 'Book8', 'Page', 'Book9', 'Page', 'Book10', 'Page', 'Book 12 I', 'Page', 'Book12 II', 'Page', 'Book12 III', 'Page', 'Book12 IV', 'Page', 'Book13 office', 'Book 13 sabalingam', 'Book 13 class', 'Book 13 oba sponser', 'book 13 s.Parliment', 'book 13 school are', 'book 13 reserch room', 'Page', 'Book14 bio', 'book 14 maths', 'book 14 physics', 'book 14 ICT', 'Page', 'Book15 Makilakam', 'Book 15 SMART CLASS ROOM', 'Page', 'Book 16 Video conferencing', 'Book 16 photography', 'Page'],
    [1, '0-6v (Volt.Metter)', 'I', 1, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1],
    [2, '13 Amp plug clipsal model', 'I', 1, 7, 7, 7, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 7, null, null, null, null, null, null, 5],
    [3, '15 Amp plug clipsal model', 'I', 1, 2, 2, 2, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 2, null, null, null, null, null, null, 5],
    [4, '5 Amp plug clipsal model', 'I', 1, 7, 7, 7, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 7, null, null, null, null, null, null, 5],
    [5, '16 channel canovis DVR', 'I', 1, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 73],
    [6, 'Air condition Plant (LG)', 'I', 1, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
    [7, 'Core i7 Computer System', 'I', 1, 15, 15, 15, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 15, 12, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]
  ];
  const wsMaster = XLSX.utils.aoa_to_sheet(masterData);
  XLSX.utils.book_append_sheet(wb, wsMaster, 'Master inventory');

  // 2. Sheet: Responsible person
  const respData: any[][] = [
    ['Jaffna Hindu College - Departmental Inventory Officers Directory'],
    ['Book No', 'Description', 'Responsible person', 'Contact number', 'Signature'],
    ['Book 1', 'Office', 'Mr. S. Suhierathan', '0776169716', 'Signed'],
    ['Book 1', 'YMHA', 'Mr. S. Suhierathan', '0776169716', 'Signed'],
    ['Book 1', 'Band', 'Mr. T. Nitharshan', '0771066721', 'Signed'],
    ['Book 1', 'Scout', 'Mr. S. Jeevaganathan', '0766585489', 'Signed'],
    ['Book 2', 'Art & Music', 'Mrs. V. Puvana', '0773344556', 'Signed'],
    ['Book 3', 'Sports', 'Mr. K. Nimal', '0778899112', 'Signed'],
    ['Book 5', 'Physics Lab', 'Mr. P. Kumar', '0771234567', 'Signed'],
    ['Book 6', 'ICT Computer Lab', 'Mrs. K. Suren', '0777654321', 'Signed'],
    ['book 7', 'General Science Lab', 'Mrs. R. Shanthi', '0779988776', 'Signed'],
    ['book 8', 'Bio Lab', 'Miss. K. Thadshajini', '0761122334', 'Signed'],
    ['book 9', 'Chemistry Lab', 'Mr. V. Tharmarajah', '0774455667', 'Signed'],
    ['book 10', 'Hostel', 'Mr. T. Balan', '0775566778', 'Signed'],
    ['book 12 I', 'Bio Tech', 'Miss. K. Thadshajini', '0761122334', 'Signed'],
    ['Book 12 II', 'E.Tech (Electrical)', 'Mr. N. Srikaran', '0776677889', 'Signed'],
    ['Book 12 III', 'Mechanic Lab', 'Mr. V. Kanesalingam', '0778899001', 'Signed'],
    ['Book 12 IV', 'Civil Lab', 'Mr. V. Kanesalingam', '0778899001', 'Signed'],
    ['Book 13', 'Administration', 'Mr. S. Suhierathan', '0776169716', 'Signed'],
    ['Book 14', 'Mahinthothaya Lab', 'Mr. M. Selvi', '0772233445', 'Signed'],
    ['book 15', 'Mahilakam & Smart Classroom', 'Mrs. N. Sangeetha', '0773322110', 'Signed'],
    ['book 16', 'Photography & Media', 'Mr. R. Ramanan', '0775544332', 'Signed'],
    ['Book 17', 'Mahinthothaya Physics Lab', 'Mr. P. Kumar', '0771234567', 'Signed']
  ];
  const wsResp = XLSX.utils.aoa_to_sheet(respData);
  XLSX.utils.book_append_sheet(wb, wsResp, 'Responsible person');

  // 3. Sheet: Book 1 Admin
  const b1Data: any[][] = [
    [null, 'Department', null, 'Jaffna Hindu College', null, null, 'Book 1'],
    [null, null, null, null, null, null, 'Admin'],
    ['No'],
    ['No', 'Articles', 'Section and Number', 'Balance as Shown Per Ledger', 'Actual Balance on Hand', 'Surplus', 'Deficiency', 'Remarks'],
    [1, 'Ahey Hom 18"', 1, 3, 3, 0, 0, 'Good condition'],
    [2, 'Ahey Unit 60 W', 1, 2, 2, 0, 0, 'In service'],
    [3, 'Air condition (National)', 1, 1, 1, 0, 0, 'Principal Room'],
    [4, 'Executive Conference Table', 1, 4, 4, 0, 0, 'Board Room']
  ];
  const wsB1 = XLSX.utils.aoa_to_sheet(b1Data);
  XLSX.utils.book_append_sheet(wb, wsB1, 'Book 1 Admin');

  // 4. Sheet: Book 6 ICT
  const b6Data: any[][] = [
    ['Department', null, 'Jaffna Hindu College', null, null, null, null, 'Book 6'],
    [null, null, null, null, 'Responsible person Mrs.K.Suren', null, null, 'Ict'],
    ['No', 'Articles', 'Section and Number3.', 'Balance as Shown Per Ledger', 'Actual Balance on Hand', 'Surplus', 'Deficiency', 'Remarks'],
    [1, 'Air condition Plant (fuji pure)', 1, 1, 1, 0, 0, 'Server room'],
    [2, 'Air condition plant (LG)', 1, 1, 1, 0, 0, 'Lab 1'],
    [3, 'Core i7 Workstations', 1, 25, 25, 0, 0, 'Student Lab'],
    [4, 'Cisco 24-Port Gigabit Switch', 1, 4, 4, 0, 0, 'Rack 1'],
    [5, 'Epson Multimedia Projector', 1, 3, 3, 0, 0, 'Wall mounted']
  ];
  const wsB6 = XLSX.utils.aoa_to_sheet(b6Data);
  XLSX.utils.book_append_sheet(wb, wsB6, 'Book 6 ICT');

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return Buffer.from(wbout);
}

