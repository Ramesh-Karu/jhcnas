import { 
  NextcloudConfig, 
  SupabaseConfig, 
  SyncSettings, 
  NextcloudFile, 
  WorkbookAnalysis, 
  WorksheetMapping, 
  ImportLog, 
  ImportAuditError, 
  LogMessage 
} from '../types';
import { createComplexSampleWorkbook, getDefaultSampleMappings } from './sampleWorkbook';
import { ExcelAnalyzer } from './excelAnalyzer';

const STORAGE_KEYS = {
  NEXTCLOUD: 'nes_nextcloud_config',
  SUPABASE: 'nes_supabase_config',
  SETTINGS: 'nes_sync_settings',
  FILES: 'nes_files',
  CURRENT_ANALYSIS: 'nes_current_analysis',
  MAPPINGS: 'nes_mappings',
  IMPORT_LOGS: 'nes_import_logs',
  IMPORT_ERRORS: 'nes_import_errors',
  WORKER_LOGS: 'nes_worker_logs',
  DB_STATE: 'nes_db_state',
  SAMPLE_LOADED: 'nes_sample_loaded_v1'
};

export class StorageService {
  static getNextcloudConfig(): NextcloudConfig {
    const raw = localStorage.getItem(STORAGE_KEYS.NEXTCLOUD);
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed.url === 'https://cloud.internal.truenas.net' || !parsed.url) {
          parsed.url = 'https://cloud.jhcnexus.space';
          parsed.webdavUrl = 'https://cloud.jhcnexus.space/remote.php/dav/files/truenas_admin/';
          parsed.username = 'truenas_admin';
          parsed.appPassword = 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT';
          this.saveNextcloudConfig(parsed);
        }
        return parsed;
      } catch {}
    }
    return {
      url: 'https://cloud.jhcnexus.space',
      webdavUrl: 'https://cloud.jhcnexus.space/remote.php/dav/files/truenas_admin/',
      username: 'truenas_admin',
      appPassword: 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT',
      sourceFolder: '/ExcelImports',
      isConnected: true,
      lastChecked: new Date().toISOString(),
      statusMessage: 'Connected to Nextcloud via WebDAV (cloud.jhcnexus.space)'
    };
  }

  static saveNextcloudConfig(cfg: NextcloudConfig): void {
    localStorage.setItem(STORAGE_KEYS.NEXTCLOUD, JSON.stringify(cfg));
  }

  static getSupabaseConfig(): SupabaseConfig {
    const raw = localStorage.getItem(STORAGE_KEYS.SUPABASE);
    if (raw) return JSON.parse(raw);
    return {
      url: 'https://dbsync-prod.supabase.co',
      anonKey: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyZWYiOiJkYnN5bmMtcHJvZCIsInJvbGUiOiJhbm9uIn0.sample_anon_key',
      serviceRoleKey: '••••••••••••••••••••••••••••••••••••••••••••••',
      isConnected: true,
      lastChecked: new Date().toISOString(),
      statusMessage: 'Connected to Supabase PostgreSQL (Schema verified)'
    };
  }

  static saveSupabaseConfig(cfg: SupabaseConfig): void {
    localStorage.setItem(STORAGE_KEYS.SUPABASE, JSON.stringify(cfg));
  }

  static getSyncSettings(): SyncSettings {
    const raw = localStorage.getItem(STORAGE_KEYS.SETTINGS);
    if (raw) return JSON.parse(raw);
    return {
      syncInterval: '15m',
      autoSyncEnabled: true,
      backupToStorage: true,
      storageBucket: 'excel-archives',
      workerUrl: 'http://coolify-worker:8000',
      workerStatus: 'healthy',
      lastSyncAt: new Date(Date.now() - 1000 * 60 * 12).toISOString(),
      nextSyncAt: new Date(Date.now() + 1000 * 60 * 3).toISOString()
    };
  }

  static saveSyncSettings(s: SyncSettings): void {
    localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(s));
  }

  static getFiles(): NextcloudFile[] {
    const raw = localStorage.getItem(STORAGE_KEYS.FILES);
    if (raw) {
      try {
        const parsed: NextcloudFile[] = JSON.parse(raw);
        if (!parsed.some(f => f.filename === 'students.xlsx')) {
          parsed.unshift({
            id: 'nc-students.xlsx',
            filename: 'students.xlsx',
            path: '/remote.php/dav/files/truenas_admin/ExcelImports/students.xlsx',
            fileSize: 411834,
            fileSizeFormatted: '402.2 KB',
            lastModified: new Date().toISOString(),
            fileHash: '055fb66f39721e7c8085d06914085ba14f33eafa53668598c0d974edcc018a0c',
            status: 'Synced',
            worksheetsCount: 1
          });
          this.saveFiles(parsed);
        }
        return parsed;
      } catch {}
    }
    return [
      {
        id: 'nc-students.xlsx',
        filename: 'students.xlsx',
        path: '/remote.php/dav/files/truenas_admin/ExcelImports/students.xlsx',
        fileSize: 411834,
        fileSizeFormatted: '402.2 KB',
        lastModified: new Date().toISOString(),
        fileHash: '055fb66f39721e7c8085d06914085ba14f33eafa53668598c0d974edcc018a0c',
        status: 'Synced',
        lastProcessedAt: new Date().toISOString(),
        worksheetsCount: 1
      },
      {
        id: 'file-1',
        filename: 'students_complex.xlsx',
        path: '/ExcelImports/students_complex.xlsx',
        fileSize: 42150,
        fileSizeFormatted: '41.2 KB',
        lastModified: new Date(Date.now() - 1000 * 60 * 45).toISOString(),
        fileHash: 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
        status: 'Synced',
        lastProcessedAt: new Date(Date.now() - 1000 * 60 * 12).toISOString(),
        worksheetsCount: 5
      }
    ];
  }

  static saveFiles(files: NextcloudFile[]): void {
    localStorage.setItem(STORAGE_KEYS.FILES, JSON.stringify(files));
  }

  static getCurrentAnalysis(): WorkbookAnalysis | null {
    const raw = localStorage.getItem(STORAGE_KEYS.CURRENT_ANALYSIS);
    if (raw) return JSON.parse(raw);
    return null;
  }

  static saveCurrentAnalysis(analysis: WorkbookAnalysis | null): void {
    if (!analysis) {
      localStorage.removeItem(STORAGE_KEYS.CURRENT_ANALYSIS);
    } else {
      localStorage.setItem(STORAGE_KEYS.CURRENT_ANALYSIS, JSON.stringify(analysis));
    }
  }

  static getMappings(): WorksheetMapping[] {
    const raw = localStorage.getItem(STORAGE_KEYS.MAPPINGS);
    if (raw) return JSON.parse(raw);
    return getDefaultSampleMappings('students_complex.xlsx');
  }

  static saveMappings(mappings: WorksheetMapping[]): void {
    localStorage.setItem(STORAGE_KEYS.MAPPINGS, JSON.stringify(mappings));
  }

  static getImportLogs(): ImportLog[] {
    const raw = localStorage.getItem(STORAGE_KEYS.IMPORT_LOGS);
    if (raw) return JSON.parse(raw);
    return [
      {
        id: 'log-101',
        filename: 'students_complex.xlsx',
        filePath: '/ExcelImports/students_complex.xlsx',
        fileHash: 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
        status: 'Success',
        isDryRun: false,
        numberOfWorksheets: 5,
        rowsProcessed: 43,
        rowsInserted: 35,
        rowsUpdated: 8,
        rowsFailed: 0,
        startedAt: new Date(Date.now() - 1000 * 60 * 12 - 2100).toISOString(),
        completedAt: new Date(Date.now() - 1000 * 60 * 12).toISOString(),
        durationMs: 2100
      },
      {
        id: 'log-102',
        filename: 'term2_attendance_weekly.xlsx',
        filePath: '/ExcelImports/term2_attendance_weekly.xlsx',
        fileHash: 'a7c5b98298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7899a12c',
        status: 'Skipped',
        isDryRun: false,
        numberOfWorksheets: 1,
        rowsProcessed: 0,
        rowsInserted: 0,
        rowsUpdated: 0,
        rowsFailed: 0,
        startedAt: new Date(Date.now() - 1000 * 60 * 12).toISOString(),
        completedAt: new Date(Date.now() - 1000 * 60 * 12).toISOString(),
        durationMs: 45,
        errorSummary: 'File unchanged (SHA-256 hash matched previous import)'
      },
      {
        id: 'log-103',
        filename: 'q1_admissions_draft.xlsx',
        filePath: '/ExcelImports/q1_admissions_draft.xlsx',
        fileHash: 'c2f8e12398fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7800ba89',
        status: 'Partial Success',
        isDryRun: false,
        numberOfWorksheets: 2,
        rowsProcessed: 28,
        rowsInserted: 26,
        rowsUpdated: 0,
        rowsFailed: 2,
        startedAt: new Date(Date.now() - 1000 * 60 * 360).toISOString(),
        completedAt: new Date(Date.now() - 1000 * 60 * 359).toISOString(),
        durationMs: 1350,
        errorSummary: '2 rows failed validation (Missing Student ID & Invalid Date)'
      }
    ];
  }

  static saveImportLogs(logs: ImportLog[]): void {
    localStorage.setItem(STORAGE_KEYS.IMPORT_LOGS, JSON.stringify(logs));
  }

  static getImportErrors(): ImportAuditError[] {
    const raw = localStorage.getItem(STORAGE_KEYS.IMPORT_ERRORS);
    if (raw) return JSON.parse(raw);
    return [
      {
        id: 'err-1',
        importLogId: 'log-103',
        filename: 'q1_admissions_draft.xlsx',
        worksheetName: 'Candidate List',
        rowNumber: 14,
        excelColumn: 'A',
        columnName: 'student_number',
        rawValue: '',
        errorMessage: 'Row 14: Required field \'student_number\' is missing or empty.',
        errorType: 'missing_required',
        createdAt: new Date(Date.now() - 1000 * 60 * 359).toISOString()
      },
      {
        id: 'err-2',
        importLogId: 'log-103',
        filename: 'q1_admissions_draft.xlsx',
        worksheetName: 'Candidate List',
        rowNumber: 22,
        excelColumn: 'D',
        columnName: 'dob',
        rawValue: '19/99/2026',
        errorMessage: 'Row 22: Invalid date format \'19/99/2026\'. Expected YYYY-MM-DD.',
        errorType: 'invalid_date',
        createdAt: new Date(Date.now() - 1000 * 60 * 359).toISOString()
      }
    ];
  }

  static saveImportErrors(errs: ImportAuditError[]): void {
    localStorage.setItem(STORAGE_KEYS.IMPORT_ERRORS, JSON.stringify(errs));
  }

  static getWorkerLogs(): LogMessage[] {
    const raw = localStorage.getItem(STORAGE_KEYS.WORKER_LOGS);
    if (raw) return JSON.parse(raw);
    return [
      { id: 'l1', timestamp: new Date(Date.now() - 1000 * 60 * 12).toISOString(), level: 'info', component: 'Worker', message: 'Triggering scheduled synchronization cycle (interval: 15m)' },
      { id: 'l2', timestamp: new Date(Date.now() - 1000 * 60 * 12 + 100).toISOString(), level: 'info', component: 'WebDAV', message: 'Connecting to Nextcloud TrueNAS SCALE at https://cloud.internal.truenas.net/remote.php/dav/files/excel-sync/' },
      { id: 'l3', timestamp: new Date(Date.now() - 1000 * 60 * 12 + 300).toISOString(), level: 'info', component: 'WebDAV', message: 'Scanned /ExcelImports: found 3 files' },
      { id: 'l4', timestamp: new Date(Date.now() - 1000 * 60 * 12 + 500).toISOString(), level: 'info', component: 'Parser', message: 'Analyzing students_complex.xlsx: detected 5 worksheets, 3 merged ranges (A1:H1 title, A3:H3 & A13:H13 section headers)' },
      { id: 'l5', timestamp: new Date(Date.now() - 1000 * 60 * 12 + 900).toISOString(), level: 'info', component: 'MappingEngine', message: 'Applied downward propagation: "CLASS 10A" mapped to 7 student records, "CLASS 10B" mapped to 5 student records' },
      { id: 'l6', timestamp: new Date(Date.now() - 1000 * 60 * 12 + 1500).toISOString(), level: 'info', component: 'Supabase', message: 'Batch upsert executed: 35 inserted, 8 updated into tables: students, attendance, sports, medical, results' },
      { id: 'l7', timestamp: new Date(Date.now() - 1000 * 60 * 12 + 2100).toISOString(), level: 'success', component: 'Worker', message: 'Cycle completed successfully. Updated SHA-256 hash.' }
    ];
  }

  static saveWorkerLogs(logs: LogMessage[]): void {
    localStorage.setItem(STORAGE_KEYS.WORKER_LOGS, JSON.stringify(logs));
  }

  static getDatabaseState(): Record<string, any[]> {
    const raw = localStorage.getItem(STORAGE_KEYS.DB_STATE);
    if (raw) return JSON.parse(raw);
    return {
      students: [
        { student_number: 'STU-1001', name: 'Alice Smith', class: 'CLASS 10A', gender: 'Female', dob: '2008-04-12', phone: '+15550192831', status: 'Active' },
        { student_number: 'STU-1002', name: 'Bob Jones', class: 'CLASS 10A', gender: 'Male', dob: '2008-09-21', phone: '+15550184422', status: 'Active' },
        { student_number: 'STU-1003', name: 'Charlie Davis', class: 'CLASS 10A', gender: 'Male', dob: '2007-12-05', phone: '+15550173311', status: 'Active' },
        { student_number: 'STU-1004', name: 'Diana Prince', class: 'CLASS 10A', gender: 'Female', dob: '2008-06-18', phone: '+15550165544', status: 'Active' },
        { student_number: 'STU-1005', name: 'Ethan Hunt', class: 'CLASS 10A', gender: 'Male', dob: '2008-01-30', phone: '+15550156677', status: 'Active' }
      ],
      attendance: [
        { student_number: 'STU-1001', attendance_date: '2025-09-15', status: 'Present', remarks: 'On time' },
        { student_number: 'STU-1002', attendance_date: '2025-09-15', status: 'Absent', remarks: 'Excused medical leave' },
        { student_number: 'STU-1003', attendance_date: '2025-09-15', status: 'Late', remarks: 'Bus delay' }
      ],
      sports: [
        { student_number: 'STU-1001', sport_name: 'Basketball', position: 'Point Guard', medical_clearance: true },
        { student_number: 'STU-1004', sport_name: 'Soccer', position: 'Midfielder', medical_clearance: true }
      ],
      medical: [
        { student_number: 'STU-1001', blood_group: 'O+', allergies: 'Peanuts', doctor_contact: 'Dr. Wilson (555-0100)' },
        { student_number: 'STU-1002', blood_group: 'A+', allergies: 'None', doctor_contact: 'Dr. Gomez (555-0200)' }
      ],
      results: [
        { student_number: 'STU-1001', term: 'Term 1', mathematics: 94.5, science: 88.0, english: 92.0, total_score: 274.5, grade: 'A' },
        { student_number: 'STU-1002', term: 'Term 1', mathematics: 76.0, science: 82.5, english: 79.0, total_score: 237.5, grade: 'B' }
      ]
    };
  }

  static saveDatabaseState(state: Record<string, any[]>): void {
    localStorage.setItem(STORAGE_KEYS.DB_STATE, JSON.stringify(state));
  }

  static initSampleIfNeeded(): void {
    if (!localStorage.getItem(STORAGE_KEYS.SAMPLE_LOADED)) {
      const sample = createComplexSampleWorkbook();
      const { analysis } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
      analysis.fileHash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
      this.saveCurrentAnalysis(analysis);
      localStorage.setItem(STORAGE_KEYS.SAMPLE_LOADED, 'true');
    }
  }
}
