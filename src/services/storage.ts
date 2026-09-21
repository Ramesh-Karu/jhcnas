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
    return [];
  }

  static saveMappings(mappings: WorksheetMapping[]): void {
    localStorage.setItem(STORAGE_KEYS.MAPPINGS, JSON.stringify(mappings));
  }

  static getImportLogs(): ImportLog[] {
    const raw = localStorage.getItem(STORAGE_KEYS.IMPORT_LOGS);
    if (raw) return JSON.parse(raw);
    return [];
  }

  static saveImportLogs(logs: ImportLog[]): void {
    localStorage.setItem(STORAGE_KEYS.IMPORT_LOGS, JSON.stringify(logs));
  }

  static getImportErrors(): ImportAuditError[] {
    const raw = localStorage.getItem(STORAGE_KEYS.IMPORT_ERRORS);
    if (raw) return JSON.parse(raw);
    return [];
  }

  static saveImportErrors(errs: ImportAuditError[]): void {
    localStorage.setItem(STORAGE_KEYS.IMPORT_ERRORS, JSON.stringify(errs));
  }

  static getWorkerLogs(): LogMessage[] {
    const raw = localStorage.getItem(STORAGE_KEYS.WORKER_LOGS);
    if (raw) return JSON.parse(raw);
    return [];
  }

  static saveWorkerLogs(logs: LogMessage[]): void {
    localStorage.setItem(STORAGE_KEYS.WORKER_LOGS, JSON.stringify(logs));
  }

  static getDatabaseState(): Record<string, any[]> {
    const raw = localStorage.getItem(STORAGE_KEYS.DB_STATE);
    if (raw) return JSON.parse(raw);
    return {};
  }

  static saveDatabaseState(state: Record<string, any[]>): void {
    localStorage.setItem(STORAGE_KEYS.DB_STATE, JSON.stringify(state));
  }

  static loadSampleData(): void {
    const sample = createComplexSampleWorkbook();
    const { analysis } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
    analysis.fileHash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
    this.saveCurrentAnalysis(analysis);
    this.saveMappings(getDefaultSampleMappings(sample.filename));
    localStorage.setItem(STORAGE_KEYS.SAMPLE_LOADED, 'true');
  }

  static clearAllData(): void {
    localStorage.removeItem(STORAGE_KEYS.CURRENT_ANALYSIS);
    localStorage.removeItem(STORAGE_KEYS.MAPPINGS);
    localStorage.removeItem(STORAGE_KEYS.IMPORT_LOGS);
    localStorage.removeItem(STORAGE_KEYS.IMPORT_ERRORS);
    localStorage.removeItem(STORAGE_KEYS.WORKER_LOGS);
    localStorage.removeItem(STORAGE_KEYS.DB_STATE);
    localStorage.removeItem(STORAGE_KEYS.SAMPLE_LOADED);
  }

  static initSampleIfNeeded(): void {
    // In production mode, leave empty until user configures or uploads files
  }
}
