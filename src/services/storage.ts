import { 
  NextcloudConfig, 
  SupabaseConfig, 
  SyncSettings, 
  NextcloudFile, 
  WorkbookAnalysis, 
  WorksheetMapping, 
  ImportLog, 
  ImportAuditError, 
  LogMessage,
  SyncConflictRecord,
  SyncBaselineRecord,
  TwoWaySyncSettings
} from '../types';
import { createComplexSampleWorkbook, getDefaultSampleMappings, getSampleInitialDatabaseState } from './sampleWorkbook';
import { ExcelAnalyzer } from './excelAnalyzer';

function safeRangeToString(val: any): string {
  if (!val) return 'A1';
  if (typeof val === 'string') return val;
  if (typeof val === 'object' && 's' in val && 'e' in val) {
    const colToLetter = (c: number) => {
      let temp = c;
      let letter = '';
      while (temp >= 0) {
        letter = String.fromCharCode((temp % 26) + 65) + letter;
        temp = Math.floor(temp / 26) - 1;
      }
      return letter || 'A';
    };
    const sCol = colToLetter(val.s?.c ?? 0);
    const sRow = (val.s?.r ?? 0) + 1;
    const eCol = colToLetter(val.e?.c ?? 0);
    const eRow = (val.e?.r ?? 0) + 1;
    return `${sCol}${sRow}:${eCol}${eRow}`;
  }
  return String(val);
}

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
  CONFLICTS: 'nes_sync_conflicts',
  SYNC_BASELINES: 'nes_sync_baselines',
  TWOWAY_SETTINGS: 'nes_twoway_settings',
  SAMPLE_LOADED: 'nes_sample_loaded_v1'
};

export class StorageService {
  static getNextcloudConfig(): NextcloudConfig {
    const raw = localStorage.getItem(STORAGE_KEYS.NEXTCLOUD);
    if (raw) {
      try {
        const parsed: NextcloudConfig = JSON.parse(raw);
        if (parsed.url === 'https://cloud.internal.truenas.net' || !parsed.url) {
          parsed.url = 'https://cloud.jhcnexus.space';
          parsed.webdavUrl = 'https://cloud.jhcnexus.space/remote.php/dav/files/truenas_admin/';
          parsed.username = 'truenas_admin';
          parsed.appPassword = 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT';
        }
        // Do not keep unverified fake connected status
        if (parsed.isConnected && !parsed.lastChecked) {
          parsed.isConnected = false;
          parsed.statusMessage = 'Unverified. Click Test Nextcloud Connection to verify live access.';
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
      isConnected: false,
      lastChecked: undefined,
      statusMessage: 'Unverified. Click Test Connection to verify live WebDAV access.'
    };
  }

  static saveNextcloudConfig(cfg: NextcloudConfig): void {
    localStorage.setItem(STORAGE_KEYS.NEXTCLOUD, JSON.stringify(cfg));
  }

  static getSupabaseConfig(): SupabaseConfig {
    const raw = localStorage.getItem(STORAGE_KEYS.SUPABASE);
    if (raw) {
      try {
        const parsed: SupabaseConfig = JSON.parse(raw);
        // Handle migration from legacy serviceRoleKey to serviceKey
        if (parsed.serviceRoleKey && !parsed.serviceKey) {
          parsed.serviceKey = parsed.serviceRoleKey;
        } else if (!parsed.serviceKey) {
          parsed.serviceKey = '';
        }
        // Clean out any fake dummy/placeholder credentials
        if (!parsed.url || parsed.url.includes('dbsync-prod.supabase.co') || parsed.anonKey?.includes('sample_anon_key')) {
          parsed.url = '';
          parsed.anonKey = '';
          parsed.serviceKey = '';
          parsed.serviceRoleKey = '';
          parsed.isConnected = false;
          parsed.lastChecked = undefined;
          parsed.statusMessage = 'Not connected. Enter your Supabase Project URL and API Key to connect.';
          this.saveSupabaseConfig(parsed);
        } else if (!parsed.url || (!parsed.anonKey && !parsed.serviceKey)) {
          parsed.isConnected = false;
          parsed.statusMessage = 'Not connected. Missing Supabase credentials.';
        }
        return parsed;
      } catch {}
    }
    return {
      url: '',
      anonKey: '',
      serviceKey: '',
      isConnected: false,
      lastChecked: undefined,
      statusMessage: 'Not connected. Enter your Supabase Project URL and API Key to connect.'
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

  private static _cachedAnalysis: WorkbookAnalysis | null = null;

  static getCurrentAnalysis(): WorkbookAnalysis | null {
    if (this._cachedAnalysis) {
      return this._cachedAnalysis;
    }
    const raw = localStorage.getItem(STORAGE_KEYS.CURRENT_ANALYSIS);
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.worksheets)) {
          let modified = false;
          parsed.worksheets.forEach((ws: any) => {
            if (typeof ws.usedRange !== 'string') {
              ws.usedRange = safeRangeToString(ws.usedRange);
              modified = true;
            }
            if (Array.isArray(ws.mergedRanges)) {
              ws.mergedRanges.forEach((mr: any) => {
                if (typeof mr.range !== 'string') {
                  mr.range = safeRangeToString(mr.range);
                  modified = true;
                }
                if (typeof mr.value === 'object' && mr.value !== null) {
                  mr.value = JSON.stringify(mr.value);
                  modified = true;
                }
              });
            }
          });
          if (modified) {
            this.saveCurrentAnalysis(parsed);
          }
        }
        this._cachedAnalysis = parsed;
        return parsed;
      } catch {}
    }
    return null;
  }

  static saveCurrentAnalysis(analysis: WorkbookAnalysis | null): void {
    this._cachedAnalysis = analysis;
    if (!analysis) {
      localStorage.removeItem(STORAGE_KEYS.CURRENT_ANALYSIS);
    } else {
      try {
        localStorage.setItem(STORAGE_KEYS.CURRENT_ANALYSIS, JSON.stringify(analysis));
      } catch (quotaErr) {
        console.warn('LocalStorage quota exceeded for analysis with base64Data, saving metadata only to localStorage:', quotaErr);
        try {
          const stripped = { ...analysis, base64Data: undefined };
          localStorage.setItem(STORAGE_KEYS.CURRENT_ANALYSIS, JSON.stringify(stripped));
        } catch {}
      }
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

  static getConflicts(): SyncConflictRecord[] {
    const raw = localStorage.getItem(STORAGE_KEYS.CONFLICTS);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {}
    }
    return [];
  }

  static saveConflicts(conflicts: SyncConflictRecord[]): void {
    localStorage.setItem(STORAGE_KEYS.CONFLICTS, JSON.stringify(conflicts));
  }

  static getSyncBaselines(): Record<string, SyncBaselineRecord> {
    const raw = localStorage.getItem(STORAGE_KEYS.SYNC_BASELINES);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {}
    }
    return {};
  }

  static saveSyncBaselines(baselines: Record<string, SyncBaselineRecord>): void {
    localStorage.setItem(STORAGE_KEYS.SYNC_BASELINES, JSON.stringify(baselines));
  }

  static updateSyncBaseline(recordKey: string, baseline: SyncBaselineRecord): void {
    const current = this.getSyncBaselines();
    current[recordKey] = baseline;
    this.saveSyncBaselines(current);
  }

  static getTwoWaySyncSettings(): TwoWaySyncSettings {
    const raw = localStorage.getItem(STORAGE_KEYS.TWOWAY_SETTINGS);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {}
    }
    return {
      autoPushNonConflicting: true, // Default to true: automatic bi-directional sync without manual clicks
      defaultConflictResolution: 'manual', // True concurrent conflicts require manual decision
      createBackupBeforeExcelWrite: true,
      backupFolder: 'Sync_Backups'
    };
  }

  static saveTwoWaySyncSettings(settings: TwoWaySyncSettings): void {
    localStorage.setItem(STORAGE_KEYS.TWOWAY_SETTINGS, JSON.stringify(settings));
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
    localStorage.removeItem(STORAGE_KEYS.CONFLICTS);
    localStorage.removeItem(STORAGE_KEYS.SAMPLE_LOADED);
  }

  static initSampleIfNeeded(): void {
    const isLoaded = localStorage.getItem(STORAGE_KEYS.SAMPLE_LOADED);
    if (!isLoaded) {
      this.loadSampleData();
    }
  }
}
