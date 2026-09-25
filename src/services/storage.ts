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
  TwoWaySyncSettings,
  AiWorkbookPreset
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
  SAMPLE_LOADED: 'nes_sample_loaded_v1',
  PRESETS: 'nes_ai_presets'
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
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        // Clean out legacy fake placeholder worker URL and dummy timestamps
        if (parsed.workerUrl === 'http://coolify-worker:8000') {
          parsed.workerUrl = '';
        }
        if (parsed.workerStatus === 'healthy' && !parsed.workerUrl) {
          parsed.workerStatus = 'idle';
        }
        return parsed;
      } catch {}
    }
    const defaultSettings: SyncSettings = {
      syncInterval: '15m',
      autoSyncEnabled: false,
      backupToStorage: false,
      storageBucket: 'excel-archives',
      workerUrl: '',
      workerSecretKey: '',
      workerStatus: 'idle',
      lastSyncAt: undefined,
      nextSyncAt: undefined
    };
    this.saveSyncSettings(defaultSettings);
    return defaultSettings;
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

  static getPresets(): AiWorkbookPreset[] {
    const raw = localStorage.getItem(STORAGE_KEYS.PRESETS);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {}
    }
    return [];
  }

  static savePresets(presets: AiWorkbookPreset[]): void {
    try {
      localStorage.setItem(STORAGE_KEYS.PRESETS, JSON.stringify(presets));
    } catch (e) {
      console.warn('LocalStorage limit for presets:', e);
    }
  }

  private static _cachedDbState: Record<string, any[]> | null = null;

  static getImportLogs(): ImportLog[] {
    const raw = localStorage.getItem(STORAGE_KEYS.IMPORT_LOGS);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {}
    }
    return [];
  }

  static saveImportLogs(logs: ImportLog[]): void {
    try {
      const trimmed = logs.slice(0, 100);
      localStorage.setItem(STORAGE_KEYS.IMPORT_LOGS, JSON.stringify(trimmed));
    } catch (e) {
      console.warn('LocalStorage quota limit reached for import logs, saving recent logs only:', e);
      try {
        localStorage.setItem(STORAGE_KEYS.IMPORT_LOGS, JSON.stringify(logs.slice(0, 20)));
      } catch {}
    }
  }

  static getImportErrors(): ImportAuditError[] {
    const raw = localStorage.getItem(STORAGE_KEYS.IMPORT_ERRORS);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {}
    }
    return [];
  }

  static saveImportErrors(errs: ImportAuditError[]): void {
    try {
      const trimmed = errs.slice(0, 150);
      localStorage.setItem(STORAGE_KEYS.IMPORT_ERRORS, JSON.stringify(trimmed));
    } catch (e) {
      console.warn('LocalStorage quota limit reached for import errors:', e);
      try {
        localStorage.setItem(STORAGE_KEYS.IMPORT_ERRORS, JSON.stringify(errs.slice(0, 30)));
      } catch {}
    }
  }

  static getWorkerLogs(): LogMessage[] {
    const raw = localStorage.getItem(STORAGE_KEYS.WORKER_LOGS);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {}
    }
    return [];
  }

  static saveWorkerLogs(logs: LogMessage[]): void {
    try {
      const trimmed = logs.slice(0, 150);
      localStorage.setItem(STORAGE_KEYS.WORKER_LOGS, JSON.stringify(trimmed));
    } catch (e) {
      console.warn('LocalStorage quota limit reached for worker logs:', e);
      try {
        localStorage.setItem(STORAGE_KEYS.WORKER_LOGS, JSON.stringify(logs.slice(0, 30)));
      } catch {}
    }
  }

  static getDatabaseState(): Record<string, any[]> {
    if (this._cachedDbState) {
      return this._cachedDbState;
    }
    const raw = localStorage.getItem(STORAGE_KEYS.DB_STATE);
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        this._cachedDbState = parsed;
        return parsed;
      } catch {}
    }
    return {};
  }

  static saveDatabaseState(state: Record<string, any[]>): void {
    // Keep full in-memory cache for all tables (even with 50,000+ rows)
    this._cachedDbState = state;

    if (!state || typeof state !== 'object') {
      localStorage.removeItem(STORAGE_KEYS.DB_STATE);
      return;
    }

    try {
      // First try saving full state
      localStorage.setItem(STORAGE_KEYS.DB_STATE, JSON.stringify(state));
    } catch (quotaErr) {
      console.warn('LocalStorage quota exceeded for full database state. Saving top 100 rows per table for offline preview:', quotaErr);
      try {
        // Sample down to first 100 rows per table for offline browser storage
        const lightweightState: Record<string, any[]> = {};
        for (const [table, rows] of Object.entries(state)) {
          if (Array.isArray(rows)) {
            lightweightState[table] = rows.slice(0, 100);
          } else {
            lightweightState[table] = rows;
          }
        }
        localStorage.setItem(STORAGE_KEYS.DB_STATE, JSON.stringify(lightweightState));
      } catch (nestedErr) {
        console.warn('Could not persist database preview to localStorage, operating in memory-only mode:', nestedErr);
      }
    }
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
    try {
      const trimmed = conflicts.slice(0, 150);
      localStorage.setItem(STORAGE_KEYS.CONFLICTS, JSON.stringify(trimmed));
    } catch (e) {
      console.warn('LocalStorage quota limit reached for conflicts:', e);
      try {
        localStorage.setItem(STORAGE_KEYS.CONFLICTS, JSON.stringify(conflicts.slice(0, 30)));
      } catch {}
    }
  }

  private static _cachedBaselines: Record<string, SyncBaselineRecord> | null = null;

  static getSyncBaselines(): Record<string, SyncBaselineRecord> {
    if (this._cachedBaselines) {
      return this._cachedBaselines;
    }
    const raw = localStorage.getItem(STORAGE_KEYS.SYNC_BASELINES);
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') {
          this._cachedBaselines = parsed;
          return parsed;
        }
      } catch {}
    }
    return {};
  }

  static saveSyncBaselines(baselines: Record<string, SyncBaselineRecord>): void {
    // Retain full baseline records in memory
    this._cachedBaselines = baselines;

    if (!baselines || typeof baselines !== 'object') {
      localStorage.removeItem(STORAGE_KEYS.SYNC_BASELINES);
      return;
    }

    // Build a compact, lightweight representation by stripping heavy syncedValues payload (which takes 95% of space)
    const compactBaselines: Record<string, Partial<SyncBaselineRecord>> = {};
    for (const [key, b] of Object.entries(baselines)) {
      if (!b) continue;
      compactBaselines[key] = {
        recordKey: b.recordKey,
        tableName: b.tableName,
        primaryKeyCol: b.primaryKeyCol,
        primaryKeyValue: b.primaryKeyValue,
        excelHash: b.excelHash,
        supabaseHash: b.supabaseHash,
        lastSyncedAt: b.lastSyncedAt
      };
    }

    try {
      localStorage.setItem(STORAGE_KEYS.SYNC_BASELINES, JSON.stringify(compactBaselines));
    } catch (quotaErr) {
      console.warn('LocalStorage quota limit reached for sync baselines, pruning older baselines:', quotaErr);
      try {
        // Keep most recently synced 500 baseline records
        const sortedEntries = Object.entries(compactBaselines).sort((a, b) => {
          const timeA = a[1]?.lastSyncedAt ? new Date(a[1].lastSyncedAt).getTime() : 0;
          const timeB = b[1]?.lastSyncedAt ? new Date(b[1].lastSyncedAt).getTime() : 0;
          return timeB - timeA;
        });

        const pruned = Object.fromEntries(sortedEntries.slice(0, 500));
        localStorage.setItem(STORAGE_KEYS.SYNC_BASELINES, JSON.stringify(pruned));
      } catch (nestedErr) {
        console.warn('LocalStorage quota still exceeded for baselines, keeping in-memory only:', nestedErr);
        try {
          // Minimal fallback with top 100
          const sortedEntries = Object.entries(compactBaselines).slice(0, 100);
          localStorage.setItem(STORAGE_KEYS.SYNC_BASELINES, JSON.stringify(Object.fromEntries(sortedEntries)));
        } catch {}
      }
    }
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

  static getAiPresets(): AiWorkbookPreset[] {
    const raw = localStorage.getItem(STORAGE_KEYS.PRESETS);
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      } catch {}
    }
    return [];
  }

  static saveAiPresets(presets: AiWorkbookPreset[]): void {
    localStorage.setItem(STORAGE_KEYS.PRESETS, JSON.stringify(presets));
  }

  static getMergedGradeRecords(): any[] {
    const raw = localStorage.getItem('nes_merged_grade_records');
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      } catch {}
    }
    return [];
  }

  static saveMergedGradeRecords(records: any[]): void {
    try {
      localStorage.setItem('nes_merged_grade_records', JSON.stringify(records));
    } catch (e) {
      console.warn('LocalStorage quota limit reached for merged records, saving top slice:', e);
      try {
        localStorage.setItem('nes_merged_grade_records', JSON.stringify(records.slice(0, 500)));
      } catch {}
    }
  }

  static getStudentHistories(): import('../types').StudentHistoryRecord[] {
    const raw = localStorage.getItem('nes_student_histories');
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      } catch {}
    }
    return [];
  }

  static saveStudentHistories(histories: import('../types').StudentHistoryRecord[]): void {
    try {
      localStorage.setItem('nes_student_histories', JSON.stringify(histories));
    } catch (e) {
      console.warn('LocalStorage quota limit reached for student histories:', e);
    }
  }

  static addStudentHistory(history: import('../types').StudentHistoryRecord): import('../types').StudentHistoryRecord[] {
    const current = this.getStudentHistories();
    const updated = [history, ...current.filter(h => h.id !== history.id)];
    this.saveStudentHistories(updated);
    return updated;
  }

  static getNextcloudGradeSyncConfig(): import('../types').NextcloudGradeSyncConfig {
    const raw = localStorage.getItem('nes_nextcloud_grade_sync_cfg');
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        return {
          url: parsed.url || 'https://cloud.jhcnexus.space',
          username: parsed.username || 'truenas_admin',
          appPassword: parsed.appPassword || 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT',
          sourceFolder: parsed.sourceFolder || '/ExcelImports',
          selectedFiles: Array.isArray(parsed.selectedFiles) ? parsed.selectedFiles : [],
          autoSyncEnabled: !!parsed.autoSyncEnabled,
          syncIntervalMinutes: Number(parsed.syncIntervalMinutes) || 15,
          lastSyncTime: parsed.lastSyncTime || null,
          nextScheduledSyncTime: parsed.nextScheduledSyncTime || null,
          lastSyncStatus: parsed.lastSyncStatus || 'idle',
          lastSyncMessage: parsed.lastSyncMessage || '',
          lastSyncCount: Number(parsed.lastSyncCount) || 0,
          history: Array.isArray(parsed.history) ? parsed.history : []
        };
      } catch {}
    }
    return {
      url: 'https://cloud.jhcnexus.space',
      username: 'truenas_admin',
      appPassword: 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT',
      sourceFolder: '/ExcelImports',
      selectedFiles: [],
      autoSyncEnabled: false,
      syncIntervalMinutes: 15,
      lastSyncTime: null,
      nextScheduledSyncTime: null,
      lastSyncStatus: 'idle',
      lastSyncMessage: '',
      lastSyncCount: 0,
      history: []
    };
  }

  static saveNextcloudGradeSyncConfig(cfg: import('../types').NextcloudGradeSyncConfig): void {
    try {
      localStorage.setItem('nes_nextcloud_grade_sync_cfg', JSON.stringify(cfg));
    } catch (e) {
      console.warn('LocalStorage error saving Nextcloud grade sync config:', e);
    }
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
    this._cachedAnalysis = null;
    this._cachedDbState = null;
    this._cachedBaselines = null;
    localStorage.removeItem(STORAGE_KEYS.CURRENT_ANALYSIS);
    localStorage.removeItem(STORAGE_KEYS.MAPPINGS);
    localStorage.removeItem(STORAGE_KEYS.IMPORT_LOGS);
    localStorage.removeItem(STORAGE_KEYS.IMPORT_ERRORS);
    localStorage.removeItem(STORAGE_KEYS.WORKER_LOGS);
    localStorage.removeItem(STORAGE_KEYS.DB_STATE);
    localStorage.removeItem(STORAGE_KEYS.CONFLICTS);
    localStorage.removeItem(STORAGE_KEYS.SYNC_BASELINES);
    localStorage.removeItem(STORAGE_KEYS.SAMPLE_LOADED);
  }

  static initSampleIfNeeded(): void {
    const isLoaded = localStorage.getItem(STORAGE_KEYS.SAMPLE_LOADED);
    const existingMappings = this.getMappings();
    const existingAnalysis = this.getCurrentAnalysis();
    if (!isLoaded && existingMappings.length === 0 && !existingAnalysis) {
      this.loadSampleData();
    }
  }
}
