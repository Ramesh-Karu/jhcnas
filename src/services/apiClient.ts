import { 
  NextcloudConfig, 
  SupabaseConfig, 
  NextcloudFile, 
  WorkbookAnalysis, 
  SupabaseTableInfo, 
  LiveSchedulerStatus, 
  SyncInterval, 
  SyncSettings,
  WorksheetMapping, 
  AutomatedRunDiagnosticResult,
  WorkerConnectionStatus,
  LiveSyncHistoryItem,
  LogMessage,
  AiWorkbookAnalysisResult,
  AiWorkbookPreset
} from '../types';

export interface SupabaseDiagnostic {
  errorCode: string;
  httpStatus: number;
  tableName: string;
  cause: string;
  details: string;
  remediation: string;
  suggestedSql?: string;
}

export interface UpsertBatchProgress {
  processed: number;
  total: number;
  percentage: number;
  currentBatch: number;
  totalBatches: number;
  successfulCount: number;
  failedCount: number;
  tableName: string;
  status: 'running' | 'success' | 'error';
  message: string;
}

export interface SupabaseTestResult {
  success: boolean;
  isConnected: boolean;
  latencyMs?: number;
  tablesCount?: number;
  tables?: string[];
  matchingSyncTables?: string[];
  checks: {
    hostReachability: boolean;
    authValid: boolean;
    schemaDetected: boolean;
  };
  message?: string;
  error?: string;
}

export class ApiClient {
  static async testNextcloudConnection(config: NextcloudConfig): Promise<{
    success: boolean;
    checks: {
      hostReachability: boolean;
      webdavHandshake: boolean;
      authValid: boolean;
      folderExists: boolean;
      details?: any;
    };
    message: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/nextcloud/test-connection', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          webdavUrl: config.webdavUrl,
          username: config.username,
          appPassword: config.appPassword,
          sourceFolder: config.sourceFolder,
        }),
      });

      const data = await res.json();
      return data;
    } catch (e: any) {
      return {
        success: false,
        checks: {
          hostReachability: false,
          webdavHandshake: false,
          authValid: false,
          folderExists: false,
        },
        message: 'Network error calling backend test endpoint',
        error: e.message,
      };
    }
  }

  static async listNextcloudFiles(config: NextcloudConfig): Promise<{
    success: boolean;
    folderPath?: string;
    count?: number;
    files?: NextcloudFile[];
    mappings?: WorksheetMapping[];
    error?: string;
  }> {
    try {
      const res = await fetch('/api/nextcloud/list-files', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          username: config.username,
          appPassword: config.appPassword,
          sourceFolder: config.sourceFolder,
        }),
      });

      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async fetchAndParseWorkbook(
    config: NextcloudConfig,
    filePath?: string,
    filename?: string
  ): Promise<{
    success: boolean;
    analysis?: WorkbookAnalysis;
    base64Data?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/nextcloud/fetch-and-parse', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          username: config.username,
          appPassword: config.appPassword,
          filePath,
          filename,
        }),
      });

      const data = await res.json();
      if (!data.success) {
        return { success: false, error: data.error };
      }

      const analysis: WorkbookAnalysis = {
        filename: data.filename,
        fileSize: data.fileSize,
        fileSizeFormatted: data.fileSizeFormatted,
        totalWorksheets: data.totalWorksheets,
        worksheets: data.worksheets,
        analyzedAt: data.analyzedAt,
        fileHash: data.fileHash,
        base64Data: data.base64Data,
        rawWorkbookBase64: data.rawWorkbookBase64 || data.base64Data,
        detectedArchetype: data.detectedArchetype,
        archetypeTitle: data.archetypeTitle,
        archetypeBadge: data.archetypeBadge,
        archetypeSummary: data.archetypeSummary,
        archetypeFeatures: data.archetypeFeatures,
        archetypeRecommendations: data.archetypeRecommendations,
      };

      return {
        success: true,
        analysis,
        base64Data: data.base64Data,
      };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async fetchGoogleSheet(urlOrId: string): Promise<{
    success: boolean;
    analysis?: WorkbookAnalysis;
    base64Data?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/sheets/fetch-google-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url: urlOrId }),
      });

      const data = await res.json();
      if (!data.success) {
        return { success: false, error: data.error };
      }

      const analysis: WorkbookAnalysis = {
        filename: data.filename,
        fileSize: data.fileSize,
        fileSizeFormatted: data.fileSizeFormatted,
        totalWorksheets: data.totalWorksheets,
        worksheets: data.worksheets,
        analyzedAt: data.analyzedAt,
        fileHash: data.fileHash,
        base64Data: data.base64Data,
        rawWorkbookBase64: data.rawWorkbookBase64 || data.base64Data,
        detectedArchetype: data.detectedArchetype,
        archetypeTitle: data.archetypeTitle,
        archetypeBadge: data.archetypeBadge,
        archetypeSummary: data.archetypeSummary,
        archetypeFeatures: data.archetypeFeatures,
        archetypeRecommendations: data.archetypeRecommendations,
      };

      return {
        success: true,
        analysis,
        base64Data: data.base64Data,
      };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async loadPresetWorkbook(presetId: 'timetable' | 'donations' | 'teacher_allocations' | 'jhc_inventory' | 'preset-jhc-inventory' | string): Promise<{
    success: boolean;
    analysis?: WorkbookAnalysis;
    base64Data?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/sheets/load-preset', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ presetId }),
      });

      const data = await res.json();
      if (!data.success) {
        return { success: false, error: data.error };
      }

      const analysis: WorkbookAnalysis = {
        filename: data.filename,
        fileSize: data.fileSize,
        fileSizeFormatted: data.fileSizeFormatted,
        totalWorksheets: data.totalWorksheets,
        worksheets: data.worksheets,
        analyzedAt: data.analyzedAt,
        fileHash: data.fileHash,
        base64Data: data.base64Data,
        rawWorkbookBase64: data.rawWorkbookBase64 || data.base64Data,
        detectedArchetype: data.detectedArchetype,
        archetypeTitle: data.archetypeTitle,
        archetypeBadge: data.archetypeBadge,
        archetypeSummary: data.archetypeSummary,
        archetypeFeatures: data.archetypeFeatures,
        archetypeRecommendations: data.archetypeRecommendations,
      };

      return {
        success: true,
        analysis,
        base64Data: data.base64Data,
      };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async uploadExcelFile(
    config: NextcloudConfig,
    filename: string,
    base64Content: string,
    folder?: string
  ): Promise<{
    success: boolean;
    message?: string;
    uploadUrl?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/nextcloud/upload-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          username: config.username,
          appPassword: config.appPassword,
          filename,
          base64Content,
          folder: folder || config.sourceFolder,
        }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async createNextcloudFolder(config: NextcloudConfig, folderPathOrName?: string): Promise<{
    success: boolean;
    message?: string;
    error?: string;
  }> {
    try {
      const folder = folderPathOrName || config.sourceFolder || 'ExcelImports';
      const res = await fetch('/api/nextcloud/create-folder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          username: config.username,
          appPassword: config.appPassword,
          sourceFolder: folder,
          folderName: folder,
        }),
      });

      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async testSupabaseConnection(config: SupabaseConfig): Promise<SupabaseTestResult> {
    try {
      const res = await fetch('/api/supabase/test-connection', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          anonKey: config.anonKey,
          serviceKey: config.serviceKey || config.serviceRoleKey,
        }),
      });

      const data = await res.json();
      return data;
    } catch (e: any) {
      return {
        success: false,
        isConnected: false,
        checks: {
          hostReachability: false,
          authValid: false,
          schemaDetected: false,
        },
        error: `Network error connecting to backend: ${e.message}`,
      };
    }
  }

  static async fetchSupabaseTable(
    config: SupabaseConfig,
    tableName: string,
    limit: number = 50
  ): Promise<{
    success: boolean;
    tableName?: string;
    count?: number;
    rows?: any[];
    error?: string;
  }> {
    try {
      const res = await fetch('/api/supabase/fetch-table', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          anonKey: config.anonKey,
          serviceKey: config.serviceKey || config.serviceRoleKey,
          tableName,
          limit,
        }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async upsertSupabaseRecords(
    config: SupabaseConfig,
    tableName: string,
    records: any[],
    onConflict?: string,
    onProgress?: (progress: UpsertBatchProgress) => void
  ): Promise<{
    success: boolean;
    tableName?: string;
    upsertedCount?: number;
    insertedCount?: number;
    updatedCount?: number;
    skippedCount?: number;
    failedCount?: number;
    totalRecords?: number;
    records?: any[];
    error?: string;
    warning?: string;
    message?: string;
    diagnostic?: SupabaseDiagnostic;
  }> {
    if (!records || records.length === 0) {
      return { success: true, tableName, upsertedCount: 0, insertedCount: 0, updatedCount: 0, skippedCount: 0, failedCount: 0, totalRecords: 0, records: [] };
    }

    const CHUNK_SIZE = 250;
    const totalRecords = records.length;
    const totalBatches = Math.ceil(totalRecords / CHUNK_SIZE);
    let totalUpserted = 0;
    let totalInserted = 0;
    let totalUpdated = 0;
    let totalSkipped = 0;
    let totalFailed = 0;
    const allInserted: any[] = [];

    for (let b = 0; b < totalBatches; b++) {
      const start = b * CHUNK_SIZE;
      const end = Math.min(start + CHUNK_SIZE, totalRecords);
      const batch = records.slice(start, end);
      const currentBatchNum = b + 1;

      if (onProgress) {
        onProgress({
          processed: start,
          total: totalRecords,
          percentage: Math.round((start / totalRecords) * 100),
          currentBatch: currentBatchNum,
          totalBatches,
          successfulCount: totalUpserted,
          failedCount: totalFailed,
          tableName,
          status: 'running',
          message: `Pushing batch ${currentBatchNum}/${totalBatches} (${batch.length} rows) to public.${tableName}...`
        });
      }

      try {
        const res = await fetch('/api/supabase/upsert-records', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            url: config.url,
            anonKey: config.anonKey,
            serviceKey: config.serviceKey || config.serviceRoleKey,
            tableName,
            records: batch,
            onConflict,
          }),
        });

        const data = await res.json();

        if (!data.success) {
          totalFailed += batch.length;
          if (onProgress) {
            onProgress({
              processed: start,
              total: totalRecords,
              percentage: Math.round((start / totalRecords) * 100),
              currentBatch: currentBatchNum,
              totalBatches,
              successfulCount: totalUpserted,
              failedCount: totalFailed,
              tableName,
              status: 'error',
              message: `Batch ${currentBatchNum}/${totalBatches} failed: ${data.error || 'Check Supabase table'}`
            });
          }
          return {
            success: false,
            tableName,
            upsertedCount: totalUpserted,
            insertedCount: totalInserted,
            updatedCount: totalUpdated,
            skippedCount: totalSkipped,
            failedCount: totalFailed,
            totalRecords,
            error: data.error,
            diagnostic: data.diagnostic,
          };
        }

        const count = data.upsertedCount || (data.insertedCount || 0) + (data.updatedCount || 0);
        totalUpserted += count;
        totalInserted += data.insertedCount || 0;
        totalUpdated += data.updatedCount || 0;
        totalSkipped += data.skippedCount || 0;

        if (Array.isArray(data.records)) {
          allInserted.push(...data.records);
        }

        if (onProgress) {
          onProgress({
            processed: end,
            total: totalRecords,
            percentage: Math.round((end / totalRecords) * 100),
            currentBatch: currentBatchNum,
            totalBatches,
            successfulCount: totalUpserted,
            failedCount: 0,
            tableName,
            status: end === totalRecords ? 'success' : 'running',
            message: `Batch ${currentBatchNum}/${totalBatches} confirmed! (${end}/${totalRecords} rows checked, ${totalSkipped} identical skipped)`
          });
        }
      } catch (e: any) {
        totalFailed += batch.length;
        if (onProgress) {
          onProgress({
            processed: start,
            total: totalRecords,
            percentage: Math.round((start / totalRecords) * 100),
            currentBatch: currentBatchNum,
            totalBatches,
            successfulCount: totalUpserted,
            failedCount: totalFailed,
            tableName,
            status: 'error',
            message: `Network error on batch ${currentBatchNum}: ${e.message}`
          });
        }
        return {
          success: false,
          tableName,
          upsertedCount: totalUpserted,
          insertedCount: totalInserted,
          updatedCount: totalUpdated,
          skippedCount: totalSkipped,
          failedCount: totalFailed,
          totalRecords,
          error: `Network error pushing to Supabase: ${e.message}`,
        };
      }
    }

    return {
      success: true,
      tableName,
      upsertedCount: totalUpserted,
      insertedCount: totalInserted,
      updatedCount: totalUpdated,
      skippedCount: totalSkipped,
      failedCount: totalFailed,
      totalRecords,
      records: allInserted,
      message: `Successfully processed ${totalRecords} records: ${totalSkipped} identical skipped (no duplicates), ${totalUpdated} cell-updated in-place, ${totalInserted} new records inserted.`
    };
  }

  static async pushMappingsToSupabase(
    config: SupabaseConfig,
    mappings: WorksheetMapping[]
  ): Promise<{
    success: boolean;
    pushedWorksheetMappingsCount: number;
    pushedColumnMappingsCount: number;
    message?: string;
    error?: string;
    diagnostic?: SupabaseDiagnostic;
  }> {
    if (!config.url || (!config.anonKey && !config.serviceKey && !config.serviceRoleKey)) {
      return {
        success: false,
        pushedWorksheetMappingsCount: 0,
        pushedColumnMappingsCount: 0,
        error: 'Supabase URL and API Key are required. Please configure them in Supabase tab or Secrets Vault.'
      };
    }

    const userMappings = mappings.filter(m => m.isUserConfigured || m.enabled !== false);
    if (userMappings.length === 0) {
      return {
        success: false,
        pushedWorksheetMappingsCount: 0,
        pushedColumnMappingsCount: 0,
        error: 'No active user-configured mappings found to push. Please configure or enable mappings first.'
      };
    }

    try {
      const res = await fetch('/api/supabase/push-mappings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          anonKey: config.anonKey,
          serviceKey: config.serviceKey || config.serviceRoleKey,
          mappings: userMappings,
        }),
      });

      const data = await res.json();
      return data;
    } catch (e: any) {
      return {
        success: false,
        pushedWorksheetMappingsCount: 0,
        pushedColumnMappingsCount: 0,
        error: `Network error calling backend push-mappings: ${e.message}`,
      };
    }
  }

  static async parseRawExcelOrCsv(params: {
    base64Data?: string;
    rawText?: string;
    filename?: string;
  }): Promise<{
    success: boolean;
    analysis?: WorkbookAnalysis;
    base64Data?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/excel/parse-raw', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });

      const data = await res.json();
      if (!data.success) {
        return { success: false, error: data.error };
      }

      const analysis: WorkbookAnalysis = {
        filename: data.filename,
        fileSize: data.fileSize,
        fileSizeFormatted: data.fileSizeFormatted,
        totalWorksheets: data.totalWorksheets,
        worksheets: data.worksheets,
        analyzedAt: data.analyzedAt,
        fileHash: data.fileHash,
        base64Data: data.base64Data,
      };

      return {
        success: true,
        analysis,
        base64Data: data.base64Data,
      };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async executeFullPipelineSync(params: {
    nextcloud: NextcloudConfig;
    supabase: SupabaseConfig;
    mappings: any[];
    targetFilename?: string;
    syncAllFiles?: boolean;
    base64Workbook?: string;
  }): Promise<{
    success: boolean;
    filename?: string;
    filesSynced?: string[];
    fileHash?: string;
    totalInserted?: number;
    totalUpdated?: number;
    totalFailed?: number;
    syncResults?: any[];
    errors?: any[];
    executedAt?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/sync/execute-full-pipeline', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async refreshAllNextcloudWorkbooks(params: {
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    mappings?: any[];
  }): Promise<{
    success: boolean;
    refreshed?: {
      downloadedFiles: string[];
      deletedFiles?: string[];
      prunedMappingsCount?: number;
      mappingsAdded: number;
      mappings?: WorksheetMapping[];
      executedAt: string;
    };
    pipeline?: any;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/nextcloud/refresh-all-workbooks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async getWorkerDiagnostics(params: {
    nextcloud: NextcloudConfig;
    supabase: SupabaseConfig;
    workerUrl?: string;
  }): Promise<{
    success: boolean;
    diagnostics?: {
      timestamp: string;
      nextcloud: { reachable: boolean; status: string; details?: any; latencyMs?: number };
      supabase: { reachable: boolean; status: string; tablesCount?: number; tables?: string[]; latencyMs?: number };
      workerService: { reachable: boolean; status: string; endpoint: string; latencyMs?: number };
      allSystemsReady: boolean;
    };
    error?: string;
  }> {
    try {
      const res = await fetch('/api/worker/diagnostics', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async getSupabaseSchema(config: SupabaseConfig): Promise<{
    success: boolean;
    tables: SupabaseTableInfo[];
    tablesCount?: number;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/supabase/schema', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          anonKey: config.anonKey,
          serviceKey: config.serviceKey,
          serviceRoleKey: config.serviceRoleKey,
        }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, tables: [], error: e.message };
    }
  }

  static async executeSupabaseDdl(
    config: SupabaseConfig,
    sql: string,
    tableName?: string
  ): Promise<{
    success: boolean;
    directExecuted?: boolean;
    message?: string;
    sql?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/supabase/execute-ddl', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          anonKey: config.anonKey,
          serviceKey: config.serviceKey,
          serviceRoleKey: config.serviceRoleKey,
          sql,
          tableName,
        }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async fetchSupabaseTableRows(
    config: SupabaseConfig,
    tableName: string,
    limit: number = 50
  ): Promise<{
    success: boolean;
    tableName?: string;
    count?: number;
    rows?: any[];
    error?: string;
  }> {
    try {
      const res = await fetch('/api/supabase/fetch-table', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          anonKey: config.anonKey,
          serviceKey: config.serviceKey,
          serviceRoleKey: config.serviceRoleKey,
          tableName,
          limit,
        }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, rows: [], error: e.message };
    }
  }

  static async testAllSecrets(payload: {
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    workerUrl?: string;
  }): Promise<{
    success: boolean;
    report?: {
      timestamp: string;
      allConnected: boolean;
      nextcloud: {
        isConnected: boolean;
        latencyMs: number;
        statusText: string;
        checks: { hostReachability: boolean; authValid: boolean; folderAccessible: boolean };
        error?: string;
      };
      supabase: {
        isConnected: boolean;
        latencyMs: number;
        statusText: string;
        tablesCount: number;
        tables: string[];
        checks: { hostReachability: boolean; authValid: boolean; schemaDetected: boolean };
        error?: string;
      };
      worker: {
        isConnected: boolean;
        latencyMs: number;
        statusText: string;
        endpoint: string;
        error?: string;
      };
    };
    error?: string;
  }> {
    try {
      const res = await fetch('/api/secrets/test-all', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async updateWorkerSecrets(payload: any): Promise<{
    success: boolean;
    relayedToWorker?: boolean;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/worker/secrets/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async getSchedulerStatus(): Promise<{
    success: boolean;
    status?: LiveSchedulerStatus;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/scheduler/status');
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async configureScheduler(payload: {
    enabled?: boolean;
    intervalLabel?: SyncInterval;
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    mappings?: any[];
    workerUrl?: string;
  }): Promise<{
    success: boolean;
    status?: LiveSchedulerStatus;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/scheduler/configure', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async triggerSchedulerNow(payload?: {
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    mappings?: WorksheetMapping[];
  }): Promise<{
    success: boolean;
    result?: any;
    status?: LiveSchedulerStatus;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/scheduler/trigger-now', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: payload ? JSON.stringify(payload) : undefined,
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async testAutomatedRun(payload: {
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    mappings?: WorksheetMapping[];
    base64Workbook?: string;
    filename?: string;
  }): Promise<AutomatedRunDiagnosticResult> {
    try {
      const res = await fetch('/api/scheduler/test-automated-run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return {
        success: false,
        verdict: 'FAILED',
        totalDurationMs: 0,
        timestamp: new Date().toISOString(),
        stages: [
          {
            name: 'COMMUNICATION',
            label: 'Client-Server Communication',
            status: 'FAILED',
            durationMs: 0,
            message: `Could not reach backend test runner: ${e.message}`,
          }
        ],
        summary: {
          sourceType: 'UNKNOWN',
          filename: payload.filename || 'unknown.xlsx',
          sheetsProcessed: 0,
          targetTables: [],
          rowsInserted: 0,
          rowsUpdated: 0,
          rowsFailed: 0,
          verificationRowCount: 0,
        },
        recommendations: [
          'Verify that the local development server or backend service is running and responsive.'
        ],
        rawError: e.message,
      };
    }
  }

  static async pingWorker(payload?: {
    workerUrl?: string;
    secretKey?: string;
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
  }): Promise<WorkerConnectionStatus> {
    try {
      const res = await fetch('/api/worker/ping', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: payload ? JSON.stringify(payload) : undefined,
      });
      const data = await res.json();
      return data;
    } catch (e: any) {
      return {
        connected: false,
        workerMode: 'INTEGRATED_PRODUCTION_ENGINE',
        workerEndpoint: payload?.workerUrl || 'internal',
        latencyMs: 0,
        handshakeVerified: false,
        version: 'unknown',
        uptimeSeconds: 0,
        lastHeartbeat: new Date().toISOString(),
        state: 'ERROR',
        activeInterval: '15m',
        nextRunAt: null,
        secretsMatched: false,
        diagnostics: {
          nextcloudReachable: false,
          supabaseReachable: false,
          message: `Network error pinging worker daemon: ${e.message}`,
        },
      };
    }
  }

  static async getLiveLogs(params?: {
    level?: string;
    search?: string;
    limit?: number;
  }): Promise<{
    success: boolean;
    count: number;
    totalCount: number;
    logs: LogMessage[];
    error?: string;
  }> {
    try {
      const qs = new URLSearchParams();
      if (params?.level) qs.set('level', params.level);
      if (params?.search) qs.set('search', params.search);
      if (params?.limit) qs.set('limit', String(params.limit));

      const res = await fetch(`/api/logs?${qs.toString()}`, { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, count: 0, totalCount: 0, logs: [], error: e.message };
    }
  }

  static async addServerLog(log: {
    level: 'info' | 'warn' | 'error' | 'success';
    component: string;
    message: string;
    details?: any;
  }): Promise<{ success: boolean; log?: any }> {
    try {
      const res = await fetch('/api/logs/add', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(log),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false };
    }
  }

  static async clearServerLogs(): Promise<{ success: boolean }> {
    try {
      const res = await fetch('/api/logs/clear', { method: 'POST' });
      return await res.json();
    } catch (e: any) {
      return { success: false };
    }
  }

  static async getSyncHistory(): Promise<{
    success: boolean;
    count: number;
    history: LiveSyncHistoryItem[];
    error?: string;
  }> {
    try {
      const res = await fetch('/api/sync/history', { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, count: 0, history: [], error: e.message };
    }
  }

  static async clearSyncHistory(): Promise<{ success: boolean }> {
    try {
      const res = await fetch('/api/sync/history/clear', { method: 'POST' });
      return await res.json();
    } catch (e: any) {
      return { success: false };
    }
  }

  static async triggerLiveSync(payload: {
    nextcloud: NextcloudConfig;
    supabase: SupabaseConfig;
    mappings: WorksheetMapping[];
    targetFilename?: string;
    base64Workbook?: string;
  }): Promise<{
    success: boolean;
    result?: any;
    status?: LiveSchedulerStatus;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/sync/trigger', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  // ==========================================
  // Permanent Secrets & Mappings Persistence
  // ==========================================
  static async savePermanentSecrets(payload: {
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    syncSettings?: SyncSettings;
    geminiApiKey?: string;
  }): Promise<{
    success: boolean;
    savedOnDisk?: boolean;
    savedInDb?: boolean;
    savedAt?: string;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/secrets/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async loadPermanentSecrets(): Promise<{
    success: boolean;
    secrets?: {
      nextcloud?: NextcloudConfig;
      supabase?: SupabaseConfig;
      syncSettings?: SyncSettings;
      geminiApiKey?: string;
      savedAt?: string;
    } | null;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/secrets/load', { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async importPermanentSecrets(secrets: any): Promise<{
    success: boolean;
    secrets?: any;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/secrets/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(secrets),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async savePermanentMappings(payload: {
    mappings: WorksheetMapping[];
    workbookInfo?: any;
    supabase?: SupabaseConfig;
  }): Promise<{
    success: boolean;
    savedOnDisk?: boolean;
    syncedToSupabase?: boolean;
    savedAt?: string;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/mappings/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async loadPermanentMappings(): Promise<{
    success: boolean;
    mappings?: WorksheetMapping[];
    workbookInfo?: any;
    savedAt?: string;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/mappings/load', { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  // ==========================================
  // Permanent Served Sheets & Mapping Hubs
  // ==========================================
  static async saveServedSheets(payload: {
    servedSheets?: any[];
    presets?: any[];
    mappings?: WorksheetMapping[];
    workbookInfo?: any;
  }): Promise<{
    success: boolean;
    savedAt?: string;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/served-sheets/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async loadServedSheets(): Promise<{
    success: boolean;
    servedSheets?: any[];
    presets?: any[];
    workbooks?: Array<{ filename: string; size: number; lastModified: string }>;
    mappings?: WorksheetMapping[];
    savedAt?: string | null;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/served-sheets', { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async getStoredWorkbooks(): Promise<{
    success: boolean;
    workbooks?: Array<{ filename: string; size: number; lastModified: string }>;
    totalCount?: number;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/workbooks', { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async syncMappingsToSupabase(payload: {
    mappings: WorksheetMapping[];
    workbookInfo?: any;
    supabase: SupabaseConfig;
  }): Promise<{
    success: boolean;
    worksheetsCount?: number;
    columnsCount?: number;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/mappings/sync-to-supabase', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async analyzeWorkbookWithAi(payload: {
    workbook: WorkbookAnalysis;
    supabaseTables?: SupabaseTableInfo[];
    consolidationMode?: string;
  }): Promise<{
    success: boolean;
    aiPowered?: boolean;
    fallbackActive?: boolean;
    fallbackNotice?: string;
    result?: AiWorkbookAnalysisResult;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/ai/analyze-workbook', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  // ==========================================
  // Industrial Standard AI Presets Management
  // ==========================================
  static async getPresets(): Promise<{
    success: boolean;
    presets: AiWorkbookPreset[];
    count?: number;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/presets', { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, presets: [], error: e.message };
    }
  }

  static async savePreset(payload: {
    preset: Partial<AiWorkbookPreset>;
    supabase?: SupabaseConfig;
  }): Promise<{
    success: boolean;
    preset?: AiWorkbookPreset;
    savedOnDisk?: boolean;
    syncedToSupabase?: boolean;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/presets/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async deletePreset(
    id: string,
    supabase?: SupabaseConfig
  ): Promise<{ success: boolean; message?: string; error?: string }> {
    try {
      const res = await fetch(`/api/presets/${encodeURIComponent(id)}`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ supabase }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async loadPresetDetails(presetId: string): Promise<{
    success: boolean;
    preset?: AiWorkbookPreset;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/presets/load', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ presetId }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async createPresetWithAi(payload: {
    workbook: WorkbookAnalysis;
    supabaseTables?: SupabaseTableInfo[];
    customName?: string;
    customDescription?: string;
    tags?: string[];
  }): Promise<{
    success: boolean;
    preset?: AiWorkbookPreset;
    aiPowered?: boolean;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/ai/create-preset', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  // ==========================================
  // Coolify & Server Unified State Management
  // ==========================================
  static async getFullServerState(): Promise<{
    success: boolean;
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    syncSettings?: SyncSettings;
    mappings?: WorksheetMapping[];
    workbookInfo?: any;
    presets?: AiWorkbookPreset[];
    geminiApiKey?: string;
    workerStatus?: LiveSchedulerStatus;
    envSources?: Record<string, boolean>;
    savedAt?: string;
    serverBootTime?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/config/full-state', { method: 'GET' });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async saveFullServerState(payload: {
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    syncSettings?: SyncSettings;
    mappings?: WorksheetMapping[];
    presets?: AiWorkbookPreset[];
    workbookInfo?: any;
    geminiApiKey?: string;
  }): Promise<{
    success: boolean;
    savedOnDisk?: boolean;
    syncedToSupabase?: boolean;
    savedAt?: string;
    status?: LiveSchedulerStatus;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/config/save-full-state', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async syncWithBrowser(payload: {
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    syncSettings?: SyncSettings;
    mappings?: WorksheetMapping[];
    presets?: AiWorkbookPreset[];
    workbookInfo?: any;
    geminiApiKey?: string;
  }): Promise<{
    success: boolean;
    nextcloud?: NextcloudConfig;
    supabase?: SupabaseConfig;
    syncSettings?: SyncSettings;
    mappings?: WorksheetMapping[];
    presets?: AiWorkbookPreset[];
    workerStatus?: LiveSchedulerStatus;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/config/sync-with-browser', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async startWorkerDaemon(): Promise<{ success: boolean; message?: string; status?: LiveSchedulerStatus }> {
    try {
      const res = await fetch('/api/worker/start', { method: 'POST' });
      return await res.json();
    } catch (e: any) {
      return { success: false, message: e.message };
    }
  }

  static async stopWorkerDaemon(): Promise<{ success: boolean; message?: string; status?: LiveSchedulerStatus }> {
    try {
      const res = await fetch('/api/worker/stop', { method: 'POST' });
      return await res.json();
    } catch (e: any) {
      return { success: false, message: e.message };
    }
  }

  static async triggerWorkerRunNow(payload?: any): Promise<{
    success: boolean;
    result?: any;
    status?: LiveSchedulerStatus;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/worker/run-now', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload || {}),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async getGradesServerStorage(): Promise<{
    success: boolean;
    exists?: boolean;
    data?: {
      files: any[];
      records: any[];
      referenceColumns: string[];
      summary: any;
      histories: any[];
      lastSaved: string | null;
    };
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/storage');
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async saveGradesServerStorage(payload: {
    files: any[];
    records: any[];
    referenceColumns: string[];
    summary: any;
    histories: any[];
  }): Promise<{ success: boolean; message?: string; lastSaved?: string; error?: string }> {
    try {
      const res = await fetch('/api/grades/storage', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async saveWorkbookFileOnServer(fileName: string, fileData: string): Promise<{
    success: boolean;
    message?: string;
    path?: string;
    fileSize?: number;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/save-workbook-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileName, fileData }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async saveStudentHistoryOnServer(history: any): Promise<{
    success: boolean;
    histories?: any[];
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/history', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ history }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async clearGradesServerStorage(): Promise<{ success: boolean; message?: string; error?: string }> {
    try {
      const res = await fetch('/api/grades/storage', { method: 'DELETE' });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async getNextcloudGradeSyncConfig(): Promise<{
    success: boolean;
    config?: import('../types').NextcloudGradeSyncConfig;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/nextcloud/sync-config');
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async saveNextcloudGradeSyncConfig(config: Partial<import('../types').NextcloudGradeSyncConfig>): Promise<{
    success: boolean;
    message?: string;
    config?: import('../types').NextcloudGradeSyncConfig;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/nextcloud/sync-config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(config),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async browseNextcloudGradeFiles(params: {
    url?: string;
    username?: string;
    appPassword?: string;
    sourceFolder?: string;
  }): Promise<{
    success: boolean;
    folderPath?: string;
    files?: {
      filename: string;
      path: string;
      fileSize: number;
      fileSizeFormatted: string;
      lastModified: string;
      etag: string;
      isDirectory: boolean;
      isSpreadsheet: boolean;
    }[];
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/nextcloud/browse', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async pullNextcloudGradeWorkbooks(params: {
    url?: string;
    username?: string;
    appPassword?: string;
    sourceFolder?: string;
    selectedFiles?: string[];
  }): Promise<{
    success: boolean;
    filesPulled?: number;
    recordsCount?: number;
    message?: string;
    files?: any[];
    records?: any[];
    referenceColumns?: string[];
    summary?: any;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/nextcloud/pull', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async triggerNextcloudGradeSyncNow(): Promise<{
    success: boolean;
    filesPulled?: number;
    recordsCount?: number;
    message?: string;
    files?: any[];
    records?: any[];
    referenceColumns?: string[];
    summary?: any;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/grades/nextcloud/sync-now', {
        method: 'POST',
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async downloadNextcloudFileBase64(params: {
    url?: string;
    username?: string;
    appPassword?: string;
    filePath?: string;
    filename?: string;
  }): Promise<{
    success: boolean;
    filename?: string;
    fileSize?: number;
    fileSizeFormatted?: string;
    base64Data?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/nextcloud/download-file-base64', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }

  static async cacheActiveWorkbook(payload: {
    base64Data: string;
    filename?: string;
  }): Promise<{
    success: boolean;
    filename?: string;
    fileSize?: number;
    fileHash?: string;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/workbook/cache-active', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}

