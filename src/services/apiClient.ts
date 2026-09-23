import { NextcloudConfig, SupabaseConfig, NextcloudFile, WorkbookAnalysis, SupabaseTableInfo, LiveSchedulerStatus, SyncInterval } from '../types';

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

  static async createNextcloudFolder(config: NextcloudConfig, folderName: string): Promise<{
    success: boolean;
    message?: string;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/nextcloud/create-folder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          username: config.username,
          appPassword: config.appPassword,
          folderName,
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
    onConflict?: string
  ): Promise<{
    success: boolean;
    tableName?: string;
    upsertedCount?: number;
    records?: any[];
    error?: string;
  }> {
    try {
      const res = await fetch('/api/supabase/upsert-records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: config.url,
          anonKey: config.anonKey,
          serviceKey: config.serviceKey || config.serviceRoleKey,
          tableName,
          records,
          onConflict,
        }),
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
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
  }): Promise<{
    success: boolean;
    filename?: string;
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

  static async triggerSchedulerNow(): Promise<{
    success: boolean;
    result?: any;
    status?: LiveSchedulerStatus;
    error?: string;
  }> {
    try {
      const res = await fetch('/api/scheduler/trigger-now', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
      });
      return await res.json();
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
