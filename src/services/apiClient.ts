import { NextcloudConfig, SupabaseConfig, NextcloudFile, WorkbookAnalysis } from '../types';

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
}
