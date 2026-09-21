import { NextcloudConfig, NextcloudFile, WorkbookAnalysis } from '../types';

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
}
