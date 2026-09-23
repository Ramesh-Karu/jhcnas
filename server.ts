import express from 'express';
import path from 'path';
import crypto from 'crypto';
import * as XLSX from 'xlsx';
import { GoogleGenAI } from '@google/genai';
import { createServer as createViteServer } from 'vite';

const app = express();
const PORT = Number(process.env.PORT) || 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Healthcheck endpoint for Coolify / Docker
app.get('/health', (_req, res) => {
  res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
});

let aiClient: GoogleGenAI | null = null;
function getGemini(): GoogleGenAI | null {
  if (!aiClient && process.env.GEMINI_API_KEY) {
    aiClient = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
  }
  return aiClient;
}

// Helper to construct basic auth header
function getBasicAuth(username: string, appPassword: string) {
  return Buffer.from(`${username}:${appPassword}`).toString('base64');
}

// 1. Live Nextcloud Connection Test
app.post('/api/nextcloud/test-connection', async (req, res) => {
  try {
    const { url, webdavUrl, username, appPassword, sourceFolder } = req.body;
    const host = (url || 'https://cloud.jhcnexus.space').replace(/\/+$/, '');
    const user = username || 'truenas_admin';
    const pass = appPassword || 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT';
    const folder = (sourceFolder || '/ExcelImports').replace(/^\/+/, '');

    const checks: {
      hostReachability: boolean;
      webdavHandshake: boolean;
      authValid: boolean;
      folderExists: boolean;
      details?: any;
    } = {
      hostReachability: false,
      webdavHandshake: false,
      authValid: false,
      folderExists: false,
    };

    // Step 1: Check host reachability via Nextcloud status.php
    try {
      const statusRes = await fetch(`${host}/status.php`, { method: 'GET' });
      if (statusRes.ok) {
        checks.hostReachability = true;
        const statusJson = await statusRes.json();
        checks.details = statusJson;
      }
    } catch (e: any) {
      return res.status(502).json({
        success: false,
        checks,
        error: `Failed to reach Nextcloud host at ${host}: ${e.message}`,
      });
    }

    // Step 2 & 3: Check WebDAV handshake and authentication
    const authHeader = `Basic ${getBasicAuth(user, pass)}`;
    const rootWebdav = `${host}/remote.php/dav/files/${encodeURIComponent(user)}/`;

    try {
      const davRes = await fetch(rootWebdav, {
        method: 'PROPFIND',
        headers: {
          Authorization: authHeader,
          Depth: '0',
        },
      });

      if (davRes.status === 401 || davRes.status === 403) {
        return res.status(401).json({
          success: false,
          checks: { ...checks, webdavHandshake: true, authValid: false },
          error: `Authentication failed (HTTP ${davRes.status}). Verify app password for user '${user}'.`,
        });
      }

      if (davRes.status === 207 || davRes.ok) {
        checks.webdavHandshake = true;
        checks.authValid = true;
      } else {
        checks.webdavHandshake = true;
      }
    } catch (e: any) {
      return res.status(502).json({
        success: false,
        checks,
        error: `WebDAV handshake error at ${rootWebdav}: ${e.message}`,
      });
    }

    // Step 4: Check if source folder exists
    const folderUrl = `${host}/remote.php/dav/files/${encodeURIComponent(user)}/${folder ? `${folder}/` : ''}`;
    try {
      const folderRes = await fetch(folderUrl, {
        method: 'PROPFIND',
        headers: {
          Authorization: authHeader,
          Depth: '0',
        },
      });

      if (folderRes.status === 207 || folderRes.ok) {
        checks.folderExists = true;
      } else if (folderRes.status === 404) {
        // Folder doesn't exist yet, but credentials and host are valid
        checks.folderExists = false;
      }
    } catch {
      checks.folderExists = false;
    }

    return res.json({
      success: true,
      checks,
      message: checks.folderExists
        ? `Successfully connected to Nextcloud ${checks.details?.versionstring || ''} and verified WebDAV access to /${folder}`
        : `Connected to Nextcloud successfully, but folder /${folder} does not exist yet. You can create it with 1 click.`,
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 2. List Real Files from Nextcloud WebDAV
app.post('/api/nextcloud/list-files', async (req, res) => {
  try {
    const { url, username, appPassword, sourceFolder } = req.body;
    const host = (url || 'https://cloud.jhcnexus.space').replace(/\/+$/, '');
    const user = username || 'truenas_admin';
    const pass = appPassword || 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT';
    const folder = (sourceFolder || '/ExcelImports').replace(/^\/+/, '').replace(/\/+$/, '');

    const authHeader = `Basic ${getBasicAuth(user, pass)}`;
    const folderUrl = folder
      ? `${host}/remote.php/dav/files/${encodeURIComponent(user)}/${folder}/`
      : `${host}/remote.php/dav/files/${encodeURIComponent(user)}/`;

    const davRes = await fetch(folderUrl, {
      method: 'PROPFIND',
      headers: {
        Authorization: authHeader,
        Depth: '1',
      },
    });

    if (!davRes.ok && davRes.status !== 207) {
      return res.status(davRes.status).json({
        success: false,
        error: `Nextcloud WebDAV returned HTTP ${davRes.status} (${davRes.statusText}) for ${folderUrl}`,
      });
    }

    const xml = await davRes.text();

    // Parse multi-status XML items
    const responseBlocks = xml.split(/<d:response>/i).slice(1);
    const files: any[] = [];

    for (const block of responseBlocks) {
      const hrefMatch = block.match(/<d:href>(.*?)<\/d:href>/i);
      if (!hrefMatch) continue;

      const rawHref = decodeURIComponent(hrefMatch[1]);
      const isDir = /<d:resourcetype>[\s\S]*?<d:collection\/>[\s\S]*?<\/d:resourcetype>/i.test(block);

      const lengthMatch = block.match(/<d:getcontentlength>(.*?)<\/d:getcontentlength>/i);
      const modifiedMatch = block.match(/<d:getlastmodified>(.*?)<\/d:getlastmodified>/i);
      const etagMatch = block.match(/<d:getetag>(.*?)<\/d:getetag>/i);
      const contentTypeMatch = block.match(/<d:getcontenttype>(.*?)<\/d:getcontenttype>/i);

      // Extract filename from href
      const trimmedHref = rawHref.replace(/\/+$/, '');
      const filename = trimmedHref.split('/').pop() || '';

      // Skip the folder itself (the root response)
      const targetFolderClean = folder ? `/${folder}` : '';
      if (trimmedHref.endsWith(targetFolderClean) || trimmedHref.endsWith(`/files/${user}`)) {
        continue;
      }

      const fileSize = lengthMatch ? parseInt(lengthMatch[1], 10) : 0;
      const cleanEtag = etagMatch ? etagMatch[1].replace(/"/g, '') : '';
      const sha256Hash = crypto.createHash('sha256').update(cleanEtag || filename).digest('hex');

      files.push({
        id: `nc-${filename}-${Date.now()}`,
        filename,
        path: rawHref,
        fileSize,
        fileSizeFormatted: formatBytes(fileSize),
        lastModified: modifiedMatch ? new Date(modifiedMatch[1]).toISOString() : new Date().toISOString(),
        fileHash: sha256Hash,
        etag: cleanEtag,
        isDirectory: isDir,
        contentType: contentTypeMatch ? contentTypeMatch[1] : '',
        status: 'Synced',
      });
    }

    return res.json({
      success: true,
      folderPath: folderUrl,
      count: files.length,
      files,
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 3. Create Folder in Nextcloud WebDAV
app.post('/api/nextcloud/create-folder', async (req, res) => {
  try {
    const { url, username, appPassword, folderName } = req.body;
    const host = (url || 'https://cloud.jhcnexus.space').replace(/\/+$/, '');
    const user = username || 'truenas_admin';
    const pass = appPassword || 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT';
    const folder = (folderName || 'ExcelImports').replace(/^\/+/, '').replace(/\/+$/, '');

    const authHeader = `Basic ${getBasicAuth(user, pass)}`;
    const folderUrl = `${host}/remote.php/dav/files/${encodeURIComponent(user)}/${folder}/`;

    const mkcolRes = await fetch(folderUrl, {
      method: 'MKCOL',
      headers: {
        Authorization: authHeader,
      },
    });

    if (mkcolRes.status === 201) {
      return res.json({ success: true, message: `Created folder /${folder} in Nextcloud` });
    } else if (mkcolRes.status === 405) {
      return res.json({ success: true, message: `Folder /${folder} already exists in Nextcloud` });
    } else {
      return res.status(mkcolRes.status).json({
        success: false,
        error: `Nextcloud MKCOL returned HTTP ${mkcolRes.status}: ${mkcolRes.statusText}`,
      });
    }
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 4. Download Real Excel Workbook from Nextcloud WebDAV and Parse It
app.post('/api/nextcloud/fetch-and-parse', async (req, res) => {
  try {
    const { url, username, appPassword, filePath, filename } = req.body;
    const host = (url || 'https://cloud.jhcnexus.space').replace(/\/+$/, '');
    const user = username || 'truenas_admin';
    const pass = appPassword || 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT';

    let targetUrl: string;
    if (filePath && filePath.startsWith('http')) {
      targetUrl = filePath;
    } else if (filePath) {
      targetUrl = `${host}${filePath.startsWith('/') ? '' : '/'}${filePath}`;
    } else {
      targetUrl = `${host}/remote.php/dav/files/${encodeURIComponent(user)}/ExcelImports/${filename || 'students.xlsx'}`;
    }

    const authHeader = `Basic ${getBasicAuth(user, pass)}`;
    const fileRes = await fetch(targetUrl, {
      method: 'GET',
      headers: { Authorization: authHeader },
    });

    if (!fileRes.ok) {
      return res.status(fileRes.status).json({
        success: false,
        error: `Failed to download file from ${targetUrl}: HTTP ${fileRes.status} ${fileRes.statusText}`,
      });
    }

    const arrayBuffer = await fileRes.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');

    // Parse with XLSX
    const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const worksheets: any[] = [];

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      const totalRows = range.e.r + 1;
      const totalColumns = range.e.c + 1;
      const merges = ws['!merges'] || [];

      const usedRangeStr = ws['!ref'] || `${XLSX.utils.encode_col(range.s.c)}${range.s.r + 1}:${XLSX.utils.encode_col(range.e.c)}${range.e.r + 1}`;

      // Helper to resolve merged cells value
      const getResolvedVal = (r: number, c: number) => {
        const direct = ws[XLSX.utils.encode_cell({ r, c })];
        if (direct && direct.v !== undefined && direct.v !== null && String(direct.v).trim() !== '') {
          return direct.w !== undefined ? direct.w : direct.v;
        }
        const m = merges.find((m: any) => m.s.r <= r && r <= m.e.r && m.s.c <= c && c <= m.e.c);
        if (m) {
          const orig = ws[XLSX.utils.encode_cell(m.s)];
          if (orig && orig.v !== undefined && orig.v !== null) {
            return orig.w !== undefined ? orig.w : orig.v;
          }
        }
        return '';
      };

      // Extract detected headers (Row 0 / Row 1 in 1-based indexing)
      const detectedHeaderRow = 1;
      const detectedDataStartRow = 2;

      const headers = [];
      for (let c = 0; c < totalColumns; c++) {
        const colLetter = XLSX.utils.encode_col(c);
        const headerVal = getResolvedVal(0, c);
        const name = String(headerVal || `Column ${colLetter}`).trim();
        const sampleValues = [];
        for (let r = 1; r < Math.min(totalRows, 7); r++) {
          sampleValues.push(String(getResolvedVal(r, c)));
        }
        headers.push({
          colLetter,
          colIndex: c + 1,
          name,
          sampleValues,
        });
      }

      // Sample rows with resolved merged cells
      const sampleRows = [];
      for (let r = 1; r < Math.min(totalRows, 15); r++) {
        const rowData: Record<string, any> = {};
        headers.forEach((h, hIdx) => {
          const val = getResolvedVal(r, hIdx);
          rowData[h.colLetter] = val;
          if (h.name) {
            rowData[h.name] = val;
          }
        });
        sampleRows.push({
          rowNumber: r + 1,
          data: rowData,
        });
      }

      worksheets.push({
        sheetName,
        totalRows,
        totalColumns,
        usedRange: usedRangeStr,
        mergedRanges: merges.map((m: any) => ({
          range: `${XLSX.utils.encode_col(m.s.c)}${m.s.r + 1}:${XLSX.utils.encode_col(m.e.c)}${m.e.r + 1}`,
          startCol: XLSX.utils.encode_col(m.s.c),
          endCol: XLSX.utils.encode_col(m.e.c),
          startRow: m.s.r + 1,
          endRow: m.e.r + 1,
          value: String(ws[XLSX.utils.encode_cell(m.s)]?.w ?? ws[XLSX.utils.encode_cell(m.s)]?.v ?? ''),
          type: 'data_span',
        })),
        detectedHeaderRow,
        detectedDataStartRow,
        candidateHeaderRows: [{ row: 1, headers: headers.map((h) => h.name), confidence: 0.95 }],
        titleRows: [],
        sectionHeadings: [],
        emptyRowsCount: 0,
        repeatedHeadersCount: 0,
        headers,
        sampleRows,
      });
    }

    const base64Data = buffer.toString('base64');

    return res.json({
      success: true,
      filename: filename || 'students.xlsx',
      fileSize: buffer.byteLength,
      fileSizeFormatted: formatBytes(buffer.byteLength),
      fileHash: sha256,
      totalWorksheets: worksheets.length,
      worksheets,
      analyzedAt: new Date().toISOString(),
      base64Data,
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 5. Upload File to Nextcloud WebDAV
app.post('/api/nextcloud/upload-file', async (req, res) => {
  try {
    const { url, username, appPassword, filename, base64Content, folder } = req.body;
    const host = (url || 'https://cloud.jhcnexus.space').replace(/\/+$/, '');
    const user = username || 'truenas_admin';
    const pass = appPassword || 'mpxC4-dk7jn-4GYCH-WByRo-jEQdT';
    const targetFolder = (folder || 'ExcelImports').replace(/^\/+/, '').replace(/\/+$/, '');

    const authHeader = `Basic ${getBasicAuth(user, pass)}`;
    const uploadUrl = `${host}/remote.php/dav/files/${encodeURIComponent(user)}/${targetFolder}/${filename}`;

    const buffer = Buffer.from(base64Content, 'base64');

    const putRes = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        Authorization: authHeader,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: buffer,
    });

    if (putRes.status === 201 || putRes.status === 204 || putRes.ok) {
      return res.json({
        success: true,
        message: `Successfully uploaded ${filename} to Nextcloud /${targetFolder}/`,
        uploadUrl,
      });
    } else {
      return res.status(putRes.status).json({
        success: false,
        error: `Nextcloud upload failed with HTTP ${putRes.status}: ${putRes.statusText}`,
      });
    }
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 6. Live Supabase Connection Test & PostgreSQL Schema Validation
app.post('/api/supabase/test-connection', async (req, res) => {
  try {
    const { url, anonKey, serviceKey, serviceRoleKey } = req.body;
    if (!url || typeof url !== 'string' || !url.trim()) {
      return res.status(400).json({
        success: false,
        isConnected: false,
        error: 'Supabase Project URL is required (e.g. https://your-project.supabase.co).',
        checks: { hostReachability: false, authValid: false, schemaDetected: false },
      });
    }

    const key = (serviceKey && serviceKey.trim()) || (serviceRoleKey && serviceRoleKey.trim()) || (anonKey && anonKey.trim());
    if (!key) {
      return res.status(400).json({
        success: false,
        isConnected: false,
        error: 'Supabase API Key (Anon Key or Service Key) is required to authenticate.',
        checks: { hostReachability: false, authValid: false, schemaDetected: false },
      });
    }

    const cleanUrl = url.trim().replace(/\/+$/, '');
    if (!cleanUrl.startsWith('http://') && !cleanUrl.startsWith('https://')) {
      return res.status(400).json({
        success: false,
        isConnected: false,
        error: 'Supabase URL must start with https:// or http://',
        checks: { hostReachability: false, authValid: false, schemaDetected: false },
      });
    }

    const startTime = Date.now();
    const restEndpoint = `${cleanUrl}/rest/v1/`;

    let response: Response;
    try {
      response = await fetch(restEndpoint, {
        method: 'GET',
        headers: {
          'apikey': key,
          'Authorization': `Bearer ${key}`,
          'Accept': 'application/openapi+json, application/json',
        },
      });
    } catch (netErr: any) {
      return res.status(502).json({
        success: false,
        isConnected: false,
        latencyMs: Date.now() - startTime,
        error: `Could not reach Supabase host at ${cleanUrl}: ${netErr.message}`,
        checks: { hostReachability: false, authValid: false, schemaDetected: false },
      });
    }

    const latencyMs = Date.now() - startTime;

    if (response.status === 401 || response.status === 403) {
      const errText = await response.text().catch(() => '');
      return res.status(401).json({
        success: false,
        isConnected: false,
        latencyMs,
        error: `Authentication failed (HTTP ${response.status}). The provided API key is invalid for project ${cleanUrl}. ${errText}`,
        checks: { hostReachability: true, authValid: false, schemaDetected: false },
      });
    }

    if (!response.ok && response.status !== 200 && response.status !== 204) {
      const errText = await response.text().catch(() => '');
      return res.status(response.status).json({
        success: false,
        isConnected: false,
        latencyMs,
        error: `Supabase returned HTTP ${response.status} (${response.statusText}): ${errText}`,
        checks: { hostReachability: true, authValid: false, schemaDetected: false },
      });
    }

    // Parse OpenAPI schema to discover available tables
    let discoveredTables: string[] = [];
    try {
      const schemaData = await response.json();
      if (schemaData.definitions) {
        discoveredTables = Object.keys(schemaData.definitions);
      } else if (schemaData.paths) {
        discoveredTables = Object.keys(schemaData.paths)
          .map((p) => p.replace(/^\//, ''))
          .filter((p) => !p.startsWith('rpc/'));
      }
    } catch {
      // JSON parse might fail if not openapi, but status was 200 OK
    }

    const expectedTables = ['sync_settings', 'workbooks', 'worksheet_mappings', 'column_mappings', 'import_logs', 'import_errors', 'sync_conflicts'];
    const matchingExpected = expectedTables.filter((t) => discoveredTables.includes(t));
    const schemaDetected = matchingExpected.length > 0;

    let hostName = cleanUrl;
    try {
      hostName = new URL(cleanUrl).hostname;
    } catch {}

    return res.json({
      success: true,
      isConnected: true,
      latencyMs,
      tablesCount: discoveredTables.length,
      tables: discoveredTables,
      matchingSyncTables: matchingExpected,
      checks: {
        hostReachability: true,
        authValid: true,
        schemaDetected,
      },
      message: schemaDetected
        ? `Successfully connected to Supabase PostgreSQL at ${hostName} (${latencyMs}ms). Verified ${matchingExpected.length} sync schema tables.`
        : `Successfully authenticated with Supabase PostgreSQL at ${hostName} (${latencyMs}ms). Note: Sync schema not yet created. Run the migration SQL in the SQL tab.`,
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, isConnected: false, error: error.message });
  }
});

// 7. Live Supabase Fetch Table Records
app.post('/api/supabase/fetch-table', async (req, res) => {
  try {
    const { url, anonKey, serviceKey, serviceRoleKey, tableName, limit } = req.body;
    if (!url || (!anonKey && !serviceKey && !serviceRoleKey) || !tableName) {
      return res.status(400).json({ success: false, error: 'url, API key, and tableName are required' });
    }

    const key = (serviceKey && serviceKey.trim()) || (serviceRoleKey && serviceRoleKey.trim()) || (anonKey && anonKey.trim());
    const cleanUrl = url.trim().replace(/\/+$/, '');
    const maxLimit = Math.min(Number(limit) || 50, 200);
    const queryUrl = `${cleanUrl}/rest/v1/${encodeURIComponent(tableName)}?select=*&limit=${maxLimit}`;

    const fetchRes = await fetch(queryUrl, {
      method: 'GET',
      headers: {
        'apikey': key,
        'Authorization': `Bearer ${key}`,
        'Range': `0-${maxLimit - 1}`,
      },
    });

    if (!fetchRes.ok) {
      const errText = await fetchRes.text().catch(() => '');
      return res.status(fetchRes.status).json({
        success: false,
        error: `Supabase query failed with HTTP ${fetchRes.status}: ${errText}`,
      });
    }

    const data = await fetchRes.json();
    return res.json({
      success: true,
      tableName,
      count: Array.isArray(data) ? data.length : 0,
      rows: Array.isArray(data) ? data : [],
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 8. Live Supabase Upsert Records
app.post('/api/supabase/upsert-records', async (req, res) => {
  try {
    const { url, anonKey, serviceKey, serviceRoleKey, tableName, records, onConflict } = req.body;
    if (!url || (!anonKey && !serviceKey && !serviceRoleKey) || !tableName || !Array.isArray(records) || records.length === 0) {
      return res.status(400).json({ success: false, error: 'url, API key, tableName, and records array are required' });
    }

    const key = (serviceKey && serviceKey.trim()) || (serviceRoleKey && serviceRoleKey.trim()) || (anonKey && anonKey.trim());
    const cleanUrl = url.trim().replace(/\/+$/, '');
    let upsertUrl = `${cleanUrl}/rest/v1/${encodeURIComponent(tableName)}`;
    if (onConflict) {
      upsertUrl += `?on_conflict=${encodeURIComponent(onConflict)}`;
    }

    const upsertRes = await fetch(upsertUrl, {
      method: 'POST',
      headers: {
        'apikey': key,
        'Authorization': `Bearer ${key}`,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=merge-duplicates, return=representation',
      },
      body: JSON.stringify(records),
    });

    if (!upsertRes.ok) {
      const errText = await upsertRes.text().catch(() => '');
      return res.status(upsertRes.status).json({
        success: false,
        error: `Supabase upsert failed with HTTP ${upsertRes.status}: ${errText}`,
      });
    }

    const data = await upsertRes.json().catch(() => []);
    return res.json({
      success: true,
      tableName,
      upsertedCount: records.length,
      records: data,
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// Gemini AI Schema and Column Mapping Endpoint
app.post('/api/gemini/suggest-mapping', async (req, res) => {
  try {
    const ai = getGemini();
    if (!ai) {
      return res.status(200).json({
        available: false,
        message: 'GEMINI_API_KEY not configured on server',
      });
    }

    const { prompt } = req.body;
    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        responseMimeType: 'application/json',
      },
    });

    const text = response.text;
    if (text) {
      const parsed = JSON.parse(text);
      return res.json({
        available: true,
        success: true,
        data: parsed,
      });
    }

    return res.json({ available: false, message: 'Empty response from model' });
  } catch (err: any) {
    console.warn('[Gemini] API error on server:', err.message);
    return res.json({
      available: false,
      error: err.message,
    });
  }
});

function formatBytes(bytes: number): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
}

// Start Server and mount Vite middleware
async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true, port: PORT, host: '0.0.0.0' },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
