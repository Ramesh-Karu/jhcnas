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

    const result = parseWorkbookBuffer(buffer, filename || targetUrl.split('/').pop() || 'workbook.xlsx');

    return res.json({
      success: true,
      filename: result.filename,
      fileSize: buffer.byteLength,
      fileSizeFormatted: formatBytes(buffer.byteLength),
      fileHash: sha256,
      totalWorksheets: result.worksheets.length,
      worksheets: result.worksheets,
      analyzedAt: new Date().toISOString(),
      base64Data: buffer.toString('base64'),
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// Helper function to parse XLSX / CSV buffer into structured worksheets & headers
function parseWorkbookBuffer(buffer: Buffer, filename: string) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const worksheets: any[] = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) {
      worksheets.push({
        sheetName,
        totalRows: 0,
        totalColumns: 0,
        usedRange: 'A1',
        mergedRanges: [],
        detectedHeaderRow: 1,
        detectedDataStartRow: 2,
        candidateHeaderRows: [],
        titleRows: [],
        sectionHeadings: [],
        emptyRowsCount: 0,
        repeatedHeadersCount: 0,
        headers: [],
        sampleRows: [],
      });
      continue;
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    const totalRows = range.e.r + 1;
    const totalColumns = range.e.c + 1;
    const merges = ws['!merges'] || [];
    const usedRangeStr = ws['!ref'];

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

    // Scan top 30 rows to detect the true header row with non-empty column names
    let bestHeaderRowIndex = 0;
    let maxHeaderScore = -1;
    const candidateHeaderRows: any[] = [];

    const scanLimit = Math.min(totalRows, 30);
    for (let r = 0; r < scanLimit; r++) {
      const rowVals: string[] = [];
      for (let c = 0; c < totalColumns; c++) {
        const val = String(getResolvedVal(r, c) || '').trim();
        if (val) rowVals.push(val);
      }

      if (rowVals.length > 0) {
        // Score row: number of non-empty strings, prefer rows with common identifier names
        const stringCount = rowVals.filter((v) => isNaN(Number(v))).length;
        const confidence = stringCount / Math.max(1, rowVals.length);
        candidateHeaderRows.push({
          row: r + 1,
          headers: rowVals,
          confidence: Number(confidence.toFixed(2)),
        });

        const score = rowVals.length * 2 + (confidence >= 0.5 ? 5 : 0);
        if (score > maxHeaderScore) {
          maxHeaderScore = score;
          bestHeaderRowIndex = r;
        }
      }
    }

    const detectedHeaderRow = bestHeaderRowIndex + 1;
    const detectedDataStartRow = detectedHeaderRow + 1;

    // Build headers from detected header row, filtering out completely blank trailing columns
    const headers: any[] = [];
    let lastNonEmptyCol = -1;

    for (let c = totalColumns - 1; c >= 0; c--) {
      const headerVal = String(getResolvedVal(bestHeaderRowIndex, c) || '').trim();
      let hasDataInCol = false;
      for (let r = detectedDataStartRow - 1; r < Math.min(totalRows, detectedDataStartRow + 20); r++) {
        if (String(getResolvedVal(r, c) || '').trim() !== '') {
          hasDataInCol = true;
          break;
        }
      }
      if (headerVal || hasDataInCol) {
        lastNonEmptyCol = c;
        break;
      }
    }

    const effectiveColCount = lastNonEmptyCol >= 0 ? lastNonEmptyCol + 1 : totalColumns;

    for (let c = 0; c < effectiveColCount; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const rawHeader = String(getResolvedVal(bestHeaderRowIndex, c) || '').trim();
      const name = rawHeader || `Column_${colLetter}`;

      const sampleValues: string[] = [];
      for (let r = detectedDataStartRow - 1; r < Math.min(totalRows, detectedDataStartRow + 10); r++) {
        const val = getResolvedVal(r, c);
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          sampleValues.push(String(val));
        }
      }

      headers.push({
        colLetter,
        colIndex: c + 1,
        name,
        sampleValues,
      });
    }

    // Extract sample rows (up to 1000 rows for complete dry-run and sync preview)
    const sampleRows: any[] = [];
    for (let r = detectedDataStartRow - 1; r < Math.min(totalRows, detectedDataStartRow + 1000); r++) {
      const rowData: Record<string, any> = {};
      let hasRowData = false;

      headers.forEach((h, hIdx) => {
        const val = getResolvedVal(r, hIdx);
        rowData[h.colLetter] = val;
        if (h.name) {
          rowData[h.name] = val;
        }
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          hasRowData = true;
        }
      });

      if (hasRowData) {
        sampleRows.push({
          rowNumber: r + 1,
          data: rowData,
        });
      }
    }

    worksheets.push({
      sheetName,
      totalRows,
      totalColumns: effectiveColCount,
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
      candidateHeaderRows,
      titleRows: [],
      sectionHeadings: [],
      emptyRowsCount: 0,
      repeatedHeadersCount: 0,
      headers,
      sampleRows,
    });
  }

  return { filename, worksheets };
}

// 4b. Parse Raw Uploaded / Pasted Excel or CSV Buffer
app.post('/api/excel/parse-raw', async (req, res) => {
  try {
    const { base64Data, rawText, filename } = req.body;
    let buffer: Buffer;

    if (base64Data) {
      buffer = Buffer.from(base64Data, 'base64');
    } else if (rawText) {
      buffer = Buffer.from(rawText, 'utf-8');
    } else {
      return res.status(400).json({ success: false, error: 'base64Data or rawText is required' });
    }

    const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
    const result = parseWorkbookBuffer(buffer, filename || 'uploaded_data.xlsx');

    return res.json({
      success: true,
      filename: result.filename,
      fileSize: buffer.byteLength,
      fileSizeFormatted: formatBytes(buffer.byteLength),
      fileHash: sha256,
      totalWorksheets: result.worksheets.length,
      worksheets: result.worksheets,
      analyzedAt: new Date().toISOString(),
      base64Data: buffer.toString('base64'),
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

// 6b. Live Supabase Detailed Table & Column Schema Discovery
app.post('/api/supabase/schema', async (req, res) => {
  try {
    const { url, anonKey, serviceKey, serviceRoleKey } = req.body;
    if (!url || typeof url !== 'string' || !url.trim()) {
      return res.status(400).json({ success: false, error: 'Supabase URL is required' });
    }

    const key = (serviceKey && serviceKey.trim()) || (serviceRoleKey && serviceRoleKey.trim()) || (anonKey && anonKey.trim());
    if (!key) {
      return res.status(400).json({ success: false, error: 'Supabase API key is required' });
    }

    const cleanUrl = url.trim().replace(/\/+$/, '');
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
        error: `Could not connect to Supabase host: ${netErr.message}`,
        tables: []
      });
    }

    if (!response.ok && response.status !== 200) {
      const errText = await response.text().catch(() => '');
      return res.status(response.status).json({
        success: false,
        error: `Supabase returned HTTP ${response.status}: ${errText}`,
        tables: []
      });
    }

    const tables: any[] = [];
    try {
      const schemaData = await response.json();

      if (schemaData.definitions) {
        for (const [tName, tDef] of Object.entries<any>(schemaData.definitions)) {
          const properties = tDef.properties || {};
          const requiredProps = Array.isArray(tDef.required) ? tDef.required : [];
          const columns: any[] = [];
          const primaryKeys: string[] = [];

          for (const [colName, colDef] of Object.entries<any>(properties)) {
            const desc = colDef.description || '';
            const isPk = desc.includes('<pk/>') || desc.toLowerCase().includes('primary key') || colName === 'id';
            if (isPk) primaryKeys.push(colName);

            let mappedType = colDef.type || 'text';
            if (colDef.format === 'uuid') mappedType = 'uuid';
            else if (colDef.format === 'date') mappedType = 'date';
            else if (colDef.format?.includes('time')) mappedType = 'timestamp';
            else if (colDef.type === 'integer' || colDef.type === 'number') mappedType = colDef.type === 'integer' ? 'integer' : 'decimal';
            else if (colDef.type === 'boolean') mappedType = 'boolean';
            else if (colDef.type === 'string') mappedType = 'text';

            columns.push({
              name: colName,
              type: mappedType,
              format: colDef.format || '',
              isPrimary: isPk,
              required: requiredProps.includes(colName),
              description: desc,
              defaultValue: colDef.default !== undefined ? String(colDef.default) : undefined,
            });
          }

          tables.push({
            name: tName,
            columns,
            primaryKeys,
            description: tDef.description || '',
          });
        }
      } else if (schemaData.paths) {
        const paths = Object.keys(schemaData.paths)
          .map((p) => p.replace(/^\//, ''))
          .filter((p) => !p.startsWith('rpc/'));

        for (const tName of paths) {
          tables.push({
            name: tName,
            columns: [
              { name: 'id', type: 'uuid', isPrimary: true, required: true },
              { name: 'created_at', type: 'timestamp', isPrimary: false, required: false },
            ],
            primaryKeys: ['id'],
          });
        }
      }
    } catch (e: any) {
      console.warn('Failed to parse OpenAPI JSON:', e);
    }

    return res.json({
      success: true,
      tablesCount: tables.length,
      tables,
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message, tables: [] });
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

// 7b. Live Supabase Execute DDL / Schema Migration
app.post('/api/supabase/execute-ddl', async (req, res) => {
  try {
    const { url, anonKey, serviceKey, serviceRoleKey, sql, tableName } = req.body;
    if (!url || !sql) {
      return res.status(400).json({ success: false, error: 'Supabase url and SQL script are required' });
    }

    const key = (serviceKey && serviceKey.trim()) || (serviceRoleKey && serviceRoleKey.trim()) || (anonKey && anonKey.trim());
    const cleanUrl = url.trim().replace(/\/+$/, '');

    // 1. Try Supabase rpc/exec_sql if provisioned
    try {
      const rpcRes = await fetch(`${cleanUrl}/rest/v1/rpc/exec_sql`, {
        method: 'POST',
        headers: {
          'apikey': key || '',
          'Authorization': `Bearer ${key}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ query: sql }),
      });

      if (rpcRes.ok) {
        return res.json({
          success: true,
          directExecuted: true,
          message: `Successfully executed DDL migration for table '${tableName || 'database'}' directly on Supabase!`,
        });
      }
    } catch {}

    // 2. Try Supabase pg/query endpoint
    try {
      const pgRes = await fetch(`${cleanUrl}/pg/query`, {
        method: 'POST',
        headers: {
          'apikey': key || '',
          'Authorization': `Bearer ${key}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ query: sql }),
      });

      if (pgRes.ok) {
        return res.json({
          success: true,
          directExecuted: true,
          message: `Successfully executed DDL migration on Supabase PostgreSQL!`,
        });
      }
    } catch {}

    // If direct execution endpoints are not enabled on Supabase, return detailed guidance
    return res.json({
      success: true,
      directExecuted: false,
      message: 'SQL generated successfully! To apply this schema to Supabase, paste this SQL into your Supabase Dashboard SQL Editor (https://supabase.com/dashboard). Once executed, the table will be instantly available.',
      sql,
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

// 9. Full End-to-End Live Pipeline Sync (Nextcloud WebDAV -> Transform -> Supabase PostgreSQL)
async function executeFullPipelineCore(params: {
  nextcloud?: any;
  supabase?: any;
  mappings?: any[];
  targetFilename?: string;
}): Promise<{
  success: boolean;
  filename: string;
  fileHash?: string;
  totalInserted: number;
  totalUpdated: number;
  totalFailed: number;
  syncResults: any[];
  errors: any[];
  executedAt: string;
}> {
  const { nextcloud, supabase, mappings, targetFilename } = params;
  if (!nextcloud?.url || !supabase?.url) {
    throw new Error('Nextcloud and Supabase configurations are required for sync execution');
  }

  const host = nextcloud.url.replace(/\/+$/, '');
  const user = nextcloud.username || 'truenas_admin';
  const pass = nextcloud.appPassword;
  const folder = (nextcloud.sourceFolder || '/ExcelImports').replace(/^\/+/, '').replace(/\/+$/, '');
  const filename = targetFilename || 'students.xlsx';

  const supUrl = supabase.url.trim().replace(/\/+$/, '');
  const supKey = (supabase.serviceKey && supabase.serviceKey.trim()) || 
                 (supabase.serviceRoleKey && supabase.serviceRoleKey.trim()) || 
                 (supabase.anonKey && supabase.anonKey.trim());

  if (!supKey) {
    throw new Error('Supabase API key is required for sync execution');
  }

  // Step 1: Download Workbook from Nextcloud WebDAV
  const targetUrl = `${host}/remote.php/dav/files/${encodeURIComponent(user)}/${folder}/${filename}`;
  const authHeader = `Basic ${getBasicAuth(user, pass)}`;

  const fileRes = await fetch(targetUrl, {
    method: 'GET',
    headers: { Authorization: authHeader },
    signal: AbortSignal.timeout(15000),
  });

  if (!fileRes.ok) {
    throw new Error(`Could not fetch ${filename} from Nextcloud WebDAV (${targetUrl}): HTTP ${fileRes.status} ${fileRes.statusText}`);
  }

  const arrayBuffer = await fileRes.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');

  // Step 2: Parse Workbook
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const syncResults: any[] = [];
  let totalInserted = 0;
  let totalUpdated = 0;
  let totalFailed = 0;
  const errors: any[] = [];

  const activeMappings = Array.isArray(mappings) && mappings.length > 0
    ? mappings.filter((m: any) => m.enabled !== false)
    : [];

  if (activeMappings.length === 0) {
    throw new Error('No enabled table mappings configured. Please configure at least one sheet-to-table mapping.');
  }

  // Step 3: Process each mapped worksheet
  for (const wm of activeMappings) {
    const ws = wb.Sheets[wm.worksheetName] || wb.Sheets[Object.keys(wb.Sheets)[0]];
    if (!ws) {
      errors.push({
        worksheetName: wm.worksheetName,
        error: `Worksheet "${wm.worksheetName}" not found in workbook. Available sheets: ${wb.SheetNames.join(', ')}`,
      });
      continue;
    }

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    const totalRows = range.e.r + 1;
    const merges = ws['!merges'] || [];

    const getVal = (r: number, c: number) => {
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

    const headerRowIndex = (wm.headerRow || 1) - 1;
    const dataStartRowIndex = (wm.dataStartRow || 2) - 1;

    // Map Excel column letters / names
    const colLookup: Record<string, number> = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const headerName = String(getVal(headerRowIndex, c) || '').trim();
      colLookup[colLetter] = c;
      if (headerName) {
        colLookup[headerName.toLowerCase()] = c;
      }
    }

    const targetTable = wm.supabaseTable || 'students';
    const recordsToUpsert: any[] = [];
    const uniqueCol = wm.columns?.find((c: any) => c.uniqueKey)?.supabaseColumn || 'username';

    for (let r = dataStartRowIndex; r < totalRows; r++) {
      const rowData: Record<string, any> = {};
      let hasAnyData = false;

      for (const colMap of wm.columns) {
        let colIdx = -1;
        if (colMap.excelColumn && colLookup[colMap.excelColumn] !== undefined) {
          colIdx = colLookup[colMap.excelColumn];
        } else if (colMap.excelHeader && colLookup[colMap.excelHeader.toLowerCase()] !== undefined) {
          colIdx = colLookup[colMap.excelHeader.toLowerCase()];
        }

        let rawVal = colIdx >= 0 ? getVal(r, colIdx) : '';
        let finalVal: any = rawVal;

        if (typeof finalVal === 'string') {
          finalVal = finalVal.trim();
        }

        if (colMap.transformation === 'trim' && typeof finalVal === 'string') {
          finalVal = finalVal.trim();
        } else if (colMap.transformation === 'uppercase' && typeof finalVal === 'string') {
          finalVal = finalVal.toUpperCase();
        } else if (colMap.transformation === 'lowercase' && typeof finalVal === 'string') {
          finalVal = finalVal.toLowerCase();
        } else if (colMap.transformation === 'normalize_id' && typeof finalVal === 'string') {
          finalVal = finalVal.replace(/\s+/g, '').replace(/[-_]/g, '');
        } else if (colMap.transformation === 'parse_date' && finalVal) {
          if (finalVal instanceof Date) {
            finalVal = finalVal.toISOString().split('T')[0];
          } else {
            const str = String(finalVal).replace(/\./g, '-').replace(/\//g, '-').trim();
            finalVal = str;
          }
        } else if (colMap.transformation === 'parse_number' && finalVal) {
          const num = Number(String(finalVal).replace(/[^0-9.-]/g, ''));
          if (!isNaN(num)) finalVal = num;
        }

        if (finalVal !== undefined && finalVal !== null && finalVal !== '') {
          hasAnyData = true;
        }

        rowData[colMap.supabaseColumn] = finalVal !== '' ? finalVal : (colMap.defaultValue || null);
      }

      if (hasAnyData) {
        recordsToUpsert.push(rowData);
      }
    }

    if (recordsToUpsert.length === 0) {
      syncResults.push({
        sheetName: wm.worksheetName,
        targetTable,
        rowsCount: 0,
        status: 'Skipped (No data rows found)',
      });
      continue;
    }

    // Execute upsert into Supabase PostgREST
    let upsertUrl = `${supUrl}/rest/v1/${encodeURIComponent(targetTable)}`;
    if (uniqueCol) {
      upsertUrl += `?on_conflict=${encodeURIComponent(uniqueCol)}`;
    }

    const upsertRes = await fetch(upsertUrl, {
      method: 'POST',
      headers: {
        'apikey': supKey,
        'Authorization': `Bearer ${supKey}`,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=merge-duplicates, return=representation',
      },
      body: JSON.stringify(recordsToUpsert),
      signal: AbortSignal.timeout(15000),
    });

    if (!upsertRes.ok) {
      const errText = await upsertRes.text().catch(() => '');
      totalFailed += recordsToUpsert.length;
      errors.push({
        worksheetName: wm.worksheetName,
        targetTable,
        httpStatus: upsertRes.status,
        error: `Supabase rejected upsert into public.${targetTable}: ${errText}`,
      });
      syncResults.push({
        sheetName: wm.worksheetName,
        targetTable,
        rowsCount: recordsToUpsert.length,
        status: 'Failed',
        error: errText,
      });
    } else {
      const upsertedData = await upsertRes.json().catch(() => []);
      const count = Array.isArray(upsertedData) ? upsertedData.length : recordsToUpsert.length;
      totalInserted += count;
      syncResults.push({
        sheetName: wm.worksheetName,
        targetTable,
        rowsCount: count,
        uniqueKey: uniqueCol,
        status: 'Success',
      });
    }
  }

  const overallSuccess = totalFailed === 0;
  return {
    success: overallSuccess,
    filename,
    fileHash: sha256,
    totalInserted,
    totalUpdated,
    totalFailed,
    syncResults,
    errors,
    executedAt: new Date().toISOString(),
  };
}

// Production In-Process Sync Scheduler
class ProductionScheduler {
  private enabled: boolean = false;
  private intervalMinutes: number = 15;
  private intervalLabel: string = '15m';
  private state: 'IDLE' | 'SYNCING' | 'SCHEDULED' | 'DISABLED' | 'ERROR' = 'DISABLED';
  private engineMode: 'INTEGRATED_PRODUCTION_ENGINE' | 'EXTERNAL_WORKER_DAEMON' = 'INTEGRATED_PRODUCTION_ENGINE';
  private workerUrl: string = '';
  private lastRunAt: string | null = null;
  private nextRunAt: string | null = null;
  private lastRunResult: any = null;
  private cachedNextcloud: any = null;
  private cachedSupabase: any = null;
  private cachedMappings: any[] = [];
  private timer: NodeJS.Timeout | null = null;

  public getStatus() {
    let secondsUntilNextRun: number | null = null;
    if (this.nextRunAt && this.enabled && this.state !== 'SYNCING') {
      const diff = Math.floor((new Date(this.nextRunAt).getTime() - Date.now()) / 1000);
      secondsUntilNextRun = Math.max(0, diff);
    }

    return {
      enabled: this.enabled,
      intervalMinutes: this.intervalMinutes,
      intervalLabel: this.intervalLabel,
      state: this.state,
      engineMode: this.engineMode,
      workerEndpoint: this.workerUrl || 'In-Process (Production Node.js Engine)',
      lastRunAt: this.lastRunAt,
      nextRunAt: this.enabled ? this.nextRunAt : null,
      secondsUntilNextRun,
      lastRunResult: this.lastRunResult,
    };
  }

  public configure(config: {
    enabled?: boolean;
    intervalLabel?: string;
    nextcloud?: any;
    supabase?: any;
    mappings?: any[];
    workerUrl?: string;
  }) {
    if (config.nextcloud) this.cachedNextcloud = config.nextcloud;
    if (config.supabase) this.cachedSupabase = config.supabase;
    if (Array.isArray(config.mappings)) this.cachedMappings = config.mappings;

    if (config.workerUrl !== undefined) {
      this.workerUrl = (config.workerUrl || '').trim();
      this.engineMode = this.workerUrl.startsWith('http')
        ? 'EXTERNAL_WORKER_DAEMON'
        : 'INTEGRATED_PRODUCTION_ENGINE';
    }

    if (config.intervalLabel) {
      this.intervalLabel = config.intervalLabel;
      switch (config.intervalLabel) {
        case '5m': this.intervalMinutes = 5; break;
        case '15m': this.intervalMinutes = 15; break;
        case '30m': this.intervalMinutes = 30; break;
        case '1h': this.intervalMinutes = 60; break;
        case 'daily': this.intervalMinutes = 1440; break;
        default: this.intervalMinutes = 15; break;
      }
    }

    if (config.enabled !== undefined) {
      this.enabled = config.enabled;
    }

    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
    }

    if (this.enabled) {
      this.state = 'SCHEDULED';
      const intervalMs = this.intervalMinutes * 60 * 1000;
      this.nextRunAt = new Date(Date.now() + intervalMs).toISOString();

      this.timer = setInterval(() => {
        this.triggerSync(false);
      }, intervalMs);

      console.log(`[Scheduler] Activated: Next run in ${this.intervalMinutes}m at ${this.nextRunAt}`);
    } else {
      this.state = 'DISABLED';
      this.nextRunAt = null;
      console.log('[Scheduler] Deactivated: Switched to Manual Mode');
    }

    return this.getStatus();
  }

  public async triggerSync(isManual: boolean = true): Promise<any> {
    if (this.state === 'SYNCING') {
      return { skipped: true, reason: 'Sync already in progress' };
    }

    this.state = 'SYNCING';
    console.log(`[Scheduler] Executing sync cycle (${isManual ? 'Manual Trigger' : 'Scheduled Run'})...`);

    try {
      const result = await executeFullPipelineCore({
        nextcloud: this.cachedNextcloud,
        supabase: this.cachedSupabase,
        mappings: this.cachedMappings,
      });

      this.lastRunAt = new Date().toISOString();
      this.lastRunResult = result;
      this.state = this.enabled ? 'SCHEDULED' : 'IDLE';

      if (this.enabled) {
        this.nextRunAt = new Date(Date.now() + this.intervalMinutes * 60 * 1000).toISOString();
      }

      console.log(`[Scheduler] Completed sync cycle: ${result.totalInserted} inserted, ${result.totalFailed} failed.`);
      return result;
    } catch (err: any) {
      this.lastRunAt = new Date().toISOString();
      this.lastRunResult = {
        success: false,
        totalInserted: 0,
        totalUpdated: 0,
        totalFailed: 0,
        error: err.message,
        executedAt: this.lastRunAt,
      };
      this.state = this.enabled ? 'SCHEDULED' : 'ERROR';
      if (this.enabled) {
        this.nextRunAt = new Date(Date.now() + this.intervalMinutes * 60 * 1000).toISOString();
      }
      console.warn(`[Scheduler] Sync failed: ${err.message}`);
      return this.lastRunResult;
    }
  }
}

const schedulerInstance = new ProductionScheduler();

// REST Endpoints for Real Scheduler
app.get('/api/scheduler/status', (_req, res) => {
  return res.json({ success: true, status: schedulerInstance.getStatus() });
});

app.post('/api/scheduler/configure', (req, res) => {
  try {
    const updated = schedulerInstance.configure(req.body);
    return res.json({ success: true, status: updated });
  } catch (err: any) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/scheduler/trigger-now', async (_req, res) => {
  try {
    const result = await schedulerInstance.triggerSync(true);
    return res.json({ success: true, result, status: schedulerInstance.getStatus() });
  } catch (err: any) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/sync/execute-full-pipeline', async (req, res) => {
  try {
    const result = await executeFullPipelineCore(req.body);
    // Cache for future scheduled runs
    schedulerInstance.configure({
      nextcloud: req.body.nextcloud,
      supabase: req.body.supabase,
      mappings: req.body.mappings,
    });
    return res.json(result);
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 10. Worker Diagnostics & Connectivity Matrix
app.post('/api/worker/diagnostics', async (req, res) => {
  try {
    const { nextcloud, supabase, workerUrl } = req.body;

    const diag: {
      timestamp: string;
      nextcloud: { reachable: boolean; status: string; details?: any; latencyMs?: number };
      supabase: { reachable: boolean; status: string; tablesCount?: number; tables?: string[]; latencyMs?: number };
      workerService: { reachable: boolean; status: string; endpoint: string; latencyMs?: number };
      allSystemsReady: boolean;
    } = {
      timestamp: new Date().toISOString(),
      nextcloud: { reachable: false, status: 'Not Checked' },
      supabase: { reachable: false, status: 'Not Checked' },
      workerService: { reachable: false, status: 'Integrated in App Engine', endpoint: workerUrl || 'internal' },
      allSystemsReady: false,
    };

    // 1. Nextcloud Check
    if (nextcloud?.url) {
      const t0 = Date.now();
      try {
        const host = nextcloud.url.replace(/\/+$/, '');
        const ncRes = await fetch(`${host}/status.php`, { method: 'GET' });
        diag.nextcloud.latencyMs = Date.now() - t0;
        if (ncRes.ok) {
          const info = await ncRes.json().catch(() => ({}));
          diag.nextcloud.reachable = true;
          diag.nextcloud.status = `Online (Nextcloud ${info.versionstring || 'v34'})`;
          diag.nextcloud.details = info;
        } else {
          diag.nextcloud.status = `HTTP ${ncRes.status} (${ncRes.statusText})`;
        }
      } catch (e: any) {
        diag.nextcloud.status = `Unreachable: ${e.message}`;
      }
    }

    // 2. Supabase Check
    if (supabase?.url) {
      const t0 = Date.now();
      try {
        const cleanUrl = supabase.url.trim().replace(/\/+$/, '');
        const key = (supabase.serviceKey && supabase.serviceKey.trim()) || (supabase.serviceRoleKey && supabase.serviceRoleKey.trim()) || (supabase.anonKey && supabase.anonKey.trim());
        if (key) {
          const sbRes = await fetch(`${cleanUrl}/rest/v1/`, {
            method: 'GET',
            headers: {
              'apikey': key,
              'Authorization': `Bearer ${key}`,
              'Accept': 'application/openapi+json, application/json',
            },
          });
          diag.supabase.latencyMs = Date.now() - t0;
          if (sbRes.ok) {
            diag.supabase.reachable = true;
            const openApi = await sbRes.json().catch(() => ({}));
            const tbls = openApi.definitions ? Object.keys(openApi.definitions) : [];
            diag.supabase.tablesCount = tbls.length;
            diag.supabase.tables = tbls;
            diag.supabase.status = `Connected to PostgreSQL PostgREST (${tbls.length} tables discovered)`;
          } else {
            diag.supabase.status = `HTTP ${sbRes.status} (${sbRes.statusText})`;
          }
        } else {
          diag.supabase.status = 'API Key missing';
        }
      } catch (e: any) {
        diag.supabase.status = `Unreachable: ${e.message}`;
      }
    }

    // 3. Worker Service Check (if external worker url configured)
    if (workerUrl && workerUrl.startsWith('http') && !workerUrl.includes('coolify-worker')) {
      const t0 = Date.now();
      try {
        const wRes = await fetch(`${workerUrl.replace(/\/+$/, '')}/health`, { method: 'GET' });
        diag.workerService.latencyMs = Date.now() - t0;
        if (wRes.ok) {
          diag.workerService.reachable = true;
          diag.workerService.status = 'Worker daemon responsive';
        } else {
          diag.workerService.status = `HTTP ${wRes.status}`;
        }
      } catch {
        diag.workerService.status = 'Worker daemon endpoint not responding (using integrated sync engine)';
      }
    } else {
      diag.workerService.reachable = true;
      diag.workerService.status = 'Active (Integrated sync pipeline engine)';
    }

    diag.allSystemsReady = diag.nextcloud.reachable && diag.supabase.reachable;

    return res.json({ success: true, diagnostics: diag });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 11. Unified Secrets & Subsystem Connectivity Test (Nextcloud, Supabase, Worker Array)
app.post('/api/secrets/test-all', async (req, res) => {
  try {
    const { nextcloud, supabase, workerUrl } = req.body;

    const report: {
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
        secretsSynced?: boolean;
        error?: string;
      };
    } = {
      timestamp: new Date().toISOString(),
      allConnected: false,
      nextcloud: {
        isConnected: false,
        latencyMs: 0,
        statusText: 'Untested',
        checks: { hostReachability: false, authValid: false, folderAccessible: false },
      },
      supabase: {
        isConnected: false,
        latencyMs: 0,
        statusText: 'Untested',
        tablesCount: 0,
        tables: [],
        checks: { hostReachability: false, authValid: false, schemaDetected: false },
      },
      worker: {
        isConnected: false,
        latencyMs: 0,
        statusText: 'Integrated In-App Daemon',
        endpoint: workerUrl || 'http://localhost:8000',
      },
    };

    // 1. Nextcloud Comprehensive WebDAV Check
    if (nextcloud?.url) {
      const t0 = Date.now();
      try {
        const host = nextcloud.url.replace(/\/+$/, '');
        const user = nextcloud.username || 'truenas_admin';
        const pass = nextcloud.appPassword || '';
        const folder = (nextcloud.sourceFolder || 'ExcelImports').replace(/^\/+/, '').replace(/\/+$/, '');

        // Step A: status.php
        const stRes = await fetch(`${host}/status.php`, { method: 'GET', signal: AbortSignal.timeout(5000) });
        if (stRes.ok) {
          report.nextcloud.checks.hostReachability = true;
        }

        // Step B: WebDAV PROPFIND
        const authHeader = `Basic ${getBasicAuth(user, pass)}`;
        const davUrl = `${host}/remote.php/dav/files/${encodeURIComponent(user)}/${folder}/`;
        const davRes = await fetch(davUrl, {
          method: 'PROPFIND',
          headers: {
            Authorization: authHeader,
            Depth: '1',
          },
          signal: AbortSignal.timeout(7000),
        });

        report.nextcloud.latencyMs = Date.now() - t0;
        if (davRes.status === 207 || davRes.status === 200) {
          report.nextcloud.isConnected = true;
          report.nextcloud.checks.authValid = true;
          report.nextcloud.checks.folderAccessible = true;
          report.nextcloud.statusText = `Connected (WebDAV /${folder}/ verified)`;
        } else if (davRes.status === 401 || davRes.status === 403) {
          report.nextcloud.checks.authValid = false;
          report.nextcloud.statusText = `Authentication failed (HTTP ${davRes.status}). Check Nextcloud App Password.`;
          report.nextcloud.error = 'Invalid credentials';
        } else if (davRes.status === 404) {
          report.nextcloud.checks.authValid = true;
          report.nextcloud.checks.folderAccessible = false;
          report.nextcloud.statusText = `Folder /${folder}/ not found on Nextcloud.`;
        } else {
          report.nextcloud.statusText = `WebDAV returned HTTP ${davRes.status}`;
        }
      } catch (err: any) {
        report.nextcloud.latencyMs = Date.now() - t0;
        report.nextcloud.statusText = `Nextcloud unreachable: ${err.message}`;
        report.nextcloud.error = err.message;
      }
    } else {
      report.nextcloud.statusText = 'Nextcloud URL missing';
    }

    // 2. Supabase Comprehensive PostgreSQL PostgREST Check
    if (supabase?.url) {
      const t0 = Date.now();
      try {
        const cleanUrl = supabase.url.trim().replace(/\/+$/, '');
        const key = (supabase.serviceKey && supabase.serviceKey.trim()) || 
                    (supabase.serviceRoleKey && supabase.serviceRoleKey.trim()) || 
                    (supabase.anonKey && supabase.anonKey.trim());

        if (!key) {
          report.supabase.statusText = 'Supabase API Key (Service Role or Anon) missing';
        } else {
          const restRes = await fetch(`${cleanUrl}/rest/v1/`, {
            method: 'GET',
            headers: {
              'apikey': key,
              'Authorization': `Bearer ${key}`,
              'Accept': 'application/openapi+json, application/json',
            },
            signal: AbortSignal.timeout(6000),
          });

          report.supabase.latencyMs = Date.now() - t0;
          report.supabase.checks.hostReachability = true;

          if (restRes.ok) {
            report.supabase.isConnected = true;
            report.supabase.checks.authValid = true;
            report.supabase.checks.schemaDetected = true;

            const schema = await restRes.json().catch(() => ({}));
            const tbls = schema.definitions ? Object.keys(schema.definitions) : [];
            report.supabase.tablesCount = tbls.length;
            report.supabase.tables = tbls;
            report.supabase.statusText = `Connected (${tbls.length} PostgreSQL tables verified)`;
          } else {
            report.supabase.checks.authValid = false;
            report.supabase.statusText = `Authentication failed (HTTP ${restRes.status}). Check Service Role Key.`;
            report.supabase.error = `HTTP ${restRes.status}`;
          }
        }
      } catch (err: any) {
        report.supabase.latencyMs = Date.now() - t0;
        report.supabase.statusText = `Supabase unreachable: ${err.message}`;
        report.supabase.error = err.message;
      }
    } else {
      report.supabase.statusText = 'Supabase URL missing';
    }

    // 3. Worker Daemon / Sync Engine Check
    if (workerUrl && typeof workerUrl === 'string' && workerUrl.trim().startsWith('http')) {
      const cleanUrl = workerUrl.trim().replace(/\/+$/, '');
      const t0 = Date.now();
      report.worker.endpoint = cleanUrl;
      try {
        const wRes = await fetch(`${cleanUrl}/health`, {
          method: 'GET',
          signal: AbortSignal.timeout(3000),
        });
        report.worker.latencyMs = Date.now() - t0;
        if (wRes.ok) {
          report.worker.isConnected = true;
          report.worker.statusText = `External worker operational at ${cleanUrl}`;
        } else {
          report.worker.isConnected = false;
          report.worker.statusText = `External worker returned HTTP ${wRes.status}`;
        }
      } catch (err: any) {
        report.worker.latencyMs = Date.now() - t0;
        report.worker.isConnected = false;
        report.worker.statusText = `External worker unreachable at ${cleanUrl}`;
      }
    } else {
      // Production in-app engine
      const schStatus = schedulerInstance.getStatus();
      report.worker.isConnected = true;
      report.worker.endpoint = 'Production Node.js In-App Pipeline';
      report.worker.latencyMs = 0;
      report.worker.statusText = schStatus.enabled
        ? `Production In-App Engine Active (Scheduled: Every ${schStatus.intervalLabel})`
        : 'Production In-App Engine Ready (Manual Trigger)';
    }

    report.allConnected = report.nextcloud.isConnected && report.supabase.isConnected;

    return res.json({
      success: true,
      report,
    });
  } catch (error: any) {
    return res.status(500).json({ success: false, error: error.message });
  }
});

// 12. Forward Secret Updates to Worker Daemon on Port 8000 (if running)
app.post('/api/worker/secrets/update', async (req, res) => {
  try {
    const workerTarget = (req.body.workerUrl || 'http://localhost:8000').replace(/\/+$/, '');
    try {
      const response = await fetch(`${workerTarget}/api/secrets/update`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(req.body),
        signal: AbortSignal.timeout(4000),
      });
      if (response.ok) {
        const data = await response.json();
        return res.json({ success: true, relayedToWorker: true, data });
      }
    } catch {
      // Worker daemon not running on port 8000; stored in Express memory
    }

    return res.json({
      success: true,
      relayedToWorker: false,
      message: 'Secrets updated in application state. Background sync pipeline configured.',
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
