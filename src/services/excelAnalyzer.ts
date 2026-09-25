import * as XLSX from 'xlsx';
import { WorkbookAnalysis, SheetAnalysis, MergedRange, SheetHeader, DataType } from '../types';
import { MatrixTransformer } from './matrixTransformer';
import { smartSanitizeIdentifier, containsTamil, translateTamilHeader } from './tamilTranslator';

export class ExcelAnalyzer {
  static async computeSHA256(data: ArrayBuffer): Promise<string> {
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  }

  static bufferToBase64(buffer: ArrayBuffer | Uint8Array): string {
    try {
      const bytes = new Uint8Array(buffer);
      let binary = '';
      const len = bytes.byteLength;
      for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      return btoa(binary);
    } catch (e) {
      console.warn('Could not encode buffer to base64:', e);
      return '';
    }
  }

  static parseBuffer(buffer: ArrayBuffer | Uint8Array, filename: string): { wb: XLSX.WorkBook; analysis: WorkbookAnalysis } {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true, cellNF: true });
    const fileSize = buffer.byteLength;
    const worksheets: SheetAnalysis[] = [];
    const base64Data = this.bufferToBase64(buffer);
    const rawPreviewData: { sheetName: string; rows: any[][] }[] = [];

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const sheetAnalysis = this.analyzeSheet(ws, sheetName);
      worksheets.push(sheetAnalysis);

      // Collect sample rows for archetype detection
      try {
        const previewRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
        rawPreviewData.push({ sheetName, rows: previewRows.slice(0, 25) });
      } catch {
        rawPreviewData.push({ sheetName, rows: [] });
      }
    }

    // Run Archetype Detection Engine
    const archetypeResult = MatrixTransformer.detectArchetype(wb.SheetNames, rawPreviewData);

    const analysis: WorkbookAnalysis = {
      filename,
      fileSize,
      fileSizeFormatted: this.formatBytes(fileSize),
      totalWorksheets: worksheets.length,
      worksheets,
      analyzedAt: new Date().toISOString(),
      fileHash: '', // Will be populated with SHA-256
      base64Data,
      rawWorkbookBase64: base64Data,
      detectedArchetype: archetypeResult.archetype,
      archetypeTitle: archetypeResult.title,
      archetypeBadge: archetypeResult.badge,
      archetypeSummary: archetypeResult.summary,
      archetypeFeatures: archetypeResult.features,
      archetypeRecommendations: archetypeResult.recommendations
    };

    return { wb, analysis };
  }

  static getResolvedCellValue(ws: XLSX.WorkSheet, r: number, c: number, rawMerges: XLSX.Range[]): any {
    const directCell = ws[XLSX.utils.encode_cell({ r, c })];
    if (directCell && directCell.v !== undefined && directCell.v !== null && String(directCell.v).trim() !== '') {
      return directCell.w !== undefined ? directCell.w : directCell.v;
    }
    // Check if covered by any merged cell range
    const matchingMerge = rawMerges.find(m => m.s.r <= r && r <= m.e.r && m.s.c <= c && c <= m.e.c);
    if (matchingMerge) {
      const originCell = ws[XLSX.utils.encode_cell(matchingMerge.s)];
      if (originCell && originCell.v !== undefined && originCell.v !== null) {
        return originCell.w !== undefined ? originCell.w : originCell.v;
      }
    }
    return null;
  }

  /**
   * Industrial Standard: Detects if a row is a merged year header, section divider,
   * or category banner (e.g. "Year 2023", "2020", "CLASS 10A") so that it is NEVER
   * used, sampled, or inserted into the database as a data row.
   */
  static isYearOrSectionDividerRow(
    ws: XLSX.WorkSheet,
    r: number,
    totalCols: number,
    rawMerges: XLSX.Range[]
  ): { isDivider: boolean; extractedHeading?: string } {
    // 1. Check if this row intersects a merged range spanning 2+ columns with a year or section definition
    const mergesInRow = rawMerges.filter(m => m.s.r <= r && r <= m.e.r && m.e.c > m.s.c);
    for (const m of mergesInRow) {
      const originCell = ws[XLSX.utils.encode_cell(m.s)];
      const rawText = String(originCell?.w ?? originCell?.v ?? '').trim();
      if (!rawText) continue;

      const spanCols = m.e.c - m.s.c + 1;
      const isYearPattern =
        /\b(19\d{2}|20\d{2})\b/.test(rawText) ||
        /year\s*[-:]?\s*\d{4}/i.test(rawText) ||
        /\d{4}\s*[-/]\s*\d{2,4}/.test(rawText) ||
        /^(year|academic\s*year|batch)/i.test(rawText);
      const isSectionPattern = /^(grade|class|section|term|semester)\s*[:;]?\s*\w+/i.test(rawText);

      if (spanCols >= 2 && (isYearPattern || isSectionPattern)) {
        return { isDivider: true, extractedHeading: rawText };
      }

      if (spanCols >= Math.max(2, Math.floor(totalCols / 2))) {
        return { isDivider: true, extractedHeading: rawText };
      }
    }

    // 2. Check individual cell values in this row
    const nonNullVals: string[] = [];
    for (let c = 0; c < totalCols; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
        nonNullVals.push(String(cell.w ?? cell.v).trim());
      }
    }

    if (nonNullVals.length === 0) {
      return { isDivider: true }; // blank divider row
    }

    // If only 1 or 2 cells populated across the row and matches year or section pattern
    if (nonNullVals.length <= 2) {
      const joined = nonNullVals.join(' ').trim();
      if (
        /^(year\s*[-:]?\s*)?(19\d{2}|20\d{2})([-/\s]+(19\d{2}|20\d{2}))?$/i.test(joined) ||
        /year[- ]?\d{4}/i.test(joined) ||
        /^(academic\s*year|batch|grade|class|section)\s*[:;]?\s*[0-9a-zA-Z\s_-]+/i.test(joined) ||
        /^(total|grand\s*total|subtotal)$/i.test(joined)
      ) {
        return { isDivider: true, extractedHeading: joined };
      }
    }

    return { isDivider: false };
  }

  // Automatic Data Type Inference Engine
  static inferColumnDataType(values: any[]): {
    dataType: DataType;
    nullCount: number;
    uniqueCount: number;
    isCandidateKey: boolean;
  } {
    let nullCount = 0;
    const nonNulls: any[] = [];
    const uniqueSet = new Set<string>();

    for (const v of values) {
      if (v === null || v === undefined || String(v).trim() === '') {
        nullCount++;
      } else {
        const str = String(v).trim();
        nonNulls.push(v);
        uniqueSet.add(str.toLowerCase());
      }
    }

    const totalCount = nonNulls.length;
    if (totalCount === 0) {
      return { dataType: 'text', nullCount, uniqueCount: 0, isCandidateKey: false };
    }

    let intCount = 0;
    let decimalCount = 0;
    let dateCount = 0;
    let boolCount = 0;
    let uuidCount = 0;

    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
    const isoDateRegex = /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/;
    const dmyDateRegex = /^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/;

    for (const v of nonNulls) {
      if (v instanceof Date && !isNaN(v.getTime())) {
        dateCount++;
        continue;
      }

      const str = String(v).trim();

      // Boolean
      const lower = str.toLowerCase();
      if (['true', 'false', 'yes', 'no', 'y', 'n', 't', 'f'].includes(lower)) {
        boolCount++;
        continue;
      }

      // UUID
      if (uuidRegex.test(str)) {
        uuidCount++;
        continue;
      }

      // Date string or Excel date serial
      if (isoDateRegex.test(str) || dmyDateRegex.test(str)) {
        const testD = new Date(str);
        if (!isNaN(testD.getTime())) {
          dateCount++;
          continue;
        }
      }

      // Number test
      const cleanNumStr = str.replace(/[$,€£₹% ]/g, '').replace(/,/g, '');
      const num = Number(cleanNumStr);
      if (!isNaN(num) && cleanNumStr !== '') {
        // Check if integer or decimal
        if (Number.isInteger(num)) {
          intCount++;
        } else {
          decimalCount++;
        }
        continue;
      }
    }

    let dataType: DataType = 'text';
    const threshold = totalCount * 0.8;

    if (boolCount >= threshold) {
      dataType = 'boolean';
    } else if (dateCount >= threshold) {
      dataType = 'date';
    } else if (intCount >= threshold) {
      dataType = 'integer';
    } else if ((intCount + decimalCount) >= threshold) {
      dataType = 'decimal';
    } else {
      dataType = 'text';
    }

    // High uniqueness indicates candidate primary/merge key
    const isCandidateKey = nullCount === 0 && uniqueSet.size === totalCount && totalCount > 0;

    return {
      dataType,
      nullCount,
      uniqueCount: uniqueSet.size,
      isCandidateKey
    };
  }

  /**
   * Deduplicates headers across a row so that multiple columns with identical
   * names (e.g. "Job", "Job" or "Phone", "Phone") receive unique, standard names
   * (e.g. "Job 1", "Job 2" or "Phone 1", "Phone 2") for seamless database schema mapping.
   */
  static deduplicateHeaders(
    rawList: { colIndex: number; colLetter: string; rawName: string }[]
  ): { colIndex: number; colLetter: string; uniqueName: string; originalName: string }[] {
    const nameCounts = new Map<string, number>();
    for (const item of rawList) {
      const norm = item.rawName.trim().toLowerCase();
      nameCounts.set(norm, (nameCounts.get(norm) || 0) + 1);
    }

    const usedUniqueNames = new Set<string>();
    const nameOccurrenceIndex = new Map<string, number>();
    const result: { colIndex: number; colLetter: string; uniqueName: string; originalName: string }[] = [];

    for (const item of rawList) {
      const rawTrimmed = item.rawName.trim() || `Column_${item.colLetter}`;
      const norm = rawTrimmed.toLowerCase();
      const totalCount = nameCounts.get(norm) || 1;

      let candidateName = rawTrimmed;

      if (totalCount > 1) {
        const currentOcc = (nameOccurrenceIndex.get(norm) || 0) + 1;
        nameOccurrenceIndex.set(norm, currentOcc);
        candidateName = `${rawTrimmed} ${currentOcc}`;
      }

      let finalName = candidateName;
      let suffix = 2;
      while (usedUniqueNames.has(finalName.toLowerCase())) {
        finalName = `${rawTrimmed} ${suffix}`;
        suffix++;
      }

      usedUniqueNames.add(finalName.toLowerCase());
      result.push({
        colIndex: item.colIndex,
        colLetter: item.colLetter,
        uniqueName: finalName,
        originalName: rawTrimmed,
      });
    }

    return result;
  }

  static analyzeSheet(
    ws: XLSX.WorkSheet,
    sheetName: string,
    headerRowOverride?: number,
    dataStartRowOverride?: number
  ): SheetAnalysis {
    // 1. Scan actual cell keys to find the true bounds of the sheet (not just !ref)
    let minR = 0, maxR = 0, minC = 0, maxC = 0;
    let foundCells = false;

    for (const k of Object.keys(ws)) {
      if (k.startsWith('!')) continue;
      foundCells = true;
      try {
        const cell = XLSX.utils.decode_cell(k);
        if (cell.r > maxR) maxR = cell.r;
        if (cell.c > maxC) maxC = cell.c;
      } catch {}
    }

    if (ws['!ref']) {
      try {
        const refRange = XLSX.utils.decode_range(ws['!ref']);
        maxR = Math.max(maxR, refRange.e.r);
        maxC = Math.max(maxC, refRange.e.c);
        foundCells = true;
      } catch {}
    }

    const totalRows = foundCells ? maxR + 1 : 1;
    const totalColumns = foundCells ? maxC + 1 : 1;
    const usedRange = ws['!ref'] || `A1:${XLSX.utils.encode_col(maxC)}${maxR + 1}`;

    // 2. Process Merged Cells
    const rawMerges: XLSX.Range[] = ws['!merges'] || [];
    const mergedRanges: MergedRange[] = [];

    for (const m of rawMerges) {
      const startCol = XLSX.utils.encode_col(m.s.c);
      const endCol = XLSX.utils.encode_col(m.e.c);
      const startRow = m.s.r + 1;
      const endRow = m.e.r + 1;
      const rangeStr = `${startCol}${startRow}:${endCol}${endRow}`;

      const cellAddress = XLSX.utils.encode_cell(m.s);
      const cell = ws[cellAddress];
      const value = cell ? String(cell.v || cell.w || '') : '';

      const spanCols = m.e.c - m.s.c + 1;
      let type: 'title' | 'section_heading' | 'data_span' = 'data_span';
      if (
        (spanCols >= 2 && (/\b(19\d{2}|20\d{2})\b/.test(value) || /year/i.test(value) || /^(grade|class|section|batch)/i.test(value))) ||
        spanCols >= Math.max(2, Math.floor(totalColumns / 2))
      ) {
        if (startRow <= 2) {
          type = 'title';
        } else {
          type = 'section_heading';
        }
      }

      mergedRanges.push({
        range: rangeStr,
        startCol,
        endCol,
        startRow,
        endRow,
        value,
        type
      });
    }

    // 3. Scan top rows (rows 1-15) to detect candidate headers
    const emptyRows: number[] = [];
    const titleRows: { row: number; text: string }[] = [];
    const sectionHeadings: { row: number; text: string; range: string }[] = [];
    const candidateHeaders: { row: number; headers: string[]; confidence: number; score: number }[] = [];

    const headerKeywords = [
      'name', 'id', 'user', 'code', 'date', 'dob', 'class', 'grade', 'status', 'email', 'phone', 'number', 'roll', 'index', 'gender', 'mark', 'score', 'address',
      'balance', 'ledger', 'articles', 'description', 'actual', 'surplus', 'deficiency', 'remarks', 'page', 'signature', 'responsible', 'person', 'contact', 'total', 'section', 'quantity', 'amount',
      // Tamil common keywords
      'பெயர்', 'இலக்கம்', 'எண்', 'மீதி', 'கையிருப்பு', 'பொருட்கள்', 'விளக்கம்', 'உபரி', 'பற்றாக்குறை', 'குறிப்பு', 'பொறுப்பாளர்', 'ஒப்பம்', 'தொடர்பு', 'மொத்தம்', 'பிரிவு', 'தொகை', 'திகதி', 'தேதி'
    ];

    const scanLimit = Math.min(totalRows, 20);
    for (let r = 0; r < scanLimit; r++) {
      const rowNumber = r + 1;
      const rowValues: string[] = [];
      let textCount = 0;
      let numCount = 0;
      let keywordHits = 0;
      const uniqueHeaders = new Set<string>();

      for (let c = 0; c < totalColumns; c++) {
        const val = this.getResolvedCellValue(ws, r, c, rawMerges);
        if (val !== null && val !== undefined && String(val).trim() !== '') {
          const str = String(val).trim();
          rowValues.push(str);
          uniqueHeaders.add(str.toLowerCase());

          const isNum = !isNaN(Number(str.replace(/[,%$ ]/g, '')));
          if (isNum) numCount++;
          else textCount++;

          const sanitized = smartSanitizeIdentifier(str);
          if (
            headerKeywords.some(kw => str.toLowerCase().includes(kw) || sanitized.includes(kw)) ||
            containsTamil(str)
          ) {
            keywordHits++;
          }
        }
      }

      if (rowValues.length === 0) {
        emptyRows.push(rowNumber);
        continue;
      }

      // Check if merged title banner
      const matchingMerge = mergedRanges.find(m => m.startRow <= rowNumber && rowNumber <= m.endRow);
      if (rowValues.length <= 2 && matchingMerge) {
        if (matchingMerge.type === 'title') {
          titleRows.push({ row: rowNumber, text: matchingMerge.value });
          continue;
        } else if (matchingMerge.type === 'section_heading') {
          sectionHeadings.push({ row: rowNumber, text: matchingMerge.value, range: matchingMerge.range });
          continue;
        }
      }

      // Header scoring formula:
      // High unique text headers + keyword presence - pure numbers - depth penalty
      const confidence = rowValues.length > 0 ? textCount / rowValues.length : 0;
      const score = (uniqueHeaders.size * 3) + (keywordHits * 4) + (textCount * 2) - (numCount * 2.5) - (r * 1.5);

      if (uniqueHeaders.size >= 1 && (confidence >= 0.3 || rowNumber <= 3)) {
        candidateHeaders.push({
          row: rowNumber,
          headers: rowValues,
          confidence: Number(confidence.toFixed(2)),
          score
        });
      }
    }

    // Determine Best Header Row (prefer highest score or manual override)
    let detectedHeaderRow = headerRowOverride && headerRowOverride >= 1 ? headerRowOverride : 1;
    if (!headerRowOverride && candidateHeaders.length > 0) {
      const best = [...candidateHeaders].sort((a, b) => b.score - a.score)[0];
      detectedHeaderRow = best.row;
    }

    // Determine Data Start Row (default to detectedHeaderRow + 1)
    let detectedDataStartRow = dataStartRowOverride && dataStartRowOverride > detectedHeaderRow
      ? dataStartRowOverride
      : detectedHeaderRow + 1;

    while (emptyRows.includes(detectedDataStartRow) && detectedDataStartRow < totalRows) {
      detectedDataStartRow++;
    }

    // 4. Extract Headers at detectedHeaderRow
    // Scan all columns up to totalColumns
    const headerR = detectedHeaderRow - 1;
    let lastActiveCol = -1;

    // First pass to find the true last active column with either a header or data
    for (let c = totalColumns - 1; c >= 0; c--) {
      const hVal = this.getResolvedCellValue(ws, headerR, c, rawMerges);
      let colHasData = false;
      for (let r = detectedDataStartRow - 1; r < Math.min(totalRows, detectedDataStartRow + 25); r++) {
        const dVal = this.getResolvedCellValue(ws, r, c, rawMerges);
        if (dVal !== null && dVal !== undefined && String(dVal).trim() !== '') {
          colHasData = true;
          break;
        }
      }
      if ((hVal !== null && String(hVal).trim() !== '') || colHasData) {
        lastActiveCol = c;
        break;
      }
    }

    const effectiveColCount = lastActiveCol >= 0 ? lastActiveCol + 1 : totalColumns;
    const rawHeaderList: { colIndex: number; colLetter: string; rawName: string }[] = [];

    // Extract raw header strings from header row
    for (let c = 0; c < effectiveColCount; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const val = this.getResolvedCellValue(ws, headerR, c, rawMerges);
      let rawName = val !== null && val !== undefined && String(val).trim() !== ''
        ? String(val).trim().replace(/[\r\n\t]+/g, ' ')
        : `Column_${colLetter}`;
      rawHeaderList.push({ colIndex: c, colLetter, rawName });
    }

    // Deduplicate headers: if there are duplicate column names (e.g. Job, Job), rename them to Job 1, Job 2
    const deduplicatedHeaders = this.deduplicateHeaders(rawHeaderList);
    const headers: SheetHeader[] = [];

    // 5. Build Headers and Auto-Detect Data Types
    for (let c = 0; c < effectiveColCount; c++) {
      const { colLetter, uniqueName, originalName } = deduplicatedHeaders[c];

      // Sample values from data rows for this column
      const sampleValues: string[] = [];
      const rawSampleList: any[] = [];
      const sampleLimit = Math.min(totalRows, detectedDataStartRow + 100);

      for (let r = detectedDataStartRow - 1; r < sampleLimit; r++) {
        // Exclude merged year and section divider rows from column type inference
        if (this.isYearOrSectionDividerRow(ws, r, effectiveColCount, rawMerges).isDivider) {
          continue;
        }
        const scVal = this.getResolvedCellValue(ws, r, c, rawMerges);
        if (scVal !== null && scVal !== undefined && String(scVal).trim() !== '') {
          rawSampleList.push(scVal);
          if (sampleValues.length < 8) {
            sampleValues.push(String(scVal));
          }
        }
      }

      // Run Automated Type Inference
      const { dataType, nullCount, uniqueCount, isCandidateKey } = this.inferColumnDataType(rawSampleList);

      // Check if header name hints at candidate key (id, username, code, roll, number, receipt, etc.)
      const nameNorm = smartSanitizeIdentifier(uniqueName);
      const keyHints = ['id', 'username', 'user', 'code', 'index', 'rollnumber', 'studentid', 'indexnumber', 'number', 'serial', 'receipt', 'item_no', 'section_and_number', 'book_no'];
      const hasKeyHint = keyHints.some(k => nameNorm.includes(k) || uniqueName.toLowerCase().includes(k));

      headers.push({
        colLetter,
        colIndex: c + 1,
        name: uniqueName,
        originalName,
        sampleValues,
        inferredType: dataType,
        nullCount,
        uniqueCount,
        isCandidateKey: isCandidateKey || (c === 0 && hasKeyHint)
      });
    }

    // 6. Extract Sample Rows (extracting all data rows, up to 50,000 rows for complete 3000+ row processing)
    const sampleRows: { rowNumber: number; data: Record<string, any> }[] = [];
    const maxSampleRows = Math.min(totalRows, detectedDataStartRow + 50000);

    for (let r = detectedDataStartRow - 1; r < maxSampleRows; r++) {
      const rowNumber = r + 1;
      if (emptyRows.includes(rowNumber)) continue;

      // Industrial standard: exclude merged year header rows and section dividers from sample data
      if (this.isYearOrSectionDividerRow(ws, r, effectiveColCount, rawMerges).isDivider) {
        continue;
      }

      const rowData: Record<string, any> = {};
      let hasData = false;

      for (let c = 0; c < effectiveColCount; c++) {
        const colLetter = XLSX.utils.encode_col(c);
        const val = this.getResolvedCellValue(ws, r, c, rawMerges);
        rowData[colLetter] = val !== null && val !== undefined ? val : '';
        const h = headers[c];
        if (h) {
          rowData[h.name] = val !== null && val !== undefined ? val : '';
          if (h.originalName && rowData[h.originalName] === undefined) {
            rowData[h.originalName] = val !== null && val !== undefined ? val : '';
          }
        }
        if (val !== null && val !== undefined && String(val).trim() !== '') {
          hasData = true;
        }
      }

      if (hasData) {
        sampleRows.push({ rowNumber, data: rowData });
      }
    }

    return {
      sheetName,
      totalRows,
      totalColumns: effectiveColCount,
      usedRange,
      mergedRanges,
      detectedHeaderRow,
      detectedDataStartRow,
      candidateHeaderRows: candidateHeaders.map(ch => ({
        row: ch.row,
        headers: ch.headers,
        confidence: ch.confidence
      })),
      titleRows,
      sectionHeadings,
      emptyRowsCount: emptyRows.length,
      repeatedHeadersCount: 0,
      headers,
      sampleRows
    };
  }

  static extractRecordsWithMergedSections(
    ws: XLSX.WorkSheet,
    headerRow: number,
    dataStartRow: number,
    dataEndRow?: number,
    sectionHeadingTargetCol?: string
  ): { records: Record<string, any>[]; headers: SheetHeader[] } {
    let maxR = 0, maxC = 0;
    for (const k of Object.keys(ws)) {
      if (k.startsWith('!')) continue;
      try {
        const cell = XLSX.utils.decode_cell(k);
        if (cell.r > maxR) maxR = cell.r;
        if (cell.c > maxC) maxC = cell.c;
      } catch {}
    }
    if (ws['!ref']) {
      try {
        const r = XLSX.utils.decode_range(ws['!ref']);
        maxR = Math.max(maxR, r.e.r);
        maxC = Math.max(maxC, r.e.c);
      } catch {}
    }

    const totalCols = maxC + 1;
    const maxRow = dataEndRow ? Math.min(maxR + 1, dataEndRow) : maxR + 1;

    // Detect section headings from merges
    const rawMerges = ws['!merges'] || [];
    const sectionHeadingsMap: Record<number, string> = {};

    for (const m of rawMerges) {
      const spanCols = m.e.c - m.s.c + 1;
      const startRow = m.s.r + 1;
      if (spanCols >= Math.max(2, Math.floor(totalCols / 2)) && startRow > 2) {
        const cell = ws[XLSX.utils.encode_cell(m.s)];
        if (cell && cell.v) {
          sectionHeadingsMap[startRow] = String(cell.w || cell.v).trim();
        }
      }
    }

    // Build header lookup with deduplication
    const headerR = headerRow - 1;
    const rawHeaderList: { colIndex: number; colLetter: string; rawName: string }[] = [];

    for (let c = 0; c < totalCols; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const val = this.getResolvedCellValue(ws, headerR, c, rawMerges);
      const rawName = val !== null && val !== undefined && String(val).trim() !== ''
        ? String(val).trim().replace(/[\r\n\t]+/g, ' ')
        : `Column_${colLetter}`;
      rawHeaderList.push({ colIndex: c, colLetter, rawName });
    }

    const deduplicated = this.deduplicateHeaders(rawHeaderList);
    const headers: SheetHeader[] = [];
    const colIndexToName: Record<number, string> = {};
    const colIndexToOriginal: Record<number, string> = {};

    for (let c = 0; c < totalCols; c++) {
      const { colLetter, uniqueName, originalName } = deduplicated[c];
      colIndexToName[c] = uniqueName;
      colIndexToOriginal[c] = originalName;
      headers.push({ colLetter, colIndex: c + 1, name: uniqueName, originalName, sampleValues: [] });
    }

    const records: Record<string, any>[] = [];
    let currentSectionHeading: string | null = null;

    for (let r = dataStartRow - 1; r < maxRow; r++) {
      const rowNumber = r + 1;

      if (sectionHeadingsMap[rowNumber]) {
        currentSectionHeading = sectionHeadingsMap[rowNumber];
        continue;
      }

      // Check if this row is an embedded section/year divider (merged columns with year or category)
      const dividerCheck = this.isYearOrSectionDividerRow(ws, r, totalCols, rawMerges);
      if (dividerCheck.isDivider) {
        if (dividerCheck.extractedHeading) {
          currentSectionHeading = dividerCheck.extractedHeading;
        }
        continue; // CRITICAL: NEVER treat merged year or divider rows as database records!
      }

      const rowObj: Record<string, any> = { __rowNumber: rowNumber };
      let hasData = false;
      const nonNullVals: string[] = [];

      for (let c = 0; c < totalCols; c++) {
        const colLetter = XLSX.utils.encode_col(c);
        let val = this.getResolvedCellValue(ws, r, c, rawMerges);

        if (val !== null && val !== undefined && String(val).trim() !== '') {
          hasData = true;
          const str = String(val).trim();
          nonNullVals.push(str);

          // Clean currency strings: e.g. "Rs 40000" or "$ 1500" -> number
          if (/^Rs\.?\s*[\d,]+/i.test(str) || /^[$€£₹]\s*[\d,]+/i.test(str)) {
            const numStr = str.replace(/[^0-9.]/g, '');
            const parsedNum = parseFloat(numStr);
            if (!isNaN(parsedNum)) {
              val = parsedNum;
            }
          }
        }

        const uniqueCol = colIndexToName[c];
        const origCol = colIndexToOriginal[c];

        rowObj[colLetter] = val;
        rowObj[`header_${uniqueCol}`] = val;
        rowObj[uniqueCol] = val;
        if (origCol && origCol !== uniqueCol && rowObj[origCol] === undefined) {
          rowObj[origCol] = val;
        }
      }

      if (!hasData) continue;

      // Secondary check: single cell value with year or category pattern
      if (nonNullVals.length === 1) {
        const singleVal = nonNullVals[0];
        if (/^(year\s*[-:]?\s*)?(19\d{2}|20\d{2})$/i.test(singleVal) || /year[- ]?\d{4}/i.test(singleVal) || /^(grade|class|section)\s*[:;]?\s*\w+/i.test(singleVal)) {
          currentSectionHeading = singleVal;
          continue; // Skip the divider row so it does not corrupt SQL table rows
        }
      }

      if (currentSectionHeading) {
        rowObj['__sectionHeading'] = currentSectionHeading;
        if (sectionHeadingTargetCol) {
          rowObj[sectionHeadingTargetCol] = currentSectionHeading;
        }
      }

      records.push(rowObj);
    }

    return { records, headers };
  }

  static formatBytes(bytes: number): string {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }
}
