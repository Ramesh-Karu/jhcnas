import * as XLSX from 'xlsx';
import { WorkbookAnalysis, SheetAnalysis, MergedRange, SheetHeader } from '../types';

export class ExcelAnalyzer {
  static async computeSHA256(data: ArrayBuffer): Promise<string> {
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  }

  static parseBuffer(buffer: ArrayBuffer | Uint8Array, filename: string): { wb: XLSX.WorkBook; analysis: WorkbookAnalysis } {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const fileSize = buffer.byteLength;
    const worksheets: SheetAnalysis[] = [];

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const sheetAnalysis = this.analyzeSheet(ws, sheetName);
      worksheets.push(sheetAnalysis);
    }

    const analysis: WorkbookAnalysis = {
      filename,
      fileSize,
      fileSizeFormatted: this.formatBytes(fileSize),
      totalWorksheets: worksheets.length,
      worksheets,
      analyzedAt: new Date().toISOString(),
      fileHash: '' // Will be populated with SHA-256
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

  static analyzeSheet(ws: XLSX.WorkSheet, sheetName: string): SheetAnalysis {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    const totalRows = range.e.r + 1;
    const totalColumns = range.e.c + 1;
    const usedRange = ws['!ref'] || 'A1';

    // 1. Process Merged Cells
    const rawMerges: XLSX.Range[] = ws['!merges'] || [];
    const mergedRanges: MergedRange[] = [];

    for (const m of rawMerges) {
      const startCol = XLSX.utils.encode_col(m.s.c);
      const endCol = XLSX.utils.encode_col(m.e.c);
      const startRow = m.s.r + 1;
      const endRow = m.e.r + 1;
      const rangeStr = `${startCol}${startRow}:${endCol}${endRow}`;

      // Get value from top-left cell
      const cellAddress = XLSX.utils.encode_cell(m.s);
      const cell = ws[cellAddress];
      const value = cell ? String(cell.v || cell.w || '') : '';

      const spanCols = m.e.c - m.s.c + 1;
      let type: 'title' | 'section_heading' | 'data_span' = 'data_span';
      if (spanCols >= Math.max(2, Math.floor(totalColumns / 2))) {
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

    // 2. Scan Rows to detect titles, section headings, empty rows, candidate headers
    const emptyRows: number[] = [];
    const titleRows: { row: number; text: string }[] = [];
    const sectionHeadings: { row: number; text: string; range: string }[] = [];
    const candidateHeaders: { row: number; headers: string[]; confidence: number }[] = [];

    const scanLimit = Math.min(totalRows, 100);
    for (let r = 0; r < scanLimit; r++) {
      const rowNumber = r + 1;
      const rowValues: string[] = [];

      for (let c = 0; c < totalColumns; c++) {
        const val = this.getResolvedCellValue(ws, r, c, rawMerges);
        if (val !== null && val !== undefined && String(val).trim() !== '') {
          rowValues.push(String(val).trim());
        }
      }

      if (rowValues.length === 0) {
        emptyRows.push(rowNumber);
        continue;
      }

      // Check if row matches a merged range
      const matchingMerge = mergedRanges.find(m => m.startRow <= rowNumber && rowNumber <= m.endRow);

      if (rowValues.length === 1 && matchingMerge) {
        if (matchingMerge.type === 'title') {
          titleRows.push({ row: rowNumber, text: matchingMerge.value });
        } else if (matchingMerge.type === 'section_heading') {
          sectionHeadings.push({ row: rowNumber, text: matchingMerge.value, range: matchingMerge.range });
        }
      } else if (rowValues.length >= 2) {
        // High string density indicates headers
        const stringCount = rowValues.filter(v => isNaN(Number(v))).length;
        const confidence = stringCount / rowValues.length;
        if (confidence >= 0.6) {
          candidateHeaders.push({
            row: rowNumber,
            headers: rowValues,
            confidence: Number(confidence.toFixed(2))
          });
        }
      }
    }

    // Determine Best Header Row
    let detectedHeaderRow = 1;
    if (candidateHeaders.length > 0) {
      detectedHeaderRow = candidateHeaders[0].row;
    } else if (totalRows > 0) {
      detectedHeaderRow = 1;
    }

    // Determine Data Start Row
    let detectedDataStartRow = detectedHeaderRow + 1;
    while (emptyRows.includes(detectedDataStartRow) && detectedDataStartRow < totalRows) {
      detectedDataStartRow++;
    }

    // Extract Headers at detectedHeaderRow with merged cell resolution
    const headers: SheetHeader[] = [];
    const headerR = detectedHeaderRow - 1;
    for (let c = 0; c < totalColumns; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const val = this.getResolvedCellValue(ws, headerR, c, rawMerges);
      const name = val !== null && val !== undefined && String(val).trim() !== '' 
        ? String(val).trim() 
        : `Column_${colLetter}`;

      // Sample a couple values for this column with merged cell resolution
      const sampleValues: string[] = [];
      for (let r = detectedDataStartRow - 1; r < Math.min(totalRows, detectedDataStartRow + 4); r++) {
        const scVal = this.getResolvedCellValue(ws, r, c, rawMerges);
        if (scVal !== null && scVal !== undefined) {
          sampleValues.push(String(scVal));
        }
      }

      headers.push({
        colLetter,
        colIndex: c + 1,
        name,
        sampleValues
      });
    }

    // Extract Sample Rows (first 6 rows of data) with resolved merged cells
    const sampleRows: { rowNumber: number; data: Record<string, any> }[] = [];
    const sampleLimit = Math.min(totalRows, detectedDataStartRow + 7);

    for (let r = detectedDataStartRow - 1; r < sampleLimit; r++) {
      const rowNumber = r + 1;
      if (emptyRows.includes(rowNumber)) continue;

      const rowData: Record<string, any> = {};
      let hasData = false;

      for (let c = 0; c < totalColumns; c++) {
        const colLetter = XLSX.utils.encode_col(c);
        const val = this.getResolvedCellValue(ws, r, c, rawMerges);
        rowData[colLetter] = val !== null && val !== undefined ? val : '';
        if (val) hasData = true;
      }

      if (hasData) {
        sampleRows.push({ rowNumber, data: rowData });
      }
    }

    return {
      sheetName,
      totalRows,
      totalColumns,
      usedRange,
      mergedRanges,
      detectedHeaderRow,
      detectedDataStartRow,
      candidateHeaderRows: candidateHeaders,
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
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    const totalCols = range.e.c + 1;
    const maxRow = dataEndRow ? Math.min(range.e.r + 1, dataEndRow) : range.e.r + 1;

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

    // Build header lookup
    const headerR = headerRow - 1;
    const headers: SheetHeader[] = [];
    const colIndexToName: Record<number, string> = {};

    for (let c = 0; c < totalCols; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const val = this.getResolvedCellValue(ws, headerR, c, rawMerges);
      const name = val !== null && val !== undefined && String(val).trim() !== '' ? String(val).trim() : `Col_${colLetter}`;
      colIndexToName[c] = name;
      headers.push({ colLetter, colIndex: c + 1, name, sampleValues: [] });
    }

    const records: Record<string, any>[] = [];
    let currentSectionHeading: string | null = null;

    for (let r = dataStartRow - 1; r < maxRow; r++) {
      const rowNumber = r + 1;

      // Check if this row is a section heading banner
      if (sectionHeadingsMap[rowNumber]) {
        currentSectionHeading = sectionHeadingsMap[rowNumber];
        continue;
      }

      const rowObj: Record<string, any> = { __rowNumber: rowNumber };
      let hasData = false;

      for (let c = 0; c < totalCols; c++) {
        const colLetter = XLSX.utils.encode_col(c);
        const val = this.getResolvedCellValue(ws, r, c, rawMerges);

        if (val !== null && val !== undefined && String(val).trim() !== '') {
          hasData = true;
        }

        rowObj[colLetter] = val;
        rowObj[`header_${colIndexToName[c]}`] = val;
      }

      if (!hasData) continue;

      // Downward propagation of merged section heading
      if (sectionHeadingTargetCol && currentSectionHeading) {
        rowObj['__sectionHeading'] = currentSectionHeading;
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
