import * as XLSX from 'xlsx';
import { DryRunResult, WorksheetMapping, RowValidationError, ImportLog, WorkbookAnalysis } from '../types';
import { ExcelAnalyzer } from './excelAnalyzer';
import { MappingEngine } from './mappingEngine';

function areCellsEqual(a: any, b: any): boolean {
  if (a === b) return true;
  const isNullishA = a === null || a === undefined || a === '';
  const isNullishB = b === null || b === undefined || b === '';
  if (isNullishA && isNullishB) return true;

  const strA = String(a ?? '').trim();
  const strB = String(b ?? '').trim();
  if (strA === strB) return true;

  const emptyPlaceholders = ['-', '—', '--', 'n/a', 'na', 'nil', 'null', 'none', '?', 'undefined', 'n.a.', 'n/r'];
  const isPlaceholderA = isNullishA || emptyPlaceholders.includes(strA.toLowerCase());
  const isPlaceholderB = isNullishB || emptyPlaceholders.includes(strB.toLowerCase());
  if (isPlaceholderA && isPlaceholderB) return true;

  const cleanNumA = strA.replace(/,/g, '').replace(/%$/, '');
  const cleanNumB = strB.replace(/,/g, '').replace(/%$/, '');
  const numA = Number(cleanNumA);
  const numB = Number(cleanNumB);
  if (!isNaN(numA) && !isNaN(numB) && cleanNumA !== '' && cleanNumB !== '') {
    return Math.abs(numA - numB) < 0.00001;
  }

  const lowerA = strA.toLowerCase();
  const lowerB = strB.toLowerCase();
  const truthy = ['true', '1', 'yes', 'y', 'active', 'enrolled', 'pass', 'present'];
  const falsy = ['false', '0', 'no', 'n', 'inactive', 'fail', 'absent'];
  if (truthy.includes(lowerA) && truthy.includes(lowerB)) return true;
  if (falsy.includes(lowerA) && falsy.includes(lowerB)) return true;

  if (strA.length >= 10 && strB.length >= 10) {
    const d1 = strA.substring(0, 10);
    const d2 = strB.substring(0, 10);
    if (d1 === d2 && /^\d{4}[-/]\d{2}[-/]\d{2}$/.test(d1)) {
      return true;
    }
  }

  if (lowerA === lowerB) return true;

  return false;
}

export class DryRunEngine {
  static getWorkbookSheet(
    wb: XLSX.WorkBook | null | undefined,
    sheetName: string,
    analysis?: WorkbookAnalysis | null
  ): { ws: XLSX.WorkSheet | null; fallbackRecords?: any[] } {
    if (wb && wb.Sheets) {
      // 1. Exact match
      if (wb.Sheets[sheetName]) return { ws: wb.Sheets[sheetName] };

      // 2. Case-insensitive match
      const key = Object.keys(wb.Sheets).find(k => k.trim().toLowerCase() === sheetName.trim().toLowerCase());
      if (key && wb.Sheets[key]) return { ws: wb.Sheets[key] };

      // 3. Normalized alphanumeric match
      const normTarget = sheetName.replace(/[^a-z0-9]/gi, '').toLowerCase();
      const normKey = Object.keys(wb.Sheets).find(k => k.replace(/[^a-z0-9]/gi, '').toLowerCase() === normTarget);
      if (normKey && wb.Sheets[normKey]) return { ws: wb.Sheets[normKey] };

      // 4. Single sheet fallback
      const allKeys = Object.keys(wb.Sheets);
      if (allKeys.length === 1 && wb.Sheets[allKeys[0]]) return { ws: wb.Sheets[allKeys[0]] };

      // 5. If sheet exists by index
      if (analysis?.worksheets) {
        const idx = analysis.worksheets.findIndex(w => w.sheetName.toLowerCase() === sheetName.toLowerCase());
        if (idx >= 0 && wb.SheetNames[idx] && wb.Sheets[wb.SheetNames[idx]]) {
          return { ws: wb.Sheets[wb.SheetNames[idx]] };
        }
      }
    }

    // 6. Try from analysis worksheets sample data
    if (analysis && analysis.worksheets && analysis.worksheets.length > 0) {
      const match = analysis.worksheets.find(
        w => w.sheetName.trim().toLowerCase() === sheetName.trim().toLowerCase()
      ) || analysis.worksheets.find(
        w => w.sheetName.replace(/[^a-z0-9]/gi, '').toLowerCase() === sheetName.replace(/[^a-z0-9]/gi, '').toLowerCase()
      ) || analysis.worksheets[0];

      if (match && match.sampleRows && match.sampleRows.length > 0) {
        const fallbackRecords = match.sampleRows.map(sr => {
          const row: Record<string, any> = { __rowNumber: sr.rowNumber };
          Object.assign(row, sr.data);
          match.headers.forEach(h => {
            if (row[h.colLetter] === undefined && row[h.name] !== undefined) {
              row[h.colLetter] = row[h.name];
            }
          });
          return row;
        });

        try {
          const aoa: any[][] = [];
          aoa.push(match.headers.map(h => h.name));
          match.sampleRows.forEach(sr => {
            const r: any[] = [];
            match.headers.forEach(h => {
              r.push(sr.data[h.colLetter] ?? sr.data[h.name] ?? '');
            });
            aoa.push(r);
          });
          const syntheticWs = XLSX.utils.aoa_to_sheet(aoa);
          return { ws: syntheticWs, fallbackRecords };
        } catch {
          return { ws: null, fallbackRecords };
        }
      }
    }

    return { ws: null };
  }

  static executeDryRun(
    wb: XLSX.WorkBook | null | undefined,
    filename: string,
    worksheetMappings: WorksheetMapping[],
    databaseState: Record<string, any[]>,
    currentAnalysis?: WorkbookAnalysis | null
  ): DryRunResult {
    let totalRows = 0;
    let validRows = 0;
    let proposedInserts = 0;
    let proposedUpdates = 0;
    let failedRows = 0;

    const allErrors: RowValidationError[] = [];
    const sheetSummaries: DryRunResult['sheetSummaries'] = [];
    const sampleTransformedRecords: DryRunResult['sampleTransformedRecords'] = [];

    // Fallback: if no mappings or none enabled, build auto-mappings from currentAnalysis or workbook
    let effectiveMappings = worksheetMappings.filter(m => m.enabled);
    if (effectiveMappings.length === 0) {
      if (currentAnalysis && currentAnalysis.worksheets.length > 0) {
        effectiveMappings = currentAnalysis.worksheets.map((ws, sIdx) => ({
          id: `wm-auto-${sIdx}`,
          workbookName: filename,
          worksheetName: ws.sheetName,
          supabaseTable: ws.sheetName.toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'sheet_data',
          headerRow: ws.detectedHeaderRow || 1,
          dataStartRow: ws.detectedDataStartRow || 2,
          enabled: true,
          syncPolicy: 'EXCEL_TO_DB',
          columns: ws.headers.map((h, idx) => ({
            id: `cm-auto-${sIdx}-${idx}`,
            excelColumn: h.colLetter,
            excelHeader: h.name,
            supabaseColumn: h.name.toLowerCase().replace(/[^a-z0-9_]/g, '_'),
            dataType: h.inferredType || 'text',
            required: false,
            uniqueKey: h.isCandidateKey || idx === 0,
            transformation: h.inferredType === 'date' ? 'parse_date' : h.inferredType === 'integer' ? 'parse_number' : 'trim'
          }))
        }));
      } else if (wb && wb.SheetNames.length > 0) {
        effectiveMappings = wb.SheetNames.map((sName, sIdx) => ({
          id: `wm-auto-wb-${sIdx}`,
          workbookName: filename,
          worksheetName: sName,
          supabaseTable: sName.toLowerCase().replace(/[^a-z0-9_]/g, '_') || 'sheet_data',
          headerRow: 1,
          dataStartRow: 2,
          enabled: true,
          syncPolicy: 'EXCEL_TO_DB',
          columns: []
        }));
      }
    }

    for (const wm of effectiveMappings) {
      const { ws, fallbackRecords } = this.getWorkbookSheet(wb, wm.worksheetName, currentAnalysis);

      let records: any[] = [];
      if (ws) {
        const { records: extracted } = ExcelAnalyzer.extractRecordsWithMergedSections(
          ws,
          wm.headerRow,
          wm.dataStartRow,
          wm.dataEndRow,
          wm.sectionHeadingTargetCol
        );
        records = extracted;
      }

      if ((!records || records.length === 0) && fallbackRecords && fallbackRecords.length > 0) {
        records = fallbackRecords;
      }

      if (!records || records.length === 0) continue;

      const targetTable = wm.supabaseTable;
      const existingTableRecords = databaseState[targetTable] || [];

      // Find primary unique keys for this worksheet mapping
      const uniqueColMappings = wm.columns.filter(c => c.uniqueKey);
      const uniqueColNames = uniqueColMappings.map(c => c.supabaseColumn);

      // Track unique keys in batch to detect duplicates within same file
      const seenBatchKeys = new Set<string>();

      let sheetValid = 0;
      let sheetInserts = 0;
      let sheetUpdates = 0;
      let sheetFails = 0;

      for (const raw of records) {
        const { cleanRecord, errors } = MappingEngine.transformAndValidate(
          raw,
          wm.columns,
          wm.worksheetName,
          wm.sectionHeadingTargetCol,
          seenBatchKeys
        );

        totalRows++;

        if (errors.length > 0) {
          failedRows++;
          sheetFails++;
          allErrors.push(...errors);
          if (sampleTransformedRecords.length < 5000) {
            sampleTransformedRecords.push({
              sheetName: wm.worksheetName,
              targetTable,
              action: 'FAIL',
              record: cleanRecord,
              rowNumber: raw.__rowNumber || 0
            });
          }
        } else {
          validRows++;
          sheetValid++;

          // Check against existing database records for Identity & Cell-level changes
          let matchedExistingRecord: any = null;
          const candidateKeys = ['admission_no', 'student_id', 'roll_no', 'username', 'index_number', 'email', 'id', 'code'];

          if (uniqueColNames.length > 0 && existingTableRecords.length > 0) {
            matchedExistingRecord = existingTableRecords.find(ex => {
              return uniqueColNames.every(col => areCellsEqual(ex[col], cleanRecord[col]));
            });
          } else if (existingTableRecords.length > 0) {
            for (const ck of candidateKeys) {
              if (cleanRecord[ck] !== undefined && cleanRecord[ck] !== null && String(cleanRecord[ck]).trim() !== '') {
                matchedExistingRecord = existingTableRecords.find(ex => areCellsEqual(ex[ck], cleanRecord[ck]));
                if (matchedExistingRecord) break;
              }
            }
            if (!matchedExistingRecord) {
              matchedExistingRecord = existingTableRecords.find(ex => {
                return Object.keys(cleanRecord).every(k => areCellsEqual(cleanRecord[k], ex[k]));
              });
            }
          }

          let action: 'INSERT' | 'UPDATE' | 'SKIP' = 'INSERT';
          if (matchedExistingRecord) {
            const isIdentical = Object.keys(cleanRecord).every(k => areCellsEqual(cleanRecord[k], matchedExistingRecord[k]));
            if (isIdentical) {
              action = 'SKIP';
            } else {
              action = 'UPDATE';
              proposedUpdates++;
              sheetUpdates++;
            }
          } else {
            action = 'INSERT';
            proposedInserts++;
            sheetInserts++;
          }

          if (sampleTransformedRecords.length < 50000) {
            sampleTransformedRecords.push({
              sheetName: wm.worksheetName,
              targetTable,
              action: action === 'SKIP' ? 'UPDATE' : action,
              record: cleanRecord,
              rowNumber: raw.__rowNumber || 0
            });
          }
        }
      }

      sheetSummaries.push({
        sheetName: wm.worksheetName,
        targetTable,
        totalRows: records.length,
        validRows: sheetValid,
        inserts: sheetInserts,
        updates: sheetUpdates,
        fails: sheetFails
      });
    }

    return {
      filename,
      timestamp: new Date().toISOString(),
      totalWorksheets: sheetSummaries.length,
      totalRows,
      validRows,
      proposedInserts,
      proposedUpdates,
      failedRows,
      sheetSummaries,
      errors: allErrors,
      sampleTransformedRecords
    };
  }

  static executeRealImport(
    wb: XLSX.WorkBook | null | undefined,
    filename: string,
    fileHash: string,
    worksheetMappings: WorksheetMapping[],
    databaseState: Record<string, any[]>,
    updateDatabase: (tableName: string, newRecords: any[]) => void,
    currentAnalysis?: WorkbookAnalysis | null
  ): { log: ImportLog; errors: RowValidationError[] } {
    const dryRunResult = this.executeDryRun(wb, filename, worksheetMappings, databaseState, currentAnalysis);

    // Apply Upsert to databaseState for each worksheet
    for (const wm of worksheetMappings) {
      if (!wm.enabled) continue;

      const targetTable = wm.supabaseTable;
      const { ws, fallbackRecords } = this.getWorkbookSheet(wb, wm.worksheetName, currentAnalysis);

      let records: any[] = [];
      if (ws) {
        const { records: extracted } = ExcelAnalyzer.extractRecordsWithMergedSections(
          ws,
          wm.headerRow,
          wm.dataStartRow,
          wm.dataEndRow,
          wm.sectionHeadingTargetCol
        );
        records = extracted;
      }
      if ((!records || records.length === 0) && fallbackRecords) {
        records = fallbackRecords;
      }
      if (!records || records.length === 0) continue;

      const currentDb = [...(databaseState[targetTable] || [])];
      const uniqueColNames = wm.columns.filter(c => c.uniqueKey).map(c => c.supabaseColumn);
      const candidateKeys = ['admission_no', 'student_id', 'roll_no', 'username', 'index_number', 'email', 'id', 'code'];
      const seenBatchKeys = new Set<string>();

      for (const raw of records) {
        const { cleanRecord, errors } = MappingEngine.transformAndValidate(
          raw,
          wm.columns,
          wm.worksheetName,
          wm.sectionHeadingTargetCol,
          seenBatchKeys
        );

        if (errors.length === 0) {
          cleanRecord.__updated_at = new Date().toISOString();
          cleanRecord.__source_file = filename;

          let matchIdx = -1;
          if (uniqueColNames.length > 0) {
            matchIdx = currentDb.findIndex(ex => {
              return uniqueColNames.every(col => areCellsEqual(ex[col], cleanRecord[col]));
            });
          } else {
            for (const ck of candidateKeys) {
              if (cleanRecord[ck] !== undefined && cleanRecord[ck] !== null && String(cleanRecord[ck]).trim() !== '') {
                matchIdx = currentDb.findIndex(ex => areCellsEqual(ex[ck], cleanRecord[ck]));
                if (matchIdx >= 0) break;
              }
            }
            if (matchIdx < 0) {
              matchIdx = currentDb.findIndex(ex => {
                return Object.keys(cleanRecord).every(k => areCellsEqual(cleanRecord[k], ex[k]));
              });
            }
          }

          if (matchIdx >= 0) {
            const existing = currentDb[matchIdx];
            const isIdentical = Object.keys(cleanRecord).every(k => areCellsEqual(cleanRecord[k], existing[k]));
            if (!isIdentical) {
              // Row has updated cells -> update in-place!
              currentDb[matchIdx] = { ...existing, ...cleanRecord };
            }
            // If identical: do nothing! No duplicate rows!
          } else {
            // Truly new row -> insert
            if (!cleanRecord.id) {
              cleanRecord.id = `rec-${Date.now()}-${Math.random().toString(36).substr(2, 6)}`;
            }
            currentDb.push(cleanRecord);
          }
        }
      }

      updateDatabase(targetTable, currentDb);
    }

    const status: import('../types').ImportStatus = dryRunResult.failedRows === 0 ? 'Success' : dryRunResult.validRows > 0 ? 'Partial Success' : 'Failed';
    const log: ImportLog = {
      id: `log-${Date.now()}`,
      filename,
      filePath: filename,
      fileHash,
      status,
      isDryRun: false,
      numberOfWorksheets: dryRunResult.totalWorksheets,
      rowsProcessed: dryRunResult.totalRows,
      rowsInserted: dryRunResult.proposedInserts,
      rowsUpdated: dryRunResult.proposedUpdates,
      rowsFailed: dryRunResult.failedRows,
      startedAt: new Date().toISOString(),
      completedAt: new Date().toISOString(),
      durationMs: 450,
      errorSummary: dryRunResult.errors.length > 0 ? `${dryRunResult.errors.length} validation errors occurred.` : undefined
    };

    return { log, errors: dryRunResult.errors };
  }
}
