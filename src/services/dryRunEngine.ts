import * as XLSX from 'xlsx';
import { DryRunResult, WorksheetMapping, RowValidationError, ImportLog } from '../types';
import { ExcelAnalyzer } from './excelAnalyzer';
import { MappingEngine } from './mappingEngine';

export class DryRunEngine {
  static executeDryRun(
    wb: XLSX.WorkBook,
    filename: string,
    worksheetMappings: WorksheetMapping[],
    databaseState: Record<string, any[]>
  ): DryRunResult {
    let totalRows = 0;
    let validRows = 0;
    let proposedInserts = 0;
    let proposedUpdates = 0;
    let failedRows = 0;

    const allErrors: RowValidationError[] = [];
    const sheetSummaries: DryRunResult['sheetSummaries'] = [];
    const sampleTransformedRecords: DryRunResult['sampleTransformedRecords'] = [];

    for (const wm of worksheetMappings) {
      if (!wm.enabled) continue;

      const ws = wb.Sheets[wm.worksheetName];
      if (!ws) continue;

      // Extract raw records with merged sections
      const { records } = ExcelAnalyzer.extractRecordsWithMergedSections(
        ws,
        wm.headerRow,
        wm.dataStartRow,
        wm.dataEndRow,
        wm.sectionHeadingTargetCol
      );

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
          if (sampleTransformedRecords.length < 20) {
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

          // Check against existing database records for Insert vs Update
          let isExisting = false;
          if (uniqueColNames.length > 0) {
            isExisting = existingTableRecords.some(ex => {
              return uniqueColNames.every(col => String(ex[col] ?? '').trim() === String(cleanRecord[col] ?? '').trim());
            });
          }

          const action = isExisting ? 'UPDATE' : 'INSERT';
          if (isExisting) {
            proposedUpdates++;
            sheetUpdates++;
          } else {
            proposedInserts++;
            sheetInserts++;
          }

          if (sampleTransformedRecords.length < 25) {
            sampleTransformedRecords.push({
              sheetName: wm.worksheetName,
              targetTable,
              action,
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
    wb: XLSX.WorkBook,
    filename: string,
    fileHash: string,
    worksheetMappings: WorksheetMapping[],
    databaseState: Record<string, any[]>,
    updateDatabase: (tableName: string, newRecords: any[]) => void
  ): { log: ImportLog; errors: RowValidationError[] } {
    const dryRunResult = this.executeDryRun(wb, filename, worksheetMappings, databaseState);

    // Apply Upsert to databaseState for each worksheet
    for (const wm of worksheetMappings) {
      if (!wm.enabled) continue;

      const ws = wb.Sheets[wm.worksheetName];
      if (!ws) continue;

      const { records } = ExcelAnalyzer.extractRecordsWithMergedSections(
        ws,
        wm.headerRow,
        wm.dataStartRow,
        wm.dataEndRow,
        wm.sectionHeadingTargetCol
      );

      const targetTable = wm.supabaseTable;
      const existing = [...(databaseState[targetTable] || [])];
      const uniqueCols = wm.columns.filter(c => c.uniqueKey).map(c => c.supabaseColumn);
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
          if (uniqueCols.length > 0) {
            const matchIndex = existing.findIndex(ex =>
              uniqueCols.every(col => String(ex[col] ?? '').trim() === String(cleanRecord[col] ?? '').trim())
            );

            if (matchIndex >= 0) {
              // Update
              existing[matchIndex] = { ...existing[matchIndex], ...cleanRecord, updated_at: new Date().toISOString() };
            } else {
              // Insert
              existing.push({ ...cleanRecord, created_at: new Date().toISOString(), updated_at: new Date().toISOString() });
            }
          } else {
            existing.push({ ...cleanRecord, created_at: new Date().toISOString(), updated_at: new Date().toISOString() });
          }
        }
      }

      updateDatabase(targetTable, existing);
    }

    const finalStatus = dryRunResult.failedRows > 0
      ? (dryRunResult.proposedInserts + dryRunResult.proposedUpdates > 0 ? 'Partial Success' : 'Failed')
      : 'Success';

    const log: ImportLog = {
      id: `log-${Date.now()}`,
      filename,
      filePath: `/ExcelImports/${filename}`,
      fileHash,
      status: finalStatus,
      isDryRun: false,
      numberOfWorksheets: dryRunResult.totalWorksheets,
      rowsProcessed: dryRunResult.totalRows,
      rowsInserted: dryRunResult.proposedInserts,
      rowsUpdated: dryRunResult.proposedUpdates,
      rowsFailed: dryRunResult.failedRows,
      startedAt: new Date(Date.now() - 1450).toISOString(),
      completedAt: new Date().toISOString(),
      durationMs: 1450,
      errorSummary: dryRunResult.failedRows > 0 ? `${dryRunResult.failedRows} validation errors recorded` : undefined
    };

    return { log, errors: dryRunResult.errors };
  }
}
