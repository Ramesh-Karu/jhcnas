import * as XLSX from 'xlsx';
import { 
  WorksheetMapping, 
  TwoWaySyncResult, 
  SyncConflictRecord, 
  ConflictFieldDiff, 
  ConflictResolutionChoice,
  SyncBaselineRecord,
  TableSyncPolicy
} from '../types';
import { ExcelAnalyzer } from './excelAnalyzer';
import { MappingEngine } from './mappingEngine';

export class TwoWaySyncEngine {
  /**
   * Generates a stable deterministic hash string for a record based on mapped columns
   */
  static computeRecordHash(record: Record<string, any>, columns: { supabaseColumn: string }[]): string {
    const normalized: Record<string, string> = {};
    for (const col of columns) {
      const val = record[col.supabaseColumn];
      normalized[col.supabaseColumn] = val !== undefined && val !== null ? String(val).trim() : '';
    }
    return JSON.stringify(normalized);
  }

  /**
   * Compare Nextcloud Excel worksheets with Supabase PostgreSQL database state
   * using Sync Baselines and Per-Table Sync Policies (EXCEL_TO_DB, DB_TO_EXCEL, BIDIRECTIONAL, READ_ONLY).
   */
  static detectTwoWayDiffs(
    wb: XLSX.WorkBook,
    filename: string,
    worksheetMappings: WorksheetMapping[],
    databaseState: Record<string, any[]>,
    baselines: Record<string, SyncBaselineRecord> = {}
  ): TwoWaySyncResult {
    const allConflictRecords: SyncConflictRecord[] = [];
    let inSyncCount = 0;
    let autoPushToDbCount = 0;
    let autoPushToExcelCount = 0;
    let policyBlockedCount = 0;
    let excelOnlyCount = 0;
    let supabaseOnlyCount = 0;
    let conflictCount = 0;
    let totalCompared = 0;

    for (const wm of worksheetMappings) {
      if (!wm.enabled) continue;

      const ws = wb.Sheets[wm.worksheetName];
      if (!ws) continue;

      const tablePolicy: TableSyncPolicy = wm.syncPolicy || 'BIDIRECTIONAL';

      // Extract raw records from Excel sheet
      const { records: rawExcelRows } = ExcelAnalyzer.extractRecordsWithMergedSections(
        ws,
        wm.headerRow,
        wm.dataStartRow,
        wm.dataEndRow,
        wm.sectionHeadingTargetCol
      );

      const targetTable = wm.supabaseTable;
      const existingDbRows: any[] = databaseState[targetTable] || [];

      // Identify unique primary key columns
      const uniqueColMappings = wm.columns.filter(c => c.uniqueKey);
      const uniqueColNames = uniqueColMappings.length > 0 
        ? uniqueColMappings.map(c => c.supabaseColumn)
        : [wm.columns[0]?.supabaseColumn || 'id'];

      const primaryKeyCol = uniqueColNames[0] || 'id';

      // Build cleaned Excel records with validation
      const seenBatchKeys = new Set<string>();
      const cleanExcelMap = new Map<string, { record: Record<string, any>; rowNumber: number }>();

      for (const raw of rawExcelRows) {
        const { cleanRecord } = MappingEngine.transformAndValidate(
          raw,
          wm.columns,
          wm.worksheetName,
          wm.sectionHeadingTargetCol,
          seenBatchKeys
        );

        const pkVal = String(cleanRecord[primaryKeyCol] ?? '').trim();
        if (pkVal) {
          cleanExcelMap.set(pkVal, {
            record: cleanRecord,
            rowNumber: raw.__rowNumber || 0
          });
        }
      }

      // Map Supabase DB rows by primary key
      const dbMap = new Map<string, Record<string, any>>();
      for (const dbRow of existingDbRows) {
        const pkVal = String(dbRow[primaryKeyCol] ?? '').trim();
        if (pkVal) {
          dbMap.set(pkVal, dbRow);
        }
      }

      // 1. Process records found in Excel
      for (const [pkVal, { record: excelRec, rowNumber }] of cleanExcelMap.entries()) {
        totalCompared++;
        const recordKey = `${targetTable}:${pkVal}`;
        const dbRec = dbMap.get(pkVal);
        const baseline = baselines[recordKey];
        const excelHash = this.computeRecordHash(excelRec, wm.columns);

        if (!dbRec) {
          // EXCEL_ONLY: Exists in Excel but not yet in Supabase
          excelOnlyCount++;

          const fieldDiffs: ConflictFieldDiff[] = wm.columns.map(col => ({
            supabaseColumn: col.supabaseColumn,
            excelHeader: col.excelHeader,
            excelColumn: col.excelColumn,
            excelValue: excelRec[col.supabaseColumn] ?? '',
            supabaseValue: '(Not in DB)',
            selectedSource: 'excel'
          }));

          if (tablePolicy === 'DB_TO_EXCEL') {
            // Protected: Table is Supabase-Authoritative! Excel cannot insert rows into DB.
            policyBlockedCount++;
            allConflictRecords.push({
              id: `blocked-ex-${wm.worksheetName}-${pkVal}`,
              worksheetName: wm.worksheetName,
              tableName: targetTable,
              primaryKeyCol,
              primaryKeyValue: pkVal,
              detectedAt: new Date().toISOString(),
              state: 'POLICY_BLOCKED',
              tableSyncPolicy: tablePolicy,
              changeOrigin: 'NEW_IN_EXCEL',
              policyNotice: `Table "${wm.worksheetName}" is set to Supabase-Authoritative (DB Master). Nextcloud Excel cannot insert records into live Supabase database.`,
              autoApplyReason: `Blocked by table policy: Only Supabase can create records for ${wm.worksheetName}. Rogue Excel rows are protected from polluting the DB.`,
              excelRowNumber: rowNumber,
              fieldDiffs,
              excelFullRecord: excelRec,
              supabaseFullRecord: {},
              resolutionChoice: undefined
            });
          } else if (tablePolicy === 'READ_ONLY') {
            policyBlockedCount++;
            allConflictRecords.push({
              id: `blocked-ro-${wm.worksheetName}-${pkVal}`,
              worksheetName: wm.worksheetName,
              tableName: targetTable,
              primaryKeyCol,
              primaryKeyValue: pkVal,
              detectedAt: new Date().toISOString(),
              state: 'POLICY_BLOCKED',
              tableSyncPolicy: tablePolicy,
              changeOrigin: 'NEW_IN_EXCEL',
              policyNotice: `Table "${wm.worksheetName}" is configured as Read-Only. Automated insert is disabled.`,
              autoApplyReason: `Audit only: Table ${wm.worksheetName} is read-only.`,
              excelRowNumber: rowNumber,
              fieldDiffs,
              excelFullRecord: excelRec,
              supabaseFullRecord: {},
              resolutionChoice: undefined
            });
          } else {
            // EXCEL_TO_DB or BIDIRECTIONAL: allowed to auto-push into Supabase
            autoPushToDbCount++;
            allConflictRecords.push({
              id: `diff-ex-${wm.worksheetName}-${pkVal}`,
              worksheetName: wm.worksheetName,
              tableName: targetTable,
              primaryKeyCol,
              primaryKeyValue: pkVal,
              detectedAt: new Date().toISOString(),
              state: 'AUTO_PUSH_TO_DB',
              tableSyncPolicy: tablePolicy,
              changeOrigin: 'NEW_IN_EXCEL',
              autoApplyReason: tablePolicy === 'EXCEL_TO_DB'
                ? `Nextcloud Master: New record in "${wm.worksheetName}" auto-inserts to Supabase DB with 0 clicks.`
                : 'New record created in Nextcloud Excel. Staged to automatically insert into Supabase with 0 clicks.',
              excelRowNumber: rowNumber,
              fieldDiffs,
              excelFullRecord: excelRec,
              supabaseFullRecord: {},
              resolutionChoice: 'excel'
            });
          }
        } else {
          // Both exist: check field-by-field for differences
          const dbHash = this.computeRecordHash(dbRec, wm.columns);
          const fieldDiffs: ConflictFieldDiff[] = [];

          for (const col of wm.columns) {
            const colName = col.supabaseColumn;
            const exVal = excelRec[colName];
            const dbVal = dbRec[colName];

            // Normalize for comparison
            const exStr = exVal !== undefined && exVal !== null ? String(exVal).trim() : '';
            const dbStr = dbVal !== undefined && dbVal !== null ? String(dbVal).trim() : '';

            if (exStr !== dbStr) {
              fieldDiffs.push({
                supabaseColumn: colName,
                excelHeader: col.excelHeader,
                excelColumn: col.excelColumn,
                excelValue: exVal ?? '',
                supabaseValue: dbVal ?? '',
                selectedSource: 'excel'
              });
            }
          }

          if (fieldDiffs.length === 0) {
            inSyncCount++;
            allConflictRecords.push({
              id: `insync-${wm.worksheetName}-${pkVal}`,
              worksheetName: wm.worksheetName,
              tableName: targetTable,
              primaryKeyCol,
              primaryKeyValue: pkVal,
              detectedAt: new Date().toISOString(),
              state: 'IN_SYNC',
              tableSyncPolicy: tablePolicy,
              changeOrigin: 'IDENTICAL',
              autoApplyReason: 'Identical in Nextcloud Excel and Supabase.',
              excelRowNumber: rowNumber,
              fieldDiffs: [],
              excelFullRecord: excelRec,
              supabaseFullRecord: dbRec
            });
          } else {
            // Differences exist! Evaluate based on TableSyncPolicy and actual baseline or timestamps
            const excelChanged = baseline ? (excelHash !== baseline.excelHash) : true;
            const dbChanged = baseline 
              ? (dbHash !== baseline.supabaseHash) 
              : Boolean(dbRec.updated_at && dbRec.created_at && dbRec.updated_at !== dbRec.created_at);

            if (tablePolicy === 'EXCEL_TO_DB') {
              // Nextcloud Excel is the Master Source of Truth for this table
              if (excelChanged) {
                autoPushToDbCount++;
                allConflictRecords.push({
                  id: `autopush-db-${wm.worksheetName}-${pkVal}`,
                  worksheetName: wm.worksheetName,
                  tableName: targetTable,
                  primaryKeyCol,
                  primaryKeyValue: pkVal,
                  detectedAt: new Date().toISOString(),
                  state: 'AUTO_PUSH_TO_DB',
                  tableSyncPolicy: tablePolicy,
                  changeOrigin: 'EXCEL_ONLY_EDIT',
                  autoApplyReason: `Nextcloud Master Policy: Excel changes for "${wm.worksheetName}" auto-sync to Supabase DB with 0 clicks.`,
                  excelRowNumber: rowNumber,
                  fieldDiffs,
                  excelFullRecord: excelRec,
                  supabaseFullRecord: dbRec,
                  resolutionChoice: 'excel'
                });
              } else {
                // Only Supabase changed, but Excel is Master! Blocked from overwriting Excel!
                policyBlockedCount++;
                allConflictRecords.push({
                  id: `blocked-sup-${wm.worksheetName}-${pkVal}`,
                  worksheetName: wm.worksheetName,
                  tableName: targetTable,
                  primaryKeyCol,
                  primaryKeyValue: pkVal,
                  detectedAt: new Date().toISOString(),
                  state: 'POLICY_BLOCKED',
                  tableSyncPolicy: tablePolicy,
                  changeOrigin: 'SUPABASE_ONLY_EDIT',
                  policyNotice: `Table "${wm.worksheetName}" is set to Nextcloud-Authoritative (Excel Master). Supabase edits are blocked from overwriting Nextcloud Excel.`,
                  autoApplyReason: `Protected by table policy: Nextcloud Excel is the single source of truth for "${wm.worksheetName}". Supabase updates will not overwrite your spreadsheet.`,
                  excelRowNumber: rowNumber,
                  fieldDiffs,
                  excelFullRecord: excelRec,
                  supabaseFullRecord: dbRec,
                  resolutionChoice: undefined
                });
              }
            } else if (tablePolicy === 'DB_TO_EXCEL') {
              // Supabase DB is the Master Source of Truth for this table (e.g. attendance, exams)
              if (dbChanged) {
                autoPushToExcelCount++;
                allConflictRecords.push({
                  id: `autopush-ex-${wm.worksheetName}-${pkVal}`,
                  worksheetName: wm.worksheetName,
                  tableName: targetTable,
                  primaryKeyCol,
                  primaryKeyValue: pkVal,
                  detectedAt: new Date().toISOString(),
                  state: 'AUTO_PUSH_TO_EXCEL',
                  tableSyncPolicy: tablePolicy,
                  changeOrigin: 'SUPABASE_ONLY_EDIT',
                  autoApplyReason: `Supabase Master Policy: Live DB changes for "${wm.worksheetName}" auto-mirror to Nextcloud Excel with 0 clicks.`,
                  excelRowNumber: rowNumber,
                  fieldDiffs,
                  excelFullRecord: excelRec,
                  supabaseFullRecord: dbRec,
                  resolutionChoice: 'supabase'
                });
              } else {
                // Only Excel changed, but DB is Master! Blocked from overwriting DB!
                policyBlockedCount++;
                allConflictRecords.push({
                  id: `blocked-exc-${wm.worksheetName}-${pkVal}`,
                  worksheetName: wm.worksheetName,
                  tableName: targetTable,
                  primaryKeyCol,
                  primaryKeyValue: pkVal,
                  detectedAt: new Date().toISOString(),
                  state: 'POLICY_BLOCKED',
                  tableSyncPolicy: tablePolicy,
                  changeOrigin: 'EXCEL_ONLY_EDIT',
                  policyNotice: `Table "${wm.worksheetName}" is set to Supabase-Authoritative (DB Master). Excel edits are blocked from overwriting live database state.`,
                  autoApplyReason: `Protected by table policy: Supabase is the single source of truth for "${wm.worksheetName}". Spreadsheets cannot overwrite live database records.`,
                  excelRowNumber: rowNumber,
                  fieldDiffs,
                  excelFullRecord: excelRec,
                  supabaseFullRecord: dbRec,
                  resolutionChoice: undefined
                });
              }
            } else if (tablePolicy === 'READ_ONLY') {
              policyBlockedCount++;
              allConflictRecords.push({
                id: `blocked-ro-${wm.worksheetName}-${pkVal}`,
                worksheetName: wm.worksheetName,
                tableName: targetTable,
                primaryKeyCol,
                primaryKeyValue: pkVal,
                detectedAt: new Date().toISOString(),
                state: 'POLICY_BLOCKED',
                tableSyncPolicy: tablePolicy,
                changeOrigin: 'CONCURRENT_COLLISION',
                policyNotice: `Table "${wm.worksheetName}" is configured as Read-Only. Synchronization is disabled.`,
                autoApplyReason: `Audit only: Table ${wm.worksheetName} is read-only.`,
                excelRowNumber: rowNumber,
                fieldDiffs,
                excelFullRecord: excelRec,
                supabaseFullRecord: dbRec,
                resolutionChoice: undefined
              });
            } else {
              // Full BIDIRECTIONAL policy
              if (excelChanged && !dbChanged) {
                autoPushToDbCount++;
                allConflictRecords.push({
                  id: `autopush-db-${wm.worksheetName}-${pkVal}`,
                  worksheetName: wm.worksheetName,
                  tableName: targetTable,
                  primaryKeyCol,
                  primaryKeyValue: pkVal,
                  detectedAt: new Date().toISOString(),
                  state: 'AUTO_PUSH_TO_DB',
                  tableSyncPolicy: tablePolicy,
                  changeOrigin: 'EXCEL_ONLY_EDIT',
                  autoApplyReason: 'Edited in Nextcloud Excel only. Supabase was untouched. Auto-syncs to database with 0 clicks.',
                  excelRowNumber: rowNumber,
                  fieldDiffs,
                  excelFullRecord: excelRec,
                  supabaseFullRecord: dbRec,
                  resolutionChoice: 'excel'
                });
              } else if (!excelChanged && dbChanged) {
                autoPushToExcelCount++;
                allConflictRecords.push({
                  id: `autopush-ex-${wm.worksheetName}-${pkVal}`,
                  worksheetName: wm.worksheetName,
                  tableName: targetTable,
                  primaryKeyCol,
                  primaryKeyValue: pkVal,
                  detectedAt: new Date().toISOString(),
                  state: 'AUTO_PUSH_TO_EXCEL',
                  tableSyncPolicy: tablePolicy,
                  changeOrigin: 'SUPABASE_ONLY_EDIT',
                  autoApplyReason: 'Edited in Supabase only. Nextcloud Excel was untouched. Auto-syncs back to spreadsheet with 0 clicks.',
                  excelRowNumber: rowNumber,
                  fieldDiffs,
                  excelFullRecord: excelRec,
                  supabaseFullRecord: dbRec,
                  resolutionChoice: 'supabase'
                });
              } else {
                // True Concurrent Collision on a bidirectional table
                conflictCount++;
                allConflictRecords.push({
                  id: `conflict-${wm.worksheetName}-${pkVal}`,
                  worksheetName: wm.worksheetName,
                  tableName: targetTable,
                  primaryKeyCol,
                  primaryKeyValue: pkVal,
                  detectedAt: new Date().toISOString(),
                  state: 'CONFLICT',
                  tableSyncPolicy: tablePolicy,
                  changeOrigin: 'CONCURRENT_COLLISION',
                  autoApplyReason: 'Simultaneous edit collision! Both Nextcloud and Supabase were modified concurrently on this bidirectional table. Requires your manual decision.',
                  excelRowNumber: rowNumber,
                  fieldDiffs,
                  excelFullRecord: excelRec,
                  supabaseFullRecord: dbRec,
                  resolutionChoice: undefined
                });
              }
            }
          }
        }
      }

      // 2. Process records in Supabase that are NOT in Excel
      for (const [pkVal, dbRec] of dbMap.entries()) {
        if (!cleanExcelMap.has(pkVal)) {
          totalCompared++;
          supabaseOnlyCount++;

          const fieldDiffs: ConflictFieldDiff[] = wm.columns.map(col => ({
            supabaseColumn: col.supabaseColumn,
            excelHeader: col.excelHeader,
            excelColumn: col.excelColumn,
            excelValue: '(Not in Excel)',
            supabaseValue: dbRec[col.supabaseColumn] ?? '',
            selectedSource: 'supabase'
          }));

          if (tablePolicy === 'EXCEL_TO_DB') {
            // Nextcloud is Master! Supabase records cannot push into Excel
            policyBlockedCount++;
            allConflictRecords.push({
              id: `blocked-sup-new-${wm.worksheetName}-${pkVal}`,
              worksheetName: wm.worksheetName,
              tableName: targetTable,
              primaryKeyCol,
              primaryKeyValue: pkVal,
              detectedAt: new Date().toISOString(),
              state: 'POLICY_BLOCKED',
              tableSyncPolicy: tablePolicy,
              changeOrigin: 'NEW_IN_SUPABASE',
              policyNotice: `Table "${wm.worksheetName}" is set to Nextcloud-Authoritative (Excel Master). Supabase records will not be exported to Nextcloud Excel.`,
              autoApplyReason: `Protected: Table policy is Nextcloud Master. Nextcloud Excel will not be modified by external Supabase records.`,
              fieldDiffs,
              excelFullRecord: {},
              supabaseFullRecord: dbRec,
              resolutionChoice: undefined
            });
          } else if (tablePolicy === 'READ_ONLY') {
            policyBlockedCount++;
            allConflictRecords.push({
              id: `blocked-ro-new-${wm.worksheetName}-${pkVal}`,
              worksheetName: wm.worksheetName,
              tableName: targetTable,
              primaryKeyCol,
              primaryKeyValue: pkVal,
              detectedAt: new Date().toISOString(),
              state: 'POLICY_BLOCKED',
              tableSyncPolicy: tablePolicy,
              changeOrigin: 'NEW_IN_SUPABASE',
              policyNotice: `Table "${wm.worksheetName}" is configured as Read-Only. Synchronization is disabled.`,
              autoApplyReason: `Audit only: Table ${wm.worksheetName} is read-only.`,
              fieldDiffs,
              excelFullRecord: {},
              supabaseFullRecord: dbRec,
              resolutionChoice: undefined
            });
          } else {
            // DB_TO_EXCEL or BIDIRECTIONAL: auto-push into Excel
            autoPushToExcelCount++;
            allConflictRecords.push({
              id: `diff-db-${wm.worksheetName}-${pkVal}`,
              worksheetName: wm.worksheetName,
              tableName: targetTable,
              primaryKeyCol,
              primaryKeyValue: pkVal,
              detectedAt: new Date().toISOString(),
              state: 'AUTO_PUSH_TO_EXCEL',
              tableSyncPolicy: tablePolicy,
              changeOrigin: 'NEW_IN_SUPABASE',
              autoApplyReason: tablePolicy === 'DB_TO_EXCEL'
                ? `Supabase Master: New row from live app auto-mirrors to Nextcloud Excel with 0 clicks.`
                : 'New record created in Supabase. Staged to auto-sync back to Nextcloud Excel.',
              fieldDiffs,
              excelFullRecord: {},
              supabaseFullRecord: dbRec,
              resolutionChoice: 'supabase'
            });
          }
        }
      }
    }

    return {
      timestamp: new Date().toISOString(),
      filename,
      totalRecordsCompared: totalCompared,
      inSyncCount,
      autoPushToDbCount,
      autoPushToExcelCount,
      policyBlockedCount,
      excelOnlyCount,
      supabaseOnlyCount,
      conflictCount,
      records: allConflictRecords
    };
  }

  /**
   * Automatically executes bi-directional synchronization for all unilateral changes:
   * - Records changed only in Excel -> automatically written to databaseState
   * - Records changed only in Supabase -> automatically written to Excel workbook
   * - Baselines updated to lock in the new synchronized state
   * - ONLY true concurrent collisions and policy-blocked records are left untouched
   */
  static executeAutoSync(
    records: SyncConflictRecord[],
    currentWorkbook: XLSX.WorkBook | null,
    worksheetMappings: WorksheetMapping[],
    databaseState: Record<string, any[]>,
    baselines: Record<string, SyncBaselineRecord> = {}
  ): {
    updatedDatabaseState: Record<string, any[]>;
    updatedWorkbook: XLSX.WorkBook | null;
    updatedBaselines: Record<string, SyncBaselineRecord>;
    syncedToDbCount: number;
    syncedToExcelCount: number;
    remainingConflicts: SyncConflictRecord[];
  } {
    let updatedDatabaseState = { ...databaseState };
    const updatedBaselines = { ...baselines };
    let syncedToDbCount = 0;
    let syncedToExcelCount = 0;

    // 1. Group records by worksheet for reverse Excel updates
    const excelUpdatesBySheet = new Map<string, Record<string, any>[]>();

    for (const rec of records) {
      if (rec.state === 'AUTO_PUSH_TO_DB' || rec.state === 'EXCEL_ONLY') {
        // Auto-Push to Supabase database (0 clicks)
        updatedDatabaseState = this.applyResolutionToDatabase(rec, 'excel', updatedDatabaseState);
        syncedToDbCount++;

        // Update baseline
        const recordKey = `${rec.tableName}:${rec.primaryKeyValue}`;
        const mapping = worksheetMappings.find(m => m.worksheetName === rec.worksheetName);
        const hash = mapping ? this.computeRecordHash(rec.excelFullRecord, mapping.columns) : '';
        updatedBaselines[recordKey] = {
          recordKey,
          tableName: rec.tableName,
          primaryKeyCol: rec.primaryKeyCol,
          primaryKeyValue: rec.primaryKeyValue,
          excelHash: hash,
          supabaseHash: hash,
          lastSyncedAt: new Date().toISOString(),
          syncedValues: { ...rec.excelFullRecord }
        };
      } else if (rec.state === 'AUTO_PUSH_TO_EXCEL' || rec.state === 'SUPABASE_ONLY') {
        // Auto-Push to Nextcloud Excel (0 clicks)
        const list = excelUpdatesBySheet.get(rec.worksheetName) || [];
        list.push(rec.supabaseFullRecord);
        excelUpdatesBySheet.set(rec.worksheetName, list);
        syncedToExcelCount++;

        // Update baseline
        const recordKey = `${rec.tableName}:${rec.primaryKeyValue}`;
        const mapping = worksheetMappings.find(m => m.worksheetName === rec.worksheetName);
        const hash = mapping ? this.computeRecordHash(rec.supabaseFullRecord, mapping.columns) : '';
        updatedBaselines[recordKey] = {
          recordKey,
          tableName: rec.tableName,
          primaryKeyCol: rec.primaryKeyCol,
          primaryKeyValue: rec.primaryKeyValue,
          excelHash: hash,
          supabaseHash: hash,
          lastSyncedAt: new Date().toISOString(),
          syncedValues: { ...rec.supabaseFullRecord }
        };
      }
    }

    // 2. Write updates to Excel workbook if available
    let updatedWorkbook = currentWorkbook;
    if (currentWorkbook && excelUpdatesBySheet.size > 0) {
      for (const [sheetName, sheetRecords] of excelUpdatesBySheet.entries()) {
        const mapping = worksheetMappings.find(m => m.worksheetName === sheetName);
        if (mapping && currentWorkbook.Sheets[sheetName]) {
          try {
            this.writeRecordsToExcelWorkbook(currentWorkbook, sheetName, mapping, sheetRecords);
          } catch (err) {
            console.error(`Failed to write auto-sync records to sheet ${sheetName}:`, err);
          }
        }
      }
    }

    // 3. Filter remaining records: collisions and policy-blocked records remain
    const remainingConflicts = records
      .filter(r => r.state === 'CONFLICT' || r.state === 'POLICY_BLOCKED')
      .map(r => ({ ...r }));

    return {
      updatedDatabaseState,
      updatedWorkbook,
      updatedBaselines,
      syncedToDbCount,
      syncedToExcelCount,
      remainingConflicts
    };
  }

  /**
   * Apply a resolution choice (Excel, Supabase, or Custom Merge) to database state
   */
  static applyResolutionToDatabase(
    record: SyncConflictRecord,
    choice: ConflictResolutionChoice,
    databaseState: Record<string, any[]>
  ): Record<string, any[]> {
    const table = record.tableName;
    const existing = [...(databaseState[table] || [])];
    const pkCol = record.primaryKeyCol;
    const pkVal = record.primaryKeyValue;

    let finalRecord: Record<string, any> = {};

    if (choice === 'excel') {
      finalRecord = { ...record.excelFullRecord };
    } else if (choice === 'supabase') {
      finalRecord = { ...record.supabaseFullRecord };
    } else {
      // Custom Merge: combine field choices
      finalRecord = { ...record.supabaseFullRecord, ...record.excelFullRecord };
      for (const diff of record.fieldDiffs) {
        finalRecord[diff.supabaseColumn] = diff.selectedSource === 'excel'
          ? diff.excelValue
          : diff.supabaseValue;
      }
    }

    finalRecord.updated_at = new Date().toISOString();

    const idx = existing.findIndex(r => String(r[pkCol] ?? '').trim() === pkVal);
    if (idx >= 0) {
      existing[idx] = { ...existing[idx], ...finalRecord };
    } else {
      finalRecord.created_at = new Date().toISOString();
      existing.push(finalRecord);
    }

    return {
      ...databaseState,
      [table]: existing
    };
  }

  /**
   * Reverse-Sync: Update or insert rows back into an Excel workbook (preserving existing sheets and layout)
   */
  static writeRecordsToExcelWorkbook(
    wb: XLSX.WorkBook,
    worksheetName: string,
    mapping: WorksheetMapping,
    recordsToWrite: Record<string, any>[]
  ): { updatedCount: number; newRowsCount: number } {
    const ws = wb.Sheets[worksheetName];
    if (!ws) {
      throw new Error(`Worksheet "${worksheetName}" not found in workbook`);
    }

    const ref = ws['!ref'] || 'A1:A1';
    const range = XLSX.utils.decode_range(ref);
    const pkMapping = mapping.columns.find(c => c.uniqueKey) || mapping.columns[0];
    const pkCol = pkMapping?.supabaseColumn || 'id';
    const pkExcelCol = pkMapping?.excelColumn || 'A';
    const pkColIndex = XLSX.utils.decode_col(pkExcelCol);

    let updatedCount = 0;
    let newRowsCount = 0;

    // Build row index map based on the Excel PK column
    const existingRowByPk = new Map<string, number>();

    for (let R = mapping.dataStartRow - 1; R <= range.e.r; ++R) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: pkColIndex });
      const cell = ws[cellAddress];
      if (cell && cell.v !== undefined && cell.v !== null) {
        const valStr = String(cell.v).trim();
        if (valStr && !existingRowByPk.has(valStr)) {
          existingRowByPk.set(valStr, R);
        }
      }
    }

    let nextAvailableRow = range.e.r + 1;

    for (const record of recordsToWrite) {
      const pkVal = String(record[pkCol] ?? '').trim();
      if (!pkVal) continue;

      let targetRowIndex: number;
      if (existingRowByPk.has(pkVal)) {
        targetRowIndex = existingRowByPk.get(pkVal)!;
        updatedCount++;
      } else {
        targetRowIndex = nextAvailableRow++;
        newRowsCount++;
        existingRowByPk.set(pkVal, targetRowIndex);
      }

      // Write mapped columns
      for (const colMap of mapping.columns) {
        const colIndex = XLSX.utils.decode_col(colMap.excelColumn);
        const cellAddress = XLSX.utils.encode_cell({ r: targetRowIndex, c: colIndex });
        const val = record[colMap.supabaseColumn];

        if (val !== undefined && val !== null) {
          let cellType: 's' | 'n' | 'b' = 's';
          let finalVal: any = val;

          if (colMap.dataType === 'integer' || colMap.dataType === 'decimal') {
            const num = Number(val);
            if (!isNaN(num)) {
              cellType = 'n';
              finalVal = num;
            }
          } else if (colMap.dataType === 'boolean') {
            cellType = 'b';
            finalVal = Boolean(val);
          } else {
            finalVal = String(val);
          }

          ws[cellAddress] = { t: cellType, v: finalVal };
        }
      }
    }

    // Expand worksheet reference range if new rows were appended
    if (nextAvailableRow - 1 > range.e.r) {
      range.e.r = nextAvailableRow - 1;
      ws['!ref'] = XLSX.utils.encode_range(range);
    }

    return { updatedCount, newRowsCount };
  }
}
