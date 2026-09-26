import {
  WorkbookAnalysis,
  SheetAnalysis,
  MultiSheetConsolidationMode,
  TableSchemaPlan,
  TableSchemaColumn,
  SupabaseTableInfo,
  WorksheetMapping,
  ColumnMapping,
  DataType,
  TransformationType
} from '../types';
import { smartSanitizeIdentifier, translateTamilHeader, containsTamil } from './tamilTranslator';

export class SchemaGenerator {
  static sanitizeIdentifier(raw: string): string {
    return smartSanitizeIdentifier(raw, 'col');
  }

  /**
   * Intelligently maps workbook archetypes and filenames to canonical destination tables
   * so tables are NOT changed erratically or fragmented across arbitrary sheet names.
   */
  static detectCanonicalTargetTable(
    filename: string,
    sheetNames: string[] = [],
    archetype?: string,
    supabaseTables?: SupabaseTableInfo[]
  ): string {
    const norm = (filename || '').toLowerCase().trim();

    // Check if live Supabase tables already exist that match
    if (supabaseTables && supabaseTables.length > 0) {
      if ((norm.includes('stu') || norm.includes('student') || /\b20\d\d\b/.test(norm)) && supabaseTables.some(t => t.name === 'students')) {
        return 'students';
      }
      if ((norm.includes('timetable') || norm.includes('classwise')) && supabaseTables.some(t => t.name === 'timetable_master' || t.name === 'class_timetables')) {
        return supabaseTables.find(t => t.name === 'timetable_master' || t.name === 'class_timetables')!.name;
      }
      if (norm.includes('donation') && supabaseTables.some(t => t.name === 'jhc_donations' || t.name === 'donation_details')) {
        return supabaseTables.find(t => t.name === 'jhc_donations' || t.name === 'donation_details')!.name;
      }
      if (norm.includes('teacher') && supabaseTables.some(t => t.name === 'subject_teacher_allocations' || t.name === 'teacher_allocations')) {
        return supabaseTables.find(t => t.name === 'subject_teacher_allocations' || t.name === 'teacher_allocations')!.name;
      }
      if (norm.includes('inventory') && supabaseTables.some(t => t.name === 'college_inventory' || t.name === 'jhc_inventory')) {
        return supabaseTables.find(t => t.name === 'college_inventory' || t.name === 'jhc_inventory')!.name;
      }
    }

    if (archetype === 'TIMETABLE_MATRIX' || norm.includes('timetable') || norm.includes('time_table') || norm.includes('classwise')) {
      return 'timetable_master';
    }
    if (archetype === 'MULTI_SHEET_LEDGER' || norm.includes('donation') || norm.includes('contribution') || norm.includes('ledger')) {
      return 'jhc_donations';
    }
    if (archetype === 'PIVOT_ALLOCATION_MATRIX' || norm.includes('teacher') || norm.includes('subjectteacher') || norm.includes('allocation')) {
      return 'subject_teacher_allocations';
    }
    if (norm.includes('inventory') || norm.includes('assets') || norm.includes('equipment') || norm.includes('stock')) {
      return 'college_inventory';
    }

    // Default canonical destination for student workbooks (e.g. 2026_stu, 2032_stu_new, students_complex)
    if (norm.includes('stu') || norm.includes('student') || norm.includes('admission') || /\b20\d\d\b/.test(norm)) {
      return 'students';
    }

    const combinedSheets = sheetNames.join(' ').toLowerCase();
    if (combinedSheets.includes('grade') || combinedSheets.includes('6a') || combinedSheets.includes('6b') || combinedSheets.includes('7a')) {
      return 'students';
    }

    return 'students';
  }

  static dataTypeToPostgresType(type: DataType): string {
    switch (type) {
      case 'integer':
        return 'integer';
      case 'decimal':
        return 'numeric(12, 2)';
      case 'date':
        return 'date';
      case 'timestamp':
        return 'timestamptz';
      case 'boolean':
        return 'boolean';
      case 'json':
        return 'jsonb';
      case 'text':
      default:
        return 'text';
    }
  }

  static dataTypeToTransformation(type: DataType): TransformationType {
    switch (type) {
      case 'date':
        return 'parse_date';
      case 'integer':
      case 'decimal':
        return 'parse_number';
      case 'boolean':
        return 'yes_no_to_boolean';
      default:
        return 'trim';
    }
  }

  // Merge multiple data types (e.g. integer + decimal -> decimal; text + anything -> text)
  static combineDataTypes(a: DataType, b: DataType): DataType {
    if (a === b) return a;
    if (a === 'text' || b === 'text') return 'text';
    if ((a === 'integer' && b === 'decimal') || (a === 'decimal' && b === 'integer')) return 'decimal';
    if ((a === 'date' && b === 'timestamp') || (a === 'timestamp' && b === 'date')) return 'timestamp';
    return 'text';
  }

  /**
   * Generates schemas for the entire workbook based on the chosen consolidation mode:
   * 1. 'UNIFIED_TABLE' (all sheets into 1 common schema, e.g. 'students')
   * 2. 'SEPARATE_TABLES' (1:1 per sheet)
   * 3. 'CUSTOM_GROUPING' (user chooses which sheets merge into which tables)
   */
  static generateSchemas(
    analysis: WorkbookAnalysis,
    mode: MultiSheetConsolidationMode,
    sheetToTableMap?: Record<string, string>,
    supabaseTables?: SupabaseTableInfo[],
    customPrimaryKeys?: Record<string, string>
  ): TableSchemaPlan[] {
    const plans: TableSchemaPlan[] = [];

    if (!analysis || !analysis.worksheets || analysis.worksheets.length === 0) {
      return plans;
    }

    const canonicalDefaultTable = this.detectCanonicalTargetTable(
      analysis.filename,
      analysis.worksheets.map(w => w.sheetName),
      analysis.detectedArchetype,
      supabaseTables
    );

    if (mode === 'UNIFIED_TABLE') {
      // 1 Common Destination Table for ALL worksheets
      const targetTable = sheetToTableMap?.['__unified__'] || canonicalDefaultTable;
      const userSelectedPk = customPrimaryKeys?.[targetTable] ?? customPrimaryKeys?.['__unified__'];

      const columnMap = new Map<string, TableSchemaColumn>();
      const sourceSheetNames = analysis.worksheets.map(w => w.sheetName);

      // Collect union of columns across all sheets
      for (const ws of analysis.worksheets) {
        for (const h of ws.headers) {
          let colName = this.sanitizeIdentifier(h.name);
          const inferredType = h.inferredType || 'text';

          // If target is students, align common header variations
          if (targetTable === 'students') {
            const rawLower = h.name.toLowerCase().trim();
            if (rawLower.includes('admission') || rawLower.includes('adm_no') || rawLower.includes('அனுமதி')) {
              colName = 'admission_no';
            } else if (rawLower.includes('full name') || rawLower === 'name' || rawLower.includes('பெயர்')) {
              colName = containsTamil(h.name) ? 'tamil_name' : 'full_name';
            } else if (rawLower.includes('initial') || rawLower.includes('name with')) {
              colName = 'name_with_initials';
            } else if (rawLower.includes('dob') || rawLower.includes('birth') || rawLower.includes('பிறந்த')) {
              colName = 'date_of_birth';
            } else if (rawLower.includes('gender') || rawLower.includes('sex') || rawLower.includes('பாலினம்')) {
              colName = 'gender';
            } else if (rawLower.includes('religion') || rawLower.includes('மதம்')) {
              colName = 'religion';
            } else if (rawLower.includes('house') || rawLower.includes('இல்லம்')) {
              colName = 'house';
            } else if (rawLower.includes('index') || rawLower.includes('சுட்டெண்')) {
              colName = 'index_number';
            } else if (rawLower.includes('phone') || rawLower.includes('contact') || rawLower.includes('mobile') || rawLower.includes('தொலைபேசி')) {
              colName = 'phone_number';
            } else if (rawLower.includes('guardian') || rawLower.includes('parent') || rawLower.includes('பெற்றோர்')) {
              colName = 'guardian_name';
            } else if (rawLower.includes('address') || rawLower.includes('முகவரி')) {
              colName = 'address';
            }
          }

          const existing = columnMap.get(colName);
          const isPrimary = userSelectedPk !== undefined
            ? (userSelectedPk !== '__NONE__' && (userSelectedPk.toLowerCase() === colName.toLowerCase() || userSelectedPk.toLowerCase() === h.name.toLowerCase()))
            : (colName === 'admission_no' || h.isCandidateKey || false);

          if (!existing) {
            columnMap.set(colName, {
              name: colName,
              originalHeaders: [h.name],
              dataType: inferredType,
              sqlType: this.dataTypeToPostgresType(inferredType),
              isPrimary,
              required: isPrimary || false,
              sampleValues: [...h.sampleValues]
            });
          } else {
            const combinedType = this.combineDataTypes(existing.dataType, inferredType);
            existing.dataType = combinedType;
            existing.sqlType = this.dataTypeToPostgresType(combinedType);
            if (!existing.originalHeaders.includes(h.name)) {
              existing.originalHeaders.push(h.name);
            }
            if (isPrimary) {
              existing.isPrimary = true;
            }
            for (const sv of h.sampleValues) {
              if (existing.sampleValues.length < 8 && !existing.sampleValues.includes(sv)) {
                existing.sampleValues.push(sv);
              }
            }
          }
        }
      }

      // Add sheet source tracker column so rows preserve source tab
      if (analysis.worksheets.length > 1 && !columnMap.has('_sheet_source')) {
        columnMap.set('_sheet_source', {
          name: '_sheet_source',
          originalHeaders: ['[Workbook Tab Origin]'],
          dataType: 'text',
          sqlType: 'text',
          isPrimary: false,
          required: false,
          sampleValues: analysis.worksheets.map(w => w.sheetName).slice(0, 5)
        });
      }

      const plan = this.buildTableSchemaPlan(targetTable, sourceSheetNames, Array.from(columnMap.values()), supabaseTables);
      plans.push(plan);

    } else if (mode === 'SEPARATE_TABLES') {
      // 1 Table Per Sheet
      const usedTableNames = new Set<string>();
      for (const ws of analysis.worksheets) {
        const customTable = sheetToTableMap?.[ws.sheetName];
        let targetTable = customTable
          ? this.sanitizeIdentifier(customTable)
          : canonicalDefaultTable;

        if (usedTableNames.has(targetTable)) {
          let suffix = 2;
          while (usedTableNames.has(`${targetTable}_${suffix}`)) {
            suffix++;
          }
          targetTable = `${targetTable}_${suffix}`;
        }
        usedTableNames.add(targetTable);

        const userSelectedPk = customPrimaryKeys?.[targetTable] ?? customPrimaryKeys?.[ws.sheetName];

        const usedColNames = new Set<string>();
        const columns: TableSchemaColumn[] = ws.headers.map(h => {
          let colName = this.sanitizeIdentifier(h.name);
          if (usedColNames.has(colName)) {
            let suffix = 2;
            while (usedColNames.has(`${colName}_${suffix}`)) {
              suffix++;
            }
            colName = `${colName}_${suffix}`;
          }
          usedColNames.add(colName);

          const isPrimary = userSelectedPk !== undefined
            ? (userSelectedPk !== '__NONE__' && (userSelectedPk.toLowerCase() === colName.toLowerCase() || userSelectedPk.toLowerCase() === h.name.toLowerCase()))
            : (colName === 'admission_no' || h.isCandidateKey || false);

          const inferredType = h.inferredType || 'text';
          return {
            name: colName,
            originalHeaders: [h.name],
            dataType: inferredType,
            sqlType: this.dataTypeToPostgresType(inferredType),
            isPrimary,
            required: isPrimary || false,
            sampleValues: [...h.sampleValues]
          };
        });

        const plan = this.buildTableSchemaPlan(targetTable, [ws.sheetName], columns, supabaseTables);
        plans.push(plan);
      }

    } else {
      // CUSTOM_GROUPING
      const tableToSheets = new Map<string, SheetAnalysis[]>();

      for (const ws of analysis.worksheets) {
        const assignedTable = sheetToTableMap?.[ws.sheetName]
          ? this.sanitizeIdentifier(sheetToTableMap[ws.sheetName])
          : canonicalDefaultTable;

        if (!tableToSheets.has(assignedTable)) {
          tableToSheets.set(assignedTable, []);
        }
        tableToSheets.get(assignedTable)!.push(ws);
      }

      for (const [targetTable, sheets] of tableToSheets.entries()) {
        const columnMap = new Map<string, TableSchemaColumn>();
        const sourceSheetNames = sheets.map(s => s.sheetName);
        const userSelectedPk = customPrimaryKeys?.[targetTable];

        for (const ws of sheets) {
          for (const h of ws.headers) {
            const colName = this.sanitizeIdentifier(h.name);
            const existing = columnMap.get(colName);
            const inferredType = h.inferredType || 'text';

            const isPrimary = userSelectedPk !== undefined
              ? (userSelectedPk !== '__NONE__' && (userSelectedPk.toLowerCase() === colName.toLowerCase() || userSelectedPk.toLowerCase() === h.name.toLowerCase()))
              : (colName === 'admission_no' || h.isCandidateKey || false);

            if (!existing) {
              columnMap.set(colName, {
                name: colName,
                originalHeaders: [h.name],
                dataType: inferredType,
                sqlType: this.dataTypeToPostgresType(inferredType),
                isPrimary,
                required: isPrimary || false,
                sampleValues: [...h.sampleValues]
              });
            } else {
              const combinedType = this.combineDataTypes(existing.dataType, inferredType);
              existing.dataType = combinedType;
              existing.sqlType = this.dataTypeToPostgresType(combinedType);
              if (!existing.originalHeaders.includes(h.name)) {
                existing.originalHeaders.push(h.name);
              }
              if (isPrimary) {
                existing.isPrimary = true;
              }
              for (const sv of h.sampleValues) {
                if (existing.sampleValues.length < 8 && !existing.sampleValues.includes(sv)) {
                  existing.sampleValues.push(sv);
                }
              }
            }
          }
        }

        if (sheets.length > 1 && !columnMap.has('_sheet_source')) {
          columnMap.set('_sheet_source', {
            name: '_sheet_source',
            originalHeaders: ['[Workbook Tab Origin]'],
            dataType: 'text',
            sqlType: 'text',
            isPrimary: false,
            required: false,
            sampleValues: sheets.map(w => w.sheetName).slice(0, 5)
          });
        }

        const plan = this.buildTableSchemaPlan(targetTable, sourceSheetNames, Array.from(columnMap.values()), supabaseTables);
        plans.push(plan);
      }
    }

    return plans;
  }

  private static buildTableSchemaPlan(
    tableName: string,
    sourceSheetNames: string[],
    columns: TableSchemaColumn[],
    supabaseTables?: SupabaseTableInfo[]
  ): TableSchemaPlan {
    // Check against live Supabase
    const liveTable = supabaseTables?.find(t => t.name.toLowerCase() === tableName.toLowerCase());
    const liveColNames = new Set(liveTable ? liveTable.columns.map(c => c.name.toLowerCase()) : []);

    const alreadyInSupabase: string[] = [];
    const missingInSupabase: string[] = [];

    for (const c of columns) {
      if (liveColNames.has(c.name.toLowerCase())) {
        c.matchesSupabaseColumn = true;
        const liveCol = liveTable?.columns.find(x => x.name.toLowerCase() === c.name.toLowerCase());
        c.supabaseMatchDetails = `Matches '${liveCol?.name}' (${liveCol?.type})`;
        alreadyInSupabase.push(c.name);
      } else {
        c.matchesSupabaseColumn = false;
        c.supabaseMatchDetails = liveTable ? 'Column missing in Supabase' : 'New Table';
        missingInSupabase.push(c.name);
      }
    }

    let diffStatus: TableSchemaPlan['diffStatus'] = 'NEW_TABLE';
    if (liveTable) {
      if (missingInSupabase.length === 0) {
        diffStatus = 'EXACT_MATCH';
      } else {
        diffStatus = 'NEEDS_ALTER';
      }
    }

    const colDefs: string[] = [];
    const uniqueCols: string[] = [];

    for (const c of columns) {
      if (c.isPrimary) {
        colDefs.push(`  "${c.name}" ${c.sqlType} PRIMARY KEY`);
        uniqueCols.push(c.name);
      } else if (c.required) {
        colDefs.push(`  "${c.name}" ${c.sqlType} NOT NULL`);
      } else {
        colDefs.push(`  "${c.name}" ${c.sqlType}`);
      }
    }

    colDefs.push('  created_at timestamptz DEFAULT now()');
    colDefs.push('  updated_at timestamptz DEFAULT now()');

    const createTableSql = `-- 1. Create table public.${tableName}\nCREATE TABLE IF NOT EXISTS public.${tableName} (\n${colDefs.join(',\n')}\n);`;

    const alterStatements: string[] = [];
    if (liveTable && missingInSupabase.length > 0) {
      for (const colName of missingInSupabase) {
        const col = columns.find(c => c.name === colName);
        if (col) {
          alterStatements.push(`ALTER TABLE public.${tableName} ADD COLUMN IF NOT EXISTS ${col.name} ${col.sqlType};`);
        }
      }
    }
    const alterTableSql = alterStatements.length > 0
      ? `-- Missing columns to add to existing public.${tableName}\n${alterStatements.join('\n')}`
      : `-- Table public.${tableName} already contains all columns!`;

    const indexStatements: string[] = [];
    for (const u of uniqueCols) {
      indexStatements.push(`CREATE UNIQUE INDEX IF NOT EXISTS idx_${tableName}_${u} ON public.${tableName} (${u});`);
    }
    const indexesSql = indexStatements.join('\n');

    const completeSql = [
      `-- ==========================================`,
      `-- SUPABASE POSTGRESQL SCHEMA FOR: public.${tableName}`,
      `-- Generated automatically from Excel: [${sourceSheetNames.join(', ')}]`,
      `-- ==========================================`,
      ``,
      createTableSql,
      ``,
      alterStatements.length > 0 ? alterTableSql + '\n' : '',
      indexesSql ? `-- Unique Merge Indexes\n${indexesSql}\n` : '',
      `-- Enable Row Level Security (RLS)`,
      `ALTER TABLE public.${tableName} ENABLE ROW LEVEL SECURITY;`,
      ``,
      `-- Permissive RLS Policy for Nextcloud Sync & App Access`,
      `DO $$`,
      `BEGIN`,
      `  IF NOT EXISTS (`,
      `    SELECT 1 FROM pg_policies WHERE tablename = '${tableName}' AND policyname = 'Allow service and auth sync access'`,
      `  ) THEN`,
      `    CREATE POLICY "Allow service and auth sync access" ON public.${tableName}`,
      `      FOR ALL USING (true);`,
      `  END IF;`,
      `END $$;`,
      ``,
      `-- Grant permissions`,
      `GRANT ALL ON public.${tableName} TO anon, authenticated, service_role;`
    ].filter(Boolean).join('\n');

    return {
      tableName,
      sourceSheetNames,
      columns,
      createTableSql,
      alterTableSql,
      indexesSql,
      completeSql,
      diffStatus,
      missingInSupabaseColumns: missingInSupabase,
      alreadyInSupabaseColumns: alreadyInSupabase
    };
  }

  /**
   * Converts generated schema plans into fully populated WorksheetMapping[]
   * so that the application mappings are immediately synchronized with the schema!
   */
  static generateMappingsFromPlans(
    analysis: WorkbookAnalysis,
    plans: TableSchemaPlan[],
    mode: MultiSheetConsolidationMode,
    sheetToTableMap?: Record<string, string>,
    supabaseTables?: SupabaseTableInfo[]
  ): WorksheetMapping[] {
    const mappings: WorksheetMapping[] = [];
    const canonicalTarget = this.detectCanonicalTargetTable(
      analysis.filename,
      analysis.worksheets.map(w => w.sheetName),
      analysis.detectedArchetype,
      supabaseTables
    );

    for (const ws of analysis.worksheets) {
      let targetPlan: TableSchemaPlan | undefined;

      if (mode === 'UNIFIED_TABLE') {
        targetPlan = plans[0];
      } else if (mode === 'SEPARATE_TABLES') {
        const cleanName = this.sanitizeIdentifier(ws.sheetName);
        targetPlan = plans.find(p => p.tableName === cleanName || p.sourceSheetNames.includes(ws.sheetName));
      } else {
        const assigned = sheetToTableMap?.[ws.sheetName];
        if (assigned) {
          const cleanAssigned = this.sanitizeIdentifier(assigned);
          targetPlan = plans.find(p => p.tableName === cleanAssigned);
        }
        if (!targetPlan) {
          targetPlan = plans.find(p => p.sourceSheetNames.includes(ws.sheetName));
        }
      }

      const targetTableName = targetPlan ? targetPlan.tableName : (sheetToTableMap?.[ws.sheetName] || canonicalTarget);

      // Create column mappings for this sheet with unique database column names
      const usedColNames = new Set<string>();
      const columnMappings: ColumnMapping[] = ws.headers.map((h, idx) => {
        let cleanCol = this.sanitizeIdentifier(h.name);

        if (targetTableName === 'students') {
          const rawLower = h.name.toLowerCase().trim();
          if (rawLower.includes('admission') || rawLower.includes('adm_no') || rawLower.includes('அனுமதி')) {
            cleanCol = 'admission_no';
          } else if (rawLower.includes('full name') || rawLower === 'name' || rawLower.includes('பெயர்')) {
            cleanCol = containsTamil(h.name) ? 'tamil_name' : 'full_name';
          } else if (rawLower.includes('initial') || rawLower.includes('name with')) {
            cleanCol = 'name_with_initials';
          } else if (rawLower.includes('dob') || rawLower.includes('birth') || rawLower.includes('பிறந்த')) {
            cleanCol = 'date_of_birth';
          } else if (rawLower.includes('gender') || rawLower.includes('sex') || rawLower.includes('பாலினம்')) {
            cleanCol = 'gender';
          } else if (rawLower.includes('religion') || rawLower.includes('மதம்')) {
            cleanCol = 'religion';
          } else if (rawLower.includes('house') || rawLower.includes('இல்லம்')) {
            cleanCol = 'house';
          } else if (rawLower.includes('index') || rawLower.includes('சுட்டெண்')) {
            cleanCol = 'index_number';
          } else if (rawLower.includes('phone') || rawLower.includes('contact') || rawLower.includes('mobile') || rawLower.includes('தொலைபேசி')) {
            cleanCol = 'phone_number';
          } else if (rawLower.includes('guardian') || rawLower.includes('parent') || rawLower.includes('பெற்றோர்')) {
            cleanCol = 'guardian_name';
          } else if (rawLower.includes('address') || rawLower.includes('முகவரி')) {
            cleanCol = 'address';
          }
        }

        if (usedColNames.has(cleanCol)) {
          let suffix = 2;
          while (usedColNames.has(`${cleanCol}_${suffix}`)) {
            suffix++;
          }
          cleanCol = `${cleanCol}_${suffix}`;
        }
        usedColNames.add(cleanCol);

        const planCol = targetPlan?.columns.find(c => c.name === cleanCol);
        const dataType = planCol ? planCol.dataType : (h.inferredType || 'text');
        const isKey = planCol ? planCol.isPrimary : (cleanCol === 'admission_no' || h.isCandidateKey || false);

        return {
          id: `cm-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
          excelColumn: h.colLetter,
          excelHeader: h.name,
          supabaseColumn: cleanCol,
          dataType,
          required: isKey,
          uniqueKey: isKey,
          transformation: this.dataTypeToTransformation(dataType)
        };
      });

      mappings.push({
        id: `wm-${Date.now()}-${Math.random().toString(36).substr(2, 6)}`,
        workbookName: analysis.filename,
        worksheetName: ws.sheetName,
        supabaseTable: targetTableName,
        headerRow: ws.detectedHeaderRow,
        dataStartRow: ws.detectedDataStartRow,
        enabled: true,
        syncPolicy: 'EXCEL_TO_DB',
        columns: columnMappings
      });
    }

    return mappings;
  }
}
