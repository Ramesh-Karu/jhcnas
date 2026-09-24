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

export class SchemaGenerator {
  static sanitizeIdentifier(raw: string): string {
    if (!raw) return 'col';
    let s = raw.trim().toLowerCase();
    // Replace non-alphanumeric with underscore
    s = s.replace(/[^a-z0-9_]/g, '_');
    // Collapse multiple underscores
    s = s.replace(/_+/g, '_');
    // Remove leading/trailing underscores
    s = s.replace(/^_+|_+$/g, '');
    // Avoid starting with number
    if (/^[0-9]/.test(s)) {
      s = `col_${s}`;
    }
    // Reserved keywords in Postgres
    const reserved = ['user', 'order', 'group', 'table', 'select', 'where', 'limit', 'offset', 'primary', 'check', 'index', 'column', 'values', 'database'];
    if (reserved.includes(s)) {
      s = `${s}_val`;
    }
    return s || 'column_val';
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
   * 1. 'SEPARATE_TABLES' (1:1 per sheet)
   * 2. 'UNIFIED_TABLE' (all sheets into 1 common schema)
   * 3. 'CUSTOM_GROUPING' (user chooses which sheets merge into which tables)
   */
  static generateSchemas(
    analysis: WorkbookAnalysis,
    mode: MultiSheetConsolidationMode,
    sheetToTableMap?: Record<string, string>,
    supabaseTables?: SupabaseTableInfo[]
  ): TableSchemaPlan[] {
    const plans: TableSchemaPlan[] = [];

    if (!analysis || !analysis.worksheets || analysis.worksheets.length === 0) {
      return plans;
    }

    if (mode === 'UNIFIED_TABLE') {
      // 1 Common Table for ALL worksheets
      const defaultTableName = this.sanitizeIdentifier(
        analysis.filename.replace(/\.[^/.]+$/, '').toLowerCase()
      ) || 'excel_data';

      // Let user override common table name via sheetToTableMap['__unified__']
      const targetTable = sheetToTableMap?.['__unified__'] || defaultTableName;

      const columnMap = new Map<string, TableSchemaColumn>();
      const sourceSheetNames = analysis.worksheets.map(w => w.sheetName);

      // Collect union of columns across all sheets
      for (const ws of analysis.worksheets) {
        for (const h of ws.headers) {
          const colName = this.sanitizeIdentifier(h.name);
          const existing = columnMap.get(colName);
          const inferredType = h.inferredType || 'text';

          if (!existing) {
            columnMap.set(colName, {
              name: colName,
              originalHeaders: [h.name],
              dataType: inferredType,
              sqlType: this.dataTypeToPostgresType(inferredType),
              isPrimary: h.isCandidateKey || colName === 'id',
              required: h.isCandidateKey || false,
              sampleValues: [...h.sampleValues]
            });
          } else {
            // Merge type & samples
            const combinedType = this.combineDataTypes(existing.dataType, inferredType);
            existing.dataType = combinedType;
            existing.sqlType = this.dataTypeToPostgresType(combinedType);
            if (!existing.originalHeaders.includes(h.name)) {
              existing.originalHeaders.push(h.name);
            }
            if (h.isCandidateKey) {
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
      // 1 Table Per Sheet (guaranteeing unique table names)
      const usedTableNames = new Set<string>();
      for (const ws of analysis.worksheets) {
        const customTable = sheetToTableMap?.[ws.sheetName];
        let targetTable = customTable
          ? this.sanitizeIdentifier(customTable)
          : this.sanitizeIdentifier(ws.sheetName);

        if (usedTableNames.has(targetTable)) {
          let suffix = 2;
          while (usedTableNames.has(`${targetTable}_${suffix}`)) {
            suffix++;
          }
          targetTable = `${targetTable}_${suffix}`;
        }
        usedTableNames.add(targetTable);

        const columns: TableSchemaColumn[] = ws.headers.map(h => {
          const colName = this.sanitizeIdentifier(h.name);
          const inferredType = h.inferredType || 'text';
          return {
            name: colName,
            originalHeaders: [h.name],
            dataType: inferredType,
            sqlType: this.dataTypeToPostgresType(inferredType),
            isPrimary: h.isCandidateKey || colName === 'id',
            required: h.isCandidateKey || false,
            sampleValues: [...h.sampleValues]
          };
        });

        const plan = this.buildTableSchemaPlan(targetTable, [ws.sheetName], columns, supabaseTables);
        plans.push(plan);
      }

    } else {
      // CUSTOM_GROUPING (Hybrid)
      // Group sheets by target table
      const tableToSheets = new Map<string, SheetAnalysis[]>();

      for (const ws of analysis.worksheets) {
        const assignedTable = sheetToTableMap?.[ws.sheetName]
          ? this.sanitizeIdentifier(sheetToTableMap[ws.sheetName])
          : this.sanitizeIdentifier(ws.sheetName);

        if (!tableToSheets.has(assignedTable)) {
          tableToSheets.set(assignedTable, []);
        }
        tableToSheets.get(assignedTable)!.push(ws);
      }

      for (const [targetTable, sheets] of tableToSheets.entries()) {
        const columnMap = new Map<string, TableSchemaColumn>();
        const sourceSheetNames = sheets.map(s => s.sheetName);

        for (const ws of sheets) {
          for (const h of ws.headers) {
            const colName = this.sanitizeIdentifier(h.name);
            const existing = columnMap.get(colName);
            const inferredType = h.inferredType || 'text';

            if (!existing) {
              columnMap.set(colName, {
                name: colName,
                originalHeaders: [h.name],
                dataType: inferredType,
                sqlType: this.dataTypeToPostgresType(inferredType),
                isPrimary: h.isCandidateKey || colName === 'id',
                required: h.isCandidateKey || false,
                sampleValues: [...h.sampleValues]
              });
            } else {
              const combinedType = this.combineDataTypes(existing.dataType, inferredType);
              existing.dataType = combinedType;
              existing.sqlType = this.dataTypeToPostgresType(combinedType);
              if (!existing.originalHeaders.includes(h.name)) {
                existing.originalHeaders.push(h.name);
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

    // Generate CREATE TABLE SQL
    const colDefs: string[] = [];
    // Ensure primary id column exists
    const hasExplicitId = columns.some(c => c.name === 'id');
    if (!hasExplicitId) {
      colDefs.push('  id uuid PRIMARY KEY DEFAULT gen_random_uuid()');
    }

    const uniqueCols: string[] = [];

    for (const c of columns) {
      if (c.name === 'id') {
        colDefs.push(`  id ${c.sqlType} PRIMARY KEY DEFAULT gen_random_uuid()`);
      } else {
        // Keep columns nullable so that empty cells in Excel rows are stored seamlessly as NULL
        colDefs.push(`  ${c.name} ${c.sqlType}`);
        if (c.isPrimary) {
          uniqueCols.push(c.name);
        }
      }
    }

    colDefs.push('  created_at timestamptz DEFAULT now()');
    colDefs.push('  updated_at timestamptz DEFAULT now()');

    const createTableSql = `-- 1. Create table public.${tableName}\nCREATE TABLE IF NOT EXISTS public.${tableName} (\n${colDefs.join(',\n')}\n);`;

    // Generate ALTER TABLE SQL for missing columns
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

    // Generate Indexes SQL
    const indexStatements: string[] = [];
    for (const u of uniqueCols) {
      indexStatements.push(`CREATE UNIQUE INDEX IF NOT EXISTS idx_${tableName}_${u} ON public.${tableName} (${u});`);
    }
    const indexesSql = indexStatements.join('\n');

    // Complete executable script including RLS policies
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
    sheetToTableMap?: Record<string, string>
  ): WorksheetMapping[] {
    const mappings: WorksheetMapping[] = [];

    for (const ws of analysis.worksheets) {
      // Determine which plan this worksheet belongs to
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

      const targetTableName = targetPlan ? targetPlan.tableName : this.sanitizeIdentifier(ws.sheetName);

      // Create column mappings for this sheet
      const columnMappings: ColumnMapping[] = ws.headers.map((h, idx) => {
        const cleanCol = this.sanitizeIdentifier(h.name);
        const planCol = targetPlan?.columns.find(c => c.name === cleanCol);
        const dataType = planCol ? planCol.dataType : (h.inferredType || 'text');
        const isKey = planCol ? planCol.isPrimary : (h.isCandidateKey || false);

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

      // If mode is UNIFIED and there's a _sheet_source column in plan, add it as default mapping if possible
      if (mode === 'UNIFIED_TABLE' && targetPlan?.columns.some(c => c.name === '_sheet_source')) {
        // Handled automatically via sheet origin
      }

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
