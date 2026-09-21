export type NavigationTab = 
  | 'dashboard'
  | 'nextcloud'
  | 'supabase'
  | 'files'
  | 'analyzer'
  | 'mappings'
  | 'import'
  | 'history'
  | 'errors'
  | 'logs'
  | 'settings';

export type SyncInterval = '5m' | '15m' | '30m' | '1h' | 'daily';

export type ImportStatus = 'Pending' | 'Processing' | 'Success' | 'Partial Success' | 'Failed' | 'Skipped';

export type DataType = 'text' | 'integer' | 'decimal' | 'boolean' | 'date' | 'timestamp' | 'json';

export type TransformationType = 
  | 'none'
  | 'trim'
  | 'uppercase'
  | 'lowercase'
  | 'parse_date'
  | 'parse_number'
  | 'yes_no_to_boolean'
  | 'pa_to_status'
  | 'normalize_phone'
  | 'normalize_id';

export interface NextcloudConfig {
  url: string;
  webdavUrl: string;
  username: string;
  appPassword: string;
  sourceFolder: string;
  isConnected: boolean;
  lastChecked?: string;
  statusMessage?: string;
}

export interface SupabaseConfig {
  url: string;
  anonKey: string;
  serviceRoleKey: string;
  isConnected: boolean;
  lastChecked?: string;
  statusMessage?: string;
}

export interface SyncSettings {
  syncInterval: SyncInterval;
  autoSyncEnabled: boolean;
  backupToStorage: boolean;
  storageBucket: string;
  workerUrl: string;
  workerStatus: 'healthy' | 'idle' | 'syncing' | 'error';
  lastSyncAt?: string;
  nextSyncAt?: string;
}

export interface NextcloudFile {
  id: string;
  filename: string;
  path: string;
  fileSize: number;
  fileSizeFormatted: string;
  lastModified: string;
  fileHash: string; // SHA-256
  status: 'New' | 'Unchanged' | 'Modified' | 'Synced' | 'Error';
  lastProcessedAt?: string;
  worksheetsCount?: number;
}

export interface MergedRange {
  range: string; // e.g. "A1:H1"
  startCol: string;
  endCol: string;
  startRow: number;
  endRow: number;
  value: string;
  type: 'title' | 'section_heading' | 'data_span';
}

export interface SheetHeader {
  colLetter: string;
  colIndex: number;
  name: string;
  sampleValues: string[];
}

export interface SheetAnalysis {
  sheetName: string;
  totalRows: number;
  totalColumns: number;
  usedRange: string;
  mergedRanges: MergedRange[];
  detectedHeaderRow: number;
  detectedDataStartRow: number;
  candidateHeaderRows: { row: number; headers: string[]; confidence: number }[];
  titleRows: { row: number; text: string }[];
  sectionHeadings: { row: number; text: string; range: string }[];
  emptyRowsCount: number;
  repeatedHeadersCount: number;
  headers: SheetHeader[];
  sampleRows: { rowNumber: number; data: Record<string, any> }[];
}

export interface WorkbookAnalysis {
  filename: string;
  fileSize: number;
  fileSizeFormatted: string;
  totalWorksheets: number;
  worksheets: SheetAnalysis[];
  analyzedAt: string;
  fileHash: string;
}

export interface ColumnMapping {
  id: string;
  excelColumn: string; // e.g. "A"
  excelHeader: string; // e.g. "Student ID"
  supabaseColumn: string; // e.g. "student_number"
  dataType: DataType;
  required: boolean;
  uniqueKey: boolean;
  defaultValue?: string;
  transformation: TransformationType;
  validationRegex?: string;
}

export interface WorksheetMapping {
  id: string;
  workbookName: string;
  worksheetName: string;
  supabaseTable: string;
  headerRow: number;
  dataStartRow: number;
  dataEndRow?: number;
  sectionHeadingTargetCol?: string; // e.g. 'class' for A3:H3 CLASS 10A
  enabled: boolean;
  columns: ColumnMapping[];
}

export interface WorkbookMapping {
  id: string;
  workbookName: string;
  nextcloudPath: string;
  enabled: boolean;
  worksheetMappings: WorksheetMapping[];
}

export interface RowValidationError {
  worksheetName: string;
  rowNumber: number;
  excelColumn: string;
  columnName: string;
  rawValue: string;
  errorMessage: string;
  errorType: 'missing_required' | 'type_mismatch' | 'invalid_date' | 'duplicate' | 'foreign_key' | 'validation';
}

export interface DryRunResult {
  filename: string;
  timestamp: string;
  totalWorksheets: number;
  totalRows: number;
  validRows: number;
  proposedInserts: number;
  proposedUpdates: number;
  failedRows: number;
  sheetSummaries: {
    sheetName: string;
    targetTable: string;
    totalRows: number;
    validRows: number;
    inserts: number;
    updates: number;
    fails: number;
  }[];
  errors: RowValidationError[];
  sampleTransformedRecords: {
    sheetName: string;
    targetTable: string;
    action: 'INSERT' | 'UPDATE' | 'FAIL';
    record: Record<string, any>;
    rowNumber: number;
  }[];
}

export interface ImportLog {
  id: string;
  filename: string;
  filePath: string;
  fileHash: string;
  status: ImportStatus;
  isDryRun: boolean;
  numberOfWorksheets: number;
  rowsProcessed: number;
  rowsInserted: number;
  rowsUpdated: number;
  rowsFailed: number;
  startedAt: string;
  completedAt?: string;
  durationMs?: number;
  errorSummary?: string;
  details?: Record<string, any>;
}

export interface ImportAuditError extends RowValidationError {
  id: string;
  importLogId: string;
  filename: string;
  createdAt: string;
}

export interface LogMessage {
  id: string;
  timestamp: string;
  level: 'info' | 'warn' | 'error' | 'success';
  component: 'WebDAV' | 'Worker' | 'Parser' | 'Supabase' | 'MappingEngine';
  message: string;
}
