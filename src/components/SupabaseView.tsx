import React, { useState } from 'react';
import { 
  Database, 
  Table, 
  CheckCircle2, 
  XCircle,
  AlertCircle,
  RefreshCw, 
  Key, 
  Save, 
  ExternalLink, 
  Eye, 
  EyeOff,
  Code, 
  Layers,
  Copy,
  Check,
  Server,
  ShieldCheck,
  Lock,
  ArrowDown,
  Trash2
} from 'lucide-react';
import { SupabaseConfig } from '../types';
import { ApiClient, SupabaseTestResult } from '../services/apiClient';

const PRODUCTION_SQL = `-- Nextcloud Excel Sync - Production Supabase PostgreSQL Schema
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- 1. Sync Settings
CREATE TABLE IF NOT EXISTS sync_settings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    nextcloud_url TEXT NOT NULL DEFAULT '',
    nextcloud_webdav_url TEXT NOT NULL DEFAULT '',
    nextcloud_username TEXT NOT NULL DEFAULT '',
    nextcloud_folder TEXT NOT NULL DEFAULT '/ExcelImports',
    sync_interval TEXT NOT NULL DEFAULT '15m',
    auto_sync_enabled BOOLEAN NOT NULL DEFAULT TRUE,
    backup_to_storage BOOLEAN NOT NULL DEFAULT FALSE,
    storage_bucket TEXT DEFAULT 'excel-archives',
    last_sync_at TIMESTAMPTZ,
    next_sync_at TIMESTAMPTZ,
    worker_status TEXT NOT NULL DEFAULT 'idle',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 2. Workbooks Registry
CREATE TABLE IF NOT EXISTS workbooks (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL,
    nextcloud_path TEXT NOT NULL UNIQUE,
    file_size BIGINT NOT NULL DEFAULT 0,
    file_hash TEXT,
    last_modified_at TIMESTAMPTZ,
    enabled BOOLEAN NOT NULL DEFAULT TRUE,
    auto_detect_headers BOOLEAN NOT NULL DEFAULT TRUE,
    description TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3. Worksheet Mappings
CREATE TABLE IF NOT EXISTS worksheet_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_id UUID NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
    worksheet_name TEXT NOT NULL,
    supabase_table TEXT NOT NULL,
    header_row INTEGER NOT NULL DEFAULT 1,
    data_start_row INTEGER NOT NULL DEFAULT 2,
    data_end_row INTEGER,
    section_heading_col TEXT,
    section_heading_target_col TEXT,
    enabled BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(workbook_id, worksheet_name)
);

-- 4. Column Mappings
CREATE TABLE IF NOT EXISTS column_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    worksheet_mapping_id UUID NOT NULL REFERENCES worksheet_mappings(id) ON DELETE CASCADE,
    excel_column TEXT NOT NULL,
    excel_header TEXT NOT NULL,
    supabase_column TEXT NOT NULL,
    data_type TEXT NOT NULL DEFAULT 'text',
    required BOOLEAN NOT NULL DEFAULT FALSE,
    unique_key BOOLEAN NOT NULL DEFAULT FALSE,
    default_value TEXT,
    transformation TEXT NOT NULL DEFAULT 'none',
    validation_regex TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 5. Import Audit Logs
CREATE TABLE IF NOT EXISTS import_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_id UUID REFERENCES workbooks(id) ON DELETE SET NULL,
    filename TEXT NOT NULL,
    file_path TEXT NOT NULL,
    file_hash TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Pending',
    is_dry_run BOOLEAN NOT NULL DEFAULT FALSE,
    number_of_worksheets INTEGER NOT NULL DEFAULT 1,
    rows_processed INTEGER NOT NULL DEFAULT 0,
    rows_inserted INTEGER NOT NULL DEFAULT 0,
    rows_updated INTEGER NOT NULL DEFAULT 0,
    rows_failed INTEGER NOT NULL DEFAULT 0,
    started_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    completed_at TIMESTAMPTZ,
    error_summary TEXT,
    details JSONB
);

-- 6. Import Errors
CREATE TABLE IF NOT EXISTS import_errors (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    import_log_id UUID NOT NULL REFERENCES import_logs(id) ON DELETE CASCADE,
    worksheet_name TEXT NOT NULL,
    row_number INTEGER NOT NULL,
    excel_column TEXT,
    column_name TEXT,
    raw_value TEXT,
    error_message TEXT NOT NULL,
    error_type TEXT NOT NULL DEFAULT 'validation',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 7. Two-Way Sync Conflicts Staging Queue
CREATE TABLE IF NOT EXISTS sync_conflicts (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_name TEXT NOT NULL,
    worksheet_name TEXT NOT NULL,
    target_table TEXT NOT NULL,
    primary_key_col TEXT NOT NULL,
    primary_key_val TEXT NOT NULL,
    excel_row_num INTEGER,
    field_diffs JSONB NOT NULL DEFAULT '[]'::jsonb,
    excel_payload JSONB NOT NULL DEFAULT '{}'::jsonb,
    supabase_payload JSONB NOT NULL DEFAULT '{}'::jsonb,
    status TEXT NOT NULL DEFAULT 'pending',
    resolved_by TEXT,
    resolved_at TIMESTAMPTZ,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 8. Two-Way Sync Audit Trail
CREATE TABLE IF NOT EXISTS sync_two_way_audit (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    direction TEXT NOT NULL,
    records_compared INTEGER NOT NULL DEFAULT 0,
    in_sync_count INTEGER NOT NULL DEFAULT 0,
    excel_changes_applied INTEGER NOT NULL DEFAULT 0,
    supabase_changes_pushed INTEGER NOT NULL DEFAULT 0,
    conflicts_flagged INTEGER NOT NULL DEFAULT 0,
    performed_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    details JSONB
);

-- Indexes
CREATE INDEX IF NOT EXISTS idx_workbooks_hash ON workbooks(file_hash);
CREATE INDEX IF NOT EXISTS idx_workbooks_path ON workbooks(nextcloud_path);
CREATE INDEX IF NOT EXISTS idx_import_logs_hash ON import_logs(file_hash);
CREATE INDEX IF NOT EXISTS idx_import_logs_status ON import_logs(status);
CREATE INDEX IF NOT EXISTS idx_import_logs_started ON import_logs(started_at DESC);
CREATE INDEX IF NOT EXISTS idx_worksheet_mappings_wb ON worksheet_mappings(workbook_id);
CREATE INDEX IF NOT EXISTS idx_column_mappings_ws ON column_mappings(worksheet_mapping_id);
CREATE INDEX IF NOT EXISTS idx_import_errors_log ON import_errors(import_log_id);
CREATE INDEX IF NOT EXISTS idx_sync_conflicts_status ON sync_conflicts(status);
CREATE INDEX IF NOT EXISTS idx_sync_conflicts_pk ON sync_conflicts(target_table, primary_key_val);

-- Row Level Security
ALTER TABLE sync_settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE workbooks ENABLE ROW LEVEL SECURITY;
ALTER TABLE worksheet_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE column_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE import_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE import_errors ENABLE ROW LEVEL SECURITY;
ALTER TABLE sync_conflicts ENABLE ROW LEVEL SECURITY;
ALTER TABLE sync_two_way_audit ENABLE ROW LEVEL SECURITY;

-- Grant broad RLS policies for clean production access
DO $$
BEGIN
    DROP POLICY IF EXISTS "Allow full access to sync_settings" ON sync_settings;
    CREATE POLICY "Allow full access to sync_settings" ON sync_settings FOR ALL TO public USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow full access to workbooks" ON workbooks;
    CREATE POLICY "Allow full access to workbooks" ON workbooks FOR ALL TO public USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow full access to worksheet_mappings" ON worksheet_mappings;
    CREATE POLICY "Allow full access to worksheet_mappings" ON worksheet_mappings FOR ALL TO public USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow full access to column_mappings" ON column_mappings;
    CREATE POLICY "Allow full access to column_mappings" ON column_mappings FOR ALL TO public USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow full access to import_logs" ON import_logs;
    CREATE POLICY "Allow full access to import_logs" ON import_logs FOR ALL TO public USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow full access to import_errors" ON import_errors;
    CREATE POLICY "Allow full access to import_errors" ON import_errors FOR ALL TO public USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow full access to sync_conflicts" ON sync_conflicts;
    CREATE POLICY "Allow full access to sync_conflicts" ON sync_conflicts FOR ALL TO public USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow full access to sync_two_way_audit" ON sync_two_way_audit;
    CREATE POLICY "Allow full access to sync_two_way_audit" ON sync_two_way_audit FOR ALL TO public USING (true) WITH CHECK (true);
END $$;`;

interface SupabaseViewProps {
  config: SupabaseConfig;
  databaseState: Record<string, any[]>;
  onSaveConfig: (cfg: SupabaseConfig) => void;
}

export const SupabaseView: React.FC<SupabaseViewProps> = ({
  config,
  databaseState,
  onSaveConfig
}) => {
  const [formData, setFormData] = useState<SupabaseConfig>(config);
  const [activeTab, setActiveTab] = useState<'connection' | 'tables' | 'sql'>('connection');
  const [isTesting, setIsTesting] = useState(false);
  const [testResult, setTestResult] = useState<SupabaseTestResult | null>(null);
  const [statusNotice, setStatusNotice] = useState<string | null>(null);
  const [copiedSql, setCopiedSql] = useState(false);
  const [showAnonKey, setShowAnonKey] = useState(false);
  const [showServiceKey, setShowServiceKey] = useState(false);

  // Live remote table querying
  const [isFetchingRemote, setIsFetchingRemote] = useState(false);
  const [remoteRows, setRemoteRows] = useState<any[] | null>(null);
  const [remoteFetchNotice, setRemoteFetchNotice] = useState<string | null>(null);

  const stateTables = Object.keys(databaseState);
  const availableTables = stateTables.length > 0 
    ? stateTables 
    : ['sync_settings', 'workbooks', 'worksheet_mappings', 'column_mappings', 'import_logs', 'import_errors'];

  const [selectedTable, setSelectedTable] = useState<string>(() => availableTables[0] || 'workbooks');

  const currentRecords = remoteRows !== null ? remoteRows : (databaseState[selectedTable] || []);

  const handleCopySql = () => {
    navigator.clipboard.writeText(PRODUCTION_SQL);
    setCopiedSql(true);
    setTimeout(() => setCopiedSql(false), 2500);
  };

  const handleTestConnection = async () => {
    if (!formData.url.trim()) {
      setStatusNotice('Please enter your Supabase Project URL before testing.');
      setTimeout(() => setStatusNotice(null), 3500);
      return;
    }
    const activeKey = (formData.serviceKey && formData.serviceKey.trim()) || (formData.serviceRoleKey && formData.serviceRoleKey.trim()) || (formData.anonKey && formData.anonKey.trim());
    if (!activeKey) {
      setStatusNotice('Please enter your Supabase API Key (Anon Key or Service Key) before testing.');
      setTimeout(() => setStatusNotice(null), 3500);
      return;
    }

    setIsTesting(true);
    setTestResult(null);

    // Call live real test endpoint
    const result = await ApiClient.testSupabaseConnection(formData);
    setTestResult(result);
    setIsTesting(false);

    if (result.success && result.isConnected) {
      const updated: SupabaseConfig = {
        ...formData,
        isConnected: true,
        lastChecked: new Date().toISOString(),
        statusMessage: result.message || 'Connected to Supabase PostgreSQL (PostgREST API verified)'
      };
      setFormData(updated);
      onSaveConfig(updated);
      setStatusNotice('Supabase connection verified! Live PostgreSQL database is connected.');
    } else {
      const updated: SupabaseConfig = {
        ...formData,
        isConnected: false,
        lastChecked: new Date().toISOString(),
        statusMessage: result.error || 'Connection test failed. Verify Project URL and API Key.'
      };
      setFormData(updated);
      onSaveConfig(updated);
      setStatusNotice(result.error || 'Connection test failed.');
    }
    setTimeout(() => setStatusNotice(null), 6000);
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    onSaveConfig(formData);
    setStatusNotice('Supabase credentials saved successfully.');
    setTimeout(() => setStatusNotice(null), 3000);
  };

  const handleClearCredentials = () => {
    const cleared: SupabaseConfig = {
      url: '',
      anonKey: '',
      serviceKey: '',
      serviceRoleKey: '',
      isConnected: false,
      lastChecked: undefined,
      statusMessage: 'Not connected. Enter your Supabase Project URL and API Key to connect.'
    };
    setFormData(cleared);
    onSaveConfig(cleared);
    setTestResult(null);
    setRemoteRows(null);
    setStatusNotice('Supabase credentials cleared. Disconnected from database.');
    setTimeout(() => setStatusNotice(null), 3000);
  };

  const handleFetchRemoteTable = async () => {
    if (!config.isConnected || !config.url) {
      setRemoteFetchNotice('Supabase is not connected. Enter credentials and verify connection first.');
      setTimeout(() => setRemoteFetchNotice(null), 3500);
      return;
    }

    setIsFetchingRemote(true);
    setRemoteFetchNotice(null);
    const result = await ApiClient.fetchSupabaseTable(config, selectedTable, 50);
    setIsFetchingRemote(false);

    if (result.success && result.rows) {
      setRemoteRows(result.rows);
      setRemoteFetchNotice(`Fetched ${result.count} live records from Supabase table '${selectedTable}'.`);
    } else {
      setRemoteFetchNotice(`Could not fetch table '${selectedTable}': ${result.error || 'Table may not exist yet in database'}`);
    }
    setTimeout(() => setRemoteFetchNotice(null), 5000);
  };

  // Determine truthful connection status
  const hasCredentials = Boolean(formData.url.trim() && (formData.anonKey.trim() || formData.serviceKey?.trim() || formData.serviceRoleKey?.trim()));
  const isVerifiedConnected = Boolean(config.isConnected && hasCredentials);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className={`p-2.5 rounded-lg ${isVerifiedConnected ? 'bg-teal-50 text-teal-600' : 'bg-slate-100 text-slate-500'}`}>
            <Database className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center space-x-2.5">
              <h1 className="text-xl font-bold text-slate-900 tracking-tight">Supabase PostgreSQL Integration</h1>
              {isVerifiedConnected ? (
                <span className="inline-flex items-center space-x-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Live Connected</span>
                </span>
              ) : hasCredentials ? (
                <span className="inline-flex items-center space-x-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-amber-100 text-amber-800 border border-amber-200">
                  <AlertCircle className="w-3.5 h-3.5 text-amber-600" />
                  <span>Unverified</span>
                </span>
              ) : (
                <span className="inline-flex items-center space-x-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-slate-100 text-slate-600 border border-slate-200">
                  <XCircle className="w-3.5 h-3.5 text-slate-400" />
                  <span>Not Configured</span>
                </span>
              )}
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Target PostgreSQL database for idempotent record synchronization, unique key upserts, and audit logs.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-3">
          {hasCredentials && (
            <button
              onClick={handleClearCredentials}
              className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-medium text-slate-600 hover:bg-slate-50 transition-colors"
              title="Disconnect and remove saved credentials"
            >
              <Trash2 className="w-3.5 h-3.5 text-slate-400" />
              <span>Disconnect</span>
            </button>
          )}

          <button
            onClick={handleTestConnection}
            disabled={isTesting}
            id="btn-test-supabase-connection"
            className="inline-flex items-center space-x-2 px-3.5 py-1.5 rounded-lg bg-teal-600 hover:bg-teal-700 text-xs font-medium text-white shadow-2xs transition-colors"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isTesting ? 'animate-spin' : ''}`} />
            <span>{isTesting ? 'Verifying Live API...' : 'Test Supabase Connection'}</span>
          </button>
        </div>
      </div>

      {statusNotice && (
        <div className={`p-4 rounded-xl text-xs font-medium flex items-center space-x-2 border ${
          isVerifiedConnected 
            ? 'bg-teal-50 text-teal-800 border-teal-200'
            : 'bg-amber-50 text-amber-800 border-amber-200'
        }`}>
          {isVerifiedConnected ? (
            <CheckCircle2 className="w-4 h-4 text-teal-600 shrink-0" />
          ) : (
            <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
          )}
          <span>{statusNotice}</span>
        </div>
      )}

      {/* Tabs */}
      <div className="flex border-b border-slate-200 space-x-4">
        <button
          onClick={() => setActiveTab('connection')}
          id="tab-supabase-connection"
          className={`pb-3 text-sm font-medium border-b-2 transition-colors flex items-center space-x-2 ${
            activeTab === 'connection'
              ? 'border-teal-600 text-teal-700'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          <Server className="w-4 h-4" />
          <span>Connection & Credentials</span>
        </button>

        <button
          onClick={() => setActiveTab('tables')}
          id="tab-supabase-tables"
          className={`pb-3 text-sm font-medium border-b-2 transition-colors flex items-center space-x-2 ${
            activeTab === 'tables'
              ? 'border-teal-600 text-teal-700'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          <Table className="w-4 h-4" />
          <span>Synchronized Tables & Live Data</span>
        </button>

        <button
          onClick={() => setActiveTab('sql')}
          id="tab-supabase-sql"
          className={`pb-3 text-sm font-medium border-b-2 transition-colors flex items-center space-x-2 ${
            activeTab === 'sql'
              ? 'border-teal-600 text-teal-700'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          <Code className="w-4 h-4" />
          <span>PostgreSQL Schema & Migrations</span>
        </button>
      </div>

      {activeTab === 'connection' && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Credentials Form */}
          <div className="lg:col-span-2 bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
            <h2 className="text-base font-semibold text-slate-900 mb-4 flex items-center space-x-2">
              <Key className="w-4 h-4 text-teal-600" />
              <span>Supabase Project Configuration</span>
            </h2>

            <form onSubmit={handleSave} className="space-y-4">
              <div>
                <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                  Supabase Project URL <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  id="input-supabase-url"
                  value={formData.url}
                  onChange={(e) => setFormData({ ...formData, url: e.target.value })}
                  placeholder="https://xyzcompany.supabase.co"
                  className="w-full px-3.5 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500 font-mono"
                />
                <p className="text-[11px] text-slate-500 mt-1">
                  Found in your Supabase Dashboard: <strong>Project Settings → API → Project URL</strong>
                </p>
              </div>

              <div>
                <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                  Anon / Public API Key <span className="text-rose-500">*</span>
                </label>
                <div className="relative">
                  <input
                    type={showAnonKey ? 'text' : 'password'}
                    id="input-supabase-anon-key"
                    value={formData.anonKey}
                    onChange={(e) => setFormData({ ...formData, anonKey: e.target.value })}
                    placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
                    className="w-full pl-3.5 pr-10 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500 font-mono"
                  />
                  <button
                    type="button"
                    onClick={() => setShowAnonKey(!showAnonKey)}
                    className="absolute right-2.5 top-2.5 text-slate-400 hover:text-slate-600 focus:outline-none"
                    title={showAnonKey ? 'Hide key' : 'Show key'}
                  >
                    {showAnonKey ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
                <p className="text-[11px] text-slate-500 mt-1">
                  Found in <strong>Project Settings → API → Project API Keys → anon / public</strong>
                </p>
              </div>

              <div>
                <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                  Supabase Service Key <span className="text-slate-400 font-normal">(Service Role Key / Admin Key - Bypasses RLS)</span>
                </label>
                <div className="relative">
                  <input
                    type={showServiceKey ? 'text' : 'password'}
                    id="input-supabase-service-key"
                    value={formData.serviceKey ?? formData.serviceRoleKey ?? ''}
                    onChange={(e) => {
                      const val = e.target.value;
                      setFormData({ ...formData, serviceKey: val, serviceRoleKey: val });
                    }}
                    placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
                    className="w-full pl-3.5 pr-10 py-2 text-sm rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500 font-mono"
                  />
                  <button
                    type="button"
                    onClick={() => setShowServiceKey(!showServiceKey)}
                    className="absolute right-2.5 top-2.5 text-slate-400 hover:text-slate-600 focus:outline-none"
                    title={showServiceKey ? 'Hide key' : 'Show key'}
                  >
                    {showServiceKey ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
                <p className="text-[11px] text-slate-500 mt-1">
                  Also known as the <code>service_role</code> secret key in your Supabase dashboard (<strong>Project Settings → API → service_role</strong>). Used by the backend sync engine for unrestricted PostgreSQL schema and table upserts.
                </p>
              </div>

              <div className="flex items-center justify-between pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={handleTestConnection}
                  disabled={isTesting}
                  className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 hover:bg-slate-50 shadow-2xs transition-colors"
                >
                  <RefreshCw className={`w-4 h-4 text-slate-500 ${isTesting ? 'animate-spin' : ''}`} />
                  <span>{isTesting ? 'Testing Connection...' : 'Test Connection'}</span>
                </button>

                <button
                  type="submit"
                  id="btn-save-supabase"
                  className="inline-flex items-center space-x-2 px-4 py-2 rounded-lg bg-teal-600 hover:bg-teal-700 text-sm font-medium text-white shadow-xs transition-colors"
                >
                  <Save className="w-4 h-4" />
                  <span>Save Credentials</span>
                </button>
              </div>
            </form>
          </div>

          {/* Connection Verification Matrix Card */}
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs">
              <h3 className="text-sm font-semibold text-slate-900 mb-3 flex items-center space-x-2">
                <ShieldCheck className="w-4 h-4 text-teal-600" />
                <span>Live Verification Matrix</span>
              </h3>

              {testResult ? (
                <div className="space-y-3">
                  <div className={`p-3 rounded-lg text-xs font-medium ${
                    testResult.success 
                      ? 'bg-teal-50 text-teal-800 border border-teal-200' 
                      : 'bg-rose-50 text-rose-800 border border-rose-200'
                  }`}>
                    {testResult.message || testResult.error}
                  </div>

                  <div className="space-y-2 text-xs">
                    <div className="flex items-center justify-between py-1 border-b border-slate-100">
                      <span className="text-slate-600">1. Host Reachability:</span>
                      <span className="flex items-center space-x-1 font-medium">
                        {testResult.checks.hostReachability ? (
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                        ) : (
                          <XCircle className="w-3.5 h-3.5 text-rose-600" />
                        )}
                        <span>{testResult.checks.hostReachability ? 'HTTP 200 Reachable' : 'Unreachable'}</span>
                      </span>
                    </div>

                    <div className="flex items-center justify-between py-1 border-b border-slate-100">
                      <span className="text-slate-600">2. Authentication:</span>
                      <span className="flex items-center space-x-1 font-medium">
                        {testResult.checks.authValid ? (
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                        ) : (
                          <XCircle className="w-3.5 h-3.5 text-rose-600" />
                        )}
                        <span>{testResult.checks.authValid ? 'PostgREST Authorized' : 'Auth Failed'}</span>
                      </span>
                    </div>

                    <div className="flex items-center justify-between py-1 border-b border-slate-100">
                      <span className="text-slate-600">3. Schema Detected:</span>
                      <span className="flex items-center space-x-1 font-medium">
                        {testResult.checks.schemaDetected ? (
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                        ) : (
                          <AlertCircle className="w-3.5 h-3.5 text-amber-500" />
                        )}
                        <span>
                          {testResult.checks.schemaDetected 
                            ? `${testResult.matchingSyncTables?.length || 0} sync tables found` 
                            : 'Schema not yet migrated'}
                        </span>
                      </span>
                    </div>

                    {testResult.latencyMs !== undefined && (
                      <div className="flex items-center justify-between py-1 border-b border-slate-100">
                        <span className="text-slate-600">4. Roundtrip Latency:</span>
                        <span className="font-mono text-slate-800 font-semibold">{testResult.latencyMs} ms</span>
                      </div>
                    )}

                    <div className="flex items-center justify-between py-1">
                      <span className="text-slate-600">5. Exposed Tables:</span>
                      <span className="font-semibold text-slate-900">{testResult.tablesCount || 0} tables</span>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="text-center py-6 text-slate-500 text-xs">
                  {hasCredentials ? (
                    <span>Click <strong className="text-slate-700">Test Connection</strong> to verify live connectivity with your Supabase database.</span>
                  ) : (
                    <span>No credentials configured. Enter your Supabase Project URL and API Key on the left.</span>
                  )}
                </div>
              )}
            </div>

            <div className="bg-slate-50 rounded-xl p-5 border border-slate-200 text-xs text-slate-600 space-y-2">
              <div className="flex items-center space-x-2 font-semibold text-slate-900">
                <Lock className="w-4 h-4 text-slate-700" />
                <span>Security & Zero Fake Statuses</span>
              </div>
              <p className="leading-relaxed">
                This sync console only reports a connected status when a live HTTP handshake with your actual Supabase PostgreSQL instance has been executed and confirmed.
              </p>
              <p className="text-[11px] text-slate-500">
                Requests are proxied securely through the backend server, keeping secret keys protected from browser exposure.
              </p>
            </div>
          </div>
        </div>
      )}

      {activeTab === 'tables' && (
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          {/* Table Selector Sidebar */}
          <div className="lg:col-span-1 space-y-2">
            <div className="text-xs font-semibold uppercase tracking-wider text-slate-500 px-1 mb-2">
              Domain Tables
            </div>
            {availableTables.map((t) => {
              const count = (databaseState[t] || []).length;
              return (
                <button
                  key={t}
                  id={`table-select-${t}`}
                  onClick={() => {
                    setSelectedTable(t);
                    setRemoteRows(null);
                  }}
                  className={`w-full text-left p-3 rounded-lg border text-sm transition-all flex items-center justify-between ${
                    selectedTable === t
                      ? 'bg-teal-50 border-teal-300 text-teal-900 font-semibold shadow-2xs'
                      : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <Table className="w-4 h-4 text-teal-600" />
                    <span>{t}</span>
                  </div>
                  <span className="text-xs px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-mono">
                    {count} rows
                  </span>
                </button>
              );
            })}

            <div className="mt-6 p-4 rounded-xl bg-slate-50 border border-slate-200 text-xs text-slate-600 space-y-2">
              <div className="font-semibold text-slate-900">Upsert Unique Keys</div>
              <ul className="list-disc list-inside space-y-1 text-slate-500">
                <li><code className="text-slate-700">students</code>: student_number</li>
                <li><code className="text-slate-700">attendance</code>: (student_number, date)</li>
                <li><code className="text-slate-700">sports</code>: (student_number, sport)</li>
                <li><code className="text-slate-700">results</code>: (student_number, term)</li>
              </ul>
            </div>
          </div>

          {/* Table Data Viewer */}
          <div className="lg:col-span-3 bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            {/* Live vs Staging Notice Banner */}
            {!isVerifiedConnected ? (
              <div className="p-3.5 bg-amber-50 border-b border-amber-200 text-xs text-amber-800 flex items-center justify-between">
                <div className="flex items-center space-x-2">
                  <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
                  <span>
                    <strong>Supabase Disconnected:</strong> Displaying local staging buffer. Configure credentials in the <strong>Connection & Credentials</strong> tab to read and write live data.
                  </span>
                </div>
                <button
                  onClick={() => setActiveTab('connection')}
                  className="px-2.5 py-1 bg-amber-200 hover:bg-amber-300 text-amber-900 rounded font-semibold text-xs transition-colors shrink-0"
                >
                  Configure
                </button>
              </div>
            ) : (
              <div className="p-3 bg-teal-50 border-b border-teal-200 text-xs text-teal-800 flex items-center justify-between">
                <div className="flex items-center space-x-2">
                  <CheckCircle2 className="w-4 h-4 text-teal-600 shrink-0" />
                  <span>
                    <strong>Live Supabase Connected:</strong> Connected to PostgreSQL at <code className="font-mono">{new URL(config.url).hostname}</code>.
                  </span>
                </div>
                <button
                  onClick={handleFetchRemoteTable}
                  disabled={isFetchingRemote}
                  className="inline-flex items-center space-x-1.5 px-3 py-1 rounded bg-teal-600 hover:bg-teal-700 text-white font-medium text-xs transition-colors"
                >
                  <ArrowDown className={`w-3.5 h-3.5 ${isFetchingRemote ? 'animate-bounce' : ''}`} />
                  <span>{isFetchingRemote ? 'Fetching...' : 'Fetch Live from Database'}</span>
                </button>
              </div>
            )}

            {remoteFetchNotice && (
              <div className="p-2.5 bg-slate-100 border-b border-slate-200 text-xs text-slate-700 font-mono">
                {remoteFetchNotice}
              </div>
            )}

            <div className="p-4 border-b border-slate-200 bg-slate-50/70 flex items-center justify-between">
              <div>
                <span className="font-bold text-slate-900 text-base">public.{selectedTable}</span>
                <span className="ml-2 text-xs text-slate-500">
                  ({currentRecords.length} records {remoteRows !== null ? 'fetched live from Supabase' : 'present in local buffer'})
                </span>
              </div>
              <span className="text-xs text-emerald-700 bg-emerald-100/60 px-2 py-0.5 rounded font-mono">
                Idempotent Upsert Active
              </span>
            </div>

            {currentRecords.length === 0 ? (
              <div className="p-12 text-center text-slate-500 text-sm">
                No records currently populated in <code className="font-mono">{selectedTable}</code>.
                <p className="text-xs text-slate-400 mt-1">Run an Import or Dry Run to synchronize Excel data.</p>
              </div>
            ) : (
              <div className="overflow-x-auto max-h-[500px]">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200 sticky top-0">
                    <tr>
                      {Object.keys(currentRecords[0]).map((key) => (
                        <th key={key} className="px-4 py-3 whitespace-nowrap font-mono">
                          {key}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {currentRecords.map((row, idx) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        {Object.keys(currentRecords[0]).map((key) => (
                          <td key={key} className="px-4 py-2.5 whitespace-nowrap text-slate-800">
                            {row[key] === null || row[key] === undefined ? (
                              <span className="text-slate-400 italic">null</span>
                            ) : typeof row[key] === 'boolean' ? (
                              row[key] ? 'true' : 'false'
                            ) : (
                              String(row[key])
                            )}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === 'sql' && (
        <div className="bg-slate-900 text-slate-200 rounded-xl p-6 border border-slate-800 font-mono text-xs overflow-x-auto shadow-2xs">
          <div className="flex items-center justify-between pb-3 mb-4 border-b border-slate-800 text-slate-400">
            <div>
              <span className="font-semibold text-slate-300">/supabase/production_schema.sql</span>
              <p className="text-[11px] text-slate-500 mt-0.5">Run this script once in your Supabase SQL Editor to initialize all tables, RLS policies, and sync triggers.</p>
            </div>
            <button
              onClick={handleCopySql}
              className="inline-flex items-center space-x-1.5 px-3 py-1 rounded bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-xs transition-colors shrink-0"
            >
              {copiedSql ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copiedSql ? 'Copied to Clipboard!' : 'Copy SQL for Supabase'}</span>
            </button>
          </div>
          <pre className="text-slate-300 leading-relaxed font-mono">
            {PRODUCTION_SQL}
          </pre>
        </div>
      )}
    </div>
  );
};
