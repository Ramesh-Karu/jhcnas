-- ============================================================================
-- Nextcloud Excel Sync - Production PostgreSQL Schema & RLS Policies
-- Execute this script in your Supabase Dashboard SQL Editor for a fresh production setup.
-- ============================================================================

-- Enable required extensions
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- Function to automatically update 'updated_at' timestamps
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- ----------------------------------------------------------------------------
-- 1. Sync Settings Table
-- ----------------------------------------------------------------------------
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

DROP TRIGGER IF EXISTS trg_sync_settings_updated_at ON sync_settings;
CREATE TRIGGER trg_sync_settings_updated_at
    BEFORE UPDATE ON sync_settings
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 2. Workbooks Registry (Files discovered in Nextcloud)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS workbooks (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL,
    nextcloud_path TEXT NOT NULL UNIQUE,
    file_size BIGINT NOT NULL DEFAULT 0,
    file_hash TEXT, -- SHA-256 checksum for change detection
    last_modified_at TIMESTAMPTZ,
    enabled BOOLEAN NOT NULL DEFAULT TRUE,
    auto_detect_headers BOOLEAN NOT NULL DEFAULT TRUE,
    description TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

DROP TRIGGER IF EXISTS trg_workbooks_updated_at ON workbooks;
CREATE TRIGGER trg_workbooks_updated_at
    BEFORE UPDATE ON workbooks
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 3. Worksheet Mappings (Sheet to Target Table Mapping)
-- ----------------------------------------------------------------------------
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

DROP TRIGGER IF EXISTS trg_worksheet_mappings_updated_at ON worksheet_mappings;
CREATE TRIGGER trg_worksheet_mappings_updated_at
    BEFORE UPDATE ON worksheet_mappings
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 4. Column Mappings (Excel Column to Supabase Table Column)
-- ----------------------------------------------------------------------------
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

DROP TRIGGER IF EXISTS trg_column_mappings_updated_at ON column_mappings;
CREATE TRIGGER trg_column_mappings_updated_at
    BEFORE UPDATE ON column_mappings
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 5. Import Audit Logs
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS import_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_id UUID REFERENCES workbooks(id) ON DELETE SET NULL,
    filename TEXT NOT NULL,
    file_path TEXT NOT NULL,
    file_hash TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Pending', -- Pending, Processing, Success, Partial Success, Failed, Skipped
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

-- ----------------------------------------------------------------------------
-- 6. Import Row-Level Error Details
-- ----------------------------------------------------------------------------
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

-- ----------------------------------------------------------------------------
-- Performance Indexes
-- ----------------------------------------------------------------------------
CREATE INDEX IF NOT EXISTS idx_workbooks_hash ON workbooks(file_hash);
CREATE INDEX IF NOT EXISTS idx_workbooks_path ON workbooks(nextcloud_path);
CREATE INDEX IF NOT EXISTS idx_workbooks_enabled ON workbooks(enabled);
CREATE INDEX IF NOT EXISTS idx_import_logs_hash ON import_logs(file_hash);
CREATE INDEX IF NOT EXISTS idx_import_logs_status ON import_logs(status);
CREATE INDEX IF NOT EXISTS idx_import_logs_started ON import_logs(started_at DESC);
CREATE INDEX IF NOT EXISTS idx_worksheet_mappings_wb ON worksheet_mappings(workbook_id);
CREATE INDEX IF NOT EXISTS idx_column_mappings_ws ON column_mappings(worksheet_mapping_id);
CREATE INDEX IF NOT EXISTS idx_import_errors_log ON import_errors(import_log_id);

-- ----------------------------------------------------------------------------
-- Row Level Security (RLS) & Universal Key Compatibility
-- Allows authenticated users, anon public keys, and service_role to manage sync
-- ----------------------------------------------------------------------------
ALTER TABLE sync_settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE workbooks ENABLE ROW LEVEL SECURITY;
ALTER TABLE worksheet_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE column_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE import_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE import_errors ENABLE ROW LEVEL SECURITY;

-- Grant broad RLS policies for clean production access
DO $$
BEGIN
    -- sync_settings
    DROP POLICY IF EXISTS "Allow full access to sync_settings" ON sync_settings;
    CREATE POLICY "Allow full access to sync_settings" ON sync_settings
        FOR ALL TO public USING (true) WITH CHECK (true);

    -- workbooks
    DROP POLICY IF EXISTS "Allow full access to workbooks" ON workbooks;
    CREATE POLICY "Allow full access to workbooks" ON workbooks
        FOR ALL TO public USING (true) WITH CHECK (true);

    -- worksheet_mappings
    DROP POLICY IF EXISTS "Allow full access to worksheet_mappings" ON worksheet_mappings;
    CREATE POLICY "Allow full access to worksheet_mappings" ON worksheet_mappings
        FOR ALL TO public USING (true) WITH CHECK (true);

    -- column_mappings
    DROP POLICY IF EXISTS "Allow full access to column_mappings" ON column_mappings;
    CREATE POLICY "Allow full access to column_mappings" ON column_mappings
        FOR ALL TO public USING (true) WITH CHECK (true);

    -- import_logs
    DROP POLICY IF EXISTS "Allow full access to import_logs" ON import_logs;
    CREATE POLICY "Allow full access to import_logs" ON import_logs
        FOR ALL TO public USING (true) WITH CHECK (true);

    -- import_errors
    DROP POLICY IF EXISTS "Allow full access to import_errors" ON import_errors;
    CREATE POLICY "Allow full access to import_errors" ON import_errors
        FOR ALL TO public USING (true) WITH CHECK (true);
END $$;
