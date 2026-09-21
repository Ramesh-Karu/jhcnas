-- Nextcloud Excel Sync - Supabase PostgreSQL Schema Migration
-- Migration 001: Core Architecture Tables

-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 1. Sync Settings Table
CREATE TABLE IF NOT EXISTS sync_settings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    nextcloud_url TEXT NOT NULL,
    nextcloud_webdav_url TEXT NOT NULL,
    nextcloud_username TEXT NOT NULL,
    nextcloud_folder TEXT NOT NULL DEFAULT '/ExcelImports',
    sync_interval TEXT NOT NULL DEFAULT '15m', -- 5m, 15m, 30m, 1h, daily
    auto_sync_enabled BOOLEAN NOT NULL DEFAULT TRUE,
    backup_to_storage BOOLEAN NOT NULL DEFAULT FALSE,
    storage_bucket TEXT DEFAULT 'excel-archives',
    last_sync_at TIMESTAMPTZ,
    next_sync_at TIMESTAMPTZ,
    worker_status TEXT NOT NULL DEFAULT 'idle', -- idle, running, error, paused
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 2. Workbooks Registry
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

-- 3. Worksheet Mappings
CREATE TABLE IF NOT EXISTS worksheet_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_id UUID NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
    worksheet_name TEXT NOT NULL,
    supabase_table TEXT NOT NULL,
    header_row INTEGER NOT NULL DEFAULT 1,
    data_start_row INTEGER NOT NULL DEFAULT 2,
    data_end_row INTEGER, -- Optional limit or empty for full range
    section_heading_col TEXT, -- E.g. 'A' if merged headings apply to rows below
    section_heading_target_col TEXT, -- Target Supabase column for section heading (e.g. 'class')
    enabled BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(workbook_id, worksheet_name)
);

-- 4. Column Mappings
CREATE TABLE IF NOT EXISTS column_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    worksheet_mapping_id UUID NOT NULL REFERENCES worksheet_mappings(id) ON DELETE CASCADE,
    excel_column TEXT NOT NULL, -- e.g. "A", "B" or detected column name
    excel_header TEXT NOT NULL, -- e.g. "Student ID", "Student Name"
    supabase_column TEXT NOT NULL, -- e.g. "student_number", "name"
    data_type TEXT NOT NULL DEFAULT 'text', -- text, integer, decimal, boolean, date, timestamp, json
    required BOOLEAN NOT NULL DEFAULT FALSE,
    unique_key BOOLEAN NOT NULL DEFAULT FALSE, -- Used for Upsert duplicate prevention
    default_value TEXT,
    transformation TEXT NOT NULL DEFAULT 'none', 
    -- transformations: none, trim, uppercase, lowercase, parse_date, parse_number, 
    --                  yes_no_to_boolean, pa_to_status, normalize_phone, normalize_id
    validation_regex TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 5. Import Logs Table
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

-- 6. Import Row-Level Errors
CREATE TABLE IF NOT EXISTS import_errors (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    import_log_id UUID NOT NULL REFERENCES import_logs(id) ON DELETE CASCADE,
    worksheet_name TEXT NOT NULL,
    row_number INTEGER NOT NULL,
    excel_column TEXT,
    column_name TEXT,
    raw_value TEXT,
    error_message TEXT NOT NULL,
    error_type TEXT NOT NULL DEFAULT 'validation', -- validation, type_mismatch, missing_required, foreign_key, duplicate
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Indices for rapid querying & worker lookup
CREATE INDEX IF NOT EXISTS idx_workbooks_hash ON workbooks(file_hash);
CREATE INDEX IF NOT EXISTS idx_workbooks_path ON workbooks(nextcloud_path);
CREATE INDEX IF NOT EXISTS idx_import_logs_hash ON import_logs(file_hash);
CREATE INDEX IF NOT EXISTS idx_import_logs_status ON import_logs(status);
CREATE INDEX IF NOT EXISTS idx_import_logs_started ON import_logs(started_at DESC);
CREATE INDEX IF NOT EXISTS idx_worksheet_mappings_wb ON worksheet_mappings(workbook_id);
CREATE INDEX IF NOT EXISTS idx_column_mappings_ws ON column_mappings(worksheet_mapping_id);
CREATE INDEX IF NOT EXISTS idx_import_errors_log ON import_errors(import_log_id);
