-- ============================================================================
-- Nextcloud Excel Sync - TOTAL SYSTEM PRODUCTION POSTGRESQL SCHEMA
-- Complete database schema for Supabase: Destination Business Tables & Sync Metadata
-- ============================================================================

-- Enable required PostgreSQL extensions
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

-- Optional Helper RPC to execute dynamic SQL from authorized services
CREATE OR REPLACE FUNCTION exec_sql(query TEXT)
RETURNS VOID AS $$
BEGIN
    EXECUTE query;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- ============================================================================
-- SECTION A: CORE DESTINATION BUSINESS TABLES (Excel Data Repositories)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. Students Repository Table (Target for all Student Admission Workbooks)
-- Target for: 2026_stu.xlsx ... 2034_stu.xlsx, students_complex.xlsx, etc.
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.students (
    admission_no TEXT PRIMARY KEY,
    name_with_initials TEXT,
    full_name TEXT,
    tamil_name TEXT,
    grade TEXT,
    class_section TEXT,
    medium TEXT DEFAULT 'Tamil',
    date_of_birth DATE,
    gender TEXT,
    religion TEXT,
    house TEXT,
    index_number TEXT,
    guardian_name TEXT,
    phone_number TEXT,
    address TEXT,
    admission_date DATE,
    academic_year TEXT,
    status TEXT DEFAULT 'Active',
    national_id TEXT,
    email TEXT,
    emergency_contact TEXT,
    previous_school TEXT,
    special_notes TEXT,
    _sheet_source TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

DROP TRIGGER IF EXISTS trg_students_updated_at ON public.students;
CREATE TRIGGER trg_students_updated_at
    BEFORE UPDATE ON public.students
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 2. Master Timetable Matrix Table (Target for Class Wise Time Table.xlsx)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.timetable_master (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    day_of_week TEXT NOT NULL,
    grade TEXT NOT NULL,
    division TEXT NOT NULL,
    period_number INTEGER NOT NULL,
    subject_code TEXT NOT NULL,
    teacher_name TEXT,
    room_number TEXT,
    academic_year TEXT DEFAULT '2026',
    remarks TEXT,
    _sheet_source TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(day_of_week, grade, division, period_number)
);

DROP TRIGGER IF EXISTS trg_timetable_master_updated_at ON public.timetable_master;
CREATE TRIGGER trg_timetable_master_updated_at
    BEFORE UPDATE ON public.timetable_master
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 3. Subject Teacher Allocations Table (Target for Subject teacher Class wise.xlsx)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.subject_teacher_allocations (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    teacher_name TEXT NOT NULL,
    subject TEXT NOT NULL,
    grade TEXT NOT NULL,
    class_section TEXT NOT NULL,
    periods_per_week INTEGER DEFAULT 1,
    medium TEXT DEFAULT 'Tamil',
    academic_year TEXT DEFAULT '2026',
    remarks TEXT,
    _sheet_source TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(teacher_name, subject, grade, class_section)
);

DROP TRIGGER IF EXISTS trg_teacher_allocations_updated_at ON public.subject_teacher_allocations;
CREATE TRIGGER trg_teacher_allocations_updated_at
    BEFORE UPDATE ON public.subject_teacher_allocations
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 4. Donations & Contributions Ledger (Target for JHC Donation Details.xlsx)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.jhc_donations (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    receipt_no TEXT,
    donation_date DATE,
    donor_name TEXT NOT NULL,
    donor_type TEXT DEFAULT 'Individual',
    category TEXT DEFAULT 'Cash Donation',
    item_description TEXT,
    quantity INTEGER DEFAULT 1,
    amount NUMERIC(14, 2) DEFAULT 0.00,
    payment_mode TEXT DEFAULT 'Cash',
    purpose TEXT,
    folio TEXT,
    financial_year TEXT DEFAULT '2026',
    remarks TEXT,
    _sheet_source TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

DROP TRIGGER IF EXISTS trg_donations_updated_at ON public.jhc_donations;
CREATE TRIGGER trg_donations_updated_at
    BEFORE UPDATE ON public.jhc_donations
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 5. College Inventory & Assets (Target for Jaffna Hindu College Inventory.xlsx)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.college_inventory (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    item_code TEXT,
    item_name TEXT NOT NULL,
    tamil_item_name TEXT,
    category TEXT,
    location TEXT,
    quantity INTEGER DEFAULT 1,
    unit_of_measure TEXT DEFAULT 'Nos',
    condition TEXT DEFAULT 'Good',
    purchase_date DATE,
    unit_price NUMERIC(12, 2) DEFAULT 0.00,
    total_valuation NUMERIC(14, 2) DEFAULT 0.00,
    custodian TEXT,
    section_number TEXT,
    ledger_balance INTEGER,
    actual_stock INTEGER,
    remarks TEXT,
    _sheet_source TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

DROP TRIGGER IF EXISTS trg_inventory_updated_at ON public.college_inventory;
CREATE TRIGGER trg_inventory_updated_at
    BEFORE UPDATE ON public.college_inventory
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();


-- ============================================================================
-- SECTION B: SYSTEM SYNCHRONIZATION & METADATA INFRASTRUCTURE
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 6. Sync Settings Table
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.sync_settings (
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

DROP TRIGGER IF EXISTS trg_sync_settings_updated_at ON public.sync_settings;
CREATE TRIGGER trg_sync_settings_updated_at
    BEFORE UPDATE ON public.sync_settings
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 7. Workbooks Registry (Files discovered in Nextcloud)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.workbooks (
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

DROP TRIGGER IF EXISTS trg_workbooks_updated_at ON public.workbooks;
CREATE TRIGGER trg_workbooks_updated_at
    BEFORE UPDATE ON public.workbooks
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 8. Worksheet Mappings (Sheet to Target Supabase Table Alignment)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.worksheet_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_id UUID NOT NULL REFERENCES public.workbooks(id) ON DELETE CASCADE,
    worksheet_name TEXT NOT NULL,
    supabase_table TEXT NOT NULL,
    header_row INTEGER NOT NULL DEFAULT 1,
    data_start_row INTEGER NOT NULL DEFAULT 2,
    data_end_row INTEGER,
    section_heading_col TEXT,
    section_heading_target_col TEXT,
    sync_policy TEXT DEFAULT 'EXCEL_TO_DB',
    enabled BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(workbook_id, worksheet_name)
);

DROP TRIGGER IF EXISTS trg_worksheet_mappings_updated_at ON public.worksheet_mappings;
CREATE TRIGGER trg_worksheet_mappings_updated_at
    BEFORE UPDATE ON public.worksheet_mappings
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 9. Column Mappings (Excel Column to Supabase Table Column)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.column_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    worksheet_mapping_id UUID NOT NULL REFERENCES public.worksheet_mappings(id) ON DELETE CASCADE,
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

DROP TRIGGER IF EXISTS trg_column_mappings_updated_at ON public.column_mappings;
CREATE TRIGGER trg_column_mappings_updated_at
    BEFORE UPDATE ON public.column_mappings
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 10. Import Audit Logs
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.import_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_id UUID REFERENCES public.workbooks(id) ON DELETE SET NULL,
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

-- ----------------------------------------------------------------------------
-- 11. Import Row-Level Error Details
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.import_errors (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    import_log_id UUID NOT NULL REFERENCES public.import_logs(id) ON DELETE CASCADE,
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
-- 12. AI Presets & Archetype Rules
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.ai_presets (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    description TEXT,
    archetype TEXT NOT NULL DEFAULT 'STANDARD_TABULAR',
    badge TEXT,
    is_system BOOLEAN NOT NULL DEFAULT FALSE,
    filename_pattern TEXT,
    sheet_mappings JSONB NOT NULL DEFAULT '[]'::jsonb,
    unpivot_config JSONB,
    skip_merged_year_rows BOOLEAN NOT NULL DEFAULT TRUE,
    tags TEXT[] DEFAULT '{}',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

DROP TRIGGER IF EXISTS trg_ai_presets_updated_at ON public.ai_presets;
CREATE TRIGGER trg_ai_presets_updated_at
    BEFORE UPDATE ON public.ai_presets
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- ----------------------------------------------------------------------------
-- 13. Two-Way Sync Conflict Staging Queue
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.sync_conflicts (
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

-- ----------------------------------------------------------------------------
-- 14. Two-Way Sync Audit Trail
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.sync_two_way_audit (
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

-- ============================================================================
-- SECTION C: HIGH-PERFORMANCE INDEXES
-- ============================================================================
CREATE INDEX IF NOT EXISTS idx_students_grade_class ON public.students(grade, class_section);
CREATE INDEX IF NOT EXISTS idx_students_tamil_name ON public.students(tamil_name);
CREATE INDEX IF NOT EXISTS idx_students_full_name ON public.students(full_name);
CREATE INDEX IF NOT EXISTS idx_students_dob ON public.students(date_of_birth);

CREATE INDEX IF NOT EXISTS idx_timetable_day_grade ON public.timetable_master(day_of_week, grade, division);
CREATE INDEX IF NOT EXISTS idx_timetable_teacher ON public.timetable_master(teacher_name);

CREATE INDEX IF NOT EXISTS idx_teacher_alloc_teacher ON public.subject_teacher_allocations(teacher_name);
CREATE INDEX IF NOT EXISTS idx_teacher_alloc_grade ON public.subject_teacher_allocations(grade, class_section);

CREATE INDEX IF NOT EXISTS idx_donations_donor ON public.jhc_donations(donor_name);
CREATE INDEX IF NOT EXISTS idx_donations_date ON public.jhc_donations(donation_date);

CREATE INDEX IF NOT EXISTS idx_inventory_code ON public.college_inventory(item_code);
CREATE INDEX IF NOT EXISTS idx_inventory_category ON public.college_inventory(category);

CREATE INDEX IF NOT EXISTS idx_workbooks_hash ON public.workbooks(file_hash);
CREATE INDEX IF NOT EXISTS idx_workbooks_path ON public.workbooks(nextcloud_path);
CREATE INDEX IF NOT EXISTS idx_import_logs_hash ON public.import_logs(file_hash);
CREATE INDEX IF NOT EXISTS idx_import_logs_status ON public.import_logs(status);
CREATE INDEX IF NOT EXISTS idx_import_logs_started ON public.import_logs(started_at DESC);
CREATE INDEX IF NOT EXISTS idx_worksheet_mappings_wb ON public.worksheet_mappings(workbook_id);
CREATE INDEX IF NOT EXISTS idx_column_mappings_ws ON public.column_mappings(worksheet_mapping_id);
CREATE INDEX IF NOT EXISTS idx_import_errors_log ON public.import_errors(import_log_id);
CREATE INDEX IF NOT EXISTS idx_sync_conflicts_status ON public.sync_conflicts(status);
CREATE INDEX IF NOT EXISTS idx_sync_conflicts_pk ON public.sync_conflicts(target_table, primary_key_val);

-- ============================================================================
-- SECTION D: ROW LEVEL SECURITY & UNIVERSAL ACCESS POLICIES
-- ============================================================================
ALTER TABLE public.students ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.timetable_master ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.subject_teacher_allocations ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.jhc_donations ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.college_inventory ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.sync_settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.workbooks ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.worksheet_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.column_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.import_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.import_errors ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.ai_presets ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.sync_conflicts ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.sync_two_way_audit ENABLE ROW LEVEL SECURITY;

DO $$
DECLARE
    tbl text;
    tables text[] := ARRAY[
        'students', 'timetable_master', 'subject_teacher_allocations', 
        'jhc_donations', 'college_inventory', 'sync_settings', 
        'workbooks', 'worksheet_mappings', 'column_mappings', 
        'import_logs', 'import_errors', 'ai_presets', 
        'sync_conflicts', 'sync_two_way_audit'
    ];
BEGIN
    FOREACH tbl IN ARRAY tables LOOP
        EXECUTE format('DROP POLICY IF EXISTS "Allow full access to %I" ON public.%I;', tbl, tbl);
        EXECUTE format('CREATE POLICY "Allow full access to %I" ON public.%I FOR ALL TO public USING (true) WITH CHECK (true);', tbl, tbl);
        EXECUTE format('GRANT ALL ON public.%I TO anon, authenticated, service_role;', tbl);
    END LOOP;
END $$;
