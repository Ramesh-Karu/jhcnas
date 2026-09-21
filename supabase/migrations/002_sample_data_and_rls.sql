-- Nextcloud Excel Sync - Supabase PostgreSQL Schema Migration
-- Migration 002: Target Domain Tables, RLS, and Seed Mappings

-- 1. Students Master Table (Target for "Student Details" worksheet)
CREATE TABLE IF NOT EXISTS students (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    student_number TEXT NOT NULL UNIQUE, -- Unique Upsert Key
    name TEXT NOT NULL,
    class TEXT NOT NULL,
    gender TEXT,
    dob DATE,
    phone TEXT,
    emergency_contact TEXT,
    status TEXT DEFAULT 'Active',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 2. Attendance Table (Target for "Attendance" worksheet)
CREATE TABLE IF NOT EXISTS attendance (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    student_number TEXT NOT NULL REFERENCES students(student_number) ON UPDATE CASCADE ON DELETE CASCADE,
    attendance_date DATE NOT NULL,
    status TEXT NOT NULL, -- 'Present', 'Absent', 'Late', 'Excused'
    remarks TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(student_number, attendance_date) -- Composite unique key for upsert
);

-- 3. Sports Table (Target for "Sports" worksheet)
CREATE TABLE IF NOT EXISTS sports (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    student_number TEXT NOT NULL REFERENCES students(student_number) ON UPDATE CASCADE ON DELETE CASCADE,
    sport_name TEXT NOT NULL,
    position TEXT,
    medical_clearance BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(student_number, sport_name)
);

-- 4. Medical Records Table (Target for "Medical" worksheet)
CREATE TABLE IF NOT EXISTS medical (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    student_number TEXT NOT NULL UNIQUE REFERENCES students(student_number) ON UPDATE CASCADE ON DELETE CASCADE,
    blood_group TEXT,
    allergies TEXT,
    doctor_contact TEXT,
    last_checkup_date DATE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 5. Results Table (Target for "Results" worksheet)
CREATE TABLE IF NOT EXISTS results (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    student_number TEXT NOT NULL REFERENCES students(student_number) ON UPDATE CASCADE ON DELETE CASCADE,
    term TEXT NOT NULL,
    mathematics NUMERIC(5,2),
    science NUMERIC(5,2),
    english NUMERIC(5,2),
    total_score NUMERIC(6,2),
    grade TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(student_number, term)
);

-- 6. Enable Row Level Security (RLS) on all configuration and log tables
ALTER TABLE sync_settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE workbooks ENABLE ROW LEVEL SECURITY;
ALTER TABLE worksheet_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE column_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE import_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE import_errors ENABLE ROW LEVEL SECURITY;

-- 7. RLS Policies: Authenticated Admin Read/Write, Service Role Full Access
DO $$
BEGIN
    -- Authenticated Admin Access
    DROP POLICY IF EXISTS "Allow authenticated admin full access to sync_settings" ON sync_settings;
    CREATE POLICY "Allow authenticated admin full access to sync_settings" ON sync_settings
        FOR ALL TO authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow authenticated admin full access to workbooks" ON workbooks;
    CREATE POLICY "Allow authenticated admin full access to workbooks" ON workbooks
        FOR ALL TO authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow authenticated admin full access to worksheet_mappings" ON worksheet_mappings;
    CREATE POLICY "Allow authenticated admin full access to worksheet_mappings" ON worksheet_mappings
        FOR ALL TO authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow authenticated admin full access to column_mappings" ON column_mappings;
    CREATE POLICY "Allow authenticated admin full access to column_mappings" ON column_mappings
        FOR ALL TO authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow authenticated admin full access to import_logs" ON import_logs;
    CREATE POLICY "Allow authenticated admin full access to import_logs" ON import_logs
        FOR ALL TO authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow authenticated admin full access to import_errors" ON import_errors;
    CREATE POLICY "Allow authenticated admin full access to import_errors" ON import_errors
        FOR ALL TO authenticated USING (true) WITH CHECK (true);

    -- Service Role bypasses RLS automatically in Supabase, but explicit policy guarantees worker execution
    DROP POLICY IF EXISTS "Service role sync_settings access" ON sync_settings;
    CREATE POLICY "Service role sync_settings access" ON sync_settings FOR ALL TO service_role USING (true) WITH CHECK (true);
END $$;
