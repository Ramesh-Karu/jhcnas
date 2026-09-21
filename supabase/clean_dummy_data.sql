-- ============================================================================
-- Nextcloud Excel Sync - Clean / Reset All Sync Data in Supabase
-- Run this if you want to wipe out all previous dummy logs and sample data.
-- ============================================================================

TRUNCATE TABLE import_errors CASCADE;
TRUNCATE TABLE import_logs CASCADE;
TRUNCATE TABLE column_mappings CASCADE;
TRUNCATE TABLE worksheet_mappings CASCADE;
TRUNCATE TABLE workbooks CASCADE;

-- Optional: Reset sync settings worker status to idle
UPDATE sync_settings
SET worker_status = 'idle',
    last_sync_at = NULL,
    next_sync_at = NULL;
