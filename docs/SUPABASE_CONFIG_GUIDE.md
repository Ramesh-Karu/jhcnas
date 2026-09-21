# Supabase PostgreSQL Configuration Guide

## 1. PostgreSQL Schema Installation
Run the migration scripts located in `/supabase/migrations/`:
1. In the Supabase Dashboard, open the **SQL Editor**.
2. Run `001_initial_schema.sql` to generate:
   - `sync_settings`
   - `workbooks`
   - `worksheet_mappings`
   - `column_mappings`
   - `import_logs`
   - `import_errors`
3. Run `002_sample_data_and_rls.sql` to configure:
   - Target domain tables: `students`, `attendance`, `sports`, `medical`, `results`
   - Unique constraints for idempotent Upsert
   - Foreign key relationships
   - Row Level Security (RLS) policies

## 2. API Credentials & Roles
Under **Project Settings** -> **API**:
- **URL**: `https://<project-ref>.supabase.co`
- **anon key**: Public key used client-side for authenticated UI sessions
- **service_role key**: Secret elevated key used **ONLY** by the Coolify worker and Vercel server-side APIs to bypass RLS and perform batch upserts.

## 3. Storage Bucket (Optional Archive)
If automatic archiving of successfully processed workbooks is enabled (`BACKUP_TO_STORAGE=true`):
1. Navigate to **Storage** in Supabase.
2. Create a private bucket named `excel-archives`.
