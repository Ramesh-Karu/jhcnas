# Nextcloud Excel Sync

A production-ready platform for automatically reading complex Excel workbooks stored in Nextcloud on TrueNAS SCALE, transforming their data into structured records with merged-cell intelligence, and synchronizing those records into a Supabase PostgreSQL database.

## System Architecture

```text
Nextcloud on TrueNAS SCALE
        │
        │ WebDAV / HTTPS
        ▼
  Coolify Worker (Python / FastAPI / openpyxl / pandas)
        │
        │ Excel processing & Idempotent Upsert
        ▼
   Supabase PostgreSQL
        ▲
        │ Configuration & Auditing API
        │
  Vercel Next.js Web Application
        │
        ▼
 Administrator Dashboard
```

- **Vercel Application**: Administrator dashboard, visual workbook analyzer, mapping management, connection testing, error auditing, and manual dry runs.
- **Coolify Worker**: Background synchronization worker handling polling, SHA-256 change detection, openpyxl merged-cell parsing, data transformations, validations, and idempotent upserts.
- **Supabase PostgreSQL**: Database storing synchronization settings, mappings, import history logs, row-level errors, and target domain tables.
- **Nextcloud on TrueNAS SCALE**: Source of Excel workbooks accessed securely via WebDAV App Password.
