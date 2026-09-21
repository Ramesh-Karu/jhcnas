# Coolify Deployment Guide: Background Sync Worker

## Overview
The synchronization worker runs as a dedicated Python Docker service managed by **Coolify** on your Oracle Cloud (OCI) VPS or dedicated server.

It continuously executes:
- Scheduled Nextcloud WebDAV polling (5m, 15m, 30m, 1h, daily)
- File change detection using SHA-256 hash comparison
- Multi-sheet Excel workbook analysis with openpyxl
- Intelligent merged cell handling & section heading propagation
- Configurable transformations (P/A attendance, dates, phone, etc.)
- In-batch data validation & duplicate prevention
- Idempotent upserting to Supabase PostgreSQL
- Import logging & row-level error reporting

---

## 1. Coolify Application Setup
1. In your Coolify dashboard, select your Server (e.g. Oracle Cloud VM) and Project.
2. Click **+ New Resource** -> **Git Repository** (or **Docker Compose**).
3. Set Repository URL to your repository and Base Directory to `/worker`.
4. Select **Dockerfile** as the Build Pack.
5. Set the Port to `8000`.

---

## 2. Environment Variables in Coolify
Add the following under **Environment Variables** in Coolify:

```env
PORT=8000
HOST=0.0.0.0

# Nextcloud WebDAV on TrueNAS SCALE
NEXTCLOUD_URL=https://cloud.example.com
NEXTCLOUD_WEBDAV_URL=https://cloud.example.com/remote.php/dav/files/excel-sync/
NEXTCLOUD_USERNAME=excel-sync
NEXTCLOUD_APP_PASSWORD=your-secure-app-password
NEXTCLOUD_FOLDER=/ExcelImports

# Supabase Credentials (from Oracle Cloud self-hosted or Supabase Cloud)
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_SERVICE_ROLE_KEY=eyJhbGciOiJIUzI1NiIsInR5cCI6...

# Scheduling
SYNC_INTERVAL=15m
```

---

## 3. Healthcheck & Coolify Restarts
The worker includes a native healthcheck at `GET /health`.
In Coolify:
- Healthcheck URL: `http://localhost:8000/health`
- Interval: 30s
- Retries: 3
- Restart Policy: `unless-stopped`
