# Vercel Deployment Guide: Nextcloud Excel Sync

## Overview
The Nextcloud Excel Sync web application is deployed to **Vercel** as a high-performance Next.js application. It serves the administrative dashboard, configuration APIs, interactive workbook analyzer, mapping manager, and dry-run engine.

> **Architecture Principle:** The Vercel application is **NOT** responsible for continuously monitoring Nextcloud. Background synchronization is delegated to the Coolify worker.

---

## 1. Environment Variables Configuration
In your Vercel Project Settings (`Settings` -> `Environment Variables`), configure:

| Variable Name | Required | Environment | Description |
|---|---|---|---|
| `NEXT_PUBLIC_SUPABASE_URL` | **Yes** | Production, Preview | Your Supabase project URL (e.g. `https://xyz.supabase.co`) |
| `NEXT_PUBLIC_SUPABASE_ANON_KEY` | **Yes** | Production, Preview | Supabase public anonymous API key |
| `SUPABASE_SERVICE_ROLE_KEY` | **Yes** | Production, Preview | Supabase service-role secret key (Server-side ONLY) |
| `NEXTCLOUD_URL` | **Yes** | Production, Preview | Nextcloud instance URL (e.g. `https://cloud.example.com`) |
| `NEXTCLOUD_WEBDAV_URL` | **Yes** | Production, Preview | WebDAV endpoint `.../remote.php/dav/files/user/` |
| `NEXTCLOUD_USERNAME` | **Yes** | Production, Preview | Nextcloud WebDAV user |
| `NEXTCLOUD_APP_PASSWORD` | **Yes** | Production, Preview | Generated Nextcloud App Password (never personal password) |
| `NEXTCLOUD_FOLDER` | Optional | Production, Preview | Folder to scan (default: `/ExcelImports`) |
| `WORKER_URL` | Optional | Production, Preview | Public/internal URL of your Coolify worker for manual dispatch |
| `GEMINI_API_KEY` | Optional | Production, Preview | Google Gemini API key for AI-assisted schema mapping suggestions |

> **Security Rule:** Never prefix `SUPABASE_SERVICE_ROLE_KEY` or `NEXTCLOUD_APP_PASSWORD` with `NEXT_PUBLIC_`. They are strictly kept server-side in API Route Handlers.

---

## 2. Deploying via Vercel CLI
```bash
# Install Vercel CLI
npm i -g vercel

# Login and link project
vercel login
vercel link

# Push environment variables
vercel env add NEXT_PUBLIC_SUPABASE_URL
vercel env add NEXT_PUBLIC_SUPABASE_ANON_KEY
vercel env add SUPABASE_SERVICE_ROLE_KEY
vercel env add NEXTCLOUD_URL
vercel env add NEXTCLOUD_WEBDAV_URL
vercel env add NEXTCLOUD_USERNAME
vercel env add NEXTCLOUD_APP_PASSWORD

# Deploy to production
vercel --prod
```

---

## 3. Deploying via GitHub Integration
1. Push your repository to GitHub.
2. In the Vercel Dashboard, click **Add New Project** -> **Import Git Repository**.
3. Set Framework Preset to **Next.js** (or Vite/React if using SPA client build).
4. Fill in the Environment Variables listed above.
5. Click **Deploy**.
