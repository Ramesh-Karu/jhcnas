"""
Nextcloud Excel Sync - Background Worker Service
Deploys on Coolify to handle scheduled Nextcloud WebDAV polling, Excel workbook analysis,
merged-cell intelligence, pandas transformations, and Supabase PostgreSQL upserts.
"""
import os
import shutil
import tempfile
from contextlib import asynccontextmanager
from typing import Dict, Any, Optional

from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

from nextcloud.client import NextcloudClient
from excel.parser import ExcelParser
from mappings.manager import MappingManager
from validation.validator import DataValidator
from database.client import SupabaseSyncClient
from sync_logger.logger import SyncLogger

# Environment Configuration
PORT = int(os.getenv("PORT", 8000))
HOST = os.getenv("HOST", "0.0.0.0")

NEXTCLOUD_URL = os.getenv("NEXTCLOUD_URL", "")
NEXTCLOUD_WEBDAV_URL = os.getenv("NEXTCLOUD_WEBDAV_URL", "")
NEXTCLOUD_USERNAME = os.getenv("NEXTCLOUD_USERNAME", "")
NEXTCLOUD_APP_PASSWORD = os.getenv("NEXTCLOUD_APP_PASSWORD", "")
NEXTCLOUD_FOLDER = os.getenv("NEXTCLOUD_FOLDER", "/ExcelImports")

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "") or os.getenv("SUPABASE_ANON_KEY", "") or os.getenv("SUPABASE_KEY", "")
SYNC_INTERVAL = os.getenv("SYNC_INTERVAL", "15m")

# Global scheduler
scheduler = BackgroundScheduler()

def get_interval_minutes(interval_str: str) -> int:
    mapping = {
        "5m": 5,
        "15m": 15,
        "30m": 30,
        "1h": 60,
        "daily": 1440
    }
    return mapping.get(interval_str.lower(), 15)

def run_synchronization_cycle(dry_run: bool = False, specific_file: Optional[str] = None) -> Dict[str, Any]:
    """
    Core Synchronization Orchestrator:
    1. Connects to Nextcloud WebDAV on TrueNAS SCALE
    2. Lists Excel files in configured folder
    3. Detects file modifications using SHA-256 hashes
    4. Downloads, analyzes merged cells, headers, and sheets
    5. Applies transformations & validates data
    6. Upserts records to Supabase or simulates dry-run
    7. Emits structured import_logs and import_errors
    """
    print(f"[Worker] Starting sync cycle (dry_run={dry_run}, specific_file={specific_file})...")

    if not NEXTCLOUD_WEBDAV_URL or not SUPABASE_URL:
        print("[Worker] Credentials not configured. Skipping cycle.")
        return {"status": "skipped", "reason": "Missing Nextcloud or Supabase configuration."}

    nc = NextcloudClient(NEXTCLOUD_WEBDAV_URL, NEXTCLOUD_USERNAME, NEXTCLOUD_APP_PASSWORD, NEXTCLOUD_FOLDER)
    supa = SupabaseSyncClient(SUPABASE_URL, SUPABASE_KEY)
    logger = SyncLogger(supa.client)

    files = nc.list_excel_files()
    if specific_file:
        files = [f for f in files if f["filename"] == specific_file or f["relative_path"] == specific_file]

    results = []
    temp_dir = tempfile.mkdtemp(prefix="excel_sync_")

    try:
        for f_info in files:
            file_name = f_info["filename"]
            file_path = f_info["relative_path"]
            temp_local_file = os.path.join(temp_dir, file_name)

            # Step 1: Download file & calculate hash
            success, sha256_hash, file_size = nc.download_file(file_path, temp_local_file)
            if not success:
                print(f"[Worker] Failed to download {file_name}")
                continue

            # Step 2: File Change Detection (Section 14)
            if not dry_run and supa.check_file_unchanged(file_path, sha256_hash):
                print(f"[Worker] File {file_name} unchanged (SHA-256: {sha256_hash[:8]}...). Skipping.")
                continue

            # Step 3: Fetch mappings from Supabase
            wb_config = supa.get_workbook_config(file_path)
            workbook_id = wb_config["id"] if wb_config else None

            # Step 4: Initialize Import Log
            log_id = logger.create_import_log(
                filename=file_name,
                file_path=file_path,
                file_hash=sha256_hash,
                workbook_id=workbook_id,
                is_dry_run=dry_run
            )

            # Step 5: Parse Excel Workbook with OpenPyXL
            try:
                parser = ExcelParser(temp_local_file)
                analysis = parser.analyze_workbook()

                total_worksheets = len(analysis["worksheets"])
                total_processed = 0
                total_inserted = 0
                total_updated = 0
                total_failed = 0
                file_errors = []

                # Active worksheet mappings
                active_mappings = {}
                has_configured_mappings = False
                if wb_config and wb_config.get("worksheet_mappings"):
                    has_configured_mappings = True
                    for wm in wb_config["worksheet_mappings"]:
                        if wm.get("enabled", True):
                            active_mappings[wm["worksheet_name"]] = wm

                # Step 6: Process Each Worksheet
                sheets_synced_count = 0
                for ws_info in analysis["worksheets"]:
                    sheet_name = ws_info["sheet_name"]
                    wm = active_mappings.get(sheet_name)

                    # Multi-sheet protection: If workbook has explicit mappings, ONLY process active mapped sheets!
                    if has_configured_mappings and not wm:
                        print(f"[Worker] Skipping sheet '{sheet_name}' in '{file_name}' (unmapped or sync disabled).")
                        continue

                    # Skip empty sheets (e.g. cover sheets, chart sheets, or empty tabs)
                    if ws_info.get("total_rows", 0) <= 1 or not ws_info.get("headers"):
                        print(f"[Worker] Skipping empty or headerless sheet '{sheet_name}'.")
                        continue

                    sheets_synced_count += 1
                    # Default fallback parameters if unmapped
                    header_row = wm.get("header_row", ws_info["detected_header_row"]) if wm else ws_info["detected_header_row"]
                    data_start = wm.get("data_start_row", ws_info["detected_data_start_row"]) if wm else ws_info["detected_data_start_row"]
                    target_table = wm.get("supabase_table", sheet_name.lower().replace(" ", "_")) if wm else sheet_name.lower().replace(" ", "_")
                    section_col = wm.get("section_heading_target_col", "class") if wm else None

                    # Extract records with merged-cell downward propagation
                    raw_records, headers = parser.extract_sheet_records(
                        worksheet_name=sheet_name,
                        header_row=header_row,
                        data_start_row=data_start,
                        section_heading_target_col=section_col
                    )

                    col_maps = wm.get("column_mappings", []) if wm else []
                    if not col_maps:
                        # Auto-generate identity mappings from detected headers
                        col_maps = [{
                            "excel_column": h["col_letter"],
                            "excel_header": h["name"],
                            "supabase_column": h["name"].lower().replace(" ", "_").replace(".", ""),
                            "data_type": "text",
                            "required": False,
                            "unique_key": (h["col_letter"] == "A"),
                            "transformation": "trim"
                        } for h in ws_info["headers"]]

                    # Apply transformations
                    transformed_records = [
                        MappingManager.transform_record(r, col_maps, section_col)
                        for r in raw_records
                    ]

                    # Validate records
                    valid_recs, failed_recs, sheet_errs = DataValidator.validate_records(
                        transformed_records, col_maps, sheet_name
                    )

                    total_processed += len(transformed_records)
                    total_failed += len(failed_recs)
                    file_errors.extend(sheet_errs)

                    # Unique keys for upsert
                    unique_cols = [cm["supabase_column"] for cm in col_maps if cm.get("unique_key")]

                    # Step 7: Database Interaction or Dry Run
                    if dry_run:
                        dry_stats = supa.execute_dry_run_analysis(target_table, unique_cols, valid_recs)
                        total_inserted += dry_stats["proposed_inserts"]
                        total_updated += dry_stats["proposed_updates"]
                    else:
                        ins, upd, err = supa.upsert_records(target_table, valid_recs, unique_cols)
                        total_inserted += ins
                        total_updated += upd
                        if err:
                            file_errors.append({
                                "worksheet_name": sheet_name,
                                "row_number": 0,
                                "error_message": f"Database write error: {err}",
                                "error_type": "database_error"
                            })

                # Step 8: Finalize Log Status
                final_status = "Success"
                if total_failed > 0 and total_inserted + total_updated > 0:
                    final_status = "Partial Success"
                elif total_failed > 0 and total_inserted + total_updated == 0:
                    final_status = "Failed"

                logger.finalize_import_log(
                    log_id=log_id,
                    status=final_status,
                    number_of_worksheets=total_worksheets,
                    rows_processed=total_processed,
                    rows_inserted=total_inserted,
                    rows_updated=total_updated,
                    rows_failed=total_failed,
                    error_summary=f"{total_failed} errors across {total_worksheets} sheets" if total_failed else None,
                    details={"worksheets_processed": [w["sheet_name"] for w in analysis["worksheets"]]}
                )

                # Record individual row errors
                if file_errors:
                    logger.record_row_errors(log_id, file_errors)

                # Update workbook hash if non-dry-run success
                if not dry_run and final_status in ("Success", "Partial Success"):
                    supa.update_workbook_hash(file_path, file_name, sha256_hash, file_size)

                results.append({
                    "filename": file_name,
                    "status": final_status,
                    "rows_processed": total_processed,
                    "rows_inserted": total_inserted,
                    "rows_updated": total_updated,
                    "rows_failed": total_failed,
                    "log_id": log_id
                })

            except Exception as e:
                print(f"[Worker] Error processing {file_name}: {e}")
                logger.finalize_import_log(
                    log_id=log_id,
                    status="Failed",
                    number_of_worksheets=1,
                    rows_processed=0,
                    rows_inserted=0,
                    rows_updated=0,
                    rows_failed=0,
                    error_summary=str(e)
                )
                results.append({"filename": file_name, "status": "Failed", "error": str(e)})

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return {"status": "completed", "files_processed": len(results), "results": results}


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Start background synchronization scheduler
    minutes = get_interval_minutes(SYNC_INTERVAL)
    scheduler.add_job(
        run_synchronization_cycle,
        trigger=IntervalTrigger(minutes=minutes),
        id="sync_job",
        replace_existing=True
    )
    scheduler.start()
    print(f"[Worker] Scheduler started with interval {SYNC_INTERVAL} ({minutes}m)")
    yield
    scheduler.shutdown()

app = FastAPI(
    title="Nextcloud Excel Sync Worker",
    description="Coolify background synchronization worker for Nextcloud and Supabase",
    version="1.0.0",
    lifespan=lifespan
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class TriggerRequest(BaseModel):
    dry_run: bool = False
    filename: Optional[str] = None

class UpdateSecretsRequest(BaseModel):
    nextcloud_url: Optional[str] = None
    nextcloud_webdav_url: Optional[str] = None
    nextcloud_username: Optional[str] = None
    nextcloud_app_password: Optional[str] = None
    nextcloud_folder: Optional[str] = None
    supabase_url: Optional[str] = None
    supabase_key: Optional[str] = None
    sync_interval: Optional[str] = None

def mask_secret(val: str) -> str:
    if not val:
        return ""
    if len(val) <= 4:
        return "••••"
    return val[:2] + "••••••••" + val[-2:]

def get_current_secrets_summary(reveal: bool = False) -> List[Dict[str, Any]]:
    global NEXTCLOUD_URL, NEXTCLOUD_WEBDAV_URL, NEXTCLOUD_USERNAME, NEXTCLOUD_APP_PASSWORD, NEXTCLOUD_FOLDER
    global SUPABASE_URL, SUPABASE_KEY, SYNC_INTERVAL, PORT, HOST

    items = [
        {
            "id": "nc_webdav_url",
            "key": "NEXTCLOUD_WEBDAV_URL",
            "label": "Nextcloud WebDAV Endpoint",
            "category": "Nextcloud",
            "value": NEXTCLOUD_WEBDAV_URL if reveal else mask_secret(NEXTCLOUD_WEBDAV_URL),
            "raw_value": NEXTCLOUD_WEBDAV_URL,
            "is_configured": bool(NEXTCLOUD_WEBDAV_URL),
            "is_secret": False,
            "description": "Full WebDAV endpoint URL (e.g. https://cloud.jhcnexus.space/remote.php/dav/files/truenas_admin/)"
        },
        {
            "id": "nc_url",
            "key": "NEXTCLOUD_URL",
            "label": "Nextcloud Host URL",
            "category": "Nextcloud",
            "value": NEXTCLOUD_URL if reveal else mask_secret(NEXTCLOUD_URL),
            "raw_value": NEXTCLOUD_URL,
            "is_configured": bool(NEXTCLOUD_URL),
            "is_secret": False,
            "description": "Base domain of TrueNAS SCALE Nextcloud instance"
        },
        {
            "id": "nc_username",
            "key": "NEXTCLOUD_USERNAME",
            "label": "Nextcloud Username",
            "category": "Nextcloud",
            "value": NEXTCLOUD_USERNAME if reveal else mask_secret(NEXTCLOUD_USERNAME),
            "raw_value": NEXTCLOUD_USERNAME,
            "is_configured": bool(NEXTCLOUD_USERNAME),
            "is_secret": False,
            "description": "WebDAV authorized user account"
        },
        {
            "id": "nc_app_password",
            "key": "NEXTCLOUD_APP_PASSWORD",
            "label": "Nextcloud App Password",
            "category": "Nextcloud",
            "value": NEXTCLOUD_APP_PASSWORD if reveal else mask_secret(NEXTCLOUD_APP_PASSWORD),
            "raw_value": NEXTCLOUD_APP_PASSWORD,
            "is_configured": bool(NEXTCLOUD_APP_PASSWORD),
            "is_secret": True,
            "description": "Granular application password generated in Nextcloud Security Settings"
        },
        {
            "id": "nc_folder",
            "key": "NEXTCLOUD_FOLDER",
            "label": "Source Folder Path",
            "category": "Nextcloud",
            "value": NEXTCLOUD_FOLDER,
            "raw_value": NEXTCLOUD_FOLDER,
            "is_configured": bool(NEXTCLOUD_FOLDER),
            "is_secret": False,
            "description": "Directory monitored for Excel spreadsheets (e.g. /ExcelImports)"
        },
        {
            "id": "supa_url",
            "key": "SUPABASE_URL",
            "label": "Supabase Project URL",
            "category": "Supabase",
            "value": SUPABASE_URL if reveal else mask_secret(SUPABASE_URL),
            "raw_value": SUPABASE_URL,
            "is_configured": bool(SUPABASE_URL),
            "is_secret": False,
            "description": "PostgreSQL REST API base URL (e.g. https://db.jhcnexus.space)"
        },
        {
            "id": "supa_key",
            "key": "SUPABASE_SERVICE_ROLE_KEY",
            "label": "Supabase Service Role Key",
            "category": "Supabase",
            "value": SUPABASE_KEY if reveal else mask_secret(SUPABASE_KEY),
            "raw_value": SUPABASE_KEY,
            "is_configured": bool(SUPABASE_KEY),
            "is_secret": True,
            "description": "Privileged API key used by background worker for RLS bypass upserts"
        },
        {
            "id": "sync_interval",
            "key": "SYNC_INTERVAL",
            "label": "Polling Schedule Interval",
            "category": "Worker Engine",
            "value": SYNC_INTERVAL,
            "raw_value": SYNC_INTERVAL,
            "is_configured": bool(SYNC_INTERVAL),
            "is_secret": False,
            "description": "Automated WebDAV check frequency (e.g. 5m, 15m, 30m, 1h)"
        }
    ]
    return items

@app.get("/api/secrets")
def api_get_secrets():
    """Returns secrets configuration status and masked values."""
    return {
        "success": True,
        "all_configured": bool(NEXTCLOUD_WEBDAV_URL and NEXTCLOUD_USERNAME and NEXTCLOUD_APP_PASSWORD and SUPABASE_URL and SUPABASE_KEY),
        "secrets": get_current_secrets_summary(reveal=False)
    }

@app.post("/api/secrets/reveal")
def api_reveal_secrets():
    """Returns unmasked secrets for operator verification in visual menu."""
    return {
        "success": True,
        "secrets": get_current_secrets_summary(reveal=True)
    }

@app.post("/api/secrets/update")
def api_update_secrets(req: UpdateSecretsRequest):
    """Dynamically updates worker secrets in memory and re-evaluates clients."""
    global NEXTCLOUD_URL, NEXTCLOUD_WEBDAV_URL, NEXTCLOUD_USERNAME, NEXTCLOUD_APP_PASSWORD, NEXTCLOUD_FOLDER
    global SUPABASE_URL, SUPABASE_KEY, SYNC_INTERVAL

    updated = []
    if req.nextcloud_url is not None:
        NEXTCLOUD_URL = req.nextcloud_url.strip()
        os.environ["NEXTCLOUD_URL"] = NEXTCLOUD_URL
        updated.append("NEXTCLOUD_URL")

    if req.nextcloud_webdav_url is not None:
        NEXTCLOUD_WEBDAV_URL = req.nextcloud_webdav_url.strip()
        os.environ["NEXTCLOUD_WEBDAV_URL"] = NEXTCLOUD_WEBDAV_URL
        updated.append("NEXTCLOUD_WEBDAV_URL")

    if req.nextcloud_username is not None:
        NEXTCLOUD_USERNAME = req.nextcloud_username.strip()
        os.environ["NEXTCLOUD_USERNAME"] = NEXTCLOUD_USERNAME
        updated.append("NEXTCLOUD_USERNAME")

    if req.nextcloud_app_password is not None:
        NEXTCLOUD_APP_PASSWORD = req.nextcloud_app_password.strip()
        os.environ["NEXTCLOUD_APP_PASSWORD"] = NEXTCLOUD_APP_PASSWORD
        updated.append("NEXTCLOUD_APP_PASSWORD")

    if req.nextcloud_folder is not None:
        NEXTCLOUD_FOLDER = "/" + req.nextcloud_folder.strip().lstrip("/")
        os.environ["NEXTCLOUD_FOLDER"] = NEXTCLOUD_FOLDER
        updated.append("NEXTCLOUD_FOLDER")

    if req.supabase_url is not None:
        SUPABASE_URL = req.supabase_url.strip()
        os.environ["SUPABASE_URL"] = SUPABASE_URL
        updated.append("SUPABASE_URL")

    if req.supabase_key is not None:
        SUPABASE_KEY = req.supabase_key.strip()
        os.environ["SUPABASE_SERVICE_ROLE_KEY"] = SUPABASE_KEY
        updated.append("SUPABASE_SERVICE_ROLE_KEY")

    if req.sync_interval is not None and req.sync_interval != SYNC_INTERVAL:
        SYNC_INTERVAL = req.sync_interval.strip()
        os.environ["SYNC_INTERVAL"] = SYNC_INTERVAL
        updated.append("SYNC_INTERVAL")
        # Reschedule background job
        try:
            mins = get_interval_minutes(SYNC_INTERVAL)
            scheduler.reschedule_job("sync_job", trigger=IntervalTrigger(minutes=mins))
            print(f"[Worker] Rescheduled sync job to every {mins} minutes.")
        except Exception as e:
            print(f"[Worker] Could not reschedule job: {e}")

    return {
        "success": True,
        "message": f"Successfully updated {len(updated)} secrets: {', '.join(updated)}",
        "updated_keys": updated,
        "secrets": get_current_secrets_summary(reveal=False)
    }

@app.post("/api/secrets/test")
def api_test_secrets():
    """Live connectivity verification of Nextcloud WebDAV and Supabase database."""
    nc_connected = bool(NEXTCLOUD_WEBDAV_URL and NEXTCLOUD_USERNAME and NEXTCLOUD_APP_PASSWORD)
    supa_connected = bool(SUPABASE_URL and SUPABASE_KEY)

    nc_result = {"success": False, "error": "Nextcloud credentials incomplete"}
    supa_result = {"success": False, "error": "Supabase credentials incomplete"}

    if nc_connected:
        try:
            nc = NextcloudClient(NEXTCLOUD_WEBDAV_URL, NEXTCLOUD_USERNAME, NEXTCLOUD_APP_PASSWORD, NEXTCLOUD_FOLDER)
            nc_result = nc.test_connection()
        except Exception as e:
            nc_result = {"success": False, "error": str(e)}

    if supa_connected:
        try:
            supa = SupabaseSyncClient(SUPABASE_URL, SUPABASE_KEY)
            supa_result = supa.test_connection()
        except Exception as e:
            supa_result = {"success": False, "error": str(e)}

    all_connected = nc_result.get("success", False) and supa_result.get("success", False)

    return {
        "success": True,
        "all_connected": all_connected,
        "nextcloud": nc_result,
        "supabase": supa_result,
        "timestamp": os.getenv("TZ", "UTC")
    }

@app.get("/", response_class=HTMLResponse)
def root_dashboard():
    """Serves a rich, interactive Visual Menu & Secrets Management Console on port 8000."""
    nc_connected = bool(NEXTCLOUD_WEBDAV_URL and NEXTCLOUD_USERNAME and NEXTCLOUD_APP_PASSWORD)
    supa_connected = bool(SUPABASE_URL and SUPABASE_KEY)
    all_ready = nc_connected and supa_connected
    job = scheduler.get_job("sync_job")
    next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S UTC") if job and job.next_run_time else "Not scheduled"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Worker Console & Secrets Vault (Port {PORT})</title>
    <style>
        :root {{
            --bg-canvas: #090d16;
            --bg-card: #131b2e;
            --bg-subtle: #1c2640;
            --border: #263554;
            --text-main: #f1f5f9;
            --text-muted: #94a3b8;
            --emerald: #10b981;
            --emerald-bg: rgba(16, 185, 129, 0.12);
            --amber: #f59e0b;
            --amber-bg: rgba(245, 158, 11, 0.12);
            --rose: #ef4444;
            --rose-bg: rgba(239, 68, 68, 0.12);
            --blue: #38bdf8;
            --blue-bg: rgba(56, 189, 248, 0.12);
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            background: var(--bg-canvas);
            color: var(--text-main);
            min-height: 100vh;
            padding: 24px 16px;
            display: flex;
            justify-content: center;
        }}
        .container {{
            max-width: 960px;
            width: 100%;
            display: flex;
            flex-col;
            flex-direction: column;
            gap: 20px;
        }}
        /* Top Bar */
        .topbar {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 20px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.25);
        }}
        .brand {{
            display: flex;
            align-items: center;
            gap: 14px;
        }}
        .logo-box {{
            width: 44px;
            height: 44px;
            border-radius: 10px;
            background: linear-gradient(135deg, #10b981, #0284c7);
            color: white;
            font-weight: 800;
            font-size: 18px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: monospace;
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        }}
        .brand-text h1 {{
            font-size: 20px;
            font-weight: 700;
            color: #ffffff;
            letter-spacing: -0.02em;
        }}
        .brand-text p {{
            font-size: 13px;
            color: var(--text-muted);
            margin-top: 2px;
        }}
        .status-badge {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 13px;
            font-weight: 600;
        }}
        .status-ready {{
            background: var(--emerald-bg);
            color: #6ee7b7;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }}
        .status-warning {{
            background: var(--amber-bg);
            color: #fde68a;
            border: 1px solid rgba(245, 158, 11, 0.3);
        }}
        .dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
            display: inline-block;
        }}
        .dot-green {{ background: #10b981; box-shadow: 0 0 8px #10b981; }}
        .dot-amber {{ background: #f59e0b; box-shadow: 0 0 8px #f59e0b; }}

        /* Navigation Menu Tabs */
        .nav-tabs {{
            display: flex;
            gap: 8px;
            background: var(--bg-card);
            padding: 6px;
            border: 1px solid var(--border);
            border-radius: 12px;
            overflow-x: auto;
        }}
        .tab-btn {{
            background: transparent;
            color: var(--text-muted);
            border: none;
            padding: 10px 18px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            white-space: nowrap;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .tab-btn:hover {{
            color: #ffffff;
            background: var(--bg-subtle);
        }}
        .tab-btn.active {{
            background: #2563eb;
            color: #ffffff;
            box-shadow: 0 2px 8px rgba(37, 99, 235, 0.4);
        }}

        /* Content Card */
        .card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 24px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.25);
        }}
        .card-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
            padding-bottom: 14px;
            border-bottom: 1px solid var(--border);
            flex-wrap: wrap;
            gap: 12px;
        }}
        .card-title {{
            font-size: 16px;
            font-weight: 700;
            color: #ffffff;
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        /* Button styles */
        .btn {{
            background: #2563eb;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s ease;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            text-decoration: none;
        }}
        .btn:hover {{ background: #1d4ed8; }}
        .btn-emerald {{ background: #059669; }}
        .btn-emerald:hover {{ background: #047857; }}
        .btn-outline {{
            background: transparent;
            border: 1px solid var(--border);
            color: #cbd5e1;
        }}
        .btn-outline:hover {{
            background: var(--bg-subtle);
            border-color: #475569;
            color: white;
        }}
        .btn-sm {{
            padding: 5px 10px;
            font-size: 12px;
            border-radius: 6px;
        }}

        /* Grid */
        .grid-2 {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 16px;
            margin-bottom: 20px;
        }}
        .stat-box {{
            background: var(--bg-subtle);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 16px;
        }}
        .stat-label {{
            font-size: 11px;
            font-weight: 700;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 6px;
        }}
        .stat-val {{
            font-size: 15px;
            font-weight: 600;
            word-break: break-all;
        }}

        /* Alert Banner */
        .alert {{
            padding: 14px 18px;
            border-radius: 10px;
            font-size: 13px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            flex-wrap: wrap;
        }}
        .alert-warning {{
            background: var(--amber-bg);
            border: 1px solid rgba(245, 158, 11, 0.4);
            color: #fde68a;
        }}
        .alert-success {{
            background: var(--emerald-bg);
            border: 1px solid rgba(16, 185, 129, 0.4);
            color: #a7f3d0;
        }}

        /* Secrets Table */
        .secrets-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        .secrets-table th {{
            text-align: left;
            padding: 12px 14px;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-muted);
            background: var(--bg-subtle);
            border-bottom: 1px solid var(--border);
        }}
        .secrets-table td {{
            padding: 14px;
            border-bottom: 1px solid var(--border);
            vertical-align: middle;
        }}
        .secrets-table tr:hover td {{
            background: rgba(255,255,255,0.02);
        }}
        .secret-key {{
            font-family: monospace;
            font-weight: 700;
            font-size: 12px;
            color: #7dd3fc;
        }}
        .secret-val {{
            font-family: monospace;
            font-size: 12px;
            background: #090d16;
            padding: 4px 8px;
            border-radius: 6px;
            border: 1px solid var(--border);
            display: inline-block;
            max-width: 260px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        .badge-pill {{
            padding: 4px 10px;
            border-radius: 9999px;
            font-size: 11px;
            font-weight: 700;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }}
        .badge-green {{ background: var(--emerald-bg); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }}
        .badge-red {{ background: var(--rose-bg); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }}

        /* Modal Overlay */
        .modal {{
            display: none;
            position: fixed;
            inset: 0;
            background: rgba(0,0,0,0.7);
            backdrop-filter: blur(4px);
            z-index: 1000;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }}
        .modal.open {{
            display: flex;
        }}
        .modal-box {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 14px;
            max-width: 540px;
            width: 100%;
            padding: 24px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.5);
        }}
        .form-group {{
            margin-bottom: 16px;
        }}
        .form-label {{
            display: block;
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-muted);
            margin-bottom: 6px;
        }}
        .form-input {{
            width: 100%;
            background: var(--bg-canvas);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 10px 14px;
            color: #ffffff;
            font-family: monospace;
            font-size: 13px;
            outline: none;
            transition: border-color 0.15s;
        }}
        .form-input:focus {{
            border-color: #38bdf8;
            box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.2);
        }}
        .modal-actions {{
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 20px;
        }}
        .toast {{
            position: fixed;
            bottom: 24px;
            right: 24px;
            background: #1e293b;
            border: 1px solid #38bdf8;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.5);
            font-size: 13px;
            z-index: 2000;
            display: none;
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Top Bar Header -->
        <header class="topbar">
            <div class="brand">
                <div class="logo-box">NX</div>
                <div class="brand-text">
                    <h1>Worker Console & Secrets Vault</h1>
                    <p>Coolify Background Array • Port {PORT} • TrueNAS & Supabase Hub</p>
                </div>
            </div>
            <div>
                <span class="status-badge {'status-ready' if all_ready else 'status-warning'}" id="global-status-badge">
                    <span class="dot {'dot-green' if all_ready else 'dot-amber'}"></span>
                    <span>{'All Secrets Connected' if all_ready else 'Secrets Verification Needed'}</span>
                </span>
            </div>
        </header>

        <!-- Visual Menu Navigation Tabs -->
        <nav class="nav-tabs">
            <button class="tab-btn active" onclick="switchTab('secrets')" id="tab-btn-secrets">
                🔑 Secrets & Credentials Vault
            </button>
            <button class="tab-btn" onclick="switchTab('overview')" id="tab-btn-overview">
                📊 Architecture & Diagnostics
            </button>
            <button class="tab-btn" onclick="switchTab('sync')" id="tab-btn-sync">
                ⚡ Manual Sync & Dry Run
            </button>
            <button class="tab-btn" onclick="switchTab('endpoints')" id="tab-btn-endpoints">
                🌐 OpenAPI Endpoints
            </button>
        </nav>

        <!-- Notification Banner if Missing -->
        <div id="missing-alert-box" class="alert alert-warning" style="display: {'none' if all_ready else 'flex'};">
            <div>
                <strong>⚠️ Missing or Unverified Secrets Detected:</strong>
                <span id="missing-alert-text">Configure Nextcloud or Supabase credentials to enable background synchronization.</span>
            </div>
            <button class="btn btn-sm btn-emerald" onclick="openBulkEditModal()">Configure All Secrets Now</button>
        </div>

        <!-- TAB 1: Secrets & Credentials Vault -->
        <section id="panel-secrets" class="card">
            <div class="card-header">
                <div>
                    <h2 class="card-title">System Secrets & Credentials</h2>
                    <p style="font-size: 13px; color: var(--text-muted); margin-top: 4px;">
                        Manage, reveal, edit, and test live connectivity of Nextcloud WebDAV and Supabase keys.
                    </p>
                </div>
                <div style="display: flex; gap: 8px; flex-wrap: wrap;">
                    <button class="btn btn-outline btn-sm" id="btn-toggle-all-reveal" onclick="toggleRevealAll()">
                        👁️ <span id="reveal-btn-text">Reveal All Secrets</span>
                    </button>
                    <button class="btn btn-outline btn-sm" onclick="runLiveConnectionTest()" id="btn-test-secrets">
                        🔄 Test Live Connections
                    </button>
                    <button class="btn btn-emerald btn-sm" onclick="openBulkEditModal()">
                        ✏️ Edit / Change Secrets
                    </button>
                </div>
            </div>

            <!-- Secrets Table -->
            <div style="overflow-x: auto;">
                <table class="secrets-table">
                    <thead>
                        <tr>
                            <th>Environment Key</th>
                            <th>Description</th>
                            <th>Current Value</th>
                            <th>Status</th>
                            <th style="text-align: right;">Action</th>
                        </tr>
                    </thead>
                    <tbody id="secrets-tbody">
                        <!-- Populated by JavaScript -->
                    </tbody>
                </table>
            </div>

            <!-- Test Results Box -->
            <div id="test-results-panel" style="display: none; margin-top: 20px; background: var(--bg-subtle); border: 1px solid var(--border); border-radius: 10px; padding: 16px;">
                <div style="font-size: 13px; font-weight: 700; margin-bottom: 10px; color: #38bdf8;">
                    Live Diagnostic Test Feedback:
                </div>
                <div id="test-results-content" style="font-family: monospace; font-size: 12px; line-height: 1.6;"></div>
            </div>
        </section>

        <!-- TAB 2: Architecture & Diagnostics -->
        <section id="panel-overview" class="card" style="display: none;">
            <div class="card-header">
                <h2 class="card-title">Architecture & System Matrix</h2>
                <span style="font-size: 12px; color: var(--text-muted);">TrueNAS SCALE ➔ Coolify ➔ Supabase</span>
            </div>

            <div class="grid-2">
                <div class="stat-box">
                    <div class="stat-label">Nextcloud WebDAV Engine</div>
                    <div class="stat-val" style="color: {'#34d399' if nc_connected else '#f87171'};">
                        {('Active (' + NEXTCLOUD_FOLDER + ')') if nc_connected else 'Credentials Missing'}
                    </div>
                    <div style="font-size: 12px; color: var(--text-muted); margin-top: 4px;">
                        {NEXTCLOUD_URL or 'https://cloud.jhcnexus.space'}
                    </div>
                </div>

                <div class="stat-box">
                    <div class="stat-label">Supabase PostgreSQL</div>
                    <div class="stat-val" style="color: {'#34d399' if supa_connected else '#f87171'};">
                        {'Connected & RLS Ready' if supa_connected else 'Credentials Missing'}
                    </div>
                    <div style="font-size: 12px; color: var(--text-muted); margin-top: 4px;">
                        {SUPABASE_URL or 'https://db.jhcnexus.space'}
                    </div>
                </div>

                <div class="stat-box">
                    <div class="stat-label">APScheduler Polling Interval</div>
                    <div class="stat-val" style="color: #38bdf8;">Every {SYNC_INTERVAL}</div>
                    <div style="font-size: 12px; color: var(--text-muted); margin-top: 4px;">
                        Next run: <span id="scheduler-next-run">{next_run}</span>
                    </div>
                </div>

                <div class="stat-box">
                    <div class="stat-label">Worker Daemon Port</div>
                    <div class="stat-val">0.0.0.0:{PORT}</div>
                    <div style="font-size: 12px; color: var(--text-muted); margin-top: 4px;">
                        FastAPI • OpenPyXL • Pandas • Supabase-py
                    </div>
                </div>
            </div>
        </section>

        <!-- TAB 3: Manual Sync & Dry Run -->
        <section id="panel-sync" class="card" style="display: none;">
            <div class="card-header">
                <h2 class="card-title">Manual Synchronization Trigger</h2>
                <span style="font-size: 12px; color: var(--text-muted);">Execute on-demand ingest cycle</span>
            </div>

            <p style="font-size: 13px; color: var(--text-muted); margin-bottom: 20px;">
                Trigger a complete synchronization pass immediately, or run a simulated Dry Run to evaluate transformed rows without committing to Supabase PostgreSQL.
            </p>

            <div style="display: flex; gap: 12px; flex-wrap: wrap;">
                <button class="btn btn-emerald" onclick="triggerWorkerSync(false)" id="btn-run-sync">
                    ⚡ Execute Production Sync Pass
                </button>
                <button class="btn btn-outline" onclick="triggerWorkerSync(true)" id="btn-run-dry">
                    🧪 Run Safe Dry Run Simulation
                </button>
            </div>

            <div id="sync-output-box" style="display: none; margin-top: 20px;">
                <div style="font-size: 12px; font-weight: 700; color: #94a3b8; margin-bottom: 6px;">EXECUTION OUTPUT:</div>
                <pre id="sync-output-pre" style="background: #090d16; padding: 14px; border-radius: 8px; border: 1px solid var(--border); font-size: 12px; color: #a5f3fc; overflow-x: auto; max-height: 250px;"></pre>
            </div>
        </section>

        <!-- TAB 4: OpenAPI Endpoints -->
        <section id="panel-endpoints" class="card" style="display: none;">
            <div class="card-header">
                <h2 class="card-title">FastAPI Worker Endpoints</h2>
            </div>
            <p style="font-size: 13px; color: var(--text-muted); margin-bottom: 16px;">
                Interactive documentation and direct API routes:
            </p>
            <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                <a href="/docs" class="btn" target="_blank">Open Swagger API Docs (/docs)</a>
                <a href="/status" class="btn btn-outline" target="_blank">Status JSON (/status)</a>
                <a href="/health" class="btn btn-outline" target="_blank">Healthcheck (/health)</a>
                <a href="/api/secrets" class="btn btn-outline" target="_blank">Secrets API (/api/secrets)</a>
            </div>
        </section>
    </div>

    <!-- Edit Secret Modal -->
    <div id="edit-modal" class="modal">
        <div class="modal-box">
            <h3 style="font-size: 18px; font-weight: 700; margin-bottom: 16px; color: #ffffff;" id="modal-title">
                Edit Secret
            </h3>
            <form id="edit-form" onsubmit="saveSecretForm(event)">
                <div class="form-group">
                    <label class="form-label" id="edit-field-label">Secret Value</label>
                    <input type="text" class="form-input" id="edit-field-input" placeholder="Enter new value..." required />
                    <input type="hidden" id="edit-field-key" />
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="closeModal()">Cancel</button>
                    <button type="submit" class="btn btn-emerald" id="btn-modal-save">Save & Apply</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Bulk Edit Modal -->
    <div id="bulk-edit-modal" class="modal">
        <div class="modal-box" style="max-width: 680px;">
            <h3 style="font-size: 18px; font-weight: 700; margin-bottom: 16px; color: #ffffff;">
                Configure All System Secrets & Credentials
            </h3>
            <form id="bulk-edit-form" onsubmit="saveBulkSecretsForm(event)">
                <div class="grid-2">
                    <div class="form-group">
                        <label class="form-label">Nextcloud Host URL</label>
                        <input type="text" class="form-input" id="bulk-nc-url" placeholder="https://cloud.jhcnexus.space" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Nextcloud WebDAV URL</label>
                        <input type="text" class="form-input" id="bulk-nc-webdav" placeholder="https://cloud.jhcnexus.space/remote.php/dav/files/user/" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Nextcloud Username</label>
                        <input type="text" class="form-input" id="bulk-nc-user" placeholder="truenas_admin" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Nextcloud App Password</label>
                        <input type="password" class="form-input" id="bulk-nc-pass" placeholder="mpxC4-dk7jn-4GYCH-..." />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Nextcloud Folder</label>
                        <input type="text" class="form-input" id="bulk-nc-folder" placeholder="/ExcelImports" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Sync Interval</label>
                        <select class="form-input" id="bulk-sync-interval">
                            <option value="5m">Every 5 minutes</option>
                            <option value="15m" selected>Every 15 minutes</option>
                            <option value="30m">Every 30 minutes</option>
                            <option value="1h">Hourly</option>
                        </select>
                    </div>
                </div>

                <div class="form-group">
                    <label class="form-label">Supabase Project URL</label>
                    <input type="text" class="form-input" id="bulk-supa-url" placeholder="https://db.jhcnexus.space" />
                </div>

                <div class="form-group">
                    <label class="form-label">Supabase Service Role Key</label>
                    <input type="password" class="form-input" id="bulk-supa-key" placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6..." />
                </div>

                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="closeBulkModal()">Cancel</button>
                    <button type="submit" class="btn btn-emerald" id="btn-bulk-save">Save & Apply All Secrets</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Toast Notification -->
    <div id="toast" class="toast"></div>

    <script>
        let currentSecrets = [];
        let isRevealedAll = false;
        let revealedSecretsMap = {{}};

        function showToast(msg, isSuccess = true) {{
            const t = document.getElementById('toast');
            t.innerText = msg;
            t.style.borderColor = isSuccess ? '#10b981' : '#ef4444';
            t.style.display = 'block';
            setTimeout(() => {{ t.style.display = 'none'; }}, 3500);
        }}

        function switchTab(tab) {{
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.getElementById('tab-btn-' + tab).classList.add('active');

            document.getElementById('panel-secrets').style.display = tab === 'secrets' ? 'block' : 'none';
            document.getElementById('panel-overview').style.display = tab === 'overview' ? 'block' : 'none';
            document.getElementById('panel-sync').style.display = tab === 'sync' ? 'block' : 'none';
            document.getElementById('panel-endpoints').style.display = tab === 'endpoints' ? 'block' : 'none';
        }}

        async function fetchSecrets() {{
            try {{
                const res = await fetch('/api/secrets');
                const data = await res.json();
                if (data.success) {{
                    currentSecrets = data.secrets;
                    renderSecretsTable();
                }}
            }} catch (e) {{
                console.error('Failed to load secrets:', e);
            }}
        }}

        function renderSecretsTable() {{
            const tbody = document.getElementById('secrets-tbody');
            tbody.innerHTML = '';

            let missingCount = 0;

            currentSecrets.forEach(sec => {{
                if (!sec.is_configured) missingCount++;
                const isRevealed = isRevealedAll || revealedSecretsMap[sec.key];
                const displayVal = isRevealed ? (sec.raw_value || 'Not configured') : sec.value;

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>
                        <div class="secret-key">${{sec.key}}</div>
                        <div style="font-size: 11px; color: var(--text-muted);">${{sec.label}}</div>
                    </td>
                    <td style="color: var(--text-muted); font-size: 12px; max-width: 240px;">
                        ${{sec.description}}
                    </td>
                    <td>
                        <span class="secret-val" id="val-${{sec.key}}" title="${{isRevealed ? (sec.raw_value || '') : 'Masked for security'}}">
                            ${{displayVal || '<span style="color: #f87171;">Not configured</span>'}}
                        </span>
                    </td>
                    <td>
                        <span class="badge-pill ${{sec.is_configured ? 'badge-green' : 'badge-red'}}">
                            ${{sec.is_configured ? '● Configured' : '○ Missing'}}
                        </span>
                    </td>
                    <td style="text-align: right; white-space: nowrap;">
                        <button class="btn btn-outline btn-sm" onclick="toggleSingleSecret('${{sec.key}}')" title="View or mask secret">
                            ${{isRevealed ? '🔒' : '👁️'}}
                        </button>
                        <button class="btn btn-sm btn-outline" style="margin-left: 4px;" onclick="openSingleEditModal('${{sec.key}}', '${{sec.label}}')">
                            ✏️ Edit
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            }});

            const alertBox = document.getElementById('missing-alert-box');
            if (missingCount > 0) {{
                alertBox.style.display = 'flex';
                document.getElementById('missing-alert-text').innerText = `${{missingCount}} secret(s) not configured. Background synchronization will be skipped until configured.`;
            }} else {{
                alertBox.style.display = 'none';
            }}
        }}

        async function toggleRevealAll() {{
            isRevealedAll = !isRevealedAll;
            const btnText = document.getElementById('reveal-btn-text');

            if (isRevealedAll) {{
                try {{
                    const res = await fetch('/api/secrets/reveal', {{ method: 'POST' }});
                    const data = await res.json();
                    if (data.success) {{
                        currentSecrets = data.secrets;
                        btnText.innerText = 'Mask All Secrets';
                    }}
                }} catch (e) {{
                    showToast('Failed to reveal secrets: ' + e.message, false);
                }}
            }} else {{
                btnText.innerText = 'Reveal All Secrets';
                await fetchSecrets();
            }}
            renderSecretsTable();
        }}

        function toggleSingleSecret(key) {{
            revealedSecretsMap[key] = !revealedSecretsMap[key];
            renderSecretsTable();
        }}

        function openSingleEditModal(key, label) {{
            document.getElementById('modal-title').innerText = `Edit ${{label}}`;
            document.getElementById('edit-field-label').innerText = key;
            document.getElementById('edit-field-key').value = key;
            
            const sec = currentSecrets.find(s => s.key === key);
            document.getElementById('edit-field-input').value = sec?.raw_value || '';
            document.getElementById('edit-modal').classList.add('open');
        }}

        function closeModal() {{
            document.getElementById('edit-modal').classList.remove('open');
        }}

        async function saveSecretForm(e) {{
            e.preventDefault();
            const key = document.getElementById('edit-field-key').value;
            const val = document.getElementById('edit-field-input').value;

            const payload = {{}};
            if (key === 'NEXTCLOUD_URL') payload.nextcloud_url = val;
            else if (key === 'NEXTCLOUD_WEBDAV_URL') payload.nextcloud_webdav_url = val;
            else if (key === 'NEXTCLOUD_USERNAME') payload.nextcloud_username = val;
            else if (key === 'NEXTCLOUD_APP_PASSWORD') payload.nextcloud_app_password = val;
            else if (key === 'NEXTCLOUD_FOLDER') payload.nextcloud_folder = val;
            else if (key === 'SUPABASE_URL') payload.supabase_url = val;
            else if (key === 'SUPABASE_SERVICE_ROLE_KEY') payload.supabase_key = val;
            else if (key === 'SYNC_INTERVAL') payload.sync_interval = val;

            try {{
                const res = await fetch('/api/secrets/update', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify(payload)
                }});
                const data = await res.json();
                if (data.success) {{
                    showToast(`Updated ${{key}} successfully!`);
                    closeModal();
                    await fetchSecrets();
                }} else {{
                    showToast(data.message || 'Update failed', false);
                }}
            }} catch (err) {{
                showToast('Save error: ' + err.message, false);
            }}
        }}

        function openBulkEditModal() {{
            const getVal = (k) => (currentSecrets.find(s => s.key === k)?.raw_value || '');
            document.getElementById('bulk-nc-url').value = getVal('NEXTCLOUD_URL');
            document.getElementById('bulk-nc-webdav').value = getVal('NEXTCLOUD_WEBDAV_URL');
            document.getElementById('bulk-nc-user').value = getVal('NEXTCLOUD_USERNAME');
            document.getElementById('bulk-nc-pass').value = getVal('NEXTCLOUD_APP_PASSWORD');
            document.getElementById('bulk-nc-folder').value = getVal('NEXTCLOUD_FOLDER') || '/ExcelImports';
            document.getElementById('bulk-supa-url').value = getVal('SUPABASE_URL');
            document.getElementById('bulk-supa-key').value = getVal('SUPABASE_SERVICE_ROLE_KEY');
            document.getElementById('bulk-sync-interval').value = getVal('SYNC_INTERVAL') || '15m';

            document.getElementById('bulk-edit-modal').classList.add('open');
        }}

        function closeBulkModal() {{
            document.getElementById('bulk-edit-modal').classList.remove('open');
        }}

        async function saveBulkSecretsForm(e) {{
            e.preventDefault();
            const payload = {{
                nextcloud_url: document.getElementById('bulk-nc-url').value,
                nextcloud_webdav_url: document.getElementById('bulk-nc-webdav').value,
                nextcloud_username: document.getElementById('bulk-nc-user').value,
                nextcloud_app_password: document.getElementById('bulk-nc-pass').value,
                nextcloud_folder: document.getElementById('bulk-nc-folder').value,
                supabase_url: document.getElementById('bulk-supa-url').value,
                supabase_key: document.getElementById('bulk-supa-key').value,
                sync_interval: document.getElementById('bulk-sync-interval').value
            }};

            try {{
                const res = await fetch('/api/secrets/update', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify(payload)
                }});
                const data = await res.json();
                if (data.success) {{
                    showToast('All secrets saved and reloaded!');
                    closeBulkModal();
                    await fetchSecrets();
                    runLiveConnectionTest();
                }} else {{
                    showToast(data.message || 'Update failed', false);
                }}
            }} catch (err) {{
                showToast('Save error: ' + err.message, false);
            }}
        }}

        async function runLiveConnectionTest() {{
            const panel = document.getElementById('test-results-panel');
            const content = document.getElementById('test-results-content');
            panel.style.display = 'block';
            content.innerHTML = '<span style="color: #94a3b8;">Testing Nextcloud WebDAV and Supabase PostgreSQL live...</span>';

            try {{
                const res = await fetch('/api/secrets/test', {{ method: 'POST' }});
                const data = await res.json();

                let html = '';
                const ncOk = data.nextcloud?.success;
                const supaOk = data.supabase?.success;

                html += `<div style="color: ${{ncOk ? '#34d399' : '#f87171'}};">`;
                html += `Nextcloud WebDAV: ${{ncOk ? '✔ CONNECTED & VERIFIED' : '✖ FAILED (' + (data.nextcloud?.error || 'Unreachable') + ')'}}`;
                html += `</div>`;

                html += `<div style="color: ${{supaOk ? '#34d399' : '#f87171'}}; margin-top: 4px;">`;
                html += `Supabase Database: ${{supaOk ? '✔ CONNECTED & TABLES VERIFIED' : '✖ FAILED (' + (data.supabase?.error || 'Auth failure') + ')'}}`;
                html += `</div>`;

                if (data.all_connected) {{
                    html += `<div style="margin-top: 8px; color: #a7f3d0; font-weight: bold;">All systems operational. Scheduled background sync ready!</div>`;
                    showToast('All secrets verified successfully!');
                }} else {{
                    html += `<div style="margin-top: 8px; color: #fde68a;">Some connections failed. Use "Edit / Change Secrets" above to correct credentials.</div>`;
                    showToast('Connection test failed for 1 or more secrets', false);
                }}

                content.innerHTML = html;
            }} catch (err) {{
                content.innerHTML = `<span style="color: #f87171;">Test execution failed: ${{err.message}}</span>`;
            }}
        }}

        async function triggerWorkerSync(dryRun) {{
            const outBox = document.getElementById('sync-output-box');
            const outPre = document.getElementById('sync-output-pre');
            outBox.style.display = 'block';
            outPre.innerText = dryRun ? 'Executing Dry Run simulation...' : 'Triggering production synchronization cycle...';

            const endpoint = dryRun ? '/sync/dry-run' : '/sync/trigger';
            try {{
                const res = await fetch(endpoint, {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify({{ dry_run: dryRun }})
                }});
                const data = await res.json();
                outPre.innerText = JSON.stringify(data, null, 2);
                showToast(dryRun ? 'Dry run completed!' : 'Sync pass dispatched!');
            }} catch (err) {{
                outPre.innerText = 'Sync error: ' + err.message;
                showToast('Sync error: ' + err.message, false);
            }}
        }}

        // Initialize on load
        fetchSecrets();
    </script>
</body>
</html>"""
    return HTMLResponse(content=html)

@app.get("/health")
def health_check():
    """Coolify health check endpoint."""
    return {"status": "healthy"}

@app.get("/status")
def get_worker_status():
    """Returns worker status, interval, and next scheduled run."""
    job = scheduler.get_job("sync_job")
    next_run = job.next_run_time.isoformat() if job and job.next_run_time else None
    return {
        "status": "active" if scheduler.running else "idle",
        "sync_interval": SYNC_INTERVAL,
        "next_run": next_run,
        "nextcloud_folder": NEXTCLOUD_FOLDER
    }

@app.post("/sync/trigger")
def trigger_sync(req: TriggerRequest, background_tasks: BackgroundTasks):
    """Manually triggers synchronization in background."""
    background_tasks.add_task(run_synchronization_cycle, req.dry_run, req.filename)
    return {"message": "Sync cycle dispatched", "dry_run": req.dry_run, "filename": req.filename}

@app.post("/sync/dry-run")
def trigger_dry_run(req: TriggerRequest):
    """Synchronous Dry Run for instant UI preview."""
    res = run_synchronization_cycle(dry_run=True, specific_file=req.filename)
    return res

@app.post("/test-connections")
def test_connections():
    """Tests Nextcloud WebDAV and Supabase connectivity."""
    nc = NextcloudClient(NEXTCLOUD_WEBDAV_URL, NEXTCLOUD_USERNAME, NEXTCLOUD_APP_PASSWORD, NEXTCLOUD_FOLDER)
    nc_result = nc.test_connection()

    supa = SupabaseSyncClient(SUPABASE_URL, SUPABASE_KEY)
    supa_result = supa.test_connection()

    return {
        "nextcloud": nc_result,
        "supabase": supa_result
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=HOST, port=PORT)
