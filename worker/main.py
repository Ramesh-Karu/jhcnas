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
from pydantic import BaseModel
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

from nextcloud.client import NextcloudClient
from excel.parser import ExcelParser
from mappings.manager import MappingManager
from validation.validator import DataValidator
from database.client import SupabaseSyncClient
from logging.logger import SyncLogger

# Environment Configuration
PORT = int(os.getenv("PORT", 8000))
HOST = os.getenv("HOST", "0.0.0.0")

NEXTCLOUD_URL = os.getenv("NEXTCLOUD_URL", "")
NEXTCLOUD_WEBDAV_URL = os.getenv("NEXTCLOUD_WEBDAV_URL", "")
NEXTCLOUD_USERNAME = os.getenv("NEXTCLOUD_USERNAME", "")
NEXTCLOUD_APP_PASSWORD = os.getenv("NEXTCLOUD_APP_PASSWORD", "")
NEXTCLOUD_FOLDER = os.getenv("NEXTCLOUD_FOLDER", "/ExcelImports")

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "")
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
    supa = SupabaseSyncClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)
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
                if wb_config and wb_config.get("worksheet_mappings"):
                    for wm in wb_config["worksheet_mappings"]:
                        if wm.get("enabled", True):
                            active_mappings[wm["worksheet_name"]] = wm

                # Step 6: Process Each Worksheet
                for ws_info in analysis["worksheets"]:
                    sheet_name = ws_info["sheet_name"]
                    wm = active_mappings.get(sheet_name)

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

    supa = SupabaseSyncClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)
    supa_result = supa.test_connection()

    return {
        "nextcloud": nc_result,
        "supabase": supa_result
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=HOST, port=PORT)
