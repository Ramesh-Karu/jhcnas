"""
Import Logging and Error Auditing Module.
Maintains tamper-proof sync history in Supabase import_logs and import_errors tables.
"""
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
from supabase import Client

class SyncLogger:
    def __init__(self, supabase_client: Client):
        self.supabase = supabase_client

    def create_import_log(
        self,
        filename: str,
        file_path: str,
        file_hash: str,
        workbook_id: Optional[str] = None,
        is_dry_run: bool = False
    ) -> str:
        """Initializes a new import log entry with 'Processing' status."""
        data = {
            "workbook_id": workbook_id,
            "filename": filename,
            "file_path": file_path,
            "file_hash": file_hash,
            "status": "Processing",
            "is_dry_run": is_dry_run,
            "started_at": datetime.now(timezone.utc).isoformat(),
            "rows_processed": 0,
            "rows_inserted": 0,
            "rows_updated": 0,
            "rows_failed": 0
        }
        res = self.supabase.table("import_logs").insert(data).execute()
        if res.data and len(res.data) > 0:
            return res.data[0]["id"]
        return ""

    def finalize_import_log(
        self,
        log_id: str,
        status: str, # Success, Partial Success, Failed, Skipped
        number_of_worksheets: int,
        rows_processed: int,
        rows_inserted: int,
        rows_updated: int,
        rows_failed: int,
        error_summary: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ) -> None:
        """Completes an import log entry with final tallies and status."""
        update_data = {
            "status": status,
            "number_of_worksheets": number_of_worksheets,
            "rows_processed": rows_processed,
            "rows_inserted": rows_inserted,
            "rows_updated": rows_updated,
            "rows_failed": rows_failed,
            "completed_at": datetime.now(timezone.utc).isoformat(),
            "error_summary": error_summary,
            "details": details or {}
        }
        self.supabase.table("import_logs").update(update_data).eq("id", log_id).execute()

    def record_row_errors(self, log_id: str, errors: List[Dict[str, Any]]) -> None:
        """Inserts detailed row validation errors into import_errors."""
        if not errors or not log_id:
            return

        payload = []
        for err in errors:
            payload.append({
                "import_log_id": log_id,
                "worksheet_name": err.get("worksheet_name", ""),
                "row_number": err.get("row_number", 0),
                "excel_column": err.get("excel_column", ""),
                "column_name": err.get("column_name", ""),
                "raw_value": str(err.get("raw_value", ""))[:500], # truncate long strings
                "error_message": err.get("error_message", ""),
                "error_type": err.get("error_type", "validation")
            })

        # Batch insert errors in chunks of 200
        for i in range(0, len(payload), 200):
            chunk = payload[i:i+200]
            try:
                self.supabase.table("import_errors").insert(chunk).execute()
            except Exception as e:
                print(f"[SyncLogger] Failed to write import errors batch: {e}")
