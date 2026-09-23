from __future__ import annotations
"""
Supabase PostgreSQL Client with Upsert, Dry Run, and Schema Management.
Uses Supabase Service Role credentials securely in the background worker.
"""
import os
from typing import Any, Dict, List, Optional, Tuple
from supabase import create_client, Client

class SupabaseSyncClient:
    def __init__(self, supabase_url: str, service_role_key: str):
        self.supabase_url = supabase_url.rstrip("/")
        self.service_role_key = service_role_key
        self.client: Client = create_client(self.supabase_url, self.service_role_key)

    def test_connection(self) -> Dict[str, Any]:
        """Verifies database connectivity and configuration tables accessibility."""
        try:
            res = self.client.table("sync_settings").select("id").limit(1).execute()
            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_workbook_config(self, nextcloud_path: str) -> Optional[Dict[str, Any]]:
        """Retrieves workbook configuration and mappings for a Nextcloud file path."""
        try:
            wb_res = self.client.table("workbooks") \
                .select("*, worksheet_mappings(*, column_mappings(*))") \
                .eq("nextcloud_path", nextcloud_path) \
                .eq("enabled", True) \
                .execute()

            if wb_res.data and len(wb_res.data) > 0:
                return wb_res.data[0]
            return None
        except Exception as e:
            print(f"[Supabase] Error fetching workbook config: {e}")
            return None

    def check_file_unchanged(self, file_path: str, sha256_hash: str) -> bool:
        """
        File Change Detection (Section 14):
        Checks if file was previously imported with the exact same SHA-256 hash.
        """
        try:
            res = self.client.table("workbooks") \
                .select("file_hash") \
                .eq("nextcloud_path", file_path) \
                .execute()

            if res.data and len(res.data) > 0:
                stored_hash = res.data[0].get("file_hash")
                if stored_hash == sha256_hash:
                    return True # Unchanged
            return False
        except Exception as e:
            print(f"[Supabase] Error checking file hash: {e}")
            return False

    def update_workbook_hash(self, file_path: str, file_name: str, sha256_hash: str, file_size: int) -> None:
        """Updates stored hash and size after successful processing."""
        try:
            self.client.table("workbooks").upsert({
                "nextcloud_path": file_path,
                "name": file_name,
                "file_hash": sha256_hash,
                "file_size": file_size,
                "enabled": True
            }, on_conflict="nextcloud_path").execute()
        except Exception as e:
            print(f"[Supabase] Error updating workbook hash: {e}")

    def execute_dry_run_analysis(
        self,
        table_name: str,
        unique_key_cols: List[str],
        valid_records: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Dry Run Comparison:
        Queries Supabase for existing keys to determine proposed inserts vs updates
        WITHOUT modifying the database.
        """
        if not valid_records:
            return {"proposed_inserts": 0, "proposed_updates": 0, "existing_keys_count": 0}

        if not unique_key_cols:
            # If no unique key is designated, all valid rows would be treated as inserts
            return {"proposed_inserts": len(valid_records), "proposed_updates": 0, "existing_keys_count": 0}

        primary_key = unique_key_cols[0]
        keys_to_check = [r[primary_key] for r in valid_records if primary_key in r and r[primary_key] is not None]

        existing_keys = set()
        try:
            # Batch check in chunks of 500
            for i in range(0, len(keys_to_check), 500):
                chunk = keys_to_check[i:i+500]
                resp = self.client.table(table_name).select(primary_key).in_(primary_key, chunk).execute()
                for item in resp.data:
                    existing_keys.add(item[primary_key])
        except Exception as e:
            print(f"[Supabase] Dry run check error on table {table_name}: {e}")

        inserts = 0
        updates = 0
        for r in valid_records:
            if r.get(primary_key) in existing_keys:
                updates += 1
            else:
                inserts += 1

        return {
            "proposed_inserts": inserts,
            "proposed_updates": updates,
            "existing_keys_count": len(existing_keys)
        }

    def upsert_records(
        self,
        table_name: str,
        records: List[Dict[str, Any]],
        unique_key_cols: List[str]
    ) -> Tuple[int, int, Optional[str]]:
        """
        Performs idempotent upsert into Supabase table.
        Returns: (inserted_count, updated_count, error)
        """
        if not records:
            return 0, 0, None

        conflict_clause = ",".join(unique_key_cols) if unique_key_cols else None

        # To track exact inserts vs updates, query existing keys first
        existing_keys = set()
        if unique_key_cols:
            pk = unique_key_cols[0]
            batch_keys = [r[pk] for r in records if pk in r and r[pk] is not None]
            try:
                for i in range(0, len(batch_keys), 500):
                    c = batch_keys[i:i+500]
                    res = self.client.table(table_name).select(pk).in_(pk, c).execute()
                    for item in res.data:
                        existing_keys.add(item[pk])
            except Exception as e:
                print(f"[Supabase] Warning fetching existing keys: {e}")

        inserted = 0
        updated = 0

        # Upsert in batches of 200
        batch_size = 200
        for i in range(0, len(records), batch_size):
            batch = records[i:i+batch_size]
            try:
                if conflict_clause:
                    self.client.table(table_name).upsert(batch, on_conflict=conflict_clause).execute()
                else:
                    self.client.table(table_name).insert(batch).execute()

                # Tally inserts vs updates
                for r in batch:
                    if unique_key_cols and r.get(unique_key_cols[0]) in existing_keys:
                        updated += 1
                    else:
                        inserted += 1

            except Exception as e:
                return inserted, updated, str(e)

        return inserted, updated, None
