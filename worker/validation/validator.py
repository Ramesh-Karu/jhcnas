"""
Data Validation Engine.
Ensures records meet all constraints before writing to Supabase.
Never silently discards invalid records; captures detailed row numbers, fields, and error messages.
"""
import re
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple

class DataValidator:
    @staticmethod
    def validate_records(
        records: List[Dict[str, Any]],
        col_mappings: List[Dict[str, Any]],
        worksheet_name: str
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        Validates all records for a worksheet.
        Returns:
            - valid_records: List of valid dicts
            - failed_records: List of failed raw dicts with error annotations
            - errors: List of detailed error objects for import_errors table
        """
        valid_records = []
        failed_records = []
        all_errors = []

        # Track unique keys seen in this batch to prevent intra-file duplicates
        unique_keys = [cm["supabase_column"] for cm in col_mappings if cm.get("unique_key")]
        seen_unique_values: Dict[str, set] = {k: set() for k in unique_keys}

        for idx, rec in enumerate(records):
            row_num = rec.get("__row_number", idx + 1)
            row_errors = []

            for cm in col_mappings:
                col_name = cm.get("supabase_column")
                excel_col = cm.get("excel_column")
                data_type = cm.get("data_type", "text")
                required = cm.get("required", False)
                is_unique = cm.get("unique_key", False)
                val = rec.get(col_name)

                # 1. Required field check
                if required and (val is None or (isinstance(val, str) and val.strip() == "")):
                    row_errors.append({
                        "worksheet_name": worksheet_name,
                        "row_number": row_num,
                        "excel_column": excel_col,
                        "column_name": col_name,
                        "raw_value": str(val),
                        "error_message": f"Required field '{col_name}' is missing or empty.",
                        "error_type": "missing_required"
                    })
                    continue

                if val is None:
                    continue

                # 2. Data Type checks
                val_str = str(val).strip()

                if data_type == "integer":
                    try:
                        int(val)
                    except (ValueError, TypeError):
                        row_errors.append({
                            "worksheet_name": worksheet_name,
                            "row_number": row_num,
                            "excel_column": excel_col,
                            "column_name": col_name,
                            "raw_value": val_str,
                            "error_message": f"Value '{val_str}' is not a valid integer.",
                            "error_type": "type_mismatch"
                        })

                elif data_type in ("decimal", "numeric", "float"):
                    try:
                        float(val)
                    except (ValueError, TypeError):
                        row_errors.append({
                            "worksheet_name": worksheet_name,
                            "row_number": row_num,
                            "excel_column": excel_col,
                            "column_name": col_name,
                            "raw_value": val_str,
                            "error_message": f"Value '{val_str}' is not a valid decimal number.",
                            "error_type": "type_mismatch"
                        })

                elif data_type == "date":
                    if not isinstance(val, (datetime, date)):
                        if not re.match(r"^\d{4}-\d{2}-\d{2}$", val_str):
                            row_errors.append({
                                "worksheet_name": worksheet_name,
                                "row_number": row_num,
                                "excel_column": excel_col,
                                "column_name": col_name,
                                "raw_value": val_str,
                                "error_message": f"Invalid date format '{val_str}'. Expected YYYY-MM-DD.",
                                "error_type": "invalid_date"
                            })

                elif data_type == "boolean":
                    if not isinstance(val, bool) and str(val).lower() not in ("true", "false", "1", "0"):
                        row_errors.append({
                            "worksheet_name": worksheet_name,
                            "row_number": row_num,
                            "excel_column": excel_col,
                            "column_name": col_name,
                            "raw_value": val_str,
                            "error_message": f"Value '{val_str}' cannot be interpreted as boolean.",
                            "error_type": "type_mismatch"
                        })

                # 3. In-batch Duplicate Key Check
                if is_unique and val is not None:
                    u_key = str(val).strip()
                    if u_key in seen_unique_values[col_name]:
                        row_errors.append({
                            "worksheet_name": worksheet_name,
                            "row_number": row_num,
                            "excel_column": excel_col,
                            "column_name": col_name,
                            "raw_value": val_str,
                            "error_message": f"Duplicate key value '{val_str}' detected in batch at row {row_num}.",
                            "error_type": "duplicate"
                        })
                    else:
                        seen_unique_values[col_name].add(u_key)

            if row_errors:
                all_errors.extend(row_errors)
                failed_records.append({"row": rec, "errors": row_errors})
            else:
                # Strip internal metadata before database insertion
                clean_rec = {k: v for k, v in rec.items() if not k.startswith("__")}
                valid_records.append(clean_rec)

        return valid_records, failed_records, all_errors
