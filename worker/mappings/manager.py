from __future__ import annotations
"""
Mapping and Transformation Engine.
Reads dynamic mapping configurations from Supabase and transforms raw Excel rows into structured database records.
"""
import re
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple

class MappingManager:
    @staticmethod
    def apply_transformations(
        raw_val: Any,
        transformation: str,
        default_val: Optional[str] = None
    ) -> Any:
        """
        Applies configured transformation rules:
        - Trim whitespace
        - Uppercase / Lowercase
        - Parse Dates (ISO format YYYY-MM-DD)
        - Parse Numbers (int/float)
        - Yes/No to Boolean
        - Attendance P/A conversion (P->Present, A->Absent, L->Late)
        - Phone normalization
        - ID normalization
        """
        if raw_val is None or (isinstance(raw_val, str) and raw_val.strip() == ""):
            return default_val if default_val is not None else None

        val_str = str(raw_val).strip()

        if transformation == "trim":
            return val_str

        elif transformation == "uppercase":
            return val_str.upper()

        elif transformation == "lowercase":
            return val_str.lower()

        elif transformation == "parse_date":
            # If already datetime/date object from openpyxl
            if isinstance(raw_val, (datetime, date)):
                return raw_val.strftime("%Y-%m-%d")
            # Try parsing various date formats
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d %b %Y"):
                try:
                    return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
                except ValueError:
                    continue
            return val_str # Return as-is if unparseable, validator will flag it

        elif transformation == "parse_number":
            # Strip currency symbols and commas
            clean_num = re.sub(r"[^\d.-]", "", val_str)
            try:
                if "." in clean_num:
                    return float(clean_num)
                return int(clean_num)
            except ValueError:
                return raw_val

        elif transformation == "yes_no_to_boolean":
            norm = val_str.lower()
            if norm in ("yes", "y", "true", "1", "t"):
                return True
            elif norm in ("no", "n", "false", "0", "f"):
                return False
            return bool(default_val) if default_val else None

        elif transformation == "pa_to_status":
            norm = val_str.upper()
            mapping = {
                "P": "Present",
                "PRESENT": "Present",
                "A": "Absent",
                "ABSENT": "Absent",
                "L": "Late",
                "LATE": "Late",
                "E": "Excused",
                "EXCUSED": "Excused"
            }
            return mapping.get(norm, val_str)

        elif transformation == "normalize_phone":
            # Standardize phone numbers: keep leading +, remove all other non-digits
            has_plus = val_str.startswith("+")
            digits_only = re.sub(r"\D", "", val_str)
            return f"+{digits_only}" if has_plus else digits_only

        elif transformation == "normalize_id":
            # Strip spaces, make uppercase, standardize alphanumeric
            return re.sub(r"\s+", "", val_str).upper()

        # Default: just trim
        return val_str if isinstance(raw_val, str) else raw_val

    @staticmethod
    def transform_record(
        raw_row: Dict[str, Any],
        col_mappings: List[Dict[str, Any]],
        section_heading_target_col: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Transforms a raw Excel row dict into a Supabase record dict based on column mappings.
        """
        record: Dict[str, Any] = {}

        for cm in col_mappings:
            excel_col = cm.get("excel_column") # e.g. "A" or header name
            excel_header = cm.get("excel_header")
            supa_col = cm.get("supabase_column")
            transformation = cm.get("transformation", "none")
            default_val = cm.get("default_value")

            # Try finding value by letter or header name
            raw_val = None
            if excel_col in raw_row:
                raw_val = raw_row[excel_col]
            elif f"__header_{excel_header}" in raw_row:
                raw_val = raw_row[f"__header_{excel_header}"]

            transformed = MappingManager.apply_transformations(raw_val, transformation, default_val)
            record[supa_col] = transformed

        # Attach merged section context if designated (e.g. class = "CLASS 10A")
        if section_heading_target_col and "__section_context" in raw_row:
            record[section_heading_target_col] = raw_row["__section_context"]

        return record
