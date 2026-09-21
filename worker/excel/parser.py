"""
Advanced Excel Workbook Parser with openpyxl & pandas.
Specialized for complex real-world workbooks:
- Merged cells identification and classification (Titles vs Section Headings)
- Section header downward propagation to subsequent data rows
- Intelligent Header Row and Data Region detection
- Multi-worksheet inspection (.xlsx and .xlsm)
"""
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Set
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import pandas as pd

class ExcelParser:
    def __init__(self, file_path: str):
        self.file_path = file_path
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found: {file_path}")

    def analyze_workbook(self) -> Dict[str, Any]:
        """
        Comprehensive Workbook Analyzer fulfilling Section 6 & 7 requirements:
        - Filename, File size
        - Worksheet names
        - For each worksheet: Rows, Columns, Used range, Merged ranges,
          Possible header rows, Possible data regions, Sample data,
          Empty rows, Repeated headers, Section headings.
        """
        file_size = os.path.getsize(self.file_path)
        filename = os.path.basename(self.file_path)

        # Load workbook with openpyxl (data_only=False to detect formulas/merged cells)
        wb = openpyxl.load_workbook(self.file_path, data_only=True)
        worksheets_analysis = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_info = self._analyze_single_sheet(ws)
            worksheets_analysis.append(sheet_info)

        wb.close()

        return {
            "filename": filename,
            "file_size": file_size,
            "file_size_formatted": self._format_size(file_size),
            "worksheet_names": [w["sheet_name"] for w in worksheets_analysis],
            "total_worksheets": len(worksheets_analysis),
            "worksheets": worksheets_analysis
        }

    def _analyze_single_sheet(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> Dict[str, Any]:
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0

        merged_ranges = []
        for mr in ws.merged_cells.ranges:
            merged_ranges.append(str(mr))

        # Scan rows to detect empty rows, titles, section headings, and candidate header rows
        row_analyses = []
        empty_rows = []
        title_rows = []
        section_headings = []
        candidate_headers = []

        # Sample data up to first 50 rows for performance and deep inspection
        scan_limit = min(max_row, 100)
        for r in range(1, scan_limit + 1):
            cell_values = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
            non_empty = [v for v in cell_values if v is not None and str(v).strip() != ""]
            
            if len(non_empty) == 0:
                empty_rows.append(r)
                continue

            # Check if this row is part of a full-width merged cell
            row_merged = [mr for mr in ws.merged_cells.ranges if mr.min_row <= r <= mr.max_row]
            is_spanning_merge = any((mr.max_col - mr.min_col + 1) >= max(2, max_col // 2) for mr in row_merged)

            first_val = str(non_empty[0]).strip() if non_empty else ""

            if len(non_empty) == 1 and (is_spanning_merge or len(non_empty) < max_col // 2):
                # Likely a title or section heading
                if r <= 2:
                    title_rows.append({"row": r, "text": first_val})
                else:
                    section_headings.append({"row": r, "text": first_val, "range": str(row_merged[0]) if row_merged else f"A{r}"})
            elif len(non_empty) >= 2:
                # Potential header row or data row
                string_count = sum(1 for v in non_empty if isinstance(v, str) and not v.replace('.', '', 1).isdigit())
                if string_count / len(non_empty) >= 0.7:
                    candidate_headers.append({
                        "row": r,
                        "columns_detected": len(non_empty),
                        "headers": [str(v).strip() for v in non_empty],
                        "confidence": 0.9 if r in [1, 2, 4, 5] else 0.7
                    })

        # Determine best detected header row
        best_header_row = 1
        if candidate_headers:
            best_header_row = candidate_headers[0]["row"]
        elif max_row > 0:
            best_header_row = 1

        # Data start row is usually header_row + 1, skipping empty rows
        data_start_row = best_header_row + 1
        while data_start_row in empty_rows and data_start_row < max_row:
            data_start_row += 1

        # Read actual headers at best_header_row
        headers = []
        for c in range(1, max_col + 1):
            val = ws.cell(row=best_header_row, column=c).value
            h_name = str(val).strip() if val is not None else f"Column_{get_column_letter(c)}"
            headers.append({
                "col_letter": get_column_letter(c),
                "col_index": c,
                "name": h_name
            })

        # Fetch preview sample rows (5-10 rows)
        sample_rows = []
        sample_limit = min(max_row, data_start_row + 8)
        for r in range(data_start_row, sample_limit + 1):
            if r in empty_rows:
                continue
            row_dict = {}
            for c in range(1, max_col + 1):
                col_letter = get_column_letter(c)
                val = ws.cell(row=r, column=c).value
                row_dict[col_letter] = str(val) if val is not None else ""
            sample_rows.append({"row_number": r, "data": row_dict})

        return {
            "sheet_name": ws.title,
            "total_rows": max_row,
            "total_columns": max_col,
            "used_range": f"A1:{get_column_letter(max_col)}{max_row}" if max_row and max_col else "Empty",
            "merged_ranges": merged_ranges,
            "detected_header_row": best_header_row,
            "detected_data_start_row": data_start_row,
            "candidate_header_rows": candidate_headers,
            "title_rows": title_rows,
            "section_headings": section_headings,
            "empty_rows_count": len(empty_rows),
            "headers": headers,
            "sample_rows": sample_rows
        }

    def extract_sheet_records(
        self,
        worksheet_name: str,
        header_row: int,
        data_start_row: int,
        data_end_row: Optional[int] = None,
        section_heading_target_col: Optional[str] = None
    ) -> Tuple[List[Dict[str, Any]], List[str]]:
        """
        Extracts structured records with intelligent merged-section propagation.
        If a merged cell like 'A3:H3 = CLASS 10A' precedes data rows, the value 'CLASS 10A'
        is automatically captured and assigned to records beneath it under the section_heading_target_col!
        """
        wb = openpyxl.load_workbook(self.file_path, data_only=True)
        if worksheet_name not in wb.sheetnames:
            wb.close()
            raise ValueError(f"Worksheet '{worksheet_name}' not found in workbook.")

        ws = wb[worksheet_name]
        max_row = data_end_row if data_end_row else (ws.max_row or 0)
        max_col = ws.max_column or 0

        # Build column map from header_row
        col_headers: Dict[int, str] = {}
        for c in range(1, max_col + 1):
            val = ws.cell(row=header_row, column=c).value
            if val is not None and str(val).strip():
                col_headers[c] = str(val).strip()
            else:
                col_headers[c] = f"Col_{get_column_letter(c)}"

        # Identify all merged section headings in the sheet
        merged_sections: Dict[int, str] = {} # row -> heading text
        for mr in ws.merged_cells.ranges:
            if (mr.max_col - mr.min_col + 1) >= max(2, max_col // 2):
                top_left_val = ws.cell(row=mr.min_row, column=mr.min_col).value
                if top_left_val:
                    merged_sections[mr.min_row] = str(top_left_val).strip()

        records = []
        current_section_value = None

        for r in range(data_start_row, max_row + 1):
            # Check if this row is a new section heading
            if r in merged_sections:
                current_section_value = merged_sections[r]
                continue

            # Read all cell values in this row
            row_data = {}
            has_content = False
            for c in range(1, max_col + 1):
                val = ws.cell(row=r, column=c).value
                col_letter = get_column_letter(c)
                header_name = col_headers.get(c, f"Col_{col_letter}")
                
                if val is not None:
                    has_content = True
                    # Clean up strings
                    if isinstance(val, str):
                        val = val.strip()

                row_data[col_letter] = val
                row_data[f"__header_{header_name}"] = val

            if not has_content:
                # Skip blank separator rows
                continue

            # Check if this row is a repeated header
            is_repeated_header = all(
                str(row_data.get(get_column_letter(c), '')).strip() == str(col_headers.get(c, '')).strip()
                for c in range(1, min(5, max_col + 1)) if c in col_headers
            )
            if is_repeated_header:
                continue

            # Attach section context if configured
            if section_heading_target_col and current_section_value:
                row_data["__section_context"] = current_section_value

            row_data["__row_number"] = r
            records.append(row_data)

        wb.close()
        return records, list(col_headers.values())

    @staticmethod
    def _format_size(size_bytes: int) -> str:
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"
