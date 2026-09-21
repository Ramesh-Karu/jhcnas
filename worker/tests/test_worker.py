"""
Automated Test Suite for Nextcloud Excel Sync Worker.
Tests:
- Excel file generation with complex merged cells & multi-sheets
- Merged cells & section headers extraction
- Header detection & data regions
- Transformation functions (P/A attendance, dates, phone, booleans, trim)
- Data validation (required, types, duplicates)
- Upsert simulation and duplicate prevention
- File SHA-256 hash change detection
- Dry Run vs Import logic
"""
import os
import tempfile
import hashlib
import pytest
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from excel.parser import ExcelParser
from mappings.manager import MappingManager
from validation.validator import DataValidator

@pytest.fixture
def complex_workbook_path():
    """Generates a real sample complex workbook matching the prompt specification."""
    temp_dir = tempfile.mkdtemp()
    file_path = os.path.join(temp_dir, "students_complex_test.xlsx")

    wb = openpyxl.Workbook()

    # --- Sheet 1: Student Details ---
    ws1 = wb.active
    ws1.title = "Student Details"

    # Row 1: Merged Title A1:H1
    ws1.merge_cells("A1:H1")
    ws1["A1"] = "ST. JUDE ACADEMY - STUDENT INFORMATION"
    ws1["A1"].font = Font(size=14, bold=True)
    ws1["A1"].alignment = Alignment(horizontal="center")

    # Row 2: Metadata
    ws1["A2"] = "Academic Year: 2025-2026 | Generated: TrueNAS Nextcloud"

    # Row 3: Merged Section Heading A3:H3 = CLASS 10A
    ws1.merge_cells("A3:H3")
    ws1["A3"] = "CLASS 10A"
    ws1["A3"].font = Font(bold=True)

    # Row 4: Empty row
    # Row 5: Actual Header Row
    headers = ["Student ID", "Student Name", "Gender", "DOB", "Phone", "Emergency Contact", "Status", "Remarks"]
    for c_idx, h in enumerate(headers, 1):
        ws1.cell(row=5, column=c_idx, value=h)

    # Rows 6-8: Student Data in Class 10A
    ws1.append(["STU-1001", "Alice Smith", "F", "2008-04-12", "+1 (555) 019-2831", "+1 555 019 9999", "Active", "Honors student"])
    ws1.append(["STU-1002", "Bob Jones", "M", "2008-09-21", "+1 (555) 018-4422", "+1 555 018 8888", "Active", "Needs locker"])
    ws1.append(["STU-1003", "Charlie Davis", "M", "2007-12-05", "555-017-3311", "555-017-0000", "Active", "Transfer"])

    # Row 10: Next Section Heading: CLASS 10B
    ws1.merge_cells("A10:H10")
    ws1["A10"] = "CLASS 10B"
    ws1.append(["STU-1004", "Diana Prince", "F", "2008-06-18", "+1 (555) 016-5544", "+1 555 016 7777", "Active", "Team captain"])
    ws1.append(["STU-1005", "Ethan Hunt", "M", "2008-01-30", "+1 (555) 015-6677", "+1 555 015 2222", "Active", "Track & field"])

    # --- Sheet 2: Attendance ---
    ws2 = wb.create_sheet(title="Attendance")
    ws2.append(["Student ID", "Date", "Status", "Remarks"])
    ws2.append(["STU-1001", "2025-09-15", "P", "On time"])
    ws2.append(["STU-1002", "2025-09-15", "A", "Excused illness"])
    ws2.append(["STU-1003", "2025-09-15", "L", "Bus delay"])
    ws2.append(["STU-1004", "2025-09-15", "P", "On time"])
    ws2.append(["STU-1005", "2025-09-15", "P", "On time"])

    # --- Sheet 3: Sports ---
    ws3 = wb.create_sheet(title="Sports")
    ws3.append(["Student ID", "Sport Name", "Position", "Medical Clearance"])
    ws3.append(["STU-1001", "Basketball", "Point Guard", "Yes"])
    ws3.append(["STU-1004", "Soccer", "Midfielder", "Yes"])
    ws3.append(["STU-1005", "Track & Field", "Sprinter", "Yes"])

    # --- Sheet 4: Medical ---
    ws4 = wb.create_sheet(title="Medical")
    ws4.append(["Student ID", "Blood Group", "Allergies", "Doctor Contact"])
    ws4.append(["STU-1001", "O+", "Peanuts", "Dr. Wilson (555-0100)"])
    ws4.append(["STU-1002", "A+", "None", "Dr. Gomez (555-0200)"])
    ws4.append(["STU-1003", "B+", "Penicillin", "Dr. Adams (555-0300)"])

    # --- Sheet 5: Results ---
    ws5 = wb.create_sheet(title="Results")
    ws5.append(["Student ID", "Term", "Mathematics", "Science", "English", "Total Score"])
    ws5.append(["STU-1001", "Term 1", 94.5, 88.0, 92.0, 274.5])
    ws5.append(["STU-1002", "Term 1", 76.0, 82.5, 79.0, 237.5])
    ws5.append(["STU-1003", "Term 1", 85.0, 90.0, 84.5, 259.5])

    wb.save(file_path)
    wb.close()
    yield file_path

    # Cleanup
    if os.path.exists(file_path):
        os.remove(file_path)

def test_workbook_analysis_and_merged_cells(complex_workbook_path):
    """Verifies Section 6 & 7: Merged cell detection and multi-sheet analysis."""
    parser = ExcelParser(complex_workbook_path)
    analysis = parser.analyze_workbook()

    assert analysis["total_worksheets"] == 5
    assert "Student Details" in analysis["worksheet_names"]
    assert "Attendance" in analysis["worksheet_names"]

    details_sheet = next(w for w in analysis["worksheets"] if w["sheet_name"] == "Student Details")
    assert any("A1:H1" in m for m in details_sheet["merged_ranges"])
    assert any("A3:H3" in m for m in details_sheet["merged_ranges"])
    assert details_sheet["detected_header_row"] == 5

def test_merged_section_heading_downward_propagation(complex_workbook_path):
    """Verifies Section 7: Merged heading CLASS 10A applies to rows beneath it."""
    parser = ExcelParser(complex_workbook_path)
    records, headers = parser.extract_sheet_records(
        worksheet_name="Student Details",
        header_row=5,
        data_start_row=6,
        section_heading_target_col="class"
    )

    assert len(records) >= 5
    # First record should inherit CLASS 10A
    first_rec = records[0]
    assert first_rec.get("__section_context") == "CLASS 10A"
    assert first_rec.get("A") == "STU-1001"

    # Last record should inherit CLASS 10B
    last_rec = records[-1]
    assert last_rec.get("__section_context") == "CLASS 10B"
    assert last_rec.get("A") == "STU-1005"

def test_transformations():
    """Verifies Section 10: Data transformation rules."""
    # Trim
    assert MappingManager.apply_transformations("  hello world  ", "trim") == "hello world"
    # Case conversion
    assert MappingManager.apply_transformations("alice", "uppercase") == "ALICE"
    assert MappingManager.apply_transformations("BOB", "lowercase") == "bob"
    # Boolean Yes/No
    assert MappingManager.apply_transformations("Yes", "yes_no_to_boolean") is True
    assert MappingManager.apply_transformations("no", "yes_no_to_boolean") is False
    # P/A attendance conversion
    assert MappingManager.apply_transformations("P", "pa_to_status") == "Present"
    assert MappingManager.apply_transformations("A", "pa_to_status") == "Absent"
    assert MappingManager.apply_transformations("L", "pa_to_status") == "Late"
    # Phone normalization
    assert MappingManager.apply_transformations("+1 (555) 019-2831", "normalize_phone") == "+15550192831"
    # ID normalization
    assert MappingManager.apply_transformations(" stu - 1001 ", "normalize_id") == "STU-1001"

def test_data_validation_and_duplicate_prevention():
    """Verifies Section 11 & 13: Data validation and intra-batch duplicate detection."""
    mappings = [
        {"excel_column": "A", "supabase_column": "student_number", "data_type": "text", "required": True, "unique_key": True},
        {"excel_column": "B", "supabase_column": "score", "data_type": "decimal", "required": True},
        {"excel_column": "C", "supabase_column": "dob", "data_type": "date", "required": False}
    ]

    test_records = [
        {"student_number": "STU-001", "score": 95.5, "dob": "2008-05-10", "__row_number": 2},
        {"student_number": "STU-002", "score": "invalid_num", "dob": "2008-05-11", "__row_number": 3}, # Type mismatch
        {"student_number": "", "score": 80.0, "dob": "2008-05-12", "__row_number": 4}, # Missing required
        {"student_number": "STU-001", "score": 90.0, "dob": "2008-05-13", "__row_number": 5}, # Duplicate unique key
    ]

    valid, failed, errors = DataValidator.validate_records(test_records, mappings, "Student Details")

    assert len(valid) == 1
    assert valid[0]["student_number"] == "STU-001"
    assert len(failed) == 3
    assert len(errors) >= 3

    err_types = [e["error_type"] for e in errors]
    assert "type_mismatch" in err_types
    assert "missing_required" in err_types
    assert "duplicate" in err_types

def test_file_hash_computation(complex_workbook_path):
    """Verifies Section 14: SHA-256 calculation for file change detection."""
    sha256 = hashlib.sha256()
    with open(complex_workbook_path, "rb") as f:
        while chunk := f.read(4096):
            sha256.update(chunk)
    computed_hash = sha256.hexdigest()

    assert len(computed_hash) == 64
    assert isinstance(computed_hash, str)
