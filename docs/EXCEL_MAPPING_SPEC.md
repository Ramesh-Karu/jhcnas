# Excel Mapping & Intelligence Specification

## 1. Merged Cells Intelligence
Real-world Excel workbooks frequently use merged cells for visual hierarchy rather than tabular layout.
Nextcloud Excel Sync categorizes merged cells into two distinct types:

### Type A: Sheet Titles / Footers (e.g. `A1:H1 = STUDENT INFORMATION`)
- Spans majority of columns at the top or bottom of the worksheet.
- Handled as non-tabular metadata.
- Automatically filtered out so it is never treated as a database record.

### Type B: Merged Section Headings (e.g. `A3:H3 = CLASS 10A`, `A20:H20 = CLASS 10B`)
- Precedes a group of related data rows.
- The system automatically captures the heading value and propagates it downwards to each row beneath it until the next section heading is encountered.
- When mapped (e.g. `section_heading_target_col = "class"`), each extracted row automatically contains `class: "CLASS 10A"`.

---

## 2. Multi-Worksheet Mapping
A single workbook can synchronize to multiple target tables simultaneously:

| Worksheet | Target Table | Primary / Unique Key | Foreign Key |
|---|---|---|---|
| `Student Details` | `students` | `student_number` | - |
| `Attendance` | `attendance` | `(student_number, attendance_date)` | `students.student_number` |
| `Sports` | `sports` | `(student_number, sport_name)` | `students.student_number` |
| `Medical` | `medical` | `student_number` | `students.student_number` |
| `Results` | `results` | `(student_number, term)` | `students.student_number` |

---

## 3. Transformation Rules Reference
- `trim`: Strips leading/trailing whitespace.
- `uppercase`: Converts text to upper case (e.g. state codes, normalized IDs).
- `lowercase`: Converts text to lower case (e.g. email addresses).
- `parse_date`: Parses diverse date notations (`YYYY-MM-DD`, `DD/MM/YYYY`, `MM/DD/YYYY`) into ISO standard `YYYY-MM-DD`.
- `parse_number`: Removes commas and currency characters and converts to decimal or integer.
- `yes_no_to_boolean`: Maps `Yes`/`Y`/`1` to `True` and `No`/`N`/`0` to `False`.
- `pa_to_status`: Converts attendance notation: `P` -> `Present`, `A` -> `Absent`, `L` -> `Late`, `E` -> `Excused`.
- `normalize_phone`: Strips formatting punctuation, producing clean digits with international prefix.
- `normalize_id`: Removes internal spaces and upper-cases alphanumeric identifiers.
