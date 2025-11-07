# Excel File Format Guide for Student Risk Analyzer

## Overview

The Student Risk Analyzer now expects a **single worksheet** (`.xlsx` format) containing all required information for each student. No additional sheets or merges are needed.

---

## Required Columns

Create one worksheet with the following columns. Header names are matched case-insensitively and with flexible spacing/characters, but using the exact names below is recommended.

| Column | Data Type | Format | Example |
|--------|-----------|--------|---------|
| **Student#** | String / Integer | Unique identifier | `5708615` |
| **Student Name** | Text | Full name; may include hyperlink | `Abadi, Mahlet Chekole` |
| **Program Name** | Text | Program or course name | `Accounting, Payroll and Tax` |
| **Program Grade** | Numeric | Percentage (`88.0`) or decimal (`0.88`) | `88.0` |
| **Attended % to Date.** | Numeric | Percentage (`99.7`) or decimal (`0.997`) | `99.72` |

### Column Details

- **Student#**
  - Used as the unique key in the UI and API responses
  - Trailing `.0` values are automatically removed
  - Required for every row

- **Student Name**
  - Full student name
  - Any embedded hyperlink is preserved and used for Campus Login actions

- **Program Name**
  - Academic program or track
  - Displayed as-is; blank values appear as `Unknown`

- **Program Grade**
  - Overall grade percentage
  - Either `0-100` or `0-1` decimals are accepted
  - Converted to a numeric percentage during parsing

- **Attended % to Date.**
  - Attendance rate
  - Accepts `0-100` percentages or `0-1` decimals
  - Converted to a numeric percentage during parsing

---

## Example Worksheet

```
Student# | Student Name            | Program Name                   | Program Grade | Attended % to Date.
---------|------------------------|--------------------------------|---------------|--------------------
5708615  | Abadi, Mahlet Chekole   | Accounting, Payroll and Tax    | 0.88          | 0.9972
6012457  | John Doe                | Business Administration        | 65.0          | 72.5
1234567  | Jane Smith              | Computer Science               | 92.3          | 98.0
```

---

## Parsing & Cleaning

During upload the parser will:

1. Normalize header names (remove `%`, `#`, extra spaces) and map common aliases (e.g., `Attended % to Date#`).
2. Trim whitespace and remove trailing `.0` from `Student#`.
3. Convert `Program Grade` and `Attended % to Date.` to numeric percentages in the `0-100` range.
4. Drop duplicate `Student#` rows (keeping the first occurrence).
5. Preserve hyperlinks found in the `Student Name` column.

Rows missing `Student#` or `Student Name` are invalid and should be corrected before upload.

---

## Risk Weighting

Once parsed, each record is scored using:

- **Performance Index** = `0.8 × Grade + 0.2 × Attendance` (values expressed as decimals 0–1)
- **Risk Score** = `100 × (1 – Performance Index)` (higher = higher risk)
- Risk categories follow the rule:
  - Grade < 70 **and** Attendance < 70 → *Extremely High Risk*
  - Grade < 80 **or** Attendance < 80 → *High Risk*
  - Grade < 90 **or** Attendance < 90 → *Medium Risk*
  - Otherwise → *Low Risk*

These calculations require both grade and attendance percentages to be present in the single worksheet.

