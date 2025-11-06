# Excel File Format Guide for Student Risk Analyzer

## Overview

The Student Risk Analyzer requires an Excel file (`.xlsx` format) with **exactly two sheets** containing student grade and attendance data. The system merges these sheets using the `Student#` column as the key.

---

## Required Excel File Structure

### File Format
- **File Type**: `.xlsx` (Excel 2007+ format)
- **Number of Sheets**: Exactly 2 sheets
- **Sheet Names**: Must match exactly (case-sensitive)

---

## Sheet 1: "Students Grade"

### Sheet Name
- **Exact Name**: `Students Grade` (case-sensitive, with space)
- **Alternative**: The code looks for this exact name

### Required Columns

| Column Name | Data Type | Format | Example | Required |
|------------|-----------|--------|---------|----------|
| **Student#** | Numeric (Integer) | Whole number | `5708615` | ✅ Yes |
| **Student Name** | Text/String | Full name | `Abadi, Mahlet Chekole` | ✅ Yes |
| **Program Name** | Text/String | Program name | `Accounting, Payroll and Tax` | ✅ Yes |
| **current overall Program Grade** | Numeric (Float) | 0-1 decimal OR 0-100 percent | `0.88` or `88.0` | ✅ Yes |

### Column Details

#### 1. Student# (Required)
- **Purpose**: Unique identifier for each student
- **Format**: Integer (whole number)
- **Example Values**: `5708615`, `6012457`, `1234567`
- **Important**: 
  - Must be numeric (not text)
  - Must match `Student#` values in the attendance sheet
  - Cannot be empty or zero

#### 2. Student Name (Required)
- **Purpose**: Full name of the student
- **Format**: Text string
- **Example Values**: `Abadi, Mahlet Chekole`, `John Doe`, `Jane Smith`
- **Important**: 
  - Cannot be empty
  - Can contain hyperlinks (optional)

#### 3. Program Name (Required)
- **Purpose**: Name of the program the student is enrolled in
- **Format**: Text string
- **Example Values**: `Accounting, Payroll and Tax`, `Business Administration`, `Computer Science`
- **Important**: 
  - Cannot be empty
  - Will show as "Unknown" if missing

#### 4. current overall Program Grade (Required)
- **Purpose**: Overall grade percentage for the student
- **Format**: Numeric (Float)
- **Accepted Formats**:
  - **Decimal (0-1)**: `0.88` = 88%
  - **Percentage (0-100)**: `88.0` = 88%
- **Example Values**: `0.88`, `88.0`, `0.65`, `65.0`
- **Important**: 
  - System automatically converts 0-1 decimals to 0-100 percentages
  - Values > 1 are treated as percentages (0-100)

### Sample Data for "Students Grade" Sheet

```
Student# | Student Name              | Program Name                    | current overall Program Grade
---------|---------------------------|----------------------------------|--------------------------------
5708615  | Abadi, Mahlet Chekole     | Accounting, Payroll and Tax     | 0.88
6012457  | John Doe                  | Business Administration         | 0.65
1234567  | Jane Smith                | Computer Science                | 88.0
```

---

## Sheet 2: "Students attendance " (Note the trailing space!)

### Sheet Name
- **Exact Name**: `Students attendance ` (case-sensitive, **with trailing space**)
- **Alternative**: The code searches for any sheet containing "attendance" (case-insensitive)
- **Accepted Names**: `Students attendance `, `Students Attendance`, `students attendance`, etc.

### Required Columns

| Column Name | Data Type | Format | Example | Required |
|------------|-----------|--------|---------|----------|
| **Student#** | Numeric (Integer) | Whole number | `5708615` | ✅ Yes |
| **Student Name** | Text/String | Full name | `Abadi, Mahlet Chekole` | ✅ Yes |
| **Scheduled Hours to Date** | Text/String OR Numeric | "HH:MM" format OR decimal hours | `90:00` or `90.0` | ✅ Yes |
| **Attended Hours to Date** | Text/String OR Numeric | "HH:MM" format OR decimal hours | `89:45` or `89.75` | ✅ Yes |
| **Attended % to Date.** | Numeric (Float) | 0-1 decimal OR 0-100 percent | `0.997222` or `99.72` | ✅ Yes |
| **Missed Hours to Date** | Text/String OR Numeric | "HH:MM" format OR decimal hours | `0:15` or `0.25` | ✅ Yes |
| **% Missed** | Numeric (Float) | 0-1 decimal OR 0-100 percent | `0.002778` or `0.28` | ✅ Yes |
| **Missed Minus Excused to date** | Text/String OR Numeric | "HH:MM" format OR decimal hours | `0:15` or `0.25` | ✅ Yes |

### Column Details

#### 1. Student# (Required)
- **Purpose**: Unique identifier (must match Grades sheet)
- **Format**: Integer (whole number)
- **Example Values**: `5708615`, `6012457`, `1234567`
- **Important**: 
  - Must match `Student#` values in the Grades sheet
  - Must be numeric (not text)
  - Cannot be empty or zero

#### 2. Student Name (Required)
- **Purpose**: Full name of the student
- **Format**: Text string
- **Example Values**: `Abadi, Mahlet Chekole`, `John Doe`, `Jane Smith`
- **Important**: 
  - Cannot be empty
  - Can contain hyperlinks to Campus Login (optional but recommended)

#### 3. Scheduled Hours to Date (Required)
- **Purpose**: Total scheduled hours
- **Format**: 
  - **Text**: "HH:MM" format (e.g., `90:00` = 90 hours)
  - **Numeric**: Decimal hours (e.g., `90.0` = 90 hours)
- **Example Values**: `90:00`, `90.0`, `120:30`, `120.5`
- **Important**: System automatically converts "HH:MM" to decimal hours

#### 4. Attended Hours to Date (Required)
- **Purpose**: Total hours attended
- **Format**: 
  - **Text**: "HH:MM" format (e.g., `89:45` = 89 hours 45 minutes = 89.75 hours)
  - **Numeric**: Decimal hours (e.g., `89.75` = 89.75 hours)
- **Example Values**: `89:45`, `89.75`, `119:30`, `119.5`
- **Important**: System automatically converts "HH:MM" to decimal hours

#### 5. Attended % to Date. (Required)
- **Purpose**: Attendance percentage
- **Format**: Numeric (Float)
- **Accepted Formats**:
  - **Decimal (0-1)**: `0.997222` = 99.72%
  - **Percentage (0-100)**: `99.72` = 99.72%
- **Example Values**: `0.997222`, `99.72`, `0.88`, `88.0`
- **Important**: 
  - System automatically converts 0-1 decimals to 0-100 percentages
  - Values > 1 are treated as percentages (0-100)
  - **This is the primary column used for risk assessment**

#### 6. Missed Hours to Date (Required)
- **Purpose**: Total hours missed
- **Format**: 
  - **Text**: "HH:MM" format (e.g., `0:15` = 15 minutes = 0.25 hours)
  - **Numeric**: Decimal hours (e.g., `0.25` = 0.25 hours)
- **Example Values**: `0:15`, `0.25`, `5:00`, `5.0`

#### 7. % Missed (Required)
- **Purpose**: Percentage of hours missed
- **Format**: Numeric (Float)
- **Accepted Formats**:
  - **Decimal (0-1)**: `0.002778` = 0.28%
  - **Percentage (0-100)**: `0.28` = 0.28%
- **Example Values**: `0.002778`, `0.28`, `0.12`, `12.0`

#### 8. Missed Minus Excused to date (Required)
- **Purpose**: Hours missed minus excused hours
- **Format**: 
  - **Text**: "HH:MM" format (e.g., `0:15` = 15 minutes = 0.25 hours)
  - **Numeric**: Decimal hours (e.g., `0.25` = 0.25 hours)
- **Example Values**: `0:15`, `0.25`, `5:00`, `5.0`

### Sample Data for "Students attendance " Sheet

```
Student# | Student Name              | Scheduled Hours to Date | Attended Hours to Date | Attended % to Date. | Missed Hours to Date | % Missed | Missed Minus Excused to date
---------|---------------------------|-------------------------|------------------------|---------------------|----------------------|-----------|------------------------------
5708615  | Abadi, Mahlet Chekole     | 90:00                   | 89:45                  | 0.997222            | 0:15                 | 0.002778  | 0:15
6012457  | John Doe                  | 100:00                  | 62:30                  | 0.625               | 37:30                | 0.375     | 37:30
1234567  | Jane Smith                | 80.0                    | 79.5                   | 99.375              | 0.5                  | 0.625     | 0.5
```

---

## Important Notes

### 1. Sheet Names Must Match
- **Grades Sheet**: Must be named `Students Grade` (exact match, case-sensitive)
- **Attendance Sheet**: Must contain "attendance" in the name (case-insensitive)
  - Examples: `Students attendance `, `Students Attendance`, `students attendance`

### 2. Student# Matching
- **Critical**: The `Student#` values in both sheets must match for the merge to work
- **Format**: Must be numeric (integer), not text
- **Example**: If a student has `Student# = 5708615` in the Grades sheet, they must have the same `Student# = 5708615` in the Attendance sheet

### 3. Data Cleaning
The system automatically:
- Removes rows with empty `Student#` or `Student Name`
- Removes summary/total rows (rows containing "Total" or "Summary" in Student Name)
- Converts "HH:MM" time strings to decimal hours
- Normalizes percentages (0-1 decimals → 0-100 percentages)
- Handles missing data (fills with 0.0)

### 4. Missing Data Handling
- **Students with only Grades**: Will have `attendance_pct = 0.0`
- **Students with only Attendance**: Will have `grade_pct = 0.0`
- **Missing Program Name**: Will show as "Unknown"
- **Missing Student Name**: Will show as "Unknown"

### 5. Hyperlinks (Optional)
- **Campus Login URLs**: Can be embedded as hyperlinks in the `Student Name` column
- **Preferred**: Hyperlinks in the Attendance sheet take precedence
- **Format**: Standard Excel hyperlinks

---

## Common Issues and Solutions

### Issue 1: Only 2 Students Showing
**Cause**: 
- Student# values don't match between sheets
- Student# is stored as text instead of numeric
- Missing required columns

**Solution**:
1. Ensure `Student#` is numeric (not text) in both sheets
2. Verify `Student#` values match exactly between sheets
3. Check that all required columns are present

### Issue 2: Attendance Showing as 0.0%
**Cause**:
- `Attended % to Date.` column is missing or incorrectly named
- Values are in wrong format (e.g., text instead of numeric)

**Solution**:
1. Verify column name is exactly `Attended % to Date.` (with period at end)
2. Ensure values are numeric (0-1 decimal or 0-100 percent)
3. Check for empty cells or invalid data

### Issue 3: Program Name Showing as "nan"
**Cause**:
- `Program Name` column is missing or empty
- Values are stored as NaN (Not a Number)

**Solution**:
1. Ensure `Program Name` column exists in Grades sheet
2. Fill in all program names (cannot be empty)
3. Check for NaN values in Excel

### Issue 4: Sheet Not Found Error
**Cause**:
- Sheet names don't match exactly
- Sheets are missing

**Solution**:
1. Verify Grades sheet is named exactly `Students Grade`
2. Verify Attendance sheet contains "attendance" in the name
3. Check that both sheets exist in the Excel file

---

## Example Excel File Structure

```
📁 Student_Data.xlsx
├── 📄 Sheet 1: "Students Grade"
│   ├── Column A: Student# (5708615, 6012457, ...)
│   ├── Column B: Student Name (Abadi, Mahlet Chekole, ...)
│   ├── Column C: Program Name (Accounting, Payroll and Tax, ...)
│   └── Column D: current overall Program Grade (0.88, 0.65, ...)
│
└── 📄 Sheet 2: "Students attendance "
    ├── Column A: Student# (5708615, 6012457, ...)
    ├── Column B: Student Name (Abadi, Mahlet Chekole, ...)
    ├── Column C: Scheduled Hours to Date (90:00, 100:00, ...)
    ├── Column D: Attended Hours to Date (89:45, 62:30, ...)
    ├── Column E: Attended % to Date. (0.997222, 0.625, ...)
    ├── Column F: Missed Hours to Date (0:15, 37:30, ...)
    ├── Column G: % Missed (0.002778, 0.375, ...)
    └── Column H: Missed Minus Excused to date (0:15, 37:30, ...)
```

---

## Quick Checklist

Before uploading your Excel file, verify:

- [ ] File is `.xlsx` format (not `.xls` or `.csv`)
- [ ] Exactly 2 sheets exist
- [ ] Grades sheet is named `Students Grade` (exact match)
- [ ] Attendance sheet contains "attendance" in the name
- [ ] `Student#` column exists in both sheets and is numeric
- [ ] `Student Name` column exists in both sheets
- [ ] `Program Name` column exists in Grades sheet
- [ ] `current overall Program Grade` column exists in Grades sheet
- [ ] `Attended % to Date.` column exists in Attendance sheet
- [ ] All `Student#` values match between sheets
- [ ] No empty `Student#` or `Student Name` values
- [ ] No summary/total rows in the data

---

## Need Help?

If you're still experiencing issues:
1. Check the server console logs for error messages
2. Verify your Excel file matches the structure above
3. Test with a small sample file (5-10 students) first
4. Ensure all required columns are present and properly formatted

