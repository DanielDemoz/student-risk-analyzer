"""Excel file parsing and data normalization."""

import pandas as pd
import numpy as np
import re
from typing import Tuple, Dict, Optional
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell


def clean_student_id(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize Student# column so merge keys match.
    Extracts digits only and converts to numeric.
    """
    if "Student#" in df.columns:
        # Convert to string first, then extract digits only
        df["Student#"] = (
            df["Student#"]
            .astype(str)
            .str.extract(r"(\d+)", expand=False)  # Extract digits only
        )
        df["Student#"] = pd.to_numeric(df["Student#"], errors="coerce")
    return df


def clean_attendance_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean, convert, and normalize attendance data.
    Removes total/summary rows, converts time columns, normalizes percentages.
    """
    df = df.copy()
    
    # Remove total/summary rows
    if "Student Name" in df.columns:
        # Remove rows where Student Name contains "Total", "Summary", or is empty
        mask = (
            ~df["Student Name"].astype(str).str.contains("Total", case=False, na=False) &
            ~df["Student Name"].astype(str).str.contains("Summary", case=False, na=False) &
            ~df["Student Name"].astype(str).str.strip().eq("") &
            df["Student Name"].notna()
        )
        df = df[mask].copy()
        print(f"DEBUG: Removed summary rows, remaining rows: {len(df)}")
    
    # Convert HH:MM columns to decimal hours
    time_columns = [
        "Scheduled Hours to Date",
        "Attended Hours to Date",
        "Missed Hours to Date",
        "Missed Minus Excused to date",
    ]
    
    for col in time_columns:
        if col in df.columns:
            df[col] = df[col].apply(to_hours)
    
    # Normalize % columns
    percentage_columns = ["Attended % to Date.", "% Missed"]
    
    for col in percentage_columns:
        if col in df.columns:
            df[col] = df[col].apply(normalize_pct)
    
    # Replace NaN and invalid values
    df = df.replace([np.inf, -np.inf, np.nan], 0)
    
    return df


def to_hours(value) -> float:
    """
    Convert time strings like '90:00' to numeric hours.
    Handles both string formats and numeric values.
    
    Args:
        value: String in format 'HH:MM', 'H:MM', or numeric value
    
    Returns:
        Decimal hours (e.g., '90:00' -> 90.0, '0:15' -> 0.25, 90.0 -> 90.0)
    """
    if isinstance(value, (int, float)):
        return float(value)
    
    if pd.isna(value) or value == '':
        return 0.0
    
    if isinstance(value, str) and ":" in value:
        try:
            hours, minutes = value.split(":")
            return float(hours) + float(minutes) / 60.0
        except (ValueError, TypeError):
            return 0.0
    
    # Try to parse as decimal number
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def parse_duration(duration_str: str) -> float:
    """
    Parse duration string like '90:00' or '0:15' to decimal hours.
    (Alias for to_hours for backward compatibility)
    
    Args:
        duration_str: String in format 'HH:MM' or 'H:MM'
    
    Returns:
        Decimal hours (e.g., '90:00' -> 90.0, '0:15' -> 0.25)
    """
    return to_hours(duration_str)


def normalize_attendance_pct(x) -> float:
    """
    Normalize attendance percentage values.
    Handles both 0-1 decimals (e.g., 0.88) and 0-100 percentages (e.g., 88).
    
    Args:
        x: Value that might be in 0-1 range or 0-100 range
    
    Returns:
        Percentage in 0-100 range
    """
    if pd.isna(x) or x == '':
        return 0.0
    
    try:
        # Handle string values like "85%" or "0.85"
        if isinstance(x, str):
            val_str = str(x).strip().replace('%', '').strip()
            val = float(val_str)
        else:
            val = float(x)
        
        # If value is <= 1, assume it's a decimal (0-1 range) and multiply by 100
        # If value > 1, assume it's already a percentage (0-100 range)
        if val <= 1.0:
            return val * 100.0
        return val
    except (ValueError, TypeError):
        return 0.0


def normalize_pct(x) -> float:
    """
    Normalize percentage values (simpler version for direct column application).
    Handles both 0-1 decimals (e.g., 0.88) and 0-100 percentages (e.g., 88).
    
    Args:
        x: Value that might be in 0-1 range or 0-100 range
    
    Returns:
        Percentage in 0-100 range
    """
    # Handle NaN/None first
    if pd.isna(x) or x is None:
        return 0.0
    
    try:
        # Convert to float, handling string values
        if isinstance(x, str):
            # Remove % sign and whitespace
            val_str = str(x).strip().replace('%', '').strip()
            if not val_str or val_str == '':
                return 0.0
            val = float(val_str)
        else:
            val = float(x)
        
        # Check for invalid values
        if np.isnan(val) or np.isinf(val):
            return 0.0
        
        # If value is <= 1, assume it's a decimal (0-1 range) and multiply by 100
        # If value > 1, assume it's already a percentage (0-100 range)
        if val <= 1.0:
            return val * 100.0
        return val
    except (ValueError, TypeError) as e:
        print(f"DEBUG: normalize_pct error for value '{x}' (type: {type(x)}): {e}")
        return 0.0


def normalize_attendance_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Preprocessing step to automatically clean and normalize attendance data.
    
    This function:
    - Converts all "HH:MM" time strings into decimal hours (e.g., "89:45" -> 89.75)
    - Normalizes attendance percentages (e.g., 0.997 -> 99.7)
    - Replaces NaN or invalid values with 0
    
    Args:
        df: Raw attendance DataFrame
    
    Returns:
        Cleaned and normalized attendance DataFrame
    """
    # Make a copy to avoid modifying the original
    df = df.copy()
    
    # Debug: Print available columns before processing
    print(f"\n=== DEBUG: normalize_attendance_data - Available columns ===")
    print(f"Columns: {list(df.columns)}")
    
    # Convert HH:MM text columns to decimal hours
    time_columns = [
        "Scheduled Hours to Date",
        "Attended Hours to Date",
        "Missed Hours to Date",
        "Missed Minus Excused to date",
    ]
    
    for col in time_columns:
        if col in df.columns:
            print(f"DEBUG: Converting time column '{col}' - Sample values before: {df[col].head(3).tolist()}")
            df[col] = df[col].apply(to_hours)
            print(f"DEBUG: After conversion - Sample values: {df[col].head(3).tolist()}")
        else:
            # Try case-insensitive search
            for actual_col in df.columns:
                if col.lower().strip() == actual_col.lower().strip():
                    print(f"DEBUG: Found column '{actual_col}' (case variation of '{col}')")
                    print(f"DEBUG: Sample values before: {df[actual_col].head(3).tolist()}")
                    df[col] = df[actual_col].apply(to_hours)
                    print(f"DEBUG: After conversion - Sample values: {df[col].head(3).tolist()}")
                    break
    
    # Normalize percentage columns
    percentage_columns = ["Attended % to Date.", "% Missed"]
    
    for col in percentage_columns:
        if col in df.columns:
            print(f"DEBUG: Normalizing percentage column '{col}' - Sample values before: {df[col].head(3).tolist()}")
            df[col] = df[col].apply(normalize_pct)
            print(f"DEBUG: After normalization - Sample values: {df[col].head(3).tolist()}")
        else:
            # Try case-insensitive search
            for actual_col in df.columns:
                if col.lower().strip() == actual_col.lower().strip():
                    print(f"DEBUG: Found column '{actual_col}' (case variation of '{col}')")
                    print(f"DEBUG: Sample values before: {df[actual_col].head(3).tolist()}")
                    df[col] = df[actual_col].apply(normalize_pct)
                    print(f"DEBUG: After normalization - Sample values: {df[col].head(3).tolist()}")
                    break
    
    # Replace invalid or missing values (but preserve valid zeros)
    # Only replace NaN, Infinity, and -Infinity
    df = df.replace([np.inf, -np.inf], 0)
    # Fill NaN values with 0 only for numeric columns
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        df[col] = df[col].fillna(0)
    
    print(f"=== END DEBUG: normalize_attendance_data ===\n")
    
    return df


def normalize_percentage(value: float, max_value: float = 1.0) -> float:
    """
    Normalize percentage to 0-100 range.
    
    If max_value <= 1.0, multiply by 100.
    """
    if pd.isna(value):
        return 0.0
    
    if max_value <= 1.0:
        return float(value) * 100.0
    return float(value)


def extract_hyperlink_from_cell(cell: Cell) -> Optional[str]:
    """Extract hyperlink from an openpyxl cell if present."""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return None


def load_excel(file_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, str]]:
    """
    Load Excel file and extract data from both sheets.
    
    Expected Excel structure:
    - Sheet 1: "Students Grade"
      - Student#: numeric (unique student ID)
      - Student Name: string
      - Program Name: string
      - current overall Program Grade: float (0-1 decimal or 0-100 percent)
    
    - Sheet 2: "Students attendance " (note trailing space)
      - Student#: numeric (same ID key as Grades sheet)
      - Student Name: string (some have hyperlinks to Campus Login)
      - Scheduled Hours to Date: string "HH:MM" (e.g., "90:00")
      - Attended Hours to Date: string "HH:MM" (e.g., "89:45")
      - Attended % to Date.: float 0-1 (e.g., 0.997222 = 99.7%)
      - Missed Hours to Date: string "HH:MM" (e.g., "5:00")
      - % Missed: float 0-1
      - Missed Minus Excused to date: string "HH:MM" or number (e.g., "0:15" or 0)
    
    Args:
        file_bytes: Raw bytes of the Excel file
    
    Returns:
        Tuple of (grades_df, attendance_df, name_hyperlinks_dict)
        where name_hyperlinks_dict maps student_id -> hyperlink_url
    """
    # Load workbook to extract hyperlinks
    workbook = load_workbook(filename=BytesIO(file_bytes), data_only=False)
    
    # Extract hyperlinks from Grades sheet
    name_hyperlinks = {}
    
    if 'Students Grade' in workbook.sheetnames:
        grades_sheet = workbook['Students Grade']
        
        # Find Student# and Student Name columns
        header_row = None
        student_id_col = None
        student_name_col = None
        
        for row_idx, row in enumerate(grades_sheet.iter_rows(min_row=1, max_row=20), start=1):
            for col_idx, cell in enumerate(row, start=1):
                cell_value = str(cell.value).strip() if cell.value else ''
                if 'Student#' in cell_value or 'Student #' in cell_value:
                    header_row = row_idx
                    student_id_col = col_idx
                elif 'Student Name' in cell_value:
                    student_name_col = col_idx
            
            if header_row and student_id_col and student_name_col:
                break
        
        # Extract hyperlinks
        if header_row and student_id_col and student_name_col:
            for row in grades_sheet.iter_rows(min_row=header_row + 1):
                # Check if row has enough columns
                if len(row) < max(student_id_col, student_name_col):
                    continue
                try:
                    student_id_cell = row[student_id_col - 1]
                    student_name_cell = row[student_name_col - 1]
                    
                    if student_id_cell.value and student_name_cell.value:
                        # Convert Student# to numeric for consistent matching
                        try:
                            student_id = str(int(float(student_id_cell.value)))
                        except (ValueError, TypeError):
                            student_id = str(student_id_cell.value).strip()
                        hyperlink = extract_hyperlink_from_cell(student_name_cell)
                        if hyperlink:
                            name_hyperlinks[student_id] = hyperlink
                except (IndexError, TypeError) as e:
                    # Skip rows that don't have enough columns or have invalid data
                    continue
    
    # Load dataframes using pandas
    excel_file = BytesIO(file_bytes)
    
    # Load Grades sheet
    try:
        grades_df = pd.read_excel(excel_file, sheet_name='Students Grade', engine='openpyxl')
    except ValueError as e:
        raise ValueError(f"Could not find 'Students Grade' sheet. Available sheets: {workbook.sheetnames}")
    
    # Reset file pointer for second sheet
    excel_file.seek(0)
    
    # Load Attendance sheet (note the trailing space) and extract hyperlinks
    attendance_sheet_name = None
    for sheet_name in workbook.sheetnames:
        if 'attendance' in sheet_name.lower():
            attendance_sheet_name = sheet_name
            break
    
    if not attendance_sheet_name:
        raise ValueError(f"Could not find attendance sheet. Available sheets: {workbook.sheetnames}")
    
    # Extract hyperlinks from Attendance sheet
    attendance_sheet = workbook[attendance_sheet_name]
    attendance_hyperlinks = {}
    
    # Find Student# and Student Name columns in attendance sheet
    att_header_row = None
    att_student_id_col = None
    att_student_name_col = None
    
    for row_idx, row in enumerate(attendance_sheet.iter_rows(min_row=1, max_row=20), start=1):
        for col_idx, cell in enumerate(row, start=1):
            cell_value = str(cell.value).strip() if cell.value else ''
            if 'Student#' in cell_value or 'Student #' in cell_value:
                att_header_row = row_idx
                att_student_id_col = col_idx
            elif 'Student Name' in cell_value:
                att_student_name_col = col_idx
        
        if att_header_row and att_student_id_col and att_student_name_col:
            break
    
    # Extract hyperlinks from attendance sheet
    if att_header_row and att_student_id_col and att_student_name_col:
        for row in attendance_sheet.iter_rows(min_row=att_header_row + 1):
            # Check if row has enough columns
            if len(row) < max(att_student_id_col, att_student_name_col):
                continue
            try:
                student_id_cell = row[att_student_id_col - 1]
                student_name_cell = row[att_student_name_col - 1]
                
                if student_id_cell.value and student_name_cell.value:
                    # Convert Student# to numeric for consistent matching
                    try:
                        student_id = str(int(float(student_id_cell.value)))
                    except (ValueError, TypeError):
                        student_id = str(student_id_cell.value).strip()
                    hyperlink = extract_hyperlink_from_cell(student_name_cell)
                    if hyperlink:
                        attendance_hyperlinks[student_id] = hyperlink
                        # Attendance sheet hyperlinks take precedence over grades sheet
                        name_hyperlinks[student_id] = hyperlink
            except IndexError:
                # Skip rows that don't have enough columns
                continue
    
    # Load attendance dataframe
    attendance_df = pd.read_excel(excel_file, sheet_name=attendance_sheet_name, engine='openpyxl')
    
    # Normalize column names (strip whitespace, handle variations) BEFORE cleaning
    grades_df.columns = grades_df.columns.str.strip()
    attendance_df.columns = attendance_df.columns.str.strip()
    
    # Debug: Print raw attendance DataFrame info
    print(f"\n=== DEBUG: Raw attendance DataFrame after pd.read_excel ===")
    print(f"Shape: {attendance_df.shape}")
    print(f"Columns: {list(attendance_df.columns)}")
    if 'Attended % to Date.' in attendance_df.columns:
        print(f"'Attended % to Date.' dtype: {attendance_df['Attended % to Date.'].dtype}")
        print(f"'Attended % to Date.' sample values (raw): {attendance_df['Attended % to Date.'].head(5).tolist()}")
        print(f"'Attended % to Date.' non-null count: {attendance_df['Attended % to Date.'].notna().sum()}")
        print(f"'Attended % to Date.' null count: {attendance_df['Attended % to Date.'].isna().sum()}")
    else:
        print("WARNING: 'Attended % to Date.' column NOT found in raw DataFrame!")
        # Try to find similar column names
        for col in attendance_df.columns:
            if 'attended' in col.lower() and '%' in col:
                print(f"Found similar column: '{col}' with values: {attendance_df[col].head(3).tolist()}")
    print(f"=== END DEBUG: Raw attendance DataFrame ===\n")
    
    # Clean Student# in both DataFrames to ensure matching
    print("=== DEBUG: Cleaning Student# columns ===")
    grades_df = clean_student_id(grades_df)
    attendance_df = clean_student_id(attendance_df)
    print(f"After cleaning - grades_df Student# sample: {grades_df['Student#'].head(3).tolist() if 'Student#' in grades_df.columns else 'N/A'}")
    print(f"After cleaning - attendance_df Student# sample: {attendance_df['Student#'].head(3).tolist() if 'Student#' in attendance_df.columns else 'N/A'}")
    
    # Clean attendance DataFrame (remove totals, convert times, normalize percentages)
    print("=== DEBUG: Cleaning attendance DataFrame ===")
    attendance_df = clean_attendance_df(attendance_df)
    print(f"After cleaning attendance_df, shape: {attendance_df.shape}")
    
    # Drop rows missing Student# or Student Name
    print("=== DEBUG: Dropping rows with missing Student# or Student Name ===")
    initial_grades = len(grades_df)
    initial_attendance = len(attendance_df)
    
    if 'Student#' in grades_df.columns and 'Student Name' in grades_df.columns:
        grades_df = grades_df.dropna(subset=["Student#", "Student Name"])
        print(f"Grades: Dropped {initial_grades - len(grades_df)} rows with missing Student# or Student Name")
    
    if 'Student#' in attendance_df.columns and 'Student Name' in attendance_df.columns:
        attendance_df = attendance_df.dropna(subset=["Student#", "Student Name"])
        print(f"Attendance: Dropped {initial_attendance - len(attendance_df)} rows with missing Student# or Student Name")
    
    # Ensure both DataFrames use the same data type for Student#
    if 'Student#' in grades_df.columns and 'Student#' in attendance_df.columns:
        try:
            grades_df["Student#"] = grades_df["Student#"].fillna(0).astype(int)
            attendance_df["Student#"] = attendance_df["Student#"].fillna(0).astype(int)
            print(f"Converted Student# to int - grades sample: {grades_df['Student#'].head(3).tolist()}, attendance sample: {attendance_df['Student#'].head(3).tolist()}")
        except (ValueError, TypeError) as e:
            print(f"Warning: Could not convert Student# to int: {e}, keeping as numeric")
            grades_df["Student#"] = pd.to_numeric(grades_df["Student#"], errors='coerce').fillna(0)
            attendance_df["Student#"] = pd.to_numeric(attendance_df["Student#"], errors='coerce').fillna(0)
    
    # Add Campus Login URL column to attendance_df (after column normalization)
    # First, ensure Student# is numeric for consistent matching
    if 'Student#' in attendance_df.columns:
        # Convert to numeric, then to string for mapping
        try:
            student_ids = pd.to_numeric(attendance_df['Student#'], errors='coerce').fillna(0).astype(int).astype(str)
        except (ValueError, TypeError):
            student_ids = attendance_df['Student#'].astype(str).str.strip()
        attendance_df['Campus Login URL'] = student_ids.map(
            lambda x: attendance_hyperlinks.get(x, None)
        )
    else:
        # Find Student# column
        student_id_col = None
        for col in attendance_df.columns:
            if 'student#' in col.lower().strip():
                student_id_col = col
                break
        if student_id_col:
            try:
                student_ids = pd.to_numeric(attendance_df[student_id_col], errors='coerce').fillna(0).astype(int).astype(str)
            except (ValueError, TypeError):
                student_ids = attendance_df[student_id_col].astype(str).str.strip()
            attendance_df['Campus Login URL'] = student_ids.map(
                lambda x: attendance_hyperlinks.get(x, None)
            )
        else:
            attendance_df['Campus Login URL'] = None
    
    # Validate required columns (case-insensitive)
    required_grades_cols = ['Student#', 'Student Name', 'Program Name', 'current overall Program Grade']
    required_attendance_cols = ['Student#', 'Student Name', 'Scheduled Hours to Date', 
                                'Attended Hours to Date', 'Attended % to Date.', 
                                'Missed Hours to Date', '% Missed', 'Missed Minus Excused to date']
    
    # Check for column name variations (case-insensitive, handle whitespace)
    grades_cols_lower = {col.lower().strip(): col for col in grades_df.columns}
    attendance_cols_lower = {col.lower().strip(): col for col in attendance_df.columns}
    
    missing_grades = []
    missing_attendance = []
    
    for req_col in required_grades_cols:
        if req_col.lower().strip() not in grades_cols_lower:
            missing_grades.append(req_col)
    
    for req_col in required_attendance_cols:
        if req_col.lower().strip() not in attendance_cols_lower:
            missing_attendance.append(req_col)
    
    if missing_grades or missing_attendance:
        error_msg = "Missing required columns:\n"
        if missing_grades:
            error_msg += f"  Grades sheet: {missing_grades}\n"
        if missing_attendance:
            error_msg += f"  Attendance sheet: {missing_attendance}\n"
        error_msg += f"\nFound columns:\n"
        error_msg += f"  Grades: {list(grades_df.columns)}\n"
        error_msg += f"  Attendance: {list(attendance_df.columns)}"
        raise ValueError(error_msg)
    
    return grades_df, attendance_df, name_hyperlinks


def normalize_data(grades_df: pd.DataFrame, attendance_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Normalize data: convert percentages, durations, and clean student IDs.
    
    Args:
        grades_df: Raw grades dataframe
        attendance_df: Raw attendance dataframe
    
    Returns:
        Tuple of normalized dataframes
    """
    # Normalize Grades sheet
    grades_normalized = grades_df.copy()

    # Find Student# column (case-insensitive)
    student_id_col = None
    for col in grades_normalized.columns:
        if 'student#' in col.lower().strip():
            student_id_col = col
            break

    if student_id_col:
        # Convert Student# to numeric (int) for consistent matching
        try:
            # First try to convert to numeric, handling any string values
            student_ids = pd.to_numeric(grades_normalized[student_id_col], errors='coerce')
            # Replace NaN with 0, but keep track of which ones failed
            grades_normalized['Student#'] = student_ids.fillna(0).astype(int)
            # If we got all zeros, the conversion probably failed - use string instead
            if (grades_normalized['Student#'] == 0).all():
                print("WARNING: All Student# values converted to 0 in grades, using string format instead")
                grades_normalized['Student#'] = grades_normalized[student_id_col].astype(str).str.strip()
            else:
                print(f"Successfully converted grades Student# to int, sample: {grades_normalized['Student#'].head(3).tolist()}")
        except (ValueError, TypeError) as e:
            print(f"Warning: Could not convert grades Student# to numeric: {e}, using string format")
            # Fallback to string if conversion fails
            grades_normalized['Student#'] = grades_normalized[student_id_col].astype(str).str.strip()
    
    # Normalize grade percentage (find column case-insensitively)
    # "current overall Program Grade" is between 0-1, needs ×100
    grade_col = None
    for col in grades_normalized.columns:
        if 'current overall program grade' in col.lower().strip():
            grade_col = col
            break
    
    if grade_col:
        # Apply normalize_pct to convert 0-1 to 0-100
        print(f"DEBUG: Normalizing grade column '{grade_col}' - Sample values before: {grades_normalized[grade_col].head(3).tolist()}")
        grades_normalized['grade_pct'] = grades_normalized[grade_col].apply(normalize_pct)
        print(f"DEBUG: After normalization - Sample values: {grades_normalized['grade_pct'].head(3).tolist()}")
    else:
        print("WARNING: 'current overall Program Grade' column not found")
        grades_normalized['grade_pct'] = 0.0
    
    # Normalize Attendance sheet
    attendance_normalized = attendance_df.copy()

    # Preserve Campus Login URL column if it exists
    campus_login_url_col = None
    for col in attendance_normalized.columns:
        if 'campus login url' in col.lower().strip():
            campus_login_url_col = col
            break

    # Find Student# column (case-insensitive)
    student_id_col = None
    for col in attendance_normalized.columns:
        if 'student#' in col.lower().strip():
            student_id_col = col
            break

    if student_id_col:
        # Convert Student# to numeric (int) for consistent matching
        try:
            # First try to convert to numeric, handling any string values
            student_ids = pd.to_numeric(attendance_normalized[student_id_col], errors='coerce')
            # Replace NaN with 0, but keep track of which ones failed
            attendance_normalized['Student#'] = student_ids.fillna(0).astype(int)
            # If we got all zeros, the conversion probably failed - use string instead
            if (attendance_normalized['Student#'] == 0).all():
                print("WARNING: All Student# values converted to 0 in attendance, using string format instead")
                attendance_normalized['Student#'] = attendance_normalized[student_id_col].astype(str).str.strip()
            else:
                print(f"Successfully converted attendance Student# to int, sample: {attendance_normalized['Student#'].head(3).tolist()}")
        except (ValueError, TypeError) as e:
            print(f"Warning: Could not convert attendance Student# to numeric: {e}, using string format")
            # Fallback to string if conversion fails
            attendance_normalized['Student#'] = attendance_normalized[student_id_col].astype(str).str.strip()

    # Preserve Campus Login URL column
    if campus_login_url_col and campus_login_url_col != 'Campus Login URL':
        attendance_normalized['Campus Login URL'] = attendance_normalized[campus_login_url_col]
    elif 'Campus Login URL' not in attendance_normalized.columns:
        attendance_normalized['Campus Login URL'] = None

    # PREPROCESSING STEP: Automatically clean and normalize attendance data
    # This converts HH:MM strings to decimal hours, normalizes percentages, and replaces invalid values
    print(f"\n=== DEBUG: Before normalize_attendance_data ===")
    print(f"Columns in attendance_normalized: {list(attendance_normalized.columns)}")
    if 'Attended % to Date.' in attendance_normalized.columns:
        print(f"Sample 'Attended % to Date.' values: {attendance_normalized['Attended % to Date.'].head(3).tolist()}")
    attendance_normalized = normalize_attendance_data(attendance_normalized)
    
    # Extract attendance_pct from normalized "Attended % to Date." column
    if 'Attended % to Date.' in attendance_normalized.columns:
        attendance_normalized['attendance_pct'] = attendance_normalized['Attended % to Date.']
        print(f"DEBUG: Set attendance_pct from 'Attended % to Date.' - Sample: {attendance_normalized['attendance_pct'].head(3).tolist()}")
    else:
        # Try case-insensitive search for the column
        found_col = None
        for col in attendance_normalized.columns:
            if 'attended' in col.lower() and '%' in col and 'date' in col.lower():
                found_col = col
                print(f"DEBUG: Found attendance percentage column: '{found_col}'")
                attendance_normalized['attendance_pct'] = attendance_normalized[found_col]
                print(f"DEBUG: Set attendance_pct from '{found_col}' - Sample: {attendance_normalized['attendance_pct'].head(3).tolist()}")
                break
        
        if not found_col:
            # Column not found - try to calculate from hours
            print("DEBUG: 'Attended % to Date.' column not found, trying to calculate from hours")
            if 'Attended Hours to Date' in attendance_normalized.columns and 'Scheduled Hours to Date' in attendance_normalized.columns:
                # Calculate attendance percentage from hours
                scheduled_hours = attendance_normalized['Scheduled Hours to Date']
                attended_hours = attendance_normalized['Attended Hours to Date']
                
                print(f"DEBUG: Scheduled Hours sample: {scheduled_hours.head(3).tolist()}")
                print(f"DEBUG: Attended Hours sample: {attended_hours.head(3).tolist()}")
                
                # Calculate percentage, handling division by zero
                attendance_normalized['attendance_pct'] = (
                    (attended_hours / scheduled_hours.replace(0, np.nan) * 100.0)
                    .replace([np.inf, -np.inf, np.nan], 0.0)
                    .clip(0.0, 100.0)
                )
                print(f"DEBUG: Calculated attendance_pct from hours - Sample: {attendance_normalized['attendance_pct'].head(3).tolist()}")
            else:
                # Last resort: set to 0
                print("DEBUG: Could not find hours columns either, setting attendance_pct to 0")
                attendance_normalized['attendance_pct'] = 0.0
    
    # Extract missed_pct from normalized "% Missed" column
    if '% Missed' in attendance_normalized.columns:
        attendance_normalized['missed_pct'] = attendance_normalized['% Missed']
    else:
        attendance_normalized['missed_pct'] = 0.0
    
    # Debug: Print sample data to verify transformations
    if 'Student Name' in attendance_normalized.columns:
        sample_cols = ['Student Name']
        if 'Attended Hours to Date' in attendance_normalized.columns:
            sample_cols.append('Attended Hours to Date')
        if 'Attended % to Date.' in attendance_normalized.columns:
            sample_cols.append('Attended % to Date.')
        print("\n=== Attendance DataFrame Sample (after normalization) ===")
        print(attendance_normalized[sample_cols].head())
        print("========================================================\n")
    
    return grades_normalized, attendance_normalized


def merge_data(grades_df: pd.DataFrame, attendance_df: pd.DataFrame) -> pd.DataFrame:
    """
    Merge grades and attendance data on Student#.
    
    Prefers Grades sheet name for display.
    Preserves Campus Login URL from attendance sheet.
    """
    # Debug: Check what we're merging
    print(f"\n=== DEBUG: merge_data - BEFORE MERGE ===")
    print(f"grades_df columns: {list(grades_df.columns)}")
    print(f"grades_df shape: {grades_df.shape}")
    if 'attendance_pct' in grades_df.columns:
        print(f"grades_df has attendance_pct: {grades_df['attendance_pct'].head(3).tolist()}")
    
    print(f"attendance_df columns: {list(attendance_df.columns)}")
    print(f"attendance_df shape: {attendance_df.shape}")
    if 'attendance_pct' in attendance_df.columns:
        print(f"attendance_df has attendance_pct: {attendance_df['attendance_pct'].head(3).tolist()}")
    if 'Attended % to Date.' in attendance_df.columns:
        print(f"attendance_df has 'Attended % to Date.': {attendance_df['Attended % to Date.'].head(3).tolist()}")
    
    # Check Student# columns
    if 'Student#' in grades_df.columns:
        print(f"grades_df Student# sample: {grades_df['Student#'].head(3).tolist()}")
        print(f"grades_df Student# dtype: {grades_df['Student#'].dtype}")
    if 'Student#' in attendance_df.columns:
        print(f"attendance_df Student# sample: {attendance_df['Student#'].head(3).tolist()}")
        print(f"attendance_df Student# dtype: {attendance_df['Student#'].dtype}")
    
    # Ensure Student# types match in both DataFrames before merging
    # Check if Student# is already numeric or string
    if 'Student#' in grades_df.columns and 'Student#' in attendance_df.columns:
        grades_dtype = grades_df['Student#'].dtype
        attendance_dtype = attendance_df['Student#'].dtype
        
        print(f"Before merge conversion - grades_df Student# dtype: {grades_dtype}, attendance_df Student# dtype: {attendance_dtype}")
        
        # If one is numeric and one is string, convert both to string for matching
        if (pd.api.types.is_numeric_dtype(grades_dtype) and not pd.api.types.is_numeric_dtype(attendance_dtype)) or \
           (not pd.api.types.is_numeric_dtype(grades_dtype) and pd.api.types.is_numeric_dtype(attendance_dtype)):
            print("WARNING: Student# types don't match, converting both to string for matching")
            grades_df['Student#'] = grades_df['Student#'].astype(str).str.strip()
            attendance_df['Student#'] = attendance_df['Student#'].astype(str).str.strip()
        elif not pd.api.types.is_numeric_dtype(grades_dtype) and not pd.api.types.is_numeric_dtype(attendance_dtype):
            # Both are strings, ensure they're clean
            grades_df['Student#'] = grades_df['Student#'].astype(str).str.strip()
            attendance_df['Student#'] = attendance_df['Student#'].astype(str).str.strip()
        else:
            # Both are numeric, ensure they're the same type
            try:
                grades_df['Student#'] = pd.to_numeric(grades_df['Student#'], errors='coerce').fillna(0).astype(int)
                attendance_df['Student#'] = pd.to_numeric(attendance_df['Student#'], errors='coerce').fillna(0).astype(int)
                print(f"Converted both to int, grades sample: {grades_df['Student#'].head(3).tolist()}, attendance sample: {attendance_df['Student#'].head(3).tolist()}")
            except (ValueError, TypeError) as e:
                print(f"Warning: Could not convert to numeric: {e}, using string format")
                grades_df['Student#'] = grades_df['Student#'].astype(str).str.strip()
                attendance_df['Student#'] = attendance_df['Student#'].astype(str).str.strip()
    
    # Merge on Student# using outer join to include ALL students from both sheets
    # This ensures students with only grades OR only attendance are included
    print(f"Before merge - grades_df: {len(grades_df)} rows, attendance_df: {len(attendance_df)} rows")
    print(f"Unique Student# in grades: {grades_df['Student#'].nunique()}, in attendance: {attendance_df['Student#'].nunique()}")
    
    merged = pd.merge(
        grades_df,
        attendance_df,
        on='Student#',
        how='outer',  # Use outer join to include ALL students from both sheets
        suffixes=('_grades', '_attendance')
    )
    
    print(f"After merge with outer join, merged shape: {merged.shape}")
    print(f"Total unique students after merge: {merged['Student#'].nunique()}")
    
    # If merge resulted in no records, something is wrong
    if len(merged) == 0:
        print("WARNING: Outer join resulted in 0 records. This should not happen.")
        print(f"grades_df Student# unique count: {grades_df['Student#'].nunique()}")
        print(f"attendance_df Student# unique count: {attendance_df['Student#'].nunique()}")
        print(f"grades_df Student# sample values: {sorted(grades_df['Student#'].unique().tolist())[:10]}")
        print(f"attendance_df Student# sample values: {sorted(attendance_df['Student#'].unique().tolist())[:10]}")
        
        raise ValueError(
            f"No records found after merge. "
            f"Grades sheet has {len(grades_df)} rows, "
            f"Attendance sheet has {len(attendance_df)} rows. "
            f"Student# values may not match between sheets."
        )
    
    print(f"\n=== DEBUG: merge_data - AFTER MERGE ===")
    print(f"merged shape: {merged.shape}")
    print(f"merged columns: {list(merged.columns)}")
    if 'attendance_pct' in merged.columns:
        print(f"merged has attendance_pct: {merged['attendance_pct'].head(3).tolist()}")
    if 'attendance_pct_attendance' in merged.columns:
        print(f"merged has attendance_pct_attendance: {merged['attendance_pct_attendance'].head(3).tolist()}")
    if 'attendance_pct_grades' in merged.columns:
        print(f"merged has attendance_pct_grades: {merged['attendance_pct_grades'].head(3).tolist()}")
    if 'Attended % to Date.' in merged.columns:
        print(f"merged has 'Attended % to Date.': {merged['Attended % to Date.'].head(3).tolist()}")
    
    # Prefer Grades sheet name, but use attendance if grades is missing
    if 'Student Name_grades' in merged.columns and 'Student Name_attendance' in merged.columns:
        merged['Student Name'] = merged['Student Name_grades'].fillna(merged['Student Name_attendance'])
        merged = merged.drop(columns=['Student Name_grades', 'Student Name_attendance'])
    elif 'Student Name_grades' in merged.columns:
        merged['Student Name'] = merged['Student Name_grades']
        merged = merged.drop(columns=['Student Name_grades'])
    elif 'Student Name_attendance' in merged.columns:
        merged['Student Name'] = merged['Student Name_attendance']
        merged = merged.drop(columns=['Student Name_attendance'])
    else:
        # If neither exists, create empty column
        merged['Student Name'] = 'Unknown'
    
    # Prefer Grades sheet program name, but use attendance if grades is missing
    if 'Program Name' in merged.columns:
        # Already from grades, keep as is
        pass
    elif 'Program Name_grades' in merged.columns:
        merged['Program Name'] = merged['Program Name_grades']
        merged = merged.drop(columns=['Program Name_grades'])
    elif 'Program Name_attendance' in merged.columns:
        merged['Program Name'] = merged['Program Name_attendance']
        merged = merged.drop(columns=['Program Name_attendance'])
    else:
        # If neither exists, create empty column
        merged['Program Name'] = 'Unknown'
    
    # Preserve Campus Login URL from attendance sheet (if it exists)
    if 'Campus Login URL' not in merged.columns:
        # Check if it exists with suffix
        if 'Campus Login URL_attendance' in merged.columns:
            merged['Campus Login URL'] = merged['Campus Login URL_attendance']
            merged = merged.drop(columns=['Campus Login URL_attendance'])
        else:
            merged['Campus Login URL'] = None
    
    # Preserve attendance_pct from attendance sheet (critical!)
    # Also ensure grade_pct is preserved from grades sheet
    print(f"\n=== DEBUG: merge_data - PRESERVING attendance_pct and grade_pct ===")
    
    # Handle attendance_pct
    if 'attendance_pct' not in merged.columns:
        print("attendance_pct not in merged.columns, checking for suffixed versions...")
        # Check if it exists with suffix
        if 'attendance_pct_attendance' in merged.columns:
            print(f"Found attendance_pct_attendance, values: {merged['attendance_pct_attendance'].head(3).tolist()}")
            merged['attendance_pct'] = merged['attendance_pct_attendance'].fillna(0.0)
            merged = merged.drop(columns=['attendance_pct_attendance'])
            print(f"Set attendance_pct from attendance_pct_attendance: {merged['attendance_pct'].head(3).tolist()}")
        elif 'attendance_pct_grades' in merged.columns:
            print(f"Found attendance_pct_grades, values: {merged['attendance_pct_grades'].head(3).tolist()}")
            merged['attendance_pct'] = merged['attendance_pct_grades'].fillna(0.0)
            merged = merged.drop(columns=['attendance_pct_grades'])
            print(f"Set attendance_pct from attendance_pct_grades: {merged['attendance_pct'].head(3).tolist()}")
        else:
            # Try to find any attendance percentage column
            print("Searching for attendance percentage columns...")
            found_col = None
            for col in merged.columns:
                col_lower = str(col).lower()
                if 'attended' in col_lower and ('%' in str(col) or 'pct' in col_lower):
                    found_col = col
                    print(f"Found column '{col}', values: {merged[col].head(3).tolist()}")
                    merged['attendance_pct'] = merged[col].fillna(0.0)
                    print(f"Set attendance_pct from '{col}': {merged['attendance_pct'].head(3).tolist()}")
                    break
            if not found_col:
                print("WARNING: No attendance percentage column found, setting to 0.0")
                merged['attendance_pct'] = 0.0
    else:
        merged['attendance_pct'] = merged['attendance_pct'].fillna(0.0)
        print(f"attendance_pct already in merged.columns, values: {merged['attendance_pct'].head(3).tolist()}")
    
    # Handle grade_pct - ensure it exists and fill missing values with 0
    if 'grade_pct' not in merged.columns:
        if 'grade_pct_grades' in merged.columns:
            merged['grade_pct'] = merged['grade_pct_grades'].fillna(0.0)
            merged = merged.drop(columns=['grade_pct_grades'])
        elif 'grade_pct_attendance' in merged.columns:
            merged['grade_pct'] = merged['grade_pct_attendance'].fillna(0.0)
            merged = merged.drop(columns=['grade_pct_attendance'])
        else:
            merged['grade_pct'] = 0.0
    else:
        merged['grade_pct'] = merged['grade_pct'].fillna(0.0)
    
    print(f"Final merged shape: {merged.shape}, grade_pct non-null: {merged['grade_pct'].notna().sum()}, attendance_pct non-null: {merged['attendance_pct'].notna().sum()}")
    print(f"=== END DEBUG: merge_data ===\n")
    
    # Deduplicate by Student# (keep last row)
    merged = merged.drop_duplicates(subset=['Student#'], keep='last')
    
    return merged

