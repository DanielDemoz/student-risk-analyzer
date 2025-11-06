"""Excel file parsing and data normalization."""

import pandas as pd
import numpy as np
import re
from typing import Tuple, Dict, Optional
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell


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
                student_id_cell = row[student_id_col - 1]
                student_name_cell = row[student_name_col - 1]
                
                if student_id_cell.value and student_name_cell.value:
                    student_id = str(student_id_cell.value).strip()
                    hyperlink = extract_hyperlink_from_cell(student_name_cell)
                    if hyperlink:
                        name_hyperlinks[student_id] = hyperlink
    
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
            student_id_cell = row[att_student_id_col - 1]
            student_name_cell = row[att_student_name_col - 1]
            
            if student_id_cell.value and student_name_cell.value:
                student_id = str(student_id_cell.value).strip()
                hyperlink = extract_hyperlink_from_cell(student_name_cell)
                if hyperlink:
                    attendance_hyperlinks[student_id] = hyperlink
                    # Attendance sheet hyperlinks take precedence over grades sheet
                    name_hyperlinks[student_id] = hyperlink
    
    # Load attendance dataframe
    attendance_df = pd.read_excel(excel_file, sheet_name=attendance_sheet_name, engine='openpyxl')
    
    # Normalize column names (strip whitespace, handle variations) BEFORE adding Campus Login URL
    grades_df.columns = grades_df.columns.str.strip()
    attendance_df.columns = attendance_df.columns.str.strip()
    
    # Add Campus Login URL column to attendance_df (after column normalization)
    if 'Student#' in attendance_df.columns:
        attendance_df['Campus Login URL'] = attendance_df['Student#'].astype(str).str.strip().map(
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
            attendance_df['Campus Login URL'] = attendance_df[student_id_col].astype(str).str.strip().map(
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
        grades_normalized['Student#'] = grades_normalized[student_id_col].astype(str).str.strip()
    
    # Normalize grade percentage (find column case-insensitively)
    grade_col = None
    for col in grades_normalized.columns:
        if 'current overall program grade' in col.lower().strip():
            grade_col = col
            break
    
    if grade_col:
        max_grade = grades_normalized[grade_col].max()
        if pd.notna(max_grade):
            grades_normalized['grade_pct'] = grades_normalized[grade_col].apply(
                lambda x: normalize_percentage(x, max_grade)
            )
        else:
            grades_normalized['grade_pct'] = 0.0
    else:
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
        attendance_normalized['Student#'] = attendance_normalized[student_id_col].astype(str).str.strip()
    
    # Preserve Campus Login URL column
    if campus_login_url_col and campus_login_url_col != 'Campus Login URL':
        attendance_normalized['Campus Login URL'] = attendance_normalized[campus_login_url_col]
    elif 'Campus Login URL' not in attendance_normalized.columns:
        attendance_normalized['Campus Login URL'] = None
    
    # Convert time-based columns from "HH:MM" strings to numeric hours
    # Find and convert "Scheduled Hours to Date"
    scheduled_col = None
    for col in attendance_normalized.columns:
        if 'scheduled hours' in col.lower().strip() and 'date' in col.lower().strip():
            scheduled_col = col
            break
    
    if scheduled_col:
        attendance_normalized['Scheduled Hours to Date'] = attendance_normalized[scheduled_col].apply(to_hours)
    
    # Find and convert "Attended Hours to Date"
    attended_hours_col = None
    for col in attendance_normalized.columns:
        if 'attended hours' in col.lower().strip() and 'date' in col.lower().strip():
            attended_hours_col = col
            break
    
    if attended_hours_col:
        attendance_normalized['Attended Hours to Date'] = attendance_normalized[attended_hours_col].apply(to_hours)
    
    # Find and convert "Missed Hours to Date"
    missed_hours_col = None
    for col in attendance_normalized.columns:
        if 'missed hours' in col.lower().strip() and 'date' in col.lower().strip():
            missed_hours_col = col
            break
    
    if missed_hours_col:
        attendance_normalized['Missed Hours to Date'] = attendance_normalized[missed_hours_col].apply(to_hours)
    
    # Find and convert "Missed Minus Excused to date"
    missed_excused_col = None
    for col in attendance_normalized.columns:
        if 'missed minus excused' in col.lower().strip() and 'date' in col.lower().strip():
            missed_excused_col = col
            break
    
    if missed_excused_col:
        attendance_normalized['Missed Minus Excused to date'] = attendance_normalized[missed_excused_col].apply(to_hours)
    
    # Normalize "Attended % to Date." column
    att_pct_col = None
    possible_patterns = [
        'attended % to date.',
        'attended% to date.',
        'attended % to date',
        'attended% to date',
        'attendance %',
        'attendance%',
        'attended %',
        'attended%'
    ]
    
    for pattern in possible_patterns:
        for col in attendance_normalized.columns:
            col_lower = col.lower().strip()
            if pattern in col_lower:
                att_pct_col = col
                break
        if att_pct_col:
            break
    
    if att_pct_col:
        # Apply normalize_attendance_pct function
        attendance_normalized['Attended % to Date.'] = attendance_normalized[att_pct_col].apply(normalize_attendance_pct)
        attendance_normalized['attendance_pct'] = attendance_normalized['Attended % to Date.']
    else:
        # Column not found - try to calculate from hours
        if 'Attended Hours to Date' in attendance_normalized.columns and 'Scheduled Hours to Date' in attendance_normalized.columns:
            # Calculate attendance percentage from hours
            scheduled_hours = attendance_normalized['Scheduled Hours to Date']
            attended_hours = attendance_normalized['Attended Hours to Date']
            
            # Calculate percentage, handling division by zero
            attendance_normalized['attendance_pct'] = (
                (attended_hours / scheduled_hours.replace(0, np.nan) * 100.0)
                .replace([np.inf, -np.inf, np.nan], 0.0)
                .clip(0.0, 100.0)
            )
        else:
            # Last resort: set to 0
            attendance_normalized['attendance_pct'] = 0.0
    
    # Normalize "% Missed" column if it exists
    missed_pct_col = None
    for col in attendance_normalized.columns:
        col_lower = col.lower().strip()
        if '% missed' in col_lower or '%missed' in col_lower:
            missed_pct_col = col
            break
    
    if missed_pct_col:
        attendance_normalized['% Missed'] = attendance_normalized[missed_pct_col].apply(normalize_attendance_pct)
        attendance_normalized['missed_pct'] = attendance_normalized['% Missed']
    else:
        attendance_normalized['missed_pct'] = 0.0
    
    # Ensure all numeric columns are clean (no NaN, Infinity)
    numeric_cols = attendance_normalized.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        attendance_normalized[col] = attendance_normalized[col].replace([np.inf, -np.inf, np.nan], 0.0)
    
    return grades_normalized, attendance_normalized


def merge_data(grades_df: pd.DataFrame, attendance_df: pd.DataFrame) -> pd.DataFrame:
    """
    Merge grades and attendance data on Student#.
    
    Prefers Grades sheet name for display.
    Preserves Campus Login URL from attendance sheet.
    """
    # Merge on Student#
    merged = pd.merge(
        grades_df,
        attendance_df,
        on='Student#',
        how='left',
        suffixes=('_grades', '_attendance')
    )
    
    # Prefer Grades sheet name
    if 'Student Name_grades' in merged.columns and 'Student Name_attendance' in merged.columns:
        merged['Student Name'] = merged['Student Name_grades'].fillna(merged['Student Name_attendance'])
        merged = merged.drop(columns=['Student Name_grades', 'Student Name_attendance'])
    elif 'Student Name_grades' in merged.columns:
        merged['Student Name'] = merged['Student Name_grades']
        merged = merged.drop(columns=['Student Name_grades'])
    elif 'Student Name_attendance' in merged.columns:
        merged['Student Name'] = merged['Student Name_attendance']
        merged = merged.drop(columns=['Student Name_attendance'])
    
    # Prefer Grades sheet program name
    if 'Program Name' in merged.columns:
        pass  # Already from grades
    elif 'Program Name_attendance' in merged.columns:
        merged['Program Name'] = merged['Program Name_attendance']
        merged = merged.drop(columns=['Program Name_attendance'])
    
    # Preserve Campus Login URL from attendance sheet (if it exists)
    if 'Campus Login URL' not in merged.columns:
        # Check if it exists with suffix
        if 'Campus Login URL_attendance' in merged.columns:
            merged['Campus Login URL'] = merged['Campus Login URL_attendance']
            merged = merged.drop(columns=['Campus Login URL_attendance'])
        else:
            merged['Campus Login URL'] = None
    
    # Preserve attendance_pct from attendance sheet (critical!)
    if 'attendance_pct' not in merged.columns:
        # Check if it exists with suffix
        if 'attendance_pct_attendance' in merged.columns:
            merged['attendance_pct'] = merged['attendance_pct_attendance']
            merged = merged.drop(columns=['attendance_pct_attendance'])
        elif 'attendance_pct_grades' in merged.columns:
            merged['attendance_pct'] = merged['attendance_pct_grades']
            merged = merged.drop(columns=['attendance_pct_grades'])
        else:
            # Try to find any attendance percentage column
            for col in merged.columns:
                if 'attendance' in str(col).lower() and ('%' in str(col) or 'pct' in str(col).lower()):
                    merged['attendance_pct'] = merged[col]
                    break
            else:
                # If still not found, set to 0
                merged['attendance_pct'] = 0.0
    
    # Deduplicate by Student# (keep last row)
    merged = merged.drop_duplicates(subset=['Student#'], keep='last')
    
    return merged

