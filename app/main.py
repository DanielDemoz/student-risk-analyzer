"""FastAPI main application for Student Risk Analyzer."""

import os
import json
import csv
from io import BytesIO, StringIO
from typing import Dict, List, Optional
from datetime import datetime

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.exceptions import RequestValidationError
from starlette.exceptions import HTTPException as StarletteHTTPException
from dotenv import load_dotenv
import pandas as pd
import numpy as np
import traceback

from app.models import StudentRiskResult, UploadResponse, EmailDraftRequest, EmailDraftResponse
from app.parsers import load_excel, normalize_data, merge_data
from app.risk import simple_rule, train_or_fallback_score, get_risk_category, get_explanation
from app.email_templates import generate_email_draft

# Load environment variables
load_dotenv()

app = FastAPI(title="Student Risk Analyzer", version="1.0.0")

# CORS configuration
allow_origins = os.getenv('ALLOW_ORIGINS', '*').split(',')
app.add_middleware(
    CORSMiddleware,
    allow_origins=allow_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Override default exception handlers to return JSON (register specific handlers first)
@app.exception_handler(StarletteHTTPException)
async def http_exception_handler_json(request: Request, exc: StarletteHTTPException):
    """Handle HTTP exceptions and return JSON."""
    return JSONResponse(
        status_code=exc.status_code,
        content={"detail": exc.detail}
    )

@app.exception_handler(RequestValidationError)
async def validation_exception_handler_json(request: Request, exc: RequestValidationError):
    """Handle validation errors and return JSON."""
    return JSONResponse(
        status_code=422,
        content={"detail": exc.errors(), "body": exc.body}
    )

# Global exception handler to ensure JSON responses for unhandled exceptions
# This will only catch exceptions not handled by the specific handlers above
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    """Handle all unhandled exceptions and return JSON."""
    error_detail = str(exc)
    # In production, you might want to hide traceback details
    if os.getenv('DEBUG', 'False').lower() == 'true':
        error_detail = f"{str(exc)}\n\n{traceback.format_exc()}"
    
    return JSONResponse(
        status_code=500,
        content={
            "detail": f"Internal server error: {error_detail}",
            "type": type(exc).__name__
        }
    )

# Configuration
CAMPUS_LOGIN_BASE_URL = os.getenv(
    'CAMPUS_LOGIN_BASE_URL',
    'https://compuslogin.example.com?student_id={student_id}'
)

# Parse risk thresholds (including 'failed' category)
thresholds_str = os.getenv('RISK_THRESHOLDS', 'low:0,medium:60,high:80,failed:90')
RISK_THRESHOLDS = {}
for item in thresholds_str.split(','):
    key, value = item.split(':')
    RISK_THRESHOLDS[key.strip()] = float(value.strip())

MAX_UPLOAD_SIZE_MB = int(os.getenv('MAX_UPLOAD_SIZE_MB', '10'))
MAX_UPLOAD_SIZE = MAX_UPLOAD_SIZE_MB * 1024 * 1024

# In-memory storage for results (session-based)
results_cache: Dict[str, List[StudentRiskResult]] = {}


def build_campus_login_url(student_id: str, hyperlink: Optional[str] = None) -> str:
    """Build campus login URL from hyperlink or base URL."""
    if hyperlink:
        return hyperlink
    return CAMPUS_LOGIN_BASE_URL.format(student_id=student_id)


def clean_numeric_value(value) -> float:
    """
    Clean numeric values to ensure JSON compliance.
    Replaces NaN, Infinity, and -Infinity with 0.0.
    """
    if pd.isna(value) or value is None:
        return 0.0
    try:
        val = float(value)
        if np.isnan(val) or np.isinf(val):
            return 0.0
        return val
    except (ValueError, TypeError):
        return 0.0


@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the main HTML page."""
    html_path = os.path.join(os.path.dirname(__file__), 'static', 'index.html')
    if os.path.exists(html_path):
        with open(html_path, 'r', encoding='utf-8') as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>Student Risk Analyzer</h1><p>Static files not found. Please ensure the static directory exists.</p>")


@app.get("/health")
async def health_check():
    """Health check endpoint to test server connectivity."""
    return JSONResponse(content={"status": "ok", "message": "Server is running"})


@app.post("/debug-columns")
async def debug_columns(file: UploadFile = File(...)):
    """Debug endpoint to see what columns are in the Excel file."""
    try:
        file_bytes = await file.read()
        grades_df, attendance_df, name_hyperlinks = load_excel(file_bytes)
        
        return JSONResponse(content={
            "grades_columns": list(grades_df.columns),
            "attendance_columns": list(attendance_df.columns),
            "grades_sample": grades_df.head(3).to_dict(orient='records') if len(grades_df) > 0 else [],
            "attendance_sample": attendance_df.head(3).to_dict(orient='records') if len(attendance_df) > 0 else []
        })
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e), "traceback": traceback.format_exc()}
        )


@app.post("/upload", response_model=UploadResponse)
async def upload_file(
    file: UploadFile = File(...)
):
    """
    Upload and process Excel file.
    
    Accepts Excel file with 'Students Grade' and attendance sheets.
    Returns processed results with risk scores.
    """
    # Check file size
    file_bytes = await file.read()
    if len(file_bytes) > MAX_UPLOAD_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Maximum size: {MAX_UPLOAD_SIZE_MB}MB"
        )
    
    # Check file type
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Please upload an Excel file (.xlsx or .xls)"
        )
    
    try:
        # Load and parse Excel - use original functions but only read required columns
        try:
            grades_df, attendance_df, name_hyperlinks = load_excel(file_bytes)
        except Exception as e:
            error_msg = f"Error loading Excel file: {str(e)}"
            print(f"ERROR: {error_msg}")
            print(f"Exception type: {type(e).__name__}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=error_msg)
        
        # DO NOT filter columns here - let normalize_and_rename_columns handle it
        # The normalization function will ensure all required columns exist
        # We just need to verify they exist after normalization
        print(f"DEBUG: Before normalization - grades_df columns: {list(grades_df.columns)}")
        print(f"DEBUG: Before normalization - attendance_df columns: {list(attendance_df.columns)}")
        
        # Normalize data (convert percentages, etc.)
        try:
            grades_normalized, attendance_normalized = normalize_data(grades_df, attendance_df)
        except Exception as e:
            error_msg = f"Error normalizing data: {str(e)}"
            print(f"ERROR: {error_msg}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=error_msg)
        
        # Merge data using INNER JOIN
        try:
            merged_df = merge_data(grades_normalized, attendance_normalized)
        except Exception as e:
            error_msg = f"Error merging data: {str(e)}"
            print(f"ERROR: {error_msg}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=error_msg)
        
        print(f"Processed: {len(merged_df)} students")
        
        # Audit and validate risk scoring calculations
        from app.risk import audit_and_recalculate_risk
        try:
            merged_df_audit = audit_and_recalculate_risk(merged_df)
            # Use audited values if available for validation
            if 'Risk Score' in merged_df_audit.columns:
                merged_df['audited_risk_score'] = merged_df_audit['Risk Score']
            if 'Weighted Index P' in merged_df_audit.columns:
                merged_df['performance_index'] = merged_df_audit['Weighted Index P']
            print("✅ Risk scoring audit completed successfully")
        except Exception as e:
            print(f"⚠️ Risk scoring audit warning: {e}")
            # Continue with normal processing
        
        if len(merged_df) == 0:
            # Provide more detailed error message
            error_detail = (
                "No valid student records found in the Excel file after merging. "
                f"Grades sheet has {len(grades_normalized)} rows, "
                f"Attendance sheet has {len(attendance_normalized)} rows. "
                "This usually means Student# values don't match between the two sheets. "
                "Please check that Student# values are consistent in both sheets."
            )
            print(f"ERROR: {error_detail}")
            raise HTTPException(
                status_code=400,
                detail=error_detail
            )
        
        # Clean invalid or missing numeric values before JSON conversion
        # Replace NaN, Infinity, and -Infinity with 0 for numeric columns
        numeric_cols = merged_df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            merged_df[col] = merged_df[col].fillna(0)
            merged_df[col] = merged_df[col].replace([np.inf, -np.inf, np.nan], 0)
        
        # Check for target labels
        has_target = 'is_at_risk' in merged_df.columns
        
        # Compute risk scores
        feature_cols = []
        if 'grade_pct' in merged_df.columns:
            feature_cols.append('grade_pct')
        if 'attendance_pct' in merged_df.columns:
            feature_cols.append('attendance_pct')
        if 'missed_pct' in merged_df.columns:
            feature_cols.append('missed_pct')
        if 'Missed Hours to Date_hours' in merged_df.columns:
            feature_cols.append('Missed Hours to Date_hours')
        
        # Compute risk scores using ML model or fallback heuristic
        risk_scores, categories, model, scaler = train_or_fallback_score(
            merged_df, RISK_THRESHOLDS, has_target
        )
        
        # Clean risk_scores array to ensure no NaN or Infinity values
        if isinstance(risk_scores, np.ndarray):
            risk_scores = np.nan_to_num(risk_scores, nan=0.0, posinf=100.0, neginf=0.0)
            risk_scores = np.clip(risk_scores, 0.0, 100.0)
        else:
            risk_scores = [clean_numeric_value(score) for score in risk_scores]
        
        # Reset index to ensure sequential indexing for risk_scores and categories
        merged_df = merged_df.reset_index(drop=True)
        
        print(f"DEBUG: Before building results - merged_df shape: {merged_df.shape}")
        print(f"DEBUG: risk_scores length: {len(risk_scores)}")
        print(f"DEBUG: categories length: {len(categories)}")
        
        # Build results - use enumerate to get sequential index
        # Ensure index is reset to avoid serial number in output
        merged_df = merged_df.reset_index(drop=True)
        
        # With simplified loading, columns are already standardized
        # No need for column renaming - merged_df already has: Student ID, Student Name, Program, Grade %, Attendance %
        
        # Debug: Print column names to understand merge structure
        # After merge_data, Student Name should be consolidated into a single column
        print(f"\n=== DEBUG: After merge_data ===")
        print(f"DEBUG: Merged DataFrame columns: {list(merged_df.columns)}")
        print(f"DEBUG: Merged DataFrame shape: {merged_df.shape}")
        if len(merged_df) > 0:
            print(f"DEBUG: First row sample columns: {list(merged_df.iloc[0].index)}")
            print(f"DEBUG: First row Student#: {merged_df.iloc[0].get('Student#', 'NOT FOUND')}")
            print(f"DEBUG: First row Student Name: {merged_df.iloc[0].get('Student Name', 'NOT FOUND')}")
            # Check for Student Name columns
            student_name_cols = [col for col in merged_df.columns if 'Student Name' in col or 'Name' in col]
            student_id_cols = [col for col in merged_df.columns if 'Student#' in col or 'Student ID' in col]
            print(f"DEBUG: Student Name columns found: {student_name_cols}")
            print(f"DEBUG: Student ID columns found: {student_id_cols}")
            
            # Check if Student Name column has actual values (not all Unknown)
            if 'Student Name' in merged_df.columns:
                student_names = merged_df['Student Name'].astype(str)
                unknown_count = (student_names == 'Unknown').sum()
                empty_count = (student_names.str.strip() == '').sum()
                valid_count = len(merged_df) - unknown_count - empty_count
                print(f"DEBUG: Student Name statistics - Total: {len(merged_df)}, Valid: {valid_count}, Unknown: {unknown_count}, Empty: {empty_count}")
                if valid_count > 0:
                    valid_names = student_names[student_names.str.strip() != '']
                    valid_names = valid_names[valid_names != 'Unknown']
                    print(f"DEBUG: Sample Valid Student Names (first 10): {valid_names.head(10).tolist()}")
                if unknown_count > 0:
                    print(f"WARNING: {unknown_count} rows have 'Unknown' as Student Name!")
                    # Show first few rows with Unknown names
                    unknown_rows = merged_df[merged_df['Student Name'].astype(str) == 'Unknown']
                    if len(unknown_rows) > 0:
                        print(f"DEBUG: First 3 rows with Unknown names:")
                        for idx, (_, row) in enumerate(unknown_rows.head(3).iterrows()):
                            print(f"  Row {idx}: Student#={row.get('Student#', 'N/A')}, Student Name={row.get('Student Name', 'N/A')}")
            else:
                print(f"ERROR: 'Student Name' column NOT found in merged DataFrame!")
        print(f"=== END DEBUG: After merge_data ===\n")
        
        results = []
        for row_idx, (_, row) in enumerate(merged_df.iterrows()):
            if row_idx < 5:  # Debug first 5 rows
                print(f"DEBUG: Processing row {row_idx}")
                print(f"  Available columns: {list(row.index)}")
                # Print all Student# related columns
                student_id_cols = [col for col in row.index if 'Student' in col and ('#' in col or 'ID' in col or 'id' in col.lower())]
                print(f"  Student ID related columns: {student_id_cols}")
                for col in student_id_cols:
                    print(f"    {col}: {row.get(col)}")
            
            # Extract Student ID - handle merge suffixes (_grades, _attendance)
            # NOTE: normalize_data() renames Student# to Student ID, so check Student ID first
            # IMPORTANT: Do NOT use DataFrame index - it's just a row number, not the Student ID
            student_id_val = None
            student_id_source = None
            
            # Priority: Student ID (from normalize_data) > Student# (legacy) > suffixed versions
            if 'Student ID' in row.index and pd.notna(row.get('Student ID')):
                student_id_val = row.get('Student ID')
                student_id_source = 'Student ID'
            elif 'Student#' in row.index and pd.notna(row.get('Student#')):
                student_id_val = row.get('Student#')
                student_id_source = 'Student#'
            elif 'Student ID_grades' in row.index and pd.notna(row.get('Student ID_grades')):
                student_id_val = row.get('Student ID_grades')
                student_id_source = 'Student ID_grades'
            elif 'Student ID_attendance' in row.index and pd.notna(row.get('Student ID_attendance')):
                student_id_val = row.get('Student ID_attendance')
                student_id_source = 'Student ID_attendance'
            elif 'Student#_grades' in row.index and pd.notna(row.get('Student#_grades')):
                student_id_val = row.get('Student#_grades')
                student_id_source = 'Student#_grades'
            elif 'Student#_attendance' in row.index and pd.notna(row.get('Student#_attendance')):
                student_id_val = row.get('Student#_attendance')
                student_id_source = 'Student#_attendance'
            else:
                # Try alternative column names
                for col in ['Student Number', 'student_id']:
                    if col in row.index and pd.notna(row.get(col)):
                        student_id_val = row.get(col)
                        student_id_source = col
                        break
            
            if student_id_val is None or pd.isna(student_id_val):
                student_id = 'Unknown'
                if row_idx < 5:
                    print(f"  WARNING: No Student ID found in row {row_idx}")
            else:
                # Convert to string and strip - this is the numeric ID
                student_id_str = str(student_id_val).strip()
                # Remove any decimal point if it's a float (e.g., "609571.0" -> "609571")
                if '.' in student_id_str and student_id_str.replace('.', '').isdigit():
                    student_id_str = student_id_str.split('.')[0]
                
                # Validate: Real Student IDs are typically 6-7 digits (e.g., 5686877, 609571)
                # If it's a small number (< 1000), it might be the index or wrong column
                try:
                    student_id_num = int(float(student_id_str))
                    if student_id_num < 1000:
                        print(f"WARNING: Row {row_idx} - Student ID '{student_id_str}' from '{student_id_source}' is suspiciously small (< 1000). This might be wrong column.")
                        # Try to find the actual Student ID in other columns
                        for col in row.index:
                            if 'student' in col.lower() and ('#' in col or 'id' in col.lower() or 'number' in col.lower()):
                                if col == student_id_source:  # Skip the one we already tried
                                    continue
                                alt_val = row.get(col)
                                if pd.notna(alt_val):
                                    alt_str = str(alt_val).strip()
                                    # Remove decimal point if present
                                    if '.' in alt_str and alt_str.replace('.', '').isdigit():
                                        alt_str = alt_str.split('.')[0]
                                    try:
                                        alt_num = int(float(alt_str))
                                        if alt_num >= 1000:  # More likely to be a real ID
                                            print(f"  Found alternative Student ID in '{col}': {alt_str} (using this instead)")
                                            student_id_str = alt_str
                                            student_id_source = col
                                            break
                                    except (ValueError, TypeError):
                                        pass
                    else:
                        # Valid Student ID (>= 1000)
                        if row_idx < 5:
                            print(f"  Using Student ID from '{student_id_source}': {student_id_str}")
                except (ValueError, TypeError):
                    # Not a number, but might still be valid
                    if row_idx < 5:
                        print(f"  Student ID '{student_id_str}' is not numeric, but using it anyway")
                student_id = student_id_str
            
            # Extract Student Name - AFTER merge, Student Name should already be consolidated by merge_data
            # Priority: Student Name (consolidated) > Student Name_grades > Student Name_attendance
            student_name_val = None
            student_name_source = None
            
            # First check for consolidated Student Name (this should exist after merge_data)
            if 'Student Name' in row.index:
                student_name_val = row.get('Student Name')
                student_name_source = 'Student Name'
                if row_idx < 5 and (pd.isna(student_name_val) or str(student_name_val).strip().lower() in ['nan', 'none', '']):
                    print(f"  WARNING: Student Name column exists but is empty/null: '{student_name_val}'")
            
            # Fallback to suffixed versions if consolidated column doesn't exist or is empty
            if (student_name_val is None or pd.isna(student_name_val) or str(student_name_val).strip().lower() in ['nan', 'none', '']) and 'Student Name_grades' in row.index:
                student_name_val = row.get('Student Name_grades')
                student_name_source = 'Student Name_grades'
                if row_idx < 5:
                    print(f"  Using Student Name from grades suffix: '{student_name_val}'")
            
            if (student_name_val is None or pd.isna(student_name_val) or str(student_name_val).strip().lower() in ['nan', 'none', '']) and 'Student Name_attendance' in row.index:
                student_name_val = row.get('Student Name_attendance')
                student_name_source = 'Student Name_attendance'
                if row_idx < 5:
                    print(f"  Using Student Name from attendance suffix: '{student_name_val}'")
            
            # Last resort: try alternative column names
            if student_name_val is None or pd.isna(student_name_val) or str(student_name_val).strip().lower() in ['nan', 'none', '']:
                for col in ['Name', 'student_name', 'Full Name']:
                    if col in row.index and pd.notna(row.get(col)):
                        alt_val = row.get(col)
                        if str(alt_val).strip().lower() not in ['nan', 'none', '']:
                            student_name_val = alt_val
                            student_name_source = col
                            if row_idx < 5:
                                print(f"  Found Student Name in alternative column '{col}': '{student_name_val}'")
                            break
            
            # Final check and assignment
            if student_name_val is None or pd.isna(student_name_val) or str(student_name_val).strip().lower() in ['nan', 'none', '']:
                student_name = 'Unknown'
                if row_idx < 5:
                    print(f"  ERROR: No Student Name found in row {row_idx} from any source!")
                    print(f"    Available columns: {[col for col in row.index if 'name' in col.lower() or 'student' in col.lower()]}")
            else:
                student_name = str(student_name_val).strip()
                if row_idx < 5:
                    print(f"  ✅ Extracted Student Name from '{student_name_source}': '{student_name}'")
            
            # Final safety check: if Student Name looks like a number (ID), it might be misaligned
            # Only swap if we're CERTAIN they're swapped (both conditions must be true):
            # 1. Student Name is all digits and >= 1000 (looks like an ID)
            # 2. Student ID is < 1000 (looks like an index, not a real ID)
            # IMPORTANT: If we can't find a better name, KEEP the original (don't set to Unknown)
            try:
                # Check if student_name is all digits and looks like an ID
                if student_name.isdigit():
                    student_name_as_id = int(float(student_name))
                    if student_name_as_id >= 1000:  # Looks like a Student ID (6-7 digits)
                        # Check if Student ID is suspiciously small - if so, they're swapped!
                        try:
                            student_id_num = int(float(student_id))
                            if student_id_num < 1000:
                                # Both conditions met: Name has ID, ID has small number -> SWAP
                                print(f"WARNING: Row {row_idx} - Detected column swap: Student Name='{student_name}' (ID) and Student ID='{student_id}' (small). SWAPPING!")
                                temp_id = student_id
                                student_id = student_name  # Use the name value as the ID
                                
                                # Try to find the actual Student Name in other columns
                                found_real_name = False
                                for col in row.index:
                                    if 'name' in col.lower() and 'student' in col.lower() and col != student_name_source:
                                        alt_val = row.get(col)
                                        if pd.notna(alt_val):
                                            alt_str = str(alt_val).strip()
                                            # If it's not all digits and not empty, it might be the actual name
                                            if not alt_str.isdigit() and alt_str.lower() not in ['nan', 'none', '']:
                                                print(f"  Found actual Student Name in '{col}': {alt_str}")
                                                student_name = alt_str
                                                found_real_name = True
                                                break
                                
                                # If we couldn't find a real name, check if temp_id is a valid name
                                if not found_real_name:
                                    # temp_id might actually be a name if it's not numeric
                                    if not temp_id.isdigit() and temp_id.lower() not in ['nan', 'none', '']:
                                        student_name = temp_id
                                        print(f"  Using swapped Student ID as name: {student_name}")
                                    else:
                                        # Last resort: keep the numeric ID as the name (better than Unknown)
                                        print(f"  WARNING: Could not find actual name. Keeping numeric ID '{student_name}' as name (better than Unknown)")
                                        # Don't set to Unknown - keep the ID as the name
                        except (ValueError, TypeError):
                            # Student ID is not numeric, so don't swap
                            pass
                    # If student_name is numeric but < 1000, it might be wrong, but don't auto-swap
                    # Just log a warning but keep the name
                    elif student_name_as_id < 1000:
                        if row_idx < 5:
                            print(f"  Note: Student Name '{student_name}' is numeric but < 1000. Keeping as-is.")
                else:
                    # Student Name is not all digits - it's likely a real name, keep it!
                    if row_idx < 5:
                        print(f"  Student Name '{student_name}' looks valid (not all digits), keeping it")
            except (AttributeError, ValueError, TypeError) as e:
                # If validation fails, keep the original student_name
                if row_idx < 5:
                    print(f"  Note: Could not validate Student Name format: {e}, keeping original: {student_name}")
                pass
            
            if row_idx < 5:  # Debug first 5 rows
                print(f"  Student# value: {row.get('Student#', 'NOT FOUND')}")
                print(f"  Student Name value: {row.get('Student Name', 'NOT FOUND')}")
                print(f"  Final - Student ID: {student_id}, Student Name: {student_name}")
                print(f"  Final - Student ID type: {type(student_id)}, Student Name type: {type(student_name)}")
                print(f"  Final - Student ID value: '{student_id}', Student Name value: '{student_name}'")
            
            # Handle Program Name - clean NaN values
            program_name_val = row.get('Program Name', 'Unknown')
            if pd.isna(program_name_val) or str(program_name_val).strip().lower() in ['nan', 'none', '']:
                program_name = 'Unknown'
            else:
                program_name = str(program_name_val).strip()
            
            # Clean numeric values to ensure JSON compliance
            # Use grade_pct and attendance_pct columns (created by normalize_data)
            grade_pct = clean_numeric_value(row.get('grade_pct', 0))
            
            # Get attendance percentage
            attendance_pct = clean_numeric_value(row.get('attendance_pct', 0))
            
            # Get data status
            data_status = str(row.get('data_status', 'Complete')).strip()
            if data_status not in ['Complete', 'Missing Grade', 'Missing Attendance', 'Missing Both']:
                data_status = 'Complete'
            
            # Use row_idx (sequential) instead of DataFrame index
            try:
                risk_score = clean_numeric_value(risk_scores[row_idx] if row_idx < len(risk_scores) else 0)
            except (IndexError, KeyError, TypeError):
                risk_score = 0.0
            
            # Adjust risk category based on data status
            # Use direct grade/attendance classification for more accurate results
            from app.risk import classify_risk_by_grade_attendance, get_risk_color
            
            try:
                if data_status == 'Missing Both':
                    risk_category = 'Insufficient Data'
                    risk_color = '#6B7280'  # Gray for missing data
                elif data_status == 'Missing Grade':
                    risk_category = 'Missing Grade'
                    risk_color = '#6B7280'  # Gray for missing data
                elif data_status == 'Missing Attendance':
                    risk_category = 'Missing Attendance'
                    risk_color = '#6B7280'  # Gray for missing data
                elif data_status == 'Complete':
                    # Use direct classification based on grade and attendance
                    risk_category = classify_risk_by_grade_attendance(grade_pct, attendance_pct)
                    risk_color = get_risk_color(risk_category)
                else:
                    # Fallback to risk score-based category
                    base_category = categories[row_idx] if row_idx < len(categories) else 'Low'
                    risk_category = base_category
                    risk_color = get_risk_color(risk_category)
            except (IndexError, KeyError, TypeError) as e:
                print(f"Warning: Error classifying risk for row {row_idx}: {e}")
                risk_category = 'Low'
                risk_color = get_risk_color('Low')
            
            # Simple rule only applies if both values are available
            if data_status == 'Complete':
                simple_rule_flagged = simple_rule(grade_pct, attendance_pct)
            else:
                # If data is missing, flag based on available data
                if data_status == 'Missing Grade':
                    simple_rule_flagged = attendance_pct < 70.0
                elif data_status == 'Missing Attendance':
                    simple_rule_flagged = grade_pct < 70.0
                else:
                    simple_rule_flagged = False
            
            # Get Campus Login URL from attendance sheet (preferred) or grades sheet hyperlink
            campus_login_url_from_df = row.get('Campus Login URL')
            if pd.notna(campus_login_url_from_df) and campus_login_url_from_df:
                campus_login_url = str(campus_login_url_from_df)
            else:
                hyperlink = name_hyperlinks.get(student_id)
                campus_login_url = build_campus_login_url(student_id, hyperlink)
            
            # Get explanation - use row_idx for sequential access
            explanation = get_explanation(row_idx, merged_df, model, scaler, feature_cols)
            
            # Final validation before creating result object
            # Ensure Student ID is numeric string (not index) and Student Name is text (not ID)
            if row_idx < 5:
                print(f"  Before creating result: student_id='{student_id}', student_name='{student_name}'")
                print(f"    student_id is numeric: {student_id.isdigit() if isinstance(student_id, str) else False}")
                print(f"    student_name is numeric: {student_name.isdigit() if isinstance(student_name, str) else False}")
            
            result = StudentRiskResult(
                student_id=str(student_id),  # Ensure it's a string
                student_name=str(student_name),  # Ensure it's a string
                program_name=program_name,
                grade_pct=grade_pct,
                attendance_pct=attendance_pct,
                risk_score=risk_score,
                risk_category=risk_category,
                risk_color=risk_color,
                simple_rule_flagged=simple_rule_flagged,
                campus_login_url=campus_login_url,
                data_status=data_status,
                explanation=explanation
            )
            results.append(result)
            
            if row_idx < 5:  # Debug first 5 results
                print(f"  After creating result: student_id='{result.student_id}', student_name='{result.student_name}'")
                print(f"DEBUG: Added result {row_idx}: {student_name}, data_status: {data_status}, risk_category: {risk_category}")
        
        print(f"DEBUG: Total results built: {len(results)}")
        
        # Sort by data status first (Complete > Missing Attendance > Missing Grade > Missing Both)
        # Then by risk score (descending) - highest risk first
        status_priority = {'Complete': 0, 'Missing Attendance': 1, 'Missing Grade': 2, 'Missing Both': 3}
        results.sort(key=lambda x: (status_priority.get(x.data_status, 3), -x.risk_score))
        
        # Generate session ID
        session_id = datetime.now().isoformat()
        results_cache[session_id] = results
        
        # Summary counts - match exact category names from classify_risk_by_grade_attendance
        summary = {
            'Failed': sum(1 for r in results if r.risk_category == 'Extremely High Risk'),
            'High': sum(1 for r in results if r.risk_category == 'High Risk'),
            'Medium': sum(1 for r in results if r.risk_category == 'Medium Risk'),
            'Low': sum(1 for r in results if r.risk_category == 'Low Risk'),
            'Total': len(results),
            'At Risk (Simple Rule)': sum(1 for r in results if r.simple_rule_flagged)
        }
        
        # Also count old category names for backward compatibility
        summary['High'] += sum(1 for r in results if r.risk_category == 'High')
        summary['Medium'] += sum(1 for r in results if r.risk_category == 'Medium')
        summary['Low'] += sum(1 for r in results if r.risk_category == 'Low')
        
        print(f"Results: {summary['Total']} students ({summary['Failed']} Extremely High Risk, {summary['High']} High Risk, {summary['Medium']} Medium Risk, {summary['Low']} Low Risk)")
        
        return UploadResponse(
            success=True,
            message=f"Successfully processed {len(results)} students",
            results=results,
            summary=summary
        )
    
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing file: {str(e)}"
        )


@app.get("/results")
async def get_results():
    """Get the last processed results."""
    if not results_cache:
        raise HTTPException(status_code=404, detail="No results available")
    
    # Get most recent session
    latest_session = max(results_cache.keys())
    results = results_cache[latest_session]
    
    summary = {
        'Failed': sum(1 for r in results if r.risk_category == 'Extremely High Risk'),
        'High': sum(1 for r in results if r.risk_category == 'High Risk'),
        'Medium': sum(1 for r in results if r.risk_category == 'Medium Risk'),
        'Low': sum(1 for r in results if r.risk_category == 'Low Risk'),
        'Total': len(results)
    }
    
    # Also count old category names for backward compatibility
    summary['High'] += sum(1 for r in results if r.risk_category == 'High')
    summary['Medium'] += sum(1 for r in results if r.risk_category == 'Medium')
    summary['Low'] += sum(1 for r in results if r.risk_category == 'Low')
    
    return {
        'session_id': latest_session,
        'results': [r.dict() for r in results],
        'summary': summary
    }


@app.post("/email-draft", response_model=EmailDraftResponse)
async def generate_email_draft_endpoint(request: EmailDraftRequest):
    """Generate email draft for a student."""
    try:
        # Get student name from results if available
        student_name = "Student"
        if results_cache:
            latest_session = max(results_cache.keys())
            for result in results_cache[latest_session]:
                if result.student_id == request.student_id:
                    student_name = result.student_name
                    break
        
        email = generate_email_draft(
            student_name=student_name,
            program=request.program,
            grade_pct=request.grade_pct,
            attendance_pct=request.attendance_pct,
            risk_category=request.risk_category
        )
        
        return EmailDraftResponse(**email)
    
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating email draft: {str(e)}"
        )


@app.get("/download.csv")
async def download_csv():
    """Download processed results as CSV."""
    if not results_cache:
        raise HTTPException(status_code=404, detail="No results available")
    
    # Get most recent session
    latest_session = max(results_cache.keys())
    results = results_cache[latest_session]
    
    # Create CSV
    output = StringIO()
    writer = csv.writer(output)
    
    # Write header
    writer.writerow([
        'Student ID',
        'Student Name',
        'Program',
        'Grade %',
        'Attendance %',
        'Risk Score',
        'Risk Category',
        'Campus Login URL'
    ])
    
    # Write rows
    for result in results:
        writer.writerow([
            result.student_id,
            result.student_name,
            result.program_name,
            f"{result.grade_pct:.2f}",
            f"{result.attendance_pct:.2f}",
            f"{result.risk_score:.2f}",
            result.risk_category,
            result.campus_login_url
        ])
    
    output.seek(0)
    
    return StreamingResponse(
        iter([output.getvalue()]),
        media_type="text/csv",
        headers={
            "Content-Disposition": f"attachment; filename=student_risk_results_{latest_session[:10]}.csv"
        }
    )


# Mount static files
static_path = os.path.join(os.path.dirname(__file__), 'static')
if os.path.exists(static_path):
    app.mount("/static", StaticFiles(directory=static_path), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

