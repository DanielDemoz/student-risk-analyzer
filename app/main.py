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
        # Load and parse Excel
        try:
            grades_df, attendance_df, name_hyperlinks = load_excel(file_bytes)
        except Exception as e:
            error_msg = f"Error loading Excel file: {str(e)}"
            print(f"ERROR: {error_msg}")
            print(f"Exception type: {type(e).__name__}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=error_msg)
        
        # Verify Student Name exists in both DataFrames
        if 'Student Name' not in grades_df.columns:
            print(f"ERROR: 'Student Name' missing in grades_df. Columns: {list(grades_df.columns)}")
            raise HTTPException(
                status_code=400,
                detail=f"'Student Name' column missing in grades sheet. Found columns: {list(grades_df.columns)}"
            )
        
        if 'Student Name' not in attendance_df.columns:
            print(f"ERROR: 'Student Name' missing in attendance_df. Columns: {list(attendance_df.columns)}")
            raise HTTPException(
                status_code=400,
                detail=f"'Student Name' column missing in attendance sheet. Found columns: {list(attendance_df.columns)}"
            )
        
        # Normalize data
        try:
            grades_normalized, attendance_normalized = normalize_data(grades_df, attendance_df)
        except Exception as e:
            error_msg = f"Error normalizing data: {str(e)}"
            print(f"ERROR: {error_msg}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=error_msg)
        
        # Merge data
        try:
            merged_df = merge_data(grades_normalized, attendance_normalized)
        except Exception as e:
            error_msg = f"Error merging data: {str(e)}"
            print(f"ERROR: {error_msg}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=error_msg)
        
        print(f"Processed: {len(merged_df)} students")
        
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
        results = []
        for row_idx, (_, row) in enumerate(merged_df.iterrows()):
            if row_idx < 5:  # Debug first 5 rows
                print(f"DEBUG: Processing row {row_idx}, Student#: {row.get('Student#')}, Student Name: {row.get('Student Name', 'Unknown')}")
            student_id = str(row['Student#']).strip()
            
            # Handle Student Name - clean NaN values
            student_name_val = row.get('Student Name', 'Unknown')
            if pd.isna(student_name_val) or str(student_name_val).strip().lower() in ['nan', 'none', '']:
                student_name = 'Unknown'
            else:
                student_name = str(student_name_val).strip()
            
            # Handle Program Name - clean NaN values
            program_name_val = row.get('Program Name', 'Unknown')
            if pd.isna(program_name_val) or str(program_name_val).strip().lower() in ['nan', 'none', '']:
                program_name = 'Unknown'
            else:
                program_name = str(program_name_val).strip()
            
            # Clean numeric values to ensure JSON compliance
            grade_pct = clean_numeric_value(row.get('grade_pct', 0))
            
            # Get attendance_pct - try multiple possible column names
            attendance_pct = 0.0
            if 'attendance_pct' in row.index:
                attendance_pct = clean_numeric_value(row.get('attendance_pct', 0))
            elif 'attendance_pct_attendance' in row.index:
                attendance_pct = clean_numeric_value(row.get('attendance_pct_attendance', 0))
            elif 'attendance_pct_grades' in row.index:
                attendance_pct = clean_numeric_value(row.get('attendance_pct_grades', 0))
            else:
                # Try to find any column with attendance percentage
                for col in row.index:
                    if 'attended' in str(col).lower() and ('%' in str(col) or 'pct' in str(col).lower()):
                        attendance_pct = clean_numeric_value(row.get(col, 0))
                        break
            
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
            try:
                base_category = categories[row_idx] if row_idx < len(categories) else 'Low'
            except (IndexError, KeyError, TypeError):
                base_category = 'Low'
            
            # Override risk category if data is missing
            if data_status == 'Missing Both':
                risk_category = 'Insufficient Data'
            elif data_status == 'Missing Grade':
                risk_category = 'Missing Grade'
            elif data_status == 'Missing Attendance':
                risk_category = 'Missing Attendance'
            else:
                risk_category = base_category
            
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
            
            result = StudentRiskResult(
                student_id=student_id,
                student_name=student_name,
                program_name=program_name,
                grade_pct=grade_pct,
                attendance_pct=attendance_pct,
                risk_score=risk_score,
                risk_category=risk_category,
                simple_rule_flagged=simple_rule_flagged,
                campus_login_url=campus_login_url,
                data_status=data_status,
                explanation=explanation
            )
            results.append(result)
            if row_idx < 5:  # Debug first 5 results
                print(f"DEBUG: Added result {row_idx}: {student_name}, data_status: {data_status}, risk_category: {risk_category}")
        
        print(f"DEBUG: Total results built: {len(results)}")
        
        # Sort by data status first (Complete > Missing Attendance > Missing Grade > Missing Both)
        # Then by risk score (descending) - highest risk first
        status_priority = {'Complete': 0, 'Missing Attendance': 1, 'Missing Grade': 2, 'Missing Both': 3}
        results.sort(key=lambda x: (status_priority.get(x.data_status, 3), -x.risk_score))
        
        # Generate session ID
        session_id = datetime.now().isoformat()
        results_cache[session_id] = results
        
        # Summary counts
        summary = {
            'Failed': sum(1 for r in results if r.risk_category == 'Extremely High Risk'),
            'High': sum(1 for r in results if r.risk_category == 'High'),
            'Medium': sum(1 for r in results if r.risk_category == 'Medium'),
            'Low': sum(1 for r in results if r.risk_category == 'Low'),
            'Total': len(results),
            'At Risk (Simple Rule)': sum(1 for r in results if r.simple_rule_flagged)
        }
        
        print(f"Results: {summary['Total']} students ({summary['Failed']} Extremely High Risk, {summary['High']} High, {summary['Medium']} Medium, {summary['Low']} Low risk)")
        
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
        'High': sum(1 for r in results if r.risk_category == 'High'),
        'Medium': sum(1 for r in results if r.risk_category == 'Medium'),
        'Low': sum(1 for r in results if r.risk_category == 'Low'),
        'Total': len(results)
    }
    
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

