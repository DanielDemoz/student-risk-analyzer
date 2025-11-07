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
from app.parsers import load_excel, normalize_data
from app.risk import simple_rule, classify_risk_by_grade_attendance, get_risk_color
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
    """Debug endpoint to inspect the uploaded Excel file."""
    try:
        file_bytes = await file.read()
        df, _ = load_excel(file_bytes)
        return JSONResponse(content={
            "columns": list(df.columns),
            "sample": df.head(3).to_dict(orient="records") if not df.empty else []
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
    """Upload and process the single-sheet Excel file."""
    # Check file size
    file_bytes = await file.read()
    if len(file_bytes) > MAX_UPLOAD_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Maximum size: {MAX_UPLOAD_SIZE_MB}MB"
        )

    # Check file type
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Please upload an Excel file (.xlsx or .xls)"
        )

    try:
        # Load and normalize the single worksheet
        try:
            raw_df, name_hyperlinks = load_excel(file_bytes)
            student_df = normalize_data(raw_df)
        except Exception as e:
            error_msg = f"Error loading Excel file: {str(e)}"
            print(f"ERROR: {error_msg}")
            print(f"Exception type: {type(e).__name__}")
            print(f"Traceback: {traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=error_msg)

        if student_df.empty:
            raise HTTPException(status_code=400, detail="No student records found in the uploaded file.")

        # Attach hyperlinks if present
        student_df["Campus Login URL"] = student_df["Student#"].map(name_hyperlinks)

        # Performance index: 80% grade, 20% attendance
        grade_pct = student_df["grade_pct"].astype(float)
        attendance_pct = student_df["attendance_pct"].astype(float)

        grade_fraction = (grade_pct / 100.0).clip(lower=0.0, upper=1.0)
        attendance_fraction = (attendance_pct / 100.0).clip(lower=0.0, upper=1.0)
        performance_index = (0.8 * grade_fraction + 0.2 * attendance_fraction).clip(lower=0.0, upper=1.0)
        risk_scores = ((1.0 - performance_index) * 100.0).round(1)

        categories = [
            classify_risk_by_grade_attendance(g, a)
            for g, a in zip(grade_pct.tolist(), attendance_pct.tolist())
        ]
        colors = [get_risk_color(cat) for cat in categories]

        results: List[StudentRiskResult] = []
        summary_counts = {
            'No Risk': 0,
            'Medium Risk': 0,
            'High Risk': 0,
            'Extremely High Risk': 0
        }
        simple_rule_count = 0

        for idx, row in student_df.iterrows():
            student_id = str(row.get("Student#", "")).strip() or "Unknown"
            student_name = str(row.get("Student Name", "")).strip() or "Unknown"
            program_name = str(row.get("Program Name", "")).strip() or "Unknown"

            grade_value = clean_numeric_value(row.get("grade_pct", 0.0))
            attendance_value = clean_numeric_value(row.get("attendance_pct", 0.0))
            risk_score = clean_numeric_value(risk_scores.iloc[idx])

            risk_category = categories[idx]
            risk_color = colors[idx]

            if risk_category == 'Extremely High Risk':
                summary_counts['Extremely High Risk'] += 1
            elif risk_category == 'High Risk':
                summary_counts['High Risk'] += 1
            elif risk_category == 'Medium Risk':
                summary_counts['Medium Risk'] += 1
            elif risk_category == 'No Risk':
                summary_counts['No Risk'] += 1

            is_simple_rule = simple_rule(grade_value, attendance_value)
            if is_simple_rule:
                simple_rule_count += 1

            hyperlink = row.get("Campus Login URL")
            campus_login_url = build_campus_login_url(student_id, hyperlink)

            result = StudentRiskResult(
                student_id=student_id,
                student_name=student_name,
                program_name=program_name,
                grade_pct=grade_value,
                attendance_pct=attendance_value,
                risk_score=risk_score,
                risk_category=risk_category,
                risk_color=risk_color,
                simple_rule_flagged=is_simple_rule,
                campus_login_url=campus_login_url,
                data_status='Complete',
                explanation=None
            )
            results.append(result)

        # Sort by risk score descending (higher risk first)
        results.sort(key=lambda r: (-r.risk_score, r.student_name))

        # Cache results for downstream endpoints (download/email)
        session_id = datetime.now().isoformat()
        results_cache[session_id] = results

        summary = {
            'No Risk': summary_counts['No Risk'],
            'Medium Risk': summary_counts['Medium Risk'],
            'High Risk': summary_counts['High Risk'],
            'Extremely High Risk': summary_counts['Extremely High Risk'],
            'Total': len(results),
            'At Risk (Simple Rule)': simple_rule_count
        }

        print(
            f"Results: {summary['Total']} students ("
            f"{summary['Extremely High Risk']} Extremely High Risk, {summary['High Risk']} High Risk, "
            f"{summary['Medium Risk']} Medium Risk, {summary['No Risk']} No Risk)"
        )

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
        'No Risk': sum(1 for r in results if r.risk_category == 'No Risk'),
        'Medium Risk': sum(1 for r in results if r.risk_category == 'Medium Risk'),
        'High Risk': sum(1 for r in results if r.risk_category == 'High Risk'),
        'Extremely High Risk': sum(1 for r in results if r.risk_category == 'Extremely High Risk'),
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

