"""Data models for the Student Risk Analyzer application."""

from typing import Optional, Dict, List
from pydantic import BaseModel


class StudentRiskResult(BaseModel):
    """Individual student risk analysis result."""
    student_id: str
    student_name: str
    program_name: str
    grade_pct: float
    attendance_pct: float
    risk_score: float
    risk_category: str
    simple_rule_flagged: bool
    campus_login_url: str
    explanation: Optional[str] = None


class UploadResponse(BaseModel):
    """Response from file upload endpoint."""
    success: bool
    message: str
    results: List[StudentRiskResult]
    summary: Dict[str, int]


class EmailDraftRequest(BaseModel):
    """Request for email draft generation."""
    student_id: str
    risk_category: str
    program: str
    grade_pct: float
    attendance_pct: float


class EmailDraftResponse(BaseModel):
    """Email draft response."""
    subject: str
    body: str

