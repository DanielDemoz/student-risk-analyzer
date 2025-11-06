"""Email template generation for different risk categories."""

import os
from typing import Dict


def get_advisor_info() -> Dict[str, str]:
    """Get advisor name and email from environment or defaults."""
    return {
        'name': os.getenv('ADVISOR_NAME', 'Academic Advisor'),
        'email': os.getenv('ADVISOR_EMAIL', 'advisor@example.com')
    }


def generate_email_draft(
    student_name: str,
    program: str,
    grade_pct: float,
    attendance_pct: float,
    risk_category: str
) -> Dict[str, str]:
    """
    Generate email draft based on risk category.
    
    Args:
        student_name: Student's name
        program: Program name
        grade_pct: Grade percentage (0-100)
        attendance_pct: Attendance percentage (0-100)
        risk_category: 'High', 'Medium', or 'Low'
    
    Returns:
        Dict with 'subject' and 'body' keys
    """
    advisor = get_advisor_info()
    
    # Format percentages
    grade_str = f"{grade_pct:.1f}"
    attendance_str = f"{attendance_pct:.1f}"
    
    if risk_category.lower() == 'high':
        return _high_risk_email(student_name, program, grade_str, attendance_str, advisor)
    elif risk_category.lower() == 'medium':
        return _medium_risk_email(student_name, program, grade_str, attendance_str, advisor)
    else:
        return _low_risk_email(student_name, program, grade_str, attendance_str, advisor)


def _high_risk_email(
    student_name: str,
    program: str,
    grade_pct: str,
    attendance_pct: str,
    advisor: Dict[str, str]
) -> Dict[str, str]:
    """Generate high risk warning email."""
    subject = f"Action Needed: Let's get you back on track in {program}"
    
    body = f"""Hi {student_name},

We're concerned about your recent progress in {program}. Your current grade is {grade_pct}% and attendance is {attendance_pct}%.

Let's meet this week to make a plan—tutoring, time management, and catch-up resources are available.

Please reply with your availability for a 15–20 minute check-in.

Best,

{advisor['name']}
{advisor['email']}"""
    
    return {'subject': subject, 'body': body}


def _medium_risk_email(
    student_name: str,
    program: str,
    grade_pct: str,
    attendance_pct: str,
    advisor: Dict[str, str]
) -> Dict[str, str]:
    """Generate medium risk notice email."""
    subject = f"Quick check-in for {program}"
    
    body = f"""Hi {student_name},

I'm checking in about {program}. You're currently at {grade_pct}% with {attendance_pct}% attendance.

If you'd like, we can review upcoming deadlines and support options to keep you on track.

Let me know a good time to connect.

Best,

{advisor['name']}
{advisor['email']}"""
    
    return {'subject': subject, 'body': body}


def _low_risk_email(
    student_name: str,
    program: str,
    grade_pct: str,
    attendance_pct: str,
    advisor: Dict[str, str]
) -> Dict[str, str]:
    """Generate low risk encouragement email."""
    subject = f"Staying on track in {program}"
    
    body = f"""Hi {student_name},

Nice work so far in {program}. You're at {grade_pct}% with {attendance_pct}% attendance.

If you want to boost your results further, I can share study tips and campus resources.

Keep it up!

{advisor['name']}
{advisor['email']}"""
    
    return {'subject': subject, 'body': body}

