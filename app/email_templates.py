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
    """Generate an email draft tailored to the student's risk category."""
    advisor = get_advisor_info()
    grade_str = f"{grade_pct:.1f}"
    attendance_str = f"{attendance_pct:.1f}"

    category = risk_category.lower()
    if category == "no risk":
        return _no_risk_email(student_name, program, grade_str, attendance_str, advisor)
    if category == "medium risk":
        return _medium_risk_email(student_name, program, grade_str, attendance_str, advisor)
    if category == "high risk":
        return _high_risk_email(student_name, program, grade_str, attendance_str, advisor)
    return _extremely_high_risk_email(student_name, program, grade_str, attendance_str, advisor)


def _no_risk_email(student_name: str, program: str, grade_pct: str, attendance_pct: str, advisor: Dict[str, str]) -> Dict[str, str]:
    subject = f"Great Work, {student_name} — Keep It Up!"
    body = f"""Hi {student_name},

Excellent work so far in {program}! You’re maintaining a strong academic record with a grade of {grade_pct}% and {attendance_pct}% attendance.

Keep up the consistency — your progress shows commitment and discipline. If you’d like, I can share advanced study tips or ways to stay engaged with campus enrichment programs.

Great job!

{advisor['name']}
{advisor['email']}"""
    return {'subject': subject, 'body': body}


def _medium_risk_email(student_name: str, program: str, grade_pct: str, attendance_pct: str, advisor: Dict[str, str]) -> Dict[str, str]:
    subject = f"Let’s Talk About Attendance, {student_name}"
    body = f"""Hi {student_name},

You’re doing very well in {program} with a grade of {grade_pct}%, but your attendance is currently at {attendance_pct}%.

Regular participation helps maintain your performance and ensures you don’t miss key content. Please reach out to your instructor or the Student Success Office if something is affecting your attendance.

We want to help you stay on track for success.

{advisor['name']}
{advisor['email']}"""
    return {'subject': subject, 'body': body}


def _high_risk_email(student_name: str, program: str, grade_pct: str, attendance_pct: str, advisor: Dict[str, str]) -> Dict[str, str]:
    subject = f"Support to Improve Your Academic Performance, {student_name}"
    body = f"""Hi {student_name},

Thanks for keeping strong attendance at {attendance_pct}% in {program} — that shows real dedication. However, your current grade is {grade_pct}%, which suggests you may need some extra academic support.

I recommend meeting with your instructor or the Student Success team to review study techniques, tutoring options, and learning resources available to you.

You’re already putting in the effort — let’s work together to lift your results.

{advisor['name']}
{advisor['email']}"""
    return {'subject': subject, 'body': body}


def _extremely_high_risk_email(student_name: str, program: str, grade_pct: str, attendance_pct: str, advisor: Dict[str, str]) -> Dict[str, str]:
    subject = f"Let’s Work Together to Get You Back on Track, {student_name}"
    body = f"""Hi {student_name},

I’m reaching out about your current progress in {program}. Your grade is {grade_pct}% and attendance is {attendance_pct}%.

These indicators suggest you might be struggling with both coursework and attendance.

Please contact the Student Success Office or your instructor as soon as possible to discuss a recovery plan. We can help you explore tutoring, time management strategies, and support services designed to help you succeed.

You’re not alone in this — we’re here to help you get back on track.

{advisor['name']}
{advisor['email']}"""
    return {'subject': subject, 'body': body}

