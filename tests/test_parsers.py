"""Unit tests for parsers module."""

import pytest
import pandas as pd
import numpy as np
from io import BytesIO

from app.parsers import (
    parse_duration,
    normalize_percentage,
    normalize_data,
    merge_data
)


def test_parse_duration():
    """Test duration parsing."""
    assert parse_duration("90:00") == 90.0
    assert parse_duration("0:15") == 0.25
    assert parse_duration("1:30") == 1.5
    assert parse_duration("0:00") == 0.0
    assert parse_duration("") == 0.0
    assert parse_duration(None) == 0.0
    assert parse_duration(pd.NA) == 0.0


def test_normalize_percentage():
    """Test percentage normalization."""
    # Values in 0-1 range should be multiplied by 100
    assert normalize_percentage(0.5, max_value=1.0) == 50.0
    assert normalize_percentage(0.7, max_value=1.0) == 70.0
    assert normalize_percentage(1.0, max_value=1.0) == 100.0
    
    # Values already in 0-100 range should stay the same
    assert normalize_percentage(50.0, max_value=100.0) == 50.0
    assert normalize_percentage(70.0, max_value=100.0) == 70.0
    
    # Handle NaN
    assert normalize_percentage(pd.NA, max_value=1.0) == 0.0


def test_normalize_data():
    """Test data normalization."""
    # Create sample dataframes
    grades_df = pd.DataFrame({
        'Student#': ['001', '002', '003'],
        'Student Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
        'Program Name': ['Program A', 'Program B', 'Program A'],
        'current overall Program Grade': [0.85, 0.65, 0.95]
    })
    
    attendance_df = pd.DataFrame({
        'Student#': ['001', '002', '003'],
        'Student Name': ['Doe, John', 'Smith, Jane', 'Johnson, Bob'],
        'Scheduled Hours to Date': ['100:00', '100:00', '100:00'],
        'Attended Hours to Date': ['90:00', '70:00', '95:00'],
        'Attended % to Date.': [0.90, 0.70, 0.95],
        'Missed Hours to Date': ['10:00', '30:00', '5:00'],
        '% Missed': [0.10, 0.30, 0.05],
        'Missed Minus Excused to date': ['5:00', '20:00', '2:00']
    })
    
    grades_norm, attendance_norm = normalize_data(grades_df, attendance_df)
    
    # Check grade normalization
    assert 'grade_pct' in grades_norm.columns
    assert grades_norm['grade_pct'].iloc[0] == 85.0
    assert grades_norm['grade_pct'].iloc[1] == 65.0
    assert grades_norm['grade_pct'].iloc[2] == 95.0
    
    # Check attendance normalization
    assert 'attendance_pct' in attendance_norm.columns
    assert attendance_norm['attendance_pct'].iloc[0] == 90.0
    assert attendance_norm['attendance_pct'].iloc[1] == 70.0
    assert attendance_norm['attendance_pct'].iloc[2] == 95.0
    
    # Check duration parsing
    assert 'Scheduled Hours to Date_hours' in attendance_norm.columns
    assert attendance_norm['Scheduled Hours to Date_hours'].iloc[0] == 100.0


def test_merge_data():
    """Test data merging."""
    # Create normalized dataframes
    grades_df = pd.DataFrame({
        'Student#': ['001', '002', '003'],
        'Student Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
        'Program Name': ['Program A', 'Program B', 'Program A'],
        'grade_pct': [85.0, 65.0, 95.0]
    })
    
    attendance_df = pd.DataFrame({
        'Student#': ['001', '002', '003'],
        'Student Name': ['Doe, John', 'Smith, Jane', 'Johnson, Bob'],
        'attendance_pct': [90.0, 70.0, 95.0],
        'missed_pct': [10.0, 30.0, 5.0]
    })
    
    merged = merge_data(grades_df, attendance_df)
    
    # Check merge
    assert len(merged) == 3
    assert 'Student#' in merged.columns
    assert 'Student Name' in merged.columns
    assert 'Program Name' in merged.columns
    assert 'grade_pct' in merged.columns
    assert 'attendance_pct' in merged.columns
    
    # Check that Grades sheet name is preferred
    assert merged['Student Name'].iloc[0] == 'John Doe'
    assert merged['Program Name'].iloc[0] == 'Program A'


def test_merge_data_deduplication():
    """Test that duplicates are handled correctly."""
    grades_df = pd.DataFrame({
        'Student#': ['001', '001', '002'],
        'Student Name': ['John Doe', 'John Doe', 'Jane Smith'],
        'Program Name': ['Program A', 'Program A', 'Program B'],
        'grade_pct': [85.0, 90.0, 65.0]
    })
    
    attendance_df = pd.DataFrame({
        'Student#': ['001', '002'],
        'Student Name': ['Doe, John', 'Smith, Jane'],
        'attendance_pct': [90.0, 70.0]
    })
    
    merged = merge_data(grades_df, attendance_df)
    
    # Should deduplicate and keep last row
    assert len(merged) == 2
    assert '001' in merged['Student#'].values
    assert '002' in merged['Student#'].values
    
    # Check that last row for 001 is kept (grade_pct = 90.0)
    student_001 = merged[merged['Student#'] == '001']
    assert student_001['grade_pct'].iloc[0] == 90.0

