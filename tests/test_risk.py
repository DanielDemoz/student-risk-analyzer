"""Unit tests for risk scoring module."""

import pytest
import pandas as pd
import numpy as np

from app.risk import (
    simple_rule,
    get_risk_category,
    train_or_fallback_score
)


def test_simple_rule():
    """Test simple rule logic."""
    # Grade < 70 should flag
    assert simple_rule(65.0, 80.0) == True
    
    # Attendance < 70 should flag
    assert simple_rule(80.0, 65.0) == True
    
    # Both < 70 should flag
    assert simple_rule(65.0, 65.0) == True
    
    # Both >= 70 should not flag
    assert simple_rule(75.0, 75.0) == False
    
    # Edge case: exactly 70
    assert simple_rule(70.0, 70.0) == False
    assert simple_rule(69.9, 70.0) == True
    assert simple_rule(70.0, 69.9) == True


def test_get_risk_category():
    """Test risk category assignment."""
    thresholds = {'low': 0, 'medium': 60, 'high': 80}
    
    # High risk
    assert get_risk_category(85.0, thresholds) == 'High'
    assert get_risk_category(80.0, thresholds) == 'High'
    
    # Medium risk
    assert get_risk_category(70.0, thresholds) == 'Medium'
    assert get_risk_category(60.0, thresholds) == 'Medium'
    assert get_risk_category(79.9, thresholds) == 'Medium'
    
    # Low risk
    assert get_risk_category(50.0, thresholds) == 'Low'
    assert get_risk_category(59.9, thresholds) == 'Low'
    assert get_risk_category(0.0, thresholds) == 'Low'


def test_train_or_fallback_score_no_labels():
    """Test fallback heuristic scoring when no labels are present."""
    df = pd.DataFrame({
        'Student#': ['001', '002', '003', '004', '005'],
        'grade_pct': [95.0, 85.0, 75.0, 65.0, 55.0],
        'attendance_pct': [95.0, 85.0, 75.0, 65.0, 55.0],
        'missed_pct': [5.0, 15.0, 25.0, 35.0, 45.0]
    })
    
    thresholds = {'low': 0, 'medium': 60, 'high': 80}
    
    risk_scores, categories, model, scaler = train_or_fallback_score(
        df, thresholds, has_target=False
    )
    
    # Check that scores are returned
    assert len(risk_scores) == len(df)
    assert len(categories) == len(df)
    
    # Check score bounds (0-100)
    assert all(0 <= score <= 100 for score in risk_scores)
    
    # Check categories
    assert all(cat in ['Low', 'Medium', 'High'] for cat in categories)
    
    # Lower grades/attendance should have higher risk scores
    assert risk_scores[0] < risk_scores[4]  # Student 001 (high grade) < Student 005 (low grade)


def test_train_or_fallback_score_with_labels():
    """Test supervised model training when labels are present."""
    df = pd.DataFrame({
        'Student#': ['001', '002', '003', '004', '005', '006', '007', '008', '009', '010'],
        'grade_pct': [95.0, 85.0, 75.0, 65.0, 55.0, 90.0, 80.0, 70.0, 60.0, 50.0],
        'attendance_pct': [95.0, 85.0, 75.0, 65.0, 55.0, 90.0, 80.0, 70.0, 60.0, 50.0],
        'is_at_risk': [0, 0, 0, 1, 1, 0, 0, 1, 1, 1]
    })
    
    thresholds = {'low': 0, 'medium': 60, 'high': 80}
    
    risk_scores, categories, model, scaler = train_or_fallback_score(
        df, thresholds, has_target=True
    )
    
    # Check that scores are returned
    assert len(risk_scores) == len(df)
    assert len(categories) == len(df)
    
    # Check score bounds (0-100)
    assert all(0 <= score <= 100 for score in risk_scores)
    
    # Check categories
    assert all(cat in ['Low', 'Medium', 'High'] for cat in categories)
    
    # Students labeled as at-risk should generally have higher scores
    at_risk_indices = df[df['is_at_risk'] == 1].index
    not_at_risk_indices = df[df['is_at_risk'] == 0].index
    
    avg_risk_at_risk = np.mean([risk_scores[i] for i in at_risk_indices])
    avg_risk_not_at_risk = np.mean([risk_scores[i] for i in not_at_risk_indices])
    
    # At-risk students should have higher average risk scores
    assert avg_risk_at_risk >= avg_risk_not_at_risk


def test_risk_score_bounds():
    """Test that risk scores are always within 0-100 bounds."""
    # Test with various data configurations
    test_cases = [
        pd.DataFrame({
            'grade_pct': [100.0, 0.0, 50.0],
            'attendance_pct': [100.0, 0.0, 50.0]
        }),
        pd.DataFrame({
            'grade_pct': [95.0, 85.0, 75.0, 65.0, 55.0],
            'attendance_pct': [95.0, 85.0, 75.0, 65.0, 55.0],
            'missed_pct': [5.0, 15.0, 25.0, 35.0, 45.0]
        })
    ]
    
    thresholds = {'low': 0, 'medium': 60, 'high': 80}
    
    for df in test_cases:
        risk_scores, categories, model, scaler = train_or_fallback_score(
            df, thresholds, has_target=False
        )
        
        # All scores should be in [0, 100]
        assert all(0 <= score <= 100 for score in risk_scores), \
            f"Scores out of bounds: {risk_scores}"

