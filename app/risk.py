"""Risk scoring logic: simple rules and ML models."""

import numpy as np
import pandas as pd
from typing import Tuple, Optional, Dict, List
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import GradientBoostingClassifier, IsolationForest
from sklearn.preprocessing import StandardScaler
from sklearn.calibration import CalibratedClassifierCV
from sklearn.model_selection import train_test_split
import warnings

warnings.filterwarnings('ignore')

try:
    import shap
    SHAP_AVAILABLE = True
except ImportError:
    SHAP_AVAILABLE = False


def simple_rule(grade_pct: float, attendance_pct: float) -> bool:
    """
    Simple rule: flagged if Grade < 70 OR Attendance < 70.
    
    Args:
        grade_pct: Grade percentage (0-100)
        attendance_pct: Attendance percentage (0-100)
    
    Returns:
        True if at risk by simple rule
    """
    return grade_pct < 70.0 or attendance_pct < 70.0


def get_risk_category(risk_score: float, thresholds: Dict[str, float]) -> str:
    """
    Categorize risk score into Low/Medium/High/Extremely High Risk.
    
    Args:
        risk_score: Risk score (0-100)
        thresholds: Dict with 'low', 'medium', 'high', 'failed' threshold values
    
    Returns:
        Risk category string (Low, Medium, High, or Extremely High Risk)
    """
    # Check thresholds in descending order
    if risk_score >= thresholds.get('failed', 90):
        return 'Extremely High Risk'
    elif risk_score >= thresholds.get('high', 80):
        return 'High'
    elif risk_score >= thresholds.get('medium', 60):
        return 'Medium'
    else:
        return 'Low'


# Risk category color mapping
RISK_COLOR_MAP = {
    "Low Risk": "#00B050",            # Green
    "Medium Risk": "#FFC000",         # Orange
    "High Risk": "#FF0000",           # Red
    "Extremely High Risk": "#7030A0"  # Dark purple / maroon
}


def get_risk_color(risk_category: str) -> str:
    """
    Get color code for a risk category.
    
    Args:
        risk_category: Risk category string
    
    Returns:
        Hex color code string
    """
    return RISK_COLOR_MAP.get(risk_category, "#6B7280")  # Default gray


def classify_risk_by_grade_attendance(grade_pct: float, attendance_pct: float) -> str:
    """
    Classify risk category directly from grade and attendance percentages.
    Uses explicit numeric conversion and direct threshold comparisons.
    
    This ensures that low grade alone (even if attendance is high) still triggers High Risk.
    
    Args:
        grade_pct: Grade percentage (0-100, numeric or string like "24.0%")
        attendance_pct: Attendance percentage (0-100, numeric or string like "88.9%")
    
    Returns:
        Risk category string (Low Risk, Medium Risk, High Risk, or Extremely High Risk)
    """
    # Ensure explicit numeric conversion (handle string inputs like "24.0%")
    if isinstance(grade_pct, str):
        grade_pct = float(str(grade_pct).replace("%", "").strip())
    else:
        grade_pct = float(grade_pct)
    
    if isinstance(attendance_pct, str):
        attendance_pct = float(str(attendance_pct).replace("%", "").strip())
    else:
        attendance_pct = float(attendance_pct)
    
    # Clamp to valid range
    grade = max(0.0, min(100.0, grade_pct))
    attendance = max(0.0, min(100.0, attendance_pct))
    
    # Classification logic: low grade alone (even if attendance is high) triggers High Risk
    if grade < 70 and attendance < 70:
        return "Extremely High Risk"
    elif grade < 80 or attendance < 80:
        return "High Risk"
    elif grade < 90 or attendance < 90:
        return "Medium Risk"
    else:
        return "Low Risk"


def train_or_fallback_score(
    df: pd.DataFrame,
    thresholds: Dict[str, float],
    has_target: bool = False
) -> Tuple[np.ndarray, List[str], Optional[object], Optional[object]]:
    """
    Train ML model or use fallback heuristic to compute risk scores.
    
    Args:
        df: Merged dataframe with student data
        thresholds: Risk category thresholds
        has_target: Whether dataframe has 'is_at_risk' column
    
    Returns:
        Tuple of (risk_scores_0_100, categories, model, scaler)
    """
    # Prepare features
    feature_cols = []
    
    if 'grade_pct' in df.columns:
        feature_cols.append('grade_pct')
    if 'attendance_pct' in df.columns:
        feature_cols.append('attendance_pct')
    if 'missed_pct' in df.columns:
        feature_cols.append('missed_pct')
    if 'Missed Hours to Date_hours' in df.columns:
        feature_cols.append('Missed Hours to Date_hours')
    
    if not feature_cols:
        # Fallback: use simple weighted score
        return _fallback_heuristic_score(df, thresholds)
    
    X = df[feature_cols].fillna(0).values
    
    # Standardize features
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    # Check if we have target labels
    if has_target and 'is_at_risk' in df.columns:
        y = df['is_at_risk'].fillna(False).astype(int).values
        
        if len(np.unique(y)) > 1:  # Need both classes
            # Train supervised model
            try:
                # Try XGBoost first (better for nonlinearity)
                try:
                    from xgboost import XGBClassifier
                    model = XGBClassifier(random_state=42, n_estimators=100, max_depth=3)
                except ImportError:
                    # Fallback to GradientBoosting
                    model = GradientBoostingClassifier(random_state=42, n_estimators=100, max_depth=3)
                
                # Split data for calibration
                trained_model = model
                if len(df) > 10:
                    try:
                        X_train, X_cal, y_train, y_cal = train_test_split(
                            X_scaled, y, test_size=0.2, random_state=42, stratify=y
                        )
                        
                        # Train model
                        model.fit(X_train, y_train)
                        
                        # Calibrate probabilities
                        calibrated_model = CalibratedClassifierCV(model, method='isotonic', cv=min(3, len(X_cal)))
                        calibrated_model.fit(X_cal, y_cal)
                        
                        # Predict probabilities
                        proba = calibrated_model.predict_proba(X_scaled)[:, 1]
                        trained_model = calibrated_model
                    except Exception:
                        # If calibration fails, train without calibration
                        model.fit(X_scaled, y)
                        proba = model.predict_proba(X_scaled)[:, 1]
                        trained_model = model
                else:
                    # Too few samples, train without calibration
                    model.fit(X_scaled, y)
                    proba = model.predict_proba(X_scaled)[:, 1]
                    trained_model = model
                
                # Convert probability to risk score (0-100)
                risk_scores = proba * 100.0
                
                categories = [get_risk_category(score, thresholds) for score in risk_scores]
                
                return risk_scores, categories, trained_model, scaler
                
            except Exception as e:
                # If training fails, fall back to heuristic
                print(f"Model training failed: {e}, using fallback heuristic")
                return _fallback_heuristic_score(df, thresholds)
    
    # No labels or training failed: use fallback heuristic
    return _fallback_heuristic_score(df, thresholds)


def audit_and_recalculate_risk(df: pd.DataFrame) -> pd.DataFrame:
    """
    Audit and correct risk scoring for attendance-based model (presence rate).
    Ensures exponential weighting, consistent risk logic, and category accuracy.
    
    This function validates that:
    - Attendance % represents presence (higher = better)
    - Risk Score increases as grades or attendance drop
    - Exponential weighting emphasizes low performance more sharply
    - Categories align with the corrected attendance interpretation
    
    Args:
        df: DataFrame with 'grade_pct' and 'attendance_pct' columns
    
    Returns:
        DataFrame with validated and corrected risk calculations
    """
    # Create a copy to avoid modifying original
    df_audit = df.copy()
    
    # --- Step 1. Clean and convert percentage columns with explicit numeric conversion ---
    if 'grade_pct' in df_audit.columns:
        df_audit['Grade %'] = (
            df_audit['grade_pct']
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.strip()
            .astype(float)
            .clip(0, 100)
        )
    else:
        df_audit['Grade %'] = 0.0
    
    if 'attendance_pct' in df_audit.columns:
        df_audit['Attendance %'] = (
            df_audit['attendance_pct']
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.strip()
            .astype(float)
            .clip(0, 100)
        )
    else:
        df_audit['Attendance %'] = 0.0
    
    # --- Step 2. Normalize 0–1 scale ---
    g = df_audit['Grade %'] / 100.0
    a = df_audit['Attendance %'] / 100.0
    
    # --- Step 3. Exponential weighting ---
    # Higher = better; penalize low scores more sharply
    performance_index = 0.6 * (g ** 1.2) + 0.4 * (a ** 1.2)
    df_audit['Weighted Index P'] = performance_index.round(2)
    
    # --- Step 4. Compute Risk Score (higher = worse) ---
    df_audit['Risk Score'] = (100.0 * (1.0 - performance_index)).round(1)
    df_audit['Risk Score'] = df_audit['Risk Score'].clip(0.0, 100.0)
    
    # --- Step 5. Add audit summary columns ---
    df_audit['Audit Notes'] = np.where(
        (df_audit['Risk Score'] > 100) | (df_audit['Risk Score'] < 0),
        "⚠️ Out-of-Range Score",
        "OK"
    )
    
    # Validate attendance represents presence (not absence)
    # If attendance % > 50, it's likely presence; if < 50, might be absence
    # This is a sanity check - we assume attendance is already presence-based
    high_attendance_mask = df_audit['Attendance %'] > 50.0
    low_attendance_mask = df_audit['Attendance %'] < 50.0
    
    # Add validation note if attendance seems inverted
    if low_attendance_mask.sum() > high_attendance_mask.sum() * 2:
        print("⚠️ WARNING: More students have <50% attendance than >50%. Verify attendance represents presence (not absence).")
    
    return df_audit


def _fallback_heuristic_score(
    df: pd.DataFrame,
    thresholds: Dict[str, float]
) -> Tuple[np.ndarray, List[str], Optional[object], Optional[object]]:
    """
    Fallback heuristic risk scoring using exponential weighting.
    
    Uses exponential weighting formula that penalizes very low grades or attendance more strongly:
    - Normalize grades and attendance to 0-1
    - Apply exponential weighting: performance_index = 0.6 * (g^1.2) + 0.4 * (a^1.2)
    - Risk Score = 100 * (1 - performance_index)
    
    This ensures:
    - Attendance % represents presence (higher = better)
    - Risk Score increases as grades or attendance drop
    - Exponential weighting emphasizes low performance more sharply
    """
    grade_pct = df.get('grade_pct', pd.Series([0.0] * len(df))).fillna(0.0)
    attendance_pct = df.get('attendance_pct', pd.Series([0.0] * len(df))).fillna(0.0)
    
    # Validate: Ensure attendance represents presence (higher = better)
    # If values seem inverted (most students have <50%), log a warning
    if len(attendance_pct) > 0:
        high_attendance_count = (attendance_pct > 50.0).sum()
        low_attendance_count = (attendance_pct < 50.0).sum()
        if low_attendance_count > high_attendance_count * 2:
            print("⚠️ WARNING: Attendance values may be inverted. Expected attendance to represent presence (higher = better).")
    
    # Normalize to 0-1 (clamp values to valid range)
    # Ensure values are in 0-100 range (attendance should be presence %)
    g = np.clip(grade_pct, 0.0, 100.0) / 100.0
    a = np.clip(attendance_pct, 0.0, 100.0) / 100.0
    
    # Exponential weighting — more penalty for lower scores
    # Using power of 1.2 to create exponential curve
    # Formula: performance_index = 0.6 * (g^1.2) + 0.4 * (a^1.2)
    # This means:
    # - Higher grades/attendance → higher performance_index → lower risk
    # - Lower grades/attendance → lower performance_index → higher risk
    performance_index = 0.6 * (g ** 1.2) + 0.4 * (a ** 1.2)
    
    # Risk Score (higher = worse), rounded to 1 decimal place
    # Formula: Risk Score = 100 * (1 - performance_index)
    # This ensures risk increases as performance decreases
    risk_scores = np.round(100.0 * (1.0 - performance_index), 1)
    
    # Clip to 0-100 range (safety check)
    risk_scores = np.clip(risk_scores, 0.0, 100.0)
    
    # Optional: refine with IsolationForest for outliers
    try:
        if len(df) > 5:
            features = np.column_stack([
                grade_pct.fillna(0).values,
                attendance_pct.fillna(0).values
            ])
            
            iso_forest = IsolationForest(contamination=0.1, random_state=42)
            outlier_scores = iso_forest.fit_predict(features)
            
            # Boost risk for outliers
            outlier_mask = outlier_scores == -1
            risk_scores[outlier_mask] = np.minimum(risk_scores[outlier_mask] + 10, 100.0)
    except Exception:
        pass  # If IsolationForest fails, continue with base scores
    
    categories = [get_risk_category(score, thresholds) for score in risk_scores]
    
    return risk_scores.values, categories, None, None


def get_explanation(
    student_idx: int,
    df: pd.DataFrame,
    model: Optional[object],
    scaler: Optional[object],
    feature_cols: List[str]
) -> Optional[str]:
    """
    Generate explanation for why a student is at risk.
    
    Uses SHAP if available, otherwise model coefficients/feature importances.
    """
    if model is None or scaler is None:
        # Use simple feature-based explanation
        # Check bounds before accessing
        if student_idx < 0 or student_idx >= len(df):
            return None
        
        try:
            row = df.iloc[student_idx]
        except (IndexError, KeyError):
            return None
        
        reasons = []
        if 'grade_pct' in df.columns and row.get('grade_pct', 100) < 70:
            reasons.append(f"Low grade ({row.get('grade_pct', 0):.1f}%)")
        if 'attendance_pct' in df.columns and row.get('attendance_pct', 100) < 70:
            reasons.append(f"Low attendance ({row.get('attendance_pct', 0):.1f}%)")
        
        if reasons:
            return " | ".join(reasons)
        return None
    
    # Try SHAP explanation
    if SHAP_AVAILABLE:
        try:
            # Check bounds before accessing
            if student_idx < 0 or student_idx >= len(df):
                return None
            
            X = df[feature_cols].fillna(0).values
            X_scaled = scaler.transform(X)
            
            # Check if X_scaled has enough rows
            if student_idx >= len(X_scaled):
                return None
            
            # Create SHAP explainer
            if hasattr(model, 'predict_proba'):
                explainer = shap.TreeExplainer(model.base_estimator if hasattr(model, 'base_estimator') else model)
                shap_values = explainer.shap_values(X_scaled[student_idx:student_idx+1])
                
                if isinstance(shap_values, list):
                    shap_values = shap_values[1]  # Use positive class
                
                # Get top 2 contributing features
                feature_importance = np.abs(shap_values[0])
                top_indices = np.argsort(feature_importance)[-2:][::-1]
                
                explanations = []
                for idx in top_indices:
                    feature_name = feature_cols[idx]
                    contribution = shap_values[0][idx]
                    
                    if contribution > 0:
                        explanations.append(f"{feature_name} increases risk")
                    else:
                        explanations.append(f"{feature_name} decreases risk")
                
                return " | ".join(explanations)
        except Exception:
            pass  # Fall back to coefficients
    
    # Use model coefficients/feature importances
    try:
        if hasattr(model, 'feature_importances_'):
            importances = model.feature_importances_
        elif hasattr(model, 'base_estimator') and hasattr(model.base_estimator, 'feature_importances_'):
            importances = model.base_estimator.feature_importances_
        elif hasattr(model, 'coef_'):
            importances = np.abs(model.coef_[0])
        else:
            return None
        
        top_indices = np.argsort(importances)[-2:][::-1]
        top_features = [feature_cols[idx] for idx in top_indices]
        
        return f"Top factors: {', '.join(top_features)}"
    except Exception:
        return None

