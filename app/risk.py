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
    Categorize risk score into Low/Medium/High.
    
    Args:
        risk_score: Risk score (0-100)
        thresholds: Dict with 'low', 'medium', 'high' threshold values
    
    Returns:
        Risk category string
    """
    if risk_score >= thresholds.get('high', 80):
        return 'High'
    elif risk_score >= thresholds.get('medium', 60):
        return 'Medium'
    else:
        return 'Low'


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


def _fallback_heuristic_score(
    df: pd.DataFrame,
    thresholds: Dict[str, float]
) -> Tuple[np.ndarray, List[str], Optional[object], Optional[object]]:
    """
    Fallback heuristic risk scoring.
    
    Uses weighted combination: risk = 0.6*(100-Grade%) + 0.4*(100-Att%)
    """
    grade_pct = df.get('grade_pct', pd.Series([0.0] * len(df))).fillna(0.0)
    attendance_pct = df.get('attendance_pct', pd.Series([0.0] * len(df))).fillna(0.0)
    
    # Weighted risk score
    risk_scores = 0.6 * (100.0 - grade_pct) + 0.4 * (100.0 - attendance_pct)
    
    # Clip to 0-100 range
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

