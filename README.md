# Student Risk Analyzer

> **Live Demo:** https://<your-github-username>.github.io/student-risk-analyzer/

> **📋 Excel File Format**: See [EXCEL_FORMAT_GUIDE.md](./EXCEL_FORMAT_GUIDE.md) for detailed instructions on how to structure your Excel file (single worksheet).

A web application for analyzing student risk levels based on grades and attendance data. The application uses both simple rules and machine learning models to predict at-risk students and provides actionable insights for academic advisors.

## Features

- **Excel File Upload**: Upload a single-sheet Excel file containing student identifiers, names, program, grade, and attendance
- **Dual Risk Assessment**:
  - Simple Rule: Flags students with grade < 70% OR attendance < 70%
  - Weighted Score: 80% grade + 20% attendance produces a continuous risk score used for dashboards
- **Rule-Based Categories**: Automatically labels students as No Risk, Medium Risk, High Risk, or Extremely High Risk based on grade/attendance combinations
- **Personalized Email Drafts**: Generates category-specific outreach templates ready to copy
- **Campus Login Integration**: Supports hyperlinks from Excel or configurable base URLs
- **Email Draft Generation**: Generates personalized email drafts based on risk category
- **Export Functionality**: Export results as CSV
- **Interactive UI**: Single-page web interface with search, filtering, and sorting

## Tech Stack

- **Backend**: Python FastAPI
- **Frontend**: HTML5, JavaScript (vanilla), Bootstrap 5
- **ML**: scikit-learn (LogisticRegression, GradientBoostingClassifier), XGBoost
- **Data Processing**: pandas, openpyxl
- **Explainability**: SHAP (optional)

## Project Structure

```
student-risk-analyzer/
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI application and endpoints
│   ├── models.py            # Pydantic models
│   ├── parsers.py           # Excel parsing and data normalization
│   ├── risk.py              # Risk scoring logic (simple rule + ML)
│   ├── email_templates.py   # Email draft generation
│   └── static/
│       ├── index.html       # Main HTML page
│       ├── app.js           # Frontend JavaScript
│       └── styles.css      # Custom styles
├── tests/
│   ├── __init__.py
│   ├── test_parsers.py     # Parser unit tests
│   └── test_risk.py         # Risk scoring unit tests
├── requirements.txt
├── Dockerfile
└── README.md
```

## Installation

### Prerequisites

- Python 3.11 or higher
- pip

### Local Setup

1. Clone the repository:
```bash
git clone <repository-url>
cd student-risk-analyzer
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file (optional, uses defaults if not present):
```bash
CAMPUS_LOGIN_BASE_URL=https://compuslogin.example.com?student_id={student_id}
RISK_THRESHOLDS=low:0,medium:60,high:80
ALLOW_ORIGINS=*
MAX_UPLOAD_SIZE_MB=10
ADVISOR_NAME=Academic Advisor
ADVISOR_EMAIL=advisor@example.com
```

5. Run the application:
```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

6. Open your browser and navigate to:
```
http://localhost:8000
```

## Docker Setup

1. Build the Docker image:
```bash
docker build -t student-risk-analyzer .
```

2. Run the container:
```bash
docker run -p 8000:8000 --env-file .env student-risk-analyzer
```

Or run without env file (uses defaults):
```bash
docker run -p 8000:8000 student-risk-analyzer
```

3. Access the application at:
```
http://localhost:8000
```

## Excel File Format

The application expects an Excel file with **one worksheet** that contains the following columns:

| Column | Description |
|--------|-------------|
| `Student#` | Unique student identifier |
| `Student Name`