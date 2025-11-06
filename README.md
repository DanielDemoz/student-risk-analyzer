# Student Risk Analyzer

> **📋 Excel File Format**: See [EXCEL_FORMAT_GUIDE.md](./EXCEL_FORMAT_GUIDE.md) for detailed instructions on how to structure your Excel file.

A web application for analyzing student risk levels based on grades and attendance data. The application uses both simple rules and machine learning models to predict at-risk students and provides actionable insights for academic advisors.

## Features

- **Excel File Upload**: Upload Excel files with student grades and attendance data
- **Dual Risk Assessment**:
  - Simple Rule: Flags students with grade < 70% OR attendance < 70%
  - Advanced ML Model: Uses logistic regression or gradient boosting to predict risk scores
- **Risk Categorization**: Automatically categorizes students as Low/Medium/High risk
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

The application expects an Excel file with **two sheets**:

### Sheet 1: "Students Grade"

Required columns:
- `Student#` (string or number)
- `Student Name` (string; may include hyperlink)
- `Program Name` (string)
- `current overall Program Grade` (0-1 decimal or 0-100; will be normalized to 0-100)

### Sheet 2: "Students attendance " (note trailing space)

Required columns:
- `Student#`
- `Student Name` (format may be "Last, First")
- `Scheduled Hours to Date` (e.g., "90:00")
- `Attended Hours to Date` (e.g., "80:00")
- `Attended % to Date.` (0-1 decimal; will be normalized to 0-100)
- `Missed Hours to Date` (e.g., "5:00")
- `% Missed` (0-1 decimal; will be normalized to 0-100)
- `Missed Minus Excused to date` (duration string like "0:15")

**Note**: Column names are matched case-insensitively and with flexible whitespace handling.

**Merging**: Data is merged on `Student#`. If both sheets have student names, the Grades sheet name is preferred.

**Hyperlinks**: If the "Student Name" cell in the Grades sheet has a hyperlink, it will be used for the Campus Login action.

## API Endpoints

### `GET /`
Serves the main HTML page.

### `POST /upload`
Upload and process an Excel file.

**Request**: multipart/form-data
- `file`: Excel file (.xlsx or .xls)
- `campus_login_base_url` (optional): Base URL for campus login

**Response**: JSON with processed results
```json
{
  "success": true,
  "message": "Successfully processed 50 students",
  "results": [...],
  "summary": {
    "High": 10,
    "Medium": 15,
    "Low": 25,
    "Total": 50
  }
}
```

### `GET /results`
Get the last processed results (in-memory cache).

### `POST /email-draft`
Generate an email draft for a student.

**Request**: JSON
```json
{
  "student_id": "001",
  "risk_category": "High",
  "program": "Computer Science",
  "grade_pct": 65.0,
  "attendance_pct": 70.0
}
```

**Response**: JSON
```json
{
  "subject": "Action Needed: Let's get you back on track in Computer Science",
  "body": "Hi John Doe,\n\n..."
}
```

### `GET /download.csv`
Download processed results as CSV.

## Risk Scoring Logic

### Simple Rule
A student is flagged if:
- Grade < 70% OR
- Attendance < 70%

### Advanced ML Model

**Supervised Mode** (if `is_at_risk` column present):
- Trains a logistic regression or gradient boosting classifier
- Uses calibrated probabilities for risk scores
- Falls back to heuristic if training fails

**Unsupervised Mode** (no labels):
- Uses weighted heuristic: `risk = 0.6*(100-Grade%) + 0.4*(100-Att%)`
- Optionally refines with IsolationForest for outliers
- Scores are rescaled to 0-100

### Risk Categories

- **High**: Risk Score ≥ 80
- **Medium**: Risk Score 60-79
- **Low**: Risk Score < 60

Thresholds are configurable via `RISK_THRESHOLDS` environment variable.

## Testing

Run unit tests:
```bash
pytest tests/
```

Run with coverage:
```bash
pytest tests/ --cov=app --cov-report=html
```

## Configuration

Environment variables (`.env` file):

- `CAMPUS_LOGIN_BASE_URL`: Base URL template for campus login (default: `https://compuslogin.example.com?student_id={student_id}`)
- `RISK_THRESHOLDS`: Risk category thresholds (default: `low:0,medium:60,high:80`)
- `ALLOW_ORIGINS`: CORS allowed origins (default: `*`)
- `MAX_UPLOAD_SIZE_MB`: Maximum upload size in MB (default: `10`)
- `ADVISOR_NAME`: Advisor name for email templates (default: `Academic Advisor`)
- `ADVISOR_EMAIL`: Advisor email for email templates (default: `advisor@example.com`)

## Usage

1. **Upload File**: Click "Choose File" and select your Excel file
2. **Optional**: Enter a custom Campus Login Base URL
3. **Process**: Click "Process File" button
4. **View Results**: Results table shows all students with risk scores and categories
5. **Search**: Use the search box to filter by name or program
6. **Actions**:
   - **Open Campus Login**: Opens the student's campus login page
   - **Email Draft**: Opens a modal with a pre-filled email draft
7. **Export**: Click "Export CSV" to download results

## Email Templates

The application generates three types of email drafts:

- **High Risk**: Warning email with action items
- **Medium Risk**: Check-in email with support options
- **Low Risk**: Encouragement email with resources

All emails include placeholders for advisor name and contact information.

## Limitations

- **In-Memory Storage**: Results are stored in memory and lost on server restart
- **Single Session**: Only the most recent upload is cached
- **No Authentication**: Application is open to all users
- **File Size**: Limited by `MAX_UPLOAD_SIZE_MB` (default 10MB)

## Troubleshooting

### "Could not find 'Students Grade' sheet"
- Ensure your Excel file has a sheet named exactly "Students Grade"
- Check that sheet names match exactly (case-sensitive)

### "Missing required columns"
- Verify all required columns are present in both sheets
- Column names are matched case-insensitively but must contain the expected text

### "File too large"
- Reduce file size or increase `MAX_UPLOAD_SIZE_MB` in `.env`

### SHAP not available
- SHAP is optional; the application will use model coefficients/feature importances instead
- To install SHAP: `pip install shap`

## License

This project is provided as-is for educational and internal use.

## Contributing

Contributions are welcome! Please ensure:
- Code follows PEP 8 style guidelines
- Tests are added for new features
- Documentation is updated

## Support

For issues or questions, please open an issue in the repository.

