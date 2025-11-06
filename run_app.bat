@echo off
REM Run the Student Risk Analyzer application using the virtual environment
C:\Users\asbda\PycharmProjects\pythonProject\.venv\Scripts\python.exe -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

