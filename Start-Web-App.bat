@echo off
rem =====================================================================
rem  Elite Prep - SAT Class Test Result Analysis (web app)
rem  Port 8502: another app (TutoringReminder) already uses 8501.
rem  Double-click this file. It starts the app and opens your browser.
rem  Close this window (or press Ctrl+C) to stop the app.
rem =====================================================================
cd /d "%~dp0"
echo Starting the SAT report app ... (keep this window open)
start "" http://localhost:8502
python -m streamlit run app.py --server.port 8502 --browser.gatherUsageStats false
pause
