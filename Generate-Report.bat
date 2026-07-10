@echo off
rem =====================================================================
rem  Elite Prep - SAT Class Results Analysis Report Generator
rem  Usage: drag a folder of student score-report PDFs onto this file,
rem         or double-click to use a "PDFs" subfolder next to this file.
rem =====================================================================
setlocal
set "SCRIPT_DIR=%~dp0"

if "%~1"=="" (
    if exist "%SCRIPT_DIR%PDFs\" (
        python "%SCRIPT_DIR%generate_class_report.py" "%SCRIPT_DIR%PDFs"
    ) else (
        echo Drag a folder of student PDF score reports onto this file,
        echo or create a "PDFs" subfolder next to it and put the PDFs there.
    )
) else (
    python "%SCRIPT_DIR%generate_class_report.py" "%~1"
)

echo.
pause
