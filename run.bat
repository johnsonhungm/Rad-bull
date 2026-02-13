@echo off
title RIS Full Workflow

REM Set your Hugging Face token and endpoint URL here
REM Deploy MedGemma 1.5 at https://endpoints.huggingface.co/ to get your endpoint URL
set HF_TOKEN=your-hf-token-here
set HF_ENDPOINT_URL=https://your-endpoint.endpoints.huggingface.cloud

REM Change to the directory where this batch file is located
cd /d "%~dp0"

echo ============================================================
echo RIS Full Workflow - Radiology Report Automation
echo ============================================================
echo Current folder: %cd%
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/
    pause
    exit /b 1
)

REM Run the Python script from the same directory
echo Starting script...
echo.
python "ris_full_workflow.py"

if errorlevel 1 (
    echo.
    echo ============================================================
    echo ERROR: Script exited with an error.
    echo ============================================================
)

echo.
echo ============================================================
echo Script completed. Press any key to close...
pause >nul
