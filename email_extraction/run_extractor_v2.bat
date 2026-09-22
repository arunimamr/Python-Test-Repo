@echo off
REM Email Extractor v2 Launcher - Windows Batch
REM Usage: run_extractor_v2.bat [limit] [batch_size]

setlocal enabledelayedexpansion

echo.
echo ======================================================================
echo Email Contact Extractor v2
echo ======================================================================
echo.

REM Check Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8+
    exit /b 1
)

REM Get Python path
for /f "delims=" %%A in ('python -c "import sys; print(sys.executable)"') do (
    set PYTHON_PATH=%%A
)
echo Using Python: %PYTHON_PATH%
echo.

REM Build arguments
set ARGS=
if "%1" neq "" (
    set ARGS=!ARGS! --limit %1
    echo Limit: %1 emails
)
if "%2" neq "" (
    set ARGS=!ARGS! --batch-size %2
    echo Batch size: %2
) else (
    echo Batch size: 20 (default)
)
echo.

REM Run extractor
python email_extractor_v2.py %ARGS%

if errorlevel 1 (
    echo.
    echo Extraction completed with errors
    exit /b 1
) else (
    echo.
    echo Extraction completed successfully!
    echo Check 'output/' directory for results
    exit /b 0
)

