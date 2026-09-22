@echo off
REM Turbify Email Extractor - Windows Batch Script
REM This script activates the virtual environment and runs the email extractor

setlocal enabledelayedexpansion

echo.
echo ======================================================
echo TURBIFY EMAIL CONTACT EXTRACTOR
echo ======================================================
echo.

REM Check if .env file exists
if not exist ".env" (
    echo ERROR: .env file not found!
    echo.
    echo Please create a .env file with your Turbify credentials:
    echo.
    echo   EMAIL=your-email@abluae.com
    echo   PASSWORD=your-password
    echo   IMAP_SERVER=abluae.com
    echo   IMAP_PORT=993
    echo   BATCH_SIZE=50
    echo.
    pause
    exit /b 1
)

REM Check if password is set
for /f "tokens=2 delims==" %%a in ('findstr "^PASSWORD=" .env') do set PASSWORD=%%a

if "!PASSWORD!"=="" (
    echo ERROR: PASSWORD not set in .env file!
    echo.
    echo Please edit .env and add your password:
    echo   PASSWORD=your-actual-password
    echo.
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    echo Virtual environment created.
)

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat

REM Check if dependencies are installed
pip show python-dotenv >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

echo.
echo ======================================================
echo Running Email Extractor
echo ======================================================
echo.

REM Run the extractor
python General\extractContacts_Turbify.py

REM Check result
if errorlevel 1 (
    echo.
    echo ERROR: Extraction failed. Check the log above.
    echo.
    pause
) else (
    echo.
    echo ======================================================
    echo SUCCESS! Check the output files:
    echo   - contacts_*.xlsx (extracted contacts)
    echo   - logos\ (downloaded company logos)
    echo   - email_extraction.log (detailed log)
    echo ======================================================
    echo.
    pause
)

endlocal

