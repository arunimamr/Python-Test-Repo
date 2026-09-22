# Turbify Email Extractor - PowerShell Script
# This script activates the virtual environment and runs the email extractor

Write-Host ""
Write-Host "======================================================"
Write-Host "TURBIFY EMAIL CONTACT EXTRACTOR"
Write-Host "======================================================"
Write-Host ""

# Check if .env file exists
if (-not (Test-Path ".env")) {
    Write-Host "ERROR: .env file not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please create a .env file with your Turbify credentials:"
    Write-Host ""
    Write-Host "  EMAIL=your-email@abluae.com"
    Write-Host "  PASSWORD=your-password"
    Write-Host "  IMAP_SERVER=abluae.com"
    Write-Host "  IMAP_PORT=993"
    Write-Host "  BATCH_SIZE=50"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Read .env file to check password
$envContent = Get-Content ".env" | Select-String "^PASSWORD=" | Select-Object -First 1
if ($null -eq $envContent) {
    Write-Host "ERROR: PASSWORD not set in .env file!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please edit .env and add your password:"
    Write-Host "  PASSWORD=your-actual-password"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if virtual environment exists
if (-not (Test-Path ".venv\Scripts\Activate.ps1")) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to create virtual environment" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "Virtual environment created." -ForegroundColor Green
}

# Activate virtual environment
Write-Host "Activating virtual environment..." -ForegroundColor Yellow
& .\.venv\Scripts\Activate.ps1

# Check if dependencies are installed
$depCheck = python -c "import python_dotenv, openpyxl, bs4" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing dependencies..." -ForegroundColor Yellow
    pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to install dependencies" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host ""
Write-Host "======================================================"
Write-Host "Running Email Extractor"
Write-Host "======================================================"
Write-Host ""

# Run the extractor
python General\extractContacts_Turbify.py

# Check result
if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "======================================================"
    Write-Host "SUCCESS! Check the output files:" -ForegroundColor Green
    Write-Host "  - contacts_*.xlsx (extracted contacts)"
    Write-Host "  - logos\ (downloaded company logos)"
    Write-Host "  - email_extraction.log (detailed log)"
    Write-Host "======================================================"
} else {
    Write-Host ""
    Write-Host "ERROR: Extraction failed. Check the log above." -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"

