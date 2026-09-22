# Email Extractor v2 Launcher - PowerShell
# Usage: ./run_extractor_v2.ps1 [--limit 100] [--batch-size 10]

param(
    [int]$limit = 0,
    [int]$batch_size = 20
)

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "Email Contact Extractor v2" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""

# Check Python installation
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "ERROR: Python not found. Please install Python 3.8+" -ForegroundColor Red
    exit 1
}

Write-Host "Using Python: $($python.Source)" -ForegroundColor Green

# Check .env file
$envFile = Join-Path (Get-Location) ".env"
if (-not (Test-Path $envFile)) {
    Write-Host "WARNING: .env file not found at $envFile" -ForegroundColor Yellow
    Write-Host "Please create .env with: EMAIL, PASSWORD, IMAP_SERVER, IMAP_PORT" -ForegroundColor Yellow
}

# Build command arguments
$args_list = @()
if ($limit -gt 0) {
    $args_list += "--limit", $limit
    Write-Host "Limit: $limit emails" -ForegroundColor Cyan
}
if ($batch_size -gt 0) {
    $args_list += "--batch-size", $batch_size
    Write-Host "Batch size: $batch_size" -ForegroundColor Cyan
}

Write-Host ""

# Run extractor
try {
    & python email_extractor_v2.py @args_list
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "Extraction completed successfully!" -ForegroundColor Green
        Write-Host "Check 'output/' directory for results" -ForegroundColor Green
    }
    else {
        Write-Host ""
        Write-Host "Extraction completed with errors (exit code: $LASTEXITCODE)" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "ERROR: Failed to run extractor: $_" -ForegroundColor Red
    exit 1
}

