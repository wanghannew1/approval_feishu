# Feishu Approval Print Tool - One-Click Launcher
# Double-click this file or run: .\start.ps1

$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Feishu Approval Print Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. Check .env config ---
if (-not (Test-Path ".env")) {
    Write-Host "[ERROR] .env file not found" -ForegroundColor Red
    Write-Host "Please copy .env.example to .env and fill in credentials:" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_ID=your_app_id" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_SECRET=your_app_secret" -ForegroundColor Yellow
    Write-Host "  FEISHU_APPROVAL_CODE=your_approval_code" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# --- 2. Activate virtual environment ---
if (Test-Path ".venv\Scripts\Activate.ps1") {
    Write-Host "[1/3] Activating virtual environment..." -ForegroundColor Green
    . .venv\Scripts\Activate.ps1
}
else {
    Write-Host "[1/3] No virtual environment found, using system Python..." -ForegroundColor Yellow
}

# --- 3. Check and install dependencies ---
Write-Host "[2/3] Checking dependencies..." -ForegroundColor Green
$check = python -c "import streamlit, requests, dotenv, openpyxl" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "      Installing dependencies (Tsinghua mirror)..." -ForegroundColor Yellow
    pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Dependency installation failed" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}
Write-Host "      Dependencies ready" -ForegroundColor Green

# --- 4. Launch Streamlit ---
Write-Host "[3/3] Starting Streamlit..." -ForegroundColor Green
Write-Host ""
Write-Host "Browser will open automatically. If not, visit: http://localhost:8501" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop" -ForegroundColor Cyan
Write-Host ""

streamlit run app/app.py --server.port 8501 --server.headless false

# --- 5. Exit ---
Write-Host ""
Write-Host "Server stopped." -ForegroundColor Yellow
Read-Host "Press Enter to exit"
