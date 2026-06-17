# Feishu Approval Print Tool - One-Click Launcher
# Manage your own Python environment. This script just checks deps and runs.

Set-Location $PSScriptRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Feishu Approval Print Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. Check .env ---
if (-not (Test-Path ".env")) {
    Write-Host "[ERROR] .env file not found" -ForegroundColor Red
    Write-Host "Please copy .env.example to .env and fill in:" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_ID, FEISHU_APP_SECRET, FEISHU_APPROVAL_CODE" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# --- 2. Check Python ---
python --version 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] python is not available on PATH" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# --- 3. Install dependencies if missing ---
Write-Host "[1/2] Checking dependencies..." -ForegroundColor Green

python -c "import streamlit, requests, dotenv, openpyxl" 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "      Installing..." -ForegroundColor Yellow
    python -m pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] pip install failed" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}
Write-Host "      Ready" -ForegroundColor Green

# --- 4. Launch ---
Write-Host "[2/2] Starting Streamlit..." -ForegroundColor Green
Write-Host ""
Write-Host "Browser: http://localhost:8501  |  Ctrl+C to stop" -ForegroundColor Cyan
Write-Host ""

python -m streamlit run app/app.py --server.port 8501

Write-Host ""
Write-Host "Server stopped." -ForegroundColor Yellow
Read-Host "Press Enter to exit"
