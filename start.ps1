# Feishu Approval Print Tool - One-Click Launcher
# Double-click this file or run: .\start.ps1

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

# --- 2. Locate Python ---
# Priority: 1) already-active venv  2) .venv/ dir  3) venv/ dir  4) system python
$PY = $null
$venvName = ""

if ($env:VIRTUAL_ENV) {
    # A virtual environment is already active - use it
    $venvName = Split-Path $env:VIRTUAL_ENV -Leaf
    $PY = Join-Path $env:VIRTUAL_ENV "Scripts\python.exe"
    Write-Host "[1/3] Using active virtual environment: $venvName" -ForegroundColor Green
}
elseif (Test-Path ".venv\Scripts\python.exe") {
    $venvName = ".venv"
    $PY = ".venv\Scripts\python.exe"
    Write-Host "[1/3] Using virtual environment: .venv" -ForegroundColor Green
}
elseif (Test-Path "venv\Scripts\python.exe") {
    $venvName = "venv"
    $PY = "venv\Scripts\python.exe"
    Write-Host "[1/3] Using virtual environment: venv" -ForegroundColor Green
}
else {
    $PY = "python"
    Write-Host "[1/3] No virtual environment found, using system Python" -ForegroundColor Yellow
    Write-Host "      (Create one with: uv venv  or  python -m venv .venv)" -ForegroundColor Yellow
}

# Quick sanity check
$savedErrorPref = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& $PY --version 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Cannot run Python: $PY" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
$ErrorActionPreference = $savedErrorPref

# --- 3. Install dependencies ---
Write-Host "[2/3] Checking dependencies..." -ForegroundColor Green

$ErrorActionPreference = "Continue"
& $PY -c "import streamlit, requests, dotenv, openpyxl" 2>&1 | Out-Null
$importFailed = ($LASTEXITCODE -ne 0)
$ErrorActionPreference = $savedErrorPref

if ($importFailed) {
    Write-Host "      Installing dependencies (Tsinghua mirror)..." -ForegroundColor Yellow
    $ErrorActionPreference = "Continue"
    & $PY -m pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
    $installFailed = ($LASTEXITCODE -ne 0)
    $ErrorActionPreference = $savedErrorPref

    if ($installFailed) {
        Write-Host "[ERROR] Dependency installation failed." -ForegroundColor Red
        Write-Host "        Try manually: $PY -m pip install -r requirements.txt" -ForegroundColor Yellow
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

& $PY -m streamlit run app/app.py --server.port 8501 --server.headless false

# --- 5. Exit ---
Write-Host ""
Write-Host "Server stopped." -ForegroundColor Yellow
Read-Host "Press Enter to exit"
