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

# --- 2. Set up Python / pip paths ---
# Use venv Python when available, otherwise fall back to system Python
$venvPython = ".venv\Scripts\python.exe"
$venvPip    = ".venv\Scripts\pip.exe"
$venvStreamlit = ".venv\Scripts\streamlit.exe"

if ((Test-Path $venvPython) -and (Test-Path $venvPip)) {
    $PY = $venvPython
    $PIP = $venvPip
    Write-Host "[1/3] Using virtual environment Python" -ForegroundColor Green
} else {
    $PY = "python"
    $PIP = "pip"
    Write-Host "[1/3] Virtual environment not found, using system Python" -ForegroundColor Yellow
}

# --- 3. Ensure pip itself is available ---
$savedErrorPref = $ErrorActionPreference
$ErrorActionPreference = "Continue"

& $PY -m pip --version 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] pip is not available. Please install pip first." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# --- 4. Check and install dependencies ---
Write-Host "[2/3] Checking dependencies..." -ForegroundColor Green

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
        Write-Host "[ERROR] Dependency installation failed. Try manually:" -ForegroundColor Red
        Write-Host "        $PIP install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }
}
Write-Host "      Dependencies ready" -ForegroundColor Green

# --- 5. Launch Streamlit ---
Write-Host "[3/3] Starting Streamlit..." -ForegroundColor Green
Write-Host ""
Write-Host "Browser will open automatically. If not, visit: http://localhost:8501" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $venvStreamlit) {
    & $venvStreamlit run app/app.py --server.port 8501 --server.headless false
} else {
    & $PY -m streamlit run app/app.py --server.port 8501 --server.headless false
}

# --- 6. Exit ---
Write-Host ""
Write-Host "Server stopped." -ForegroundColor Yellow
Read-Host "Press Enter to exit"
