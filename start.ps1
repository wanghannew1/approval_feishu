# Feishu Approval Print Tool - One-Click Launcher
# Double-click this file in project folder. No manual venv activation needed.

Set-Location $PSScriptRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Feishu Approval Print Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. Check .env ---
if (-not (Test-Path ".env")) {
    Write-Host "[ERROR] .env file not found" -ForegroundColor Red
    Write-Host "Please copy .env.example to .env and fill in credentials:" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_ID=your_app_id" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_SECRET=your_app_secret" -ForegroundColor Yellow
    Write-Host "  FEISHU_APPROVAL_CODE=your_approval_code" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# --- 2. Locate Python (uv .venv preferred) ---
if (Test-Path ".venv\Scripts\python.exe") {
    $_py = ".venv\Scripts\python.exe"
    Write-Host "[venv] Found uv virtual environment: .venv" -ForegroundColor Green
}
elseif (Test-Path "venv\Scripts\python.exe") {
    $_py = "venv\Scripts\python.exe"
    Write-Host "[venv] Found virtual environment: venv" -ForegroundColor Green
}
else {
    # No venv found, try system python
    python --version 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $_py = "python"
        Write-Host "[python] Using system Python (no venv found)" -ForegroundColor Yellow
    } else {
        # Try to create venv with uv
        $uv = Get-Command uv -ErrorAction SilentlyContinue
        if ($uv) {
            Write-Host "[venv] Creating uv virtual environment..." -ForegroundColor Green
            uv venv
        } else {
            Write-Host "[ERROR] Neither Python nor uv found on PATH" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
        if (Test-Path ".venv\Scripts\python.exe") {
            $_py = ".venv\Scripts\python.exe"
            Write-Host "[venv] Created: .venv" -ForegroundColor Green
        } else {
            Write-Host "[ERROR] Failed to create virtual environment" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
}

# --- 3. Install dependencies ---
Write-Host "[deps] Checking..." -ForegroundColor Green

& $_py -c "import streamlit, requests, dotenv, openpyxl" 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[deps] Installing (Tsinghua mirror)..." -ForegroundColor Yellow
    & $_py -m pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Dependency installation failed" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}
Write-Host "[deps] Ready" -ForegroundColor Green

# --- 4. Launch Streamlit ---
Write-Host ""
Write-Host "Starting Streamlit..." -ForegroundColor Green
Write-Host "Browser: http://localhost:8501  |  Ctrl+C to stop" -ForegroundColor Cyan
Write-Host ""

& $_py -m streamlit run app/app.py --server.port 8501

Write-Host ""
Write-Host "Server stopped." -ForegroundColor Yellow
Read-Host "Press Enter to exit"
