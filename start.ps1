# Feishu Approval Print Tool - One-Click Launcher
# Double-click this file in project folder. No manual venv activation needed.

Set-Location $PSScriptRoot

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Feishu Approval Print Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. Check .env ---
if (-not (Test-Path ".env")) {
    Write-Host "[ERROR] .env file not found" -ForegroundColor Red
    Write-Host "Copy .env.example to .env and fill in credentials:" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_ID, FEISHU_APP_SECRET, FEISHU_APPROVAL_CODE" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# --- 2. Locate Python (uv 虚拟环境优先) ---
$uv = Get-Command uv -ErrorAction SilentlyContinue
if ($uv) {
    # 有 uv 就用 uv 维护 .venv，无论系统 Python 是否存在
    if (-not (Test-Path ".venv\Scripts\python.exe")) {
        Write-Host "[venv] Creating uv virtual environment..." -ForegroundColor Green
        & uv venv
        if ($LASTEXITCODE -ne 0) {
            Write-Host "[ERROR] uv venv 创建失败" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
    $_py = ".venv\Scripts\python.exe"
    Write-Host "[venv] Using uv virtual environment: .venv" -ForegroundColor Green
}
elseif (Test-Path ".venv\Scripts\python.exe") {
    $_py = ".venv\Scripts\python.exe"
    Write-Host "[venv] Found virtual environment: .venv" -ForegroundColor Green
}
elseif (Test-Path "venv\Scripts\python.exe") {
    $_py = "venv\Scripts\python.exe"
    Write-Host "[venv] Found virtual environment: venv" -ForegroundColor Green
}
else {
    # 没有 uv，回退系统 Python
    python --version 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $_py = "python"
        Write-Host "[python] Using system Python (no venv & uv not found)" -ForegroundColor Yellow
        Write-Host "         建议: 安装 uv 让脚本自动管理虚拟环境 (https://docs.astral.sh/uv/getting-started/installation/)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "[ERROR] Neither Python nor uv found on PATH" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# --- 3. Pick installer (uv pip first, then python -m pip) ---
if ($uv) {
    $_install = "uv pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple"
    Write-Host "[deps] Using uv pip (Tsinghua mirror)" -ForegroundColor Green
}
else {
    # uv not available, check if pip exists in the venv
    & $_py -m pip --version 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Neither uv nor pip is available" -ForegroundColor Red
        Write-Host "Install uv: https://docs.astral.sh/uv/getting-started/installation/" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }
    $_install = "$_py -m pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple"
    Write-Host "[deps] Using pip (Tsinghua mirror)" -ForegroundColor Green
}

# --- 4. Install dependencies (含版本校验，streamlit 1.33+ 才有 st.button width) ---
Write-Host "[deps] Checking..." -ForegroundColor Green

# 校验 streamlit >= 1.33 (炸 width= 参数 API). 其他依赖按 import 存在性即可.
& $_py -c "import streamlit as s; v = tuple(map(int, s.__version__.split('.')[:2])); assert v >= (1, 33), f'streamlit {s.__version__} < 1.33'; import requests, dotenv, openpyxl" 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[deps] Installing..." -ForegroundColor Yellow
    Invoke-Expression $_install
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Dependency installation failed" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}
Write-Host "[deps] Ready" -ForegroundColor Green

# --- 5. Launch Streamlit ---
Write-Host ""
Write-Host "Starting Streamlit..." -ForegroundColor Green
Write-Host "Browser: http://localhost:8501  |  Ctrl+C to stop" -ForegroundColor Cyan
Write-Host ""

& $_py -m streamlit run app/app.py --server.port 8501

Write-Host ""
Write-Host "Server stopped." -ForegroundColor Yellow
Read-Host "Press Enter to exit"
