# 飞书审批打印工具 — 一键启动脚本
# 双击此文件或在 PowerShell 中运行：.\start.ps1

$ErrorActionPreference = "Stop"

# 切换到脚本所在目录
Set-Location $PSScriptRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  飞书审批打印工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. 检查 .env 配置 ---
if (-not (Test-Path ".env")) {
    Write-Host "[错误] 未找到 .env 文件" -ForegroundColor Red
    Write-Host "请复制 .env.example 为 .env 并填入飞书应用凭证：" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_ID=your_app_id" -ForegroundColor Yellow
    Write-Host "  FEISHU_APP_SECRET=your_app_secret" -ForegroundColor Yellow
    Write-Host "  FEISHU_APPROVAL_CODE=your_approval_code" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "按 Enter 退出"
    exit 1
}

# --- 2. 激活虚拟环境 ---
if (Test-Path ".venv\Scripts\Activate.ps1") {
    Write-Host "[1/3] 激活虚拟环境..." -ForegroundColor Green
    . .venv\Scripts\Activate.ps1
} else {
    Write-Host "[1/3] 未找到虚拟环境，使用系统 Python..." -ForegroundColor Yellow
}

# --- 3. 检查并安装依赖 ---
Write-Host "[2/3] 检查依赖..." -ForegroundColor Green
$packages = python -c "import streamlit, requests, python_dotenv, openpyxl" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "      正在安装依赖（清华镜像）..." -ForegroundColor Yellow
    pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[错误] 依赖安装失败" -ForegroundColor Red
        Read-Host "按 Enter 退出"
        exit 1
    }
}
Write-Host "      依赖就绪" -ForegroundColor Green

# --- 4. 启动 Streamlit ---
Write-Host "[3/3] 启动 Streamlit..." -ForegroundColor Green
Write-Host ""
Write-Host "浏览器将自动打开，如未打开请访问: http://localhost:8501" -ForegroundColor Cyan
Write-Host "按 Ctrl+C 可停止服务" -ForegroundColor Cyan
Write-Host ""

streamlit run app/app.py --server.port 8501 --server.headless false

# --- 5. 退出 ---
Write-Host ""
Write-Host "服务已停止。" -ForegroundColor Yellow
Read-Host "按 Enter 退出"
