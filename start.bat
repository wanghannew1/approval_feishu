@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo ========================================
echo   飞书审批打印工具
echo ========================================
echo.

REM 检查 .env 配置
if not exist ".env" (
    echo [错误] 未找到 .env 文件
    echo 请复制 .env.example 为 .env 并填入飞书应用凭证
    pause
    exit /b 1
)

REM 激活虚拟环境
if exist ".venv\Scripts\activate.bat" (
    echo [1/3] 激活虚拟环境...
    call .venv\Scripts\activate.bat
) else (
    echo [1/3] 未找到虚拟环境，使用系统 Python...
)

REM 检查依赖
echo [2/3] 检查依赖...
python -c "import streamlit, requests, dotenv, openpyxl" 2>nul
if errorlevel 1 (
    echo       正在安装依赖（清华镜像）...
    pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        echo [错误] 依赖安装失败
        pause
        exit /b 1
    )
)
echo       依赖就绪

REM 启动 Streamlit
echo [3/3] 启动 Streamlit...
echo.
echo 浏览器将自动打开，如未打开请访问: http://localhost:8501
echo 按 Ctrl+C 可停止服务
echo.

streamlit run app/app.py --server.port 8501 --server.headless false

echo.
echo 服务已停止。
pause
