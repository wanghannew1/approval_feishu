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

REM --- 定位 Python (uv 虚拟环境优先) ---
where uv >nul 2>nul
if %errorlevel% equ 0 (
    if not exist ".venv\Scripts\python.exe" (
        echo [venv] Creating uv virtual environment...
        call uv venv
        if errorlevel 1 (
            echo [错误] uv venv 创建失败
            pause
            exit /b 1
        )
    )
    set "_PY=.venv\Scripts\python.exe"
    set "_INSTALL=uv pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple"
    echo [venv] Using uv virtual environment: .venv
    goto :check_deps
)

if exist ".venv\Scripts\activate.bat" (
    echo [1/3] 激活虚拟环境...
    call .venv\Scripts\activate.bat
    set "_PY=python"
    set "_INSTALL=python -m pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple"
    goto :check_deps
)

if exist "venv\Scripts\activate.bat" (
    echo [1/3] 激活虚拟环境...
    call venv\Scripts\activate.bat
    set "_PY=python"
    set "_INSTALL=python -m pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple"
    goto :check_deps
)

echo [1/3] 未找到 uv 也无虚拟环境，使用系统 Python...
python --version >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 Python 或 uv，请安装其一
    pause
    exit /b 1
)
set "_PY=python"
set "_INSTALL=python -m pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple"

:check_deps
REM 检查依赖 (含 streamlit 版本: 1.33+ 才有 st.button width 参数)
echo [2/3] 检查依赖...
%_PY% -c "import streamlit as s; v=tuple(map(int,s.__version__.split('.')[:2])); assert v>=(1,33), f'streamlit {s.__version__}<1.33'; import requests, dotenv, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo       正在安装依赖（清华镜像）...
    call %_INSTALL%
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

%_PY% -m streamlit run app/app.py --server.port 8501 --server.headless false

echo.
echo 服务已停止。
pause