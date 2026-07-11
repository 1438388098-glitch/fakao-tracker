@echo off
chcp 65001 >nul
title 法考打卡

echo.
echo   正在启动法考打卡...
echo.

cd /d "%~dp0"

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [错误] 未检测到 Python，请先安装 Python 3.7+
    echo   下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Install dependencies
echo   [1/2] 安装依赖...
pip install -r requirements.txt -q

:: Start server
echo   [2/2] 启动服务器...
echo.
echo   ================================
echo   浏览器打开 http://127.0.0.1:5000
echo   按 Ctrl+C 停止服务器
echo   ================================
echo.

start "" http://127.0.0.1:5000
python app.py
pause
