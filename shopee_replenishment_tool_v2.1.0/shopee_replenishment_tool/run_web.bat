@echo off
chcp 65001 >nul
title Shopee 官方仓智能补货工具 - 本地服务
cd /d "%~dp0"

echo 正在检查运行环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 python 命令。
    echo        请先安装 Python 3.9+ 并勾选 "Add Python to PATH"。
    pause
    exit /b 1
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖 openpyxl ...
    python -m pip install openpyxl
)

echo 正在启动本地服务并打开浏览器（关闭本窗口即停止服务）...
python server.py
pause
