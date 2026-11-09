@echo off
chcp 65001 > nul
title Shopee 官方仓全流程智能补货工作台

echo ========================================================
echo   Shopee 官方仓全流程智能补货工作台 启动中...
echo ========================================================

cd /d "%~dp0"

echo [*] 正在为您打开网页版智能工作台 (无需任何环境依赖，支持即开即用)...
start "" "index.html"

where python >nul 2>nul
if %errorlevel% equ 0 (
    echo [*] 检测到本地 Python 环境，正在同步启动桌面客户端...
    start "" python gui.py
) else if exist "D:\Program Files\Python313\python.exe" (
    echo [*] 正在启动桌面客户端...
    start "" "D:\Program Files\Python313\python.exe" gui.py
)

echo [√] 工作台已成功启动！
exit
