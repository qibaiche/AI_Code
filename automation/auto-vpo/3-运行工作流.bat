@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 自动化工作流工具
echo ========================================
echo.

REM 保持在当前目录（automation/auto-vpo/），以便Python能找到workflow_automation模块
python -m workflow_automation.main

echo.
echo 按任意键退出...
pause >nul

