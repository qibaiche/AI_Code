@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 自动化工作流工具
echo ========================================
echo.

REM 切换到父目录，以便Python能找到workflow_automation模块
cd /d "%~dp0.."
python -m workflow_automation.main

echo.
echo 按任意键退出...
pause >nul

