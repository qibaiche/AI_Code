@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 自动化工作流工具
echo ========================================
echo.

python -m workflow_automation.main --mole-only

echo.
echo 按任意键退出...
pause >nul

