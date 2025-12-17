@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo    统一自动化工具
echo    (Lab TP + PRD LOT)
echo ========================================
echo.

python unified_automation.py

echo.
if errorlevel 1 (
    echo ❌ 运行失败，请查看上方错误信息
) else (
    echo ✅ 运行完成
)
echo.
pause

