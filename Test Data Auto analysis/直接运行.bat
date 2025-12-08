@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo    PRD LOT 自动化工具 - 直接运行
echo ========================================
echo.

python -m prd_lot_automation.main --config prd_lot_automation\config.yaml

echo.
if errorlevel 1 (
    echo ❌ 运行失败，请查看上方错误信息
) else (
    echo ✅ 运行完成
)
echo.
pause

