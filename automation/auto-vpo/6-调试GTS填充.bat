@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo GTS 填充调试 - 使用最新 MIR/VPO CSV
echo ========================================
echo.

REM 运行填充脚本：读取 output 下最新 MIR_Results*.csv
REM 使用 input\GTS_Submit.xlsx 模板，生成 output\GTS_Submit_filled_*.xlsx
python -m workflow_automation.gts_excel_filler

echo.
echo 完成后请在 output 目录查看生成的 GTS_Submit_filled_*.xlsx
echo.
echo 按任意键退出...
pause >nul


