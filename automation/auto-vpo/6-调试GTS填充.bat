@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo GTS 填充调试 - 使用 SPARK 文件夹最新文件
echo ========================================
echo.

REM 运行填充脚本：读取 02_SPARK 文件夹下最新 MIR_Results_For_Spark_*.xlsx
REM 使用 input\GTS_Submit.xlsx 模板，生成到 03_GTS\GTS_Submit_filled_*.xlsx
python -m workflow_automation.gts_excel_filler

echo.
echo 完成后请在 output\03_GTS 目录查看生成的 GTS_Submit_filled_*.xlsx
echo.
echo 按任意键退出...
pause >nul


