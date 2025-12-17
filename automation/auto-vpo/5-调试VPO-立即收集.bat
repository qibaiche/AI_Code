@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo Spark VPO 调试 - 立即收集（不等待）
echo ========================================
echo.

REM 提示：如需真正等待30分钟，请直接运行：python spark\debug_collect_vpo.py
REM 本批处理仅用于快速调试，通过 --no-wait 参数跳过等待。

python spark\debug_collect_vpo.py --no-wait

echo.
echo 按任意鍵退出...
pause >nul


