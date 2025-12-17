@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo Auto GTS 开发 / 测试入口
echo ========================================
echo.

python gts\test_gts.py

echo.
echo 按任意鍵退出...
pause >nul


