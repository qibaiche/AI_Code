@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo    调试 Correlation_UnitLevel 表
echo ========================================
echo.

python -m prd_lot_automation.main --config prd_lot_automation\config.yaml --skip-fetch

echo.
if errorlevel 1 (
    echo ❌ 运行失败，请查看上方错误信息
    pause
    exit /b 1
) else (
    echo ✅ 运行完成
    echo.
    echo 正在查找并打开最新的报告文件...
    
    REM 获取今天的日期
    for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
    set today=%datetime:~0,8%
    set date_str=%today:~0,4%-%today:~4,2%-%today:~6,2%
    
    REM 查找最新的报告文件
    set "report_file=reports\Bin_Pareto_Report_%date_str%.xlsx"
    
    if exist "%report_file%" (
        echo 找到报告文件: %report_file%
        echo 正在打开...
        start "" "%report_file%"
        echo ✅ 已打开报告文件
    ) else (
        echo ⚠️ 未找到今天的报告文件: %report_file%
        echo 尝试查找最新的报告文件...
        
        REM 查找reports目录下最新的Bin_Pareto_Report_*.xlsx文件
        for /f "delims=" %%F in ('dir /b /o-d "reports\Bin_Pareto_Report_*.xlsx" 2^>nul') do (
            set "latest_report=reports\%%F"
            goto :found
        )
        
        :found
        if defined latest_report (
            echo 找到最新的报告文件: %latest_report%
            echo 正在打开...
            start "" "%latest_report%"
            echo ✅ 已打开报告文件
        ) else (
            echo ❌ 未找到任何报告文件
        )
    )
)

echo.
pause

