@echo off
cd /d "%~dp0"
echo ========================================
echo Spark 自动化工具 - 调试模式
echo ========================================
echo.
echo 此模式将从"添加 Lot"步骤开始执行
echo 请确保浏览器中已完成前面的步骤：
echo   - 已点击 Add New
echo   - 已填写 TP 路径并点击 Add New Experiment
echo   - 已选择 VPO 类别并填写实验信息
echo.
echo ========================================
echo.
python spark\test_spark.py --from-lot
pause

