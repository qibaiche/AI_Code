@echo off
cd /d "%~dp0"
echo Testing Spark Web Page...
python spark\test_spark.py
pause

