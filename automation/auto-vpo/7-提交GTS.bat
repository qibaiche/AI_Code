@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo GTS 自动提交工具
echo ========================================
echo.
echo 此工具将：
echo 1. 查找最新的 GTS_Submit_filled_*.xlsx 文件
echo 2. 打开 GTS 页面并自动填充 Title 和 Description
echo 3. 等待用户确认后自动提交
echo.

REM 运行 GTS 提交（通过 workflow_main）
python -c "from workflow_automation.config_loader import load_config; from workflow_automation.workflow_main import WorkflowController; from pathlib import Path; import logging; logging.basicConfig(level=logging.INFO, format='%%(asctime)s - %%(name)s - %%(levelname)s - %%(message)s'); config = load_config(Path('workflow_automation/config.yaml')); controller = WorkflowController(config); controller._step_submit_to_gts()"

echo.
echo 按任意键退出...
pause >nul

