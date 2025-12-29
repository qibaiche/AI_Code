@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo GTS 自动提交工具
echo ========================================
echo.
echo 此工具将：
echo 1. 在 03_GTS 文件夹中查找最新的 GTS_Submit_filled_*.xlsx 文件
echo 2. 打开 GTS 页面并自动填充 Title 和 Description
echo 3. 等待用户确认后自动提交
echo.

REM 运行 GTS 提交（明确指定从 03_GTS 文件夹查找文件）
python -c "from workflow_automation.config_loader import load_config; from workflow_automation.gts_submitter import GTSSubmitter, GTSConfig; from pathlib import Path; import logging; logging.basicConfig(level=logging.INFO, format='%%(asctime)s - %%(name)s - %%(levelname)s - %%(message)s'); base_dir = Path('.').resolve(); config = load_config(base_dir / 'workflow_automation/config.yaml'); gts_config = GTSConfig(**config.gts.__dict__); gts_config.output_dir = base_dir / 'output' / '03_GTS'; submitter = GTSSubmitter(gts_config); submitter.fill_ticket_with_latest_output()"

echo.
echo 按任意键退出...
pause >nul

