@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 测试 Mole 配置 UI
echo ========================================
echo.
echo 此脚本仅测试配置UI界面
echo 不会实际执行Mole自动化
echo.

python -c "from workflow_automation.mole_config_ui import show_mole_config_ui; from pathlib import Path; config_path = Path('workflow_automation/config.yaml'); result = show_mole_config_ui(config_path); print('\n配置结果:'); print(result) if result else print('用户取消了配置')"

echo.
echo 按任意键退出...
pause >nul

