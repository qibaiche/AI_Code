#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置验证脚本 - 检查 PRD LOT 自动化工具的所有配置
"""

import sys
from pathlib import Path

def check_item(name, condition, message=""):
    """检查单项配置"""
    status = "✅" if condition else "❌"
    print(f"{status} {name}")
    if message:
        print(f"   {message}")
    return condition

def main():
    print("=" * 60)
    print("PRD LOT 自动化工具 - 配置验证")
    print("=" * 60)
    print()
    
    base_dir = Path(__file__).parent
    all_ok = True
    
    # 1. 检查 Python 环境
    print("📦 [1/5] Python 环境检查")
    print("-" * 60)
    
    import platform
    print(f"   Python 版本: {platform.python_version()}")
    
    try:
        import pywinauto
        check_item("pywinauto", True, f"版本: {pywinauto.__version__}")
    except ImportError:
        all_ok &= check_item("pywinauto", False, "未安装，请运行: pip install pywinauto")
    
    try:
        import pyautogui
        check_item("pyautogui", True, f"版本: {pyautogui.__version__}")
    except ImportError:
        all_ok &= check_item("pyautogui", False, "未安装，请运行: pip install pyautogui")
    
    try:
        import pandas as pd
        check_item("pandas", True, f"版本: {pd.__version__}")
    except ImportError:
        all_ok &= check_item("pandas", False, "未安装，请运行: pip install pandas")
    
    try:
        import openpyxl
        check_item("openpyxl", True, f"版本: {openpyxl.__version__}")
    except ImportError:
        all_ok &= check_item("openpyxl", False, "未安装，请运行: pip install openpyxl")
    
    try:
        import yaml
        check_item("pyyaml", True, f"版本: {yaml.__version__}")
    except ImportError:
        all_ok &= check_item("pyyaml", False, "未安装，请运行: pip install pyyaml")
    
    try:
        import win32com.client
        check_item("pywin32", True)
    except ImportError:
        all_ok &= check_item("pywin32", False, "未安装，请运行: pip install pywin32")
    
    print()
    
    # 2. 检查文件路径
    print("📁 [2/5] 文件路径检查")
    print("-" * 60)
    
    lots_file = base_dir / "Lot info.txt"
    all_ok &= check_item(
        "Lot info.txt",
        lots_file.exists(),
        f"路径: {lots_file}"
    )
    
    vg2_file = base_dir / "Get_Sort_or_Test_Unit_Results_HBASE_By_Lot.VG2"
    all_ok &= check_item(
        "VG2 文件",
        vg2_file.exists(),
        f"路径: {vg2_file}"
    )
    
    config_file = base_dir / "prd_lot_automation" / "config.yaml"
    all_ok &= check_item(
        "config.yaml",
        config_file.exists(),
        f"路径: {config_file}"
    )
    
    requirements_file = base_dir / "prd_lot_automation" / "requirements.txt"
    all_ok &= check_item(
        "requirements.txt",
        requirements_file.exists(),
        f"路径: {requirements_file}"
    )
    
    print()
    
    # 3. 检查 SQLPathFinder
    print("🚀 [3/5] SQLPathFinder 检查")
    print("-" * 60)
    
    spf_path = Path("C:/Program Files/SQLPathFinder/SQLPathFinder.exe")
    spf_exists = spf_path.exists()
    if spf_exists:
        check_item("SQLPathFinder.exe", True, f"路径: {spf_path}")
    else:
        check_item(
            "SQLPathFinder.exe (可选)",
            True,
            f"⚠️ 未在默认路径找到，将通过 VG2 文件关联自动打开"
        )
    
    output_csv = Path("C:/Users/qibaiche/AppData/Local/Temp/SQLPathFinder_Temp")
    check_item(
        "CSV 输出目录可写",
        output_csv.parent.exists(),
        f"路径: {output_csv.parent}"
    )
    
    print()
    
    # 4. 检查 assets 资源
    print("🖼️ [4/5] 资源文件检查")
    print("-" * 60)
    
    assets_dir = base_dir / "assets"
    check_item("assets 目录", assets_dir.exists())
    
    run_button_img = assets_dir / "run_button.png"
    if run_button_img.exists():
        check_item("run_button.png", True, f"路径: {run_button_img}")
    else:
        print("⚠️  run_button.png (待创建)")
        print(f"   路径: {run_button_img}")
        print("   说明: 首次运行前需要截取 SQLPathFinder 的 Run 按钮")
        print("   指导: 参考 assets/如何截取Run按钮.md")
        print("   备选: 可在 config.yaml 中配置 run_button_automation_id")
    
    readme = assets_dir / "如何截取Run按钮.md"
    check_item("截图指导文档", readme.exists())
    
    print()
    
    # 5. 检查配置内容
    print("⚙️ [5/5] 配置内容检查")
    print("-" * 60)
    
    try:
        import yaml
        with open(config_file, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        check_item("config.yaml 格式", True)
        
        email_to = config.get('email', {}).get('to', [])
        if email_to and email_to[0] not in ["someone@example.com", ""]:
            check_item("邮箱配置", True, f"收件人: {', '.join(email_to)}")
        else:
            print("⚠️  邮箱配置")
            print("   建议: 在 config.yaml 中配置实际的收件人邮箱")
        
        email_mode = config.get('email', {}).get('mode', 'outlook')
        check_item("邮件模式", True, f"当前: {email_mode}")
        
    except Exception as e:
        all_ok &= check_item("config.yaml", False, f"读取失败: {e}")
    
    print()
    
    # 6. 检查目录结构
    print("📂 [补充] 输出目录检查")
    print("-" * 60)
    
    reports_dir = base_dir / "reports"
    all_ok &= check_item("reports 目录", reports_dir.exists(), f"路径: {reports_dir}")
    
    logs_dir = base_dir / "logs"
    all_ok &= check_item("logs 目录", logs_dir.exists(), f"路径: {logs_dir}")
    
    print()
    print("=" * 60)
    
    if all_ok:
        print("✅ 核心配置检查通过！")
        print()
        if not run_button_img.exists():
            print("📝 运行前提示：")
            print("   首次运行前，请先截取 Run 按钮图片")
            print("   详见: assets/如何截取Run按钮.md")
            print()
        print("运行方法：")
        print("  1. 双击 运行工具.bat")
        print("  2. 或运行: python -m prd_lot_automation.main")
    else:
        print("❌ 部分配置存在问题，请根据上述提示修复后再运行。")
        print()
        print("常见修复方法：")
        print("  1. 安装依赖: pip install -r prd_lot_automation\\requirements.txt")
        print("  2. 截取 Run 按钮: 参考 assets/如何截取Run按钮.md")
        print("  3. 配置邮箱: 编辑 prd_lot_automation/config.yaml")
    
    print("=" * 60)
    return 0 if all_ok else 1

if __name__ == "__main__":
    sys.exit(main())
