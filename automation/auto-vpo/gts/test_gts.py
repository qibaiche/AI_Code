"""GTS 自动化开发/测试入口

目前只是加载配置并初始化 GTSSubmitter，后续可以在这里逐步开发、调试
具体的 GTS 网页 / 应用操作。
"""

import sys
from pathlib import Path

# 将 auto-vpo 根目录加入 sys.path，方便导入 workflow_automation 模块
current_dir = Path(__file__).parent
parent_dir = current_dir.parent  # automation/auto-vpo/
sys.path.insert(0, str(parent_dir))

from workflow_automation.config_loader import load_config  # noqa: E402
from workflow_automation.gts_submitter import GTSSubmitter  # noqa: E402


def main() -> None:
    print("=" * 80)
    print("🚀 GTS 自动化开发 / 测试")
    print("=" * 80)
    print()

    # 加载主配置
    config_path = parent_dir / "workflow_automation" / "config.yaml"
    if not config_path.exists():
        print("❌ 错误：配置文件不存在！")
        print(f"   预期位置：{config_path}")
        print()
        print("💡 解决方法：")
        print("   1. 确认文件路径是否正确")
        print("   2. 查看 workflow_automation/README.md 了解配置方法")
        input("\n按 Enter 键退出...")
        return

    config = load_config(config_path)

    print("📋 当前 GTS 配置：")
    print(f"   URL: {getattr(config.gts, 'url', '(未配置)')}")
    print(f"   超时: {getattr(config.gts, 'timeout', '(默认)')} 秒")
    print()

    submitter = GTSSubmitter(config.gts)

    try:
        print("🚀 打开页面并填充最新输出（Title + Description TSV）...")
        submitter.fill_ticket_with_latest_output()
        print("✅ 已填充完毕，请在浏览器中确认后手动提交/保存。")
        print()
    except Exception as e:
        print()
        print("=" * 80)
        print("❌ 初始化或填充过程中发生错误")
        print("=" * 80)
        print(f"错误信息: {e}")
        print()
        import traceback
        traceback.print_exc()
    finally:
        print()
        print("=" * 80)
        print("提示：若页面元素选择器不同，可在 config.yaml 的 gts.title_selector / description_selector 调整。")
        print("输出数据来源：output/GTS_Submit_filled_*.xlsx，取最新一份。")
        print("=" * 80)
        input("按 Enter 键退出...")


if __name__ == "__main__":
    main()


