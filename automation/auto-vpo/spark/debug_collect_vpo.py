"""调试脚本：只执行“等待30分钟后从 Spark Dashboard 收集 VPO 并写回 CSV”的步骤。

用法（在 automation/auto-vpo 目录下执行）：

    python spark\\debug_collect_vpo.py

逻辑：
1. 在 output 目录中找到最新的 MIR_Results_*.csv
2. 读取行数 N，作为期望 VPO 数量
3. 调用 SparkSubmitter.collect_recent_vpos_from_dashboard(N)
4. 将收集到的 VPO 反向匹配回 CSV，每行新增一列 VPO
5. 生成 MIR_Results_with_VPO_DEBUG_*.csv

注意：
- 等待时间由 config.yaml 中 spark.vpo_collect_wait_minutes 控制，默认 0 分钟（0 表示不等待）。
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# 将 auto-vpo 根目录加入 sys.path，方便导入 workflow_automation 模块
current_dir = Path(__file__).parent          # .../automation/auto-vpo/spark
parent_dir = current_dir.parent              # .../automation/auto-vpo
sys.path.insert(0, str(parent_dir))

from workflow_automation.config_loader import load_config  # noqa: E402
from workflow_automation.spark_submitter import SparkSubmitter  # noqa: E402


def main(no_wait: bool = False) -> None:
    print("=" * 80)
    print("🧪 Spark VPO 收集调试脚本")
    print("=" * 80)
    print()

    # 加载配置
    config_path = parent_dir / "workflow_automation" / "config.yaml"
    if not config_path.exists():
        print("❌ 错误：配置文件不存在！")
        print(f"   预期位置：{config_path}")
        input("\n按 Enter 键退出...")
        return

    config = load_config(config_path)

    # 查找最新 MIR_Results_*.csv
    output_dir = config.paths.output_dir
    print(f"📁 输出目录：{output_dir}")

    if not output_dir.exists():
        print("❌ 错误：output 目录不存在！")
        input("\n按 Enter 键退出...")
        return

    mir_files = sorted(output_dir.glob("MIR_Results_*.csv"), reverse=True)
    if not mir_files:
        print("❌ 错误：未在 output 目录找到 MIR_Results_*.csv 文件！")
        input("\n按 Enter 键退出...")
        return

    selected_file = mir_files[0]
    print(f"✅ 使用最新 MIR 文件：{selected_file.name}")

    df = pd.read_csv(selected_file)
    if df.empty:
        print("❌ 错误：MIR 结果文件为空！")
        input("\n按 Enter 键退出...")
        return

    expected_count = len(df)
    print(f"🔢 MIR 行数：{expected_count}")
    print()

    # 根据参数决定是否跳过等待
    if no_wait:
        wait_minutes = 0
        print("⏱  调试模式：不等待，立即从 Spark Dashboard 收集 VPO ...")
    else:
        wait_minutes = getattr(config.spark, "vpo_collect_wait_minutes", 0)
        print(f"⏱  将在 {wait_minutes} 分钟后从 Spark Dashboard 收集 VPO（0 表示不等待）...")
        print("    （如需等待一段时间，可以在 config.yaml 中设置 vpo_collect_wait_minutes）")
    print()

    submitter = SparkSubmitter(config.spark)

    try:
        # 将等待时间通过 config 传递给 submitter（collect_recent_vpos_from_dashboard 内部使用）
        submitter.config.vpo_collect_wait_minutes = wait_minutes
        vpo_list = submitter.collect_recent_vpos_from_dashboard(expected_count=expected_count)
    finally:
        # 确保浏览器关闭
        submitter._close_driver()

    if not vpo_list:
        print("⚠️ 未收集到任何 VPO，结束。")
        input("\n按 Enter 键退出...")
        return

    print(f"✅ 收集到 {len(vpo_list)} 个 VPO：{vpo_list}")

    # 反向匹配回原 CSV
    vpo_list_reversed = list(reversed(vpo_list))
    mir_with_vpo = df.copy()
    vpo_col_name = "VPO"

    if vpo_col_name in mir_with_vpo.columns:
        print(f"⚠️ 警告：原文件中已存在列 '{vpo_col_name}'，将覆盖该列的值。")

    mir_with_vpo[vpo_col_name] = ""

    max_count = min(len(mir_with_vpo), len(vpo_list_reversed))
    for i in range(max_count):
        mir_with_vpo.at[mir_with_vpo.index[i], vpo_col_name] = vpo_list_reversed[i]

    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_dir / f"MIR_Results_with_VPO_DEBUG_{date_str}.csv"
    mir_with_vpo.to_csv(output_file, index=False, encoding="utf-8-sig")

    print()
    print("✅ 已生成带 VPO 的调试文件：")
    print(f"   {output_file}")
    print(f"   共写入 {max_count} 条 VPO 记录（总行数：{len(mir_with_vpo)}）")
    print()
    input("按 Enter 键退出...")


if __name__ == "__main__":
    # 命令行可选参数：--no-wait  表示不等待，立即收集 VPO
    no_wait_flag = "--no-wait" in sys.argv
    main(no_wait=no_wait_flag)


