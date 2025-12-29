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

    # 查找最新 MIR_Results_*.csv（支持新的目录结构）
    output_dir = config.paths.output_dir
    print(f"📁 输出目录：{output_dir}")

    if not output_dir.exists():
        print("❌ 错误：output 目录不存在！")
        input("\n按 Enter 键退出...")
        return

    mir_files = []
    
    # 新结构：直接在 01_MIR/ 目录中查找
    mir_results_dir = output_dir / "01_MIR"
    if mir_results_dir.exists():
        found_files = list(mir_results_dir.glob("MIR_Results_*.csv"))
        found_files.extend(mir_results_dir.glob("MIR_Results_*.xlsx"))
        if found_files:
            mir_files.extend(found_files)
            print(f"   📁 在 01_MIR 目录中找到文件")
    
    # 向后兼容：查找旧的工作目录结构（run_YYYYMMDD_HHMMSS/01_MIR/）
    if not mir_files:
        work_dirs = sorted(output_dir.glob("run_*"), reverse=True)
        for work_dir in work_dirs:
            # 新结构：01_MIR 目录
            mir_results_dir = work_dir / "01_MIR"
            if not mir_results_dir.exists():
                # 向后兼容：01_MIR_Results 目录
                mir_results_dir = work_dir / "01_MIR_Results"
            
            if mir_results_dir.exists():
                found_files = list(mir_results_dir.glob("MIR_Results_*.csv"))
                found_files.extend(mir_results_dir.glob("MIR_Results_*.xlsx"))
                if found_files:
                    mir_files.extend(found_files)
                    print(f"   📁 在工作目录中找到文件: {work_dir.name}")
                    break
    
    # 向后兼容：在output根目录查找（旧格式）
    if not mir_files:
        mir_files = sorted(output_dir.glob("MIR_Results_*.csv"), reverse=True)
        mir_files.extend(sorted(output_dir.glob("MIR_Results_*.xlsx"), reverse=True))
        if mir_files:
            print(f"   📁 在output根目录中找到文件（旧格式）")
    
    if not mir_files:
        print("❌ 错误：未找到 MIR_Results_*.csv 或 *.xlsx 文件！")
        work_dirs = list(output_dir.glob("run_*"))
        if work_dirs:
            print(f"   找到 {len(work_dirs)} 个工作目录，但未找到MIR结果文件")
        print("   文件应位于：output/01_MIR/ 或 output/run_*/01_MIR/ 或 output/")
        input("\n按 Enter 键退出...")
        return

    # 使用最新的文件（按修改时间排序）
    selected_file = max(mir_files, key=lambda p: p.stat().st_mtime)
    print(f"✅ 使用最新 MIR 文件：{selected_file.name}")

    # 根据文件扩展名选择正确的读取方法
    if selected_file.suffix.lower() == '.xlsx':
        try:
            df = pd.read_excel(selected_file, engine='openpyxl')
        except ImportError:
            print("❌ 错误：需要安装 openpyxl 来读取 Excel 文件！")
            print("   请运行: pip install openpyxl")
            input("\n按 Enter 键退出...")
            return
    else:
        df = pd.read_csv(selected_file, encoding='utf-8-sig')
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

    # 保存到 01_MIR 目录
    mir_dir = output_dir / "01_MIR"
    mir_dir.mkdir(parents=True, exist_ok=True)
    
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = mir_dir / f"MIR_Results_with_VPO_DEBUG_{date_str}.csv"
    mir_with_vpo.to_csv(output_file, index=False, encoding="utf-8-sig")

    print()
    print("✅ 已生成带 VPO 的调试文件：")
    print(f"   {output_file}")
    print(f"   共写入 {max_count} 条 VPO 记录（总行数：{len(mir_with_vpo)}）")
    print()
    
    # 将 VPO 填写到 SPARK 文件夹里最新的文件中
    print("=" * 80)
    print("📝 正在将 VPO 填写到 SPARK 文件夹...")
    print("=" * 80)
    
    spark_files = []
    
    # 查找 SPARK 文件夹（支持新旧两种结构）
    # 新结构：output/run_*/Spark/
    work_dirs = sorted(output_dir.glob("run_*"), reverse=True)
    for work_dir in work_dirs:
        spark_dir = work_dir / "Spark"
        if spark_dir.exists():
            found_files = list(spark_dir.glob("MIR_Results_For_Spark_*.xlsx"))
            found_files.extend(list(spark_dir.glob("MIR_Results_For_Spark_*.csv")))
            if found_files:
                spark_files.extend(found_files)
                print(f"   📁 在工作目录中找到 SPARK 文件: {work_dir.name}")
                break
    
    # 向后兼容：旧结构 output/02_SPARK/
    if not spark_files:
        spark_dir = output_dir / "02_SPARK"
        if spark_dir.exists():
            found_files = list(spark_dir.glob("MIR_Results_For_Spark_*.xlsx"))
            found_files.extend(list(spark_dir.glob("MIR_Results_For_Spark_*.csv")))
            if found_files:
                spark_files.extend(found_files)
                print(f"   📁 在 02_SPARK 目录中找到文件")
    
    if not spark_files:
        print("⚠️ 警告：未找到 SPARK 文件夹中的 MIR_Results_For_Spark_*.xlsx 文件")
        print("   跳过填写 VPO 到 SPARK 文件")
    else:
        # 使用最新的文件（按修改时间排序）
        selected_spark_file = max(spark_files, key=lambda p: p.stat().st_mtime)
        print(f"✅ 使用最新 SPARK 文件：{selected_spark_file.name}")
        
        try:
            # 读取 SPARK 文件
            if selected_spark_file.suffix.lower() == '.xlsx':
                try:
                    spark_df = pd.read_excel(selected_spark_file, engine='openpyxl')
                except ImportError:
                    print("❌ 错误：需要安装 openpyxl 来读取 Excel 文件！")
                    print("   请运行: pip install openpyxl")
                else:
                    # 填写 VPO 到 SPARK 文件
                    vpo_col_name = "VPO"
                    
                    if vpo_col_name in spark_df.columns:
                        print(f"⚠️ 警告：SPARK 文件中已存在列 '{vpo_col_name}'，将覆盖该列的值。")
                    
                    spark_df[vpo_col_name] = ""
                    
                    # 使用反向匹配，和 MIR 文件一样的方式
                    spark_max_count = min(len(spark_df), len(vpo_list_reversed))
                    for i in range(spark_max_count):
                        spark_df.at[spark_df.index[i], vpo_col_name] = vpo_list_reversed[i]
                    
                    # 保存文件（覆盖原文件）
                    try:
                        spark_df.to_excel(selected_spark_file, index=False, engine='openpyxl')
                        print(f"✅ 已更新 SPARK 文件：{selected_spark_file.name}")
                        print(f"   共写入 {spark_max_count} 条 VPO 记录（总行数：{len(spark_df)}）")
                    except Exception as e:
                        # 如果 Excel 保存失败，尝试保存为 CSV
                        print(f"⚠️ 保存 Excel 文件失败: {e}，尝试保存为 CSV 格式...")
                        csv_file = selected_spark_file.with_suffix('.csv')
                        spark_df.to_csv(csv_file, index=False, encoding="utf-8-sig")
                        print(f"✅ 已保存为 CSV 文件：{csv_file.name}")
            else:
                # CSV 文件
                spark_df = pd.read_csv(selected_spark_file, encoding='utf-8-sig')
                
                # 填写 VPO 到 SPARK 文件
                vpo_col_name = "VPO"
                
                if vpo_col_name in spark_df.columns:
                    print(f"⚠️ 警告：SPARK 文件中已存在列 '{vpo_col_name}'，将覆盖该列的值。")
                
                spark_df[vpo_col_name] = ""
                
                # 使用反向匹配，和 MIR 文件一样的方式
                spark_max_count = min(len(spark_df), len(vpo_list_reversed))
                for i in range(spark_max_count):
                    spark_df.at[spark_df.index[i], vpo_col_name] = vpo_list_reversed[i]
                
                # 保存文件（覆盖原文件）
                spark_df.to_csv(selected_spark_file, index=False, encoding="utf-8-sig")
                print(f"✅ 已更新 SPARK 文件：{selected_spark_file.name}")
                print(f"   共写入 {spark_max_count} 条 VPO 记录（总行数：{len(spark_df)}）")
        except Exception as e:
            print(f"❌ 错误：更新 SPARK 文件时出错: {e}")
            import traceback
            print(traceback.format_exc())
    
    print()
    input("按 Enter 键退出...")


if __name__ == "__main__":
    # 命令行可选参数：--no-wait  表示不等待，立即收集 VPO
    no_wait_flag = "--no-wait" in sys.argv
    main(no_wait=no_wait_flag)


