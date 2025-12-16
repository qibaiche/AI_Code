"""Spark自动化测试 - Material和Flow标签"""
import sys
import time
from pathlib import Path
import pandas as pd

# 添加父目录到路径（workflow_automation在父目录）
current_dir = Path(__file__).parent
parent_dir = current_dir.parent  # automation/auto-vpo/
sys.path.insert(0, str(parent_dir))

from workflow_automation.config_loader import load_config
from workflow_automation.spark_submitter import SparkSubmitter

def main():
    print("=" * 80)
    print("🚀 Spark 自动化工具")
    print("=" * 80)
    print()
    
    # 加载配置
    config_path = parent_dir / "workflow_automation" / "config.yaml"
    
    # 检查配置文件是否存在
    if not config_path.exists():
        print("❌ 错误：配置文件不存在！")
        print(f"   预期位置：{config_path}")
        print()
        print("💡 解决方法：")
        print("   1. 确认文件路径是否正确")
        print("   2. 查看 README.md 了解配置方法")
        input("\n按 Enter 键退出...")
        return
    
    config = load_config(config_path)
    
    # 显示当前配置
    print("📋 当前配置：")
    print(f"   TP路径: {config.paths.tp_path}")
    print(f"   VPO类别: {config.spark.vpo_category}")
    print(f"   Step: {config.spark.step}")
    print(f"   Tags: {config.spark.tags}")
    print()
    
    # 提示用户确认
    print("⚠️  请确认以上配置是否正确")
    print("   如需修改，请编辑：workflow_automation/config.yaml")
    print("   详细说明请查看：spark/README.md")
    print()
    
    # 只在 output 目录查找最新的 MIR 结果文件
    print("🔍 查找 MIR 结果文件（仅output目录）...")
    output_dir = parent_dir / "output"
    if not output_dir.exists():
        print("❌ 错误：output 目录不存在！")
        print(f"   预期位置：{output_dir}")
        print()
        print("💡 解决方法：")
        print("   1. 确认 Mole 步骤已成功执行并生成 MIR_Results_*.csv")
        print("   2. 确认 output 目录存在且位于 auto-vpo 根目录下")
        input("\n按 Enter 键退出...")
        return
    
    mir_files = sorted(output_dir.glob("MIR_Results_*.csv"), reverse=True)
    
    if not mir_files:
        print()
        print("❌ 错误：未在 output 目录找到 MIR 结果文件！")
        print(f"   已检查目录：{output_dir}")
        print()
        print("💡 解决方法：")
        print("   1. 确认文件名格式为：MIR_Results_*.csv")
        print("   2. 确认 Mole 步骤已成功生成 MIR 结果文件")
        print("   3. 查看 README.md 了解详细说明")
        input("\n按 Enter 键退出...")
        return
    
    # 使用 output 目录中最新的文件
    selected_file = mir_files[0]
    print(f"   📄 使用文件：{selected_file.name}")
    print()
    
    df = pd.read_csv(selected_file)
    if df.empty:
        print("❌ 错误：MIR 结果文件为空！")
        print(f"   文件：{selected_file}")
        print()
        print("💡 解决方法：")
        print("   1. 检查文件是否损坏")
        print("   2. 重新生成 MIR 结果文件")
        input("\n按 Enter 键退出...")
        return
    
    print(f"✅ 成功读取 MIR 数据：{len(df)} 行")
    print()
    
    # 显示文件中的列名（用于调试）
    print(f"📋 文件列名: {df.columns.tolist()}")
    print()

    submitter = SparkSubmitter(config.spark)
    
    try:
        print("=" * 80)
        print("开始自动化流程（按CSV顺序依次提交所有MIR）...")
        print("=" * 80)
        print()
        
        # 打开网页
        print("步骤 1: 打开Spark网页...")
        submitter._init_driver()
        submitter._navigate_to_page()
        print("✅ 完成\n")
        
        total_rows = len(df)
        
        for row_num, (idx, row) in enumerate(df.iterrows(), start=1):
            print("=" * 80)
            print(f"处理第 {row_num}/{total_rows} 行 MIR 数据 (DataFrame索引: {idx})")
            print("=" * 80)
            print(f"行数据: {row.to_dict()}")
            print()
            
            # 查找SourceLot
            source_lot = None
            for col in row.index:
                col_upper = str(col).strip().upper()
                if col_upper in ['SOURCELOT', 'SOURCE LOT', 'SOURCE_LOT']:
                    source_lot = str(row[col]).strip() if pd.notna(row[col]) else ''
                    break
            
            if not source_lot:
                print(f"⚠️ 第 {row_num} 行SourceLot为空，跳过")
                continue
            
            # 查找Part Type
            part_type = None
            for col in row.index:
                col_upper = str(col).strip().upper()
                if col_upper in ['PART TYPE', 'PARTTYPE', 'PART_TYPE']:
                    part_type = str(row[col]).strip() if pd.notna(row[col]) else ''
                    break
            
            if not part_type:
                print(f"⚠️ 第 {row_num} 行Part Type为空，跳过")
                continue
            
            # Operation（可选）
            operation = None
            for col in row.index:
                col_upper = str(col).strip().upper()
                if col_upper in ['OPERATION', 'OP', 'OPN']:
                    if pd.notna(row[col]) and str(row[col]).strip():
                        operation = str(row[col]).strip()
                    break
            
            # Eng ID（可选）
            eng_id = None
            for col in row.index:
                col_upper = str(col).strip().upper()
                if col_upper in ['ENG ID', 'ENGID', 'ENG_ID', 'ENGINEERING ID', 'ENGINEERING_ID']:
                    if pd.notna(row[col]) and str(row[col]).strip():
                        eng_id = str(row[col]).strip()
                    break
            
            # More options 字段
            unit_test_time = row.get('Unit test time', None)
            retest_rate = row.get('Retest rate', None)
            hri_mrv = row.get('HRI / MRV:', None)
            
            if pd.isna(unit_test_time) or str(unit_test_time).strip() == '':
                unit_test_time = None
            else:
                unit_test_time = str(unit_test_time).strip()
            
            if pd.isna(retest_rate) or str(retest_rate).strip() == '':
                retest_rate = None
            else:
                retest_rate = str(retest_rate).strip()
            
            if pd.isna(hri_mrv) or str(hri_mrv).strip() == '':
                hri_mrv = None
            else:
                hri_mrv = str(hri_mrv).strip()
            
            # 对于每一行都执行一套完整流程
            print("步骤 2: 点击Add New...")
            if not submitter._click_add_new_button():
                print("❌ 失败：点击Add New失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 3: 填写TP路径...")
            if not submitter._fill_test_program_path(config.paths.tp_path):
                print("❌ 失败：填写TP路径失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 4: Add New Experiment...")
            if not submitter._click_add_new_experiment():
                print("❌ 失败：点击Add New Experiment失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 5: 选择VPO类别...")
            if not submitter._select_vpo_category(config.spark.vpo_category):
                print("❌ 失败：选择VPO类别失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 6: 填写实验信息...")
            if not submitter._fill_experiment_info(config.spark.step, config.spark.tags):
                print("❌ 失败：填写实验信息失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 7: 添加Lot name...")
            if not submitter._add_lot_name(str(source_lot)):
                print("❌ 失败：添加Lot name失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 8: 选择Part Type...")
            if not submitter._select_parttype(str(part_type)):
                print("❌ 失败：选择Part Type失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 9: 点击Flow标签...")
            if not submitter._click_flow_tab():
                print("❌ 失败：点击Flow标签失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            if operation:
                print("步骤 10: 选择Operation...")
                if not submitter._select_operation(str(operation)):
                    print("❌ 失败：选择Operation失败")
                    input("\n按 Enter 键退出...")
                    return
                print("✅ 完成\n")
            else:
                print("步骤 10: 跳过Operation（文件中未提供）\n")
            
            if eng_id:
                print("步骤 11: 选择Eng ID...")
                if not submitter._select_eng_id(str(eng_id)):
                    print("❌ 失败：选择Eng ID失败")
                    input("\n按 Enter 键退出...")
                    return
                print("✅ 完成\n")
            else:
                print("步骤 11: 跳过Eng ID（文件中未提供）\n")
            
            print("步骤 12: 点击More options标签...")
            if not submitter._click_more_options_tab():
                print("❌ 失败：点击More options标签失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 13: 填写More options字段...")
            if not submitter._fill_more_options(unit_test_time, retest_rate, hri_mrv):
                print("❌ 失败：填写More options字段失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            print("步骤 14: 点击Roll按钮提交当前MIR...")
            if not submitter._click_roll_button():
                print("❌ 失败：点击Roll按钮失败")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
            
            # 简单等待，给页面一点时间处理提交
            time.sleep(2.0)
        
        print()
        print("=" * 80)
        print("🎉 所有MIR已按CSV顺序依次提交完成！")
        print("=" * 80)
        print()
        print("   请在浏览器中检查各个Lot的提交结果")
        print("=" * 80)
        print()
        input("按 Enter 键关闭浏览器...")
                
    except Exception as e:
        print()
        print("=" * 80)
        print("❌ 执行过程中发生错误")
        print("=" * 80)
        print(f"错误信息: {e}")
        print()
        import traceback
        traceback.print_exc()
        print()
        print("💡 建议：")
        print("   1. 检查配置文件是否正确")
        print("   2. 查看上方的错误信息")
        print("   3. 查阅 README.md 获取帮助")
        print()
        input("按 Enter 键关闭...")
    finally:
        submitter._close_driver()

if __name__ == "__main__":
    main()

