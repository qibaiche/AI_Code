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
    
    # 读取MIR结果（在多个位置查找）
    print("🔍 查找 MIR 结果文件...")
    mir_files = []
    
    # 优先在output目录查找（MIR结果文件默认保存在这里）
    output_dir = parent_dir / "output"
    
    search_locations = [
        (output_dir, "output目录（推荐）"),
        (current_dir, "当前目录 (spark/)"),
        (parent_dir, "父目录 (auto-vpo/)"),
    ]
    
    mole_dir = parent_dir / "mole"
    if mole_dir.exists():
        search_locations.append((mole_dir, "mole目录"))
    
    for location, description in search_locations:
        files = list(location.glob("MIR_Results_*.csv"))
        if files:
            print(f"   ✅ 在 {description} 找到 {len(files)} 个文件")
        mir_files.extend(files)
    
    if not mir_files:
        print()
        print("❌ 错误：未找到 MIR 结果文件！")
        print()
        print("📁 已搜索以下位置：")
        for location, description in search_locations:
            print(f"   - {location}")
        print()
        print("💡 解决方法：")
        print("   1. 确认文件名格式为：MIR_Results_*.csv")
        print("   2. 将文件放在以上任一目录")
        print("   3. 查看 README.md 了解详细说明")
        input("\n按 Enter 键退出...")
        return
    
    # 使用最新的文件
    selected_file = sorted(mir_files, reverse=True)[0]
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

    # 使用第一个SourceLot的第一行（不再考虑多个Operation的情形）
    # 查找SourceLot列（支持多种命名格式）
    source_lot_col = None
    for col in df.columns:
        col_upper = str(col).strip().upper()
        if col_upper in ['SOURCELOT', 'SOURCE LOT', 'SOURCE_LOT', 'SOURCELOTS', 'SOURCE LOTS']:
            source_lot_col = col
            break
    
    if source_lot_col is None:
        print("❌ 错误：未找到SourceLot列！")
        print(f"   可用列: {df.columns.tolist()}")
        input("\n按 Enter 键退出...")
        return
    
    first_lot_value = df[source_lot_col].iloc[0]
    first_row = df[df[source_lot_col] == first_lot_value].iloc[0]
    
    # 安全地获取列值（支持多种列名格式）
    first_lot = str(first_row.get(source_lot_col, '')).strip()
    
    # 查找Part Type列
    part_type_col = None
    for col in df.columns:
        col_upper = str(col).strip().upper()
        if col_upper in ['PART TYPE', 'PARTTYPE', 'PART_TYPE']:
            part_type_col = col
            break
    first_part_type = str(first_row.get(part_type_col, '')).strip() if part_type_col else ''
    
    # 查找Operation列（可选）
    operation_col = None
    for col in df.columns:
        col_upper = str(col).strip().upper()
        if col_upper in ['OPERATION', 'OP', 'OPN']:
            operation_col = col
            break
    first_operation = str(first_row.get(operation_col, '')).strip() if operation_col else None
    
    # 查找Eng ID列（支持多种命名格式）
    eng_id_col = None
    for col in df.columns:
        col_upper = str(col).strip().upper()
        if col_upper in ['ENG ID', 'ENGID', 'ENG_ID', 'ENGINEERING ID', 'ENGINEERING_ID']:
            eng_id_col = col
            break
    first_eng_id = str(first_row.get(eng_id_col, '')).strip() if eng_id_col else None
    
    # 验证必需字段
    if not first_lot:
        print("❌ 错误：SourceLot值为空！")
        input("\n按 Enter 键退出...")
        return
    
    if not first_part_type:
        print("❌ 错误：Part Type值为空！")
        input("\n按 Enter 键退出...")
        return
    
    # 读取More options字段（如果存在）
    unit_test_time = first_row.get('Unit test time', None)
    retest_rate = first_row.get('Retest rate', None)
    hri_mrv = first_row.get('HRI / MRV:', None)
    
    # 处理空值并转换为字符串（处理numpy.int64等类型）
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

    submitter = SparkSubmitter(config.spark)
    
    try:
        print("=" * 80)
        print("开始自动化流程...")
        print("=" * 80)
        print()
        
        print("步骤 1/13: 打开网页...")
        submitter._init_driver()
        submitter._navigate_to_page()
        print("✅ 完成\n")
                
        print("步骤 2/13: 点击Add New...")
        if not submitter._click_add_new_button():
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 3/13: 填写TP路径...")
        if not submitter._fill_test_program_path(config.paths.tp_path):
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 4/13: Add New Experiment...")
        if not submitter._click_add_new_experiment():
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 5/13: 选择VPO类别...")
        if not submitter._select_vpo_category(config.spark.vpo_category):
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 6/13: 填写实验信息...")
        if not submitter._fill_experiment_info(config.spark.step, config.spark.tags):
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 7/13: 添加Lot name...")
        if not submitter._add_lot_name(str(first_lot)):
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 8/13: 选择Part Type...")
        if not submitter._select_parttype(str(first_part_type)):
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 9/13: 点击Flow标签...")
        if not submitter._click_flow_tab():
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        # Operation是可选的，如果存在则选择
        if first_operation:
            print("步骤 10/13: 选择Operation...")
            if not submitter._select_operation(str(first_operation)):
                print("❌ 失败\n")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
        else:
            print("步骤 10/13: 跳过Operation（文件中未提供）\n")
        
        # Eng ID是可选的，如果存在则选择
        if first_eng_id:
            print("步骤 11/13: 选择Eng ID...")
            if not submitter._select_eng_id(str(first_eng_id)):
                print("❌ 失败\n")
                input("\n按 Enter 键退出...")
                return
            print("✅ 完成\n")
        else:
            print("步骤 11/13: 跳过Eng ID（文件中未提供）\n")
        
        print("步骤 12/13: 点击More options标签...")
        if not submitter._click_more_options_tab():
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 13/13: 填写More options字段...")
        if not submitter._fill_more_options(unit_test_time, retest_rate, hri_mrv):
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print("步骤 14/14: 点击Roll按钮...")
        if not submitter._click_roll_button():
            print("❌ 失败\n")
            input("\n按 Enter 键退出...")
            return
        print("✅ 完成\n")
        
        print()
        print("=" * 80)
        print("🎉 所有步骤完成！")
        print("=" * 80)
        print()
        print("📊 填写摘要：")
        print(f"   Material:")
        print(f"      - Lot: {first_lot}")
        print(f"      - Part Type: {first_part_type}")
        print(f"   Flow:")
        print(f"      - Operation: {first_operation or '(未提供)'}")
        print(f"      - Eng ID: {first_eng_id or '(未提供)'}")
        print(f"   More options:")
        print(f"      - Unit test time: {unit_test_time or '(未填写)'}")
        print(f"      - Retest rate: {retest_rate or '(未填写)'}")
        print(f"      - HRI / MRV: {hri_mrv or 'DEFAULT'}")
        print()
        print("=" * 80)
        print("✅ 自动化流程执行成功！")
        print("   请在浏览器中检查填写结果")
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

