"""Spark自动化测试 - Material、Flow和More options标签"""
import sys
from pathlib import Path
import pandas as pd

# 添加父目录到路径
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from workflow_automation.config_loader import load_config
from workflow_automation.spark_submitter import SparkSubmitter

def main():
    print("=" * 60)
    print("Spark自动化 - Material、Flow和More options标签")
    print("=" * 60)
    
    config_path = current_dir / "workflow_automation" / "config.yaml"
    config = load_config(config_path)
    
    # 读取MIR结果
    mir_files = list(current_dir.glob("MIR_Results_*.csv"))
    if not mir_files:
        print("❌ 未找到MIR结果文件")
        input()
        return
    
    df = pd.read_csv(sorted(mir_files, reverse=True)[0])
    if df.empty:
        print("❌ MIR结果文件为空")
        input()
        return

    # 使用第一个SourceLot对应的所有行，支持同一个Lot多个Operation/EngID
    first_lot_value = df['SourceLot'].iloc[0]
    lot_group = df[df['SourceLot'] == first_lot_value].reset_index(drop=True)

    # 第一行用于Material + Flow第一个condition + More options
    first_row = lot_group.iloc[0]
    first_lot = first_row['SourceLot']
    first_part_type = first_row['Part Type']
    first_operation = first_row['Operation']
    first_eng_id = first_row['Eng ID']

    # 读取More options字段（按Lot第一行）
    first_unit_test_time = first_row['Unit test time']
    first_retest_rate = first_row['Retest rate']
    first_hri_mrv = first_row['HRI / MRV:'] if 'HRI / MRV:' in lot_group.columns else None

    # 其余行用于Add new condition
    additional_conditions = lot_group.iloc[1:]

    submitter = SparkSubmitter(config.spark)
    
    try:
        print("1. 打开网页...")
        submitter._init_driver()
        submitter._navigate_to_page()
        print("✅\n")
        
        print("2. 点击Add New...")
        if not submitter._click_add_new_button():
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("3. 填写TP路径...")
        if not submitter._fill_test_program_path(config.paths.tp_path):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("4. Add New Experiment...")
        if not submitter._click_add_new_experiment():
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("5. 选择VPO类别...")
        if not submitter._select_vpo_category(config.spark.vpo_category):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("6. 填写实验信息...")
        if not submitter._fill_experiment_info(config.spark.step, config.spark.tags):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("7. 添加Lot name...")
        if not submitter._add_lot_name(str(first_lot)):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("8. 选择Part Type...")
        if not submitter._select_parttype(str(first_part_type)):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("9. 点击Flow标签...")
        if not submitter._click_flow_tab():
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("10. 选择Operation（第1个condition）...")
        if not submitter._select_operation(str(first_operation), condition_index=0):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("11. 选择Eng ID（第1个condition）...")
        if not submitter._select_eng_id(str(first_eng_id), condition_index=0):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")

        # 如果同一个Lot还有更多行，则为每一行添加一个新的condition
        if not additional_conditions.empty:
            print(f"11+. 为Lot {first_lot} 添加 {len(additional_conditions)} 个额外Flow condition...\n")

        # ⚠️ 关键修复：使用enumerate而不是iterrows的idx，确保condition_index从1开始连续
        for i, (_, row) in enumerate(additional_conditions.iterrows(), start=1):
            condition_index = i  # 第2个condition是i=1，第3个是i=2...
            op = str(row['Operation'])
            eng = str(row['Eng ID'])

            print(f"11.{i} 点击Add new condition（添加第{condition_index + 1}个condition）...")
            if not submitter._click_add_new_condition():
                print("❌ 失败\n")
                input()
                return
            print("✅\n")

            print(f"11.{i} 选择Operation（第{condition_index + 1}个condition，值={op}）...")
            if not submitter._select_operation(op, condition_index=condition_index):
                print("❌ 失败\n")
                input()
                return
            print("✅\n")

            print(f"11.{i} 选择Eng ID（第{condition_index + 1}个condition，值={eng}）...")
            if not submitter._select_eng_id(eng, condition_index=condition_index):
                print("❌ 失败\n")
                input()
                return
            print("✅\n")
        
        print("12. 点击More options标签...")
        if not submitter._click_more_options_tab():
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("13. 填写More options...")
        if not submitter._fill_more_options(
            str(first_unit_test_time),
            str(first_retest_rate),
            str(first_hri_mrv) if first_hri_mrv and pd.notna(first_hri_mrv) else None
        ):
            print("❌ 失败\n")
            input()
            return
        print("✅\n")
        
        print("=" * 60)
        print("✅ 所有标签填写完成！")
        print("=" * 60)
        print(f"Material: Lot={first_lot}, PartType={first_part_type}")
        print(f"Flow: 第1个condition Operation={first_operation}, EngID={first_eng_id}")
        if not additional_conditions.empty:
            for idx, row in additional_conditions.iterrows():
                cond_no = idx + 2
                print(f"      第{cond_no}个condition Operation={row['Operation']}, EngID={row['Eng ID']}")
        print(f"More options: UnitTestTime={first_unit_test_time}, RetestRate={first_retest_rate}, HRI/MRV={first_hri_mrv or 'default'}")
        print("\n按Enter关闭浏览器...")
        input()
        
    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        input()
    finally:
        submitter._close_driver()

if __name__ == "__main__":
    main()

