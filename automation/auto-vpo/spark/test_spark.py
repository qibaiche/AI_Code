"""Spark自动化测试 - Material和Flow标签"""
import sys
import time
from pathlib import Path

# 添加父目录到路径（workflow_automation在父目录）
current_dir = Path(__file__).parent
parent_dir = current_dir.parent  # automation/auto-vpo/
sys.path.insert(0, str(parent_dir))

# Optional dependency: when pandas is unavailable (e.g., in lightweight test
# environments), skip importing this helper script instead of failing
# collection.
try:
    import pandas as pd
except ImportError:
    print("❌ 错误：pandas 模块未安装！")
    print("   请运行: pip install pandas")
    input("\n按 Enter 键退出...")
    sys.exit(1)

from workflow_automation.config_loader import load_config
from workflow_automation.main import configure_logging
from workflow_automation.spark_submitter import SparkSubmitter
from workflow_automation.utils.keyboard_listener import start_global_listener, is_esc_pressed, stop_global_listener

def main(skip_to_lot: bool = False):
    print("=" * 80)
    print("🚀 Spark 自动化工具")
    if skip_to_lot:
        print("   [调试模式：从添加 Lot 开始]")
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
    
    # 配置日志系统（使用统一的日志配置）
    configure_logging(config)
    
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
    
    # 查找最新的 MIR 结果文件（支持新的目录结构）
    print("🔍 查找 MIR 结果文件...")
    output_dir = parent_dir / "output"
    
    mir_files = []
    if output_dir.exists():
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
                    # 向后兼容：MIR 或 01_MIR_Results 目录
                    mir_results_dir = work_dir / "MIR"
                    if not mir_results_dir.exists():
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
        print()
        print("❌ 错误：未找到 MIR 结果文件！")
        print(f"   已检查目录：{output_dir}")
        if output_dir.exists():
            work_dirs = list(output_dir.glob("run_*"))
            if work_dirs:
                print(f"   找到 {len(work_dirs)} 个工作目录，但未找到MIR结果文件")
        print()
        print("💡 解决方法：")
        print("   1. 确认文件名格式为：MIR_Results_*.csv 或 MIR_Results_*.xlsx")
        print("   2. 确认 Mole 步骤已成功生成 MIR 结果文件")
        print("   3. 文件应位于：output/01_MIR/ 或 output/run_*/01_MIR/ 或 output/")
        print("   4. 查看 README.md 了解详细说明")
        input("\n按 Enter 键退出...")
        return
    
    # 使用最新的文件（按修改时间排序）
    selected_file = max(mir_files, key=lambda p: p.stat().st_mtime)
    print(f"   📄 使用文件：{selected_file.name}")
    print()
    
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
        print(f"   文件：{selected_file}")
        print()
        print("💡 解决方法：")
        print("   1. 检查文件是否损坏")
        print("   2. 重新生成 MIR 结果文件")
        input("\n按 Enter 键退出...")
        return
    
    print(f"✅ 成功读取 MIR 数据：{len(df)} 行")
    print()
    
    # 尝试合并 For Spark.csv
    print("🔍 查找并合并 For Spark.csv 文件...")
    spark_config_file = None
    possible_paths = [
        parent_dir / "input" / "For Spark.csv",
        parent_dir / "For Spark.csv",
    ]
    
    for path in possible_paths:
        if path.exists():
            spark_config_file = path
            break
    
    if spark_config_file:
        try:
            print(f"   📄 找到 For Spark.csv: {spark_config_file.name}")
            spark_df = pd.read_csv(spark_config_file, encoding='utf-8-sig')
            
            # 查找 SourceLot 列
            mir_source_lot_col = None
            for col in df.columns:
                col_upper = str(col).strip().upper()
                if col_upper in ['SOURCELOT', 'SOURCE LOT', 'SOURCE_LOT']:
                    mir_source_lot_col = col
                    break
            
            spark_source_lot_col = None
            for col in spark_df.columns:
                col_upper = str(col).strip().upper()
                if col_upper in ['SOURCELOT', 'SOURCE LOT', 'SOURCE_LOT']:
                    spark_source_lot_col = col
                    break
            
            if mir_source_lot_col and spark_source_lot_col:
                # 标准化列名
                if mir_source_lot_col != 'Source Lot':
                    df = df.rename(columns={mir_source_lot_col: 'Source Lot'})
                if spark_source_lot_col != 'Source Lot':
                    spark_df = spark_df.rename(columns={spark_source_lot_col: 'Source Lot'})
                
                # 标准化 Source Lot 值
                df['Source Lot'] = df['Source Lot'].astype(str).str.strip()
                spark_df['Source Lot'] = spark_df['Source Lot'].astype(str).str.strip()
                
                # 建立映射
                spark_config_dict = {}
                for _, row in spark_df.iterrows():
                    source_lot = str(row['Source Lot']).strip()
                    if source_lot and source_lot != 'nan':
                        spark_config_dict[source_lot] = row.to_dict()
                
                # 合并数据
                # 特别处理：Part Type 字段始终以 For Spark.csv 为准（即使 MIR 结果中有值）
                part_type_cols = ['Part Type', 'PartType', 'PART_TYPE', 'Part_Type']  # 可能的 Part Type 列名
                
                for source_lot, config_row in spark_config_dict.items():
                    mask = df['Source Lot'] == source_lot
                    if mask.any():
                        for col, value in config_row.items():
                            if col == 'Source Lot':
                                continue
                            if pd.notna(value) and str(value).strip():
                                if col not in df.columns:
                                    df[col] = None
                                
                                # 检查是否是 Part Type 相关的列
                                col_upper = str(col).strip().upper()
                                is_part_type = any(pt_col.upper() == col_upper for pt_col in part_type_cols)
                                
                                if is_part_type:
                                    # Part Type 字段：始终以 For Spark.csv 为准，直接覆盖
                                    df.loc[mask, col] = value
                                    print(f"   Source Lot '{source_lot}': Part Type 以 For Spark.csv 为准，值='{value}'")
                                else:
                                    # 其他字段：直接覆盖（For Spark.csv 优先）
                                    df.loc[mask, col] = value
                
                print(f"   ✅ 成功合并 {len(spark_config_dict)} 个 Source Lot 的配置")
            else:
                print("   ⚠️ 未找到 SourceLot 列，跳过合并")
        except Exception as e:
            print(f"   ⚠️ 合并 For Spark.csv 时出错: {e}，将使用原始 MIR 结果")
    else:
        print("   ℹ️ 未找到 For Spark.csv 文件，将使用原始 MIR 结果")
    
    # 特殊处理：将 MIR 结果中的 Units_Count_Expected 映射到 Quantity 列（无论是否有 For Spark.csv）
    if 'Units_Count_Expected' in df.columns:
        print("   🔄 发现 Units_Count_Expected 列，将其映射到 Quantity 列...")
        if 'Quantity' not in df.columns:
            df['Quantity'] = None
        # 将 Units_Count_Expected 的值复制到 Quantity（如果 Quantity 为空或不存在）
        mask_quantity_empty = df['Quantity'].isna() | (df['Quantity'] == '')
        df.loc[mask_quantity_empty, 'Quantity'] = df.loc[mask_quantity_empty, 'Units_Count_Expected']
        print(f"   ✅ 已将 {mask_quantity_empty.sum()} 行的 Units_Count_Expected 映射到 Quantity")
    elif 'Units_Count_Actual' in df.columns:
        # 如果没有 Units_Count_Expected，尝试使用 Units_Count_Actual
        print("   🔄 发现 Units_Count_Actual 列，将其映射到 Quantity 列...")
        if 'Quantity' not in df.columns:
            df['Quantity'] = None
        mask_quantity_empty = df['Quantity'].isna() | (df['Quantity'] == '')
        df.loc[mask_quantity_empty, 'Quantity'] = df.loc[mask_quantity_empty, 'Units_Count_Actual']
        print(f"   ✅ 已将 {mask_quantity_empty.sum()} 行的 Units_Count_Actual 映射到 Quantity")
    
    # 保存合并后的文件到 output/02_SPARK 目录
    print()
    print("💾 保存合并后的文件...")
    output_dir = parent_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 保存到 02_SPARK 目录
    spark_dir = output_dir / "02_SPARK"
    spark_dir.mkdir(parents=True, exist_ok=True)
    
    from datetime import datetime
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    merged_file = spark_dir / f"MIR_Results_For_Spark_{date_str}.xlsx"
    
    try:
        df.to_excel(merged_file, index=False, engine='openpyxl')
        print(f"   ✅ 已生成合并文件: {merged_file.name}")
        print(f"   📁 完整路径: {merged_file.absolute()}")
        print(f"   📊 包含 {len(df)} 行数据")
        print(f"   📋 列: {df.columns.tolist()}")
        
        # 验证文件是否真的被创建
        if merged_file.exists():
            file_size = merged_file.stat().st_size
            print(f"   📦 文件大小: {file_size:,} 字节")
        else:
            print(f"   ❌ 警告：文件未成功创建")
    except Exception as e:
        # 如果Excel保存失败，尝试保存为CSV
        print(f"   ⚠️ 保存Excel文件失败: {e}，尝试保存为CSV格式...")
        merged_file = spark_dir / f"MIR_Results_For_Spark_{date_str}.csv"
        try:
            df.to_csv(merged_file, index=False, encoding='utf-8-sig')
            print(f"   ✅ 已生成合并文件: {merged_file.name} (CSV格式)")
            print(f"   📁 完整路径: {merged_file.absolute()}")
            print(f"   📊 包含 {len(df)} 行数据")
            
            if merged_file.exists():
                file_size = merged_file.stat().st_size
                print(f"   📦 文件大小: {file_size:,} 字节")
        except Exception as e2:
            print(f"   ❌ 保存CSV文件也失败: {e2}")
    
    print()
    
    # 显示文件中的列名（用于调试）
    print(f"📋 文件列名: {df.columns.tolist()}")
    print()

    submitter = SparkSubmitter(config.spark)
    
    # 启动键盘监听器（ESC 键停止）
    def on_escape():
        """ESC 键按下时的处理"""
        print("\n" + "=" * 80)
        print("⚠️  检测到 ESC 键，正在停止程序...")
        print("=" * 80)
        submitter._close_driver()
    
    start_global_listener(on_escape)
    print("💡 提示：按 ESC 键可随时停止程序\n")
    
    try:
        print("=" * 80)
        print("开始自动化流程（按CSV顺序依次提交所有MIR）...")
        print("=" * 80)
        print()
        
        # 检查 ESC 键
        if is_esc_pressed():
            print("程序已停止（ESC 键）")
            return
        
        # 打开网页
        print("步骤 1: 打开Spark网页...")
        submitter._init_driver()
        submitter._navigate_to_page()
        print("✅ 完成\n")
        
        # 检查 ESC 键
        if is_esc_pressed():
            print("程序已停止（ESC 键）")
            submitter._close_driver()
            return
        
        # 如果使用 --from-lot 参数，跳过前面的步骤
        if skip_to_lot:
            print("=" * 80)
            print("⏭️  调试模式：跳过前面的步骤，直接从添加 Lot 开始")
            print("=" * 80)
            print()
            print("💡 请确保浏览器中已完成以下步骤：")
            print("   1. 已点击 'Add New' 按钮")
            print("   2. 已填写 TP 路径并点击 'Add New Experiment'")
            print("   3. 已选择 VPO 类别并填写实验信息（Step 和 Tags）")
            print("   4. 当前页面已准备好添加 Lot name（Material 标签页）")
            print()
            print("⚠️  如果还未完成，请在浏览器中手动完成上述步骤")
            print()
            input("完成后按 Enter 键继续...")
            print()
        
        total_rows = len(df)
        
        for row_num, (idx, row) in enumerate(df.iterrows(), start=1):
            # 检查 ESC 键
            if is_esc_pressed():
                print("\n" + "=" * 80)
                print(f"⚠️  程序已停止（ESC 键）")
                print(f"   已处理 {row_num - 1}/{total_rows} 行")
                print("=" * 80)
                break
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
            
            # MIR（已移除 instructions 步骤，不再需要）
            # mir_value = None
            # for col in row.index:
            #     col_upper = str(col).strip().upper()
            #     if col_upper in ['MIR', 'MIR#', 'MIR_NUMBER']:
            #         if pd.notna(row[col]) and str(row[col]).strip():
            #             mir_value = str(row[col]).strip()
            #         break
            
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
            
            # 根据 skip_to_lot 参数决定是否跳过前面的步骤
            if not skip_to_lot:
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
            else:
                # 调试模式：跳过前面的步骤，直接从添加 Lot 开始
                print("⏭️  跳过步骤 2-6（Add New、TP路径、Experiment、VPO类别、实验信息）\n")
            
            print("步骤 7: 添加Lot name...")
            # 查找Quantity列（用于设置units数量）
            quantity = None
            for col in row.index:
                col_upper = str(col).strip().upper()
                if col_upper in ['QUANTITY', 'QTY', 'UNITS', 'UNIT COUNT', 'COUNT']:
                    if pd.notna(row[col]) and str(row[col]).strip():
                        try:
                            # 确保转换为纯整数，去除所有空格和占位符
                            raw_value = str(row[col]).strip()
                            # 移除所有空格
                            raw_value = raw_value.replace(' ', '').replace('\t', '').replace('\n', '')
                            quantity = int(float(raw_value))
                            print(f"   从数据中读取Quantity: {quantity} (纯数字格式)")
                        except (ValueError, TypeError):
                            print(f"   ⚠️ Quantity值无效: {row[col]}，跳过设置units数量")
                            quantity = None
                    break
            
            if not submitter._add_lot_name(str(source_lot), quantity):
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
            
            # 已移除 instructions 步骤
            # if mir_value:
            #     print("步骤 11.5: 点击instructions图标并填写MIR#...")
            #     if not submitter._click_instructions_and_fill_mir(str(mir_value)):
            #         print("❌ 失败：点击instructions并填写MIR#失败")
            #         input("\n按 Enter 键退出...")
            #         return
            #     print("✅ 完成\n")
            # else:
            #     print("步骤 11.5: 跳过填写MIR#（文件中未提供MIR值）\n")
            
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
            
            # 检查 ESC 键
            if is_esc_pressed():
                print("\n" + "=" * 80)
                print(f"⚠️  程序已停止（ESC 键）")
                print(f"   已处理 {row_num}/{total_rows} 行")
                print("=" * 80)
                break
        
        if not is_esc_pressed():
            print()
            print("=" * 80)
            print("🎉 所有MIR已按CSV顺序依次提交完成！")
            print("=" * 80)
            print()
            print("   请在浏览器中检查各个Lot的提交结果")
            print("=" * 80)
            print()
            input("按 Enter 键关闭浏览器...")
        else:
            print("\n程序已停止（ESC 键）")
                
    except KeyboardInterrupt:
        print("\n" + "=" * 80)
        print("⚠️  程序被用户中断（Ctrl+C）")
        print("=" * 80)
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
        # 停止键盘监听器
        stop_global_listener()
        # 关闭浏览器
        submitter._close_driver()

if __name__ == "__main__":
    # 检查命令行参数
    skip_to_lot = "--from-lot" in sys.argv or "--skip-to-lot" in sys.argv
    
    if skip_to_lot:
        print("💡 提示：使用 --from-lot 参数，将从添加 Lot 开始执行")
        print("   请确保已手动完成前面的步骤（Add New、TP路径、Experiment等）")
        print()
    
    main(skip_to_lot=skip_to_lot)

