"""
统一自动化主程序
同时运行 PRD LOT 和 Lab TP Performance 自动化，并将结果放在一封邮件中
"""
import sys
import logging
from pathlib import Path
from datetime import datetime

# 添加路径
sys.path.insert(0, str(Path(__file__).parent))

from prd_lot_automation.config_loader import load_config
from prd_lot_automation.lot_reader import read_lots, split_batches
from prd_lot_automation.spf_runner import SQLPathFinderRunner
from prd_lot_automation.data_processing import (
    load_dataset,
    normalize_columns,
    apply_filters,
    build_quantity_table,
    build_pareto,
    collect_exceptions,
)
from prd_lot_automation.report_builder import build_pareto_table, save_report
from prd_lot_automation.close_sqlpathfinder import close_sqlpathfinder


def configure_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "unified_automation.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding="utf-8"),
        ],
    )


def run_prd_lot_automation(config_path: Path) -> Path:
    """运行 PRD LOT 自动化，返回报告路径"""
    logging.info("=" * 80)
    logging.info("开始执行 PRD LOT 自动化")
    logging.info("=" * 80)
    
    config = load_config(config_path)
    
    lots = read_lots(config.paths.lots_file)
    batches = split_batches(lots, config.processing.max_lots_per_batch)
    
    runner = SQLPathFinderRunner(config)
    all_data = []
    for idx, batch in enumerate(batches, start=1):
        logging.info("执行批次 %s/%s，LOT 数：%s", idx, len(batches), len(batch))
        runner.execute(batch)
        csv_path = runner.wait_for_output()
        df = load_dataset(csv_path)
        all_data.append(df)
    
    import pandas as pd
    
    merged_df = pd.concat(all_data).reset_index(drop=True)
    normalized_df = normalize_columns(merged_df, config)
    
    # 按 process_step 分组统计 Functional_bin Pareto
    functional_pareto_by_step = {}
    process_step_col = config.fields.process_step
    
    if process_step_col in normalized_df.columns:
        # 获取所有 process_step 类型
        process_steps = normalized_df[process_step_col].dropna().unique()
        logging.info(f"找到 {len(process_steps)} 个 process_step 类型: {list(process_steps)}")
        
        # 为每个 process_step 生成 Pareto 表
        for step in process_steps:
            step_df = normalized_df[normalized_df[process_step_col] == step].copy()
            if not step_df.empty:
                logging.info(f"处理 process_step: {step}，数据行数: {len(step_df)}")
                quantity, _ = build_quantity_table(step_df, config, bin_type="functional")
                if not quantity.empty:
                    percentages, row_totals, total_percentage = build_pareto(quantity)
                    pareto_table = build_pareto_table(quantity, percentages, row_totals, total_percentage)
                    functional_pareto_by_step[str(step)] = pareto_table
                    logging.info(f"✅ {step} 的 Functional_bin Pareto 生成完成")
                else:
                    logging.warning(f"⚠️ {step} 的数量表为空，跳过")
            else:
                logging.warning(f"⚠️ {step} 的数据为空，跳过")
        
        # 如果没有 process_step 数据，生成一个总的 Pareto 表
        if not functional_pareto_by_step:
            logging.warning("所有 process_step 的数据都为空，生成总的 Functional_bin Pareto")
            quantity, _ = build_quantity_table(normalized_df, config, bin_type="functional")
            percentages, row_totals, total_percentage = build_pareto(quantity)
            functional_pareto = build_pareto_table(quantity, percentages, row_totals, total_percentage)
        else:
            # 如果按 process_step 分组成功，不生成总的 Pareto 表
            functional_pareto = None
            logging.info("已按 process_step 分组，不生成总的 Functional_bin Pareto 表")
    else:
        logging.warning(f"⚠️ 未找到 process_step 列（{process_step_col}），生成总的 Functional_bin Pareto")
        # Functional_bin Pareto
        quantity, _ = build_quantity_table(normalized_df, config, bin_type="functional")
        percentages, row_totals, total_percentage = build_pareto(quantity)
        functional_pareto = build_pareto_table(quantity, percentages, row_totals, total_percentage)
    
    # 按 process_step 分组统计 Interface_bin Pareto（如果存在）
    interface_pareto = None
    interface_pareto_by_step = {}
    if config.fields.interface_bin in normalized_df.columns:
        if process_step_col in normalized_df.columns:
            # 按 process_step 分组
            process_steps = normalized_df[process_step_col].dropna().unique()
            for step in process_steps:
                step_df = normalized_df[normalized_df[process_step_col] == step].copy()
                if not step_df.empty:
                    interface_quantity, _ = build_quantity_table(step_df, config, bin_type="interface")
                    if not interface_quantity.empty:
                        interface_percentages, interface_row_totals, interface_total_percentage = build_pareto(interface_quantity)
                        pareto_table = build_pareto_table(
                            interface_quantity, interface_percentages, interface_row_totals, interface_total_percentage
                        )
                        interface_pareto_by_step[str(step)] = pareto_table
                        logging.info(f"✅ {step} 的 Interface_bin Pareto 生成完成")
            
            if interface_pareto_by_step:
                # 如果按 process_step 分组成功，不生成总的 Pareto 表
                interface_pareto = None
                logging.info("已按 process_step 分组，不生成总的 Interface_bin Pareto 表")
        else:
            # 没有 process_step 列，生成总的 Interface_bin Pareto
            interface_quantity, _ = build_quantity_table(normalized_df, config, bin_type="interface")
            if not interface_quantity.empty:
                interface_percentages, interface_row_totals, interface_total_percentage = build_pareto(interface_quantity)
                interface_pareto = build_pareto_table(
                    interface_quantity, interface_percentages, interface_row_totals, interface_total_percentage
                )
    
    # 按 process_step 分组统计 Retest Bin Pareto（使用不同的过滤条件：mut_within_subflow_latest_flag = N）
    retest_pareto = None
    retest_pareto_by_step = {}
    logging.info("开始生成 Retest Bin Pareto（mut_within_subflow_latest_flag = N）")
    retest_filters = {
        "mut_within_subflow_latest_flag": "N",
        "SUBSTRUCTURE_ID": "U1.U2"
    }
    retest_df = apply_filters(merged_df, retest_filters)
    retest_df[config.fields.functional_bin] = retest_df[config.fields.functional_bin].fillna("UNKNOWN")
    if config.fields.interface_bin in retest_df.columns:
        retest_df[config.fields.interface_bin] = retest_df[config.fields.interface_bin].fillna("UNKNOWN")
    retest_df[config.fields.devrevstep] = retest_df[config.fields.devrevstep].fillna("UNKNOWN")
    
    if not retest_df.empty:
        # 处理 process_step 列
        if config.fields.process_step in retest_df.columns:
            retest_df[config.fields.process_step] = retest_df[config.fields.process_step].fillna("UNKNOWN")
            # 按 process_step 分组
            process_steps = retest_df[config.fields.process_step].dropna().unique()
            logging.info(f"Retest 数据中找到 {len(process_steps)} 个 process_step 类型: {list(process_steps)}")
            
            for step in process_steps:
                step_df = retest_df[retest_df[config.fields.process_step] == step].copy()
                if not step_df.empty:
                    retest_quantity, _ = build_quantity_table(step_df, config, bin_type="functional")
                    if not retest_quantity.empty:
                        retest_percentages, retest_row_totals, retest_total_percentage = build_pareto(retest_quantity)
                        pareto_table = build_pareto_table(
                            retest_quantity, retest_percentages, retest_row_totals, retest_total_percentage
                        )
                        retest_pareto_by_step[str(step)] = pareto_table
                        logging.info(f"✅ {step} 的 Retest Bin Pareto 生成完成")
            
            if retest_pareto_by_step:
                # 如果按 process_step 分组成功，不生成总的 Pareto 表
                retest_pareto = None
                logging.info("已按 process_step 分组，不生成总的 Retest Bin Pareto 表")
            else:
                logging.warning("所有 process_step 的 Retest 数据都为空，跳过 Retest Bin Pareto")
        else:
            # 没有 process_step 列，生成总的 Retest Bin Pareto
            retest_quantity, _ = build_quantity_table(retest_df, config, bin_type="functional")
            if not retest_quantity.empty:
                retest_percentages, retest_row_totals, retest_total_percentage = build_pareto(retest_quantity)
                retest_pareto = build_pareto_table(
                    retest_quantity, retest_percentages, retest_row_totals, retest_total_percentage
                )
                logging.info("Retest Bin Pareto 生成完成")
            else:
                logging.warning("Retest 数据为空，跳过 Retest Bin Pareto")
    else:
        logging.warning("没有 Retest 数据（mut_within_subflow_latest_flag = N），跳过 Retest Bin Pareto")
    
    exceptions_df = collect_exceptions(normalized_df, config)
    
    report_path = save_report(
        report_dir=config.paths.report_dir,
        df=normalized_df,
        functional_pareto=functional_pareto,
        functional_pareto_by_step=functional_pareto_by_step if functional_pareto_by_step else None,
        interface_pareto=interface_pareto,
        interface_pareto_by_step=interface_pareto_by_step if interface_pareto_by_step else None,
        retest_pareto=retest_pareto,
        retest_pareto_by_step=retest_pareto_by_step if retest_pareto_by_step else None,
        exceptions=exceptions_df,
        config=config,
    )
    
    logging.info("✅ PRD LOT 自动化完成：%s", report_path)
    
    # 关闭 SQLPathFinder
    logging.info("关闭 SQLPathFinder...")
    close_sqlpathfinder(config.ui.main_window_title)
    
    return report_path


def run_lab_tp_automation(config_path: Path):
    """运行 Lab TP Performance 自动化，返回报告路径和数据"""
    logging.info("=" * 80)
    logging.info("开始执行 Lab TP Performance 自动化")
    logging.info("=" * 80)
    
    config = load_config(config_path)
    
    lots = read_lots(config.paths.lots_file)
    
    runner = SQLPathFinderRunner(config)
    runner.execute(lots)
    csv_path = runner.wait_for_output()
    
    # 处理数据
    from lab_tp_automation.lab_tp_main import process_lab_data, save_report as save_lab_report
    df = process_lab_data(csv_path, config)
    report_path = save_lab_report(config.paths.report_dir, df)
    
    logging.info("✅ Lab TP Performance 自动化完成：%s", report_path)
    
    # 关闭 SQLPathFinder
    logging.info("关闭 SQLPathFinder...")
    close_sqlpathfinder(config.ui.main_window_title)
    
    return report_path, df


def send_unified_email(config, prd_report: Path, lab_report: Path, lab_df, lots_count: int):
    """发送统一邮件，包含两个报告和 Lab TP 数据表格"""
    import pandas as pd
    
    # 准备邮件摘要
    summary = {
        "date": f"{datetime.now():%Y-%m-%d}",
        "lot_count": lots_count,
        "top_bins": [],
        "exception_count": 0,
        "grand_total": 0.0,
    }
    
    # 修改邮件主题，包含两个报告
    original_subject = config.email.subject_template
    config.email.subject_template = f"[Daily Report] Lab TP Performance & PRD LOT - {{date}}"
    
    # 生成 Lab TP Performance 的 HTML 表格
    lab_html_table = generate_lab_tp_html_table(lab_df)
    
    # 发送邮件，附带两个报告和 HTML 表格
    from prd_lot_automation.mailer import send_unified_report_email
    send_unified_report_email(config, summary, [lab_report, prd_report], lab_html_table)
    
    # 恢复原始主题模板
    config.email.subject_template = original_subject


def load_phi_data(phi_file: Path):
    """加载 PHI 数据"""
    import pandas as pd
    
    if not phi_file.exists():
        logging.warning(f"PHI 文件不存在: {phi_file}")
        return pd.DataFrame()
    
    try:
        # 尝试读取 Excel 文件
        phi_df = pd.read_excel(phi_file)
        logging.info(f"成功加载 PHI 数据: {phi_df.shape}")
        logging.info(f"PHI 数据列名: {phi_df.columns.tolist()}")
        
        # 处理合并单元格：填充 NaN 值（forward fill）
        # 对于 Operation 和 Sub Flow Step 列，如果上一行有值而当前行为 NaN，则使用上一行的值
        # 这是因为 Excel 中合并单元格在 pandas 读取时只有第一个单元格有值，其他为 NaN
        if 'Operation' in phi_df.columns:
            phi_df['Operation'] = phi_df['Operation'].ffill()  # 向前填充
        if 'Sub Flow Step' in phi_df.columns:
            phi_df['Sub Flow Step'] = phi_df['Sub Flow Step'].ffill()  # 向前填充
        
        logging.info(f"处理合并单元格后的 PHI 数据:")
        if not phi_df.empty:
            logging.info(f"PHI 数据前5行:")
            for idx, row in phi_df.head(5).iterrows():
                row_dict = {}
                for col in phi_df.columns:
                    row_dict[col] = row[col]
                logging.info(f"  行 {idx}: {row_dict}")
        
        return phi_df
    except Exception as e:
        logging.error(f"加载 PHI 数据失败: {e}")
        import traceback
        logging.error(traceback.format_exc())
        return pd.DataFrame()

def get_phi_value(phi_df, row) -> dict:
    """根据 Operation, Sub Flow Step, Devrevstep 精确匹配对应的 PHI 值"""
    import pandas as pd
    
    if phi_df.empty:
        return {'Yield PHI': None, 'TTG PHI': None, 'RCS PHI': None}
    
    phi_value = {'Yield PHI': None, 'TTG PHI': None, 'RCS PHI': None}
    
    try:
        # 获取当前行的匹配键值（转换为字符串并去除空格，确保精确匹配）
        op_value = row.get('Operation', None)
        if op_value is not None and not pd.isna(op_value):
            # 处理浮点数：6248.0 -> "6248"
            op_value = str(int(float(op_value))) if isinstance(op_value, (int, float)) else str(op_value).strip()
        else:
            op_value = None
            
        subflow_value = str(row.get('Sub Flow Step', '')).strip() if 'Sub Flow Step' in row.index and not pd.isna(row.get('Sub Flow Step', None)) else None
        devrev_value = str(row.get('Devrevstep', '')).strip() if 'Devrevstep' in row.index and not pd.isna(row.get('Devrevstep', None)) else None
        
        # 根据 Operation, Sub Flow Step, Devrevstep 精确匹配
        match_conditions = []
        
        if 'Operation' in phi_df.columns and op_value:
            # 处理 PHI 文件中的 Operation：可能是浮点数或字符串，转换为字符串比较
            # 6248.0 -> "6248", "6248" -> "6248"
            phi_op_series = phi_df['Operation'].apply(lambda x: str(int(float(x))) if pd.notna(x) and isinstance(x, (int, float)) else (str(x).strip() if pd.notna(x) else None))
            match_conditions.append(phi_op_series == op_value)
        if 'Sub Flow Step' in phi_df.columns and subflow_value:
            # 处理 NaN：只匹配非 NaN 的行
            phi_subflow_series = phi_df['Sub Flow Step'].apply(lambda x: str(x).strip() if pd.notna(x) else None)
            match_conditions.append(phi_subflow_series == subflow_value)
        if 'Devrevstep' in phi_df.columns and devrev_value:
            # 关键：确保 Devrevstep 精确匹配，避免 4PXA2VDB 匹配到 4PXA4VDB
            phi_devrev_series = phi_df['Devrevstep'].apply(lambda x: str(x).strip() if pd.notna(x) else None)
            match_conditions.append(phi_devrev_series == devrev_value)
        
        if match_conditions:
            # 组合所有条件（必须全部匹配）
            combined_condition = match_conditions[0]
            for condition in match_conditions[1:]:
                combined_condition = combined_condition & condition
            
            match = phi_df[combined_condition]
            
            # 如果精确匹配失败，且 Operation 或 Sub Flow Step 为 None，尝试只匹配 Devrevstep
            if match.empty and devrev_value:
                # 只匹配 Devrevstep（当 Operation 或 Sub Flow Step 为空时）
                if 'Devrevstep' in phi_df.columns:
                    devrev_only_match = phi_df[phi_df['Devrevstep'].apply(lambda x: str(x).strip() if pd.notna(x) else None) == devrev_value]
                    if not devrev_only_match.empty:
                        # 检查是否 Operation 或 Sub Flow Step 为 NaN（允许匹配）
                        valid_rows = []
                        for idx, phi_row in devrev_only_match.iterrows():
                            phi_op = phi_row.get('Operation', None)
                            phi_subflow = phi_row.get('Sub Flow Step', None)
                            # 如果 PHI 文件中的 Operation 或 Sub Flow Step 是 NaN，且我们的值也是 None，则匹配
                            # 或者如果 PHI 文件中的值与我们匹配
                            op_match = (pd.isna(phi_op) and op_value is None) or (pd.notna(phi_op) and op_value and str(int(float(phi_op))) == op_value)
                            subflow_match = (pd.isna(phi_subflow) and subflow_value is None) or (pd.notna(phi_subflow) and subflow_value and str(phi_subflow).strip() == subflow_value)
                            
                            if op_match and subflow_match:
                                valid_rows.append(idx)
                        
                        if valid_rows:
                            match = devrev_only_match.loc[valid_rows]
            if not match.empty:
                # 获取匹配的第一行
                matched_row = match.iloc[0]
                
                logging.info(f"✅ 匹配到 PHI: Operation={op_value}, Sub Flow Step={subflow_value}, Devrevstep={devrev_value}")
                logging.info(f"匹配到的行数据: {dict(matched_row)}")
                
                # 查找 PHI 列（支持多种列名格式，但优先精确匹配）
                for col in phi_df.columns:
                    col_str = str(col).strip()
                    col_lower = col_str.lower()
                    
                    # 匹配 Yield PHI（优先精确匹配列名）
                    if phi_value['Yield PHI'] is None:
                        if col_str == 'Yield PHI':
                            val = matched_row[col]
                            # 如果是小数（如 0.753），转换为百分比（75.30%）
                            if pd.notna(val):
                                try:
                                    val_float = float(val)
                                    if 0 <= val_float <= 1:  # 如果是小数格式
                                        phi_value['Yield PHI'] = f"{val_float * 100:.2f}%"
                                    else:
                                        phi_value['Yield PHI'] = f"{val_float:.2f}%"
                                except:
                                    phi_value['Yield PHI'] = str(val) if '%' in str(val) else f"{val}%"
                            logging.info(f"  找到 Yield PHI 列: {col_str}, 原始值: {matched_row[col]}, 转换后: {phi_value['Yield PHI']}")
                        elif 'yield' in col_lower and 'phi' in col_lower and 'rcs' not in col_lower:
                            val = matched_row[col]
                            if pd.notna(val):
                                try:
                                    val_float = float(val)
                                    if 0 <= val_float <= 1:
                                        phi_value['Yield PHI'] = f"{val_float * 100:.2f}%"
                                    else:
                                        phi_value['Yield PHI'] = f"{val_float:.2f}%"
                                except:
                                    phi_value['Yield PHI'] = str(val) if '%' in str(val) else f"{val}%"
                            logging.info(f"  找到 Yield PHI 列（模糊匹配）: {col_str}, 原始值: {matched_row[col]}, 转换后: {phi_value['Yield PHI']}")
                    
                    # 匹配 TTG PHI
                    if phi_value['TTG PHI'] is None:
                        if col_str == 'TTG PHI':
                            val = matched_row[col]
                            phi_value['TTG PHI'] = val if pd.notna(val) else None
                            logging.info(f"  找到 TTG PHI 列: {col_str}, 值: {matched_row[col]}")
                        elif 'ttg' in col_lower and 'phi' in col_lower:
                            val = matched_row[col]
                            phi_value['TTG PHI'] = val if pd.notna(val) else None
                            logging.info(f"  找到 TTG PHI 列（模糊匹配）: {col_str}, 值: {matched_row[col]}")
                    
                    # 匹配 RCS PHI（优先精确匹配列名）
                    if phi_value['RCS PHI'] is None:
                        if col_str == 'RCS PHI':
                            val = matched_row[col]
                            # 如果是小数（如 0.2000），转换为百分比（20.00%）
                            if pd.notna(val):
                                try:
                                    val_float = float(val)
                                    if 0 <= val_float <= 1:  # 如果是小数格式
                                        phi_value['RCS PHI'] = f"{val_float * 100:.2f}%"
                                    else:
                                        phi_value['RCS PHI'] = f"{val_float:.2f}%"
                                except:
                                    phi_value['RCS PHI'] = str(val) if '%' in str(val) else f"{val}%"
                            logging.info(f"  找到 RCS PHI 列: {col_str}, 原始值: {matched_row[col]}, 转换后: {phi_value['RCS PHI']}")
                        elif 'rcs' in col_lower and 'phi' in col_lower and 'yield' not in col_lower:
                            val = matched_row[col]
                            if pd.notna(val):
                                try:
                                    val_float = float(val)
                                    if 0 <= val_float <= 1:
                                        phi_value['RCS PHI'] = f"{val_float * 100:.2f}%"
                                    else:
                                        phi_value['RCS PHI'] = f"{val_float:.2f}%"
                                except:
                                    phi_value['RCS PHI'] = str(val) if '%' in str(val) else f"{val}%"
                            logging.info(f"  找到 RCS PHI 列（模糊匹配）: {col_str}, 原始值: {matched_row[col]}, 转换后: {phi_value['RCS PHI']}")
                
                logging.info(f"最终提取的 PHI 值: Yield PHI={phi_value['Yield PHI']}, TTG PHI={phi_value['TTG PHI']}, RCS PHI={phi_value['RCS PHI']}")
            else:
                logging.warning(f"⚠️ 未找到匹配的 PHI 值: Operation={op_value}, Sub Flow Step={subflow_value}, Devrevstep={devrev_value}")
                # 显示 PHI 文件中的所有唯一值用于调试
                if 'Operation' in phi_df.columns:
                    logging.info(f"  PHI 文件中的 Operation 值: {phi_df['Operation'].unique().tolist()}")
                if 'Sub Flow Step' in phi_df.columns:
                    logging.info(f"  PHI 文件中的 Sub Flow Step 值: {phi_df['Sub Flow Step'].unique().tolist()}")
                if 'Devrevstep' in phi_df.columns:
                    logging.info(f"  PHI 文件中的 Devrevstep 值: {phi_df['Devrevstep'].unique().tolist()}")
        else:
            logging.warning(f"⚠️ 缺少匹配条件: Operation={op_value}, Sub Flow Step={subflow_value}, Devrevstep={devrev_value}")
            logging.info(f"  PHI 文件中的列: {phi_df.columns.tolist()}")
            
    except Exception as e:
        logging.warning(f"获取 PHI 值失败: {e}")
        import traceback
        logging.debug(traceback.format_exc())
    
    return phi_value

def generate_lab_tp_html_table(df) -> str:
    """生成 Lab TP Performance 的 HTML 表格，合并相同值并按指定顺序排列 Sub Flow Step
    
    Sub Flow Step 排序顺序：CLASSHOT, CLASSCOLD, PHMHOT, PHMCOLD, 其他的
    """
    import pandas as pd
    
    # 加载 PHI 数据
    phi_file = Path(__file__).parent / "Prod_PHI.xlsx"
    logging.info(f"尝试加载 PHI 文件: {phi_file}")
    phi_df = load_phi_data(phi_file)
    
    if phi_df.empty:
        logging.warning("⚠️ PHI 数据为空，将无法应用条件格式")
    else:
        logging.info(f"✅ PHI 数据加载成功，共 {len(phi_df)} 行")
    
    # 检查并修正列名（容错处理）
    # LOT 是可选的，如果存在则添加到合并列中（在 Program Name 之后）
    expected_merge_cols = ['Facility', 'Operation', 'Sub Flow Step', 'Devrevstep', 'Program Name']
    expected_data_cols = ['Total Tested', 'Tested Good', 'Yield', 'TTG', 'ETT', 'RCS', 'Recovery Rate']
    
    # 获取实际存在的列名
    actual_cols = df.columns.tolist()
    merge_cols = [col for col in expected_merge_cols if col in actual_cols]
    data_cols = [col for col in expected_data_cols if col in actual_cols]
    
    # 如果存在 LOT 列，将其添加到 merge_cols 中（在 Program Name 之后）
    if 'LOT' in actual_cols:
        # 找到 Program Name 的位置，在其后插入 LOT
        if 'Program Name' in merge_cols:
            program_name_idx = merge_cols.index('Program Name')
            merge_cols.insert(program_name_idx + 1, 'LOT')
        else:
            # 如果 Program Name 不存在，添加到末尾
            merge_cols.append('LOT')
    
    # 检查必需的列是否存在
    missing_merge = set(expected_merge_cols) - set(merge_cols)
    missing_data = set(expected_data_cols) - set(data_cols)
    if missing_merge or missing_data:
        logging.warning(f"缺少列: merge={missing_merge}, data={missing_data}")
        logging.info(f"实际列名: {actual_cols}")
    
    # 1. 按 Sub Flow Step 自定义排序：CLASSHOT, CLASSCOLD, PHMHOT, PHMCOLD, 其他的
    if 'Sub Flow Step' in df.columns:
        # 定义排序顺序
        sort_order = ['CLASSHOT', 'CLASSCOLD', 'PHMHOT', 'PHMCOLD']
        
        # 创建排序键：如果在预定义顺序中，使用索引；否则使用很大的数字（排在最后）
        def get_sort_key(value):
            if pd.isna(value):
                return (999, '')  # NaN 值排在最后
            value_str = str(value).strip().upper()
            try:
                index = sort_order.index(value_str)
                return (index, value_str)
            except ValueError:
                # 不在预定义顺序中，排在最后，但保持原有顺序
                return (len(sort_order), value_str)
        
        # 创建临时排序列
        df['_sort_key'] = df['Sub Flow Step'].apply(get_sort_key)
        df_sorted = df.sort_values('_sort_key').reset_index(drop=True)
        df_sorted = df_sorted.drop(columns=['_sort_key'])
    else:
        df_sorted = df.copy()
    
    # 为每列计算 rowspan：计算连续相同值的数量
    def calculate_rowspans_for_column(df, col):
        """计算每列需要合并的行数"""
        rowspans = []
        i = 0
        while i < len(df):
            value = df.iloc[i][col]
            count = 1
            # 计算连续相同值的数量
            for j in range(i + 1, len(df)):
                if str(df.iloc[j][col]) == str(value):  # 使用 str 比较避免类型问题
                    count += 1
                else:
                    break
            rowspans.append((i, count, value))
            i += count
        return rowspans
    
    # 为每列计算 rowspan
    rowspans_dict = {}
    for col in merge_cols:
        rowspans_dict[col] = calculate_rowspans_for_column(df_sorted, col)
    
    # 创建 HTML 表格样式（与截图完全一致）
    html = """
    <html>
    <head>
        <style>
            body { 
                font-family: Calibri, Arial, sans-serif; 
                margin: 0; 
                padding: 0; 
            }
            h2 { 
                color: #2E75B6; 
                margin-top: 10px; 
                margin-bottom: 5px; 
                font-size: 14pt;
                font-weight: normal;
            }
            table { 
                border-collapse: collapse; 
                width: 100%; 
                margin-top: 5px;
                font-size: 11pt;
                line-height: 1.0;
            }
            th { 
                background-color: #92D050; 
                color: #000000; 
                padding: 2px 4px; 
                text-align: center;
                border: 1px solid #000000;
                font-weight: bold;
                line-height: 1.2;
            }
            td { 
                padding: 2px 4px; 
                text-align: center;
                border: 1px solid #000000;
                vertical-align: middle;
                line-height: 1.2;
                color: #000000;
            }
            td.merged { 
                text-align: center;
                vertical-align: middle;
                line-height: 1.2;
                color: #000000;
            }
            tr { 
                line-height: 1.0;
            }
            /* 数据行配色：奇数行白色，偶数行浅绿色（与截图完全一致） */
            tbody tr:nth-child(odd) td { 
                background-color: #FFFFFF; 
            }
            tbody tr:nth-child(even) td { 
                background-color: #E6F2D9; 
            }
            /* 确保合并单元格也应用背景色 */
            tbody tr:nth-child(odd) td.merged { 
                background-color: #FFFFFF; 
            }
            tbody tr:nth-child(even) td.merged { 
                background-color: #E6F2D9; 
            }
        </style>
    </head>
    <body>
        <h2>Lot Level Performance</h2>
        <table>
            <thead>
                <tr>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Facility</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Operation</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Sub Flow Step</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Devrevstep</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Program Name</th>"""
    # 如果存在 LOT 列，添加 LOT 表头
    if 'LOT' in df.columns:
        html += """
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">LOT</th>"""
    html += """
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Total Tested</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Tested Good</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Yield</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Yield PHI</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">TTG</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">TTG PHI</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">ETT</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">RCS</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">RCS PHI</th>
                    <th style="background-color: #92D050; color: #000000; padding: 2px 4px; text-align: center; border: 1px solid #000000; font-weight: bold;">Recovery Rate</th>
                </tr>
            </thead>
            <tbody>
    """
    
    # 跟踪每列当前应该显示的行索引
    merge_positions = {col: 0 for col in merge_cols}
    
    # 添加数据行（直接在td标签上添加背景色样式，确保邮件客户端支持）
    for row_idx in range(len(df_sorted)):
        row = df_sorted.iloc[row_idx]
        # 确定行背景色：奇数行白色，偶数行浅绿色
        bg_color = "#FFFFFF" if (row_idx % 2 == 0) else "#E6F2D9"
        html += "                <tr>\n"
        
        # 处理需要合并的列
        for col in merge_cols:
            rowspans = rowspans_dict[col]
            
            # 检查是否应该在这个位置显示合并单元格
            if merge_positions[col] < len(rowspans):
                start_idx, span, value = rowspans[merge_positions[col]]
                if row_idx == start_idx:
                    # 这是合并单元格的起始位置
                    display_value = str(value) if not pd.isna(value) else ""
                    style_attr = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000;"'
                    if span > 1:
                        html += f'                    <td class="merged" rowspan="{span}" {style_attr}>{display_value}</td>\n'
                    else:
                        html += f'                    <td class="merged" {style_attr}>{display_value}</td>\n'
                    merge_positions[col] += 1
                # 否则跳过（因为已经被合并了，不输出这个单元格）
        
        # 获取 PHI 值（在处理数据列之前）
        phi_values = get_phi_value(phi_df, row)
        
        # 调试：记录第一行的 PHI 值
        if row_idx == 0:
            logging.info(f"第一行 PHI 值: {phi_values}")
            logging.info(f"第一行数据: Operation={row.get('Operation')}, Sub Flow Step={row.get('Sub Flow Step')}, Devrevstep={row.get('Devrevstep')}")
        
        # 处理数据列（不需要合并）
        # 需要特殊处理 Yield、TTG、RCS，在它们后面添加对应的 PHI 列
        for col in data_cols:
            if col not in row.index:
                style_attr = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000;"'
                html += f'                    <td {style_attr}></td>\n'
                # 如果是 Yield、TTG、RCS，还需要添加对应的 PHI 列
                if col == 'Yield':
                    phi_style = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000; font-weight: bold;"'
                    phi_display = f"{phi_values.get('Yield PHI', '')}" if phi_values.get('Yield PHI') is not None else ""
                    html += f'                    <td {phi_style}>{phi_display}</td>\n'
                elif col == 'TTG':
                    phi_style = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000; font-weight: bold;"'
                    phi_display = f"{phi_values.get('TTG PHI', '')}" if phi_values.get('TTG PHI') is not None else ""
                    html += f'                    <td {phi_style}>{phi_display}</td>\n'
                elif col == 'RCS':
                    phi_style = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000; font-weight: bold;"'
                    phi_display = f"{phi_values.get('RCS PHI', '')}" if phi_values.get('RCS PHI') is not None else ""
                    html += f'                    <td {phi_style}>{phi_display}</td>\n'
                continue
            
            value = row[col]
            
            # 特殊处理 Yield、TTG、RCS、ETT：应用条件格式
            if col == 'Yield':
                yield_phi = phi_values.get('Yield PHI', None)
                # 格式化 Yield 值
                if pd.isna(value):
                    display_value = ""
                    cell_bg_color = bg_color
                    cell_text_color = "#000000"
                else:
                    try:
                        yield_num = abs(float(str(value).replace('%', '')))  # 使用绝对值
                        display_value = f"{yield_num:.2f}%"
                        # 应用条件格式：绝对值 > Yield PHI → 绿色，< Yield PHI → 红色
                        if yield_phi is not None:
                            try:
                                yield_phi_num = abs(float(str(yield_phi).replace('%', '')))  # PHI 也取绝对值
                                if yield_num > yield_phi_num:
                                    cell_bg_color = "#00B050"  # 绿色
                                    cell_text_color = "#FFFFFF"  # 白色
                                else:
                                    cell_bg_color = "#FF0000"  # 红色
                                    cell_text_color = "#FFFFFF"  # 白色
                            except:
                                cell_bg_color = bg_color
                                cell_text_color = "#000000"
                        else:
                            cell_bg_color = bg_color
                            cell_text_color = "#000000"
                    except:
                        display_value = str(value)
                        cell_bg_color = bg_color
                        cell_text_color = "#000000"
                
                style_attr = f'style="background-color: {cell_bg_color}; color: {cell_text_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000;"'
                html += f'                    <td {style_attr}>{display_value}</td>\n'
                
                # 添加 Yield PHI 列（保留百分号，不应用颜色）
                phi_style = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000; font-weight: bold;"'
                if yield_phi is not None:
                    try:
                        # 确保保留百分号
                        phi_str = str(yield_phi).replace('%', '')
                        phi_num = float(phi_str)
                        phi_display = f"{phi_num:.2f}%"
                    except:
                        phi_display = str(yield_phi)
                else:
                    phi_display = ""
                html += f'                    <td {phi_style}>{phi_display}</td>\n'
                
            elif col == 'TTG':
                ttg_phi = phi_values.get('TTG PHI', None)
                # 格式化 TTG 值
                if pd.isna(value):
                    display_value = ""
                    cell_bg_color = bg_color
                    cell_text_color = "#000000"
                else:
                    try:
                        ttg_num = abs(float(value))  # 使用绝对值
                        display_value = f"{ttg_num:.2f}"
                        # 应用条件格式：绝对值 > TTG PHI → 红色，< TTG PHI → 绿色
                        if ttg_phi is not None:
                            try:
                                ttg_phi_num = abs(float(ttg_phi))  # PHI 也取绝对值
                                if ttg_num > ttg_phi_num:
                                    cell_bg_color = "#FF0000"  # 红色
                                    cell_text_color = "#FFFFFF"  # 白色
                                else:
                                    cell_bg_color = "#00B050"  # 绿色
                                    cell_text_color = "#FFFFFF"  # 白色
                            except:
                                cell_bg_color = bg_color
                                cell_text_color = "#000000"
                        else:
                            cell_bg_color = bg_color
                            cell_text_color = "#000000"
                    except:
                        display_value = str(value)
                        cell_bg_color = bg_color
                        cell_text_color = "#000000"
                
                style_attr = f'style="background-color: {cell_bg_color}; color: {cell_text_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000;"'
                html += f'                    <td {style_attr}>{display_value}</td>\n'
                
                # 添加 TTG PHI 列（只显示值，不应用颜色）
                phi_style = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000; font-weight: bold;"'
                if ttg_phi is not None:
                    try:
                        phi_num = float(ttg_phi)
                        phi_display = f"{phi_num:.2f}"
                    except:
                        phi_display = str(ttg_phi)
                else:
                    phi_display = ""
                html += f'                    <td {phi_style}>{phi_display}</td>\n'
                
            elif col == 'ETT':
                # ETT 使用 TTG PHI 进行比较
                ttg_phi = phi_values.get('TTG PHI', None)
                # 格式化 ETT 值
                if pd.isna(value):
                    display_value = ""
                    cell_bg_color = bg_color
                    cell_text_color = "#000000"
                else:
                    try:
                        ett_num = abs(float(value))  # 使用绝对值
                        display_value = f"{ett_num:.2f}"
                        # 应用条件格式：绝对值 > TTG PHI → 红色，< TTG PHI → 绿色
                        if ttg_phi is not None:
                            try:
                                ttg_phi_num = abs(float(ttg_phi))  # PHI 也取绝对值
                                if ett_num > ttg_phi_num:
                                    cell_bg_color = "#FF0000"  # 红色
                                    cell_text_color = "#FFFFFF"  # 白色
                                else:
                                    cell_bg_color = "#00B050"  # 绿色
                                    cell_text_color = "#FFFFFF"  # 白色
                            except:
                                cell_bg_color = bg_color
                                cell_text_color = "#000000"
                        else:
                            cell_bg_color = bg_color
                            cell_text_color = "#000000"
                    except:
                        display_value = str(value)
                        cell_bg_color = bg_color
                        cell_text_color = "#000000"
                
                style_attr = f'style="background-color: {cell_bg_color}; color: {cell_text_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000;"'
                html += f'                    <td {style_attr}>{display_value}</td>\n'
                
            elif col == 'RCS':
                rcs_phi = phi_values.get('RCS PHI', None)
                # 格式化 RCS 值
                if pd.isna(value):
                    display_value = ""
                    cell_bg_color = bg_color
                    cell_text_color = "#000000"
                else:
                    try:
                        rcs_num = abs(float(str(value).replace('%', '')))  # 使用绝对值
                        display_value = f"{rcs_num:.2f}%"
                        # 应用条件格式：绝对值 > RCS PHI → 红色，< RCS PHI → 绿色
                        if rcs_phi is not None:
                            try:
                                rcs_phi_num = abs(float(str(rcs_phi).replace('%', '')))  # PHI 也取绝对值
                                if rcs_num > rcs_phi_num:
                                    cell_bg_color = "#FF0000"  # 红色
                                    cell_text_color = "#FFFFFF"  # 白色
                                else:
                                    cell_bg_color = "#00B050"  # 绿色
                                    cell_text_color = "#FFFFFF"  # 白色
                            except:
                                cell_bg_color = bg_color
                                cell_text_color = "#000000"
                        else:
                            cell_bg_color = bg_color
                            cell_text_color = "#000000"
                    except:
                        display_value = str(value)
                        cell_bg_color = bg_color
                        cell_text_color = "#000000"
                
                style_attr = f'style="background-color: {cell_bg_color}; color: {cell_text_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000;"'
                html += f'                    <td {style_attr}>{display_value}</td>\n'
                
                # 添加 RCS PHI 列（保留百分号，不应用颜色）
                phi_style = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000; font-weight: bold;"'
                if rcs_phi is not None:
                    try:
                        # 确保保留百分号
                        phi_str = str(rcs_phi).replace('%', '')
                        phi_num = float(phi_str)
                        phi_display = f"{phi_num:.2f}%"
                    except:
                        phi_display = str(rcs_phi)
                else:
                    phi_display = ""
                html += f'                    <td {phi_style}>{phi_display}</td>\n'
                
            else:
                # 其他列正常处理（Total Tested, Tested Good, ETT, Recovery Rate）
                if pd.isna(value):
                    display_value = ""
                elif isinstance(value, (int, float)):
                    if col == 'Recovery Rate':
                        display_value = f"{value:.2f}%" if not pd.isna(value) else ""
                    elif col == 'ETT':
                        display_value = f"{value:.2f}" if not pd.isna(value) else ""
                    else:
                        display_value = str(int(value)) if value == int(value) else f"{value:.2f}"
                else:
                    display_value = str(value)
                style_attr = f'style="background-color: {bg_color}; text-align: center; vertical-align: middle; padding: 2px 4px; border: 1px solid #000000;"'
                html += f'                    <td {style_attr}>{display_value}</td>\n'
        
        html += "                </tr>\n"
    
    html += """
            </tbody>
        </table>
    </body>
    </html>
    """
    
    return html


def main():
    """主程序：依次运行两个自动化，最后发送统一邮件"""
    base_dir = Path(__file__).parent
    
    # 配置日志
    log_dir = base_dir / "logs"
    configure_logging(log_dir)
    
    logging.info("=" * 80)
    logging.info("统一自动化开始")
    logging.info("=" * 80)
    
    try:
        # 1. 先运行 Lab TP Performance 自动化
        lab_config_path = base_dir / "lab_tp_automation" / "config.yaml"
        lab_report, lab_df = run_lab_tp_automation(lab_config_path)
        
        # 在执行下一个 VG2 之前，先关闭所有 SQLPathFinder 窗口
        logging.info("=" * 80)
        logging.info("准备执行下一个 VG2，先关闭所有 SQLPathFinder 窗口")
        logging.info("=" * 80)
        close_sqlpathfinder("SQLPathFinder")
        
        # 等待一下，确保 SQLPathFinder 完全关闭
        import time
        time.sleep(2)
        
        # 2. 运行 PRD LOT 自动化
        prd_config_path = base_dir / "prd_lot_automation" / "config.yaml"
        prd_report = run_prd_lot_automation(prd_config_path)
        
        # 3. 发送统一邮件
        logging.info("=" * 80)
        logging.info("发送统一邮件")
        logging.info("=" * 80)
        
        # 读取 LOT 数量
        prd_config = load_config(prd_config_path)
        lots = read_lots(prd_config.paths.lots_file)
        
        send_unified_email(prd_config, prd_report, lab_report, lab_df, len(lots))
        
        logging.info("=" * 80)
        logging.info("✅ 统一自动化完成")
        logging.info(f"  Lab TP 报告: {lab_report}")
        logging.info(f"  PRD LOT 报告: {prd_report}")
        logging.info("=" * 80)
        
    except Exception as e:
        logging.error(f"❌ 统一自动化失败: {e}")
        import traceback
        logging.error(traceback.format_exc())
        raise


if __name__ == "__main__":
    main()

