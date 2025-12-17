import argparse
import logging
from datetime import datetime
from pathlib import Path

from .config_loader import load_config
from .lot_reader import read_lots, split_batches
from .spf_runner import SQLPathFinderRunner
from .data_processing import (
    load_dataset,
    normalize_columns,
    apply_filters,
    build_quantity_table,
    build_pareto,
    collect_exceptions,
)
from .report_builder import build_pareto_table, save_report
from .mailer import send_report_email
from .close_sqlpathfinder import close_sqlpathfinder


def configure_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "prd_lot_automation.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding="utf-8"),
        ],
    )


def run_pipeline(config_path: Path) -> Path:
    config = load_config(config_path)
    configure_logging(config.paths.log_dir)
    logging.info("PRD LOT 自动化开始")

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
        LOGGER.info(f"找到 {len(process_steps)} 个 process_step 类型: {list(process_steps)}")
        
        # 为每个 process_step 生成 Pareto 表
        for step in process_steps:
            step_df = normalized_df[normalized_df[process_step_col] == step].copy()
            if not step_df.empty:
                LOGGER.info(f"处理 process_step: {step}，数据行数: {len(step_df)}")
                quantity, _ = build_quantity_table(step_df, config, bin_type="functional")
                if not quantity.empty:
                    percentages, row_totals, total_percentage = build_pareto(quantity)
                    pareto_table = build_pareto_table(quantity, percentages, row_totals, total_percentage)
                    functional_pareto_by_step[str(step)] = pareto_table
                    LOGGER.info(f"✅ {step} 的 Functional_bin Pareto 生成完成")
                else:
                    LOGGER.warning(f"⚠️ {step} 的数量表为空，跳过")
            else:
                LOGGER.warning(f"⚠️ {step} 的数据为空，跳过")
        
        # 如果没有 process_step 数据，生成一个总的 Pareto 表
        if not functional_pareto_by_step:
            LOGGER.warning("所有 process_step 的数据都为空，生成总的 Functional_bin Pareto")
            quantity, _ = build_quantity_table(normalized_df, config, bin_type="functional")
            percentages, row_totals, total_percentage = build_pareto(quantity)
            functional_pareto = build_pareto_table(quantity, percentages, row_totals, total_percentage)
        else:
            # 如果按 process_step 分组成功，不生成总的 Pareto 表
            functional_pareto = None
            LOGGER.info("已按 process_step 分组，不生成总的 Functional_bin Pareto 表")
    else:
        LOGGER.warning(f"⚠️ 未找到 process_step 列（{process_step_col}），生成总的 Functional_bin Pareto")
        # Functional_bin Pareto（使用配置中的过滤条件，通常是 mut_within_subflow_latest_flag = Y）
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
                        LOGGER.info(f"✅ {step} 的 Interface_bin Pareto 生成完成")
            
            if interface_pareto_by_step:
                # 如果按 process_step 分组成功，不生成总的 Pareto 表
                interface_pareto = None
                LOGGER.info("已按 process_step 分组，不生成总的 Interface_bin Pareto 表")
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
    # 重新从原始数据开始，应用 retest 过滤条件
    retest_filters = {
        "mut_within_subflow_latest_flag": "N",
        "SUBSTRUCTURE_ID": "U1.U2"
    }
    retest_df = apply_filters(merged_df, retest_filters)
    # 标准化列名（填充缺失值）
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
            LOGGER.info(f"Retest 数据中找到 {len(process_steps)} 个 process_step 类型: {list(process_steps)}")
            
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
                        LOGGER.info(f"✅ {step} 的 Retest Bin Pareto 生成完成")
            
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

    # 准备邮件摘要（如果有按 process_step 分组的表，使用第一个；否则使用总的表）
    summary_pareto = None
    if functional_pareto_by_step:
        summary_pareto = list(functional_pareto_by_step.values())[0]
    elif functional_pareto is not None:
        summary_pareto = functional_pareto
    
    summary = {
        "date": f"{datetime.now():%Y-%m-%d}",
        "lot_count": len(lots),
        "top_bins": summary_pareto.index[:3].tolist() if summary_pareto is not None and not summary_pareto.empty else [],
        "exception_count": len(exceptions_df),
        "grand_total": float(summary_pareto[('Grand Total', 'Quantity')].sum()) if summary_pareto is not None and not summary_pareto.empty else 0.0,
    }
    send_report_email(config, summary, [report_path])
    logging.info("流程完成：%s", report_path)
    
    # 关闭 SQLPathFinder
    logging.info("关闭 SQLPathFinder...")
    close_sqlpathfinder(config.ui.main_window_title)
    
    return report_path


def main() -> None:
    parser = argparse.ArgumentParser(description="PRD LOT 自动抓数与报表自动化工具")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path(__file__).parent / "config.yaml",
        help="配置文件路径",
    )
    args = parser.parse_args()
    run_pipeline(args.config)


if __name__ == "__main__":
    main()

