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
    
    # Functional_bin Pareto（使用配置中的过滤条件，通常是 mut_within_subflow_latest_flag = Y）
    quantity, _ = build_quantity_table(normalized_df, config, bin_type="functional")
    percentages, row_totals, total_percentage = build_pareto(quantity)
    functional_pareto = build_pareto_table(quantity, percentages, row_totals, total_percentage)
    
    # Interface_bin Pareto（如果存在）
    interface_pareto = None
    if config.fields.interface_bin in normalized_df.columns:
        interface_quantity, _ = build_quantity_table(normalized_df, config, bin_type="interface")
        if not interface_quantity.empty:
            interface_percentages, interface_row_totals, interface_total_percentage = build_pareto(interface_quantity)
            interface_pareto = build_pareto_table(
                interface_quantity, interface_percentages, interface_row_totals, interface_total_percentage
            )
    
    # Retest Bin Pareto（使用不同的过滤条件：mut_within_subflow_latest_flag = N）
    retest_pareto = None
    logging.info("开始生成 Retest Bin Pareto（mut_within_subflow_latest_flag = N）")
    # 重新从原始数据开始，应用 retest 过滤条件
    retest_filters = {
        "mut_within_subflow_latest_flag": "N",
        "SUBSTRUCTURE_ID": "U1.U1"
    }
    retest_df = apply_filters(merged_df, retest_filters)
    # 标准化列名（填充缺失值）
    retest_df[config.fields.functional_bin] = retest_df[config.fields.functional_bin].fillna("UNKNOWN")
    if config.fields.interface_bin in retest_df.columns:
        retest_df[config.fields.interface_bin] = retest_df[config.fields.interface_bin].fillna("UNKNOWN")
    retest_df[config.fields.devrevstep] = retest_df[config.fields.devrevstep].fillna("UNKNOWN")
    
    if not retest_df.empty:
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
        interface_pareto=interface_pareto,
        retest_pareto=retest_pareto,
        exceptions=exceptions_df,
        config=config,
    )

    summary = {
        "date": f"{datetime.now():%Y-%m-%d}",
        "lot_count": len(lots),
        "top_bins": functional_pareto.index[:3].tolist() if not functional_pareto.empty else [],
        "exception_count": len(exceptions_df),
        "grand_total": float(functional_pareto[('Grand Total', 'Quantity')].sum()) if not functional_pareto.empty else 0.0,
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

