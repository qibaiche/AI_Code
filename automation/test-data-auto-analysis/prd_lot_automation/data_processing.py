import logging
from pathlib import Path
from typing import Tuple

import pandas as pd

from .config_loader import AppConfig


LOGGER = logging.getLogger(__name__)


def load_dataset(csv_path: Path) -> pd.DataFrame:
    LOGGER.info("读取抓取结果：%s", csv_path)
    df = pd.read_csv(csv_path)
    if df.empty:
        raise ValueError(f"{csv_path} 为空")
    return df


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    """
    应用自定义过滤条件
    
    Args:
        df: 数据框
        filters: 过滤条件字典，例如 {"mut_within_subflow_latest_flag": "Y", "SUBSTRUCTURE_ID": "U1.U1"}
    
    Returns:
        过滤后的数据框
    """
    filtered_df = df.copy()
    
    if filters:
        initial_count = len(filtered_df)
        for filter_col, filter_value in filters.items():
            if filter_col in filtered_df.columns:
                before_count = len(filtered_df)
                filtered_df = filtered_df[filtered_df[filter_col] == filter_value]
                after_count = len(filtered_df)
                LOGGER.info("应用过滤条件: %s = %s，从 %s 行过滤到 %s 行", 
                           filter_col, filter_value, before_count, after_count)
            else:
                LOGGER.warning("过滤列 %s 不存在于数据中，跳过此过滤", filter_col)
        LOGGER.info("过滤完成：从原始 %s 行过滤到 %s 行", initial_count, len(filtered_df))
    
    return filtered_df


def normalize_columns(df: pd.DataFrame, config: AppConfig) -> pd.DataFrame:
    required = [
        config.fields.lot,
        config.fields.functional_bin,
        config.fields.devrevstep,
        config.fields.visual_id,
    ]
    missing = [col for col in required if col not in df.columns]
    if missing:
        # 提供详细的错误信息，包括实际存在的列名
        error_msg = f"缺少列：{missing}\n"
        error_msg += f"实际存在的列：{list(df.columns)}\n"
        error_msg += f"CSV 文件共有 {len(df.columns)} 列，{len(df)} 行\n"
        error_msg += "\n"
        error_msg += "可能的原因：\n"
        error_msg += "1. SQLPathFinder 查询输出的列名与配置不匹配\n"
        error_msg += "2. 列名可能有大小写、下划线或空格差异\n"
        error_msg += "3. 请运行 '检查CSV列名.py' 查看实际的列名\n"
        error_msg += "\n"
        error_msg += "解决方案：\n"
        error_msg += "1. 检查 config.yaml 中的 fields 配置\n"
        error_msg += "2. 根据实际列名更新 config.yaml\n"
        raise KeyError(error_msg)

    normalized = df.copy()
    
    # 应用配置中的过滤条件
    filters = config.processing.filters
    if filters:
        normalized = apply_filters(normalized, filters)
    
    normalized[config.fields.functional_bin] = normalized[
        config.fields.functional_bin
    ].fillna("UNKNOWN")
    if config.fields.interface_bin in normalized.columns:
        normalized[config.fields.interface_bin] = normalized[
            config.fields.interface_bin
        ].fillna("UNKNOWN")
    normalized[config.fields.devrevstep] = normalized[
        config.fields.devrevstep
    ].fillna("UNKNOWN")
    return normalized


def build_quantity_table(df: pd.DataFrame, config: AppConfig, bin_type: str = "functional") -> Tuple[pd.DataFrame, bool]:
    """
    构建数量表
    
    Args:
        df: 数据框
        config: 配置
        bin_type: "functional" 或 "interface"
    
    Returns:
        (数量表, 是否为数值)
    """
    if bin_type == "functional":
        bin_col = config.fields.functional_bin
    elif bin_type == "interface":
        bin_col = config.fields.interface_bin
        if bin_col not in df.columns:
            LOGGER.warning("interface_bin 列不存在，跳过 Interface_bin 分析")
            return pd.DataFrame(), False
    else:
        raise ValueError(f"未知的bin_type: {bin_type}")
    
    dev_col = config.fields.devrevstep
    visual_col = config.fields.visual_id

    visual_numeric = pd.to_numeric(df[visual_col], errors="coerce")
    is_numeric = visual_numeric.notna().all()
    if is_numeric:
        df = df.assign(_quantity=visual_numeric)
        agg = df.groupby([bin_col, dev_col])["_quantity"].sum()
        LOGGER.info(f"{bin_type.upper()}_BIN: VISUAL_ID 以数值方式求和")
    else:
        agg = df.groupby([bin_col, dev_col])[visual_col].nunique()
        LOGGER.warning(f"{bin_type.upper()}_BIN: VISUAL_ID 存在非数值，使用去重计数")

    quantity = agg.unstack(dev_col).fillna(0)
    quantity = quantity.sort_index(axis=0)
    return quantity, is_numeric


def build_pareto(quantity: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    col_totals = quantity.sum(axis=0)
    row_totals = quantity.sum(axis=1)
    grand_total = col_totals.sum()
    percentages = quantity.divide(col_totals.where(col_totals != 0, 1), axis=1).fillna(0)
    total_percentage = row_totals / grand_total if grand_total else row_totals * 0
    return percentages, row_totals, total_percentage


def collect_exceptions(df: pd.DataFrame, config: AppConfig) -> pd.DataFrame:
    required = [
        config.fields.lot,
        config.fields.functional_bin,
        config.fields.devrevstep,
        config.fields.visual_id,
    ]
    mask = df[required].isna().any(axis=1)
    return df.loc[mask].copy()

