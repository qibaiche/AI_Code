import logging
from datetime import datetime
from pathlib import Path
from typing import Tuple, List, Optional

import pandas as pd
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .config_loader import AppConfig


LOGGER = logging.getLogger(__name__)


def _assemble_pareto_table(
    quantity: pd.DataFrame,
    percentages: pd.DataFrame,
    row_totals: pd.Series,
    total_percentage: pd.Series,
) -> pd.DataFrame:
    """组装Pareto表格，使用多级列索引，并在底部添加summary行"""
    import logging
    LOGGER = logging.getLogger(__name__)
    
    # 检查 quantity 是否为空
    if quantity.empty:
        LOGGER.warning("⚠️ 警告：quantity 表为空，生成的 Pareto 表将只包含 Grand Total 行（值为 0）")
        LOGGER.warning("   可能的原因：")
        LOGGER.warning("   1. 数据被过滤条件全部过滤掉了")
        LOGGER.warning("   2. 数据中没有有效的 functional_bin 或 devrevstep 值")
        LOGGER.warning("   3. 列名不匹配")
    
    data = {}
    for dev in quantity.columns:
        data[(dev, "Quantity")] = quantity[dev]
        data[(dev, "Percentage")] = percentages[dev]
    data[("Grand Total", "Quantity")] = row_totals
    data[("Grand Total", "Percentage")] = total_percentage
    table = pd.DataFrame(data)
    table.index.name = quantity.index.name if not quantity.empty else None
    if not quantity.empty:
        table = table.sort_values(("Grand Total", "Quantity"), ascending=False)
    
    # 在表格底部添加summary行（Grand Total行）
    summary_row = {}
    for dev in quantity.columns:
        # 计算每个devrevstep的总计
        col_total_qty = quantity[dev].sum()
        col_total_pct = 1.0 if col_total_qty > 0 else 0.0  # 总计行的百分比应该是100%
        summary_row[(dev, "Quantity")] = col_total_qty
        summary_row[(dev, "Percentage")] = col_total_pct
    
    # Grand Total列的总计
    grand_total_qty = row_totals.sum()
    grand_total_pct = 1.0 if grand_total_qty > 0 else 0.0
    summary_row[("Grand Total", "Quantity")] = grand_total_qty
    summary_row[("Grand Total", "Percentage")] = grand_total_pct
    
    # 将summary行添加到表格底部
    summary_df = pd.DataFrame([summary_row], index=["Grand Total"])
    table = pd.concat([table, summary_df])
    
    return table


def _write_pareto_sheet_with_style(
    writer: pd.ExcelWriter,
    pareto_table: pd.DataFrame,
    sheet_name: str,
    config: AppConfig,
) -> None:
    """写入Pareto表格到Excel，并应用样式"""
    # 写入数据（从第6行开始，前5行留给标题和自定义表头）
    pareto_table.to_excel(writer, sheet_name=sheet_name, startrow=5, index=True)
    
    workbook = writer.book
    ws = workbook[sheet_name]
    
    # 获取阈值
    red_threshold = config.processing.percentage_thresholds.get("red", 0.05)
    yellow_threshold = config.processing.percentage_thresholds.get("yellow", 0.02)
    
    # 设置样式
    _format_pareto_sheet(ws, pareto_table, sheet_name, red_threshold, yellow_threshold)


def _format_pareto_sheet(ws, pareto_table: pd.DataFrame, sheet_name: str, red_threshold: float, yellow_threshold: float) -> None:
    """格式化Pareto表格，使其像Excel透视表"""
    max_row = ws.max_row
    max_col = ws.max_column
    
    # 清除所有合并单元格
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))
    
    # 删除pandas写入的表头行（第6-8行，即第6、7、8行）
    # 注意：pandas从第6行开始写入（startrow=5），所以第6、7、8行是pandas的表头
    ws.delete_rows(6, 3)  # 从第6行开始删除3行
    
    # 清除前5行的内容（自定义表头区域）
    for row in range(1, 6):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.value = None
    
    # 定义样式
    # 标题样式（浅粉色背景，深紫色字体）
    title_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # 浅粉色/灰色
    title_font = Font(bold=True, color="7030A0", size=14)  # 深紫色
    
    # 表头样式
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    subheader_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    subheader_font = Font(bold=True, size=10)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_align = Alignment(horizontal="center", vertical="center")
    
    # 第1-2行：标题行
    title_text = sheet_name  # 例如 "Interface_bin Pareto" 或 "Functional_bin Pareto"
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = title_text
    title_cell.fill = title_fill
    title_cell.font = title_font
    title_cell.alignment = center_align
    # 合并标题单元格（跨所有列，跨2行）
    end_col_letter = get_column_letter(max_col)
    ws.merge_cells(f"A1:{end_col_letter}2")
    # 设置第2行也为标题样式
    for col in range(1, max_col + 1):
        cell = ws.cell(row=2, column=col)
        cell.fill = title_fill
        cell.border = border
    
    # 处理多级表头（第3、4、5行）
    if isinstance(pareto_table.columns, pd.MultiIndex):
        # 定义行号
        row3 = 3
        row4 = 4
        row5 = 5
        col = 2  # 从B列开始（A列是索引）
        
        # 获取所有唯一的devrevstep值
        devrevsteps = [col[0] for col in pareto_table.columns if col[0] != "Grand Total"]
        devrevsteps = list(dict.fromkeys(devrevsteps))  # 去重保持顺序
        
        # 写入"Devrevstep"主表头（合并所有devrevstep列，只跨第3行）
        devrevstep_start_col = col
        devrevstep_end_col = col
        for dev in devrevsteps:
            sub_cols = [c for c in pareto_table.columns if c[0] == dev]
            devrevstep_end_col += len(sub_cols)
        
        if devrevstep_end_col > devrevstep_start_col:
            cell = ws.cell(row=row3, column=devrevstep_start_col)
            cell.value = "Devrevstep"
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
            end_col_letter = get_column_letter(devrevstep_end_col - 1)
            ws.merge_cells(f"{cell.coordinate}:{end_col_letter}{row3}")  # 只跨第3行
        
        # Grand Total列（第3行，跨第3-4行，跨所有 Grand Total 列）
        gt_cols = [c for c in pareto_table.columns if c[0] == "Grand Total"]
        if gt_cols:
            cell = ws.cell(row=row3, column=devrevstep_end_col)
            cell.value = "Grand Total"
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
            end_col_letter = get_column_letter(devrevstep_end_col + len(gt_cols) - 1)
            # 跨第3-4行，跨所有 Grand Total 列
            ws.merge_cells(f"{cell.coordinate}:{end_col_letter}{row4}")
        
        # 第4行：子表头（具体的devrevstep值和Grand Total，以及索引列标题）
        col = 2
        
        # 索引列标题（第4行，跨第3-5行）
        index_name = pareto_table.index.name or "Index"
        cell_a4 = ws.cell(row=row4, column=1)
        cell_a4.value = index_name
        cell_a4.fill = header_fill
        cell_a4.font = header_font
        cell_a4.alignment = center_align
        cell_a4.border = border
        ws.merge_cells(f"A3:A5")  # 索引列跨第3-5行
        
        # 写入具体的devrevstep值（第4行）
        for dev in devrevsteps:
            sub_cols = [c for c in pareto_table.columns if c[0] == dev]
            span = len(sub_cols)
            
            if span > 0:
                cell = ws.cell(row=row4, column=col)
                cell.value = dev
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = border
                
                # 合并单元格（只在第4行合并，不跨第5行）
                if span > 1:
                    end_col_letter = get_column_letter(col + span - 1)
                    ws.merge_cells(f"{cell.coordinate}:{end_col_letter}{row4}")  # 只跨第4行
                # 如果只有一列，不需要合并
                
                col += span
        
        # Grand Total列（第4行，不显示文字，只保留样式，因为第5行会显示 Quantity 和 Percentage）
        # 第3行的 "Grand Total" 已经跨所有列（Quantity 和 Percentage），第4行不需要显示文字
        if gt_cols:
            gt_col = devrevstep_end_col  # Grand Total列的起始列
            
            # 确保第4行的单元格没有被合并（如果被合并了，先取消）
            cell_coord = ws.cell(row=row4, column=gt_col).coordinate
            for merged_range in list(ws.merged_cells.ranges):
                if cell_coord in merged_range:
                    ws.unmerge_cells(str(merged_range))
                    break
            
            # 第4行不显示 "Grand Total" 文字，只保留样式（蓝色背景）
            # 对 Grand Total 的所有列应用样式
            for i, gt_col_tuple in enumerate(gt_cols):
                cell = ws.cell(row=row4, column=gt_col + i)
                cell.value = None  # 不显示文字
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = border
        
        # 第5行：子表头（Quantity和Percentage）
        col = 2
        for col_tuple in pareto_table.columns:
            cell = ws.cell(row=row5, column=col)
            cell.value = col_tuple[1]  # Quantity 或 Percentage
            cell.fill = subheader_fill
            cell.font = subheader_font
            cell.alignment = center_align
            cell.border = border
            col += 1
        
        # 索引列已经在上面合并了（A3:A5），不需要单独处理第5行
        
        data_start_row = 6
    else:
        # 单级表头的情况
        data_start_row = 2
        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
    
    # 格式化数据行
    # 检查最后一行是否是"Grand Total"（summary行）
    # summary行使用与Devrevstep相同的蓝色背景
    summary_row_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # 蓝色背景
    summary_row_font = Font(bold=True, color="FFFFFF", size=11)  # 白色字体
    
    for row in range(data_start_row, max_row + 1):
        # 检查是否是summary行（最后一行，且索引为"Grand Total"）
        is_summary_row = False
        if row == max_row:
            try:
                cell_value = ws.cell(row=row, column=1).value
                if cell_value and "Grand Total" in str(cell_value):
                    is_summary_row = True
            except:
                pass
        
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            if col == 1:  # 索引列
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = center_align
            
            # 如果是summary行，应用特殊样式
            if is_summary_row:
                cell.fill = summary_row_fill
                cell.font = summary_row_font
    
    # 删除summary行下面的空白行（如果有的话）
    # 检查summary行是否是最后一行
    summary_row_num = None
    for row in range(data_start_row, max_row + 1):
        try:
            cell_value = ws.cell(row=row, column=1).value
            if cell_value and "Grand Total" in str(cell_value):
                summary_row_num = row
                break
        except:
            pass
    
    if summary_row_num:
        # 删除summary行后面的所有空白行
        rows_to_delete = max_row - summary_row_num
        if rows_to_delete > 0:
            ws.delete_rows(summary_row_num + 1, rows_to_delete)
            max_row = summary_row_num  # 更新max_row
    
    # 格式化百分比列
    percentage_columns: List[int] = []
    col = 2
    for col_tuple in pareto_table.columns:
        if isinstance(col_tuple, tuple) and col_tuple[1] == "Percentage":
            percentage_columns.append(col)
        elif isinstance(col_tuple, str) and "Percentage" in str(col_tuple):
            percentage_columns.append(col)
        col += 1
    
    for col in percentage_columns:
        for row in range(data_start_row, max_row + 1):
            cell = ws.cell(row=row, column=col)
            cell.number_format = "0.00%"
    
    # 识别需要特殊处理的 bin（functional bin 100 和 interface bin 1）
    # 判断是 functional bin 还是 interface bin
    is_functional_bin = "Func" in sheet_name or "Functional" in sheet_name
    is_interface_bin = "Intf" in sheet_name or "Interface" in sheet_name
    
    # 需要排除的行（bin 100 用于 functional，bin 1 用于 interface）
    excluded_rows = set()
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    
    # 遍历数据行，识别需要特殊处理的行
    for row in range(data_start_row, max_row):
        try:
            bin_value = ws.cell(row=row, column=1).value  # A列是 bin 值
            if bin_value is not None:
                bin_str = str(bin_value).strip()
                # functional bin 100 或 interface bin 1
                if (is_functional_bin and bin_str == "100") or (is_interface_bin and bin_str == "1"):
                    excluded_rows.add(row)
                    # 直接填充绿色（对所有百分比列）
                    for col in percentage_columns:
                        cell = ws.cell(row=row, column=col)
                        cell.fill = green_fill
        except:
            pass
    
    # 应用条件格式（仅对百分比列，排除summary行和特殊bin行）
    if percentage_columns:
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")
        
        # 确定数据行的范围（排除summary行）
        data_end_row = max_row - 1  # 排除最后一行（summary行）
        if data_end_row < data_start_row:
            data_end_row = data_start_row  # 如果没有数据行，至少包含起始行
        
        for col in percentage_columns:
            col_letter = get_column_letter(col)
            
            # 构建不包含排除行的范围列表
            ranges_to_format = []
            current_start = None
            
            for row in range(data_start_row, data_end_row + 1):
                if row not in excluded_rows:
                    if current_start is None:
                        current_start = row
                else:
                    # 遇到排除行，如果之前有连续范围，添加到列表
                    if current_start is not None:
                        if current_start == row - 1:
                            # 单个单元格
                            ranges_to_format.append(f"{col_letter}{current_start}")
                        else:
                            # 范围
                            ranges_to_format.append(f"{col_letter}{current_start}:{col_letter}{row - 1}")
                        current_start = None
            
            # 处理最后一段
            if current_start is not None:
                if current_start == data_end_row:
                    ranges_to_format.append(f"{col_letter}{current_start}")
                else:
                    ranges_to_format.append(f"{col_letter}{current_start}:{col_letter}{data_end_row}")
            
            # 对每个范围应用条件格式
            for range_ref in ranges_to_format:
                # 红色：大于red_threshold
                ws.conditional_formatting.add(
                    range_ref,
                    CellIsRule(operator="greaterThan", formula=[str(red_threshold)], fill=red_fill),
                )
                # 黄色：在yellow_threshold和red_threshold之间
                ws.conditional_formatting.add(
                    range_ref,
                    CellIsRule(operator="between", formula=[str(yellow_threshold), str(red_threshold)], fill=yellow_fill),
                )
                # 绿色：小于等于yellow_threshold
                ws.conditional_formatting.add(
                    range_ref,
                    CellIsRule(operator="lessThanOrEqual", formula=[str(yellow_threshold)], fill=green_fill),
                )
    
    # 设置列宽
    ws.column_dimensions["A"].width = 15
    for col in range(2, max_col + 1):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 12


def save_report(
    report_dir: Path,
    df: pd.DataFrame,
    functional_pareto: pd.DataFrame,
    interface_pareto: Optional[pd.DataFrame],
    retest_pareto: Optional[pd.DataFrame],
    exceptions: pd.DataFrame,
    config: AppConfig,
    functional_pareto_by_step: Optional[dict] = None,
    interface_pareto_by_step: Optional[dict] = None,
    retest_pareto_by_step: Optional[dict] = None,
) -> Path:
    """保存报告到Excel文件"""
    report_name = f"Bin_Pareto_Report_{datetime.now():%Y-%m-%d}.xlsx"
    report_path = report_dir / report_name
    LOGGER.info("生成报表：%s", report_path)

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        # Functional_bin Pareto（总的，如果存在）
        if functional_pareto is not None and not functional_pareto.empty:
            _write_pareto_sheet_with_style(
                writer, functional_pareto, "Functional_bin Pareto", config
            )
        
        # 按 process_step 分组的 Functional_bin Pareto（如果存在）
        if functional_pareto_by_step:
            LOGGER.info(f"生成 {len(functional_pareto_by_step)} 个按 process_step 分组的 Functional_bin Pareto 表")
            for step_name, pareto_table in functional_pareto_by_step.items():
                if not pareto_table.empty:
                    # Excel sheet 名称限制为 31 个字符
                    sheet_name = f"FuncBin_{step_name}"[:31]
                    _write_pareto_sheet_with_style(
                        writer, pareto_table, sheet_name, config
                    )
                    LOGGER.info(f"✅ 已生成 {step_name} 的 Functional_bin Pareto 表")
        
        # Interface_bin Pareto（总的，如果存在）
        if interface_pareto is not None and not interface_pareto.empty:
            _write_pareto_sheet_with_style(
                writer, interface_pareto, "Interface_bin Pareto", config
            )
        
        # 按 process_step 分组的 Interface_bin Pareto（如果存在）
        if interface_pareto_by_step:
            LOGGER.info(f"生成 {len(interface_pareto_by_step)} 个按 process_step 分组的 Interface_bin Pareto 表")
            for step_name, pareto_table in interface_pareto_by_step.items():
                if not pareto_table.empty:
                    # Excel sheet 名称限制为 31 个字符
                    sheet_name = f"IntfBin_{step_name}"[:31]
                    _write_pareto_sheet_with_style(
                        writer, pareto_table, sheet_name, config
                    )
                    LOGGER.info(f"✅ 已生成 {step_name} 的 Interface_bin Pareto 表")
        
        # Retest Bin Pareto（总的，如果存在）
        if retest_pareto is not None and not retest_pareto.empty:
            _write_pareto_sheet_with_style(
                writer, retest_pareto, "Retest Bin Pareto", config
            )
        
        # 按 process_step 分组的 Retest Bin Pareto（如果存在）
        if retest_pareto_by_step:
            LOGGER.info(f"生成 {len(retest_pareto_by_step)} 个按 process_step 分组的 Retest Bin Pareto 表")
            for step_name, pareto_table in retest_pareto_by_step.items():
                if not pareto_table.empty:
                    # Excel sheet 名称限制为 31 个字符
                    sheet_name = f"RetestBin_{step_name}"[:31]
                    _write_pareto_sheet_with_style(
                        writer, pareto_table, sheet_name, config
                    )
                    LOGGER.info(f"✅ 已生成 {step_name} 的 Retest Bin Pareto 表")
        
        # Details
        df.to_excel(writer, sheet_name="Details", index=False)
        
        # Exceptions
        exceptions.to_excel(writer, sheet_name="Exceptions", index=False)

    return report_path


def build_pareto_table(
    quantity: pd.DataFrame,
    percentages: pd.DataFrame,
    row_totals: pd.Series,
    total_percentage: pd.Series,
) -> pd.DataFrame:
    """构建Pareto表格"""
    return _assemble_pareto_table(quantity, percentages, row_totals, total_percentage)
