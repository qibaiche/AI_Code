"""主工作流控制器"""
import logging
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional
import pandas as pd

from .config_loader import load_config, WorkflowConfig
from .data_reader import read_excel_file, save_result_excel, validate_data
from .mole_submitter import MoleSubmitter
from .spark_submitter import SparkSubmitter
from .gts_submitter import GTSSubmitter

LOGGER = logging.getLogger(__name__)


class WorkflowError(Exception):
    """工作流异常"""
    pass


class WorkflowController:
    """工作流控制器"""
    
    def __init__(self, config: WorkflowConfig):
        self.config = config
        self.mole_submitter = MoleSubmitter(config.mole)
        self.spark_submitter = SparkSubmitter(config.spark)
        self.gts_submitter = GTSSubmitter(config.gts)
        self.results = []
        self.errors = []
    
    def run_workflow(self, excel_file_path: str | Path) -> Path | None:
        """
        运行完整的工作流
        
        Args:
            excel_file_path: Excel文件路径
        
        Returns:
            输出Excel文件路径
        
        Raises:
            WorkflowError: 如果工作流执行失败
        """
        excel_file_path = Path(excel_file_path)
        
        LOGGER.info("=" * 80)
        LOGGER.info("开始执行自动化工作流")
        LOGGER.info(f"输入文件: {excel_file_path}")
        LOGGER.info("=" * 80)
        
        start_time = datetime.now()
        
        try:
            # 步骤0: 预先启动Mole工具
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 0/2: 启动Mole工具")
            LOGGER.info("=" * 80)
            self._step_start_mole()
            
            # 步骤1: 读取文件（Excel或CSV）
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 1/2: 读取source lot文件")
            LOGGER.info("=" * 80)
            df = self._step_read_excel(excel_file_path)
            
            # 步骤2: 提交MIR数据到Mole工具（循环处理所有行）
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 2/2: 提交MIR数据到Mole工具（循环处理所有行）")
            LOGGER.info("=" * 80)
            # 获取Source Lot文件路径
            source_lot_file_path = self._get_source_lot_file_path(excel_file_path)
            self._step_submit_to_mole(df, source_lot_file_path)
            
            # ========================================================================
            # 以下步骤已暂停（根据用户要求，执行到Add to Summary就停止）
            # ========================================================================
            # # 步骤3: 提交VPO数据到Spark网页
            # LOGGER.info("\n" + "=" * 80)
            # LOGGER.info("步骤 3/4: 提交VPO数据到Spark网页")
            # LOGGER.info("=" * 80)
            # self._step_submit_to_spark(df)
            # 
            # # 步骤4: 提交最终数据到GTS网站
            # LOGGER.info("\n" + "=" * 80)
            # LOGGER.info("步骤 4/4: 提交最终数据到GTS网站")
            # LOGGER.info("=" * 80)
            # self._step_submit_to_gts(df)
            # 
            # # 保存结果
            # LOGGER.info("\n" + "=" * 80)
            # LOGGER.info("保存处理结果")
            # LOGGER.info("=" * 80)
            # output_path = self._step_save_results(df)
            # ========================================================================
            
            output_path = None  # 不再保存结果文件
            
            # 计算执行时间
            elapsed_time = (datetime.now() - start_time).total_seconds()
            
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("✅ 工作流执行成功（已处理所有Source Lot行）")
            LOGGER.info(f"执行时间: {elapsed_time:.2f} 秒")
            if output_path:
                LOGGER.info(f"输出文件: {output_path}")
            else:
                LOGGER.info("注意: MIR结果已保存为CSV文件")
            LOGGER.info("=" * 80)
            
            return output_path
            
        except Exception as e:
            elapsed_time = (datetime.now() - start_time).total_seconds()
            error_msg = f"工作流执行失败: {e}"
            LOGGER.error("\n" + "=" * 80)
            LOGGER.error(f"❌ {error_msg}")
            LOGGER.error(f"执行时间: {elapsed_time:.2f} 秒")
            LOGGER.error("=" * 80)
            LOGGER.error(traceback.format_exc())
            raise WorkflowError(error_msg) from e
    
    def _get_source_lot_file_path(self, excel_file_path: Path) -> Path:
        """获取Source Lot文件路径"""
        # 优先使用配置文件中的路径
        if self.config.paths.source_lot_file and self.config.paths.source_lot_file.exists():
            return self.config.paths.source_lot_file
        
        # 否则在excel_file_path的同目录中查找
        parent_dir = excel_file_path.parent
        possible_names = [
            "Source Lot.csv",
            "Source Lot.xlsx",
            "Source Lot.xls",
            "source lot.csv",
            "source lot.xlsx",
            "source lot.xls",
        ]
        
        for filename in possible_names:
            file_path = parent_dir / filename
            if file_path.exists():
                return file_path
        
        raise WorkflowError(f"未找到Source Lot文件。请确保文件存在于: {parent_dir}")
    
    def _save_all_mir_results(self, source_lot_file_path: Path, mir_results: list) -> None:
        """
        保存所有MIR结果到CSV文件
        
        Args:
            source_lot_file_path: Source Lot文件路径
            mir_results: MIR结果列表，每个元素是一个字典，包含原始行数据+MIR列
        """
        try:
            # 创建DataFrame
            result_df = pd.DataFrame(mir_results)
            
            # 生成输出文件名（包含日期和时间）
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = source_lot_file_path.parent / f"MIR_Results_{date_str}.csv"
            
            # 保存到CSV
            result_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            LOGGER.info(f"✅ 所有MIR结果已保存到: {output_file}")
            LOGGER.info(f"   共 {len(mir_results)} 行数据")
            
            # 显示每行的详细信息
            for idx, result in enumerate(mir_results):
                source_lot = result.get('SourceLot', 'N/A')
                mir = result.get('MIR', 'N/A')
                LOGGER.info(f"   第 {idx + 1} 行: SourceLot={source_lot}, MIR={mir}")
            
        except Exception as e:
            LOGGER.error(f"保存MIR结果到CSV失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
    
    def _step_start_mole(self) -> None:
        """步骤0: 预先启动Mole工具"""
        try:
            LOGGER.info("正在启动Mole工具...")
            # 调用_ensure_application来启动Mole工具
            self.mole_submitter._ensure_application()
            LOGGER.info("✅ Mole工具已启动")
        except Exception as e:
            raise WorkflowError(f"启动Mole工具失败: {e}")
    
    def _step_read_excel(self, excel_file_path: Path) -> pd.DataFrame:
        """步骤1: 读取文件（Excel或CSV）"""
        try:
            df = read_excel_file(excel_file_path)
            file_type = "CSV文件" if excel_file_path.suffix.lower() == '.csv' else "Excel文件"
            LOGGER.info(f"✅ 成功读取{file_type}: {len(df)} 行，{len(df.columns)} 列")
            return df
        except Exception as e:
            raise WorkflowError(f"读取文件失败: {e}")
    
    def _step_submit_to_mole(self, df: pd.DataFrame, source_lot_file_path: Path) -> None:
        """步骤2: 提交MIR数据到Mole工具（循环处理所有行）"""
        try:
            # 读取Source Lot文件
            LOGGER.info("读取Source Lot文件...")
            source_lot_df = read_excel_file(source_lot_file_path)
            
            LOGGER.info(f"Source Lot文件列名: {source_lot_df.columns.tolist()}")
            LOGGER.info(f"Source Lot文件共有 {len(source_lot_df)} 行数据")
            
            # 查找SourceLot列（不区分大小写）
            source_lot_col = None
            for col in source_lot_df.columns:
                col_upper = str(col).strip().upper()
                if col_upper in ['SOURCELOT', 'SOURCE LOT', 'SOURCE_LOT', 'SOURCELOTS', 'SOURCE LOTS']:
                    source_lot_col = col
                    LOGGER.info(f"找到SourceLot列: '{col}'")
                    break
            
            if source_lot_col is None:
                raise WorkflowError(f"在Source Lot文件中未找到SourceLot列。可用列: {source_lot_df.columns.tolist()}")
            
            if source_lot_df.empty:
                raise WorkflowError("Source Lot文件为空")
            
            # 存储所有MIR结果
            mir_results = []
            
            # 循环处理每一行
            for row_index, row in source_lot_df.iterrows():
                source_lot_value = row[source_lot_col]
                
                if pd.isna(source_lot_value):
                    LOGGER.warning(f"第 {row_index + 1} 行的SourceLot值为空，跳过")
                    continue
                
                source_lot_value = str(source_lot_value).strip()
                
                LOGGER.info("=" * 80)
                LOGGER.info(f"处理第 {row_index + 1}/{len(source_lot_df)} 行: SourceLot = {source_lot_value}")
                LOGGER.info("=" * 80)
                
                try:
                    # 打开File菜单 -> New MIR Request
                    LOGGER.info("开始Mole工具操作流程...")
                    success = self.mole_submitter.submit_mir_data({})
                    
                    if success:
                        # 填写VPO搜索对话框
                        LOGGER.info("填写VPO搜索对话框...")
                        self.mole_submitter._fill_vpo_search_dialog(source_lot_value)
                        
                        # 检查搜索结果行状态并执行相应操作
                        LOGGER.info("检查搜索结果行状态...")
                        self.mole_submitter._check_row_status_and_select()
                        
                        # 点击Submit按钮
                        LOGGER.info("点击Submit按钮...")
                        self.mole_submitter._click_submit_button()
                        
                        # 处理最终成功对话框并获取MIR号码
                        LOGGER.info("处理最终成功对话框并获取MIR号码...")
                        mir_number = self.mole_submitter._handle_final_success_dialog_and_get_mir()
                        
                        if mir_number:
                            # 保存该行数据和MIR号码
                            result_row = row.to_dict()
                            result_row['MIR'] = mir_number
                            mir_results.append(result_row)
                            
                            LOGGER.info(f"✅ 第 {row_index + 1} 行处理成功: SourceLot={source_lot_value}, MIR={mir_number}")
                            
                            self.results.append({
                                'row_index': row_index,
                                'step': 'Mole',
                                'status': 'success',
                                'source_lot': source_lot_value,
                                'mir': mir_number,
                                'timestamp': datetime.now().isoformat()
                            })
                        else:
                            LOGGER.error(f"❌ 第 {row_index + 1} 行未能获取MIR号码")
                            self.errors.append({
                                'row_index': row_index,
                                'step': 'Mole',
                                'error': '未能获取MIR号码',
                                'source_lot': source_lot_value,
                                'timestamp': datetime.now().isoformat()
                            })
                    else:
                        error_msg = f"第 {row_index + 1} 行Mole工具操作失败"
                        LOGGER.error(f"❌ {error_msg}")
                        self.errors.append({
                            'row_index': row_index,
                            'step': 'Mole',
                            'error': error_msg,
                            'source_lot': source_lot_value,
                            'timestamp': datetime.now().isoformat()
                        })
                
                except Exception as e:
                    error_msg = f"第 {row_index + 1} 行处理失败: {e}"
                    LOGGER.error(f"❌ {error_msg}")
                    LOGGER.error(traceback.format_exc())
                    self.errors.append({
                        'row_index': row_index,
                        'step': 'Mole',
                        'error': str(e),
                        'source_lot': source_lot_value,
                        'timestamp': datetime.now().isoformat()
                    })
                    # 继续处理下一行，不中断整个流程
                
                # 在处理下一行之前，等待一下，确保界面准备好
                if row_index < len(source_lot_df) - 1:
                    LOGGER.info("等待2秒后处理下一行...")
                    time.sleep(2.0)
            
            # 所有行处理完后，保存结果到CSV
            if mir_results:
                LOGGER.info("=" * 80)
                LOGGER.info(f"所有行处理完成，保存结果...")
                self._save_all_mir_results(source_lot_file_path, mir_results)
            else:
                LOGGER.warning("没有成功获取任何MIR号码")
            
            # 输出汇总信息
            LOGGER.info("=" * 80)
            LOGGER.info(f"处理汇总:")
            LOGGER.info(f"  总行数: {len(source_lot_df)}")
            LOGGER.info(f"  成功: {len(mir_results)}")
            LOGGER.info(f"  失败: {len(self.errors)}")
            LOGGER.info("=" * 80)
                
        except Exception as e:
            raise WorkflowError(f"提交MIR数据到Mole工具失败: {e}")
    
    def _step_submit_to_spark(self, df: pd.DataFrame) -> None:
        """步骤3: 提交VPO数据到Spark网页"""
        try:
            # 使用上下文管理器确保WebDriver正确关闭
            with self.spark_submitter:
                # 将DataFrame转换为字典格式
                for idx, row in df.iterrows():
                    data = row.to_dict()
                    LOGGER.info(f"处理第 {idx + 1}/{len(df)} 行数据")
                    
                    success = self.spark_submitter.submit_vpo_data(data)
                    if success:
                        LOGGER.info(f"✅ 第 {idx + 1} 行数据提交成功")
                        self.results.append({
                            'row_index': idx,
                            'step': 'Spark',
                            'status': 'success',
                            'timestamp': datetime.now().isoformat()
                        })
                    else:
                        error_msg = f"第 {idx + 1} 行数据提交失败"
                        LOGGER.error(f"❌ {error_msg}")
                        self.errors.append({
                            'row_index': idx,
                            'step': 'Spark',
                            'error': error_msg,
                            'timestamp': datetime.now().isoformat()
                        })
                        raise WorkflowError(error_msg)
                
                LOGGER.info("✅ 所有VPO数据提交成功")
                
        except Exception as e:
            raise WorkflowError(f"提交VPO数据到Spark网页失败: {e}")
    
    def _step_submit_to_gts(self, df: pd.DataFrame) -> None:
        """步骤4: 提交最终数据到GTS网站"""
        try:
            # 使用上下文管理器确保WebDriver正确关闭
            with self.gts_submitter:
                # 将DataFrame转换为字典格式
                for idx, row in df.iterrows():
                    data = row.to_dict()
                    LOGGER.info(f"处理第 {idx + 1}/{len(df)} 行数据")
                    
                    success = self.gts_submitter.submit_final_data(data)
                    if success:
                        LOGGER.info(f"✅ 第 {idx + 1} 行数据提交成功")
                        self.results.append({
                            'row_index': idx,
                            'step': 'GTS',
                            'status': 'success',
                            'timestamp': datetime.now().isoformat()
                        })
                    else:
                        error_msg = f"第 {idx + 1} 行数据提交失败"
                        LOGGER.error(f"❌ {error_msg}")
                        self.errors.append({
                            'row_index': idx,
                            'step': 'GTS',
                            'error': error_msg,
                            'timestamp': datetime.now().isoformat()
                        })
                        raise WorkflowError(error_msg)
                
                LOGGER.info("✅ 所有最终数据提交成功")
                
        except Exception as e:
            raise WorkflowError(f"提交最终数据到GTS网站失败: {e}")
    
    def _step_save_results(self, df: pd.DataFrame) -> Path:
        """保存处理结果"""
        try:
            # 添加处理结果信息到数据框
            result_df = df.copy()
            
            # 添加处理状态列
            if self.results:
                results_df = pd.DataFrame(self.results)
                # 可以根据需要合并结果信息到原始数据框
                # 这里简化处理，直接保存原始数据和处理结果
            
            # 添加错误信息列
            if self.errors:
                errors_df = pd.DataFrame(self.errors)
                # 可以将错误信息合并到结果中
            
            # 生成日期字符串
            date_str = datetime.now().strftime("%Y%m%d")
            
            # 保存结果
            output_path = save_result_excel(
                result_df,
                self.config.paths.output_dir,
                date_str
            )
            
            # 如果存在错误，也保存错误日志
            if self.errors:
                error_log_path = self.config.paths.output_dir / f"workflow_errors_{date_str}.csv"
                errors_df = pd.DataFrame(self.errors)
                errors_df.to_csv(error_log_path, index=False, encoding='utf-8-sig')
                LOGGER.warning(f"存在 {len(self.errors)} 个错误，已保存到: {error_log_path}")
            
            LOGGER.info(f"✅ 结果已保存到: {output_path}")
            return output_path
            
        except Exception as e:
            raise WorkflowError(f"保存结果失败: {e}")

