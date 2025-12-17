"""主工作流控制器"""
import logging
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional
import pandas as pd

try:
    import win32gui
except ImportError:
    win32gui = None

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
        self.last_mir_result_file = None
    
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
            LOGGER.info("步骤 0/5: 启动Mole工具")
            LOGGER.info("=" * 80)
            self._step_start_mole()
            
            # 步骤1: 读取文件（Excel或CSV）
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 1/5: 读取source lot文件")
            LOGGER.info("=" * 80)
            df = self._step_read_excel(excel_file_path)
            
            # 步骤2: 提交MIR数据到Mole工具（循环处理所有行）
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 2/5: 提交MIR数据到Mole工具（循环处理所有行）")
            LOGGER.info("=" * 80)
            # 获取Source Lot文件路径
            source_lot_file_path = self._get_source_lot_file_path(excel_file_path)
            self._step_submit_to_mole(df, source_lot_file_path)
            
            # 步骤3: 提交VPO数据到Spark网页
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 3/5: 提交VPO数据到Spark网页")
            LOGGER.info("=" * 80)
            self._step_submit_to_spark(df)
            
            # 步骤4: 生成GTS填充文件
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 4/5: 生成GTS填充文件")
            LOGGER.info("=" * 80)
            self._step_generate_gts_file()
            
            # 步骤5: 自动填充并提交GTS
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("步骤 5/5: 自动填充并提交GTS")
            LOGGER.info("=" * 80)
            self._step_submit_to_gts()
            
            # 保存结果
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("保存处理结果")
            LOGGER.info("=" * 80)
            output_path = self._step_save_results(df)
            
            # 计算执行时间
            elapsed_time = (datetime.now() - start_time).total_seconds()
            
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("✅ 完整工作流执行成功！")
            LOGGER.info("=" * 80)
            LOGGER.info("已完成所有步骤:")
            LOGGER.info("  ✓ Mole: MIR 数据已提交")
            LOGGER.info("  ✓ Spark: VPO 数据已提交")
            LOGGER.info("  ✓ GTS: 填充文件已生成并提交")
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
    
    def run_mole_only(self, excel_file_path: str | Path) -> Path | None:
        """
        仅运行Mole步骤（不执行Spark/GTS）
        
        Args:
            excel_file_path: source lot文件路径
        
        Returns:
            最新的MIR结果文件路径（如果生成）
        """
        excel_file_path = Path(excel_file_path)
        
        LOGGER.info("=" * 80)
        LOGGER.info("开始执行Mole-only工作流（Spark/GTS已跳过）")
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
            source_lot_file_path = self._get_source_lot_file_path(excel_file_path)
            self._step_submit_to_mole(df, source_lot_file_path)
            
            elapsed_time = (datetime.now() - start_time).total_seconds()
            output_path = self.last_mir_result_file
            
            LOGGER.info("\n" + "=" * 80)
            LOGGER.info("✅ Mole-only工作流执行成功（已跳过Spark/GTS）")
            LOGGER.info(f"执行时间: {elapsed_time:.2f} 秒")
            if output_path:
                LOGGER.info(f"输出文件: {output_path}")
            else:
                LOGGER.info("注意: 未获取到MIR结果文件路径")
            LOGGER.info("=" * 80)
            
            return output_path
            
        except Exception as e:
            elapsed_time = (datetime.now() - start_time).total_seconds()
            error_msg = f"Mole-only工作流执行失败: {e}"
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
        
        # 父目录（Auto VPO根目录）
        parent_dir = excel_file_path.parent
        
        # 可能的文件名列表
        possible_names = [
            "Source Lot.csv",
            "Source Lot.xlsx",
            "Source Lot.xls",
            "source lot.csv",
            "source lot.xlsx",
            "source lot.xls",
        ]
        
        # 优先在 input/ 目录下查找
        input_dir = parent_dir / "input"
        if input_dir.exists():
            for filename in possible_names:
                file_path = input_dir / filename
                if file_path.exists():
                    LOGGER.info(f"在input目录找到Source Lot文件: {file_path}")
                    return file_path
        
        # 在父目录（Auto VPO根目录）中查找
        for filename in possible_names:
            file_path = parent_dir / filename
            if file_path.exists():
                LOGGER.info(f"在根目录找到Source Lot文件: {file_path}")
                return file_path
        
        raise WorkflowError(f"未找到Source Lot文件。请确保文件存在于以下位置之一:\n  - {input_dir if input_dir.exists() else parent_dir / 'input'}\n  - {parent_dir}")
    
    def _save_all_mir_results(self, source_lot_file_path: Path, mir_results: list) -> None:
        """
        保存所有MIR结果到CSV文件
        
        Args:
            source_lot_file_path: Source Lot文件路径
            mir_results: MIR结果列表，每个元素是一个字典，包含原始行数据+MIR列
        """
        try:
            if not mir_results:
                LOGGER.warning("没有MIR结果需要保存")
                return
            
            # 创建DataFrame
            result_df = pd.DataFrame(mir_results)
            
            # 查找SourceLot列（不区分大小写）
            source_lot_col = None
            for col in result_df.columns:
                col_upper = str(col).strip().upper()
                if col_upper in ['SOURCELOT', 'SOURCE LOT', 'SOURCE_LOT', 'SOURCELOTS', 'SOURCE LOTS']:
                    source_lot_col = col
                    break
            
            # 如果找到SourceLot列，按SourceLot分组，确保相同SourceLot的行放在一起
            # 同时保持source lot表的原始顺序（相同SourceLot首次出现的顺序）
            if source_lot_col:
                LOGGER.info(f"按SourceLot列 '{source_lot_col}' 分组，确保相同SourceLot的MIR放在一起...")
                
                # 添加原始索引列，用于保持原始顺序
                result_df['_original_index'] = range(len(result_df))
                
                # 记录每个SourceLot首次出现的索引
                first_occurrence = {}
                for idx, source_lot_value in enumerate(result_df[source_lot_col]):
                    if pd.notna(source_lot_value):
                        source_lot_str = str(source_lot_value).strip()
                        if source_lot_str and source_lot_str not in first_occurrence:
                            first_occurrence[source_lot_str] = idx
                
                # 创建分组键：SourceLot首次出现的索引 + SourceLot值
                def get_group_key(row):
                    source_lot_value = row[source_lot_col]
                    if pd.isna(source_lot_value):
                        return (float('inf'), '')  # NaN值放在最后
                    source_lot_str = str(source_lot_value).strip()
                    first_idx = first_occurrence.get(source_lot_str, float('inf'))
                    return (first_idx, source_lot_str)
                
                # 创建分组键列
                result_df['_group_key'] = result_df.apply(get_group_key, axis=1)
                
                # 按分组键和原始索引排序
                result_df = result_df.sort_values(by=['_group_key', '_original_index'], na_position='last')
                
                # 删除临时列
                result_df = result_df.drop(columns=['_original_index', '_group_key'])
                
                LOGGER.info(f"✅ 已按SourceLot分组，相同SourceLot的MIR已放在一起（保持source lot表的原始顺序）")
            else:
                LOGGER.warning("未找到SourceLot列，保持原始顺序")
            
            # 生成输出文件名（包含日期和时间）
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = self.config.paths.output_dir / f"MIR_Results_{date_str}.csv"
            
            # 保存到CSV
            result_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            LOGGER.info(f"✅ 所有MIR结果已保存到: {output_file}")
            LOGGER.info(f"   共 {len(mir_results)} 行数据")
            self.last_mir_result_file = output_file
            
            # 显示每行的详细信息（按排序后的顺序）
            for idx, (_, row) in enumerate(result_df.iterrows(), 1):
                source_lot = row.get(source_lot_col, 'N/A') if source_lot_col else 'N/A'
                mir = row.get('MIR', 'N/A')
                LOGGER.info(f"   第 {idx} 行: SourceLot={source_lot}, MIR={mir}")
            
            return output_file
            
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
    
    def _close_mole(self) -> None:
        """关闭Mole工具"""
        try:
            if not self.mole_submitter._window:
                LOGGER.info("Mole窗口未连接，尝试查找并关闭...")
                # 尝试查找MOLE窗口
                if win32gui:
                    def find_mole_window(hwnd, windows):
                        try:
                            if not win32gui.IsWindowVisible(hwnd):
                                return True
                            window_text = win32gui.GetWindowText(hwnd)
                            if "MOLE" in window_text.upper() and "LOGIN" not in window_text.upper():
                                windows.append(hwnd)
                        except:
                            pass
                        return True
                    
                    mole_windows = []
                    win32gui.EnumWindows(find_mole_window, mole_windows)
                    if mole_windows:
                        LOGGER.info(f"找到 {len(mole_windows)} 个Mole窗口，尝试关闭...")
                        for hwnd in mole_windows:
                            try:
                                win32gui.PostMessage(hwnd, 0x0010, 0, 0)  # WM_CLOSE
                                LOGGER.info(f"已发送关闭消息到Mole窗口")
                            except:
                                pass
                        time.sleep(2.0)
                        return
            
            # 如果已连接窗口，尝试关闭
            if self.mole_submitter._window:
                try:
                    self.mole_submitter._window.close()
                    LOGGER.info("✅ 已关闭Mole窗口")
                    time.sleep(1.0)
                except Exception as e:
                    LOGGER.warning(f"关闭Mole窗口失败: {e}，尝试其他方法...")
                    # 尝试通过进程关闭
                    try:
                        if win32gui and self.mole_submitter._window:
                            hwnd = self.mole_submitter._window.handle
                            win32gui.PostMessage(hwnd, 0x0010, 0, 0)  # WM_CLOSE
                            time.sleep(1.0)
                    except:
                        pass
        except Exception as e:
            LOGGER.warning(f"关闭Mole工具时出错: {e}")
    
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
                        
                        # 如果是最后一行，增加等待时间，确保copy MIR对话框完全弹出
                        is_last_row = (row_index == len(source_lot_df) - 1)
                        if is_last_row:
                            LOGGER.info("这是最后一行，等待更长时间确保copy MIR对话框完全弹出...")
                            time.sleep(3.0)  # 额外等待3秒
                        
                        mir_number = self.mole_submitter._handle_final_success_dialog_and_get_mir()
                        
                        # 如果是最后一行，再次等待确保对话框完全处理完成
                        if is_last_row:
                            LOGGER.info("最后一行处理完成，等待copy MIR对话框完全关闭...")
                            time.sleep(2.0)  # 再等待2秒确保对话框关闭
                            
                            # 验证对话框是否已关闭
                            if win32gui:
                                try:
                                    def check_dialog(hwnd, dialogs):
                                        try:
                                            if not win32gui.IsWindowVisible(hwnd):
                                                return True
                                            window_text = win32gui.GetWindowText(hwnd)
                                            if window_text == "Submit MIR Request":
                                                dialogs.append(hwnd)
                                        except:
                                            pass
                                        return True
                                    
                                    remaining_dialogs = []
                                    win32gui.EnumWindows(check_dialog, remaining_dialogs)
                                    if remaining_dialogs:
                                        LOGGER.warning(f"检测到仍有 {len(remaining_dialogs)} 个成功对话框未关闭，等待5秒...")
                                        time.sleep(5.0)
                                except:
                                    pass
                        
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
            
            # 所有行处理完后，等待所有对话框关闭，然后关闭MOLE
            LOGGER.info("=" * 80)
            LOGGER.info("所有MIR提交完成，等待所有对话框关闭...")
            LOGGER.info("=" * 80)
            
            # 等待所有成功对话框关闭（最多等待10秒）
            if win32gui:
                max_wait = 10
                for i in range(max_wait):
                    try:
                        def check_dialog(hwnd, dialogs):
                            try:
                                if not win32gui.IsWindowVisible(hwnd):
                                    return True
                                window_text = win32gui.GetWindowText(hwnd)
                                if window_text == "Submit MIR Request":
                                    dialogs.append(hwnd)
                            except:
                                pass
                            return True
                        
                        remaining_dialogs = []
                        win32gui.EnumWindows(check_dialog, remaining_dialogs)
                        if not remaining_dialogs:
                            LOGGER.info("✅ 所有对话框已关闭")
                            break
                        else:
                            if i % 2 == 0:
                                LOGGER.info(f"等待对话框关闭... ({i+1}/{max_wait}秒，还有{len(remaining_dialogs)}个对话框)")
                            time.sleep(1.0)
                    except:
                        time.sleep(1.0)
            
            # 关闭MOLE工具
            LOGGER.info("=" * 80)
            LOGGER.info("关闭MOLE工具...")
            LOGGER.info("=" * 80)
            try:
                self._close_mole()
                LOGGER.info("✅ MOLE工具已关闭")
            except Exception as e:
                LOGGER.warning(f"⚠️ 关闭MOLE工具时出错: {e}，继续执行...")
            
            # 保存结果到CSV
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
        """步骤3: 提交VPO数据到Spark网页（从MIR结果文件读取数据）"""
        try:
            # 仅在output目录查找最新的MIR结果文件（与test_spark一致）
            LOGGER.info("查找MIR结果文件（仅output目录）...")
            output_dir = self.config.paths.output_dir
            if not output_dir.exists():
                raise WorkflowError(f"output目录不存在: {output_dir}")
            
            mir_files = sorted(output_dir.glob("MIR_Results_*.csv"), reverse=True)
            
            if not mir_files:
                raise WorkflowError(f"未在output目录找到MIR结果文件，无法提交到Spark。请先完成Mole步骤。\n已检查目录: {output_dir}")
            
            # 使用最新的文件
            selected_file = mir_files[0]
            LOGGER.info(f"使用MIR结果文件: {selected_file.name}")
            
            # 读取MIR结果文件
            mir_df = read_excel_file(selected_file)
            if mir_df.empty:
                raise WorkflowError("MIR结果文件为空")
            
            LOGGER.info(f"成功读取MIR数据：{len(mir_df)} 行")
            LOGGER.info(f"MIR文件列名: {mir_df.columns.tolist()}")
            
            # 使用上下文管理器确保WebDriver正确关闭
            with self.spark_submitter:
                # 初始化并导航到页面（只需要一次）
                LOGGER.info("初始化Spark网页...")
                self.spark_submitter._init_driver()
                self.spark_submitter._navigate_to_page()
                
                # 循环处理每一行MIR结果
                # 使用enumerate确保行号从0开始，避免DataFrame索引问题
                for row_num, (idx, row) in enumerate(mir_df.iterrows()):
                    LOGGER.info("=" * 80)
                    LOGGER.info(f"处理第 {row_num + 1}/{len(mir_df)} 行MIR数据 (DataFrame索引: {idx})")
                    LOGGER.info("=" * 80)
                    
                    try:
                        # 提取数据（支持多种列名格式）
                        LOGGER.info(f"行数据: {row.to_dict()}")
                        
                        # 查找SourceLot列
                        source_lot = None
                        for col in row.index:
                            col_upper = str(col).strip().upper()
                            if col_upper in ['SOURCELOT', 'SOURCE LOT', 'SOURCE_LOT']:
                                source_lot = str(row[col]).strip() if pd.notna(row[col]) else ''
                                LOGGER.info(f"找到SourceLot列: '{col}' = '{source_lot}'")
                                break
                        
                        if not source_lot:
                            LOGGER.warning(f"第 {row_num + 1} 行SourceLot值为空，跳过")
                            LOGGER.warning(f"可用列: {row.index.tolist()}")
                            continue
                        
                        # 查找Part Type列
                        part_type = None
                        for col in row.index:
                            col_upper = str(col).strip().upper()
                            if col_upper in ['PART TYPE', 'PARTTYPE', 'PART_TYPE']:
                                part_type = str(row[col]).strip() if pd.notna(row[col]) else ''
                                break
                        
                        if not part_type:
                            LOGGER.warning(f"第 {row_num + 1} 行Part Type值为空，跳过")
                            continue
                        
                        # 查找Operation列（可选）
                        operation = None
                        for col in row.index:
                            col_upper = str(col).strip().upper()
                            if col_upper in ['OPERATION', 'OP', 'OPN']:
                                if pd.notna(row[col]) and str(row[col]).strip():
                                    operation = str(row[col]).strip()
                                break
                        
                        # 查找Eng ID列（可选，支持多种格式）
                        eng_id = None
                        for col in row.index:
                            col_upper = str(col).strip().upper()
                            if col_upper in ['ENG ID', 'ENGID', 'ENG_ID', 'ENGINEERING ID', 'ENGINEERING_ID']:
                                if pd.notna(row[col]) and str(row[col]).strip():
                                    eng_id = str(row[col]).strip()
                                break
                        
                        # 处理More options字段
                        unit_test_time = row.get('Unit test time', None)
                        retest_rate = row.get('Retest rate', None)
                        hri_mrv = row.get('HRI / MRV:', None)
                        
                        # 处理空值
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
                        
                        # 执行Spark提交流程
                        # 注意：第一行需要点击Add New，后续行在上一行Roll后已经点击了Add New
                        if row_num == 0:
                            LOGGER.info("步骤 1/13: 点击Add New...")
                            if not self.spark_submitter._click_add_new_button():
                                raise WorkflowError("点击Add New按钮失败")
                        else:
                            LOGGER.info("步骤 1/13: 已点击Add New（上一行Roll后已点击）")
                        
                        LOGGER.info("步骤 2/13: 填写TP路径...")
                        if not self.spark_submitter._fill_test_program_path(self.config.paths.tp_path):
                            raise WorkflowError("填写TP路径失败")
                        
                        LOGGER.info("步骤 3/13: 点击Add New Experiment...")
                        if not self.spark_submitter._click_add_new_experiment():
                            raise WorkflowError("点击Add New Experiment失败")
                        
                        LOGGER.info("步骤 4/13: 选择VPO类别...")
                        if not self.spark_submitter._select_vpo_category(self.config.spark.vpo_category):
                            raise WorkflowError("选择VPO类别失败")
                        
                        LOGGER.info("步骤 5/13: 填写实验信息...")
                        if not self.spark_submitter._fill_experiment_info(self.config.spark.step, self.config.spark.tags):
                            raise WorkflowError("填写实验信息失败")
                        
                        LOGGER.info("步骤 6/13: 添加Lot name...")
                        if not self.spark_submitter._add_lot_name(source_lot):
                            raise WorkflowError("添加Lot name失败")
                        
                        LOGGER.info("步骤 7/13: 选择Part Type...")
                        if not self.spark_submitter._select_parttype(part_type):
                            raise WorkflowError("选择Part Type失败")
                        
                        LOGGER.info("步骤 8/13: 点击Flow标签...")
                        if not self.spark_submitter._click_flow_tab():
                            raise WorkflowError("点击Flow标签失败")
                        
                        # Operation是可选的，但如果存在则必须成功选择
                        if operation:
                            LOGGER.info("步骤 9/13: 选择Operation...")
                            if not self.spark_submitter._select_operation(operation):
                                raise WorkflowError("选择Operation失败")
                        else:
                            LOGGER.info("步骤 9/13: 跳过Operation（文件中未提供）")
                        
                        # Eng ID是可选的，但如果存在则必须成功选择
                        if eng_id:
                            LOGGER.info("步骤 10/13: 选择Eng ID...")
                            if not self.spark_submitter._select_eng_id(eng_id):
                                raise WorkflowError("选择Eng ID失败")
                        else:
                            LOGGER.info("步骤 10/13: 跳过Eng ID（文件中未提供）")
                        
                        LOGGER.info("步骤 11/13: 点击More options标签...")
                        if not self.spark_submitter._click_more_options_tab():
                            raise WorkflowError("点击More options标签失败")
                        
                        LOGGER.info("步骤 12/13: 填写More options字段...")
                        if not self.spark_submitter._fill_more_options(unit_test_time, retest_rate, hri_mrv):
                            raise WorkflowError("填写More options字段失败")
                        
                        LOGGER.info("步骤 13/13: 点击Roll按钮...")
                        if not self.spark_submitter._click_roll_button():
                            raise WorkflowError("点击Roll按钮失败")
                        
                        # 等待Roll提交完成
                        LOGGER.info("等待Roll提交完成...")
                        time.sleep(3.0)  # 等待提交响应
                        
                        LOGGER.info(f"✅ 第 {row_num + 1} 行数据提交成功")
                        self.results.append({
                            'row_index': idx,
                            'step': 'Spark',
                            'status': 'success',
                            'source_lot': source_lot,
                            'timestamp': datetime.now().isoformat()
                        })
                        
                        # 如果不是最后一行，点击Add New按钮开始下一行
                        if row_num < len(mir_df) - 1:
                            LOGGER.info("=" * 80)
                            LOGGER.info(f"准备处理下一行（第 {row_num + 2}/{len(mir_df)} 行）...")
                            LOGGER.info("点击Add New按钮开始新的提交...")
                            LOGGER.info("=" * 80)
                            
                            # 等待页面稳定
                            time.sleep(2.0)
                            
                            # 点击Add New按钮
                            if not self.spark_submitter._click_add_new_button():
                                raise WorkflowError("点击Add New按钮失败，无法继续处理下一行")
                            
                            # 等待Add New对话框或页面响应
                            time.sleep(2.0)
                            
                            LOGGER.info("✅ 已点击Add New按钮，准备处理下一行")
                        else:
                            LOGGER.info("=" * 80)
                            LOGGER.info("这是最后一行，不需要点击Add New按钮")
                            LOGGER.info("=" * 80)
                        
                    except Exception as e:
                        error_msg = f"第 {row_num + 1} 行数据提交失败: {e}"
                        LOGGER.error(f"❌ {error_msg}")
                        LOGGER.error(traceback.format_exc())
                        
                        source_lot_value = source_lot if 'source_lot' in locals() and source_lot else 'N/A'
                        self.errors.append({
                            'row_index': idx,
                            'step': 'Spark',
                            'error': str(e),
                            'source_lot': source_lot_value,
                            'timestamp': datetime.now().isoformat()
                        })
                        
                        # 继续处理下一行，不中断整个流程
                        LOGGER.warning(f"⚠️ 第 {row_num + 1} 行提交失败，但将继续处理下一行...")
                        
                        # 如果不是最后一行，尝试点击Add New按钮，为下一行做准备
                        if row_num < len(mir_df) - 1:
                            try:
                                LOGGER.info("尝试点击Add New按钮，为下一行做准备...")
                                time.sleep(2.0)  # 等待页面稳定
                                if self.spark_submitter._click_add_new_button():
                                    LOGGER.info("✅ 已点击Add New按钮，可以继续处理下一行")
                                    time.sleep(2.0)
                                else:
                                    LOGGER.warning("⚠️ 点击Add New按钮失败，下一行可能无法正常处理")
                            except Exception as e2:
                                LOGGER.warning(f"⚠️ 尝试点击Add New按钮时出错: {e2}，但将继续处理下一行")
                
                LOGGER.info("=" * 80)
                LOGGER.info("✅ 所有VPO数据提交请求已完成")
                LOGGER.info(f"   总行数: {len(mir_df)}")
                LOGGER.info(f"   成功: {len([r for r in self.results if r.get('step') == 'Spark' and r.get('status') == 'success'])}")
                LOGGER.info(f"   失败: {len([e for e in self.errors if e.get('step') == 'Spark'])}")
                LOGGER.info("=" * 80)

            # ------------------------------------------------------------------
            # 所有提交完成后，等待一段时间，从Dashboard收集VPO并写回CSV
            # ------------------------------------------------------------------
            try:
                expected_vpo_count = len(mir_df)
                LOGGER.info("开始从Spark Dashboard收集VPO编号，用于回写到MIR结果CSV...")
                vpo_list = self.spark_submitter.collect_recent_vpos_from_dashboard(expected_count=expected_vpo_count)

                if not vpo_list:
                    LOGGER.warning("未能从Dashboard收集到任何VPO编号，跳过生成包含VPO的新CSV。")
                    return

                # 页面上顺序：最新在前；MIR CSV顺序：最早在前
                # 需要将列表反向后按顺序对应到每一行
                LOGGER.info("开始将收集到的VPO编号与MIR结果按顺序匹配...")
                vpo_list_reversed = list(reversed(vpo_list))

                mir_with_vpo = mir_df.copy()
                vpo_col_name = "VPO"
                if vpo_col_name in mir_with_vpo.columns:
                    LOGGER.warning(f"检测到MIR结果中已存在列 '{vpo_col_name}'，将覆盖该列的值。")

                mir_with_vpo[vpo_col_name] = ""

                max_count = min(len(mir_with_vpo), len(vpo_list_reversed))
                for i in range(max_count):
                    mir_with_vpo.at[mir_with_vpo.index[i], vpo_col_name] = vpo_list_reversed[i]
                    LOGGER.info(f"第 {i+1} 行: SourceLot={mir_with_vpo.iloc[i].get('SourceLot', 'N/A')} , VPO={vpo_list_reversed[i]}")

                # 生成新的带VPO的CSV文件
                date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = self.config.paths.output_dir / f"MIR_Results_with_VPO_{date_str}.csv"
                mir_with_vpo.to_csv(output_file, index=False, encoding="utf-8-sig")

                LOGGER.info(f"✅ 已生成包含VPO的新CSV文件: {output_file}")
                LOGGER.info(f"   共写入 {max_count} 条VPO记录（总行数: {len(mir_with_vpo)}）")
            except Exception as e:
                LOGGER.error(f"从Dashboard收集VPO并写回CSV时出错: {e}")
                LOGGER.error(traceback.format_exc())
                
        except Exception as e:
            raise WorkflowError(f"提交VPO数据到Spark网页失败: {e}")
    
    def _step_generate_gts_file(self) -> None:
        """步骤4: 生成GTS填充文件"""
        try:
            LOGGER.info("正在生成GTS填充文件...")
            
            # 导入GTS填充模块
            from .gts_fill_table import fill_gts_table
            
            # 查找最新的 MIR_Results_with_VPO_*.csv 文件
            output_dir = self.config.paths.output_dir
            vpo_files = sorted(output_dir.glob("MIR_Results_with_VPO_*.csv"), reverse=True)
            
            if not vpo_files:
                raise WorkflowError("未找到 MIR_Results_with_VPO_*.csv 文件，请先完成 Spark 步骤")
            
            input_file = vpo_files[0]
            LOGGER.info(f"使用输入文件: {input_file.name}")
            
            # 调用填充函数
            output_file = fill_gts_table(input_file, output_dir)
            
            if output_file and output_file.exists():
                LOGGER.info(f"✅ GTS填充文件已生成: {output_file.name}")
            else:
                raise WorkflowError("生成GTS填充文件失败")
                
        except Exception as e:
            raise WorkflowError(f"生成GTS填充文件失败: {e}")
    
    def _step_submit_to_gts(self) -> None:
        """步骤5: 自动填充并提交GTS"""
        try:
            LOGGER.info("正在打开GTS页面并自动填充...")
            
            # 调用新的自动填充逻辑
            self.gts_submitter.fill_ticket_with_latest_output()
            
            LOGGER.info("✅ GTS 填充和提交流程已完成")
            
        except Exception as e:
            LOGGER.error(f"GTS自动填充失败: {e}")
            LOGGER.error(traceback.format_exc())
            raise WorkflowError(f"提交GTS失败: {e}")
    
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

