"""主程序入口"""
import argparse
import logging
import sys
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

from .config_loader import load_config
from .workflow_main import WorkflowController, WorkflowError


class RobustFileHandler(logging.FileHandler):
    """增强的FileHandler，能够处理日志文件被删除的情况"""
    
    def __init__(self, filename, mode='a', encoding=None, delay=False):
        self._filename = Path(filename)
        self._encoding = encoding
        super().__init__(filename, mode, encoding, delay)
    
    def emit(self, record):
        """发送日志记录，如果文件被删除则重新创建"""
        try:
            # 检查文件是否存在，如果不存在则重新创建
            if not self._filename.exists():
                # 确保目录存在
                self._filename.parent.mkdir(parents=True, exist_ok=True)
                # 重新打开文件
                if self.stream:
                    self.stream.close()
                self.stream = self._open()
            
            super().emit(record)
        except (OSError, IOError) as e:
            # 如果写入失败，尝试重新创建文件
            try:
                if self.stream:
                    self.stream.close()
                self.stream = None
                # 确保目录存在
                self._filename.parent.mkdir(parents=True, exist_ok=True)
                self.stream = self._open()
                super().emit(record)
            except Exception:
                # 如果仍然失败，至少确保不会崩溃
                self.handleError(record)


def configure_logging(config):
    """配置日志"""
    log_file = config.paths.log_dir / config.logging.file
    log_file.parent.mkdir(parents=True, exist_ok=True)
    
    # 确保日志文件存在（即使为空）
    # 这样可以验证文件是否可以创建
    try:
        if not log_file.exists():
            log_file.touch()
    except (OSError, IOError, PermissionError) as e:
        print(f"警告: 无法创建日志文件 {log_file}: {e}")
        print("将仅使用控制台输出日志")
        # 只配置控制台输出
        logging.basicConfig(
            level=getattr(logging, config.logging.level.upper()),
            format=config.logging.format,
            handlers=[logging.StreamHandler(sys.stdout)],
            force=True  # Python 3.8+ 支持强制重新配置
        )
        return
    
    # 清除已有的处理器（如果存在）
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    try:
        # 创建文件处理器和控制台处理器
        file_handler = RobustFileHandler(log_file, encoding='utf-8')
        console_handler = logging.StreamHandler(sys.stdout)
        
        # 设置格式
        formatter = logging.Formatter(config.logging.format)
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # 设置日志级别
        log_level = getattr(logging, config.logging.level.upper())
        root_logger.setLevel(log_level)
        file_handler.setLevel(log_level)
        console_handler.setLevel(log_level)
        
        # 添加处理器
        root_logger.addHandler(file_handler)
        root_logger.addHandler(console_handler)
        
        # 验证日志文件是否可以写入
        test_logger = logging.getLogger(__name__)
        test_logger.info("日志系统初始化成功")
        
    except (OSError, IOError, PermissionError) as e:
        # 如果日志文件创建失败（例如权限问题、磁盘空间不足等），
        # 至少确保控制台输出仍然可用
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        
        console_handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(config.logging.format)
        console_handler.setFormatter(formatter)
        log_level = getattr(logging, config.logging.level.upper())
        root_logger.setLevel(log_level)
        root_logger.addHandler(console_handler)
        
        # 记录警告到控制台
        print(f"警告: 无法创建日志文件 {log_file}: {e}")
        print("将仅使用控制台输出日志")


def select_excel_file_gui() -> Path | None:
    """使用GUI选择文件（Excel或CSV）"""
    if not TKINTER_AVAILABLE:
        print("警告: tkinter未安装，无法使用GUI文件选择器")
        return None
    
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    file_path = filedialog.askopenfilename(
        title="请选择source lot 文件",
        filetypes=[
            ("所有支持的文件", "*.xlsx *.xls *.csv"),
            ("Excel文件", "*.xlsx *.xls"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*")
        ]
    )
    
    root.destroy()
    
    if file_path:
        return Path(file_path)
    return None


def select_excel_file_cli() -> Path | None:
    """使用命令行输入文件路径（Excel或CSV）"""
    print("\n" + "=" * 80)
    print("自动化工作流 - 文件选择")
    print("=" * 80)
    
    file_path_input = input("请选择source lot 文件（或拖拽文件到此窗口）: ").strip()
    
    # 处理拖拽文件时可能包含的引号
    file_path_input = file_path_input.strip('"').strip("'")
    
    if not file_path_input:
        print("错误: 未输入文件路径")
        return None
    
    file_path = Path(file_path_input)
    
    if not file_path.exists():
        print(f"错误: 文件不存在: {file_path}")
        return None
    
    if not file_path.suffix.lower() in ['.xlsx', '.xls', '.csv']:
        print(f"错误: 不支持的文件格式: {file_path.suffix}，仅支持.xlsx、.xls和.csv格式")
        return None
    
    return file_path


def run_workflow_cli(excel_file_path: Path, config_path: Path, mole_only: bool = False) -> None:
    """运行工作流（命令行模式）
    
    Args:
        excel_file_path: source lot文件路径
        config_path: 配置文件路径
        mole_only: 为True时仅运行Mole步骤，跳过Spark/GTS
    """
    try:
        # 加载配置
        config = load_config(config_path)
        
        # 配置日志
        configure_logging(config)
        
        logger = logging.getLogger(__name__)
        logger.info("=" * 80)
        logger.info("自动化工作流启动")
        logger.info("=" * 80)
        if mole_only:
            logger.info("当前模式: 仅运行Mole步骤（Spark/GTS已跳过）")
        
        # 创建控制器并运行工作流
        controller = WorkflowController(config)
        output_path = controller.run_mole_only(excel_file_path) if mole_only else controller.run_workflow(excel_file_path)
        
        # 显示成功消息
        if TKINTER_AVAILABLE:
            root = tk.Tk()
            root.withdraw()
            success_title = "Mole流程执行成功" if mole_only else "完整工作流执行成功"
            if mole_only:
                success_body = "Mole步骤已完成（Spark/GTS已跳过）"
            else:
                success_body = "完整工作流执行成功！\n\n已完成所有步骤:\n✓ Mole: MIR 数据已提交\n✓ Spark: VPO 数据已提交\n✓ GTS: 填充文件已生成并提交"
            if output_path:
                success_body += f"\n\n输出文件:\n{output_path}"
            messagebox.showinfo(success_title, success_body)
            root.destroy()
        else:
            if mole_only:
                print(f"\n✅ Mole步骤执行成功（Spark/GTS已跳过）")
            else:
                print(f"\n✅ 完整工作流执行成功！")
                print("已完成所有步骤:")
                print("  ✓ Mole: MIR 数据已提交")
                print("  ✓ Spark: VPO 数据已提交")
                print("  ✓ GTS: 填充文件已生成并提交")
            if output_path:
                print(f"输出文件: {output_path}")
    
    except WorkflowError as e:
        logger = logging.getLogger(__name__)
        logger.error(f"工作流执行失败: {e}")
        
        if TKINTER_AVAILABLE:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("工作流执行失败", f"工作流执行失败:\n\n{e}")
            root.destroy()
        else:
            print(f"\n❌ 工作流执行失败: {e}")
        
        sys.exit(1)
    
    except Exception as e:
        logger = logging.getLogger(__name__)
        logger.error(f"发生未预期的错误: {e}", exc_info=True)
        
        if TKINTER_AVAILABLE:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("错误", f"发生未预期的错误:\n\n{e}")
            root.destroy()
        else:
            print(f"\n❌ 发生未预期的错误: {e}")
        
        sys.exit(1)


def find_source_lot_file(base_dir: Path, config=None) -> Path | None:
    """自动查找Source Lot文件"""
    logger = logging.getLogger(__name__)
    
    # 首先检查配置文件中的路径
    if config and config.paths.source_lot_file and config.paths.source_lot_file.exists():
        logger.info(f"从配置中找到Source Lot文件: {config.paths.source_lot_file}")
        return config.paths.source_lot_file
    
    # 父目录（Auto VPO根目录）
    parent_dir = base_dir.parent
    
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
                logger.info(f"自动找到Source Lot文件（在input目录）: {file_path}")
                return file_path
    
    # 在父目录（Auto VPO根目录）中查找
    for filename in possible_names:
        file_path = parent_dir / filename
        if file_path.exists():
            logger.info(f"自动找到Source Lot文件: {file_path}")
            return file_path
    
    return None


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="自动化工作流入口")
    parser.add_argument(
        "--mole-only",
        action="store_true",
        help="仅运行Mole步骤，跳过Spark/GTS"
    )
    args = parser.parse_args()
    
    # 确定配置文件路径
    base_dir = Path(__file__).parent
    config_path = base_dir / "config.yaml"
    
    if not config_path.exists():
        print(f"错误: 配置文件不存在: {config_path}")
        print("请确保config.yaml文件存在于workflow_automation目录中")
        sys.exit(1)
    
    # 加载配置（用于获取source_lot_file路径）
    try:
        config = load_config(config_path)
    except Exception as e:
        print(f"警告: 加载配置失败: {e}，使用默认查找方式")
        config = None
    
    # 自动查找Source Lot文件
    excel_file_path = find_source_lot_file(base_dir, config)
    
    if not excel_file_path:
        print("警告: 未找到Source Lot文件，尝试手动选择...")
        # 如果自动查找失败，回退到GUI或命令行选择
        if TKINTER_AVAILABLE:
            try:
                excel_file_path = select_excel_file_gui()
            except Exception as e:
                print(f"GUI选择失败: {e}，改用命令行输入")
        
        # 如果GUI失败或不可用，使用命令行
        if not excel_file_path:
            excel_file_path = select_excel_file_cli()
        
        if not excel_file_path:
            print("错误: 未找到Source Lot文件且无法手动选择")
            print(f"请在 {base_dir.parent} 目录中放置 'Source Lot.csv' 文件")
            sys.exit(1)
    
    # 运行工作流
    run_workflow_cli(excel_file_path, config_path, mole_only=args.mole_only)


if __name__ == "__main__":
    main()

