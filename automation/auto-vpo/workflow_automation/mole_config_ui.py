"""Mole配置UI - 用于选择搜索方式和配置参数"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import yaml
from pathlib import Path
from typing import Dict, Optional
import logging

LOGGER = logging.getLogger(__name__)


class MoleConfigUI:
    """Mole配置界面"""
    
    def __init__(self, config_path: Path):
        """
        初始化配置UI
        
        Args:
            config_path: config.yaml文件路径
        """
        self.config_path = config_path
        self.config_data = self._load_config()
        self.result = None
        self.root = None
        
        # 获取Mole历史配置
        self.mole_history = self.config_data.get('mole_history', {})
        
        # 获取默认的 source lot 文件路径
        paths_config = self.config_data.get('paths', {})
        self.default_source_lot_file = paths_config.get('source_lot_file', '../input/Source Lot.csv')
        
    def _load_config(self) -> dict:
        """加载配置文件"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            LOGGER.error(f"加载配置文件失败: {e}")
            return {}
    
    def _save_config(self, mole_ui_config: dict) -> None:
        """
        保存配置到config.yaml
        
        Args:
            mole_ui_config: Mole UI配置数据
        """
        try:
            # 更新mole_history节
            self.config_data['mole_history'] = mole_ui_config
            
            # 保存到文件
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.config_data, f, default_flow_style=False, allow_unicode=True)
            
            LOGGER.info(f"配置已保存到: {self.config_path}")
        except Exception as e:
            LOGGER.error(f"保存配置失败: {e}")
            messagebox.showerror("保存失败", f"无法保存配置:\n{e}")
    
    def show(self) -> Optional[Dict]:
        """
        显示配置对话框
        
        Returns:
            用户配置的字典，如果取消则返回None
            {
                'search_mode': 'vpos' or 'units' or 'inactive_cage',
                'product_code': str (可选),
                'source_lot': str (可选),
                'show_available_units': bool (仅当search_mode='units'时),
                ... 其他字段
            }
        """
        self.root = tk.Tk()
        self.root.title("Mole 配置 - 自动化工作流")
        self.root.geometry("750x850")  # 增加高度确保按钮可见
        self.root.resizable(True, True)
        self.root.minsize(700, 750)  # 设置最小尺寸，确保按钮区域可见
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame,
            text="Mole 自动化配置",
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # 说明文字
        description = ttk.Label(
            main_frame,
            text="请选择搜索方式并配置参数（默认使用上次配置）",
            font=("Arial", 9)
        )
        description.pack(pady=(0, 15))
        
        # --- 搜索方式选择 ---
        search_mode_frame = ttk.LabelFrame(main_frame, text="选择搜索方式", padding="10")
        search_mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.search_mode = tk.StringVar(value=self.mole_history.get('search_mode', 'vpos'))
        
        ttk.Radiobutton(
            search_mode_frame,
            text="Search By VPOs",
            variable=self.search_mode,
            value='vpos',
            command=self._on_search_mode_change
        ).pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(
            search_mode_frame,
            text="Search By Units",
            variable=self.search_mode,
            value='units',
            command=self._on_search_mode_change
        ).pack(anchor=tk.W, pady=2)
        
        # --- 公共参数 ---
        common_params_frame = ttk.LabelFrame(main_frame, text="公共参数", padding="10")
        common_params_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 10))  # 改为 expand=False，不占用过多空间
        
        # MIR Comments
        mir_comments_label = ttk.Label(common_params_frame, text="MIR Comments:", font=("Arial", 9, "bold"))
        mir_comments_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 获取默认的 MIR Comments（从配置文件或文件）
        default_mir_comments = self.mole_history.get('mir_comments', '')
        if not default_mir_comments:
            # 尝试从 MIR Comments.txt 文件读取
            try:
                mir_comments_file = Path(__file__).parent.parent / "input" / "MIR Comments.txt"
                if mir_comments_file.exists():
                    with open(mir_comments_file, 'r', encoding='utf-8') as f:
                        default_mir_comments = f.read()
            except Exception as e:
                LOGGER.debug(f"读取MIR Comments文件失败: {e}")
        
        self.mir_comments_text = scrolledtext.ScrolledText(
            common_params_frame,
            height=4,  # 减小高度
            width=70,
            wrap=tk.WORD
        )
        self.mir_comments_text.pack(fill=tk.BOTH, expand=False, pady=(0, 5))  # 改为 expand=False
        self.mir_comments_text.insert('1.0', default_mir_comments)
        
        # --- VPOs特定参数 ---
        self.vpos_params_frame = ttk.LabelFrame(main_frame, text="VPOs 参数", padding="10")
        
        # Source Lot 文件选择
        source_lot_file_frame = ttk.Frame(self.vpos_params_frame)
        source_lot_file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(source_lot_file_frame, text="Source Lot 文件:", width=20).pack(side=tk.LEFT)
        
        # 获取保存的文件路径，如果没有则使用默认路径
        saved_source_lot_file = self.mole_history.get('source_lot_file', '')
        if not saved_source_lot_file:
            saved_source_lot_file = self.default_source_lot_file
        
        self.source_lot_file_var = tk.StringVar(value=saved_source_lot_file)
        source_lot_file_entry = ttk.Entry(source_lot_file_frame, textvariable=self.source_lot_file_var, width=45)
        source_lot_file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        def browse_source_lot_file():
            """浏览选择 Source Lot 文件"""
            initial_dir = Path(self.source_lot_file_var.get()).parent if self.source_lot_file_var.get() else None
            file_path = filedialog.askopenfilename(
                title="选择 Source Lot 文件",
                initialdir=str(initial_dir) if initial_dir and initial_dir.exists() else None,
                filetypes=[
                    ("CSV文件", "*.csv"),
                    ("Excel文件", "*.xlsx *.xls"),
                    ("所有文件", "*.*")
                ]
            )
            if file_path:
                self.source_lot_file_var.set(file_path)
        
        browse_button = ttk.Button(source_lot_file_frame, text="浏览...", command=browse_source_lot_file, width=10)
        browse_button.pack(side=tk.LEFT, padx=5)
        
        # 提示文字
        hint_label = ttk.Label(
            self.vpos_params_frame,
            text=f"提示：如果不选择文件，将使用默认路径: {self.default_source_lot_file}",
            font=("Arial", 8, "italic"),
            foreground="gray"
        )
        hint_label.pack(anchor=tk.W, pady=(5, 0))
        
        # --- Units特定参数 ---
        self.units_params_frame = ttk.LabelFrame(main_frame, text="Units 参数", padding="10")
        
        # Units 信息粘贴框
        units_label = ttk.Label(
            self.units_params_frame,
            text="粘贴 Units 信息:",
            font=("Arial", 9, "bold")
        )
        units_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 获取保存的 units 信息
        saved_units_info = self.mole_history.get('units_info', '')
        
        self.units_info_text = scrolledtext.ScrolledText(
            self.units_params_frame,
            height=5,  # 减小高度
            width=70,
            wrap=tk.WORD
        )
        self.units_info_text.pack(fill=tk.X, pady=(0, 5))  # 改为 fill=tk.X
        self.units_info_text.insert('1.0', saved_units_info)
        
        # 提示文字
        units_hint = ttk.Label(
            self.units_params_frame,
            text="提示：请粘贴要搜索的 Units 信息（每行一个，或使用其他格式）",
            font=("Arial", 8, "italic"),
            foreground="gray"
        )
        units_hint.pack(anchor=tk.W)
        
        # 初始显示相应的参数框
        self._on_search_mode_change()
        
        # --- 帮助信息 ---
        help_frame = ttk.Frame(main_frame)
        help_frame.pack(fill=tk.X, pady=(5, 10))
        
        help_text = ttk.Label(
            help_frame,
            text="提示：配置会自动保存，下次运行时将自动加载",
            font=("Arial", 8, "italic"),
            foreground="gray"
        )
        help_text.pack()
        
        # --- 按钮 ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 10))  # 增加上下边距，确保可见
        
        # 重置按钮（左侧）
        reset_btn = ttk.Button(
            button_frame,
            text="重置为默认值",
            command=self._on_reset,
            width=18
        )
        reset_btn.pack(side=tk.LEFT, padx=5)
        
        # 取消和确定按钮（右侧）
        cancel_btn = ttk.Button(
            button_frame,
            text="取消",
            command=self._on_cancel,
            width=15
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        ok_btn = ttk.Button(
            button_frame,
            text="确定",
            command=self._on_ok,
            width=15
        )
        ok_btn.pack(side=tk.RIGHT, padx=5)
        
        # 绑定 Enter 键到确定按钮
        self.root.bind('<Return>', lambda e: self._on_ok())
        self.root.bind('<Escape>', lambda e: self._on_cancel())
        
        # 居中显示窗口
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        # 运行事件循环
        self.root.mainloop()
        
        return self.result
    
    def _on_search_mode_change(self):
        """搜索方式改变时的处理"""
        mode = self.search_mode.get()
        
        # 隐藏所有特定参数框
        self.vpos_params_frame.pack_forget()
        self.units_params_frame.pack_forget()
        
        # 显示对应的参数框（在帮助信息之前）
        if mode == 'vpos':
            self.vpos_params_frame.pack(fill=tk.X, expand=False, pady=(0, 10))  # 改为 expand=False
        elif mode == 'units':
            self.units_params_frame.pack(fill=tk.X, expand=False, pady=(0, 10))  # 改为 expand=False
    
    def _on_ok(self):
        """确定按钮处理"""
        # 收集配置
        config = {
            'search_mode': self.search_mode.get(),
        }
        
        # 公共参数：MIR Comments
        mir_comments = self.mir_comments_text.get('1.0', tk.END).strip()
        if mir_comments:
            config['mir_comments'] = mir_comments
        
        # 根据搜索方式添加特定参数
        if config['search_mode'] == 'vpos':
            # VPOs 模式：Source Lot 文件路径
            source_lot_file = self.source_lot_file_var.get().strip()
            if source_lot_file:
                # 如果用户选择了文件，保存相对路径（相对于 auto-vpo 根目录）
                source_lot_path = Path(source_lot_file)
                if source_lot_path.is_absolute():
                    # 如果是绝对路径，尝试转换为相对路径
                    base_dir = Path(__file__).parent.parent  # workflow_automation -> auto-vpo
                    try:
                        relative_path = source_lot_path.relative_to(base_dir)
                        config['source_lot_file'] = str(relative_path).replace('\\', '/')
                    except ValueError:
                        # 如果无法转换为相对路径，保存绝对路径
                        config['source_lot_file'] = str(source_lot_path)
                else:
                    # 已经是相对路径，直接保存
                    config['source_lot_file'] = source_lot_file.replace('\\', '/')
            else:
                # 如果没有选择，使用默认路径（相对路径）
                config['source_lot_file'] = self.default_source_lot_file.replace('\\', '/')
        elif config['search_mode'] == 'units':
            # Units 模式：Units 信息
            units_info = self.units_info_text.get('1.0', tk.END).strip()
            if units_info:
                config['units_info'] = units_info
            else:
                messagebox.showwarning("警告", "请粘贴 Units 信息！")
                return  # 不关闭窗口，让用户继续输入
        
        # 保存到配置文件
        self._save_config(config)
        
        self.result = config
        self.root.destroy()
    
    def _on_cancel(self):
        """取消按钮处理"""
        self.result = None
        self.root.destroy()
    
    def _on_reset(self):
        """重置按钮处理"""
        # 清空所有输入
        self.mir_comments_text.delete('1.0', tk.END)
        self.source_lot_file_var.set(self.default_source_lot_file)
        self.units_info_text.delete('1.0', tk.END)
        self.search_mode.set('vpos')
        self._on_search_mode_change()


def show_mole_config_ui(config_path: Path) -> Optional[Dict]:
    """
    显示Mole配置UI并返回用户配置
    
    Args:
        config_path: config.yaml文件路径
    
    Returns:
        用户配置的字典，如果取消则返回None
    """
    ui = MoleConfigUI(config_path)
    return ui.show()


# 测试代码
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    # 测试UI
    test_config_path = Path(__file__).parent / "config.yaml"
    result = show_mole_config_ui(test_config_path)
    
    if result:
        print("用户配置:")
        for key, value in result.items():
            print(f"  {key}: {value}")
    else:
        print("用户取消了配置")

