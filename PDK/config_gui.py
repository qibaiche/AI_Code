import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
from typing import Dict, Any
from dataclasses import dataclass, asdict
import sys

@dataclass
class EmailConfig:
    """邮件配置类"""
    recipients_file: str = ""
    debug_email: str = "qibai.chen@intel.com"
    retry_times: int = 3
    retry_delay: int = 2
    max_emails_per_batch: int = 50

@dataclass
class PathConfig:
    """路径配置类"""
    source_path: str = r"\\atdfile3.ch.intel.com\PRO\Reports"
    target_path: str = ""
    output_path: str = ""
    owner_file: str = ""

@dataclass
class ColumnConfig:
    """列配置类"""
    relevant_columns: list = None
    pivot_columns: list = None
    
    def __post_init__(self):
        if self.relevant_columns is None:
            self.relevant_columns = ['ProgramName', 'ModelName', 'CCB', 'Package', 'Device', 'Revision', 'Stepping', 'DieCodeName', 'TestName']
        if self.pivot_columns is None:
            self.pivot_columns = ['DieCodeName', 'ProgramName', 'CCB', 'Package', 'Device', 'Revision', 'Stepping']

class PDKConfigGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDK周报配置工具")
        self.root.geometry("800x700")
        
        # 初始化配置
        self.email_config = EmailConfig()
        self.path_config = PathConfig()
        self.column_config = ColumnConfig()
        
        # 创建界面
        self.create_widgets()
        
        # 尝试加载现有配置
        self.load_config()
    
    def create_widgets(self):
        """创建界面组件"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置根窗口的网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # 标题
        title_label = ttk.Label(main_frame, text="PDK周报生成工具 - 配置设置", font=("Arial", 16, "bold"))
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # 路径配置区域
        path_frame = ttk.LabelFrame(main_frame, text="📁 路径配置", padding="10")
        path_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        path_frame.columnconfigure(1, weight=1)
        row += 1
        
        # 源文件路径
        ttk.Label(path_frame, text="源文件路径:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.source_path_var = tk.StringVar(value=self.path_config.source_path)
        ttk.Entry(path_frame, textvariable=self.source_path_var, width=60).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(path_frame, text="浏览", command=lambda: self.browse_folder(self.source_path_var)).grid(row=0, column=2, pady=2)
        
        # 目标路径
        ttk.Label(path_frame, text="目标路径:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.target_path_var = tk.StringVar(value=self.path_config.target_path)
        ttk.Entry(path_frame, textvariable=self.target_path_var, width=60).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(path_frame, text="浏览", command=lambda: self.browse_folder(self.target_path_var)).grid(row=1, column=2, pady=2)
        
        # 输出路径
        ttk.Label(path_frame, text="输出路径:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.output_path_var = tk.StringVar(value=self.path_config.output_path)
        ttk.Entry(path_frame, textvariable=self.output_path_var, width=60).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(path_frame, text="浏览", command=lambda: self.browse_folder(self.output_path_var)).grid(row=2, column=2, pady=2)
        
        # 产品负责人文件
        ttk.Label(path_frame, text="产品负责人文件:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.owner_file_var = tk.StringVar(value=self.path_config.owner_file)
        ttk.Entry(path_frame, textvariable=self.owner_file_var, width=60).grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(path_frame, text="浏览", command=lambda: self.browse_file(self.owner_file_var, "Excel文件", "*.xlsx")).grid(row=3, column=2, pady=2)
        
        # 邮件配置区域
        email_frame = ttk.LabelFrame(main_frame, text="📧 邮件配置", padding="10")
        email_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        email_frame.columnconfigure(1, weight=1)
        row += 1
        
        # 邮件收件人文件
        ttk.Label(email_frame, text="邮件收件人文件:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.recipients_file_var = tk.StringVar(value=self.email_config.recipients_file)
        ttk.Entry(email_frame, textvariable=self.recipients_file_var, width=60).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(email_frame, text="浏览", command=lambda: self.browse_file(self.recipients_file_var, "文本文件", "*.txt")).grid(row=0, column=2, pady=2)
        
        # 调试邮箱
        ttk.Label(email_frame, text="调试邮箱:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.debug_email_var = tk.StringVar(value=self.email_config.debug_email)
        ttk.Entry(email_frame, textvariable=self.debug_email_var, width=60).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        
        # 邮件参数
        email_params_frame = ttk.Frame(email_frame)
        email_params_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(email_params_frame, text="重试次数:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.retry_times_var = tk.IntVar(value=self.email_config.retry_times)
        ttk.Spinbox(email_params_frame, from_=1, to=10, textvariable=self.retry_times_var, width=10).grid(row=0, column=1, padx=(0, 20))
        
        ttk.Label(email_params_frame, text="重试延迟(秒):").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.retry_delay_var = tk.IntVar(value=self.email_config.retry_delay)
        ttk.Spinbox(email_params_frame, from_=1, to=60, textvariable=self.retry_delay_var, width=10).grid(row=0, column=3, padx=(0, 20))
        
        ttk.Label(email_params_frame, text="最大批量邮件数:").grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        self.max_emails_var = tk.IntVar(value=self.email_config.max_emails_per_batch)
        ttk.Spinbox(email_params_frame, from_=1, to=200, textvariable=self.max_emails_var, width=10).grid(row=0, column=5)
        
        # 列配置区域
        column_frame = ttk.LabelFrame(main_frame, text="📊 列配置", padding="10")
        column_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        column_frame.columnconfigure(1, weight=1)
        row += 1
        
        # 相关列配置
        ttk.Label(column_frame, text="比较相关列:").grid(row=0, column=0, sticky=(tk.W, tk.N), pady=2)
        self.relevant_columns_var = tk.StringVar(value=", ".join(self.column_config.relevant_columns))
        relevant_text = tk.Text(column_frame, height=3, width=60)
        relevant_text.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        relevant_text.insert("1.0", self.relevant_columns_var.get())
        self.relevant_text = relevant_text
        
        # 透视表列配置
        ttk.Label(column_frame, text="透视表列:").grid(row=1, column=0, sticky=(tk.W, tk.N), pady=2)
        self.pivot_columns_var = tk.StringVar(value=", ".join(self.column_config.pivot_columns))
        pivot_text = tk.Text(column_frame, height=3, width=60)
        pivot_text.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        pivot_text.insert("1.0", self.pivot_columns_var.get())
        self.pivot_text = pivot_text
        
        # 帮助文本
        help_text = "提示：列名请用逗号分隔，例如: ProgramName, ModelName, CCB"
        ttk.Label(column_frame, text=help_text, foreground="gray").grid(row=2, column=1, sticky=tk.W, pady=(0, 5))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="💾 保存配置", command=self.save_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="📂 加载配置", command=self.load_config_dialog).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="🔄 重置为默认", command=self.reset_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="✅ 应用并关闭", command=self.apply_and_close).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="❌ 取消", command=self.root.destroy).pack(side=tk.LEFT)
    
    def browse_folder(self, var):
        """浏览文件夹"""
        folder = filedialog.askdirectory(initialdir=var.get() or os.getcwd())
        if folder:
            var.set(folder)
    
    def browse_file(self, var, description, pattern):
        """浏览文件"""
        file = filedialog.askopenfilename(
            initialdir=os.path.dirname(var.get()) or os.getcwd(),
            title=f"选择{description}",
            filetypes=[(description, pattern), ("所有文件", "*.*")]
        )
        if file:
            var.set(file)
    
    def get_current_config(self) -> Dict[str, Any]:
        """获取当前配置"""
        # 更新列配置
        relevant_cols = [col.strip() for col in self.relevant_text.get("1.0", tk.END).strip().split(",") if col.strip()]
        pivot_cols = [col.strip() for col in self.pivot_text.get("1.0", tk.END).strip().split(",") if col.strip()]
        
        config = {
            "email": {
                "recipients_file": self.recipients_file_var.get(),
                "debug_email": self.debug_email_var.get(),
                "retry_times": self.retry_times_var.get(),
                "retry_delay": self.retry_delay_var.get(),
                "max_emails_per_batch": self.max_emails_var.get()
            },
            "paths": {
                "source_path": self.source_path_var.get(),
                "target_path": self.target_path_var.get(),
                "output_path": self.output_path_var.get(),
                "owner_file": self.owner_file_var.get()
            },
            "columns": {
                "relevant_columns": relevant_cols,
                "pivot_columns": pivot_cols
            }
        }
        return config
    
    def set_config(self, config: Dict[str, Any]):
        """设置配置"""
        # 邮件配置
        if "email" in config:
            email = config["email"]
            self.recipients_file_var.set(email.get("recipients_file", ""))
            self.debug_email_var.set(email.get("debug_email", "qibai.chen@intel.com"))
            self.retry_times_var.set(email.get("retry_times", 3))
            self.retry_delay_var.set(email.get("retry_delay", 2))
            self.max_emails_var.set(email.get("max_emails_per_batch", 50))
        
        # 路径配置
        if "paths" in config:
            paths = config["paths"]
            self.source_path_var.set(paths.get("source_path", ""))
            self.target_path_var.set(paths.get("target_path", ""))
            self.output_path_var.set(paths.get("output_path", ""))
            self.owner_file_var.set(paths.get("owner_file", ""))
        
        # 列配置
        if "columns" in config:
            columns = config["columns"]
            relevant_cols = columns.get("relevant_columns", [])
            pivot_cols = columns.get("pivot_columns", [])
            
            self.relevant_text.delete("1.0", tk.END)
            self.relevant_text.insert("1.0", ", ".join(relevant_cols))
            
            self.pivot_text.delete("1.0", tk.END)
            self.pivot_text.insert("1.0", ", ".join(pivot_cols))
    
    def validate_config(self) -> bool:
        """验证配置"""
        config = self.get_current_config()
        
        # 检查必要的路径
        required_paths = {
            "源文件路径": config["paths"]["source_path"],
            "目标路径": config["paths"]["target_path"],
            "输出路径": config["paths"]["output_path"]
        }
        
        for name, path in required_paths.items():
            if not path:
                messagebox.showerror("配置错误", f"请设置{name}")
                return False
        
        # 检查邮件配置
        if not config["email"]["recipients_file"]:
            messagebox.showerror("配置错误", "请设置邮件收件人文件")
            return False
        
        if not config["email"]["debug_email"]:
            messagebox.showerror("配置错误", "请设置调试邮箱")
            return False
        
        # 检查列配置
        if not config["columns"]["relevant_columns"]:
            messagebox.showerror("配置错误", "请设置比较相关列")
            return False
        
        if not config["columns"]["pivot_columns"]:
            messagebox.showerror("配置错误", "请设置透视表列")
            return False
        
        return True
    
    def save_config(self):
        """保存配置"""
        if not self.validate_config():
            return
        
        config = self.get_current_config()
        
        # 选择保存位置
        config_file = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            title="保存配置文件"
        )
        
        if config_file:
            try:
                with open(config_file, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, ensure_ascii=False)
                messagebox.showinfo("成功", f"配置已保存到: {config_file}")
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def load_config(self):
        """加载默认配置"""
        # 尝试加载当前目录的配置文件
        config_file = "pdk_config.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.set_config(config)
            except Exception as e:
                print(f"加载默认配置失败: {str(e)}")
    
    def load_config_dialog(self):
        """加载配置对话框"""
        config_file = filedialog.askopenfilename(
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            title="加载配置文件"
        )
        
        if config_file:
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.set_config(config)
                messagebox.showinfo("成功", f"配置已加载: {config_file}")
            except Exception as e:
                messagebox.showerror("错误", f"加载配置失败: {str(e)}")
    
    def reset_config(self):
        """重置为默认配置"""
        if messagebox.askyesno("确认", "确定要重置为默认配置吗？"):
            self.email_config = EmailConfig()
            self.path_config = PathConfig()
            self.column_config = ColumnConfig()
            
            # 重新设置界面
            self.set_config({
                "email": asdict(self.email_config),
                "paths": asdict(self.path_config),
                "columns": asdict(self.column_config)
            })
    
    def apply_and_close(self):
        """应用配置并关闭"""
        if not self.validate_config():
            return
        
        config = self.get_current_config()
        
        # 保存配置到默认文件
        try:
            with open("pdk_config.json", 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
            return
        
        # 生成配置代码
        self.generate_config_code(config)
        
        messagebox.showinfo("成功", "配置已应用并保存到 pdk_config.json")
        self.root.destroy()
    
    def generate_config_code(self, config):
        """生成配置代码"""
        code_template = f'''# 自动生成的配置文件
# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

@dataclass
class EmailConfig:
    recipients_file: str = r"{config['email']['recipients_file']}"
    debug_email: str = "{config['email']['debug_email']}"
    retry_times: int = {config['email']['retry_times']}
    retry_delay: int = {config['email']['retry_delay']}
    max_emails_per_batch: int = {config['email']['max_emails_per_batch']}

@dataclass
class PathConfig:
    source_path: str = r"{config['paths']['source_path']}"
    target_path: str = r"{config['paths']['target_path']}"
    output_path: str = r"{config['paths']['output_path']}"
    owner_file: str = r"{config['paths']['owner_file']}"

class PDKReportConfig:
    def __init__(self):
        self.email = EmailConfig()
        self.paths = PathConfig()
        self.relevant_columns = {config['columns']['relevant_columns']}
        self.pivot_columns = {config['columns']['pivot_columns']}

# 全局配置实例
CONFIG = PDKReportConfig()
'''
        
        # 保存到配置代码文件
        with open("generated_config.py", 'w', encoding='utf-8') as f:
            f.write(code_template)

def main():
    """主函数"""
    root = tk.Tk()
    app = PDKConfigGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 