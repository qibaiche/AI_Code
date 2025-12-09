import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_files():
    """
    选择需要的文件
    """
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 选择Leakage Excel文件
    messagebox.showinfo("选择文件", "请选择 包含limits的Pin_group 文件")
    leakage_file = filedialog.askopenfilename(
        title="选择 Leakage Excel 文件",
        filetypes=[
            ("Excel文件", "*.xlsx"),
            ("所有文件", "*.*")
        ],
        initialdir=os.path.expanduser("~/Desktop")
    )
    
    if not leakage_file:
        root.destroy()
        return None, None
    
    # 选择PinDefinitions CSV文件
    messagebox.showinfo("选择文件", "请选择PinDefinitions.csv 文件")
    pin_def_file = filedialog.askopenfilename(
        title="选择 PinDefinitions CSV 文件",
        filetypes=[
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*")
        ],
        initialdir=os.path.dirname(leakage_file) if leakage_file else os.path.expanduser("~/Desktop")
    )
    
    root.destroy()
    return leakage_file, pin_def_file

def match_pin_groups(leakage_file, pin_def_file, output_file=None):
    """
    根据Leakage文件中的Pin_Group匹配PinDefinitions文件中的对应行
    
    Args:
        leakage_file (str): Leakage Excel文件路径
        pin_def_file (str): PinDefinitions CSV文件路径
        output_file (str): 输出Excel文件路径，如果为None则自动生成
    """
    
    try:
        # 读取Leakage Excel文件
        print("正在读取 Leakage Excel 文件...")
        leakage_df = pd.read_excel(leakage_file)
        
        # 检查是否存在Pin_Group列
        if 'Pin_Group' not in leakage_df.columns:
            messagebox.showerror("错误", "Leakage Excel文件中未找到 'Pin_Group' 列")
            return False
        
        # 获取所有唯一的Pin_Group值
        unique_pin_groups = leakage_df['Pin_Group'].dropna().unique()
        print(f"从 Leakage 文件中找到 {len(unique_pin_groups)} 个唯一的 Pin_Group")
        
        # 读取PinDefinitions CSV文件
        print("正在读取 PinDefinitions CSV 文件...")
        pin_def_df = pd.read_csv(pin_def_file)
        
        # 检查是否存在Pin_Group列
        if 'Pin_Group' not in pin_def_df.columns:
            messagebox.showerror("错误", "PinDefinitions CSV文件中未找到 'Pin_Group' 列")
            return False
        
        # 根据Pin_Group匹配并筛选数据
        print("正在匹配 Pin_Group...")
        matched_df = pin_def_df[pin_def_df['Pin_Group'].isin(unique_pin_groups)]
        
        if matched_df.empty:
            messagebox.showwarning("警告", "未找到任何匹配的 Pin_Group")
            return False
        
        print(f"匹配到 {len(matched_df)} 行数据")
        
        # 根据PinDefinitions文件名自动检测前缀类型
        pin_def_filename = os.path.basename(pin_def_file).upper()
        if 'IP_CPU' in pin_def_filename:
            chosen_prefix = 'IP_CPU::'
            print("检测到文件名包含'IP_CPU'，自动选择前缀: IP_CPU::")
        elif 'IP_CD' in pin_def_filename:
            chosen_prefix = 'IP_CD::'
            print("检测到文件名包含'IP_CD'，自动选择前缀: IP_CD::")
        else:
            # 如果文件名中没有明确标识，默认使用IP_CD::
            chosen_prefix = 'IP_CD::'
            print("文件名中未检测到明确的前缀标识，默认使用: IP_CD::")
        
        # 在Pin_Name列前面添加检测到的前缀
        if 'Pin_Name' in matched_df.columns:
            matched_df['Pin_Name'] = chosen_prefix + matched_df['Pin_Name'].astype(str)
            print(f"已在Pin_Name列前添加'{chosen_prefix}'前缀")
        else:
            print("警告：未找到Pin_Name列")
        
        # 生成输出文件名，在输入文件名基础上加上"filtered"
        if output_file is None:
            base_dir = os.path.dirname(pin_def_file)
            # 获取原始PinDefinitions文件名（含扩展名）
            pin_def_basename = os.path.basename(pin_def_file)
            # 在文件名和扩展名之间插入"filtered"
            name_without_ext = os.path.splitext(pin_def_basename)[0]
            ext = os.path.splitext(pin_def_basename)[1]
            output_file = os.path.join(base_dir, f"{name_without_ext}_filtered{ext}")
        
        # 保存到CSV文件
        print("正在保存结果...")
        matched_df.to_csv(output_file, index=False, encoding='utf-8-sig')
        
        # 创建统计信息CSV文件，文件名与主文件保持一致
        stats_file = os.path.join(os.path.dirname(output_file), f"{os.path.splitext(os.path.basename(output_file))[0]}_Statistics.csv")
        stats_data = {
            'Pin_Group': list(unique_pin_groups),
            'Count_in_PinDefinitions': [len(matched_df[matched_df['Pin_Group'] == pg]) for pg in unique_pin_groups]
        }
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_csv(stats_file, index=False, encoding='utf-8-sig')
        
        # 显示结果统计
        total_matched_groups = len(matched_df['Pin_Group'].unique())
        total_rows = len(matched_df)
        
        result_message = f"""匹配完成！
        
结果统计：
- 原始 Leakage 文件中的唯一 Pin_Group 数量: {len(unique_pin_groups)}
- 在 PinDefinitions 中找到匹配的 Pin_Group 数量: {total_matched_groups}
- 总匹配行数: {total_rows}

输出文件：
1. 主数据文件: {output_file}
2. 统计信息文件: {stats_file}

注意：Pin_Name 列已自动添加 '{chosen_prefix}' 前缀
"""
        
        messagebox.showinfo("成功", result_message)
        print("匹配完成！")
        return True
        
    except FileNotFoundError as e:
        messagebox.showerror("错误", f"找不到文件：{e}")
        return False
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时出错：{e}")
        return False

def main():
    """
    主函数
    """
    print("Pin Group 匹配工具")
    print("=" * 50)
    
    # 选择文件
    leakage_file, pin_def_file = select_files()
    
    if not leakage_file or not pin_def_file:
        print("未选择文件，程序退出")
        return
    
    # 检查文件是否存在
    if not os.path.exists(leakage_file):
        messagebox.showerror("错误", f"Leakage文件不存在：{leakage_file}")
        return
    
    if not os.path.exists(pin_def_file):
        messagebox.showerror("错误", f"PinDefinitions文件不存在：{pin_def_file}")
        return
    
    # 执行匹配
    success = match_pin_groups(leakage_file, pin_def_file)
    
    if success:
        print("处理完成！")
    else:
        print("处理失败！")

if __name__ == "__main__":
    main() 