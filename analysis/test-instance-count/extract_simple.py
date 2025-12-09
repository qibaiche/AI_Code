#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单MTPL提取工具
只提取Test PrimeDcLeakageTestMethod的实例名和BypassPort
"""

import re
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_file():
    """文件选择对话框"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    file_path = filedialog.askopenfilename(
        title="选择MTPL文件",
        filetypes=[
            ("MTPL文件", "*.mtpl"),
            ("所有文件", "*.*")
        ],
        initialdir=os.getcwd()
    )
    
    root.destroy()
    return file_path if file_path else None

def main():
    print("🚀 简单MTPL提取工具 (支持PrimeDcLeakageTestMethod和TraceAnalyticsDcLeakage)")
    print("📁 请选择MTPL文件...")
    
    # 文件选择
    input_file = select_file()
    
    if not input_file:
        print("❌ 未选择文件，程序退出")
        return
    
    print(f"📁 已选择文件：{os.path.basename(input_file)}")
    
    # 生成输出文件名（与输入文件在同一路径）
    input_dir = os.path.dirname(input_file)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(input_dir, f"{base_name}_Instances.csv")
    
    # 检查文件是否存在
    if not os.path.exists(input_file):
        error_msg = f"文件不存在：{input_file}"
        print(f"❌ {error_msg}")
        messagebox.showerror("错误", error_msg)
        return
    
    # 读取文件
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        print(f"✅ 文件读取成功")
    except Exception as e:
        error_msg = f"文件读取失败: {e}"
        print(f"❌ {error_msg}")
        messagebox.showerror("错误", error_msg)
        return
    
    # 使用最可靠的方法：按行处理
    lines = content.splitlines()
    print(f"📄 文件共 {len(lines)} 行")
    
    # 找到所有DC泄漏测试实例的起始行（支持多种类型）
    test_lines = []
    prime_count = 0
    trace_count = 0
    
    for i, line in enumerate(lines):
        # 匹配 PrimeDcLeakageTestMethod
        if re.match(r'Test\s+PrimeDcLeakageTestMethod\s+\S+', line):
            test_lines.append((i, line, 'PrimeDcLeakageTestMethod'))
            prime_count += 1
        # 匹配 TraceAnalyticsDcLeakage
        elif re.match(r'Test\s+TraceAnalyticsDcLeakage\s+\S+', line):
            test_lines.append((i, line, 'TraceAnalyticsDcLeakage'))
            trace_count += 1
    
    print(f"🔍 找到测试实例总数: {len(test_lines)}")
    print(f"  • PrimeDcLeakageTestMethod: {prime_count} 个")
    print(f"  • TraceAnalyticsDcLeakage: {trace_count} 个")
    
    if not test_lines:
        print("❌ 未找到任何实例")
        return
    
    # 为每个实例提取信息
    instances_data = []
    
    for line_num, line, test_type in test_lines:
        # 提取实例名（根据测试类型使用不同的正则表达式）
        if test_type == 'PrimeDcLeakageTestMethod':
            name_match = re.search(r'Test\s+PrimeDcLeakageTestMethod\s+(\S+)', line)
        elif test_type == 'TraceAnalyticsDcLeakage':
            name_match = re.search(r'Test\s+TraceAnalyticsDcLeakage\s+(\S+)', line)
        else:
            continue
            
        if name_match:
            name = name_match.group(1)
            
            # 从当前行开始向下查找BypassPort（通常在接下来的10-20行内）
            bypass_port = 'UNKNOWN'
            for i in range(line_num + 1, min(line_num + 30, len(lines))):
                if 'BypassPort' in lines[i]:
                    bypass_match = re.search(r'BypassPort\s*=\s*([^;]+);', lines[i])
                    if bypass_match:
                        bypass_port = bypass_match.group(1).strip()
                        break
                # 如果遇到下一个Test，停止搜索
                if lines[i].strip().startswith('Test '):
                    break
            
            instances_data.append({
                'Instance_Name': name,
                'Test_Method': test_type,
                'BypassPort': bypass_port
            })
    
    # 保存到CSV
    try:
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            fieldnames = ['Instance_Name', 'Test_Method', 'BypassPort']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(instances_data)
        
        print(f"💾 结果已保存到：{output_file}")
        print(f"📊 总共提取了 {len(instances_data)} 个实例")
        
        # 统计分布
        bypass_1_count = sum(1 for item in instances_data if item['BypassPort'] == '1')
        bypass_neg1_count = sum(1 for item in instances_data if item['BypassPort'] == '-1')
        other_count = len(instances_data) - bypass_1_count - bypass_neg1_count
        
        prime_extracted = sum(1 for item in instances_data if item['Test_Method'] == 'PrimeDcLeakageTestMethod')
        trace_extracted = sum(1 for item in instances_data if item['Test_Method'] == 'TraceAnalyticsDcLeakage')
        
        print(f"\n📊 测试类型分布：")
        print(f"  PrimeDcLeakageTestMethod: {prime_extracted} 个")
        print(f"  TraceAnalyticsDcLeakage: {trace_extracted} 个")
        
        print(f"\n📊 BypassPort分布：")
        print(f"  BypassPort = 1:  {bypass_1_count} 个")
        print(f"  BypassPort = -1: {bypass_neg1_count} 个")
        if other_count > 0:
            print(f"  其他值: {other_count} 个")
        
        # 显示前5个示例
        print(f"\n🔍 前5个实例示例：")
        for i, item in enumerate(instances_data[:5], 1):
            print(f"  {i}. {item['Instance_Name']}")
            print(f"     测试类型: {item['Test_Method']}")
            print(f"     BypassPort: {item['BypassPort']}")
        
        print("🎉 提取完成!")
        
        # 显示成功消息框
        success_msg = f"""提取完成！

📊 统计信息：
• 总实例数：{len(instances_data)} 个
• PrimeDcLeakageTestMethod：{prime_extracted} 个
• TraceAnalyticsDcLeakage：{trace_extracted} 个
• BypassPort = 1：{bypass_1_count} 个
• BypassPort = -1：{bypass_neg1_count} 个

📁 输出文件：
• {os.path.basename(output_file)}
• 保存路径：{os.path.dirname(output_file)}
• 包含字段：Instance_Name, Test_Method, BypassPort
"""
        messagebox.showinfo("提取完成", success_msg)
        
    except Exception as e:
        error_msg = f"保存失败: {e}"
        print(f"❌ {error_msg}")
        messagebox.showerror("错误", error_msg)

if __name__ == "__main__":
    main() 