import re
import csv
import os
import tkinter as tk
from tkinter import filedialog

# 弹出文件选择对话框
root = tk.Tk()
root.withdraw()
soc_file = filedialog.askopenfilename(title='请选择SOC文件', filetypes=[('SOC Files', '*.soc'), ('All Files', '*.*')])
if not soc_file:
    print('未选择文件，程序退出。')
    exit()

csv_file = os.path.splitext(soc_file)[0] + '.csv'

rows = []
current_resource = ''

def convert_value(val):
    # 只处理单一数字，其他格式直接跳过
    num_str = re.sub(r'\[.*?\]', '', val)
    num_str = num_str.strip()
    # 如果包含非数字、点、负号的字符，或包含冒号、逗号等，直接跳过
    if not re.fullmatch(r'-?\d+(\.\d+)?', num_str):
        return ''
    try:
        num = float(num_str)
        if num != 0 and num < 100:
            return str(int(round(num * 1000)))
        else:
            return ''
    except ValueError:
        return ''

with open(soc_file, 'r', encoding='utf-8') as f:
    for line in f:
        # Detect Resource group
        resource_match = re.match(r'\s*Resource\s+(\w+)', line)
        if resource_match:
            current_resource = resource_match.group(1)
        # Match pin name and value
        match = re.match(r'\s*(IP_[^\s]+)\s+([^\s;]+);', line)
        if match:
            pin, value = match.groups()
            value_converted = convert_value(value)
            rows.append([current_resource, pin, value, value_converted])

with open(csv_file, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['Resource', 'Pin Name', 'Value_raw', 'DISTINCTIVE_VALUE'])
    writer.writerows(rows)

print(f'Extracted {len(rows)} pins, result saved as {csv_file}')