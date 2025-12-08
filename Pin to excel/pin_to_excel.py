import re
import csv
import os
import tkinter as tk
from tkinter import filedialog

# 弹出文件选择对话框
root = tk.Tk()
root.withdraw()
pin_file = filedialog.askopenfilename(title='请选择PIN文件', filetypes=[('PIN Files', '*.pin'), ('All Files', '*.*')])
if not pin_file:
    print('未选择文件，程序退出。')
    exit()

csv_file = os.path.splitext(pin_file)[0] + '.csv'

rows = []
current_group = ''
in_group = False

with open(pin_file, 'r', encoding='utf-8') as f:
    for line in f:
        original_line = line
        line = line.strip()
        
        # 跳过空行
        if not line:
            continue
            
        # 检测Group开始
        group_match = re.match(r'Group\s+(.+)', line)
        if group_match:
            current_group = group_match.group(1)
            in_group = True
            continue
            
        # 检测Group结束（遇到}）
        if '}' in line and in_group:
            in_group = False
            current_group = ''
            continue
            
        # 只提取Group内的引脚名称
        if in_group:
            pin_match = re.match(r'\s*(#?[A-Z][A-Z0-9_]*(?:__[A-Z]+)?)[;,]?\s*$', line)
            
            if pin_match:
                pin_name = pin_match.group(1)
                # 只保存 Group → Pin
                rows.append([current_group, pin_name])


# 写入CSV文件
with open(csv_file, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['Pin_Group', 'Pin_Name'])
    writer.writerows(rows)

print(f'Extracted {len(rows)} pins from groups, result saved as {csv_file}')

# 统计信息
groups = []
seen_groups = set()
for row in rows:
    if row[0] and row[0] not in seen_groups:
        groups.append(row[0])
        seen_groups.add(row[0])

print(f'Found Groups (in order): {groups}')

# 显示每个Group下的引脚数量
for group in groups:
    count = len([row for row in rows if row[0] == group])
    print(f'  {group}: {count} pins') 