#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格合并脚本
基于SIO_BSCAN_PCD_4JMP表，使用Test_Type和Configuration去匹配Leakage_LIMIT_COLD表，生成新表
"""

from datetime import datetime
from pathlib import Path

import pandas as pd

def merge_tables():
    """合并两个表格的主函数"""
    
    print("=== 表格合并工具 ===")
    print("基于SIO_BSCAN_PCD_4JMP表，使用Test_Type和Configuration匹配Leakage_LIMIT_COLD表\n")
    
    # 定义文件路径
    repo_root = Path(__file__).resolve().parents[2]
    base_dir = repo_root / "data" / "leakage-conjunction"
    sio_file = base_dir / "SIO_BSCAN_PCD_4JMP.xlsx"
    leakage_file = base_dir / "Leakage_LIMIT_COLD.xlsx"
    
    # 检查文件是否存在
    if not sio_file.exists():
        print(f"❌ 文件不存在: {sio_file}")
        print(f"当前目录: {Path.cwd()}")
        return False

    if not leakage_file.exists():
        print(f"❌ 文件不存在: {leakage_file}")
        print(f"当前目录: {Path.cwd()}")
        return False
    
    try:
        # 读取Excel文件
        print("📖 正在读取SIO_BSCAN_PCD_4JMP表...")
        sio_df = pd.read_excel(sio_file)
        print(f"   - 成功读取 {len(sio_df)} 行，{len(sio_df.columns)} 列")
        
        print("📖 正在读取Leakage_LIMIT_COLD表...")
        leakage_df = pd.read_excel(leakage_file)
        print(f"   - 成功读取 {len(leakage_df)} 行，{len(leakage_df.columns)} 列")
        
        # 显示列名
        print(f"\nSIO表列名: {list(sio_df.columns)}")
        print(f"Leakage表列名: {list(leakage_df.columns)}")
        
        # 查找匹配字段
        sio_columns = [str(col).strip() for col in sio_df.columns]
        leakage_columns = [str(col).strip() for col in leakage_df.columns]
        
        # 查找Test_Type字段
        test_type_sio = None
        test_type_leakage = None
        
        for col in sio_columns:
            if 'test_type' in col.lower() or 'testtype' in col.lower():
                test_type_sio = col
                break
        
        for col in leakage_columns:
            if 'test_type' in col.lower() or 'testtype' in col.lower():
                test_type_leakage = col
                break
        
        # 查找Configuration字段
        config_sio = None
        config_leakage = None
        
        for col in sio_columns:
            if 'configuration' in col.lower() or 'config' in col.lower():
                config_sio = col
                break
        
        for col in leakage_columns:
            if 'configuration' in col.lower() or 'config' in col.lower():
                config_leakage = col
                break
        
        print(f"\n🔍 匹配字段:")
        print(f"   SIO表 Test_Type: {test_type_sio}")
        print(f"   SIO表 Configuration: {config_sio}")
        print(f"   Leakage表 Test_Type: {test_type_leakage}")
        print(f"   Leakage表 Configuration: {config_leakage}")
        
        # 如果找不到精确匹配，显示所有列供用户参考
        if not all([test_type_sio, config_sio, test_type_leakage, config_leakage]):
            print("\n❌ 未找到所有必要的匹配字段")
            print("\n📋 SIO表前3行数据:")
            print(sio_df.head(3))
            print("\n📋 Leakage表前3行数据:")
            print(leakage_df.head(3))
            
            # 尝试直接使用列名进行合并（如果用户确认列名正确）
            print("\n💡 如果您确认列名正确，可以手动指定匹配字段")
            return False
        
        # 执行合并
        print(f"\n🔗 正在合并表格...")
        
        # 如果字段名不同，先重命名
        rename_dict = {}
        if test_type_leakage != test_type_sio:
            rename_dict[test_type_leakage] = test_type_sio
        if config_leakage != config_sio:
            rename_dict[config_leakage] = config_sio
        
        if rename_dict:
            leakage_df = leakage_df.rename(columns=rename_dict)
            print(f"   - 重命名Leakage表字段: {rename_dict}")
        
        # 执行左连接合并
        merged_df = pd.merge(
            sio_df,
            leakage_df,
            on=[test_type_sio, config_sio],
            how='left',
            suffixes=('', '_Leakage')
        )
        
        print(f"   - 合并完成，新表包含 {len(merged_df)} 行，{len(merged_df.columns)} 列")
        
        # 计算匹配统计
        # 检查有多少行成功匹配了Leakage数据
        leakage_cols = [col for col in merged_df.columns if col.endswith('_Leakage')]
        if leakage_cols:
            matched_rows = merged_df[leakage_cols].notna().any(axis=1).sum()
        else:
            # 如果没有_Leakage后缀的列，说明可能有重名列被覆盖
            matched_rows = len(merged_df) - merged_df[[test_type_sio, config_sio]].isna().any(axis=1).sum()
        
        print(f"   - 匹配率: {matched_rows}/{len(sio_df)} ({matched_rows/len(sio_df)*100:.1f}%)")
        
        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"merged_table_{timestamp}.xlsx"
        
        # 保存结果
        print(f"\n💾 正在保存到: {output_file}")
        
        # 创建ExcelWriter对象以自定义格式
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='合并结果')
            
            # 获取工作表对象
            worksheet = writer.sheets['合并结果']
            
            # 设置数字格式为普通数字显示（不使用科学计数法）
            from openpyxl.styles import NamedStyle
            
            # 创建数字格式样式
            number_style = NamedStyle(name="number_style")
            number_style.number_format = '0.000000000'  # 显示9位小数
            
            # 应用格式到所有数值列
            for col_idx, column in enumerate(merged_df.columns, 1):
                # 检查列是否包含数值数据
                if merged_df[column].dtype in ['float64', 'int64', 'float32', 'int32']:
                    # 对该列的所有单元格应用数字格式
                    for row_idx in range(2, len(merged_df) + 2):  # 从第2行开始（跳过标题）
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.number_format = '0.000000000'
            
            print("   - 已设置数字格式为普通显示（非科学计数法）")
        
        print(f"\n✅ 合并成功完成！")
        print(f"   - 输出文件: {output_file}")
        print(f"   - 总行数: {len(merged_df)}")
        print(f"   - 总列数: {len(merged_df.columns)}")
        print(f"   - 匹配成功: {matched_rows} 行")
        
        # 显示前几行预览
        print(f"\n📋 合并结果预览 (前3行):")
        print(merged_df.head(3).to_string())
        
        return True
        
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # 检查并安装必要的包
    try:
        import pandas as pd
    except ImportError:
        print("正在安装pandas...")
        import subprocess
        import sys
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
        import pandas as pd
    
    # 运行合并函数
    success = merge_tables()
    
    if not success:
        print("\n合并失败，请检查文件路径和数据格式")
    
    input("\n按Enter键退出...") 
