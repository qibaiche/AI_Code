import pandas as pd
import re
import os
from datetime import datetime
from tqdm import tqdm  # 添加tqdm库导入

def parse_test_result(result_str):
    """解析TEST_RESULT字符串，提取PRE和MAIN部分的时间"""
    try:
        # 使用正则表达式匹配时间部分
        pattern = r'PRE_(\d+\.\d+)MS_MAIN_(\d+\.\d+)MS_ET_(\d+\.\d+)MS'
        match = re.search(pattern, result_str)
        
        if match:
            pre_time = float(match.group(1))
            main_time = float(match.group(2))
            # 不再提取ET_TIME
            total_time = pre_time + main_time
            return pre_time, main_time, total_time
        else:
            return None, None, None
    except Exception as e:
        print(f"解析错误: {e}, 字符串: {result_str}")
        return None, None, None

def read_data_file(file_path):
    """
    根据文件扩展名读取数据文件
    
    Args:
        file_path: 文件路径
    
    Returns:
        pandas DataFrame
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.xlsx' or file_ext == '.xls':
        print(f"正在读取Excel文件: {file_path}")
        return pd.read_excel(file_path)
    elif file_ext == '.csv':
        print(f"正在读取CSV文件: {file_path}")
        # 尝试用不同编码读取CSV文件
        for encoding in ['utf-8', 'gbk', 'latin1']:
            try:
                # 先读取少量数据检测分隔符
                sample = pd.read_csv(file_path, nrows=5, encoding=encoding)
                # 如果成功，读取完整文件
                print(f"使用{encoding}编码成功读取CSV文件")
                return pd.read_csv(file_path, encoding=encoding)
            except Exception as e:
                print(f"使用{encoding}编码读取CSV文件失败: {e}")
                continue
        # 如果所有编码都失败，尝试自动检测分隔符
        try:
            return pd.read_csv(file_path, sep=None, engine='python')
        except Exception as e:
            print(f"自动检测分隔符读取CSV文件失败: {e}")
            raise ValueError(f"无法读取CSV文件: {file_path}")
    else:
        raise ValueError(f"不支持的文件格式: {file_ext}，请提供Excel或CSV格式的文件")

def analyze_test_times(input_file, output_dir=None):
    """
    分析测试时间数据并生成报告
    
    Args:
        input_file: 输入文件路径(Excel或CSV)
        output_dir: 输出目录，默认为输入文件所在目录
    """
    try:
        # 设置输出目录
        if output_dir is None:
            output_dir = os.path.dirname(input_file)
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取数据文件
        df = read_data_file(input_file)
        
        # 检查必要的列是否存在
        required_columns = ['VISUAL_ID', 'operation', 'devrevstep', 'Program_Name', 'DLCP', 'TEST_NAME', 'TEST_RESULT', 'lot', 'functional_bin']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"输入文件缺少必要的列: {', '.join(missing_columns)}")
        
        print(f"成功读取数据，共{len(df)}行")
        
        # 解析TEST_RESULT列，提取时间数据
        # 修改解析测试结果部分
        print("正在解析测试结果...")
        time_data = []
        # 添加进度条
        for _, row in tqdm(df.iterrows(), total=len(df), desc="解析测试结果"):
            pre_time, main_time, total_time = parse_test_result(row['TEST_RESULT'])
            if pre_time is not None:
                time_data.append({
                    'VISUAL_ID': row['VISUAL_ID'],
                    'operation': row['operation'],
                    'devrevstep': row['devrevstep'],
                    'Program_Name': row['Program_Name'],
                    'DLCP': row['DLCP'],
                    'TEST_NAME': row['TEST_NAME'],
                    'PRE_TIME': pre_time,
                    'MAIN_TIME': main_time,
                    'TOTAL_TIME': total_time,
                    'lot': row['lot'],
                    'functional_bin': row['functional_bin']
                })
        
        # 创建新的DataFrame
        time_df = pd.DataFrame(time_data)
        if time_df.empty:
            raise ValueError("无法从TEST_RESULT列解析出有效的时间数据")
        
        print(f"成功解析{len(time_df)}行有效时间数据")
        
        # 找出数量最多的functional_bin
        bin_counts = time_df['functional_bin'].value_counts()
        most_common_bin = bin_counts.idxmax()
        print(f"数量最多的functional_bin为: {most_common_bin}，共有{bin_counts[most_common_bin]}条记录")
        
        # 筛选出functional_bin值为最常见值的数据
        filtered_df = time_df[time_df['functional_bin'] == most_common_bin]
        print(f"筛选后的数据共有{len(filtered_df)}行")
        
        # 按VISUAL_ID和TEST_NAME计算平均时间
        print("正在计算每个VISUAL_ID下各TEST_NAME的平均测试时间...")
        avg_times = filtered_df.groupby(['VISUAL_ID', 'TEST_NAME']).agg({
            'PRE_TIME': 'mean',
            'MAIN_TIME': 'mean',
            'TOTAL_TIME': 'mean',
            'Program_Name': 'first',
            'DLCP': 'first',
            'lot': 'first',
            'functional_bin': 'first'
        }).reset_index()
        
        # 保存平均时间结果 - 使用CSV格式避免Excel行数限制
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        avg_output_file = os.path.join(output_dir, f'avg_test_times_{timestamp}.csv')
        avg_times.to_csv(avg_output_file, index=False)
        print(f"平均测试时间已保存至: {avg_output_file}")
        
        # 计算每个lot的TTG (Test Time per Good unit)
        print("正在计算每个lot的TTG...")
        
        # 首先计算每个VISUAL_ID的总测试时间
        visual_id_total_time = filtered_df.groupby('VISUAL_ID')['TOTAL_TIME'].sum().reset_index()
        
        # 然后按lot和Program_Name分组，计算TTG
        lot_groups = filtered_df[['lot', 'Program_Name', 'DLCP', 'VISUAL_ID']].drop_duplicates()
        lot_ttg_data = []
        
        # 添加进度条
        lot_group_iter = tqdm(lot_groups.groupby('lot'), desc="计算TTG", total=len(lot_groups['lot'].unique()))
        for lot, lot_group in lot_group_iter:
            for program_name, program_group in lot_group.groupby('Program_Name'):
                # 获取该lot下的所有VISUAL_ID
                visual_ids = program_group['VISUAL_ID'].unique()
                visual_id_count = len(visual_ids)
                
                # 计算该lot下所有VISUAL_ID的总测试时间
                lot_total_time = visual_id_total_time[visual_id_total_time['VISUAL_ID'].isin(visual_ids)]['TOTAL_TIME'].sum()
                
                # 计算TTG并转换为秒
                ttg_ms = lot_total_time / visual_id_count if visual_id_count > 0 else 0
                ttg_s = ttg_ms / 1000  # 将毫秒转换为秒
                
                lot_ttg_data.append({
                    'lot': lot,
                    'Program_Name': program_name,
                    'DLCP': program_group['DLCP'].iloc[0],
                    'VISUAL_ID_Count': visual_id_count,
                    'Total_Test_Time_ms': lot_total_time,
                    'Total_Test_Time_s': lot_total_time / 1000,  # 添加秒为单位的总时间
                    'TTG_ms': ttg_ms,
                    'TTG_s': ttg_s,  # 添加秒为单位的TTG
                    'functional_bin': most_common_bin
                })
        
        # 创建TTG DataFrame
        ttg_df = pd.DataFrame(lot_ttg_data)
        
        # 保存TTG结果 - 使用CSV格式
        ttg_output_file = os.path.join(output_dir, f'lot_ttg_{timestamp}.csv')
        ttg_df.to_csv(ttg_output_file, index=False)
        print(f"每个lot的TTG已保存至: {ttg_output_file} (TTG单位: 秒)")
        
        # 比较不同Program_Name但相同DLCP下的测试时间差异
        print("正在比较不同Program_Name下相同DLCP的测试时间差异...")
        
        # 设置最小样本量阈值
        min_visual_id_count = 5  # 可以根据实际情况调整这个值
        
        # 按DLCP和TEST_NAME分组
        dlcp_groups = filtered_df.groupby(['DLCP', 'TEST_NAME'])
        
        # 存储差异数据
        diff_data = []
        
        # 添加进度条
        dlcp_group_iter = tqdm(dlcp_groups, desc="比较测试时间差异", total=len(dlcp_groups))
        for (dlcp, test_name), group in dlcp_group_iter:
            # 获取组内所有不同的Program_Name
            program_names = group['Program_Name'].unique()
            
            # 获取该测试项目下的VISUAL_ID数量
            visual_id_count = len(group['VISUAL_ID'].unique())
            
            # 如果样本量太小，跳过该测试项目
            if visual_id_count < min_visual_id_count:
                continue
                
            # 按Program_Name计算平均时间
            program_avg = group.groupby('Program_Name').agg({
                'PRE_TIME': 'mean',
                'MAIN_TIME': 'mean',
                'TOTAL_TIME': 'mean',
                'VISUAL_ID': lambda x: len(x.unique())  # 计算每个Program_Name下的VISUAL_ID数量
            }).reset_index()
            
            # 区分CCG和EDGE的TP
            ccg_programs = []
            edge_programs = []
            
            for program in program_names:
                # 检查Program_Name的第十位字母是否为'H'
                if len(program) >= 10 and program[9] == 'H':
                    ccg_programs.append(program)
                else:
                    edge_programs.append(program)
            
            # 只有当同时存在CCG和EDGE的TP时才进行比较
            if ccg_programs and edge_programs:
                # 计算CCG的平均时间和样本量
                ccg_avg_time = program_avg[program_avg['Program_Name'].isin(ccg_programs)]['TOTAL_TIME'].mean()
                ccg_sample_count = sum(program_avg[program_avg['Program_Name'].isin(ccg_programs)]['VISUAL_ID'])
                
                # 计算EDGE的平均时间和样本量
                edge_avg_time = program_avg[program_avg['Program_Name'].isin(edge_programs)]['TOTAL_TIME'].mean()
                edge_sample_count = sum(program_avg[program_avg['Program_Name'].isin(edge_programs)]['VISUAL_ID'])
                
                # 如果任一组的样本量太小，跳过
                if ccg_sample_count < min_visual_id_count or edge_sample_count < min_visual_id_count:
                    continue
                
                # 计算delta (EDGE - CCG)
                delta = edge_avg_time - ccg_avg_time
                
                # 计算其他差异数据
                max_pre_diff = program_avg['PRE_TIME'].max() - program_avg['PRE_TIME'].min()
                max_main_diff = program_avg['MAIN_TIME'].max() - program_avg['MAIN_TIME'].min()
                
                diff_data.append({
                    'DLCP': dlcp,
                    'TEST_NAME': test_name,
                    'PRE_TIME_DIFF': max_pre_diff,
                    'MAIN_TIME_DIFF': max_main_diff,
                    'delta': delta,
                    'CCG_Programs': ', '.join(ccg_programs),
                    'EDGE_Programs': ', '.join(edge_programs),
                    'CCG_AVG_TIME': ccg_avg_time,
                    'EDGE_AVG_TIME': edge_avg_time,
                    'CCG_Sample_Count': ccg_sample_count,
                    'EDGE_Sample_Count': edge_sample_count,
                    'Total_Sample_Count': visual_id_count,
                    'PROGRAM_COUNT': len(program_names),
                    'functional_bin': most_common_bin
                })
        
        # 创建差异DataFrame
        diff_df = pd.DataFrame(diff_data)
        if not diff_df.empty:
            # 按delta的绝对值排序并获取前20个差异最大的测试（原来是10个）
            top_diff = diff_df.sort_values(by='delta', key=abs, ascending=False).head(20)
            
            # 保存差异结果 - 使用CSV格式
            diff_output_file = os.path.join(output_dir, f'test_time_diff_{timestamp}.csv')
            diff_df.to_csv(diff_output_file, index=False)
            print(f"测试时间差异已保存至: {diff_output_file}")
            
            # 保存前20个差异最大的测试 - 这个可以用Excel因为只有20行
            top_output_file = os.path.join(output_dir, f'top20_test_time_diff_{timestamp}.xlsx')
            top_diff.to_excel(top_output_file, index=False)
            print(f"前20个差异最大的测试已保存至: {top_output_file}")
            
            # 打印前20个差异最大的测试
            print(f"\n前20个差异最大的测试 (functional_bin = {most_common_bin}):")
            for i, (_, row) in enumerate(top_diff.iterrows(), 1):
                print(f"{i}. DLCP: {row['DLCP']}, TEST_NAME: {row['TEST_NAME']}")
                print(f"   delta值(EDGE-CCG): {row['delta']:.2f}ms")
                print(f"   CCG程序: {row['CCG_Programs']}, 平均时间: {row['CCG_AVG_TIME']:.2f}ms, 样本量: {row['CCG_Sample_Count']}")
                print(f"   EDGE程序: {row['EDGE_Programs']}, 平均时间: {row['EDGE_AVG_TIME']:.2f}ms, 样本量: {row['EDGE_Sample_Count']}")
                print(f"   总样本量: {row['Total_Sample_Count']}, PRE差异: {row['PRE_TIME_DIFF']:.2f}ms, MAIN差异: {row['MAIN_TIME_DIFF']:.2f}ms")
                print()
        else:
            print(f"未发现同时存在CCG和EDGE的测试项 (functional_bin = {most_common_bin})")
        
        return avg_output_file, diff_output_file if not diff_df.empty else None, ttg_output_file
        
    except Exception as e:
        print(f"分析过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None, None

if __name__ == "__main__":
    try:
        # 获取输入文件路径
        input_file = input("请输入数据文件路径(支持Excel或CSV格式): ").strip()
        
        # 处理可能包含引号的路径
        input_file = input_file.strip('"\'')
        
        if not os.path.exists(input_file):
            print(f"错误: 文件不存在 - {input_file}")
            
            # 尝试提供更多帮助信息
            dir_path = os.path.dirname(input_file)
            if os.path.exists(dir_path):
                print(f"\n目录 '{dir_path}' 存在，但找不到指定文件。")
                print("该目录下的文件列表:")
                for file in os.listdir(dir_path):
                    if file.endswith(('.xlsx', '.xls', '.csv')):
                        print(f"  - {file}")
            else:
                print(f"\n目录 '{dir_path}' 不存在。请检查路径是否正确。")
        else:
            # 执行分析
            avg_file, diff_file, ttg_file = analyze_test_times(input_file)
            if avg_file:
                print(f"\n分析完成！")
                print(f"平均测试时间文件: {avg_file}")
                if diff_file:
                    print(f"测试时间差异文件: {diff_file}")
                if ttg_file:
                    print(f"每个lot的TTG文件: {ttg_file}")
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        