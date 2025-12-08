import json
import os
import csv
from datetime import datetime
from collections import defaultdict
from tqdm import tqdm

def analyze_config_execution_modes(json_file):
    """分析JSON文件中每个ConfigSetName下ExecutionMode的出现次数"""
    try:
        print(f"正在读取JSON文件: {json_file}")
        file_size = os.path.getsize(json_file) / (1024 * 1024)
        print(f"文件大小: {file_size:.2f} MB")

        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 初始化计数器
        config_counts = defaultdict(lambda: defaultdict(int))
        valid_configs = 0
        
        # 处理顶层对象
        if isinstance(data, dict):
            process_types = data.get('ProcessTypes', [])
            if process_types and isinstance(process_types, list):
                # 遍历ProcessTypes
                for process in process_types:
                    unit_infos = process.get('UnitInfo', [])
                    if unit_infos and isinstance(unit_infos, list):
                        # 遍历UnitInfo
                        for unit in tqdm(unit_infos, desc="处理配置"):
                            config_details = unit.get('ConfigDetails', [])
                            if config_details and isinstance(config_details, list):
                                # 遍历ConfigDetails
                                for config in config_details:
                                    config_name = config.get('ConfigSetName')
                                    exec_mode = config.get('ExecutionMode')
                                    if config_name and exec_mode:
                                        config_counts[config_name][exec_mode] += 1
                                        valid_configs += 1
        
        print(f"共找到 {valid_configs} 条有效配置")
        
        if valid_configs == 0:
            print("未找到任何配置数据，请检查JSON文件结构")
            return None
        
        # 格式化结果
        result = {}
        for config_name, modes in config_counts.items():
            total = sum(modes.values())
            result[config_name] = {
                'modes': modes,
                'total_count': total
            }
        
        return result
    
    except Exception as e:
        print(f"分析JSON文件时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def print_config_stats(stats):
    """打印配置统计结果"""
    if not stats:
        print("没有可用的统计数据")
        return
    
    print("\n配置统计结果:")
    print("=" * 80)
    
    for config_name, data in sorted(stats.items()):
        print(f"\n配置名称: {config_name}")
        print(f"总数量: {data['total_count']}")
        print("执行模式统计:")
        
        # 按数量降序排列
        sorted_modes = sorted(data['modes'].items(), key=lambda x: x[1], reverse=True)
        for mode, count in sorted_modes:
            print(f"  - {mode}: {count}次 ({count/data['total_count']*100:.1f}%)")
    
    print("=" * 80)

def save_config_stats(stats, output_file):
    """保存配置统计结果到CSV文件"""
    try:
        with open(output_file, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['ConfigSetName', 'ExecutionMode', 'Count', 'Percentage'])
            
            for config_name, data in stats.items():
                total = data['total_count']
                for mode, count in data['modes'].items():
                    percentage = count/total*100 if total > 0 else 0
                    writer.writerow([config_name, mode, count, f"{percentage:.1f}%"])
        
        print(f"\n统计结果已保存至: {output_file}")
        return True
    except Exception as e:
        print(f"保存统计结果时出错: {str(e)}")
        return False

if __name__ == "__main__":
    try:
        # 交互式获取JSON文件路径
        json_file = input("请输入要分析的JSON文件路径: ").strip()
        
        # 去除可能存在的引号
        json_file = json_file.strip('"\'')
        
        if not os.path.exists(json_file):
            print(f"错误: 文件不存在 - {json_file}")
        else:
            print("开始分析配置数据...")
            stats = analyze_config_execution_modes(json_file)
            
            if stats:
                print_config_stats(stats)
                
                # 询问是否需要保存CSV文件
                save_csv = input("是否保存结果到CSV文件? (y/n): ").strip().lower()
                if save_csv == 'y':
                    # 询问输出路径或使用默认路径
                    output_path = input("请输入CSV输出路径(留空则使用默认路径): ").strip()
                    # 去除可能存在的引号
                    output_path = output_path.strip('"\'')
                    
                    if not output_path:
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        output_file = os.path.join(os.path.dirname(json_file), f"config_stats_{timestamp}.csv")
                    else:
                        output_file = output_path
                    
                    save_config_stats(stats, output_file)
    
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        import traceback
        traceback.print_exc()