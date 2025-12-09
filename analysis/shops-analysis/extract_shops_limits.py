import re
import csv
from collections import defaultdict
import os

# --- 常量定义 ---
INPUT_FILE = r"c:\Users\qibaiche\OneDrive - Intel Corporation\Documents\AI_Code\shops_analysis\TPI_SHOPS.mtpl"
OUTPUT_DIR = r"c:\Users\qibaiche\OneDrive - Intel Corporation\Documents\AI_Code\shops_analysis"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "shops_limits.csv")
DEFAULT_VALUE = "N/A"
COMMENT_PREFIX = "注释:"

# 定义五个子表的输出文件名
SUBTABLE_FILES = {
    "PHMCOLD": os.path.join(OUTPUT_DIR, "shops_limits_PHMCOLD.csv"),
    "PHMHOT": os.path.join(OUTPUT_DIR, "shops_limits_PHMHOT.csv"),
    "CLASSCOLD": os.path.join(OUTPUT_DIR, "shops_limits_COLD.csv"),
    "CLASSHOT": os.path.join(OUTPUT_DIR, "shops_limits_HOT.csv"),
    "ALL": os.path.join(OUTPUT_DIR, "shops_limits_ALL.csv")
}

# --- 正则表达式模式 ---
# 使用非捕获组 (?:) 处理 upper|lower 前缀
LIMIT_PATTERNS = [
    # 添加新模式：Selector_LOWERDIODE_limit_high_EVG 和 Selector_LOWERDIODE_limit_low_EVG
    (r'(?:upper|lower)_diode_limit_high\s*=\s*TPI_SHOPS_Rules\.Selector_(?:UPPER|LOWER)DIODE_limit_high_EVG\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     r'(?:upper|lower)_diode_limit_low\s*=\s*TPI_SHOPS_Rules\.Selector_(?:UPPER|LOWER)DIODE_limit_low_EVG\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     False),
    # 带引号的赋值模式
    (r'(?:upper|lower)_diode_limit_high\s*=\s*[^(]*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     r'(?:upper|lower)_diode_limit_low\s*=\s*[^(]*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     False),
    # 不带引号的赋值模式
    (r'(?:upper|lower)_diode_limit_high\s*=\s*[^(]*\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,\)]+)\s*\)',
     r'(?:upper|lower)_diode_limit_low\s*=\s*[^(]*\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,\)]+)\s*\)',
     False),
    # 带引号的函数调用模式
    (r'(?:upper|lower)_diode_limit_high\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     r'(?:upper|lower)_diode_limit_low\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     False),
    # 不带引号的函数调用模式
    (r'(?:upper|lower)_diode_limit_high\s*\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,\)]+)\s*\)',
     r'(?:upper|lower)_diode_limit_low\s*\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,\)]+)\s*\)',
     False),
    # 注释掉的模式
    (r'#\s*(?:upper|lower)_diode_limit_high\s*=\s*[^(]*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     r'#\s*(?:upper|lower)_diode_limit_low\s*=\s*[^(]*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)',
     True),
]

def extract_limits(block):
    """从测试块中提取上限和下限值"""
    high_values = [DEFAULT_VALUE] * 5
    low_values = [DEFAULT_VALUE] * 5
    
    # 遍历所有模式，找到第一个匹配的
    for high_pattern, low_pattern, is_comment in LIMIT_PATTERNS:
        # 尝试匹配上限
        high_match = re.search(high_pattern, block)
        if high_match:
            values = list(high_match.groups())
            high_values = [val.strip('"\'').strip() if val else DEFAULT_VALUE for val in values]
            if is_comment:
                high_values = [f"{COMMENT_PREFIX}{val}" if val != DEFAULT_VALUE else val for val in high_values]
        
        # 尝试匹配下限
        low_match = re.search(low_pattern, block)
        if low_match:
            values = list(low_match.groups())
            low_values = [val.strip('"\'').strip() if val else DEFAULT_VALUE for val in values]
            if is_comment:
                low_values = [f"{COMMENT_PREFIX}{val}" if val != DEFAULT_VALUE else val for val in low_values]
        
        # 如果找到了匹配，就不再继续查找
        if high_match or low_match:
            break
    
    return high_values, low_values

def determine_category(test_name):
    """根据测试名称确定类别为 UPPERDIODE 或 LOWERDIODE"""
    if "UPPERDIODE" in test_name:
        return "UPPERDIODE"
    elif "LOWERDIODE" in test_name:
        return "LOWERDIODE"
    else:
        return "OTHER"  # 对于既不是 UPPERDIODE 也不是 LOWERDIODE 的测试

def main():
    # 存储结果的数据结构
    results = defaultdict(lambda: defaultdict(dict))
    
    # 为五个子表创建单独的结果集
    subtable_results = {
        "PHMCOLD": defaultdict(lambda: defaultdict(dict)),
        "PHMHOT": defaultdict(lambda: defaultdict(dict)),
        "CLASSCOLD": defaultdict(lambda: defaultdict(dict)),
        "CLASSHOT": defaultdict(lambda: defaultdict(dict)),
        "ALL": defaultdict(lambda: defaultdict(dict))
    }
    
    try:
        # 读取输入文件
        with open(INPUT_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 查找所有测试定义
        test_pattern = re.compile(
            r"^(#?)\s*Test\s+iCShopsTest\s*(\S+)(.*?)(?=^\s*(?:#\s*)?Test\s+iCShopsTest|\Z)",
            re.MULTILINE | re.DOTALL
        )
        
        # 处理每个测试块
        for match in test_pattern.finditer(content):
            is_commented = match.group(1) == "#"
            test_name = match.group(2).strip()
            block = match.group(3)
            
            # 跳过注释的测试和以_FF结尾的测试
            if is_commented:
                print(f"跳过注释的测试: {test_name}")
                continue
            if test_name.endswith("_FF"):
                print(f"跳过_FF测试: {test_name}")
                continue
            
            # 根据测试名称确定类别（UPPERDIODE 或 LOWERDIODE）
            category = determine_category(test_name)
            high_values, low_values = extract_limits(block)
            
            # 存储主表结果
            results[category][test_name] = {
                "PHMCOLD_high": high_values[0], "PHMHOT_high": high_values[1],
                "CLASSCOLD_high": high_values[2], "CLASSHOT_high": high_values[3],
                "ALL_high": high_values[4],
                "PHMCOLD_low": low_values[0], "PHMHOT_low": low_values[1],
                "CLASSCOLD_low": low_values[2], "CLASSHOT_low": low_values[3],
                "ALL_low": low_values[4]
            }
            
            # 存储子表结果
            subtable_results["PHMCOLD"][category][test_name] = {
                "high": high_values[0],
                "low": low_values[0]
            }
            
            subtable_results["PHMHOT"][category][test_name] = {
                "high": high_values[1],
                "low": low_values[1]
            }
            
            subtable_results["CLASSCOLD"][category][test_name] = {
                "high": high_values[2],
                "low": low_values[2]
            }
            
            subtable_results["CLASSHOT"][category][test_name] = {
                "high": high_values[3],
                "low": low_values[3]
            }
            
            subtable_results["ALL"][category][test_name] = {
                "high": high_values[4],
                "low": low_values[4]
            }
        
        # 写入主CSV文件
        with open(OUTPUT_FILE, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'Category', 'Test_Name',
                'PHMCOLD_high', 'PHMHOT_high', 'CLASSCOLD_high', 'CLASSHOT_high', 'ALL_high',
                'PHMCOLD_low', 'PHMHOT_low', 'CLASSCOLD_low', 'CLASSHOT_low', 'ALL_low'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # 按类别和测试名称排序写入
            for category in sorted(results.keys()):
                for test_name in sorted(results[category].keys()):
                    row = {'Category': category, 'Test_Name': test_name}
                    row.update(results[category][test_name])
                    writer.writerow(row)
        
        # 写入五个子表CSV文件
        for subtable_name, subtable_data in subtable_results.items():
            output_file = SUBTABLE_FILES[subtable_name]
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['Category', 'Test_Name', 'High_Limit', 'Low_Limit']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                # 按类别和测试名称排序写入
                for category in sorted(subtable_data.keys()):
                    for test_name in sorted(subtable_data[category].keys()):
                        row = {
                            'Category': category, 
                            'Test_Name': test_name,
                            'High_Limit': subtable_data[category][test_name]['high'],
                            'Low_Limit': subtable_data[category][test_name]['low']
                        }
                        writer.writerow(row)
            
            print(f"子表 {subtable_name} 已保存到 {output_file}")
        
        print(f"提取完成！主表结果已保存到 {OUTPUT_FILE}")
        print(f"五个子表已保存到 {OUTPUT_DIR} 目录下")
    
    except Exception as e:
        print(f"处理过程中出错: {e}")

if __name__ == "__main__":
    main()