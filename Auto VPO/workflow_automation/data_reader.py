"""数据文件读取模块（支持Excel和CSV）"""
import logging
from pathlib import Path
from typing import Optional
import pandas as pd
import pandas as pd

LOGGER = logging.getLogger(__name__)


def read_excel_file(file_path: str | Path) -> pd.DataFrame:
    """
    从文件读取数据（支持Excel和CSV格式）
    
    Args:
        file_path: 文件路径（支持.xlsx, .xls, .csv格式）
    
    Returns:
        pandas DataFrame对象
    
    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 文件格式不支持或无法读取
    """
    file_path = Path(file_path)
    
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    file_ext = file_path.suffix.lower()
    
    if file_ext in ['.xlsx', '.xls']:
        # 读取Excel文件
        LOGGER.info(f"正在读取Excel文件: {file_path}")
        try:
            df = pd.read_excel(file_path)
            LOGGER.info(f"成功读取Excel文件: {len(df)} 行，{len(df.columns)} 列")
            LOGGER.debug(f"列名: {df.columns.tolist()}")
            
            if df.empty:
                LOGGER.warning("Excel文件为空")
            
            return df
        except Exception as e:
            LOGGER.error(f"读取Excel文件失败: {e}")
            raise ValueError(f"无法读取Excel文件 {file_path}: {e}")
    
    elif file_ext == '.csv':
        # 读取CSV文件
        LOGGER.info(f"正在读取CSV文件: {file_path}")
        try:
            # 尝试不同的编码格式
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin1', 'cp1252']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    LOGGER.info(f"成功使用 {encoding} 编码读取CSV文件: {len(df)} 行，{len(df.columns)} 列")
                    break
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    # 如果遇到其他错误，尝试下一种编码
                    LOGGER.warning(f"使用 {encoding} 编码读取失败: {e}")
                    continue
            
            if df is None:
                # 如果所有编码都失败，尝试自动检测
                LOGGER.warning("所有编码尝试失败，尝试自动检测编码...")
                df = pd.read_csv(file_path, encoding='utf-8', errors='ignore')
            
            LOGGER.debug(f"列名: {df.columns.tolist()}")
            
            if df.empty:
                LOGGER.warning("CSV文件为空")
            
            return df
        except Exception as e:
            LOGGER.error(f"读取CSV文件失败: {e}")
            raise ValueError(f"无法读取CSV文件 {file_path}: {e}")
    
    else:
        raise ValueError(f"不支持的文件格式: {file_ext}，仅支持.xlsx、.xls和.csv格式")


def validate_data(df: pd.DataFrame, required_columns: Optional[list] = None) -> bool:
    """
    验证数据框是否包含必需的列
    
    Args:
        df: 数据框
        required_columns: 必需的列名列表，如果为None则不验证
    
    Returns:
        True如果验证通过
    
    Raises:
        ValueError: 如果缺少必需的列
    """
    if required_columns is None:
        return True
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        raise ValueError(f"数据缺少必需的列: {missing_columns}")
    
    LOGGER.info(f"数据验证通过，包含必需的列: {required_columns}")
    return True


def save_result_excel(df: pd.DataFrame, output_path: str | Path, 
                     date_str: Optional[str] = None) -> Path:
    """
    保存结果到Excel文件
    
    Args:
        df: 要保存的数据框
        output_path: 输出目录或文件路径
        date_str: 日期字符串（格式：YYYYMMDD），如果为None则自动生成
    
    Returns:
        保存的文件路径
    """
    from datetime import datetime
    
    output_path = Path(output_path)
    
    # 如果输出路径是目录，生成文件名
    if output_path.is_dir() or not output_path.suffix:
        if date_str is None:
            date_str = datetime.now().strftime("%Y%m%d")
        filename = f"workflow_result_{date_str}.xlsx"
        output_path = output_path / filename
    
    # 确保输出目录存在
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    LOGGER.info(f"正在保存结果到: {output_path}")
    
    try:
        # 保存为Excel文件
        df.to_excel(output_path, index=False, engine='openpyxl')
        LOGGER.info(f"成功保存结果: {output_path} ({len(df)} 行，{len(df.columns)} 列)")
        
        return output_path
    
    except Exception as e:
        LOGGER.error(f"保存Excel文件失败: {e}")
        raise ValueError(f"无法保存Excel文件 {output_path}: {e}")

