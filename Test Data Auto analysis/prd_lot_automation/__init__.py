"""
PRD LOT 自动化工具包。

包含以下模块：
- config_loader: 读取/校验配置
- lot_reader: LOT 列表处理
- spf_runner: SQLPathFinder UI 控制
- data_processing: CSV 解析与聚合
- report_builder: Excel 报表生成
- mailer: 邮件发送
- main: 入口
"""

__all__ = [
    "config_loader",
    "lot_reader",
    "spf_runner",
    "data_processing",
    "report_builder",
    "mailer",
]

