"""配置文件加载模块"""
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import yaml

from .mole_submitter import MoleConfig
from .spark_submitter import SparkConfig
from .gts_submitter import GTSConfig

LOGGER = logging.getLogger(__name__)


@dataclass
class PathsConfig:
    """路径配置"""
    input_dir: Path
    output_dir: Path
    log_dir: Path
    source_lot_file: Optional[Path] = None
    file_path_config: Optional[Path] = None
    tp_path: Optional[str] = None


@dataclass
class TimeoutsConfig:
    """超时配置"""
    excel_read: int = 30
    excel_write: int = 30
    step_execution: int = 300
    overall: int = 900


@dataclass
class LoggingConfig:
    """日志配置"""
    level: str = "INFO"
    format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    file: str = "workflow_automation.log"


@dataclass
class WorkflowConfig:
    """工作流配置"""
    paths: PathsConfig
    mole: MoleConfig
    spark: SparkConfig
    gts: GTSConfig
    timeouts: TimeoutsConfig
    logging: LoggingConfig


def _ensure_path(value: str, base_dir: Path) -> Path:
    """确保路径是绝对路径"""
    path = Path(value)
    if not path.is_absolute():
        path = (base_dir / path).resolve()
    return path


def load_config(config_path: Path) -> WorkflowConfig:
    """加载配置文件"""
    base_dir = config_path.parent
    
    with open(config_path, "r", encoding="utf-8") as fp:
        data = yaml.safe_load(fp) or {}
    
    # 尝试加载Spark自动化配置文件（如果存在）
    paths_section = data.get("paths", {})
    spark_config_file = paths_section.get("spark_config")
    if spark_config_file:
        spark_config_path = _ensure_path(spark_config_file, base_dir)
        if spark_config_path.exists():
            try:
                with open(spark_config_path, "r", encoding="utf-8") as fp:
                    spark_auto_config = yaml.safe_load(fp) or {}
                    LOGGER.info(f"已加载Spark自动化配置: {spark_config_path}")
                    
                    # 合并Spark配置（spark_automation_config.yaml优先）
                    if "spark" in spark_auto_config and "spark" in data:
                        for key, value in spark_auto_config["spark"].items():
                            if value is not None:  # 只覆盖非空值
                                data["spark"][key] = value
                    
                    # 合并TP路径
                    if "test_program" in spark_auto_config:
                        tp_path = spark_auto_config["test_program"].get("tp_path")
                        if tp_path:
                            data["paths"]["tp_path"] = tp_path
                            
            except Exception as e:
                LOGGER.warning(f"加载Spark自动化配置文件失败: {e}")
    
    # 路径配置
    paths_section = data.get("paths", {})
    source_lot_file_path = paths_section.get("source_lot_file")
    file_path_config = paths_section.get("file_path_config")
    tp_path = paths_section.get("tp_path")
    
    paths_cfg = PathsConfig(
        input_dir=_ensure_path(paths_section.get("input_dir", ".."), base_dir),
        output_dir=_ensure_path(paths_section.get("output_dir", "./reports"), base_dir),
        log_dir=_ensure_path(paths_section.get("log_dir", "./logs"), base_dir),
        source_lot_file=_ensure_path(source_lot_file_path, base_dir) if source_lot_file_path else None,
        file_path_config=_ensure_path(file_path_config, base_dir) if file_path_config else None,
        tp_path=tp_path,
    )
    
    # 确保目录存在
    paths_cfg.output_dir.mkdir(parents=True, exist_ok=True)
    paths_cfg.log_dir.mkdir(parents=True, exist_ok=True)
    
    # Mole配置
    mole_section = data.get("mole", {})
    mole_executable = mole_section.get("executable_path")
    mole_cfg = MoleConfig(
        executable_path=_ensure_path(mole_executable, base_dir) if mole_executable else None,
        window_title=mole_section.get("window_title", "Mole"),
        timeout=mole_section.get("timeout", 60),
        retry_count=mole_section.get("retry_count", 3),
        retry_delay=mole_section.get("retry_delay", 2),
    )
    
    # Spark配置
    spark_section = data.get("spark", {})
    spark_url = spark_section.get("url")
    if not spark_url:
        LOGGER.warning("Spark URL未配置，请在config.yaml中设置spark.url")
    spark_cfg = SparkConfig(
        url=spark_url or "",
        vpo_category=spark_section.get("vpo_category", "correlation"),
        step=spark_section.get("step", "B5"),
        tags=spark_section.get("tags", "CCG_24J-TEST"),
        timeout=spark_section.get("timeout", 60),
        retry_count=spark_section.get("retry_count", 3),
        retry_delay=spark_section.get("retry_delay", 2),
        wait_after_submit=spark_section.get("wait_after_submit", 5),
        headless=spark_section.get("headless", False),
        implicit_wait=spark_section.get("implicit_wait", 10),
        explicit_wait=spark_section.get("explicit_wait", 20),
    )
    
    # GTS配置
    gts_section = data.get("gts", {})
    gts_url = gts_section.get("url")
    if not gts_url:
        LOGGER.warning("GTS URL未配置，请在config.yaml中设置gts.url")
    gts_cfg = GTSConfig(
        url=gts_url or "",
        timeout=gts_section.get("timeout", 60),
        retry_count=gts_section.get("retry_count", 3),
        retry_delay=gts_section.get("retry_delay", 2),
        wait_after_submit=gts_section.get("wait_after_submit", 5),
        headless=gts_section.get("headless", False),
        implicit_wait=gts_section.get("implicit_wait", 10),
        explicit_wait=gts_section.get("explicit_wait", 20),
    )
    
    # 超时配置
    timeouts_section = data.get("timeouts", {})
    timeouts_cfg = TimeoutsConfig(
        excel_read=timeouts_section.get("excel_read", 30),
        excel_write=timeouts_section.get("excel_write", 30),
        step_execution=timeouts_section.get("step_execution", 300),
        overall=timeouts_section.get("overall", 900),
    )
    
    # 日志配置
    logging_section = data.get("logging", {})
    logging_cfg = LoggingConfig(
        level=logging_section.get("level", "INFO"),
        format=logging_section.get("format", "%(asctime)s - %(name)s - %(levelname)s - %(message)s"),
        file=logging_section.get("file", "workflow_automation.log"),
    )
    
    LOGGER.debug(f"配置加载完成: {config_path}")
    
    return WorkflowConfig(
        paths=paths_cfg,
        mole=mole_cfg,
        spark=spark_cfg,
        gts=gts_cfg,
        timeouts=timeouts_cfg,
        logging=logging_cfg,
    )

