import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Dict, Any

import yaml


LOGGER = logging.getLogger(__name__)


@dataclass
class EmailConfig:
    mode: str = "outlook"  # outlook / smtp
    to: List[str] = field(default_factory=list)
    cc: List[str] = field(default_factory=list)
    subject_template: str = "[LOT日报] {date}（共{lot_count}个 LOT）"
    use_outlook_body_template: bool = True
    outlook_account: Optional[str] = None
    smtp_host: Optional[str] = None
    smtp_port: int = 587
    smtp_username: Optional[str] = None
    smtp_password: Optional[str] = None
    use_tls: bool = True


@dataclass
class PathsConfig:
    lots_file: Path
    vg2_file: Path
    spf_executable: Optional[Path]
    output_csv: Path
    report_dir: Path
    config_dir: Path
    log_dir: Path


@dataclass
class FieldsConfig:
    lot: str = "lot"
    functional_bin: str = "functional_bin"
    interface_bin: str = "interface_bin"
    devrevstep: str = "devrevstep"
    visual_id: str = "VISUAL_ID"


@dataclass
class TimeoutsConfig:
    spf_launch: int = 60
    ui_action: int = 15
    execution_poll_interval: int = 5
    file_stabilize_checks: int = 4
    file_stabilize_interval: int = 2
    overall_timeout: int = 900


@dataclass
class UIConfig:
    main_window_title: str = "SQLPathFinder"
    run_button_automation_id: Optional[str] = None
    lot_popup_title: str = "Enter Values for Lot Filtering"
    lot_input_automation_id: Optional[str] = None
    ok_button_automation_id: Optional[str] = None
    run_button_image: Optional[Path] = None


@dataclass
class ProcessingConfig:
    percentage_thresholds: Dict[str, float] = field(
        default_factory=lambda: {
            "red": 0.05,
            "yellow": 0.02,
        }
    )
    max_lots_per_batch: Optional[int] = None
    filters: Dict[str, Any] = field(default_factory=dict)


@dataclass
class AppConfig:
    paths: PathsConfig
    email: EmailConfig
    fields: FieldsConfig = field(default_factory=FieldsConfig)
    timeouts: TimeoutsConfig = field(default_factory=TimeoutsConfig)
    ui: UIConfig = field(default_factory=UIConfig)
    processing: ProcessingConfig = field(default_factory=ProcessingConfig)
    extra: Dict[str, Any] = field(default_factory=dict)


def _ensure_path(value: str, base_dir: Path) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = (base_dir / path).resolve()
    return path


def load_config(config_path: Path) -> AppConfig:
    base_dir = config_path.parent
    with open(config_path, "r", encoding="utf-8") as fp:
        data = yaml.safe_load(fp) or {}

    paths_section = data.get("paths", {})
    lots_file = _ensure_path(paths_section["lots_file"], base_dir)
    vg2_file = _ensure_path(paths_section["vg2_file"], base_dir)
    spf_exe = paths_section.get("spf_executable")
    spf_executable = (
        _ensure_path(spf_exe, base_dir) if spf_exe else None
    )
    output_csv = _ensure_path(paths_section["output_csv"], base_dir)
    report_dir = _ensure_path(paths_section["report_dir"], base_dir)
    log_dir = _ensure_path(paths_section.get("log_dir", "logs"), base_dir)

    report_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    paths_cfg = PathsConfig(
        lots_file=lots_file,
        vg2_file=vg2_file,
        spf_executable=spf_executable,
        output_csv=output_csv,
        report_dir=report_dir,
        config_dir=base_dir,
        log_dir=log_dir,
    )

    email_cfg = EmailConfig(
        mode=data.get("email", {}).get("mode", "outlook"),
        to=data.get("email", {}).get("to", []),
        cc=data.get("email", {}).get("cc", []),
        subject_template=data.get("email", {}).get(
            "subject_template",
            EmailConfig.subject_template,
        ),
        use_outlook_body_template=data.get("email", {}).get(
            "use_outlook_body_template", True
        ),
        outlook_account=data.get("email", {}).get("outlook_account"),
        smtp_host=data.get("email", {}).get("smtp", {}).get("host"),
        smtp_port=data.get("email", {}).get("smtp", {}).get("port", 587),
        smtp_username=data.get("email", {}).get("smtp", {}).get("username"),
        smtp_password=data.get("email", {}).get("smtp", {}).get("password"),
        use_tls=data.get("email", {}).get("smtp", {}).get("use_tls", True),
    )

    fields_cfg = FieldsConfig(
        lot=data.get("fields", {}).get("lot", "lot"),
        functional_bin=data.get("fields", {}).get("functional_bin", "functional_bin"),
        interface_bin=data.get("fields", {}).get("interface_bin", "interface_bin"),
        devrevstep=data.get("fields", {}).get("devrevstep", "devrevstep"),
        visual_id=data.get("fields", {}).get("visual_id", "VISUAL_ID"),
    )

    timeouts_section = data.get("timeouts", {})
    timeouts_cfg = TimeoutsConfig(
        spf_launch=timeouts_section.get("spf_launch", 60),
        ui_action=timeouts_section.get("ui_action", 15),
        execution_poll_interval=timeouts_section.get("execution_poll_interval", 5),
        file_stabilize_checks=timeouts_section.get("file_stabilize_checks", 4),
        file_stabilize_interval=timeouts_section.get("file_stabilize_interval", 2),
        overall_timeout=timeouts_section.get("overall_timeout", 900),
    )

    ui_section = data.get("ui", {})
    ui_cfg = UIConfig(
        main_window_title=ui_section.get("main_window_title", "SQLPathFinder"),
        run_button_automation_id=ui_section.get("run_button_automation_id"),
        lot_popup_title=ui_section.get("lot_popup_title", "Enter Values for Lot Filtering"),
        lot_input_automation_id=ui_section.get("lot_input_automation_id"),
        ok_button_automation_id=ui_section.get("ok_button_automation_id"),
        run_button_image=(
            _ensure_path(ui_section["run_button_image"], base_dir)
            if ui_section.get("run_button_image")
            else None
        ),
    )

    processing_section = data.get("processing", {})
    processing_cfg = ProcessingConfig(
        percentage_thresholds=processing_section.get(
            "percentage_thresholds",
            {"red": 0.05, "yellow": 0.02},
        ),
        max_lots_per_batch=processing_section.get("max_lots_per_batch"),
        filters=processing_section.get("filters", {}),
    )

    LOGGER.debug("配置加载完成：%s", config_path)
    return AppConfig(
        paths=paths_cfg,
        email=email_cfg,
        fields=fields_cfg,
        timeouts=timeouts_cfg,
        ui=ui_cfg,
        processing=processing_cfg,
        extra=data.get("extra", {}),
    )

