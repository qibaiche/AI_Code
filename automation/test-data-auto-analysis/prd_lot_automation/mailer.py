import logging
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Sequence, Optional

from .config_loader import AppConfig

LOGGER = logging.getLogger(__name__)


def _build_subject(config: AppConfig, date_str: str, lot_count: int) -> str:
    template = config.email.subject_template or "[LOT日报] {date}"
    return template.format(date=date_str, lot_count=lot_count)


def _build_body(summary: dict) -> str:
    # 确保 top_bins 中的元素都是字符串（可能是整数）
    top_bins_str = [str(bin_name) for bin_name in summary['top_bins']]
    return (
        "LOT 自动化任务完成：\n"
        f"- LOT 总数：{summary['lot_count']}\n"
        f"- Top3 Functional Bin：{', '.join(top_bins_str)}\n"
        f"- 异常记录：{summary['exception_count']}\n"
        f"- Functional Bin Grand Total：{summary['grand_total']}\n"
    )


def send_email_via_outlook(
    config: AppConfig,
    subject: str,
    body: str,
    attachments: Sequence[Path],
) -> None:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("请先安装 pywin32 以使用 Outlook 通道") from exc

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = ";".join(config.email.to)
    mail.CC = ";".join(config.email.cc)
    mail.Subject = subject
    mail.Body = body
    for attach in attachments:
        mail.Attachments.Add(str(attach))
    mail.Send()
    LOGGER.info("Outlook 邮件已发送")


def send_email_via_smtp(
    config: AppConfig,
    subject: str,
    body: str,
    attachments: Sequence[Path],
) -> None:
    if not config.email.smtp_host:
        raise RuntimeError("未配置 SMTP host")
    message = EmailMessage()
    message["Subject"] = subject
    message["To"] = ", ".join(config.email.to)
    if config.email.cc:
        message["Cc"] = ", ".join(config.email.cc)
    message.set_content(body)
    for attach in attachments:
        data = attach.read_bytes()
        message.add_attachment(
            data,
            maintype="application",
            subtype="octet-stream",
            filename=attach.name,
        )
    recipients = config.email.to + config.email.cc
    with smtplib.SMTP(config.email.smtp_host, config.email.smtp_port) as smtp:
        if config.email.use_tls:
            smtp.starttls()
        if config.email.smtp_username and config.email.smtp_password:
            smtp.login(config.email.smtp_username, config.email.smtp_password)
        smtp.send_message(message, to_addrs=recipients)
    LOGGER.info("SMTP 邮件已发送")


def send_report_email(
    config: AppConfig,
    summary: dict,
    attachments: Sequence[Path],
) -> None:
    subject = _build_subject(config, summary["date"], summary["lot_count"])
    body = _build_body(summary)
    if config.email.mode.lower() == "smtp":
        send_email_via_smtp(config, subject, body, attachments)
    else:
        send_email_via_outlook(config, subject, body, attachments)


def send_unified_report_email(
    config: AppConfig,
    summary: dict,
    attachments: Sequence[Path],
    html_body: str,
) -> None:
    """发送统一报告邮件，包含 HTML 格式的正文"""
    subject = _build_subject(config, summary["date"], summary["lot_count"])
    
    if config.email.mode.lower() == "smtp":
        # SMTP 不支持 HTML，使用纯文本
        body = _build_body(summary)
        send_email_via_smtp(config, subject, body, attachments)
    else:
        # Outlook 支持 HTML
        send_html_email_via_outlook(config, subject, html_body, attachments)


def send_html_email_via_outlook(
    config: AppConfig,
    subject: str,
    html_body: str,
    attachments: Sequence[Path],
) -> None:
    """通过 Outlook 发送 HTML 格式的邮件"""
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("请先安装 pywin32 以使用 Outlook 通道") from exc

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = ";".join(config.email.to)
    mail.CC = ";".join(config.email.cc)
    mail.Subject = subject
    mail.HTMLBody = html_body  # 使用 HTMLBody 而不是 Body
    for attach in attachments:
        mail.Attachments.Add(str(attach))
    mail.Send()
    LOGGER.info("Outlook HTML 邮件已发送")
