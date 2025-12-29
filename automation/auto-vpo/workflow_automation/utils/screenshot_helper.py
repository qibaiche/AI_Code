"""截图辅助工具 - 自动截图并记录错误"""
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, Union
from selenium import webdriver

try:
    from PIL import ImageGrab
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

LOGGER = logging.getLogger(__name__)


def capture_error_screenshot(
    driver: webdriver.Chrome,
    error_message: str,
    output_dir: Path,
    prefix: str = "error"
) -> Optional[Path]:
    """
    捕获错误时的截图
    
    Args:
        driver: Selenium WebDriver实例
        error_message: 错误消息
        output_dir: 输出目录
        prefix: 文件名前缀
        
    Returns:
        截图文件路径，如果失败则返回None
    """
    try:
        # 确保输出目录存在
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # 保留毫秒
        screenshot_path = output_dir / f"{prefix}_{timestamp}.png"
        
        # 截图
        driver.save_screenshot(str(screenshot_path))
        
        LOGGER.info(f"📸 错误截图已保存: {screenshot_path.name}")
        return screenshot_path
        
    except Exception as e:
        LOGGER.debug(f"截图失败: {e}")
        return None


def log_error_with_screenshot(
    driver: webdriver.Chrome,
    error_message: str,
    output_dir: Path,
    exception: Optional[Exception] = None,
    prefix: str = "error"
) -> None:
    """
    记录错误并自动截图
    
    Args:
        driver: Selenium WebDriver实例
        error_message: 错误消息
        output_dir: 输出目录
        exception: 异常对象（可选）
        prefix: 文件名前缀
    """
    # 截图
    screenshot_path = capture_error_screenshot(driver, error_message, output_dir, prefix)
    
    # 记录错误
    if screenshot_path:
        LOGGER.error(f"❌ {error_message} [截图: {screenshot_path.name}]")
    else:
        LOGGER.error(f"❌ {error_message}")
    
    # 如果有异常对象，记录详细信息
    if exception:
        import traceback
        LOGGER.error(f"异常详情: {str(exception)}")
        LOGGER.debug(traceback.format_exc())


def capture_debug_screenshot(
    driver: webdriver.Chrome,
    description: str,
    output_dir: Path,
    prefix: str = "debug"
) -> Optional[Path]:
    """
    捕获调试截图（不记录为错误）
    
    Args:
        driver: Selenium WebDriver实例
        description: 截图描述
        output_dir: 输出目录
        prefix: 文件名前缀
        
    Returns:
        截图文件路径，如果失败则返回None
    """
    try:
        # 确保输出目录存在
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        screenshot_path = output_dir / f"{prefix}_{timestamp}.png"
        
        # 截图
        driver.save_screenshot(str(screenshot_path))
        
        LOGGER.debug(f"📸 调试截图已保存: {screenshot_path.name} ({description})")
        return screenshot_path
        
    except Exception as e:
        LOGGER.debug(f"调试截图失败: {e}")
        return None


def capture_screen_screenshot(
    output_dir: Path,
    error_message: str = "",
    prefix: str = "screen_error"
) -> Optional[Path]:
    """
    捕获整个屏幕的截图（用于非浏览器应用，如 Mole）
    
    Args:
        output_dir: 输出目录
        error_message: 错误消息
        prefix: 文件名前缀
        
    Returns:
        截图文件路径，如果失败则返回None
    """
    if not PIL_AVAILABLE:
        LOGGER.debug("PIL/Pillow 未安装，无法截取屏幕")
        return None
    
    try:
        # 确保输出目录存在
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        screenshot_path = output_dir / f"{prefix}_{timestamp}.png"
        
        # 截取整个屏幕
        screenshot = ImageGrab.grab()
        screenshot.save(str(screenshot_path))
        
        LOGGER.info(f"📸 屏幕截图已保存: {screenshot_path.name}")
        return screenshot_path
        
    except Exception as e:
        LOGGER.debug(f"屏幕截图失败: {e}")
        return None


def log_error_with_screen_screenshot(
    error_message: str,
    output_dir: Path,
    exception: Optional[Exception] = None,
    prefix: str = "mole_error"
) -> None:
    """
    记录错误并自动截取屏幕（用于非浏览器应用）
    
    Args:
        error_message: 错误消息
        output_dir: 输出目录
        exception: 异常对象（可选）
        prefix: 文件名前缀
    """
    # 截图
    screenshot_path = capture_screen_screenshot(output_dir, error_message, prefix)
    
    # 记录错误
    if screenshot_path:
        LOGGER.error(f"❌ {error_message} [截图: {screenshot_path.name}]")
    else:
        LOGGER.error(f"❌ {error_message}")
    
    # 如果有异常对象，记录详细信息
    if exception:
        import traceback
        LOGGER.error(f"异常详情: {str(exception)}")
        LOGGER.debug(traceback.format_exc())

