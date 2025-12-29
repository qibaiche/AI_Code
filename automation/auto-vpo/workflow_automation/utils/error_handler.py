"""统一错误处理机制"""
import logging
import functools
from typing import Callable, Optional, Any
from pathlib import Path

LOGGER = logging.getLogger(__name__)


class WorkflowError(Exception):
    """工作流异常基类"""
    pass


class ElementNotFoundError(WorkflowError):
    """元素未找到异常"""
    pass


class OperationTimeoutError(WorkflowError):
    """操作超时异常"""
    pass


class ValidationError(WorkflowError):
    """验证失败异常"""
    pass


def retry_on_exception(
    max_retries: int = 3,
    delay: float = 2.0,
    exceptions: tuple = (Exception,),
    on_retry: Optional[Callable] = None
):
    """
    重试装饰器
    
    Args:
        max_retries: 最大重试次数
        delay: 重试延迟（秒）
        exceptions: 需要重试的异常类型
        on_retry: 重试时的回调函数
        
    Example:
        @retry_on_exception(max_retries=3, delay=2.0)
        def unstable_operation():
            # 可能失败的操作
            pass
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            import time
            
            last_exception = None
            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e
                    if attempt < max_retries:
                        LOGGER.warning(
                            f"操作失败 (尝试 {attempt + 1}/{max_retries + 1}): {e}"
                        )
                        if on_retry:
                            on_retry(attempt, e)
                        time.sleep(delay)
                    else:
                        LOGGER.error(
                            f"操作失败，已重试 {max_retries} 次: {e}"
                        )
            
            raise last_exception
        
        return wrapper
    return decorator


def handle_errors(
    default_return: Any = False,
    log_traceback: bool = True,
    capture_screenshot: bool = False,
    screenshot_prefix: str = "error"
):
    """
    错误处理装饰器
    
    Args:
        default_return: 发生异常时的默认返回值
        log_traceback: 是否记录完整的堆栈跟踪
        capture_screenshot: 是否捕获截图
        screenshot_prefix: 截图文件名前缀
        
    Example:
        @handle_errors(default_return=False, capture_screenshot=True)
        def risky_operation(self):
            # 可能抛出异常的操作
            pass
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # 记录错误
                LOGGER.error(f"执行 {func.__name__} 时发生错误: {e}")
                
                # 记录堆栈跟踪
                if log_traceback:
                    import traceback
                    LOGGER.debug(traceback.format_exc())
                
                # 捕获截图（如果可用）
                if capture_screenshot:
                    try:
                        # 尝试从 self 获取 driver 和 debug_dir
                        if args and hasattr(args[0], '_driver') and hasattr(args[0], 'debug_dir'):
                            from .screenshot_helper import log_error_with_screenshot
                            log_error_with_screenshot(
                                args[0]._driver,
                                f"{func.__name__} 失败",
                                args[0].debug_dir,
                                e,
                                screenshot_prefix
                            )
                    except Exception as screenshot_error:
                        LOGGER.debug(f"截图失败: {screenshot_error}")
                
                return default_return
        
        return wrapper
    return decorator


class ErrorContext:
    """错误上下文管理器"""
    
    def __init__(
        self,
        operation_name: str,
        logger: logging.Logger = LOGGER,
        raise_on_error: bool = True,
        default_return: Any = None
    ):
        """
        Args:
            operation_name: 操作名称
            logger: 日志记录器
            raise_on_error: 是否在错误时抛出异常
            default_return: 错误时的默认返回值
        """
        self.operation_name = operation_name
        self.logger = logger
        self.raise_on_error = raise_on_error
        self.default_return = default_return
        self.exception = None
    
    def __enter__(self):
        self.logger.debug(f"开始执行: {self.operation_name}")
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            self.exception = exc_val
            self.logger.error(f"{self.operation_name} 失败: {exc_val}")
            
            if self.raise_on_error:
                return False  # 重新抛出异常
            else:
                return True  # 抑制异常
        else:
            self.logger.debug(f"完成执行: {self.operation_name}")
            return True


def safe_execute(
    func: Callable,
    *args,
    default_return: Any = None,
    error_message: str = None,
    **kwargs
) -> Any:
    """
    安全执行函数，捕获所有异常
    
    Args:
        func: 要执行的函数
        *args: 函数参数
        default_return: 异常时的默认返回值
        error_message: 自定义错误消息
        **kwargs: 函数关键字参数
        
    Returns:
        函数返回值，或异常时的默认返回值
        
    Example:
        result = safe_execute(
            driver.find_element,
            By.ID, "button",
            default_return=None,
            error_message="未找到按钮"
        )
    """
    try:
        return func(*args, **kwargs)
    except Exception as e:
        if error_message:
            LOGGER.debug(f"{error_message}: {e}")
        else:
            LOGGER.debug(f"执行 {func.__name__} 时发生错误: {e}")
        return default_return


# 使用示例
"""
# 1. 使用重试装饰器
@retry_on_exception(max_retries=3, delay=2.0)
def click_button(self):
    button = self._driver.find_element(By.ID, "submit")
    button.click()

# 2. 使用错误处理装饰器
@handle_errors(default_return=False, capture_screenshot=True)
def fill_form(self, data):
    # 填写表单
    pass

# 3. 使用错误上下文
with ErrorContext("填写表单", raise_on_error=False):
    # 执行操作
    pass

# 4. 使用安全执行
element = safe_execute(
    driver.find_element,
    By.ID, "button",
    default_return=None,
    error_message="未找到按钮"
)
"""

