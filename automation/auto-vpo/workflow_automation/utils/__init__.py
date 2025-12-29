"""工具模块"""
from .screenshot_helper import (
    capture_error_screenshot,
    log_error_with_screenshot,
    capture_debug_screenshot,
    capture_screen_screenshot,
    log_error_with_screen_screenshot
)

from .wait_helpers import (
    wait_for_element,
    wait_for_element_disappear,
    wait_for_condition,
    wait_for_page_load,
    wait_for_ajax_complete,
    wait_for_angular_load,
    smart_wait,
    wait_and_click,
    wait_and_send_keys
)

from .error_handler import (
    WorkflowError,
    ElementNotFoundError,
    OperationTimeoutError,
    ValidationError,
    retry_on_exception,
    handle_errors,
    ErrorContext,
    safe_execute
)

from .keyboard_listener import (
    KeyboardListener,
    start_global_listener,
    stop_global_listener,
    is_esc_pressed,
    check_esc_and_exit
)

__all__ = [
    # 截图工具
    'capture_error_screenshot',
    'log_error_with_screenshot',
    'capture_debug_screenshot',
    'capture_screen_screenshot',
    'log_error_with_screen_screenshot',
    # 等待工具
    'wait_for_element',
    'wait_for_element_disappear',
    'wait_for_condition',
    'wait_for_page_load',
    'wait_for_ajax_complete',
    'wait_for_angular_load',
    'smart_wait',
    'wait_and_click',
    'wait_and_send_keys',
    # 错误处理
    'WorkflowError',
    'ElementNotFoundError',
    'OperationTimeoutError',
    'ValidationError',
    'retry_on_exception',
    'handle_errors',
    'ErrorContext',
    'safe_execute',
    # 键盘监听
    'KeyboardListener',
    'start_global_listener',
    'stop_global_listener',
    'is_esc_pressed',
    'check_esc_and_exit',
]

