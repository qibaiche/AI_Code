"""键盘监听工具 - 监听 ESC 键停止程序"""
import threading
import logging
from typing import Callable, Optional

LOGGER = logging.getLogger(__name__)

try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False
    LOGGER.warning("keyboard 模块未安装，ESC 键监听功能不可用。请运行: pip install keyboard")


class KeyboardListener:
    """键盘监听器 - 监听 ESC 键停止程序"""
    
    def __init__(self, on_escape: Optional[Callable] = None):
        """
        初始化键盘监听器
        
        Args:
            on_escape: ESC 键按下时的回调函数
        """
        self.on_escape = on_escape
        self._listening = False
        self._thread: Optional[threading.Thread] = None
        self._stop_flag = threading.Event()
        self._esc_pressed = False
    
    def start(self) -> bool:
        """
        开始监听键盘
        
        Returns:
            True 如果成功启动，False 如果失败
        """
        if not KEYBOARD_AVAILABLE:
            LOGGER.warning("⚠️ keyboard 模块未安装，无法监听 ESC 键")
            return False
        
        if self._listening:
            LOGGER.warning("键盘监听器已在运行")
            return False
        
        self._listening = True
        self._stop_flag.clear()
        self._esc_pressed = False
        
        # 在后台线程中监听
        self._thread = threading.Thread(target=self._listen_loop, daemon=True)
        self._thread.start()
        
        LOGGER.info("✅ 键盘监听器已启动（按 ESC 键可停止程序）")
        return True
    
    def stop(self) -> None:
        """停止监听键盘"""
        if not self._listening:
            return
        
        self._listening = False
        self._stop_flag.set()
        
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)
        
        LOGGER.info("键盘监听器已停止")
    
    def _listen_loop(self) -> None:
        """监听循环（在后台线程中运行）"""
        try:
            # 注册 ESC 键监听
            keyboard.on_press_key('esc', self._on_esc_pressed)
            
            # 等待停止信号
            while self._listening and not self._stop_flag.is_set():
                self._stop_flag.wait(timeout=0.1)
            
            # 取消注册
            keyboard.unhook_all()
            
        except Exception as e:
            LOGGER.error(f"键盘监听出错: {e}")
            self._listening = False
    
    def _on_esc_pressed(self, event) -> None:
        """ESC 键按下时的处理"""
        if event.name == 'esc' and not self._esc_pressed:
            self._esc_pressed = True
            LOGGER.warning("⚠️ 检测到 ESC 键按下，程序将停止...")
            
            # 调用回调函数
            if self.on_escape:
                try:
                    self.on_escape()
                except Exception as e:
                    LOGGER.error(f"执行 ESC 回调函数时出错: {e}")
    
    def is_esc_pressed(self) -> bool:
        """
        检查 ESC 键是否被按下
        
        Returns:
            True 如果 ESC 键被按下
        """
        return self._esc_pressed
    
    def reset_esc_flag(self) -> None:
        """重置 ESC 键标志（用于继续执行）"""
        self._esc_pressed = False


# 全局键盘监听器实例
_global_listener: Optional[KeyboardListener] = None


def start_global_listener(on_escape: Optional[Callable] = None) -> bool:
    """
    启动全局键盘监听器
    
    Args:
        on_escape: ESC 键按下时的回调函数
        
    Returns:
        True 如果成功启动
    """
    global _global_listener
    
    if _global_listener is None:
        _global_listener = KeyboardListener(on_escape)
    
    return _global_listener.start()


def stop_global_listener() -> None:
    """停止全局键盘监听器"""
    global _global_listener
    
    if _global_listener:
        _global_listener.stop()
        _global_listener = None


def is_esc_pressed() -> bool:
    """
    检查 ESC 键是否被按下（全局）
    
    Returns:
        True 如果 ESC 键被按下
    """
    if _global_listener:
        return _global_listener.is_esc_pressed()
    return False


def check_esc_and_exit(message: str = "程序已停止") -> None:
    """
    检查 ESC 键，如果按下则退出程序
    
    Args:
        message: 退出时的消息
    """
    if is_esc_pressed():
        LOGGER.warning(f"⚠️ {message}")
        import sys
        sys.exit(0)

