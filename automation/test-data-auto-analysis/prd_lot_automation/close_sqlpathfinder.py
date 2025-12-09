"""
关闭 SQLPathFinder 窗口的辅助函数
"""
import logging
import time

try:
    import pyautogui
except ImportError:
    pyautogui = None

try:
    import win32gui
    import win32con
except ImportError:
    win32gui = None
    win32con = None


LOGGER = logging.getLogger(__name__)


def close_sqlpathfinder(main_window_title: str = "SQLPathFinder") -> bool:
    """
    关闭 SQLPathFinder 窗口（使用 Alt+F4 快捷键）
    
    Args:
        main_window_title: SQLPathFinder 主窗口标题
    
    Returns:
        是否成功关闭
    """
    closed_count = 0
    
    # 方法1: 使用 win32gui 查找并关闭所有 SQLPathFinder 窗口
    if win32gui:
        LOGGER.info("使用 win32gui 查找并关闭 SQLPathFinder 窗口...")
        
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if main_window_title in window_text:
                    results.append((hwnd, window_text))
        
        windows = []
        try:
            win32gui.EnumWindows(enum_windows_callback, windows)
            LOGGER.info(f"找到 {len(windows)} 个 SQLPathFinder 窗口")
            
            for hwnd, window_text in windows:
                try:
                    LOGGER.info(f"关闭窗口: {window_text}")
                    
                    # 激活窗口
                    win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3)
                    
                    # 发送 Alt+F4
                    if pyautogui:
                        pyautogui.hotkey('alt', 'F4')
                        LOGGER.info(f"✅ 已发送 Alt+F4")
                        time.sleep(0.5)
                        closed_count += 1
                    else:
                        # 如果没有 pyautogui，使用 WM_CLOSE
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        LOGGER.info(f"✅ 已发送 WM_CLOSE")
                        time.sleep(0.5)
                        closed_count += 1
                        
                except Exception as e:
                    LOGGER.warning(f"关闭窗口失败: {e}")
                    # 尝试强制关闭
                    try:
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        time.sleep(0.3)
                        closed_count += 1
                    except:
                        pass
                        
        except Exception as e:
            LOGGER.error(f"枚举窗口失败: {e}")
    
    # 方法2: 如果 win32gui 不可用或没找到窗口，尝试强制终止进程
    if closed_count == 0:
        LOGGER.info("尝试强制终止 SQLPathFinder 进程...")
        try:
            import psutil
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'SQLPathFinder' in proc.info['name']:
                        LOGGER.info(f"终止进程: {proc.info['name']} (PID: {proc.info['pid']})")
                        proc.kill()
                        closed_count += 1
                        time.sleep(0.5)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except ImportError:
            LOGGER.warning("psutil 未安装，无法强制终止进程")
        except Exception as e:
            LOGGER.error(f"强制终止进程失败: {e}")
    
    if closed_count > 0:
        LOGGER.info(f"✅ 成功关闭 {closed_count} 个 SQLPathFinder 窗口/进程")
        return True
    else:
        LOGGER.info("未找到需要关闭的 SQLPathFinder 窗口")
        return True

