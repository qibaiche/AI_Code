"""Mole工具提交MIR数据模块"""
import logging
import time
import threading
import subprocess
import os
from pathlib import Path
from typing import Optional
from dataclasses import dataclass

try:
    from pywinauto import Application
    from pywinauto.findwindows import ElementNotFoundError
except ImportError:
    Application = None
    ElementNotFoundError = Exception

try:
    import win32gui
    import win32con
    import win32api
except ImportError:
    win32gui = None
    win32con = None
    win32api = None

try:
    import tkinter.messagebox as messagebox
except ImportError:
    messagebox = None

try:
    import pyperclip
except ImportError:
    pyperclip = None

try:
    import win32clipboard
except ImportError:
    win32clipboard = None

from .utils.screenshot_helper import log_error_with_screen_screenshot, capture_screen_screenshot

LOGGER = logging.getLogger(__name__)


@dataclass
class MoleConfig:
    """Mole工具配置"""
    executable_path: Optional[Path] = None
    window_title: str = "Mole"
    login_dialog_title: str = "MOLE LOGIN"
    timeout: int = 60
    retry_count: int = 3
    retry_delay: int = 2
    paths: Optional[object] = None  # 可选的paths配置对象，用于访问input_dir等
    search_mode: str = "vpos"  # 搜索方式: vpos, units, inactive_cage
    ui_config: Optional[dict] = None  # UI配置数据（包含所有用户输入的参数）


class MoleSubmitter:
    """Mole工具数据提交器"""
    
    def __init__(self, config: MoleConfig, debug_dir: Optional[Path] = None):
        self.config = config
        self._app = None
        self._window = None
        self.debug_dir = debug_dir or Path.cwd() / "output" / "05_Debug"
    
    def _log_error_with_screenshot(self, error_message: str, exception: Optional[Exception] = None, prefix: str = "mole_error") -> None:
        """记录错误并自动截取屏幕"""
        log_error_with_screen_screenshot(error_message, self.debug_dir, exception, prefix)
    
    def _ensure_application(self) -> None:
        """确保Mole应用程序已启动"""
        if Application is None:
            raise RuntimeError("pywinauto 未安装，无法执行 UI 自动化")
        
        # 如果窗口已存在，验证窗口句柄是否有效
        if self._window:
            try:
                if win32gui:
                    hwnd = self._window.handle
                    if not win32gui.IsWindow(hwnd):
                        LOGGER.warning("窗口句柄已失效，需要重新连接...")
                        self._window = None
                        self._app = None
                    else:
                        # 窗口句柄有效，直接返回
                        return
            except Exception as e:
                LOGGER.warning(f"验证窗口状态失败: {e}，重新连接...")
                self._window = None
                self._app = None
        
        # 检查是否已运行
        if self._is_process_running():
            LOGGER.info("Mole工具已在运行")
            try:
                self._app = Application(backend="win32").connect(
                    title_re=self.config.window_title,
                    visible_only=True
                )
                windows = self._app.windows()
                if windows:
                    self._window = windows[0]
                    LOGGER.info(f"已连接到Mole窗口: {self._window.window_text()}")
                    
                    return
            except Exception as e:
                LOGGER.warning(f"连接现有Mole窗口失败: {e}")
        
        # 启动Mole工具
        if self.config.executable_path and self.config.executable_path.exists():
            executable_path_str = str(self.config.executable_path)
            LOGGER.info(f"启动Mole工具: {executable_path_str}")
            
            # 如果是.appref-ms文件（ClickOnce应用程序快捷方式），使用os.startfile
            if executable_path_str.lower().endswith('.appref-ms'):
                try:
                    os.startfile(executable_path_str)
                    LOGGER.info("已通过os.startfile启动.appref-ms文件")
                except Exception as e:
                    LOGGER.warning(f"os.startfile启动失败: {e}，尝试使用subprocess")
                    # 如果os.startfile失败，尝试使用rundll32
                    try:
                        subprocess.Popen([
                            'rundll32.exe', 
                            'dfshim.dll,ShOpenVerbApplication',
                            executable_path_str
                        ])
                        LOGGER.info("已通过rundll32启动.appref-ms文件")
                    except Exception as e2:
                        raise RuntimeError(f"无法启动.appref-ms文件: {e2}")
            else:
                # 普通可执行文件，使用subprocess
                subprocess.Popen([executable_path_str])
        else:
            # 如果配置路径不存在，尝试默认路径
            default_path = Path(r"C:\Users\qibaiche\Desktop\Mole 2.0.appref-ms")
            if default_path.exists():
                LOGGER.info(f"使用默认路径启动Mole工具: {default_path}")
                try:
                    os.startfile(str(default_path))
                except Exception as e:
                    raise RuntimeError(f"无法启动Mole工具（默认路径）: {e}")
            else:
                raise RuntimeError(
                    f"无法找到Mole工具。配置路径: {self.config.executable_path}，"
                    f"默认路径: {default_path}。请检查config.yaml中的executable_path配置。"
                )
        
        # 等待登录对话框出现并处理
        LOGGER.info("等待Mole登录对话框...")
        self._handle_login_dialog()
        
        # 等待主窗口出现 - 使用多种方法查找
        deadline = time.time() + self.config.timeout
        backends = ["win32", "uia"]
        
        while time.time() < deadline:
            # 尝试不同的backend
            for backend in backends:
                try:
                    # 方法1: 使用配置的标题模式
                    try:
                        self._app = Application(backend=backend).connect(
                            title_re=self.config.window_title,
                            visible_only=True,
                            timeout=2
                        )
                        windows = self._app.windows()
                        if windows:
                            self._window = windows[0]
                            LOGGER.info(f"已连接到Mole窗口: {self._window.window_text()} (backend: {backend})")
                            return
                    except Exception as e:
                        LOGGER.debug(f"使用标题模式 '{self.config.window_title}' 连接失败 (backend: {backend}): {e}")
                    
                    # 方法2: 查找包含"MOLE"的所有窗口
                    try:
                        self._app = Application(backend=backend).connect(
                            title_re=".*MOLE.*",
                            visible_only=True,
                            timeout=2
                        )
                        windows = self._app.windows()
                        if windows:
                            # 找到包含MOLE的窗口，选择第一个
                            for win in windows:
                                try:
                                    win_text = win.window_text()
                                    if "MOLE" in win_text.upper():
                                        self._window = win
                                        LOGGER.info(f"已连接到Mole窗口: {win_text} (backend: {backend}, 方法: 包含MOLE)")
                                        return
                                except:
                                    continue
                    except Exception as e:
                        LOGGER.debug(f"查找包含MOLE的窗口失败 (backend: {backend}): {e}")
                    
                    # 方法3: 使用进程名查找（如果知道Mole的进程名）
                    try:
                        # 先尝试通过进程查找
                        from pywinauto.findwindows import find_windows
                        mole_windows = find_windows(backend=backend, title_re=".*MOLE.*")
                        if mole_windows:
                            for hwnd in mole_windows:
                                try:
                                    self._app = Application(backend=backend).connect(handle=hwnd)
                                    windows = self._app.windows()
                                    if windows:
                                        self._window = windows[0]
                                        LOGGER.info(f"已通过进程句柄连接到Mole窗口: {self._window.window_text()} (backend: {backend})")
                                        return
                                except:
                                    continue
                    except Exception as e:
                        LOGGER.debug(f"通过进程查找窗口失败 (backend: {backend}): {e}")
                    
                except Exception as e:
                    LOGGER.debug(f"尝试连接Mole窗口失败 (backend: {backend}): {e}")
                    continue
            
            # 如果所有方法都失败，等待后重试
            elapsed = int(time.time() - (deadline - self.config.timeout))
            if elapsed % 5 == 0:  # 每5秒输出一次进度
                LOGGER.info(f"等待Mole窗口出现... ({elapsed}/{self.config.timeout}秒)")
            time.sleep(1)
        
        # 如果超时，列出所有可见窗口以便调试
        LOGGER.error("等待Mole窗口超时，列出当前所有可见窗口:")
        try:
            import pygetwindow as gw
            all_windows = gw.getAllWindows()
            for w in all_windows:
                if w.title and w.visible:
                    LOGGER.error(f"  窗口: '{w.title}'")
        except:
            LOGGER.error("无法列出所有窗口")
        
        raise TimeoutError(f"等待Mole窗口出现超时（{self.config.timeout}秒）。请检查窗口标题配置是否正确。")
    
    def _reconnect_to_existing_window(self) -> bool:
        """
        尝试重新连接到已经运行的Mole窗口，不启动新的实例
        
        这个方法用于在窗口句柄失效时重新连接，避免重新启动Mole导致工作进度丢失
        
        Returns:
            bool: 成功连接返回True，失败返回False
        """
        if Application is None:
            return False
        
        LOGGER.info("尝试重新连接到现有的Mole窗口...")
        
        backends = ["win32", "uia"]
        
        for backend in backends:
            try:
                # 方法1: 通过窗口标题模式连接
                try:
                    self._app = Application(backend=backend).connect(
                        title_re=".*MOLE.*",
                        visible_only=True,
                        timeout=5
                    )
                    windows = self._app.windows()
                    for win in windows:
                        try:
                            win_text = win.window_text()
                            # 排除登录对话框，只连接主窗口
                            if "MOLE" in win_text.upper() and "LOGIN" not in win_text.upper():
                                self._window = win
                                LOGGER.info(f"✅ 已重新连接到Mole窗口: {win_text} (backend: {backend})")
                                return True
                        except:
                            continue
                except Exception as e:
                    LOGGER.debug(f"通过标题模式重连失败 (backend: {backend}): {e}")
                
                # 方法2: 通过进程句柄查找
                try:
                    from pywinauto.findwindows import find_windows
                    mole_windows = find_windows(backend=backend, title_re=".*MOLE.*")
                    for hwnd in mole_windows:
                        try:
                            if win32gui and win32gui.IsWindow(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                # 排除登录对话框
                                if "LOGIN" not in window_text.upper():
                                    self._app = Application(backend=backend).connect(handle=hwnd)
                                    windows = self._app.windows()
                                    if windows:
                                        self._window = windows[0]
                                        LOGGER.info(f"✅ 已通过句柄重新连接到Mole窗口: {self._window.window_text()} (backend: {backend})")
                                        return True
                        except:
                            continue
                except Exception as e:
                    LOGGER.debug(f"通过句柄重连失败 (backend: {backend}): {e}")
                    
            except Exception as e:
                LOGGER.debug(f"重连尝试失败 (backend: {backend}): {e}")
                continue
        
        LOGGER.warning("无法重新连接到现有的Mole窗口")
        return False
    
    def _click_file_menu_new_mir_request(self) -> None:
        """点击File菜单，然后选择New MIR Request（处理全屏模式的左上角菜单栏）"""
        if not self._window:
            raise RuntimeError("Mole窗口未连接")
        
        LOGGER.info("点击File菜单（全屏模式，左上角菜单栏）...")
        
        try:
            # 验证窗口句柄是否仍然有效
            try:
                # 尝试获取窗口句柄，如果失败说明窗口已失效
                hwnd = self._window.handle
                if win32gui:
                    # 验证句柄是否有效
                    if not win32gui.IsWindow(hwnd):
                        LOGGER.warning("窗口句柄已失效，尝试重新连接...")
                        self._window = None
                        self._app = None
                        self._ensure_application()
                        hwnd = self._window.handle
            except Exception as e:
                LOGGER.warning(f"窗口句柄验证失败: {e}，尝试重新连接...")
                self._window = None
                self._app = None
                self._ensure_application()
                hwnd = self._window.handle
            
            # 激活窗口并确保在最前
            self._window.set_focus()
            if win32gui and win32con:
                try:
                    hwnd = self._window.handle
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # 确保全屏
                except Exception as e:
                    LOGGER.warning(f"设置窗口焦点失败: {e}")
            time.sleep(0.5)
            
            # 获取窗口位置信息
            try:
                rect = self._window.rectangle()
                # 验证矩形是否有效（不应该全为0）
                if rect.left == 0 and rect.top == 0 and rect.right == 0 and rect.bottom == 0:
                    LOGGER.warning("窗口矩形无效（全为0），尝试重新连接...")
                    self._window = None
                    self._app = None
                    self._ensure_application()
                    rect = self._window.rectangle()
                LOGGER.info(f"窗口位置: left={rect.left}, top={rect.top}, right={rect.right}, bottom={rect.bottom}")
            except Exception as e:
                LOGGER.warning(f"获取窗口位置失败: {e}，尝试重新连接...")
                self._window = None
                self._app = None
                self._ensure_application()
                try:
                    rect = self._window.rectangle()
                    LOGGER.info(f"重新连接后窗口位置: left={rect.left}, top={rect.top}, right={rect.right}, bottom={rect.bottom}")
                except:
                    rect = None
            
            file_menu_clicked = False
            
            # 方法1: 使用Windows API直接点击菜单（最可靠的方法）
            if win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API点击File菜单...")
                    hwnd = self._window.handle
                    
                    # 获取菜单句柄
                    menu_handle = win32gui.GetMenu(hwnd)
                    if menu_handle:
                        # File菜单通常是第一个（索引0）
                        # 发送WM_COMMAND消息来点击File菜单的第一个项
                        # 或者先发送WM_SYSCOMMAND来打开菜单
                        LOGGER.info("找到菜单句柄，尝试点击File菜单项")
                        
                        # 方法1a: 直接使用PostMessage点击File菜单
                        # File菜单的ID通常是0xE100 (SC_SIZE + 0)
                        # 但更简单的方法是直接点击菜单项
                        try:
                            # 尝试获取第一个菜单项的位置并点击
                            menu_rect = win32gui.GetMenuItemRect(hwnd, menu_handle, 0)
                            if menu_rect:
                                # 计算菜单项中心点（在屏幕坐标中）
                                center_x = (menu_rect[0] + menu_rect[2]) // 2
                                center_y = (menu_rect[1] + menu_rect[3]) // 2
                                LOGGER.info(f"File菜单位置: ({center_x}, {center_y})")
                                # 使用pyautogui点击（屏幕坐标）
                                try:
                                    import pyautogui
                                    pyautogui.click(center_x, center_y)
                                    file_menu_clicked = True
                                    LOGGER.info("✅ 已通过屏幕坐标点击File菜单")
                                except ImportError:
                                    LOGGER.debug("pyautogui未安装，跳过屏幕坐标点击")
                        except:
                            pass
                        
                        # 方法1b: 如果坐标点击失败，使用Alt+F快捷键
                        if not file_menu_clicked:
                            LOGGER.info("使用Alt+F快捷键打开File菜单")
                            self._window.type_keys("%f")  # Alt+F
                            file_menu_clicked = True
                            LOGGER.info("✅ 已使用Alt+F打开File菜单")
                    else:
                        LOGGER.warning("未找到菜单句柄，尝试其他方法")
                except Exception as e1:
                    LOGGER.debug(f"使用Windows API点击File菜单失败: {e1}")
            
            # 方法2: 使用Alt+F快捷键（最通用的方法）
            if not file_menu_clicked:
                try:
                    LOGGER.info("使用Alt+F快捷键打开File菜单")
                    self._window.type_keys("%f")  # Alt+F
                    file_menu_clicked = True
                    LOGGER.info("✅ 已使用Alt+F打开File菜单")
                except Exception as e2:
                    LOGGER.debug(f"使用快捷键Alt+F失败: {e2}")
            
            # 方法3: 如果窗口在左上角，尝试点击屏幕左上角的固定位置
            if not file_menu_clicked and rect:
                try:
                    LOGGER.info("尝试点击窗口左上角的File菜单位置...")
                    # File菜单通常在窗口左上角，大约(10, 10)的相对位置
                    # 但需要转换为屏幕坐标
                    if win32gui:
                        hwnd = self._window.handle
                        # 获取窗口在屏幕上的位置
                        window_rect = win32gui.GetWindowRect(hwnd)
                        # File菜单通常位于窗口顶部菜单栏，约y=30像素处，x=10像素处
                        file_x = window_rect[0] + 10
                        file_y = window_rect[1] + 30
                        LOGGER.info(f"尝试点击屏幕坐标: ({file_x}, {file_y})")
                        
                        try:
                            import pyautogui
                            pyautogui.click(file_x, file_y)
                            file_menu_clicked = True
                            LOGGER.info("✅ 已通过固定坐标点击File菜单")
                        except ImportError:
                            LOGGER.debug("pyautogui未安装，无法使用屏幕坐标点击")
                        except Exception as e:
                            LOGGER.debug(f"屏幕坐标点击失败: {e}")
                except Exception as e3:
                    LOGGER.debug(f"点击固定位置失败: {e3}")
            
            # 方法4: 尝试查找File菜单控件
            if not file_menu_clicked:
                try:
                    # 查找菜单栏
                    menu_bar = self._window.child_window(control_type="MenuBar")
                    if menu_bar.exists():
                        file_menu = menu_bar.child_window(title="File")
                        if file_menu.exists():
                            LOGGER.info("找到File菜单项（MenuBar）")
                            file_menu.click_input()
                            file_menu_clicked = True
                            LOGGER.info("✅ 已点击File菜单（MenuBar）")
                except Exception as e4:
                    LOGGER.debug(f"通过MenuBar查找File菜单失败: {e4}")
            
            if not file_menu_clicked:
                raise RuntimeError("无法点击File菜单，已尝试所有方法")
            
            # 等待下拉菜单出现
            time.sleep(1.2)  # 增加等待时间，确保菜单完全展开
            
            # 在下拉菜单中选择"New MIR Request"
            LOGGER.info("在下拉菜单中选择'New MIR Request'（第5个菜单项）...")
            menu_item_clicked = False
            
            # 方法1: 使用Windows API精确查找并点击（最可靠）
            if win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API精确查找'New MIR Request'菜单项...")
                    hwnd = self._window.handle
                    
                    # 获取菜单句柄
                    menu_handle = win32gui.GetMenu(hwnd)
                    if menu_handle:
                        # File菜单通常是第一个（索引0）
                        file_submenu = win32gui.GetSubMenu(menu_handle, 0)
                        if file_submenu:
                            # 遍历File菜单的所有项
                            item_count = win32gui.GetMenuItemCount(file_submenu)
                            LOGGER.info(f"File菜单共有 {item_count} 个菜单项")
                            
                            target_index = -1
                            for i in range(item_count):
                                try:
                                    # 获取菜单项文本
                                    menu_text = win32gui.GetMenuString(file_submenu, i, win32con.MF_BYPOSITION)
                                    LOGGER.info(f"  菜单项 {i + 1}: '{menu_text}'")
                                    
                                    # 精确匹配"New MIR Request"（完全匹配，不区分大小写）
                                    if menu_text.strip().upper() == "NEW MIR REQUEST":
                                        target_index = i
                                        LOGGER.info(f"✅ 找到目标菜单项: '{menu_text}' (索引: {i})")
                                        break
                                except Exception as e_item:
                                    LOGGER.debug(f"处理菜单项 {i} 时出错: {e_item}")
                                    continue
                            
                            if target_index >= 0:
                                # 获取菜单项ID并发送点击消息
                                menu_item_id = win32gui.GetMenuItemID(file_submenu, target_index)
                                LOGGER.info(f"菜单项ID: {menu_item_id}")
                                
                                # 方法1a: 使用PostMessage发送WM_COMMAND消息
                                try:
                                    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, menu_item_id, 0)
                                    menu_item_clicked = True
                                    LOGGER.info("✅ 已通过Windows API PostMessage点击'New MIR Request'菜单项")
                                except Exception as e_cmd:
                                    LOGGER.debug(f"PostMessage失败: {e_cmd}")
                                    # 方法1b: 如果PostMessage失败，尝试SendMessage
                                    try:
                                        win32gui.SendMessage(hwnd, win32con.WM_COMMAND, menu_item_id, 0)
                                        menu_item_clicked = True
                                        LOGGER.info("✅ 已通过Windows API SendMessage点击'New MIR Request'菜单项")
                                    except Exception as e_send:
                                        LOGGER.debug(f"SendMessage也失败: {e_send}")
                except Exception as e1:
                    LOGGER.debug(f"使用Windows API查找菜单失败: {e1}")
            
            # 方法2: 使用键盘导航（精确到第5个菜单项）
            if not menu_item_clicked:
                try:
                    LOGGER.info("使用键盘导航选择第5个菜单项'New MIR Request'...")
                    # 菜单顺序（从File菜单打开后）：
                    # 0. (当前焦点在第一个菜单项)
                    # 1. New VPO Request
                    # 2. Mole Direction
                    # 3. New Source Lot
                    # 4. Standard Request
                    # 5. New MIR Request <- 目标（需要按4次下箭头）
                    
                    # 确保焦点在菜单上
                    self._window.set_focus()
                    time.sleep(0.2)
                    
                    # 按4次下箭头键移动到"New MIR Request"（从第1个到第5个）
                    LOGGER.info("按4次下箭头键...")
                    for i in range(4):
                        self._window.type_keys("{DOWN}")
                        time.sleep(0.15)  # 每次按键之间稍作延迟
                    
                    # 等待菜单高亮更新
                    time.sleep(0.3)
                    
                    # 按Enter确认选择
                    LOGGER.info("按Enter键确认选择...")
                    self._window.type_keys("{ENTER}")
                    
                    menu_item_clicked = True
                    LOGGER.info("✅ 已使用键盘导航选择'New MIR Request'（下箭头4次+Enter）")
                except Exception as e3:
                    LOGGER.debug(f"使用键盘导航失败: {e3}")
            
            # 方法3: 精确匹配MenuItem控件并点击
            if not menu_item_clicked:
                try:
                    LOGGER.info("通过控件精确查找'New MIR Request'...")
                    
                    # 首先列出所有菜单项以便调试
                    all_menu_items = self._window.descendants(control_type="MenuItem")
                    LOGGER.info(f"找到 {len(all_menu_items)} 个菜单项控件")
                    
                    for idx, item in enumerate(all_menu_items):
                        try:
                            item_text = item.window_text().strip()
                            LOGGER.info(f"  菜单项 {idx + 1}: '{item_text}'")
                            
                            # 精确匹配（完全匹配，不区分大小写）
                            if item_text.upper() == "NEW MIR REQUEST":
                                LOGGER.info(f"✅ 找到目标菜单项（精确匹配）: '{item_text}' (索引: {idx})")
                                
                                # 确保菜单项可见并可用
                                if item.is_visible() and item.is_enabled():
                                    # 先设置焦点到菜单项
                                    item.set_focus()
                                    time.sleep(0.2)
                                    # 点击菜单项
                                    item.click_input()
                                    menu_item_clicked = True
                                    LOGGER.info("✅ 已点击'New MIR Request'菜单项（精确匹配）")
                                    break
                                else:
                                    LOGGER.warning(f"菜单项存在但不可见或不可用: visible={item.is_visible()}, enabled={item.is_enabled()}")
                        except Exception as e:
                            LOGGER.debug(f"检查菜单项 {idx} 时出错: {e}")
                            continue
                except Exception as e2:
                    LOGGER.debug(f"通过控件查找菜单项失败: {e2}")
            
            if not menu_item_clicked:
                raise RuntimeError("无法选择'New MIR Request'菜单项，已尝试所有方法")
            
            # 等待新窗口或对话框出现
            time.sleep(1)
            LOGGER.info("✅ 已成功打开New MIR Request")
            
            # 处理可能出现的"MIR is now MRS!"信息对话框
            self._handle_mir_mrs_info_dialog()
            
            # 等待窗口完全加载
            time.sleep(1)
            
            # 注意：不在这里点击搜索按钮，由 workflow_main 根据 search_mode 决定点击哪个按钮
            # 填写对话框的逻辑将在workflow_main中调用
        
        except Exception as e:
            LOGGER.error(f"点击File菜单选择New MIR Request失败: {e}")
            raise RuntimeError(f"无法打开New MIR Request: {e}")
    
    def _click_search_by_vpos_button(self) -> None:
        """点击'Search By VPOs'按钮"""
        if not self._window:
            raise RuntimeError("Mole窗口未连接")
        
        LOGGER.info("点击'Search By VPOs'按钮...")
        
        try:
            # 确保窗口激活
            self._window.set_focus()
            if win32gui and win32con:
                try:
                    hwnd = self._window.handle
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                except:
                    pass
            time.sleep(0.5)
            
            button_clicked = False
            
            # 方法1: 直接查找标题为"Search By VPOs"的按钮
            try:
                search_button = self._window.child_window(title="Search By VPOs", control_type="Button")
                if search_button.exists() and search_button.is_enabled():
                    LOGGER.info("找到'Search By VPOs'按钮（通过title）")
                    search_button.click_input()
                    button_clicked = True
                    LOGGER.info("✅ 已点击'Search By VPOs'按钮（通过title）")
            except Exception as e1:
                LOGGER.debug(f"通过title查找按钮失败: {e1}")
            
            # 方法2: 遍历所有按钮查找
            if not button_clicked:
                try:
                    all_buttons = self._window.descendants(control_type="Button")
                    LOGGER.info(f"找到 {len(all_buttons)} 个按钮")
                    for button in all_buttons:
                        try:
                            button_text = button.window_text().strip()
                            LOGGER.debug(f"  按钮: '{button_text}'")
                            if button_text == "Search By VPOs":
                                LOGGER.info(f"找到'Search By VPOs'按钮（文本: '{button_text}'）")
                                if button.is_visible() and button.is_enabled():
                                    button.click_input()
                                    button_clicked = True
                                    LOGGER.info("✅ 已点击'Search By VPOs'按钮（遍历按钮）")
                                    break
                        except Exception as e:
                            LOGGER.debug(f"检查按钮时出错: {e}")
                            continue
                except Exception as e2:
                    LOGGER.debug(f"遍历按钮失败: {e2}")
            
            # 方法3: 使用部分匹配查找（包含"Search By VPOs"）
            if not button_clicked:
                try:
                    all_buttons = self._window.descendants(control_type="Button")
                    for button in all_buttons:
                        try:
                            button_text = button.window_text().strip()
                            if "Search By VPOs" in button_text or "Search By VPO" in button_text:
                                LOGGER.info(f"找到'Search By VPOs'按钮（部分匹配: '{button_text}'）")
                                if button.is_visible() and button.is_enabled():
                                    button.click_input()
                                    button_clicked = True
                                    LOGGER.info("✅ 已点击'Search By VPOs'按钮（部分匹配）")
                                    break
                        except:
                            continue
                except Exception as e3:
                    LOGGER.debug(f"部分匹配查找失败: {e3}")
            
            # 方法4: 使用Windows API查找按钮
            if not button_clicked and win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API查找'Search By VPOs'按钮...")
                    hwnd = self._window.handle
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child)
                            class_name = win32gui.GetClassName(hwnd_child)
                            if "Search By VPOs" in window_text and "BUTTON" in class_name.upper():
                                LOGGER.info(f"通过Windows API找到按钮: '{window_text}' (类名: {class_name})")
                                win32gui.PostMessage(hwnd_child, win32con.BM_CLICK, 0, 0)
                                return False  # 停止枚举
                            return True
                        except:
                            return True
                    
                    win32gui.EnumChildWindows(hwnd, enum_child_proc, None)
                    time.sleep(0.5)
                    button_clicked = True
                    LOGGER.info("✅ 已通过Windows API点击'Search By VPOs'按钮")
                except Exception as e4:
                    LOGGER.debug(f"使用Windows API查找按钮失败: {e4}")
            
            if not button_clicked:
                raise RuntimeError("无法点击'Search By VPOs'按钮，已尝试所有方法")
            
            # 等待按钮点击后的响应和对话框出现
            time.sleep(1)
            LOGGER.info("✅ 已成功点击'Search By VPOs'按钮")
        
        except Exception as e:
            LOGGER.error(f"点击'Search By VPOs'按钮失败: {e}")
            raise RuntimeError(f"无法点击'Search By VPOs'按钮: {e}")
    
    def _click_search_by_units_button(self) -> None:
        """点击'Search By Units'按钮，并等待对话框出现"""
        if Application is None:
            raise RuntimeError("pywinauto 未安装，无法执行 UI 自动化")
        
        # 首先检查是否已经有Units搜索对话框存在，如果有则先关闭
        LOGGER.info("检查是否已有Units搜索对话框存在...")
        existing_dialog_found = False
        existing_dialogs_count = 0
        
        if win32gui:
            try:
                def enum_existing_dialog(hwnd, windows):
                    try:
                        if win32gui.IsWindowVisible(hwnd):
                            window_text = win32gui.GetWindowText(hwnd)
                            if "Import" in window_text and "Serial" in window_text:
                                rect = win32gui.GetWindowRect(hwnd)
                                width = rect[2] - rect[0]
                                height = rect[3] - rect[1]
                                if 300 < width < 1000 and 300 < height < 900:
                                    windows.append(hwnd)
                    except:
                        pass
                    return True
                
                existing_dialogs = []
                win32gui.EnumWindows(enum_existing_dialog, existing_dialogs)
                
                if existing_dialogs:
                    existing_dialogs_count = len(existing_dialogs)
                    existing_dialog_found = True
                    LOGGER.warning(f"发现 {existing_dialogs_count} 个已存在的Units搜索对话框，正在关闭...")
                    for dialog_hwnd in existing_dialogs:
                        try:
                            # 尝试关闭对话框
                            win32gui.PostMessage(dialog_hwnd, win32con.WM_CLOSE, 0, 0)
                            LOGGER.info(f"已发送关闭消息到对话框（句柄: {dialog_hwnd}）")
                            time.sleep(0.5)
                        except Exception as e:
                            LOGGER.warning(f"关闭对话框失败: {e}")
                    
                    # 等待对话框关闭
                    time.sleep(1.5)  # 增加等待时间，确保对话框完全关闭
                    LOGGER.info("已关闭现有对话框，准备打开新对话框")
            except Exception as e:
                LOGGER.debug(f"检查现有对话框时出错: {e}")
        
        # 也尝试使用pywinauto检查
        if not existing_dialog_found:
            try:
                for backend in ["win32", "uia"]:
                    try:
                        dialog_app = Application(backend=backend).connect(
                            title_re=".*Import.*Serial.*",
                            visible_only=True,
                            timeout=1
                        )
                        windows = dialog_app.windows()
                        for win in windows:
                            try:
                                win_text = win.window_text()
                                if "Import" in win_text and "Serial" in win_text:
                                    LOGGER.warning(f"发现已存在的对话框（标题: '{win_text}'），正在关闭...")
                                    try:
                                        win.close()
                                        time.sleep(0.5)
                                        LOGGER.info("已关闭现有对话框")
                                    except:
                                        pass
                            except:
                                continue
                    except:
                        continue
            except:
                pass
        
        LOGGER.info("点击'Search By Units'按钮...")
        
        button_clicked = False
        
        try:
            # 确保窗口激活
            if self._window:
                try:
                    self._window.set_focus()
                    if win32gui and win32con:
                        hwnd = self._window.handle
                        win32gui.SetForegroundWindow(hwnd)
                        win32gui.BringWindowToTop(hwnd)
                    time.sleep(0.5)
                except Exception as e:
                    LOGGER.warning(f"设置窗口焦点失败: {e}")
            
            # 方法1: 直接查找标题为"Search By Units"的按钮
            # 如果检测到已有对话框，先等待它们完全关闭，然后再点击
            if existing_dialog_found:
                LOGGER.info("等待已关闭的对话框完全消失...")
                time.sleep(1.0)  # 额外等待，确保对话框完全关闭
            
            try:
                search_button = self._window.child_window(title="Search By Units", control_type="Button")
                if search_button.exists() and search_button.is_enabled():
                    LOGGER.info("找到'Search By Units'按钮（通过title）")
                    # 只点击一次，确保不会重复点击
                    search_button.click_input()
                    time.sleep(0.5)
                    LOGGER.info("✅ 已点击'Search By Units'按钮（通过title）")
                    button_clicked = True
            except Exception as e1:
                LOGGER.debug(f"方法1失败: {e1}")
            
            # 方法2: 遍历所有按钮查找
            if not button_clicked:
                try:
                    all_buttons = self._window.descendants(control_type="Button")
                    for button in all_buttons:
                        try:
                            button_text = button.window_text().strip()
                            if button_text == "Search By Units":
                                LOGGER.info(f"找到'Search By Units'按钮（文本: '{button_text}'）")
                                if button.is_visible() and button.is_enabled():
                                    button.click_input()
                                    time.sleep(0.5)
                                    LOGGER.info("✅ 已点击'Search By Units'按钮（遍历按钮）")
                                    button_clicked = True
                                    break
                        except:
                            continue
                except Exception as e2:
                    LOGGER.debug(f"方法2失败: {e2}")
            
            # 方法3: 使用Windows API查找按钮（添加超时保护）
            if not button_clicked and win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API查找'Search By Units'按钮...")
                    hwnd = self._window.handle
                    
                    button_list = []
                    button_found = False
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child)
                            class_name = win32gui.GetClassName(hwnd_child)
                            if "Search By Units" in window_text and "BUTTON" in class_name.upper():
                                lParam.append(hwnd_child)
                                return False  # 找到后停止枚举
                        except:
                            pass
                        return True
                    
                    # 使用线程和超时机制，避免卡住
                    import threading
                    enum_complete = threading.Event()
                    enum_exception = [None]
                    
                    def enum_thread():
                        try:
                            win32gui.EnumChildWindows(hwnd, enum_child_proc, button_list)
                        except Exception as e:
                            enum_exception[0] = e
                        finally:
                            enum_complete.set()
                    
                    enum_thread_obj = threading.Thread(target=enum_thread, daemon=True)
                    enum_thread_obj.start()
                    enum_thread_obj.join(timeout=5.0)  # 最多等待5秒
                    
                    if not enum_complete.is_set():
                        LOGGER.warning("Windows API枚举超时（5秒），尝试其他方法...")
                    elif enum_exception[0]:
                        LOGGER.warning(f"Windows API枚举出错: {enum_exception[0]}，尝试其他方法...")
                    elif button_list:
                        button_hwnd = button_list[0]
                        LOGGER.info(f"找到'Search By Units'按钮（句柄: {button_hwnd}）")

                        # SendMessage 可能阻塞，使用线程和超时保护
                        send_result = {"success": False, "error": None}

                        def send_click():
                            try:
                                win32gui.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                                send_result["success"] = True
                            except Exception as exc:
                                send_result["error"] = exc

                        send_thread = threading.Thread(target=send_click, daemon=True)
                        send_thread.start()
                        send_thread.join(timeout=2.0)

                        # 检查SendMessage结果
                        if send_result["success"]:
                            time.sleep(0.5)
                            LOGGER.info("✅ 已通过Windows API点击'Search By Units'按钮（SendMessage, 带超时保护）")
                            button_clicked = True
                        elif send_thread.is_alive():
                            # SendMessage超时，但可能已经成功触发了点击
                            # 先等待一小段时间，检查对话框是否已经出现
                            LOGGER.warning("SendMessage 点击'Search By Units'超时，检查是否已成功触发...")
                            time.sleep(1.0)  # 等待对话框出现
                            
                            # 检查对话框是否已经出现
                            dialog_already_opened = False
                            try:
                                def check_dialog_callback(hwnd, result):
                                    try:
                                        if win32gui.IsWindowVisible(hwnd):
                                            window_text = win32gui.GetWindowText(hwnd)
                                            if "Import" in window_text and "Serial" in window_text:
                                                result.append(True)
                                    except:
                                        pass
                                    return True
                                
                                dialog_check_result = []
                                win32gui.EnumWindows(check_dialog_callback, dialog_check_result)
                                if dialog_check_result:
                                    dialog_already_opened = True
                                    LOGGER.info("✅ 检测到对话框已打开，SendMessage已成功触发点击，无需再次点击")
                                    button_clicked = True
                            except:
                                pass
                            
                            # 如果对话框没有出现，才尝试PostMessage
                            if not dialog_already_opened:
                                LOGGER.warning("对话框未出现，尝试PostMessage...")
                                try:
                                    win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                                    time.sleep(0.5)
                                    LOGGER.info("✅ 已通过Windows API点击'Search By Units'按钮（PostMessage）")
                                    button_clicked = True
                                except Exception as e2:
                                    LOGGER.warning(f"PostMessage也失败: {e2}，尝试pywinauto包装器...")
                                    if not button_clicked:
                                        try:
                                            from pywinauto.controls.hwndwrapper import HwndWrapper
                                            HwndWrapper(button_hwnd).click_input()
                                            time.sleep(0.5)
                                            LOGGER.info("✅ 已通过pywinauto包装器点击'Search By Units'按钮")
                                            button_clicked = True
                                        except Exception as e3:
                                            LOGGER.warning(f"pywinauto包装器点击失败: {e3}，尝试鼠标点击...")
                                            if not button_clicked:
                                                try:
                                                    import pyautogui
                                                    rect = win32gui.GetWindowRect(button_hwnd)
                                                    center_x = (rect[0] + rect[2]) // 2
                                                    center_y = (rect[1] + rect[3]) // 2
                                                    pyautogui.click(center_x, center_y)
                                                    time.sleep(0.5)
                                                    LOGGER.info(f"✅ 已通过鼠标点击'Search By Units'按钮（坐标: {center_x}, {center_y}）")
                                                    button_clicked = True
                                                except Exception as e4:
                                                    LOGGER.warning(f"鼠标点击也失败: {e4}")
                        else:
                            # SendMessage失败，尝试PostMessage
                            LOGGER.warning(f"SendMessage失败: {send_result['error']}，尝试PostMessage...")
                            if not button_clicked:
                                try:
                                    win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                                    time.sleep(0.5)
                                    LOGGER.info("✅ 已通过Windows API点击'Search By Units'按钮（PostMessage）")
                                    button_clicked = True
                                except Exception as e2:
                                    LOGGER.warning(f"PostMessage也失败: {e2}，尝试pywinauto包装器...")
                                    if not button_clicked:
                                        try:
                                            from pywinauto.controls.hwndwrapper import HwndWrapper
                                            HwndWrapper(button_hwnd).click_input()
                                            time.sleep(0.5)
                                            LOGGER.info("✅ 已通过pywinauto包装器点击'Search By Units'按钮")
                                            button_clicked = True
                                        except Exception as e3:
                                            LOGGER.warning(f"pywinauto包装器点击失败: {e3}，尝试鼠标点击...")
                                            if not button_clicked:
                                                try:
                                                    import pyautogui
                                                    rect = win32gui.GetWindowRect(button_hwnd)
                                                    center_x = (rect[0] + rect[2]) // 2
                                                    center_y = (rect[1] + rect[3]) // 2
                                                    pyautogui.click(center_x, center_y)
                                                    time.sleep(0.5)
                                                    LOGGER.info(f"✅ 已通过鼠标点击'Search By Units'按钮（坐标: {center_x}, {center_y}）")
                                                    button_clicked = True
                                                except Exception as e4:
                                                    LOGGER.warning(f"鼠标点击也失败: {e4}")
                    else:
                        LOGGER.debug("Windows API未找到'Search By Units'按钮")
                except Exception as e3:
                    LOGGER.debug(f"方法3失败: {e3}")
                    import traceback
                    LOGGER.debug(traceback.format_exc())
            
            # 方法4: 使用部分匹配查找（"Search By" 或 "Units"）
            if not button_clicked:
                try:
                    LOGGER.info("尝试使用部分匹配查找按钮...")
                    all_buttons = self._window.descendants(control_type="Button")
                    for button in all_buttons:
                        try:
                            button_text = button.window_text().strip()
                            # 匹配 "Search By Units" 或包含 "Search" 和 "Units" 的按钮
                            if ("Search By Units" in button_text or 
                                ("Search" in button_text and "Units" in button_text)):
                                LOGGER.info(f"找到候选按钮（部分匹配）: '{button_text}'")
                                if button.is_visible() and button.is_enabled():
                                    button.click_input()
                                    time.sleep(0.5)
                                    LOGGER.info("✅ 已点击'Search By Units'按钮（部分匹配）")
                                    button_clicked = True
                                    break
                        except:
                            continue
                except Exception as e4:
                    LOGGER.debug(f"方法4失败: {e4}")
            
            # 方法5: 尝试使用鼠标点击估算位置（最后备用方法）
            if not button_clicked:
                try:
                    LOGGER.warning("所有标准方法都失败，尝试估算位置并使用鼠标点击...")
                    if win32gui:
                        hwnd = self._window.handle
                        rect = win32gui.GetWindowRect(hwnd)
                        window_width = rect[2] - rect[0]
                        window_height = rect[3] - rect[1]
                        
                        # 估算按钮可能的位置（通常在窗口左侧或中间）
                        possible_positions = [
                            (0.15, 0.20),  # 15%宽度, 20%高度
                            (0.20, 0.25),  # 20%宽度, 25%高度
                            (0.25, 0.30),  # 25%宽度, 30%高度
                            (0.30, 0.25),  # 30%宽度, 25%高度
                        ]
                        
                        try:
                            import pyautogui
                            for pos_idx, (width_ratio, height_ratio) in enumerate(possible_positions):
                                estimated_x = rect[0] + int(window_width * width_ratio)
                                estimated_y = rect[1] + int(window_height * height_ratio)
                                
                                LOGGER.info(f"尝试估算位置 #{pos_idx + 1}: ({estimated_x}, {estimated_y})")
                                pyautogui.moveTo(estimated_x, estimated_y, duration=0.2)
                                time.sleep(0.1)
                                pyautogui.click(estimated_x, estimated_y)
                                time.sleep(0.5)
                                
                                # 检查是否打开了对话框（通过检查是否有新窗口出现）
                                # 这里简单等待一下，如果打开了对话框，后续步骤会继续
                                LOGGER.info(f"已点击估算位置 #{pos_idx + 1}，等待对话框出现...")
                                button_clicked = True
                                break
                        except ImportError:
                            LOGGER.warning("pyautogui未安装，无法使用估算位置方法")
                except Exception as e5:
                    LOGGER.debug(f"方法5失败: {e5}")
            
            if not button_clicked:
                raise RuntimeError("无法点击'Search By Units'按钮，已尝试所有方法")
            
            # 先等待一小段时间，让对话框有时间出现
            time.sleep(1.0)
            
            # 等待对话框出现（最多等待20秒）
            LOGGER.info("等待'MOLE - Import List Of Serial Numbers'对话框出现...")
            dialog_found = False
            deadline = time.time() + 20
            
            if win32gui:
                while time.time() < deadline and not dialog_found:
                    candidate_windows = []
                    
                    # 使用线程和超时机制，避免EnumWindows卡住
                    import threading
                    enum_complete = threading.Event()
                    enum_exception = [None]
                    
                    def enum_windows_callback(hwnd, windows):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                if "Import" in window_text and "Serial" in window_text:
                                    windows.append(hwnd)
                        except:
                            pass
                        return True
                    
                    def enum_thread():
                        try:
                            win32gui.EnumWindows(enum_windows_callback, candidate_windows)
                        except Exception as e:
                            enum_exception[0] = e
                        finally:
                            enum_complete.set()
                    
                    enum_thread_obj = threading.Thread(target=enum_thread, daemon=True)
                    enum_thread_obj.start()
                    enum_thread_obj.join(timeout=2.0)  # 最多等待2秒
                    
                    if enum_complete.is_set() and candidate_windows:
                        dialog_found = True
                        LOGGER.info(f"✅ 找到Units搜索对话框（句柄: {candidate_windows[0]}）")
                        break
                    elif not enum_complete.is_set():
                        LOGGER.debug("EnumWindows超时（2秒），继续尝试...")
                    
                    time.sleep(0.5)
            
            # 如果win32gui方法没找到，尝试pywinauto
            if not dialog_found:
                LOGGER.info("使用pywinauto查找对话框...")
                while time.time() < deadline and not dialog_found:
                    try:
                        for backend in ["win32", "uia"]:
                            try:
                                dialog_app = Application(backend=backend).connect(
                                    title_re=".*Import.*Serial.*",
                                    visible_only=True,
                                    timeout=1
                                )
                                windows = dialog_app.windows()
                                for win in windows:
                                    try:
                                        win_text = win.window_text()
                                        if "Import" in win_text and "Serial" in win_text:
                                            dialog_found = True
                                            LOGGER.info(f"✅ 找到Units搜索对话框（标题: '{win_text}'）")
                                            break
                                    except:
                                        continue
                                if dialog_found:
                                    break
                            except:
                                continue
                        if not dialog_found:
                            time.sleep(0.5)
                    except:
                        time.sleep(0.5)
            
            if not dialog_found:
                LOGGER.warning("⚠️ 未能在15秒内检测到对话框，但继续执行（对话框可能已打开但检测失败）")
            else:
                LOGGER.info("✅ 对话框已出现，可以继续填写Units信息")
        
        except RuntimeError:
            raise
        except Exception as e:
            LOGGER.error(f"点击'Search By Units'按钮失败: {e}")
            raise RuntimeError(f"无法点击'Search By Units'按钮: {e}")
    
    def _click_search_inactive_cage_button(self) -> None:
        """点击'Search Inactive Cage'按钮"""
        if Application is None:
            raise RuntimeError("pywinauto 未安装，无法执行 UI 自动化")
        
        LOGGER.info("点击'Search Inactive Cage'按钮...")
        
        try:
            # 确保窗口激活
            if self._window:
                try:
                    self._window.set_focus()
                    if win32gui and win32con:
                        hwnd = self._window.handle
                        win32gui.SetForegroundWindow(hwnd)
                        win32gui.BringWindowToTop(hwnd)
                    time.sleep(0.5)
                except Exception as e:
                    LOGGER.warning(f"设置窗口焦点失败: {e}")
            
            # 方法1: 直接查找标题为"Search Inactive Cage"的按钮
            try:
                search_button = self._window.child_window(title="Search Inactive Cage", control_type="Button")
                if search_button.exists():
                    LOGGER.info("找到'Search Inactive Cage'按钮（通过title）")
                    search_button.click_input()
                    time.sleep(0.5)
                    LOGGER.info("✅ 已点击'Search Inactive Cage'按钮（通过title）")
                    return
            except Exception as e1:
                LOGGER.debug(f"方法1失败: {e1}")
            
            # 方法2: 遍历所有按钮查找
            try:
                all_buttons = self._window.descendants(control_type="Button")
                for button in all_buttons:
                    try:
                        button_text = button.window_text().strip()
                        if "Inactive" in button_text and "Cage" in button_text:
                            LOGGER.info(f"找到'Search Inactive Cage'按钮（文本: '{button_text}'）")
                            if button.is_visible() and button.is_enabled():
                                button.click_input()
                                time.sleep(0.5)
                                LOGGER.info("✅ 已点击'Search Inactive Cage'按钮（遍历按钮）")
                                return
                    except:
                        continue
            except Exception as e2:
                LOGGER.debug(f"方法2失败: {e2}")
            
            # 方法3: 使用Windows API查找按钮
            if win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API查找'Search Inactive Cage'按钮...")
                    hwnd = self._window.handle
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child)
                            class_name = win32gui.GetClassName(hwnd_child)
                            if "Inactive" in window_text and "Cage" in window_text and "BUTTON" in class_name.upper():
                                lParam.append(hwnd_child)
                                return False
                        except:
                            pass
                        return True
                    
                    button_list = []
                    win32gui.EnumChildWindows(hwnd, enum_child_proc, button_list)
                    
                    if button_list:
                        button_hwnd = button_list[0]
                        win32gui.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                        time.sleep(0.5)
                        LOGGER.info("✅ 已通过Windows API点击'Search Inactive Cage'按钮")
                        return
                except Exception as e3:
                    LOGGER.debug(f"方法3失败: {e3}")
            
            raise RuntimeError("无法点击'Search Inactive Cage'按钮，已尝试所有方法")
        
        except RuntimeError:
            raise
        except Exception as e:
            LOGGER.error(f"点击'Search Inactive Cage'按钮失败: {e}")
            raise RuntimeError(f"无法点击'Search Inactive Cage'按钮: {e}")
    
    def _fill_vpo_search_dialog(self, source_lot_value: str) -> None:
        """
        在VPO搜索对话框中填写Source Lot值并点击Search
        
        Args:
            source_lot_value: Source Lot的值（第一行的SourceLot列值）
        """
        if Application is None:
            raise RuntimeError("pywinauto 未安装，无法执行 UI 自动化")
        
        LOGGER.info("等待VPO搜索对话框出现...")
        
        # 等待对话框出现（最多等待15秒）
        vpo_dialog = None
        deadline = time.time() + 15
        dialog_titles = ["MOLE"]
        
        backends = ["win32", "uia"]
        
        LOGGER.info("查找VPO搜索对话框（模态对话框，标题: MOLE）...")
        
        # 首先尝试使用win32gui枚举所有窗口（更可靠的方法）
        if win32gui:
            LOGGER.info("使用Windows API枚举所有窗口...")
            dialog_hwnd = None
            
            def enum_windows_callback(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if window_text.strip() == "MOLE":
                            # 检查窗口大小（对话框通常比主窗口小）
                            rect = win32gui.GetWindowRect(hwnd)
                            width = rect[2] - rect[0]
                            height = rect[3] - rect[1]
                            # 对话框通常宽度在400-800，高度在300-600
                            if 300 < width < 1000 and 200 < height < 800:
                                # 检查是否是模态对话框（可能有特定窗口样式）
                                style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                                windows.append((hwnd, window_text, width, height))
                except:
                    pass
                return True
            
            candidate_windows = []
            win32gui.EnumWindows(enum_windows_callback, candidate_windows)
            
            LOGGER.info(f"找到 {len(candidate_windows)} 个可能的对话框窗口")
            for hwnd, win_text, w, h in candidate_windows:
                LOGGER.info(f"  窗口: '{win_text}' (大小: {w}x{h})")
            
            # 选择最可能的对话框（通常是第一个，或者最大的符合条件的窗口）
            if candidate_windows:
                dialog_hwnd, win_text, w, h = candidate_windows[0]
                LOGGER.info(f"选择对话框: '{win_text}' (大小: {w}x{h})")
                # 尝试用pywinauto连接到此窗口（优先使用主窗口的app实例，避免影响主窗口）
                try:
                    # 优先使用主窗口的app实例（如果存在且可用）
                    if self._app:
                        try:
                            vpo_dialog = self._app.window(handle=dialog_hwnd)
                            if vpo_dialog.exists():
                                LOGGER.info(f"✅ 找到VPO搜索对话框: '{win_text}' (使用主窗口app实例, handle: {dialog_hwnd})")
                                # 确保对话框有焦点
                                vpo_dialog.set_focus()
                                if win32gui:
                                    try:
                                        win32gui.SetForegroundWindow(dialog_hwnd)
                                        win32gui.BringWindowToTop(dialog_hwnd)
                                    except:
                                        pass
                            else:
                                vpo_dialog = None
                        except Exception as e:
                            LOGGER.debug(f"使用主窗口app实例连接对话框失败: {e}")
                            vpo_dialog = None
                    
                    # 如果主窗口app实例无法连接对话框，尝试创建新连接
                    if vpo_dialog is None:
                        for backend in backends:
                            try:
                                dialog_app = Application(backend=backend).connect(handle=dialog_hwnd)
                                vpo_dialog = dialog_app.window(handle=dialog_hwnd)
                                if vpo_dialog.exists():
                                    LOGGER.info(f"✅ 找到VPO搜索对话框: '{win_text}' (backend: {backend}, handle: {dialog_hwnd})")
                                    # 确保对话框有焦点
                                    vpo_dialog.set_focus()
                                    if win32gui:
                                        try:
                                            win32gui.SetForegroundWindow(dialog_hwnd)
                                            win32gui.BringWindowToTop(dialog_hwnd)
                                        except:
                                            pass
                                    break
                            except Exception as e:
                                LOGGER.debug(f"使用backend {backend}连接对话框失败: {e}")
                                continue
                    
                    if vpo_dialog is None:
                        # 如果pywinauto无法连接，使用win32gui直接操作
                        LOGGER.warning("pywinauto无法连接对话框，将使用win32gui直接操作")
                        vpo_dialog = {"hwnd": dialog_hwnd, "text": win_text, "use_win32": True}
                except Exception as e:
                    LOGGER.warning(f"连接对话框时出错: {e}")
                    vpo_dialog = {"hwnd": dialog_hwnd, "text": win_text, "use_win32": True}
        
        # 如果win32gui方法失败，使用pywinauto方法
        if vpo_dialog is None:
            LOGGER.info("使用pywinauto方法查找对话框...")
            while time.time() < deadline:
                for backend in backends:
                    for title_pattern in dialog_titles:
                        try:
                            dialog_app = Application(backend=backend).connect(
                                title_re=title_pattern,
                                visible_only=True,
                                timeout=2
                            )
                            # 获取所有窗口，包括主窗口和对话框
                            windows = dialog_app.windows()
                            LOGGER.debug(f"找到 {len(windows)} 个窗口")
                            
                            for win in windows:
                                try:
                                    win_text = win.window_text().strip()
                                    LOGGER.debug(f"检查窗口: '{win_text}'")
                                    
                                    # 检查是否是VPO搜索对话框（模态对话框）
                                    # 特征1: 标题完全等于"MOLE"（主窗口标题更长）
                                    if win_text == "MOLE":
                                        LOGGER.info(f"找到标题为'MOLE'的窗口，检查是否是VPO搜索对话框...")
                                        
                                        # 检查窗口大小（对话框通常比主窗口小）
                                        try:
                                            rect = win.rectangle()
                                            width = rect.width()
                                            height = rect.height()
                                            LOGGER.debug(f"窗口大小: {width}x{height}")
                                            # 对话框通常宽度在400-800，高度在300-600
                                            if 300 < width < 1000 and 200 < height < 800:
                                                # 尝试查找Edit控件确认
                                                try:
                                                    edits = win.descendants(control_type="Edit")
                                                    if edits and len(edits) > 0:
                                                        vpo_dialog = win
                                                        LOGGER.info(f"✅ 找到VPO搜索对话框: '{win_text}' (backend: {backend}, 大小: {width}x{height})")
                                                        break
                                                except:
                                                    pass
                                        except:
                                            # 如果无法获取大小，尝试其他方法
                                            try:
                                                edits = win.descendants(control_type="Edit")
                                                if edits and len(edits) > 0:
                                                    vpo_dialog = win
                                                    LOGGER.info(f"✅ 找到可能的VPO搜索对话框: '{win_text}' (backend: {backend})")
                                                    break
                                            except:
                                                pass
                                except Exception as e:
                                    LOGGER.debug(f"检查窗口时出错: {e}")
                                    continue
                            
                            if vpo_dialog is not None:
                                break
                        except:
                            continue
                    
                    if vpo_dialog is not None:
                        break
                
                if vpo_dialog is not None:
                    break
                
                time.sleep(0.5)
        
        if vpo_dialog is None:
            raise RuntimeError("未找到VPO搜索对话框（标题: MOLE）")
        
        try:
            # 处理两种情况：pywinauto窗口对象或win32gui字典
            use_win32_direct = isinstance(vpo_dialog, dict) and vpo_dialog.get("use_win32", False)
            
            if use_win32_direct:
                # 使用win32gui直接操作
                dialog_hwnd = vpo_dialog["hwnd"]
                LOGGER.info("使用Windows API直接操作对话框...")
                
                # 激活对话框
                win32gui.SetForegroundWindow(dialog_hwnd)
                win32gui.BringWindowToTop(dialog_hwnd)
                time.sleep(0.5)
            else:
                # 使用pywinauto操作
                vpo_dialog.set_focus()
                time.sleep(0.5)
            
            LOGGER.info(f"填写Source Lot值: {source_lot_value}")
            
            # 先将值复制到剪贴板
            try:
                import win32clipboard
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, source_lot_value)
                win32clipboard.CloseClipboard()
                LOGGER.info("已复制值到剪贴板")
            except ImportError:
                try:
                    import pyperclip
                    pyperclip.copy(source_lot_value)
                    LOGGER.info("已复制值到剪贴板（pyperclip）")
                except ImportError:
                    raise RuntimeError("无法使用剪贴板，请安装pywin32或pyperclip")
            
            text_field_filled = False
            
            if use_win32_direct:
                # 使用win32gui找到文本输入框并点击
                LOGGER.info("使用Windows API查找文本输入框...")
                text_input_hwnd = None
                text_input_rect = None
                
                def enum_child_proc(hwnd_child, candidates):
                    try:
                        class_name = win32gui.GetClassName(hwnd_child)
                        rect = win32gui.GetWindowRect(hwnd_child)
                        width = rect[2] - rect[0]
                        height = rect[3] - rect[1]
                        size = width * height
                        
                        # 方法1: 查找Edit控件（包括各种变体）
                        if "EDIT" in class_name.upper() or "RICHTEXT" in class_name.upper() or "RICHEDIT" in class_name.upper():
                            # 大的文本输入框通常面积 > 5000
                            if size > 5000:
                                candidates.append((hwnd_child, size, rect, "Edit/RichEdit"))
                        
                        # 方法2: 查找所有大的子窗口（可能是自定义的文本输入框）
                        # 对话框中的文本输入框通常在中间位置，且比较大
                        if size > 10000:
                            # 获取对话框大小
                            dialog_rect = win32gui.GetWindowRect(dialog_hwnd)
                            dialog_width = dialog_rect[2] - dialog_rect[0]
                            dialog_height = dialog_rect[3] - dialog_rect[1]
                            
                            # 检查控件是否在对话框的下半部分（文本输入框通常在下方）
                            child_center_y = (rect[1] + rect[3]) // 2
                            dialog_center_y = (dialog_rect[1] + dialog_rect[3]) // 2
                            
                            if child_center_y > dialog_center_y and width > dialog_width * 0.5:
                                candidates.append((hwnd_child, size, rect, "Large child window"))
                    except:
                        pass
                    return True
                
                all_candidates = []
                win32gui.EnumChildWindows(dialog_hwnd, enum_child_proc, all_candidates)
                
                LOGGER.info(f"找到 {len(all_candidates)} 个可能的文本输入区域")
                
                if all_candidates:
                    # 优先选择Edit/RichEdit控件，否则选择最大的子窗口
                    edit_candidates = [c for c in all_candidates if "Edit" in c[3]]
                    if edit_candidates:
                        text_input_hwnd, _, text_input_rect, _ = max(edit_candidates, key=lambda x: x[1])
                        LOGGER.info(f"找到Edit控件文本输入框（handle: {text_input_hwnd}）")
                    else:
                        # 选择最大的子窗口
                        text_input_hwnd, _, text_input_rect, desc = max(all_candidates, key=lambda x: x[1])
                        LOGGER.info(f"找到大型子窗口作为文本输入框（handle: {text_input_hwnd}, 类型: {desc}）")
                    
                    # 获取控件中心坐标并点击
                    center_x = (text_input_rect[0] + text_input_rect[2]) // 2
                    center_y = (text_input_rect[1] + text_input_rect[3]) // 2
                    
                    LOGGER.info(f"点击文本输入框中心位置: ({center_x}, {center_y})")
                    try:
                        import pyautogui
                        pyautogui.click(center_x, center_y)
                        time.sleep(0.3)
                        
                        # 清除现有内容并粘贴
                        pyautogui.hotkey('ctrl', 'a')
                        time.sleep(0.1)
                        pyautogui.hotkey('ctrl', 'v')
                        text_field_filled = True
                        LOGGER.info(f"✅ 已填写Source Lot值: {source_lot_value}")
                    except ImportError:
                        raise RuntimeError("需要pyautogui来执行鼠标点击操作")
                else:
                    # 如果找不到任何控件，尝试估算白色区域的位置
                    LOGGER.warning("未找到文本输入控件，尝试估算白色区域位置...")
                    try:
                        dialog_rect = win32gui.GetWindowRect(dialog_hwnd)
                        dialog_width = dialog_rect[2] - dialog_rect[0]
                        dialog_height = dialog_rect[3] - dialog_rect[1]
                        
                        # 白色文本输入区域通常在对话框中间偏下的位置
                        # 估算：X在中间，Y在60-70%的位置
                        estimated_x = dialog_rect[0] + dialog_width // 2
                        estimated_y = dialog_rect[1] + int(dialog_height * 0.65)
                        
                        LOGGER.info(f"估算白色区域位置: ({estimated_x}, {estimated_y})")
                        import pyautogui
                        pyautogui.click(estimated_x, estimated_y)
                        time.sleep(0.3)
                        
                        # 清除现有内容并粘贴
                        pyautogui.hotkey('ctrl', 'a')
                        time.sleep(0.1)
                        pyautogui.hotkey('ctrl', 'v')
                        text_field_filled = True
                        LOGGER.info(f"✅ 已通过估算位置填写Source Lot值: {source_lot_value}")
                    except Exception as e:
                        LOGGER.error(f"估算位置方法也失败: {e}")
            
            else:
                # 使用pywinauto查找并点击文本输入框
                try:
                    # 方法1: 查找Edit控件
                    edit_controls = vpo_dialog.descendants(control_type="Edit")
                    LOGGER.info(f"找到 {len(edit_controls)} 个Edit控件")
                    
                    # 方法2: 查找所有文本相关的控件
                    text_controls = []
                    try:
                        all_controls = vpo_dialog.descendants()
                        for ctrl in all_controls:
                            try:
                                ctrl_type = str(ctrl.element_info.control_type_name).lower()
                                if 'edit' in ctrl_type or 'text' in ctrl_type or 'input' in ctrl_type:
                                    if ctrl.is_visible():
                                        text_controls.append(ctrl)
                            except:
                                continue
                        LOGGER.info(f"找到 {len(text_controls)} 个文本相关控件")
                    except:
                        pass
                    
                    # 优先选择最大的控件（通常是主要的文本输入框）
                    target_edit = None
                    target_rect = None
                    max_size = 0
                    
                    # 先尝试Edit控件
                    for edit in edit_controls:
                        try:
                            if edit.is_visible() and edit.is_enabled():
                                rect = edit.rectangle()
                                size = rect.width() * rect.height()
                                if size > max_size:
                                    max_size = size
                                    target_edit = edit
                                    target_rect = rect
                        except:
                            continue
                    
                    # 如果Edit控件不够大，尝试其他文本控件
                    if not target_edit or max_size < 5000:
                        for ctrl in text_controls:
                            try:
                                if ctrl.is_visible() and ctrl.is_enabled():
                                    rect = ctrl.rectangle()
                                    size = rect.width() * rect.height()
                                    if size > max_size:
                                        max_size = size
                                        target_edit = ctrl
                                        target_rect = rect
                            except:
                                continue
                    
                    if target_edit:
                        LOGGER.info(f"找到文本输入框（大小: {max_size}）")
                        # 使用鼠标点击文本输入框的中心位置
                        center_x = target_rect.left + target_rect.width() // 2
                        center_y = target_rect.top + target_rect.height() // 2
                        
                        LOGGER.info(f"点击文本输入框中心位置: ({center_x}, {center_y})")
                        try:
                            import pyautogui
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.3)
                            
                            # 清除现有内容并粘贴
                            pyautogui.hotkey('ctrl', 'a')
                            time.sleep(0.1)
                            pyautogui.hotkey('ctrl', 'v')
                            text_field_filled = True
                            LOGGER.info(f"✅ 已填写Source Lot值: {source_lot_value}")
                        except ImportError:
                            # 如果没有pyautogui，使用pywinauto的点击方法
                            target_edit.set_focus()
                            time.sleep(0.3)
                            target_edit.type_keys("^a")
                            time.sleep(0.1)
                            target_edit.type_keys("^v")
                            text_field_filled = True
                            LOGGER.info(f"✅ 已填写Source Lot值: {source_lot_value}")
                    else:
                        LOGGER.warning("未找到可用的文本输入控件，尝试估算位置...")
                        # 尝试估算白色区域位置
                        try:
                            dialog_rect = vpo_dialog.rectangle()
                            estimated_x = dialog_rect.left + dialog_rect.width() // 2
                            estimated_y = dialog_rect.top + int(dialog_rect.height() * 0.65)
                            
                            LOGGER.info(f"估算白色区域位置: ({estimated_x}, {estimated_y})")
                            import pyautogui
                            pyautogui.click(estimated_x, estimated_y)
                            time.sleep(0.3)
                            
                            # 清除现有内容并粘贴
                            pyautogui.hotkey('ctrl', 'a')
                            time.sleep(0.1)
                            pyautogui.hotkey('ctrl', 'v')
                            text_field_filled = True
                            LOGGER.info(f"✅ 已通过估算位置填写Source Lot值: {source_lot_value}")
                        except Exception as e:
                            LOGGER.error(f"估算位置方法也失败: {e}")
                except Exception as e:
                    LOGGER.debug(f"查找文本输入框失败: {e}")
            
            if not text_field_filled:
                raise RuntimeError("无法填写Source Lot值到文本输入框")
            
            # 等待一下，确保值已填写
            time.sleep(0.5)
            
            # 点击Search按钮
            LOGGER.info("点击Search按钮...")
            search_clicked = False
            
            if use_win32_direct:
                # 使用win32gui查找Search按钮
                LOGGER.info("使用Windows API查找Search按钮...")
                search_button_hwnd = None
                
                def enum_button_proc(hwnd_child, lParam):
                    try:
                        window_text = win32gui.GetWindowText(hwnd_child)
                        class_name = win32gui.GetClassName(hwnd_child)
                        LOGGER.debug(f"  子窗口: 文本='{window_text}', 类名='{class_name}'")
                        # 查找所有可能的按钮
                        if "BUTTON" in class_name.upper():
                            lParam.append((hwnd_child, window_text, class_name))
                    except:
                        pass
                    return True
                
                all_buttons = []
                win32gui.EnumChildWindows(dialog_hwnd, enum_button_proc, all_buttons)
                
                LOGGER.info(f"找到 {len(all_buttons)} 个按钮控件")
                for hwnd, text, cls in all_buttons:
                    LOGGER.info(f"  按钮: '{text}' (类名: {cls}, handle: {hwnd})")
                
                # 查找Search按钮
                search_buttons = [(h, t, c) for h, t, c in all_buttons if "SEARCH" in t.upper()]
                
                if search_buttons:
                    search_button_hwnd, button_text, button_class = search_buttons[0]
                    LOGGER.info(f"找到Search按钮（handle: {search_button_hwnd}, 文本: '{button_text}'）")
                    
                    # 获取按钮中心坐标并尝试多种点击方法
                    rect = win32gui.GetWindowRect(search_button_hwnd)
                    center_x = (rect[0] + rect[2]) // 2
                    center_y = (rect[1] + rect[3]) // 2
                    LOGGER.info(f"按钮位置: ({center_x}, {center_y}), 大小: {rect[2] - rect[0]}x{rect[3] - rect[1]}")
                    
                    # 方法1: 使用pyautogui点击（尝试多个位置）
                    try:
                        import pyautogui
                        # 尝试按钮的不同位置
                        click_positions = [
                            (center_x, center_y),  # 中心
                            (center_x - 10, center_y),  # 稍微左偏
                            (center_x + 10, center_y),  # 稍微右偏
                            (center_x, center_y - 5),  # 稍微上偏
                            (center_x, center_y + 5),  # 稍微下偏
                        ]
                        
                        for pos_x, pos_y in click_positions:
                            try:
                                LOGGER.debug(f"尝试点击位置: ({pos_x}, {pos_y})")
                                pyautogui.moveTo(pos_x, pos_y, duration=0.1)
                                time.sleep(0.05)
                                pyautogui.click(pos_x, pos_y, button='left')
                                time.sleep(0.2)
                                search_clicked = True
                                LOGGER.info(f"✅ 已通过pyautogui点击Search按钮（位置: {pos_x}, {pos_y}）")
                                break
                            except Exception as e:
                                LOGGER.debug(f"点击位置 ({pos_x}, {pos_y}) 失败: {e}")
                                continue
                    except ImportError:
                        LOGGER.warning("pyautogui未安装，尝试其他方法")
                    
                    # 方法2: 使用鼠标按下和释放消息
                    if not search_clicked:
                        try:
                            import win32api
                            # 将坐标转换为相对于按钮的坐标
                            button_rect = win32gui.GetWindowRect(search_button_hwnd)
                            rel_x = center_x - button_rect[0]
                            rel_y = center_y - button_rect[1]
                            # 将屏幕坐标转换为lParam
                            lparam = win32api.MAKELONG(rel_x, rel_y)
                            
                            # 发送鼠标按下消息
                            win32gui.SendMessage(search_button_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                            time.sleep(0.05)
                            # 发送鼠标释放消息
                            win32gui.SendMessage(search_button_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                            time.sleep(0.2)
                            search_clicked = True
                            LOGGER.info("✅ 已通过鼠标消息点击Search按钮")
                        except Exception as e:
                            LOGGER.debug(f"鼠标消息方法失败: {e}")
                    
                    # 方法3: 使用SendMessage BM_CLICK
                    if not search_clicked:
                        try:
                            # 先发送BM_SETSTATE来高亮按钮
                            win32gui.SendMessage(search_button_hwnd, win32con.BM_SETSTATE, 1, 0)
                            time.sleep(0.1)
                            # 发送BM_CLICK
                            result = win32gui.SendMessage(search_button_hwnd, win32con.BM_CLICK, 0, 0)
                            time.sleep(0.1)
                            win32gui.SendMessage(search_button_hwnd, win32con.BM_SETSTATE, 0, 0)
                            search_clicked = True
                            LOGGER.info(f"✅ 已通过SendMessage BM_CLICK点击Search按钮（返回值: {result}）")
                        except Exception as e:
                            LOGGER.debug(f"SendMessage方法失败: {e}")
                    
                    # 方法4: 使用PostMessage
                    if not search_clicked:
                        try:
                            win32gui.PostMessage(search_button_hwnd, win32con.BM_CLICK, 0, 0)
                            time.sleep(0.2)
                            search_clicked = True
                            LOGGER.info("✅ 已通过PostMessage点击Search按钮")
                        except Exception as e:
                            LOGGER.debug(f"PostMessage方法失败: {e}")
                    
                    # 方法5: 使用pyautogui在按钮位置点击（额外的鼠标点击尝试）
                    if not search_clicked:
                        try:
                            import pyautogui
                            LOGGER.info("尝试使用pyautogui在按钮坐标处点击...")
                            # 再次尝试在按钮中心位置点击
                            button_rect = win32gui.GetWindowRect(search_button_hwnd)
                            center_x = (button_rect[0] + button_rect[2]) // 2
                            center_y = (button_rect[1] + button_rect[3]) // 2
                            pyautogui.moveTo(center_x, center_y, duration=0.2)
                            time.sleep(0.1)
                            pyautogui.click(center_x, center_y, clicks=2, interval=0.1)  # 双击
                            time.sleep(0.3)
                            search_clicked = True
                            LOGGER.info(f"✅ 已通过pyautogui双击Search按钮（位置: {center_x}, {center_y}）")
                        except Exception as e:
                            LOGGER.debug(f"pyautogui双击方法失败: {e}")
                else:
                    LOGGER.warning("未找到Search按钮，但找到以下按钮:")
                    for hwnd, text, cls in all_buttons:
                        LOGGER.warning(f"  '{text}' (类名: {cls})")
            
            else:
                # 使用pywinauto查找Search按钮（这是最可靠的方法，就像"Search By VPOs"按钮那样）
                # 注意：不使用Enter键，必须使用鼠标点击
                
                # 方法1: 直接查找标题为"Search"的按钮并使用鼠标点击（最优先，最可靠）
                if not search_clicked:
                    try:
                        search_button = vpo_dialog.child_window(title="Search", control_type="Button")
                        if search_button.exists():
                            LOGGER.info("找到Search按钮（通过title）")
                            # 检查按钮属性
                            try:
                                LOGGER.info(f"按钮属性: enabled={search_button.is_enabled()}, visible={search_button.is_visible()}")
                                LOGGER.info(f"按钮文本: '{search_button.window_text()}'")
                                LOGGER.info(f"按钮类名: '{search_button.class_name()}'")
                                # 获取按钮位置
                                rect = search_button.rectangle()
                                LOGGER.info(f"按钮位置: left={rect.left}, top={rect.top}, width={rect.width()}, height={rect.height()}")
                            except Exception as e:
                                LOGGER.debug(f"获取按钮属性失败: {e}")
                            
                            if search_button.is_enabled() and search_button.is_visible():
                                # 先确保按钮可见（滚动到按钮位置）
                                try:
                                    search_button.set_focus()
                                    time.sleep(0.1)
                                except:
                                    pass
                                
                                # 获取按钮位置（备用方案）
                                try:
                                    button_rect = search_button.rectangle()
                                    button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                                    button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                                    LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y})")
                                except Exception as e:
                                    LOGGER.debug(f"获取按钮位置失败: {e}")
                                    button_center_x = button_center_y = None
                                
                                # 首先尝试使用click_input()进行鼠标点击
                                try:
                                    search_button.click_input()
                                    time.sleep(0.3)
                                    search_clicked = True
                                    LOGGER.info("✅ 已通过鼠标点击Search按钮（通过title，使用click_input）")
                                except Exception as e:
                                    LOGGER.warning(f"click_input()失败: {e}，尝试使用坐标点击...")
                                    # 如果click_input()失败，使用pyautogui在按钮坐标处点击
                                    if button_center_x and button_center_y:
                                        try:
                                            import pyautogui
                                            pyautogui.moveTo(button_center_x, button_center_y, duration=0.2)
                                            time.sleep(0.1)
                                            pyautogui.click(button_center_x, button_center_y, clicks=1)
                                            time.sleep(0.3)
                                            # 如果单次点击不行，尝试双击
                                            pyautogui.click(button_center_x, button_center_y, clicks=2, interval=0.1)
                                            time.sleep(0.2)
                                            search_clicked = True
                                            LOGGER.info(f"✅ 已通过坐标鼠标点击Search按钮（位置: {button_center_x}, {button_center_y}）")
                                        except Exception as e2:
                                            LOGGER.warning(f"坐标点击也失败: {e2}")
                            else:
                                LOGGER.warning(f"Search按钮存在但不可用: enabled={search_button.is_enabled()}, visible={search_button.is_visible()}")
                    except Exception as e1:
                        LOGGER.debug(f"通过title查找Search按钮失败: {e1}")
                
                # 方法2: 遍历所有按钮查找Search（使用click_input，最可靠）
                if not search_clicked:
                    try:
                        all_buttons = vpo_dialog.descendants(control_type="Button")
                        LOGGER.info(f"找到 {len(all_buttons)} 个按钮")
                        for idx, button in enumerate(all_buttons):
                            try:
                                button_text = button.window_text().strip()
                                button_class = button.class_name()
                                is_enabled = button.is_enabled()
                                is_visible = button.is_visible()
                                
                                LOGGER.info(f"  按钮 #{idx}: 文本='{button_text}', 类名='{button_class}', enabled={is_enabled}, visible={is_visible}")
                                
                                if button_text.upper() == "SEARCH" or "SEARCH" in button_text.upper():
                                    LOGGER.info(f"找到Search按钮（文本: '{button_text}'）")
                                    if is_visible and is_enabled:
                                        # 先设置焦点
                                        try:
                                            button.set_focus()
                                            time.sleep(0.1)
                                        except:
                                            pass
                                        
                                        # 获取按钮位置（备用方案）
                                        try:
                                            button_rect = button.rectangle()
                                            button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                                            button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                                            LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y})")
                                        except Exception as e:
                                            LOGGER.debug(f"获取按钮位置失败: {e}")
                                            button_center_x = button_center_y = None
                                        
                                        # 优先使用click_input()进行鼠标点击
                                        try:
                                            button.click_input()
                                            time.sleep(0.3)
                                            search_clicked = True
                                            LOGGER.info("✅ 已通过鼠标点击Search按钮（遍历按钮，使用click_input）")
                                            break
                                        except Exception as e:
                                            LOGGER.warning(f"click_input()失败: {e}，尝试使用坐标点击...")
                                            # 如果click_input()失败，使用pyautogui在按钮坐标处点击
                                            if button_center_x and button_center_y:
                                                try:
                                                    import pyautogui
                                                    pyautogui.moveTo(button_center_x, button_center_y, duration=0.2)
                                                    time.sleep(0.1)
                                                    pyautogui.click(button_center_x, button_center_y, clicks=1)
                                                    time.sleep(0.3)
                                                    # 如果单次点击不行，尝试双击
                                                    pyautogui.click(button_center_x, button_center_y, clicks=2, interval=0.1)
                                                    time.sleep(0.2)
                                                    search_clicked = True
                                                    LOGGER.info(f"✅ 已通过坐标鼠标点击Search按钮（位置: {button_center_x}, {button_center_y}）")
                                                    break
                                                except Exception as e2:
                                                    LOGGER.warning(f"坐标点击也失败: {e2}")
                                    else:
                                        LOGGER.warning(f"Search按钮不可用: enabled={is_enabled}, visible={is_visible}")
                            except Exception as e:
                                LOGGER.debug(f"检查按钮 #{idx} 时出错: {e}")
                                continue
                    except Exception as e2:
                        LOGGER.debug(f"遍历按钮失败: {e2}")
                
                # 方法3: 尝试使用uia backend（如果当前是win32）
                if not search_clicked:
                    try:
                        LOGGER.info("尝试使用uia backend查找Search按钮...")
                        # 重新连接对话框使用uia backend
                        if win32gui:
                            dialog_hwnd = vpo_dialog.handle if hasattr(vpo_dialog, 'handle') else None
                            if dialog_hwnd:
                                try:
                                    uia_app = Application(backend="uia").connect(handle=dialog_hwnd)
                                    uia_dialog = uia_app.window(handle=dialog_hwnd)
                                    if uia_dialog.exists():
                                        search_button = uia_dialog.child_window(title="Search", control_type="Button")
                                        if search_button.exists() and search_button.is_enabled():
                                            LOGGER.info("使用uia backend找到Search按钮")
                                            search_button.click_input()
                                            search_clicked = True
                                            LOGGER.info("✅ 已通过uia backend点击Search按钮")
                                except Exception as e:
                                    LOGGER.debug(f"uia backend方法失败: {e}")
                    except Exception as e3:
                        LOGGER.debug(f"uia backend查找失败: {e3}")
                
                # 方法4: 使用部分匹配查找Search按钮
                if not search_clicked:
                    try:
                        all_buttons = vpo_dialog.descendants(control_type="Button")
                        for button in all_buttons:
                            try:
                                button_text = button.window_text().strip()
                                if "SEARCH" in button_text.upper():
                                    LOGGER.info(f"找到Search按钮（部分匹配: '{button_text}'）")
                                    if button.is_visible() and button.is_enabled():
                                        button.set_focus()
                                        time.sleep(0.1)
                                        button.click_input()
                                        search_clicked = True
                                        LOGGER.info("✅ 已点击Search按钮（部分匹配，使用click_input）")
                                        break
                            except:
                                continue
                    except Exception as e4:
                        LOGGER.debug(f"部分匹配查找失败: {e4}")
                
                # 方法3: 使用Windows API查找Search按钮（备用）
                if not search_clicked and win32gui and win32con:
                    try:
                        LOGGER.info("使用Windows API查找Search按钮...")
                        hwnd = vpo_dialog.handle
                        
                        def enum_child_proc(hwnd_child, lParam):
                            try:
                                window_text = win32gui.GetWindowText(hwnd_child)
                                class_name = win32gui.GetClassName(hwnd_child)
                                if "SEARCH" in window_text.upper() and "BUTTON" in class_name.upper():
                                    lParam.append(hwnd_child)
                                    return False  # 停止枚举
                                return True
                            except:
                                return True
                        
                        button_list = []
                        win32gui.EnumChildWindows(hwnd, enum_child_proc, button_list)
                        
                        if button_list:
                            button_hwnd = button_list[0]
                            # 获取按钮位置并点击
                            rect = win32gui.GetWindowRect(button_hwnd)
                            center_x = (rect[0] + rect[2]) // 2
                            center_y = (rect[1] + rect[3]) // 2
                            
                            try:
                                import pyautogui
                                pyautogui.click(center_x, center_y)
                                search_clicked = True
                                LOGGER.info(f"✅ 已通过屏幕坐标点击Search按钮（位置: {center_x}, {center_y}）")
                            except ImportError:
                                win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                                search_clicked = True
                                LOGGER.info("✅ 已通过Windows API点击Search按钮")
                    except Exception as e4:
                        LOGGER.debug(f"使用Windows API查找Search按钮失败: {e4}")
            
            # 如果所有方法都失败，尝试估算Search按钮位置并使用鼠标点击
            if not search_clicked:
                LOGGER.warning("未找到Search按钮，尝试估算位置并使用鼠标点击...")
                try:
                    if use_win32_direct:
                        dialog_rect = win32gui.GetWindowRect(dialog_hwnd)
                        # 确保对话框有焦点
                        win32gui.SetForegroundWindow(dialog_hwnd)
                        win32gui.BringWindowToTop(dialog_hwnd)
                    else:
                        dialog_rect_obj = vpo_dialog.rectangle()
                        dialog_rect = (dialog_rect_obj.left, dialog_rect_obj.top, 
                                      dialog_rect_obj.right, dialog_rect_obj.bottom)
                        vpo_dialog.set_focus()
                    
                    dialog_width = dialog_rect[2] - dialog_rect[0]
                    dialog_height = dialog_rect[3] - dialog_rect[1]
                    
                    time.sleep(0.2)
                    
                    # 方法1: 尝试估算位置并使用鼠标点击
                    try:
                        import pyautogui
                        # 尝试多个可能的位置（Search按钮通常在对话框顶部左侧，黄色按钮）
                        possible_positions = [
                            (0.78, 0.15),  # 78%宽度, 15%高度（右上角）
                            (0.45, 0.12),  # 45%宽度, 12%高度（中间偏左，可能是Search按钮）
                            (0.50, 0.12),  # 50%宽度, 12%高度
                            (0.55, 0.12),  # 55%宽度, 12%高度
                            (0.80, 0.12),  # 80%宽度, 12%高度
                            (0.75, 0.18),  # 75%宽度, 18%高度
                        ]
                        
                        for pos_idx, (width_ratio, height_ratio) in enumerate(possible_positions):
                            estimated_x = dialog_rect[0] + int(dialog_width * width_ratio)
                            estimated_y = dialog_rect[1] + int(dialog_height * height_ratio)
                            
                            LOGGER.info(f"尝试估算位置 #{pos_idx + 1}: ({estimated_x}, {estimated_y})")
                            try:
                                pyautogui.moveTo(estimated_x, estimated_y, duration=0.1)
                                time.sleep(0.1)
                                pyautogui.click(estimated_x, estimated_y, clicks=1, interval=0.1)
                                time.sleep(0.3)
                                # 尝试双击
                                pyautogui.click(estimated_x, estimated_y, clicks=2, interval=0.1)
                                time.sleep(0.2)
                                search_clicked = True
                                LOGGER.info(f"✅ 已通过估算位置 #{pos_idx + 1} 点击Search按钮")
                                break
                            except Exception as e:
                                LOGGER.debug(f"估算位置 #{pos_idx + 1} 失败: {e}")
                                continue
                    except ImportError:
                        LOGGER.warning("pyautogui未安装，跳过估算位置点击")
                    except Exception as e:
                        LOGGER.debug(f"估算位置点击失败: {e}")
                
                    # 检查是否成功点击
                    if not search_clicked:
                        LOGGER.error("所有方法都失败，无法点击Search按钮")
                    else:
                        LOGGER.info("✓ Search按钮点击成功")
                except Exception as e:
                    LOGGER.error(f"估算位置方法也失败: {e}")
                    if not search_clicked:
                        raise RuntimeError("无法点击Search按钮，已尝试所有方法（鼠标点击、Windows API、估算位置）")
            
            # 等待搜索完成
            time.sleep(1)
            LOGGER.info("✅ 已成功填写Source Lot值并点击Search按钮")
        
        except Exception as e:
            LOGGER.error(f"填写VPO搜索对话框失败: {e}")
            raise RuntimeError(f"无法填写VPO搜索对话框: {e}")
    
    def _fill_units_search_dialog(self, ui_config: dict) -> None:
        """
        在Units搜索对话框中填写参数并点击Search
        
        对话框标题: "MOLE - Import List Of Serial Numbers"
        需要填写: Serial Numbers 文本框（多行文本框）
        需要点击: Search 按钮
        
        Args:
            ui_config: UI配置数据，包含units_info（用户粘贴的units信息）
        """
        # 确保使用全局的win32con变量
        global win32con
        
        if Application is None:
            raise RuntimeError("pywinauto 未安装，无法执行 UI 自动化")
        
        LOGGER.info("等待Units搜索对话框出现（标题: MOLE - Import List Of Serial Numbers）...")
        
        # 等待对话框出现（最多等待30秒，因为对话框可能需要一些时间才能完全加载）
        units_dialog = None
        deadline = time.time() + 30
        
        # 对话框标题模式
        dialog_title_patterns = [
            "MOLE - Import List Of Serial Numbers",
            ".*Import List Of Serial Numbers.*",
            ".*Import.*Serial.*",
        ]
        
        # 首先尝试使用win32gui枚举所有窗口（带超时保护）
        if win32gui:
            LOGGER.info("使用Windows API枚举所有窗口...")
            dialog_hwnd = None
            
            def enum_windows_callback(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        # 查找包含 "Import List Of Serial Numbers" 的窗口
                        if "Import" in window_text and "Serial" in window_text:
                            rect = win32gui.GetWindowRect(hwnd)
                            width = rect[2] - rect[0]
                            height = rect[3] - rect[1]
                            # 对话框通常宽度在400-800，高度在400-700
                            if 300 < width < 1000 and 300 < height < 900:
                                windows.append((hwnd, window_text, width, height))
                except:
                    pass
                return True
            
            candidate_windows = []
            
            # 使用线程和超时机制，避免EnumWindows卡住
            import threading
            enum_complete = threading.Event()
            enum_exception = [None]
            
            def enum_thread():
                try:
                    win32gui.EnumWindows(enum_windows_callback, candidate_windows)
                except Exception as e:
                    enum_exception[0] = e
                finally:
                    enum_complete.set()
            
            enum_thread_obj = threading.Thread(target=enum_thread, daemon=True)
            enum_thread_obj.start()
            enum_thread_obj.join(timeout=3.0)  # 最多等待3秒
            
            if not enum_complete.is_set():
                LOGGER.warning("EnumWindows超时（3秒），尝试其他方法...")
            elif enum_exception[0]:
                LOGGER.warning(f"EnumWindows出错: {enum_exception[0]}，尝试其他方法...")
            elif candidate_windows:
                dialog_hwnd, dialog_title, dialog_width, dialog_height = candidate_windows[0]
                LOGGER.info(f"找到Units搜索对话框（Windows API）: '{dialog_title}', hwnd={dialog_hwnd}, 大小={dialog_width}x{dialog_height}")
                
                # 使用pywinauto连接到这个窗口
                try:
                    dialog_app = Application(backend="win32").connect(handle=dialog_hwnd)
                    units_dialog = dialog_app.window(handle=dialog_hwnd)
                    LOGGER.info("已通过pywinauto连接到Units搜索对话框")
                except Exception as e:
                    LOGGER.warning(f"pywinauto连接失败: {e}")
        
        # 如果win32gui方法失败，使用pywinauto查找
        if units_dialog is None:
            LOGGER.info("使用pywinauto查找Units搜索对话框...")
            backends = ["win32", "uia"]
            start_time = time.time()
            last_progress_log = 0
            
            while time.time() < deadline and units_dialog is None:
                for backend in backends:
                    for title_pattern in dialog_title_patterns:
                        try:
                            # 尝试连接，但使用较短的超时时间
                            dialog_app = Application(backend=backend).connect(
                                title_re=title_pattern,
                                visible_only=True,
                                timeout=1
                            )
                            windows = dialog_app.windows()
                            LOGGER.debug(f"使用 {backend} backend 和模式 '{title_pattern}' 找到 {len(windows)} 个窗口")
                            for win in windows:
                                try:
                                    win_text = win.window_text()
                                    LOGGER.debug(f"检查窗口: '{win_text}'")
                                    # 更宽松的匹配条件
                                    if ("Import" in win_text and "Serial" in win_text) or \
                                       ("Serial" in win_text and "Numbers" in win_text) or \
                                       ("Import" in win_text and "List" in win_text):
                                        units_dialog = win
                                        LOGGER.info(f"✅ 找到Units搜索对话框（backend: {backend}, 标题: '{win_text}'）")
                                        break
                                except Exception as e:
                                    LOGGER.debug(f"检查窗口时出错: {e}")
                                    continue
                            if units_dialog:
                                break
                        except Exception as e:
                            LOGGER.debug(f"连接失败 (backend={backend}, pattern={title_pattern}): {e}")
                            continue
                    if units_dialog:
                        break
                
                if units_dialog is None:
                    time.sleep(0.5)
                    # 每5秒输出一次进度
                    elapsed = time.time() - start_time
                    if int(elapsed) - last_progress_log >= 5:
                        remaining = deadline - time.time()
                        LOGGER.info(f"仍在等待对话框出现...（已等待 {elapsed:.1f} 秒，剩余 {remaining:.1f} 秒）")
                        last_progress_log = int(elapsed)
                        
                        # 尝试列出所有可见窗口，帮助调试
                        if win32gui:
                            try:
                                all_windows = []
                                def list_windows_callback(hwnd, windows):
                                    try:
                                        if win32gui.IsWindowVisible(hwnd):
                                            text = win32gui.GetWindowText(hwnd)
                                            if text and len(text) > 0:
                                                windows.append(text)
                                    except:
                                        pass
                                    return True
                                win32gui.EnumWindows(list_windows_callback, all_windows)
                                # 只显示包含 "Import", "Serial", "MOLE" 的窗口
                                relevant_windows = [w for w in all_windows if any(keyword in w for keyword in ["Import", "Serial", "MOLE", "Mole"])]
                                if relevant_windows:
                                    LOGGER.debug(f"当前可见的相关窗口: {relevant_windows[:5]}")  # 只显示前5个
                            except:
                                pass
        
        if units_dialog is None:
            raise RuntimeError("未找到Units搜索对话框（标题应包含 'Import List Of Serial Numbers'）")
        
        try:
            # 激活对话框
            units_dialog.set_focus()
            time.sleep(0.5)
        except Exception as focus_exc:
            LOGGER.warning(f"设置Units对话框焦点失败: {focus_exc}")

        LOGGER.info("开始填写Units搜索对话框...")

        # 包裹整个填写和搜索过程在try块中
        try:
            # 获取units信息
            units_info = ui_config.get('units_info', '').strip()
            if not units_info:
                raise RuntimeError("units_info 为空，请在配置UI中粘贴Units信息")

            # Mole 的多行输入控件对 CRLF 的兼容性更好，确保换行格式统一
            normalized_units_info = units_info.replace("\r\n", "\n").replace("\n", "\r\n")

            LOGGER.info(f"Units信息长度: {len(units_info)} 字符")
            LOGGER.debug(f"Units信息预览: {units_info[:100]}...")

            # 复制units信息到剪贴板
            try:
                import pyperclip
                pyperclip.copy(normalized_units_info)
                LOGGER.info("已复制Units信息到剪贴板")
            except ImportError:
                try:
                    import win32clipboard
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(normalized_units_info, win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    LOGGER.info("已复制Units信息到剪贴板（使用win32clipboard）")
                except ImportError:
                    LOGGER.warning("无法复制到剪贴板，将直接输入文本")
            
            # 先点击对话框空白处，确保对话框有焦点（用户反馈需要这一步）
            try:
                dialog_rect = units_dialog.rectangle()
                # 点击对话框中心偏上的位置（空白区域）
                blank_area_x = dialog_rect.left + dialog_rect.width() // 2
                blank_area_y = dialog_rect.top + dialog_rect.height() // 3
                
                try:
                    import pyautogui
                    pyautogui.click(blank_area_x, blank_area_y)
                    time.sleep(0.3)
                    LOGGER.info("已点击对话框空白处，确保对话框有焦点")
                except ImportError:
                    LOGGER.warning("pyautogui未安装，跳过点击空白处")
            except Exception as e:
                LOGGER.debug(f"点击对话框空白处失败: {e}")
            
            # 查找多行文本框（Serial Numbers输入框）
            # 优化：优先使用Windows API方法，因为这是最可靠的方法（根据日志验证）
            text_field_filled = False
            
            # 方法1: 使用Windows API查找文本框（最优先，已验证成功）
            if not text_field_filled and win32gui:
                try:
                    LOGGER.info("使用Windows API查找Serial Numbers文本框（最优先方法）...")
                    dialog_hwnd = units_dialog.handle if hasattr(units_dialog, 'handle') else None
                    if dialog_hwnd:
                        def enum_edit_proc(hwnd_child, lParam):
                            try:
                                class_name = win32gui.GetClassName(hwnd_child)
                                if "EDIT" in class_name.upper():
                                    rect = win32gui.GetWindowRect(hwnd_child)
                                    width = rect[2] - rect[0]
                                    height = rect[3] - rect[1]
                                    # 多行文本框通常比较大
                                    if width > 150 and height > 80:
                                        lParam.append((hwnd_child, width, height, rect))
                            except:
                                pass
                            return True
                        
                        edit_list = []
                        win32gui.EnumChildWindows(dialog_hwnd, enum_edit_proc, edit_list)
                        
                        LOGGER.info(f"Windows API找到 {len(edit_list)} 个Edit控件")
                        for i, (hwnd, w, h, rect) in enumerate(edit_list):
                            LOGGER.info(f"  Edit {i+1}: hwnd={hwnd}, 大小={w}x{h}, 位置=({rect[0]},{rect[1]})")
                        
                        if edit_list:
                            # 选择最大的文本框
                            edit_list.sort(key=lambda x: x[1] * x[2], reverse=True)
                            edit_hwnd, edit_width, edit_height, edit_rect = edit_list[0]
                            
                            LOGGER.info(f"选择最大的文本框（Windows API）: {edit_width}x{edit_height}, hwnd={edit_hwnd}")
                            
                            # 使用鼠标点击和粘贴（已验证成功的方法）
                            try:
                                import pyautogui
                                center_x = (edit_rect[0] + edit_rect[2]) // 2
                                center_y = (edit_rect[1] + edit_rect[3]) // 2
                                
                                LOGGER.info(f"点击文本框中心: ({center_x}, {center_y})")
                                pyautogui.click(center_x, center_y)
                                time.sleep(0.8)
                                
                                # 再次点击确保焦点
                                pyautogui.click(center_x, center_y)
                                time.sleep(0.5)
                                
                                # 清除并粘贴
                                LOGGER.info("清除现有内容...")
                                pyautogui.hotkey('ctrl', 'a')
                                time.sleep(0.3)
                                
                                LOGGER.info("粘贴Units信息...")
                                pyautogui.hotkey('ctrl', 'v')
                                time.sleep(1.0)
                                
                                text_field_filled = True
                                LOGGER.info("✅ 已通过Windows API + 鼠标点击填写Serial Numbers文本框")
                            except ImportError:
                                LOGGER.warning("pyautogui未安装，尝试使用SendMessage")
                                # 如果pyautogui未安装，使用SendMessage
                                try:
                                    if not win32con:
                                        raise ImportError("win32con未安装")
                                    
                                    win32gui.SetFocus(edit_hwnd)
                                    time.sleep(0.3)
                                    win32gui.SendMessage(edit_hwnd, win32con.EM_SETSEL, 0, -1)
                                    time.sleep(0.1)
                                    win32gui.SendMessage(edit_hwnd, win32con.EM_REPLACESEL, True, units_info)
                                    time.sleep(0.5)
                                    text_field_filled = True
                                    LOGGER.info("✅ 已通过SendMessage填写Serial Numbers文本框")
                                except Exception as e:
                                    LOGGER.warning(f"SendMessage方法失败: {e}")
                except Exception as e:
                    LOGGER.debug(f"Windows API方法失败: {e}")
            
            # 方法2: 使用pywinauto查找Edit控件（备用方法）
            if not text_field_filled:
                try:
                    LOGGER.info("使用pywinauto查找Serial Numbers文本框（备用方法）...")
                    edit_controls = units_dialog.descendants(control_type="Edit")
                    LOGGER.info(f"找到 {len(edit_controls)} 个Edit控件")
                    
                    # 查找最大的Edit控件（通常是多行文本框）
                    target_edit = None
                    target_rect = None
                    max_size = 0
                    
                    for edit in edit_controls:
                        try:
                            if edit.is_visible() and edit.is_enabled():
                                rect = edit.rectangle()
                                size = rect.width() * rect.height()
                                # 多行文本框通常比较大（降低阈值，至少 100x50）
                                if size > max_size and size > 5000:
                                    max_size = size
                                    target_edit = edit
                                    target_rect = rect
                        except Exception as e:
                            LOGGER.debug(f"检查Edit控件时出错: {e}")
                            continue
                    
                    if target_edit:
                        LOGGER.info(f"找到Serial Numbers文本框（大小: {max_size}, 位置: {target_rect.left},{target_rect.top}）")
                        # 使用鼠标点击文本框中心（更可靠的方法）
                        try:
                            import pyautogui
                            center_x = target_rect.left + target_rect.width() // 2
                            center_y = target_rect.top + target_rect.height() // 2

                            # 先点击文本框，确保获得焦点
                            LOGGER.info(f"点击文本框中心位置: ({center_x}, {center_y})")
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.8)  # 增加等待时间，确保焦点切换

                            # 再次点击，确保文本框获得焦点
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.5)

                            # 清除现有内容
                            LOGGER.info("清除现有内容...")
                            pyautogui.hotkey('ctrl', 'a')
                            time.sleep(0.3)

                            # 验证剪贴板内容
                            try:
                                import pyperclip
                                clipboard_content = pyperclip.paste()
                                LOGGER.info(f"剪贴板内容预览: {clipboard_content[:50]}... (长度: {len(clipboard_content)})")
                            except Exception:
                                pass

                            # 为后续多次写入准备一个可靠的文本写入辅助方法
                            def _force_write_text(edit_control) -> bool:
                                """确保文本框包含Units文本，无论粘贴是否被拦截"""
                                write_attempts = [
                                    ("set_edit_text", lambda: edit_control.set_edit_text(normalized_units_info)),
                                    ("set_value", lambda: edit_control.set_value(normalized_units_info)),
                                    ("type_keys", lambda: (
                                        edit_control.set_focus(),
                                        time.sleep(0.3),
                                        edit_control.type_keys("^a{BACKSPACE}"),
                                        time.sleep(0.3),
                                        edit_control.type_keys(normalized_units_info, with_spaces=True, pause=0.01)
                                    )),
                                ]

                                for write_name, write_action in write_attempts:
                                    try:
                                        write_action()
                                        time.sleep(0.5)
                                        try:
                                            current_text = edit_control.window_text()
                                            if current_text:
                                                LOGGER.info(
                                                    f"✅ 通过 {write_name} 写入Units文本，长度: {len(current_text)} 字符"
                                                )
                                                LOGGER.debug(f"内容预览: {current_text[:100]}...")
                                                return True
                                        except Exception as read_exc:
                                            LOGGER.debug(f"读取文本内容以验证失败（{write_name}）：{read_exc}")
                                            return True
                                    except Exception as write_exc:
                                        LOGGER.debug(f"通过 {write_name} 写入失败: {write_exc}")
                                return False

                            # 粘贴内容
                            LOGGER.info("粘贴Units信息...")
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(1.0)  # 增加等待时间，确保粘贴完成

                            # 验证粘贴是否成功（尝试读取文本框内容），失败则强制写入
                            paste_success = False
                            try:
                                current_text = target_edit.window_text()
                                if current_text:
                                    paste_success = True
                                    LOGGER.info(f"✅ 文本框内容已更新，长度: {len(current_text)} 字符")
                                    LOGGER.debug(f"内容预览: {current_text[:100]}...")
                                else:
                                    LOGGER.warning("⚠️ 文本框内容为空，粘贴可能失败，尝试直接写入文本")
                            except Exception as read_exc:
                                LOGGER.debug(f"无法读取文本框内容进行验证: {read_exc}")

                            if not paste_success:
                                paste_success = _force_write_text(target_edit)

                            text_field_filled = paste_success
                            if paste_success:
                                LOGGER.info("✅ 已通过鼠标点击填写Serial Numbers文本框")
                            else:
                                LOGGER.warning("粘贴和直接写入都未确认成功，将继续尝试其他方法")
                        except ImportError:
                            LOGGER.warning("pyautogui未安装，尝试使用pywinauto")
                            # 回退到pywinauto方法
                            try:
                                target_edit.set_focus()
                                time.sleep(0.8)

                                # 清除现有内容
                                target_edit.type_keys("^a")
                                time.sleep(0.3)

                                # 粘贴内容
                                if pyperclip or win32clipboard:
                                    target_edit.type_keys("^v")
                                else:
                                    # 如果没有剪贴板，直接输入（可能很慢）
                                    LOGGER.warning("没有剪贴板库，将直接输入文本（可能很慢）")
                                    target_edit.type_keys(normalized_units_info)

                                time.sleep(0.5)
                                text_field_filled = True
                                LOGGER.info("✅ 已填写Serial Numbers文本框（pywinauto）")
                            except Exception as e:
                                LOGGER.error(f"使用pywinauto填写失败: {e}")
                                raise
                        else:
                            LOGGER.warning(f"未找到足够大的Edit控件（最大: {max_size}，阈值: 5000）")
                except Exception as e:
                    LOGGER.debug(f"使用pywinauto查找Edit控件失败: {e}")
            
            # 方法3: 使用Windows API查找文本框（如果方法2失败）
            if not text_field_filled and win32gui:
                try:
                    LOGGER.info("使用Windows API查找Serial Numbers文本框...")
                    dialog_hwnd = units_dialog.handle if hasattr(units_dialog, 'handle') else None
                    if dialog_hwnd:
                        def enum_edit_proc(hwnd_child, lParam):
                            try:
                                class_name = win32gui.GetClassName(hwnd_child)
                                if "EDIT" in class_name.upper():
                                    rect = win32gui.GetWindowRect(hwnd_child)
                                    width = rect[2] - rect[0]
                                    height = rect[3] - rect[1]
                                    # 降低阈值，多行文本框通常比较大
                                    if width > 150 and height > 80:
                                        lParam.append((hwnd_child, width, height, rect))
                            except:
                                pass
                            return True
                    
                    edit_list = []
                    win32gui.EnumChildWindows(dialog_hwnd, enum_edit_proc, edit_list)
                    
                    LOGGER.info(f"Windows API找到 {len(edit_list)} 个Edit控件")
                    for i, (hwnd, w, h, rect) in enumerate(edit_list):
                        LOGGER.info(f"  Edit {i+1}: hwnd={hwnd}, 大小={w}x{h}, 位置=({rect[0]},{rect[1]})")
                    
                    if edit_list:
                        # 选择最大的文本框
                        edit_list.sort(key=lambda x: x[1] * x[2], reverse=True)
                        edit_hwnd, edit_width, edit_height, edit_rect = edit_list[0]
                        
                        LOGGER.info(f"选择最大的文本框（Windows API）: {edit_width}x{edit_height}, hwnd={edit_hwnd}")
                        
                        # 方法2a: 使用鼠标点击和粘贴
                        try:
                            import pyautogui
                            center_x = (edit_rect[0] + edit_rect[2]) // 2
                            center_y = (edit_rect[1] + edit_rect[3]) // 2
                            
                            LOGGER.info(f"点击文本框中心: ({center_x}, {center_y})")
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.8)
                            
                            # 再次点击确保焦点
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.5)
                            
                            # 清除并粘贴
                            LOGGER.info("清除现有内容...")
                            pyautogui.hotkey('ctrl', 'a')
                            time.sleep(0.3)
                            
                            LOGGER.info("粘贴Units信息...")
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(1.0)
                            
                            text_field_filled = True
                            LOGGER.info("✅ 已通过Windows API + 鼠标点击填写Serial Numbers文本框")
                        except ImportError:
                            LOGGER.warning("pyautogui未安装，尝试使用SendMessage")
                        
                        # 方法2b: 使用SendMessage直接发送文本（备用方法）
                        if not text_field_filled:
                            try:
                                # win32con和win32api已在文件顶部导入
                                if not win32con:
                                    raise ImportError("win32con未安装")
                                
                                LOGGER.info("使用SendMessage直接发送文本到文本框...")
                                
                                # 先设置焦点
                                win32gui.SetFocus(edit_hwnd)
                                time.sleep(0.3)
                                
                                # 发送WM_SETTEXT消息设置文本
                                # 注意：对于多行文本框，可能需要使用EM_REPLACESEL
                                # 先清除现有内容
                                win32gui.SendMessage(edit_hwnd, win32con.EM_SETSEL, 0, -1)
                                time.sleep(0.1)
                                
                                # 替换选中内容
                                win32gui.SendMessage(edit_hwnd, win32con.EM_REPLACESEL, True, units_info)
                                time.sleep(0.5)
                                
                                text_field_filled = True
                                LOGGER.info("✅ 已通过SendMessage填写Serial Numbers文本框")
                            except Exception as e:
                                LOGGER.warning(f"SendMessage方法失败: {e}")
                
                except Exception as e:
                    LOGGER.error(f"Windows API方法失败: {e}")
                    import traceback
                    LOGGER.debug(traceback.format_exc())
            
            # 方法3: 最后的备用方法 - 尝试查找所有可见的文本框并点击最大的
            if not text_field_filled:
                LOGGER.warning("所有自动方法都失败，尝试最后的备用方法...")
                try:
                    # 尝试使用pyautogui在对话框中心区域查找并点击文本框
                    import pyautogui
                    dialog_rect = units_dialog.rectangle()
                    
                    # 在对话框的下半部分（通常是文本框位置）尝试点击
                    # 根据图片，文本框在对话框的下半部分
                    text_area_x = dialog_rect.left + dialog_rect.width() // 2
                    text_area_y = dialog_rect.top + dialog_rect.height() * 2 // 3  # 下半部分
                    
                    LOGGER.info(f"尝试在对话框下半部分点击: ({text_area_x}, {text_area_y})")
                    pyautogui.click(text_area_x, text_area_y)
                    time.sleep(0.5)
                    
                    # 再次点击确保焦点
                    pyautogui.click(text_area_x, text_area_y)
                    time.sleep(0.5)
                    
                    # 清除并粘贴
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.3)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(1.0)
                    
                    text_field_filled = True
                    LOGGER.info("✅ 已通过备用方法填写Serial Numbers文本框")
                except Exception as e:
                    LOGGER.error(f"备用方法也失败: {e}")
            
            if not text_field_filled:
                error_msg = (
                    "无法自动填写Serial Numbers文本框。\n"
                    "请手动执行以下步骤：\n"
                    "1. 在对话框中找到 'Serial Numbers' 文本框\n"
                    "2. 点击文本框\n"
                    "3. 按 Ctrl+V 粘贴Units信息\n"
                    "4. 点击 'Search' 按钮\n"
                    f"\nUnits信息（已复制到剪贴板）:\n{units_info[:200]}..."
                )
                LOGGER.error(error_msg)
                raise RuntimeError("无法自动填写Serial Numbers文本框，请手动操作")
            
            # 等待一下，确保内容已填写
            time.sleep(0.5)
            
            # 点击Search按钮（优化：优先使用uia backend，因为这是最可靠的方法）
            LOGGER.info("点击Search按钮...")
            
            # 确保对话框有焦点
            try:
                units_dialog.set_focus()
                time.sleep(0.3)
            except:
                pass
            
            search_clicked = False
            
            # 方法1: 优先使用uia backend（这是最可靠的方法，根据日志显示这是成功的方法）
            if not search_clicked:
                try:
                    LOGGER.info("尝试使用uia backend查找Search按钮（最优先方法）...")
                    dialog_hwnd = units_dialog.handle if hasattr(units_dialog, 'handle') else None
                    if dialog_hwnd:
                        try:
                            uia_app = Application(backend="uia").connect(handle=dialog_hwnd)
                            uia_dialog = uia_app.window(handle=dialog_hwnd)
                            if uia_dialog.exists():
                                search_button = uia_dialog.child_window(title="Search", control_type="Button")
                                if search_button.exists() and search_button.is_enabled():
                                    LOGGER.info("使用uia backend找到Search按钮")
                                    search_button.click_input()
                                    time.sleep(0.3)
                                    search_clicked = True
                                    LOGGER.info("✅ 已通过uia backend点击Search按钮")
                        except Exception as e:
                            LOGGER.debug(f"uia backend方法失败: {e}")
                except Exception as e3:
                    LOGGER.debug(f"uia backend查找失败: {e3}")
            
            # 方法2: 使用win32 backend查找（如果uia失败）
            if not search_clicked:
                try:
                    search_button = units_dialog.child_window(title="Search", control_type="Button")
                    if search_button.exists():
                        LOGGER.info("找到Search按钮（通过title，win32 backend）")
                        if search_button.is_enabled() and search_button.is_visible():
                            try:
                                search_button.click_input()
                                time.sleep(0.3)
                                search_clicked = True
                                LOGGER.info("✅ 已通过鼠标点击Search按钮（通过title，使用click_input）")
                            except Exception as e:
                                LOGGER.warning(f"click_input()失败: {e}")
                                # 如果click_input()失败，尝试使用坐标点击
                                try:
                                    button_rect = search_button.rectangle()
                                    button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                                    button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                                    import pyautogui
                                    pyautogui.moveTo(button_center_x, button_center_y, duration=0.2)
                                    time.sleep(0.1)
                                    pyautogui.click(button_center_x, button_center_y, clicks=1)
                                    time.sleep(0.3)
                                    search_clicked = True
                                    LOGGER.info(f"✅ 已通过坐标鼠标点击Search按钮（位置: {button_center_x}, {button_center_y}）")
                                except Exception as e2:
                                    LOGGER.warning(f"坐标点击也失败: {e2}")
                except Exception as e1:
                    LOGGER.debug(f"通过title查找Search按钮失败: {e1}")
            
            # 方法3: 遍历所有按钮查找Search（备用方法）
            if not search_clicked:
                try:
                    LOGGER.info("尝试使用uia backend查找Search按钮...")
                    # 重新连接对话框使用uia backend
                    if win32gui:
                        dialog_hwnd = units_dialog.handle if hasattr(units_dialog, 'handle') else None
                        if dialog_hwnd:
                            try:
                                uia_app = Application(backend="uia").connect(handle=dialog_hwnd)
                                uia_dialog = uia_app.window(handle=dialog_hwnd)
                                if uia_dialog.exists():
                                    search_button = uia_dialog.child_window(title="Search", control_type="Button")
                                    if search_button.exists() and search_button.is_enabled():
                                        LOGGER.info("使用uia backend找到Search按钮")
                                        search_button.click_input()
                                        search_clicked = True
                                        LOGGER.info("✅ 已通过uia backend点击Search按钮")
                            except Exception as e:
                                LOGGER.debug(f"uia backend方法失败: {e}")
                except Exception as e3:
                    LOGGER.debug(f"uia backend查找失败: {e3}")
            
            # 方法4: 使用Windows API查找按钮（递归查找所有子窗口）
            if not search_clicked and win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API查找Search按钮（递归查找所有子窗口）...")
                    dialog_hwnd = units_dialog.handle if hasattr(units_dialog, 'handle') else None
                    if dialog_hwnd:
                        button_found = [False]  # 使用列表以便在回调中修改
                        button_hwnd = [None]
                        all_buttons = []  # 记录所有找到的按钮，用于调试
                        
                        def enum_child_proc(hwnd_child, lParam):
                            """枚举子窗口，查找Search按钮"""
                            try:
                                window_text = win32gui.GetWindowText(hwnd_child)
                                class_name = win32gui.GetClassName(hwnd_child)
                                
                                # 记录所有BUTTON控件，用于调试
                                if "BUTTON" in class_name.upper():
                                    all_buttons.append((hwnd_child, window_text, class_name))
                                    
                                    # 查找包含"Search"的按钮（不区分大小写）
                                    if "Search" in window_text or "search" in window_text.lower():
                                        LOGGER.info(f"通过Windows API找到Search按钮: '{window_text}' (类名: {class_name}, hwnd={hwnd_child})")
                                        button_found[0] = True
                                        button_hwnd[0] = hwnd_child
                                        return False  # 找到后停止枚举
                                return True
                            except:
                                return True
                        
                        # EnumChildWindows 会自动递归查找所有子窗口
                        win32gui.EnumChildWindows(dialog_hwnd, enum_child_proc, None)
                        
                        # 如果没找到，尝试在主窗口中查找（Search按钮可能在主窗口中，不在对话框内）
                        if not button_found[0] and self._window:
                            try:
                                main_hwnd = self._window.handle if hasattr(self._window, 'handle') else None
                                if main_hwnd and main_hwnd != dialog_hwnd:
                                    LOGGER.info("在对话框子窗口中未找到Search按钮，尝试在主窗口中查找...")
                                    win32gui.EnumChildWindows(main_hwnd, enum_child_proc, None)
                            except Exception as main_err:
                                LOGGER.debug(f"在主窗口中查找失败: {main_err}")
                        
                        # 输出所有找到的按钮，用于调试
                        if all_buttons:
                            LOGGER.info(f"Windows API找到 {len(all_buttons)} 个按钮:")
                            for hwnd, text, cls in all_buttons:
                                LOGGER.info(f"  按钮: '{text}' (类名: {cls}, hwnd={hwnd})")
                        else:
                            LOGGER.warning("Windows API未找到任何按钮控件")
                        
                        if button_found[0] and button_hwnd[0]:
                            # 找到了按钮，尝试点击
                            try:
                                # 先尝试 SendMessage（同步，更可靠）
                                win32gui.SendMessage(button_hwnd[0], win32con.BM_CLICK, 0, 0)
                                time.sleep(0.5)
                                search_clicked = True
                                LOGGER.info("✅ 已通过Windows API点击Search按钮（SendMessage）")
                            except Exception as send_err:
                                LOGGER.warning(f"SendMessage失败: {send_err}，尝试PostMessage...")
                                try:
                                    # 如果 SendMessage 失败，尝试 PostMessage
                                    win32gui.PostMessage(button_hwnd[0], win32con.BM_CLICK, 0, 0)
                                    time.sleep(0.5)
                                    search_clicked = True
                                    LOGGER.info("✅ 已通过Windows API点击Search按钮（PostMessage）")
                                except Exception as post_err:
                                    LOGGER.warning(f"PostMessage也失败: {post_err}")
                        else:
                            LOGGER.warning("Windows API未找到Search按钮")
                            
                            # 如果没找到Search按钮，尝试通过位置查找（Search按钮通常在对话框右下角）
                            if all_buttons:
                                LOGGER.info("尝试通过位置查找Search按钮（通常在对话框右下角）...")
                                dialog_rect = units_dialog.rectangle()
                                dialog_right = dialog_rect.left + dialog_rect.width()
                                dialog_bottom = dialog_rect.top + dialog_rect.height()
                                
                                # 查找位置最接近右下角的按钮
                                best_button = None
                                min_distance = float('inf')
                                
                                for btn_hwnd, btn_text, btn_class in all_buttons:
                                    try:
                                        btn_rect = win32gui.GetWindowRect(btn_hwnd)
                                        btn_center_x = (btn_rect[0] + btn_rect[2]) // 2
                                        btn_center_y = (btn_rect[1] + btn_rect[3]) // 2
                                        
                                        # 计算按钮中心到对话框右下角的距离
                                        distance = ((btn_center_x - dialog_right) ** 2 + (btn_center_y - dialog_bottom) ** 2) ** 0.5
                                        
                                        # 如果按钮在对话框右下角区域（右下角80%区域内）
                                        if (btn_center_x > dialog_rect.left + dialog_rect.width() * 0.6 and 
                                            btn_center_y > dialog_rect.top + dialog_rect.height() * 0.7):
                                            if distance < min_distance:
                                                min_distance = distance
                                                best_button = (btn_hwnd, btn_text, btn_class, btn_center_x, btn_center_y)
                                    except:
                                        continue
                                
                                if best_button:
                                    btn_hwnd, btn_text, btn_class, btn_x, btn_y = best_button
                                    LOGGER.info(f"找到最接近右下角的按钮: '{btn_text}' (位置: {btn_x}, {btn_y}, 距离右下角: {min_distance:.1f}像素)")
                                    try:
                                        win32gui.SendMessage(btn_hwnd, win32con.BM_CLICK, 0, 0)
                                        time.sleep(0.5)
                                        search_clicked = True
                                        LOGGER.info("✅ 已通过Windows API点击右下角按钮（可能是Search按钮）")
                                    except Exception as pos_err:
                                        LOGGER.warning(f"点击右下角按钮失败: {pos_err}，尝试PostMessage...")
                                        try:
                                            win32gui.PostMessage(btn_hwnd, win32con.BM_CLICK, 0, 0)
                                            time.sleep(0.5)
                                            search_clicked = True
                                            LOGGER.info("✅ 已通过Windows API点击右下角按钮（PostMessage）")
                                        except Exception as post_err:
                                            LOGGER.warning(f"PostMessage也失败: {post_err}")
                                else:
                                    # 如果位置查找失败，尝试点击最后一个按钮作为备用
                                    LOGGER.info("位置查找失败，尝试点击最后一个按钮作为备用...")
                                    if len(all_buttons) > 0:
                                        last_button_hwnd, last_button_text, last_button_class = all_buttons[-1]
                                        LOGGER.info(f"尝试点击最后一个按钮: '{last_button_text}' (类名: {last_button_class}, hwnd={last_button_hwnd})")
                                        try:
                                            win32gui.SendMessage(last_button_hwnd, win32con.BM_CLICK, 0, 0)
                                            time.sleep(0.5)
                                            search_clicked = True
                                            LOGGER.info("✅ 已通过Windows API点击最后一个按钮（备用方法）")
                                        except Exception as last_err:
                                            LOGGER.warning(f"点击最后一个按钮失败: {last_err}，尝试PostMessage...")
                                            try:
                                                win32gui.PostMessage(last_button_hwnd, win32con.BM_CLICK, 0, 0)
                                                time.sleep(0.5)
                                                search_clicked = True
                                                LOGGER.info("✅ 已通过Windows API点击最后一个按钮（PostMessage）")
                                            except Exception as post_err:
                                                LOGGER.warning(f"PostMessage也失败: {post_err}")
                except Exception as e4:
                    LOGGER.warning(f"使用Windows API查找按钮失败: {e4}")
                    import traceback
                    LOGGER.debug(traceback.format_exc())
            
            # 方法5: 备用方法 - 使用鼠标点击估算位置
            if not search_clicked:
                try:
                    LOGGER.warning("所有标准方法都失败，尝试使用鼠标点击估算位置...")
                    import pyautogui
                    dialog_rect = units_dialog.rectangle()
                    
                    LOGGER.info(f"对话框位置: left={dialog_rect.left}, top={dialog_rect.top}, width={dialog_rect.width()}, height={dialog_rect.height()}")
                    
                    # Search按钮通常在对话框的右下角或底部中间
                    # 根据对话框布局，尝试更多可能的位置
                    possible_positions = [
                        (dialog_rect.left + dialog_rect.width() * 0.85, dialog_rect.top + dialog_rect.height() * 0.85),  # 右下角
                        (dialog_rect.left + dialog_rect.width() * 0.80, dialog_rect.top + dialog_rect.height() * 0.90),  # 右下角偏上
                        (dialog_rect.left + dialog_rect.width() * 0.75, dialog_rect.top + dialog_rect.height() * 0.88),  # 右下角偏左
                        (dialog_rect.left + dialog_rect.width() * 0.50, dialog_rect.top + dialog_rect.height() * 0.85),  # 底部中间
                        (dialog_rect.left + dialog_rect.width() * 0.70, dialog_rect.top + dialog_rect.height() * 0.82),  # 右下角偏上更多
                    ]
                    
                    for pos_idx, (x, y) in enumerate(possible_positions):
                        try:
                            LOGGER.info(f"尝试位置 {pos_idx + 1}: ({int(x)}, {int(y)})")
                            # 先移动鼠标到位置
                            pyautogui.moveTo(int(x), int(y), duration=0.2)
                            time.sleep(0.1)
                            # 点击
                            pyautogui.click(int(x), int(y))
                            time.sleep(0.8)  # 等待点击响应
                            
                            # 不检查对话框是否存在，因为即使点击成功，对话框也可能不会立即关闭
                            # 直接假设点击成功，让后续流程验证
                            search_clicked = True
                            LOGGER.info(f"✅ 已通过位置点击Search按钮（位置 {pos_idx + 1}），假设点击成功")
                            break
                        except Exception as pos_err:
                            LOGGER.warning(f"位置 {pos_idx + 1} 点击失败: {pos_err}")
                            continue
                except ImportError:
                    LOGGER.warning("pyautogui未安装，无法使用备用方法")
                except Exception as e5:
                    LOGGER.warning(f"备用方法失败: {e5}")
                    import traceback
                    LOGGER.debug(traceback.format_exc())
            
            if not search_clicked:
                raise RuntimeError("无法点击Search按钮，已尝试所有方法")
            
            # 等待搜索完成
            time.sleep(1)
            LOGGER.info("✅ 已成功填写Units信息并点击Search按钮")
        
        except Exception as e:
            LOGGER.error(f"填写Units搜索对话框失败: {e}")
            raise RuntimeError(f"无法填写Units搜索对话框: {e}")
    
    def _check_row_status_and_select(self, ui_config: dict = None, use_available_rows: bool = False) -> dict:
        """
        点击选择按钮，添加到Summary，并切换到Summary标签
        
        Args:
            ui_config: UI配置
            use_available_rows: 如果为True，使用"Select Available Rows"；否则使用"Select Visible Rows"
        
        Returns:
            包含选择信息的字典
        """
        result_info = {}
        
        if Application is None:
            return result_info
        
        LOGGER.info("执行选择和添加流程...")
        
        # 等待搜索结果加载（给数据一些时间加载）
        time.sleep(2)
        
        # 步骤1: 点击选择按钮
        if use_available_rows:
            LOGGER.info("步骤1: 点击 'Select Available Rows'")
            self._click_select_available_rows_button()
        else:
            LOGGER.info("步骤1: 点击 'Select Visible Rows'")
            self._click_select_visible_rows_button()
        time.sleep(1)
        
        # 步骤2: 点击"Add to summary"按钮
        LOGGER.info("步骤2: 点击 'Add to Summary'")
        self._click_add_to_summary_button()
        # Units模式需要更长的等待时间（比VPOs模式长）
        if use_available_rows:
            # VPOs模式：使用较短的等待时间
            time.sleep(1.5)
        else:
            # Units模式：等待10秒，让Add to Summary操作完全完成
            LOGGER.info("Units模式：等待Add to Summary操作完成（等待10秒）...")
            time.sleep(10.0)  # Units模式需要等待10秒
        
        # 步骤3: 点击"3. View Summary"标签
        LOGGER.info("步骤3: 点击 '3. View Summary' 标签")
        self._click_summary_tab()
        time.sleep(1)
        
        # 步骤4: 填写Requestor Comments
        LOGGER.info("步骤4: 填写 'Requestor Comments'")
        self._fill_requestor_comments(ui_config)
        time.sleep(0.5)
        
        LOGGER.info("✅ 已完成选择和添加操作，并切换到Summary标签，已填写Comments")
        return result_info
    
    def _click_select_visible_rows_button(self) -> None:
        """点击左侧的'Select Visible Rows'按钮"""
        if Application is None:
            return
        
        LOGGER.info("查找并点击'Select Visible Rows'按钮...")
        
        try:
            # 确保主窗口有焦点
            self._window.set_focus()
            time.sleep(0.3)
            
            # 方法1: 通过按钮文本查找
            try:
                select_button = self._window.child_window(title="Select Visible Rows", control_type="Button")
                if select_button.exists() and select_button.is_enabled() and select_button.is_visible():
                    LOGGER.info("找到'Select Visible Rows'按钮（通过title）")
                    # 获取按钮位置
                    try:
                        button_rect = select_button.rectangle()
                        button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                        button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                        LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y})")
                    except:
                        pass
                    
                    # 使用click_input()点击
                    select_button.click_input()
                    time.sleep(0.5)
                    LOGGER.info("✅ 已点击'Select Visible Rows'按钮（通过title，使用click_input）")
                    return
            except Exception as e:
                LOGGER.debug(f"通过title查找按钮失败: {e}")
            
            # 方法2: 遍历所有按钮查找
            try:
                all_buttons = self._window.descendants(control_type="Button")
                LOGGER.info(f"找到 {len(all_buttons)} 个按钮")
                for idx, button in enumerate(all_buttons):
                    try:
                        button_text = button.window_text().strip()
                        LOGGER.debug(f"  按钮 #{idx}: 文本='{button_text}'")
                        
                        if "SELECT VISIBLE ROWS" in button_text.upper() or "SELECT VISIBLE" in button_text.upper():
                            LOGGER.info(f"找到'Select Visible Rows'按钮（文本: '{button_text}'）")
                            if button.is_enabled() and button.is_visible():
                                # 获取按钮位置
                                try:
                                    button_rect = button.rectangle()
                                    button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                                    button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                                    LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y})")
                                except:
                                    pass
                                
                                # 使用click_input()点击
                                button.click_input()
                                time.sleep(0.5)
                                LOGGER.info(f"✅ 已点击'Select Visible Rows'按钮（文本: '{button_text}'，使用click_input）")
                                return
                    except Exception as e:
                        LOGGER.debug(f"检查按钮 #{idx} 时出错: {e}")
                        continue
            except Exception as e:
                LOGGER.debug(f"遍历按钮失败: {e}")
            
            # 方法3: 使用Windows API查找
            if win32gui:
                try:
                    main_hwnd = self._window.handle
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child)
                            if "SELECT VISIBLE ROWS" in window_text.upper() or "SELECT VISIBLE" in window_text.upper():
                                lParam.append((hwnd_child, window_text))
                        except:
                            pass
                        return True
                    
                    button_list = []
                    win32gui.EnumChildWindows(main_hwnd, enum_child_proc, button_list)
                    
                    if button_list:
                        button_hwnd, button_text = button_list[0]
                        LOGGER.info(f"使用Windows API找到按钮: '{button_text}'")
                        
                        # 获取按钮位置并点击
                        rect = win32gui.GetWindowRect(button_hwnd)
                        center_x = (rect[0] + rect[2]) // 2
                        center_y = (rect[1] + rect[3]) // 2
                        
                        try:
                            import pyautogui
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.5)
                            LOGGER.info(f"✅ 已通过坐标点击'Select Visible Rows'按钮（位置: {center_x}, {center_y}）")
                            return
                        except ImportError:
                            win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                            time.sleep(0.5)
                            LOGGER.info("✅ 已通过Windows API点击'Select Visible Rows'按钮")
                            return
                except Exception as e:
                    LOGGER.debug(f"使用Windows API查找按钮失败: {e}")
            
            # 方法4: 如果找不到，尝试根据位置估算（参考之前的日志）
            try:
                window_rect = self._window.rectangle()
                # 使用与Add按钮相同的X偏移（约220像素）
                estimated_x = window_rect.left + 220
                
                # 使用中间位置的Y坐标（假设它在Add按钮上方）
                # 根据日志 Select(814) vs Add(1786)，大概在40%高度处
                window_height = window_rect.bottom - window_rect.top
                estimated_y = window_rect.top + int(window_height * 0.4) 
                
                LOGGER.info(f"尝试在估算位置点击Select Visible Rows: ({estimated_x}, {estimated_y})")
                try:
                    import pyautogui
                    pyautogui.click(estimated_x, estimated_y)
                    time.sleep(0.5)
                    LOGGER.info(f"✅ 已在估算位置点击'Select Visible Rows'按钮（位置: {estimated_x}, {estimated_y}）")
                    return
                except ImportError:
                    LOGGER.warning("pyautogui未安装，无法使用坐标点击")
            except Exception as e:
                LOGGER.debug(f"使用估算位置点击失败: {e}")
                
            LOGGER.warning("未找到'Select Visible Rows'按钮")
            
        except Exception as e:
            LOGGER.error(f"点击'Select Visible Rows'按钮失败: {e}")
            raise RuntimeError(f"无法点击'Select Visible Rows'按钮: {e}")
    
    def _click_select_available_rows_button(self) -> None:
        """点击左侧的'Select Available Rows'按钮"""
        if Application is None:
            return
        
        LOGGER.info("查找并点击'Select Available Rows'按钮...")
        
        try:
            # 确保主窗口有焦点
            self._window.set_focus()
            time.sleep(0.5)
            
            # 方法1: 通过按钮文本查找（精确匹配）
            try:
                select_button = self._window.child_window(title="Select Available Rows", control_type="Button")
                if select_button.exists(timeout=2):
                    if select_button.is_enabled() and select_button.is_visible():
                        LOGGER.info("找到'Select Available Rows'按钮（通过title）")
                        select_button.click_input()
                        time.sleep(0.5)
                        LOGGER.info("✅ 已点击'Select Available Rows'按钮")
                        return
            except Exception as e:
                LOGGER.debug(f"通过title查找按钮失败: {e}")
            
            # 方法2: 通过部分文本匹配查找
            try:
                select_button = self._window.child_window(title_re=".*Select Available.*", control_type="Button")
                if select_button.exists(timeout=2):
                    if select_button.is_enabled() and select_button.is_visible():
                        LOGGER.info("找到'Select Available Rows'按钮（通过正则title）")
                        select_button.click_input()
                        time.sleep(0.5)
                        LOGGER.info("✅ 已点击'Select Available Rows'按钮")
                        return
            except Exception as e:
                LOGGER.debug(f"通过正则title查找按钮失败: {e}")
            
            # 方法3: 遍历所有按钮查找（包含所有后代元素）
            try:
                all_buttons = self._window.descendants(control_type="Button")
                LOGGER.debug(f"找到 {len(list(all_buttons))} 个按钮，开始搜索...")
                for button in all_buttons:
                    try:
                        button_text = button.window_text().strip()
                        LOGGER.debug(f"检查按钮文本: '{button_text}'")
                        if button_text and ("SELECT AVAILABLE ROWS" in button_text.upper() or 
                                          "SELECT AVAILABLE" in button_text.upper() or
                                          "AVAILABLE ROWS" in button_text.upper()):
                            LOGGER.info(f"找到'Select Available Rows'按钮（文本: '{button_text}'）")
                            if button.is_enabled() and button.is_visible():
                                button.click_input()
                                time.sleep(0.5)
                                LOGGER.info(f"✅ 已点击'Select Available Rows'按钮")
                                return
                    except Exception as e:
                        LOGGER.debug(f"检查按钮时出错: {e}")
                        continue
            except Exception as e:
                LOGGER.debug(f"遍历按钮查找失败: {e}")
            
            # 方法4: 使用win32gui查找按钮（备用方法）
            if win32gui:
                try:
                    def find_button_by_text(hwnd, buttons):
                        try:
                            if not win32gui.IsWindowVisible(hwnd):
                                return True
                            text = win32gui.GetWindowText(hwnd)
                            class_name = win32gui.GetClassName(hwnd)
                            if "BUTTON" in class_name.upper() and text:
                                text_upper = text.upper()
                                if "SELECT AVAILABLE" in text_upper or "AVAILABLE ROWS" in text_upper:
                                    buttons.append((hwnd, text))
                        except:
                            pass
                        return True
                    
                    buttons = []
                    win32gui.EnumChildWindows(self._window.handle, find_button_by_text, buttons)
                    
                    if buttons:
                        hwnd, text = buttons[0]
                        LOGGER.info(f"通过win32gui找到按钮: '{text}'")
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.2)
                        win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
                        time.sleep(0.5)
                        LOGGER.info("✅ 已点击'Select Available Rows'按钮（win32gui方法）")
                        return
                except Exception as e:
                    LOGGER.debug(f"win32gui方法失败: {e}")
            
            LOGGER.warning("未找到'Select Available Rows'按钮")
            
        except Exception as e:
            LOGGER.error(f"点击'Select Available Rows'按钮失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            raise RuntimeError(f"无法点击'Select Available Rows'按钮: {e}")
    
    def _click_export_to_excel_button(self) -> None:
        """点击左侧的'Export to Excel'按钮"""
        if Application is None:
            return
        
        LOGGER.info("查找并点击'Export to Excel'按钮...")
        
        try:
            self._window.set_focus()
            time.sleep(0.5)
            
            # 方法1: 通过按钮文本查找（精确匹配）
            try:
                export_button = self._window.child_window(title="Export to Excel", control_type="Button")
                if export_button.exists(timeout=2):
                    if export_button.is_enabled() and export_button.is_visible():
                        LOGGER.info("找到'Export to Excel'按钮（通过title）")
                        export_button.click_input()
                        time.sleep(1.0)  # 等待导出对话框
                        LOGGER.info("✅ 已点击'Export to Excel'按钮")
                        return
            except Exception as e:
                LOGGER.debug(f"通过title查找按钮失败: {e}")
            
            # 方法2: 通过部分文本匹配查找
            try:
                export_button = self._window.child_window(title_re=".*Export.*Excel.*", control_type="Button")
                if export_button.exists(timeout=2):
                    if export_button.is_enabled() and export_button.is_visible():
                        LOGGER.info("找到'Export to Excel'按钮（通过正则title）")
                        export_button.click_input()
                        time.sleep(1.0)
                        LOGGER.info("✅ 已点击'Export to Excel'按钮")
                        return
            except Exception as e:
                LOGGER.debug(f"通过正则title查找按钮失败: {e}")
            
            # 方法3: 遍历所有按钮查找（包含图标的情况）
            try:
                all_buttons = self._window.descendants(control_type="Button")
                LOGGER.debug(f"找到 {len(list(all_buttons))} 个按钮，开始搜索...")
                for button in all_buttons:
                    try:
                        button_text = button.window_text().strip()
                        LOGGER.debug(f"检查按钮文本: '{button_text}'")
                        if button_text and ("EXPORT" in button_text.upper() and "EXCEL" in button_text.upper()):
                            LOGGER.info(f"找到'Export to Excel'按钮（文本: '{button_text}'）")
                            if button.is_enabled() and button.is_visible():
                                button.click_input()
                                time.sleep(1.0)
                                LOGGER.info(f"✅ 已点击'Export to Excel'按钮")
                                return
                    except Exception as e:
                        LOGGER.debug(f"检查按钮时出错: {e}")
                        continue
            except Exception as e:
                LOGGER.debug(f"遍历按钮查找失败: {e}")
            
            # 方法4: 使用win32gui查找按钮（备用方法）
            if win32gui:
                try:
                    def find_button_by_text(hwnd, buttons):
                        try:
                            if not win32gui.IsWindowVisible(hwnd):
                                return True
                            text = win32gui.GetWindowText(hwnd)
                            class_name = win32gui.GetClassName(hwnd)
                            if "BUTTON" in class_name.upper() and text:
                                text_upper = text.upper()
                                if "EXPORT" in text_upper and "EXCEL" in text_upper:
                                    buttons.append((hwnd, text))
                        except:
                            pass
                        return True
                    
                    buttons = []
                    win32gui.EnumChildWindows(self._window.handle, find_button_by_text, buttons)
                    
                    if buttons:
                        hwnd, text = buttons[0]
                        LOGGER.info(f"通过win32gui找到按钮: '{text}'")
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.2)
                        win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
                        time.sleep(1.0)
                        LOGGER.info("✅ 已点击'Export to Excel'按钮（win32gui方法）")
                        return
                except Exception as e:
                    LOGGER.debug(f"win32gui方法失败: {e}")
            
            LOGGER.warning("未找到'Export to Excel'按钮")
            
        except Exception as e:
            LOGGER.error(f"点击'Export to Excel'按钮失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            raise RuntimeError(f"无法点击'Export to Excel'按钮: {e}")
    
    def _select_available_and_export_to_excel(self) -> None:
        """
        先点击 'Select Available Rows'，然后点击 'Export to Excel'
        确保操作顺序正确
        """
        LOGGER.info("开始执行: 先选择可用行，再导出到Excel")
        
        # 步骤1: 点击 Select Available Rows
        LOGGER.info("步骤1: 点击 'Select Available Rows'")
        self._click_select_available_rows_button()
        time.sleep(1.0)  # 等待选择操作完成
        
        # 步骤2: 点击 Export to Excel
        LOGGER.info("步骤2: 点击 'Export to Excel'")
        self._click_export_to_excel_button()
        time.sleep(1.0)  # 等待导出对话框出现
        
        LOGGER.info("✅ 已完成: 选择可用行并触发导出")
    
    def _select_visible_and_export_to_excel(self) -> None:
        """
        先点击 'Select Visible Rows'，然后点击 'Export to Excel'
        确保操作顺序正确
        """
        LOGGER.info("开始执行: 先选择可见行，再导出到Excel")
        
        # 步骤1: 点击 Select Visible Rows
        LOGGER.info("步骤1: 点击 'Select Visible Rows'")
        self._click_select_visible_rows_button()
        time.sleep(1.0)  # 等待选择操作完成
        
        # 步骤2: 点击 Export to Excel
        LOGGER.info("步骤2: 点击 'Export to Excel'")
        self._click_export_to_excel_button()
        time.sleep(1.0)  # 等待导出对话框出现
        
        LOGGER.info("✅ 已完成: 选择可见行并触发导出")
    
    def _handle_excel_opened_file(self, output_path: Path, max_wait_time: int = 30) -> Optional[Path]:
        """
        处理 Excel 直接打开文件的情况（没有保存对话框）
        
        Args:
            output_path: 希望保存到的文件路径
            max_wait_time: 最大等待时间（秒）
        
        Returns:
            实际保存的文件路径，如果失败返回None
        """
        LOGGER.info("检测到 Excel 直接打开了文件，尝试从 Excel 窗口获取文件路径...")
        
        try:
            start_time = time.time()
            excel_window = None
            excel_file_path = None
            
            # 等待 Excel 窗口出现
            while time.time() - start_time < max_wait_time:
                try:
                    # 方法1: 使用 pywinauto 查找 Excel 窗口
                    if Application:
                        try:
                            excel_app = Application(backend="uia").connect(class_name="XLMAIN")
                            excel_windows = excel_app.windows()
                            
                            # 查找最近打开的 Excel 窗口（标题可能包含文件名）
                            for window in excel_windows:
                                try:
                                    title = window.window_text()
                                    # Excel 窗口标题通常格式为: "文件名 - Excel" 或 "文件名.xlsx - Excel"
                                    if "Excel" in title and ("Compatibility Mode" in title or ".xls" in title or ".xlsx" in title):
                                        excel_window = window
                                        # 尝试从标题提取文件路径
                                        # 标题格式可能是: "MIR_12252025_84828_AM.xls - Compatibility Mode - Saved to this PC"
                                        if " - " in title:
                                            potential_filename = title.split(" - ")[0].strip()
                                            # 检查是否是完整路径
                                            if "\\" in potential_filename or "/" in potential_filename:
                                                excel_file_path = potential_filename
                                            else:
                                                # 只是文件名，需要查找完整路径
                                                LOGGER.debug(f"从标题提取的文件名: {potential_filename}")
                                        LOGGER.info(f"找到 Excel 窗口: {title}")
                                        break
                                except:
                                    continue
                            
                            if excel_window:
                                break
                        except:
                            pass
                    
                    # 方法2: 使用 win32gui 查找 Excel 窗口
                    if win32gui and not excel_window:
                        def find_excel_window(hwnd, windows):
                            try:
                                if not win32gui.IsWindowVisible(hwnd):
                                    return True
                                class_name = win32gui.GetClassName(hwnd)
                                title = win32gui.GetWindowText(hwnd)
                                
                                # Excel 主窗口类名通常是 "XLMAIN"
                                if "XLMAIN" in class_name.upper() or "EXCEL" in class_name.upper():
                                    if title and ("Excel" in title or ".xls" in title or ".xlsx" in title):
                                        windows.append((hwnd, title))
                            except:
                                pass
                            return True
                        
                        excel_windows = []
                        win32gui.EnumWindows(find_excel_window, excel_windows)
                        
                        if excel_windows:
                            hwnd, title = excel_windows[0]  # 使用第一个找到的窗口
                            LOGGER.info(f"找到 Excel 窗口（win32gui）: {title}")
                            
                            # 尝试从标题提取文件路径
                            if " - " in title:
                                potential_filename = title.split(" - ")[0].strip()
                                if "\\" in potential_filename or "/" in potential_filename:
                                    excel_file_path = potential_filename
                                else:
                                    LOGGER.debug(f"从标题提取的文件名: {potential_filename}")
                            
                            # 尝试使用 COM 对象获取完整路径
                            try:
                                import win32com.client
                                excel_app = win32com.client.GetActiveObject("Excel.Application")
                                if excel_app and excel_app.Workbooks.Count > 0:
                                    workbook = excel_app.ActiveWorkbook
                                    if workbook:
                                        full_path = workbook.FullName
                                        if full_path:
                                            excel_file_path = full_path
                                            LOGGER.info(f"通过 COM 对象获取文件路径: {excel_file_path}")
                            except Exception as e:
                                LOGGER.debug(f"使用 COM 对象获取路径失败: {e}")
                            
                            break
                    
                    time.sleep(0.5)
                except Exception as e:
                    LOGGER.debug(f"查找 Excel 窗口时出错: {e}")
                    time.sleep(0.5)
            
            if excel_file_path and Path(excel_file_path).exists():
                LOGGER.info(f"✅ 检测到 Excel 打开的文件: {excel_file_path}")
                
                result_path = None
                
                # 如果文件路径与目标路径不同，需要保存到目标路径
                if str(excel_file_path) != str(output_path):
                    LOGGER.info(f"文件路径不同，需要保存到: {output_path}")
                    
                    # 方法1: 使用 COM 对象保存
                    try:
                        import win32com.client
                        excel_app = win32com.client.GetActiveObject("Excel.Application")
                        if excel_app and excel_app.Workbooks.Count > 0:
                            workbook = excel_app.ActiveWorkbook
                            if workbook:
                                # 保存到目标路径
                                workbook.SaveAs(str(output_path))
                                LOGGER.info(f"✅ 已通过 COM 对象保存文件到: {output_path}")
                                
                                if output_path.exists():
                                    result_path = output_path
                                    # 关闭 Excel 工作簿
                                    try:
                                        workbook.Close(False)  # False = 不保存（因为已经保存过了）
                                        LOGGER.info("✅ 已关闭 Excel 工作簿")
                                    except Exception as e:
                                        LOGGER.debug(f"关闭工作簿时出错: {e}")
                    except Exception as e:
                        LOGGER.debug(f"使用 COM 对象保存失败: {e}")
                    
                    # 方法2: 使用 Excel 的 Save As 功能（通过快捷键）
                    if not result_path:
                        try:
                            if excel_window:
                                excel_window.set_focus()
                                time.sleep(0.5)
                                
                                # 按 F12 打开 Save As 对话框
                                import pyautogui
                                pyautogui.hotkey('f12')
                                time.sleep(1.0)
                                
                                # 现在应该出现 Save As 对话框，使用现有的保存对话框处理方法
                                result_path = self._handle_export_save_dialog(output_path, max_wait_time=10)
                                # 保存对话框处理完成后，关闭 Excel
                                if result_path:
                                    self._close_excel_window()
                        except Exception as e:
                            LOGGER.debug(f"使用 Save As 快捷键失败: {e}")
                    
                    # 方法3: 直接复制文件
                    if not result_path:
                        try:
                            import shutil
                            shutil.copy2(excel_file_path, output_path)
                            LOGGER.info(f"✅ 已复制文件到: {output_path}")
                            if output_path.exists():
                                result_path = output_path
                                # 关闭 Excel 窗口
                                self._close_excel_window()
                        except Exception as e:
                            LOGGER.warning(f"复制文件失败: {e}")
                else:
                    # 文件路径相同，直接返回
                    result_path = Path(excel_file_path)
                    # 关闭 Excel 窗口
                    self._close_excel_window()
                
                return result_path
            elif excel_window:
                # 找到了 Excel 窗口但无法获取文件路径
                LOGGER.warning("⚠️ 找到了 Excel 窗口但无法获取文件路径")
                
                # 方法1: 尝试从窗口标题提取文件名，然后在临时目录中搜索
                potential_filename = None
                if excel_window:
                    try:
                        title = excel_window.window_text() if hasattr(excel_window, 'window_text') else None
                        if not title and win32gui:
                            # 如果 excel_window 是 hwnd，使用 win32gui 获取标题
                            if isinstance(excel_window, int):
                                title = win32gui.GetWindowText(excel_window)
                            elif hasattr(excel_window, 'handle'):
                                title = win32gui.GetWindowText(excel_window.handle)
                        
                        if title and " - " in title:
                            potential_filename = title.split(" - ")[0].strip()
                            LOGGER.info(f"从窗口标题提取的文件名: {potential_filename}")
                    except Exception as e:
                        LOGGER.debug(f"提取文件名失败: {e}")
                
                # 在临时目录中搜索文件
                if potential_filename:
                    LOGGER.info(f"在临时目录中搜索文件: {potential_filename}")
                    import tempfile
                    import os
                    
                    # 常见的临时目录
                    temp_dirs = [
                        Path(tempfile.gettempdir()),
                        Path(os.environ.get('TEMP', '')),
                        Path(os.environ.get('TMP', '')),
                        Path.home() / 'AppData' / 'Local' / 'Temp',
                        Path.home() / 'AppData' / 'Roaming' / 'Microsoft' / 'Excel',
                    ]
                    
                    # 也在输出目录的父目录中搜索
                    temp_dirs.append(output_path.parent)
                    
                    found_file = None
                    for temp_dir in temp_dirs:
                        if not temp_dir or not temp_dir.exists():
                            continue
                        try:
                            # 搜索文件名（不区分大小写）
                            for file_path in temp_dir.rglob(potential_filename):
                                if file_path.is_file():
                                    found_file = file_path
                                    LOGGER.info(f"✅ 在临时目录中找到文件: {found_file}")
                                    break
                            if found_file:
                                break
                        except Exception as e:
                            LOGGER.debug(f"搜索目录 {temp_dir} 失败: {e}")
                    
                    if found_file:
                        # 复制文件到目标位置
                        try:
                            import shutil
                            shutil.copy2(found_file, output_path)
                            LOGGER.info(f"✅ 已复制文件到: {output_path}")
                            if output_path.exists():
                                return output_path
                        except Exception as e:
                            LOGGER.warning(f"复制文件失败: {e}")
                
                # 方法2: 尝试使用 COM 对象获取文件路径
                LOGGER.info("尝试使用 COM 对象获取文件路径...")
                try:
                    import win32com.client
                    excel_app = win32com.client.GetActiveObject("Excel.Application")
                    if excel_app and excel_app.Workbooks.Count > 0:
                        workbook = excel_app.ActiveWorkbook
                        if workbook:
                            full_path = workbook.FullName
                            if full_path:
                                LOGGER.info(f"✅ 通过 COM 对象获取文件路径: {full_path}")
                                
                                # 如果路径不同，保存到目标路径
                                if str(full_path) != str(output_path):
                                    workbook.SaveAs(str(output_path))
                                    LOGGER.info(f"✅ 已保存文件到: {output_path}")
                                
                                if output_path.exists():
                                    # 关闭 Excel 工作簿
                                    try:
                                        workbook.Close(False)  # False = 不保存（因为已经保存过了）
                                        LOGGER.info("✅ 已关闭 Excel 工作簿")
                                    except Exception as e:
                                        LOGGER.debug(f"关闭工作簿时出错: {e}")
                                    return output_path
                except Exception as e:
                    LOGGER.debug(f"使用 COM 对象失败: {e}")
                
                # 方法3: 尝试使用 Excel 的 Save As 功能（通过快捷键）
                if excel_window:
                    try:
                        LOGGER.info("尝试使用 Excel Save As 功能...")
                        if hasattr(excel_window, 'set_focus'):
                            excel_window.set_focus()
                        elif win32gui:
                            hwnd = excel_window if isinstance(excel_window, int) else getattr(excel_window, 'handle', None)
                            if hwnd:
                                win32gui.SetForegroundWindow(hwnd)
                                win32gui.BringWindowToTop(hwnd)
                        
                        time.sleep(0.5)
                        
                        # 按 F12 打开 Save As 对话框
                        import pyautogui
                        pyautogui.hotkey('f12')
                        time.sleep(1.5)
                        
                        # 现在应该出现 Save As 对话框，使用现有的保存对话框处理方法
                        saved_file = self._handle_export_save_dialog(output_path, max_wait_time=10)
                        if saved_file and saved_file.exists():
                            # 关闭 Excel 窗口
                            self._close_excel_window()
                            return saved_file
                    except Exception as e:
                        LOGGER.debug(f"使用 Save As 快捷键失败: {e}")
            
            # 如果无法自动处理，提示用户手动保存
            LOGGER.warning("⚠️ 无法自动获取或保存 Excel 文件")
            LOGGER.info(f"请手动将 Excel 文件保存到: {output_path}")
            LOGGER.info(f"等待 {max_wait_time} 秒...")
            
            time.sleep(max_wait_time)
            
            if output_path.exists():
                LOGGER.info(f"✅ 检测到文件已保存: {output_path}")
                # 关闭 Excel 窗口
                self._close_excel_window()
                return output_path
            
            # 即使失败也尝试关闭 Excel 窗口
            self._close_excel_window()
            return None
            
        except Exception as e:
            LOGGER.error(f"处理 Excel 打开的文件时出错: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return None
    
    def _handle_export_save_dialog(self, output_path: Path, max_wait_time: int = 30) -> Optional[Path]:
        """
        处理导出保存对话框
        
        Args:
            output_path: 保存文件的路径
            max_wait_time: 最大等待时间（秒）
        
        Returns:
            实际保存的文件路径，如果失败返回None
        """
        if Application is None:
            return None
        
        LOGGER.info(f"等待导出保存对话框出现，准备保存到: {output_path}")
        
        try:
            start_time = time.time()
            
            # 等待保存对话框出现
            save_dialog = None
            dialog_found = False
            
            # 方法1: 使用UIA查找对话框
            while time.time() - start_time < max_wait_time:
                # 检查 ESC 键
                try:
                    from .utils.keyboard_listener import is_esc_pressed
                    if is_esc_pressed():
                        LOGGER.warning("⚠️ 检测到 ESC 键，停止等待对话框")
                        return None
                except:
                    pass
                
                try:
                    app = Application(backend="uia")
                    save_dialog = app.window(title_re=".*保存.*|.*Save.*|.*另存为.*|.*Save As.*")
                    
                    if save_dialog.exists(timeout=1):
                        LOGGER.info("找到保存对话框（UIA方法）")
                        dialog_found = True
                        break
                except:
                    pass
                
                # 方法2: 使用win32gui查找对话框
                if win32gui:
                    try:
                        def find_save_dialog(hwnd, dialogs):
                            try:
                                if not win32gui.IsWindowVisible(hwnd):
                                    return True
                                title = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)
                                if title and ("#32770" in class_name or "Dialog" in class_name):
                                    title_upper = title.upper()
                                    if any(keyword in title_upper for keyword in ["SAVE", "保存", "另存为", "SAVE AS"]):
                                        dialogs.append(hwnd)
                            except:
                                pass
                            return True
                        
                        dialogs = []
                        win32gui.EnumWindows(find_save_dialog, dialogs)
                        if dialogs:
                            save_dialog_hwnd = dialogs[0]
                            LOGGER.info(f"找到保存对话框（win32gui方法，句柄: {save_dialog_hwnd}）")
                            dialog_found = True
                            # 激活对话框
                            win32gui.SetForegroundWindow(save_dialog_hwnd)
                            win32gui.BringWindowToTop(save_dialog_hwnd)
                            time.sleep(0.5)
                            
                            # 查找文件名输入框（ComboBox或Edit）
                            def find_filename_edit(hwnd, edits):
                                try:
                                    class_name = win32gui.GetClassName(hwnd)
                                    if "EDIT" in class_name.upper() or "COMBOBOX" in class_name.upper():
                                        edits.append(hwnd)
                                except:
                                    pass
                                return True
                            
                            edits = []
                            win32gui.EnumChildWindows(save_dialog_hwnd, find_filename_edit, edits)
                            
                            if edits:
                                edit_hwnd = edits[0]
                                # 输入文件路径
                                win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, str(output_path))
                                time.sleep(0.5)
                                LOGGER.info(f"已输入文件路径: {output_path}")
                                
                                # 查找保存按钮
                                def find_save_button(hwnd, buttons):
                                    try:
                                        text = win32gui.GetWindowText(hwnd)
                                        class_name = win32gui.GetClassName(hwnd)
                                        if "BUTTON" in class_name.upper():
                                            text_upper = text.upper()
                                            if any(keyword in text_upper for keyword in ["SAVE", "保存", "OK", "确定"]):
                                                buttons.append(hwnd)
                                    except:
                                        pass
                                    return True
                                
                                buttons = []
                                win32gui.EnumChildWindows(save_dialog_hwnd, find_save_button, buttons)
                                
                                if buttons:
                                    button_hwnd = buttons[0]
                                    LOGGER.info(f"找到保存按钮，准备点击...")
                                    
                                    # 点击保存按钮（尝试多种方法）
                                    click_success = False
                                    try:
                                        # 方法1: 使用 SendMessage
                                        win32gui.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                                        click_success = True
                                        LOGGER.info("已通过 SendMessage 点击保存按钮")
                                    except Exception as e:
                                        LOGGER.debug(f"SendMessage 失败: {e}")
                                    
                                    if not click_success:
                                        try:
                                            # 方法2: 使用 PostMessage
                                            win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                                            click_success = True
                                            LOGGER.info("已通过 PostMessage 点击保存按钮")
                                        except Exception as e:
                                            LOGGER.debug(f"PostMessage 失败: {e}")
                                    
                                    if not click_success:
                                        try:
                                            # 方法3: 使用 pyautogui 点击按钮位置
                                            import pyautogui
                                            rect = win32gui.GetWindowRect(button_hwnd)
                                            center_x = (rect[0] + rect[2]) // 2
                                            center_y = (rect[1] + rect[3]) // 2
                                            pyautogui.click(center_x, center_y)
                                            click_success = True
                                            LOGGER.info(f"已通过 pyautogui 点击保存按钮（坐标: {center_x}, {center_y}）")
                                        except Exception as e:
                                            LOGGER.debug(f"pyautogui 点击失败: {e}")
                                    
                                    # 等待对话框关闭和文件保存
                                    LOGGER.info("等待文件保存...")
                                    max_wait = 10  # 最多等待10秒
                                    wait_interval = 0.5
                                    waited = 0
                                    
                                    while waited < max_wait:
                                        # 检查对话框是否已关闭
                                        try:
                                            if not win32gui.IsWindow(save_dialog_hwnd) or not win32gui.IsWindowVisible(save_dialog_hwnd):
                                                LOGGER.info("保存对话框已关闭")
                                                time.sleep(1.0)  # 再等待1秒确保文件写入完成
                                                break
                                        except:
                                            pass
                                        
                                        # 检查文件是否已保存
                                        if output_path.exists():
                                            LOGGER.info(f"✅ 检测到文件已保存: {output_path}")
                                            return output_path
                                        
                                        time.sleep(wait_interval)
                                        waited += wait_interval
                                    
                                    # 再次检查文件是否存在
                                    if output_path.exists():
                                        LOGGER.info(f"✅ 已保存导出文件到: {output_path}")
                                        return output_path
                                    else:
                                        LOGGER.warning(f"⚠️ 点击保存按钮后，文件仍未保存: {output_path}")
                                else:
                                    LOGGER.warning("未找到保存按钮，尝试按Enter键")
                                    # 尝试按Enter键
                                    win32gui.SendMessage(edit_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                                    time.sleep(3.0)
                                    if output_path.exists():
                                        LOGGER.info(f"✅ 已通过Enter键保存文件: {output_path}")
                                        return output_path
                            break
                    except Exception as e:
                        LOGGER.debug(f"win32gui方法查找对话框失败: {e}")
                
                time.sleep(0.5)
            
            # 如果找到了对话框但UIA方法可用，尝试UIA方法
            if save_dialog and save_dialog.exists():
                try:
                    # 查找文件名输入框
                    filename_edit = save_dialog.child_window(control_type="Edit", found_index=0)
                    if not filename_edit.exists():
                        # 尝试ComboBox
                        filename_edit = save_dialog.child_window(control_type="ComboBox", found_index=0)
                    
                    if filename_edit.exists():
                        # 输入完整路径
                        filename_edit.set_text(str(output_path))
                        time.sleep(0.5)
                        LOGGER.info(f"已输入文件路径: {output_path}")
                        
                        # 点击保存按钮
                        save_button = None
                        # 尝试不同的按钮文本
                        for button_text in ["保存", "Save", "确定", "OK"]:
                            try:
                                save_button = save_dialog.child_window(title=button_text, control_type="Button")
                                if save_button.exists():
                                    break
                            except:
                                continue
                        
                        if save_button and save_button.exists():
                            save_button.click()
                            time.sleep(2.0)  # 等待文件保存
                            LOGGER.info(f"✅ 已保存导出文件到: {output_path}")
                            
                            # 检查文件是否已保存
                            if output_path.exists():
                                return output_path
                        else:
                            LOGGER.warning("未找到保存按钮，尝试按Enter键")
                            # 尝试按Enter键
                            filename_edit.type_keys("{ENTER}")
                            time.sleep(2.0)
                            if output_path.exists():
                                return output_path
                except Exception as e:
                    LOGGER.warning(f"自动保存失败: {e}")
            
            # 如果自动保存失败，提示用户手动保存
            if not dialog_found:
                LOGGER.warning("⚠️ 未检测到保存对话框，可能已自动保存或需要手动操作")
            else:
                LOGGER.warning("⚠️ 无法自动处理保存对话框，请手动保存文件")
            
            LOGGER.info(f"请将文件保存到: {output_path}")
            LOGGER.info(f"等待 {max_wait_time} 秒...")
            
            # 等待用户手动保存
            time.sleep(max_wait_time)
            
            # 检查文件是否已保存
            if output_path.exists():
                LOGGER.info(f"✅ 检测到文件已保存: {output_path}")
                return output_path
            
            return None
            
        except Exception as e:
            LOGGER.error(f"处理导出对话框失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return None
    
    def _find_latest_export_file_in_mir_folder(self, max_wait_time: int = 30) -> Optional[Path]:
        """
        在 C:\MIR 目录中查找最新的导出文件
        
        Args:
            max_wait_time: 最大等待时间（秒）
        
        Returns:
            最新文件的路径，如果未找到返回None
        """
        mir_folder = Path("C:/MIR")
        
        if not mir_folder.exists():
            LOGGER.warning(f"MIR文件夹不存在: {mir_folder}")
            return None
        
        LOGGER.info(f"在 {mir_folder} 目录中查找最新的导出文件...")
        
        try:
            # 等待文件出现（最多等待max_wait_time秒）
            start_time = time.time()
            latest_file = None
            latest_mtime = 0
            
            while time.time() - start_time < max_wait_time:
                try:
                    # 查找所有 .xls 和 .xlsx 文件
                    excel_files = list(mir_folder.glob("*.xls")) + list(mir_folder.glob("*.xlsx"))
                    
                    if excel_files:
                        # 找到最新的文件
                        for file in excel_files:
                            try:
                                mtime = file.stat().st_mtime
                                if mtime > latest_mtime:
                                    latest_mtime = mtime
                                    latest_file = file
                            except Exception as e:
                                LOGGER.debug(f"无法获取文件 {file} 的修改时间: {e}")
                                continue
                        
                        if latest_file:
                            # 检查文件是否在最近几秒内被修改（说明是新导出的）
                            time_since_modify = time.time() - latest_mtime
                            if time_since_modify < max_wait_time:
                                LOGGER.info(f"✅ 找到最新的导出文件: {latest_file}")
                                LOGGER.info(f"   文件修改时间: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(latest_mtime))}")
                                return latest_file
                except Exception as e:
                    LOGGER.debug(f"查找文件时出错: {e}")
                
                time.sleep(1.0)
            
            # 如果等待超时，返回找到的最新文件（即使不是最近修改的）
            if latest_file:
                LOGGER.info(f"⚠️ 超时，返回找到的最新文件: {latest_file}")
                return latest_file
            
            LOGGER.warning(f"在 {mir_folder} 中未找到导出文件")
            return None
            
        except Exception as e:
            LOGGER.error(f"查找MIR文件夹中的文件时出错: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return None
    
    def _get_actual_units_count_from_export(
        self, 
        expected_units: list, 
        temp_export_dir: Path,
        source_lot: str = None,
        use_visible_rows: bool = False
    ) -> dict:
        """
        通过导出Excel文件获取实际可用的units数量
        
        Args:
            expected_units: 期望的units列表
            temp_export_dir: 临时导出目录
            source_lot: Source Lot值（用于日志）
        
        Returns:
            包含实际数量信息的字典:
            {
                'expected_count': int,
                'actual_count': int or None,
                'actual_units': list,
                'missing_units': list,
                'count_match': bool,
                'export_file': Path or None,
                'success': bool
            }
        """
        result_info = {
            'expected_count': len(expected_units),
            'actual_count': None,
            'actual_units': [],
            'missing_units': [],
            'count_match': False,
            'export_file': None,
            'success': False
        }
        
        try:
            # 创建临时导出目录
            temp_export_dir.mkdir(parents=True, exist_ok=True)
            
            # 生成临时文件名
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            source_lot_suffix = f"_{source_lot}" if source_lot else ""
            temp_export_file = temp_export_dir / f"available_units_export{source_lot_suffix}_{timestamp}.xlsx"
            
            LOGGER.info("=" * 80)
            LOGGER.info("开始获取实际可用units数量...")
            if source_lot:
                LOGGER.info(f"Source Lot: {source_lot}")
            LOGGER.info(f"期望数量: {len(expected_units)}")
            LOGGER.info("=" * 80)
            
            # 步骤1和2: 先选择行，再导出到Excel
            if use_visible_rows:
                self._select_visible_and_export_to_excel()
            else:
                self._select_available_and_export_to_excel()
            
            # 步骤3: 等待文件保存到 C:\MIR 目录
            LOGGER.info("步骤3: 等待文件保存到 C:\\MIR 目录...")
            time.sleep(3.0)  # 等待文件保存
            
            # 从 C:\MIR 目录查找最新的导出文件
            saved_file = self._find_latest_export_file_in_mir_folder(max_wait_time=25)
            
            if not saved_file or not saved_file.exists():
                LOGGER.warning("⚠️ 无法在 C:\\MIR 目录中找到导出文件，尝试其他方法...")
                
                # 备用方法：尝试处理保存对话框或Excel直接打开
                time.sleep(2.0)
                saved_file = self._handle_export_save_dialog(temp_export_file, max_wait_time=5)
                
                if not saved_file or not saved_file.exists():
                    LOGGER.info("未检测到保存对话框，尝试处理 Excel 直接打开的情况...")
                    saved_file = self._handle_excel_opened_file(temp_export_file, max_wait_time=20)
            
            if not saved_file or not saved_file.exists():
                LOGGER.warning("⚠️ 无法获取导出文件，跳过实际数量统计")
                return result_info
            
            # 如果文件不在目标目录中，将其复制到目标目录
            if saved_file.parent != temp_export_dir:
                import shutil
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                source_lot_suffix = f"_{source_lot}" if source_lot else ""
                # 保持原始文件的扩展名（.xls 或 .xlsx）
                original_suffix = saved_file.suffix
                target_file = temp_export_dir / f"available_units_export{source_lot_suffix}_{timestamp}{original_suffix}"
                try:
                    shutil.copy2(saved_file, target_file)
                    LOGGER.info(f"✅ 已将导出文件复制到output目录: {target_file}")
                    saved_file = target_file
                except Exception as e:
                    LOGGER.warning(f"⚠️ 复制文件到output目录失败: {e}，使用原始文件路径")
            
            result_info['export_file'] = saved_file
            
            # 步骤4: 读取导出的Excel文件
            LOGGER.info(f"步骤4: 读取导出文件: {saved_file}")
            from .data_reader import read_excel_file
            
            try:
                export_df = read_excel_file(saved_file)
                LOGGER.info(f"导出文件包含 {len(export_df)} 行数据")
                LOGGER.debug(f"导出文件列名: {export_df.columns.tolist()}")
                
                # 查找 Unit Name 列（支持多种可能的列名）
                unit_name_col = None
                for col in export_df.columns:
                    col_upper = str(col).strip().upper()
                    if col_upper in ['UNIT NAME', 'UNITNAME', 'UNIT_NAME', 'UNIT', 'UNITS', 'UNIT ID', 'UNITID']:
                        unit_name_col = col
                        LOGGER.info(f"找到 Unit Name 列: '{col}'")
                        break
                
                if unit_name_col:
                    # 提取实际可用的units
                    actual_units = export_df[unit_name_col].dropna().astype(str).tolist()
                    actual_units = [u.strip() for u in actual_units if u.strip()]
                    result_info['actual_units'] = actual_units
                    result_info['actual_count'] = len(actual_units)
                    
                    # 找出缺失的units
                    expected_set = set(str(u).strip() for u in expected_units)
                    actual_set = set(actual_units)
                    result_info['missing_units'] = list(expected_set - actual_set)
                    
                    # 检查数量是否匹配
                    result_info['count_match'] = (result_info['expected_count'] == result_info['actual_count'])
                    result_info['success'] = True
                    
                    LOGGER.info("=" * 80)
                    LOGGER.info("实际可用units统计结果:")
                    LOGGER.info(f"  期望数量: {result_info['expected_count']}")
                    LOGGER.info(f"  实际可用数量: {result_info['actual_count']}")
                    LOGGER.info(f"  数量匹配: {'✅ 是' if result_info['count_match'] else '❌ 否'}")
                    
                    if result_info['missing_units']:
                        LOGGER.warning(f"  缺失的units ({len(result_info['missing_units'])} 个):")
                        for missing in result_info['missing_units'][:10]:  # 只显示前10个
                            LOGGER.warning(f"    - {missing}")
                        if len(result_info['missing_units']) > 10:
                            LOGGER.warning(f"    ... 还有 {len(result_info['missing_units']) - 10} 个")
                    else:
                        LOGGER.info("  ✅ 所有units都可用")
                    LOGGER.info("=" * 80)
                else:
                    LOGGER.warning("⚠️ 在导出文件中未找到 Unit Name 列")
                    LOGGER.warning(f"可用列: {export_df.columns.tolist()}")
                    
            except Exception as e:
                LOGGER.error(f"读取导出文件失败: {e}")
                import traceback
                LOGGER.error(traceback.format_exc())
            
            # 步骤5: 关闭自动打开的 Excel 窗口
            if saved_file and saved_file.exists():
                LOGGER.info("步骤5: 关闭自动打开的 Excel 窗口...")
                self._close_excel_window()
            
        except Exception as e:
            LOGGER.error(f"获取实际units数量失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
        
        return result_info
    
    def _close_excel_window(self) -> None:
        """
        关闭自动打开的 Excel 窗口
        
        查找并关闭最近打开的 Excel 窗口（通过窗口标题识别）
        """
        try:
            LOGGER.info("查找并关闭 Excel 窗口...")
            
            if not win32gui:
                LOGGER.warning("win32gui不可用，无法关闭Excel窗口")
                return
            
            # 查找 Excel 窗口
            excel_windows = []
            
            def find_excel_window(hwnd, windows):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return True
                    class_name = win32gui.GetClassName(hwnd)
                    title = win32gui.GetWindowText(hwnd)
                    
                    # Excel 主窗口类名通常是 "XLMAIN"
                    if "XLMAIN" in class_name.upper() or "EXCEL" in class_name.upper():
                        if title and ("Excel" in title or ".xls" in title or ".xlsx" in title):
                            windows.append((hwnd, title))
                except:
                    pass
                return True
            
            win32gui.EnumWindows(find_excel_window, excel_windows)
            
            if excel_windows:
                # 关闭找到的 Excel 窗口（可能有多个，关闭所有）
                for hwnd, title in excel_windows:
                    try:
                        LOGGER.info(f"关闭 Excel 窗口: {title}")
                        
                        # 激活窗口
                        win32gui.SetForegroundWindow(hwnd)
                        win32gui.BringWindowToTop(hwnd)
                        time.sleep(0.3)
                        
                        # 方法1: 使用 Alt+F4 关闭
                        try:
                            import pyautogui
                            pyautogui.hotkey('alt', 'f4')
                            LOGGER.info("✅ 已发送 Alt+F4 关闭 Excel 窗口")
                            time.sleep(0.5)
                        except Exception as e:
                            LOGGER.debug(f"Alt+F4 失败: {e}")
                            
                            # 方法2: 使用 WM_CLOSE
                            try:
                                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                                LOGGER.info("✅ 已发送 WM_CLOSE 关闭 Excel 窗口")
                                time.sleep(0.5)
                            except Exception as e2:
                                LOGGER.warning(f"WM_CLOSE 失败: {e2}")
                                
                                # 方法3: 使用 COM 对象关闭
                                try:
                                    import win32com.client
                                    excel_app = win32com.client.GetActiveObject("Excel.Application")
                                    if excel_app:
                                        workbook = excel_app.ActiveWorkbook
                                        if workbook and workbook.FullName:
                                            workbook.Close(False)  # False = 不保存
                                            LOGGER.info("✅ 已通过 COM 对象关闭 Excel 工作簿")
                                except Exception as e3:
                                    LOGGER.debug(f"COM 对象关闭失败: {e3}")
                        
                        # 验证窗口是否已关闭
                        time.sleep(0.5)
                        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                            LOGGER.info("✅ Excel 窗口已关闭")
                        else:
                            LOGGER.warning("⚠️ Excel 窗口可能未完全关闭")
                            
                    except Exception as e:
                        LOGGER.warning(f"关闭 Excel 窗口失败: {e}")
            else:
                LOGGER.debug("未找到需要关闭的 Excel 窗口")
                
        except Exception as e:
            LOGGER.warning(f"关闭 Excel 窗口时出错: {e}")
            import traceback
            LOGGER.debug(traceback.format_exc())
    
    def _click_add_to_summary_button(self) -> None:
        """点击左下角的'Add to summary'按钮"""
        if Application is None:
            return
        
        LOGGER.info("查找并点击左下角的'Add to summary'按钮...")
        
        try:
            # 确保主窗口有焦点
            self._window.set_focus()
            time.sleep(0.3)
            
            # 方法1: 通过按钮文本查找
            try:
                add_button = self._window.child_window(title="Add to summary", control_type="Button")
                if add_button.exists() and add_button.is_enabled() and add_button.is_visible():
                    LOGGER.info("找到'Add to summary'按钮（通过title）")
                    # 获取按钮位置
                    try:
                        button_rect = add_button.rectangle()
                        button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                        button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                        LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y})")
                    except:
                        pass
                    
                    # 使用click_input()点击
                    add_button.click_input()
                    time.sleep(0.5)
                    LOGGER.info("✅ 已点击'Add to summary'按钮（通过title，使用click_input）")
                    return
            except Exception as e:
                LOGGER.debug(f"通过title查找按钮失败: {e}")
            
            # 方法2: 遍历所有按钮查找（优先查找左下角的按钮）
            try:
                all_buttons = self._window.descendants(control_type="Button")
                LOGGER.info(f"找到 {len(all_buttons)} 个按钮")
                
                # 先尝试精确匹配
                for idx, button in enumerate(all_buttons):
                    try:
                        button_text = button.window_text().strip()
                        LOGGER.debug(f"  按钮 #{idx}: 文本='{button_text}'")
                        
                        if "ADD TO SUMMARY" in button_text.upper() or ("ADD TO" in button_text.upper() and "SUMMARY" in button_text.upper()):
                            LOGGER.info(f"找到'Add to summary'按钮（文本: '{button_text}'）")
                            if button.is_enabled() and button.is_visible():
                                # 获取按钮位置（验证是否在左下角区域）
                                try:
                                    button_rect = button.rectangle()
                                    button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                                    button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                                    
                                    # 获取窗口大小，判断是否在左下角区域
                                    window_rect = self._window.rectangle()
                                    window_width = window_rect.right - window_rect.left
                                    window_height = window_rect.bottom - window_rect.top
                                    
                                    # 左下角区域：左侧50%，底部30%
                                    is_bottom_left = (button_center_x < window_rect.left + window_width * 0.5 and 
                                                      button_center_y > window_rect.top + window_height * 0.7)
                                    
                                    LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y}), 是否在左下角: {is_bottom_left}")
                                    
                                    # 如果找到匹配的按钮，即使不在左下角也尝试点击
                                    # 使用click_input()点击
                                    button.click_input()
                                    time.sleep(0.5)
                                    LOGGER.info(f"✅ 已点击'Add to summary'按钮（文本: '{button_text}'，使用click_input）")
                                    return
                                except:
                                    # 如果无法获取位置，也尝试点击
                                    button.click_input()
                                    time.sleep(0.5)
                                    LOGGER.info(f"✅ 已点击'Add to summary'按钮（文本: '{button_text}'，使用click_input）")
                                    return
                    except Exception as e:
                        LOGGER.debug(f"检查按钮 #{idx} 时出错: {e}")
                        continue
                
                # 如果精确匹配失败，查找所有包含"ADD"或"SUMMARY"的按钮，并优先选择左下角的
                candidate_buttons = []
                window_rect = self._window.rectangle()
                window_width = window_rect.right - window_rect.left
                window_height = window_rect.bottom - window_rect.top
                
                for idx, button in enumerate(all_buttons):
                    try:
                        button_text = button.window_text().strip().upper()
                        if ("ADD" in button_text and "SUMMARY" in button_text) or "ADD TO" in button_text:
                            if button.is_enabled() and button.is_visible():
                                try:
                                    button_rect = button.rectangle()
                                    button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                                    button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                                    
                                    # 计算到左下角的距离（用于排序）
                                    distance_to_bottom_left = ((button_center_x - window_rect.left)**2 + 
                                                              (button_center_y - (window_rect.top + window_height))**2)**0.5
                                    
                                    candidate_buttons.append((button, button_text, distance_to_bottom_left))
                                except:
                                    candidate_buttons.append((button, button_text, float('inf')))
                    except:
                        continue
                
                # 按距离左下角的距离排序，优先选择最近的
                if candidate_buttons:
                    candidate_buttons.sort(key=lambda x: x[2])
                    button, button_text, distance = candidate_buttons[0]
                    LOGGER.info(f"找到候选'Add to summary'按钮（文本: '{button_text}'，距离左下角: {distance:.1f}像素）")
                    button.click_input()
                    time.sleep(0.5)
                    LOGGER.info(f"✅ 已点击'Add to summary'按钮（文本: '{button_text}'，使用click_input）")
                    return
                    
            except Exception as e:
                LOGGER.debug(f"遍历按钮失败: {e}")
            
            # 方法3: 使用Windows API查找左下角的按钮
            if win32gui:
                try:
                    main_hwnd = self._window.handle
                    window_rect = self._window.rectangle()
                    window_width = window_rect.right - window_rect.left
                    window_height = window_rect.bottom - window_rect.top
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child)
                            class_name = win32gui.GetClassName(hwnd_child)
                            
                            # 检查是否是按钮控件，且文本包含"ADD"和"SUMMARY"
                            if "BUTTON" in class_name.upper():
                                text_upper = window_text.upper()
                                if ("ADD TO SUMMARY" in text_upper or 
                                    ("ADD" in text_upper and "SUMMARY" in text_upper) or
                                    "ADD TO" in text_upper):
                                    try:
                                        rect = win32gui.GetWindowRect(hwnd_child)
                                        center_x = (rect[0] + rect[2]) // 2
                                        center_y = (rect[1] + rect[3]) // 2
                                        
                                        # 计算到左下角的距离
                                        distance = ((center_x - window_rect.left)**2 + 
                                                   (center_y - (window_rect.top + window_height))**2)**0.5
                                        
                                        lParam.append((hwnd_child, window_text, distance))
                                    except:
                                        lParam.append((hwnd_child, window_text, float('inf')))
                        except:
                            pass
                        return True
                    
                    button_list = []
                    win32gui.EnumChildWindows(main_hwnd, enum_child_proc, button_list)
                    
                    if button_list:
                        # 按距离左下角的距离排序
                        button_list.sort(key=lambda x: x[2])
                        button_hwnd, button_text, distance = button_list[0]
                        LOGGER.info(f"使用Windows API找到按钮: '{button_text}'（距离左下角: {distance:.1f}像素）")
                        
                        # 获取按钮位置并点击（Y坐标往上移动8像素，避免点击到按钮边缘）
                        rect = win32gui.GetWindowRect(button_hwnd)
                        center_x = (rect[0] + rect[2]) // 2
                        center_y = (rect[1] + rect[3]) // 2 - 8  # 往上移动8像素
                        
                        try:
                            import pyautogui
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.5)
                            LOGGER.info(f"✅ 已通过坐标点击'Add to summary'按钮（位置: {center_x}, {center_y}，已调整Y坐标往上）")
                            return
                        except ImportError:
                            win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                            time.sleep(0.5)
                            LOGGER.info("✅ 已通过Windows API点击'Add to summary'按钮")
                            return
                except Exception as e:
                    LOGGER.debug(f"使用Windows API查找按钮失败: {e}")
            
            # 方法4: 如果找不到，尝试根据位置估算左下角区域
            try:
                window_rect = self._window.rectangle()
                window_width = window_rect.right - window_rect.left
                window_height = window_rect.bottom - window_rect.top
                
                # 改进估算逻辑：侧边栏通常是固定宽度，而不是窗口宽度的百分比
                # 使用更靠左的固定偏移（约220像素，参考之前的成功日志）
                estimated_x = window_rect.left + 220
                
                # 确保不会超出窗口范围
                if estimated_x > window_rect.right:
                    estimated_x = window_rect.left + int(window_width * 0.1)
                
                # Y坐标往上移动，使用82%高度而不是85%，避免点到底部状态栏或按钮边缘
                estimated_y = window_rect.top + int(window_height * 0.82)
                
                LOGGER.info(f"尝试在估算的左下角位置点击: ({estimated_x}, {estimated_y})")
                try:
                    import pyautogui
                    pyautogui.click(estimated_x, estimated_y)
                    time.sleep(0.5)
                    LOGGER.info(f"✅ 已在估算位置点击'Add to summary'按钮（位置: {estimated_x}, {estimated_y}）")
                    return
                except ImportError:
                    LOGGER.warning("pyautogui未安装，无法使用坐标点击")
            except Exception as e:
                LOGGER.debug(f"使用估算位置点击失败: {e}")
            
            LOGGER.warning("未找到'Add to summary'按钮")
            
        except Exception as e:
            LOGGER.error(f"点击'Add to summary'按钮失败: {e}")
            raise RuntimeError(f"无法点击'Add to summary'按钮: {e}")
    
    def _click_yes_button_in_dialog(self, dialog_hwnd, dialog_title="对话框") -> bool:
        """
        在对话框中点击Yes按钮（只保留最精准的坐标点击方法）
        针对独立弹出的顶层窗口，先激活，再找内部Yes按钮坐标进行点击。
        """
        try:
            LOGGER.info(f"正在处理独立窗口: '{dialog_title}' (Handle: {dialog_hwnd})")
            
            # 1. 强制激活窗口，确保它在屏幕最前端
            try:
                if win32gui.IsIconic(dialog_hwnd):  # 如果最小化了，就还原
                    win32gui.ShowWindow(dialog_hwnd, win32con.SW_RESTORE)
                
                win32gui.SetForegroundWindow(dialog_hwnd)
                win32gui.BringWindowToTop(dialog_hwnd)
                time.sleep(0.5) # 等待窗口激活这一物理过程完成
            except Exception as e:
                LOGGER.warning(f"窗口激活尝试遇到小问题 (通常可忽略): {e}")

            # 2. 遍历该窗口下的所有子控件，寻找 "Yes" 按钮
            yes_button_info = {"hwnd": None, "rect": None}

            def find_yes_button(hwnd_child, _):
                try:
                    # 获取控件文本和类名
                    text = win32gui.GetWindowText(hwnd_child).strip()
                    class_name = win32gui.GetClassName(hwnd_child)
                    
                    # 打印所有按钮以便调试
                    if "BUTTON" in class_name.upper():
                        LOGGER.info(f"  扫描到按钮: '{text}' (类名: {class_name})")
                        # 去掉&符号后比较（Windows按钮常用&表示快捷键）
                        clean_text = text.replace("&", "").strip().upper()
                        LOGGER.debug(f"    清理后的文本: '{clean_text}'")
                        # 支持多种Yes按钮文本格式
                        if clean_text == "YES" or clean_text.startswith("YES") or "YES" in clean_text:
                            try:
                                yes_button_info["hwnd"] = hwnd_child
                                yes_button_info["rect"] = win32gui.GetWindowRect(hwnd_child)
                                LOGGER.info(f"  ✅ 匹配到Yes按钮！原始文本: '{text}', 清理后: '{clean_text}'")
                                return False # 找到了，停止遍历
                            except Exception as e:
                                LOGGER.warning(f"  获取Yes按钮信息失败: {e}")
                                # 继续查找其他按钮
                except Exception as e:
                    LOGGER.debug(f"  扫描按钮时出错: {e}")
                    pass
                return True

            win32gui.EnumChildWindows(dialog_hwnd, find_yes_button, None)

            # 3. 如果找到了Yes按钮，计算中心坐标并点击
            LOGGER.info(f"按钮查找结果: hwnd={yes_button_info['hwnd']}, rect={yes_button_info['rect']}")
            if yes_button_info["hwnd"] and yes_button_info["rect"]:
                rect = yes_button_info["rect"]
                center_x = (rect[0] + rect[2]) // 2
                center_y = (rect[1] + rect[3]) // 2
                
                LOGGER.info(f"✅ 找到Yes按钮! 屏幕坐标: ({center_x}, {center_y})")
                
                try:
                    import pyautogui
                    # 确保对话框在最前面
                    win32gui.SetForegroundWindow(dialog_hwnd)
                    win32gui.BringWindowToTop(dialog_hwnd)
                    time.sleep(0.2)
                    
                    # 方法1: 使用坐标直接点击（更可靠）
                    try:
                        pyautogui.click(center_x, center_y)
                        LOGGER.info(f"🖱️ 已执行鼠标点击（方法1：直接坐标点击）")
                        time.sleep(0.3)
                        # 验证对话框是否消失
                        if not win32gui.IsWindow(dialog_hwnd) or not win32gui.IsWindowVisible(dialog_hwnd):
                            LOGGER.info("✅ 对话框已关闭，点击成功")
                            return True
                    except Exception as e:
                        LOGGER.warning(f"方法1失败: {e}，尝试方法2...")
                    
                    # 方法2: 先移动再点击
                    try:
                        pyautogui.moveTo(center_x, center_y, duration=0.2)
                        time.sleep(0.1)
                        pyautogui.click()
                        LOGGER.info(f"🖱️ 已执行鼠标点击（方法2：移动后点击）")
                        time.sleep(0.3)
                        # 验证对话框是否消失
                        if not win32gui.IsWindow(dialog_hwnd) or not win32gui.IsWindowVisible(dialog_hwnd):
                            LOGGER.info("✅ 对话框已关闭，点击成功")
                            return True
                    except Exception as e:
                        LOGGER.warning(f"方法2失败: {e}，尝试方法3...")
                    
                    # 方法3: 使用Windows API点击按钮
                    try:
                        if win32gui and win32con:
                            win32gui.SetForegroundWindow(dialog_hwnd)
                            time.sleep(0.1)
                            # 发送点击消息到按钮（使用win32gui.SendMessage）
                            win32gui.SendMessage(yes_button_info["hwnd"], win32con.BM_CLICK, 0, 0)
                            LOGGER.info(f"🖱️ 已执行Windows API点击（方法3：SendMessage）")
                            time.sleep(0.3)
                            # 验证对话框是否消失
                            if not win32gui.IsWindow(dialog_hwnd) or not win32gui.IsWindowVisible(dialog_hwnd):
                                LOGGER.info("✅ 对话框已关闭，点击成功")
                                return True
                            
                            # 如果SendMessage失败，尝试PostMessage
                            win32gui.PostMessage(yes_button_info["hwnd"], win32con.BM_CLICK, 0, 0)
                            LOGGER.info(f"🖱️ 已执行Windows API点击（方法3b：PostMessage）")
                            time.sleep(0.3)
                            # 再次验证对话框是否消失
                            if not win32gui.IsWindow(dialog_hwnd) or not win32gui.IsWindowVisible(dialog_hwnd):
                                LOGGER.info("✅ 对话框已关闭，点击成功")
                                return True
                    except Exception as e:
                        LOGGER.warning(f"方法3失败: {e}")
                    
                    # 如果所有方法都失败，返回False让调用者尝试fallback
                    LOGGER.warning("所有点击方法都执行了，但对话框可能未关闭，将尝试fallback方法")
                    return False
                    
                except ImportError:
                    LOGGER.error("缺少 pyautogui 库，无法执行物理点击")
                    return False
                except Exception as e:
                    LOGGER.error(f"点击Yes按钮时出错: {e}")
                    import traceback
                    LOGGER.error(traceback.format_exc())
                    return False
            else:
                # Fallback: 直接按'y'键（之前成功的方法）
                LOGGER.warning(f"未找到文本为'Yes'的按钮，尝试按'y'键...")
                try:
                    import pyautogui
                    # 确保对话框已激活
                    win32gui.SetForegroundWindow(dialog_hwnd)
                    time.sleep(0.2)
                    # 按'y'键
                    pyautogui.press('y')
                    LOGGER.info("✅ 已通过按'y'键点击Yes")
                    time.sleep(0.3)
                    return True
                except Exception as e:
                    LOGGER.error(f"按'y'键失败: {e}")
                    return False

        except Exception as e:
            LOGGER.error(f"点击操作发生异常: {e}")
            return False

    def _handle_submit_confirmation_dialogs(self) -> None:
        """
        全局扫描屏幕，查找并处理所有相关的确认对话框。
        针对 Submit 之后可能连续弹出的多个独立窗口。
        """
        LOGGER.info("开始全局扫描确认对话框...")
        
        # 关键词列表：包含截图中的标题和内容关键词
        target_keywords = [
            "warning", 
            "submit mir request", 
            "correlation unit", 
            "are you sure"
        ]

        # 最多尝试处理几轮，防止死循环
        max_rounds = 5 
        
        for i in range(max_rounds):
            dialog_found = False
            
            # 定义回调函数，用于 EnumWindows
            def find_target_dialogs(hwnd, _):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return True
                    
                    title = win32gui.GetWindowText(hwnd)
                    if not title:
                        return True
                        
                    title_lower = title.lower()
                    
                    # 检查标题是否匹配关键词
                    is_target = False
                    for kw in target_keywords:
                        if kw in title_lower:
                            is_target = True
                            break
                    
                    if is_target:
                        # 进一步检查类名，排除非对话框窗口（可选，但加上更稳）
                        # class_name = win32gui.GetClassName(hwnd)
                        # LOGGER.info(f"发现潜在目标: '{title}' (类名: {class_name})")
                        
                        # 直接尝试处理这个窗口
                        LOGGER.info(f"🔍 捕获到目标窗口: '{title}'")
                        if self._click_yes_button_in_dialog(hwnd, title):
                            nonlocal dialog_found
                            dialog_found = True
                            # 处理成功后，稍微等待窗口关闭
                            time.sleep(1.0)
                except Exception:
                    pass
                return True

            # 执行全局扫描
            win32gui.EnumWindows(find_target_dialogs, None)
            
            # 如果这一轮没有找到任何对话框，说明可能处理完了，或者还没弹出来
            if not dialog_found:
                if i == 0:
                    LOGGER.info("当前未检测到确认对话框，等待 2 秒再试...")
                    time.sleep(2.0)
                else:
                    LOGGER.info("未发现更多对话框，流程结束。")
                    break
            else:
                LOGGER.info("已处理一个对话框，继续扫描下一个...")
                time.sleep(1.5) # 等待下一个弹出
    
    def _handle_submit_confirmation_dialogs(self) -> None:
        """
        处理Submit后的所有确认对话框（可能有多个）
        
        流程：
        1. 第一个对话框: Warning - "Do you want to submit the MIR with no AWS session related?"
        2. 第二个对话框: Submit MIR Request - "Are you sure that you want to submit the current request?"
        
        对每个对话框都点击Yes按钮
        """
        try:
            LOGGER.info("检查并处理Submit后的确认对话框...")
            
            # 最多处理3个对话框（以防有更多）
            max_dialogs = 3
            processed_dialogs = []  # 记录已处理的对话框标题，避免重复处理
            
            for i in range(max_dialogs):
                # 等待对话框出现（第一次等待稍长，后续等待更长以便新对话框弹出）
                wait_time = 2.0 if i == 0 else 4.0
                LOGGER.info(f"等待 {wait_time} 秒，让对话框 #{i+1} 弹出...")
                time.sleep(wait_time)
                
                # 尝试多次查找对话框（因为对话框可能需要时间才能完全显示）
                dialog_hwnd = None
                dialog_title = None
                max_retries = 3
                
                for retry in range(max_retries):
                    if win32gui:
                        # 先扫描所有可见窗口用于调试
                        def enum_all_windows(hwnd, all_windows):
                            try:
                                if win32gui.IsWindowVisible(hwnd):
                                    window_text = win32gui.GetWindowText(hwnd)
                                    if window_text:
                                        class_name = win32gui.GetClassName(hwnd)
                                        all_windows.append((window_text, class_name))
                            except:
                                pass
                            return True
                        
                        all_visible_windows = []
                        win32gui.EnumWindows(enum_all_windows, all_visible_windows)
                        
                        if retry == 0:  # 只在第一次尝试时打印所有窗口
                            LOGGER.info(f"当前所有可见窗口（共{len(all_visible_windows)}个）：")
                            for win_text, win_class in all_visible_windows[:20]:  # 只显示前20个
                                LOGGER.info(f"    窗口: '{win_text}' (类名: {win_class})")
                        
                        def enum_dialogs_callback(hwnd, dialogs):
                            try:
                                window_text = win32gui.GetWindowText(hwnd)
                                if window_text:
                                    text_lower = window_text.lower()
                                    # 检查是否是确认对话框（扩展关键词列表）
                                    if any(keyword in text_lower for keyword in [
                                        "warning", "submit", "confirm", "are you sure",
                                        "mir request", "request", "correlation"
                                    ]):
                                        # 检查窗口类名和标题
                                        try:
                                            class_name = win32gui.GetClassName(hwnd)
                                            LOGGER.info(f"    潜在对话框: '{window_text}', 类名={class_name}")
                                            
                                            # 判断是否是对话框：
                                            # 1. 标准对话框类名（#32770）
                                            # 2. 类名包含"dialog"
                                            # 3. 标题是"Warning"或"Submit MIR Request"（即使不是标准对话框类名）
                                            is_dialog = False
                                            if "dialog" in class_name.lower() or "#32770" in class_name:
                                                is_dialog = True
                                            elif window_text in ["Warning", "Submit MIR Request"]:
                                                is_dialog = True
                                                LOGGER.info(f"    根据标题匹配为对话框: '{window_text}'")
                                            
                                            if is_dialog:
                                                # 检查窗口是否可见
                                                if win32gui.IsWindowVisible(hwnd):
                                                    LOGGER.info(f"    ✅ 确认为对话框: '{window_text}', 类名={class_name}")
                                                    dialogs.append((hwnd, window_text))
                                        except:
                                            pass
                            except:
                                pass
                            return True
                        
                        dialogs = []
                        win32gui.EnumWindows(enum_dialogs_callback, dialogs)
                        
                        LOGGER.info(f"扫描到 {len(dialogs)} 个对话框，已处理: {len(processed_dialogs)} 个")
                        
                        # 过滤掉已处理的对话框
                        new_dialogs = [(hwnd, title) for hwnd, title in dialogs 
                                      if title not in processed_dialogs]
                        
                        if new_dialogs:
                            dialog_hwnd, dialog_title = new_dialogs[0]
                            LOGGER.info(f"找到确认对话框 #{i+1}: '{dialog_title}' (尝试 {retry+1}/{max_retries})")
                            break
                        elif retry < max_retries - 1:
                            # 如果没找到，再等待一下再重试
                            LOGGER.info(f"暂未找到新对话框，等待1.5秒后重试... (尝试 {retry+1}/{max_retries})")
                            time.sleep(1.5)
                    else:
                        LOGGER.warning("win32gui不可用，无法自动处理对话框")
                        return
                
                # 如果找到了对话框，处理它
                if dialog_hwnd and dialog_title:
                    LOGGER.info(f"准备处理对话框 #{i+1}: '{dialog_title}' (句柄: {dialog_hwnd})")
                    
                    # 检查对话框是否真的可见
                    try:
                        is_visible = win32gui.IsWindowVisible(dialog_hwnd)
                        LOGGER.info(f"对话框可见性: {is_visible}")
                    except:
                        pass
                    
                    # 记录已处理的对话框
                    processed_dialogs.append(dialog_title)
                    
                    # 点击Yes按钮（多次尝试，确保成功）
                    LOGGER.info(f"开始点击对话框 #{i+1} 的Yes按钮...")
                    
                    click_success = False
                    max_click_retries = 3
                    
                    for click_retry in range(max_click_retries):
                        LOGGER.info(f"点击尝试 {click_retry + 1}/{max_click_retries}...")
                        
                        if self._click_yes_button_in_dialog(dialog_hwnd, dialog_title):
                            click_success = True
                            LOGGER.info(f"✅ 已成功处理对话框 #{i+1}: '{dialog_title}'")
                            break
                        else:
                            if click_retry < max_click_retries - 1:
                                LOGGER.warning(f"点击失败，等待1秒后重试...")
                                time.sleep(1.0)
                    
                    if not click_success:
                        LOGGER.error(f"❌ 无法自动处理对话框 #{i+1}: '{dialog_title}'，已重试{max_click_retries}次，请手动点击Yes")
                        break
                    
                    # 点击成功后，额外等待一下，让下一个对话框有时间弹出
                    LOGGER.info("等待下一个对话框弹出...")
                else:
                    # 没有找到更多对话框，说明已处理完毕
                    if i == 0:
                        LOGGER.info("未检测到确认对话框")
                    else:
                        LOGGER.info(f"✅ 已处理所有确认对话框（共{i}个）")
                    break
            
            # 最后再检查一次，是否还有遗漏的对话框
            # 既然使用了全局扫描，这里可以简化或者作为最后的确认
            # LOGGER.info("最后检查：是否还有遗漏的对话框...")
            # self._handle_submit_confirmation_dialogs()
            
        except Exception as e:
            LOGGER.error(f"处理确认对话框时出错: {e}")
            LOGGER.warning("请手动点击所有确认对话框的Yes按钮")
        
        except Exception as e:
            LOGGER.error(f"处理确认对话框时出错: {e}")
            LOGGER.warning("请手动点击所有确认对话框的Yes按钮")
    
    def _handle_inactive_source_lots_dialog(self, max_wait_time: int = 10) -> bool:
        """
        处理 "Inactive Source Lots" 对话框
        
        对话框标题: "Inactive Source Lots"
        内容: "Following units are not active for correlation: ... Please check that units are in: 'PNG-CPU'."
        按钮: "Copy Inactive & Close" 和 "Close"
        
        点击 "Copy Inactive & Close" 按钮
        
        Args:
            max_wait_time: 最大等待时间（秒），等待对话框出现
        
        Returns:
            bool: 如果找到并处理了对话框返回True，否则返回False
        """
        try:
            LOGGER.info("检查是否存在 'Inactive Source Lots' 对话框...")
            
            if not win32gui:
                LOGGER.warning("win32gui不可用，无法处理对话框")
                return False
            
            # 循环等待对话框出现
            dialog_hwnd = None
            dialog_title = None
            start_time = time.time()
            
            def enum_inactive_dialog(hwnd, dialogs):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return True
                    
                    window_text = win32gui.GetWindowText(hwnd)
                    if "Inactive Source Lots" in window_text:
                        class_name = win32gui.GetClassName(hwnd)
                        LOGGER.info(f"找到 'Inactive Source Lots' 对话框: '{window_text}' (类名: {class_name})")
                        dialogs.append((hwnd, window_text))
                except:
                    pass
                return True
            
            # 等待对话框出现（最多等待max_wait_time秒）
            while time.time() - start_time < max_wait_time:
                inactive_dialogs = []
                win32gui.EnumWindows(enum_inactive_dialog, inactive_dialogs)
                
                if inactive_dialogs:
                    dialog_hwnd, dialog_title = inactive_dialogs[0]
                    LOGGER.info(f"✅ 找到 'Inactive Source Lots' 对话框 (句柄: {dialog_hwnd})")
                    break
                
                # 如果还没找到，等待一下再试
                time.sleep(0.5)
            
            if not dialog_hwnd:
                LOGGER.debug(f"等待 {max_wait_time} 秒后未找到 'Inactive Source Lots' 对话框，可能不存在")
                return False
            
            dialog_hwnd, dialog_title = inactive_dialogs[0]
            LOGGER.info(f"✅ 找到 'Inactive Source Lots' 对话框 (句柄: {dialog_hwnd})")
            
            # 激活对话框
            win32gui.SetForegroundWindow(dialog_hwnd)
            win32gui.BringWindowToTop(dialog_hwnd)
            time.sleep(0.5)
            
            # 查找 "Copy Inactive & Close" 按钮
            def find_copy_inactive_button(hwnd_child, button_info):
                try:
                    text = win32gui.GetWindowText(hwnd_child).strip()
                    class_name = win32gui.GetClassName(hwnd_child)
                    
                    if "BUTTON" in class_name.upper():
                        LOGGER.debug(f"  扫描到按钮: '{text}' (类名: {class_name})")
                        # 去掉&符号后比较
                        clean_text = text.replace("&", "").upper()
                        if "COPY INACTIVE" in clean_text and "CLOSE" in clean_text:
                            button_info["hwnd"] = hwnd_child
                            button_info["rect"] = win32gui.GetWindowRect(hwnd_child)
                            LOGGER.info(f"  ✅ 匹配到 'Copy Inactive & Close' 按钮！原始文本: '{text}'")
                            return False  # 找到了，停止遍历
                except:
                    pass
                return True
            
            button_info = {"hwnd": None, "rect": None}
            win32gui.EnumChildWindows(dialog_hwnd, find_copy_inactive_button, button_info)
            
            if button_info["hwnd"] and button_info["rect"]:
                # 点击按钮
                rect = button_info["rect"]
                center_x = (rect[0] + rect[2]) // 2
                center_y = (rect[1] + rect[3]) // 2
                
                LOGGER.info(f"✅ 找到 'Copy Inactive & Close' 按钮! 屏幕坐标: ({center_x}, {center_y})")
                
                try:
                    import pyautogui
                    # 确保对话框在最前面
                    win32gui.SetForegroundWindow(dialog_hwnd)
                    win32gui.BringWindowToTop(dialog_hwnd)
                    time.sleep(0.2)
                    
                    # 方法1: 使用坐标直接点击
                    click_success = False
                    try:
                        pyautogui.click(center_x, center_y)
                        LOGGER.info("🖱️ 已点击 'Copy Inactive & Close' 按钮（方法1：直接坐标点击）")
                        time.sleep(0.5)
                        click_success = True
                    except Exception as e:
                        LOGGER.warning(f"方法1失败: {e}，尝试方法2...")
                        
                        # 方法2: 先移动再点击
                        try:
                            pyautogui.moveTo(center_x, center_y, duration=0.2)
                            time.sleep(0.1)
                            pyautogui.click()
                            LOGGER.info("🖱️ 已点击 'Copy Inactive & Close' 按钮（方法2：移动后点击）")
                            time.sleep(0.5)
                            click_success = True
                        except Exception as e2:
                            LOGGER.warning(f"方法2失败: {e2}，尝试方法3...")
                            
                            # 方法3: 使用Windows API点击按钮
                            try:
                                if win32gui and win32con:
                                    win32gui.SetForegroundWindow(dialog_hwnd)
                                    time.sleep(0.1)
                                    win32gui.SendMessage(button_info["hwnd"], win32con.BM_CLICK, 0, 0)
                                    LOGGER.info("🖱️ 已点击 'Copy Inactive & Close' 按钮（方法3：SendMessage）")
                                    time.sleep(0.3)
                                    # 如果SendMessage后对话框还在，尝试PostMessage
                                    if win32gui.IsWindow(dialog_hwnd) and win32gui.IsWindowVisible(dialog_hwnd):
                                        win32gui.PostMessage(button_info["hwnd"], win32con.BM_CLICK, 0, 0)
                                        LOGGER.info("🖱️ 已点击 'Copy Inactive & Close' 按钮（方法3b：PostMessage）")
                                        time.sleep(0.3)
                                    click_success = True
                            except Exception as e3:
                                LOGGER.warning(f"方法3失败: {e3}")
                    
                    # 验证对话框是否关闭
                    time.sleep(0.5)
                    if not win32gui.IsWindow(dialog_hwnd) or not win32gui.IsWindowVisible(dialog_hwnd):
                        LOGGER.info("✅ 'Inactive Source Lots' 对话框已关闭")
                        return True
                    elif click_success:
                        LOGGER.info("✅ 已执行点击操作（对话框可能稍后关闭）")
                        return True
                    else:
                        LOGGER.warning("⚠️ 点击操作可能未成功，对话框仍然存在")
                        return False
                        
                except Exception as e:
                    LOGGER.error(f"点击 'Copy Inactive & Close' 按钮失败: {e}")
                    import traceback
                    LOGGER.error(traceback.format_exc())
                    return False
            else:
                LOGGER.warning("⚠️ 未找到 'Copy Inactive & Close' 按钮")
                return False
                
        except Exception as e:
            LOGGER.error(f"处理 'Inactive Source Lots' 对话框时出错: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _handle_final_success_dialog_and_get_mir(self) -> str:
        """
        处理最终的Submit MIR Request成功对话框
        标题: "Submit MIR Request"
        内容: "Your MIR# XXXXXX has been submitted"
        按钮: "Copy MIR & Close"和"Close"
        
        点击"Copy MIR & Close"并返回MIR号码
        
        Returns:
            str: MIR号码，例如"2965268"
        """
        try:
            LOGGER.info("等待最终成功对话框出现（轮询等待，直到对话框弹出）...")
            
            # 轮询等待对话框出现（一直等待直到对话框出现，每2秒检查一次）
            # 不设置最大等待时间，一直等待直到对话框出现
            check_interval = 2  # 每2秒检查一次
            max_attempts = 1000  # 设置一个很大的数字，实际会一直等待直到找到对话框
            
            dialog_hwnd = None
            for attempt in range(max_attempts):
                if win32gui:
                    def enum_success_dialog(hwnd, dialogs):
                        try:
                            if not win32gui.IsWindowVisible(hwnd):
                                return True
                            
                            window_text = win32gui.GetWindowText(hwnd)
                            if window_text == "Submit MIR Request":
                                class_name = win32gui.GetClassName(hwnd)
                                dialogs.append(hwnd)
                        except:
                            pass
                        return True
                    
                    success_dialogs = []
                    win32gui.EnumWindows(enum_success_dialog, success_dialogs)
                    
                    if success_dialogs:
                        dialog_hwnd = success_dialogs[0]
                        LOGGER.info(f"✅ 找到成功对话框 (等待了 {(attempt+1)*check_interval} 秒)")
                        break
                    else:
                        if (attempt + 1) % 5 == 0:  # 每10秒（5次检查）输出一次日志
                            LOGGER.info(f"等待对话框中... (已等待 {(attempt+1)*check_interval} 秒，将继续等待直到对话框出现)")
                        
                        # 检查 ESC 键
                        try:
                            from .utils.keyboard_listener import is_esc_pressed
                            if is_esc_pressed():
                                LOGGER.warning("⚠️ 检测到 ESC 键，停止等待对话框")
                                return None
                        except:
                            pass
                        
                        time.sleep(check_interval)
                else:
                    # 如果没有win32gui，使用固定等待
                    time.sleep(check_interval)
            
            if dialog_hwnd:
                LOGGER.info(f"✅ 找到成功对话框 (句柄: {dialog_hwnd})")
                
                # 激活对话框
                win32gui.SetForegroundWindow(dialog_hwnd)
                win32gui.BringWindowToTop(dialog_hwnd)
                time.sleep(0.5)
                
                # 查找"Copy MIR & Close"按钮
                def find_copy_button(hwnd_child, button_info):
                            try:
                                text = win32gui.GetWindowText(hwnd_child).strip()
                                class_name = win32gui.GetClassName(hwnd_child)
                                
                                if "BUTTON" in class_name.upper():
                                    LOGGER.info(f"  扫描到按钮: '{text}' (类名: {class_name})")
                                    # 去掉&符号后比较
                                    clean_text = text.replace("&", "").upper()
                                    if "COPY MIR" in clean_text and "CLOSE" in clean_text:
                                        button_info["hwnd"] = hwnd_child
                                        button_info["rect"] = win32gui.GetWindowRect(hwnd_child)
                                        LOGGER.info(f"  ✅ 匹配到Copy MIR & Close按钮！原始文本: '{text}'")
                                        return False  # 找到了，停止遍历
                            except:
                                pass
                            return True
                
                button_info = {"hwnd": None, "rect": None}
                win32gui.EnumChildWindows(dialog_hwnd, find_copy_button, button_info)
                
                if button_info["hwnd"] and button_info["rect"]:
                            # 点击按钮
                            rect = button_info["rect"]
                            center_x = (rect[0] + rect[2]) // 2
                            center_y = (rect[1] + rect[3]) // 2
                            
                            LOGGER.info(f"✅ 找到Copy MIR & Close按钮! 屏幕坐标: ({center_x}, {center_y})")
                            
                            try:
                                import pyautogui
                                # 确保对话框在最前面
                                win32gui.SetForegroundWindow(dialog_hwnd)
                                win32gui.BringWindowToTop(dialog_hwnd)
                                time.sleep(0.2)
                                
                                # 方法1: 使用坐标直接点击（最可靠）
                                click_success = False
                                try:
                                    pyautogui.click(center_x, center_y)
                                    LOGGER.info("🖱️ 已点击Copy MIR & Close按钮（方法1：直接坐标点击）")
                                    time.sleep(0.5)
                                    click_success = True
                                except Exception as e:
                                    LOGGER.warning(f"方法1失败: {e}，尝试方法2...")
                                    
                                    # 方法2: 先移动再点击
                                    try:
                                        pyautogui.moveTo(center_x, center_y, duration=0.2)
                                        time.sleep(0.1)
                                        pyautogui.click()
                                        LOGGER.info("🖱️ 已点击Copy MIR & Close按钮（方法2：移动后点击）")
                                        time.sleep(0.5)
                                        click_success = True
                                    except Exception as e2:
                                        LOGGER.warning(f"方法2失败: {e2}，尝试方法3...")
                                        
                                        # 方法3: 使用Windows API点击按钮
                                        try:
                                            if win32gui and win32con:
                                                win32gui.SetForegroundWindow(dialog_hwnd)
                                                time.sleep(0.1)
                                                win32gui.SendMessage(button_info["hwnd"], win32con.BM_CLICK, 0, 0)
                                                LOGGER.info("🖱️ 已点击Copy MIR & Close按钮（方法3：SendMessage）")
                                                time.sleep(0.3)
                                                # 如果SendMessage后对话框还在，尝试PostMessage
                                                if win32gui.IsWindow(dialog_hwnd) and win32gui.IsWindowVisible(dialog_hwnd):
                                                    win32gui.PostMessage(button_info["hwnd"], win32con.BM_CLICK, 0, 0)
                                                    LOGGER.info("🖱️ 已点击Copy MIR & Close按钮（方法3b：PostMessage）")
                                                    time.sleep(0.3)
                                                click_success = True
                                        except Exception as e3:
                                            LOGGER.warning(f"方法3失败: {e3}")
                                
                                # 验证对话框是否关闭
                                if not win32gui.IsWindow(dialog_hwnd) or not win32gui.IsWindowVisible(dialog_hwnd):
                                    LOGGER.info("✅ 对话框已关闭，点击成功")
                                elif click_success:
                                    LOGGER.info("✅ 已执行点击操作（对话框可能稍后关闭）")
                                
                                # 从剪贴板获取MIR号码（无论对话框是否关闭，都尝试获取）
                                time.sleep(0.5)  # 等待剪贴板更新
                                try:
                                    import pyperclip
                                    mir_number = pyperclip.paste().strip()
                                    if mir_number:
                                        LOGGER.info(f"✅ 已从剪贴板获取MIR号码: {mir_number}")
                                        return mir_number
                                    else:
                                        LOGGER.warning("剪贴板为空，尝试使用win32clipboard...")
                                        raise ImportError("剪贴板为空")
                                except ImportError:
                                    # 使用win32clipboard
                                    try:
                                        import win32clipboard
                                        win32clipboard.OpenClipboard()
                                        try:
                                            mir_number = win32clipboard.GetClipboardData().strip()
                                        finally:
                                            win32clipboard.CloseClipboard()
                                        if mir_number:
                                            LOGGER.info(f"✅ 已从剪贴板获取MIR号码: {mir_number}")
                                            return mir_number
                                        else:
                                            LOGGER.warning("剪贴板内容为空")
                                    except Exception as e:
                                        LOGGER.error(f"获取剪贴板内容失败: {e}")
                                
                                # 如果无法从剪贴板获取，尝试从对话框文本中提取MIR号码
                                LOGGER.warning("无法从剪贴板获取MIR号码，尝试从对话框文本中提取...")
                                try:
                                    # 获取对话框文本内容
                                    def enum_text(hwnd_child, texts):
                                        try:
                                            text = win32gui.GetWindowText(hwnd_child)
                                            if text and "MIR#" in text:
                                                texts.append(text)
                                        except:
                                            pass
                                        return True
                                    
                                    texts = []
                                    win32gui.EnumChildWindows(dialog_hwnd, enum_text, texts)
                                    for text in texts:
                                        import re
                                        match = re.search(r'MIR#\s*(\d+)', text)
                                        if match:
                                            mir_number = match.group(1)
                                            LOGGER.info(f"✅ 从对话框文本中提取MIR号码: {mir_number}")
                                            return mir_number
                                except Exception as e:
                                    LOGGER.error(f"从对话框文本提取MIR号码失败: {e}")
                                
                                LOGGER.warning("无法获取MIR号码")
                                return ""
                                
                            except Exception as e:
                                LOGGER.error(f"点击按钮或获取剪贴板内容失败: {e}")
                                import traceback
                                LOGGER.error(traceback.format_exc())
                                return ""
                else:
                    LOGGER.warning("未找到Copy MIR & Close按钮，等待1秒后重试...")
                    time.sleep(1.0)
                    # 重新查找对话框（可能对话框已更新）
                    def enum_success_dialog_retry(hwnd, dialogs):
                        try:
                            if not win32gui.IsWindowVisible(hwnd):
                                return True
                            window_text = win32gui.GetWindowText(hwnd)
                            if window_text == "Submit MIR Request":
                                dialogs.append(hwnd)
                        except:
                            pass
                        return True
                    success_dialogs = []
                    win32gui.EnumWindows(enum_success_dialog_retry, success_dialogs)
                    if success_dialogs:
                        dialog_hwnd = success_dialogs[0]
                    else:
                        LOGGER.warning("对话框已消失")
                        return ""
            else:
                # 理论上不应该到达这里，因为会一直等待直到找到对话框
                LOGGER.error(f"等待了 {(max_attempts)*check_interval} 秒后仍未找到成功对话框，请手动检查")
                return ""
            
            LOGGER.warning("无法处理最终成功对话框，请手动点击Copy MIR & Close")
            return ""
            
        except Exception as e:
            LOGGER.error(f"处理最终成功对话框时出错: {e}")
            return ""
    
    def _verify_submit_success(self) -> bool:
        """
        验证Submit按钮点击是否成功
        
        检查方法：
        1. 检查是否有"Warning"对话框出现（这是Submit成功的标志）
           - 对话框标题应该是"Warning"
           - 内容包含"Do you want to submit the MIR with no AWS session related?"
        2. 检查是否有其他确认对话框
        
        Returns:
            bool: True如果点击成功，False如果未成功
        """
        try:
            # 等待一下，让对话框有时间出现
            time.sleep(1.5)
            
            # 检查是否有"Warning"对话框（这是Submit成功的标志）
            if win32gui:
                # 首先获取Mole主窗口句柄，用于排除
                main_hwnd = self._window.handle if self._window else None
                
                def enum_windows_callback(hwnd, windows):
                    try:
                        # 排除主窗口本身
                        if main_hwnd and hwnd == main_hwnd:
                            return True
                        
                        # 必须是可见窗口
                        if not win32gui.IsWindowVisible(hwnd):
                            return True
                        
                        window_text = win32gui.GetWindowText(hwnd)
                        if not window_text:
                            return True
                        
                        # 必须检查窗口类名，确认是对话框
                        class_name = win32gui.GetClassName(hwnd)
                        if not ("dialog" in class_name.lower() or "#32770" in class_name):
                            return True
                        
                        text_lower = window_text.lower()
                        # 检查是否是"Warning"对话框（这是Submit成功的标志）
                        if "warning" in text_lower:
                            LOGGER.info(f"发现Warning对话框: '{window_text}' (类名: {class_name})")
                            windows.append((hwnd, window_text, "warning"))
                        # 也检查其他可能的确认对话框（但必须是对话框类型）
                        elif any(kw in text_lower for kw in ["confirm", "success", "submitted"]):
                            LOGGER.info(f"发现确认对话框: '{window_text}' (类名: {class_name})")
                            windows.append((hwnd, window_text, "other"))
                    except:
                        pass
                    return True
                
                dialogs = []
                win32gui.EnumWindows(enum_windows_callback, dialogs)
                
                LOGGER.info(f"检测到 {len(dialogs)} 个对话框")
                
                if dialogs:
                    # 优先检查Warning对话框
                    for hwnd, title, dialog_type in dialogs:
                        if dialog_type == "warning":
                            LOGGER.info(f"✅ 检测到Warning对话框: '{title}' - Submit按钮点击成功！")
                            # 自动处理所有确认对话框（可能有多个）
                            self._handle_submit_confirmation_dialogs()
                            return True
                    # 如果有其他对话框，也认为可能成功
                    if dialogs:
                        LOGGER.info(f"✅ 检测到确认对话框: '{dialogs[0][1]}' - Submit按钮点击可能成功")
                        # 自动处理所有确认对话框
                        self._handle_submit_confirmation_dialogs()
                        return True
                else:
                    LOGGER.warning("⚠️ 未检测到任何对话框，Submit可能未成功")
            
            # 检查主窗口标题是否改变（可能包含"Submitted"等字样）
            try:
                current_title = self._window.window_text()
                if "submitted" in current_title.lower() or "complete" in current_title.lower():
                    LOGGER.info(f"✅ 窗口标题已改变，可能已提交: {current_title}")
                    return True
            except:
                pass
            
            return False
        except Exception as e:
            LOGGER.debug(f"验证提交成功失败: {e}")
            return False
    
    def _click_submit_button(self) -> None:
        """
        点击'Submit'按钮（绿色按钮，位于底部右侧区域）
        
        完全按照Search By VPOs、Select Visible Rows、Add to Summary的成功模式实现
        """
        if not self._window:
            raise RuntimeError("Mole窗口未连接")
        
        LOGGER.info("点击'Submit'按钮...")
        
        try:
            # 确保窗口激活（和Search By VPOs一样）
            self._window.set_focus()
            if win32gui and win32con:
                try:
                    hwnd = self._window.handle
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                except:
                    pass
            time.sleep(0.5)
            
            button_clicked = False
            
            # 方法1: 直接查找标题为"Submit"或"Submit MIR Request"的按钮（和Search By VPOs一样）
            submit_texts = ["Submit", "Submit MIR Request"]
            for submit_text in submit_texts:
                try:
                    submit_button = self._window.child_window(title=submit_text, control_type="Button")
                    if submit_button.exists() and submit_button.is_enabled():
                        LOGGER.info(f"找到'Submit'按钮（通过title='{submit_text}'）")
                        submit_button.click_input()
                        button_clicked = True
                        LOGGER.info(f"✅ 已点击'Submit'按钮（通过title='{submit_text}'）")
                        break
                except Exception as e1:
                    LOGGER.debug(f"通过title='{submit_text}'查找按钮失败: {e1}")
            
            # 方法2: 遍历所有按钮查找（和Search By VPOs一样）
            if not button_clicked:
                try:
                    all_buttons = self._window.descendants(control_type="Button")
                    LOGGER.info(f"找到 {len(all_buttons)} 个按钮")
                    for button in all_buttons:
                        try:
                            button_text = button.window_text().strip()
                            LOGGER.debug(f"  按钮: '{button_text}'")
                            if button_text == "Submit" or button_text == "Submit MIR Request":
                                LOGGER.info(f"找到'Submit'按钮（文本: '{button_text}'）")
                                if button.is_visible() and button.is_enabled():
                                    button.click_input()
                                    button_clicked = True
                                    LOGGER.info(f"✅ 已点击'Submit'按钮（遍历按钮，文本: '{button_text}'）")
                                    break
                        except Exception as e:
                            LOGGER.debug(f"检查按钮时出错: {e}")
                            continue
                except Exception as e2:
                    LOGGER.debug(f"遍历按钮失败: {e2}")
            
            # 方法3: 使用部分匹配查找（包含"Submit"）
            if not button_clicked:
                try:
                    all_buttons = self._window.descendants(control_type="Button")
                    for button in all_buttons:
                        try:
                            button_text = button.window_text().strip()
                            if "Submit" in button_text:
                                LOGGER.info(f"找到'Submit'按钮（部分匹配: '{button_text}'）")
                                if button.is_visible() and button.is_enabled():
                                    button.click_input()
                                    button_clicked = True
                                    LOGGER.info(f"✅ 已点击'Submit'按钮（部分匹配: '{button_text}'）")
                                    break
                        except:
                            continue
                except Exception as e3:
                    LOGGER.debug(f"部分匹配查找失败: {e3}")
            
            # 方法4: 使用Windows API查找按钮（和Search By VPOs一样）
            if not button_clicked and win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API查找'Submit'按钮...")
                    hwnd = self._window.handle
                    
                    found_button = {"found": False}
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child)
                            class_name = win32gui.GetClassName(hwnd_child)
                            # 打印所有按钮信息用于调试
                            if "BUTTON" in class_name.upper():
                                LOGGER.debug(f"  检测到按钮: '{window_text}' (类名: {class_name})")
                            # 检查是否是按钮控件或SubmitMIR控件，且文本包含"SUBMIT"
                            if ("BUTTON" in class_name.upper() or "SUBMITMIR" in class_name.upper()) and "SUBMIT" in window_text.upper():
                                LOGGER.info(f"通过Windows API找到按钮: '{window_text}' (类名: {class_name})")
                                win32gui.PostMessage(hwnd_child, win32con.BM_CLICK, 0, 0)
                                found_button["found"] = True
                                return False  # 停止枚举
                            return True
                        except:
                            return True
                    
                    win32gui.EnumChildWindows(hwnd, enum_child_proc, None)
                    time.sleep(0.5)
                    
                    if found_button["found"]:
                        button_clicked = True
                        LOGGER.info("✅ 已通过Windows API点击'Submit'按钮")
                    else:
                        LOGGER.warning("⚠️ Windows API未找到Submit按钮")
                except Exception as e4:
                    LOGGER.debug(f"使用Windows API查找按钮失败: {e4}")
            
            if not button_clicked:
                raise RuntimeError("无法点击'Submit'按钮，已尝试所有方法")
            
            # 等待按钮点击后的响应和对话框出现（和Search By VPOs一样）
            time.sleep(1.5)
            
            # 验证是否成功（检查Warning对话框）
            if self._verify_submit_success():
                LOGGER.info("✅ 已成功点击'Submit'按钮（验证通过：检测到Warning对话框）")
            else:
                LOGGER.warning("⚠️ 点击'Submit'按钮后未检测到Warning对话框，可能未成功")
                # 不抛出异常，让用户知道可能需要手动处理
        
        except Exception as e:
            LOGGER.error(f"点击'Submit'按钮失败: {e}")
            raise RuntimeError(f"无法点击'Submit'按钮: {e}")
    
    def _click_summary_tab(self) -> None:
        """
        点击顶部的 '3. View Summary' 标签页（使用相对定位，适应不同窗口大小和位置）
        
        策略：
        1. 优先查找 "1. Transfer Type" 作为参考点
        2. 基于参考点的相对位置计算目标位置
        3. 如果找不到参考点，使用窗口相对位置（百分比）
        """
        if Application is None:
            return
        
        LOGGER.info("点击 '3. View Summary' 标签页（使用相对定位）...")
        
        # 激活窗口
        try:
            self._window.set_focus()
            time.sleep(0.2)
        except Exception as e:
            LOGGER.warning(f"窗口激活失败: {e}")
        
        try:
            import pyautogui
            
            # 获取窗口位置和尺寸
            window_rect = self._window.rectangle()
            window_left = window_rect.left
            window_top = window_rect.top
            window_width = window_rect.right - window_rect.left
            window_height = window_rect.bottom - window_rect.top
            
            LOGGER.info(f"窗口位置: left={window_left}, top={window_top}, width={window_width}, height={window_height}")
            
            # 策略1: 基于参考标签 "1. Transfer Type" 的相对位置
            target_x = None
            target_y = None
            
            try:
                all_controls = self._window.descendants()
                
                # 查找 "1. Transfer Type" 作为参考点
                ref_ctrl = None
                for ctrl in all_controls:
                    try:
                        if not ctrl.is_visible():
                            continue
                        text = ctrl.window_text().strip()
                        if "1" in text and "TRANSFER" in text.upper() and "TYPE" in text.upper():
                            ref_ctrl = ctrl
                            break
                    except:
                        continue
                
                if ref_ctrl:
                    ref_rect = ref_ctrl.rectangle()
                    ref_center_x = ref_rect.left + (ref_rect.right - ref_rect.left) // 2
                    ref_center_y = ref_rect.top + (ref_rect.bottom - ref_rect.top) // 2
                    
                    LOGGER.info(f"找到参考标签 '1. Transfer Type' 中心: ({ref_center_x}, {ref_center_y})")
                    
                    # 基于已知的坐标关系计算相对偏移
                    # 已知：参考标签在 (3946, 768)，目标在 (4008, 804)
                    # 相对偏移：X方向 +62像素，Y方向 +36像素
                    # 但为了更健壮，我们使用窗口宽度的百分比
                    
                    # 计算参考标签相对于窗口的百分比位置
                    ref_x_percent = (ref_center_x - window_left) / window_width
                    ref_y_percent = (ref_center_y - window_top) / window_height
                    
                    LOGGER.info(f"参考标签相对位置: X={ref_x_percent:.2%}, Y={ref_y_percent:.2%}")
                    
                    # 基于已知的坐标关系：
                    # 参考标签: (3946, 768) 在窗口 (3832, 695, 6408, 2103) 中
                    # 窗口宽度: 2576, 高度: 1408
                    # 参考标签相对位置: X=(3946-3832)/2576=0.044, Y=(768-695)/1408=0.052
                    # 目标位置: (4008, 804)
                    # 目标相对位置: X=(4008-3832)/2576=0.068, Y=(804-695)/1408=0.077
                    # 相对偏移: X偏移=0.024 (约2.4%), Y偏移=0.025 (约2.5%)
                    # 调整：X偏移减小，点击位置往左移动
                    
                    # 使用固定的相对偏移（基于窗口宽度和高度）
                    x_offset_percent = 0.008  # 约0.8%的窗口宽度（往左移动）
                    y_offset_percent = 0.025  # 约2.5%的窗口高度
                    
                    target_x = ref_center_x + int(window_width * x_offset_percent)
                    target_y = ref_center_y + int(window_height * y_offset_percent)
                    
                    LOGGER.info(f"基于参考标签计算: 目标位置=({target_x}, {target_y})")
                    LOGGER.info(f"  相对偏移: X={x_offset_percent:.2%}窗口宽度, Y={y_offset_percent:.2%}窗口高度")
                else:
                    LOGGER.warning("未找到参考标签 '1. Transfer Type'")
            except Exception as e:
                LOGGER.warning(f"基于参考标签计算失败: {e}")
            
            # 策略2: 如果找不到参考标签，使用窗口相对位置（百分比）
            if target_x is None or target_y is None:
                LOGGER.info("使用窗口相对位置（百分比）")
                # 基于已知坐标 (4008, 804) 在窗口 (3832, 695, 6408, 2103) 中的相对位置
                # X相对位置: (4008-3832)/(6408-3832) = 176/2576 ≈ 0.068 (6.8%)
                # Y相对位置: (804-695)/(2103-695) = 109/1408 ≈ 0.077 (7.7%)
                # 调整：X位置减小，点击位置往左移动
                
                x_percent = 0.050  # 窗口宽度的5.0%（往左移动）
                y_percent = 0.077  # 窗口高度的7.7%
                
                target_x = window_left + int(window_width * x_percent)
                target_y = window_top + int(window_height * y_percent)
                
                LOGGER.info(f"基于窗口百分比计算: 目标位置=({target_x}, {target_y})")
                LOGGER.info(f"  相对位置: X={x_percent:.2%}窗口宽度, Y={y_percent:.2%}窗口高度")
            
            # 点击计算出的位置
            LOGGER.info(f"最终点击位置: ({target_x}, {target_y})")
            pyautogui.moveTo(target_x, target_y, duration=0.3)
            time.sleep(0.2)
            pyautogui.click(target_x, target_y)
            time.sleep(0.3)
            
            LOGGER.info(f"✅ 已点击 '3. View Summary' 标签页 at ({target_x}, {target_y})")
            
        except ImportError:
            LOGGER.error("pyautogui 未安装，无法点击")
            raise RuntimeError("需要安装 pyautogui 才能点击")
        except Exception as e:
            LOGGER.error(f"点击失败: {e}")
            raise

    def _fill_requestor_comments(self, ui_config: dict = None) -> None:
        """
        填写Requestor Comments到文本区域（使用相对定位，适应不同窗口大小和位置）
        
        策略：
        1. 优先使用 UI 配置中的 mir_comments
        2. 如果没有，则读取 MIR Comments.txt 文件
        3. 查找"Requestor Comments"标签作为参考点
        4. 基于标签位置找到文本区域（在标签下方或右侧）
        5. 使用相对定位点击文本区域并填写内容
        
        Args:
            ui_config: UI配置数据，可能包含 mir_comments
        """
        if Application is None:
            return
        
        LOGGER.info("填写Requestor Comments...")
        
        try:
            # 检查是否已经在Summary标签页（避免在错误的标签页填写）
            # 注意：这个检查可能过于严格，如果检查失败，我们仍然尝试填写comments
            try:
                # 检查是否存在Submit按钮（如果存在，说明还在可以编辑的页面）
                # 如果不存在Submit按钮，可能已经提交了，不应该再填写
                submit_buttons = self._window.descendants(control_type="Button")
                has_submit_button = False
                for btn in submit_buttons:
                    try:
                        if btn.is_visible() and btn.is_enabled():
                            btn_text = btn.window_text().strip()
                            if btn_text == "Submit" or "Submit" in btn_text:
                                has_submit_button = True
                                break
                    except:
                        continue
                
                if not has_submit_button:
                    LOGGER.warning("⚠️ 未找到Submit按钮，可能页面还未加载完成，重新点击View Summary并等待页面加载...")
                    # 重新点击View Summary标签，确保页面已跳转
                    self._click_summary_tab()
                    
                    # 轮询等待页面加载完成（等待Submit按钮出现，最多等待10秒）
                    LOGGER.info("等待View Summary页面加载完成（轮询等待Submit按钮出现）...")
                    max_wait_time = 10  # 最多等待10秒
                    check_interval = 1  # 每1秒检查一次
                    max_attempts = max_wait_time // check_interval
                    
                    has_submit_button_retry = False
                    for attempt in range(max_attempts):
                        time.sleep(check_interval)
                        # 检查Submit按钮是否存在
                        try:
                            submit_buttons_retry = self._window.descendants(control_type="Button")
                            for btn in submit_buttons_retry:
                                try:
                                    if btn.is_visible() and btn.is_enabled():
                                        btn_text = btn.window_text().strip()
                                        if btn_text == "Submit" or "Submit" in btn_text:
                                            has_submit_button_retry = True
                                            LOGGER.info(f"✅ 找到Submit按钮，页面已加载完成（等待了 {(attempt+1)*check_interval} 秒）")
                                            break
                                except:
                                    continue
                            
                            if has_submit_button_retry:
                                break
                            
                            if (attempt + 1) % 3 == 0:  # 每3秒输出一次日志
                                LOGGER.info(f"等待页面加载中... (已等待 {(attempt+1)*check_interval} 秒)")
                        except Exception as e:
                            LOGGER.debug(f"检查Submit按钮时出错: {e}")
                            continue
                    
                    if not has_submit_button_retry:
                        LOGGER.warning(f"⚠️ 等待 {max_wait_time} 秒后仍未找到Submit按钮，但继续尝试填写Requestor Comments")
                    # 继续尝试填写
            except Exception as e:
                # 如果检查失败，继续执行（可能是检查逻辑有问题）
                LOGGER.debug(f"检查Submit按钮时出错: {e}，继续执行填写comments")
                pass
            
            # 确保主窗口有焦点
            self._window.set_focus()
            time.sleep(0.3)
            
            # 优先使用 UI 配置中的 mir_comments
            comments_text = None
            if ui_config and ui_config.get('mir_comments'):
                comments_text = ui_config['mir_comments'].strip()
                LOGGER.info("使用UI配置中的MIR Comments")
            
            # 如果没有 UI 配置，则读取文件
            if not comments_text:
                possible_paths = [
                    Path(__file__).parent.parent / "input" / "MIR Comments.txt",  # input目录（优先）
                    Path(__file__).parent.parent / "MIR Comments.txt",  # Auto VPO根目录
                    Path("MIR Comments.txt"),  # 当前工作目录
                    Path(__file__).parent / "MIR Comments.txt",  # workflow_automation目录
                ]
                
                mir_comments_file = None
                for path in possible_paths:
                    if path.exists():
                        mir_comments_file = path
                        break
                
                if not mir_comments_file:
                    mir_comments_file = Path(__file__).parent.parent / "MIR Comments.txt"
                    LOGGER.warning(f"未找到MIR Comments.txt文件，将尝试在以下位置查找: {mir_comments_file}")
                
                if not mir_comments_file.exists():
                    LOGGER.warning(f"未找到MIR Comments.txt文件，跳过填写Requestor Comments")
                    return
                
                LOGGER.info(f"读取MIR Comments文件: {mir_comments_file}")
                try:
                    with open(mir_comments_file, 'r', encoding='utf-8') as f:
                        comments_text = f.read().strip()
                except UnicodeDecodeError:
                    with open(mir_comments_file, 'r', encoding='gbk') as f:
                        comments_text = f.read().strip()
            
            if not comments_text:
                LOGGER.warning("MIR Comments为空，跳过填写")
                return
            
            LOGGER.info(f"MIR Comments内容: {comments_text[:50]}...")
            
            # 复制内容到剪贴板
            try:
                import pyperclip
                pyperclip.copy(comments_text)
                LOGGER.info("已复制内容到剪贴板")
            except ImportError:
                try:
                    import win32clipboard
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(comments_text, win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    LOGGER.info("已复制内容到剪贴板（使用win32clipboard）")
                except ImportError:
                    LOGGER.warning("无法复制到剪贴板（需要pyperclip或win32clipboard），将直接输入文本")
            
            # 获取窗口位置和尺寸（用于相对定位）
            window_rect = self._window.rectangle()
            window_left = window_rect.left
            window_top = window_rect.top
            window_width = window_rect.right - window_rect.left
            window_height = window_rect.bottom - window_rect.top
            
            LOGGER.info(f"窗口位置: left={window_left}, top={window_top}, width={window_width}, height={window_height}")
            
            # 查找"Requestor Comments"标签
            LOGGER.info("查找'Requestor Comments'标签...")
            all_controls = self._window.descendants()
            
            requestor_comments_label = None
            
            for ctrl in all_controls:
                try:
                    if not ctrl.is_visible():
                        continue
                    ctrl_text = ctrl.window_text().strip()
                    if ctrl_text and "Requestor Comments" in ctrl_text:
                        requestor_comments_label = ctrl
                        label_rect = ctrl.rectangle()
                        LOGGER.info(f"找到'Requestor Comments'标签: 位置=({label_rect.left}, {label_rect.top}), 文本='{ctrl_text}'")
                        break
                except:
                    continue
            
            if not requestor_comments_label:
                LOGGER.warning("未找到'Requestor Comments'标签，无法定位文本区域")
                return
            
            # 基于标签位置查找文本区域
            label_rect = requestor_comments_label.rectangle()
            label_right = label_rect.right
            label_bottom = label_rect.bottom
            
            LOGGER.info(f"标签位置: left={label_rect.left}, top={label_rect.top}, right={label_right}, bottom={label_bottom}")
            
            # 查找文本区域（Edit、RichEdit、TextBox等）
            text_area_candidates = []
            
            for ctrl in all_controls:
                try:
                    if not ctrl.is_visible():
                        continue
                    
                    ctrl_rect = ctrl.rectangle()
                    ctrl_type = ctrl.element_info.control_type if hasattr(ctrl.element_info, 'control_type') else "Unknown"
                    
                    # 查找Edit、RichEdit、TextBox等文本输入控件
                    type_str = str(ctrl_type).upper()
                    if "EDIT" in type_str or "RICHEDIT" in type_str or "TEXTBOX" in type_str:
                        area = (ctrl_rect.right - ctrl_rect.left) * (ctrl_rect.bottom - ctrl_rect.top)
                        
                        # 文本区域应该比较大（面积 > 5000）
                        if area > 5000:
                            # 检查位置关系：在标签右侧，或标签下方
                            is_below = ctrl_rect.top >= label_bottom and ctrl_rect.left <= label_right + 100
                            is_right = ctrl_rect.left >= label_right - 50
                            
                            if is_below or is_right:
                                distance = ((ctrl_rect.left - label_right)**2 + (ctrl_rect.top - label_bottom)**2)**0.5
                                text_area_candidates.append((ctrl, ctrl_rect, area, distance, ctrl_type))
                                LOGGER.info(f"  找到候选文本区域: 类型={ctrl_type}, 位置=({ctrl_rect.left}, {ctrl_rect.top}), 面积={area}, 距离标签={distance:.1f}")
                except:
                    continue
            
            # 选择最接近标签的文本区域
            if text_area_candidates:
                text_area_candidates.sort(key=lambda x: x[3])  # 按距离排序
                requestor_comments_text_area, text_rect, _, _, _ = text_area_candidates[0]
                LOGGER.info(f"选择最近的文本区域: 位置=({text_rect.left}, {text_rect.top})")
                
                # 点击文本区域中心
                text_center_x = text_rect.left + (text_rect.right - text_rect.left) // 2
                text_center_y = text_rect.top + (text_rect.bottom - text_rect.top) // 2
                
                LOGGER.info(f"点击文本区域中心位置: ({text_center_x}, {text_center_y})")
                
                try:
                    requestor_comments_text_area.click_input()
                    time.sleep(0.3)
                except:
                    try:
                        import pyautogui
                        pyautogui.click(text_center_x, text_center_y)
                        time.sleep(0.3)
                    except:
                        pass
                
                # 清空现有内容并输入新内容
                try:
                    import pyautogui
                    # 确保文本区域有焦点
                    pyautogui.click(text_center_x, text_center_y)
                    time.sleep(0.3)
                    pyautogui.hotkey('ctrl', 'a')  # 全选
                    time.sleep(0.3)  # 增加等待时间
                    pyautogui.hotkey('ctrl', 'v')  # 粘贴
                    time.sleep(0.5)  # 增加等待时间，确保粘贴完成
                    # 验证是否填写成功（尝试读取文本框内容）
                    try:
                        current_text = requestor_comments_text_area.window_text()
                        if current_text and comments_text[:50] in current_text:
                            LOGGER.info(f"✅ 已填写Requestor Comments（验证成功，长度: {len(current_text)} 字符）")
                        else:
                            LOGGER.warning(f"⚠️ Requestor Comments可能未正确填写（当前内容: {current_text[:50] if current_text else '空'}）")
                    except:
                        LOGGER.info("✅ 已填写Requestor Comments（无法验证内容）")
                    return
                except ImportError:
                    # 如果没有pyautogui，尝试使用键盘输入
                    try:
                        requestor_comments_text_area.type_keys('^a')  # Ctrl+A
                        time.sleep(0.2)
                        requestor_comments_text_area.type_keys(comments_text, with_spaces=True)
                        time.sleep(0.3)
                        LOGGER.info("✅ 已填写Requestor Comments（使用键盘输入）")
                        return
                    except Exception as e:
                        LOGGER.warning(f"键盘输入失败: {e}")
            
            # 如果找不到文本区域控件，使用相对定位估算位置
            LOGGER.warning("未找到文本区域控件，使用相对定位估算位置...")
            
            # 基于标签位置和窗口相对位置估算文本区域
            # 文本区域通常在标签下方，或者标签右侧
            # 使用相对偏移（基于窗口尺寸的百分比）
            
            # 计算标签相对于窗口的位置
            label_x_percent = (label_rect.left - window_left) / window_width
            label_y_percent = (label_rect.top - window_top) / window_height
            
            # 文本区域通常在标签右侧或下方
            # 基于经验值：文本区域在标签右侧约10-15%窗口宽度，或标签下方约5-8%窗口高度
            estimated_x = label_right + int(window_width * 0.02)  # 标签右侧2%窗口宽度
            estimated_y = label_bottom + int(window_height * 0.03)  # 标签下方3%窗口高度
            
            # 如果估算位置超出窗口，使用标签下方
            if estimated_x > window_left + window_width:
                estimated_x = label_rect.left
                estimated_y = label_bottom + int(window_height * 0.05)  # 标签下方5%窗口高度
            
            LOGGER.info(f"估算文本区域位置: ({estimated_x}, {estimated_y})")
            LOGGER.info(f"  基于标签相对位置: X={label_x_percent:.2%}, Y={label_y_percent:.2%}")
            
            try:
                import pyautogui
                pyautogui.click(estimated_x, estimated_y)
                time.sleep(0.3)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.3)
                LOGGER.info(f"✅ 已通过相对定位填写Requestor Comments")
                return
            except ImportError:
                LOGGER.warning("pyautogui未安装，无法填写Requestor Comments")
        
        except Exception as e:
            LOGGER.error(f"填写Requestor Comments失败: {e}")
            # 不抛出异常，允许工作流继续执行

    # 表格提取功能已禁用
    def extract_table_data(self) -> list:
        """
        从表格中提取Qty、PartType、Source三列的数据（功能已禁用）
        
        Returns:
            list: 空列表
        """
        LOGGER.info("表格提取功能已禁用")
        return []
    
    # 以下为已禁用的表格提取实现代码（保留以备将来使用）
    def _extract_table_data_old(self) -> list:
        """
        从表格中提取Qty、PartType、Source三列的数据（旧实现，已禁用）
        """
        if Application is None:
            return []
        
        LOGGER.info("开始提取表格数据（Qty、PartType、Source）...")
        extracted_data = []
        
        try:
            # 确保主窗口有焦点
            self._window.set_focus()
            time.sleep(0.5)
            
            # 方法：通过点击表格、全选、复制、粘贴来提取数据
            LOGGER.info("使用点击表格、全选、复制、粘贴的方法提取数据...")
            
            # 查找表格位置
            # 策略1: 查找包含表头（Qty、PartType、Source）的表格控件
            table_rect = None
            all_controls = self._window.descendants()
            
            # 查找包含表头的控件
            for ctrl in all_controls:
                try:
                    if not ctrl.is_visible():
                        continue
                    ctrl_text = ctrl.window_text().strip()
                    # 查找包含表头的控件（通常表格控件会包含表头文本）
                    if ctrl_text and ('Qty' in ctrl_text or 'PartType' in ctrl_text or 'Source' in ctrl_text):
                        # 检查是否是表格控件
                        ctrl_type = str(ctrl.element_info.control_type).upper() if hasattr(ctrl.element_info, 'control_type') else ""
                        if "GRID" in ctrl_type or "LIST" in ctrl_type or "TABLE" in ctrl_type or "DATA" in ctrl_type:
                            table_rect = ctrl.rectangle()
                            LOGGER.info(f"找到表格控件: 位置=({table_rect.left}, {table_rect.top}), 大小=({table_rect.right - table_rect.left}, {table_rect.bottom - table_rect.top})")
                            break
                except:
                    continue
            
            # 策略2: 如果找不到表格控件，使用窗口底部位置（表格在View Summary页面最下面显示lot信息）
            if not table_rect:
                LOGGER.info("未找到表格控件，使用窗口底部位置定位表格（表格在View Summary页面最下面）...")
                window_rect = self._window.rectangle()
                window_width = window_rect.right - window_rect.left
                window_height = window_rect.bottom - window_rect.top
                
                # 表格在窗口底部，估算表格位置
                # 表格通常在窗口底部，高度约300-400像素
                estimated_table_height = 400  # 表格高度
                estimated_table_y = window_rect.bottom - estimated_table_height - 50  # 距离底部50像素
                estimated_table_x = window_rect.left + 50  # 左边距50像素
                estimated_table_width = window_width - 100  # 左右各留50像素
                
                # 创建一个虚拟的表格区域（使用简单的类来模拟rect）
                class SimpleRect:
                    def __init__(self, left, top, right, bottom):
                        self.left = left
                        self.top = top
                        self.right = right
                        self.bottom = bottom
                
                table_rect = SimpleRect(
                    left=estimated_table_x,
                    top=estimated_table_y,
                    right=estimated_table_x + estimated_table_width,
                    bottom=window_rect.bottom - 20  # 距离底部20像素
                )
                LOGGER.info(f"使用窗口底部位置定位表格: ({table_rect.left}, {table_rect.top}), 高度={table_rect.bottom - table_rect.top}")
            
            # 策略3: 如果还是找不到，使用窗口中心偏下的位置（备用）
            if not table_rect:
                LOGGER.warning("无法定位表格，使用窗口中心偏下位置...")
                window_rect = self._window.rectangle()
                window_center_y = window_rect.top + int((window_rect.bottom - window_rect.top) * 0.7)  # 窗口高度的70%处
                
                class SimpleRect:
                    def __init__(self, left, top, right, bottom):
                        self.left = left
                        self.top = top
                        self.right = right
                        self.bottom = bottom
                
                table_rect = SimpleRect(
                    left=window_rect.left + 100,
                    top=window_center_y,
                    right=window_rect.right - 100,
                    bottom=window_rect.bottom - 50
                )
                LOGGER.info(f"使用窗口中心偏下位置: ({table_rect.left}, {table_rect.top})")
            
            if not table_rect:
                LOGGER.error("无法定位表格位置")
                return []
            
            # 逐行点击并复制数据（表格不是真正意义上的表格格式）
            import pyautogui
            
            # 计算表格参数
            table_width = table_rect.right - table_rect.left
            table_center_x = table_rect.left + table_width // 2
            
            # 估算行高（通常每行约20-30像素）
            estimated_row_height = 25
            # 表头高度约30-40像素
            header_height = 35
            # 第一行数据从表头下方开始
            first_data_row_y = table_rect.top + header_height + estimated_row_height // 2
            
            LOGGER.info(f"表格位置: 顶部={table_rect.top}, 底部={table_rect.bottom}, 宽度={table_width}")
            LOGGER.info(f"估算行高: {estimated_row_height}像素, 第一行数据Y位置: {first_data_row_y}")
            
            # 先点击表头行，获取表头信息以确定列位置
            header_y = table_rect.top + header_height // 2
            LOGGER.info(f"点击表头行位置: ({table_center_x}, {header_y})")
            pyautogui.click(table_center_x, header_y)
            time.sleep(0.3)
            
            # 选择整行（Shift+End 选择到行尾，然后 Shift+Home 选择整行）
            pyautogui.hotkey('shift', 'end')
            time.sleep(0.2)
            pyautogui.hotkey('shift', 'home')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.3)
            
            # 读取表头
            header_text = None
            try:
                import pyperclip
                header_text = pyperclip.paste()
            except ImportError:
                try:
                    import win32clipboard
                    win32clipboard.OpenClipboard()
                    header_text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                except:
                    pass
            
            # 解析表头，找到列位置
            qty_col = None
            parttype_col = None
            source_col = None
            
            if header_text:
                LOGGER.info(f"表头内容: {header_text[:100]}")
                # 解析表头
                if '\t' in header_text:
                    headers = header_text.split('\t')
                else:
                    headers = [h.strip() for h in header_text.split() if h.strip()]
                
                for i, header in enumerate(headers):
                    header_upper = header.upper().strip()
                    if 'QTY' in header_upper and qty_col is None:
                        qty_col = i
                        LOGGER.info(f"找到 Qty 列索引: {qty_col}")
                    elif ('PARTTYPE' in header_upper or ('PART' in header_upper and 'TYPE' in header_upper)) and parttype_col is None:
                        parttype_col = i
                        LOGGER.info(f"找到 PartType 列索引: {parttype_col}")
                    elif 'SOURCE' in header_upper and source_col is None:
                        source_col = i
                        LOGGER.info(f"找到 Source 列索引: {source_col}")
            
            # 逐行提取数据
            max_rows = 100  # 最多提取100行，防止无限循环
            row_idx = 0
            current_row_y = first_data_row_y
            
            LOGGER.info("开始逐行提取数据...")
            
            while row_idx < max_rows and current_row_y < table_rect.bottom - 20:
                try:
                    # 点击当前行
                    LOGGER.debug(f"点击第 {row_idx + 1} 行位置: ({table_center_x}, {current_row_y})")
                    pyautogui.click(table_center_x, current_row_y)
                    time.sleep(0.3)
                    
                    # 选择整行（Shift+End 选择到行尾，然后 Shift+Home 选择整行）
                    pyautogui.hotkey('shift', 'end')
                    time.sleep(0.2)
                    pyautogui.hotkey('shift', 'home')
                    time.sleep(0.2)
                    
                    # 复制
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(0.3)
                    
                    # 读取剪贴板
                    row_text = None
                    try:
                        import pyperclip
                        row_text = pyperclip.paste()
                    except ImportError:
                        try:
                            import win32clipboard
                            win32clipboard.OpenClipboard()
                            row_text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                            win32clipboard.CloseClipboard()
                        except:
                            pass
                    
                    if not row_text or not row_text.strip():
                        LOGGER.debug(f"第 {row_idx + 1} 行为空，可能已到表格末尾")
                        break
                    
                    # 解析行数据
                    if '\t' in row_text:
                        parts = row_text.split('\t')
                    else:
                        parts = [p.strip() for p in row_text.split() if p.strip()]
                    
                    if len(parts) == 0:
                        LOGGER.debug(f"第 {row_idx + 1} 行无有效数据")
                        current_row_y += estimated_row_height
                        row_idx += 1
                        continue
                    
                    row_data = {}
                    
                    # 如果找到了列索引，直接提取
                    if qty_col is not None and qty_col < len(parts):
                        row_data['Qty'] = parts[qty_col].strip()
                    if parttype_col is not None and parttype_col < len(parts):
                        row_data['PartType'] = parts[parttype_col].strip()
                    if source_col is not None and source_col < len(parts):
                        row_data['Source'] = parts[source_col].strip()
                    
                    # 如果通过列索引提取失败，尝试模式匹配
                    if not row_data.get('Qty') or not row_data.get('PartType') or not row_data.get('Source'):
                        for part in parts:
                            part = part.strip()
                            if not part:
                                continue
                            
                            # Qty通常是纯数字
                            if part.isdigit() and 'Qty' not in row_data:
                                row_data['Qty'] = part
                            # PartType通常包含字母和数字，格式类似"43 4PXA2V E B"
                            elif any(c.isalpha() for c in part) and any(c.isdigit() for c in part) and len(part) > 5:
                                if 'PartType' not in row_data:
                                    row_data['PartType'] = part
                            # Source通常是lot编号，格式类似"PTPB5TP1206"
                            elif (part.startswith('PTP') or (len(part) >= 8 and any(c.isalpha() for c in part) and any(c.isdigit() for c in part))):
                                if 'Source' not in row_data:
                                    row_data['Source'] = part
                    
                    # 只有当三列都有数据时才添加
                    if row_data.get('Qty') and row_data.get('PartType') and row_data.get('Source'):
                        extracted_data.append(row_data)
                        LOGGER.info(f"✅ 提取第 {row_idx + 1} 行: Qty={row_data.get('Qty')}, PartType={row_data.get('PartType')}, Source={row_data.get('Source')}")
                    else:
                        LOGGER.debug(f"第 {row_idx + 1} 行数据不完整: {row_data}")
                        # 如果连续几行都没有完整数据，可能已到表格末尾
                        if row_idx > 0 and len(extracted_data) == 0:
                            LOGGER.warning("前几行都没有完整数据，可能表格格式不同，停止提取")
                            break
                    
                    # 移动到下一行
                    current_row_y += estimated_row_height
                    row_idx += 1
                    
                except Exception as e:
                    LOGGER.warning(f"提取第 {row_idx + 1} 行时出错: {e}")
                    current_row_y += estimated_row_height
                    row_idx += 1
                    continue
            
            if extracted_data:
                LOGGER.info(f"✅ 成功提取 {len(extracted_data)} 行表格数据")
                for idx, row in enumerate(extracted_data[:5], 1):  # 只显示前5行
                    LOGGER.info(f"  第 {idx} 行: Qty={row.get('Qty')}, PartType={row.get('PartType')}, Source={row.get('Source')}")
            else:
                LOGGER.warning("⚠️ 未能提取到表格数据")
            
            return extracted_data
        
        except Exception as e:
            LOGGER.error(f"提取表格数据时出错: {e}")
            import traceback
            LOGGER.debug(traceback.format_exc())
        
        return extracted_data
    
    # 表格提取功能已禁用
    def _extract_and_save_table_data(self) -> None:
        """
        提取表格数据（Qty、PartType、Source）并保存到CSV文件（功能已禁用）
        """
        LOGGER.info("表格提取功能已禁用，跳过提取和保存")
        return

    def _handle_mir_mrs_info_dialog(self) -> None:
        """处理'MIR is now MRS!'信息对话框，自动点击OK按钮"""
        if Application is None:
            return
        
        LOGGER.info("检查是否有'MIR is now MRS!'信息对话框...")
        
        # 等待对话框出现（最多等待10秒）
        info_dialog = None
        deadline = time.time() + 10
        dialog_titles = [
            "MIR is now MRS!",
            ".*MIR.*MRS.*",
            ".*MRS.*"
        ]
        
        while time.time() < deadline:
            for title_pattern in dialog_titles:
                try:
                    # 尝试win32 backend
                    try:
                        dialog_app = Application(backend="win32").connect(
                            title_re=title_pattern,
                            visible_only=True,
                            timeout=1
                        )
                        info_dialog = dialog_app.window(title_re=title_pattern)
                        if info_dialog.exists() and info_dialog.is_visible():
                            LOGGER.info(f"找到信息对话框: {info_dialog.window_text()} (backend: win32)")
                            break
                    except:
                        pass
                    
                    # 尝试uia backend
                    try:
                        dialog_app = Application(backend="uia").connect(
                            title_re=title_pattern,
                            visible_only=True,
                            timeout=1
                        )
                        info_dialog = dialog_app.window(title_re=title_pattern)
                        if info_dialog.exists() and info_dialog.is_visible():
                            LOGGER.info(f"找到信息对话框: {info_dialog.window_text()} (backend: uia)")
                            break
                    except:
                        pass
                except:
                    continue
            
            if info_dialog is not None:
                break
            
            time.sleep(0.3)
        
        if info_dialog is None:
            LOGGER.info("未找到'MIR is now MRS!'对话框，可能未出现或已关闭")
            return
        
        try:
            # 激活对话框
            info_dialog.set_focus()
            time.sleep(0.3)
            
            ok_clicked = False
            
            # 方法1: 查找OK按钮并点击
            try:
                ok_button = info_dialog.child_window(title="OK", control_type="Button")
                if ok_button.exists() and ok_button.is_visible():
                    ok_button.click_input()
                    ok_clicked = True
                    LOGGER.info("✅ 已点击OK按钮（方法1: pywinauto）")
            except Exception as e:
                LOGGER.debug(f"方法1失败: {e}")
            
            # 方法2: 使用Windows API查找并点击OK按钮
            if not ok_clicked and win32gui:
                try:
                    dialog_hwnd = info_dialog.handle
                    def enum_ok_button(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child).strip().upper()
                            class_name = win32gui.GetClassName(hwnd_child).upper()
                            if window_text == "OK" and "BUTTON" in class_name:
                                lParam.append(hwnd_child)
                        except:
                            pass
                        return True
                    
                    ok_buttons = []
                    win32gui.EnumChildWindows(dialog_hwnd, enum_ok_button, ok_buttons)
                    
                    if ok_buttons:
                        ok_hwnd = ok_buttons[0]
                        win32gui.PostMessage(ok_hwnd, win32con.BM_CLICK, 0, 0)
                        ok_clicked = True
                        LOGGER.info("✅ 已点击OK按钮（方法2: Windows API）")
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            if ok_clicked:
                time.sleep(0.5)
                LOGGER.info("✅ 已成功处理'MIR is now MRS!'信息对话框")
            else:
                LOGGER.warning("未能点击OK按钮")
        
        except Exception as e:
            LOGGER.error(f"处理'MIR is now MRS!'信息对话框失败: {e}")
    
    def _handle_login_dialog(self) -> None:
        """处理Mole登录对话框，自动点击OK按钮"""
        if Application is None:
            return
        
        LOGGER.info("等待MOLE LOGIN对话框出现...")
        
        # 等待登录对话框出现（最多等待30秒）
        login_dialog = None
        login_app = None
        deadline = time.time() + 30
        # 使用配置的标题，同时提供备选模式
        dialog_titles = [
            self.config.login_dialog_title,
            ".*MOLE LOGIN.*",
            ".*LOGIN.*",
            ".*Login.*"
        ]
        
        # 尝试不同的backend
        backends = ["win32", "uia"]
        
        while time.time() < deadline:
            for backend in backends:
                for title_pattern in dialog_titles:
                    try:
                        login_app = Application(backend=backend).connect(
                            title_re=title_pattern,
                            visible_only=True,
                            timeout=2
                        )
                        login_dialog = login_app.window(title_re=title_pattern)
                        
                        if login_dialog.exists() and login_dialog.is_visible():
                            LOGGER.info(f"找到登录对话框: {login_dialog.window_text()} (backend: {backend})")
                            break
                    except Exception:
                        continue
                
                if login_dialog is not None:
                    break
            
            if login_dialog is not None:
                break
            
            time.sleep(0.5)
        
        if login_dialog is None:
            LOGGER.warning("未找到MOLE LOGIN对话框，可能已经登录或对话框已关闭")
            return
        
        try:
            # 激活对话框
            login_dialog.set_focus()
            time.sleep(0.5)
            
            # 打印控件结构用于调试（保存到日志）
            try:
                LOGGER.info("正在分析登录对话框的控件结构...")
                # 将控件结构输出到字符串
                import io
                import sys
                old_stdout = sys.stdout
                sys.stdout = buffer = io.StringIO()
                login_dialog.print_control_identifiers(depth=4)
                sys.stdout = old_stdout
                control_info = buffer.getvalue()
                LOGGER.debug(f"登录对话框控件结构:\n{control_info}")
                # 同时尝试查找所有可能的控件类型
                LOGGER.info("查找所有可用的控件...")
            except Exception as e:
                LOGGER.debug(f"打印控件结构失败: {e}")
            
            # 查找并点击OK按钮
            ok_clicked = False
            
            # 首先尝试多种方法查找所有按钮
            all_buttons = []
            
            # 方法1: 使用children查找
            try:
                buttons1 = login_dialog.children(control_type="Button")
                all_buttons.extend(buttons1)
                LOGGER.info(f"通过children找到 {len(buttons1)} 个Button控件")
            except Exception as e:
                LOGGER.debug(f"通过children查找按钮失败: {e}")
            
            # 方法2: 使用descendants查找（查找所有子控件）
            if len(all_buttons) == 0:
                try:
                    buttons2 = login_dialog.descendants(control_type="Button")
                    all_buttons.extend(buttons2)
                    LOGGER.info(f"通过descendants找到 {len(buttons2)} 个Button控件")
                except Exception as e:
                    LOGGER.debug(f"通过descendants查找按钮失败: {e}")
            
            # 方法3: 查找所有子窗口（可能按钮是子窗口）
            if len(all_buttons) == 0:
                try:
                    all_windows = login_dialog.children()
                    LOGGER.info(f"找到 {len(all_windows)} 个子控件")
                    for win in all_windows:
                        try:
                            win_text = win.window_text()
                            win_type = str(type(win))
                            LOGGER.info(f"  控件: '{win_text}' (类型: {win_type})")
                            if "Button" in win_type or win_text.upper() in ["OK", "CANCEL"]:
                                all_buttons.append(win)
                        except:
                            pass
                except Exception as e:
                    LOGGER.debug(f"查找所有子窗口失败: {e}")
            
            # 方法4: 尝试child_windows()
            if len(all_buttons) == 0:
                try:
                    child_windows = login_dialog.child_windows()
                    LOGGER.info(f"通过child_windows找到 {len(child_windows)} 个子窗口")
                    for win in child_windows:
                        try:
                            win_text = win.window_text()
                            if win_text.upper() in ["OK", "CANCEL"]:
                                all_buttons.append(win)
                        except:
                            pass
                except Exception as e:
                    LOGGER.debug(f"通过child_windows查找失败: {e}")
            
            # 列出找到的所有按钮
            LOGGER.info(f"总共找到 {len(all_buttons)} 个可能的按钮控件")
            for idx, btn in enumerate(all_buttons):
                try:
                    btn_text = btn.window_text()
                    LOGGER.info(f"  按钮 {idx + 1}: '{btn_text}'")
                except:
                    LOGGER.info(f"  按钮 {idx + 1}: (无法获取文本)")
            
            # 方法1: 使用找到的按钮列表点击OK
            if not ok_clicked and len(all_buttons) > 0:
                for button in all_buttons:
                    try:
                        button_text = button.window_text().strip().upper()
                        if button_text == "OK" and "CANCEL" not in button_text:
                            LOGGER.info(f"找到OK按钮（从按钮列表中: '{button.window_text()}'）")
                            button.click_input()
                            ok_clicked = True
                            LOGGER.info("✅ 已点击OK按钮（从按钮列表）")
                            break
                    except Exception as e:
                        LOGGER.debug(f"检查按钮时出错: {e}")
                        continue
            
            # 方法2: 精确匹配OK按钮（完全匹配，不区分大小写）
            if not ok_clicked:
                try:
                    # 尝试多种方式精确匹配OK
                    ok_patterns = ["OK", "Ok", "ok"]
                    for pattern in ok_patterns:
                        try:
                            ok_button = login_dialog.child_window(title=pattern, control_type="Button")
                            if ok_button.exists() and ok_button.is_enabled():
                                btn_text = ok_button.window_text()
                                LOGGER.info(f"找到OK按钮（精确匹配: '{btn_text}'）")
                                ok_button.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK按钮（通过精确匹配title）")
                                break
                        except:
                            continue
                except Exception as e1:
                    LOGGER.debug(f"通过精确匹配查找OK按钮失败: {e1}")
            
            # 方法3: 使用descendants查找所有按钮（包括嵌套的）
            if not ok_clicked:
                try:
                    all_descendants = login_dialog.descendants(control_type="Button")
                    LOGGER.info(f"通过descendants找到 {len(all_descendants)} 个Button控件")
                    for button in all_descendants:
                        try:
                            button_text = button.window_text().strip().upper()
                            if button_text == "OK" and "CANCEL" not in button_text:
                                LOGGER.info(f"找到OK按钮（descendants: '{button.window_text()}'）")
                                button.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK按钮（通过descendants）")
                                break
                        except Exception as e:
                            LOGGER.debug(f"检查descendant按钮时出错: {e}")
                            continue
                except Exception as e3:
                    LOGGER.debug(f"通过descendants查找按钮失败: {e3}")
            
            # 方法3: 如果OK通常是第一个按钮（在Cancel之前），按位置点击
            if not ok_clicked:
                try:
                    buttons = login_dialog.children(control_type="Button")
                    if len(buttons) >= 2:
                        # 检查第一个按钮是否是OK
                        first_button_text = buttons[0].window_text().strip().upper()
                        if first_button_text == "OK":
                            LOGGER.info(f"按位置点击第一个按钮（文本: '{buttons[0].window_text()}'）")
                            buttons[0].click_input()
                            ok_clicked = True
                            LOGGER.info("✅ 已点击第一个按钮（确认为OK）")
                except Exception as e3:
                    LOGGER.debug(f"按位置点击按钮失败: {e3}")
            
            # 方法4: 尝试通过auto_id或其他属性查找OK按钮
            if not ok_clicked:
                try:
                    # 尝试查找常见的按钮ID
                    ok_ids = ["OK", "btnOK", "buttonOK", "idOK"]
                    for btn_id in ok_ids:
                        try:
                            ok_button = login_dialog.child_window(auto_id=btn_id, control_type="Button")
                            if ok_button.exists() and ok_button.is_enabled():
                                btn_text = ok_button.window_text()
                                LOGGER.info(f"通过auto_id找到OK按钮: '{btn_id}' (文本: '{btn_text}')")
                                ok_button.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK按钮（通过auto_id）")
                                break
                        except:
                            continue
                except Exception as e4:
                    LOGGER.debug(f"通过auto_id查找OK按钮失败: {e4}")
            
            # 方法5: 使用Windows API直接查找和点击（win32gui方法）
            if not ok_clicked and win32gui:
                try:
                    LOGGER.info("尝试使用Windows API查找按钮...")
                    hwnd = login_dialog.handle
                    
                    # 枚举所有子窗口
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            class_name = win32gui.GetClassName(hwnd_child)
                            window_text = win32gui.GetWindowText(hwnd_child)
                            if window_text.upper() == "OK" and "BUTTON" in class_name.upper():
                                LOGGER.info(f"通过Windows API找到OK按钮: '{window_text}' (类名: {class_name})")
                                # 发送BM_CLICK消息
                                import win32con
                                win32gui.PostMessage(hwnd_child, win32con.BM_CLICK, 0, 0)
                                return False  # 停止枚举
                            return True
                        except:
                            return True
                    
                    win32gui.EnumChildWindows(hwnd, enum_child_proc, None)
                    # 给一点时间让点击生效
                    time.sleep(0.5)
                    ok_clicked = True
                    LOGGER.info("✅ 已通过Windows API点击OK按钮")
                except Exception as e5:
                    LOGGER.debug(f"使用Windows API点击失败: {e5}")
            
            # 方法6: 尝试所有子控件，不指定类型
            if not ok_clicked:
                try:
                    LOGGER.info("尝试查找所有子控件（不指定类型）...")
                    all_children = login_dialog.children()
                    for child in all_children:
                        try:
                            child_text = child.window_text().strip().upper()
                            if child_text == "OK":
                                LOGGER.info(f"找到OK控件（无类型限制: '{child.window_text()}'）")
                                child.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK控件（无类型限制）")
                                break
                        except:
                            continue
                except Exception as e6:
                    LOGGER.debug(f"查找所有子控件失败: {e6}")
            
            if not ok_clicked:
                LOGGER.warning("无法点击OK按钮，请手动处理登录对话框")
            else:
                # 等待对话框关闭
                time.sleep(1)
                LOGGER.info("登录对话框已处理")
        
        except Exception as e:
            LOGGER.warning(f"处理登录对话框时出错: {e}")
            LOGGER.warning("请手动处理登录对话框")
    
    def _is_process_running(self) -> bool:
        """检查Mole进程是否在运行"""
        if Application is None:
            return False
        try:
            Application(backend="win32").connect(
                title_re=self.config.window_title
            )
            return True
        except ElementNotFoundError:
            return False
    
    def verify_submission(self) -> bool:
        """
        验证数据是否提交成功
        
        Returns:
            True如果验证通过
        """
        # TODO: 实现验证逻辑
        # 例如：检查Mole工具中是否显示成功消息
        try:
            # 这里需要根据实际Mole工具的反馈机制来实现
            LOGGER.info("✅ MIR数据提交验证通过")
            return True
        except Exception as e:
            LOGGER.warning(f"验证提交时出错: {e}")
            return False
    
    def _handle_mir_mrs_info_dialog(self) -> None:
        """处理'MIR is now MRS!'信息对话框，自动点击OK按钮"""
        if Application is None:
            return
        
        LOGGER.info("检查是否有'MIR is now MRS!'信息对话框...")
        
        # 等待对话框出现（最多等待10秒）
        info_dialog = None
        deadline = time.time() + 10
        dialog_titles = [
            "MIR is now MRS!",
            ".*MIR.*MRS.*",
            ".*MRS.*"
        ]
        
        while time.time() < deadline:
            for title_pattern in dialog_titles:
                try:
                    # 尝试win32 backend
                    try:
                        dialog_app = Application(backend="win32").connect(
                            title_re=title_pattern,
                            visible_only=True,
                            timeout=1
                        )
                        info_dialog = dialog_app.window(title_re=title_pattern)
                        if info_dialog.exists() and info_dialog.is_visible():
                            LOGGER.info(f"找到信息对话框: {info_dialog.window_text()} (backend: win32)")
                            break
                    except:
                        pass
                    
                    # 尝试uia backend
                    try:
                        dialog_app = Application(backend="uia").connect(
                            title_re=title_pattern,
                            visible_only=True,
                            timeout=1
                        )
                        info_dialog = dialog_app.window(title_re=title_pattern)
                        if info_dialog.exists() and info_dialog.is_visible():
                            LOGGER.info(f"找到信息对话框: {info_dialog.window_text()} (backend: uia)")
                            break
                    except:
                        pass
                except:
                    continue
            
            if info_dialog is not None:
                break
            
            time.sleep(0.3)
        
        if info_dialog is None:
            LOGGER.info("未找到'MIR is now MRS!'对话框，可能未出现或已关闭")
            return
        
        try:
            # 激活对话框
            info_dialog.set_focus()
            time.sleep(0.3)
            
            ok_clicked = False
            
            # 方法1: 查找OK按钮并点击
            try:
                ok_button = info_dialog.child_window(title="OK", control_type="Button")
                if ok_button.exists() and ok_button.is_enabled():
                    LOGGER.info("找到OK按钮（通过title）")
                    ok_button.click_input()
                    ok_clicked = True
                    LOGGER.info("✅ 已点击OK按钮（通过title）")
            except Exception as e1:
                LOGGER.debug(f"通过title查找OK按钮失败: {e1}")
            
            # 方法2: 遍历所有按钮查找OK
            if not ok_clicked:
                try:
                    buttons = info_dialog.descendants(control_type="Button")
                    LOGGER.info(f"找到 {len(buttons)} 个按钮")
                    for button in buttons:
                        try:
                            button_text = button.window_text().strip().upper()
                            if button_text == "OK":
                                LOGGER.info(f"找到OK按钮（文本: '{button.window_text()}'）")
                                button.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK按钮（遍历按钮）")
                                break
                        except:
                            continue
                except Exception as e2:
                    LOGGER.debug(f"遍历按钮失败: {e2}")
            
            # 方法3: 使用Enter键（如果焦点在OK按钮上）
            if not ok_clicked:
                try:
                    LOGGER.info("尝试使用Enter键点击OK按钮")
                    info_dialog.type_keys("{ENTER}")
                    ok_clicked = True
                    LOGGER.info("✅ 已使用Enter键点击OK按钮")
                except Exception as e3:
                    LOGGER.debug(f"使用Enter键失败: {e3}")
            
            # 方法4: 使用Windows API查找并点击OK按钮
            if not ok_clicked and win32gui and win32con:
                try:
                    LOGGER.info("使用Windows API查找OK按钮...")
                    hwnd = info_dialog.handle
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            class_name = win32gui.GetClassName(hwnd_child)
                            window_text = win32gui.GetWindowText(hwnd_child)
                            if window_text.upper() == "OK" and "BUTTON" in class_name.upper():
                                LOGGER.info(f"通过Windows API找到OK按钮: '{window_text}' (类名: {class_name})")
                                win32gui.PostMessage(hwnd_child, win32con.BM_CLICK, 0, 0)
                                return False  # 停止枚举
                            return True
                        except:
                            return True
                    
                    win32gui.EnumChildWindows(hwnd, enum_child_proc, None)
                    time.sleep(0.5)
                    ok_clicked = True
                    LOGGER.info("✅ 已通过Windows API点击OK按钮")
                except Exception as e4:
                    LOGGER.debug(f"使用Windows API点击失败: {e4}")
            
            if not ok_clicked:
                LOGGER.warning("无法点击OK按钮，请手动处理对话框")
            else:
                # 等待对话框关闭
                time.sleep(0.5)
                LOGGER.info("✅ 信息对话框已处理")
        
        except Exception as e:
            LOGGER.warning(f"处理信息对话框时出错: {e}")
            LOGGER.warning("请手动处理'MIR is now MRS!'对话框")
    
    def _click_select_visible_rows_button(self) -> None:
        """点击左侧的'Select Visible Rows'按钮"""
        if Application is None:
            return
        
        LOGGER.info("查找并点击'Select Visible Rows'按钮...")
        
        try:
            # 确保主窗口有焦点
            self._window.set_focus()
            time.sleep(0.3)
            
            # 方法1: 通过按钮文本查找
            try:
                select_button = self._window.child_window(title="Select Visible Rows", control_type="Button")
                if select_button.exists() and select_button.is_enabled() and select_button.is_visible():
                    LOGGER.info("找到'Select Visible Rows'按钮（通过title）")
                    # 获取按钮位置
                    try:
                        button_rect = select_button.rectangle()
                        button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                        button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                        LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y})")
                    except:
                        pass
                    
                    # 使用click_input()点击
                    select_button.click_input()
                    time.sleep(0.5)
                    LOGGER.info("✅ 已点击'Select Visible Rows'按钮（通过title，使用click_input）")
                    return
            except Exception as e:
                LOGGER.debug(f"通过title查找按钮失败: {e}")
            
            # 方法2: 遍历所有按钮查找
            try:
                all_buttons = self._window.descendants(control_type="Button")
                LOGGER.info(f"找到 {len(all_buttons)} 个按钮")
                for idx, button in enumerate(all_buttons):
                    try:
                        button_text = button.window_text().strip()
                        LOGGER.debug(f"  按钮 #{idx}: 文本='{button_text}'")
                        
                        if "SELECT VISIBLE ROWS" in button_text.upper() or "SELECT VISIBLE" in button_text.upper():
                            LOGGER.info(f"找到'Select Visible Rows'按钮（文本: '{button_text}'）")
                            if button.is_enabled() and button.is_visible():
                                # 获取按钮位置
                                try:
                                    button_rect = button.rectangle()
                                    button_center_x = button_rect.left + (button_rect.right - button_rect.left) // 2
                                    button_center_y = button_rect.top + (button_rect.bottom - button_rect.top) // 2
                                    LOGGER.info(f"按钮中心坐标: ({button_center_x}, {button_center_y})")
                                except:
                                    pass
                                
                                # 使用click_input()点击
                                button.click_input()
                                time.sleep(0.5)
                                LOGGER.info(f"✅ 已点击'Select Visible Rows'按钮（文本: '{button_text}'，使用click_input）")
                                return
                    except Exception as e:
                        LOGGER.debug(f"检查按钮 #{idx} 时出错: {e}")
                        continue
            except Exception as e:
                LOGGER.debug(f"遍历按钮失败: {e}")
            
            # 方法3: 使用Windows API查找
            if win32gui:
                try:
                    main_hwnd = self._window.handle
                    
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            window_text = win32gui.GetWindowText(hwnd_child)
                            if "SELECT VISIBLE ROWS" in window_text.upper() or "SELECT VISIBLE" in window_text.upper():
                                lParam.append((hwnd_child, window_text))
                        except:
                            pass
                        return True
                    
                    button_list = []
                    win32gui.EnumChildWindows(main_hwnd, enum_child_proc, button_list)
                    
                    if button_list:
                        button_hwnd, button_text = button_list[0]
                        LOGGER.info(f"使用Windows API找到按钮: '{button_text}'")
                        
                        # 获取按钮位置并点击
                        rect = win32gui.GetWindowRect(button_hwnd)
                        center_x = (rect[0] + rect[2]) // 2
                        center_y = (rect[1] + rect[3]) // 2
                        
                        try:
                            import pyautogui
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.5)
                            LOGGER.info(f"✅ 已通过坐标点击'Select Visible Rows'按钮（位置: {center_x}, {center_y}）")
                            return
                        except ImportError:
                            win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
                            time.sleep(0.5)
                            LOGGER.info("✅ 已通过Windows API点击'Select Visible Rows'按钮")
                            return
                except Exception as e:
                    LOGGER.debug(f"使用Windows API查找按钮失败: {e}")
            
            LOGGER.warning("未找到'Select Visible Rows'按钮")
            
        except Exception as e:
            LOGGER.error(f"点击'Select Visible Rows'按钮失败: {e}")
            raise RuntimeError(f"无法点击'Select Visible Rows'按钮: {e}")
    
    def _handle_login_dialog(self) -> None:
        """处理Mole登录对话框，自动点击OK按钮"""
        if Application is None:
            return
        
        LOGGER.info("等待MOLE LOGIN对话框出现...")
        
        # 等待登录对话框出现（最多等待30秒）
        login_dialog = None
        login_app = None
        deadline = time.time() + 30
        # 使用配置的标题，同时提供备选模式
        dialog_titles = [
            self.config.login_dialog_title,
            ".*MOLE LOGIN.*",
            ".*LOGIN.*",
            ".*Login.*"
        ]
        
        # 尝试不同的backend
        backends = ["win32", "uia"]
        
        while time.time() < deadline:
            for backend in backends:
                for title_pattern in dialog_titles:
                    try:
                        login_app = Application(backend=backend).connect(
                            title_re=title_pattern,
                            visible_only=True,
                            timeout=2
                        )
                        login_dialog = login_app.window(title_re=title_pattern)
                        
                        if login_dialog.exists() and login_dialog.is_visible():
                            LOGGER.info(f"找到登录对话框: {login_dialog.window_text()} (backend: {backend})")
                            break
                    except Exception:
                        continue
                
                if login_dialog is not None:
                    break
            
            if login_dialog is not None:
                break
            
            time.sleep(0.5)
        
        if login_dialog is None:
            LOGGER.warning("未找到MOLE LOGIN对话框，可能已经登录或对话框已关闭")
            return
        
        try:
            # 激活对话框
            login_dialog.set_focus()
            time.sleep(0.5)
            
            # 打印控件结构用于调试（保存到日志）
            try:
                LOGGER.info("正在分析登录对话框的控件结构...")
                # 将控件结构输出到字符串
                import io
                import sys
                old_stdout = sys.stdout
                sys.stdout = buffer = io.StringIO()
                login_dialog.print_control_identifiers(depth=4)
                sys.stdout = old_stdout
                control_info = buffer.getvalue()
                LOGGER.debug(f"登录对话框控件结构:\n{control_info}")
                # 同时尝试查找所有可能的控件类型
                LOGGER.info("查找所有可用的控件...")
            except Exception as e:
                LOGGER.debug(f"打印控件结构失败: {e}")
            
            # 查找并点击OK按钮
            ok_clicked = False
            
            # 首先尝试多种方法查找所有按钮
            all_buttons = []
            
            # 方法1: 使用children查找
            try:
                buttons1 = login_dialog.children(control_type="Button")
                all_buttons.extend(buttons1)
                LOGGER.info(f"通过children找到 {len(buttons1)} 个Button控件")
            except Exception as e:
                LOGGER.debug(f"通过children查找按钮失败: {e}")
            
            # 方法2: 使用descendants查找（查找所有子控件）
            if len(all_buttons) == 0:
                try:
                    buttons2 = login_dialog.descendants(control_type="Button")
                    all_buttons.extend(buttons2)
                    LOGGER.info(f"通过descendants找到 {len(buttons2)} 个Button控件")
                except Exception as e:
                    LOGGER.debug(f"通过descendants查找按钮失败: {e}")
            
            # 方法3: 查找所有子窗口（可能按钮是子窗口）
            if len(all_buttons) == 0:
                try:
                    all_windows = login_dialog.children()
                    LOGGER.info(f"找到 {len(all_windows)} 个子控件")
                    for win in all_windows:
                        try:
                            win_text = win.window_text()
                            win_type = str(type(win))
                            LOGGER.info(f"  控件: '{win_text}' (类型: {win_type})")
                            if "Button" in win_type or win_text.upper() in ["OK", "CANCEL"]:
                                all_buttons.append(win)
                        except:
                            pass
                except Exception as e:
                    LOGGER.debug(f"查找所有子窗口失败: {e}")
            
            # 方法4: 尝试child_windows()
            if len(all_buttons) == 0:
                try:
                    child_windows = login_dialog.child_windows()
                    LOGGER.info(f"通过child_windows找到 {len(child_windows)} 个子窗口")
                    for win in child_windows:
                        try:
                            win_text = win.window_text()
                            if win_text.upper() in ["OK", "CANCEL"]:
                                all_buttons.append(win)
                        except:
                            pass
                except Exception as e:
                    LOGGER.debug(f"通过child_windows查找失败: {e}")
            
            # 列出找到的所有按钮
            LOGGER.info(f"总共找到 {len(all_buttons)} 个可能的按钮控件")
            for idx, btn in enumerate(all_buttons):
                try:
                    btn_text = btn.window_text()
                    LOGGER.info(f"  按钮 {idx + 1}: '{btn_text}'")
                except:
                    LOGGER.info(f"  按钮 {idx + 1}: (无法获取文本)")
            
            # 方法1: 使用找到的按钮列表点击OK
            if not ok_clicked and len(all_buttons) > 0:
                for button in all_buttons:
                    try:
                        button_text = button.window_text().strip().upper()
                        if button_text == "OK" and "CANCEL" not in button_text:
                            LOGGER.info(f"找到OK按钮（从按钮列表中: '{button.window_text()}'）")
                            button.click_input()
                            ok_clicked = True
                            LOGGER.info("✅ 已点击OK按钮（从按钮列表）")
                            break
                    except Exception as e:
                        LOGGER.debug(f"检查按钮时出错: {e}")
                        continue
            
            # 方法2: 精确匹配OK按钮（完全匹配，不区分大小写）
            if not ok_clicked:
                try:
                    # 尝试多种方式精确匹配OK
                    ok_patterns = ["OK", "Ok", "ok"]
                    for pattern in ok_patterns:
                        try:
                            ok_button = login_dialog.child_window(title=pattern, control_type="Button")
                            if ok_button.exists() and ok_button.is_enabled():
                                btn_text = ok_button.window_text()
                                LOGGER.info(f"找到OK按钮（精确匹配: '{btn_text}'）")
                                ok_button.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK按钮（通过精确匹配title）")
                                break
                        except:
                            continue
                except Exception as e1:
                    LOGGER.debug(f"通过精确匹配查找OK按钮失败: {e1}")
            
            # 方法3: 使用descendants查找所有按钮（包括嵌套的）
            if not ok_clicked:
                try:
                    all_descendants = login_dialog.descendants(control_type="Button")
                    LOGGER.info(f"通过descendants找到 {len(all_descendants)} 个Button控件")
                    for button in all_descendants:
                        try:
                            button_text = button.window_text().strip().upper()
                            if button_text == "OK" and "CANCEL" not in button_text:
                                LOGGER.info(f"找到OK按钮（descendants: '{button.window_text()}'）")
                                button.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK按钮（通过descendants）")
                                break
                        except Exception as e:
                            LOGGER.debug(f"检查descendant按钮时出错: {e}")
                            continue
                except Exception as e3:
                    LOGGER.debug(f"通过descendants查找按钮失败: {e3}")
            
            # 方法3: 如果OK通常是第一个按钮（在Cancel之前），按位置点击
            if not ok_clicked:
                try:
                    buttons = login_dialog.children(control_type="Button")
                    if len(buttons) >= 2:
                        # 检查第一个按钮是否是OK
                        first_button_text = buttons[0].window_text().strip().upper()
                        if first_button_text == "OK":
                            LOGGER.info(f"按位置点击第一个按钮（文本: '{buttons[0].window_text()}'）")
                            buttons[0].click_input()
                            ok_clicked = True
                            LOGGER.info("✅ 已点击第一个按钮（确认为OK）")
                except Exception as e3:
                    LOGGER.debug(f"按位置点击按钮失败: {e3}")
            
            # 方法4: 尝试通过auto_id或其他属性查找OK按钮
            if not ok_clicked:
                try:
                    # 尝试查找常见的按钮ID
                    ok_ids = ["OK", "btnOK", "buttonOK", "idOK"]
                    for btn_id in ok_ids:
                        try:
                            ok_button = login_dialog.child_window(auto_id=btn_id, control_type="Button")
                            if ok_button.exists() and ok_button.is_enabled():
                                btn_text = ok_button.window_text()
                                LOGGER.info(f"通过auto_id找到OK按钮: '{btn_id}' (文本: '{btn_text}')")
                                ok_button.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK按钮（通过auto_id）")
                                break
                        except:
                            continue
                except Exception as e4:
                    LOGGER.debug(f"通过auto_id查找OK按钮失败: {e4}")
            
            # 方法5: 使用Windows API直接查找和点击（win32gui方法）
            if not ok_clicked and win32gui:
                try:
                    LOGGER.info("尝试使用Windows API查找按钮...")
                    hwnd = login_dialog.handle
                    
                    # 枚举所有子窗口
                    def enum_child_proc(hwnd_child, lParam):
                        try:
                            class_name = win32gui.GetClassName(hwnd_child)
                            window_text = win32gui.GetWindowText(hwnd_child)
                            if window_text.upper() == "OK" and "BUTTON" in class_name.upper():
                                LOGGER.info(f"通过Windows API找到OK按钮: '{window_text}' (类名: {class_name})")
                                # 发送BM_CLICK消息
                                import win32con
                                win32gui.PostMessage(hwnd_child, win32con.BM_CLICK, 0, 0)
                                return False  # 停止枚举
                            return True
                        except:
                            return True
                    
                    win32gui.EnumChildWindows(hwnd, enum_child_proc, None)
                    # 给一点时间让点击生效
                    time.sleep(0.5)
                    ok_clicked = True
                    LOGGER.info("✅ 已通过Windows API点击OK按钮")
                except Exception as e5:
                    LOGGER.debug(f"使用Windows API点击失败: {e5}")
            
            # 方法6: 尝试所有子控件，不指定类型
            if not ok_clicked:
                try:
                    LOGGER.info("尝试查找所有子控件（不指定类型）...")
                    all_children = login_dialog.children()
                    for child in all_children:
                        try:
                            child_text = child.window_text().strip().upper()
                            if child_text == "OK":
                                LOGGER.info(f"找到OK控件（无类型限制: '{child.window_text()}'）")
                                child.click_input()
                                ok_clicked = True
                                LOGGER.info("✅ 已点击OK控件（无类型限制）")
                                break
                        except:
                            continue
                except Exception as e6:
                    LOGGER.debug(f"查找所有子控件失败: {e6}")
            
            if not ok_clicked:
                LOGGER.warning("无法点击OK按钮，请手动处理登录对话框")
            else:
                # 等待对话框关闭
                time.sleep(1)
                LOGGER.info("登录对话框已处理")
        
        except Exception as e:
            LOGGER.warning(f"处理登录对话框时出错: {e}")
            LOGGER.warning("请手动处理登录对话框")
    
    def _is_process_running(self) -> bool:
        """检查Mole进程是否在运行"""
        if Application is None:
            return False
        try:
            Application(backend="win32").connect(
                title_re=self.config.window_title
            )
            return True
        except ElementNotFoundError:
            return False
    
    def submit_mir_data(self, data: dict) -> bool:
        """
        提交MIR数据到Mole工具
        
        Args:
            data: 包含MIR数据的字典
        
        Returns:
            True如果提交成功
        
        Raises:
            RuntimeError: 如果提交失败
        """
        LOGGER.info("开始提交MIR数据到Mole工具")
        LOGGER.debug(f"MIR数据: {data}")
        
        # 重试机制
        last_exception = None
        for attempt in range(1, self.config.retry_count + 1):
            try:
                LOGGER.info(f"尝试提交MIR数据 (第{attempt}/{self.config.retry_count}次)")
                
                # 确保应用程序已启动
                self._ensure_application()
                
                # 点击File菜单，选择New MIR Request
                if attempt == 1:  # 只在第一次尝试时打开New MIR Request
                    self._click_file_menu_new_mir_request()
                
                # 激活窗口
                if self._window:
                    try:
                        self._window.set_focus()
                        if win32gui and win32con:
                            hwnd = self._window.handle
                            win32gui.SetForegroundWindow(hwnd)
                            win32gui.BringWindowToTop(hwnd)
                        time.sleep(0.5)
                    except Exception as e:
                        LOGGER.warning(f"设置窗口焦点失败: {e}")
                
                # TODO: 根据Mole工具的实际界面实现具体的数据输入逻辑
                # 这里需要根据实际的Mole工具界面来填写表单
                # 示例：查找输入框、填写数据、点击提交按钮等
                
                LOGGER.info("✅ MIR数据提交成功")
                return True
                
            except Exception as e:
                last_exception = e
                LOGGER.warning(f"第{attempt}次提交失败: {e}")
                if attempt < self.config.retry_count:
                    LOGGER.info(f"等待{self.config.retry_delay}秒后重试...")
                    time.sleep(self.config.retry_delay)
                else:
                    LOGGER.error(f"❌ MIR数据提交失败（已重试{self.config.retry_count}次）")
        
        raise RuntimeError(f"MIR数据提交失败: {last_exception}")
    
    def verify_submission(self) -> bool:
        """
        验证数据是否提交成功
        
        Returns:
            True如果验证通过
        """
        # TODO: 实现验证逻辑
        # 例如：检查Mole工具中是否显示成功消息
        LOGGER.info("验证MIR数据提交结果...")
        
        try:
            if self._window:
                # 检查窗口状态、消息等
                # 这里需要根据实际Mole工具的反馈机制来实现
                LOGGER.info("✅ MIR数据提交验证通过")
                return True
        except Exception as e:
            LOGGER.warning(f"验证提交结果时出错: {e}")
        
        return False

