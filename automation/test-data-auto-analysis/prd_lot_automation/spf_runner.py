import logging
import os
import subprocess
import time
from pathlib import Path
from typing import Sequence

try:
    from pywinauto import Application
    from pywinauto.findwindows import ElementNotFoundError
except ImportError:  # pragma: no cover
    Application = None  # type: ignore
    ElementNotFoundError = Exception  # type: ignore

try:
    import pyautogui
except ImportError:  # pragma: no cover
    pyautogui = None

try:
    import win32gui
    import win32con
    import win32process
    import win32api
except ImportError:  # pragma: no cover
    win32gui = None
    win32con = None
    win32process = None
    win32api = None

from .config_loader import AppConfig


LOGGER = logging.getLogger(__name__)


class SQLPathFinderRunner:
    def __init__(self, config: AppConfig):
        self.config = config
        self._app = None
        self._window = None

    def _close_existing_windows(self) -> None:
        """在启动新的 VG2 之前，关闭所有现有的 SQLPathFinder 窗口"""
        if Application is None:
            return
        
        try:
            LOGGER.info("检查并关闭现有的 SQLPathFinder 窗口...")
            from .close_sqlpathfinder import close_sqlpathfinder
            close_sqlpathfinder(self.config.ui.main_window_title)
            time.sleep(1)  # 等待窗口完全关闭
        except Exception as e:
            LOGGER.warning(f"关闭现有窗口时出错（可能没有窗口）: {e}")

    def _handle_update_dialog(self) -> bool:
        """
        检查并处理 SQLPathFinder 更新对话框
        
        Returns:
            True 如果处理了更新对话框（需要重新打开 SQLPathFinder），False 如果没有更新对话框
        """
        if Application is None:
            return False
        
        try:
            # 尝试查找更新对话框
            # 对话框标题是 "Update Recommended"
            update_patterns = [
                "Update Recommended",
                ".*Update.*Recommended.*",
                ".*Update.*",
            ]
            
            for pattern in update_patterns:
                try:
                    # 使用更短的超时时间，避免阻塞太久
                    app = Application(backend="win32").connect(title_re=pattern, timeout=1)
                    dialog = app.window(title_re=pattern)
                    
                    if dialog.exists() and dialog.is_visible():
                        dialog_text = dialog.window_text()
                        LOGGER.info(f"✅ 检测到更新对话框: {dialog_text}")
                        
                        # 确保对话框在最前面并获得焦点
                        try:
                            if win32gui and win32con:
                                hwnd = dialog.handle
                                # 将对话框置于最前
                                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                                win32gui.SetForegroundWindow(hwnd)
                                time.sleep(0.3)
                                # 恢复为非最顶层（但仍在前面）
                                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                            else:
                                dialog.set_focus()
                                time.sleep(0.3)
                        except Exception as e:
                            LOGGER.warning(f"设置对话框焦点失败: {e}")
                            try:
                                dialog.set_focus()
                                time.sleep(0.3)
                            except:
                                pass
                        
                        # 点击 Yes 按钮（必须点击 Yes，不能用 Enter 键，因为 Enter 可能触发 No）
                        yes_clicked = False
                        
                        # 方法1: 直接点击第一个按钮（从测试得知，第一个按钮就是 Yes）
                        try:
                            LOGGER.info("尝试点击第一个按钮（Yes）...")
                            buttons = dialog.children(class_name="Button")
                            LOGGER.info(f"找到 {len(buttons)} 个按钮")
                            
                            if buttons:
                                first_btn = buttons[0]
                                btn_text = first_btn.window_text()
                                LOGGER.info(f"第一个按钮文本: '{btn_text}'")
                                
                                if "Yes" in btn_text:
                                    LOGGER.info("准备点击 Yes 按钮...")
                                    first_btn.click_input()
                                    LOGGER.info("✅ 已点击 Yes 按钮")
                                    yes_clicked = True
                                else:
                                    LOGGER.warning(f"第一个按钮不是 Yes（文本: {btn_text}），尝试其他方法...")
                            else:
                                LOGGER.warning("未找到任何按钮")
                        except Exception as e:
                            LOGGER.warning(f"点击第一个按钮失败: {e}")
                        
                        # 方法2: 通过标题 "Yes" 或 "&Yes" 查找按钮
                        if not yes_clicked:
                            try:
                                LOGGER.info("尝试通过标题查找 Yes 按钮...")
                                # 尝试 "Yes"
                                try:
                                    yes_btn = dialog.child_window(title="Yes", control_type="Button")
                                    if yes_btn.exists():
                                        LOGGER.info("找到 Yes 按钮（title='Yes'）")
                                        yes_btn.click_input()
                                        LOGGER.info("✅ 已点击 Yes 按钮")
                                        yes_clicked = True
                                except:
                                    pass
                                
                                # 尝试 "&Yes"
                                if not yes_clicked:
                                    try:
                                        yes_btn = dialog.child_window(title="&Yes", control_type="Button")
                                        if yes_btn.exists():
                                            LOGGER.info("找到 Yes 按钮（title='&Yes'）")
                                            yes_btn.click_input()
                                            LOGGER.info("✅ 已点击 Yes 按钮")
                                            yes_clicked = True
                                    except:
                                        pass
                            except Exception as e:
                                LOGGER.warning(f"通过标题查找失败: {e}")
                        
                        # 方法3: 遍历所有按钮，查找包含 "Yes" 的按钮
                        if not yes_clicked:
                            try:
                                LOGGER.info("遍历所有按钮，查找 Yes...")
                                buttons = dialog.children(class_name="Button")
                                for idx, btn in enumerate(buttons):
                                    try:
                                        btn_text = btn.window_text()
                                        LOGGER.info(f"  按钮 {idx}: '{btn_text}'")
                                        if "Yes" in btn_text or "yes" in btn_text.lower():
                                            LOGGER.info(f"尝试点击按钮 {idx}...")
                                            btn.click_input()
                                            LOGGER.info(f"✅ 已点击按钮 {idx}")
                                            yes_clicked = True
                                            break
                                    except Exception as e:
                                        LOGGER.warning(f"处理按钮 {idx} 失败: {e}")
                            except Exception as e:
                                LOGGER.warning(f"遍历按钮失败: {e}")
                        
                        # 方法4: 使用 pyautogui 模拟鼠标点击
                        if not yes_clicked and pyautogui:
                            try:
                                LOGGER.info("尝试使用 pyautogui 点击 Yes 按钮...")
                                buttons = dialog.children(class_name="Button")
                                if buttons:
                                    first_btn = buttons[0]
                                    btn_rect = first_btn.rectangle()
                                    # 计算按钮中心位置
                                    center_x = (btn_rect.left + btn_rect.right) // 2
                                    center_y = (btn_rect.top + btn_rect.bottom) // 2
                                    LOGGER.info(f"按钮位置: ({center_x}, {center_y})")
                                    
                                    # 在远程桌面环境中，临时禁用 fail-safe
                                    original_failsafe = pyautogui.FAILSAFE
                                    try:
                                        pyautogui.FAILSAFE = False
                                        # 移动鼠标并点击
                                        pyautogui.moveTo(center_x, center_y)
                                        time.sleep(0.2)
                                        pyautogui.click()
                                        LOGGER.info("✅ 已通过 pyautogui 点击 Yes 按钮（已禁用 fail-safe）")
                                        yes_clicked = True
                                    finally:
                                        # 恢复原始设置
                                        pyautogui.FAILSAFE = original_failsafe
                            except Exception as e:
                                LOGGER.warning(f"pyautogui 点击失败: {e}")
                        
                        if not yes_clicked:
                            LOGGER.error("❌ 所有方法都失败，无法点击 Yes 按钮")
                            LOGGER.error("请手动点击 Yes 按钮，然后脚本会继续执行")
                            # 给用户 30 秒时间手动点击
                            manual_deadline = time.time() + 30
                            while time.time() < manual_deadline:
                                try:
                                    if not dialog.exists() or not dialog.is_visible():
                                        LOGGER.info("✅ 对话框已关闭（用户可能已手动点击）")
                                        yes_clicked = True
                                        break
                                except:
                                    yes_clicked = True
                                    break
                                time.sleep(1)
                            
                            if not yes_clicked:
                                LOGGER.error("❌ 超时，对话框仍未关闭")
                                return False
                        
                        # 运行 StartSQLPathFinder.bat 来触发升级（不等待对话框关闭，立即运行）
                        LOGGER.info("立即运行 StartSQLPathFinder.bat 来触发升级...")
                        
                        # 从配置文件读取升级脚本路径，如果没有配置则使用默认路径
                        if self.config.paths.upgrade_script:
                            upgrade_bat = self.config.paths.upgrade_script
                        else:
                            # 默认路径
                            upgrade_bat = Path(r"C:\Users\qibaiche\My Programs\SQLPathFinder3\StartSQLPathFinder.bat")
                        
                        LOGGER.info(f"升级脚本路径: {upgrade_bat}")
                        
                        if upgrade_bat.exists():
                            try:
                                LOGGER.info(f"找到升级脚本: {upgrade_bat}")
                                # 使用 subprocess 运行 bat 文件
                                subprocess.Popen([str(upgrade_bat)], shell=True)
                                LOGGER.info("✅ 升级脚本已启动")
                                
                                # 等待新的 "Update SQLPathFinder?" 窗口出现
                                LOGGER.info("等待 'Update SQLPathFinder?' 窗口出现...")
                                time.sleep(3)  # 给窗口一点时间出现
                                
                                # 查找并点击 "Update" 按钮
                                update_dialog = None
                                update_patterns = [
                                    "Update SQLPathFinder?",
                                    ".*Update.*SQLPathFinder.*",
                                ]
                                
                                deadline = time.time() + 30
                                while time.time() < deadline:
                                    for pattern in update_patterns:
                                        try:
                                            app = Application(backend="win32").connect(title_re=pattern, timeout=1)
                                            test_dialog = app.window(title_re=pattern)
                                            
                                            if test_dialog.exists() and test_dialog.is_visible():
                                                update_dialog = test_dialog
                                                LOGGER.info(f"✅ 找到升级窗口: {update_dialog.window_text()}")
                                                break
                                        except:
                                            continue
                                    
                                    if update_dialog:
                                        break
                                    time.sleep(1)
                                
                                if update_dialog:
                                    # 确保窗口获得焦点
                                    try:
                                        if win32gui and win32con:
                                            hwnd = update_dialog.handle
                                            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                                                 win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                                            win32gui.SetForegroundWindow(hwnd)
                                            time.sleep(0.3)
                                            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                                                 win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                                        else:
                                            update_dialog.set_focus()
                                        time.sleep(0.5)
                                    except:
                                        pass
                                    
                                    # 点击 Update 按钮
                                    update_clicked = False
                                    
                                    # 查找并点击 Update 按钮
                                    try:
                                        all_children = update_dialog.descendants()
                                        
                                        for child in all_children:
                                            try:
                                                child_class = child.class_name()
                                                child_text = child.window_text()
                                                
                                                # 只查找按钮控件，且文本精确匹配 "Update"
                                                if "BUTTON" in child_class.upper() and child_text.strip() == "Update":
                                                    child.click_input()
                                                    LOGGER.info("✅ 已点击 Update 按钮")
                                                    update_clicked = True
                                                    break
                                            except:
                                                continue
                                    except Exception as e:
                                        LOGGER.warning(f"点击 Update 按钮失败: {e}")
                                    
                                    if not update_clicked:
                                        LOGGER.warning("⚠️ 未能自动点击 Update 按钮，请手动点击")
                                        time.sleep(10)
                                    else:
                                        # 等待升级完成（5 秒）
                                        LOGGER.info("等待 SQLPathFinder 升级完成（5 秒）...")
                                        time.sleep(5)
                                        LOGGER.info("✅ 升级等待完成")
                                else:
                                    LOGGER.warning("⚠️ 未找到 'Update SQLPathFinder?' 窗口")
                                    LOGGER.warning("将等待 5 秒后继续...")
                                    time.sleep(5)
                                
                            except Exception as e:
                                LOGGER.error(f"❌ 运行升级脚本失败: {e}")
                                import traceback
                                LOGGER.error(traceback.format_exc())
                        else:
                            LOGGER.warning(f"⚠️ 未找到升级脚本: {upgrade_bat}")
                            LOGGER.warning("将尝试继续执行，但升级可能未完成")
                        
                        # 关闭所有 SQLPathFinder 窗口（确保完全关闭）
                        LOGGER.info("关闭所有 SQLPathFinder 窗口...")
                        from .close_sqlpathfinder import close_sqlpathfinder
                        close_sqlpathfinder(self.config.ui.main_window_title)
                        time.sleep(2)  # 等待完全关闭
                        
                        # 再次确认所有进程已关闭
                        LOGGER.info("确认所有 SQLPathFinder 进程已关闭...")
                        try:
                            import psutil
                            for proc in psutil.process_iter(['pid', 'name']):
                                try:
                                    if 'SQLPathFinder' in proc.info['name']:
                                        LOGGER.info(f"终止残留进程: {proc.info['name']} (PID: {proc.info['pid']})")
                                        proc.kill()
                                        time.sleep(0.5)
                                except (psutil.NoSuchProcess, psutil.AccessDenied):
                                    pass
                        except:
                            pass
                        
                        # 等待一下，确保升级完成
                        LOGGER.info("等待 3 秒，确保升级完成...")
                        time.sleep(3)
                        
                        LOGGER.info("✅ 更新对话框已处理，需要重新打开 SQLPathFinder")
                        return True
                except ElementNotFoundError:
                    continue
                except Exception as e:
                    LOGGER.debug(f"检查更新对话框时出错（可能不存在）: {e}")
                    continue
            
            return False
        except Exception as e:
            LOGGER.debug(f"检查更新对话框时出错: {e}")
            return False

    def _ensure_application(self) -> None:
        if Application is None:
            raise RuntimeError("pywinauto 未安装，无法执行 UI 自动化")

        if self._window:
            return

        # 在启动新窗口之前，先关闭所有现有窗口
        self._close_existing_windows()

        # 检查是否有更新对话框（可能在已打开的 SQLPathFinder 中）
        if self._is_process_running():
            LOGGER.info("检测到 SQLPathFinder 正在运行，检查是否有更新对话框...")
            if self._handle_update_dialog():
                # 如果处理了更新对话框，需要重新打开 SQLPathFinder
                LOGGER.info("更新对话框已处理，等待后重新打开 SQLPathFinder...")
                time.sleep(2)

        if self.config.paths.spf_executable and not self._is_process_running():
            LOGGER.info("启动 SQLPathFinder：%s", self.config.paths.spf_executable)
            subprocess.Popen(
                [str(self.config.paths.spf_executable), str(self.config.paths.vg2_file)]
            )
        else:
            LOGGER.info("直接打开 VG2：%s", self.config.paths.vg2_file)
            os.startfile(self.config.paths.vg2_file)

        # 在等待连接之前，再次检查更新对话框（可能在打开 VG2 时弹出）
        time.sleep(2)  # 给一点时间让对话框出现
        if self._handle_update_dialog():
            # 如果处理了更新对话框，需要重新打开 SQLPathFinder
            LOGGER.info("更新对话框已处理，重新打开 SQLPathFinder...")
            time.sleep(2)
            if self.config.paths.spf_executable:
                LOGGER.info("重新启动 SQLPathFinder：%s", self.config.paths.spf_executable)
                subprocess.Popen(
                    [str(self.config.paths.spf_executable), str(self.config.paths.vg2_file)]
                )
            else:
                LOGGER.info("重新打开 VG2：%s", self.config.paths.vg2_file)
                os.startfile(self.config.paths.vg2_file)
            time.sleep(2)  # 等待重新打开

        deadline = time.time() + self.config.timeouts.spf_launch
        while time.time() < deadline:
            try:
                # 使用 win32 backend（SQLPathFinder 是 WindowsForms 应用）
                self._app = Application(backend="win32").connect(
                    title_re=self.config.ui.main_window_title,
                    visible_only=True  # 只连接可见窗口
                )
                
                # 获取所有匹配的窗口
                windows = self._app.windows()
                LOGGER.info(f"找到 {len(windows)} 个 SQLPathFinder 窗口")
                
                # 选择最新打开的窗口（通常是第一个）
                # 如果有多个窗口，选择包含当前VG2文件名的窗口
                target_window = None
                vg2_name = self.config.paths.vg2_file.stem  # 获取文件名（不含扩展名）
                
                for window in windows:
                    try:
                        window_title = window.window_text()
                        LOGGER.debug(f"检查窗口: {window_title}")
                        if vg2_name in window_title:
                            target_window = window
                            LOGGER.info(f"✅ 找到匹配的窗口: {window_title}")
                            break
                    except:
                        continue
                
                # 如果没找到匹配的，使用第一个
                if target_window is None and windows:
                    target_window = windows[0]
                    LOGGER.info(f"使用第一个窗口: {target_window.window_text()}")
                
                if target_window:
                    self._window = target_window
                    LOGGER.info("已连接到 SQLPathFinder 主窗口 (win32 backend)")
                    return
                
                # 如果没有找到窗口，抛出异常
                raise ElementNotFoundError("未找到可用的窗口")
                    
            except ElementNotFoundError:
                time.sleep(1)
        raise TimeoutError("无法连接到 SQLPathFinder 主窗口")

    def _is_process_running(self) -> bool:
        if Application is None:
            return False
        try:
            Application(backend="win32").connect(
                title_re=self.config.ui.main_window_title
            )
            return True
        except ElementNotFoundError:
            return False

    def _click_run_button(self) -> None:
        """使用 F8 快捷键触发 Run（最可靠的方式）"""
        if self._window is None:
            raise RuntimeError("SQLPathFinder 窗口未连接")
        
        LOGGER.info("使用 F8 快捷键触发 Run...")
        
        # 确保窗口在前台
        hwnd = None
        try:
            if not self._window.is_visible():
                self._window.restore()
                time.sleep(0.5)
            
            if win32gui and win32con:
                try:
                    hwnd = self._window.handle
                    # 使用多种方法确保窗口获得焦点
                    # 方法1: 基本的窗口激活
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.BringWindowToTop(hwnd)
                    win32gui.SetForegroundWindow(hwnd)
                    
                    # 方法2: 使用 AttachThreadInput 强制获得焦点（在远程桌面中更可靠）
                    if win32process and win32api:
                        try:
                            # 获取当前线程ID和窗口线程ID
                            current_thread_id = win32api.GetCurrentThreadId()
                            window_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
                            
                            # 如果线程不同，附加线程输入
                            if current_thread_id != window_thread_id:
                                win32process.AttachThreadInput(current_thread_id, window_thread_id, True)
                                try:
                                    win32gui.SetForegroundWindow(hwnd)
                                    win32gui.BringWindowToTop(hwnd)
                                finally:
                                    win32process.AttachThreadInput(current_thread_id, window_thread_id, False)
                        except Exception as e:
                            LOGGER.debug(f"AttachThreadInput 失败（可能不需要）: {e}")
                    
                    time.sleep(0.5)  # 等待窗口激活
                except Exception as e:
                    LOGGER.warning(f"Windows API 设置焦点失败: {e}")
            
            # 尝试使用 pywinauto 设置焦点（可能在远程桌面中失败）
            try:
                self._window.set_focus()
                time.sleep(0.3)
            except Exception as e:
                LOGGER.warning(f"设置窗口焦点失败（可能在远程桌面中）: {e}")
                # 在远程桌面中，set_focus 可能失败，但不影响后续操作
        except Exception as e:
            LOGGER.warning(f"设置窗口焦点失败: {e}")
        
        # 等待窗口完全激活后再发送 F8
        time.sleep(1)
        
        # 方法1: 使用 PostMessage 发送键盘消息（异步，更适合远程桌面）
        f8_sent = False
        if win32gui and win32con and hwnd:
            try:
                # F8 键的虚拟键码
                VK_F8 = 0x77
                WM_KEYDOWN = 0x0100
                WM_KEYUP = 0x0101
                
                # 使用 PostMessage 而不是 SendMessage（异步消息，更可靠）
                # 先发送按下消息
                win32gui.PostMessage(hwnd, WM_KEYDOWN, VK_F8, 0)
                time.sleep(0.1)
                # 再发送释放消息
                win32gui.PostMessage(hwnd, WM_KEYUP, VK_F8, 0)
                LOGGER.info("✅ 已通过 Windows API PostMessage 发送 F8 快捷键")
                f8_sent = True
                time.sleep(0.2)  # 等待消息处理
            except Exception as e:
                LOGGER.warning(f"Windows API PostMessage 发送 F8 失败: {e}")
        
        # 方法2: 如果 PostMessage 失败，尝试 SendMessage
        if not f8_sent and win32gui and win32con and hwnd:
            try:
                VK_F8 = 0x77
                WM_KEYDOWN = 0x0100
                WM_KEYUP = 0x0101
                win32gui.SendMessage(hwnd, WM_KEYDOWN, VK_F8, 0)
                time.sleep(0.05)
                win32gui.SendMessage(hwnd, WM_KEYUP, VK_F8, 0)
                LOGGER.info("✅ 已通过 Windows API SendMessage 发送 F8 快捷键")
                f8_sent = True
            except Exception as e:
                LOGGER.warning(f"Windows API SendMessage 发送 F8 失败: {e}")
        
        # 方法3: 使用 pywinauto type_keys（备用方案）
        if not f8_sent:
            try:
                self._window.type_keys("{F8}")
                LOGGER.info("✅ 已通过 pywinauto 发送 F8 快捷键")
                f8_sent = True
            except Exception as e:
                LOGGER.warning(f"pywinauto 发送 F8 失败: {e}")
        
        # 方法4: 使用 pyautogui（最后备用，需要禁用 fail-safe）
        if not f8_sent and pyautogui:
            try:
                # 在远程桌面环境中，临时禁用 fail-safe
                original_failsafe = pyautogui.FAILSAFE
                try:
                    pyautogui.FAILSAFE = False
                    pyautogui.press('f8')
                    LOGGER.info("✅ 已通过 pyautogui 发送 F8（已禁用 fail-safe）")
                    f8_sent = True
                finally:
                    # 恢复原始设置
                    pyautogui.FAILSAFE = original_failsafe
            except Exception as e2:
                LOGGER.warning(f"pyautogui 发送 F8 失败: {e2}")
        
        if not f8_sent:
            raise RuntimeError("所有方法都无法发送 F8 快捷键")
        
        # 等待弹窗出现（增加等待时间，确保弹窗有时间出现）
        time.sleep(2.5)

    def _process_single_popup(self, lots: Sequence[str], popup_index: int = 1) -> bool:
        """处理单个弹窗，返回是否成功"""
        # 查找弹窗 - 使用多种方法
        dialog = None
        deadline = time.time() + self.config.timeouts.ui_action
        attempt = 0
        
        # 尝试多种标题模式
        patterns_to_try = [
            ".*Prompt.*Values.*",
            "Prompt For Values (in)",
            ".*Prompt.*",
            ".*Values.*in.*",
        ]
        
        while time.time() < deadline:
            attempt += 1
            LOGGER.info(f"第 {attempt} 次尝试查找弹窗 #{popup_index}...")
            
            # 方法1: 尝试多种标题模式
            for pattern in patterns_to_try:
                try:
                    app = Application(backend="win32").connect(title_re=pattern, timeout=1)
                    dialog = app.window(title_re=pattern)
                    
                    if dialog.exists() and dialog.is_visible():
                        LOGGER.info(f"✅ 找到弹窗 #{popup_index}: {dialog.window_text()}")
                        LOGGER.info(f"   使用模式: {pattern}")
                        dialog.set_focus()
                        break  # 跳出 for 循环
                except:
                    continue
            
            # 如果找到了，跳出 while 循环
            if dialog is not None:
                break
            
            time.sleep(1)
        
        # 检查是否找到弹窗
        if dialog is None:
            LOGGER.warning(f"❌ 超时后仍未找到弹窗 #{popup_index}")
            return False
        
        # 准备数据（使用 Windows 换行符 \r\n，模拟文本编辑器的复制格式）
        # 关键：Windows 文本编辑器使用 CRLF (\r\n)，而不是 LF (\n)
        payload = "\r\n".join(lots)
        LOGGER.debug(f"弹窗 #{popup_index}: 准备粘贴 {len(lots)} 个 LOT")
        LOGGER.debug(f"   格式验证: 包含 {payload.count(chr(13))} 个回车符 (CR) 和 {payload.count(chr(10))} 个换行符 (LF)")
        LOGGER.debug(f"   使用 Windows 换行符: \\r\\n (CRLF)")
        
        # 使用 win32clipboard 设置剪贴板（模拟文本编辑器的复制行为）
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            # 设置 CF_UNICODETEXT 格式（Unicode 文本，Windows 推荐格式）
            # 这是文本编辑器通常使用的格式
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, payload)
            # 同时设置 CF_TEXT 格式（ANSI 文本，兼容旧应用）
            # 注意：CF_TEXT 需要 ANSI 编码（通常是系统默认编码）
            try:
                payload_ansi = payload.encode('mbcs')  # Windows 多字节字符集（ANSI）
                win32clipboard.SetClipboardData(win32clipboard.CF_TEXT, payload_ansi)
            except:
                pass  # 如果编码失败，只使用 Unicode
            win32clipboard.CloseClipboard()
            LOGGER.info(f"弹窗 #{popup_index}: LOT 列表已复制到剪贴板 ({len(lots)} 个，使用 win32clipboard)")
        except ImportError:
            # 如果 win32clipboard 不可用，使用 pyperclip（但可能格式不对）
            LOGGER.warning("win32clipboard 未安装，使用 pyperclip（可能格式不正确）")
            import pyperclip
            pyperclip.copy(payload)
            LOGGER.info(f"弹窗 #{popup_index}: LOT 列表已复制到剪贴板 ({len(lots)} 个，使用 pyperclip)")
        except Exception as e:
            LOGGER.warning(f"使用 win32clipboard 失败: {e}，回退到 pyperclip")
            import pyperclip
            pyperclip.copy(payload)
            LOGGER.info(f"弹窗 #{popup_index}: LOT 列表已复制到剪贴板 ({len(lots)} 个，使用 pyperclip)")
        
        # 确保弹窗获得焦点
        try:
            dialog.set_focus()
            time.sleep(0.2)
        except:
            pass
        
        # 【重要】必须点击 Paste 按钮！
        # Paste 按钮会自动将多个 LOT 分行显示
        # 直接 Ctrl+V 会导致所有 LOT 挤在一行，数据抓取失败
        paste_btn = None
        paste_clicked = False
        
        try:
            # 通过 auto_id 查找 Paste 按钮
            paste_btn = dialog.child_window(auto_id="cmdPaste")
            if paste_btn.is_enabled():
                paste_btn.click_input()
                paste_clicked = True
                LOGGER.info(f"✅ 弹窗 #{popup_index}: 已点击 Paste 按钮")
        except Exception as e1:
            try:
                # 备选：通过标题查找
                paste_btn = dialog.child_window(title="Paste")
                if paste_btn.is_enabled():
                    paste_btn.click_input()
                    paste_clicked = True
                    LOGGER.info(f"✅ 弹窗 #{popup_index}: 已点击 Paste 按钮 (通过 title)")
            except Exception as e2:
                LOGGER.error(f"❌ 无法找到 Paste 按钮！")
                LOGGER.error(f"   错误1 (auto_id): {e1}")
                LOGGER.error(f"   错误2 (title): {e2}")
                raise RuntimeError(
                    f"弹窗 #{popup_index}: 无法找到 Paste 按钮。"
                    "多个 LOT 必须通过 Paste 按钮才能正确分行。"
                )
        
        if not paste_clicked:
            raise RuntimeError(f"弹窗 #{popup_index}: Paste 按钮点击失败")
        
        # 等待 Paste 处理完成
        time.sleep(1.0)

        # 点击 OK 按钮
        try:
            ok_btn = dialog.child_window(auto_id="CmdOK")
            ok_btn.click_input()
            LOGGER.info(f"✅ 弹窗 #{popup_index}: 已点击 OK 按钮")
        except Exception as e:
            LOGGER.warning(f"通过 auto_id 点击 OK 失败: {e}")
            try:
                ok_btn = dialog.child_window(title="OK")
                ok_btn.click_input()
                LOGGER.info(f"✅ 弹窗 #{popup_index}: 已点击 OK 按钮 (通过 title)")
            except Exception as e2:
                LOGGER.warning(f"点击 OK 按钮失败: {e2}, 尝试 Enter 键")
                dialog.type_keys("{ENTER}")
                LOGGER.info(f"✅ 弹窗 #{popup_index}: 已按 Enter 键")
        
        # 等待弹窗关闭（缩短等待时间，因为后续会轮询文件）
        time.sleep(0.5)  # 减少到 0.5 秒，让弹窗有时间关闭即可
        LOGGER.info(f"✅ 弹窗 #{popup_index}: LOT 已提交，SQLPathFinder 开始执行查询...")
        return True

    def _check_query_log_window(self) -> bool:
        """检查 Query Log 窗口是否出现（表明查询已开始执行）"""
        try:
            import pygetwindow as gw
            all_windows = gw.getAllWindows()
            for w in all_windows:
                if w.title and w.visible:
                    title_lower = w.title.lower()
                    if "query log" in title_lower:
                        LOGGER.info(f"✅ 检测到 Query Log 窗口: {w.title}")
                        return True
            return False
        except:
            # 如果 pygetwindow 不可用，尝试使用 pywinauto
            try:
                if Application is None:
                    return False
                app = Application(backend="win32").connect(title_re="Query Log", timeout=1, visible_only=True)
                windows = app.windows()
                if windows:
                    LOGGER.info(f"✅ 检测到 Query Log 窗口: {windows[0].window_text()}")
                    return True
                return False
            except:
                return False

    def _enter_lots(self, lots: Sequence[str]) -> None:
        """输入 LOT 列表到所有弹窗（可能有多个参数需要输入）"""
        LOGGER.info("等待 LOT 输入弹窗：Prompt For Values")
        
        # 缩短等待时间，改为主动检测（最多等待2秒）
        LOGGER.info("等待弹窗出现...")
        time.sleep(2)
        
        # 处理第一个弹窗
        LOGGER.info("处理第一个参数弹窗...")
        if not self._process_single_popup(lots, popup_index=1):
            LOGGER.error("❌ 无法找到第一个 LOT 输入弹窗")
            LOGGER.error("")
            LOGGER.error("诊断信息：")
            LOGGER.error("  1. 请手动按 F8，确认弹窗是否出现")
            LOGGER.error("  2. 检查弹窗是否被其他窗口遮挡")
            LOGGER.error("")
            raise TimeoutError("等待第一个 LOT 输入弹窗超时")
        
        # 等待并检查 Query Log 窗口是否出现（表明查询已开始执行）
        LOGGER.info("等待 Query Log 窗口出现（表明查询已开始执行）...")
        deadline = time.time() + 10  # 最多等待 10 秒
        query_log_found = False
        
        while time.time() < deadline:
            if self._check_query_log_window():
                query_log_found = True
                break
            time.sleep(0.5)  # 每 0.5 秒检查一次
        
        if query_log_found:
            LOGGER.info("✅ Query Log 窗口已出现，查询正在执行中")
        else:
            # 如果没有检测到 Query Log 窗口，尝试检查是否有第二个弹窗（兼容旧逻辑）
            LOGGER.info("未检测到 Query Log 窗口，检查是否有第二个参数弹窗...")
            time.sleep(2)
            
            if self._process_single_popup(lots, popup_index=2):
                LOGGER.info("✅ 第二个参数弹窗已处理")
                
                # 检查是否还有第三个弹窗
                LOGGER.info("检查是否有第三个参数弹窗...")
                time.sleep(2)
                
                if self._process_single_popup(lots, popup_index=3):
                    LOGGER.info("✅ 第三个参数弹窗已处理")
                else:
                    LOGGER.info("没有第三个参数弹窗")
            else:
                LOGGER.info("没有第二个参数弹窗，假设查询已开始执行")
        
        LOGGER.info(f"✅ 所有参数弹窗已处理完成，LOT 数：{len(lots)}")

    def execute(self, lots: Sequence[str]) -> None:
        LOGGER.info("启动 SPF 运行，LOT 数：%s", len(lots))
        self._ensure_application()
        self._click_run_button()
        self._enter_lots(lots)
        LOGGER.info("LOT 已下发给 SQLPathFinder，等待 by_lot.csv 更新")

    def wait_for_output(self) -> Path:
        csv_path = self.config.paths.output_csv
        deadline = time.time() + self.config.timeouts.overall_timeout
        last_size = -1
        stable_checks = 0
        check_count = 0
        start_time = time.time()
        
        LOGGER.info("=" * 60)
        LOGGER.info("开始等待 SQLPathFinder 查询完成...")
        LOGGER.info(f"输出文件路径: {csv_path}")
        LOGGER.info(f"超时时间: {self.config.timeouts.overall_timeout} 秒")
        LOGGER.info("=" * 60)
        
        while time.time() < deadline:
            check_count += 1
            elapsed = int(time.time() - start_time)
            
            if csv_path.exists():
                size = csv_path.stat().st_size
                if size == last_size:
                    stable_checks += 1
                    if stable_checks >= self.config.timeouts.file_stabilize_checks:
                        LOGGER.info("=" * 60)
                        LOGGER.info(f"✅ CSV 文件已生成并稳定: {csv_path}")
                        LOGGER.info(f"   文件大小: {size:,} 字节")
                        LOGGER.info(f"   等待时间: {elapsed} 秒")
                        LOGGER.info("=" * 60)
                        return csv_path
                    else:
                        # 文件存在但还在变化，每 2 次检查输出一次进度
                        if check_count % 2 == 0:
                            LOGGER.info(f"⏳ 文件正在更新中... (已等待 {elapsed} 秒，稳定检查: {stable_checks}/{self.config.timeouts.file_stabilize_checks})")
                else:
                    # 文件大小变化了，说明还在写入
                    if last_size == -1:
                        LOGGER.info(f"✅ 检测到 CSV 文件已创建 (大小: {size:,} 字节，已等待 {elapsed} 秒)")
                    else:
                        LOGGER.info(f"📝 文件正在写入中... (大小: {last_size:,} → {size:,} 字节，已等待 {elapsed} 秒)")
                    last_size = size
                    stable_checks = 0
            else:
                # 文件还不存在，每 5 次检查输出一次进度（避免日志过多）
                if check_count % 5 == 0:
                    LOGGER.info(f"⏳ 等待 CSV 文件生成... (已等待 {elapsed} 秒)")
            
            time.sleep(self.config.timeouts.file_stabilize_interval)
        
        # 超时
        elapsed = int(time.time() - start_time)
        LOGGER.error("=" * 60)
        LOGGER.error(f"❌ 等待输出文件超时: {csv_path}")
        LOGGER.error(f"   已等待: {elapsed} 秒")
        LOGGER.error(f"   超时限制: {self.config.timeouts.overall_timeout} 秒")
        if csv_path.exists():
            LOGGER.error(f"   文件存在但未稳定 (最后大小: {last_size:,} 字节)")
        else:
            LOGGER.error(f"   文件不存在")
        LOGGER.error("=" * 60)
        raise TimeoutError(f"等待输出文件超时：{csv_path}")

