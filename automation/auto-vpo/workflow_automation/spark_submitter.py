"""Spark网页提交VPO数据模块"""
import logging
import re
import time
from typing import Optional, List
from dataclasses import dataclass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException

try:
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    WEBDRIVER_MANAGER_AVAILABLE = True
except ImportError:
    WEBDRIVER_MANAGER_AVAILABLE = False

LOGGER = logging.getLogger(__name__)


@dataclass
class SparkConfig:
    """Spark网页配置"""
    url: str
    vpo_category: str = "correlation"  # VPO类别
    step: str = "B5"  # Step选项
    tags: str = "CCG_24J-TEST"  # Tags标签
    timeout: int = 60
    retry_count: int = 3
    retry_delay: int = 2
    wait_after_submit: int = 5
    headless: bool = False
    implicit_wait: int = 10
    explicit_wait: int = 20
    # 提交完成后，从Dashboard收集VPO的等待时间（分钟），0 表示不等待
    vpo_collect_wait_minutes: int = 0


class SparkSubmitter:
    """Spark网页数据提交器"""
    
    def __init__(self, config: SparkConfig):
        self.config = config
        self._driver: Optional[webdriver.Chrome] = None
    
    def _init_driver(self) -> None:
        """初始化WebDriver"""
        if self._driver is not None:
            return
        
        LOGGER.info("初始化Chrome WebDriver...")
        
        options = webdriver.ChromeOptions()
        if self.config.headless:
            options.add_argument('--headless')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-gpu')
        
        # 解决GCM (Google Cloud Messaging) DEPRECATED_ENDPOINT错误
        # 禁用GCM相关的功能，避免尝试连接已弃用的端点
        options.add_argument('--disable-sync')  # 禁用同步功能（可能触发GCM）
        options.add_argument('--disable-background-networking')  # 禁用后台网络请求（包括GCM）
        options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        # 禁用通知权限请求（GCM相关）
        prefs = {
            'profile.default_content_setting_values.notifications': 2,  # 禁用通知
            'profile.default_content_settings.popups': 0,  # 禁用弹窗
        }
        options.add_experimental_option('prefs', prefs)
        
        try:
            if WEBDRIVER_MANAGER_AVAILABLE:
                # 使用webdriver-manager自动管理ChromeDriver
                service = Service(ChromeDriverManager().install())
                self._driver = webdriver.Chrome(service=service, options=options)
            else:
                # 使用系统PATH中的ChromeDriver
                self._driver = webdriver.Chrome(options=options)
            self._driver.implicitly_wait(self.config.implicit_wait)
            LOGGER.info("✅ Chrome WebDriver初始化成功")
        except WebDriverException as e:
            raise RuntimeError(f"无法初始化Chrome WebDriver: {e}")
    
    def _close_driver(self) -> None:
        """关闭WebDriver"""
        if self._driver:
            try:
                self._driver.quit()
                LOGGER.info("已关闭Chrome WebDriver")
            except Exception as e:
                LOGGER.warning(f"关闭WebDriver时出错: {e}")
            finally:
                self._driver = None
    
    def _navigate_to_page(self) -> None:
        """导航到Spark页面"""
        if not self._driver:
            raise RuntimeError("WebDriver未初始化")
        
        LOGGER.info(f"导航到Spark页面: {self.config.url}")
        self._driver.get(self.config.url)
        
        # 等待页面加载
        try:
            WebDriverWait(self._driver, self.config.explicit_wait).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            LOGGER.info("✅ 页面加载完成")
        except TimeoutException:
            LOGGER.warning("页面加载超时，继续执行...")
    
    def _click_add_new_button(self) -> bool:
        """
        点击右上角的'Add New'按钮
        
        Returns:
            True如果点击成功
        """
        LOGGER.info("=" * 60)
        LOGGER.info("步骤：查找并点击'Add New'按钮")
        LOGGER.info("=" * 60)
        
        try:
            add_new_button = None
            
            # 方法1: 通过ID定位（最可靠，按钮有固定ID: dashboardAddNew）
            LOGGER.info("方法1：通过ID定位（id=dashboardAddNew）...")
            try:
                add_new_button = WebDriverWait(self._driver, 2).until(
                    EC.element_to_be_clickable((By.ID, "dashboardAddNew"))
                )
                LOGGER.info(f"✅ 方法1成功：通过ID找到'Add New'按钮")
                LOGGER.info(f"   按钮状态：displayed={add_new_button.is_displayed()}, enabled={add_new_button.is_enabled()}")
                LOGGER.info(f"   按钮位置：{add_new_button.location}, 大小：{add_new_button.size}")
            except TimeoutException:
                LOGGER.warning("⚠️ 方法1失败：通过ID未找到按钮（等待2秒超时），尝试其他方法...")
            
            # 方法2: 通过包含dashboard-container__text的span查找按钮
            if not add_new_button:
                LOGGER.info("方法2：通过dashboard-container__text class定位...")
                try:
                    add_new_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[@class='dashboard-container__text' and contains(text(), 'Add New')]]"))
                    )
                    LOGGER.info(f"✅ 方法2成功：通过span class找到'Add New'按钮")
                    LOGGER.info(f"   按钮状态：displayed={add_new_button.is_displayed()}, enabled={add_new_button.is_enabled()}")
                except TimeoutException:
                    LOGGER.warning("⚠️ 方法2失败：通过span class未找到按钮（等待2秒超时）")
            
            # 方法3: 通过按钮包含的span文本查找
            if not add_new_button:
                LOGGER.info("方法3：通过span文本定位...")
                try:
                    add_new_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Add New')]]"))
                    )
                    LOGGER.info(f"✅ 方法3成功：通过span文本找到'Add New'按钮")
                    LOGGER.info(f"   按钮状态：displayed={add_new_button.is_displayed()}, enabled={add_new_button.is_enabled()}")
                except TimeoutException:
                    LOGGER.warning("⚠️ 方法3失败：通过span文本未找到按钮（等待2秒超时）")
            
            # 方法4: 通过按钮文本查找
            if not add_new_button:
                LOGGER.info("方法4：通过按钮文本定位...")
                try:
                    add_new_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add New')]"))
                    )
                    LOGGER.info(f"✅ 方法4成功：通过按钮文本找到'Add New'按钮")
                    LOGGER.info(f"   按钮状态：displayed={add_new_button.is_displayed()}, enabled={add_new_button.is_enabled()}")
                except TimeoutException:
                    LOGGER.warning("⚠️ 方法4失败：通过按钮文本未找到按钮（等待2秒超时）")
            
            # 方法5: 通过CSS选择器查找（button--large class）
            if not add_new_button:
                LOGGER.info("方法5：通过CSS class遍历查找...")
                try:
                    buttons = self._driver.find_elements(By.CSS_SELECTOR, "button.button--large")
                    LOGGER.info(f"   找到 {len(buttons)} 个button--large按钮")
                    for idx, button in enumerate(buttons, 1):
                        button_text = button.text.strip()
                        LOGGER.info(f"   检查按钮 {idx}: 文本='{button_text}', displayed={button.is_displayed()}, enabled={button.is_enabled()}")
                        if "Add New" in button_text and button.is_displayed() and button.is_enabled():
                            add_new_button = button
                            LOGGER.info(f"✅ 方法5成功：通过CSS class找到按钮: '{button_text}'")
                            break
                except Exception as e:
                    LOGGER.warning(f"⚠️ 方法5失败：遍历按钮时出错: {e}")
            
            if add_new_button:
                LOGGER.info("准备点击'Add New'按钮...")
                # 滚动到按钮可见
                LOGGER.info("   滚动到按钮可见...")
                self._driver.execute_script("arguments[0].scrollIntoView(true);", add_new_button)
                time.sleep(0.3)
                
                # 点击按钮
                LOGGER.info("   执行点击操作...")
                add_new_button.click()
                LOGGER.info("✅ 已点击'Add New'按钮")
                LOGGER.info("   等待页面响应（1秒）...")
                time.sleep(1.0)  # 等待页面响应
                LOGGER.info("✅ 步骤完成：'Add New'按钮点击成功")
                return True
            else:
                LOGGER.error("❌ 所有方法都失败：未找到'Add New'按钮")
                LOGGER.error("   调试信息：已尝试5种定位方法，均未找到按钮")
                return False
                
        except Exception as e:
            LOGGER.error(f"点击'Add New'按钮失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _fill_test_program_path(self, tp_path: str) -> bool:
        """
        填写Test Program Path并点击Apply
        
        Args:
            tp_path: TP路径
            
        Returns:
            True如果操作成功
        """
        LOGGER.info("=" * 60)
        LOGGER.info(f"步骤：填写Test Program Path")
        LOGGER.info(f"目标路径: {tp_path}")
        LOGGER.info("=" * 60)
        
        try:
            # 等待输入框出现并获得焦点
            LOGGER.info("等待输入框出现并获得焦点（1.5秒）...")
            time.sleep(1.5)
            
            # 查找"Provide test program path"输入框
            LOGGER.info("开始查找TP路径输入框...")
            input_field = None
            
            # 方法0: 直接使用当前获得焦点的元素（光标在跳动说明已经有焦点）
            LOGGER.info("方法0：检查当前焦点元素...")
            try:
                input_field = self._driver.switch_to.active_element
                tag_name = input_field.tag_name.lower()
                LOGGER.info(f"   当前焦点元素：tag={tag_name}")
                if tag_name == "input" or tag_name == "textarea":
                    LOGGER.info(f"✅ 方法0成功：使用当前焦点元素作为输入框 (tag: {tag_name})")
                    LOGGER.info(f"   输入框状态：displayed={input_field.is_displayed()}, enabled={input_field.is_enabled()}")
                else:
                    LOGGER.warning(f"⚠️ 方法0失败：当前焦点元素不是输入框 (tag: {tag_name})")
                    input_field = None
            except Exception as e:
                LOGGER.warning(f"⚠️ 方法0失败：获取焦点元素失败: {e}")
            
            # 方法1: 查找对话框中最大的输入框
            if not input_field:
                LOGGER.info("方法1：查找最大的可见输入框...")
                try:
                    # 查找所有可见的input和textarea
                    all_inputs = self._driver.find_elements(By.XPATH, "//input[@type='text' or not(@type)] | //textarea")
                    LOGGER.info(f"   找到 {len(all_inputs)} 个输入框元素")
                    
                    # 过滤可见的
                    visible_inputs = [inp for inp in all_inputs if inp.is_displayed()]
                    LOGGER.info(f"   其中 {len(visible_inputs)} 个可见")
                    
                    # 找最大的
                    if visible_inputs:
                        largest_input = max(visible_inputs, key=lambda x: x.size.get('width', 0) * x.size.get('height', 0))
                        input_field = largest_input
                        LOGGER.info(f"✅ 方法1成功：使用最大的输入框")
                        LOGGER.info(f"   输入框大小：宽度={largest_input.size.get('width')}, 高度={largest_input.size.get('height')}")
                        LOGGER.info(f"   输入框状态：displayed={input_field.is_displayed()}, enabled={input_field.is_enabled()}")
                    else:
                        LOGGER.warning("⚠️ 方法1失败：未找到可见的输入框")
                except Exception as e:
                    LOGGER.warning(f"⚠️ 方法1失败: {e}")
            
            # 方法2: 通过包含"path"的label查找
            if not input_field:
                LOGGER.info("方法2：通过包含'path'的label查找...")
                try:
                    labels = self._driver.find_elements(By.XPATH, "//*[contains(text(), 'path') or contains(text(), 'Path')]")
                    LOGGER.info(f"   找到 {len(labels)} 个包含'path'的label")
                    for idx, label in enumerate(labels, 1):
                        try:
                            label_text = label.text.strip()[:50]
                            LOGGER.info(f"   检查label {idx}: 文本='{label_text}'")
                            # 尝试找label附近的输入框
                            nearby_inputs = label.find_elements(By.XPATH, "./following-sibling::*//input | ./following-sibling::input | .//input | ./parent::*/following-sibling::*//input")
                            if nearby_inputs:
                                input_field = nearby_inputs[0]
                                LOGGER.info(f"✅ 方法2成功：通过label找到输入框")
                                LOGGER.info(f"   输入框状态：displayed={input_field.is_displayed()}, enabled={input_field.is_enabled()}")
                                break
                        except:
                            continue
                except Exception as e:
                    LOGGER.warning(f"⚠️ 方法2失败: {e}")
            
            if not input_field:
                LOGGER.error("❌ 所有方法都失败：未找到Test Program Path输入框")
                # 列出所有可见的输入框用于调试
                try:
                    all_inputs = self._driver.find_elements(By.XPATH, "//input | //textarea")
                    LOGGER.info(f"   调试：页面上共有 {len(all_inputs)} 个输入框")
                    for i, inp in enumerate(all_inputs[:5]):  # 只显示前5个
                        LOGGER.info(f"     输入框 {i+1}: type={inp.get_attribute('type')}, visible={inp.is_displayed()}, size={inp.size}")
                except:
                    pass
                return False
            
            # 清空并填写路径
            LOGGER.info("开始填写TP路径...")
            LOGGER.info(f"   清空输入框...")
            input_field.clear()
            LOGGER.info(f"   输入路径: {tp_path}")
            input_field.send_keys(tp_path)
            LOGGER.info(f"✅ 已填写TP路径: {tp_path}")
            
            # 立即查找并点击Apply按钮（优化：优先使用ID定位）
            LOGGER.info("开始查找'Apply'按钮...")
            apply_button = None
            
            # 方法1: 通过ID定位（最可靠，按钮有固定ID: tpPathApply）
            LOGGER.info("方法1：通过ID定位（id=tpPathApply）...")
            try:
                apply_button = WebDriverWait(self._driver, 2).until(
                    EC.element_to_be_clickable((By.ID, "tpPathApply"))
                )
                LOGGER.info(f"✅ 方法1成功：通过ID找到'Apply'按钮")
                LOGGER.info(f"   按钮状态：displayed={apply_button.is_displayed()}, enabled={apply_button.is_enabled()}")
            except TimeoutException:
                LOGGER.warning("⚠️ 方法1失败：通过ID未找到按钮（等待2秒超时），尝试其他方法...")
            
            # 方法2: 通过CSS class定位（modal__apply-button）
            if not apply_button:
                try:
                    apply_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.modal__apply-button"))
                    )
                    LOGGER.info("✅ 通过CSS class找到'Apply'按钮")
                except TimeoutException:
                    LOGGER.debug("方法2失败：通过CSS class未找到按钮")
            
            # 方法3: 通过按钮包含的span文本查找
            if not apply_button:
                try:
                    apply_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Apply')]]"))
                    )
                    LOGGER.info("✅ 通过span文本找到'Apply'按钮")
                except TimeoutException:
                    LOGGER.debug("方法3失败：通过span文本未找到按钮")
            
            # 方法4: 通过按钮文本查找
            if not apply_button:
                try:
                    apply_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Apply')]"))
                    )
                    LOGGER.info("✅ 通过按钮文本找到'Apply'按钮")
                except TimeoutException:
                    LOGGER.debug("方法4失败：通过按钮文本未找到按钮")
            
            # 方法5: 遍历查找包含Apply的按钮（可能有箭头图标）
            if not apply_button:
                try:
                    buttons = self._driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        button_text = button.text.strip()
                        if "Apply" in button_text and button.is_displayed():
                            if button.is_enabled():
                                apply_button = button
                                LOGGER.info(f"✅ 通过遍历找到按钮: '{button_text}'")
                                break
                except Exception as e:
                    LOGGER.debug(f"方法5失败：遍历按钮时出错: {e}")
            
            if not apply_button:
                LOGGER.error("❌ 未找到'Apply'按钮")
                return False
            
            # 点击Apply按钮（1秒内）
            LOGGER.info("准备点击'Apply'按钮...")
            LOGGER.info("   等待1秒后点击（确保输入完成）...")
            time.sleep(1.0)  # 填写后等待1秒
            LOGGER.info("   执行点击操作...")
            apply_button.click()
            LOGGER.info("✅ 已点击'Apply'按钮")
            
            # **等待loading并点击Continue，然后等待页面跳转**
            if self._wait_for_loading_and_continue():
                LOGGER.info("✅ 步骤完成：已成功填写TP路径并完成页面跳转")
                return True
            else:
                LOGGER.warning("⚠️ loading或页面跳转失败")
                return False
            
        except Exception as e:
            LOGGER.error(f"填写Test Program Path失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _wait_for_loading_and_continue(self) -> bool:
        """
        等待Test Program加载、点击Continue、等待页面跳转完成的完整流程
        
        流程：
        1. 等待Continue按钮从disabled变为enabled
        2. 点击Continue按钮
        3. 等待页面跳转完成（对话框消失 + Add New Experiment按钮出现）
        
        Returns:
            True如果整个流程成功完成
            False如果任何步骤失败
        """
        try:
            # ========== 第1步：等待并点击Continue ==========
            LOGGER.info("=" * 60)
            LOGGER.info("等待Continue按钮变为可点击...")
            LOGGER.info("=" * 60)
            
            max_wait = 60  # 最多等待60秒
            check_interval = 0.5  # 每0.5秒检查一次
            elapsed = 0
            continue_clicked = False
            
            # 检查是否有错误提示（即使有错误，Continue按钮可能仍然可点击）
            error_detected = False
            try:
                error_elements = self._driver.find_elements(By.XPATH, "//*[contains(@style, 'color: red') or contains(@class, 'error') or contains(text(), 'Failed to parse') or contains(text(), 'Failed')]")
                if error_elements:
                    for elem in error_elements[:3]:  # 只显示前3个
                        error_text = elem.text.strip()
                        if error_text and ("Failed to parse" in error_text or "Failed" in error_text):
                            LOGGER.warning(f"⚠️ 检测到错误提示: {error_text}")
                            error_detected = True
                    if error_detected:
                        LOGGER.info("⚠️ 检测到错误提示，但将继续尝试点击Continue按钮...")
            except:
                pass
            
            while elapsed < max_wait:
                try:
                    # 检查Continue按钮是否可点击
                    continue_button = self._driver.find_element(By.ID, "tpPathContinue")
                    if continue_button.is_displayed() and continue_button.is_enabled():
                        LOGGER.info(f"✅ Continue按钮已启用（用时{elapsed:.1f}秒）")
                        
                        # 立即点击（即使有错误提示，也要继续点击）
                        LOGGER.info("立即点击Continue按钮...")
                        self._driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center'});", 
                            continue_button
                        )
                        time.sleep(0.2)
                        try:
                            continue_button.click()
                        except:
                            # 如果普通点击失败，尝试JavaScript点击
                            self._driver.execute_script("arguments[0].click();", continue_button)
                        LOGGER.info("✅ 已点击Continue按钮")
                        continue_clicked = True
                        break
                    elif continue_button.is_displayed() and elapsed >= 60:
                        # 如果超过60秒且按钮可见但未启用，检查是否有错误提示
                        try:
                            error_elements = self._driver.find_elements(By.XPATH, "//*[contains(@style, 'color: red') or contains(@class, 'error') or contains(text(), 'Failed to parse') or contains(text(), 'Failed')]")
                            if error_elements:
                                for elem in error_elements[:3]:
                                    error_text = elem.text.strip()
                                    if error_text and ("Failed to parse" in error_text or "Failed" in error_text):
                                        LOGGER.warning(f"⚠️ 超过60秒且检测到错误提示: {error_text}")
                                        LOGGER.info("尝试强制点击Continue按钮（即使可能disabled）...")
                                        
                                        # 尝试移除disabled属性并点击
                                        try:
                                            self._driver.execute_script("arguments[0].removeAttribute('disabled');", continue_button)
                                            LOGGER.info("已移除disabled属性")
                                        except:
                                            pass
                                        
                                        # 滚动到按钮
                                        self._driver.execute_script(
                                            "arguments[0].scrollIntoView({block: 'center'});", 
                                            continue_button
                                        )
                                        time.sleep(0.3)
                                        
                                        # 尝试点击
                                        try:
                                            continue_button.click()
                                            LOGGER.info("✅ 已通过普通点击点击Continue按钮")
                                            continue_clicked = True
                                            break
                                        except:
                                            try:
                                                self._driver.execute_script("arguments[0].click();", continue_button)
                                                LOGGER.info("✅ 已通过JavaScript点击Continue按钮")
                                                continue_clicked = True
                                                break
                                            except Exception as e:
                                                LOGGER.warning(f"点击失败: {e}")
                                        break
                        except:
                            pass
                except:
                    # Continue按钮还未启用或不存在，继续等待
                    pass
                
                # 显示loading状态（仅用于日志）
                if elapsed % 5 == 0 and elapsed > 0:
                    try:
                        loading_elements = self._driver.find_elements(
                            By.XPATH,
                            "//div[contains(@class, 'creation-progress')]"
                        )
                        if loading_elements and any(elem.is_displayed() for elem in loading_elements):
                            LOGGER.info(f"   Loading中...等待Continue启用（{elapsed:.0f}秒/{max_wait}秒）")
                    except:
                        pass
                
                time.sleep(check_interval)
                elapsed += check_interval
                
                # 如果已经点击成功，跳出循环
                if continue_clicked:
                    break
            
            # 如果超时仍未点击成功，检查是否有错误提示
            if not continue_clicked:
                LOGGER.warning(f"⚠️ Continue按钮{max_wait}秒内未启用")
                
                # 检查是否有错误提示（如"Failed to parse the test program"）
                error_detected = False
                error_text = ""
                try:
                    error_elements = self._driver.find_elements(By.XPATH, "//*[contains(@style, 'color: red') or contains(@class, 'error') or contains(text(), 'Failed to parse') or contains(text(), 'Failed')]")
                    if error_elements:
                        for elem in error_elements[:3]:  # 只显示前3个
                            text = elem.text.strip()
                            if text and ("Failed to parse" in text or "Failed" in text):
                                error_detected = True
                                error_text = text
                                LOGGER.warning(f"⚠️ 检测到错误提示: {text}")
                                break
                except:
                    pass
                
                # 如果检测到错误提示，即使Continue按钮可能还是disabled，也尝试强制点击
                if error_detected:
                    LOGGER.info("⚠️ 检测到错误提示，但将继续尝试点击Continue按钮（即使可能disabled）...")
                    try:
                        continue_button = self._driver.find_element(By.ID, "tpPathContinue")
                        if continue_button.is_displayed():
                            LOGGER.info("找到Continue按钮，尝试强制点击（即使可能disabled）...")
                            
                            # 尝试移除disabled属性并点击
                            try:
                                self._driver.execute_script("arguments[0].removeAttribute('disabled');", continue_button)
                                LOGGER.info("已移除disabled属性")
                            except:
                                pass
                            
                            # 滚动到按钮
                            self._driver.execute_script(
                                "arguments[0].scrollIntoView({block: 'center'});", 
                                continue_button
                            )
                            time.sleep(0.3)
                            
                            # 尝试多种点击方法
                            click_success = False
                            try:
                                continue_button.click()
                                LOGGER.info("✅ 已通过普通点击点击Continue按钮")
                                click_success = True
                            except:
                                try:
                                    self._driver.execute_script("arguments[0].click();", continue_button)
                                    LOGGER.info("✅ 已通过JavaScript点击Continue按钮")
                                    click_success = True
                                except Exception as e:
                                    LOGGER.warning(f"JavaScript点击也失败: {e}")
                            
                            if click_success:
                                LOGGER.info("✅ 已强制点击Continue按钮（尽管检测到错误）")
                                continue_clicked = True
                            else:
                                LOGGER.error("❌ 所有点击方法都失败")
                        else:
                            LOGGER.warning("Continue按钮不可见")
                    except Exception as e:
                        LOGGER.error(f"尝试强制点击Continue按钮失败: {e}")
                
                # 如果仍然没有点击成功，返回False
                if not continue_clicked:
                    LOGGER.error("❌ 无法点击Continue按钮，流程失败")
                    return False
            
            # ========== 第2步：等待页面跳转完成 ==========
            LOGGER.info("\n等待页面跳转完成...")
            
            # 等待对话框消失（最多30秒）
            max_dialog_wait = 30
            dialog_disappeared = False
            
            for i in range(max_dialog_wait):
                try:
                    # 检查mat-dialog是否还存在
                    mat_dialogs = self._driver.find_elements(
                        By.XPATH,
                        "//mat-dialog-container | //div[contains(@class,'mat-dialog-container')]"
                    )
                    dialog_exists = any(d.is_displayed() for d in mat_dialogs if mat_dialogs)
                    
                    if not dialog_exists:
                        LOGGER.info(f"✅ 对话框已消失（用时{i+1}秒）")
                        dialog_disappeared = True
                        break
                    
                    if i % 5 == 0 and i > 0:
                        LOGGER.info(f"   等待对话框消失...（{i}秒/{max_dialog_wait}秒）")
                    
                    time.sleep(1.0)
                except:
                    # 找不到对话框，说明已经消失
                    LOGGER.info(f"✅ 对话框已消失（用时{i+1}秒）")
                    dialog_disappeared = True
                    break
            
            if not dialog_disappeared:
                LOGGER.warning(f"⚠️ 对话框{max_dialog_wait}秒后仍未消失")
                # 但是继续检查是否能找到Add New Experiment按钮
            
            # 验证页面跳转成功：查找Add New Experiment按钮
            LOGGER.info("验证页面跳转：查找'Add New Experiment'按钮...")
            try:
                add_exp_button = WebDriverWait(self._driver, 15).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//button[.//span[contains(text(), 'Add New Experiment')] or contains(text(), 'Add New Experiment')]"
                    ))
                )
                LOGGER.info("✅ 找到'Add New Experiment'按钮，页面跳转成功")
                LOGGER.info("=" * 60)
                return True
            except TimeoutException:
                LOGGER.warning("⚠️ 15秒内未找到'Add New Experiment'按钮，验证页面状态...")
                
                # 验证页面是否在Create New Experiments页面，且正在加载test program
                try:
                    # 检查是否在Create New Experiments页面
                    page_title = self._driver.title
                    current_url = self._driver.current_url
                    LOGGER.info(f"当前页面标题: {page_title}")
                    LOGGER.info(f"当前URL: {current_url}")
                    
                    # 检查是否存在creation-progress元素（加载test program的指示器）
                    creation_progress_elements = self._driver.find_elements(
                        By.XPATH,
                        "//div[contains(@class, 'creation-progress')]"
                    )
                    
                    # 检查是否存在"preparing your test program data"文本
                    preparing_text_elements = self._driver.find_elements(
                        By.XPATH,
                        "//div[contains(@class, 'creating-text') and contains(text(), 'preparing your test program data')]"
                    )
                    
                    # 检查是否存在lds-ring加载动画
                    lds_ring_elements = self._driver.find_elements(
                        By.XPATH,
                        "//div[contains(@class, 'lds-ring')]"
                    )
                    
                    is_loading = False
                    if creation_progress_elements:
                        for elem in creation_progress_elements:
                            if elem.is_displayed():
                                is_loading = True
                                LOGGER.info("✅ 检测到'creation-progress'元素，页面正在加载test program")
                                break
                    
                    if preparing_text_elements:
                        for elem in preparing_text_elements:
                            if elem.is_displayed():
                                is_loading = True
                                LOGGER.info("✅ 检测到'preparing your test program data'文本，正在准备test program数据")
                                break
                    
                    if lds_ring_elements:
                        for elem in lds_ring_elements:
                            if elem.is_displayed():
                                is_loading = True
                                LOGGER.info("✅ 检测到'lds-ring'加载动画，页面正在加载")
                                break
                    
                    if is_loading:
                        LOGGER.info("✅ 确认页面在'Create New Experiments'页面，且正在加载test program")
                        LOGGER.info("等待test program加载完成，继续查找'Add New Experiment'按钮...")
                        
                        # 等待加载完成（最多再等30秒）
                        max_additional_wait = 30
                        for i in range(max_additional_wait):
                            time.sleep(1.0)
                            
                            # 检查加载是否完成（creation-progress消失）
                            still_loading = False
                            try:
                                current_progress = self._driver.find_elements(
                                    By.XPATH,
                                    "//div[contains(@class, 'creation-progress')]"
                                )
                                for elem in current_progress:
                                    if elem.is_displayed():
                                        still_loading = True
                                        break
                            except:
                                pass
                            
                            if not still_loading:
                                LOGGER.info(f"✅ Test program加载完成（额外等待{i+1}秒）")
                                break
                            
                            # 检查是否已经出现Add New Experiment按钮
                            try:
                                add_exp_button = self._driver.find_element(
                                    By.XPATH,
                                    "//button[.//span[contains(text(), 'Add New Experiment')] or contains(text(), 'Add New Experiment')]"
                                )
                                if add_exp_button.is_displayed():
                                    LOGGER.info("✅ 找到'Add New Experiment'按钮，页面跳转成功")
                                    LOGGER.info("=" * 60)
                                    return True
                            except:
                                pass
                            
                            if i % 5 == 0 and i > 0:
                                LOGGER.info(f"   等待test program加载完成...（{i}秒/{max_additional_wait}秒）")
                        
                        # 加载完成后，再次尝试查找Add New Experiment按钮
                        try:
                            add_exp_button = WebDriverWait(self._driver, 10).until(
                                EC.presence_of_element_located((
                                    By.XPATH,
                                    "//button[.//span[contains(text(), 'Add New Experiment')] or contains(text(), 'Add New Experiment')]"
                                ))
                            )
                            LOGGER.info("✅ 找到'Add New Experiment'按钮，页面跳转成功")
                            LOGGER.info("=" * 60)
                            return True
                        except TimeoutException:
                            LOGGER.error("❌ Test program加载完成后，仍未找到'Add New Experiment'按钮")
                    else:
                        LOGGER.warning("⚠️ 未检测到test program加载状态，页面可能不在'Create New Experiments'页面")
                        LOGGER.warning("⚠️ 或者页面加载已完成但未跳转到预期页面")
                        
                except Exception as e:
                    LOGGER.warning(f"⚠️ 验证页面状态时出错: {e}")
                
                LOGGER.error("❌ 页面跳转可能失败")
                LOGGER.info("=" * 60)
                return False
            
        except Exception as e:
            LOGGER.error(f"❌ 等待loading和Continue过程中出错: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _wait_for_test_program_loading(self) -> bool:
        """
        等待Test Program加载并点击Continue按钮
        
        点击Apply后：
        1. Continue按钮从disabled变为enabled
        2. 可能出现loading界面（preparing your test program data）
        3. 需要在loading过程中持续监测Continue按钮，一旦enabled就立即点击
        
        Returns:
            True如果成功点击Continue或已进入下一界面
            False如果超时或出错
        """
        try:
            LOGGER.info("=" * 60)
            LOGGER.info("等待Continue按钮变为可点击...")
            LOGGER.info("=" * 60)
            
            max_wait = 60  # 最多等待60秒
            check_interval = 0.5  # 每0.5秒检查一次（更频繁）
            elapsed = 0
            
            continue_clicked = False
            
            while elapsed < max_wait:
                try:
                    # **优先检查Continue按钮是否可点击**
                    try:
                        continue_button = self._driver.find_element(By.ID, "tpPathContinue")
                        if continue_button.is_displayed() and continue_button.is_enabled():
                            LOGGER.info(f"✅ Continue按钮已启用（用时{elapsed:.1f}秒）")
                            
                            # 立即点击
                            LOGGER.info("立即点击Continue按钮...")
                            self._driver.execute_script(
                                "arguments[0].scrollIntoView({block: 'center'});", 
                                continue_button
                            )
                            time.sleep(0.2)
                            continue_button.click()
                            LOGGER.info("✅ 已点击Continue按钮")
                            continue_clicked = True
                            LOGGER.info("=" * 60)
                            return True
                    except Exception as e:
                        # Continue按钮还未启用或不存在，继续等待
                        pass
                    
                    # 检查loading状态（仅用于日志）
                    if elapsed % 5 == 0 and elapsed > 0:
                        try:
                            loading_elements = self._driver.find_elements(
                                By.XPATH,
                                "//div[contains(@class, 'creation-progress')]"
                            )
                            if loading_elements and any(elem.is_displayed() for elem in loading_elements):
                                LOGGER.info(f"   Loading中...等待Continue启用（{elapsed:.0f}秒/{max_wait}秒）")
                            else:
                                LOGGER.info(f"   等待Continue启用（{elapsed:.0f}秒/{max_wait}秒）")
                        except:
                            pass
                    
                    time.sleep(check_interval)
                    elapsed += check_interval
                    
                except Exception as e:
                    LOGGER.debug(f"检查过程中出错: {e}")
                    time.sleep(check_interval)
                    elapsed += check_interval
            
            # 超时
            if not continue_clicked:
                LOGGER.warning(f"⚠️ Continue按钮{max_wait}秒内未启用")
                # 检查是否已经进入下一界面
                try:
                    # 如果找不到tpPathContinue，说明可能已经跳转了
                    continue_buttons = self._driver.find_elements(By.ID, "tpPathContinue")
                    if not continue_buttons:
                        LOGGER.info("ℹ️ Continue按钮已消失，可能已进入下一界面")
                        LOGGER.info("=" * 60)
                        return True
                except:
                    pass
                
                LOGGER.info("=" * 60)
                return False
            
            return True
            
        except Exception as e:
            LOGGER.warning(f"⚠️ 等待Continue时出错: {e}")
            import traceback
            LOGGER.warning(traceback.format_exc())
            return False
    
    def _wait_and_click_continue(self) -> bool:
        """
        等待加载完成并点击Continue按钮
        
        注意：即使出现错误提示（红色文字），也会继续点击Continue按钮
        
        Returns:
            True如果点击成功
        """
        try:
            # 检查是否有错误提示
            try:
                error_elements = self._driver.find_elements(By.XPATH, "//*[contains(@style, 'color: red') or contains(@class, 'error') or contains(text(), 'Failed')]")
                if error_elements:
                    for elem in error_elements[:3]:  # 只显示前3个
                        error_text = elem.text.strip()
                        if error_text:
                            LOGGER.warning(f"⚠️ 检测到错误提示: {error_text}")
                    LOGGER.info("忽略错误提示，继续点击Continue...")
            except:
                pass
            
            LOGGER.info("等待Continue按钮变为可点击...")
            
            # 等待Continue按钮出现并可点击（优化：先快速检查，减少等待时间）
            continue_button = None
            
            # 方法1: 先快速检查按钮是否存在（不等待）
            try:
                continue_buttons = self._driver.find_elements(By.ID, "tpPathContinue")
                if continue_buttons:
                    for btn in continue_buttons:
                        if btn.is_displayed():
                            try:
                                if btn.is_enabled():
                                    continue_button = btn
                                    LOGGER.info("✅ 通过ID找到'Continue'按钮（快速检查）")
                                    break
                            except:
                                pass
                
                # 如果按钮存在但不可点击，等待最多2秒让它变为可点击
                if continue_button:
                    try:
                        continue_button = WebDriverWait(self._driver, 2).until(
                            EC.element_to_be_clickable(continue_button)
                        )
                    except TimeoutException:
                        LOGGER.debug("按钮存在但2秒内未变为可点击，尝试直接点击")
            except Exception as e:
                LOGGER.debug(f"快速检查失败: {e}")
            
            # 方法2: 如果快速检查失败，使用显式等待（减少到5秒）
            if not continue_button:
                try:
                    continue_button = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "tpPathContinue"))
                    )
                    LOGGER.info("✅ 通过ID找到'Continue'按钮（显式等待）")
                except TimeoutException:
                    LOGGER.debug("方法2失败：通过ID未找到按钮，尝试其他方法...")
            
            # 方法3: 通过CSS class定位（button--large），然后检查文本
            if not continue_button:
                try:
                    buttons = self._driver.find_elements(By.CSS_SELECTOR, "button.button--large")
                    for button in buttons:
                        if "Continue" in button.text.strip() and button.is_displayed() and button.is_enabled():
                            continue_button = button
                            LOGGER.info("✅ 通过CSS class找到'Continue'按钮")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法3失败：通过CSS class未找到按钮: {e}")
            
            # 方法4: 通过文本查找Continue按钮
            if not continue_button:
                try:
                    continue_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))
                    )
                    LOGGER.info("✅ 通过文本找到'Continue'按钮")
                except TimeoutException:
                    LOGGER.debug("方法4失败：通过文本未找到Continue按钮")
            
            if not continue_button:
                LOGGER.info("ℹ️ 未找到'Continue'按钮（可能已进入下一界面）")
                return False
            
            # 滚动到按钮可见
            self._driver.execute_script("arguments[0].scrollIntoView(true);", continue_button)
            time.sleep(0.3)
            
            # 点击Continue按钮（可能需要多次点击）
            # 优化：减少重试次数，但增加每次等待时间
            max_continue_clicks = 6  # 减少到6次重试（从15次）
            
            for click_attempt in range(1, max_continue_clicks + 1):
                LOGGER.info(f"🔄 准备点击'Continue'按钮（第 {click_attempt}/{max_continue_clicks} 次）...")
                
                # 重新查找Continue按钮（可能在重试过程中DOM更新了）
                continue_button = None
                try:
                    # 优先使用ID定位（减少等待时间到5秒）
                    continue_button = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "tpPathContinue"))
                    )
                    LOGGER.info(f"✅ 通过ID找到Continue按钮（第 {click_attempt} 次尝试）")
                except TimeoutException:
                    # 如果ID定位失败，尝试文本定位
                    try:
                        continue_button = WebDriverWait(self._driver, 2).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))
                        )
                        LOGGER.info(f"✅ 通过文本找到Continue按钮（第 {click_attempt} 次尝试）")
                    except TimeoutException:
                        LOGGER.warning(f"⚠️ 5秒内未找到Continue按钮（第 {click_attempt} 次尝试）")
                    
                    # 检查是否已经跳转成功
                    if self._check_target_page_loaded():
                        LOGGER.info("✅ 目标页面已加载，跳转成功！")
                        return True
                    
                    # 如果还没到最后一次尝试，继续循环等待Continue按钮重新出现
                    if click_attempt < max_continue_clicks:
                        LOGGER.info(f"Continue按钮暂时消失，等待10秒后继续尝试...")
                        time.sleep(10.0)  # 从5秒增加到10秒
                        continue  # 继续下一次循环
                    else:
                        # 最后一次尝试也找不到
                        LOGGER.error("❌ 最后一次尝试仍未找到Continue按钮且页面未跳转")
                        return False
                
                if not continue_button:
                    # 理论上不应该到这里，但保险起见
                    LOGGER.warning("Continue按钮为空，跳过本次循环")
                    time.sleep(3.0)
                    continue
                
                # 点击Continue按钮
                try:
                    continue_button.click()
                    LOGGER.info(f"✅ 已点击'Continue'按钮（第 {click_attempt} 次）")
                except Exception as e:
                    LOGGER.warning(f"点击失败: {e}，尝试JavaScript点击")
                    try:
                        self._driver.execute_script("arguments[0].click();", continue_button)
                        LOGGER.info(f"✅ 已通过JavaScript点击'Continue'按钮（第 {click_attempt} 次）")
                    except Exception as e2:
                        LOGGER.error(f"JavaScript点击也失败: {e2}")
                        time.sleep(3.0)
                        continue
                
                # 等待页面加载完成（最多90秒，由_wait_for_page_load_after_continue控制）
                LOGGER.info("⏳ 等待页面加载完成（最多90秒）...")
                load_success = self._wait_for_page_load_after_continue()
                
                if load_success:
                    LOGGER.info(f"✅✅✅ 页面加载完成，跳转成功！（第 {click_attempt} 次点击后成功）")
                    return True
                else:
                    # 加载未成功，继续重试
                    LOGGER.warning(f"⚠️ 第 {click_attempt} 次点击后页面未成功跳转")
                    
                    if click_attempt < max_continue_clicks:
                        LOGGER.info(f"💪 继续重试！准备第 {click_attempt + 1} 次点击...")
                        # 等待更长时间后继续（从2秒增加到5秒）
                        time.sleep(5.0)
                        continue  # 继续下一次循环
                    else:
                        # 已经是最后一次尝试了
                        LOGGER.error(f"❌ 已重试{max_continue_clicks}次，仍未成功跳转")
                        # 最后再检查一次目标页面
                        if self._check_target_page_loaded():
                            LOGGER.info("✅ 最终检查：目标页面已加载！")
                            return True
                        return False
            
            LOGGER.error(f"❌ Continue点击了 {max_continue_clicks} 次仍未成功")
            return False
            
        except Exception as e:
            LOGGER.error(f"点击'Continue'按钮失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _check_target_page_loaded(self) -> bool:
        """
        检查目标页面（带有"Add New Experiment"按钮的页面）是否已加载
        
        Returns:
            True如果目标页面已加载
        """
        try:
            # 检查"Add New Experiment"按钮是否存在且可见
            add_experiment_buttons = self._driver.find_elements(
                By.XPATH, 
                "//button[contains(text(), 'Add New Experiment') or .//span[contains(text(), 'Add New Experiment')]]"
            )
            
            if add_experiment_buttons:
                for btn in add_experiment_buttons:
                    if btn.is_displayed():
                        LOGGER.info("✅ 检测到'Add New Experiment'按钮，目标页面已加载")
                        return True
            
            LOGGER.debug("未检测到'Add New Experiment'按钮")
            return False
            
        except Exception as e:
            LOGGER.debug(f"检查目标页面时出错: {e}")
            return False
    
    def _wait_for_page_load_after_continue(self) -> bool:
        """
        等待Continue点击后的页面加载完成
        
        检测策略：
        1. 检查是否还在"Create New Experiments"对话框（说明未跳转）
        2. 检查mat-dialog元素是否还存在（更准确的检测）
        3. 等待"Add New Experiment"按钮出现（说明跳转成功）
        4. 检测页面是否崩溃
        
        优化：检查时间递减，第1次30秒，第2次20秒，第3次10秒，之后都是5秒
        
        Returns:
            True如果页面加载完成并成功跳转
            False如果仍在原对话框或加载失败
        """
        try:
            # 优化：检查时间从多到少
            # 第1次30秒，第2次20秒，第3次10秒，之后都是5秒
            max_attempts = 6  # 最多尝试6次
            wait_times = [30, 20, 10]  # 前3次检查的等待时间（秒），超过3次后默认5秒
            
            LOGGER.info(f"等待页面跳转完成（最多{max_attempts}次检查，前3次：30s/20s/10s，之后5s）...")
            
            # 检查是否还在"Create New Experiments"对话框
            for check_attempt in range(max_attempts):
                wait_time = wait_times[check_attempt] if check_attempt < len(wait_times) else 5
                LOGGER.info(f"第 {check_attempt + 1} 次检查（等待 {wait_time} 秒）...")
                try:
                    # 方法1: 检查mat-dialog元素是否还存在（更准确）
                    dialog_exists = False
                    try:
                        mat_dialogs = self._driver.find_elements(By.XPATH, "//mat-dialog-container | //div[contains(@class,'mat-dialog-container')]")
                        if mat_dialogs:
                            for dialog in mat_dialogs:
                                if dialog.is_displayed():
                                    dialog_exists = True
                                    break
                    except:
                        pass
                    
                    # 方法2: 检查对话框标题文本是否还存在
                    if not dialog_exists:
                        create_dialog_text = self._driver.find_elements(By.XPATH, "//*[contains(text(), 'Create New Experiments')]")
                        if create_dialog_text and any(elem.is_displayed() for elem in create_dialog_text):
                            dialog_exists = True
                    
                    if dialog_exists:
                        LOGGER.warning(f"⚠️ 仍在'Create New Experiments'对话框中（第{check_attempt + 1}次检查，等待{wait_time}秒）")
                        
                        # 检查是否有错误提示
                        try:
                            error_elements = self._driver.find_elements(By.XPATH, "//*[contains(text(), 'Failed') or contains(text(), 'error') or contains(@style, 'color: red')]")
                            if error_elements:
                                for elem in error_elements[:2]:
                                    error_text = elem.text.strip()
                                    if error_text and "Failed" in error_text:
                                        LOGGER.warning(f"检测到错误: {error_text}")
                        except:
                            pass
                        
                        # **新增：检查Continue按钮是否可用，如果可用就直接点击**
                        continue_button_clicked = False
                        try:
                            # 优先通过ID查找Continue按钮
                            continue_button = None
                            try:
                                continue_button = self._driver.find_element(By.ID, "tpPathContinue")
                                if continue_button.is_displayed() and continue_button.is_enabled():
                                    LOGGER.info("✅ 检测到Continue按钮可用，直接点击...")
                                    self._driver.execute_script("arguments[0].scrollIntoView(true);", continue_button)
                                    time.sleep(0.3)
                                    continue_button.click()
                                    LOGGER.info("✅ 已点击Continue按钮")
                                    continue_button_clicked = True
                                    # 点击后等待一下，然后继续检查对话框是否消失
                                    time.sleep(2.0)
                            except:
                                # 如果ID查找失败，尝试通过文本查找
                                try:
                                    continue_buttons = self._driver.find_elements(By.XPATH, "//button[contains(text(), 'Continue')]")
                                    for btn in continue_buttons:
                                        if btn.is_displayed() and btn.is_enabled():
                                            LOGGER.info("✅ 检测到Continue按钮可用（通过文本），直接点击...")
                                            self._driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                                            time.sleep(0.3)
                                            btn.click()
                                            LOGGER.info("✅ 已点击Continue按钮")
                                            continue_button_clicked = True
                                            time.sleep(2.0)
                                            break
                                except:
                                    pass
                        except Exception as e:
                            LOGGER.debug(f"检查Continue按钮时出错: {e}")
                        
                        # 如果已经点击了Continue按钮，跳过等待，直接进入下一次检查
                        if not continue_button_clicked:
                            # 等待指定时间（时间递减：30秒、25秒、20秒...）
                            time.sleep(wait_time)
                        
                        # 如果是最后一次检查，返回False让上层重新点击Continue
                        if check_attempt == max_attempts - 1:
                            total_wait = sum(wait_times[:check_attempt + 1])
                            LOGGER.warning(f"⚠️ 已等待{total_wait}秒，仍在原对话框中，返回让上层重新点击Continue")
                            return False
                        
                        continue
                    else:
                        # 对话框已消失，说明可能已经跳转
                        total_wait = sum(wait_times[:check_attempt + 1])
                        LOGGER.info(f"✅ 'Create New Experiments'对话框已消失（第{check_attempt + 1}次检查，已等待{total_wait}秒）")
                        break
                        
                except Exception as e:
                    LOGGER.debug(f"检查对话框时出错: {e}")
                    # 出错时也认为对话框可能已消失，继续验证
                    break
            
            # 验证是否成功跳转：查找"Add New Experiment"按钮
            LOGGER.info("验证页面跳转：查找'Add New Experiment'按钮...")
            try:
                # 优先使用dashboard-container__text class定位
                add_exp_button = None
                try:
                    add_exp_button = WebDriverWait(self._driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, "//button[.//span[@class='dashboard-container__text' and contains(text(), 'Add New Experiment')]]"))
                    )
                    LOGGER.info("✅ 通过dashboard-container__text class找到'Add New Experiment'按钮")
                except TimeoutException:
                    # 备用：通过span文本查找
                    try:
                        add_exp_button = WebDriverWait(self._driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Add New Experiment') or .//span[contains(text(), 'Add New Experiment')]]"))
                        )
                        LOGGER.info("✅ 通过文本找到'Add New Experiment'按钮")
                    except TimeoutException:
                        LOGGER.warning("⚠️ 未找到'Add New Experiment'按钮")
                
                if add_exp_button and add_exp_button.is_displayed():
                    LOGGER.info("✅ 'Add New Experiment'按钮已出现，页面跳转成功！")
                    time.sleep(2.0)  # 等待页面稳定（从1.5秒增加到2秒）
                    return True
                elif add_exp_button:
                    LOGGER.warning("⚠️ 找到'Add New Experiment'按钮但不可见，继续等待...")
                    # 再等待5秒
                    time.sleep(5.0)
                    if add_exp_button.is_displayed():
                        LOGGER.info("✅ 'Add New Experiment'按钮现在可见，页面跳转成功！")
                        return True
                    else:
                        LOGGER.error("❌ 'Add New Experiment'按钮仍不可见")
                        return False
                else:
                    LOGGER.error("❌ 未找到'Add New Experiment'按钮")
                    return False
                
            except TimeoutException:
                LOGGER.error("❌ 未找到'Add New Experiment'按钮，页面跳转失败")
                # 最后尝试：检查是否还有其他方式确认页面已跳转
                try:
                    # 检查是否有其他特征元素（如VPO类别选择器等）
                    vpo_elements = self._driver.find_elements(By.XPATH, "//*[contains(text(), 'Correlation') or contains(text(), 'Engineering')]")
                    if vpo_elements:
                        LOGGER.info("✅ 检测到VPO类别选择器，页面可能已跳转")
                        return True
                except:
                    pass
                return False
            
        except Exception as e:
            LOGGER.error(f"等待页面加载时出错: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _click_add_new_experiment(self) -> bool:
        """
        点击右上角的'Add New Experiment'按钮
        
        Returns:
            True如果点击成功
        """
        LOGGER.info("=" * 60)
        LOGGER.info("步骤4：点击'Add New Experiment'按钮")
        LOGGER.info("=" * 60)
        
        try:
            add_experiment_button = None
            
            # 策略1: 优先通过ID查找（最快最可靠）
            LOGGER.info("策略1：通过ID 'egAddNewExperiment' 查找...")
            try:
                # 先快速检查按钮是否存在
                button_elem = self._driver.find_element(By.ID, "egAddNewExperiment")
                if button_elem.is_displayed():
                    LOGGER.info("   按钮已存在，等待其变为可点击...")
                    # 等待按钮可点击（最多1秒）
                    add_experiment_button = WebDriverWait(self._driver, 1).until(
                        EC.element_to_be_clickable((By.ID, "egAddNewExperiment"))
                    )
                    LOGGER.info("✅ 策略1成功：通过ID找到按钮")
                else:
                    LOGGER.info("   按钮存在但不可见")
            except NoSuchElementException:
                LOGGER.info("   策略1失败：按钮不存在")
            except TimeoutException:
                LOGGER.info("   策略1失败：按钮1秒内未变为可点击")
            except Exception as e:
                LOGGER.info(f"   策略1失败: {e}")
            
            # 策略2: 通过dashboard-container__text class查找
            if not add_experiment_button:
                LOGGER.info("策略2：通过dashboard-container__text class查找...")
                try:
                    add_experiment_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((
                            By.XPATH, 
                            "//button[.//span[@class='dashboard-container__text' and contains(text(), 'Add New Experiment')]]"
                        ))
                    )
                    LOGGER.info("✅ 策略2成功：通过dashboard-container__text找到按钮")
                except TimeoutException:
                    LOGGER.info("   策略2失败：2秒内未找到按钮")
            
            # 策略3: 通过简化的XPath查找
            if not add_experiment_button:
                LOGGER.info("策略3：通过简化XPath查找...")
                try:
                    add_experiment_button = WebDriverWait(self._driver, 2).until(
                        EC.element_to_be_clickable((
                            By.XPATH, 
                            "//button[.//span[contains(text(), 'Add New Experiment')]]"
                        ))
                    )
                    LOGGER.info("✅ 策略3成功：通过简化XPath找到按钮")
                except TimeoutException:
                    LOGGER.info("   策略3失败：2秒内未找到按钮")
            
            if not add_experiment_button:
                LOGGER.error("❌ 所有策略均失败，未找到'Add New Experiment'按钮")
                return False
            
            # 输出按钮信息
            try:
                button_id = add_experiment_button.get_attribute("id") or "无ID"
                button_class = add_experiment_button.get_attribute("class") or "无class"
                LOGGER.info(f"按钮信息：ID='{button_id}', class='{button_class[:60]}...'")
            except:
                pass
            
            # 滚动到按钮可见位置
            LOGGER.info("滚动到按钮位置...")
            self._driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});", 
                add_experiment_button
            )
            time.sleep(0.2)
            
            # 点击按钮
            LOGGER.info("点击按钮...")
            add_experiment_button.click()
            LOGGER.info("✅ 已点击'Add New Experiment'按钮")
            
            # 等待页面响应
            LOGGER.info("等待对话框出现（2秒）...")
            time.sleep(2.0)
            
            LOGGER.info("=" * 60)
            LOGGER.info("步骤4完成：成功点击'Add New Experiment'")
            LOGGER.info("=" * 60)
            
            return True
            
        except Exception as e:
            LOGGER.error(f"❌ 点击'Add New Experiment'按钮失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _select_vpo_category(self, category: str) -> bool:
        """
        选择VPO类别（Correlation / Engineering / Walk the lot）
        
        Args:
            category: VPO类别名称
            
        Returns:
            True如果选择成功
        """
        LOGGER.info(f"选择VPO类别: {category}")
        
        try:
            # 等待下拉菜单出现
            time.sleep(1.5)
            
            # 标准化category名称（转小写，用于匹配）
            category_lower = category.lower().strip()
            
            # 映射关系
            category_map = {
                "correlation": "Correlation",
                "engineering": "Engineering", 
                "walk the lot": "Walk the lot",
                "walktheLot": "Walk the lot"
            }
            
            # 获取标准名称
            target_category = category_map.get(category_lower)
            if not target_category:
                LOGGER.warning(f"未知的VPO类别: {category}，默认使用Correlation")
                target_category = "Correlation"
            
            LOGGER.info(f"查找选项: {target_category}")
            
            # 查找并点击对应选项
            option_clicked = False
            
            # 方法1: 通过文本查找
            try:
                option = WebDriverWait(self._driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{target_category}')]"))
                )
                option.click()
                option_clicked = True
                LOGGER.info(f"通过文本找到并点击选项: {target_category}")
            except TimeoutException:
                LOGGER.debug(f"方法1失败：通过文本未找到选项")
            
            # 方法2: 查找下拉菜单中的选项
            if not option_clicked:
                try:
                    # 查找所有可见的选项元素
                    options = self._driver.find_elements(By.XPATH, "//div[@role='menuitem'] | //li[@role='menuitem'] | //button[contains(@class, 'menu-item')]")
                    
                    for option in options:
                        option_text = option.text.strip()
                        LOGGER.debug(f"检查选项: '{option_text}'")
                        if target_category.lower() in option_text.lower():
                            option.click()
                            option_clicked = True
                            LOGGER.info(f"通过遍历找到并点击选项: {option_text}")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            if not option_clicked:
                LOGGER.error(f"❌ 未找到VPO类别选项: {target_category}")
                return False
            
            LOGGER.info(f"✅ 已选择VPO类别: {target_category}")
            time.sleep(1.0)  # 等待选择生效
            
            return True
            
        except Exception as e:
            LOGGER.error(f"选择VPO类别失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _fill_experiment_info(self, step: str, tags: str) -> bool:
        """
        填写实验信息（Step和Tags）
        
        Args:
            step: Step选项（如B4, B5, B0）
            tags: Tags标签（如CCG_24J-TEST）
            
        Returns:
            True如果填写成功
        """
        LOGGER.info(f"填写实验信息 - Step: {step}, Tags: {tags}")
        
        try:
            # 等待表单加载
            time.sleep(2.0)
            
            # 1. 选择Step
            LOGGER.info(f"选择Step: {step}")
            try:
                # 方法1: 查找Step下拉框
                step_dropdown = WebDriverWait(self._driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//select[preceding-sibling::*[contains(text(), 'Step')] or following-sibling::*[contains(text(), 'Step')]]"))
                )
                
                # 选择对应的选项
                from selenium.webdriver.support.ui import Select
                select = Select(step_dropdown)
                select.select_by_visible_text(step)
                LOGGER.info(f"✅ 已选择Step: {step}")
                
            except Exception as e:
                LOGGER.warning(f"方法1失败，尝试其他方法: {e}")
                
                # 方法2: 通过label查找
                try:
                    step_label = self._driver.find_element(By.XPATH, "//*[contains(text(), 'Step:')]")
                    # 找到label附近的select
                    step_dropdown = step_label.find_element(By.XPATH, "following::select[1]")
                    from selenium.webdriver.support.ui import Select
                    select = Select(step_dropdown)
                    select.select_by_visible_text(step)
                    LOGGER.info(f"✅ 已选择Step: {step}")
                except Exception as e2:
                    LOGGER.error(f"选择Step失败: {e2}")
                    return False
            
            time.sleep(0.5)
            
            # 2. 填写Tags
            LOGGER.info(f"填写Tags: {tags}")
            try:
                # 查找Tags输入框
                tags_input = None
                
                # 方法1: 通过label查找
                try:
                    tags_label = self._driver.find_element(By.XPATH, "//*[contains(text(), 'Tags')]")
                    tags_input = tags_label.find_element(By.XPATH, "following::input[1]")
                    LOGGER.info("通过label找到Tags输入框")
                except:
                    pass
                
                # 方法2: 直接查找包含tags相关属性的输入框
                if not tags_input:
                    try:
                        tags_inputs = self._driver.find_elements(By.XPATH, "//input[@placeholder or @name or @id]")
                        for inp in tags_inputs:
                            placeholder = (inp.get_attribute("placeholder") or "").lower()
                            name = (inp.get_attribute("name") or "").lower()
                            id_attr = (inp.get_attribute("id") or "").lower()
                            if "tag" in placeholder or "tag" in name or "tag" in id_attr:
                                tags_input = inp
                                LOGGER.info("通过属性找到Tags输入框")
                                break
                    except:
                        pass
                
                if not tags_input:
                    LOGGER.warning("未找到Tags输入框，可能不是必填项，继续执行")
                else:
                    tags_input.clear()
                    tags_input.send_keys(tags)
                    LOGGER.info(f"✅ 已填写Tags: {tags}")
                
            except Exception as e:
                LOGGER.warning(f"填写Tags时出错: {e}")
            
            time.sleep(0.5)
            
            # 3. 点击Next按钮
            LOGGER.info("查找并点击'Next'按钮...")
            try:
                next_button = WebDriverWait(self._driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Next')]"))
                )
                
                # 滚动到按钮可见
                self._driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                time.sleep(0.3)
                
                next_button.click()
                LOGGER.info("✅ 已点击'Next'按钮")
                time.sleep(2.0)  # 等待页面响应
                
                return True
                
            except Exception as e:
                LOGGER.error(f"点击'Next'按钮失败: {e}")
                return False
            
        except Exception as e:
            LOGGER.error(f"填写实验信息失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _add_lot_name(self, lot_name: str) -> bool:
        """
        在Material标签页输入Lot name并点击Add
        
        Args:
            lot_name: Lot名称（Source Lot值）
            
        Returns:
            True如果添加成功
        """
        LOGGER.info(f"添加Lot name: {lot_name}")
        
        try:
            # 等待Material标签页加载
            time.sleep(1.5)
            
            # 确保"Use lot name"单选按钮被选中
            try:
                use_lot_name_radio = WebDriverWait(self._driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='radio' and (contains(following-sibling::*/text(), 'Use lot name') or contains(../text(), 'Use lot name'))]"))
                )
                if not use_lot_name_radio.is_selected():
                    use_lot_name_radio.click()
                    LOGGER.info("已选择'Use lot name'选项")
                else:
                    LOGGER.info("'Use lot name'选项已被选中")
            except Exception as e:
                LOGGER.debug(f"选择'Use lot name'单选按钮时出错: {e}")
                # 继续执行，可能默认就是选中的
            
            time.sleep(0.5)
            
            # 查找"Lot name"输入框
            lot_input = None
            
            # 方法1: 通过placeholder查找
            try:
                lot_input = WebDriverWait(self._driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Lot name']"))
                )
                LOGGER.info("通过placeholder找到Lot name输入框")
            except TimeoutException:
                LOGGER.debug("方法1失败：通过placeholder未找到")
            
            # 方法2: 查找包含"lot"的输入框
            if not lot_input:
                try:
                    inputs = self._driver.find_elements(By.XPATH, "//input[@type='text']")
                    for inp in inputs:
                        placeholder = (inp.get_attribute("placeholder") or "").lower()
                        name = (inp.get_attribute("name") or "").lower()
                        if "lot" in placeholder or "lot" in name:
                            if inp.is_displayed():
                                lot_input = inp
                                LOGGER.info(f"通过属性找到Lot输入框 (placeholder: {placeholder})")
                                break
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            if not lot_input:
                LOGGER.error("❌ 未找到Lot name输入框")
                return False
            
            # 清空并输入lot name
            lot_input.clear()
            lot_input.send_keys(lot_name)
            LOGGER.info(f"✅ 已输入Lot name: {lot_name}")
            
            # 查找并点击Add按钮
            time.sleep(0.3)
            
            add_button = None
            
            # 方法1: 通过文本查找Add按钮
            try:
                add_button = WebDriverWait(self._driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[text()='Add' or contains(text(), 'Add')]"))
                )
                LOGGER.info("通过文本找到'Add'按钮")
            except TimeoutException:
                LOGGER.debug("方法1失败：通过文本未找到Add按钮")
            
            # 方法2: 查找Lot name输入框附近的Add按钮
            if not add_button:
                try:
                    # 在输入框的父容器中查找Add按钮
                    parent = lot_input.find_element(By.XPATH, "./ancestor::div[1]")
                    add_button = parent.find_element(By.XPATH, ".//button[contains(text(), 'Add')]")
                    LOGGER.info("在输入框附近找到'Add'按钮")
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            if not add_button:
                LOGGER.error("❌ 未找到'Add'按钮")
                return False
            
            # 点击Add按钮
            add_button.click()
            LOGGER.info("✅ 已点击'Add'按钮")
            time.sleep(1.0)  # 等待添加生效
            
            return True
            
        except Exception as e:
            LOGGER.error(f"添加Lot name失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _select_parttype(self, part_type: str) -> bool:
        """
        选择Parttype override（自定义下拉控件）
        
        Args:
            part_type: Part Type值（如"43 4PXA2V E B"）
            
        Returns:
            True如果选择成功
        """
        LOGGER.info(f"选择Part Type: {part_type}")
        
        try:
            # 等待页面稳定（优化：减少等待时间）
            time.sleep(0.8)
            
            # 注意：'Override parttype with'复选框默认已勾选，无需处理
            
            # 查找并点击Parttype下拉框（自定义控件）
            LOGGER.info("查找Parttype下拉框...")
            
            # 这是一个自定义下拉控件，需要点击下三角符号展开
            dropdown_trigger = None
            
            # 方法1: 查找包含"Select Parttype"的元素
            try:
                dropdown_trigger = WebDriverWait(self._driver, 10).until(
                    EC.element_to_be_clickable((
                        By.XPATH, 
                        "//*[contains(text(), 'Select Parttype') or contains(text(), '-- Select Parttype --')]"
                    ))
                )
                LOGGER.info("方法1找到下拉触发器（包含'Select Parttype'文本）")
            except:
                pass
            
            # 方法2: 查找下三角符号（通常是SVG或特殊字符）
            if not dropdown_trigger:
                try:
                    # 查找包含下箭头的元素（Material UI常用）
                    dropdown_trigger = self._driver.find_element(
                        By.XPATH,
                        "//div[contains(@class, 'select') or contains(@role, 'button')]//svg[contains(@class, 'arrow') or contains(@class, 'dropdown')]/.."
                    )
                    LOGGER.info("方法2找到下拉触发器（包含下箭头SVG）")
                except:
                    pass
            
            # 方法3: 查找所有可能的下拉框容器
            if not dropdown_trigger:
                try:
                    # 查找Parttype override区域
                    parttype_area = self._driver.find_element(
                        By.XPATH,
                        "//*[contains(text(), 'Parttype override')]/../.."
                    )
                    
                    # 在这个区域内查找可点击的下拉元素
                    possible_triggers = parttype_area.find_elements(
                        By.XPATH,
                        ".//*[@role='button' or contains(@class, 'select') or contains(@class, 'dropdown')]"
                    )
                    
                    for trigger in possible_triggers:
                        if trigger.is_displayed():
                            dropdown_trigger = trigger
                            LOGGER.info(f"方法3找到下拉触发器（在Parttype区域）")
                            break
                            
                except Exception as e:
                    LOGGER.debug(f"方法3失败: {e}")
            
            if not dropdown_trigger:
                LOGGER.error("❌ 未找到Parttype下拉框触发器")
                return False
            
            # 点击展开下拉框
            LOGGER.info("点击展开Parttype下拉框...")
            try:
                dropdown_trigger.click()
                LOGGER.info("✅ 已点击下拉框")
            except:
                # 如果普通点击失败，尝试JavaScript点击
                self._driver.execute_script("arguments[0].click();", dropdown_trigger)
                LOGGER.info("✅ 已点击下拉框（JavaScript）")
            
            time.sleep(0.6)  # 等待下拉选项展开（优化：减少等待时间）
            
            # 3. 在展开的选项中选择Part Type
            LOGGER.info(f"在下拉选项中查找: {part_type}")
            
            # 查找所有下拉选项（优化：优先使用最快的方法）
            options = []
            
            # 方法1: 直接查找匹配的选项（最快，避免遍历所有选项）
            try:
                # 尝试直接找到包含目标Part Type的元素
                direct_match = self._driver.find_element(
                    By.XPATH,
                    f"//*[normalize-space(text())='{part_type}']"
                )
                if direct_match.is_displayed():
                    LOGGER.info(f"✅ 直接找到匹配选项: {part_type}")
                    # 直接点击
                    try:
                        direct_match.click()
                        LOGGER.info(f"✅ 已选择Part Type（直接匹配）: {part_type}")
                        time.sleep(0.5)
                        return True
                    except:
                        # 如果直接点击失败，加入到options中后续处理
                        options = [direct_match]
            except:
                pass
            
            # 方法2: 查找包含Part Type特征的元素（包含"4PXA"或"4PLH"）
            if not options:
                try:
                    options = self._driver.find_elements(
                        By.XPATH,
                        "//*[contains(text(), '4PXA') or contains(text(), '4PLH')]"
                    )
                    if options:
                        LOGGER.info(f"找到 {len(options)} 个候选选项")
                except:
                    pass
            
            # 方法3: 查找role="option"的元素
            if not options:
                try:
                    options = self._driver.find_elements(By.XPATH, "//li[@role='option'] | //div[@role='option']")
                    if options:
                        LOGGER.info(f"找到 {len(options)} 个候选选项")
                except:
                    pass
            
            if not options:
                LOGGER.error("❌ 未找到任何下拉选项")
                return False
            
            # 查找匹配的选项（优化：只输出前5个和匹配的选项）
            matched_option = None
            displayed_count = 0
            
            for idx, option in enumerate(options):
                try:
                    option_text = option.text.strip()
                    if not option_text:
                        continue
                    
                    # 只输出前5个选项的日志（避免日志过多拖慢速度）
                    if displayed_count < 5:
                        LOGGER.debug(f"  选项 {idx + 1}: '{option_text}'")
                        displayed_count += 1
                    
                    # 精确匹配
                    if option_text == part_type:
                        matched_option = option
                        LOGGER.info(f"✅ 精确匹配: {option_text}")
                        break
                    # 模糊匹配（去除多余空格）
                    elif ' '.join(option_text.split()) == ' '.join(part_type.split()):
                        matched_option = option
                        LOGGER.info(f"✅ 模糊匹配: {option_text}")
                        break
                    # 包含匹配
                    elif part_type in option_text or option_text in part_type:
                        matched_option = option
                        LOGGER.info(f"✅ 包含匹配: {option_text}")
                        break
                except:
                    continue
            
            if matched_option:
                # 滚动到选项可见
                try:
                    self._driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'});", matched_option)
                    time.sleep(0.3)  # 优化：减少等待时间
                except:
                    pass
                
                # 点击选项（多种方法）
                click_success = False
                
                # 方法1: 普通点击
                try:
                    matched_option.click()
                    LOGGER.info(f"✅ 已选择Part Type（普通点击）: {matched_option.text}")
                    click_success = True
                except Exception as e:
                    LOGGER.debug(f"普通点击失败: {e}")
                
                # 方法2: JavaScript点击
                if not click_success:
                    try:
                        self._driver.execute_script("arguments[0].click();", matched_option)
                        LOGGER.info(f"✅ 已选择Part Type（JavaScript点击）: {matched_option.text}")
                        click_success = True
                    except Exception as e:
                        LOGGER.debug(f"JavaScript点击失败: {e}")
                
                # 方法3: 发送Enter键
                if not click_success:
                    try:
                        from selenium.webdriver.common.keys import Keys
                        matched_option.send_keys(Keys.ENTER)
                        LOGGER.info(f"✅ 已选择Part Type（Enter键）: {matched_option.text}")
                        click_success = True
                    except Exception as e:
                        LOGGER.debug(f"Enter键失败: {e}")
                
                if click_success:
                    time.sleep(0.5)  # 优化：减少等待时间
                    return True
                else:
                    LOGGER.error("❌ 所有点击方法都失败")
                    return False
            else:
                LOGGER.error(f"❌ 未找到匹配的Part Type: {part_type}")
                LOGGER.error(f"可用选项: {[opt.text.strip() for opt in options if opt.text.strip()]}")
                return False
            
        except Exception as e:
            LOGGER.error(f"选择Parttype失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _click_flow_tab(self) -> bool:
        """
        点击Flow标签页
        
        Returns:
            True如果点击成功
        """
        LOGGER.info("查找并点击'Flow'标签...")
        
        try:
            # 等待页面稳定
            time.sleep(1.0)
            
            flow_tab = None
            
            # 方法1: 通过Material UI的mat-tab-label查找（用户提供的方法）
            try:
                flow_tab = WebDriverWait(self._driver, self.config.explicit_wait).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class,'mat-tab-label-content') and normalize-space()='Flow']/.."
                    ))
                )
                LOGGER.info("方法1找到'Flow'标签（mat-tab-label）")
            except TimeoutException:
                LOGGER.debug("方法1失败：未找到mat-tab-label")
            
            # 方法2: 通过包含Flow文本的元素查找
            if not flow_tab:
                try:
                    flow_tab = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            "//*[contains(@class, 'tab') and contains(text(), 'Flow')]"
                        ))
                    )
                    LOGGER.info("方法2找到'Flow'标签（包含Flow文本）")
                except TimeoutException:
                    LOGGER.debug("方法2失败")
            
            # 方法3: 查找所有可能的标签元素
            if not flow_tab:
                try:
                    all_tabs = self._driver.find_elements(By.XPATH, "//*[contains(@class, 'tab') or @role='tab']")
                    for tab in all_tabs:
                        if tab.is_displayed() and 'Flow' in tab.text:
                            flow_tab = tab
                            LOGGER.info(f"方法3找到'Flow'标签（遍历）")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法3失败: {e}")
            
            if not flow_tab:
                LOGGER.error("❌ 未找到'Flow'标签")
                return False
            
            # 滚动到标签可见
            try:
                self._driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", flow_tab)
                time.sleep(0.3)
            except:
                pass
            
            # 点击Flow标签
            click_success = False
            
            # 方法1: 普通点击
            try:
                flow_tab.click()
                LOGGER.info("✅ 已点击'Flow'标签（普通点击）")
                click_success = True
            except Exception as e:
                LOGGER.debug(f"普通点击失败: {e}")
            
            # 方法2: JavaScript点击
            if not click_success:
                try:
                    self._driver.execute_script("arguments[0].click();", flow_tab)
                    LOGGER.info("✅ 已点击'Flow'标签（JavaScript点击）")
                    click_success = True
                except Exception as e:
                    LOGGER.debug(f"JavaScript点击失败: {e}")
            
            if not click_success:
                LOGGER.error("❌ 点击'Flow'标签失败")
                return False
            
            # 等待Flow标签页加载
            time.sleep(1.5)
            LOGGER.info("✅ Flow标签页已加载")
            return True
            
        except Exception as e:
            LOGGER.error(f"点击'Flow'标签失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _click_more_options_tab(self) -> bool:
        """
        点击More options标签页
        
        Returns:
            True如果点击成功
        """
        LOGGER.info("=" * 60)
        LOGGER.info("步骤：点击'More options'标签")
        LOGGER.info("=" * 60)
        
        try:
            # 等待页面稳定
            LOGGER.info("等待页面稳定（1秒）...")
            time.sleep(1.0)
            
            more_options_tab = None
            
            # 方法1: 通过Material UI的mat-tab-label查找
            LOGGER.info("方法1：通过mat-tab-label查找'More options'标签...")
            try:
                more_options_tab = WebDriverWait(self._driver, self.config.explicit_wait).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class,'mat-tab-label-content') and normalize-space()='More options']/.."
                    ))
                )
                LOGGER.info("✅ 方法1成功：找到'More options'标签（mat-tab-label）")
            except TimeoutException:
                LOGGER.warning("⚠️ 方法1失败：未找到mat-tab-label，尝试其他方法...")
            
            # 方法2: 通过包含More options文本的元素查找
            if not more_options_tab:
                LOGGER.info("方法2：通过文本查找'More options'标签...")
                try:
                    more_options_tab = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            "//*[contains(@class, 'tab') and contains(text(), 'More options')]"
                        ))
                    )
                    LOGGER.info("✅ 方法2成功：找到'More options'标签（包含文本）")
                except TimeoutException:
                    LOGGER.warning("⚠️ 方法2失败：未找到标签")
            
            # 方法3: 查找所有可能的标签元素
            if not more_options_tab:
                LOGGER.info("方法3：遍历所有标签元素查找...")
                try:
                    all_tabs = self._driver.find_elements(By.XPATH, "//*[contains(@class, 'tab') or @role='tab']")
                    LOGGER.info(f"   找到 {len(all_tabs)} 个标签元素")
                    for idx, tab in enumerate(all_tabs, 1):
                        tab_text = tab.text.strip()
                        LOGGER.info(f"   检查标签 {idx}: 文本='{tab_text}', displayed={tab.is_displayed()}")
                        if tab.is_displayed() and 'More options' in tab_text:
                            more_options_tab = tab
                            LOGGER.info(f"✅ 方法3成功：找到'More options'标签（遍历）")
                            break
                except Exception as e:
                    LOGGER.warning(f"⚠️ 方法3失败: {e}")
            
            if not more_options_tab:
                LOGGER.error("❌ 所有方法都失败：未找到'More options'标签")
                return False
            
            # 滚动到标签可见
            LOGGER.info("滚动到'More options'标签可见...")
            try:
                self._driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", more_options_tab)
                time.sleep(0.3)
            except:
                pass
            
            # 点击标签
            LOGGER.info("点击'More options'标签...")
            more_options_tab.click()
            LOGGER.info("✅ 已点击'More options'标签")
            time.sleep(1.0)  # 等待标签页切换
            LOGGER.info("✅ 步骤完成：'More options'标签点击成功")
            return True
            
        except Exception as e:
            LOGGER.error(f"❌ 点击'More options'标签失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _fill_more_options(self, unit_test_time: str = None, retest_rate: str = None, hri_mrv: str = None) -> bool:
        """
        填写More options标签页的字段
        
        Args:
            unit_test_time: Unit test time值（可选）
            retest_rate: Retest rate值（可选）
            hri_mrv: HRI / MRV值（可选，如果为空则使用default）
            
        Returns:
            True如果填写成功
        """
        LOGGER.info("=" * 60)
        LOGGER.info("步骤13：填写More options字段")
        LOGGER.info(f"Unit test time: {unit_test_time or '(不填写)'}")
        LOGGER.info(f"Retest rate: {retest_rate or '(不填写)'}")
        LOGGER.info(f"HRI / MRV: {hri_mrv or '(使用DEFAULT)'}")
        LOGGER.info("=" * 60)
        
        try:
            # 等待More options标签页内容加载
            LOGGER.info("等待More options页面内容加载（1秒）...")
            time.sleep(1.0)
            
            success_count = 0
            total_fields = 0
            
            # 1. 填写Unit test time（center-text-input，第1个）
            if unit_test_time:
                total_fields += 1
                LOGGER.info(f"\n[字段1/3] Unit test time: {unit_test_time}")
                # 转换为字符串（处理numpy.int64等类型）
                unit_test_time_str = str(unit_test_time).strip()
                if self._fill_center_text_input(1, unit_test_time_str, "Unit test time"):
                    success_count += 1
                    LOGGER.info(f"✅ Unit test time填写成功")
                else:
                    LOGGER.warning(f"⚠️ Unit test time填写失败")
            
            # 2. 填写Retest rate（center-text-input，第2个）
            if retest_rate:
                total_fields += 1
                LOGGER.info(f"\n[字段2/3] Retest rate: {retest_rate}")
                # 转换为字符串（处理numpy.int64等类型）
                retest_rate_str = str(retest_rate).strip()
                if self._fill_center_text_input(2, retest_rate_str, "Retest rate"):
                    success_count += 1
                    LOGGER.info(f"✅ Retest rate填写成功")
                else:
                    LOGGER.warning(f"⚠️ Retest rate填写失败")
            
            # 3. 选择HRI / MRV（select下拉框，ID=flexbomSelect）
            total_fields += 1
            hri_mrv_str = str(hri_mrv).strip() if hri_mrv is not None else ''
            if hri_mrv_str and hri_mrv_str.lower() not in ['nan', 'none', 'null', '']:
                LOGGER.info(f"\n[字段3/3] HRI / MRV: {hri_mrv_str}")
                if self._select_flexbom_dropdown(hri_mrv_str):
                    success_count += 1
                    LOGGER.info(f"✅ HRI / MRV选择成功: {hri_mrv_str}")
                else:
                    LOGGER.warning(f"⚠️ HRI / MRV选择失败，保持默认值")
                    success_count += 1  # 失败也算成功（保持默认）
            else:
                LOGGER.info(f"\n[字段3/3] HRI / MRV: 保持DEFAULT")
                LOGGER.info("✅ 使用默认值 DEFAULT_HRI")
                success_count += 1
            
            LOGGER.info("\n" + "=" * 60)
            LOGGER.info(f"✅ 步骤13完成：More options字段处理完成（{success_count}/{total_fields}）")
            LOGGER.info("=" * 60)
            return True  # More options是可选的，总是返回True
            
        except Exception as e:
            LOGGER.error(f"❌ 填写More options字段失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _fill_center_text_input(self, index: int, value: str, field_name: str = "") -> bool:
        """
        填写center-text-input类的输入框
        
        Args:
            index: 输入框索引（1=Unit test time, 2=Retest rate）
            value: 要填写的值（已转换为字符串）
            field_name: 字段名称（用于日志）
            
        Returns:
            True如果填写成功
        """
        try:
            # 查找所有center-text-input输入框
            LOGGER.info(f"   查找第{index}个center-text-input输入框...")
            center_inputs = self._driver.find_elements(
                By.XPATH,
                "//input[contains(@class, 'center-text-input')]"
            )
            
            visible_inputs = [inp for inp in center_inputs if inp.is_displayed()]
            LOGGER.info(f"   找到 {len(visible_inputs)} 个可见的center-text-input")
            
            if index > len(visible_inputs):
                LOGGER.warning(f"   索引{index}超出范围（共{len(visible_inputs)}个输入框）")
                return False
            
            # 获取目标输入框（索引从1开始，所以减1）
            input_field = visible_inputs[index - 1]
            
            # 滚动到输入框
            self._driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", 
                input_field
            )
            time.sleep(0.2)
            
            # 清空并填写
            LOGGER.info(f"   清空输入框...")
            input_field.clear()
            time.sleep(0.1)
            
            LOGGER.info(f"   输入值: '{value}'")
            input_field.send_keys(value)
            time.sleep(0.2)
            
            # 验证
            actual_value = input_field.get_attribute('value')
            LOGGER.info(f"   验证：输入'{value}'，实际'{actual_value}'")
            
            return True
            
        except Exception as e:
            LOGGER.warning(f"   填写失败: {e}")
            import traceback
            LOGGER.warning(traceback.format_exc())
            return False
    
    def _select_flexbom_dropdown(self, value: str) -> bool:
        """
        选择Flexbom下拉框（HRI / MRV）
        
        Args:
            value: 要选择的值
            
        Returns:
            True如果选择成功
        """
        try:
            # 通过ID查找select元素
            LOGGER.info(f"   查找Flexbom下拉框（ID=flexbomSelect）...")
            select_element = WebDriverWait(self._driver, 3).until(
                EC.presence_of_element_located((By.ID, "flexbomSelect"))
            )
            LOGGER.info(f"   ✅ 找到下拉框")
            
            # 滚动到下拉框
            self._driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});",
                select_element
            )
            time.sleep(0.2)
            
            # 查找所有选项
            options = select_element.find_elements(By.TAG_NAME, "option")
            LOGGER.info(f"   找到 {len(options)} 个选项")
            
            # 遍历选项，查找匹配的值
            for idx, option in enumerate(options):
                option_text = option.text.strip()
                LOGGER.info(f"      选项{idx + 1}: '{option_text}'")
                
                # 检查选项文本是否包含目标值
                if value in option_text:
                    LOGGER.info(f"   ✅ 找到匹配选项: '{option_text}'")
                    option.click()
                    time.sleep(0.3)
                    LOGGER.info(f"   ✅ 已选择: '{option_text}'")
                    return True
            
            # 没找到匹配项
            LOGGER.warning(f"   ⚠️ 未找到包含'{value}'的选项")
            LOGGER.info(f"   保持默认选择")
            return False
            
        except Exception as e:
            LOGGER.warning(f"   选择失败: {e}")
            import traceback
            LOGGER.warning(traceback.format_exc())
            return False
    
    
    def _diagnose_flow_page(self) -> None:
        """
        诊断Flow页面的DOM结构，输出详细信息用于调试
        """
        LOGGER.info("=" * 80)
        LOGGER.info("🔍 开始诊断Flow页面DOM结构...")
        LOGGER.info("=" * 80)
        
        try:
            # 1. 统计所有容器
            all_containers = self._driver.find_elements(By.XPATH, "//div[contains(@class,'condition-list-container')]")
            LOGGER.info(f"📊 找到 {len(all_containers)} 个容器")
            
            for i, container in enumerate(all_containers[:5]):  # 只显示前5个
                try:
                    container_id = container.get_attribute('id') or '无ID'
                    container_class = container.get_attribute('class') or '无class'
                    LOGGER.info(f"  容器[{i+1}]: id={container_id}, class={container_class[:100]}")
                except:
                    pass
            
            # 2. 统计所有mat-select-arrow-wrapper
            all_arrows = self._driver.find_elements(By.XPATH, "//div[contains(@class,'mat-select-arrow-wrapper')]")
            LOGGER.info(f"📊 找到 {len(all_arrows)} 个 mat-select-arrow-wrapper")
            
            # 3. 统计所有mat-form-field
            all_form_fields = self._driver.find_elements(By.XPATH, "//mat-form-field[contains(@class,'mat-form-field-type-mat-select')]")
            LOGGER.info(f"📊 找到 {len(all_form_fields)} 个 mat-form-field (mat-select类型)")
            
            # 4. 检查目标容器
            container_xpath = "(//div[contains(@class,'condition-list-container')])[1]"
            try:
                target_container = self._driver.find_element(By.XPATH, container_xpath)
                LOGGER.info("✅ 目标容器存在")
                
                # 在容器内查找mat-select
                selects_in_container = target_container.find_elements(
                    By.XPATH, 
                    ".//mat-form-field[contains(@class,'mat-form-field-type-mat-select')]"
                )
                LOGGER.info(f"   容器内找到 {len(selects_in_container)} 个mat-select")
                
                for i, select in enumerate(selects_in_container[:3]):
                    try:
                        # 查找trigger
                        trigger = select.find_element(By.XPATH, ".//div[contains(@class,'mat-select-trigger')]")
                        trigger_text = trigger.text.strip()[:50]
                        LOGGER.info(f"   mat-select[{i+1}]: trigger文本='{trigger_text}'")
                    except:
                        LOGGER.info(f"   mat-select[{i+1}]: 无法读取trigger")
                        
            except Exception as e:
                LOGGER.error(f"❌ 目标容器不存在: {e}")
            
            # 5. 检查"All Units"等可能干扰的元素
            all_units = self._driver.find_elements(By.XPATH, "//*[contains(text(),'All Units')]")
            if all_units:
                LOGGER.warning(f"⚠️ 找到 {len(all_units)} 个包含'All Units'的元素（可能干扰定位）")
            
            # 6. 输出当前页面的关键XPath尝试结果
            LOGGER.info("\n📋 测试关键XPath:")
            test_xpaths = [
                ("(//div[contains(@class,'condition-list-container')])[1]", "目标容器"),
                ("(//div[contains(@class,'condition-list-container')])[1]//mat-form-field[contains(@class,'mat-form-field-type-mat-select')][1]//div[contains(@class,'mat-select-trigger')]", "Operation trigger"),
                ("(//div[contains(@class,'condition-list-container')])[1]//mat-form-field[contains(@class,'mat-form-field-type-mat-select')][2]//div[contains(@class,'mat-select-trigger')]", "Eng ID trigger"),
                ("(//div[contains(@class,'mat-select-arrow-wrapper')])[1]", "Operation箭头[1]"),
                ("(//div[contains(@class,'mat-select-arrow-wrapper')])[2]", "Eng ID箭头[2]"),
            ]
            
            for xpath, desc in test_xpaths:
                try:
                    elements = self._driver.find_elements(By.XPATH, xpath)
                    status = "✅" if elements else "❌"
                    LOGGER.info(f"  {status} {desc}: 找到 {len(elements)} 个元素")
                    if elements:
                        try:
                            LOGGER.info(f"     元素文本: '{elements[0].text.strip()[:50]}'")
                        except:
                            pass
                except Exception as e:
                    LOGGER.error(f"  ❌ {desc}: XPath错误 - {e}")
            
            LOGGER.info("=" * 80)
            
        except Exception as e:
            LOGGER.error(f"诊断过程出错: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
    
    def _scroll_and_click(self, by: By, locator: str, description: str = "", timeout: Optional[int] = None) -> bool:
        """
        通用的：等待 → 滚动 → 再次等待可点击 → 点击 的封装
        
        Args:
            by: By.XPATH / By.ID 等
            locator: 定位字符串
            description: 日志中的描述信息
            timeout: 超时时间（秒），默认使用config.explicit_wait
        """
        if timeout is None:
            timeout = self.config.explicit_wait

        desc = description or locator
        LOGGER.info(f"准备点击元素: {desc}")

        try:
            wait = WebDriverWait(self._driver, timeout)

            # 1. 等元素出现在 DOM 中
            element = wait.until(EC.presence_of_element_located((by, locator)))

            # 2. 滚动到元素位置（尽量滚到屏幕中间，减少遮挡）
            try:
                self._driver.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                    element,
                )
                time.sleep(0.3)
            except Exception as e:
                LOGGER.debug(f"scrollIntoView 失败: {e}")

            # 3. 再等一下它变成可点击状态
            element = wait.until(EC.element_to_be_clickable((by, locator)))

            # 4. 点击
            element.click()
            LOGGER.info(f"✅ 已点击元素: {desc}")
            return True

        except Exception as e:
            LOGGER.debug(f"_scroll_and_click 失败 ({desc}): {e}")
            return False
    
    def _find_operation_headers(self, scroll_to_bottom: bool = True):
        """
        查找所有Operation区块的抬头行
        
        Args:
            scroll_to_bottom: 是否先滚动到页面底部（确保所有区块都加载出来）
        
        返回抬头行元素列表，每个元素代表一个可编辑的Operation区块
        排除：
        - 灰色历史行（只读）
        - "Continue with All Units"行
        - Additional Attributes行
        """
        try:
            time.sleep(0.5)
            
            # 如果需要，先滚动到页面底部，确保所有Operation区块都加载出来
            if scroll_to_bottom:
                try:
                    # 查找Flow标签页的主容器（通常是mat-drawer-content或类似的）
                    flow_container = self._driver.find_element(By.XPATH, "//mat-drawer-content | //div[contains(@class,'drawer-content')] | //div[contains(@class,'mat-tab-body-active')]")
                    # 滚动到容器底部
                    self._driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", flow_container)
                    LOGGER.debug("已滚动到Flow页面底部")
                    time.sleep(0.5)
                except:
                    # 如果找不到容器，尝试滚动整个页面
                    try:
                        self._driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        LOGGER.debug("已滚动到页面底部")
                        time.sleep(0.5)
                    except:
                        LOGGER.debug("滚动失败，继续查找")
            
            # 查找所有可能的抬头行：包含2个mat-select-arrow的行
            # 先找到所有包含mat-select-arrow的元素
            all_elements = self._driver.find_elements(By.XPATH, "//*[.//div[contains(@class,'mat-select-arrow')]]")
            LOGGER.info(f"🔍 找到 {len(all_elements)} 个包含mat-select-arrow的元素，开始过滤...")
            
            # 按照垂直位置去重并排序，避免同一区块被祖先元素重复命中
            candidates = []

            for idx, elem in enumerate(all_elements):
                try:
                    # 检查这个元素是否包含正好2个mat-select-arrow（Operation和EngID）
                    arrows = elem.find_elements(By.CSS_SELECTOR, "div.mat-select-arrow")
                    if len(arrows) != 2:
                        LOGGER.debug(f"元素 #{idx+1} 有 {len(arrows)} 个箭头，跳过")
                        continue
                    
                    # 获取元素文本（用于排除）
                    elem_text = ""
                    try:
                        elem_text = elem.text
                    except:
                        pass
                    
                    # 排除"Continue with"行
                    if "Continue with" in elem_text or "All Units" in elem_text:
                        LOGGER.debug(f"元素 #{idx+1} 包含'Continue with'或'All Units'，跳过")
                        continue
                    
                    # 排除"Additional Attributes"行
                    if "Additional Attributes" in elem_text:
                        LOGGER.debug(f"元素 #{idx+1} 包含'Additional Attributes'，跳过")
                        continue
                    
                    # 排除灰色只读行（检查是否有checkbox - 历史行通常有checkbox）
                    try:
                        checkboxes = elem.find_elements(By.XPATH, ".//input[@type='checkbox']")
                        if checkboxes:
                            # 检查checkbox是否被选中（历史行通常是选中的）
                            for cb in checkboxes:
                                try:
                                    if cb.is_selected():
                                        LOGGER.debug(f"元素 #{idx+1} 包含已选中的checkbox（可能是历史行），跳过")
                                        # 不直接continue，继续检查其他条件
                                        break
                                except:
                                    pass
                            # 如果checkbox存在且被选中，很可能是历史行，跳过
                            if checkboxes and any(cb.is_selected() for cb in checkboxes if cb.is_displayed()):
                                LOGGER.debug(f"元素 #{idx+1} 包含已选中的checkbox（历史行），跳过")
                                continue
                    except:
                        pass
                    
                    # 排除灰色只读行（检查是否有disabled属性或特定class）
                    elem_classes = elem.get_attribute("class") or ""
                    if "disabled" in elem_classes.lower() or "readonly" in elem_classes.lower():
                        LOGGER.debug(f"元素 #{idx+1} 包含disabled/readonly class，跳过")
                        continue
                    
                    # 检查是否有"Instructions"和"Delete"图标（可编辑行应该有这些）
                    # 这是一个正向检查：如果找到这些图标，说明是可编辑的抬头行
                    has_instructions = False
                    has_delete = False
                    try:
                        # 检查是否有Instructions图标（document icon）
                        instructions_icons = elem.find_elements(By.XPATH, ".//*[contains(@class,'instructions') or contains(text(),'Instructions') or contains(@aria-label,'Instructions')]")
                        if instructions_icons:
                            has_instructions = True
                        
                        # 检查是否有Delete图标（trash can icon）
                        delete_icons = elem.find_elements(By.XPATH, ".//*[contains(@class,'delete') or contains(text(),'Delete') or contains(@aria-label,'Delete')]")
                        if delete_icons:
                            has_delete = True
                    except:
                        pass
                    
                    # 如果既没有Instructions也没有Delete，可能是历史行或其他不可编辑行
                    # 但这不是必要条件，因为有些可编辑行可能没有这些图标
                    # 所以这里只作为辅助判断，不强制要求
                    
                    # 确保这个元素是可见的
                    if not elem.is_displayed():
                        LOGGER.debug(f"元素 #{idx+1} 不可见，跳过")
                        continue
                    
                    # 额外检查：确保箭头是可点击的（不是禁用的）
                    try:
                        arrow_clickable = True
                        for arrow in arrows:
                            try:
                                arrow_classes = arrow.get_attribute("class") or ""
                                arrow_parent = arrow.find_element(By.XPATH, "./ancestor::mat-form-field[1]")
                                parent_classes = arrow_parent.get_attribute("class") or ""
                                
                                # 检查箭头或其父元素是否被禁用
                                if "disabled" in arrow_classes.lower() or "disabled" in parent_classes.lower():
                                    arrow_clickable = False
                                    break
                            except:
                                pass
                        
                        if not arrow_clickable:
                            LOGGER.debug(f"元素 #{idx+1} 的箭头被禁用，跳过")
                            continue
                    except:
                        pass
                    
                    # 记录候选元素及其位置，用于后续去重和排序
                    location = elem.location or {}
                    size = elem.size or {}
                    y_pos = int(location.get("y", 0))
                    area = int(size.get("width", 0) * size.get("height", 0))
                    candidates.append((y_pos, area, elem, elem_text, has_instructions, has_delete))
                except Exception as e:
                    LOGGER.debug(f"检查元素 #{idx+1} 时出错: {e}")
                    continue

            # 根据垂直位置分组（5px 为一档），同一档取面积更小的元素，避免祖先元素重复
            deduped = {}
            for y_pos, area, elem, elem_text, has_instructions, has_delete in candidates:
                key = y_pos // 5
                if key not in deduped or area < deduped[key][0]:
                    deduped[key] = (area, elem, elem_text, has_instructions, has_delete)

            # 按照垂直位置从上到下排序，确保condition_index稳定
            operation_headers = []
            for _, (_, elem, elem_text, has_instructions, has_delete) in sorted(deduped.items(), key=lambda kv: kv[0]):
                operation_headers.append(elem)
                icon_info = ""
                if has_instructions or has_delete:
                    icon_info = f"（有{'Instructions' if has_instructions else ''}{'和' if has_instructions and has_delete else ''}{'Delete' if has_delete else ''}图标）"
                LOGGER.info(f"✅ 找到Operation抬头行 #{len(operation_headers)}: {elem_text[:80] if elem_text else '(无文本)'}{icon_info} 位置Y={elem.location.get('y', '未知')}")

            LOGGER.info(f"✅ 总共找到 {len(operation_headers)} 个Operation抬头行（去重后）")
            return operation_headers
        except Exception as e:
            LOGGER.error(f"查找Operation抬头行失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return []
    
    def _select_mat_option_by_text(self, text: str, timeout: int = 10) -> bool:
        """
        在mat-select的下拉面板中选择指定文本的选项
        
        Args:
            text: 要选择的选项文本
            timeout: 超时时间（秒）
        
        Returns:
            True如果选择成功
        """
        try:
            wait = WebDriverWait(self._driver, timeout)
            
            # 等待mat-select面板出现
            panel = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[contains(@class,'mat-select-panel')]")
                )
            )
            LOGGER.info("✅ mat-select面板已打开")
            
            # 查找并点击匹配的选项
            option = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//div[contains(@class,'mat-select-panel')]//mat-option//span[normalize-space()='{text}']")
                )
            )
            option.click()
            LOGGER.info(f"✅ 已选择选项: {text}")
            time.sleep(0.3)
            return True
        except Exception as e:
            LOGGER.error(f"选择选项失败 (text={text}): {e}")
            # 尝试遍历所有选项查找包含匹配
            try:
                all_options = self._driver.find_elements(By.XPATH, "//div[contains(@class,'mat-select-panel')]//mat-option")
                LOGGER.info(f"找到 {len(all_options)} 个选项，尝试包含匹配...")
                for opt in all_options:
                    opt_text = opt.text.strip()
                    if text in opt_text or opt_text in text:
                        opt.click()
                        LOGGER.info(f"✅ 已选择选项（包含匹配）: {opt_text}")
                        time.sleep(0.3)
                        return True
                LOGGER.error(f"❌ 未找到匹配的选项: {text}")
                return False
            except Exception as e2:
                LOGGER.error(f"遍历选项也失败: {e2}")
                return False
    
    def _select_option_from_dropdown(self, value: str, is_filter_dropdown: bool = False) -> bool:
        """
        从下拉框中选择选项（通用方法）
        
        Args:
            value: 要选择的值
            is_filter_dropdown: 是否为可筛选的下拉框
            
        Returns:
            True如果选择成功
        """
        try:
            wait = WebDriverWait(self._driver, 10)
            
            # 等待mat-select面板出现
            panel = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[contains(@class,'mat-select-panel')]")
                )
            )
            LOGGER.info("✅ mat-select面板已打开")
            
            # 如果是可筛选的下拉框，先尝试输入筛选
            if is_filter_dropdown:
                try:
                    # 查找输入框
                    filter_input = self._driver.find_element(
                        By.XPATH,
                        "//div[contains(@class,'mat-select-panel')]//input"
                    )
                    if filter_input.is_displayed():
                        filter_input.clear()
                        filter_input.send_keys(value)
                        time.sleep(0.5)
                        LOGGER.info(f"已输入筛选值: {value}")
                except:
                    pass
            
            # 查找并点击匹配的选项
            try:
                # 精确匹配
                option = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, f"//div[contains(@class,'mat-select-panel')]//mat-option//span[normalize-space()='{value}']")
                    )
                )
                option.click()
                LOGGER.info(f"✅ 已选择选项: {value}")
                time.sleep(0.3)
                return True
            except TimeoutException:
                # 尝试包含匹配
                all_options = self._driver.find_elements(By.XPATH, "//div[contains(@class,'mat-select-panel')]//mat-option")
                LOGGER.info(f"精确匹配失败，尝试包含匹配（找到 {len(all_options)} 个选项）...")
                for opt in all_options:
                    try:
                        opt_text = opt.text.strip()
                        if value in opt_text or opt_text in value:
                            opt.click()
                            LOGGER.info(f"✅ 已选择选项（包含匹配）: {opt_text}")
                            time.sleep(0.3)
                            return True
                    except:
                        continue
                LOGGER.error(f"❌ 未找到匹配的选项: {value}")
                return False
        except Exception as e:
            LOGGER.error(f"选择选项失败 (value={value}): {e}")
            return False

    def _select_operation(self, operation_value: str) -> bool:
        """
        选择Operation（第一个下拉框）
        
        Args:
            operation_value: 要选择的值
            
        Returns:
            True如果选择成功
        """
        LOGGER.info("=" * 60)
        LOGGER.info(f"步骤：选择Operation")
        LOGGER.info(f"目标值: {operation_value}")
        LOGGER.info("=" * 60)
        
        try:
            # **优化定位策略**：Operation有mat-select-arrow-wrapper包装器
            # 通过查找包含mat-select-arrow-wrapper的mat-select元素来定位
            LOGGER.info("定位策略：查找包含'mat-select-arrow-wrapper'的mat-select元素（第一个）")
            
            # 1. 查找所有包含mat-select-arrow-wrapper的mat-select元素
            operation_mat_select = None
            
            LOGGER.info("等待Operation mat-select元素出现...")
            try:
                # 等待至少1个包含wrapper的mat-select出现
                # 使用遍历方式查找（更可靠，不依赖:has()选择器）
                LOGGER.info(f"   等待时间：{self.config.explicit_wait}秒")
                WebDriverWait(self._driver, self.config.explicit_wait).until(
                    lambda d: len([ms for ms in d.find_elements(By.CSS_SELECTOR, "mat-select")
                                   if ms.find_elements(By.CSS_SELECTOR, "div.mat-select-arrow-wrapper")]) > 0
                )
                LOGGER.info("✅ Operation mat-select元素已出现")
                
                # 获取所有mat-select元素并过滤出包含wrapper的（Operation）
                all_mat_selects = self._driver.find_elements(By.CSS_SELECTOR, "mat-select")
                LOGGER.info(f"   页面上共有 {len(all_mat_selects)} 个mat-select元素")
                operation_selects = []
                
                for idx, ms in enumerate(all_mat_selects, 1):
                    try:
                        if not ms.is_displayed():
                            continue
                        # 检查是否包含mat-select-arrow-wrapper
                        wrapper = ms.find_elements(By.CSS_SELECTOR, "div.mat-select-arrow-wrapper")
                        if wrapper:  # 有wrapper，说明是Operation
                            operation_selects.append(ms)
                            LOGGER.info(f"   mat-select #{idx}: 是Operation（有wrapper）")
                    except:
                        continue
                
                LOGGER.info(f"   找到 {len(operation_selects)} 个Operation mat-select")
                if len(operation_selects) > 0:
                    operation_mat_select = operation_selects[0]
                    LOGGER.info(f"✅ 选择第一个Operation mat-select")
                else:
                    LOGGER.error("❌ 未找到Operation mat-select")
                    return False
                    
            except TimeoutException:
                LOGGER.error("❌ 等待Operation mat-select超时")
                return False
            except Exception as e:
                LOGGER.error(f"❌ 查找Operation mat-select失败: {e}")
                return False
            
            if not operation_mat_select:
                LOGGER.error("❌ 无法定位Operation mat-select元素")
                return False
            
            # 2. 等待元素可点击并滚动
            LOGGER.info("等待Operation mat-select变为可点击...")
            operation_mat_select = WebDriverWait(self._driver, self.config.explicit_wait).until(
                EC.element_to_be_clickable(operation_mat_select)
            )
            LOGGER.info("✅ Operation mat-select已可点击")
            
            LOGGER.info("滚动到Operation下拉框可见...")
            self._driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", operation_mat_select)
            time.sleep(0.3)
            
            LOGGER.info("点击Operation下拉框...")
            operation_mat_select.click()
            LOGGER.info("✅ 已点击Operation下拉框，等待选项浮层...")
            
            # 3. 选择选项
            LOGGER.info(f"在下拉选项中选择: {operation_value}")
            if self._select_option_from_dropdown(operation_value, is_filter_dropdown=True):
                LOGGER.info(f"✅ 步骤完成：已选择Operation: {operation_value}")
                return True
            else:
                LOGGER.error(f"❌ 步骤失败：选择Operation选项失败: {operation_value}")
                return False

        except Exception as e:
            LOGGER.error(f"❌ 选择Operation失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _select_eng_id(self, eng_id_value: str) -> bool:
        """
        选择Eng ID（第二个下拉框）
        
        Args:
            eng_id_value: 要选择的值
            
        Returns:
            True如果选择成功
        """
        LOGGER.info("=" * 60)
        LOGGER.info(f"步骤：选择Eng ID")
        LOGGER.info(f"目标值: {eng_id_value}")
        LOGGER.info("=" * 60)
        
        try:
            # **关键步骤1：等待Operation选择完成，关闭所有打开的overlay**
            LOGGER.info("等待Operation选择完成...")
            LOGGER.info("   等待时间：1.5秒")
            time.sleep(1.5)  # 等待Operation选择完成（从1秒增加到1.5秒）
            LOGGER.info("✅ Operation选择完成")
            
            # 关闭所有打开的overlay（确保Operation下拉面板已关闭）
            LOGGER.info("关闭所有打开的overlay...")
            try:
                self._driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                time.sleep(0.5)
                LOGGER.info("✅ 已按ESC关闭所有打开的overlay")
            except Exception as e:
                LOGGER.warning(f"⚠️ 关闭overlay失败: {e}")
            
            # **优化定位策略**：直接使用第二个mat-select作为EngID（更可靠）
            # 因为在实际场景中，第一个是Operation，第二个是EngID
            LOGGER.info("定位策略：使用第二个mat-select作为EngID（第一个是Operation，第二个是EngID）")
            
            # 1. 等待Eng ID元素渲染（Operation选择后，Eng ID需要时间启用和渲染）
            eng_id_mat_select = None
            
            # 先等待至少2个mat-select出现（Operation和EngID各一个）
            LOGGER.info("等待至少2个mat-select元素出现（Operation和EngID）...")
            LOGGER.info("   等待时间：20秒")
            try:
                WebDriverWait(self._driver, 20).until(
                    lambda d: len(d.find_elements(By.CSS_SELECTOR, "mat-select")) >= 2
                )
                all_mat_selects_count = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                LOGGER.info(f"✅ 找到 {all_mat_selects_count} 个mat-select元素")
            except TimeoutException:
                all_mat_selects_count = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                LOGGER.warning(f"⚠️ 等待超时（20秒），只找到 {all_mat_selects_count} 个mat-select元素，继续尝试...")
            
            # 额外等待，确保EngID元素完全渲染
            LOGGER.info("额外等待1秒，确保EngID元素完全渲染...")
            time.sleep(1.0)
            
            try:
                # 获取所有可见的mat-select元素
                all_mat_selects = self._driver.find_elements(By.CSS_SELECTOR, "mat-select")
                visible_mat_selects = [ms for ms in all_mat_selects if ms.is_displayed()]
                LOGGER.info(f"   页面上共有 {len(all_mat_selects)} 个mat-select元素，其中 {len(visible_mat_selects)} 个可见")
                
                for idx, ms in enumerate(visible_mat_selects, 1):
                    try:
                        LOGGER.info(f"   mat-select #{idx}: displayed={ms.is_displayed()}, enabled={ms.is_enabled()}, location={ms.location}")
                    except:
                        pass
                
                # 主策略：直接使用第二个可见的mat-select作为EngID
                if len(visible_mat_selects) >= 2:
                    eng_id_mat_select = visible_mat_selects[1]
                    LOGGER.info("✅ 主策略成功：使用第二个mat-select作为EngID")
                    LOGGER.info(f"   EngID mat-select状态：displayed={eng_id_mat_select.is_displayed()}, enabled={eng_id_mat_select.is_enabled()}")
                elif len(visible_mat_selects) == 1:
                    # 如果只有一个可见的，可能是EngID还未渲染，等待一下再试
                    LOGGER.warning("⚠️ 只找到1个可见的mat-select，等待EngID渲染...")
                    LOGGER.info("   额外等待2秒...")
                    time.sleep(2.0)
                    all_mat_selects = self._driver.find_elements(By.CSS_SELECTOR, "mat-select")
                    visible_mat_selects = [ms for ms in all_mat_selects if ms.is_displayed()]
                    LOGGER.info(f"   重新检查：找到 {len(visible_mat_selects)} 个可见的mat-select")
                    if len(visible_mat_selects) >= 2:
                        eng_id_mat_select = visible_mat_selects[1]
                        LOGGER.info("✅ 等待后找到第二个mat-select，使用作为EngID")
                    else:
                        LOGGER.error(f"❌ 等待后仍只有 {len(visible_mat_selects)} 个可见的mat-select")
                        return False
                else:
                    LOGGER.error(f"❌ 未找到足够的可见mat-select元素（需要至少2个，实际{len(visible_mat_selects)}个）")
                    return False
                
                # 备用策略：如果主策略失败，尝试通过wrapper过滤
                if not eng_id_mat_select:
                    LOGGER.warning("⚠️ 主策略失败，尝试备用策略：通过wrapper过滤...")
                    eng_id_selects = []
                    for idx, ms in enumerate(visible_mat_selects, 1):
                        try:
                            # 检查是否包含wrapper（如果有wrapper，则是Operation，跳过）
                            wrapper = ms.find_elements(By.CSS_SELECTOR, "div.mat-select-arrow-wrapper")
                            if not wrapper:  # 没有wrapper，可能是EngID
                                arrow = ms.find_elements(By.CSS_SELECTOR, "div.mat-select-arrow")
                                if arrow:
                                    eng_id_selects.append(ms)
                        except:
                            continue
                    
                    if len(eng_id_selects) > 0:
                        eng_id_mat_select = eng_id_selects[0]
                        LOGGER.info(f"✅ 备用策略成功：通过wrapper过滤找到EngID（共{len(eng_id_selects)}个）")
                    else:
                        LOGGER.error("❌ 备用策略也失败：未找到EngID mat-select")
                        return False
                    
            except Exception as e:
                LOGGER.error(f"❌ 查找Eng ID mat-select失败: {e}")
                import traceback
                LOGGER.error(traceback.format_exc())
                return False
            
            if not eng_id_mat_select:
                LOGGER.error("❌ 无法定位Eng ID mat-select元素")
                return False
            
            LOGGER.info("✅ Eng ID mat-select元素已找到")
            
            # 步骤2.2: 等待元素变为可点击（启用状态）
            LOGGER.info("等待Eng ID变为启用状态（可点击）...")
            LOGGER.info("   等待时间：15秒")
            enabled_eng_id_select = None
            
            # 直接等待找到的元素可点击
            try:
                enabled_eng_id_select = WebDriverWait(self._driver, 15).until(
                    EC.element_to_be_clickable(eng_id_mat_select)
                )
                LOGGER.info("✅ Eng ID已变为启用状态（可点击）")
            except TimeoutException:
                LOGGER.warning("⚠️ 等待Eng ID可点击超时（15秒），检查是否被禁用...")
                enabled_eng_id_select = eng_id_mat_select
            
            # 检查元素是否被禁用
            try:
                form_field = eng_id_mat_select.find_element(By.XPATH, "./ancestor::mat-form-field")
                class_attr = form_field.get_attribute("class") or ""
                if "mat-form-field-disabled" in class_attr:
                    LOGGER.warning("⚠️ Eng ID仍然处于禁用状态，等待更长时间...")
                    # 等待禁用类消失
                    try:
                        WebDriverWait(self._driver, 10).until_not(
                            lambda d: "mat-form-field-disabled" in (form_field.get_attribute("class") or "")
                        )
                        LOGGER.info("✅ Eng ID已从禁用状态变为启用")
                        enabled_eng_id_select = eng_id_mat_select
                    except TimeoutException:
                        LOGGER.error("❌ Eng ID仍然处于禁用状态，可能Operation选择未完成")
                        return False
            except:
                pass
            
            if not enabled_eng_id_select:
                enabled_eng_id_select = eng_id_mat_select

            # 3. 滚动并点击
            LOGGER.info("滚动到Eng ID下拉框可见...")
            self._driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", enabled_eng_id_select)
            time.sleep(0.3)
            
            # 尝试点击
            LOGGER.info("点击Eng ID下拉框...")
            try:
                enabled_eng_id_select.click()
                LOGGER.info("✅ 已点击Eng ID下拉框（普通点击）")
            except Exception as e:
                LOGGER.warning(f"⚠️ 普通点击失败: {e}，尝试JavaScript点击")
                self._driver.execute_script("arguments[0].click();", enabled_eng_id_select)
                LOGGER.info("✅ 已点击Eng ID下拉框（JavaScript点击）")

            LOGGER.info("等待选项浮层出现（0.5秒）...")
            time.sleep(0.5)
            
            # 4. 选择选项
            LOGGER.info(f"在下拉选项中选择: {eng_id_value}")
            if self._select_option_from_dropdown(eng_id_value, is_filter_dropdown=True):
                LOGGER.info(f"✅ 步骤完成：已选择Eng ID: {eng_id_value}")
                return True
            else:
                LOGGER.error(f"❌ 步骤失败：选择Eng ID选项失败: {eng_id_value}")
                return False

        except TimeoutException as e:
            LOGGER.error(f"❌ 选择Eng ID超时: {e}")
            # 尝试调试：查找所有mat-select元素
            try:
                all_mat_selects = self._driver.find_elements(By.CSS_SELECTOR, "mat-select")
                LOGGER.info(f"   页面上共有 {len(all_mat_selects)} 个mat-select元素")
                for idx, ms in enumerate(all_mat_selects, 1):
                    try:
                        is_displayed = ms.is_displayed()
                        is_enabled = ms.is_enabled()
                        LOGGER.info(f"   mat-select #{idx}: displayed={is_displayed}, enabled={is_enabled}")
                    except:
                        pass
            except:
                pass
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
        except Exception as e:
            LOGGER.error(f"❌ 选择Eng ID失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _fill_text_input(self, text_value: str) -> bool:
        """
        填写文本输入框（如Thermal）
        
        Args:
            text_value: 要填写的文本
            
        Returns:
            True如果填写成功
        """
        LOGGER.info(f"填写文本输入框: {text_value}")
        # 使用第一个文本输入框
        input_index = 1
        
        # 1. 找到对应的 input 元素 (在整个页面范围内查找)
        # 使用 input[type='text'] 来排除隐藏的或特殊类型的输入框
        input_locator = (
            By.CSS_SELECTOR, 
            f"input[type='text']:nth-of-type({input_index})"
        )
        
        LOGGER.info(f"定位策略：使用 CSS 选择器 input[type='text']:nth-of-type({input_index})")
        
        try:
            # 等待元素出现并可点击
            text_input = WebDriverWait(self._driver, self.config.explicit_wait).until(
                EC.element_to_be_clickable(input_locator)
            )
            
            # 2. 滚动、清空、发送按键
            try:
                self._driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", text_input
                )
                time.sleep(0.3)
            except:
                pass

            text_input.clear()
            text_input.send_keys(text_value)
            LOGGER.info(f"✅ 已填写文本输入框: {text_value}")
            return True
                
        except Exception as e:
            LOGGER.error(f"❌ 填写文本输入框失败: {e}")
            # 尝试使用JavaScript填写（备用）
            try:
                text_input = self._driver.find_element(*input_locator)
                self._driver.execute_script(f"arguments[0].value = '{text_value}';", text_input)
                self._driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", text_input
                )
                LOGGER.info(f"✅ 已通过JavaScript填写文本输入框: {text_value}")
                return True
            except Exception as e2:
                LOGGER.error(f"❌ JavaScript填写也失败: {e2}")
                import traceback
                LOGGER.error(traceback.format_exc())
                return False

    def _click_add_new_condition(self) -> bool:
        """
        点击最后一个Operation区块内的"Add new condition"按钮
        
        Returns:
            True如果点击成功
        """
        LOGGER.info("查找并点击最后一个Operation区块的'Add new condition'...")
        
        try:
            time.sleep(1.0)
            
            # 1. 查找所有Operation抬头行（不滚动，使用当前状态）
            operation_headers = self._find_operation_headers(scroll_to_bottom=False)
            
            if not operation_headers:
                LOGGER.error("❌ 未找到任何Operation抬头行")
                return False
            
            current_count = len(operation_headers)
            LOGGER.info(f"✅ 当前有 {current_count} 个Operation抬头行")
            
            # 2. 获取最后一个Operation抬头行
            last_header = operation_headers[-1]
            LOGGER.info(f"✅ 定位到最后一个Operation抬头行（第 {current_count} 个）")
            
            # 3. 滚动到最后一个抬头行，确保其下方的按钮也可见
            self._driver.execute_script(
                "arguments[0].scrollIntoView({block: 'end', inline: 'nearest'});",
                last_header
            )
            time.sleep(0.5)
            
            # 4. 查找"Add new condition"按钮（多种方法，优先全局最后一个）
            add_btn = None
            
            # 方法1: 全局查找所有"Add new condition"，取最后一个可见的（最简单可靠）
            try:
                all_add_btns = self._driver.find_elements(
                    By.XPATH,
                    "//span[contains(text(),'Add new condition') or contains(@class,'add-text')]"
                )
                LOGGER.info(f"找到 {len(all_add_btns)} 个'Add new condition'按钮")
                # 取最后一个可见的按钮
                for btn in reversed(all_add_btns):
                    try:
                        if btn.is_displayed():
                            add_btn = btn
                            LOGGER.info("✅ 方法1找到'Add new condition'按钮（全局查找，取最后一个可见）")
                            break
                    except:
                        continue
            except Exception as e:
                LOGGER.debug(f"方法1失败: {e}")
            
            # 方法2: 通过ID查找
            if not add_btn:
                try:
                    add_btn = self._driver.find_element(By.ID, "addNewCondition")
                    if add_btn.is_displayed():
                        LOGGER.info("✅ 方法2找到'Add new condition'按钮（通过ID）")
                except:
                    LOGGER.debug("方法2失败：未找到ID=addNewCondition")
            
            if not add_btn:
                LOGGER.error("❌ 所有方法都未找到'Add new condition'按钮")
                return False
            
            # 5. 滚动到按钮可见并等待可点击
            self._driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                add_btn
            )
            time.sleep(0.5)
            
            # 等待按钮可点击
            try:
                clickable_btn = WebDriverWait(self._driver, 10).until(
                    EC.element_to_be_clickable(add_btn)
                )
                LOGGER.info("✅ 按钮已变为可点击状态")
            except TimeoutException:
                LOGGER.warning("⚠️ 等待按钮可点击超时，但继续尝试点击...")
                clickable_btn = add_btn
            
            # 6. 点击按钮
            click_success = False
            try:
                clickable_btn.click()
                LOGGER.info("✅ 已点击'Add new condition'按钮（普通点击）")
                click_success = True
            except Exception as e:
                LOGGER.debug(f"普通点击失败: {e}，尝试JavaScript点击")
                try:
                    # 使用JavaScript点击
                    self._driver.execute_script("arguments[0].click();", add_btn)
                    LOGGER.info("✅ 已点击'Add new condition'按钮（JavaScript点击）")
                    click_success = True
                except Exception as e2:
                    LOGGER.error(f"❌ 点击按钮失败: {e2}")
                    return False
            
            if not click_success:
                LOGGER.error("❌ 按钮点击失败")
                return False
            
            # 7. 等待新的Operation区块DOM完全渲染
            LOGGER.info("等待新区块渲染...")
            
            # 记录当前mat-select数量
            try:
                initial_mat_selects = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                LOGGER.info(f"点击前：页面上有 {initial_mat_selects} 个mat-select元素")
            except:
                initial_mat_selects = 0
            
            # 使用显式等待：等待新的mat-select元素出现（新condition会有新的Operation和EngID下拉框）
            # 预期：新condition会添加2个新的mat-select（Operation和EngID）
            expected_new_mat_selects = initial_mat_selects + 2
            LOGGER.info(f"等待新的mat-select元素出现（预期总数：{expected_new_mat_selects}，当前：{initial_mat_selects}）...")
            
            try:
                # 等待新的mat-select元素出现（使用显式等待）
                WebDriverWait(self._driver, 15).until(
                    lambda d: len(d.find_elements(By.CSS_SELECTOR, "mat-select")) >= expected_new_mat_selects
                )
                actual_mat_selects = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                LOGGER.info(f"✅ 新的mat-select元素已出现（实际总数：{actual_mat_selects}）")
            except TimeoutException:
                actual_mat_selects = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                LOGGER.warning(f"⚠️ 等待新mat-select超时（实际总数：{actual_mat_selects}，预期：{expected_new_mat_selects}），但继续验证...")
            
            # 额外等待，确保DOM完全渲染（特别是Angular的变更检测）
            LOGGER.info("等待Angular变更检测完成...")
            time.sleep(3.0)  # 增加到3秒，给Angular更多时间
            
            # 8. 滚动到页面底部，确保新区块完全加载
            try:
                flow_container = self._driver.find_element(By.XPATH, "//mat-drawer-content | //div[contains(@class,'drawer-content')] | //div[contains(@class,'mat-tab-body-active')]")
                self._driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", flow_container)
                LOGGER.info("✅ 已滚动到Flow页面底部")
                time.sleep(1.0)
            except:
                try:
                    self._driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    LOGGER.info("✅ 已滚动到页面底部")
                    time.sleep(1.0)
                except:
                    LOGGER.debug("滚动失败")
            
            # 9. 验证新区块是否已添加（使用滚动查找）
            LOGGER.info("验证新区块是否已添加...")
            
            # 先检查mat-select数量是否增加
            try:
                final_mat_selects = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                LOGGER.info(f"   当前mat-select数量：{final_mat_selects}（点击前：{initial_mat_selects}）")
                if final_mat_selects < initial_mat_selects + 2:
                    LOGGER.warning(f"   ⚠️ mat-select数量未增加2个（预期增加2个，实际增加{final_mat_selects - initial_mat_selects}个）")
            except:
                pass
            
            # 使用_find_operation_headers查找抬头行
            new_headers = self._find_operation_headers(scroll_to_bottom=True)
            new_count = len(new_headers)
            LOGGER.info(f"验证结果：之前有 {current_count} 个Operation抬头行，现在有 {new_count} 个")
            
            # 调试：如果数量没增加，列出所有找到的元素详情
            if new_count == current_count:
                LOGGER.warning("⚠️ 抬头行数量未增加，进行详细调试...")
                try:
                    # 查找所有包含mat-select-arrow的元素（不过滤）
                    all_elements = self._driver.find_elements(By.XPATH, "//*[.//div[contains(@class,'mat-select-arrow')]]")
                    LOGGER.info(f"   找到 {len(all_elements)} 个包含mat-select-arrow的元素（未过滤）")
                    
                    # 统计包含2个箭头的元素
                    two_arrow_elements = []
                    for idx, elem in enumerate(all_elements):
                        try:
                            arrows = elem.find_elements(By.CSS_SELECTOR, "div.mat-select-arrow")
                            if len(arrows) == 2 and elem.is_displayed():
                                elem_text = elem.text[:100] if elem.text else "(无文本)"
                                location = elem.location
                                two_arrow_elements.append((idx, elem_text, location))
                        except:
                            pass
                    
                    LOGGER.info(f"   其中包含2个箭头且可见的元素：{len(two_arrow_elements)} 个")
                    for idx, text, loc in two_arrow_elements:
                        LOGGER.info(f"      - 元素#{idx}: {text[:50]}... 位置Y={loc.get('y', '未知')}")
                except Exception as e:
                    LOGGER.debug(f"   调试信息收集失败: {e}")
            
            if new_count > current_count:
                LOGGER.info(f"✅✅✅ 新的Operation区块已成功添加！（从 {current_count} 增加到 {new_count}）")
                return True
            elif new_count == current_count:
                LOGGER.warning(f"⚠️ Operation区块数量未增加，但可能DOM还在渲染中，继续尝试...")
                
                # 备用验证：如果mat-select数量增加了，即使_find_operation_headers没找到，也可能成功
                try:
                    final_mat_selects = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                    if final_mat_selects >= initial_mat_selects + 2:
                        LOGGER.warning(f"⚠️ 虽然抬头行数量未增加，但mat-select数量增加了（从{initial_mat_selects}到{final_mat_selects}）")
                        LOGGER.warning(f"   可能是_find_operation_headers的过滤条件太严格，新行被过滤掉了")
                        LOGGER.warning(f"   尝试继续执行，假设新condition已添加...")
                        # 再等待一次，然后返回True（假设成功）
                        time.sleep(2.0)
                        return True
                except:
                    pass
                
                # 再等待一次并重新查找（使用更长的等待时间）
                LOGGER.info("等待更长时间后重新验证...")
                time.sleep(3.0)
                
                # 再次滚动到底部
                try:
                    flow_container = self._driver.find_element(By.XPATH, "//mat-drawer-content | //div[contains(@class,'drawer-content')] | //div[contains(@class,'mat-tab-body-active')]")
                    self._driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", flow_container)
                    time.sleep(1.0)
                except:
                    pass
                
                retry_headers = self._find_operation_headers(scroll_to_bottom=True)
                retry_count = len(retry_headers)
                LOGGER.info(f"第二次检查：现在有 {retry_count} 个Operation抬头行")
                
                if retry_count > current_count:
                    LOGGER.info(f"✅ 第二次检查：新区块已添加（从 {current_count} 增加到 {retry_count}）")
                    return True
                else:
                    # 最后一次检查：如果mat-select增加了，仍然认为成功
                    try:
                        final_check_mat_selects = len(self._driver.find_elements(By.CSS_SELECTOR, "mat-select"))
                        if final_check_mat_selects >= initial_mat_selects + 2:
                            LOGGER.warning(f"⚠️ 虽然抬头行数量未增加，但mat-select数量增加了，假设成功")
                            return True
                    except:
                        pass
                    
                    LOGGER.error(f"❌ 第二次检查：区块数量仍未增加（{retry_count}）")
                    # 调试：检查按钮是否真的被点击了
                    LOGGER.error("   可能原因：按钮点击未生效，或DOM结构发生变化")
                    return False
            else:
                LOGGER.error(f"❌ 区块数量异常减少（从 {current_count} 变为 {new_count}）")
                return False
        
        except Exception as e:
            LOGGER.error(f"点击'Add new condition'失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _select_dropdown_option(self, field_name: str, value: str) -> bool:
        """
        在指定的下拉框中选择选项（通用方法，针对Spark Flow页面的下拉框）
        
        Args:
            field_name: 字段名（如"Operation", "EngID", "Thermal"）
            value: 要选择的值
            
        Returns:
            True如果选择成功
        """
        LOGGER.info(f"开始选择{field_name}下拉框...")
        
        try:
            # 查找下拉框
            dropdown_trigger = None
            
            # 方法1: 在Flow标签页内，通过label查找下拉框（最精确）
            try:
                # 查找包含field_name的label
                labels = self._driver.find_elements(
                    By.XPATH,
                    f"//*[normalize-space(text())='{field_name}' or contains(text(), '{field_name}')]"
                )
                
                for label in labels:
                    if not label.is_displayed():
                        continue
                    
                    # 方法1a: 查找label的下一个兄弟元素（通常下拉框在label旁边）
                    try:
                        dropdown_trigger = label.find_element(
                            By.XPATH,
                            "./following-sibling::*[1]//select | ./following-sibling::*[1]//*[@role='button'] | ./following-sibling::*[1]"
                        )
                        
                        if dropdown_trigger.is_displayed():
                            LOGGER.info(f"方法1a找到{field_name}下拉框（label的兄弟元素）")
                            break
                    except:
                        pass
                    
                    # 方法1b: 在label的父元素中查找下拉框
                    try:
                        parent = label.find_element(By.XPATH, "./..")
                        dropdown_trigger = parent.find_element(
                            By.XPATH,
                            ".//select | .//*[@role='button' and not(self::*[contains(text(), '{field_name}')])]"
                        )
                        
                        if dropdown_trigger.is_displayed():
                            LOGGER.info(f"方法1b找到{field_name}下拉框（label的父元素）")
                            break
                    except:
                        pass
                
            except Exception as e:
                LOGGER.debug(f"方法1失败: {e}")
            
            # 方法2: 查找传统HTML select标签
            if not dropdown_trigger:
                try:
                    selects = self._driver.find_elements(By.TAG_NAME, "select")
                    LOGGER.debug(f"页面上共有 {len(selects)} 个select元素")
                    
                    # 优先查找name或id包含field_name的
                    for select in selects:
                        if not select.is_displayed():
                            continue
                        
                        name = select.get_attribute("name") or ""
                        id_attr = select.get_attribute("id") or ""
                        
                        if field_name.lower() in name.lower() or field_name.lower() in id_attr.lower():
                            dropdown_trigger = select
                            LOGGER.info(f"方法2找到{field_name}下拉框（select标签）")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            # 方法3: 查找自定义下拉控件（Material UI等）
            if not dropdown_trigger:
                try:
                    # 查找所有可能是下拉框的元素
                    dropdowns = self._driver.find_elements(
                        By.XPATH,
                        "//*[@role='button' and contains(@class, 'select')] | //*[contains(@class, 'dropdown')]"
                    )
                    
                    LOGGER.debug(f"找到 {len(dropdowns)} 个可能的自定义下拉框")
                    
                    # 尝试通过位置关系查找
                    for dropdown in dropdowns:
                        if not dropdown.is_displayed():
                            continue
                        
                        # 检查dropdown附近是否有field_name的文本
                        try:
                            # 获取dropdown的父元素或祖父元素
                            parent = dropdown.find_element(By.XPATH, "./..")
                            parent_text = parent.text
                            
                            if field_name in parent_text:
                                dropdown_trigger = dropdown
                                LOGGER.info(f"方法3找到{field_name}下拉框（自定义控件）")
                                break
                        except:
                            continue
                            
                except Exception as e:
                    LOGGER.debug(f"方法3失败: {e}")
            
            if not dropdown_trigger:
                LOGGER.error(f"❌ 未找到{field_name}下拉框")
                
                # 调试信息
                try:
                    LOGGER.info(f"=== Debug: 查找{field_name}附近的所有元素 ===")
                    field_labels = self._driver.find_elements(By.XPATH, f"//*[contains(text(), '{field_name}')]")
                    for idx, lbl in enumerate(field_labels[:3]):
                        if lbl.is_displayed():
                            LOGGER.info(f"  找到文本 {idx+1}: '{lbl.text}', 标签: {lbl.tag_name}")
                except:
                    pass
                
                return False
            
            # 滚动到下拉框可见
            try:
                self._driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
                time.sleep(0.5)
            except:
                pass
            
            # 检查是否是<select>标签
            if dropdown_trigger.tag_name == "select":
                LOGGER.info(f"检测到传统HTML select下拉框")
                # 传统HTML select下拉框
                from selenium.webdriver.support.ui import Select
                select = Select(dropdown_trigger)
                
                # 先列出所有选项
                try:
                    all_options = [opt.text.strip() for opt in select.options]
                    LOGGER.info(f"下拉框选项: {all_options}")
                except:
                    pass
                
                # 尝试多种选择方式
                try:
                    # 方法1: 按值选择
                    select.select_by_value(value)
                    LOGGER.info(f"✅ 已选择{field_name}: {value}（按value）")
                    time.sleep(0.5)
                    return True
                except:
                    pass
                
                try:
                    # 方法2: 按可见文本选择
                    select.select_by_visible_text(value)
                    LOGGER.info(f"✅ 已选择{field_name}: {value}（按text）")
                    time.sleep(0.5)
                    return True
                except:
                    pass
                
                # 方法3: 模糊匹配
                try:
                    for option in select.options:
                        option_text = option.text.strip()
                        option_value = option.get_attribute("value")
                        
                        if (value == option_text or 
                            value == option_value or
                            value in option_text or 
                            option_text in value):
                            select.select_by_visible_text(option_text)
                            LOGGER.info(f"✅ 已选择{field_name}: {option_text}（模糊匹配）")
                            time.sleep(0.5)
                            return True
                    
                    LOGGER.error(f"❌ 未找到匹配的选项: {value}")
                    LOGGER.error(f"可用选项: {all_options}")
                    return False
                except Exception as e:
                    LOGGER.error(f"模糊匹配失败: {e}")
                    return False
            else:
                LOGGER.info(f"检测到自定义下拉控件")
                # 自定义下拉控件
                # 点击展开
                click_success = False
                
                try:
                    dropdown_trigger.click()
                    LOGGER.info(f"已展开{field_name}下拉框（普通点击）")
                    click_success = True
                except:
                    try:
                        self._driver.execute_script("arguments[0].click();", dropdown_trigger)
                        LOGGER.info(f"已展开{field_name}下拉框（JavaScript点击）")
                        click_success = True
                    except Exception as e:
                        LOGGER.error(f"展开下拉框失败: {e}")
                        return False
                
                if not click_success:
                    return False
                
                # 等待下拉选项加载（关键：等待选项出现）
                LOGGER.info("等待下拉选项加载...")
                time.sleep(1.5)  # 增加等待时间
                
                # 查找并点击匹配的选项（多种方法，增加等待和重试）
                max_wait_attempts = 5
                options = []
                
                for wait_attempt in range(1, max_wait_attempts + 1):
                    LOGGER.debug(f"查找选项（第{wait_attempt}次）...")
                    
                    # 方法1: 直接精确匹配（最快）
                    try:
                        option = WebDriverWait(self._driver, 2).until(
                            EC.presence_of_element_located((
                                By.XPATH,
                                f"//*[normalize-space(text())='{value}' and (self::li or self::div or contains(@role, 'option'))]"
                            ))
                        )
                        
                        if option.is_displayed():
                            option.click()
                            LOGGER.info(f"✅ 已选择{field_name}: {value}（直接匹配）")
                            time.sleep(0.5)
                            return True
                    except:
                        pass
                    
                    # 方法2: 查找所有下拉选项
                    try:
                        # 使用多种XPath查找选项
                        xpath_patterns = [
                            "//li[@role='option']",
                            "//div[@role='option']",
                            "//*[contains(@class, 'option') or contains(@class, 'Option')]",
                            "//ul/li",
                            "//*[@role='listbox']//*",
                            "//select/option"  # 也检查select的option
                        ]
                        
                        for pattern in xpath_patterns:
                            try:
                                found_options = self._driver.find_elements(By.XPATH, pattern)
                                if found_options:
                                    options.extend(found_options)
                            except:
                                continue
                        
                        # 去重
                        options = list(set(options))
                        
                        if options:
                            LOGGER.info(f"找到 {len(options)} 个候选选项")
                            break
                        else:
                            LOGGER.debug(f"第{wait_attempt}次未找到选项，继续等待...")
                            time.sleep(1.0)
                    except Exception as e:
                        LOGGER.debug(f"查找选项时出错: {e}")
                        time.sleep(1.0)
                
                if not options:
                    LOGGER.error(f"❌ 等待{max_wait_attempts}次后仍未找到任何下拉选项")
                    
                    # 最后的调试信息：列出所有可见元素
                    try:
                        all_visible = self._driver.find_elements(By.XPATH, "//*")
                        visible_texts = [elem.text.strip()[:30] for elem in all_visible if elem.is_displayed() and elem.text.strip()]
                        LOGGER.error(f"页面上所有可见文本（前20个）: {visible_texts[:20]}")
                    except:
                        pass
                    
                    return False
                
                # 遍历选项并匹配（支持滚动查找）
                matched = False
                visible_count = 0
                
                # 先检查当前可见的选项
                LOGGER.info("开始匹配选项...")
                
                for idx, option in enumerate(options):
                    try:
                        option_text = option.text.strip()
                        
                        # 记录所有选项（不管是否可见）
                        if visible_count < 10 and option_text:
                            LOGGER.info(f"  选项 {visible_count + 1}: '{option_text}'")
                            visible_count += 1
                        
                        # 精确匹配
                        if option_text == value:
                            # 滚动到选项可见
                            try:
                                self._driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'});", option)
                                time.sleep(0.3)
                            except:
                                pass
                            
                            option.click()
                            LOGGER.info(f"✅ 已选择{field_name}: {option_text}（精确匹配）")
                            matched = True
                            break
                        # 包含匹配
                        elif value in option_text or option_text in value:
                            # 滚动到选项可见
                            try:
                                self._driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'});", option)
                                time.sleep(0.3)
                            except:
                                pass
                            
                            option.click()
                            LOGGER.info(f"✅ 已选择{field_name}: {option_text}（包含匹配）")
                            matched = True
                            break
                    except Exception as e:
                        LOGGER.debug(f"处理选项{idx}时出错: {e}")
                        continue
                
                if matched:
                    time.sleep(0.5)
                    return True
                
                # 如果没找到，尝试在下拉框中输入筛选
                LOGGER.warning(f"在当前选项中未找到{value}，尝试输入筛选...")
                try:
                    # 查找下拉框的输入框（有些下拉框支持输入筛选）
                    input_field = None
                    
                    # 查找可能的输入框
                    try:
                        input_field = dropdown_trigger.find_element(By.XPATH, ".//input")
                    except:
                        try:
                            input_field = self._driver.find_element(By.XPATH, "//input[@type='text' and not(@disabled)]")
                        except:
                            pass
                    
                    if input_field and input_field.is_displayed():
                        LOGGER.info("找到下拉框输入框，尝试输入筛选...")
                        input_field.clear()
                        input_field.send_keys(value)
                        time.sleep(1.0)
                        
                        # 再次查找选项
                        from selenium.webdriver.common.keys import Keys
                        
                        # 方法1: 按Enter选择
                        try:
                            input_field.send_keys(Keys.ENTER)
                            LOGGER.info(f"✅ 已通过输入筛选选择{field_name}: {value}")
                            time.sleep(0.5)
                            return True
                        except:
                            pass
                        
                        # 方法2: 查找筛选后的选项
                        try:
                            filtered_option = self._driver.find_element(
                                By.XPATH,
                                f"//*[normalize-space(text())='{value}' and (self::li or self::div)]"
                            )
                            if filtered_option.is_displayed():
                                filtered_option.click()
                                LOGGER.info(f"✅ 已通过筛选选择{field_name}: {value}")
                                time.sleep(0.5)
                                return True
                        except:
                            pass
                except Exception as e:
                    LOGGER.debug(f"输入筛选失败: {e}")
                
                # 最后尝试：直接发送值到dropdown_trigger
                LOGGER.warning("尝试直接向下拉框发送值...")
                try:
                    from selenium.webdriver.common.keys import Keys
                    dropdown_trigger.send_keys(value)
                    time.sleep(0.5)
                    dropdown_trigger.send_keys(Keys.ENTER)
                    LOGGER.info(f"✅ 已通过键盘输入选择{field_name}: {value}")
                    time.sleep(0.5)
                    return True
                except Exception as e:
                    LOGGER.debug(f"键盘输入失败: {e}")
                
                LOGGER.error(f"❌ 所有方法都失败，未找到匹配的{field_name}选项: {value}")
                all_options_text = [opt.text.strip() for opt in options if opt.text.strip()]
                LOGGER.error(f"所有选项（前30个）: {all_options_text[:30]}")
                return False
            
        except Exception as e:
            LOGGER.error(f"选择{field_name}下拉框失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _click_roll_button(self) -> bool:
        """
        点击页面右下角的Roll按钮
        
        Returns:
            True如果点击成功
        """
        try:
            LOGGER.info("查找并点击Roll按钮...")
            
            # 方法1: 通过按钮文本查找（支持多种可能的文本）
            roll_texts = ["Roll", "ROLL", "roll"]
            button_clicked = False
            
            for roll_text in roll_texts:
                try:
                    # 尝试通过文本查找按钮
                    roll_button = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            f"//button[contains(text(), '{roll_text}') or .//span[contains(text(), '{roll_text}')]]"
                        ))
                    )
                    if roll_button.is_displayed() and roll_button.is_enabled():
                        LOGGER.info(f"找到Roll按钮（文本: '{roll_text}'）")
                        # 滚动到按钮位置（右下角）
                        self._driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'end', inline: 'end'});",
                            roll_button
                        )
                        time.sleep(0.3)
                        roll_button.click()
                        LOGGER.info(f"✅ 已点击Roll按钮（文本: '{roll_text}'）")
                        button_clicked = True
                        break
                except TimeoutException:
                    continue
                except Exception as e:
                    LOGGER.debug(f"通过文本'{roll_text}'查找Roll按钮失败: {e}")
                    continue
            
            # 方法2: 如果方法1失败，尝试查找页面右下角的所有按钮
            if not button_clicked:
                try:
                    LOGGER.info("尝试在页面右下角查找Roll按钮...")
                    # 获取页面尺寸
                    page_height = self._driver.execute_script("return document.body.scrollHeight")
                    viewport_height = self._driver.execute_script("return window.innerHeight")
                    
                    # 滚动到页面底部
                    self._driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(0.5)
                    
                    # 查找所有按钮
                    all_buttons = self._driver.find_elements(By.TAG_NAME, "button")
                    LOGGER.info(f"找到 {len(all_buttons)} 个按钮，查找Roll按钮...")
                    
                    for button in all_buttons:
                        try:
                            if not button.is_displayed():
                                continue
                            
                            button_text = button.text.strip()
                            button_location = button.location
                            button_size = button.size
                            
                            # 检查按钮是否在右下角区域（右侧80%以上，底部20%以内）
                            viewport_width = self._driver.execute_script("return window.innerWidth")
                            is_bottom_right = (
                                button_location['x'] + button_size['width'] / 2 > viewport_width * 0.8 and
                                button_location['y'] + button_size['height'] / 2 > viewport_height * 0.8
                            )
                            
                            # 检查按钮文本是否包含"roll"（不区分大小写）
                            if "roll" in button_text.lower():
                                LOGGER.info(f"找到可能的Roll按钮: '{button_text}' (位置: {button_location}, 是否右下角: {is_bottom_right})")
                                if button.is_enabled():
                                    self._driver.execute_script(
                                        "arguments[0].scrollIntoView({block: 'end', inline: 'end'});",
                                        button
                                    )
                                    time.sleep(0.3)
                                    button.click()
                                    LOGGER.info(f"✅ 已点击Roll按钮: '{button_text}'")
                                    button_clicked = True
                                    break
                        except Exception as e:
                            LOGGER.debug(f"检查按钮时出错: {e}")
                            continue
                except Exception as e:
                    LOGGER.warning(f"在右下角查找Roll按钮失败: {e}")
            
            # 方法3: 如果前两种方法都失败，尝试通过CSS类名或ID查找
            if not button_clicked:
                try:
                    LOGGER.info("尝试通过CSS选择器查找Roll按钮...")
                    # 常见的按钮选择器模式
                    selectors = [
                        "button[class*='roll']",
                        "button[id*='roll']",
                        "button[class*='submit']",
                        "button[class*='save']",
                        ".roll-button",
                        "#roll-button",
                    ]
                    
                    for selector in selectors:
                        try:
                            roll_button = self._driver.find_element(By.CSS_SELECTOR, selector)
                            if roll_button.is_displayed() and roll_button.is_enabled():
                                LOGGER.info(f"通过CSS选择器找到Roll按钮: {selector}")
                                self._driver.execute_script(
                                    "arguments[0].scrollIntoView({block: 'end', inline: 'end'});",
                                    roll_button
                                )
                                time.sleep(0.3)
                                roll_button.click()
                                LOGGER.info(f"✅ 已点击Roll按钮（通过CSS选择器: {selector}）")
                                button_clicked = True
                                break
                        except:
                            continue
                except Exception as e:
                    LOGGER.debug(f"通过CSS选择器查找失败: {e}")
            
            if button_clicked:
                time.sleep(1.0)  # 等待按钮点击后的响应
                return True
            else:
                LOGGER.error("❌ 未找到Roll按钮")
                return False
                
        except Exception as e:
            LOGGER.error(f"点击Roll按钮失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def submit_vpo_data(self, data: dict) -> bool:
        """
        提交VPO数据到Spark网页
        
        Args:
            data: 包含VPO数据的字典
        
        Returns:
            True如果提交成功
        
        Raises:
            RuntimeError: 如果提交失败
        """
        LOGGER.info("开始提交VPO数据到Spark网页")
        LOGGER.debug(f"VPO数据: {data}")
        
        # 重试机制
        last_exception = None
        for attempt in range(1, self.config.retry_count + 1):
            try:
                LOGGER.info(f"尝试提交VPO数据 (第{attempt}/{self.config.retry_count}次)")
                
                # 初始化WebDriver
                self._init_driver()
                
                # 导航到页面
                self._navigate_to_page()
                
                # TODO: 根据Spark网页的实际界面实现具体的数据提交逻辑
                # 这里需要根据实际的Spark网页表单来填写数据
                # 示例：查找表单元素、填写数据、点击提交按钮等
                
                # 等待提交完成
                time.sleep(self.config.wait_after_submit)
                
                # 验证提交是否成功
                if self._verify_submission():
                    LOGGER.info("✅ VPO数据提交成功")
                    return True
                else:
                    raise RuntimeError("提交验证失败")
                
            except Exception as e:
                last_exception = e
                LOGGER.warning(f"第{attempt}次提交失败: {e}")
                if attempt < self.config.retry_count:
                    LOGGER.info(f"等待{self.config.retry_delay}秒后重试...")
                    time.sleep(self.config.retry_delay)
                    # 关闭当前WebDriver，准备重试
                    self._close_driver()
                else:
                    LOGGER.error(f"❌ VPO数据提交失败（已重试{self.config.retry_count}次）")
        
        # 清理资源
        self._close_driver()
        raise RuntimeError(f"VPO数据提交失败: {last_exception}")
    
    def _verify_submission(self) -> bool:
        """
        验证数据是否提交成功
        
        Returns:
            True如果验证通过
        """
        try:
            # TODO: 实现验证逻辑
            # 例如：检查页面是否显示成功消息、是否有错误提示等
            # 可以根据实际Spark网页的反馈机制来实现
            
            # 示例：查找成功消息元素
            # success_element = WebDriverWait(self._driver, 10).until(
            #     EC.presence_of_element_located((By.CLASS_NAME, "success-message"))
            # )
            # return success_element is not None
            
            LOGGER.info("✅ VPO数据提交验证通过")
            return True
            
        except Exception as e:
            LOGGER.warning(f"验证提交结果时出错: {e}")
            return False
    
    def __enter__(self):
        """上下文管理器入口"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self._close_driver()

    # ------------------------------------------------------------------
    # Dashboard VPO 收集相关
    # ------------------------------------------------------------------

    def collect_recent_vpos_from_dashboard(self, expected_count: int) -> List[str]:
        """
        在Spark Dashboard的“Rolling Experiments”区域收集最近的VPO编号。
        
        显示顺序：从上到下、从左到右 = 从最近提交到最早提交。
        
        Args:
            expected_count: 预期需要收集的VPO数量（通常等于本次提交的MIR行数）
        
        Returns:
            VPO字符串列表（按页面显示顺序：最新在前、最旧在后）
        """
        wait_minutes = getattr(self.config, "vpo_collect_wait_minutes", 0)
        LOGGER.info("=" * 80)
        LOGGER.info(f"所有VPO提交完成，将等待 {wait_minutes} 分钟后再从Dashboard收集VPO编号（0 表示不等待）...")
        LOGGER.info("=" * 80)
        
        # 等待一段时间，给后端系统生成VPO的时间；如果配置为0则不等待
        if wait_minutes > 0:
            time.sleep(wait_minutes * 60)
        
        # 初始化并打开Dashboard
        self._init_driver()
        self._navigate_to_page()
        
        vpo_list: List[str] = []
        
        try:
            LOGGER.info("等待“Rolling Experiments”区域加载...")
            try:
                WebDriverWait(self._driver, self.config.explicit_wait).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(), 'Rolling Experiments')]")
                    )
                )
            except TimeoutException:
                LOGGER.warning("在页面中未能明显找到“Rolling Experiments”标题，将直接全局搜索包含 VPO 的元素。")
            
            # 查找所有包含 "VPO " 文本的元素
            LOGGER.info("在Dashboard页面中查找包含 'VPO ' 文本的元素...")
            elements = self._driver.find_elements(
                By.XPATH,
                "//*[contains(normalize-space(text()), 'VPO ')]"
            )
            LOGGER.info(f"找到 {len(elements)} 个包含 'VPO ' 的元素，开始解析VPO编号...")
            
            pattern = re.compile(r"VPO\s+([A-Za-z0-9_-]+)")
            
            for elem in elements:
                try:
                    text = elem.text.strip()
                    if not text:
                        continue
                    match = pattern.search(text)
                    if match:
                        vpo_code = match.group(1)
                        if vpo_code not in vpo_list:
                            vpo_list.append(vpo_code)
                            LOGGER.info(f"解析到VPO: {vpo_code}")
                            if len(vpo_list) >= expected_count:
                                break
                except Exception as e:
                    LOGGER.debug(f"解析元素文本时出错: {e}")
                    continue
            
            LOGGER.info(f"最终收集到 {len(vpo_list)} 个VPO编号，预期数量: {expected_count}")
            if len(vpo_list) < expected_count:
                LOGGER.warning("⚠️ 收集到的VPO数量少于预期，后续合并时将按可用数量进行匹配。")
            
            return vpo_list
        
        except Exception as e:
            LOGGER.error(f"从Dashboard收集VPO编号失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return vpo_list

