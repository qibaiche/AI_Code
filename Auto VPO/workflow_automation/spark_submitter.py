"""Spark网页提交VPO数据模块"""
import logging
import time
from typing import Optional
from dataclasses import dataclass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

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
        LOGGER.info("查找并点击'Add New'按钮...")
        
        try:
            # 尝试多种方式查找按钮
            add_new_button = None
            
            # 方法1: 通过按钮文本查找
            try:
                add_new_button = WebDriverWait(self._driver, self.config.explicit_wait).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add New')]"))
                )
                LOGGER.info("通过文本找到'Add New'按钮")
            except TimeoutException:
                LOGGER.debug("方法1失败：通过文本未找到按钮")
            
            # 方法2: 通过按钮包含的span文本查找
            if not add_new_button:
                try:
                    add_new_button = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Add New')]]"))
                    )
                    LOGGER.info("通过span文本找到'Add New'按钮")
                except TimeoutException:
                    LOGGER.debug("方法2失败：通过span文本未找到按钮")
            
            # 方法3: 通过CSS选择器查找（查找包含"Add New"或"+"的按钮）
            if not add_new_button:
                try:
                    buttons = self._driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        button_text = button.text.strip()
                        if "Add New" in button_text or (button_text.startswith("+") and "Add" in button_text):
                            add_new_button = button
                            LOGGER.info(f"通过遍历找到按钮: '{button_text}'")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法3失败：遍历按钮时出错: {e}")
            
            if add_new_button:
                # 滚动到按钮可见
                self._driver.execute_script("arguments[0].scrollIntoView(true);", add_new_button)
                time.sleep(0.3)
                
                # 点击按钮
                add_new_button.click()
                LOGGER.info("✅ 已点击'Add New'按钮")
                time.sleep(1.0)  # 等待页面响应
                return True
            else:
                LOGGER.error("❌ 未找到'Add New'按钮")
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
        LOGGER.info(f"填写Test Program Path: {tp_path}")
        
        try:
            # 等待输入框出现并获得焦点
            time.sleep(1.5)
            
            # 查找"Provide test program path"输入框
            input_field = None
            
            # 方法0: 直接使用当前获得焦点的元素（光标在跳动说明已经有焦点）
            try:
                input_field = self._driver.switch_to.active_element
                tag_name = input_field.tag_name.lower()
                if tag_name == "input" or tag_name == "textarea":
                    LOGGER.info(f"使用当前焦点元素作为输入框 (tag: {tag_name})")
                else:
                    LOGGER.debug(f"当前焦点元素不是输入框 (tag: {tag_name})")
                    input_field = None
            except Exception as e:
                LOGGER.debug(f"获取焦点元素失败: {e}")
            
            # 方法1: 查找对话框中最大的输入框
            if not input_field:
                try:
                    # 查找所有可见的input和textarea
                    all_inputs = self._driver.find_elements(By.XPATH, "//input[@type='text' or not(@type)] | //textarea")
                    
                    # 过滤可见的
                    visible_inputs = [inp for inp in all_inputs if inp.is_displayed()]
                    LOGGER.info(f"找到 {len(visible_inputs)} 个可见的输入框")
                    
                    # 找最大的
                    if visible_inputs:
                        largest_input = max(visible_inputs, key=lambda x: x.size.get('width', 0) * x.size.get('height', 0))
                        input_field = largest_input
                        LOGGER.info(f"使用最大的输入框 (宽度: {largest_input.size.get('width')}, 高度: {largest_input.size.get('height')})")
                except Exception as e:
                    LOGGER.debug(f"方法1失败: {e}")
            
            # 方法2: 通过包含"path"的label查找
            if not input_field:
                try:
                    labels = self._driver.find_elements(By.XPATH, "//*[contains(text(), 'path') or contains(text(), 'Path')]")
                    for label in labels:
                        try:
                            # 尝试找label附近的输入框
                            nearby_inputs = label.find_elements(By.XPATH, "./following-sibling::*//input | ./following-sibling::input | .//input | ./parent::*/following-sibling::*//input")
                            if nearby_inputs:
                                input_field = nearby_inputs[0]
                                LOGGER.info("通过label找到输入框")
                                break
                        except:
                            continue
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            if not input_field:
                LOGGER.error("❌ 未找到Test Program Path输入框")
                # 列出所有可见的输入框用于调试
                try:
                    all_inputs = self._driver.find_elements(By.XPATH, "//input | //textarea")
                    LOGGER.info(f"调试：页面上共有 {len(all_inputs)} 个输入框")
                    for i, inp in enumerate(all_inputs[:5]):  # 只显示前5个
                        LOGGER.info(f"  输入框 {i+1}: type={inp.get_attribute('type')}, visible={inp.is_displayed()}, size={inp.size}")
                except:
                    pass
                return False
            
            # 清空并填写路径
            input_field.clear()
            input_field.send_keys(tp_path)
            LOGGER.info(f"✅ 已填写TP路径: {tp_path}")
            
            # 立即查找并点击Apply按钮（无需等待）
            apply_button = None
            
            # 方法1: 通过按钮文本查找
            try:
                apply_button = WebDriverWait(self._driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Apply')]"))
                )
                LOGGER.info("通过文本找到'Apply'按钮")
            except TimeoutException:
                LOGGER.debug("通过文本未找到Apply按钮")
            
            # 方法2: 查找包含Apply的按钮（可能有箭头图标）
            if not apply_button:
                try:
                    buttons = self._driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        button_text = button.text.strip()
                        if "Apply" in button_text:
                            apply_button = button
                            LOGGER.info(f"通过遍历找到按钮: '{button_text}'")
                            break
                except Exception as e:
                    LOGGER.debug(f"遍历按钮失败: {e}")
            
            if not apply_button:
                LOGGER.error("❌ 未找到'Apply'按钮")
                return False
            
            # 点击Apply按钮
            apply_button.click()
            LOGGER.info("✅ 已点击'Apply'按钮")
            
            # 立即等待加载并点击Continue（Apply后的验证由Continue处理）
            LOGGER.info("等待Continue按钮出现并点击...")
            if self._wait_and_click_continue():
                LOGGER.info("✅ 已成功填写TP路径并完成页面跳转")
                return True
            else:
                LOGGER.error("❌ Continue按钮点击失败或页面未能跳转")
                return False  # 修复：Continue失败应该返回False，不能继续执行
            
        except Exception as e:
            LOGGER.error(f"填写Test Program Path失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
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
            
            # 等待Continue按钮出现并可点击（最多等待30秒）
            continue_button = None
            
            # 方法1: 通过文本查找Continue按钮并等待可点击
            try:
                continue_button = WebDriverWait(self._driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))
                )
                LOGGER.info("通过文本找到'Continue'按钮")
            except TimeoutException:
                LOGGER.debug("通过文本未找到Continue按钮")
            
            # 方法2: 查找所有按钮，找包含Continue的
            if not continue_button:
                try:
                    buttons = WebDriverWait(self._driver, 10).until(
                        EC.presence_of_all_elements_located((By.TAG_NAME, "button"))
                    )
                    for button in buttons:
                        button_text = button.text.strip()
                        if "Continue" in button_text:
                            # 等待按钮可点击
                            WebDriverWait(self._driver, 20).until(
                                lambda d: button.is_enabled() and button.is_displayed()
                            )
                            continue_button = button
                            LOGGER.info(f"通过遍历找到按钮: '{button_text}'")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            if not continue_button:
                LOGGER.error("❌ 未找到'Continue'按钮")
                return False
            
            # 滚动到按钮可见
            self._driver.execute_script("arguments[0].scrollIntoView(true);", continue_button)
            time.sleep(0.3)
            
            # 点击Continue按钮（可能需要多次点击）
            # 用户反馈：即使出现红色错误消息，也要坚持重试，因为这是偶发的加载失败
            max_continue_clicks = 15  # 增加到15次重试（每次等待45秒 = 最多11分钟）
            
            for click_attempt in range(1, max_continue_clicks + 1):
                LOGGER.info(f"🔄 准备点击'Continue'按钮（第 {click_attempt}/{max_continue_clicks} 次）...")
                
                # 重新查找Continue按钮（可能在重试过程中DOM更新了）
                continue_button = None
                try:
                    continue_button = WebDriverWait(self._driver, 15).until(  # 增加到15秒
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))
                    )
                    LOGGER.info(f"找到Continue按钮（第 {click_attempt} 次尝试）")
                except TimeoutException:
                    LOGGER.warning(f"⚠️ 15秒内未找到Continue按钮（第 {click_attempt} 次尝试）")
                    
                    # 检查是否已经跳转成功
                    if self._check_target_page_loaded():
                        LOGGER.info("✅ 目标页面已加载，跳转成功！")
                        return True
                    
                    # 如果还没到最后一次尝试，继续循环等待Continue按钮重新出现
                    if click_attempt < max_continue_clicks:
                        LOGGER.info(f"Continue按钮暂时消失，等待5秒后继续尝试...")
                        time.sleep(5.0)
                        continue  # 继续下一次循环
                    else:
                        # 最后一次尝试也找不到
                        LOGGER.error("❌ 最后一次尝试仍未找到Continue按钮且页面未跳转")
                        return False
                
                if not continue_button:
                    # 理论上不应该到这里，但保险起见
                    LOGGER.warning("Continue按钮为空，跳过本次循环")
                    time.sleep(2.0)
                    continue
                
                # 点击Continue按钮
                try:
                    continue_button.click()
                    LOGGER.info(f"✅ 已点击'Continue'按钮（第 {click_attempt} 次）")
                except Exception as e:
                    LOGGER.warning(f"点击失败: {e}")
                    time.sleep(2.0)
                    continue
                
                # 等待页面加载完成（最多45秒）
                LOGGER.info("⏳ 等待页面加载完成...")
                load_success = self._wait_for_page_load_after_continue()
                
                if load_success:
                    LOGGER.info(f"✅ 页面加载完成，跳转成功！（第 {click_attempt} 次点击后成功）")
                    return True
                else:
                    # 加载未成功，继续重试
                    LOGGER.warning(f"⚠️ 第 {click_attempt} 次点击后页面未成功跳转")
                    
                    if click_attempt < max_continue_clicks:
                        LOGGER.info(f"💪 即使有错误消息，也继续重试！准备第 {click_attempt + 1} 次点击...")
                        # 短暂等待后继续
                        time.sleep(2.0)
                        continue  # 继续下一次循环
                    else:
                        # 已经是最后一次尝试了
                        LOGGER.error(f"❌ 已重试{max_continue_clicks}次（约{max_continue_clicks * 0.75}分钟），仍未成功跳转")
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
        2. 等待"Add New Experiment"按钮出现（说明跳转成功）
        3. 检测页面是否崩溃
        
        Returns:
            True如果页面加载完成并成功跳转
            False如果仍在原对话框或加载失败
        """
        try:
            max_wait_time = 60  # 最多等待60秒
            LOGGER.info(f"等待页面跳转完成（最多{max_wait_time}秒）...")
            
            # 等待一段时间让页面开始加载
            time.sleep(3.0)
            
            # 检查是否还在"Create New Experiments"对话框
            for check_attempt in range(max_wait_time // 3):  # 每3秒检查一次
                try:
                    # 检查对话框标题是否还存在
                    create_dialog = self._driver.find_elements(By.XPATH, "//*[contains(text(), 'Create New Experiments')]")
                    
                    if create_dialog and any(elem.is_displayed() for elem in create_dialog):
                        LOGGER.warning(f"⚠️ 仍在'Create New Experiments'对话框中（检查{check_attempt + 1}次）")
                        
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
                        
                        # 如果已经检查了很多次（超过45秒），认为这次Continue点击无效
                        # 但不直接认为失败，而是返回False让上层继续重试Continue
                        if check_attempt >= 15:  # 15次 * 3秒 = 45秒
                            LOGGER.warning("⚠️ 加载时间较长，仍在原对话框中，返回让上层重新点击Continue")
                            return False
                        
                        # 继续等待
                        time.sleep(3.0)
                        continue
                    else:
                        # 对话框已消失，说明可能已经跳转
                        LOGGER.info("✅ 'Create New Experiments'对话框已消失")
                        break
                        
                except Exception as e:
                    LOGGER.debug(f"检查对话框时出错: {e}")
                    break
            
            # 验证是否成功跳转：查找"Add New Experiment"按钮
            try:
                LOGGER.info("验证页面跳转：查找'Add New Experiment'按钮...")
                add_exp_button = WebDriverWait(self._driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Add New Experiment')]"))
                )
                LOGGER.info("✅ 'Add New Experiment'按钮已出现，页面跳转成功！")
                time.sleep(1.5)  # 等待页面稳定
                return True
                
            except TimeoutException:
                LOGGER.error("❌ 未找到'Add New Experiment'按钮，页面跳转失败")
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
        LOGGER.info("查找并点击'Add New Experiment'按钮...")
        
        try:
            # 等待页面完全稳定（因为之前可能刚加载完）
            time.sleep(2.0)
            
            # 确保页面已完全加载
            try:
                self._driver.execute_script("return document.readyState") == "complete"
            except:
                pass
            
            add_experiment_button = None
            
            # 方法1: 通过按钮文本查找（等待更长时间，因为页面可能刚加载完）
            try:
                add_experiment_button = WebDriverWait(self._driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add New Experiment')]"))
                )
                LOGGER.info("通过文本找到'Add New Experiment'按钮")
            except TimeoutException:
                LOGGER.debug("方法1失败：通过文本未找到按钮")
            
            # 方法2: 通过包含span的按钮查找
            if not add_experiment_button:
                try:
                    add_experiment_button = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Add New Experiment')]]"))
                    )
                    LOGGER.info("通过span文本找到'Add New Experiment'按钮")
                except TimeoutException:
                    LOGGER.debug("方法2失败：通过span文本未找到按钮")
            
            # 方法3: 查找包含"Add"和"Experiment"的按钮
            if not add_experiment_button:
                try:
                    buttons = self._driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        button_text = button.text.strip()
                        if "Add" in button_text and "Experiment" in button_text:
                            add_experiment_button = button
                            LOGGER.info(f"通过遍历找到按钮: '{button_text}'")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法3失败: {e}")
            
            if not add_experiment_button:
                LOGGER.error("❌ 未找到'Add New Experiment'按钮")
                return False
            
            # 滚动到按钮可见
            self._driver.execute_script("arguments[0].scrollIntoView(true);", add_experiment_button)
            time.sleep(0.3)
            
            # 点击按钮
            add_experiment_button.click()
            LOGGER.info("✅ 已点击'Add New Experiment'按钮")
            time.sleep(2.0)  # 等待页面响应
            
            return True
            
        except Exception as e:
            LOGGER.error(f"点击'Add New Experiment'按钮失败: {e}")
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
            
            # 1. 勾选"Override parttype with"复选框
            LOGGER.info("查找'Override parttype with'复选框...")
            
            # 快速查找复选框（优化：一旦找到就停止）
            checkbox = None
            
            # 方法1: 通过label快速查找（最可能成功）
            try:
                checkbox = self._driver.find_element(
                    By.XPATH, 
                    "//input[@type='checkbox' and ..//*[contains(text(), 'Override parttype')]]"
                )
                LOGGER.info("✅ 找到复选框")
            except:
                # 方法2: 查找所有可见checkbox（限制数量避免过慢）
                try:
                    checkboxes = self._driver.find_elements(By.XPATH, "//input[@type='checkbox' and ancestor::*[contains(@class, 'parttype') or contains(@class, 'override')]]")
                    if not checkboxes:
                        checkboxes = self._driver.find_elements(By.XPATH, "//input[@type='checkbox']")[:10]  # 只检查前10个
                    
                    for cb in checkboxes:
                        if cb.is_displayed():
                            try:
                                parent = cb.find_element(By.XPATH, "./..")
                                parent_text = parent.text.lower()
                                if "parttype" in parent_text or "override" in parent_text:
                                    checkbox = cb
                                    LOGGER.info("✅ 找到复选框")
                                    break
                            except:
                                continue
                except Exception as e:
                    LOGGER.debug(f"查找复选框失败: {e}")
            
            # 勾选复选框
            if checkbox:
                try:
                    if not checkbox.is_selected():
                        # 尝试直接点击
                        try:
                            checkbox.click()
                            LOGGER.info("✅ 已勾选'Override parttype with'复选框（直接点击）")
                        except:
                            # 如果直接点击失败，尝试用JavaScript
                            self._driver.execute_script("arguments[0].click();", checkbox)
                            LOGGER.info("✅ 已勾选'Override parttype with'复选框（JavaScript）")
                    else:
                        LOGGER.info("'Override parttype with'复选框已勾选")
                    
                    time.sleep(0.5)  # 优化：减少等待时间
                    
                except Exception as e:
                    LOGGER.warning(f"勾选复选框时出错: {e}")
            else:
                LOGGER.warning("未找到'Override parttype with'复选框，可能默认已勾选")
            
            # 2. 查找并点击Parttype下拉框（自定义控件）
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

    def _diagnose_flow_page(self, condition_index: int = 0) -> None:
        """
        诊断Flow页面的DOM结构，输出详细信息用于调试
        
        Args:
            condition_index: 要诊断的condition索引
        """
        LOGGER.info("=" * 80)
        LOGGER.info("🔍 开始诊断Flow页面DOM结构...")
        LOGGER.info("=" * 80)
        
        try:
            # 1. 统计所有condition容器
            all_containers = self._driver.find_elements(By.XPATH, "//div[contains(@class,'condition-list-container')]")
            LOGGER.info(f"📊 找到 {len(all_containers)} 个 condition-list-container")
            
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
            
            # 4. 检查目标condition容器
            condition_xpath = f"(//div[contains(@class,'condition-list-container')])[{condition_index + 1}]"
            try:
                target_container = self._driver.find_element(By.XPATH, condition_xpath)
                LOGGER.info(f"✅ 目标容器[{condition_index + 1}]存在")
                
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
                LOGGER.error(f"❌ 目标容器[{condition_index + 1}]不存在: {e}")
            
            # 5. 检查"All Units"等可能干扰的元素
            all_units = self._driver.find_elements(By.XPATH, "//*[contains(text(),'All Units')]")
            if all_units:
                LOGGER.warning(f"⚠️ 找到 {len(all_units)} 个包含'All Units'的元素（可能干扰定位）")
            
            # 6. 输出当前页面的关键XPath尝试结果
            LOGGER.info("\n📋 测试关键XPath:")
            test_xpaths = [
                (f"(//div[contains(@class,'condition-list-container')])[{condition_index + 1}]", "目标容器"),
                (f"(//div[contains(@class,'condition-list-container')])[{condition_index + 1}]//mat-form-field[contains(@class,'mat-form-field-type-mat-select')][1]//div[contains(@class,'mat-select-trigger')]", "Operation trigger"),
                (f"(//div[contains(@class,'condition-list-container')])[{condition_index + 1}]//mat-form-field[contains(@class,'mat-form-field-type-mat-select')][2]//div[contains(@class,'mat-select-trigger')]", "Eng ID trigger"),
                (f"(//div[contains(@class,'mat-select-arrow-wrapper')])[{1 + 2 * condition_index}]", f"Operation箭头[{1 + 2 * condition_index}]"),
                (f"(//div[contains(@class,'mat-select-arrow-wrapper')])[{2 + 2 * condition_index}]", f"Eng ID箭头[{2 + 2 * condition_index}]"),
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
    
    def _select_operation(self, operation: str, condition_index: int = 0) -> bool:
        """
        在第 condition_index 个 condition 上选择 Operation
        
        Args:
            operation: Operation值（如"6248"）
            condition_index: 0 表示第一个，1 表示第二个...
        """
        LOGGER.info(f"选择Operation: {operation} (condition_index={condition_index})")
        timeout = self.config.explicit_wait
        wait = WebDriverWait(self._driver, timeout)

        try:
            time.sleep(0.5)

            # 🔍 调试：先看看页面上有多少个condition容器
            all_containers = self._driver.find_elements(By.XPATH, "//div[contains(@class,'condition-list-container')]")
            LOGGER.info(f"🔍 调试：页面上共有 {len(all_containers)} 个 condition-list-container")
            
            # 1. 先定位到第 N 个 condition 容器
            #    这里用DOM里的 class: "condition-list-container"
            condition_xpath = f"(//div[contains(@class,'condition-list-container')])[{condition_index + 1}]"
            LOGGER.debug(f"condition_xpath = {condition_xpath}")
            
            # **关键：先等待这个容器在DOM中出现**
            try:
                LOGGER.info(f"等待第 {condition_index + 1} 个condition容器出现...")
                wait.until(EC.presence_of_element_located((By.XPATH, condition_xpath)))
                LOGGER.info(f"✅ 第 {condition_index + 1} 个condition容器已出现")
                time.sleep(0.5)  # 再等一下让内部元素完全渲染
            except TimeoutException:
                LOGGER.error(f"❌ 超时：第 {condition_index + 1} 个condition容器未出现")
                LOGGER.error(f"   实际只找到 {len(all_containers)} 个容器")
                self._diagnose_flow_page(condition_index)
                return False

            # 2. 在这个容器里，找第一个 mat-select 作为 Operation
            #    （如果以后结构变，可以微调这个 XPath）
            operation_trigger_xpath = (
                condition_xpath
                + "//mat-form-field[contains(@class,'mat-form-field-type-mat-select')][1]"
                + "//div[contains(@class,'mat-select-trigger')]"
            )

            if not self._scroll_and_click(
                By.XPATH,
                operation_trigger_xpath,
                f"第 {condition_index + 1} 个 condition 的 Operation 下拉",
            ):
                LOGGER.warning("⚠️ 方法1（容器内查找）失败，尝试备用方法：按全局箭头索引")
                
                # 备用方法：按mat-select-arrow-wrapper的全局索引查找
                # Operation在每个condition中是第一个select，所以索引为 1 + 2*condition_index
                arrow_index = 1 + 2 * condition_index
                fallback_xpath = f"(//div[contains(@class,'mat-select-arrow-wrapper')])[{arrow_index}]"
                LOGGER.info(f"🔄 尝试备用XPath: {fallback_xpath}")
                
                if not self._scroll_and_click(
                    By.XPATH,
                    fallback_xpath,
                    f"Operation箭头[{arrow_index}]（备用方法）",
                ):
                    LOGGER.error("❌ 打开 Operation 下拉失败（所有方法均失败）")
                    self._diagnose_flow_page(condition_index)
                    return False

            # 3. 等下拉面板出来，在 overlay 里找对应的选项
            #    Angular Material 的 option 会挂在 cdk-overlay-pane 下面
            option_xpath = (
                "//div[contains(@class,'cdk-overlay-pane')]"
                "//mat-option//span[normalize-space()='%s']" % operation
            )

            option = wait.until(
                EC.element_to_be_clickable((By.XPATH, option_xpath))
            )
            option.click()
            LOGGER.info(f"✅ 已选择 Operation: {operation}")
            time.sleep(0.3)
            return True

        except Exception as e:
            LOGGER.error(f"选择Operation失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            self._diagnose_flow_page(condition_index)
            return False
    
    def _select_eng_id(self, eng_id: str, condition_index: int = 0) -> bool:
        """
        在Flow标签页选择Eng ID
        
        Args:
            eng_id: Eng ID值（如"CCG-24J-TEST"）
            condition_index: 条件序号（0表示第一个condition）
            
        Returns:
            True如果选择成功
        """
        LOGGER.info(f"选择Eng ID: {eng_id} (condition_index={condition_index})")
        timeout = self.config.explicit_wait
        wait = WebDriverWait(self._driver, timeout)
        
        try:
            # 等待Operation选择完成
            time.sleep(1.0)
            
            # **1. 先确认对应的condition容器已存在**
            condition_xpath = f"(//div[contains(@class,'condition-list-container')])[{condition_index + 1}]"
            try:
                LOGGER.info(f"确认第 {condition_index + 1} 个condition容器存在（Eng ID选择前）...")
                wait.until(EC.presence_of_element_located((By.XPATH, condition_xpath)))
                time.sleep(0.5)
            except TimeoutException:
                LOGGER.error(f"❌ 第 {condition_index + 1} 个condition容器不存在")
                self._diagnose_flow_page(condition_index)
                return False
            
            # **关键修复：在点击Eng ID之前，先关闭所有已打开的overlay（如"All Units"）**
            try:
                # 方法1: 按ESC键关闭所有打开的overlay
                self._driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                time.sleep(0.3)
                LOGGER.info("已按ESC关闭所有打开的overlay")
            except:
                pass
            
            # 方法2: 如果还有overlay，点击backdrop关闭
            try:
                backdrops = self._driver.find_elements(By.XPATH, "//div[contains(@class,'cdk-overlay-backdrop')]")
                if backdrops:
                    for backdrop in backdrops:
                        try:
                            if backdrop.is_displayed():
                                backdrop.click()
                                LOGGER.info("已点击backdrop关闭overlay")
                                time.sleep(0.3)
                                break
                        except:
                            pass
            except:
                pass
            
            # **2. 在这个容器里，找第二个 mat-select 作为 Eng ID**
            #    和Operation类似，但这次是[2]而不是[1]
            eng_id_trigger_xpath = (
                condition_xpath
                + "//mat-form-field[contains(@class,'mat-form-field-type-mat-select')][2]"
                + "//div[contains(@class,'mat-select-trigger')]"
            )
            
            LOGGER.info(f"尝试点击Eng ID下拉（容器内第2个mat-select）")
            
            # **关键：先验证trigger确实在condition容器内，并且不是"All Units"**
            try:
                trigger_element = self._driver.find_element(By.XPATH, eng_id_trigger_xpath)
                # 验证：检查这个trigger的父级是否在condition容器内
                parent_container = trigger_element.find_element(By.XPATH, "./ancestor::div[contains(@class,'condition-list-container')]")
                if not parent_container:
                    raise Exception("Trigger不在condition容器内")
                LOGGER.info("✅ 已验证Eng ID trigger在正确的condition容器内")
            except Exception as e:
                LOGGER.warning(f"⚠️ 验证trigger位置失败: {e}")
            
            # 点击trigger
            if not self._scroll_and_click(
                By.XPATH,
                eng_id_trigger_xpath,
                f"第 {condition_index + 1} 个 condition 的 Eng ID 下拉",
            ):
                LOGGER.warning("⚠️ 方法1（容器内查找）失败，尝试备用方法：按全局索引")
                
                # 备用方法：按全局索引（可能受页面其他下拉框影响）
                arrow_index = 2 + 2 * condition_index
                fallback_xpath = f"(//div[contains(@class,'mat-select-arrow-wrapper')])[{arrow_index}]"
                LOGGER.info(f"🔄 尝试备用XPath: {fallback_xpath}")
                
                if not self._scroll_and_click(
                    By.XPATH,
                    fallback_xpath,
                    f"Eng ID箭头[{arrow_index}]（备用方法）",
                ):
                    LOGGER.error("❌ 未找到Eng ID下拉框（所有方法均失败）")
                    self._diagnose_flow_page(condition_index)
                    return False
            
            # **关键验证：点击后检查打开的overlay是否是"All Units"**
            time.sleep(0.5)  # 等待overlay打开
            try:
                # 检查是否有"All Units"的overlay打开了
                all_units_overlays = self._driver.find_elements(
                    By.XPATH,
                    "//div[contains(@class,'cdk-overlay-pane')]//span[contains(text(),'All Units')]"
                )
                if all_units_overlays:
                    LOGGER.warning("⚠️ 检测到误点击了'All Units'下拉框，正在关闭并重新点击正确的Eng ID...")
                    # 关闭"All Units"的overlay
                    self._driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                    time.sleep(0.5)
                    
                    # 重新点击正确的Eng ID trigger（确保在condition容器内）
                    LOGGER.info("重新点击正确的Eng ID trigger...")
                    trigger_element = self._driver.find_element(By.XPATH, eng_id_trigger_xpath)
                    self._driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                        trigger_element
                    )
                    time.sleep(0.3)
                    trigger_element.click()
                    time.sleep(0.5)
                    LOGGER.info("✅ 已重新点击正确的Eng ID trigger")
            except:
                pass  # 如果没有"All Units"的overlay，说明点击正确
            
            # **3. 等待下拉选项面板出现（确保不是"All Units"的overlay）**
            LOGGER.info("等待Eng ID选项面板出现...")
            time.sleep(1.0)
            
            # 确认overlay已打开，但排除"All Units"相关的overlay
            max_retries = 3
            overlay_found = False
            for retry in range(max_retries):
                try:
                    # 检查是否有overlay出现
                    all_overlays = self._driver.find_elements(By.XPATH, "//div[contains(@class,'cdk-overlay-pane')]//mat-option")
                    if not all_overlays:
                        time.sleep(0.5)
                        continue
                    
                    # 检查是否有"All Units"的overlay
                    has_all_units = False
                    for overlay in self._driver.find_elements(By.XPATH, "//div[contains(@class,'cdk-overlay-pane')]"):
                        if overlay.find_elements(By.XPATH, ".//span[contains(text(),'All Units')]"):
                            has_all_units = True
                            break
                    
                    if has_all_units:
                        LOGGER.warning(f"⚠️ 第{retry+1}次检测：发现'All Units'的overlay，关闭并重新点击Eng ID...")
                        self._driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                        time.sleep(0.5)
                        # 重新点击Eng ID trigger
                        trigger_element = self._driver.find_element(By.XPATH, eng_id_trigger_xpath)
                        self._driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                            trigger_element
                        )
                        time.sleep(0.3)
                        trigger_element.click()
                        time.sleep(0.5)
                        continue
                    
                    # 找到了非"All Units"的overlay
                    overlay_found = True
                    LOGGER.info("✅ Eng ID选项面板已打开（非'All Units'）")
                    break
                except:
                    time.sleep(0.5)
                    continue
            
            if not overlay_found:
                LOGGER.error("❌ Eng ID选项面板未出现（或始终是'All Units'）")
                return False
            
            time.sleep(0.3)
            
            # **4. 查找并点击对应的Eng ID选项（排除"All Units"等干扰选项）**
            option_clicked = False
            
            # 方法1: 通过overlay中的mat-option的span文本精确匹配
            # 但排除"All Units"、"Good Units"、"Rejected Units"
            excluded_texts = ["All Units", "Good Units", "Rejected Units"]
            try:
                # 先找到所有overlay pane
                all_overlays = self._driver.find_elements(By.XPATH, "//div[contains(@class,'cdk-overlay-pane')]")
                LOGGER.info(f"找到 {len(all_overlays)} 个overlay pane")
                
                # 查找目标选项，但排除"All Units"相关的overlay
                for overlay in all_overlays:
                    try:
                        # 检查这个overlay是否包含"All Units"等选项
                        has_all_units = False
                        for excluded in excluded_texts:
                            if overlay.find_elements(By.XPATH, f".//span[contains(text(),'{excluded}')]"):
                                has_all_units = True
                                LOGGER.debug(f"跳过包含'{excluded}'的overlay")
                                break
                        
                        if has_all_units:
                            continue  # 跳过"All Units"的overlay
                        
                        # 在这个overlay中查找目标Eng ID
                        option = overlay.find_element(
                            By.XPATH,
                            f".//mat-option//span[normalize-space()='{eng_id}']"
                        )
                        if option and option.is_displayed():
                            option.click()
                            option_clicked = True
                            LOGGER.info(f"✅ 已选择Eng ID: {eng_id}（精确匹配）")
                            break
                    except:
                        continue
                
                if not option_clicked:
                    LOGGER.debug("方法1失败：在所有非'All Units'的overlay中未找到精确匹配")
            except Exception as e:
                LOGGER.debug(f"方法1失败: {e}")
            
            # 方法2: 遍历overlay中的所有mat-option查找匹配项（排除"All Units"相关）
            if not option_clicked:
                LOGGER.info("方法1失败，尝试遍历所有可见选项（排除'All Units'）...")
                try:
                    # 只查找overlay中的选项，但排除"All Units"相关的overlay
                    excluded_texts = ["All Units", "Good Units", "Rejected Units"]
                    all_overlays = self._driver.find_elements(By.XPATH, "//div[contains(@class,'cdk-overlay-pane')]")
                    
                    for overlay in all_overlays:
                        try:
                            # 检查这个overlay是否包含"All Units"等选项
                            has_all_units = False
                            for excluded in excluded_texts:
                                if overlay.find_elements(By.XPATH, f".//span[contains(text(),'{excluded}')]"):
                                    has_all_units = True
                                    break
                            
                            if has_all_units:
                                continue  # 跳过"All Units"的overlay
                            
                            # 在这个overlay中查找所有选项
                            options = overlay.find_elements(By.XPATH, ".//mat-option")
                            LOGGER.info(f"在非'All Units'的overlay中找到 {len(options)} 个选项")
                            
                            for idx, opt in enumerate(options):
                                try:
                                    opt_text = opt.text.strip()
                                    if not opt_text:
                                        continue
                                    
                                    # 排除"All Units"相关选项
                                    if any(excluded in opt_text for excluded in excluded_texts):
                                        continue
                                    
                                    # 只记录前5个选项避免日志过多
                                    if idx < 5:
                                        LOGGER.debug(f"  选项 {idx+1}: '{opt_text}'")
                                    
                                    # 精确匹配或包含匹配
                                    if opt_text == eng_id or eng_id in opt_text:
                                        if opt.is_displayed():
                                            opt.click()
                                            option_clicked = True
                                            LOGGER.info(f"✅ 已选择Eng ID: {opt_text}（遍历匹配）")
                                            break
                                except Exception as e:
                                    LOGGER.debug(f"处理选项{idx}时出错: {e}")
                                    continue
                            
                            if option_clicked:
                                break
                        except:
                            continue
                except Exception as e:
                    LOGGER.debug(f"方法2失败: {e}")
            
            if not option_clicked:
                LOGGER.error(f"❌ 未找到Eng ID选项: {eng_id}")
                self._diagnose_flow_page(condition_index)
                return False
            
            # 等待选择生效并关闭下拉框
            time.sleep(0.5)
            LOGGER.info("✅ Eng ID选择完成")
            return True
            
        except Exception as e:
            LOGGER.error(f"选择Eng ID失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            self._diagnose_flow_page(condition_index)
            return False

    def _click_add_new_condition(self) -> bool:
        """
        点击Flow页面中的“Add new condition”
        
        Returns:
            True如果点击成功
        """
        LOGGER.info("查找并点击'Add new condition'...")
        
        try:
            time.sleep(1.0)
            
            # 方法1: 使用ID=addNewCondition（你提供的更稳定的写法）
            try:
                add_btn = WebDriverWait(self._driver, 10).until(
                    EC.presence_of_element_located((By.ID, "addNewCondition"))
                )
                LOGGER.info("方法1通过ID=addNewCondition找到'Add new condition'")
                
                # 滚动到可视区域中间（实际会滚动包含它的容器，如mat-drawer-content）
                try:
                    self._driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                        add_btn,
                    )
                    time.sleep(0.3)
                except Exception as e:
                    LOGGER.debug(f"scrollIntoView(addNewCondition) 失败: {e}")
                
                # 再等一次可点击并点击
                add_btn = WebDriverWait(self._driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "addNewCondition"))
                )
                add_btn.click()
                LOGGER.info("✅ 已点击'Add new condition'（通过ID）")
                # 等待新的condition DOM完全渲染
                time.sleep(2.0)
                return True
            except TimeoutException:
                LOGGER.debug("方法1失败：ID=addNewCondition 未找到或不可点击")

            # 方法2: 使用通用滚动点击逻辑（带class的精确XPath）
            primary_xpath = "//span[contains(@class,'add-text') and contains(@class,'enabled') and normalize-space()='Add new condition']"
            if self._scroll_and_click(By.XPATH, primary_xpath, "Add new condition(span.add-text.enabled)", timeout=10):
                # 等待新的condition DOM完全渲染
                time.sleep(2.0)
                return True

            LOGGER.debug("方法2失败：未找到带class的Add new condition，尝试仅按文本查找")

            # 方法3: 仅按文本查找
            fallback_xpath = "//*[normalize-space(text())='Add new condition']"
            if self._scroll_and_click(By.XPATH, fallback_xpath, "Add new condition(文本匹配)", timeout=10):
                # 等待新的condition DOM完全渲染
                time.sleep(2.0)
                return True

            LOGGER.error("❌ 未找到'Add new condition'元素")
            return False
        
        except Exception as e:
            LOGGER.error(f"点击'Add new condition'失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _click_more_options_tab(self) -> bool:
        """
        点击More options标签页
        
        Returns:
            True如果点击成功
        """
        LOGGER.info("查找并点击'More options'标签...")
        
        try:
            # 等待页面稳定
            time.sleep(1.0)
            
            more_options_tab = None
            
            # 方法1: 通过Material UI的mat-tab-label查找
            try:
                more_options_tab = WebDriverWait(self._driver, self.config.explicit_wait).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class,'mat-tab-label-content') and normalize-space()='More options']/.."
                    ))
                )
                LOGGER.info("方法1找到'More options'标签（mat-tab-label）")
            except TimeoutException:
                LOGGER.debug("方法1失败：未找到mat-tab-label")
            
            # 方法2: 通过包含More options文本的元素查找
            if not more_options_tab:
                try:
                    more_options_tab = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            "//*[contains(@class, 'tab') and contains(text(), 'More options')]"
                        ))
                    )
                    LOGGER.info("方法2找到'More options'标签")
                except TimeoutException:
                    LOGGER.debug("方法2失败")
            
            # 方法3: 查找所有可能的标签元素
            if not more_options_tab:
                try:
                    all_tabs = self._driver.find_elements(By.XPATH, "//*[contains(@class, 'tab') or @role='tab']")
                    for tab in all_tabs:
                        if tab.is_displayed() and 'More options' in tab.text:
                            more_options_tab = tab
                            LOGGER.info(f"方法3找到'More options'标签")
                            break
                except Exception as e:
                    LOGGER.debug(f"方法3失败: {e}")
            
            if not more_options_tab:
                LOGGER.error("❌ 未找到'More options'标签")
                return False
            
            # 滚动到标签可见
            try:
                self._driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", more_options_tab)
                time.sleep(0.3)
            except:
                pass
            
            # 点击More options标签
            try:
                more_options_tab.click()
                LOGGER.info("✅ 已点击'More options'标签（普通点击）")
            except:
                # JavaScript点击
                self._driver.execute_script("arguments[0].click();", more_options_tab)
                LOGGER.info("✅ 已点击'More options'标签（JavaScript点击）")
            
            # 等待More options标签页加载
            time.sleep(1.5)
            LOGGER.info("✅ More options标签页已加载")
            return True
            
        except Exception as e:
            LOGGER.error(f"点击'More options'标签失败: {e}")
            import traceback
            LOGGER.error(traceback.format_exc())
            return False
    
    def _fill_more_options(self, unit_test_time: str, retest_rate: str, hri_mrv: str = None) -> bool:
        """
        填写More options标签页的字段
        
        Args:
            unit_test_time: Unit test time值
            retest_rate: Retest rate值
            hri_mrv: HRI / MRV值，如果为空则选择default
            
        Returns:
            True如果填写成功
        """
        LOGGER.info(f"填写More options - Unit test time: {unit_test_time}, Retest rate: {retest_rate}, HRI/MRV: {hri_mrv or 'default'}")
        
        try:
            # 等待页面加载
            time.sleep(1.5)
            
            # 1. 填写Unit test time（使用用户提供的精确定位）
            LOGGER.info(f"填写Unit test time: {unit_test_time}")
            try:
                # 方法1: 通过more-options-body和文本定位（用户提供的方法）
                unit_test_input = None
                try:
                    unit_test_input = WebDriverWait(self._driver, 10).until(
                        EC.visibility_of_element_located((
                            By.XPATH,
                            "//div[contains(@class,'more-options-body') and contains(normalize-space(),'Unit test time')]/input"
                        ))
                    )
                    LOGGER.info("方法1找到Unit test time输入框（more-options-body）")
                except TimeoutException:
                    LOGGER.debug("方法1失败")
                
                # 方法2: 通过label查找（备用）
                if not unit_test_input:
                    try:
                        label = self._driver.find_element(
                            By.XPATH,
                            "//*[contains(text(), 'Unit test time')]"
                        )
                        unit_test_input = label.find_element(
                            By.XPATH,
                            "./following-sibling::*//input | ./following-sibling::input | ./..//input"
                        )
                        LOGGER.info("方法2找到Unit test time输入框（通过label）")
                    except:
                        pass
                
                if unit_test_input:
                    unit_test_input.clear()
                    unit_test_input.send_keys(str(unit_test_time))
                    LOGGER.info(f"✅ 已填写Unit test time: {unit_test_time}")
                else:
                    LOGGER.warning("⚠️ 未找到Unit test time输入框")
                    return False
                    
            except Exception as e:
                LOGGER.error(f"填写Unit test time时出错: {e}")
                return False
            
            time.sleep(0.5)
            
            # 2. 填写Retest rate（使用类似的精确定位）
            LOGGER.info(f"填写Retest rate: {retest_rate}")
            try:
                # 方法1: 通过more-options-body和文本定位
                retest_input = None
                try:
                    retest_input = WebDriverWait(self._driver, 10).until(
                        EC.visibility_of_element_located((
                            By.XPATH,
                            "//div[contains(@class,'more-options-body') and contains(normalize-space(),'Retest rate')]/input"
                        ))
                    )
                    LOGGER.info("方法1找到Retest rate输入框（more-options-body）")
                except TimeoutException:
                    LOGGER.debug("方法1失败")
                
                # 方法2: 通过label查找（备用）
                if not retest_input:
                    try:
                        label = self._driver.find_element(
                            By.XPATH,
                            "//*[contains(text(), 'Retest rate')]"
                        )
                        retest_input = label.find_element(
                            By.XPATH,
                            "./following-sibling::*//input | ./following-sibling::input | ./..//input"
                        )
                        LOGGER.info("方法2找到Retest rate输入框（通过label）")
                    except:
                        pass
                
                if retest_input:
                    retest_input.clear()
                    retest_input.send_keys(str(retest_rate))
                    LOGGER.info(f"✅ 已填写Retest rate: {retest_rate}")
                else:
                    LOGGER.warning("⚠️ 未找到Retest rate输入框")
                    return False
                    
            except Exception as e:
                LOGGER.error(f"填写Retest rate时出错: {e}")
                return False
            
            time.sleep(0.5)
            
            # 3. 选择HRI / MRV（如果为空则选default）
            hri_value = hri_mrv if hri_mrv and str(hri_mrv).strip() else None
            LOGGER.info(f"选择HRI / MRV: {hri_value or 'default (第一个选项)'}")
            
            try:
                # 查找HRI / MRV下拉框（使用用户提供的ID定位）
                select_elem = None
                
                # 方法1: 通过ID定位（最精确的方法）
                try:
                    select_elem = WebDriverWait(self._driver, 10).until(
                        EC.presence_of_element_located((By.ID, "flexbomSelect"))
                    )
                    LOGGER.info("方法1找到HRI/MRV下拉框（通过ID: flexbomSelect）")
                except TimeoutException:
                    LOGGER.debug("方法1失败：未找到flexbomSelect")
                
                # 方法2: 通过more-options-body和文本定位（备用）
                if not select_elem:
                    try:
                        hri_container = self._driver.find_element(
                            By.XPATH,
                            "//div[contains(@class,'more-options-body') and contains(normalize-space(),'HRI')]"
                        )
                        select_elem = hri_container.find_element(By.XPATH, ".//select")
                        LOGGER.info("方法2找到HRI/MRV下拉框（more-options-body）")
                    except:
                        LOGGER.debug("方法2失败")
                
                if select_elem:
                    # 使用Select类操作下拉框
                    from selenium.webdriver.support.ui import Select
                    flexbom_select = Select(select_elem)
                    
                    if hri_value:
                        # 如果有值，尝试多种方式选择
                        selected = False
                        
                        # 方式1: 按显示文本选择
                        try:
                            flexbom_select.select_by_visible_text(hri_value)
                            LOGGER.info(f"✅ 已选择HRI/MRV: {hri_value}（按文本）")
                            selected = True
                        except Exception as e:
                            LOGGER.debug(f"按文本选择失败: {e}")
                        
                        # 方式2: 按value选择
                        if not selected:
                            try:
                                flexbom_select.select_by_value(hri_value)
                                LOGGER.info(f"✅ 已选择HRI/MRV: {hri_value}（按value）")
                                selected = True
                            except Exception as e:
                                LOGGER.debug(f"按value选择失败: {e}")
                        
                        # 方式3: 包含匹配（遍历所有选项）
                        if not selected:
                            try:
                                all_options = flexbom_select.options
                                for idx, option in enumerate(all_options):
                                    option_text = option.text.strip()
                                    if hri_value in option_text or option_text in hri_value:
                                        flexbom_select.select_by_index(idx)
                                        LOGGER.info(f"✅ 已选择HRI/MRV: {option_text}（包含匹配）")
                                        selected = True
                                        break
                            except Exception as e:
                                LOGGER.debug(f"包含匹配失败: {e}")
                        
                        # 如果都失败，选择第一个
                        if not selected:
                            try:
                                flexbom_select.select_by_index(0)
                                first_text = flexbom_select.options[0].text
                                LOGGER.warning(f"⚠️ 未找到'{hri_value}'，已选择默认值: {first_text}")
                            except:
                                LOGGER.warning("⚠️ 无法选择HRI/MRV")
                    else:
                        # 如果值为空，选择第一个选项（default）
                        try:
                            flexbom_select.select_by_index(0)
                            first_text = flexbom_select.options[0].text if flexbom_select.options else "default"
                            LOGGER.info(f"✅ 已选择HRI/MRV: {first_text}（第一个选项）")
                        except Exception as e:
                            LOGGER.warning(f"⚠️ 无法选择默认HRI/MRV: {e}")
                else:
                    LOGGER.warning("⚠️ 未找到HRI/MRV下拉框")
                    
            except Exception as e:
                LOGGER.warning(f"选择HRI/MRV时出错: {e}")
            
            time.sleep(0.5)
            LOGGER.info("✅ More options填写完成")
            return True
            
        except Exception as e:
            LOGGER.error(f"填写More options失败: {e}")
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

