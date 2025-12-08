"""GTS网站提交最终数据模块"""
import logging
import time
from typing import Optional
from dataclasses import dataclass
from selenium import webdriver
from selenium.webdriver.common.by import By
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
class GTSConfig:
    """GTS网站配置"""
    url: str
    timeout: int = 60
    retry_count: int = 3
    retry_delay: int = 2
    wait_after_submit: int = 5
    headless: bool = False
    implicit_wait: int = 10
    explicit_wait: int = 20


class GTSSubmitter:
    """GTS网站数据提交器"""
    
    def __init__(self, config: GTSConfig):
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
        """导航到GTS页面"""
        if not self._driver:
            raise RuntimeError("WebDriver未初始化")
        
        LOGGER.info(f"导航到GTS页面: {self.config.url}")
        self._driver.get(self.config.url)
        
        # 等待页面加载
        try:
            WebDriverWait(self._driver, self.config.explicit_wait).until(
                lambda d: d.ready_state == 'complete'
            )
            LOGGER.info("✅ 页面加载完成")
        except TimeoutException:
            LOGGER.warning("页面加载超时，继续执行...")
    
    def submit_final_data(self, data: dict) -> bool:
        """
        提交最终数据到GTS网站
        
        Args:
            data: 包含最终数据的字典
        
        Returns:
            True如果提交成功
        
        Raises:
            RuntimeError: 如果提交失败
        """
        LOGGER.info("开始提交最终数据到GTS网站")
        LOGGER.debug(f"最终数据: {data}")
        
        # 重试机制
        last_exception = None
        for attempt in range(1, self.config.retry_count + 1):
            try:
                LOGGER.info(f"尝试提交最终数据 (第{attempt}/{self.config.retry_count}次)")
                
                # 初始化WebDriver
                self._init_driver()
                
                # 导航到页面
                self._navigate_to_page()
                
                # TODO: 根据GTS网站的实际界面实现具体的数据提交逻辑
                # 这里需要根据实际的GTS网站表单来填写数据
                # 示例：查找表单元素、填写数据、点击提交按钮等
                
                # 等待提交完成
                time.sleep(self.config.wait_after_submit)
                
                # 验证提交是否成功
                if self._verify_submission():
                    LOGGER.info("✅ 最终数据提交成功")
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
                    LOGGER.error(f"❌ 最终数据提交失败（已重试{self.config.retry_count}次）")
        
        # 清理资源
        self._close_driver()
        raise RuntimeError(f"最终数据提交失败: {last_exception}")
    
    def _verify_submission(self) -> bool:
        """
        验证数据是否提交成功
        
        Returns:
            True如果验证通过
        """
        try:
            # TODO: 实现验证逻辑
            # 例如：检查页面是否显示成功消息、是否有错误提示等
            # 可以根据实际GTS网站的反馈机制来实现
            
            # 示例：查找成功消息元素
            # success_element = WebDriverWait(self._driver, 10).until(
            #     EC.presence_of_element_located((By.CLASS_NAME, "success-message"))
            # )
            # return success_element is not None
            
            LOGGER.info("✅ 最终数据提交验证通过")
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

