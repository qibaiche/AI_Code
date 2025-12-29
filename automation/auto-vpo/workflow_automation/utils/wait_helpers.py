"""等待辅助工具 - 优化的等待逻辑"""
import logging
import time
from typing import Callable, Optional, Any
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

LOGGER = logging.getLogger(__name__)


def wait_for_element(
    driver: webdriver.Chrome,
    by: By,
    value: str,
    timeout: int = 10,
    condition: str = "presence"
) -> Optional[Any]:
    """
    等待元素出现
    
    Args:
        driver: WebDriver 实例
        by: 定位方式（By.ID, By.XPATH 等）
        value: 定位值
        timeout: 超时时间（秒）
        condition: 等待条件
            - "presence": 元素存在于 DOM 中
            - "visible": 元素可见
            - "clickable": 元素可点击
            
    Returns:
        找到的元素，超时则返回 None
    """
    try:
        if condition == "presence":
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
        elif condition == "visible":
            element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((by, value))
            )
        elif condition == "clickable":
            element = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
        else:
            raise ValueError(f"未知的等待条件: {condition}")
        
        return element
    except TimeoutException:
        LOGGER.debug(f"等待元素超时: {by}={value}, condition={condition}, timeout={timeout}s")
        return None


def wait_for_element_disappear(
    driver: webdriver.Chrome,
    by: By,
    value: str,
    timeout: int = 10
) -> bool:
    """
    等待元素消失
    
    Args:
        driver: WebDriver 实例
        by: 定位方式
        value: 定位值
        timeout: 超时时间（秒）
        
    Returns:
        True 如果元素消失，False 如果超时
    """
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((by, value))
        )
        return True
    except TimeoutException:
        LOGGER.debug(f"等待元素消失超时: {by}={value}, timeout={timeout}s")
        return False


def wait_for_condition(
    condition_func: Callable[[], bool],
    timeout: int = 10,
    poll_frequency: float = 0.5,
    error_message: str = "等待条件超时"
) -> bool:
    """
    等待自定义条件满足
    
    Args:
        condition_func: 条件函数，返回 True 表示条件满足
        timeout: 超时时间（秒）
        poll_frequency: 轮询频率（秒）
        error_message: 超时错误消息
        
    Returns:
        True 如果条件满足，False 如果超时
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            if condition_func():
                return True
        except Exception as e:
            LOGGER.debug(f"条件检查出错: {e}")
        time.sleep(poll_frequency)
    
    LOGGER.debug(f"{error_message} (timeout={timeout}s)")
    return False


def wait_for_page_load(
    driver: webdriver.Chrome,
    timeout: int = 30
) -> bool:
    """
    等待页面加载完成
    
    Args:
        driver: WebDriver 实例
        timeout: 超时时间（秒）
        
    Returns:
        True 如果页面加载完成，False 如果超时
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        return True
    except TimeoutException:
        LOGGER.debug(f"等待页面加载超时 (timeout={timeout}s)")
        return False


def wait_for_ajax_complete(
    driver: webdriver.Chrome,
    timeout: int = 10
) -> bool:
    """
    等待 AJAX 请求完成（jQuery）
    
    Args:
        driver: WebDriver 实例
        timeout: 超时时间（秒）
        
    Returns:
        True 如果 AJAX 完成，False 如果超时或页面不使用 jQuery
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return typeof jQuery != 'undefined' && jQuery.active == 0")
        )
        return True
    except (TimeoutException, Exception):
        # 可能页面不使用 jQuery
        return False


def wait_for_angular_load(
    driver: webdriver.Chrome,
    timeout: int = 10
) -> bool:
    """
    等待 Angular 加载完成
    
    Args:
        driver: WebDriver 实例
        timeout: 超时时间（秒）
        
    Returns:
        True 如果 Angular 加载完成，False 如果超时或页面不使用 Angular
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script(
                "return typeof angular != 'undefined' && "
                "angular.element(document).injector().get('$http').pendingRequests.length === 0"
            )
        )
        return True
    except (TimeoutException, Exception):
        # 可能页面不使用 Angular
        return False


def smart_wait(
    driver: webdriver.Chrome,
    min_wait: float = 0.5,
    max_wait: int = 10
) -> None:
    """
    智能等待：先等待页面加载，再等待 AJAX/Angular
    
    Args:
        driver: WebDriver 实例
        min_wait: 最小等待时间（秒）
        max_wait: 最大等待时间（秒）
    """
    # 最小等待
    time.sleep(min_wait)
    
    # 等待页面加载
    wait_for_page_load(driver, max_wait)
    
    # 等待 AJAX（如果有）
    wait_for_ajax_complete(driver, max_wait)
    
    # 等待 Angular（如果有）
    wait_for_angular_load(driver, max_wait)


def wait_and_click(
    driver: webdriver.Chrome,
    by: By,
    value: str,
    timeout: int = 10,
    use_js: bool = False
) -> bool:
    """
    等待元素可点击并点击
    
    Args:
        driver: WebDriver 实例
        by: 定位方式
        value: 定位值
        timeout: 超时时间（秒）
        use_js: 是否使用 JavaScript 点击
        
    Returns:
        True 如果点击成功，False 如果失败
    """
    element = wait_for_element(driver, by, value, timeout, "clickable")
    if not element:
        return False
    
    try:
        if use_js:
            driver.execute_script("arguments[0].click();", element)
        else:
            element.click()
        return True
    except Exception as e:
        LOGGER.debug(f"点击元素失败: {e}")
        # 尝试 JavaScript 点击
        if not use_js:
            try:
                driver.execute_script("arguments[0].click();", element)
                return True
            except Exception as e2:
                LOGGER.debug(f"JavaScript 点击也失败: {e2}")
        return False


def wait_and_send_keys(
    driver: webdriver.Chrome,
    by: By,
    value: str,
    keys: str,
    timeout: int = 10,
    clear_first: bool = True
) -> bool:
    """
    等待元素可见并输入文本
    
    Args:
        driver: WebDriver 实例
        by: 定位方式
        value: 定位值
        keys: 要输入的文本
        timeout: 超时时间（秒）
        clear_first: 是否先清空
        
    Returns:
        True 如果输入成功，False 如果失败
    """
    element = wait_for_element(driver, by, value, timeout, "visible")
    if not element:
        return False
    
    try:
        if clear_first:
            element.clear()
        element.send_keys(keys)
        return True
    except Exception as e:
        LOGGER.debug(f"输入文本失败: {e}")
        return False

