# 工具模块说明

## 截图辅助工具 (screenshot_helper.py)

自动截图并记录错误的辅助工具，用于调试和问题追踪。

### 功能特性

1. **Selenium 浏览器截图**
   - 用于 Spark 和 GTS 模块（基于 Selenium）
   - 自动截取浏览器页面

2. **屏幕截图**
   - 用于 Mole 模块（基于 pywinauto）
   - 截取整个屏幕

3. **自动文件命名**
   - 格式：`{prefix}_{timestamp}.png`
   - 时间戳精确到毫秒

### 使用方法

#### 在 Spark/GTS 中使用（Selenium）

```python
from .utils.screenshot_helper import log_error_with_screenshot

class SparkSubmitter:
    def __init__(self, config, debug_dir=None):
        self.debug_dir = debug_dir or Path.cwd() / "output" / "05_Debug"
    
    def _log_error_with_screenshot(self, error_message, exception=None, prefix="spark_error"):
        """记录错误并自动截图"""
        if self._driver:
            log_error_with_screenshot(self._driver, error_message, self.debug_dir, exception, prefix)
    
    def some_method(self):
        try:
            # 执行操作
            pass
        except Exception as e:
            self._log_error_with_screenshot("操作失败", e, prefix="operation_failed")
```

#### 在 Mole 中使用（屏幕截图）

```python
from .utils.screenshot_helper import log_error_with_screen_screenshot

class MoleSubmitter:
    def __init__(self, config, debug_dir=None):
        self.debug_dir = debug_dir or Path.cwd() / "output" / "05_Debug"
    
    def _log_error_with_screenshot(self, error_message, exception=None, prefix="mole_error"):
        """记录错误并自动截取屏幕"""
        log_error_with_screen_screenshot(error_message, self.debug_dir, exception, prefix)
    
    def some_method(self):
        try:
            # 执行操作
            pass
        except Exception as e:
            self._log_error_with_screenshot("操作失败", e, prefix="operation_failed")
```

### API 参考

#### log_error_with_screenshot()

记录错误并自动截取浏览器截图。

**参数：**
- `driver`: Selenium WebDriver 实例
- `error_message`: 错误消息
- `output_dir`: 输出目录（Path 对象）
- `exception`: 异常对象（可选）
- `prefix`: 文件名前缀（默认："error"）

**示例：**
```python
log_error_with_screenshot(
    driver=self._driver,
    error_message="未找到元素",
    output_dir=Path("output/05_Debug"),
    exception=e,
    prefix="element_not_found"
)
```

#### log_error_with_screen_screenshot()

记录错误并自动截取屏幕。

**参数：**
- `error_message`: 错误消息
- `output_dir`: 输出目录（Path 对象）
- `exception`: 异常对象（可选）
- `prefix`: 文件名前缀（默认："mole_error"）

**示例：**
```python
log_error_with_screen_screenshot(
    error_message="窗口未找到",
    output_dir=Path("output/05_Debug"),
    exception=e,
    prefix="window_not_found"
)
```

#### capture_debug_screenshot()

捕获调试截图（不记录为错误）。

**参数：**
- `driver`: Selenium WebDriver 实例
- `description`: 截图描述
- `output_dir`: 输出目录（Path 对象）
- `prefix`: 文件名前缀（默认："debug"）

**返回：**
- 截图文件路径（Path 对象），失败则返回 None

**示例：**
```python
screenshot_path = capture_debug_screenshot(
    driver=self._driver,
    description="下拉框展开后",
    output_dir=Path("output/05_Debug"),
    prefix="dropdown_opened"
)
```

### 输出示例

截图文件将保存在指定的输出目录中，文件名格式如下：

```
output/05_Debug/
├── spark_error_20251226_025714_123.png
├── add_new_not_found_20251226_025715_456.png
├── mole_error_20251226_025716_789.png
└── debug_20251226_025717_012.png
```

### 日志输出示例

```
2025-12-26 02:57:14 - ERROR - ❌ 未找到'Add New'按钮 [截图: add_new_not_found_20251226_025714_123.png]
2025-12-26 02:57:14 - ERROR - 异常详情: TimeoutException: Message: ...
```

### 依赖项

- **Selenium**: 用于浏览器截图
- **Pillow (PIL)**: 用于屏幕截图（可选）

如果 Pillow 未安装，屏幕截图功能将不可用，但不会影响其他功能。

安装 Pillow：
```bash
pip install Pillow
```

### 注意事项

1. **性能影响**：截图操作会增加一定的执行时间（通常 < 1秒）
2. **磁盘空间**：长时间运行会产生大量截图文件，建议定期清理
3. **敏感信息**：截图可能包含敏感信息，注意保护
4. **自动清理**：可以考虑添加自动清理旧截图的功能

### 未来改进

- [ ] 添加截图压缩功能
- [ ] 添加自动清理旧截图功能
- [ ] 支持截取特定区域
- [ ] 支持视频录制

