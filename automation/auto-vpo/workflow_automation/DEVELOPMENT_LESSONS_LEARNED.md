# 自动化开发经验总结与教训

## 1. Mole自动化开发中遇到的问题及解决方案

### 1.1 对话框识别问题

#### 问题1: 误把主窗口当成对话框
**现象**: 检测对话框时，将标题包含"Submit MIR Request"的主窗口误认为是对话框。

**根本原因**: 
- 只检查窗口标题关键词，没有验证窗口类型
- 没有排除主窗口本身

**解决方案**:
```python
# 1. 排除主窗口
main_hwnd = self._window.handle
if main_hwnd and hwnd == main_hwnd:
    return True

# 2. 检查窗口类名，确认是真正的对话框
class_name = win32gui.GetClassName(hwnd)
if "dialog" in class_name.lower() or "#32770" in class_name:
    # 这才是真正的对话框
```

**教训**: 
- ✅ 必须验证窗口类型（类名），不能只看标题
- ✅ 必须排除主窗口，避免误判

---

#### 问题2: 对话框类名不是标准类型
**现象**: "Submit MIR Request"对话框的类名是`WindowsForms10.Window.8.app.0.2ac7198_r9_ad1`，不是标准的`#32770`。

**根本原因**: 
- Windows Forms应用的对话框可能使用自定义类名
- 只检查标准对话框类名(`#32770`)导致遗漏

**解决方案**:
```python
# 判断是否是对话框的多种方式：
is_dialog = False
# 1. 标准对话框类名
if "dialog" in class_name.lower() or "#32770" in class_name:
    is_dialog = True
# 2. 特定标题的窗口（即使类名不标准）
elif window_text in ["Warning", "Submit MIR Request"]:
    is_dialog = True
    LOGGER.info(f"根据标题匹配为对话框: '{window_text}'")
```

**教训**:
- ✅ 对话框识别要多种方式结合：类名 + 标题 + 可见性
- ✅ 对于已知的特定对话框，可以直接用标题匹配
- ✅ 要先扫描所有窗口，记录调试信息，了解实际的窗口结构

---

### 1.2 按钮点击问题

#### 问题3: 按钮文本包含特殊字符
**现象**: 按钮文本是`&Yes`（带&符号），代码查找`"YES"`匹配失败。

**根本原因**: 
- Windows按钮使用`&`标记快捷键（显示时会下划线字母）
- 直接比较文本不匹配

**解决方案**:
```python
# 去掉&符号后比较
clean_text = text.replace("&", "").upper()
if clean_text == "YES":
    yes_button_info["hwnd"] = hwnd_child
```

**教训**:
- ✅ 处理按钮文本前，先清理特殊字符（`&`, 空格等）
- ✅ 使用大小写不敏感的比较（`.upper()` 或 `.lower()`）

---

#### 问题4: 按键操作有副作用
**现象**: 使用`pyautogui.press('enter')`点击按钮时，默认选中了"No"按钮而不是"Yes"。

**根本原因**: 
- 对话框的默认按钮（Default Button）可能是"No"
- Enter键会点击默认按钮，而不是当前高亮的按钮

**解决方案**:
```python
# 1. 优先使用坐标点击（物理鼠标点击）
rect = win32gui.GetWindowRect(hwnd_child)
center_x = (rect[0] + rect[2]) // 2
center_y = (rect[1] + rect[3]) // 2
pyautogui.moveTo(center_x, center_y, duration=0.2)
pyautogui.click()

# 2. Fallback: 使用快捷键（'y'），但绝不用Enter
pyautogui.press('y')

# 3. 绝不使用Enter或Tab+Enter
# ❌ pyautogui.press('enter')  # 会点击默认按钮(No)
```

**教训**:
- ✅ 坐标点击最可靠，优先使用
- ✅ 使用快捷键时，选择明确的键（如'y'），不用Enter/Tab
- ✅ 仔细测试每个按键的实际效果，不要假设

---

#### 问题5: Windows API点击无效但不报错
**现象**: `win32gui.PostMessage(hwnd, BM_CLICK, 0, 0)`调用成功，但按钮没有实际被点击。

**根本原因**: 
- 按钮可能需要先激活窗口
- 某些控件不响应`PostMessage`，需要`SendMessage`
- 或者需要物理点击

**解决方案**:
```python
# 1. 先激活对话框窗口
win32gui.SetForegroundWindow(dialog_hwnd)
win32gui.BringWindowToTop(dialog_hwnd)
time.sleep(0.5)  # 等待窗口激活

# 2. 找到按钮坐标后，使用物理点击
pyautogui.moveTo(center_x, center_y, duration=0.2)
pyautogui.click()
```

**教训**:
- ✅ Windows API调用前，必须先激活目标窗口
- ✅ API调用后，通过日志和实际界面验证是否真的成功
- ✅ 不要依赖单一方法，准备多种Fallback方案

---

### 1.3 点击坐标偏移问题

#### 问题6: 点击位置不准确
**现象**: 点击"View Summary"标签时点到了右侧，点击"Add to Summary"按钮时偶尔失败。

**根本原因**: 
- 使用固定偏移量或固定百分比，不同分辨率/窗口大小不准确
- 没有考虑按钮边缘的点击容差

**解决方案**:
```python
# 1. 使用相对定位：基于参考元素计算
reference_element = find_element("1. Transfer Type")
target_x = reference_element.center_x + offset_x
target_y = reference_element.center_y + offset_y

# 2. 点击按钮中心，并可调整偏移
center_y = (rect[1] + rect[3]) // 2 - 8  # 往上移8像素，避免边缘

# 3. 使用百分比时，基于实际控件位置而非窗口大小
```

**教训**:
- ✅ 优先使用相对定位（相对于已知元素）
- ✅ 点击按钮时，稍微往中心移动，避免边缘
- ✅ 提供可调整的偏移参数，方便微调

---

### 1.4 多对话框处理问题

#### 问题7: 第二个对话框未被检测到
**现象**: 第一个对话框处理成功，但第二个对话框一直等待超时。

**根本原因**: 
- 等待时间不够（对话框弹出需要时间）
- 对话框已经在`processed_dialogs`列表中（重复检测）
- 检测逻辑太严格，漏掉了第二个对话框

**解决方案**:
```python
# 1. 增加等待时间
wait_time = 2.0 if i == 0 else 4.0  # 后续对话框等更久

# 2. 多次重试检测
for retry in range(max_retries):
    # 扫描对话框
    if new_dialogs:
        break
    elif retry < max_retries - 1:
        time.sleep(1.5)  # 等待后重试

# 3. 避免重复处理
new_dialogs = [(hwnd, title) for hwnd, title in dialogs 
               if title not in processed_dialogs]

# 4. 添加详细调试日志
LOGGER.info(f"当前所有可见窗口（共{len(all_visible_windows)}个）：")
for win_text, win_class in all_visible_windows:
    LOGGER.info(f"    窗口: '{win_text}' (类名: {win_class})")
```

**教训**:
- ✅ 对话框弹出需要时间，要有足够的等待和重试
- ✅ 跟踪已处理的对话框，避免重复
- ✅ 添加详细日志，列出所有可见窗口，便于调试
- ✅ 处理对话框后，等待一段时间再检测下一个

---

### 1.5 调试困难问题

#### 问题8: 无法知道实际发生了什么
**现象**: 日志显示"已点击按钮"，但实际界面没有变化。

**根本原因**: 
- 日志过于乐观，没有真实反映操作结果
- 缺少中间状态的详细信息

**解决方案**:
```python
# 1. 区分"尝试"和"成功"
LOGGER.info("正在尝试点击按钮...")
if button_clicked:
    LOGGER.info("✅ 已成功点击按钮")
else:
    LOGGER.error("❌ 未能点击按钮")

# 2. 记录所有扫描到的元素
LOGGER.info(f"扫描到 {len(buttons)} 个按钮：")
for button in buttons:
    LOGGER.info(f"  按钮: '{button.text}' (enabled: {button.is_enabled()})")

# 3. 记录Windows API调用的返回值
found_button = {"found": False}
# ...
if found_button["found"]:
    LOGGER.info("✅ 找到并点击了按钮")
else:
    LOGGER.warning("⚠️ Windows API未找到按钮")

# 4. 验证操作结果
time.sleep(0.5)
if self._verify_submit_success():
    LOGGER.info("✅ 验证通过")
```

**教训**:
- ✅ 日志要诚实反映实际情况，不要假设成功
- ✅ 扫描元素时，列出所有找到的元素（调试时很有用）
- ✅ 操作后要验证结果，不要盲目继续
- ✅ 使用明确的符号（✅ ❌ ⚠️）区分成功/失败/警告

---

## 2. 通用开发指南（适用于Spark和其他自动化）

### 2.1 窗口/元素识别原则
1. **多重验证**: 不要只看一个属性（标题、类名、文本等）
2. **排除法**: 先排除不可能的（隐藏的、主窗口、不可用的）
3. **调试优先**: 先扫描并记录所有元素，再写识别逻辑
4. **灵活匹配**: 准备多种匹配方式（精确、模糊、正则）

### 2.2 点击操作原则
1. **激活窗口**: 点击前先激活目标窗口/对话框
2. **等待响应**: 操作后等待界面响应（0.5-2秒）
3. **坐标优先**: 物理点击最可靠
4. **多重Fallback**: 准备3-4种点击方法
5. **验证结果**: 点击后检查是否真的成功

### 2.3 文本处理原则
1. **清理文本**: 去掉`&`、空格、换行等
2. **大小写不敏感**: 统一转大写或小写
3. **去除前后空格**: 使用`.strip()`
4. **考虑本地化**: 如果有多语言，准备多个匹配词

### 2.4 等待和重试原则
1. **合理等待**: 不同操作需要不同等待时间
   - 窗口激活: 0.5秒
   - 按钮点击后: 0.5-1秒
   - 对话框弹出: 1.5-2秒
   - 页面加载: 2-5秒
2. **重试机制**: 关键操作至少重试3次
3. **渐进等待**: 第一次失败后，增加等待时间再重试
4. **超时保护**: 设置最大重试次数，避免死循环

### 2.5 日志记录原则
1. **详细但结构化**: 每个步骤都记录，但用缩进/分隔符组织
2. **区分级别**: 正常(INFO) / 警告(WARNING) / 错误(ERROR)
3. **记录关键信息**: 窗口句柄、坐标、元素文本、操作结果
4. **便于搜索**: 使用统一的标记（✅ ❌ 🔍 ⚠️）
5. **调试开关**: 详细扫描信息可以用DEBUG级别

### 2.6 错误处理原则
1. **捕获具体异常**: 不要用空`except:`
2. **保存上下文**: 记录异常时的窗口状态、元素信息
3. **优雅降级**: 尽可能继续执行，不要轻易中断整个流程
4. **用户友好**: 错误消息要告诉用户可以做什么（如"请手动点击"）

---

## 3. 代码模板

### 3.1 查找并点击按钮的标准流程
```python
def _click_button_standard(self, button_text: str) -> bool:
    """标准按钮点击流程"""
    LOGGER.info(f"查找并点击'{button_text}'按钮...")
    
    # 1. 激活窗口
    self._window.set_focus()
    if win32gui:
        win32gui.SetForegroundWindow(self._window.handle)
        win32gui.BringWindowToTop(self._window.handle)
    time.sleep(0.5)
    
    button_clicked = False
    
    # 2. 方法1: pywinauto查找
    try:
        button = self._window.child_window(title=button_text, control_type="Button")
        if button.exists() and button.is_enabled():
            button.click_input()
            button_clicked = True
            LOGGER.info(f"✅ 通过pywinauto点击'{button_text}'")
    except Exception as e:
        LOGGER.debug(f"pywinauto方法失败: {e}")
    
    # 3. 方法2: Windows API查找
    if not button_clicked and win32gui:
        try:
            found = {"hwnd": None, "rect": None}
            
            def find_button(hwnd_child, _):
                text = win32gui.GetWindowText(hwnd_child).strip()
                class_name = win32gui.GetClassName(hwnd_child)
                if "BUTTON" in class_name.upper():
                    clean_text = text.replace("&", "").strip()
                    if clean_text.upper() == button_text.upper():
                        found["hwnd"] = hwnd_child
                        found["rect"] = win32gui.GetWindowRect(hwnd_child)
                        return False
                return True
            
            win32gui.EnumChildWindows(self._window.handle, find_button, None)
            
            if found["hwnd"] and found["rect"]:
                rect = found["rect"]
                center_x = (rect[0] + rect[2]) // 2
                center_y = (rect[1] + rect[3]) // 2
                
                import pyautogui
                pyautogui.moveTo(center_x, center_y, duration=0.2)
                pyautogui.click()
                button_clicked = True
                LOGGER.info(f"✅ 通过坐标点击'{button_text}' at ({center_x}, {center_y})")
        except Exception as e:
            LOGGER.debug(f"Windows API方法失败: {e}")
    
    # 4. 等待响应
    if button_clicked:
        time.sleep(0.5)
        return True
    else:
        LOGGER.error(f"❌ 未能点击'{button_text}'按钮")
        return False
```

### 3.2 处理对话框的标准流程
```python
def _handle_dialog_standard(self, dialog_title: str, button_text: str) -> bool:
    """标准对话框处理流程"""
    LOGGER.info(f"等待'{dialog_title}'对话框...")
    
    # 1. 等待对话框出现
    dialog_hwnd = None
    for attempt in range(5):
        if win32gui:
            def find_dialog(hwnd, dialogs):
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                
                if title == dialog_title:
                    # 验证是对话框
                    if "#32770" in class_name or "dialog" in class_name.lower():
                        dialogs.append(hwnd)
                    elif title in ["Warning", "Submit MIR Request"]:  # 已知的特殊对话框
                        dialogs.append(hwnd)
                return True
            
            dialogs = []
            win32gui.EnumWindows(find_dialog, dialogs)
            
            if dialogs:
                dialog_hwnd = dialogs[0]
                LOGGER.info(f"✅ 找到对话框: '{dialog_title}'")
                break
        
        if attempt < 4:
            time.sleep(1.0)
    
    if not dialog_hwnd:
        LOGGER.error(f"❌ 未找到'{dialog_title}'对话框")
        return False
    
    # 2. 激活对话框
    win32gui.SetForegroundWindow(dialog_hwnd)
    win32gui.BringWindowToTop(dialog_hwnd)
    time.sleep(0.5)
    
    # 3. 查找并点击按钮
    button_info = {"hwnd": None, "rect": None}
    
    def find_button(hwnd_child, _):
        text = win32gui.GetWindowText(hwnd_child).strip()
        class_name = win32gui.GetClassName(hwnd_child)
        
        if "BUTTON" in class_name.upper():
            clean_text = text.replace("&", "").strip()
            if clean_text.upper() == button_text.upper():
                button_info["hwnd"] = hwnd_child
                button_info["rect"] = win32gui.GetWindowRect(hwnd_child)
                return False
        return True
    
    win32gui.EnumChildWindows(dialog_hwnd, find_button, None)
    
    if button_info["hwnd"] and button_info["rect"]:
        rect = button_info["rect"]
        center_x = (rect[0] + rect[2]) // 2
        center_y = (rect[1] + rect[3]) // 2
        
        import pyautogui
        pyautogui.moveTo(center_x, center_y, duration=0.2)
        pyautogui.click()
        LOGGER.info(f"✅ 已点击'{button_text}'按钮")
        time.sleep(0.5)
        return True
    else:
        LOGGER.error(f"❌ 未找到'{button_text}'按钮")
        return False
```

---

## 4. 开发Spark自动化的建议

### 4.1 前期准备
1. **先手动操作**: 完整手动操作一遍Spark流程，记录每个步骤
2. **识别元素**: 使用浏览器开发者工具识别所有需要操作的元素（ID、Class、XPath）
3. **检查动态加载**: 注意哪些元素是动态加载的，需要等待
4. **测试数据**: 准备多组测试数据（正常、边界、异常）

### 4.2 开发步骤
1. **分步实现**: 每个操作写成独立方法
2. **先验证再继续**: 每步操作后验证是否成功
3. **添加详细日志**: 记录每个元素的定位和操作结果
4. **逐步测试**: 实现一步测试一步，不要写完所有代码再测试

### 4.3 可能遇到的问题
- 页面加载慢 → 增加等待时间
- 元素定位不到 → 检查是否在iframe中、是否动态加载
- 弹出窗口 → 使用`driver.switch_to.window()`切换
- 验证码 → 可能需要手动介入或OCR识别
- 会话超时 → 增加心跳保持或重新登录

---

## 5. 检查清单

开发新的自动化功能时，确保：

**设计阶段**:
- [ ] 手动操作流程已记录
- [ ] 所有元素识别方式已确定
- [ ] 异常情况已考虑（网络错误、超时、元素不存在等）

**编码阶段**:
- [ ] 每个操作都有详细日志
- [ ] 每个操作都有错误处理
- [ ] 每个操作都有验证机制
- [ ] 使用了多重Fallback方案
- [ ] 文本比较前已清理和规范化

**测试阶段**:
- [ ] 正常流程测试通过
- [ ] 异常情况已测试（网络断开、元素不存在、权限不足等）
- [ ] 在不同环境测试（不同分辨率、不同系统版本）
- [ ] 日志信息清晰可读
- [ ] 错误消息对用户友好

**部署阶段**:
- [ ] 配置项已文档化
- [ ] 依赖已列出（requirements.txt）
- [ ] 使用说明已编写
- [ ] 常见问题已记录

---

## 6. 总结

**最重要的3条经验**:
1. 🔍 **先扫描再识别**: 不要假设元素结构，先用日志列出所有元素
2. ✅ **验证每一步**: 操作后必须验证是否成功，不要盲目继续
3. 🔄 **准备Fallback**: 没有一种方法是100%可靠的，准备多种备用方案

**最容易犯的3个错误**:
1. ❌ 假设操作成功，不验证结果
2. ❌ 硬编码值（坐标、索引），不考虑不同环境
3. ❌ 日志过于简略，调试时无法定位问题

记住：**自动化开发是迭代的过程，从简单到复杂，从单一方法到多重方案，从粗略日志到详细追踪。**

