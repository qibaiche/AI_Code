# Spark Submitter 重构计划

## 现状分析

**当前文件：** `spark_submitter.py`  
**代码行数：** 4700+ 行  
**主要问题：**
1. 单个文件过大，难以维护
2. 职责不清晰，混合了配置、元素定位、操作逻辑
3. 难以测试和复用

## 重构目标

将 `spark_submitter.py` 拆分为多个模块，每个模块职责单一，易于维护和测试。

## 拆分方案

### 新的目录结构

```
spark/
├── __init__.py
├── config.py              # 配置类
├── elements.py            # 元素定位器
├── actions/               # 操作逻辑
│   ├── __init__.py
│   ├── navigation.py      # 页面导航
│   ├── material.py        # Material 标签操作
│   ├── flow.py            # Flow 标签操作
│   └── submission.py      # 提交操作
├── validators.py          # 验证逻辑
├── submitter.py           # 主入口（协调各模块）
└── README.md              # 使用说明
```

### 模块职责划分

#### 1. config.py（配置类）

**内容：**
- `SparkConfig` 类定义
- 配置验证逻辑

**代码量：** ~50 行

```python
@dataclass
class SparkConfig:
    """Spark网页配置"""
    url: str
    vpo_category: str = "correlation"
    step: str = "B5"
    tags: str = "CCG_24J-TEST"
    # ... 其他配置
```

#### 2. elements.py（元素定位器）

**内容：**
- 所有元素的定位器（XPath, CSS Selector）
- 统一管理，便于维护

**代码量：** ~200 行

```python
class SparkElements:
    """Spark 页面元素定位器"""
    
    # Add New 按钮
    ADD_NEW_BUTTON = (By.XPATH, "//button[text()='Add New']")
    
    # Test Program Path 输入框
    TP_PATH_INPUT = (By.ID, "testProgramPath")
    
    # Material 标签
    MATERIAL_TAB = (By.XPATH, "//div[contains(@class,'mat-tab-label-content') and normalize-space()='Material']/..")
    
    # ... 其他元素
```

#### 3. actions/navigation.py（页面导航）

**内容：**
- 打开页面
- 等待页面加载
- 页面刷新

**代码量：** ~100 行

```python
class NavigationActions:
    """页面导航操作"""
    
    def __init__(self, driver, config, debug_dir):
        self.driver = driver
        self.config = config
        self.debug_dir = debug_dir
    
    def navigate_to_page(self) -> bool:
        """导航到 Spark 页面"""
        # ...
    
    def wait_for_page_ready(self) -> bool:
        """等待页面就绪"""
        # ...
```

#### 4. actions/material.py（Material 标签操作）

**内容：**
- 点击 Add New
- 填写 TP Path
- 添加 Lot Name
- 选择 Part Type
- 设置 Units 数量

**代码量：** ~800 行

```python
class MaterialActions:
    """Material 标签操作"""
    
    def __init__(self, driver, config, debug_dir):
        self.driver = driver
        self.config = config
        self.debug_dir = debug_dir
    
    def click_add_new(self) -> bool:
        """点击 Add New 按钮"""
        # ...
    
    def fill_tp_path(self, path: str) -> bool:
        """填写 Test Program Path"""
        # ...
    
    def add_lot_name(self, lot_name: str, quantity: int = None) -> bool:
        """添加 Lot Name"""
        # ...
    
    def select_parttype(self, part_type: str) -> bool:
        """选择 Part Type"""
        # ...
```

#### 5. actions/flow.py（Flow 标签操作）

**内容：**
- 点击 Flow 标签
- 选择 Operation
- 选择 Eng ID

**代码量：** ~500 行

```python
class FlowActions:
    """Flow 标签操作"""
    
    def __init__(self, driver, config, debug_dir):
        self.driver = driver
        self.config = config
        self.debug_dir = debug_dir
    
    def click_flow_tab(self) -> bool:
        """点击 Flow 标签"""
        # ...
    
    def select_operation(self, operation: str) -> bool:
        """选择 Operation"""
        # ...
    
    def select_eng_id(self, eng_id: str) -> bool:
        """选择 Eng ID"""
        # ...
```

#### 6. actions/submission.py（提交操作）

**内容：**
- 填写 More Options
- 点击 Roll 按钮
- 收集 VPO 信息

**代码量：** ~600 行

```python
class SubmissionActions:
    """提交操作"""
    
    def __init__(self, driver, config, debug_dir):
        self.driver = driver
        self.config = config
        self.debug_dir = debug_dir
    
    def fill_more_options(self, unit_test_time, retest_rate, hri_mrv) -> bool:
        """填写 More Options"""
        # ...
    
    def click_roll_button(self) -> bool:
        """点击 Roll 按钮"""
        # ...
    
    def collect_vpo_info(self) -> dict:
        """收集 VPO 信息"""
        # ...
```

#### 7. validators.py（验证逻辑）

**内容：**
- 输入验证
- 状态验证
- 结果验证

**代码量：** ~200 行

```python
class SparkValidators:
    """Spark 验证逻辑"""
    
    @staticmethod
    def validate_tp_path(path: str) -> bool:
        """验证 TP Path 格式"""
        # ...
    
    @staticmethod
    def validate_lot_name(lot_name: str) -> bool:
        """验证 Lot Name 格式"""
        # ...
    
    @staticmethod
    def validate_part_type(part_type: str) -> bool:
        """验证 Part Type 格式"""
        # ...
```

#### 8. submitter.py（主入口）

**内容：**
- 协调各个模块
- 提供统一的接口
- WebDriver 管理

**代码量：** ~300 行

```python
class SparkSubmitter:
    """Spark 网页数据提交器（主入口）"""
    
    def __init__(self, config: SparkConfig, debug_dir: Optional[Path] = None):
        self.config = config
        self._driver = None
        self.debug_dir = debug_dir or Path.cwd() / "output" / "05_Debug"
        
        # 初始化各个操作模块
        self.navigation = None
        self.material = None
        self.flow = None
        self.submission = None
    
    def _init_driver(self) -> None:
        """初始化 WebDriver"""
        # ...
        
        # 初始化操作模块
        self.navigation = NavigationActions(self._driver, self.config, self.debug_dir)
        self.material = MaterialActions(self._driver, self.config, self.debug_dir)
        self.flow = FlowActions(self._driver, self.config, self.debug_dir)
        self.submission = SubmissionActions(self._driver, self.config, self.debug_dir)
    
    def submit_data(self, data: dict) -> bool:
        """提交数据（主方法）"""
        # 导航到页面
        if not self.navigation.navigate_to_page():
            return False
        
        # Material 标签操作
        if not self.material.click_add_new():
            return False
        
        if not self.material.fill_tp_path(data['tp_path']):
            return False
        
        # ... 其他操作
        
        return True
```

## 重构步骤

### 阶段 1：准备工作

1. **备份现有代码**
   ```bash
   cp spark_submitter.py spark_submitter_backup.py
   ```

2. **创建新目录结构**
   ```bash
   mkdir -p spark/actions
   ```

3. **创建空文件**
   ```bash
   touch spark/__init__.py
   touch spark/config.py
   touch spark/elements.py
   touch spark/validators.py
   touch spark/submitter.py
   touch spark/actions/__init__.py
   touch spark/actions/navigation.py
   touch spark/actions/material.py
   touch spark/actions/flow.py
   touch spark/actions/submission.py
   ```

### 阶段 2：提取配置类

1. 将 `SparkConfig` 移到 `config.py`
2. 更新导入

### 阶段 3：提取元素定位器

1. 识别所有元素定位代码
2. 统一格式，移到 `elements.py`
3. 更新所有引用

### 阶段 4：提取操作逻辑

1. **Material 操作**
   - 提取 `_click_add_new_button` 等方法
   - 移到 `actions/material.py`

2. **Flow 操作**
   - 提取 `_click_flow_tab` 等方法
   - 移到 `actions/flow.py`

3. **Submission 操作**
   - 提取 `_click_roll_button` 等方法
   - 移到 `actions/submission.py`

### 阶段 5：重构主入口

1. 简化 `submitter.py`
2. 使用组合模式调用各模块
3. 保持向后兼容

### 阶段 6：测试验证

1. 运行现有测试
2. 确保功能不变
3. 修复发现的问题

### 阶段 7：清理

1. 删除备份文件
2. 更新文档
3. 提交代码

## 兼容性策略

为了保持向后兼容，可以保留原有的 `spark_submitter.py`，让它导入新模块：

```python
# spark_submitter.py（兼容层）
from .spark.config import SparkConfig
from .spark.submitter import SparkSubmitter

# 保持原有的导入路径可用
__all__ = ['SparkConfig', 'SparkSubmitter']
```

## 优势

### 1. 可维护性

- **单一职责**：每个模块只负责一件事
- **易于定位**：问题更容易定位和修复
- **代码复用**：操作逻辑可以在不同场景复用

### 2. 可测试性

- **单元测试**：可以单独测试每个模块
- **模拟测试**：可以模拟 WebDriver
- **集成测试**：可以测试模块间的协作

### 3. 可扩展性

- **新增功能**：只需修改相关模块
- **替换实现**：可以替换某个模块的实现
- **多版本支持**：可以支持不同版本的 Spark 页面

### 4. 团队协作

- **并行开发**：多人可以同时修改不同模块
- **代码审查**：更容易进行代码审查
- **知识传递**：新人更容易理解代码结构

## 风险与对策

### 风险 1：破坏现有功能

**对策：**
- 保留备份文件
- 逐步迁移，每次迁移一个模块
- 充分测试

### 风险 2：导入路径变化

**对策：**
- 保留兼容层
- 更新所有导入语句
- 提供迁移指南

### 风险 3：性能下降

**对策：**
- 避免过度封装
- 保持方法调用简洁
- 性能测试对比

## 时间估算

- **阶段 1（准备）**：30 分钟
- **阶段 2（配置）**：30 分钟
- **阶段 3（元素）**：1 小时
- **阶段 4（操作）**：3 小时
- **阶段 5（主入口）**：1 小时
- **阶段 6（测试）**：2 小时
- **阶段 7（清理）**：30 分钟

**总计：** 约 8.5 小时

## 建议

1. **不要急于重构**：当前代码虽然大，但功能完整
2. **优先修复 Bug**：先确保功能稳定
3. **逐步迁移**：可以先迁移一个模块，验证后再继续
4. **保持向后兼容**：避免影响现有使用者

## 下一步

建议在以下情况下进行重构：

1. ✅ 功能已经稳定，没有频繁变更
2. ✅ 有充足的时间进行测试
3. ✅ 团队成员都了解重构计划
4. ✅ 有完整的测试用例

**当前建议：** 暂不执行重构，先完成其他优化任务。等代码稳定后再考虑重构。

