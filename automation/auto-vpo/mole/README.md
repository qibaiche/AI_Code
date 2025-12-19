# Mole 自动化工具使用说明

## 📋 目录

- [概述](#概述)
- [新功能：配置UI](#新功能配置ui) ⭐ NEW
- [配置说明](#配置说明)
- [需要的文件和依赖](#需要的文件和依赖)
- [使用方法](#使用方法)
- [工作流程](#工作流程)
- [常见问题](#常见问题)

## 概述

Mole 自动化工具用于自动提交 MIR（Material Information Request）数据到 Mole 工具。工具会自动执行以下操作：

1. **显示配置UI** ⭐ NEW - 选择搜索方式和配置参数
2. 启动或连接 Mole 工具
3. 处理登录对话框
4. 打开 File 菜单 -> New MIR Request
5. 根据配置选择搜索方式（VPOs / Units / Inactive Cage）⭐ NEW
6. 填写搜索对话框
7. 检查搜索结果行状态并选择
8. 点击 Submit 按钮
9. 处理对话框并获取 MIR 号码

## 新功能：配置UI

### ⭐ 版本 1.1.0 新增功能

#### 图形化配置界面

运行 Mole 自动化时，会自动弹出配置界面，让您可以：

1. **选择搜索方式**
   - Search By VPOs（默认）
   - Search By Units（新增）
   - Search Inactive Cage（新增）

2. **配置参数**
   - Product Code（可选）
   - Source Lot（可选）
   - 其他特定参数（根据搜索方式）

3. **配置管理**
   - 自动保存配置
   - 自动加载上次配置
   - 重置为默认值

#### 快速开始

```bash
# 1. 运行程序
双击 1-测试Mole.bat

# 2. 配置界面会自动弹出
#    - 选择搜索方式
#    - 填写参数（可选）
#    - 点击"确定"

# 3. 程序自动执行
```

#### 详细文档

- 📖 **使用指南**: [Mole配置UI使用说明.md](Mole配置UI使用说明.md)
- 🚀 **快速开始**: [快速开始.md](快速开始.md)
- 📝 **功能总结**: [功能总结.md](功能总结.md)
- 📋 **更新日志**: [更新日志.md](更新日志.md)

## 配置说明

### 配置文件位置

配置文件位于：`workflow_automation/config.yaml`

### Mole 配置项

在 `config.yaml` 文件中，找到 `mole:` 部分进行配置：

```yaml
mole:
  executable_path: "C:/Users/qibaiche/Desktop/Mole 2.0.appref-ms"  # Mole工具可执行文件路径
  window_title: ".*MOLE.*"  # Mole工具窗口标题（支持正则表达式）
  login_dialog_title: "MOLE LOGIN"  # 登录对话框标题
  timeout: 60  # 操作超时时间（秒）
  retry_count: 3  # 重试次数
  retry_delay: 2  # 重试延迟（秒）
```

### 配置项说明

| 配置项 | 说明 | 示例值 | 必填 |
|--------|------|--------|------|
| `executable_path` | Mole 工具可执行文件的完整路径（支持 `.appref-ms` 文件） | `"C:/Users/用户名/Desktop/Mole 2.0.appref-ms"` | ✅ 是 |
| `window_title` | Mole 工具窗口标题，用于识别窗口（支持正则表达式） | `".*MOLE.*"` | ✅ 是 |
| `login_dialog_title` | 登录对话框标题 | `"MOLE LOGIN"` | ✅ 是 |
| `timeout` | 操作超时时间（秒） | `60` | 否（默认60） |
| `retry_count` | 失败时的重试次数 | `3` | 否（默认3） |
| `retry_delay` | 重试之间的延迟时间（秒） | `2` | 否（默认2） |

### 配置示例

```yaml
mole:
  # Windows 路径使用正斜杠 / 或双反斜杠 \\
  executable_path: "C:/Users/YourName/Desktop/Mole 2.0.appref-ms"
  
  # 窗口标题支持正则表达式，匹配包含 "MOLE" 的窗口
  window_title: ".*MOLE.*"
  
  # 登录对话框标题
  login_dialog_title: "MOLE LOGIN"
  
  # 超时和重试配置
  timeout: 60
  retry_count: 3
  retry_delay: 2
```

## 需要的文件和依赖

### 1. Python 依赖包

确保已安装以下 Python 包：

```bash
pip install pywinauto pywin32 pyperclip
```

或者使用 requirements.txt：

```bash
cd workflow_automation
pip install -r requirements.txt
```

### 2. 必需文件

#### Source Lot 文件

- **位置**：`input/Source Lot.csv`（优先）或 `Auto VPO/Source Lot.csv`
- **格式**：CSV 或 Excel 文件（.xlsx, .xls）
- **必需列**：
  - `SourceLot` 或 `Source Lot`：Lot 名称
  - `Part Type`：Part Type
  - `Operation`：Operation 编号
  - `Eng ID`：Engineering ID
  - 其他可选列：`Unit test time`, `Retest rate`, `HRI / MRV:`

#### MIR Comments 文件（可选）

- **位置**：`input/MIR Comments.txt`（优先）或 `Auto VPO/MIR Comments.txt`
- **用途**：自动填写 Requestor Comments 字段
- **格式**：纯文本文件
- **示例内容**：
  ```
  Please bring this lot to 2524 level4
  ```

**文件查找顺序：**
1. `automation/auto-vpo/input/MIR Comments.txt`（优先）
2. `automation/auto-vpo/MIR Comments.txt`
3. 当前工作目录
4. `workflow_automation/` 目录

### 3. 目录结构

```
automation/auto-vpo/
├── input/
│   ├── Source Lot.csv          # Source Lot 文件（必需）
│   └── MIR Comments.txt        # MIR Comments 文件（可选）
├── output/                     # 输出目录（自动创建）
│   └── MIR_Results_*.csv       # MIR 结果文件
├── workflow_automation/
│   ├── config.yaml            # 配置文件
│   └── mole_submitter.py      # Mole 提交器代码
└── mole/
    └── README.md              # 本文档
```

## 使用方法

### 方法1：通过工作流运行（推荐）

1. **准备文件**
   - 将 `Source Lot.csv` 文件放在 `input/` 目录或 `Auto VPO/` 根目录
   - （可选）将 `MIR Comments.txt` 文件放在 `input/` 目录

2. **配置 Mole 工具路径**
   - 编辑 `workflow_automation/config.yaml`
   - 修改 `mole.executable_path` 为你的 Mole 工具路径

3. **运行工作流**
   ```bash
   # 双击运行
   3-运行工作流.bat
   
   # 或命令行运行
   cd automation/auto-vpo
   python -m workflow_automation.main
   ```

### 方法2：单独测试 Mole 功能

1. **准备测试文件**
   - 确保 `input/Source Lot.csv` 存在
   - （可选）确保 `input/MIR Comments.txt` 存在

2. **运行测试**
   ```bash
   # 双击运行
   1-测试Mole.bat
   ```

## 工作流程

### 自动化流程步骤

1. **启动 Mole 工具**
   - 检查 Mole 是否已运行
   - 如果未运行，根据配置路径启动 Mole
   - 处理登录对话框（如果出现）

2. **打开 New MIR Request**
   - 点击 File 菜单
   - 选择 New MIR Request

3. **填写 VPO 搜索**
   - 在搜索对话框中输入 SourceLot
   - 等待搜索结果

4. **选择搜索结果**
   - 检查搜索结果行状态
   - 选择符合条件的行
   - 点击 "Add to Summary" 按钮

5. **填写 Summary**
   - 切换到 "3. View Summary" 标签
   - 自动填写 Requestor Comments（如果提供了 MIR Comments.txt）

6. **提交并获取 MIR**
   - 点击 Submit 按钮
   - 处理确认对话框
   - 获取 MIR 号码

7. **保存结果**
   - 将 MIR 结果保存到 `output/MIR_Results_YYYYMMDD_HHMMSS.csv`
   - 输出文件位于：`automation/auto-vpo/output/` 目录

### 输出文件

- **MIR 结果文件**：`automation/auto-vpo/output/MIR_Results_YYYYMMDD_HHMMSS.csv`
  - **位置**：`output/` 文件夹（自动创建）
  - **内容**：包含原始 Source Lot 数据 + MIR 号码
  - **格式**：CSV 文件，UTF-8 编码（带 BOM，Excel 可直接打开）
  - **文件名格式**：`MIR_Results_YYYYMMDD_HHMMSS.csv`（包含日期和时间戳）

## 常见问题

### 1. 找不到 Mole 工具

**问题**：提示 "无法找到Mole工具"

**解决方法**：
- 检查 `config.yaml` 中的 `executable_path` 是否正确
- 确认路径使用正斜杠 `/` 或双反斜杠 `\\`
- 确认文件确实存在于指定路径

### 2. 无法连接到 Mole 窗口

**问题**：提示 "无法连接到Mole窗口"

**解决方法**：
- 检查 `window_title` 配置是否正确
- 确认 Mole 工具窗口标题包含 "MOLE"（不区分大小写）
- 如果窗口标题不同，修改 `window_title` 配置

### 3. 找不到 MIR Comments.txt 文件

**问题**：警告 "未找到MIR Comments.txt文件"

**解决方法**：
- 这是可选文件，不影响主要功能
- 如果需要自动填写 Comments，将文件放在 `input/MIR Comments.txt`
- 或者放在 `Auto VPO/` 根目录

### 4. 找不到 Source Lot 文件

**问题**：提示 "未找到Source Lot文件"

**解决方法**：
- 将 `Source Lot.csv` 文件放在以下位置之一：
  - `input/Source Lot.csv`（优先）
  - `Auto VPO/Source Lot.csv`
- 确认文件名正确（不区分大小写）

### 5. 登录对话框处理失败

**问题**：无法自动处理登录对话框

**解决方法**：
- 检查 `login_dialog_title` 配置是否正确
- 确认登录对话框标题为 "MOLE LOGIN"
- 如果标题不同，修改配置

### 6. pywinauto 未安装

**问题**：提示 "pywinauto 未安装"

**解决方法**：
```bash
pip install pywinauto pywin32 pyperclip
```

### 7. 窗口操作失败

**问题**：点击按钮或填写字段失败

**解决方法**：
- 确保 Mole 工具窗口处于可见状态
- 不要手动操作 Mole 工具窗口（自动化过程中）
- 检查 Mole 工具版本是否兼容
- 查看日志文件 `workflow_automation/logs/workflow_automation.log` 获取详细错误信息

## 日志文件

详细日志保存在：`workflow_automation/logs/workflow_automation.log`

日志包含：
- 每个步骤的执行情况
- 错误信息和堆栈跟踪
- 窗口查找和操作详情
- MIR 号码获取结果

## 技术支持

如果遇到问题：

1. 查看日志文件获取详细错误信息
2. 检查配置文件是否正确
3. 确认所有必需文件存在
4. 确认 Python 依赖包已安装

## 更新历史

- **2025-12-11**: 添加 MIR Comments.txt 文件支持，优先在 input/ 目录查找
- **2025-12-05**: 初始版本，支持基本的 MIR 提交流程

