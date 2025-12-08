# 自动化工作流工具

## 📋 功能概述

本工具实现了一个完整的自动化工作流，用于：

1. 📖 从Excel文件读取数据
2. 🔧 使用Mole工具提交MIR数据
3. 🌐 使用Spark网页提交VPO数据
4. 🌐 使用GTS网站提交最终数据
5. 💾 保存处理结果到Excel文件

## ✨ 特性

- ✅ 支持用户输入Excel文件路径或通过GUI上传文件
- ✅ 每个步骤都有错误处理和重试机制
- ✅ 自动输出文件命名：`workflow_result_<日期>.xlsx`
- ✅ 如果某个步骤失败，记录日志并通知用户
- ✅ 完整的日志记录功能

## 🚀 快速开始

### 1. 安装依赖

```bash
cd "Auto VPO/workflow_automation"
python -m pip install -r requirements.txt
```

### 2. 配置工具

编辑 `config.yaml` 文件，配置以下内容：

```yaml
# Mole工具配置
mole:
  executable_path: "C:/path/to/Mole.exe"  # Mole工具可执行文件路径
  window_title: "Mole"  # Mole工具窗口标题
  timeout: 60
  retry_count: 3
  retry_delay: 2

# Spark网页配置
spark:
  url: "https://your-spark-website.com"  # Spark网页URL
  timeout: 60
  retry_count: 3
  retry_delay: 2
  wait_after_submit: 5
  headless: false  # 是否使用无头模式

# GTS网站配置
gts:
  url: "https://your-gts-website.com"  # GTS网站URL
  timeout: 60
  retry_count: 3
  retry_delay: 2
  wait_after_submit: 5
  headless: false  # 是否使用无头模式
```

### 3. 运行工具

#### 方式1：直接运行Python脚本

```bash
python -m workflow_automation.main
```

#### 方式2：使用批处理文件（推荐）

双击 `运行工作流.bat`

## 📝 使用说明

### 输入文件格式

工具支持标准的Excel格式（.xlsx, .xls）。文件应包含需要处理的数据行。

### 工作流步骤

1. **读取Excel文件**
   - 自动读取用户指定的Excel文件
   - 验证文件格式和数据完整性

2. **提交MIR数据到Mole工具**
   - 自动启动或连接Mole工具
   - 将Excel中的数据提交到Mole工具
   - 验证提交是否成功

3. **提交VPO数据到Spark网页**
   - 打开Spark网页
   - 自动填写表单并提交数据
   - 验证提交是否成功

4. **提交最终数据到GTS网站**
   - 打开GTS网站
   - 自动填写表单并提交数据
   - 验证提交是否成功

5. **保存结果**
   - 将处理结果保存为新的Excel文件
   - 文件名格式：`workflow_result_YYYYMMDD.xlsx`
   - 如果有错误，同时保存错误日志文件

### 输出文件

- **成功结果**: `workflow_result_<日期>.xlsx` - 包含处理后的数据
- **错误日志**: `workflow_errors_<日期>.csv` - 包含错误信息（如果有）

### 日志文件

日志文件保存在 `./logs/workflow_automation.log`，包含详细的执行记录。

## ⚙️ 配置说明

### Mole工具配置

- `executable_path`: Mole工具可执行文件的完整路径。如果为null，工具会尝试在系统PATH中查找。
- `window_title`: Mole工具窗口的标题，用于识别和连接窗口。
- `timeout`: 操作超时时间（秒）。
- `retry_count`: 失败时的重试次数。
- `retry_delay`: 重试之间的延迟时间（秒）。

### Spark/GTS网页配置

- `url`: 网站的完整URL地址（必需）。
- `timeout`: 页面加载超时时间（秒）。
- `retry_count`: 失败时的重试次数。
- `retry_delay`: 重试之间的延迟时间（秒）。
- `wait_after_submit`: 提交后等待时间（秒）。
- `headless`: 是否使用无头模式（不显示浏览器窗口）。
- `implicit_wait`: Selenium隐式等待时间（秒）。
- `explicit_wait`: Selenium显式等待时间（秒）。

## 🛠️ 开发说明

### 自定义数据提交逻辑

由于Mole工具、Spark网页和GTS网站的具体界面各不相同，你可能需要根据实际情况修改以下文件中的提交逻辑：

- `mole_submitter.py` - Mole工具的数据提交逻辑
- `spark_submitter.py` - Spark网页的数据提交逻辑
- `gts_submitter.py` - GTS网站的数据提交逻辑

### 扩展功能

工具采用模块化设计，易于扩展：

- 添加新的提交器：创建新的submitter类，继承类似的模式
- 修改工作流步骤：编辑 `workflow_main.py` 中的 `WorkflowController` 类
- 自定义数据处理：修改 `data_reader.py` 中的数据处理逻辑

## ⚠️ 注意事项

1. **Chrome浏览器**: 需要安装Chrome浏览器和ChromeDriver（Selenium自动管理）。
2. **Mole工具**: 确保Mole工具可以正常启动，并且窗口标题与配置一致。
3. **网络连接**: 确保可以访问Spark和GTS网站。
4. **权限**: 确保有读写Excel文件和创建输出目录的权限。

## 🐛 故障排除

### Mole工具无法启动
- 检查 `executable_path` 配置是否正确
- 确认Mole工具是否已正确安装
- 查看日志文件了解详细错误信息

### Spark/GTS网页无法访问
- 检查URL配置是否正确
- 确认网络连接正常
- 检查是否需要VPN或特殊网络配置
- 尝试使用非无头模式查看浏览器行为

### 数据提交失败
- 检查Excel文件格式是否正确
- 查看日志文件了解具体错误
- 确认网页表单结构是否发生变化
- 检查是否需要登录或身份验证

## 📞 支持

如有问题或建议，请查看日志文件或联系开发者。

