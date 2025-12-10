# Spark 自动化工具使用说明

## 📋 简介

这是一个用于自动填写 Spark 网页表单的工具，可以自动从 MIR 结果文件读取数据并填写到 Spark 网站。

## 🚀 快速开始

### 1️⃣ 运行前准备

**必须配置：**
- ✅ 确认 **TP路径** 正确
- ✅ 确认 **MIR结果文件** 存在

**配置文件说明：**

有两个配置文件可以使用（推荐使用方法2）：

**方法1：修改主配置文件** ⚙️
```
automation/auto-vpo/workflow_automation/config.yaml
```
- 适合：需要统一管理所有配置
- 优点：所有配置在一个文件中

**方法2：修改 Spark 专用配置** ⭐ **推荐**
```
automation/auto-vpo/spark/spark_automation_config.yaml
```
- 适合：只运行 Spark 自动化
- 优点：配置文件在同一目录，更方便
- **优先级更高**：会覆盖主配置文件的设置

### 2️⃣ 修改 TP 路径

**推荐：修改 Spark 专用配置文件** ⭐

打开 `spark/spark_automation_config.yaml`，找到 TP 路径配置：

```yaml
# Test Program路径
test_program:
  tp_path: "\\\\gar.corp.intel.com\\ec\\proj\\...\\PTL_P_CLASS"
```

**修改方法：**
1. 复制你的 TP 网络路径
2. 将路径中的 `\` 改为 `\\` （双反斜杠）
3. 或者使用正斜杠 `/` 代替反斜杠

**示例：**
```yaml
# 原始路径：
\\gar.corp.intel.com\ec\proj\test

# 修改为（任选一种）：
test_program:
  tp_path: "\\\\gar.corp.intel.com\\ec\\proj\\test"  # 双反斜杠
  tp_path: "//gar.corp.intel.com/ec/proj/test"      # 正斜杠
```

**也可以修改其他参数：**
```yaml
spark:
  vpo_category: "engineering"  # 修改 VPO 类别
  step: "B5"                   # 修改 Step
  tags: "ECG_24J-DOD"          # 修改 Tags
```

### 3️⃣ 准备 MIR 结果文件

确保以下位置存在 `MIR_Results_*.csv` 文件：
- `automation/auto-vpo/spark/` （当前目录）
- `automation/auto-vpo/` （父目录）
- `automation/auto-vpo/mole/` （mole目录）

**文件格式：**
MIR结果CSV文件必须包含以下列：
- `SourceLot`: Lot名称
- `Part Type`: Part Type
- `Operation`: Operation编号
- `Eng ID`: Engineering ID
- `Unit test time`: 单位测试时间（可选）
- `Retest rate`: 重测率（可选）
- `HRI / MRV:`: HRI/MRV值（可选）

### 4️⃣ 运行工具

**方法1：双击批处理文件**
```
双击：测试Spark.bat
```

**方法2：命令行运行**
```bash
cd automation/auto-vpo/spark
python test_spark.py
```

## 📝 详细配置说明

### 配置文件优先级

```
spark_automation_config.yaml (优先级高)
         ↓ 覆盖
workflow_automation/config.yaml (基础配置)
```

### spark_automation_config.yaml 配置项 ⭐ 推荐

```yaml
# Test Program路径
test_program:
  tp_path: "你的TP路径"  # ⚠️ 必须修改

# Spark配置
spark:
  vpo_category: "engineering"  # VPO类别
  step: "B5"                   # Step选项
  tags: "ECG_24J-DOD"          # Tags标签
```

### 常见配置修改

| 配置项 | 位置 | 说明 | 示例值 | 修改频率 |
|--------|------|------|--------|----------|
| `tp_path` | test_program.tp_path | TP路径 | `\\\\server\\path\\to\\tp` | ⚠️ 每次 |
| `vpo_category` | spark.vpo_category | VPO类别 | `correlation`, `engineering` | 偶尔 |
| `step` | spark.step | Step选项 | `B0`, `B4`, `B5`, `B6` | 偶尔 |
| `tags` | spark.tags | Tags标签 | `CCG_24J-TEST`, `ECG_24J-DOD` | 偶尔 |

## 🔍 运行流程

工具会自动执行以下步骤：

1. ✅ 打开 Spark 网页
2. ✅ 点击 Add New
3. ✅ 填写 TP 路径并点击 Apply
4. ✅ 等待 loading 完成并点击 Continue
5. ✅ 点击 Add New Experiment
6. ✅ 选择 VPO 类别
7. ✅ 填写实验信息（Step、Tags）
8. ✅ 添加 Lot name
9. ✅ 选择 Part Type
10. ✅ 切换到 Flow 标签
11. ✅ 选择 Operation
12. ✅ 选择 Eng ID
13. ✅ 切换到 More options 标签
14. ✅ 填写 More options 字段

## ❓ 常见问题

### Q1: 找不到 MIR 结果文件？
**A:** 确保文件名为 `MIR_Results_*.csv` 格式，放在以下任一位置：
- `automation/auto-vpo/spark/`
- `automation/auto-vpo/`
- `automation/auto-vpo/mole/`

### Q2: TP 路径填写失败？
**A:** 检查路径格式：
- 使用双反斜杠 `\\` 或正斜杠 `/`
- 路径要用引号包围
- 确保网络路径可以访问

### Q3: 如何修改 Step 或 Tags？
**A:** 推荐在 `spark/spark_automation_config.yaml` 中修改：
```yaml
spark:
  step: "B5"
  tags: "ECG_24J-DOD"
```

### Q4: 两个配置文件有什么区别？
**A:** 
- `spark/spark_automation_config.yaml` - Spark 专用，优先级更高 ⭐
- `workflow_automation/config.yaml` - 基础配置，适用于所有模块

**推荐**：只修改 `spark/spark_automation_config.yaml`

### Q5: 运行中途出错怎么办？
**A:** 
1. 查看控制台的错误信息
2. 检查网页是否正常打开
3. 确认配置文件中的值是否正确
4. 按 Enter 键关闭浏览器，修改配置后重试

### Q6: 想看更详细的运行日志？
**A:** 日志文件保存在：
```
automation/auto-vpo/workflow_automation/logs/workflow_automation.log
```

## 📂 文件结构

```
automation/auto-vpo/
├── spark/
│   ├── test_spark.py                    # 主测试脚本（运行这个）
│   ├── 测试Spark.bat                     # 批处理文件（双击运行）
│   ├── README.md                        # 使用说明（本文件）
│   ├── spark_automation_config.yaml ⭐  # Spark配置（⚠️ 推荐修改这个）
│   └── MIR_Results_*.csv                # MIR结果文件（需要准备）
├── workflow_automation/
│   ├── config.yaml                      # 基础配置文件
│   ├── spark_submitter.py               # Spark自动化核心代码
│   └── config_loader.py                 # 配置加载器
└── mole/
    └── MIR_Results_*.csv                # 或者放这里
```

## 💡 提示

1. **第一次使用：** 修改 `spark/spark_automation_config.yaml` 中的 TP 路径 ⭐
2. **每次运行前：** 确认 MIR 结果文件存在且是最新的
3. **修改配置后：** 保存文件后再运行
4. **运行过程中：** 不要关闭浏览器窗口
5. **出错后：** 查看错误信息，修改配置后重试
6. **配置优先级：** `spark_automation_config.yaml` > `workflow_automation/config.yaml`

## 📞 需要帮助？

如果遇到问题，请检查：
1. ✅ TP 路径格式是否正确（双反斜杠）
2. ✅ MIR 结果文件是否存在
3. ✅ spark_automation_config.yaml 格式是否正确（注意缩进）
4. ✅ 网络连接是否正常
5. ✅ 浏览器是否能正常打开 Spark 网站

---

**祝使用愉快！** 🎉

