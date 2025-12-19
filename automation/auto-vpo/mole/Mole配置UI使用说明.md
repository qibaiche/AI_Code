# Mole 配置 UI 使用说明

## 📋 概述

新增的 Mole 配置 UI 允许您在运行 Mole 自动化之前，通过图形界面配置搜索方式和相关参数。

## ✨ 主要功能

### 1. 搜索方式选择

支持三种搜索方式：

- **Search By VPOs** - 按 VPO 搜索（默认）
- **Search By Units** - 按 Unit 搜索
- **Search Inactive Cage** - 搜索 Inactive Cage

### 2. 参数配置

#### 公共参数（所有搜索方式）

- **Product Code** - 产品代码（可选）
- **Source Lot** - 源批次号（可选，留空则从 Source Lot.csv 文件中读取）

#### VPOs 特定参数

- **Fuse** - Fuse 参数（可选）
- **Part Type** - 零件类型（可选）

#### Units 特定参数

- **Show available units only** - 仅显示可用的 units（复选框）
- **Unit Name** - Unit 名称（可选）

#### Inactive Cage 特定参数

- 暂无特定参数

### 3. 配置历史

- **自动保存** - 每次点击"确定"后，配置会自动保存到 `config.yaml` 的 `mole_history` 节点
- **自动加载** - 下次打开 UI 时，会自动加载上次的配置作为默认值
- **重置功能** - 点击"重置为默认值"按钮可清空所有输入

## 🚀 使用方法

### 方法 1：通过批处理文件运行（推荐）

1. 双击运行 `1-测试Mole.bat`
2. 配置 UI 会自动弹出
3. 选择搜索方式并填写参数
4. 点击"确定"开始执行
5. 点击"取消"则终止工作流

### 方法 2：通过命令行运行

```bash
cd automation/auto-vpo
python -m workflow_automation.main --mole-only
```

## 📝 配置说明

### 配置文件位置

配置保存在：`automation/auto-vpo/workflow_automation/config.yaml`

### 配置示例

```yaml
# Mole历史配置（UI自动保存）
mole_history:
  search_mode: "vpos"  # 搜索方式: vpos, units, inactive_cage
  product_code: ""  # Product Code
  source_lot: ""  # Source Lot
  fuse: ""  # Fuse (仅vpos模式)
  part_type: ""  # Part Type (仅vpos模式)
  unit_name: ""  # Unit Name (仅units模式)
  show_available_units: false  # Show available units (仅units模式)
```

## 💡 使用技巧

### 1. 留空字段

- 留空的字段将使用默认值或从 `Source Lot.csv` 文件中读取
- 例如，如果 `Source Lot` 留空，程序会从 CSV 文件的 `SourceLot` 列读取每一行的值

### 2. 覆盖文件值

- 如果在 UI 中填写了 `Source Lot`，则该值会覆盖 CSV 文件中的值
- 这对于批量处理相同的 Source Lot 很有用

### 3. 搜索方式切换

- 切换搜索方式时，UI 会自动显示/隐藏对应的参数框
- 不同搜索方式的参数会独立保存

## 🔧 工作流程

1. **启动程序** → 显示配置 UI
2. **用户配置** → 选择搜索方式和填写参数
3. **点击确定** → 保存配置并启动 Mole 工具
4. **读取文件** → 读取 Source Lot.csv
5. **循环处理** → 对每一行执行以下操作：
   - 打开 New MIR Request
   - 根据搜索方式点击对应按钮（Search By VPOs / Units / Inactive Cage）
   - 填写搜索对话框
   - 选择结果并添加到 Summary
   - 填写 Comments
   - 提交并获取 MIR 号码
6. **保存结果** → 生成 MIR 结果文件

## 📊 输出文件

- **MIR 结果文件**：`automation/auto-vpo/output/MIR_Results_YYYYMMDD_HHMMSS.csv`
- 包含每一行的 Source Lot 和对应的 MIR 号码

## ❓ 常见问题

### Q1: 配置 UI 没有弹出？

**A:** 检查以下几点：
- 确保 `tkinter` 已安装（Python 自带）
- 检查 `config.yaml` 文件是否存在
- 查看日志文件 `logs/workflow_automation.log` 获取错误信息

### Q2: 如何恢复默认配置？

**A:** 
- 方法 1：在 UI 中点击"重置为默认值"按钮
- 方法 2：手动编辑 `config.yaml`，删除或修改 `mole_history` 节点

### Q3: 配置保存在哪里？

**A:** 配置保存在 `automation/auto-vpo/workflow_automation/config.yaml` 文件的 `mole_history` 节点中

### Q4: 可以跳过配置 UI 吗？

**A:** 目前不支持跳过 UI。如果点击"取消"，工作流会终止。未来版本可能会添加命令行参数来跳过 UI。

### Q5: Search By Units 的参数如何填写？

**A:** 
- **Product Code**: 从下拉框中选择或输入产品代码
- **Source Lot**: 可选，留空则从文件读取
- **Unit Name**: 输入要搜索的 Unit 名称
- **Show available units**: 勾选此项只显示可用的 units

## 🎯 最佳实践

1. **首次使用**
   - 先使用默认的 "Search By VPOs" 模式测试
   - 确保 Source Lot.csv 文件格式正确

2. **批量处理**
   - 如果所有行使用相同的参数，可以在 UI 中填写
   - 如果每行参数不同，在 UI 中留空，让程序从文件读取

3. **配置管理**
   - 定期备份 `config.yaml` 文件
   - 为不同的项目维护不同的配置文件

4. **错误处理**
   - 查看日志文件了解详细错误信息
   - 如果某一行处理失败，程序会继续处理下一行

## 📞 技术支持

如有问题，请查看：
- 日志文件：`automation/auto-vpo/workflow_automation/logs/workflow_automation.log`
- 主 README：`automation/auto-vpo/mole/README.md`
- 配置示例：`automation/auto-vpo/mole/配置示例.yaml`

---

**版本**: 1.0  
**更新日期**: 2025-12-18  
**作者**: AI Code Assistant

