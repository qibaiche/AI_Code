# PRD LOT 自动化工具

## 功能概述
- 读取 LOT 列表（支持注释与去重）
- 通过 UI 自动化驱动 SQLPathFinder 执行 `.VG2`
- 监控 `by_lot.csv`，检测文件稳定后读取
- 生成 Functional_bin × DevRevStep 的 Pareto 表
- 导出 Excel 报表并自动发送邮件
- 日志记录与错误提示

## 快速开始
1. 安装依赖
   ```bash
   python -m pip install -r requirements.txt
   ```
2. 根据环境修改 `config.yaml`
   - `paths.lots_file`：LOT 文本
   - `paths.vg2_file`：SQLPathFinder 脚本
   - `paths.output_csv`：SQLPathFinder 导出的 by_lot.csv
   - `ui.run_button_image`：保存闪电按钮截图到 `../assets/run_button.png`
   - `email`：收件人或 SMTP 配置
3. 运行
   ```bash
   python -m prd_lot_automation.main --config prd_lot_automation/config.yaml
   ```

## 日志与报表
- 日志：`../logs/prd_lot_automation.log`
- 报表：`../reports/LOT_Report_YYYY-MM-DD.xlsx`

## 常见问题
- **未找到按钮**：确认 auto_id 或截图路径；必要时调整 `ui_action` 超时
- **CSV 读取为空**：检查 SQLPathFinder 输出路径及权限
- **邮件发送失败**：核对 Outlook 登录状态或 SMTP 账号密码

