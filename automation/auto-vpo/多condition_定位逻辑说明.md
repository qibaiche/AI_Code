# 多 Condition 定位逻辑说明

## 修改总结

根据真实页面的 UI 结构，完全重写了 `spark_submitter.py` 中 Operation 和 Eng ID 的定位逻辑。

---

## 核心改动

### 1. **新增方法：`_find_operation_headers()`**
查找所有 Operation 区块的抬头行（可编辑行）。

**定位规则：**
- 包含 2 个 `mat-select-arrow`（Operation 和 Eng ID）
- 排除 "Continue with All Units" 行
- 排除灰色只读历史行
- 只返回可见的抬头行

**返回值：**
- 抬头行元素列表：`[header_0, header_1, header_2, ...]`
- `header_0` = 第 1 个 Operation 区块
- `header_1` = 第 2 个 Operation 区块
- 以此类推...

---

### 2. **新增方法：`_select_mat_option_by_text(text)`**
在 Angular Material 的下拉面板中选择指定文本的选项。

**特点：**
- 等待 `mat-select-panel` 出现
- 精确匹配或包含匹配
- 可复用于 Operation 和 Eng ID

---

### 3. **重写方法：`_select_operation(operation, condition_index)`**
在第 `condition_index` 个 Operation 区块上选择 Operation。

**新逻辑：**
1. 调用 `_find_operation_headers()` 获取所有抬头行
2. 取第 `condition_index` 个抬头行
3. 滚动到抬头行可见
4. 在抬头行内找第 1 个 `mat-select-arrow`（Operation）
5. 点击箭头，等待下拉面板出现
6. 调用 `_select_mat_option_by_text()` 选择选项

**不再使用：**
- ❌ 全局索引（如 `(//div[contains(@class,'mat-select-arrow')])[1]`）
- ❌ `mat-form-field-type-mat-select`
- ❌ `mat-select-trigger`
- ❌ 动态 class（如 `ng-tns-c104-21`）

---

### 4. **重写方法：`_select_eng_id(eng_id, condition_index)`**
在第 `condition_index` 个 Operation 区块上选择 Eng ID。

**新逻辑：**
1. 关闭所有已打开的 overlay（避免误点）
2. 调用 `_find_operation_headers()` 获取所有抬头行
3. 取第 `condition_index` 个抬头行
4. 滚动到抬头行可见
5. 在抬头行内找第 2 个 `mat-select-arrow`（Eng ID）
6. 点击箭头，等待下拉面板出现
7. 检查是否误点了 "All Units"，如有则关闭并重新点击
8. 调用 `_select_mat_option_by_text()` 选择选项

**不再使用：**
- ❌ 全局索引
- ❌ 复杂的容器内查找（`condition-list-container`）
- ❌ 索引计算公式（`2 + 2 * condition_index`）

---

### 5. **重写方法：`_click_add_new_condition()`**
点击最后一个 Operation 区块内的 "Add new condition" 按钮。

**新逻辑：**
1. 调用 `_find_operation_headers()` 获取所有抬头行
2. 取最后一个抬头行：`last_header = operation_headers[-1]`
3. 滚动到最后一个抬头行的末尾（确保按钮可见）
4. 向上追溯找到包含该抬头行的容器
5. 在容器内查找 "Add new condition" 按钮（多种方法）：
   - 方法1：通过 ID `addNewCondition`
   - 方法2：通过 class + 文本
   - 方法3：通过文本
   - 方法4：全局查找，取最后一个可见的
6. 滚动到按钮可见并点击
7. 等待新区块渲染（2秒）
8. 验证新区块是否已添加

**不再使用：**
- ❌ 全局查找 "Add new condition"（可能点到错误的块）

---

## 工作流程（多 Condition）

```python
rows = [(op1, eng1), (op2, eng2), (op3, eng3), ...]  # 从 CSV 读取

for idx, (op_value, eng_value) in enumerate(rows):
    if idx == 0:
        # 第 1 行：使用已有的第 1 个 Operation 区块
        _select_operation(op_value, condition_index=0)
        _select_eng_id(eng_value, condition_index=0)
    else:
        # 第 2+ 行：先点击 "Add new condition"，再填值
        _click_add_new_condition()
        time.sleep(1.0)  # 等待新区块渲染
        _select_operation(op_value, condition_index=idx)
        _select_eng_id(eng_value, condition_index=idx)
```

---

## 关键优势

### 1. **更可靠**
- 不依赖动态 class（如 `ng-tns-c104-21`）
- 不使用全局索引（避免误点 "All Units"）
- 基于页面结构（抬头行列表）定位

### 2. **更简单**
- 统一的定位逻辑（所有 condition 使用同一方法）
- 无需区分第一个 condition 和后续 condition
- 代码更清晰易懂

### 3. **更灵活**
- 自动适应页面变化（只要抬头行结构不变）
- 支持任意数量的 Operation 区块
- 自动滚动到正确位置

---

## 测试建议

1. **单个 Condition：**
   - MIR CSV 中只有 1 行数据
   - 验证第 1 个 Operation 区块的填写

2. **多个 Condition（同一 SourceLot）：**
   - MIR CSV 中有多行相同 SourceLot 的数据
   - 验证循环添加新 condition 的逻辑

3. **滚动测试：**
   - MIR CSV 中有 5+ 行数据
   - 验证页面滚动和新区块定位

4. **错误处理：**
   - Operation 或 Eng ID 不存在于下拉选项中
   - 验证错误日志和跳过逻辑

---

## 注意事项

1. **等待时间：**
   - 点击 "Add new condition" 后等待 2 秒（DOM 渲染）
   - 选择 Operation 后等待 1 秒（Eng ID 前）

2. **"All Units" 误点：**
   - 在选择 Eng ID 前，先关闭所有已打开的 overlay
   - 如果误点，自动关闭并重新点击正确的箭头

3. **页面滚动：**
   - 所有点击操作前都会滚动到元素可见
   - 使用 `scrollIntoView({block: 'center'})` 确保元素在视口中央

---

## 文件修改清单

- ✅ `automation/auto-vpo/workflow_automation/spark_submitter.py`
  - 新增 `_find_operation_headers()`
  - 新增 `_select_mat_option_by_text()`
  - 重写 `_select_operation()`
  - 重写 `_select_eng_id()`
  - 重写 `_click_add_new_condition()`

- ✅ `automation/auto-vpo/test_spark.py`
  - 无需修改（已使用正确的 condition_index 参数）

---

## 下一步

1. 运行 `测试Spark.bat`
2. 观察日志输出，确认定位逻辑正确
3. 如有问题，根据日志调整

