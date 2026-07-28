# 签字提示词覆写设计决策

## 背景

工资表 Excel 中的签字提示词存在大量变体，例如：

| 原始 Excel 写法 | 飞书审批节点名 |
|----------------|---------------|
| 部长审核 | 分管领导审核 |
| 分管领导签字 | 分管领导审核 |
| 部长、分管副总签字 | 分管领导审核 |
| 财务审核 | 财务审核 |
| 财务签字 | 财务审核 |
| 业务签章 | 业务审核 |

此前，每遇到一个新变体就要手动添加归一化规则（`payroll_sheet_config.json` → `text_normalization.rules`）。这导致：

- 规则不断膨胀，难以维护。
- 总会有遗漏的变体，用户遇到就报错。
- 归一化后的单元格内容仍可能包含原始提示词的痕迹（如"分管领导审核：张三"）。

最终需要一个一劳永逸的方案：**不管 Excel 原始提示词长什么样，最终单元格内容都统一为飞书审批节点名称。**

## 决策记录

### 1. 手动归一化规则（基础层）

**规则**: 在 `payroll_sheet_config.json` 中配置 `text_normalization.rules`，每条规则定义 `source → target` 映射。

```json
{
  "text_normalization": {
    "rules": [
      {"source": "部长审核", "target": "分管领导审核"},
      {"source": "分管领导签字", "target": "分管领导审核"}
    ]
  }
}
```

**优点**: 精确、可审计、用户可通过 UI 直接编辑。
**缺点**: 每有新变体就要加规则；用户不知道有哪些变体需要加。

**commit**: `c5e1bab`, `40ce9d9`

### 2. 动态归一化规则（自动派生层）

**规则**: 根据飞书审批角色名称，自动派生"签字↔审核↔签章↔审批"后缀互换的归一化规则。

实现方式（`_generate_dynamic_normalization_rules`）:

```
输入飞书节点名: "分管领导审核"
→ 提取后缀 "审核"
→ 查找同义词: 签字、签章、审批
→ 生成规则:
    "分管领导签字" → "分管领导审核"
    "分管领导签章" → "分管领导审核"
    "分管领导审批" → "分管领导审核"
```

**优点**: 不需要为每个节点的同义词手动配规则，自动覆盖所有常见后缀变体。
**缺点**: 不能解决跨词映射（如"部长"→"分管领导"），这部分仍需手动规则。

**执行顺序**: 先应用动态规则 → 再叠加手动规则（后者优先级高）。

**commit**: `e6d37f7`

### 3. 单元格覆写（最终方案）

**规则**: 在找到签名位置后，直接将单元格内容覆写为飞书审批节点名称。

**核心逻辑**（`_insert_signature_to_excel_openpyxl`, line 1479-1486）:

```python
for approver in approvers:
    role = approver.get("role")
    if role and role in positions:
        row, col = positions[role]
        payroll_ws.cell(row=row, column=col).value = role
```

**完整流程**:

```
原始 Excel: "部长审核：张三"  "分管领导签字：李四"  "财务审核：王五"
    │
    ▼ 归一化（动态规则 + 手动规则，仅供位置检测）
 Normalized: "分管领导审核：张三"  "分管领导审核：李四"  "财务审核：王五"
    │
    ▼ find_all_signature_positions()
 Positions: {分管领导审核: (5,1), 财务审核: (5,3)}
    │
    ▼ 覆写：单元格内容 = 飞书节点名
 Overwrite:  "分管领导审核"          "分管领导审核"            "财务审核"
    │
    ▼ _split_merged_for_text() + add_image()
 Final:     [分管领导审核] [签名]    [分管领导审核] [签名]    [财务审核] [签名]
```

**commit**: `3b0806a`

### 4. 三层架构的关系

```
┌─────────────────────────────────────────────────────────┐
│  Layer 3: 单元格覆写                                      │
│  将单元格内容直接设为飞书节点名，消除一切变体                 │
│  代码位置: _insert_signature_to_excel_openpyxl, line 1479  │
│                                                          │
│  ┌─────────────────────────────────────────────────────┐  │
│  │  Layer 2: 动态归一化规则                              │  │
│  │  根据飞书节点名自动派生签字/审核/签章/审批互换规则      │  │
│  │  代码位置: _generate_dynamic_normalization_rules       │  │
│  │                                                      │  │
│  │  ┌─────────────────────────────────────────────────┐ │  │
│  │  │  Layer 1: 手动归一化规则                        │ │  │
│  │  │  payroll_sheet_config.json 中手动配置的规则      │ │  │
│  │  │  UI 编辑位置: app.py sidebar → 签字提示词映射    │ │  │
│  │  │  解决跨词映射（"部长"→"分管领导"）               │ │  │
│  │  └─────────────────────────────────────────────────┘ │  │
│  └─────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────┘

Layer 1 + Layer 2 → 位置检测（把 Excel 变体映射到配置 keywords）
Layer 3           → 最终呈现（不管 Excel 原文是什么，最终内容 = 飞书节点名）
```

## 关键设计要点

### 为什么 Layer 1 和 Layer 2 不能废除？

Layer 3（覆写）只负责"位置检测后的最终呈现"，但**位置检测本身**（`find_all_signature_positions`）依赖配置中的 keywords（如"分管领导审核"）来匹配单元格。如果 Excel 中写的是"部长审核"而飞书节点名是"分管领导审核"，必须通过 Layer 1 的手动规则把"部长审核"归一化为"分管领导审核"，才能被位置检测命中。

Layer 2 覆盖同后缀变体，Layer 1 覆盖跨词映射。两者缺一不可。

### 覆写时机

覆写发生在 `find_all_signature_positions` 之后、`_split_merged_for_text` 之前。因为位置检测依赖归一化后的单元格内容，而覆写将内容改为飞书节点名后，_split_merged_for_text 读到的值已经是规范的节点名称，不影响后续签名插入。

## 影响范围

- 仅影响 `_insert_signature_to_excel_openpyxl` 函数内的执行顺序。
- 不修改配置文件结构。
- 不改动其他模块（`test_batch_processor.py` 无新增测试，因覆写逻辑简单且与已有测试覆盖的流程一致）。

## 相关文件

- `app/batch_processor.py`: 核心逻辑，`_insert_signature_to_excel_openpyxl`、`_generate_dynamic_normalization_rules`、`_apply_normalization_rules`。
- `app/payroll_sheet_config.json`: 手动归一化规则配置。
- `app/app.py`: 归一化规则 UI 编辑（sidebar）。
