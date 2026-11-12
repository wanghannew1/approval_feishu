# 隐藏列删除与合并锚点救援设计决策

## 背景

2026-11-06 生产事故（GitHub issue #53 关联回归）：签章后的 signed 表缺少
"分管领导审核："槽位，签名图无处插入。

真实数据结构是理解本决策的关键：**签字提示词不是普通单元格，而是合并单元格的
锚点**——9 个工资表均为合并 H10:N10、锚点 H10，而锚点列 H 是隐藏列。Excel 对
合并单元格显示锚点内容且横跨整个合并区域，所以这段文字虽然"住"在隐藏列里，
实际是显示出来的。隐藏列里藏着的不是隐藏内容。

## 核心规则：隐藏列内容去留只看"是否实际显示"

| 内容形态 | 实例 | 是否显示 | 处理 |
|---|---|---|---|
| 合并锚点在隐藏列，范围延伸到存活列 | `分管领导审核：` @ H10:N10 | 是（横跨可见区） | **保留**：搬锚点、收缩合并 |
| 合并范围完全落在隐藏列内 | `应纳税所得额` @ AE3:AE5 | 否 | 随列丢弃 |
| 隐藏列普通格（无合并） | 隐藏中间计算值 | 否 | 随列丢弃 |

签字提示词全部落在第一行——因此**不需要**按关键词做特判（本决策的 v1 版本曾
加过"签名词兜底救援"，因与该规则冗余已移除，判定路径只有合并锚点一条）。

## 处理流程

入口：`_delete_remaining_hidden_columns`（`app/batch_processor.py`），删除前
调用 `_rescue_merge_anchors_in_column` 完成救援。

1. 收集全部隐藏列为 `doomed` 集合，**从右往左**逐列处理（避免删列使左侧列
   索引偏移）；
2. 对当前隐藏列 `col`，遍历合并区域：
   - 合并锚点列 ≠ `col` → 跳过（锚点在可见列的合并不受影响）；
   - 范围整体 ⊆ `doomed` → 跳过（内容不显示，随列一并丢弃）；
   - 否则（范围延伸到存活列）：
     a. 记录锚点值与四项格式（font / alignment / border / number_format）；
     b. `unmerge` 原范围；
     c. 把值与格式写入范围内**第一个存活列** `c_s`；
     d. 重建合并 `(r0, c_s)..(rmax, mxc)`；
3. `_delete_cols_with_merge` 删除该列：单元格值整体左移，合并范围同步左移
   收缩——步骤 2c 写入的值恰好落回**收缩后合并的锚点位**，显示布局与原表
   完全一致。真实输出示例：原 `H10:N10` 经三个隐藏间隔列删除后收缩为
   `E10:J10`，锚点 E10；
4. 全部删除后重置所有列 `hidden=False`，交由 `adjust_excel_for_print` 中的
   `_hide_columns` 按表头名（部门/岗位/职工号）重新设置。

## 与其他清理步骤的顺序

隐藏列删除位于 `_resolve_formulas_via_engine`（全表公式 → 静态数值固化）
**之后**、`_remove_empty_columns`（空列清理）**之前**（2026-11-12 调整，
缘由见 [hidden-column-grouped-col-entry.md](hidden-column-grouped-col-entry.md)
缺陷 C）：固化后全表不存在引用关系，删除被可见列引用的隐藏中间列
（应纳税所得额 → 个税链）不会打断派生数值；先删隐藏列则空列清理不会再
从分组 `<col>` 隐藏条目中间拆列，搬迁签字词时也不会落进隐藏列。

边界：固化引擎不可用时（既无 WPS COM 也无 soffice，返回 0），公式保持字符串
原样，此时删隐藏列会把可见公式打成 `#REF!`。生产机装有 WPS 不触发；如遇
纯 Linux 无 LibreOffice 环境需注意。

## 影响范围

- 仅影响签章流水线中的隐藏列删除，可见列的合并与内容不受影响。
- 前置约束：列删除必须经由 `_delete_cols_with_merge`（hidden 标记随数据左移），
  见 [delete-cols-column-dimensions.md](delete-cols-column-dimensions.md)。

## 相关文件

- `app/batch_processor.py`: `_rescue_merge_anchors_in_column`、
  `_delete_remaining_hidden_columns`。
- **已知缺陷（2026-11-12）**：原始表以分组 `<col min max hidden>` 条目标记隐藏时
  （如 `<col min="27" max="36" hidden="1"/>`），openpyxl 只在首列字母键下可见，
  逐列按字母查 hidden 会漏收区间内其余列，doomed 集合不完整，进而误触发本救援
  （实例：AA2:AG3「专项扣除」整段在隐藏区内却被救援保留）。修复方案见
  [hidden-column-grouped-col-entry.md](hidden-column-grouped-col-entry.md)。
- 回归测试：`test/test_batch_processor.py::TestCleanupEmptyColumns::test_delete_hidden_columns_merge_anchor_rules`
  （覆盖"可见合并保留"与"全隐藏丢弃"两种情形）。
- 事故记录：GitHub issue #53（含 8940c06 补充说明）。
