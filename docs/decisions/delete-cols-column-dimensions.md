# openpyxl 删列老雷：delete_cols 不动 column_dimensions 与合并单元格

## 背景

2026-11-06 生产事故：吉林大学第一医院 2026 年 10 月工资表签章后，signed 输出比原始表少了
**应发工资、补缴及退款保险差额（单位）、实发合计** 三个真实数据列（实发合计是汇总/验证明细
的关键对账列），而本该删除的隐藏列"应纳税所得额"反而残留。

根因不是删除逻辑本身，而是 openpyxl 的一个底层行为差异（"老雷"）。10 月 3 日的两个补丁
（48a90d1、69c21bc）都在围绕隐藏列调整策略，都没有踩到这一层。

## 老雷：openpyxl delete_cols / insert_cols 到底动什么

openpyxl 的 `ws.delete_cols()` / `ws.insert_cols()` **只移动单元格内容和合并单元格之外的东西**：

| 内容 | delete_cols 是否处理 |
|---|---|
| 单元格的值与样式 | ✅ 随列左移 |
| 合并单元格范围（merged ranges） | ❌ 保持原字母边界，留下失效的 MergedCell |
| column_dimensions（hidden / width / customWidth） | ❌ 钉死在旧列字母上，完全不移动 |

后果：删列后，旧列字母上的 hidden 标记会"贴"到左移进来的数据列上。此时任何
**按 hidden 标记做判断**的代码（跳过隐藏列、删除隐藏列、按隐藏列估宽）读到的都是错位的信息——
数据列被当隐藏列，隐藏列被当数据列。

合并单元格的同款问题见 `_delete_cols_with_merge` 的文档字符串；本篇聚焦 column_dimensions。

## 事故复盘（机制链）

以 2026-11-03 版本的签章流水线为例（原始表 34 列：D/F/H 为隐藏空白间隔列，AE 为隐藏
"应纳税所得额"，E 为"岗位"，AH 为"实发合计"）：

1. `_remove_force_delete_columns` 删"岗位"（E 列）→ 数据整体左移；
   column_dimensions 不动 → hidden 标记错位：F（现为基本工资）▨、H（现为应发工资）▨、
   AE（现为个人所得税）▨，而应纳税所得额（AE→AD）反而丢失了标记。
2. `_remove_empty_columns` 跳过 hidden 列（69c21bc 新加的逻辑）→ 错位标记被"信任"，
   真实数据列躲过了空列清理。
3. `_delete_remaining_hidden_columns` 按 hidden 删列 → 把应发工资、补缴及退款保险差额（单位）、
   实发合计当"原始隐藏列"物理删除；应纳税所得额因标记丢失而幸存。

两步都严格按 hidden 标记办事，但标记本身已经不可信——这就是老雷的引爆方式：
**错位不会在删列那一刻报错，而是延迟到下一个读 hidden 的环节才爆炸，且炸在完全无辜的列上。**

## 决策

1. **全项目禁止直接调用 `ws.delete_cols` / `ws.insert_cols` 做列删除**。列删除唯一入口是
   `_delete_cols_with_merge`，它在 `delete_cols` 之后调用 `_shift_column_dimensions` 把
   column_dimensions 同步左移。与删除区间相交的条目必须**拆段**（左段原位保留、
   右段左移后按原属性新建条目）——只保留左段会让右段的 hidden/width 随删列静默
   丢失（分组隐藏条目中间列被空列清理先删时，右侧"01-04月应纳税所得额"因此漏删，
   见 [hidden-column-grouped-col-entry.md](hidden-column-grouped-col-entry.md) 缺陷 C）；
   完全落在删除区间内的条目元数据随列丢弃。
2. 任何**读** hidden 标记的逻辑（`_remove_empty_columns` 的隐藏列跳过、
   `_delete_remaining_hidden_columns`、`_auto_column_width`）可以信任标记，前提是
   之前的删列都走了统一入口。新代码若引入新的删列路径，必须复用 `_delete_cols_with_merge`。
3. 给 openpyxl 上游行为写死注释的位置：`_shift_column_dimensions` 的 docstring。修改删列
   相关代码前先读它，不要凭直觉假设 openpyxl"应该"会帮你移动列元数据。

## 影响范围

- `app/batch_processor.py`：`_delete_cols_with_merge`（唯一删列入口）、
  `_shift_column_dimensions`（新增）。
- 回归测试：`test/test_batch_processor.py::TestCleanupEmptyColumns::test_force_delete_keeps_data_columns_despite_hidden_flags`
  ——把 `_shift_column_dimensions` 置为 no-op 即可复现丢列，用于防止老雷回归。

## 相关文件

- `app/batch_processor.py`: 删列入口与元数据同步逻辑。
- `docs/decisions/empty-columns-removal.md`: 空列删除与签字词保护（其"合并单元格"一节
  已按本篇修正）。
- 事故样本：`excel_example/paybill/202611050008_张朦_…/signed_…细胞命运与疾病实验室….xlsx`。
