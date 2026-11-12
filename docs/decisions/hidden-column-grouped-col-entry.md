# 分组 `<col>` 隐藏条目导致隐藏列漏删（诊断与修复）

**状态**：已修复并测试通过（2026-11-12）。
**样本**：`excel_example/paybill/20261015jilinbank/` 吉林农村商业银行 2026.09 四表
（驾驶员会服 / 科技大厦食堂 / 宽平食堂 / 总行食堂）。
**关联 issue**：[#54](https://github.com/wanghannew1/approval_feishu/issues/54)。

## 现象

签名后 `signed_*.xlsx` 中，原始表隐藏的中间计算列（专项扣除各子项、
01-0X 月应纳税所得额、已扣个税等）没有删除，反而公开显示；「专项扣除」
表头也出现在签名表里。按 [hidden-column-merge-anchor-rescue.md](hidden-column-merge-anchor-rescue.md)
的规则，这些列应当随隐藏列一并丢弃。以驾驶员会服为例，签名表 Z..AF 列
残留（对应原始表 AA..AJ 区）。

## 根因：分组 `<col min max hidden>` 条目只挂在首列字母键下

原始表 XML 的隐藏标记不是逐列写的，而是分组条目：

```xml
<col min="25" max="25" hidden="1"/>   <!-- Y -->
<col min="27" max="36" hidden="1"/>   <!-- AA..AJ 十列，一条分组条目 -->
<col min="37" max="37" hidden="1"/>   <!-- AK -->
```

openpyxl 读入后，`min=27..max=36` **只挂在 `column_dimensions['AA']` 一个键下**。
而 `_delete_remaining_hidden_columns` 逐列按字母查：

```python
if ws.column_dimensions[col_letter].hidden:
```

对未命中字母（'AB'..'AJ'），DimensionHolder 会**自动新建空条目**
（`hidden=False`，min/max 为 None）。按生产代码原样模拟收集结果：

```
to_delete = [25, 27, 37]   # 即 Y、AA、AK —— AB..AJ 九列漏收
```

九列中间计算列根本没进删除名单。固化公式后它们已是静态数值，
`_remove_empty_columns` 的"空列/纯公式"判定也认不出，遂留在表中显示。

**连锁伤害**：删 AA(27) 时，`_shift_column_dimensions` 处理 `min=27, max=36`
的跨区间条目落入"完全落在删除区间内 → 元数据随列丢弃"分支——整条元数据
被丢，幸存的 28..35 列从此不再带 hidden，输出里完全可见。

**连带缺陷 A（救援误触发）**：AA2:AG3「专项扣除」表头锚点在隐藏的 AA 列，
但其范围 27..33 **整段都在隐藏区间 27..36 内**，按规则应属"完全落在隐藏列内
→ 随列丢弃"。正因为 28..33 没进 doomed 集合，救援误判它"延伸到存活列"，
把表头搬到可见区保留——签名表 Z2:AE3「专项扣除」即由此来。这与 H10:N10
「分管领导审核」的救援性质不同：后者 I..N 列真实可见，救援正确；本例整段
合并实际不可见，救援是错的。**规则本身不变，错的是 doomed 集合不完整导致
"范围是否整段隐藏"的判定失真。**

**连带缺陷 B（修复过程中由回归测试暴露）**：多列隐藏区右→左逐列删除时，
锚点列排在最后处理，届时其合并已被前面的删列收缩成"整段落在 doomed 内"，
旧实现按**收缩后**范围跳过救援，锚点值随锚点列一并丢失——而合并原始范围
延伸到可见列、内容实际显示的，必须保留。故救援不能逐列挂删除循环里跑，
必须在删任何列之前、对**原始合并范围**整体跑一遍。

## 修复方案

`app/batch_processor.py` 两处：

1. **doomed 集合按维度条目的 min..max 展开**，不逐列按字母查：

```python
max_col = ws.max_column
doomed = set()
for dim in ws.column_dimensions.values():
    if not dim.hidden:
        continue
    try:
        fallback = column_index_from_string(dim.index)
    except ValueError:
        fallback = None
    lo = dim.min or fallback
    hi = min(dim.max or fallback or 0, max_col)
    if lo is None or hi < lo:
        continue
    doomed.update(range(lo, hi + 1))
```

- 程序化创建的单列条目（如 `_hide_columns` 按字母设 hidden）min/max 未落定
  （实测 `min=None, max=None, index='D'`），回退到自身键名；
- `hi` 截到 `max_column`：尾部无内容的隐藏网格列不影响打印，也避免 width 类
  超长分组（实测样本存在 `min=128 max=16371`）撑爆逐列删除。

2. **救援改为删列前一次性执行**：`_rescue_merge_anchors_in_column` 改名
   `_rescue_partial_hidden_merges`，在删除循环前对全部合并范围跑一遍；
   跳过条件（锚点在存活列 / 范围整段 doomed）只对原始合并范围生效。
   后续逐列右→左删除流程不动：`_shift_column_dimensions` 的跨区间截断分支
   在单列逐步删除下恰好让剩余隐藏元数据同步收缩
   （27..36 → 27..35 → … → 删到 27 时整条丢弃），手工推演收敛正确。

### 修复后本样本的实测行为（驾驶员会服，`_delete_remaining_hidden_columns` 直跑）

- 128 → 116 列（Y、AA..AJ、AK 共 12 列全删），hidden 标记全清；
- 专项扣除 / 大病险 / 本月应纳税所得额 随隐藏列丢弃；
- 扣款合计 / 应纳个人所得税 / 实发工资 合法保留，落位 …Y｜Z｜AA，
  即表尾收敛为 …扣款合计｜应纳个人所得税｜实发工资。

## 已知同款盲区（暂不动）

`_remove_empty_columns` 的隐藏跳过守卫（`ws.column_dimensions[...].hidden: continue`）
对分组条目同样漏判。但它是"防止误删"方向的守卫，失效方向相反（多判为可见、
可能被空列清理处理），影响小；待后续如需再统一收口。

## 回归测试（已落地）

`test/test_batch_processor.py::TestCleanupEmptyColumns` 新增两条：

1. `test_delete_hidden_columns_grouped_col_entry`：分组条目整段删除 + 整段
   隐藏合并（专项扣除形态）丢弃不救援；
2. `test_delete_hidden_columns_grouped_entry_partial_rescue`：分组区内合并
   延伸到存活列时仍须救援（该条在修复救援时机前**失败**，锚点值为 None，
   即缺陷 B 的复现）。

既有 `test_delete_hidden_columns_merge_anchor_rules`（H10:N10 形态）保持通过。
全仓 175 个测试通过。

## 相关文件

- `app/batch_processor.py`: `_delete_remaining_hidden_columns`、
  `_rescue_partial_hidden_merges`（原 `_rescue_merge_anchors_in_column`）。
- 规则上游：[hidden-column-merge-anchor-rescue.md](hidden-column-merge-anchor-rescue.md)
  （去留规则不变，本修复只修正 doomed 集合构建与救援时机）。
- 元数据左移：[delete-cols-column-dimensions.md](delete-cols-column-dimensions.md)。
