# 工资表复合表头与空列删除方案

## 1 背景

工资发放表（外包/派遣）的典型结构：

| 区域 | 行 | 说明 |
|------|-----|------|
| 标题行 | 1 | 单位名称 + 年月 + 工资表名称 |
| 单位行 | 2 | 单位名称 + 填报日期 |
| 复合表头 3-5 | 3 | 大类（序号、姓名、扣款明细、个人所得税…） |
| | 4 | 小类（养老、失业、医疗、公积金…） |
| | 5 | 子类（单位、个人） |
| 数据区 | 6+ | 人员工资数据，最后几行为合计 / 税费 / 转款 / 签字区 |

行 3-5 构成**复合表头**：行 3 是大分类，行 4 是其子分类，行 5 是孙分类。
合并单元格横向跨列（"扣款明细"跨养老~扣款合计）且纵向跨行（"个人所得税"跨 3-5 行）。

## 2 核心问题

工资表原始文件包含大量空列（如 `部门`、`职工号`、`交通补贴` 在数据区全空），
需要在打印前删除。但删除列会破坏：

- **表头合并单元格**（纵向 3 行合并 / 横向跨列合并）
- **数据区合并单元格**（如 `税费(7%)` 的公式跨 D:AE 整行合并）
- **签名区合并单元格**（签字栏横向合并）
- **隐藏列状态**（原始文件可能隐藏了某些列）

## 3 整体流程

```
原始工作表
    │
    ▼
[3.1] 展平表头合并  ───────────────── 将行 1-5 所有合并单元格解除，
    │                                  每个格子里写入原合并区域的值
    ▼
[3.2] 删除空数据列  ───────────────── 从右向左扫描，识别可删除列，
    │                                  移动公式值 → 全量快照合并 → 删列 → 恢复
    ▼
[3.3] 重建表头合并  ───────────────── 根据展平后的值重新合并（横向+纵向）
    │
    ▼
[3.4] 重算隐藏列    ───────────────── 重置所有隐藏，仅隐藏目标列
    │
    ▼
输出工作表
```

---

## 3.1 展平表头合并 (`_flatten_header_merges`)

### 目标

将行 1-5 中所有合并单元格「展平」——解除合并，把原合并区域的值写入每一个单元格。
这样后续删除列时不会受制于遗留的合并范围。

### 算法

```python
for mr in ws.merged_cells.ranges:
    if mr.min_row > header_rows:        # 只处理表头区域
        continue
    if mr.min_row == mr.max_row and mr.min_col == mr.max_col:
        continue                         # 单格跳过

    anchor_value = ws.cell(mr.min_row, mr.min_col).value
    ws.unmerge_cells(str(mr))
    for r in range(mr.min_row, mr.max_row + 1):
        for c in range(mr.min_col, mr.max_col + 1):
            ws.cell(r, c).value = anchor_value
```

展平后，行 3 的 `扣款明细`（原 Q3:AA3）变成 Q3, R3, …, AA3 每个格子都等于 `扣款明细`；
行 4-5 同理。原始合并信息完全丢失——后续删除列时不会因为列位于合并范围内而报错。

### 前提

必须在**删除列之前**调用，且只需调用一次。数据区的合并（`D12:AE12`）不做展平。

### 3.1.1 格式保存

展平解除合并后，原先的 `MergedCell` 对象变成普通 `Cell`，但**丢失了原来锚点单元格的格式**
（居中对齐、边框、字体等）。后续删除列后，任意展开的单元格都可能成为新的合并锚点，
如果该单元格没有格式，重建后的合并单元格会丧失格式。

因此展平时**将锚点格式复制给所有展开的单元格**：

```python
# 展平前快照锚点格式
fmt = {
    'font': copy(src.font),
    'alignment': copy(src.alignment),
    'border': copy(src.border),
}
ws.unmerge_cells(str(mr))

for r in range(mr_min, mr_max + 1):
    for c in range(mc, mxc + 1):
        cell = ws.cell(row=r, column=c)
        if cell.value is None:
            cell.value = anchor_val
        # 恢复格式（MergedCell → Cell 后格式丢失）
        cell.font = copy(fmt['font'])
        cell.alignment = copy(fmt['alignment'])
        cell.border = copy(fmt['border'])
```

---

## 3.2 删除空数据列 (`_remove_empty_columns`)

### 3.2.1 判空规则

数据区从行 4（`DATA_START = 4`）开始扫描。
某一列在数据区所有值满足以下**任一**条件，则判定为可删除：

1. **空值**（`None` 或 `''`）
2. **签名关键词**（如 `总经理签字`、`部长签字`）
3. **公式**（字符串以 `=` 开头）

关键词和公式在删除前会被**移动到右侧最近的非空列**（或追加在末尾），避免数据丢失。

### 3.2.2 右→左扫描

```python
for col in range(ws.max_column, 0, -1):
    non_empty = {}
    for row in range(DATA_START, ws.max_row + 1):
        v = ws.cell(row=row, column=col).value
        if v not in (None, ''):
            non_empty[row] = str(v)

    if not non_empty:
        _delete_cols_with_merge(ws, col)    # 完全空列，直接删
        continue

    if not all(_is_removable(v) for v in non_empty.values()):
        continue                            # 包含不可删除的值，跳过

    # === 可删除列 (只有签名关键词/公式) ===
    target = _find_target_column(ws, col)
    snapshot = list(ws.merged_cells.ranges)  # 全量快照！

    for row, val in non_empty.items():
        val = _evaluate_formula(val)          # 公式求值
        ws.cell(row=row, column=target).value = val

    _delete_cols_with_merge(ws, col, saved_all=snapshot)
```

**注意**：
- `snapshot` 必须在任何 `unmerge_cells` **之前**获取，否则即将解冻的合并范围会丢失。
- `saved_all` 传递给 `_delete_cols_with_merge` 用于恢复。

### 3.2.3 公式求值

由于 openpyxl 默认不计算公式，需要加载两次工作簿：

```python
wb_formula = load_workbook(path)             # 保留公式文本
wb_data    = load_workbook(path, data_only=True)  # 获取计算值

formula_values = {}
for r in range(1, max_row + 1):
    for c in range(1, max_column + 1):
        raw = wb_formula.cell(r, c).value
        computed = wb_data.cell(r, c).value
        if isinstance(raw, str) and raw.startswith('=') and computed is not None:
            formula_values[(r, c)] = computed
```

在 `_remove_empty_columns` 中，遇到公式时查表替换：

```python
if _is_formula(val):
    computed = formula_values.get((row, col))
    if computed is not None:
        val = computed
```

### 3.2.4 合并单元格的删除与恢复 (`_delete_cols_with_merge`)

openpyxl 的 `delete_cols` **不会调整任何已存在的合并单元格范围**，
删除后合并范围仍指向原列字母，产生 `MergedCell` 残留。

因此需要全量快照 + 删除 + 全量重建：

```python
def _delete_cols_with_merge(ws, col, amount=1, saved_all=None):
    if saved_all is None:
        saved_all = list(ws.merged_cells.ranges)

    ws.delete_cols(col, amount)                # openpyxl 不调合并范围

    for mr in list(ws.merged_cells.ranges):    # 清除所有残留合并
        try:
            ws.unmerge_cells(str(mr))
        except Exception:
            pass

    for mr in saved_all:                        # 从快照重建
        mc, r0, mxc, rmax = mr.bounds
        if mxc < col:                           # 完全在左侧 → 不变
            new_mc, new_mxc = mc, mxc
        elif mc >= col + amount:                # 完全在右侧 → 左移
            new_mc, new_mxc = mc - amount, mxc - amount
        else:                                   # 重叠 → 收缩 + 左移
            new_mc = mc if mc < col else max(col, mc - amount)
            new_mxc = mxc if mxc < col else mxc - amount
        if new_mc <= new_mxc:
            ws.merge_cells(...)
```

#### 边界示例

| 原始范围 | 删除列 | 重建后 |
|----------|--------|--------|
| D12:AE12 (4-31) | D (4) | D12:AD12 (4-30) |
| D12:AD12 (4-30) | F (6) | D12:AC12 (4-29) |
| H14:M14 (8-13) | D (4) | H14:M14 (8-13，不变) |
| H14:M14 (8-13) | F (6) 后 D (4) | F14:J14 (6-10) |
| AD14:AE14 (30-31) | D,F,H 三次删除 | AA14:AB14 (27-28) |

### 3.2.5 格式迁移

在将公式/关键词值移动到目标列时，一并复制源单元格的格式（对齐、边框、字体、数字格式），
确保数据区合并锚点（如 D12）的格式不丢失：

```python
tgt = ws.cell(row=row, column=target)
tgt.value = val
# 源列即将被删除，提前把格式转移到目标格
tgt.font = copy(src_cell.font)
tgt.alignment = copy(src_cell.alignment)
tgt.border = copy(src_cell.border)
tgt.number_format = copy(src_cell.number_format)
```

执行时机：在 `_delete_cols_with_merge`（实际删除列）**之前**，值移动时同步复制。

---

## 3.3 重建表头合并 (`_rebuild_header_merges`)

### 目标

根据展平后（3.1）且列删除后（3.2）的单元格值，重新生成正确的表头合并单元格。

### 合并顺序（不可调换）

```
步骤 1: 行 4 水平合并  ─ 在行 3 分组内合并行 4 相邻相同值
步骤 2: 纵向合并        ─ 三个分支 (A/B/C)
步骤 3: 兜底 3 行合并   ─ 行 4+5 都空
步骤 4: 行 3 水平合并   ─ 最后一步，不影响行 4 分组
步骤 5: 行 1-2 水平合并 ─ 简单相邻合并（标题行）
```

### 步骤详解

#### 步骤 1：行 4 水平合并

```python
vi = 1
while vi <= max_col:
    r3v = ws.cell(3, vi).value
    vj = vi + 1
    while vj <= max_col and ws.cell(3, vj).value == r3v:
        vj += 1                     # 找到行 3 分组 [vi, vj)

    vk = vi
    while vk < vj:                  # 在该分组内
        r4v = ws.cell(4, vk).value
        if r4v is None or r4v == '':
            vk += 1
            continue
        vl = vk + 1
        while vl < vj and ws.cell(4, vl).value == r4v:
            vl += 1
        if vl - 1 > vk:
            ws.merge_cells(start_row=4, start_column=vk,
                           end_row=4, end_column=vl - 1)
        vk = vl
    vi = vj
```

**关键**：行 4 的合并必须在行 3 的分组内进行。
例如行 3 `扣款明细` 跨越养老~扣款合计，行 4 的 `养老` 只能在 `扣款明细` 列范围内合并。
`工伤险`（行 4 单列）在该分组内无相邻相同值，不会合并。

#### 步骤 2：纵向合并

对每一列，判断行 3-5 的值：

| 分支 | 条件 | 合并 |
|------|------|------|
| A | r3v = r4v = r5v 且都非空 | 合并行 3-5 |
| B | r3v ≠ r4v, r4v = r5v 且都非空 | 合并行 4-5 |
| C | r4v 非空, r5v 空 | 合并行 4-5 |

典型例子：
- `个人所得税`（Z3=Z4=Z5）→ 分支 A → Z3:Z5
- `单位代理费`（W4=W5，W3 不同）→ 分支 B → W4:W5
- `工伤险`（U4 非空，U5 空）→ 分支 C → U4:U5? 
  不——U5='单位' 非空，不满足分支 C。

**注意**：用 `hmerged` 集合追踪已合并的锚点，避免重复合并。

#### 步骤 3：兜底 3 行合并

对于行 4+5 都空的列（如原文件在展平+删除后，某些列只有行 3 有值），合并行 3-5：

```python
if not r4v and not r5v:
    ws.merge_cells(start_row=3, start_column=vi,
                   end_row=header_rows, end_column=vi)
```

#### 步骤 4：行 3 水平合并（必须在最后）

合并行 3 中相邻的相同值。必须在步骤 1-3 之后，因为步骤 1-3 可能改变了行 4 的值分布，且行 3 的合并跨度过大时会干扰步骤 1 的分组计算。

```python
vi = 1
while vi <= max_col:
    r3v = ws.cell(3, vi).value
    if r3v is None or r3v == '':
        vi += 1
        continue
    vj = vi + 1
    while vj <= max_col and ws.cell(3, vj).value == r3v:
        vj += 1
    if vj - 1 > vi:
        ws.merge_cells(start_row=3, start_column=vi,
                       end_row=3, end_column=vj - 1)
    vi = vj
```

例如 `扣款明细` 原本跨 Q3:AA3（原始 17-28），删除 3 列后变为 N3:Y3（14-25）。

#### 步骤 5：行 1-2 水平合并

行 1（标题）和行 2（单位+日期）的合并方式简单：从左到右扫，值相同的相邻列合并。
不涉及纵向合并。

### 行 5 不合并

行 5（`单位`、`个人`）**不做任何水平合并**。
这是因为：

- T5='单位' 和 U5='单位' 分属不同逻辑组（工伤险单位 / 公积金单位）
- 在步骤 1-4 中行 5 仅作为纵向合并的判断依据（与行 4 的值比较），不主动生成合并

---

## 3.4 隐藏列重置 (`_hide_columns`)

### 问题

原始文件可能隐藏了某些列（如 `D(部门)`、`H(交通补贴)`）。
删除列后列位置偏移，`column_dimensions[字母].hidden` 仍指向旧字母，导致：
- 应隐藏的列未被隐藏
- 不应隐藏的列被错误隐藏（如 `应发工资`、`单位代理费`）

### 方案

重置所有隐藏状态，仅根据当前行 3 的表头值重新隐藏：

```python
headers_to_hide = {"部门", "岗位", "职工号"}
for col in range(1, ws.max_column + 1):
    cell = ws.cell(row=3, column=col)
    hidden = False  # 先全部解除隐藏
    if cell.value and str(cell.value).strip() in headers_to_hide:
        hidden = True
    ws.column_dimensions[get_column_letter(col)].hidden = hidden
```

执行时机：重建表头合并（3.3）**之后**，因为重建可能改变行 3 的值位置。

---

## 4 完整调用顺序

```python
# 1. 展平表头合并
_flatten_header_merges(ws, header_rows=5)

# 2. 加载公式计算值（另一只 data_only=True 的 workbook）
formula_values = _extract_formula_values(wb_formula, wb_data)

# 3. 删除空列（自动处理合并恢复）
_remove_empty_columns(ws, cfg, formula_values)

# 4. 重建表头合并
_rebuild_header_merges(ws, header_rows=5)

# 5. 重算隐藏列
_hide_columns(ws)
```

---

## 5 验证清单

| 检查项 | 方法 |
|--------|------|
| 表头纵向合并完整 | 逐个检查 `A3:A5`、`Z3:Z5`、`AA3:AA5`… |
| 表头横向合并正确 | `扣款明细 N3:Y3` 跨列正确，不串列 |
| 行 5 无横向合并 | `T5=单位` 和 `U5=单位` 不合并 |
| 数据区合并恢复 | `D12:AB12=3836.84` 等 |
| 签名区合并移位正确 | `AD14:AE14` → `AA14:AB14` |
| 无残留垃圾合并 | 不应有 `AC3:AC5=<None>` |
| 隐藏列正确 | 仅 `岗位`、`职工号` 等 |
| 公式已求值 | `=ROUND(L11*0.07,2)` → `3836.84` |
| 列数正确 | 31 - 删除数 = 最终列数 |
| 表头单元格格式保留 | 锚点格 `h=center`, `b-left=thin` 不变 |
| 数据区单元格格式保留 | D12 锚点 `h=center`, `border=thin` 保持 |
| 105 测试全通过 | `pytest test/ -v` |

---

## 6 已知限制

1. **合并范围的偏移计算**：`_delete_cols_with_merge` 假设每次删除连续的列（`amount=1`），
   对于 `amount>1` 未测试。
2. **表头行数固定**：假设行 1-5 为表头。其他模板需调整 `header_rows`。
3. **公式求值依赖 data_only**：`data_only=True` 仅返回最后一次保存时的缓存值，
   如果原文件从未在 Excel 中打开保存过，可能返回 `None`。
4. **关键字硬编码**：`_hide_columns` 中的 `headers_to_hide` 集合为硬编码，
   需从配置中读取。
