# Excel 打印 ### 溢出处理指南

## 问题现象

工资表在打印（或导出 PDF）时，部分数字单元格显示为 `####`，无法看到实际数值。

```
┌────────┬────────┬────────┐
│ 姓名   │ 基本工资 │ 扣款   │
├────────┼────────┼────────┤
│ 张三   │ ####   │ ####   │  ← 数字被 #### 替代
│ 李四   │ 3500   │ 412.50 │
└────────┴────────┴────────┘
```

---

## 原因分析

### 直接原因

单元格的 **列宽不足以容纳渲染后的数字宽度**。Excel/WPS 对纯数字单元格使用 `####` 作为溢出指示，而不是像文本一样跨列显示。

### 深层原因

| 因素 | 说明 |
|------|------|
| 字号过大 | 统一字号放大后，数字渲染宽度超出原列宽 |
| 公式无缓存值 | openpyxl 以 `data_only=True` 读文件时，若文件从未被 Excel/WPS 打开计算，公式单元格没有**缓存值**，代码无法得知实际数字有多长 |
| 只按表头估算 | 若表头文字较短（如"扣款"2 字），按表头算出的列宽不足以容纳下方 `12345.67` |
| 缩放冲突 | `fitToWidth=1` 将所有列压缩到一页宽，若列太多不得不压到很小字号 |

---

## 解决方案

本项目的 `batch_processor.py` 通过 **4 个机制** 协同解决 `####` 问题，在 `_auto_column_width()` 函数内按以下顺序执行：

### 步骤 1：自适应列宽（表头 + 合计行）

```python
# 只扫描表头（row 1-3）和合计行，不再遍历全部数据行
for row in range(1, min(4, ws.max_row + 1)):
    # ... 取表头最宽单元格

if total_row > 0:
    # 合计行通常是数字最大的行，取该行值估算列宽
    cell = ws.cell(row=total_row, column=col)
```

- **为什么只扫表头+合计行？** 合计行的 `=SUM(D5:D17)` 或汇总数字一般是全列最大值，用它算宽度可以覆盖所有数据。
- **公式回退**：如果公式（如 `=SUM(D5:D17)`）没有缓存值，直接用**公式文本**（如 `=SUM(D5:D17)`）估算宽度。公式文本一定比实际结果长，安全但略宽。

### 步骤 2：动态统一字号

```python
data_font_size = _calc_data_font_size(ws, col_widths)
```

根据各列「列宽 ÷ 内容宽度」的最小比值确定字号：

| 比值 | 字号 |
|------|------|
| ≥ 2.0 | 16pt（宽松） |
| ≥ 1.4 | 14pt（适中） |
| < 1.4 | 11pt（保守） |

避免字号过大导致内容溢出。

### 步骤 3：统一点阵字号

将数据区（row 4 起）所有单元格的字号统一设置为步骤 2 确定的字号，保证打印一致性。

### 步骤 4：字号放大后复查列宽

```python
if data_font_size > 11:
    _SCALE = data_font_size / 11.0
    for col ...:
        # 数字列不够宽 → 等比放大列宽
        if has_numeric and needed > cur_w:
            ws.column_dimensions[col_letter].width = round(needed + 0.5)

        # 文本列超长 → 开启自动换行，不撑爆列宽
        if has_text_overflow:
            cell.alignment = Alignment(wrap_text=True)
```

步骤 1 按 11pt 基准算宽度，步骤 4 在字号放大后复查：
- **数字列**：按 `实际字号 / 11pt` 比例放大列宽
- **文本列**：开启自动换行，避免单个长文本撑爆整列

### 全局打印设置

```python
ws.page_setup.fitToWidth = 1      # 所有列缩放到 1 页宽
ws.page_setup.fitToHeight = 0     # 不限页高
ws.page_setup.orientation = "landscape"  # A4 横向
```

`fitToWidth=1` 保证无论多少列都不会被打到多页宽。

---

## 公式缓存值机制

### 问题

openpyxl 加载文件时，公式单元格的值有两种读取方式：

```python
wb = load_workbook("file.xlsx")           # 读取公式文本：=SUM(D5:D17)
wb_data = load_workbook("file.xlsx", data_only=True)  # 读取缓存的计算结果
```

如果文件从未在 Excel/WPS 中打开计算过，`data_only` 模式返回 `None`——**没有缓存值**。

### 策略

```
公式单元格
  ├─ data_only 有缓存值 → 用缓存值估算宽度（最准确）
  └─ 无缓存值 → 用公式文本估算宽度（偏宽但安全）
```

### 谁构造 formula_values

```python
# 在 _insert_signature_to_excel_openpyxl 中
wb = load_workbook(src)
wb_data = load_workbook(src, data_only=True)

formula_values = {}
for r in range(1, ws.max_row + 1):
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=r, column=c).value
        if isinstance(v, str) and v.startswith("="):
            cv = wb_data.active.cell(row=r, column=c).value
            if cv is not None:
                formula_values[(r, c)] = cv
```

---

## 函数调用链

```
_insert_signature_to_excel_openpyxl
  └─ adjust_excel_for_print(ws, cfg, formula_values)
       ├─ ws.page_setup paperSize / orientation / fitToWidth / margins
       ├─ _hide_columns(ws)
       └─ _auto_column_width(ws, cfg, formula_values)
            ├─ 步骤 1: 表头 + 合计行 → 基础列宽
            ├─ 步骤 2: _calc_data_font_size(ws, col_widths, formula_values) → 字号
            ├─ 步骤 3: 统一点阵字号
            └─ 步骤 4: 字号复查 + 数字列加宽 / 文本列换行
```

---

## 手动排查清单

如果某个工资表打印仍有 `####`，按以下顺序排查：

1. **确认合计行是否有数据**
   - 代码自动查找包含"合计"/"总计"的行。如果工资表没有合计行，列宽估算精度下降
   - 可临时在末行加一个合计行

2. **检查字号**
   - 打开 signed 输出的 xlsx，看数据区字号是否过大（>14pt）
   - 若字号太大，可调整 `payroll_sheet_config.json` 中的列数（更多列 → `fitToWidth=1` 压缩更狠 → 需要更小字号）

3. **检查公式缓存**
   - 用 Excel/WPS 打开原文件，Ctrl+S 保存（此时公式被计算并缓存）
   - 再让程序处理该文件

4. **调整列宽上限**
   ```python
   # _auto_column_width(ws, cfg, ..., max_width=20)
   # 若某列需要更宽（如身份证号 18 位），加大 max_width
   _auto_column_width(ws, cfg, max_width=25)
   ```

5. **关闭 fitToWidth**
   （不推荐，会打多页宽）
   ```python
   ws.page_setup.fitToWidth = 0
   ```

---

## 关键代码片段

### 列宽估算

```python
def _estimate_col_width(cell_value) -> float:
    """估算单元格内容所需列宽。中文/全角字符计 2，其他计 1。"""
    text = str(cell_value)
    width = 0
    for ch in text:
        if '\u4e00' <= ch <= '\u9fff' or '\uff00' <= ch <= '\uffef':
            width += 2
        else:
            width += 1
    return width
```

### 合计行查找

```python
def _find_total_row(ws) -> int:
    keywords = ["合计", "总计", "合计金额", "合计费用", "合计支付"]
    for row in range(ws.max_row, 0, -1):  # 从下往上找
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                val = str(cell.value).replace(" ", "").replace("\u3000", "")
                for kw in keywords:
                    if kw in val and len(val) <= len(kw) + 4:
                        return row
    return 0
```

### 字号映射

```python
def _calc_data_font_size(ws, col_widths, formula_values=None):
    min_ratio = float("inf")
    # 对每个可见列，ratio = 列宽 ÷ 最宽内容宽度
    # ratio ≥ 2.0 → 16pt
    # ratio ≥ 1.4 → 14pt
    # else       → 11pt
```

---

## 修订记录

| 日期 | 变更 |
|------|------|
| 2026-07-28 | 初稿，记录 #### 溢出的 4 层防护机制 |
