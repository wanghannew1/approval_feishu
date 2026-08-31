# 下载 / 签名输出文件防覆盖机制设计决策

## 背景

飞书审批附件下载与签名输出过程中，存在文件重名导致**静默覆盖**的风险，具体有两条链路：

1. **下载阶段**（`app/feishu_api.py → download_file`）：附件文件名来自 `Content-Disposition` 头。同一审批内可能有多张同名附件（如两张 `工资表.xlsx`），用 `"wb"` 模式直接写盘会覆盖先下载的文件，且发生在签名插入之前，无法挽回。
2. **签名输出阶段**（`app/batch_processor.py`）：`_build_output_path` 会把"文件名不规范"的表重命名为标准名（`单位+年月工资表`）。但标准名来自 **worksheet 第一行标题**，而非原始文件名。两个不同来源的附件标题若相同，就会映射到同一个标准名，后写的 `signed_*.xlsx` 会覆盖前者。

本次运行日志已实际复现该问题：`2026年6月住培学员绩效补助发放表--彩虹公司人事代理-发工资.xlsx` 与 `专科医师规范化培训学员2026年6月补助明细（彩虹-人事代理）-发工资.xlsx` 两张表标题都映射到 `signed_吉林大学第二医院-助培奖金2026年07月工资表.xlsx`，后者覆盖了前者。

**核心难点**：区分信息（如 `-发工资`、`-人事代理`）只存在于**原始文件名**中，不在 worksheet 标题里，无法从标题解析。

## 决策记录

### 1. 下载阶段：数字后缀防覆盖（`_resolve_unique_download_path`）

在 `app/feishu_api.py` 新增 `_resolve_unique_download_path(path)`，保存前先解析空闲路径：

```
signed_foo.xlsx 已存在 → signed_foo_1.xlsx
signed_foo_1.xlsx 已存在 → signed_foo_2.xlsx
... 直到找到空闲路径
```

**规则**: 目标不存在则原样返回；已存在则追加 `_N` 数字后缀。已存在文件保持不动，绝不覆盖。

**commit**: `218c627`（`fix(download): 下载附件重名时追加数字后缀避免静默覆盖`）

### 2. 签名输出阶段：三层兜底

#### Layer 1 — 标准名（文件名不规范时）

`_build_output_path` 从 worksheet 标题提取 `单位+年月`，生成 `signed_<标准名>.xlsx`。若该标准名**已存在**（磁盘上被占用），则**不返回它**，落到 Layer 2。

```python
if not _is_standard_filename(original_name):
    standard = _build_standard_name(excel_path, ws)
    if standard:
        candidate = output_path.parent / f"signed_{standard}"
        if not candidate.exists():
            return candidate          # 标准名空闲 → 用标准名
        # 标准名已被占用 → 回退原始文件名（Layer 2）
        return output_path
```

#### Layer 2 — 原始文件名（标准名撞车时）

标准名冲突时，回退到**保留原始文件名** `signed_<原始stem>.xlsx`。原始文件名带唯一后缀（`-发工资` / `-人事代理`），天然区分不同来源，无需解析后缀。

```
signed_吉林大学第二医院-助培奖金2026年07月工资表.xlsx   ← Layer 1（第一个文件）
signed_专科医师规范化培训学员2026年6月补助明细（彩虹-人事代理）-发工资.xlsx  ← Layer 2（第二个，不再覆盖）
```

#### Layer 3 — 数字后缀（最终兜底，原始文件名也撞车时）

**这是用户关心的最终兜底**。`_insert_signature_to_excel_openpyxl` 在 `wb.save()` **之前**先做 `_resolve_unique_path(actual_output)`（`app/batch_processor.py` line 1970）：

```python
# Never overwrite an existing *_signed* file: resolve a free path BEFORE writing,
# bumping a numeric suffix if the target is already taken.
save_target = _resolve_unique_path(actual_output)
if save_target != actual_output:
    logger.warning(f"[SIGN] 已存在同名文件，保存为 {save_target.name}，避免覆盖")
try:
    wb.save(str(save_target))
except PermissionError:
    ...  # 文件被占用 → 时间戳后缀另存
```

- 原始文件名也重名（同一文件被处理两次 / 两个同原始名的附件）→ `_N` 数字后缀兜底。
- 文件被占用（如 Excel 处于打开状态）→ `PermissionError` 时用时间戳后缀 `_20260831_HHMMSS` 另存。

**关键点**：解析空闲路径发生在 `wb.save()` 之前，因此**任何已存在的 `signed_*` 文件都不可能被覆盖**——不管它是标准名撞车还是原始文件名撞车。

**commit**: `2a171be`（`fix(sign): 签名输出标准名冲突时回退原文件名+数字后缀兜底避免覆盖`）

### 3. 三层架构的关系

```
┌──────────────────────────────────────────────────────────────┐
│  Layer 3: 数字后缀兜底（_resolve_unique_path，保存前解析空闲路径） │
│  任何目标已存在 → signed_foo_1 / _2 / _3 ... 绝不覆盖             │
│  文件被占用 → 时间戳后缀另存（PermissionError）                  │
│                                                               │
│  ┌─────────────────────────────────────────────────────────┐  │
│  │  Layer 2: 原始文件名（_build_output_path 标准名冲突时回退） │  │
│  │  区分信息（-发工资 / -人事代理）只在原始文件名中              │  │
│  │  signed_<原始stem>.xlsx                                  │  │
│  │                                                         │  │
│  │  ┌────────────────────────────────────────────────────┐ │  │
│  │  │  Layer 1: 标准名（_build_output_path）              │ │  │
│  │  │  从 worksheet 标题提取 单位+年月 →                  │ │  │
│  │  │  signed_吉林大学第二医院-助培奖金2026年07月工资表.xlsx │ │  │
│  │  │  标题相同 → 映射到同一标准名（冲突根源）              │ │  │
│  │  └────────────────────────────────────────────────────┘ │  │
│  └─────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘

Layer 1 → 优先标准名（更规范、易识别）
Layer 2 → 标准名撞车时回退原始文件名（保留唯一后缀）
Layer 3 → 兜底：无论哪层，只要磁盘上已存在同名，就用数字后缀保证不覆盖
```

## 关键设计要点

### 为什么区分信息从原始文件名取，而不是 worksheet 标题？

`-发工资`、`-人事代理` 等后缀**只出现在原始附件文件名中**，worksheet 第一行标题是统一的"吉林大学第二医院-助培奖金2026年07月工资表"。因此 Layer 2 直接回退到 `signed_<原始stem>.xlsx`，利用原始文件名自带的差异天然区分，不做脆弱的后缀解析。

### 为什么 Layer 3 必须在 `wb.save()` 之前解析？

如果在 `save()` 之后才发现目标已存在，第一次写入已经覆盖了旧文件，为时已晚。先 `_resolve_unique_path` 拿到空闲路径再写，才能保证"已存在的文件绝不被覆盖"。下载阶段同理。

### 下载与签名是否共享逻辑？

不共享。`download_file` 在 `app/feishu_api.py`（不能反向导入 `batch_processor`，避免循环依赖），固有各自独立的 `_resolve_unique_*_path` 辅助函数，逻辑一致但实现各自独立。

## 影响范围

- 仅影响文件落盘的目标路径解析，不影响工作表内容、签名位置、打印等既有逻辑。
- 可 `revert`：`git revert 2a171be`、`git revert 218c627` 可独立回退单个阶段。

## 相关文件

- `app/batch_processor.py`: `_build_output_path`（Layer 1/2）、`_resolve_unique_path`（Layer 3）、`_insert_signature_to_excel_openpyxl`。
- `app/feishu_api.py`: `download_file`、`_resolve_unique_download_path`。
- `test/test_batch_processor.py`: `test_build_output_path_keeps_original_name_on_standard_collision`、`test_resolve_unique_path_appends_numeric_suffix`。
- `test/test_feishu_api.py`: `test_download_file_avoids_overwrite_on_collision`。
