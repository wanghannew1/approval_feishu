# 去掉空列 + 制表人居右

## TL;DR

> **Quick Summary**: 在 signed Excel 输出时，删除完全无数据的列（仅含签字关键词的列视为空列，关键词右移保存），并将"制表人"单元格设为右对齐。
>
> **Deliverables**:
> - `batch_processor.py`: 新增 `_remove_empty_columns()` 函数 + 修改 `_insert_signature_to_excel_openpyxl` 集成调用
> - `batch_processor.py`: 制表人单元格增加 `Alignment(horizontal='right')`
> - `test_batch_processor.py`: 新增 5-6 个测试用例
>
> **Estimated Effort**: Short
> **Parallel Execution**: NO - sequential (single file, single function chain)
> **Critical Path**: 修改 import → 实现 _remove_empty_columns → 集成调用 → 制表人对齐 → 测试

---

## Context

### Original Request
用户要求在 signed_Excel 输出时: (1) 制表人居右, (2) 去掉空列, 空列中的签字提示词右移保存

### Interview Summary
**Key Decisions**:
- 空列定义: 数据全为空就算空列 (仅含签字提示词的列也算, 删除前右移保存签字词)
- 关键词落点: 找右侧最近的非空列写入 (不覆盖数据, 写在该行目标列)
- 合并单元格: 删除列时让 openpyxl 自动调整合并范围
- 测试策略: Tests-after

### Metis Review
**Critical Issues Found**:
- 原有 "copy to col+1" 算法可能覆盖数据 → 改为"找下一个非空列写入"
- `_auto_column_width` 中有硬编码关键词列表 (line 499) → 需统一提取为 `_get_signature_keywords(cfg)` 辅助函数
- 空列定义与关键词保护需用户确认 → 已确认

---

## Work Objectives

### Core Objective
在 `_insert_signature_to_excel_openpyxl` 中添加两个功能: 制表人单元格右对齐, 删除数据空列(含关键词保护)

### Concrete Deliverables
1. `_get_signature_keywords(cfg)` 辅助函数 (供新函数和 `_auto_column_width` 共用)
2. `_remove_empty_columns(ws, cfg)` 函数
3. 集成到主流程 (normalize 后、find_positions 前)
4. 制表人单元格 `Alignment(horizontal='right', vertical='center')`
5. 测试用例

### Definition of Done
- [ ] `python -m pytest test/ -v` → 全部通过 (原85 + 新增)
- [ ] `_remove_empty_columns` 能删除无数据列
- [ ] 签字关键词在删除列后不丢失
- [ ] 制表人单元格对齐为 right

### Must Have
- 数据列不会被覆盖或删除
- 签字关键词在列删除后保留
- 包含合并单元格的列能安全删除 (openpyxl 自动调整)
- 硬编码关键词列表改为配置文件驱动

### Must NOT Have (Guardrails)
- 不修改配置文件结构
- 不影响非工资表 sheet
- 不改动其他模块

---

## Verification Strategy (MANDATORY)

### Test Decision
- **Infrastructure exists**: YES (pytest, 85 tests)
- **Automated tests**: Tests-after
- **Framework**: pytest

### QA Policy
所有验证通过 pytest 执行，无需人工操作。

---

## Execution Strategy

### Execution Order (Sequential - single file)

```
1. import 添加 Alignment
2. 新增 _get_signature_keywords(cfg) 辅助函数
3. 新增 _remove_empty_columns(ws, cfg) 函数
4. _auto_column_width 中硬编码关键词改为调用 _get_signature_keywords
5. _insert_signature_to_excel_openpyxl: 集成调用 + 制表人对齐
6. 新增测试用例
7. `pytest test/` 全量通过
```

---

## TODOs

- [x] 1. 添加 Alignment 导入 + 新增 _get_signature_keywords 辅助函数

  **What to do**:
  - 在 `from openpyxl.styles import Font` 中添加 `Alignment`: `from openpyxl.styles import Alignment, Font`
  - 新增函数 `_get_signature_keywords(cfg: dict) -> set[str]`:
    - 从 `cfg["sheet_filter"]["signatures"]["mandatory"]` 提取所有 key
    - 从 `cfg["sheet_filter"]["signatures"]["optional"]` 提取所有 key
    - 从 `cfg["text_normalization"]["rules"]` 提取所有 `source` 字段
    - 返回去重的集合
  - 将此函数放在 `get_payroll_config()` 附近

  **Recommended Agent Profile**:
  > - **Category**: `unspecified-high`
  >   - Reason: Requires careful refactoring of existing keyword extraction into a shared helper, plus import changes.
  > - **Skills**: `[]` (no special skills needed)

  **Parallelization**:
  > - **Can Run In Parallel**: NO (sequential — this is task 1 of 6, must complete before downstream tasks)
  > - **Wave**: Wave 1

  **Must NOT do**:
  - 不从 `_auto_column_width` 现有硬编码拷贝关键词 (用配置文件驱动)

  **Acceptance Criteria**:
  - [ ] `_get_signature_keywords(cfg)` 返回包含 ["总经理签字", "分管领导审核", "财务审核", "业务审核", "部长签字", "部长、分管副总签字"] 的集合
  - [ ] 空 cfg 返回空集合

  **QA Scenarios**:
  ```
  Scenario: 关键词提取完整
    Tool: pytest
    Steps: 调用 _get_signature_keywords(get_payroll_config())
    Expected: 返回集合包含 总经理签字/分管领导审核/财务审核/业务审核/部长签字/部长、分管副总签字
    Evidence: .omo/evidence/task-01-keywords.txt

  Scenario: 空 cfg 不崩溃
    Tool: pytest
    Steps: _get_signature_keywords({}) → 空集合
    Expected: 返回 set(), 不抛异常
    Evidence: .omo/evidence/task-01-keywords-empty.txt
  ```

- [x] 2. 实现 _remove_empty_columns 函数

  **What to do**:
  - 新增函数 `_remove_empty_columns(ws, cfg)`:
    1. 获取签名关键词集合: `keywords = _get_signature_keywords(cfg)`
    2. 从右到左扫描列 (`for col in range(ws.max_column, 0, -1)`):
       a. 收集该列所有非 None 的单元格
       b. 如果没有任何非 None 单元格 → 该列完全空, `delete_cols(col, 1)`
       c. 如果有非 None 单元格:
          - 检查每个非 None 值是否匹配关键词 (含: 后缀的也匹配)
          - 如果所有非 None 值都是关键词 → 视为数据空列:
            - 找到列右侧最近的非数据空列 (从 col+1 开始向右扫描)
            - 将关键词值写入目标列的同一行
            - `delete_cols(col, 1)`
          - 如果有非关键词数据 → 跳过(保留列)
    3. 记录日志: 删除了哪些列, 移动了哪些关键词

  **Recommended Agent Profile**:
  > - **Category**: `deep`
  >   - Reason: Core logic function — requires careful handling of edge cases (keyword scanning, right-to-left deletion, merged cells, column append). Higher risk if wrong.
  > - **Skills**: `[]`

  **Parallelization**:
  > - **Can Run In Parallel**: NO (depends on `_get_signature_keywords` from task 1, needed by task 3)
  > - **Wave**: Wave 1
  > - **Blocked By**: Task 1

  **Must NOT do**:
  - 不覆盖非空列的数据单元格
  - 不删除包含真实数据的列
  - 不修改非 payroll sheet

  **Acceptance Criteria**:
  - [ ] 完全无数据列被删除
  - [ ] 仅含关键词的列被删除, 关键词保存在右侧非空列
  - [ ] 含真实数据的列不受影响
  - [ ] 合并单元格列可删除

  **QA Scenarios**:
  ```
  Scenario: 删除完全空列
    Tool: pytest
    Steps: 创建 3 列(A有数据, B空, C有数据) → _remove_empty_columns
    Expected: max_column=2, B原值移到A旁边
    Evidence: .omo/evidence/task-02-delete-empty.txt

  Scenario: 关键词右移到非空列
    Tool: pytest
    Steps: 创建 A(数据), B(空, B2=总经理签字), C(数据) → _remove_empty_columns
    Expected: 列B删除, 关键词出现在C2
    Evidence: .omo/evidence/task-02-keyword-preserved.txt

  Scenario: 关键词找不到右侧非空列时追加到列尾
    Tool: pytest
    Steps: 创建 A(数据), B(空, B2=分管领导审核), 右侧无数据 → _remove_empty_columns
    Expected: 关键词在最后一列(C2)找到
    Evidence: .omo/evidence/task-02-keyword-append.txt

  Scenario: 合并单元格列安全删除
    Tool: pytest
    Steps: 创建 A1:B2 合并, C空 → _remove_empty_columns
    Expected: 列C删除, 合并单元格完整
    Evidence: .omo/evidence/task-02-merged.txt
  ```

- [x] 3. _auto_column_width 硬编码关键词改为调用 _get_signature_keywords

  **What to do**:
  - 在 `_auto_column_width` 函数中 (line 499):
    - 当前: `sig_keywords = {"总经理签字", "部长签字", "财务审核", "业务审核", "部长、分管副总签字", "分管副总签字"}`
    - 改为: 获取 cfg 参数, 调用 `_get_signature_keywords(cfg)`
  - 需要给 `_auto_column_width` 添加 `cfg` 参数 (可选, 默认 `get_payroll_config()`)
  - 更新 `adjust_excel_for_print` 中调用 `_auto_column_width(ws)` → `_auto_column_width(ws, cfg)`

  **Recommended Agent Profile**:
  > - **Category**: `unspecified-high`
  >   - Reason: Refactoring existing code — requires tracing the call chain and ensuring no behavioral change.
  > - **Skills**: `[]`

  **Parallelization**:
  > - **Can Run In Parallel**: NO (depends on `_get_signature_keywords` and `_remove_empty_columns` from tasks 1-2)
  > - **Wave**: Wave 1
  > - **Blocked By**: Tasks 1, 2

  **Must NOT do**:
  - 不改动其他函数

  **Acceptance Criteria**:
  - [ ] `_auto_column_width` 不再包含硬编码关键词
  - [ ] 所有测试通过

- [x] 4. 集成到主流程 + 制表人对齐

  **What to do**:
  - 在 `_insert_signature_to_excel_openpyxl` 中, 在 normalize 循环后、`find_all_signature_positions` 前:
    ```python
    # 删除数据空列 (含签字关键词保护)
    _remove_empty_columns(payroll_ws, cfg)
    ```
  - 在制表人字体设置块 (line 755-762), 添加:
    ```python
    cell.alignment = Alignment(horizontal='right', vertical='center')
    ```

  **Recommended Agent Profile**:
  > - **Category**: `unspecified-high`
  >   - Reason: Integration work — wiring the new function into the existing flow, modifying the 制表人 block.
  > - **Skills**: `[]`

  **Parallelization**:
  > - **Can Run In Parallel**: NO (depends on `_remove_empty_columns` and `_get_signature_keywords` from tasks 1-3)
  > - **Wave**: Wave 1
  > - **Blocked By**: Tasks 1, 2, 3

  **Must NOT do**:
  - 不改动其他部分的执行顺序

  **Acceptance Criteria**:
  - [ ] 制表人单元格的 `alignment.horizontal == 'right'`
  - [ ] 制表人单元格的 `alignment.vertical == 'center'`
  - [ ] 整体流程: normalize → remove_empty → find_positions → adjust_print → insert_images → save

- [x] 5. 新增测试用例

  **What to do**:
  - 在 `test_batch_processor.py` 中新增测试类:
    - `test_remove_empty_columns_removes_truly_empty()`
    - `test_remove_empty_columns_preserves_keyword()`
    - `test_remove_empty_columns_keyword_appends_at_end()`
    - `test_remove_empty_columns_no_empty_columns_noop()`
    - `test_remove_empty_columns_with_merged_cells()`
    - `test_zhibiaoren_right_alignment()`
    - `test_get_signature_keywords_from_config()`

  **Test data**: 使用 openpyxl 内存创建 Workbook, 填充模拟数据, 调用函数, 断言结果

  **Recommended Agent Profile**:
  > - **Category**: `unspecified-high`
  >   - Reason: Test writing — requires understanding the function contracts and edge cases.
  > - **Skills**: `[]`

  **Parallelization**:
  > - **Can Run In Parallel**: NO (depends on tasks 1-4 being implemented for functions to test against)
  > - **Wave**: Wave 1
  > - **Blocked By**: Tasks 1, 2, 3, 4

  **Acceptance Criteria**:
  - [ ] 新增 7 个测试, 全部 pass
  - [ ] 原 85 个测试全部 pass

- [x] 6. 全量测试验证

  **Recommended Agent Profile**:
  > - **Category**: `unspecified-high`
  >   - Reason: Final verification — run test suite, collect results, no implementation changes needed.
  > - **Skills**: `[]`

  **Parallelization**:
  > - **Can Run In Parallel**: NO (depends on all tasks 1-5)
  > - **Wave**: Wave 1
  > - **Blocked By**: Tasks 1, 2, 3, 4, 5

  **What to do**:
  - `python -m pytest test/ -v`
  - 确认全部 pass

  **Acceptance Criteria**:
  - [ ] 92/92 tests passed

- [x] 7. 生成永久决策记录文档

  **Recommended Agent Profile**:
  > - **Category**: `writing`
  >   - Reason: Documentation — no code changes, pure markdown writing.
  > - **Skills**: `[]`

  **Parallelization**:
  > - **Can Run In Parallel**: NO (depends on all implementation tasks 1-5 being done)
  > - **Wave**: Wave 1
  > - **Blocked By**: Tasks 1, 2, 3, 4, 5

  **What to do**:
  - 创建 `docs/decisions/empty-columns-removal.md`
  - 记录以下永久性决策:
    - 空列定义: 数据全为空即空列(仅含签字提示词无数据也算)
    - 关键词保护: 删除空列前将签字词右移到最近非空列的同一行
    - 合并单元格: 由 openpyxl 的 delete_cols 自动调整
    - 关键词来源: 从 payroll_sheet_config.json 配置提取(mandatory key + optional key + normalization source)
    - 制表人居右: horizontal='right', vertical='center'
  - git add + commit

  **Must NOT do**:
  - 不包含实现细节或代码片段
  - 只记录设计决策和规则, 不重复计划内容

  **Acceptance Criteria**:
  - [ ] `docs/decisions/empty-columns-removal.md` 文件存在, 包含以上5条决策

---

## Final Verification Wave

- [x] F1. **Plan Compliance Audit** — `oracle`
  逐项检查 Must Have 是否实现, Must NOT Have 是否违规

- [x] F2. **Code Quality Review** — `unspecified-high`
  `pytest test/` 全量通过, 代码无硬编码关键词

- [x] F3. **Real Manual QA** — `unspecified-high`
  用真实 xlsx 文件模拟完整流程: 下载→归一化→去空列→找位置→插签名→保存, 验证输出正确

- [x] F4. **Scope Fidelity Check** — `deep`
  只改 batch_processor.py + test_batch_processor.py, 无跨模块污染

---

## Commit Strategy

- **1+2+3+4**: `feat: 去掉空列+制表人居右 - 删除数据空列(签字词右移保护), 统一关键词获取`
  Files: `app/batch_processor.py`
  Pre-commit: `python -m pytest test/ -v`

- **5**: `test: 去掉空列+制表人对齐测试用例`
  Files: `test/test_batch_processor.py`
  Pre-commit: `python -m pytest test/ -v`

- **7**: `docs: 空列删除+签字词保护设计决策记录`
  Files: `docs/decisions/empty-columns-removal.md`

---

## Success Criteria

### Verification Commands
```bash
python -m pytest test/ -v   # Expected: 92/92 passed
```

### Final Checklist
- [x] 空列被删除, 含关键词的空列关键词已右移
- [x] 制表人单元格 horizontal='right'
- [x] `_auto_column_width` 不再硬编码关键词
- [x] 所有 92 个测试通过
- [x] `docs/decisions/empty-columns-removal.md` 永久决策文档已创建
