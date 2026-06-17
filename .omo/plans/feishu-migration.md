# 飞书审批附件工资表打印 — 钉钉→飞书完整迁移

## TL;DR

> **Quick Summary**: 将钉钉PaySignPrinter审批打印系统完整迁移为飞书API版，新建feishu_api.py模块，复用签名插入/打印逻辑，修复附件下载404和搜索API缺参两个Bug，TDD方式开发。
>
> **Deliverables**:
> - app/feishu_api.py — 飞书API封装（认证、列表、详情、附件下载、搜索）
> - test/test_feishu_api.py — API层TDD单元测试
> - 修正后的API.md文档
> - 新建 app/role_mapping.json / app/user_mapping.json（飞书版）
> - 适配后的 app/cache_manager.py / app/batch_processor.py / app/app.py
> - 完整可运行的Streamlit审批打印应用

> **Estimated Effort**: Large
> **Parallel Execution**: YES - 4 waves
> **Critical Path**: Task 1 → Task 5 → Task 8 → Task 11 → Task 14 → F1-F4

---

## Context

### Original Request
用飞书接口改造钉钉打印审批附件工资表的代码（PaySignPrinter），参考飞书开放cli代码（larksuite-cli/cli）。test_api.py有两个bug需要修复。

### Interview Summary
**Key Discussions**:
- 迁移策略: 完全替换钉钉代码，只保留飞书版
- Bug处理: 整体迁移一起修，在飞书API模块中直接实现正确逻辑
- 运行环境: Windows + WPS/Excel (COM打印)
- 附件缓存: 使用飞书临时URL(12h有效)，每次获取详情时重新获取
- UI框架: 继续用Streamlit
- 测试策略: TDD（测试先行）
- 迁移范围: 完整功能复刻（审批列表、详情、附件下载、签名插入、批量打印）
- 只处理工资表审批: approval_code=1CF34ABB-781C-40B0-9A4F-3CC416612423
- Token自动刷新: 参考 test_api.py 5分钟提前过期缓存模式

**Research Findings**:
- Bug 1根因: attachmentV2的value是完整临时URL而非file_token，不应拼接到drive API路径
- Bug 2根因: /instances/search POST端点需要 user_id, offset, limit, sort_asc 缺失字段
- 飞书审批详情的approver_list结构与钉钉operationRecords完全不同，需要重写审批状态判断
- 飞书可能不在审批详情中返回approver的user_id（只有approver_name），签名映射可能受阻
- API.md文档有多处与实际API不符，需要修正

### Latest PaySignPrinter Changes (2026-06-16)

**1. 新增 `is_ready_for_print()` — 「审批完成待出纳办理」业务逻辑** (commit `ab8a8d7`)
```python
# batch_processor.py — 新增函数
def is_ready_for_print(details: dict) -> bool:
    """审批人已全部批准、只差出纳办理 → 可以打印工资表。"""
    # 从配置中读取必签角色（总经理签字/部长签字/财务审核/业务审核）
    mandatory_roles = set(cfg["sheet_filter"]["signatures"]["mandatory"].keys())
    # 遍历 operationRecords，汇总已同意的审批人角色
    approved_roles = set()
    for record in details.get("operationRecords", []):
        if record.get("result") != "AGREE": return False
        role = get_approver_role(record.get("showName", ""))
        if role: approved_roles.add(role)
    # 所有必签角色都已同意 → 可以打印
    return mandatory_roles.issubset(approved_roles)
```

**UI 影响**: `app.py` 审批状态下拉框新增第一选项「审批完成待出纳办理」（默认选中），查询 RUNNING 实例后进一步按 `is_ready_for_print()` 过滤。

**2. 签名后删除非工资表 sheet** (commit `b61392f`)
```python
# batch_processor.py — wb.save() 前新增
sheets_to_remove = [sn for sn in wb.sheetnames if sn != payroll_ws.title]
for sn in sheets_to_remove:
    del wb[sn]  # 删除工会会费、验证结果等sheet
```
**影响**: openpyxl 操作，与 API 无关，可直接复用。

**3. 全选复选框修复** (commit `0d1c9ca`): Streamlit UI bug fix，不影响迁移。

**4. 其他 openpyxl 相关修复**（与 API 无关，全部直接复用）:
- `_hide_columns()` — 隐藏 部门/岗位/职工号 列（D/E/F），避免 openpyxl 列分组合并 bug
- `_apply_border_styles()` — 只在外部边框画粗线，内部不画
- `ws.print_options.gridLines = False` — 关闭网格线
- 自动替换 "部长、分管副总签字" → "部长签字"
- 多 sheet 工资表检测 — 找第一个匹配的 sheet，跳过工会会费、验证结果等
- `FAQ.md` + `openpyxl_column_hidden_bug_blog.md` — openpyxl 列分组 bug 文档

### 对迁移计划的影响

| PaySignPrinter 变更 | 迁移计划对应 | 需要做什么 |
|-----|------|------|
| `is_ready_for_print()` | Task 10 | 用飞书 `approver_list` 重写，替代钉钉 `operationRecords` |
| 审批状态筛选新增 | Task 14 | app.py 下拉框新增「审批完成待出纳办理」选项 |
| 非工资表 sheet 删除 | Task 11 | 直接复制代码，无需适配 |
| 全选修复 | Task 14 | 直接复制修复后的代码 |

---

## Work Objectives

### Core Objective
将钉钉PaySignPrinter审批打印系统完整迁移为飞书API版，修复现有API bug，实现TDD开发流程。

### Concrete Deliverables
- `/home/ubuntu/coding/approval_feishu/approval_feishu/app/feishu_api.py` — 飞书API封装模块
- `/home/ubuntu/coding/approval_feishu/approval_feishu/test/test_feishu_api.py` — API测试
- `/home/ubuntu/coding/approval_feishu/approval_feishu/API.md` — 修正后的API文档
- `/home/ubuntu/coding/approval_feishu/approval_feishu/role_mapping.json` — 飞书审批角色映射
- `/home/ubuntu/coding/approval_feishu/approval_feishu/user_mapping.json` — 飞书用户ID映射
- 适配后的 cache_manager.py / batch_processor.py / app.py

### Definition of Done
- [ ] `pytest test/` 全部通过
- [ ] `streamlit run app.py` 正常启动，能查询飞书审批列表
- [ ] 能下载附件并保存为本地文件
- [ ] 能在Excel中正确插入签名图片
- [ ] 能通过WPS/Excel COM发送打印任务

### Must Have
- 飞书tenant_access_token认证 + 自动刷新
- 审批实例列表查询（搜索API为主）
- 审批实例详情获取 + 表单解析
- 附件下载（正确处理attachmentV2值，无论URL还是token格式）
- Excel签名插入（总经理签字/部长签字/财务审核/业务审核）
- 批量打印（Windows COM）
- 缓存管理（适配飞书API，12h URL有效期）
- TDD单元测试覆盖所有API函数
- API.md文档修正

### Must NOT Have (Guardrails)
- **[CRITICAL] 不得修改PaySignPrinter钉钉项目任何文件** — 钉钉项目是独立远程仓库，绝不动其代码
- **[CRITICAL] 飞书项目另起仓库** — 新建独立Git仓库，不与钉钉项目共享
- 不得复用dingtalk_api.py的代码 — 飞书API模式完全不同
- 不得使用钉钉的operationRecords模式 — 飞书用approver_list
- 不得复用PaySignPrinter的user_mapping.json — 飞书用户ID不同
- 不得复用PaySignPrinter的role_mapping.json — 飞书审批节点名不同
- 不得添加用户认证/登录系统
- 不得添加PDF生成功能
- 不得添加数据库持久化
- 不得改变Excel打印布局设置（A4横向、边距、适应宽度）
- 不得超出PaySignPrinter现有错误处理级别的重试/退避逻辑
- 不得添加邮件/Webhook通知
- AI slop: 不得添加过度注释、过度抽象、通用命名(data/result/item)

---

## Verification Strategy (MANDATORY)

> **ZERO HUMAN INTERVENTION** - ALL verification is agent-executed. No exceptions.

### Test Decision
- **Infrastructure exists**: NO (原项目无测试)
- **Automated tests**: TDD
- **Framework**: pytest + requests-mock (HTTP mocking)
- **TDD**: Each task follows RED (failing test) → GREEN (minimal impl) → REFACTOR

### QA Policy
Every task MUST include agent-executed QA scenarios.
Evidence saved to `.omo/evidence/task-{N}-{scenario-slug}.{ext}`.

- **API/Backend**: Use Bash (curl) - Send requests, assert status + response fields
- **Library/Module**: Use Bash (python -c) - Import, call functions, compare output
- **UI**: Use Playwright (playwright skill) - Navigate, interact, assert DOM, screenshot

---

## Execution Strategy

### Parallel Execution Waves

```
Wave 0 (Start Immediately - assumption verification):
├── Task 1: 验证飞书API实际行为（attachmentV2值格式、approver字段、权限） [deep]

Wave 1 (After Wave 0 - foundation + scaffolding, MAX PARALLEL):
├── Task 2: 项目脚手架 + pytest配置 [quick]
├── Task 3: 修正API.md文档 [quick]
├── Task 4: 飞书认证模块 feishu_api.py (token获取+缓存) [deep]
├── Task 5: 飞书审批列表查询（搜索API + 列表API） [deep]
├── Task 6: 角色映射配置 role_mapping.json + user_mapping.json [quick]
├── Task 7: 缓存管理模块 cache_manager.py 适配 [unspecified-high]

Wave 2 (After Wave 1 - core modules, MAX PARALLEL):
├── Task 8: 飞书审批实例详情 + 表单解析 + 附件提取 [deep]
├── Task 9: 飞书附件下载（修复Bug 1） [deep]
├── Task 10: 审批状态判断重写（适配飞书approver_list） [unspecified-high]
├── Task 11: batch_processor.py签名插入逻辑适配 [unspecified-high]
├── Task 12: 打印模块适配（Windows COM） [quick]
├── Task 13: logger_config.py + payroll_sheet_config.json 复制 [quick]

Wave 3 (After Wave 2 - integration + UI):
├── Task 14: Streamlit app.py 完整UI构建 [visual-engineering]
├── Task 15: 端到端集成测试 [deep]

Wave FINAL (After ALL tasks — 4 parallel reviews):
├── Task F1: Plan compliance audit (oracle)
├── Task F2: Code quality review (unspecified-high)
├── Task F3: Real manual QA (unspecified-high + playwright)
├── Task F4: Scope fidelity check (deep)
→ Present results → Get explicit user okay
```

### Dependency Matrix

| Task | Depends On | Blocks | Wave |
|------|-----------|--------|------|
| 1 | - | 4, 5, 8, 9, 10 | 0 |
| 2 | - | 4, 5, 7, 8, 9, 10 | 1 |
| 3 | 1 | - | 1 |
| 4 | 1, 2 | 5, 8, 9 | 1 |
| 5 | 1, 2, 4 | 14 | 1 |
| 6 | 1 | 10, 11 | 1 |
| 7 | 2 | 14 | 1 |
| 8 | 1, 2, 4 | 10, 11, 14 | 2 |
| 9 | 1, 2, 4, 8 | 14 | 2 |
| 10 | 1, 6, 8 | 11, 14 | 2 |
| 11 | 6, 8, 10 | 14 | 2 |
| 12 | 2 | 14 | 2 |
| 13 | - | 14 | 2 |
| 14 | 5, 7, 8, 9, 10, 11, 12, 13 | 15 | 3 |
| 15 | 14 | F1-F4 | 3 |

### Agent Dispatch Summary

- **Wave 0**: **1** — T1 → `deep`
- **Wave 1**: **6** — T2 → `quick`, T3 → `quick`, T4 → `deep`, T5 → `deep`, T6 → `quick`, T7 → `unspecified-high`
- **Wave 2**: **6** — T8 → `deep`, T9 → `deep`, T10 → `unspecified-high`, T11 → `unspecified-high`, T12 → `quick`, T13 → `quick`
- **Wave 3**: **2** — T14 → `visual-engineering`, T15 → `deep`
- **FINAL**: **4** — F1 → `oracle`, F2 → `unspecified-high`, F3 → `unspecified-high`, F4 → `deep`

---

## TODOs

- [x] 1. 验证飞书API实际行为（关键假设验证）

  **What to do**:
  - 运行现有 test_api.py，捕获实际API响应并完整记录
  - 重点验证3个关键假设：
    1. **attachmentV2值格式**: 实际值是 file_token (如 `boxbc_xxx`) 还是完整URL? 打印 form widget 的 type 和 value
    2. **approver_list字段**: 飞书审批详情是否返回 approver 的 user_id/open_id? 仅有 approver_name 还是有 ID?
    3. **API权限**: 当前应用权限是否足够支持所有5个API调用（审批列表、详情、搜索、文件下载）
  - 验证搜索API的必填参数（user_id, offset, limit, sort_asc）
  - 将所有实际响应保存到 `feishu_api_verification.md`

  **Must NOT do**:
  - 不要修改test_api.py的API调用逻辑
  - 不要假设API行为，必须实测验证

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: 需要运行实际API调用、分析响应、记录发现，是整个项目的基础
  - **Skills**: []
  - **Skills Evaluated but Omitted**:
    - `playwright`: 不需要浏览器操作

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 0 (solo)
  - **Blocks**: Tasks 4, 5, 8, 9, 10
  - **Blocked By**: None (can start immediately)

  **References**:

  **Pattern References**:
  - `test_api.py:1-220` — 现有测试脚本，已实现5个API调用，需运行并记录实际响应
  - `test_api.py:147-152` — Bug 1 位置：attachmentV2值提取逻辑，需检查实际value格式
  - `test_api.py:195-216` — Bug 2 位置：搜索API请求体，需验证必填参数

  **API/Type References**:
  - `API.md:60-116` — 审批列表API文档（对比实际响应）
  - `API.md:118-165` — 审批详情API文档（对比实际响应，特别是approver字段）
  - `API.md:167-220` — 附件下载API文档（对比实际value格式）
  - `API.md:225-276` — 搜索API文档（对比实际必填参数）

  **External References**:
  - 飞书开放平台: https://open.feishu.cn/document/server-docs/approval-v4/approval-overview

  **WHY Each Reference Matters**:
  - test_api.py 是唯一能实际调用API的代码，必须运行它来获取真实响应
  - API.md 的描述可能与实际不符，必须用实际响应对照验证
  - approver_list 的字段结构决定了签名插入功能是否能实现

  **Acceptance Criteria**:

  > **AGENT-EXECUTABLE VERIFICATION ONLY**

  **If TDD (tests enabled):**
  - [ ] Test: N/A (验证阶段，不写测试代码)
  - [ ] 验证文件 feishu_api_verification.md 已创建

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 验证attachmentV2值格式
    Tool: Bash (python)
    Preconditions: .env文件中有正确的飞书应用凭证
    Steps:
      1. 运行 python test_api.py 2>&1 | tee verification_output.txt
      2. 检查 "测试 3" 输出中附件值的格式
      3. 判断: 值是 "boxbc_xxx" 格式(file_token) 还是 "https://..." 格式(完整URL)?
    Expected Result: 验证文件中明确记录 attachmentV2 值的实际格式
    Failure Indicators: 测试脚本报错无法运行，或附件字段为空
    Evidence: .omo/evidence/task-1-attachment-format.txt

  Scenario: 验证approver_list包含user_id
    Tool: Bash (python)
    Preconditions: 已获取到审批实例详情响应
    Steps:
      1. 在test_api.py的 test_get_instance_detail 中增加打印完整的data字段
      2. 运行测试，检查 approver_list 或 approval_node_list 中的字段
      3. 确认是否包含 user_id / open_id / approver_id 等用户标识字段
    Expected Result: 验证文件中明确记录 approver 相关字段列表，标注是否有用户ID
    Failure Indicators: 实例详情中没有审批人信息
    Evidence: .omo/evidence/task-1-approver-fields.json

  Scenario: 验证搜索API必填参数
    Tool: Bash (python)
    Preconditions: 已有有效的tenant_access_token
    Steps:
      1. 在搜索请求体中添加 user_id, offset, limit, sort_asc 参数
      2. 运行搜索API调用
      3. 确认返回 code=0 而非 99992402
    Expected Result: 搜索API返回成功 (code=0)，确认必填参数列表
    Failure Indicators: 仍然返回 field validation failed
    Evidence: .omo/evidence/task-1-search-params.txt
  ```

  **Evidence to Capture**:
  - [ ] .omo/evidence/task-1-attachment-format.txt
  - [ ] .omo/evidence/task-1-approver-fields.json
  - [ ] .omo/evidence/task-1-search-params.txt

  **Commit**: YES (groups with Wave 0)
  - Message: `docs(feishu): verify actual API behavior and document findings`
  - Files: `feishu_api_verification.md, .omo/evidence/task-1-*`
  - Pre-commit: 无

- [x] 2. 项目脚手架 + pytest配置

  **What to do**:
  - **[CRITICAL] 确认飞书项目为独立Git仓库** — 不得修改或推送PaySignPrinter钉钉项目的任何内容
  - 初始化独立Git仓库（如尚未初始化）: `git init`
  - 创建项目目录结构：
    ```
    /home/ubuntu/coding/approval_feishu/approval_feishu/
    ├── app/                      # 正式代码
    │   ├── __init__.py
    │   ├── feishu_api.py         # (待实现)
    │   ├── cache_manager.py      # (待实现)
    │   ├── batch_processor.py    # (待实现)
    │   ├── app.py                # (待实现 - Streamlit 入口)
    │   ├── logger_config.py      # (待复制)
    │   ├── role_mapping.json     # (待创建)
    │   ├── user_mapping.json     # (待创建)
    │   ├── settings.json
    │   └── payroll_sheet_config.json  # (待复制)
    ├── test/                     # 测试代码
    │   ├── __init__.py
    │   ├── test_api.py           # (已从根目录迁移)
    │   ├── conftest.py           # (待创建)
    │   ├── test_feishu_api.py    # (待实现)
    │   ├── test_cache_manager.py # (待实现)
    │   └── test_approval_status.py  # (待实现)
    ├── downloads/                # 附件下载目录
    ├── signatures/               # 签名图片目录
    ├── .env                      # (已存在)
    ├── requirements.txt          # (需更新)
    └── pyproject.toml            # (已存在)
  - 更新 requirements.txt 添加: pytest, requests-mock
  - 创建 test/conftest.py 配置共享 fixtures (mock token, mock API响应)
  - 创建 pytest.ini 或 pyproject.toml 配置 pytest
  - 验证: `pytest --collect-only` 能发现测试文件

  **Must NOT do**:
  - 不要安装不需要的依赖
  - 不要创建过度复杂的conftest.py

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 创建目录结构和配置文件，工作量小
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 3, 4, 5, 6, 7)
  - **Blocks**: Tasks 4, 5, 7, 8, 9, 10
  - **Blocked By**: None

  **References**:

  **Pattern References**:
  - `PaySignPrinter/requirements.txt` — 原项目依赖列表，需确保核心依赖保留
  - `PaySignPrinter/app.py:1-15` — 原项目import和初始化模式

  **External References**:
  - pytest 官方文档: https://docs.pytest.org/
  - requests-mock: https://requests-mock.readthedocs.io/

  **WHY Each Reference Matters**:
  - PaySignPrinter的requirements.txt确保不遗漏核心依赖(openpyxl, streamlit等)
  - conftest.py需要提供和飞书API一致的mock响应结构

  **Acceptance Criteria**:

  - [ ] `pytest --collect-only` 输出正确（发现0个测试是因为test文件还没写，但不报错）
  - [ ] `python -c "import pytest; import requests_mock"` 不报错

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: pytest配置验证
    Tool: Bash
    Preconditions: 项目脚手架已创建
    Steps:
      1. cd /home/ubuntu/coding/approval_feishu/approval_feishu
      2. pip install pytest requests-mock
      3. python -m pytest --collect-only
    Expected Result: 命令成功退出，无import错误
    Failure Indicators: ModuleNotFoundError 或 ImportError
    Evidence: .omo/evidence/task-2-pytest-setup.txt

  Scenario: 依赖安装验证
    Tool: Bash
    Preconditions: requirements.txt 已更新
    Steps:
      1. pip install -r requirements.txt
      2. python -c "import streamlit, openpyxl, pytest, requests_mock"
    Expected Result: 所有模块导入成功
    Failure Indicators: 任何模块 ImportError
    Evidence: .omo/evidence/task-2-deps-install.txt
  ```

  **Commit**: YES (groups with Wave 1)
  - Message: `build(feishu): add project scaffolding and pytest config`
  - Files: `test/conftest.py, test/__init__.py, requirements.txt, pytest.ini`
  - Pre-commit: `python -m pytest --collect-only`

- [x] 3. 修正API.md文档

  **What to do**:
  - 根据Task 1的验证结果修正API.md中的错误：
    1. 修正API 5 (搜索) 的必填参数：添加 user_id, offset, limit, sort_asc
    2. 修正attachmentV2值格式描述：如果实际返回URL，更新文档说明
    3. 修正审批详情响应中的approver_list字段描述（根据Task 1实际响应更新）
  - 确保API.md与实际API行为完全一致

  **Must NOT do**:
  - 不要添加未验证的API端点
  - 不要删除API.md中正确的部分

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 文档修正，基于Task 1的验证结果
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 4, 5, 6, 7)
  - **Blocks**: None (downstream tasks参考实际API，不依赖文档)
  - **Blocked By**: Task 1

  **References**:

  **Pattern References**:
  - `API.md:1-276` — 需要修正的文档全文
  - `API.md:167-220` — 附件下载部分，Bug 1相关
  - `API.md:225-276` — 搜索API部分，Bug 2相关

  **External References**:
  - 飞书官方API文档: https://open.feishu.cn/document/server-docs/approval-v4/approval-overview

  **WHY Each Reference Matters**:
  - API.md是开发者主要参考文档，必须与实际API行为一致

  **Acceptance Criteria**:

  - [ ] API.md中搜索API参数包含 user_id, offset, limit, sort_asc
  - [ ] API.md中attachmentV2值格式描述与Task 1验证结果一致
  - [ ] API.md中approver_list字段描述与Task 1验证结果一致

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: API.md搜索参数完整性
    Tool: Bash (grep)
    Preconditions: API.md已更新
    Steps:
      1. grep -c "user_id" API.md — 确认搜索API部分包含user_id
      2. grep -c "offset" API.md — 确认搜索API部分包含offset
      3. grep -c "limit" API.md — 确认搜索API部分包含limit
      4. grep -c "sort_asc" API.md — 确认搜索API部分包含sort_asc
    Expected Result: 每个grep至少匹配1次
    Failure Indicators: 任何字段缺失
    Evidence: .omo/evidence/task-3-api-md-check.txt
  ```

  **Commit**: YES (groups with Wave 1)
  - Message: `docs(feishu): fix API.md to match actual API behavior`
  - Files: `API.md`
  - Pre-commit: 无

- [x] 4. 飞书认证模块 feishu_api.py (token获取+缓存)

  **What to do**:
  - TDD: 先写测试 test_feishu_api.py::test_get_tenant_token
    - 测试成功获取token (mock 200响应)
    - 测试token缓存 (第二次调用不发HTTP请求)
    - 测试token过期自动刷新 (5分钟提前过期)
    - 测试认证失败 (错误的app_id/app_secret)
  - 实现 feishu_api.py::get_tenant_token()
    - POST /open-apis/auth/v3/tenant_access_token/internal
    - 文件缓存 token + expire_at (参考test_api.py:33-69)
    - 5分钟提前过期策略
    - 线程安全的文件读写
  - 实现 feishu_api.py::get_auth_headers() 辅助函数
    - 返回 {"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"}

  **Must NOT do**:
  - 不要从dingtalk_api.py复制代码
  - 不要实现OAuth用户登录流程
  - 不要使用user_access_token

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: TDD流程 + 认证是核心基础设施，需要仔细实现缓存和过期逻辑
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 5, 6, 7)
  - **Blocks**: Tasks 5, 8, 9
  - **Blocked By**: Task 1 (需要确认token API格式), Task 2 (需要pytest环境)

  **References**:

  **Pattern References**:
  - `test_api.py:33-69` — Token获取+文件缓存实现，需复用此模式
  - `test_api.py:72-76` — auth_headers() 函数实现
  - `larksuite-cli/cli/internal/credential/tat_fetch.go:33-69` — Go版token获取参考

  **API/Type References**:
  - `API.md:29-57` — 飞书认证API文档（请求/响应格式）
  - `.env` — 包含 FEISHU_APP_ID, FEISHU_APP_SECRET

  **External References**:
  - 飞书认证文档: https://open.feishu.cn/document/server-docs/authentication/tenant_access_token

  **WHY Each Reference Matters**:
  - test_api.py的token缓存实现已经验证可用，应作为主要参考
  - larksuite-cli的Go实现展示了错误分类模式，可借鉴
  - .env中的凭证是实际调用需要的

  **Acceptance Criteria**:

  **If TDD:**
  - [ ] Test file created: test/test_feishu_api.py
  - [ ] `python -m pytest test/test_feishu_api.py::test_get_tenant_token -v` → PASS

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Token获取成功
    Tool: Bash (python)
    Preconditions: .env中有正确的飞书凭证, pytest已配置
    Steps:
      1. python -m pytest test/test_feishu_api.py::test_get_tenant_token -v
      2. 检查: 成功测试、缓存测试、过期测试、失败测试全部通过
    Expected Result: 4个测试全部PASS
    Failure Indicators: 任何测试FAIL
    Evidence: .omo/evidence/task-4-token-tests.txt

  Scenario: Token缓存文件验证
    Tool: Bash (python)
    Preconditions: Token已获取
    Steps:
      1. python -c "from feishu_api import get_tenant_token; t = get_tenant_token(); print(f'Token: {t[:10]}...')"
      2. 检查 token_cache.json 文件已创建
      3. 再次调用，验证不发HTTP请求（使用缓存）
    Expected Result: 第二次调用更快，无HTTP请求
    Failure Indicators: 每次调用都发HTTP请求
    Evidence: .omo/evidence/task-4-token-cache.txt
  ```

  **Commit**: YES (groups with Wave 1)
  - Message: `feat(feishu): add tenant access token auth with caching`
  - Files: `feishu_api.py, test/test_feishu_api.py`
  - Pre-commit: `python -m pytest test/test_feishu_api.py -v`

- [x] 5. 飞书审批列表查询（搜索API + 列表API）

  **What to do**:
  - TDD: 先写测试 test_feishu_api.py::test_search_instances + test_list_instances
    - 测试搜索实例成功 (mock 200, 包含instance_list)
    - 测试搜索实例必填参数 (user_id, offset, limit, sort_asc)
    - 测试列表查询成功 (mock 200, 分页)
    - 测试空结果处理
    - 测试API错误处理 (code != 0)
    - 测试网络错误处理
   - 实现 feishu_api.py::query_instances()
     - POST /open-apis/approval/v4/instances/query（正确的飞书查询端点，非 /search）
     - 参数: approval_code, page_size, page_token, instance_status(可选), 时间范围(可选)
     - 返回 instance_list（含 instance_code, status, start_time）
  - 实现 feishu_api.py::list_instances()
    - GET /open-apis/approval/v4/instances
    - 参数: approval_code, start_time, end_time, page_size, page_token
    - 自动分页 (while page_token)
  - 实现搜索API为默认查询方式，列表API为备用

  **Must NOT do**:
  - 不要两个API同时作为主查询，选search为主
  - 不要忘记搜索API的必填参数
  - 不要复用钉钉的分页逻辑（飞书用page_token，不是offset/size）

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: TDD + 两个API实现 + 分页逻辑 + 参数验证
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 4, 6, 7)
  - **Blocks**: Task 14
  - **Blocked By**: Task 1 (需确认搜索API参数), Task 2 (需pytest), Task 4 (需auth)

  **References**:

  **Pattern References**:
  - `test_api.py:89-121` — 列表查询实现（分页、参数构建）
  - `test_api.py:195-216` — 搜索API实现（Bug 2修复参考）
  - `PaySignPrinter/dingtalk_api.py:51-98` — 钉钉列表查询模式（对比差异，不复制）

  **API/Type References**:
  - `API.md:60-116` — 列表API文档
  - `API.md:225-276` — 搜索API文档（修正后版本）

  **External References**:
  - 飞书审批实例查询: https://open.feishu.cn/document/server-docs/approval-v4/instance/query

  **WHY Each Reference Matters**:
  - test_api.py的列表查询已有分页逻辑可复用
  - 搜索API的参数修复是Bug 2的核心，必须参考修正后的API.md

  **Acceptance Criteria**:

  **If TDD:**
  - [ ] `python -m pytest test/test_feishu_api.py::test_search_instances -v` → PASS
  - [ ] `python -m pytest test/test_feishu_api.py::test_list_instances -v` → PASS

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 搜索API成功查询
    Tool: Bash (python)
    Preconditions: feishu_api.py已实现search_instances, mock已配置
    Steps:
      1. python -m pytest test/test_feishu_api.py::test_search_instances -v
      2. 验证: 成功查询、必填参数校验、空结果、错误处理、网络错误 全部PASS
    Expected Result: 5个测试全部PASS
    Failure Indicators: 任何测试FAIL
    Evidence: .omo/evidence/task-5-search-tests.txt

  Scenario: 列表API自动分页
    Tool: Bash (python)
    Preconditions: feishu_api.py已实现list_instances
    Steps:
      1. python -m pytest test/test_feishu_api.py::test_list_instances -v
      2. 验证: 分页逻辑正确处理has_more和page_token
    Expected Result: 分页测试PASS
    Failure Indicators: 分页不完整或重复
    Evidence: .omo/evidence/task-5-list-tests.txt

  Scenario: 搜索API必填参数Bug验证
    Tool: Bash (python -c)
    Preconditions: 实际API可访问
    Steps:
      1. python -c "from feishu_api import search_instances; r = search_instances(approval_code='1CF34ABB...', user_id='me', offset=0, limit=10, sort_asc=False); print(r)"
      2. 确认返回code=0而非99992402
    Expected Result: API返回成功 (code=0)
    Failure Indicators: 仍返回 field validation failed
    Evidence: .omo/evidence/task-5-search-fix.txt
  ```

  **Commit**: YES (groups with Wave 1)
  - Message: `feat(feishu): add approval instance search and list query`
  - Files: `feishu_api.py, test/test_feishu_api.py`
  - Pre-commit: `python -m pytest test/test_feishu_api.py -v`

- [x] 6. 角色映射配置 role_mapping.json + user_mapping.json

  **What to do**:
  - 根据Task 1验证的飞书审批定义，创建 role_mapping.json
    - 映射飞书审批节点名称 → Excel签名位置关键词
    - 例: {"审批人": "部长签字", "总经理": "总经理签字", ...}
    - 需要从飞书审批定义中获取实际的节点名称
  - 创建初始 user_mapping.json
    - 格式: {"飞书user_id/open_id": "姓名"}
    - 初始可为空 {}，后续通过Streamlit UI上传Excel填充
    - 参考 PaySignPrinter/user_mapping.json 的格式
  - 复制 payroll_sheet_config.json（如存在）— 工资表检测规则不依赖API

  **Must NOT do**:
  - 不要复制PaySignPrinter的role_mapping.json（飞书节点名不同）
  - 不要复制PaySignPrinter的user_mapping.json（飞书用户ID不同）
  - 不要硬编码审批节点名称，必须从实际审批定义获取

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 创建JSON配置文件
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 4, 5, 7)
  - **Blocks**: Tasks 10, 11
  - **Blocked By**: Task 1 (需要确认审批节点名称)

  **References**:

  **Pattern References**:
  - `PaySignPrinter/role_mapping.json` — 钉钉版角色映射，参考格式而非内容
  - `PaySignPrinter/user_mapping.json` — 钉钉版用户映射，参考格式而非内容
  - `PaySignPrinter/payroll_sheet_config.json` — 工资表检测配置，可原样复制

  **API/Type References**:
  - Task 1 验证结果中的审批定义节点名称

  **WHY Each Reference Matters**:
  - role_mapping.json的格式需要与batch_processor.py的签名查找逻辑一致
  - payroll_sheet_config.json的检测规则是Excel格式依赖的，与API无关

  **Acceptance Criteria**:

  - [ ] role_mapping.json 包含至少飞书工资表审批的所有节点名
  - [ ] user_mapping.json 格式正确 (JSON有效)
  - [ ] payroll_sheet_config.json 与PaySignPrinter版本一致

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: JSON文件格式验证
    Tool: Bash (python)
    Preconditions: JSON文件已创建
    Steps:
      1. python -c "import json; json.load(open('role_mapping.json')); print('role_mapping OK')"
      2. python -c "import json; json.load(open('user_mapping.json')); print('user_mapping OK')"
      3. python -c "import json; d = json.load(open('role_mapping.json')); assert len(d) >= 3; print(f'{len(d)} roles mapped')"
    Expected Result: 所有JSON文件可正确解析，role_mapping至少3条映射
    Failure Indicators: JSON解析错误或映射数量不足
    Evidence: .omo/evidence/task-6-json-config.txt

  Scenario: payroll_sheet_config兼容性
    Tool: Bash (python)
    Preconditions: payroll_sheet_config.json已复制
    Steps:
      1. python -c "import json; d = json.load(open('payroll_sheet_config.json')); print(d.keys())"
    Expected Result: 配置文件与PaySignPrinter版本键值一致
    Failure Indicators: 缺少键或格式不匹配
    Evidence: .omo/evidence/task-6-payroll-config.txt
  ```

  **Commit**: YES (groups with Wave 1)
  - Message: `feat(feishu): add role and user mapping config files`
  - Files: `role_mapping.json, user_mapping.json, payroll_sheet_config.json`
  - Pre-commit: `python -c "import json; json.load(open('role_mapping.json'))"`

- [x] 7. 缓存管理模块 cache_manager.py 适配

  **What to do**:
  - 适配 PaySignPrinter/cache_manager.py 为飞书版
    - 保留: 缓存文件读写、命中/未命中统计、force_refresh机制
    - 修改: 缓存键从钉钉参数改为飞书参数(approval_code, instance_code等)
    - 修改: 下载URL缓存TTL从15分钟改为12小时（飞书URL有效期12h）
    - 修改: 实例详情缓存需适配飞书的form JSON解析
    - 添加: 打印状态跟踪（printed标记）
  - TDD: 测试缓存命中、未命中、过期、force_refresh

  **Must NOT do**:
  - 不要添加SQLite或其他数据库
  - 不要改变缓存的文件存储方式（保持JSON文件）
  - 不要添加超出PaySignPrinter现有功能的缓存策略

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
    - Reason: 需要理解PaySignPrinter的缓存逻辑并适配飞书格式
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 4, 5, 6)
  - **Blocks**: Task 14
  - **Blocked By**: Task 2 (需要pytest)

  **References**:

  **Pattern References**:
  - `PaySignPrinter/cache_manager.py:1-204` — 完整的缓存实现，需适配飞书
  - `PaySignPrinter/cache_manager.py:72-84` — 缓存读取模式
  - `PaySignPrinter/cache_manager.py:175-203` — 下载URL缓存TTL逻辑

  **WHY Each Reference Matters**:
  - PaySignPrinter的缓存管理器已验证可靠，只需适配飞书API的缓存键和TTL
  - 12小时TTL是飞书临时URL的有效期，必须正确设置

  **Acceptance Criteria**:

  **If TDD:**
  - [ ] `python -m pytest test/test_cache_manager.py -v` → PASS

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 缓存命中和未命中
    Tool: Bash (python)
    Preconditions: cache_manager.py已实现
    Steps:
      1. python -m pytest test/test_cache_manager.py -v
      2. 验证: 命中测试、未命中测试、过期测试、force_refresh测试 全部PASS
    Expected Result: 4+个测试全部PASS
    Failure Indicators: 任何测试FAIL
    Evidence: .omo/evidence/task-7-cache-tests.txt

  Scenario: 下载URL缓存12小时TTL
    Tool: Bash (python)
    Preconditions: 缓存实现完成
    Steps:
      1. python -c "
from cache_manager import cache_download_url, get_cached_download_url
cache_download_url('test_key', 'https://example.com/file', ttl_seconds=43200)
result = get_cached_download_url('test_key')
assert result is not None, 'Cache miss'
print('Cache hit: OK')
"
    Expected Result: 缓存命中，12小时TTL
    Failure Indicators: 缓存未命中或TTL不正确
    Evidence: .omo/evidence/task-7-cache-ttl.txt
  ```

  **Commit**: YES (groups with Wave 1)
  - Message: `feat(feishu): adapt cache manager for Feishu API`
  - Files: `cache_manager.py, test/test_cache_manager.py`
  - Pre-commit: `python -m pytest test/test_cache_manager.py -v`

- [x] 8. 飞书审批实例详情 + 表单解析 + 附件提取

  **What to do**:
  - TDD: 先写测试 test_feishu_api.py::test_get_instance_detail + test_extract_attachments
    - 测试获取实例详情成功 (mock 200, 包含form/approver_list)
    - 测试表单解析 (json.loads(form) → widget列表)
    - 测试attachmentV2提取 (单个附件、多个附件、无附件)
    - 测试非attachmentV2组件过滤
    - 测试form字段为空或无效JSON
    - 测试API错误和网络错误
  - 实现 feishu_api.py::get_instance_detail(instance_code)
    - GET /open-apis/approval/v4/instances/{instance_code}
    - 返回完整detail dict
  - 实现 feishu_api.py::parse_form(detail)
    - json.loads(detail.get("form", "[]"))
    - 返回 widget 列表，每个widget含 id, type, name, value
  - 实现 feishu_api.py::extract_attachments(form_widgets)
    - 过滤 type == "attachmentV2"
    - 提取附件信息: field_name, file_token/URL (根据Task 1验证结果处理)
    - 处理value可能是file_token数组或URL数组的情况

  **Must NOT do**:
  - 不要使用钉钉的formComponentValues解析逻辑
  - 不要假设attachmentV2的value一定是file_token或一定是URL — 需兼容两种格式
  - 不要遗漏多个attachmentV2组件的情况

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: 核心业务逻辑 + TDD + 需处理多种可能的响应格式
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 9, 10, 11, 12, 13)
  - **Blocks**: Tasks 10, 11, 14
  - **Blocked By**: Task 1 (需确认form/attachment格式), Task 2 (需pytest), Task 4 (需auth)

  **References**:

  **Pattern References**:
  - `test_api.py:124-163` — 实例详情获取+表单解析+附件提取（Bug 1所在）
  - `PaySignPrinter/dingtalk_api.py:101-153` — 钉钉版详情获取+附件提取（对比差异，不复制）

  **API/Type References**:
  - `API.md:118-165` — 审批详情API文档
  - Task 1 验证结果中实际的form和attachment结构

  **WHY Each Reference Matters**:
  - test_api.py已有正确的API调用模式但Bug 1在附件提取，需要修复提取逻辑
  - 必须兼容两种attachmentV2值格式（file_token或URL），这是Bug 1的根因

  **Acceptance Criteria**:

  **If TDD:**
  - [ ] `python -m pytest test/test_feishu_api.py::test_get_instance_detail -v` → PASS
  - [ ] `python -m pytest test/test_feishu_api.py::test_extract_attachments -v` → PASS

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 详情获取+表单解析成功
    Tool: Bash (python)
    Preconditions: feishu_api.py已实现
    Steps:
      1. python -m pytest test/test_feishu_api.py::test_get_instance_detail -v
      2. 验证: 成功获取、表单解析、空form、无效JSON、API错误、网络错误 全部PASS
    Expected Result: 6个测试全部PASS
    Failure Indicators: 任何测试FAIL
    Evidence: .omo/evidence/task-8-detail-tests.txt

  Scenario: 附件提取兼容两种格式
    Tool: Bash (python)
    Preconditions: feishu_api.py已实现extract_attachments
    Steps:
      1. python -m pytest test/test_feishu_api.py::test_extract_attachments -v
      2. 验证: file_token格式、URL格式、多附件、无附件、非attachment组件 全部PASS
    Expected Result: 5+个测试全部PASS
    Failure Indicators: URL格式的附件提取失败（Bug 1未修复）
    Evidence: .omo/evidence/task-8-attachment-tests.txt

  Scenario: 实际API调用附件提取
    Tool: Bash (python -c)
    Preconditions: 有真实飞书审批实例可用
    Steps:
      1. python -c "from feishu_api import get_instance_detail, extract_attachments, parse_form; d = get_instance_detail('INSTANCE_CODE'); w = parse_form(d); a = extract_attachments(w); print(f'Attachments: {len(a)}'); [print(f'  {x}') for x in a]"
      2. 确认附件数量和格式正确
    Expected Result: 附件数量≥1，格式包含field_name和file_token/URL
    Failure Indicators: 0个附件或格式解析错误
    Evidence: .omo/evidence/task-8-real-attachment.txt
  ```

  **Commit**: YES (groups with Wave 2)
  - Message: `feat(feishu): add instance detail, form parsing, and attachment extraction`
  - Files: `feishu_api.py, test/test_feishu_api.py`
  - Pre-commit: `python -m pytest test/test_feishu_api.py -v`

- [x] 9. 飞书附件下载（修复Bug 1）

  **What to do**:
  - TDD: 先写测试 test_feishu_api.py::test_download_attachment
    - 测试file_token下载成功 (通过drive API)
    - 测试URL直接下载成功 (临时URL，12h有效)
    - 测试Content-Disposition文件名提取
    - 测试stream下载大文件
    - 测试404/403错误处理
    - 测试网络超时处理
  - 实现 feishu_api.py::download_attachment(file_token_or_url, save_dir)
    - **方案A**: 如果value是file_token → GET /open-apis/drive/v1/files/{file_token}/download
    - **方案B**: 如果value是完整URL → 直接GET URL (带Authorization header)
    - 自动检测值格式: 以 http 开头用方案B, 否则用方案A
    - 提取Content-Disposition中的文件名
    - stream=True 下载，chunk_size=8192
    - 返回保存的文件路径

  **Must NOT do**:
  - 不要把完整URL当file_token拼接到drive API路径（这是Bug 1）
  - 不要假设value一定是某种格式 — 必须自动检测
  - 不要忘记stream=True下载大文件

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: Bug 1修复核心 + TDD + 需处理两种下载路径
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 8, 10, 11, 12, 13)
  - **Blocks**: Task 14
  - **Blocked By**: Task 1 (需确认附件值格式), Task 2 (需pytest), Task 4 (需auth), Task 8 (需extract_attachments)

  **References**:

  **Pattern References**:
  - `test_api.py:166-192` — 当前下载实现（有Bug 1，需修复）
  - `test_api.py:26` — DRIVE_DOWNLOAD_URL定义
  - `larksuite-cli/cli/shortcuts/drive/drive_download.go:41-90` — Go版下载参考

  **API/Type References**:
  - `API.md:167-220` — 附件下载API文档
  - Task 1 验证结果中附件值的实际格式

  **WHY Each Reference Matters**:
  - test_api.py的下载实现是起点但需修复URL拼接bug
  - larksuite-cli展示了正确的stream下载模式
  - 自动检测格式是兼容两种可能性的关键

  **Acceptance Criteria**:

  **If TDD:**
  - [ ] `python -m pytest test/test_feishu_api.py::test_download_attachment -v` → PASS

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: file_token格式下载成功
    Tool: Bash (python)
    Preconditions: feishu_api.py已实现download_attachment
    Steps:
      1. python -m pytest test/test_feishu_api.py::test_download_attachment -v
      2. 验证: file_token下载、URL直接下载、文件名提取、大文件stream、错误处理 全部PASS
    Expected Result: 5+个测试全部PASS
    Failure Indicators: 任何测试FAIL
    Evidence: .omo/evidence/task-9-download-tests.txt

  Scenario: Bug 1修复验证 — 不再出现404
    Tool: Bash (python -c)
    Preconditions: 有真实飞书审批附件
    Steps:
      1. python -c "
from feishu_api import get_instance_detail, parse_form, extract_attachments, download_attachment
from pathlib import Path
d = get_instance_detail('INSTANCE_CODE')
w = parse_form(d)
a = extract_attachments(w)
if a:
    p = download_attachment(a[0]['file_token'], Path('./downloads'))
    print(f'Downloaded: {p}, size={p.stat().st_size}')
"
      2. 确认下载成功且文件大小>0
    Expected Result: 文件下载成功，无404错误
    Failure Indicators: 404 Not Found 或文件大小=0
    Evidence: .omo/evidence/task-9-bug1-fix.txt
  ```

  **Commit**: YES (groups with Wave 2)
  - Message: `feat(feishu): add attachment download with Bug 1 fix`
  - Files: `feishu_api.py, test/test_feishu_api.py`
  - Pre-commit: `python -m pytest test/test_feishu_api.py -v`

- [x] 10. 审批状态判断重写（适配飞书approver_list）

  **What to do**:
  - 用飞书 `approver_list` 重写 PaySignPrinter 的审批状态判断函数：

  **10a. `is_ready_for_print(details)` — 「审批完成待出纳办理」逻辑** ⭐ 最重要
  - 钉钉版依赖 `operationRecords` + `get_approver_role(showName)` — 不可复用
  - 飞书版改为遍历 `approver_list`，检查每个 approve 的 `status`
  - 从 `payroll_sheet_config.json` 读取必签角色列表 (mandatory roles)
  - 通过飞书 `role_mapping.json` 将 `approver_name` 映射到 Excel 签名角色
  - 返回 True 当且仅当所有必签角色都已状态为 APPROVED
  - 处理边界: approver_list 为空、部分审批人未完成、角色映射缺失

  **10b. `is_approval_passed(details)` — 审批是否完全通过**
  - 钉钉版遍历 `operationRecords` — 不可复用
  - 飞书版检查 `detail["status"]` 是否为 APPROVED
  - 如果 status 是 APPROVED 则全部通过

  **10c. `get_approvers_with_roles(details)` — 提取审批人及其角色**
  - 钉钉版遍历 `operationRecords` + `get_approver_role(showName)` — 不可复用
  - 飞书版遍历 `approver_list`
  - 每个 approver 返回: (approver_name, approver_id (如果有), 角色, 审批状态)
  - 角色通过 `role_mapping.json` 从 approver_name 映射
  - 如果飞书不返回 approver_id，签名查找改用 approver_name 匹配 PNG 文件

  **TDD 测试覆盖**:
  - is_ready_for_print: 全通过、部分通过、无人通过、空列表、角色映射缺失
  - is_approval_passed: APPROVED/PENDING/REJECTED/DELETED/RECALL
  - get_approvers_with_roles: 有ID、无ID(仅name)、混合状态

  **Must NOT do**:
  - 不要使用钉钉的operationRecords模式
  - 不要假设飞书返回与钉钉相同的状态值
  - 不要假设飞书一定返回approver的user_id

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
    - Reason: 核心业务逻辑重写，需深入理解两个平台的审批流程差异
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 8, 9, 11, 12, 13)
  - **Blocks**: Tasks 11, 14
  - **Blocked By**: Task 1 (需确认approver字段), Task 6 (需role_mapping), Task 8 (需detail解析)

  **References**:

  **Pattern References**:
  - `PaySignPrinter/batch_processor.py:223-337` — 钉钉版审批状态判断+签名角色映射，需重写
  - `PaySignPrinter/role_mapping.json` — 格式参考

  **API/Type References**:
  - Task 1 验证结果中approver_list的实际结构
  - `role_mapping.json` — 飞书版角色映射

  **WHY Each Reference Matters**:
  - PaySignPrinter的签名角色映射是核心业务逻辑，需理解其工作原理后在飞书语境下重写
  - approver_list的实际字段结构决定了签名映射策略

  **Acceptance Criteria**:

  **If TDD:**
  - [ ] `python -m pytest test/test_approval_status.py -v` → PASS

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 飞书审批状态判断
    Tool: Bash (python)
    Preconditions: 审批状态函数已实现
    Steps:
      1. python -m pytest test/test_approval_status.py -v
      2. 验证: APPROVED通过、PENDING待审、REJECTED拒绝、RECALL撤回、DELETED删除 全部PASS
    Expected Result: 5+个状态测试PASS
    Failure Indicators: 状态判断逻辑错误
    Evidence: .omo/evidence/task-10-status-tests.txt

  Scenario: 审批人角色映射
    Tool: Bash (python)
    Preconditions: get_approvers_with_roles已实现
    Steps:
      1. python -c "
from batch_processor import get_approvers_with_roles
detail = {'approver_list': [...]}  # 使用Task 1验证的实际结构
roles = get_approvers_with_roles(detail, 'role_mapping.json')
print(f'Roles: {roles}')
"
      2. 确认角色映射结果包含签名位置关键词
    Expected Result: 角色映射输出包含签名关键词（总经理签字/部长签字/财务审核/业务审核）
    Failure Indicators: 空结果或缺少关键角色
    Evidence: .omo/evidence/task-10-role-mapping.txt
  ```

  **Commit**: YES (groups with Wave 2)
  - Message: `feat(feishu): rewrite approval status logic for Feishu approver_list`
  - Files: `batch_processor.py, test/test_approval_status.py`
  - Pre-commit: `python -m pytest test/test_approval_status.py -v`

- [x] 11. batch_processor.py签名插入逻辑适配

  **What to do**:
  - 适配 PaySignPrinter/batch_processor.py 中的签名插入和打印逻辑
     - 保留（不修改）: find_all_signature_positions(), _insert_signature_to_excel_openpyxl()
     - 保留（不修改）: 打印布局设置（A4横向、边距、适应宽度、隐藏列）
     - 保留（不修改）: `_hide_columns()` — 隐藏部门/岗位/职工号列，含 openpyxl 列合并修复
     - 保留（不修改）: `_apply_border_styles()` — 外框粗线边框，GridLines=False
     - 保留（不修改）: payroll sheet 检测逻辑
     - **保留（不修改）: 签名后删除非工资表 sheet** (commit b61392f — 删除工会会费、验证结果等sheet)
    - 修改: process_single_approval() 中的钉钉API调用 → 飞书API调用
    - 修改: 签名查找逻辑 — 如果飞书只返回approver_name，需支持按名字查找签名PNG
    - 修改: 附件过滤逻辑 — 钉钉DDAttachment → 飞书attachmentV2
    - 保留: "汇总表"跳过逻辑
    - 保留: Windows COM打印 + Linux LibreOffice打印
  - 适配 process_single_approval() 流程:
    1. get_instance_detail(飞书) → parse_form → extract_attachments
    2. download_attachment (Bug 1已修复)
    3. get_approvers_with_roles (Task 10已重写)
    4. find_all_signature_positions (不修改)
    5. _insert_signature_to_excel_openpyxl (不修改)
    6. 打印布局调整 (不修改)
    7. print_file (不修改)

  **Must NOT do**:
  - 不要修改签名位置检测逻辑（与Excel格式绑定，与API无关）
  - 不要修改打印布局设置
  - 不要修改openpyxl签名插入核心逻辑
  - 不要改变"汇总表"过滤规则

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
    - Reason: 大量代码适配，需理解钉钉→飞书的差异并正确映射
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 8, 9, 10, 12, 13)
  - **Blocks**: Task 14
  - **Blocked By**: Task 6 (需role/user mapping), Task 8 (需detail解析), Task 10 (需审批状态函数)

  **References**:

  **Pattern References**:
  - `PaySignPrinter/batch_processor.py:140-196` — 工资表检测逻辑（保留）
  - `PaySignPrinter/batch_processor.py:284-337` — 签名插入核心逻辑（保留）
  - `PaySignPrinter/batch_processor.py:374-400` — 打印布局设置（保留）
  - `PaySignPrinter/batch_processor.py:403-476` — 签名+打印流程（适配API调用）
  - `PaySignPrinter/batch_processor.py:526-535` — 跨平台打印（保留）

  **WHY Each Reference Matters**:
  - 明确哪些代码可原样保留，哪些需要适配，避免不必要的重写

  **Acceptance Criteria**:

  - [ ] process_single_approval() 使用飞书API而非钉钉API
  - [ ] 签名查找支持按name查找（如果飞书不返回user_id）
  - [ ] openpyxl签名插入逻辑不变
  - [ ] 打印逻辑不变

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 签名插入逻辑不变验证
    Tool: Bash (python)
    Preconditions: batch_processor.py已适配
    Steps:
      1. python -c "
from batch_processor import find_all_signature_positions
# 使用PaySignPrinter中的测试Excel验证逻辑不变
"
      2. 对比适配前后的find_all_signature_positions函数签名和返回值
    Expected Result: 函数签名和返回值格式不变
    Failure Indicators: 函数签名改变或返回值格式不同
    Evidence: .omo/evidence/task-11-signature-logic.txt

  Scenario: process_single_approval使用飞书API
    Tool: Bash (python -c)
    Preconditions: 所有前置模块已实现
    Steps:
      1. python -c "
import inspect
from batch_processor import process_single_approval
source = inspect.getsource(process_single_approval)
assert 'dingtalk' not in source.lower(), 'Still references DingTalk'
assert 'feishu_api' in source.lower() or 'get_instance_detail' in source, 'Not using Feishu API'
print('OK: uses Feishu API')
"
    Expected Result: 函数中不再有钉钉API引用
    Failure Indicators: 仍包含钉钉API调用
    Evidence: .omo/evidence/task-11-api-switch.txt
  ```

  **Commit**: YES (groups with Wave 2)
  - Message: `feat(feishu): adapt batch processor for Feishu API`
  - Files: `batch_processor.py`
  - Pre-commit: `python -c "from batch_processor import process_single_approval"`

- [x] 12. 打印模块适配（Windows COM）

  **What to do**:
  - 从PaySignPrinter复制打印相关代码到batch_processor.py
    - print_file() 函数
    - _print_with_com() Windows COM打印
    - _print_with_libreoffice() Linux LibreOffice打印（保留兼容性）
    - 打印布局设置函数
  - 确保Windows COM打印能正确处理飞书下载的Excel文件
  - 验证WPS/Excel COM能打开飞书下载的.xlsx/.xls文件

  **Must NOT do**:
  - 不要改变打印布局设置（A4横向、边距等）
  - 不要删除Linux LibreOffice打印支持（保留跨平台兼容）

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 主要是复制+验证，不涉及复杂逻辑重写
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 8, 9, 10, 11, 13)
  - **Blocks**: Task 14
  - **Blocked By**: Task 2 (项目脚手架)

  **References**:

  **Pattern References**:
  - `PaySignPrinter/batch_processor.py:460-535` — Windows COM + LibreOffice打印实现
  - `PaySignPrinter/batch_processor.py:374-400` — 打印布局设置

  **WHY Each Reference Matters**:
  - 打印模块与API平台无关，可原样复制

  **Acceptance Criteria**:

  - [ ] print_file() 函数可用
  - [ ] _print_with_com() 保留Windows支持
  - [ ] _print_with_libreoffice() 保留Linux支持

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 打印函数可用性
    Tool: Bash (python)
    Preconditions: batch_processor.py包含打印函数
    Steps:
      1. python -c "from batch_processor import print_file, _print_with_com, _print_with_libreoffice; print('OK')"
    Expected Result: 导入成功无报错
    Failure Indicators: ImportError
    Evidence: .omo/evidence/task-12-print-import.txt
  ```

  **Commit**: YES (groups with Wave 2)
  - Message: `feat(feishu): add print module with Windows COM and LibreOffice support`
  - Files: `batch_processor.py`
  - Pre-commit: `python -c "from batch_processor import print_file"`

- [x] 13. logger_config.py + payroll_sheet_config.json 复制

  **What to do**:
  - 从PaySignPrinter复制 logger_config.py，修改项目名称为 approval_feishu
  - 从PaySignPrinter复制 payroll_sheet_config.json（如有），无需修改
  - 验证logger配置能正确输出到文件和stdout

  **Must NOT do**:
  - 不要修改日志格式
  - 不要修改payroll_sheet_config.json的检测规则

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 简单复制+微调
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 8, 9, 10, 11, 12)
  - **Blocks**: Task 14
  - **Blocked By**: None

  **References**:

  **Pattern References**:
  - `PaySignPrinter/logger_config.py` — 日志配置（复制+改名）
  - `PaySignPrinter/payroll_sheet_config.json` — 工资表检测规则（复制）

  **Acceptance Criteria**:

  - [ ] python -c "from logger_config import get_logger; logger = get_logger('test'); logger.info('OK')" 无报错

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Logger配置验证
    Tool: Bash (python)
    Preconditions: logger_config.py已复制
    Steps:
      1. python -c "from logger_config import get_logger; l = get_logger('test'); l.info('Test message'); print('OK')"
    Expected Result: 无报错，日志输出正常
    Failure Indicators: ImportError或日志输出异常
    Evidence: .omo/evidence/task-13-logger.txt
  ```

  **Commit**: YES (groups with Wave 2)
  - Message: `build(feishu): add logger config and payroll sheet detection config`
  - Files: `logger_config.py, payroll_sheet_config.json`
  - Pre-commit: `python -c "from logger_config import get_logger"`

- [ ] 14. Streamlit app.py 完整UI构建

  **What to do**:
  - 基于PaySignPrinter/app.py重构为飞书版
    - 保持相同的UI布局和交互逻辑
    - 替换所有钉钉API调用为飞书API调用
    - 替换钉钉状态显示为飞书状态 (APPROVED/PENDING/REJECTED/RECALL/DELETED)
    - 适配飞书表单显示 (form JSON → widget列表 → 分类型显示)
    - 适配飞书审批人显示 (approver_list → 审批节点+状态)
    - 添加附件下载功能 (调用Task 9的download_attachment)
    - 添加签名插入+打印功能 (调用Task 11的process_single_approval)
    - 添加角色映射和用户映射管理UI
    - 保留: 日期选择、状态筛选、批量操作、打印设置
  - 飞书特有的UI调整:
    - 审批列表查询改用 `query_instances` (/instances/query，page_size/page_token 分页)
    - **新增「审批完成待出纳办理」筛选** — 默认选中，查询 RUNNING 实例后按 is_ready_for_print() 过滤
    - 审批状态下拉框选项: 审批完成待出纳办理 → 已完结 → 已完结未打印 → 审批中 → 已撤销 → 全部
    - 附件显示适配attachmentV2格式
    - Token刷新集成到Streamlit session_state
    - **直接复用 PaySignPrinter 修复后的全选复选框逻辑** (commit 0d1c9ca)

  **Must NOT do**:
  - 不要改变Streamlit页面布局结构
  - 不要添加用户登录功能
  - 不要添加钉钉双平台切换
  - 不要使用钉钉的组件类型名(DDAttachment等)

  **Recommended Agent Profile**:
  - **Category**: `visual-engineering`
    - Reason: Streamlit UI构建，需确保界面正确显示和交互
  - **Skills**: [`playwright`]
    - `playwright`: UI验证需要浏览器操作

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 3 (with Task 15, but 14 must come first)
  - **Blocks**: Task 15, F1-F4
  - **Blocked By**: Tasks 5, 7, 8, 9, 10, 11, 12, 13

  **References**:

  **Pattern References**:
  - `PaySignPrinter/app.py:1-787` — 完整Streamlit UI，需基于此重构
  - `PaySignPrinter/app.py:63-96` — 初始化+token管理（适配飞书）
  - `PaySignPrinter/app.py:99-337` — 侧边栏+查询逻辑（适配飞书API）
  - `PaySignPrinter/app.py:362-463` — 批量处理UI（适配飞书）
  - `PaySignPrinter/app.py:601-786` — 单实例详情视图（适配飞书）

  **API/Type References**:
  - feishu_api.py 的所有公开函数
  - batch_processor.py 的 process_single_approval()

  **WHY Each Reference Matters**:
  - PaySignPrinter的UI布局已验证可用，保持相同的用户体验
  - 飞书API函数签名决定了UI调用方式

  **Acceptance Criteria**:

  - [ ] `streamlit run app.py --server.headless true --server.port 8501` 启动无报错
  - [ ] 侧边栏能输入日期范围和审批状态
  - [ ] 查询按钮能调用飞书搜索API获取审批列表
  - [ ] 审批列表能正确显示飞书审批状态
  - [ ] 附件下载无404错误

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Streamlit应用启动
    Tool: Bash + Playwright
    Preconditions: 所有依赖已安装
    Steps:
      1. streamlit run app.py --server.headless true --server.port 8501 &
      2. sleep 5
      3. curl -s http://localhost:8501 | head -20
      4. 检查页面包含 "飞书" 关键字
    Expected Result: 页面加载成功，包含飞书审批相关内容
    Failure Indicators: 启动报错或页面空白
    Evidence: .omo/evidence/task-14-app-start.png

  Scenario: 审批列表查询
    Tool: Playwright
    Preconditions: 应用已启动，有有效飞书凭证
    Steps:
      1. 在侧边栏选择日期范围
      2. 选择审批状态
      3. 点击查询按钮
      4. 等待审批列表加载
      5. 截图
    Expected Result: 审批列表显示至少1条记录
    Failure Indicators: 列表为空或报错
    Evidence: .omo/evidence/task-14-approval-list.png

  Scenario: 附件下载无404
    Tool: Playwright + Bash
    Preconditions: 有含附件的审批实例
    Steps:
      1. 选择含附件的审批实例
      2. 点击下载附件
      3. 检查下载目录中文件存在且大小>0
    Expected Result: 附件下载成功，无404错误
    Failure Indicators: 404错误或文件大小=0
    Evidence: .omo/evidence/task-14-download-fix.png
  ```

  **Commit**: YES (groups with Wave 3)
  - Message: `feat(feishu): add Streamlit UI with Feishu API integration`
  - Files: `app.py`
  - Pre-commit: `python -c "import app"`

- [ ] 15. 端到端集成测试

  **What to do**:
  - 编写端到端测试覆盖完整业务流程:
    1. 获取token → 查询审批列表 → 获取详情 → 解析表单 → 提取附件 → 下载附件 → 插入签名 → 打印
  - 使用mock+部分真实API混合测试:
    - token获取: 真实API调用
    - 列表查询: mock（避免依赖审批数据存在）
    - 详情获取: mock（使用Task 1验证的实际响应格式）
    - 下载: mock（避免大文件传输）
    - 签名插入: 真实文件操作（使用测试Excel和签名PNG）
    - 打印: mock（不实际打印）
  - 测试边界情况:
    - 无审批实例的日期范围
    - 含多个附件的实例
    - 附件中包含非Excel文件
    - Token过期后的自动刷新
    - 缓存命中和未命中

  **Must NOT do**:
  - 不要在实际打印机上打印
  - 不要上传/修改飞书上的真实审批数据

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: 集成测试需要精心设计mock+真实API混合策略
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 3 (after Task 14)
  - **Blocks**: F1-F4
  - **Blocked By**: Task 14

  **References**:

  **Pattern References**:
  - Task 1 验证结果 — 真实API响应格式用于构建mock
  - 所有前置Task的实现代码

  **Acceptance Criteria**:

  **If TDD:**
  - [ ] `python -m pytest test/test_integration.py -v` → PASS

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 完整业务流程E2E测试
    Tool: Bash (python)
    Preconditions: 所有模块已实现
    Steps:
      1. python -m pytest test/test_integration.py -v
      2. 验证: token获取→列表查询→详情→附件→下载→签名→打印 全链路PASS
    Expected Result: 所有集成测试PASS
    Failure Indicators: 任何链路断开
    Evidence: .omo/evidence/task-15-e2e-tests.txt

  Scenario: 边界情况测试
    Tool: Bash (python)
    Preconditions: 集成测试已编写
    Steps:
      1. python -m pytest test/test_integration.py -k "edge" -v
      2. 验证: 无实例、多附件、非Excel、token过期、缓存等边界情况PASS
    Expected Result: 边界测试全部PASS
    Failure Indicators: 任何边界情况未正确处理
    Evidence: .omo/evidence/task-15-edge-tests.txt
  ```

  **Commit**: YES (groups with Wave 3)
  - Message: `test(feishu): add end-to-end integration tests`
  - Files: `test/test_integration.py`
  - Pre-commit: `python -m pytest test/test_integration.py -v`

---

## Final Verification Wave (MANDATORY — after ALL implementation tasks)

> 4 review agents run in PARALLEL. ALL must APPROVE. Present consolidated results to user and get explicit "okay" before completing.

- [ ] F1. **Plan Compliance Audit** — `oracle`
  Read the plan end-to-end. For each "Must Have": verify implementation exists (read file, curl endpoint, run command). For each "Must NOT Have": search codebase for forbidden patterns — reject with file:line if found. Check evidence files exist in .omo/evidence/. Compare deliverables against plan.
  Output: `Must Have [N/N] | Must NOT Have [N/N] | Tasks [N/N] | VERDICT: APPROVE/REJECT`

- [ ] F2. **Code Quality Review** — `unspecified-high`
  Run the build, lint, and test commands. Review all changed files for: type suppression, empty catches, debug logging in prod, commented-out code, unused imports. Check AI slop: excessive comments, over-abstraction, generic names.
  Output: `Build [PASS/FAIL] | Lint [PASS/FAIL] | Tests [N pass/N fail] | Files [N clean/N issues] | VERDICT`

- [ ] F3. **Real Manual QA** — `unspecified-high` (+ `playwright` skill if UI)
  Start from clean state. Execute EVERY QA scenario from EVERY task — follow exact steps, capture evidence. Test cross-task integration. Test edge cases: empty state, invalid input. Save to `.omo/evidence/final-qa/`.
  Output: `Scenarios [N/N pass] | Integration [N/N] | Edge Cases [N tested] | VERDICT`

- [ ] F4. **Scope Fidelity Check** — `deep`
  For each task: read "What to do", read actual diff. Verify 1:1 — everything in spec was built, nothing beyond spec. Check "Must NOT do" compliance. Detect cross-task contamination.
  Output: `Tasks [N/N compliant] | Contamination [CLEAN/N issues] | Unaccounted [CLEAN/N files] | VERDICT`

---

## Commit Strategy

- **Wave 0**: `feat(feishu): verify API behavior and document actual responses` - feishu_api_verification.md
- **Wave 1**: `feat(feishu): add project scaffolding, auth, list query, config, cache` - feishu_api.py, test/, config files, API.md
- **Wave 2**: `feat(feishu): add detail parsing, download, approval status, signatures, printing` - batch updates
- **Wave 3**: `feat(feishu): add Streamlit UI and integration tests` - app.py, e2e tests

---

## Success Criteria

### Verification Commands
```bash
cd /home/ubuntu/coding/approval_feishu/approval_feishu
python -m pytest test/ -v              # Expected: All tests pass
python -c "from feishu_api import get_tenant_token; print(get_tenant_token()[:10])"  # Expected: t-xxx...
streamlit run app.py --server.headless true --server.port 8501  # Expected: App starts
```

### Final Checklist
- [ ] All "Must Have" present
- [ ] All "Must NOT Have" absent
- [ ] All tests pass
- [ ] Streamlit app starts and shows 飞书审批 UI
- [ ] Can query approval list from Feishu API
- [ ] Can download attachments without 404
- [ ] Can insert signatures into Excel
- [ ] Can print via WPS/Excel COM
