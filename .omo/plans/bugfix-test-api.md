# test_api.py Bug 修复

## TL;DR

> **Quick Summary**: 修复 test_api.py 中两个 bug — 附件下载 404 和搜索 API 缺少必填参数
>
> **Deliverables**:
> - test_api.py 中两个 bug 修复
> - API.md 中搜索 API 参数文档修正
>
> **Estimated Effort**: Quick
> **Parallel Execution**: NO - 两个修复有依赖（先修附件下载，验证后再修搜索）
> **Critical Path**: Task 1 → Task 2

---

## Context

### Bug 1: 附件下载 404
**位置**: `test_api.py:148-152` + `test_api.py:166-169`

**根因**: `attachmentV2` 组件的 `value` 实际返回的是完整临时下载 URL（如 `https://internal-api-drive-stream.feishu.cn/...?code=xxx`），而不是 `file_token`（如 `boxbc_xxx`）。代码把完整 URL 当 file_token 拼到 drive API 路径上，产生双重 URL。

用户测试输出的错误 URL:
```
https://open.feishu.cn/open-apis/drive/v1/files/https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/authcode/?code=...
```

**飞书官方文档确认**: 审批附件返回的 URL 是临时下载链接（12小时有效），应直接使用，不是 drive file_token。

**修复方案**: `test_download_file()` 需要自动检测值格式 — 如果是完整 URL 就直接下载；如果是 file_token 就走 drive API 下载。两种情况都要带 `Authorization: Bearer` header。

### Bug 2: 搜索 API 缺少必填参数
**位置**: `test_api.py:195-216`

**根因**: `/approval/v4/instances/search` POST 端点需要 `user_id`, `offset`, `limit`, `sort_asc`，但代码只发送了 `approval_code`, 时间范围, `page_size`。

**修复方案**: 添加所有必填参数到请求体。

---

## Work Objectives

### Core Objective
修复 test_api.py 的两个 bug，使其能成功下载附件和搜索审批实例。

### Concrete Deliverables
- test_api.py 中的 `test_get_instance_detail()` 附件提取逻辑修复
- test_api.py 中的 `test_download_file()` 下载逻辑修复
- test_api.py 中的 `test_search_instances()` 必填参数补全
- API.md 中搜索 API 参数文档修正

### Definition of Done
- [x] `python test_api.py` 全部5个测试通过，无 404 和校验失败

### Must Have
- 附件下载兼容 URL 和 file_token 两种格式
- 搜索 API 包含所有必填参数
- 下载文件名正确提取

### Must NOT Have (Guardrails)
- 不修改 token 获取逻辑（测试1已通过）
- 不修改列表查询逻辑（测试2已通过）
- 不修改主流程 main() 结构
- 不修改钉钉 PaySignPrinter 项目任何文件

---

## Verification Strategy

### Test Decision
- **Infrastructure exists**: NO (无 pytest)
- **Automated tests**: None（直接运行 test_api.py 验证）
- **Framework**: 手动运行验证

### QA Policy
直接运行 `python test_api.py` 验证所有5个测试通过。

---

## Execution Strategy

### Parallel Execution Waves

```
Wave 1 (sequential - two small fixes):
├── Task 1: 修复附件下载 Bug (test_download_file + test_get_instance_detail) [quick]
├── Task 2: 修复搜索 API 必填参数 (test_search_instances) [quick]
└── Task 3: 修正 API.md 搜索参数文档 [quick]

Wave FINAL:
└── Task F1: 运行完整 test_api.py 验证 [quick]
```

### Dependency Matrix

| Task | Depends On | Blocks | Wave |
|------|-----------|--------|------|
| 1 | - | F1 | 1 |
| 2 | - | F1 | 1 |
| 3 | - | F1 | 1 |

### Agent Dispatch Summary

- **Wave 1**: **3** — T1 → `quick`, T2 → `quick`, T3 → `quick`
- **FINAL**: **1** — F1 → `quick`

---

## TODOs

- [x] 1. 修复附件下载 404 Bug

  **What to do**:
  - 修改 `test_download_file()` 函数（line 166-192）:
    - 自动检测 `file_token_or_url` 参数的格式
    - 如果以 `http` 开头 → 直接 GET 该 URL（带 Authorization header）
    - 否则 → 使用原 drive API 路径 `DRIVE_DOWNLOAD_URL.format(file_token=xxx)`
    - 两种路径都带 `Authorization: Bearer` header，都用 `stream=True`
  - 修改 `test_get_instance_detail()` 中附件提取（line 148-152）:
    - 当前代码直接将 `ft` 作为 `file_token` 存入 dict
    - 修改为：存为 `"file_token_or_url": ft`（更准确地描述该值可能是 URL 也可能是 token）
  - 修改 `main()` 中调用下载的代码（line 296-298）:
    - 将 `att["file_token"]` 改为 `att["file_token_or_url"]`
  - 修改 `main()` 中显示附件信息的代码（line 283）:
    - 将 `att['file_token'][:20]` 改为 `att['file_token_or_url'][:20]`

  **具体代码修改**:

  1. `test_download_file` 函数改为：
  ```python
  def test_download_file(file_token_or_url: str, save_dir: Path) -> Path:
      """测试 4: 下载文件到本地（自动检测 URL 或 file_token 格式）。"""
      headers = auth_headers()

      # 自动检测：如果是完整 URL 直接下载，否则走 drive API
      if file_token_or_url.startswith("http"):
          url = file_token_or_url
      else:
          url = DRIVE_DOWNLOAD_URL.format(file_token=file_token_or_url)

      resp = requests.get(url, headers=headers, stream=True)
      resp.raise_for_status()
      # ... 后续文件名提取和保存逻辑不变
  ```

  2. 附件提取改为：
  ```python
  for ft in widget.get("value", []):
      attachments.append({
          "field_name": widget.get("name", "附件"),
          "file_token_or_url": ft,
      })
  ```

  3. main() 中对应引用改为 `att["file_token_or_url"]`

  **Must NOT do**:
  - 不删除 DRIVE_DOWNLOAD_URL 常量（file_token 格式仍然需要）
  - 不修改 main() 的整体流程结构
  - 不添加新的依赖

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 3处小改动，逻辑清晰
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3)
  - **Blocks**: F1
  - **Blocked By**: None

  **References**:

  **Pattern References**:
  - `test_api.py:166-192` — 当前 test_download_file 函数（需修改）
  - `test_api.py:148-152` — 当前附件提取逻辑（需修改 key 名）
  - `test_api.py:283,296-298` — main() 中引用 file_token 的地方

  **External References**:
  - 飞书审批附件文档: https://open.feishu.cn/document/server-docs/approval-v4/instance/get — 确认 attachmentV2 返回临时 URL

  **Acceptance Criteria**:

  - [ ] `python test_api.py` 测试4（下载附件）不再报 404
  - [ ] 下载的文件大小 > 0

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 附件下载成功（URL 格式）
    Tool: Bash (python)
    Preconditions: .env 中有正确的飞书凭证，存在含附件的审批实例
    Steps:
      1. cd /home/ubuntu/coding/approval_feishu/approval_feishu
      2. python test_api.py
      3. 检查 "测试 4" 输出是否为 ✅ 成功
    Expected Result: 下载成功，文件名和大小正确显示
    Failure Indicators: 仍然 404 或其他 HTTP 错误
    Evidence: .omo/evidence/task-1-download-fix.txt

  Scenario: 附件下载成功（file_token 格式兼容）
    Tool: Bash (python -c)
    Preconditions: test_download_file 函数已修改
    Steps:
      1. python -c "from test_api import test_download_file; print('import OK')"
      2. 验证函数签名接受 file_token_or_url 参数
    Expected Result: 导入成功，无语法错误
    Failure Indicators: ImportError 或 SyntaxError
    Evidence: .omo/evidence/task-1-import-check.txt
  ```

  **Commit**: YES
  - Message: `fix(test_api): attachment download - auto-detect URL vs file_token format`
  - Files: `test_api.py`
  - Pre-commit: `python -c "from test_api import test_download_file"`

- [x] 2. 修复搜索 API 必填参数 Bug

  **What to do**:
  - 修改 `test_search_instances()` 函数（line 195-216）:
    - 添加必填参数: `user_id`, `offset`, `limit`, `sort_asc`
    - `user_id` 使用默认值（可后续配置）
    - `offset` 默认 0
    - `limit` 默认 50（替代 page_size）
    - `sort_asc` 默认 False（降序，最新优先）
  - 修改函数签名添加 user_id 参数

  **具体代码修改**:

  ```python
  def test_search_instances(status: str | None = None, user_id: str = "") -> dict:
      """测试 5: 按状态搜索审批实例。"""
      headers = auth_headers()
      end_ms = str(int(datetime.now().timestamp() * 1000))
      start_ms = str(int((datetime.now() - timedelta(days=30)).timestamp() * 1000))

      body = {
          "approval_code": APPROVAL_CODE,
          "user_id": user_id,
          "offset": 0,
          "limit": 50,
          "sort_asc": False,
          "instance_start_time_from": start_ms,
          "instance_start_time_to": end_ms,
      }
      if status:
          body["instance_status"] = status

      resp = requests.post(SEARCH_URL, headers=headers, json=body)
      data = resp.json()

      if data.get("code") != 0:
          raise RuntimeError(f"搜索实例失败: {data}")

      return data["data"]
  ```

  > **注意**: `user_id` 如果传空字符串仍然可能校验失败。如果仍然报错，尝试不传 `user_id` 或使用 `open_id` 格式的实际用户 ID。需要实测确认飞书对此字段的具体要求。

  **Must NOT do**:
  - 不删除 `page_size` 字段如果搜索 API 也支持它
  - 不修改函数返回格式

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 添加几个必填参数
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 3)
  - **Blocks**: F1
  - **Blocked By**: None

  **References**:

  **Pattern References**:
  - `test_api.py:195-216` — 当前 test_search_instances 函数
  - 用户测试输出中的错误信息: `field_violations: user_id is required, offset is required, limit is required, sort_asc is required`

  **Acceptance Criteria**:

  - [ ] `python test_api.py` 测试5（搜索实例）不再报 field validation failed

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: 搜索API参数校验通过
    Tool: Bash (python)
    Preconditions: test_search_instances 已修改
    Steps:
      1. cd /home/ubuntu/coding/approval_feishu/approval_feishu
      2. python test_api.py
      3. 检查 "测试 5" 输出是否为 ✅ 成功
    Expected Result: 搜索成功返回实例列表
    Failure Indicators: 仍然报 field validation failed
    Evidence: .omo/evidence/task-2-search-fix.txt
  ```

  **Commit**: YES
  - Message: `fix(test_api): add required fields for search API (user_id, offset, limit, sort_asc)`
  - Files: `test_api.py`
  - Pre-commit: `python -c "from test_api import test_search_instances"`

- [x] 3. 修正 API.md 搜索参数文档

  **What to do**:
  - 在 API.md 的 API 5 搜索部分（line 245-254）添加缺失的必填参数:
    - `user_id` (string, 必填) — 查询用户 ID
    - `offset` (int, 必填) — 分页偏移
    - `limit` (int, 必填) — 分页大小
    - `sort_asc` (bool, 必填) — 是否升序排列
  - 更新请求体示例 JSON 包含这些字段

  **Must NOT do**:
  - 不修改其他 API 章节内容

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 文档更新
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 2)
  - **Blocks**: F1
  - **Blocked By**: None

  **References**:
  - `API.md:225-276` — 搜索 API 文档部分
  - 用户测试错误信息确认的必填参数

  **Acceptance Criteria**:
  - [ ] API.md 搜索 API 参数表包含 user_id, offset, limit, sort_asc

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: API.md搜索参数完整性
    Tool: Bash (grep)
    Preconditions: API.md已更新
    Steps:
      1. grep "user_id" API.md | wc -l — 确认搜索部分包含 user_id
      2. grep "offset" API.md | wc -l — 确认包含 offset
      3. grep "limit" API.md | wc -l — 确认包含 limit
      4. grep "sort_asc" API.md | wc -l — 确认包含 sort_asc
    Expected Result: 每个字段至少匹配1次
    Failure Indicators: 字段缺失
    Evidence: .omo/evidence/task-3-api-md-check.txt
  ```

  **Commit**: YES
  - Message: `docs(API): add missing required fields for search API`
  - Files: `API.md`
  - Pre-commit: 无

---

## Final Verification Wave

- [x] F1. 运行完整 test_api.py 验证 — `quick`

  运行 `python test_api.py`，确认全部5个测试通过：
  1. ✅ Token 获取成功
  2. ✅ 审批列表查询成功
  3. ✅ 实例详情 + 附件解析成功
  4. ✅ 附件下载成功（无 404）
  5. ✅ 搜索实例成功（无 field validation failed）

  **QA Scenarios:**

  ```
  Scenario: 全量测试通过
    Tool: Bash
    Preconditions: Bug 1 和 Bug 2 已修复
    Steps:
      1. cd /home/ubuntu/coding/approval_feishu/approval_feishu
      2. python test_api.py
    Expected Result: 5/5 测试全部 ✅
    Failure Indicators: 任何测试 ❌
    Evidence: .omo/evidence/f1-full-test.txt
  ```

---

## Commit Strategy

- **Task 1+2+3**: `fix(test_api): resolve attachment 404 and search API validation errors` - test_api.py, API.md
- Pre-commit: `python -c "from test_api import test_download_file, test_search_instances"`

---

## Success Criteria

### Verification Commands
```bash
cd /home/ubuntu/coding/approval_feishu/approval_feishu
python test_api.py  # Expected: 5/5 测试通过
```

### Final Checklist
- [x] 附件下载无 404 错误
- [x] 搜索 API 无 field validation failed
- [x] API.md 搜索参数文档已修正
