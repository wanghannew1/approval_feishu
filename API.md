# 飞书审批 API 接口文档

> 对应项目：approval_feishu（飞书审批打印工具）
> 飞书开放平台文档：https://open.feishu.cn/document/server-docs/approval-v4/

## 环境变量

| 变量 | 说明 |
|------|------|
| `FEISHU_APP_ID` | 飞书应用 App ID |
| `FEISHU_APP_SECRET` | 飞书应用 App Secret |
| `FEISHU_APPROVAL_CODE` | 审批定义 Code（格式如 `EB828003-9FFE-4B3F-AA50-2E199E2ED942`） |

## 通用约定

- **Base URL**: `https://open.feishu.cn/open-apis`
- **认证**: 所有 API 需在 Header 中携带 `Authorization: Bearer <tenant_access_token>`
- **时间戳**: 统一使用毫秒（ms）时间戳
- **成功响应码**: `code: 0`

---

## API 1: 获取 Tenant Access Token

**用途**：获取租户访问令牌，所有后续 API 需要此 token。

**请求**：
```
POST /open-apis/auth/v3/tenant_access_token/internal/
Content-Type: application/json; charset=utf-8

{
    "app_id": "<FEISHU_APP_ID>",
    "app_secret": "<FEISHU_APP_SECRET>"
}
```

**响应**：
```json
{
    "code": 0,
    "msg": "ok",
    "tenant_access_token": "t-xxx...",
    "expire": 7200
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `tenant_access_token` | string | 访问令牌 |
| `expire` | int | 有效期（秒），默认 7200（2小时） |

**错误码**：
| code | 说明 |
|------|------|
| 1 | App ID 或 App Secret 错误 |

---

## API 2: 查询审批实例列表

**用途**：根据审批定义 Code 和时间范围，获取该定义下所有审批实例的 `instance_code` 列表。

**请求**：
```
GET /open-apis/approval/v4/instances?approval_code=<CODE>&start_time=<MS>&end_time=<MS>&page_size=100
Authorization: Bearer <tenant_access_token>
```

**查询参数**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `approval_code` | string | ✅ | 审批定义 Code |
| `start_time` | string | ✅ | 审批创建开始时间，毫秒时间戳 |
| `end_time` | string | ✅ | 审批创建结束时间，毫秒时间戳 |
| `page_size` | int | ❌ | 分页大小，1-100，默认 100 |
| `page_token` | string | ❌ | 分页标记（翻页时传入上一次返回的 page_token） |

**响应**：
```json
{
    "code": 0,
    "msg": "success",
    "data": {
        "instance_code_list": [
            "357C21A0-2069-4F6B-955F-1DFBE6710C51",
            "A1B2C3D4-5678-90AB-CDEF-1234567890AB"
        ],
        "page_token": "nF1ZXJ5VGhlbkZldGNoCgAAAAAA6PZwFmUzSldvTC1yU",
        "has_more": false
    }
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `instance_code_list` | string[] | 审批实例 Code 列表 |
| `page_token` | string | 下一页分页标记，`has_more=false` 时为空 |
| `has_more` | bool | 是否还有更多数据 |

**分页循环**：
```python
page_token = None
all_codes = []
while True:
    params = {"approval_code": code, "start_time": start_ms, "end_time": end_ms}
    if page_token:
        params["page_token"] = page_token
    resp = requests.get(url, headers=headers, params=params)
    data = resp.json()
    all_codes.extend(data["data"]["instance_code_list"])
    if not data["data"]["has_more"]:
        break
    page_token = data["data"]["page_token"]
```

---

## API 3: 获取审批实例详情

**用途**：获取单个审批实例的完整信息，包括表单数据、审批状态、审批人、附件等。

**请求**：
```
GET /open-apis/approval/v4/instances/<instance_code>
Authorization: Bearer <tenant_access_token>
```

**路径参数**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `instance_code` | string | ✅ | 审批实例 Code（从 API 2 获取） |

**响应**：
```json
{
    "code": 0,
    "msg": "success",
    "data": {
        "instance_code": "81D31358-93AF-92D6-7425-01A5D67C4E71",
        "approval_name": "工资审批",
        "status": "APPROVED",
        "start_time": "1698316700000",
        "end_time": "1698316800000",
        "form": "[{\"id\":\"widget1\",\"name\":\"工资表\",\"type\":\"attachmentV2\",\"value\":[\"boxbc_xxx\"]}]",
        "approver_list": [
            {
                "approver_name": "张三",
                "status": "APPROVED",
                "comment": "同意"
            }
        ],
        "cc_list": [],
        "instance_url": "https://..."
    }
}
```

**关键字段**：

| 字段 | 类型 | 说明 |
|------|------|------|
| `status` | string | `PENDING` 审批中 / `APPROVED` 已通过 / `REJECTED` 已拒绝 / `CANCELED` 已撤销 |
| `form` | string | 表单数据的 JSON 字符串，需手动 `json.loads()` |
| `form[].type` | string | 控件类型，文件附件为 `attachmentV2` |
| `form[].value` | string[] | 附件控件值为**完整临时下载 URL**（12小时有效），如 `https://internal-api-drive-stream.feishu.cn/...` |
| `approver_list[].approver_name` | string | 审批人姓名 |
| `approver_list[].status` | string | 审批状态（APPROVED/REJECTED/PENDING） |
| `approver_list[].comment` | string | 审批意见 |

> ⚠️ `approver_list` **不包含 user_id/open_id**，只有 `approver_name`、`status`、`comment`。签名映射需使用姓名作为键。

**解析 form 中的附件**：
```python
form_data = json.loads(data["data"]["form"])
attachments = []
for widget in form_data:
    if widget["type"] == "attachmentV2":
        for url in widget["value"]:
            attachments.append({
                "field_name": widget.get("name", "附件"),
                "download_url": url  # 完整临时 URL，12小时有效
            })
```

> 📌 与钉钉的区别：飞书的 `form` 是 JSON 字符串，需要 `json.loads()` 解析。钉钉的 `formComponentValues` 是直接的对象数组。

---

## API 4: 下载文件

**用途**：下载审批附件。`attachmentV2` 控件的 `value` 是完整临时下载 URL（12小时有效），可直接请求。

### 方式一：直接使用临时 URL（推荐）

**请求**：
```
GET <attachmentV2 中的临时 URL>
Authorization: Bearer <tenant_access_token>
```

**说明**：
- URL 格式如 `https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/authcode/?code=...`
- 有效期 12 小时
- 必须携带 Authorization header

**下载代码**：
```python
download_url = widget["value"][0]  # 直接是完整 URL
resp = requests.get(download_url, headers=headers, stream=True)
resp.raise_for_status()
content_disp = resp.headers.get("Content-Disposition", "")
file_name = extract_filename(content_disp) or "attachment.xlsx"
with open(file_name, "wb") as f:
    for chunk in resp.iter_content(chunk_size=8192):
        f.write(chunk)
```

### 方式二：通过 Drive API

**请求**：
```
GET /open-apis/drive/v1/files/<file_token>/download
Authorization: Bearer <tenant_access_token>
```

**路径参数**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `file_token` | string | ✅ | 文件 token（从旧版 form 中提取） |

**响应**：
- 成功：返回文件的二进制流（`Content-Type` 根据文件类型变化）
- 响应头可能包含 `Content-Disposition` 指示文件名
- 可用 `Range` 请求头支持断点续传

**权限要求**：`drive:drive` 或 `drive:file:download`

---

## API 5: 按状态搜索审批实例（可选）

**用途**：比 API 2 更灵活的查询，可直接按状态筛选并支持分页。

> 📌 此接口需要应用拥有 `approval:approval.list:readonly` 权限。
> 如遇到 99991672 错误，请在飞书开放平台添加该权限。

**请求**：
```
POST /open-apis/approval/v4/instances/query
Authorization: Bearer <tenant_access_token>
Content-Type: application/json; charset=utf-8

{
    "approval_code": "<CODE>",
    "instance_status": "APPROVED",
    "instance_start_time_from": "<MS>",
    "instance_start_time_to": "<MS>",
    "page_size": 100,
    "page_token": ""
}
```

**请求体参数**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `approval_code` | string | ✅ | 审批定义 Code |
| `instance_status` | string | ❌ | `PENDING`/`APPROVED`/`REJECTED`/`CANCELED`，不传=全部 |
| `instance_start_time_from` | string | ❌ | 开始时间（ms） |
| `instance_start_time_to` | string | ❌ | 结束时间（ms） |
| `page_size` | int | ❌ | 分页大小，5-200，默认 10 |
| `page_token` | string | ❌ | 分页标记 |

**响应**（与 API 2 不同，直接返回实例摘要）：
```json
{
    "code": 0,
    "msg": "success",
    "data": {
        "instance_list": [
            {
                "instance_code": "xxx",
                "status": "APPROVED",
                "start_time": "1698316700000"
            }
        ],
        "has_more": false,
        "page_token": ""
    }
}
```

> 📌 此接口优势：可一步完成"筛选已通过 + 获取摘要"，无需先调 API 2 拿 code 列表再逐个调 API 3。

---

## 完整调用链

```
启动应用
  └── get_tenant_token()        (1) POST /auth/v3/tenant_access_token/internal/

查询审批列表
  ├── query_instances()           (2) GET /approval/v4/instances
  │     └── 遍历所有 instance_code:
  │         └── get_instance_detail()  (3) GET /approval/v4/instances/{id}
  │               └── parse_form() 提取附件 file_token

下载附件
  └── download_file()             (4) GET /drive/v1/files/{token}/download
```

---

## 钉钉 → 飞书 字段对照

| 钉钉 | 飞书 | 说明 |
|------|------|------|
| `processCode` | `approval_code` | 审批定义 ID |
| `instanceId` | `instance_code` | 审批实例 ID |
| `businessId` | `instance_code` | 业务 ID（飞书无独立 businessId） |
| `formComponentValues` | `form` (JSON 字符串) | 表单数据 |
| `title` | `approval_name` | 审批标题 |
| `COMPLETED` | `APPROVED` | 已通过的审批 |
| `originatorId` | 从 `approver_list[0]` 推断 | 发起人信息 |
| `x-acs-dingtalk-access-token` | `Authorization: Bearer` | 鉴权方式 |

## 注意事项

1. **Token 有效期**: 7200 秒（2 小时），过期前建议提前刷新
2. **Token 并发安全**: 获取新 token 时旧 token 会立即失效，建议缓存 + 加锁
3. **分页**: API 2 仅返回 `instance_code`，无标题/状态摘要，翻页需用 `page_token`
4. **文件下载**: `attachmentV2` 控件的 `value` 是完整临时下载 URL（12小时有效），可直接 GET 请求下载，无需调用 Drive API
5. **API 5 权限**: `/instances/query` 需要 `approval:approval.list:readonly` 权限，无此权限返回 99991672
6. **时间范围**: API 2 的时间范围无 120 天限制（与钉钉不同），但建议控制分页量
7. **审批状态值**: `PENDING` 审批中, `APPROVED` 已通过, `REJECTED` 已拒绝, `CANCELED` 已撤销
8. **approver_list 字段**: 仅包含 `approver_name`、`status`、`comment`，**不包含 user_id/open_id**，签名映射需使用姓名作为键

---

## API 6: 同意审批任务（出纳办理）

**用途**：出纳办结当前审批节点（如"出纳办理"），使审批流转到下一步。

> ⚠️ **手写签名限制**：如果审批节点要求手写签名（如"总经理签字"），API 会返回 `1390018 not support handwritten signature`，必须去飞书客户端操作。仅"出纳办理"这类非签名节点可用 API 审批。

**请求**：
```
POST /open-apis/approval/v4/tasks/approve
Authorization: Bearer <tenant_access_token>
Content-Type: application/json; charset=utf-8
```

**查询参数**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `user_id_type` | string | ❌ | 用户 ID 类型，默认 `open_id`。可选值：`open_id` / `union_id` / `user_id` |

**请求体**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `approval_code` | string | ✅ | 审批定义 Code |
| `instance_code` | string | ✅ | 审批实例 Code |
| `user_id` | string | ✅ | 审批人的用户 ID，类型与 `user_id_type` 一致。**必填 open_id 格式（`ou_...`）**，不要传短 `user_id` |
| `task_id` | string | ✅ | 审批任务 ID，从实例详情的 `task_list[].id` 获取 |
| `comment` | string | ❌ | 审批意见 |
| `form` | string | ❌ | 条件分支控件数据（JSON 字符串），无条件分支可不传 |

**请求体示例**：
```json
{
    "approval_code": "6636C5FC-2A9F-4A38-A7EE-21EE14FB703C",
    "instance_code": "5C048511-D345-42D4-834F-FA4E41EB6B60",
    "user_id": "ou_f5b63c60c9aee6f27a4d68671e025d2c",
    "task_id": "7662630202672974823",
    "comment": "出纳办理完成"
}
```

**响应示例（成功）**：
```json
{
    "code": 0,
    "msg": "success",
    "data": {}
}
```

**参数获取流程**：

```
获取实例详情 GET /approval/v4/instances/{instance_code}
   ↓
从 task_list 中找到 node_name="出纳办理" 且 status="PENDING" 的那个
   ↓
取出:
  task_id  = task["id"]          ← 审批任务 ID
  open_id  = task["open_id"]     ← 审批人的 open_id（传给 user_id 参数）
```

**关键字段在 task_list 中的位置**：

```json
{
    "task_list": [
        {
            "id": "7662630202672974823",      ← task_id
            "node_name": "出纳办理",           ← 审批节点名称
            "status": "PENDING",               ← 待审批状态
            "open_id": "ou_f5b63c...",         ← 审批人 open_id（传给 user_id）
            "user_id": "a493e65d",             ← 审批人短 ID（API 不可用）
            "node_id": "c60e0b98..."           ← 节点定义 ID（无关）
        }
    ]
}
```

> ⚠️ **注意**：`user_id` 参数必须传 `open_id`（`ou_...` 格式），传短 `user_id`（如 `a493e65d`）会报 `99992351` 错误。

**错误码**：

| HTTP 状态码 | 错误码 | 描述 | 排查建议 |
|-------------|--------|------|---------|
| 400 | 1390001 | param is invalid | 参数错误，检查请求参数 |
| 400 | 1390002 | approval code not found | 审批定义 Code 错误 |
| 400 | 1390003 | instance code not found | 审批实例 Code 错误 |
| 400 | 1390010 | task not found | 审批任务 ID 错误（task_id） |
| 400 | 1390018 | not support handwritten signature | **该节点需要手写签名，无法通过 API 审批**，必须去飞书客户端操作 |
| 400 | 99992351 | invalid user_id | user_id 格式错误，必须传 open_id（`ou_...`） |

**权限要求**（开启任一即可）：
- `approval:approval` — 查看、创建、更新、删除审批应用相关信息
- `approval:approval:readonly` — 访问审批应用
- `approval:task` — 同意、拒绝、退回、加签等原生审批任务操作

---

## 完整调用链（更新版）

```
启动应用
  └── get_tenant_token()        (1) POST /auth/v3/tenant_access_token/internal/

查询审批列表
  ├── query_instances()           (2) GET /approval/v4/instances
  │     └── 遍历所有 instance_code:
  │         └── get_instance_detail()  (3) GET /approval/v4/instances/{id}
  │               └── parse_form() 提取附件 URL

下载附件
  └── download_file()             (4) GET 附件临时 URL

签名打印
  ├── process_single_approval()   (3) + (4) + 签名插入 + 打印

出纳办理 （新增）
  └── approve_task()              (6) POST /approval/v4/tasks/approve
        ├── 从 task_list 获取 task_id、open_id
        └── 调用 approve API 完成出纳办理
```
