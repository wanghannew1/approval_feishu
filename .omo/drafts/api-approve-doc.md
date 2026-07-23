# API 6: 同意审批任务（出纳办理）

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
