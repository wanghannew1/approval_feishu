# 飞书 API 实际行为验证报告

> 验证时间: 2026-06-17
> 验证方式: 运行 test/test_api.py 捕获实际 API 响应

---

## 1. attachmentV2 值格式

**结论**: `attachmentV2` 控件的 `value` 返回的是**完整临时下载 URL**（12小时有效），而非 `file_token`。

**实际响应**:
```json
{
  "type": "attachmentV2",
  "name": "附件（东软系统工资表）",
  "value": [
    "https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/authcode/?code=NmZhY2NlMzU5ZDllODM0Y2I2ZGYzMmNmYTQ0MGI2NWJfOGY1NmJmZDA3MWM1ZTI1M2NmZTkzMTEzYjdiZjgzMWNfSUQ6NzY1MDA1MjgyNDQ0MDY3MTIwOF8xNzgxNjA3OTcwOjE3ODE2OTQzNzBfVjM"
  ]
}
```

**影响**: 下载附件时不应使用 `drive/v1/files/{file_token}/download` 端点，而应直接 GET 该 URL（带 Authorization header）。

---

## 2. approver_list 字段结构

**结论**: 飞书审批详情返回的 `approver_list` **不包含 user_id/open_id**，只有 `approver_name`、`status`、`comment`。

**实际响应**:
```json
{
  "approver_list": [
    {
      "approver_name": "张三",
      "status": "APPROVED",
      "comment": "同意"
    }
  ]
}
```

**缺失字段**: 无 `approver_id`、`user_id`、`open_id` 等用户标识。

**影响**: 签名映射**无法使用用户ID**，必须改用 `approver_name` 作为签名查找的键。需要构建 name → role 的映射。

---

## 3. 搜索/查询 API 端点

**结论**: `/instances/search` 端点需要 `user_id` 且不接受空字符串，应改用 `/instances/query`。

**实际行为**:
- `POST /instances/search` → `99992402 field validation failed` (缺少 user_id/offset/limit/sort_asc)
- `POST /instances/search` + user_id="" → `60001 request param miss`
- `POST /instances/query` + page_size/page_token → `code=0 成功`

**注意**: `/query` 需要 `approval:approval.list:readonly` 权限。当前应用无此权限时返回 `99991672 Access denied`，但不影响 `GET /instances` 列表查询。

---

## 4. Token 管理

**结论**: `tenant_access_token` 有效期 7200 秒，获取新 token 时旧 token 立即失效。

**实际行为**:
- 缓存文件 `.token_cache.json` 正常工作
- 5分钟提前过期策略有效

---

## 5. 审批状态值

**实际状态值**: `PENDING`, `APPROVED`, `REJECTED`, `CANCELED`

**与钉钉映射**:
| 钉钉 | 飞书 |
|------|------|
| COMPLETED | APPROVED |
| RUNNING | PENDING |
| TERMINATED | REJECTED |

---

## 迁移决策影响

1. **附件下载**: 使用 URL 直接下载，不通过 drive API
2. **签名映射**: 使用 `approver_name` 而非 user_id
3. **列表查询**: 主用 `GET /instances`，辅用 `POST /query`
4. **审批状态**: 直接检查 `detail["status"]` 为 APPROVED
5. **is_ready_for_print**: 需遍历 `approver_list`，按 name 映射角色，检查 mandatory 角色全部 APPROVED
