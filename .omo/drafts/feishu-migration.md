# Draft: 飞书审批附件工资表打印 — 钉钉→飞书迁移

## Requirements (confirmed)
- 源项目: PaySignPrinter (钉钉审批附件工资表打印)
- 目标: 改造为飞书API版
- 参考代码: larksuite-cli/cli (飞书官方CLI)
- 当前测试脚本 test_api.py 有两个bug需要修复

## Bug Reports (from user test)
1. **附件下载 404**: 把完整内部URL当file_id拼到drive API路径上
   - 错误URL: `https://open.feishu.cn/open-apis/drive/v1/files/https://internal-api-drive-stream.feishu.cn/...`
   - 应该: 正确使用飞书审批附件下载API
2. **搜索实例校验失败**: 缺少必填字段 user_id, offset, limit, sort_asc

## Research Findings

### PaySignPrinter 项目结构 (钉钉原版)
- **路径**: `/home/ubuntu/coding/approval_feishu/PaySignPrinter/`
- **入口**: `app.py` (787行, Streamlit Web UI)
- **API层**: `dingtalk_api.py` (242行, 5个钉钉API端点)
- **缓存层**: `cache_manager.py` (204行, 3层缓存+15分钟下载URL TTL)
- **业务层**: `batch_processor.py` (671行, 批量下载/签名插入/打印)
- **配置**: .env(钉钉凭证), settings.json, role_mapping.json, user_mapping.json
- **签名目录**: signatures/ (userId.png)
- **下载目录**: downloads/
- **无测试文件**

#### 钉钉API映射 (5个端点):
1. `POST /v1.0/oauth2/accessToken` → 认证
2. `POST /v1.0/workflow/processes/instanceIds/query` → 列表查询
3. `GET /v1.0/workflow/processInstances` → 实例详情+表单
4. `POST /v1.0/workflow/processInstances/spaces/files/urls/download` → 获取下载URL
5. 直接GET下载URL → 文件下载

#### 核心业务流程:
1. 获取审批实例列表 → 获取实例详情 → 解析表单提取附件ID
2. 通过附件ID获取临时下载URL (15分钟有效)
3. 下载Excel工资表附件
4. 解析Excel查找签名位置 (总经理签字/部长签字/财务审核/业务审核)
5. 根据审批人角色映射 + userId.png 签名图片 → 插入Excel
6. 跨平台打印: Windows用WPS/Excel COM, Linux用LibreOffice

### larksuite-cli 参考代码
- **路径**: `/home/ubuntu/coding/approval_feishu/larksuite-cli/cli/` (Go项目)
- 认证: `POST /open-apis/auth/v3/tenant_access_token/internal`
- Drive下载: `GET /open-apis/drive/v1/files/{file_token}/download` (需stream=True)
- **没有**审批实例列表查询的直接实现
- 参考价值: 认证流程、API客户端模式、错误处理

### Bug 1: 附件下载404 (根因已确认)
- `attachmentV2` 组件的 `value` 返回的是**完整临时下载URL** (12小时有效)
  而非 file_token (如 `boxbc_xxx`)
- 代码把完整URL当file_token拼到 `/drive/v1/files/{file_token}/download` → 产生双重URL
- **正确做法**: 直接使用附件URL下载, 不要通过drive API
  (飞书官方: "每次获取审批详情都会获得新的url, url有效期12小时")

### Bug 2: 搜索实例校验失败 (根因已确认)
- `/approval/v4/instances/search` POST 端点要求 `user_id`, `offset`, `limit`, `sort_asc`
- 代码只发送了 `approval_code`, 时间范围, `page_size`
- 需要添加缺失的必填字段

## Technical Decisions
- **迁移策略**: 完全替换钉钉代码，只保留飞书版
- **Bug处理**: 整体迁移一起修，在飞书API模块中直接实现正确逻辑
- **运行环境**: Windows + WPS/Excel (COM打印)
- **附件缓存**: 使用飞书临时URL(12h有效)，每次获取详情时重新获取
- **认证方式**: tenant_access_token (飞书自定义应用)
- **UI框架**: 继续用Streamlit
- **测试策略**: TDD — 先写测试再写实现
- **迁移范围**: 完整功能复刻（审批列表、详情、附件下载、签名插入、批量打印）
- **只处理工资表审批**: approval_code=1CF34ABB, 不处理其他审批类型
- **Token自动刷新**: 需要实现, 参考test_api.py的5分钟提前过期缓存模式
- **[CRITICAL] 不得修改PaySignPrinter钉钉项目任何文件** — 独立仓库
- **[CRITICAL] 飞书项目另起仓库** — 不与钉钉项目共享

## Scope Boundaries
- INCLUDE:
  - 飞书API完整替换钉钉（feishu_api.py 新建，不修改 dingtalk_api.py）
  - 审批列表查询（搜索API + 列表API）
  - 审批实例详情获取
  - 附件下载（修复Bug 1: 正确处理attachmentV2值）
  - Excel签名插入（复用openpyxl逻辑）
  - 批量打印（Windows COM方式）
  - 缓存管理（适配飞书API，12h URL有效期）
  - Streamlit UI重构（飞书表单解析、审批人显示）
  - TDD单元测试（pytest + requests-mock）
  - 修正API.md文档中与实际API不符的参数描述
  - 新建 role_mapping.json（飞书审批节点名称）
  - 新建 user_mapping.json（飞书用户ID）
- EXCLUDE:
  - **[CRITICAL] 不得修改PaySignPrinter钉钉项目任何文件** — 独立仓库
  - **[CRITICAL] 飞书项目另起仓库** — 不与钉钉项目共享
  - 钉钉双平台保留
  - 非工资表审批类型处理（只处理1CF34ABB工资表审批）
  - 移动端适配
  - 多租户支持
  - 用户认证/登录系统
  - PDF生成
  - 数据库持久化（保持JSON文件缓存）
  - 邮件/Webhook通知
  - 用户目录API集成（手动维护user_mapping.json）
  - 超出PaySignPrinter现有错误处理级别的重试/退避逻辑
