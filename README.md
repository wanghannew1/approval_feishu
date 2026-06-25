# 飞书审批打印工具

基于飞书开放平台 API 的审批附件下载与打印工具。从钉钉 PaySignPrinter 完整迁移至飞书版本。

## 项目结构

```
approval_feishu/
├── app/                        # 正式代码
│   ├── __init__.py
│   ├── feishu_api.py           # 飞书 API 封装（认证、列表、详情、附件下载）
│   ├── batch_processor.py      # 审批处理、签名插入、打印
│   ├── cache_manager.py        # 缓存管理（Token、下载 URL、实例详情）
│   ├── app.py                  # Streamlit UI 入口
│   ├── logger_config.py        # 日志配置
│   ├── role_mapping.json       # 审批节点 → 签名角色映射
│   ├── user_mapping.json       # 用户映射（空，待填充）
│   ├── payroll_sheet_config.json  # 工资表检测配置
│   └── settings.json           # 应用设置
├── test/                       # 测试代码
│   ├── __init__.py
│   ├── conftest.py             # pytest 共享 fixtures
│   ├── test_feishu_api.py      # API 单元测试（34 个）
│   ├── test_cache_manager.py   # 缓存管理测试（17 个）
│   ├── test_approval_status.py # 审批状态测试（9 个）
│   ├── test_batch_processor.py # 批处理测试（16 个）
│   ├── test_integration.py     # 端到端集成测试（5 个）
│   └── test_api.py             # 原始 API 验证脚本（非 pytest）
├── API.md                      # 飞书审批 API 接口文档
├── feishu_api_verification.md  # API 实际行为验证报告
├── requirements.txt            # Python 依赖
├── pyproject.toml              # 项目配置（uv + 清华镜像）
├── pytest.ini                  # pytest 配置
├── .env.example                # 环境变量模板
├── downloads/                  # 下载的附件
└── signatures/                 # 签名图片（PNG 格式，按姓名命名）
```

## 快速开始

### 1. 环境准备

```bash
# 创建虚拟环境
uv venv
source .venv/bin/activate   # Linux/macOS
.venv\Scripts\activate      # Windows

# 安装依赖（使用清华镜像加速）
uv pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

### 2. 配置凭证

```bash
cp .env.example .env
# 编辑 .env 填入飞书应用凭证
```

`.env` 内容：

```env
FEISHU_APP_ID=your_app_id
FEISHU_APP_SECRET=your_app_secret
FEISHU_APPROVAL_CODE=your_approval_code
```

### 3. 运行测试

```bash
# 运行全部 TDD 测试（85 个）
python -m pytest test/ -v

# 运行原始 API 验证脚本（需有效凭证）
python test/test_api.py
```

### 4. 启动 Streamlit UI

```bash
streamlit run app/app.py
```

## 主要功能

- **审批查询**：支持按状态筛选（全部 / 审批中 / 审批完成待出纳办理 / 已通过 / 已拒绝），默认选中"审批中"
- **附件下载**：自动检测 URL 或 file_token 格式，支持中文文件名解码
- **签名插入**：按审批人角色自动匹配签名图片插入 Excel
- **批量打印**：支持 Windows COM + LibreOffice 双平台打印
- **缓存管理**：Token 自动刷新，下载 URL 12 小时缓存

## 涉及的飞书 API

| 接口 | 用途 |
|------|------|
| `POST /auth/v3/tenant_access_token/internal/` | 获取访问令牌 |
| `GET /approval/v4/instances` | 查询审批实例列表 |
| `POST /approval/v4/instances/query` | 按状态查询实例 |
| `GET /approval/v4/instances/{code}` | 获取实例详情 |
| `GET /drive/v1/files/{token}/download` | 下载附件文件（备用）|

详见 [`API.md`](./API.md)。
