# 飞书审批打印工具

基于飞书开放平台 API 的审批附件下载与打印工具。从钉钉迁移至飞书版本。

## 项目结构

```
approval_feishu/
├── API.md              # 飞书审批 API 接口文档（5 个端点）
├── test_api.py         # API 测试脚本（逐个测试全部接口）
├── requirements.txt    # Python 依赖
├── .env.example        # 环境变量模板
└── .venv/              # uv 虚拟环境
```

## 快速开始

### 1. 环境准备

```bash
# 创建虚拟环境
uv venv
source .venv/bin/activate

# 安装依赖
uv pip install -r requirements.txt
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
python test_api.py
```

## 涉及的飞书 API

| 接口 | 用途 |
|------|------|
| `POST /auth/v3/tenant_access_token/internal/` | 获取访问令牌 |
| `GET /approval/v4/instances` | 查询审批实例列表 |
| `GET /approval/v4/instances/{code}` | 获取实例详情 |
| `GET /drive/v1/files/{token}/download` | 下载附件文件 |
| `POST /approval/v4/instances/search` | 按状态搜索 |

详见 [`API.md`](./API.md)。
