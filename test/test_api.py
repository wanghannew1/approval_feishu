#!/usr/bin/env python3
"""飞书审批 API 测试脚本 —— 逐个测试每个接口并输出结果。"""

import json
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import unquote

import requests
from dotenv import load_dotenv

# ── 配置 ──────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

APP_ID = os.getenv("FEISHU_APP_ID", "")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
APPROVAL_CODE = os.getenv("FEISHU_APPROVAL_CODE", "")

BASE_URL = "https://open.feishu.cn/open-apis"
TOKEN_URL = f"{BASE_URL}/auth/v3/tenant_access_token/internal/"
INSTANCES_URL = f"{BASE_URL}/approval/v4/instances"
INSTANCE_DETAIL_URL = f"{BASE_URL}/approval/v4/instances/{{instance_code}}"
QUERY_URL = f"{BASE_URL}/approval/v4/instances/query"
DRIVE_DOWNLOAD_URL = f"{BASE_URL}/drive/v1/files/{{file_token}}/download"

DOWNLOAD_DIR = PROJECT_ROOT / "downloads"
TOKEN_CACHE = PROJECT_ROOT / ".token_cache.json"

# ── Token 管理 ────────────────────────────────────────────


def _load_cached_token() -> dict | None:
    if not TOKEN_CACHE.exists():
        return None
    with open(TOKEN_CACHE) as f:
        data = json.load(f)
    if data.get("expire_at", 0) > datetime.now().timestamp():
        return data
    return None


def _save_token(token: str, expire: int) -> None:
    data = {
        "tenant_access_token": token,
        "expire_at": datetime.now().timestamp() + expire - 300,  # 提前 5 分钟过期
    }
    TOKEN_CACHE.parent.mkdir(parents=True, exist_ok=True)
    with open(TOKEN_CACHE, "w") as f:
        json.dump(data, f)


def get_tenant_token() -> str:
    """API 1: 获取 Tenant Access Token（带缓存）。"""
    cached = _load_cached_token()
    if cached:
        return cached["tenant_access_token"]

    payload = {"app_id": APP_ID, "app_secret": APP_SECRET}
    resp = requests.post(TOKEN_URL, json=payload)
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(f"获取 token 失败: code={data.get('code')} msg={data.get('msg')}")

    token = data["tenant_access_token"]
    _save_token(token, data.get("expire", 7200))
    return token


def auth_headers() -> dict:
    return {
        "Authorization": f"Bearer {get_tenant_token()}",
        "Content-Type": "application/json; charset=utf-8",
    }


# ── 测试函数 ──────────────────────────────────────────────


def test_token(token: str) -> bool:
    """测试 1: 验证 Token 是否有效。"""
    assert token, "token 为空"
    assert len(token) > 20, f"token 长度异常: {len(token)}"
    return True


def test_query_instances(token: str) -> list[str]:
    """测试 2: 查询审批实例列表（处理分页）。"""
    end_ms = str(int(datetime.now().timestamp() * 1000))
    start_ms = str(int((datetime.now() - timedelta(days=90)).timestamp() * 1000))

    all_codes = []
    page_token = None
    headers = auth_headers()

    while True:
        params = {
            "approval_code": APPROVAL_CODE,
            "start_time": start_ms,
            "end_time": end_ms,
            "page_size": 50,
        }
        if page_token:
            params["page_token"] = page_token

        resp = requests.get(INSTANCES_URL, headers=headers, params=params)
        data = resp.json()

        if data.get("code") != 0:
            raise RuntimeError(f"查询实例列表失败: {data}")

        batch = data["data"]["instance_code_list"]
        all_codes.extend(batch)

        if not data["data"]["has_more"]:
            break
        page_token = data["data"]["page_token"]

    return all_codes


def test_get_instance_detail(instance_code: str) -> dict:
    """测试 3: 获取单个实例详情，解析 form 中的附件。"""
    headers = auth_headers()
    resp = requests.get(
        INSTANCE_DETAIL_URL.format(instance_code=instance_code), headers=headers
    )
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(f"获取实例详情失败: {data}")

    detail = data["data"]

    # 解析 form JSON
    form_str = detail.get("form", "[]")
    try:
        form_widgets = json.loads(form_str)
    except json.JSONDecodeError:
        form_widgets = []

    # 提取附件
    attachments = []
    for widget in form_widgets:
        if widget.get("type") == "attachmentV2":
            for ft in widget.get("value", []):
                attachments.append({
                    "field_name": widget.get("name", "附件"),
                    "file_token_or_url": ft,
                })

    return {
        "instance_code": detail.get("instance_code"),
        "approval_name": detail.get("approval_name"),
        "status": detail.get("status"),
        "start_time": detail.get("start_time"),
        "attachments": attachments,
        "approver_names": [
            a.get("approver_name", "") for a in detail.get("approver_list", [])
        ],
    }


def test_download_file(file_token_or_url: str, save_dir: Path) -> Path:
    """测试 4: 下载文件到本地（自动检测 URL 或 file_token 格式）。"""
    headers = auth_headers()

    if file_token_or_url.startswith("http"):
        url = file_token_or_url
    else:
        url = DRIVE_DOWNLOAD_URL.format(file_token=file_token_or_url)

    resp = requests.get(url, headers=headers, stream=True)
    resp.raise_for_status()

    # 从 Content-Disposition 提取文件名
    filename = None
    cd = resp.headers.get("Content-Disposition", "")
    if cd:
        match = re.search(r'filename[*]?\s*=\s*(?:UTF-8\'\')?"?([^";\s]+)', cd)
        if match:
            filename = unquote(match.group(1))

    if not filename:
        filename = f"{file_token_or_url}.temp"

    save_dir.mkdir(parents=True, exist_ok=True)
    filepath = save_dir / filename

    with open(filepath, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

    return filepath


def test_search_instances(status: str | None = None) -> dict:
    """测试 5: 按状态搜索审批实例。"""
    headers = auth_headers()
    end_ms = str(int(datetime.now().timestamp() * 1000))
    start_ms = str(int((datetime.now() - timedelta(days=30)).timestamp() * 1000))

    body = {
        "approval_code": APPROVAL_CODE,
        "instance_start_time_from": start_ms,
        "instance_start_time_to": end_ms,
        "page_size": 50,
    }
    if status:
        body["instance_status"] = status

    resp = requests.post(QUERY_URL, headers=headers, json=body)
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(f"搜索实例失败: {data}")

    return data["data"]


# ── 主流程 ────────────────────────────────────────────────


def main():
    print("=" * 60)
    print("  飞书审批 API 测试")
    print("=" * 60)

    # ── 前置检查 ──
    if not all([APP_ID, APP_SECRET, APPROVAL_CODE]):
        print("\n❌ 请先配置 .env 文件：")
        print("  FEISHU_APP_ID=your_app_id")
        print("  FEISHU_APP_SECRET=your_app_secret")
        print("  FEISHU_APPROVAL_CODE=your_approval_code")
        sys.exit(1)

    print(f"\n📋 配置:")
    print(f"  App ID:     {APP_ID[:8]}...")
    print(f"  Approval:   {APPROVAL_CODE}")

    # ── 测试 1 ──
    print(f"\n{'─' * 60}")
    print("测试 1: 获取 Tenant Access Token")
    print(f"{'─' * 60}")
    try:
        token = get_tenant_token()
        test_token(token)
        print(f"  ✅ 成功: token = {token[:30]}...")
    except Exception as e:
        print(f"  ❌ 失败: {e}")
        sys.exit(1)

    # ── 测试 2 ──
    print(f"\n{'─' * 60}")
    print("测试 2: 查询审批实例列表")
    print(f"{'─' * 60}")
    try:
        codes = test_query_instances(token)
        print(f"  ✅ 成功: 共 {len(codes)} 条实例")
        if codes:
            print(f"  示例: {codes[0]}")
        else:
            print("  ⚠️  列表为空，跳过后续详情和下载测试")
    except Exception as e:
        print(f"  ❌ 失败: {e}")
        codes = []

    # ── 测试 3 ──
    if codes:
        print(f"\n{'─' * 60}")
        print("测试 3: 获取实例详情 + 解析附件")
        print(f"{'─' * 60}")
        first_attachments = []
        for i, code in enumerate(codes[:3]):  # 只测前 3 条
            try:
                detail = test_get_instance_detail(code)
                status_icon = {"APPROVED": "✅", "REJECTED": "❌", "PENDING": "⏳", "CANCELED": "🚫"}.get(
                    detail["status"], "❓"
                )
                print(f"  [{i+1}] {status_icon} {detail['approval_name']}")
                print(f"      状态: {detail['status']}  审批人: {', '.join(detail['approver_names'][:3])}")
                if detail["attachments"]:
                    print(f"      附件: {len(detail['attachments'])} 个")
                    for att in detail["attachments"]:
                        print(f"        📎 {att['field_name']}: {att['file_token_or_url'][:40]}...")
                    if i == 0:
                        first_attachments = detail["attachments"]
                else:
                    print(f"      附件: 无")
            except Exception as e:
                print(f"  [{i+1}] ❌ 失败: {e}")

        # ── 测试 4 ──
        if first_attachments:
            print(f"\n{'─' * 60}")
            print("测试 4: 下载附件")
            print(f"{'─' * 60}")
            for att in first_attachments[:2]:  # 最多下载 2 个
                try:
                    saved = test_download_file(att["file_token_or_url"], DOWNLOAD_DIR)
                    size_kb = saved.stat().st_size / 1024
                    print(f"  ✅ 下载成功: {saved.name} ({size_kb:.1f} KB)")
                    print(f"     路径: {saved}")
                except Exception as e:
                    print(f"  ❌ 下载失败 [{att['field_name']}]: {e}")

    # ── 测试 5 ──
    print(f"\n{'─' * 60}")
    print("测试 5: 按状态搜索审批实例")
    print(f"{'─' * 60}")
    try:
        result = test_search_instances(status="APPROVED")
        count = len(result.get("instance_list", []))
        print(f"  ✅ 成功: 已通过 {count} 条")
        if result.get("has_more"):
            print(f"  ⚠️  有更多数据，page_token={result.get('page_token', '')[:20]}...")
    except Exception as e:
        err_str = str(e)
        if "99991672" in err_str:
            print("  ⚠️  应用缺少权限: approval:approval.list:readonly")
            print("     请在飞书开放平台为该应用添加此权限后重试")
        else:
            print(f"  ❌ 失败: {e}")

    # ── 总结 ──
    print(f"\n{'=' * 60}")
    print("  全部测试完成")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
