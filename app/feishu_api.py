"""
Feishu API client for approval operations.

Provides functions to authenticate with Feishu Open Platform,
query approval instances, and download attachments.
"""

import json
import os
import threading
import time
from pathlib import Path
from typing import Optional

import requests


BASE_URL = "https://open.feishu.cn/open-apis"
TOKEN_URL = f"{BASE_URL}/auth/v3/tenant_access_token/internal/"
INSTANCES_URL = f"{BASE_URL}/approval/v4/instances"
INSTANCE_DETAIL_URL = f"{BASE_URL}/approval/v4/instances/{{instance_code}}"
QUERY_URL = f"{BASE_URL}/approval/v4/instances/query"
DRIVE_DOWNLOAD_URL = f"{BASE_URL}/drive/v1/files/{{file_token}}/download"

CACHE_FILE = Path(".token_cache.json")
_CACHE_LOCK = threading.Lock()


def _load_cached_token(cache_path: Path | None = None) -> dict | None:
    """Load cached token data if it exists and is not expired."""
    if cache_path is None:
        cache_path = CACHE_FILE
    if not cache_path.exists():
        return None
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return None
    if data.get("expire_at", 0) > time.time():
        return data
    return None


def _save_token(token: str, expire_in: int, cache_path: Path | None = None) -> None:
    """Save token to cache with 5-minute early expiry."""
    if cache_path is None:
        cache_path = CACHE_FILE
    data = {
        "tenant_access_token": token,
        "expire_at": time.time() + expire_in - 300,
    }
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    # Atomic write to avoid corrupted cache files
    temp_path = cache_path.with_suffix(".tmp")
    with open(temp_path, "w", encoding="utf-8") as f:
        json.dump(data, f)
        f.flush()
        os.fsync(f.fileno())
    os.replace(temp_path, cache_path)


def get_tenant_token(app_id: str, app_secret: str) -> str:
    """
    Get tenant access token from Feishu API.

    Fetches a new token if none is cached or if the cached token has
    expired (with a 5-minute early expiry margin).  The cache is
    written atomically and protected by a thread lock.

    Args:
        app_id: Feishu application ID.
        app_secret: Feishu application secret.

    Returns:
        Tenant access token string.

    Raises:
        RuntimeError: If token acquisition fails.
    """
    with _CACHE_LOCK:
        cached = _load_cached_token()
        if cached:
            return cached["tenant_access_token"]

    payload = {"app_id": app_id, "app_secret": app_secret}
    resp = requests.post(TOKEN_URL, json=payload)
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(
            f"获取 token 失败: code={data.get('code')} msg={data.get('msg')}"
        )

    token = data["tenant_access_token"]
    with _CACHE_LOCK:
        # Re-check in case another thread already saved a token
        cached = _load_cached_token()
        if cached:
            return cached["tenant_access_token"]
        _save_token(token, data.get("expire", 7200))
    return token


def get_auth_headers(token: str) -> dict:
    """
    Build authorization headers for API requests.

    Args:
        token: Valid tenant access token.

    Returns:
        Dictionary with Authorization and Content-Type headers.
    """
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }


def query_approval_instances(
    headers: dict,
    approval_code: str,
    start_time: int,
    end_time: int,
    page_size: int = 50,
    page_token: Optional[str] = None,
) -> dict:
    """
    Query approval instances with pagination.

    Args:
        headers: Authorization headers.
        approval_code: Approval definition code.
        start_time: Start timestamp in milliseconds.
        end_time: End timestamp in milliseconds.
        page_size: Number of results per page.
        page_token: Pagination token from previous response.

    Returns:
        API response data dictionary.
    """
    # TODO: Implement
    pass


def get_instance_detail(headers: dict, instance_code: str) -> dict:
    """
    Get detailed information about a specific approval instance.

    Args:
        headers: Authorization headers.
        instance_code: Unique identifier for the instance.

    Returns:
        Instance detail data including form and attachments.
    """
    # TODO: Implement
    pass


def download_file(headers: dict, file_token_or_url: str, save_dir: str) -> str:
    """
    Download a file from Feishu Drive.

    Args:
        headers: Authorization headers.
        file_token_or_url: File token or direct URL.
        save_dir: Directory to save the downloaded file.

    Returns:
        Path to the saved file.
    """
    # TODO: Implement
    pass
