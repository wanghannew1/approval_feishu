"""
Feishu API client for approval operations.

Provides functions to authenticate with Feishu Open Platform,
query approval instances, and download attachments.
"""

import requests
from typing import Optional


BASE_URL = "https://open.feishu.cn/open-apis"
TOKEN_URL = f"{BASE_URL}/auth/v3/tenant_access_token/internal/"
INSTANCES_URL = f"{BASE_URL}/approval/v4/instances"
INSTANCE_DETAIL_URL = f"{BASE_URL}/approval/v4/instances/{{instance_code}}"
QUERY_URL = f"{BASE_URL}/approval/v4/instances/query"
DRIVE_DOWNLOAD_URL = f"{BASE_URL}/drive/v1/files/{{file_token}}/download"


def get_tenant_token(app_id: str, app_secret: str) -> str:
    """
    Get tenant access token from Feishu API.

    Args:
        app_id: Feishu application ID.
        app_secret: Feishu application secret.

    Returns:
        Tenant access token string.

    Raises:
        RuntimeError: If token acquisition fails.
    """
    # TODO: Implement with caching
    pass


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
