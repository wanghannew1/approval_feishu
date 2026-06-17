"""
Pytest configuration and shared fixtures for tests.
"""

import pytest
from typing import Any


@pytest.fixture
def mock_token() -> str:
    """Provide a mock tenant access token."""
    return "test_tenant_access_token_12345"


@pytest.fixture
def mock_auth_headers(mock_token: str) -> dict:
    """Provide mock authorization headers."""
    return {
        "Authorization": f"Bearer {mock_token}",
        "Content-Type": "application/json; charset=utf-8",
    }


@pytest.fixture
def mock_api_response() -> dict:
    """Provide a mock Feishu API response structure."""
    return {
        "code": 0,
        "msg": "success",
        "data": {
            "instance_code_list": [],
            "has_more": False,
            "page_token": "",
        },
    }
