"""
Tests for Feishu API module.
"""

import pytest
from app.feishu_api import get_auth_headers, BASE_URL, TOKEN_URL


class TestFeishuApi:
    """Test suite for feishu_api module."""

    def test_get_auth_headers(self, mock_token: str):
        """Test authorization header generation."""
        headers = get_auth_headers(mock_token)
        assert "Authorization" in headers
        assert headers["Authorization"] == f"Bearer {mock_token}"
        assert "Content-Type" in headers

    def test_base_url_is_set(self):
        """Test that BASE_URL is properly configured."""
        assert BASE_URL.startswith("https://")

    def test_token_url_format(self):
        """Test that TOKEN_URL follows expected format."""
        assert "auth/v3/tenant_access_token" in TOKEN_URL
