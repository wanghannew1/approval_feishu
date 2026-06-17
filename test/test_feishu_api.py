"""
Tests for Feishu API module.
"""

import json
import time
from unittest.mock import patch, MagicMock

import pytest
from app.feishu_api import (
    get_auth_headers,
    get_tenant_token,
    BASE_URL,
    TOKEN_URL,
    CACHE_FILE,
)


@pytest.fixture
def temp_cache_file(monkeypatch, tmp_path):
    """Provide a temporary cache file path."""
    cache = tmp_path / ".token_cache.json"
    monkeypatch.setattr("app.feishu_api.CACHE_FILE", cache)
    return cache


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

    def test_get_tenant_token_success(self, temp_cache_file):
        """Test fetching tenant token from API."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "tenant_access_token": "new_token_123",
            "expire": 7200,
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            token = get_tenant_token("app_id", "app_secret")
            assert token == "new_token_123"
            mock_post.assert_called_once_with(
                TOKEN_URL, json={"app_id": "app_id", "app_secret": "app_secret"}
            )
            assert temp_cache_file.exists()

    def test_get_tenant_token_uses_cache(self, temp_cache_file):
        """Test that cached token is returned without HTTP request."""
        cache_data = {
            "tenant_access_token": "cached_token_456",
            "expire_at": time.time() + 3600,
        }
        temp_cache_file.write_text(json.dumps(cache_data))

        with patch("app.feishu_api.requests.post") as mock_post:
            token = get_tenant_token("app_id", "app_secret")
            assert token == "cached_token_456"
            mock_post.assert_not_called()

    def test_get_tenant_token_expired(self, temp_cache_file):
        """Test that expired cache triggers a new HTTP request."""
        cache_data = {
            "tenant_access_token": "old_token",
            "expire_at": time.time() - 100,
        }
        temp_cache_file.write_text(json.dumps(cache_data))

        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "tenant_access_token": "new_token_789",
            "expire": 7200,
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            token = get_tenant_token("app_id", "app_secret")
            assert token == "new_token_789"
            mock_post.assert_called_once()

    def test_get_tenant_token_auth_failure(self, temp_cache_file):
        """Test that auth failure raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 10003,
            "msg": "Invalid app_id or app_secret",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="获取 token 失败"):
                get_tenant_token("bad_id", "bad_secret")
