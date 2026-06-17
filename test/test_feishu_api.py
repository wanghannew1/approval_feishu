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
    query_instances,
    list_instances,
    BASE_URL,
    TOKEN_URL,
    QUERY_URL,
    INSTANCES_URL,
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


class TestQueryInstances:
    """Test suite for query_instances POST endpoint."""

    def test_query_instances_success(self, mock_token: str):
        """Mock 200 and verify instance_list is returned."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "instance_list": [
                    {"instance_code": "ABC-123", "status": "APPROVED"}
                ],
                "has_more": False,
                "page_token": "",
            },
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = query_instances(mock_token, "approval_1")
            assert result == {
                "instance_list": [{"instance_code": "ABC-123", "status": "APPROVED"}],
                "has_more": False,
                "page_token": "",
            }
            mock_post.assert_called_once_with(
                QUERY_URL,
                headers=get_auth_headers(mock_token),
                json={"approval_code": "approval_1", "page_size": 50},
            )

    def test_query_instances_with_status(self, mock_token: str):
        """Mock with instance_status filter."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {"instance_list": [], "has_more": False, "page_token": ""},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = query_instances(
                mock_token,
                "approval_1",
                page_size=10,
                page_token="tok1",
                instance_status="PENDING",
            )
            assert result == {"instance_list": [], "has_more": False, "page_token": ""}
            mock_post.assert_called_once_with(
                QUERY_URL,
                headers=get_auth_headers(mock_token),
                json={
                    "approval_code": "approval_1",
                    "page_size": 10,
                    "page_token": "tok1",
                    "instance_status": "PENDING",
                },
            )

    def test_query_instances_empty(self, mock_token: str):
        """Mock empty result."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {"instance_list": [], "has_more": False, "page_token": ""},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            result = query_instances(mock_token, "approval_1")
            assert result["instance_list"] == []
            assert result["has_more"] is False

    def test_query_instances_api_error(self, mock_token: str):
        """Mock code != 0 and verify RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 99991672,
            "msg": "permission denied",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="查询实例失败"):
                query_instances(mock_token, "approval_1")


class TestListInstances:
    """Test suite for list_instances GET endpoint."""

    def test_list_instances_success(self, mock_token: str):
        """Mock GET /instances and verify list returned."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "instance_code_list": ["code1", "code2"],
                "has_more": False,
                "page_token": "",
            },
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response) as mock_get:
            result = list_instances(mock_token, "approval_1", "1000", "2000")
            assert result == ["code1", "code2"]
            mock_get.assert_called_once_with(
                INSTANCES_URL,
                headers=get_auth_headers(mock_token),
                params={
                    "approval_code": "approval_1",
                    "start_time": "1000",
                    "end_time": "2000",
                    "page_size": 50,
                },
            )

    def test_list_instances_pagination(self, mock_token: str):
        """Mock has_more=True and verify auto-pagination."""
        resp1 = MagicMock()
        resp1.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "instance_code_list": ["code1"],
                "has_more": True,
                "page_token": "tok1",
            },
        }
        resp2 = MagicMock()
        resp2.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "instance_code_list": ["code2"],
                "has_more": False,
                "page_token": "",
            },
        }
        with patch(
            "app.feishu_api.requests.get", side_effect=[resp1, resp2]
        ) as mock_get:
            result = list_instances(mock_token, "approval_1", "1000", "2000")
            assert result == ["code1", "code2"]
            assert mock_get.call_count == 2
            second_call = mock_get.call_args_list[1]
            assert second_call.kwargs["params"]["page_token"] == "tok1"

    def test_list_instances_empty(self, mock_token: str):
        """Mock empty result."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "instance_code_list": [],
                "has_more": False,
                "page_token": "",
            },
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response) as mock_get:
            result = list_instances(mock_token, "approval_1", "1000", "2000")
            assert result == []
            mock_get.assert_called_once()
