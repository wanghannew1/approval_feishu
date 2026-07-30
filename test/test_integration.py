"""
End-to-end integration tests for the full Feishu approval workflow.

These tests exercise the integration between feishu_api, batch_processor,
and cache_manager modules using requests_mock to mock all HTTP requests.
"""

import json
from pathlib import Path
from unittest.mock import patch

import pytest
import requests_mock

from app.feishu_api import (
    get_tenant_token,
    get_auth_headers,
    query_instances,
    list_instances,
    get_instance_detail,
    parse_form,
    extract_attachments,
    download_file,
    BASE_URL,
    TOKEN_URL,
    QUERY_URL,
    INSTANCES_URL,
    INSTANCE_DETAIL_URL,
    DRIVE_DOWNLOAD_URL,
)
from app.batch_processor import (
    is_ready_for_print,
    is_approval_passed,
    get_approvers_with_roles,
)
from app.cache_manager import (
    InstanceDetailCache,
    DownloadURLCache,
)


@pytest.fixture
def temp_cache_file(monkeypatch, tmp_path):
    """Provide a temporary cache file path to avoid using real cache."""
    cache = tmp_path / ".token_cache.json"
    monkeypatch.setattr("app.feishu_api.CACHE_FILE", cache)
    return cache


class TestFullWorkflow:
    """End-to-end integration tests for the complete approval workflow."""

    def test_full_workflow_token_to_print_ready(
        self, requests_mock, mock_token, mock_auth_headers, tmp_path, temp_cache_file
    ):
        """
        Full workflow: token → list instances → detail → attachments → download → print ready check.
        """
        # 1. Mock token API
        requests_mock.post(
            TOKEN_URL,
            json={
                "code": 0,
                "msg": "success",
                "tenant_access_token": mock_token,
                "expire": 7200,
            },
        )

        # 2. Mock list instances
        requests_mock.get(
            INSTANCES_URL,
            json={
                "code": 0,
                "msg": "success",
                "data": {
                    "instance_code_list": ["INST-001"],
                    "has_more": False,
                    "page_token": "",
                },
            },
        )

        # 3. Mock instance detail with form and approver_list
        form_json = json.dumps([
            {
                "id": "w1",
                "type": "attachmentV2",
                "name": "工资表附件",
                "value": ["file_token_123"],
            }
        ])
        requests_mock.get(
            INSTANCE_DETAIL_URL.format(instance_code="INST-001"),
            json={
                "code": 0,
                "msg": "success",
                "data": {
                    "instance_code": "INST-001",
                    "status": "APPROVED",
                    "title": "6月工资发放审批",
                    "form": form_json,
                    "approver_list": [
                        {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                        {"approver_name": "部门负责人", "status": "APPROVED", "comment": "同意"},
                        {"approver_name": "财务", "status": "APPROVED", "comment": "同意"},
                    ],
                    "task_list": [
                        {"node_name": "出纳办理", "status": "PENDING"},
                    ],
                },
            },
        )

        # 4. Mock download
        requests_mock.get(
            DRIVE_DOWNLOAD_URL.format(file_token="file_token_123"),
            headers={"Content-Disposition": 'attachment; filename="工资表.xlsx"'},
            content=b"fake excel content",
            status_code=200,
        )

        # Step 1: Get token
        token = get_tenant_token("test_app_id", "test_app_secret")
        assert token == mock_token

        # Step 2: List instances
        instance_codes = list_instances(token, "APPROVAL-CODE", "1000", "2000")
        assert instance_codes == ["INST-001"]

        # Step 3: Get instance detail
        detail = get_instance_detail(token, "INST-001")
        assert detail["instance_code"] == "INST-001"
        assert detail["status"] == "APPROVED"

        # Step 4: Parse form and extract attachments
        form_widgets = parse_form(detail)
        attachments = extract_attachments(form_widgets)
        assert len(attachments) == 1
        assert attachments[0]["field_name"] == "工资表附件"
        assert attachments[0]["value"] == ["file_token_123"]

        # Step 5: Download file
        save_dir = str(tmp_path / "downloads")
        downloaded_path = download_file(token, "file_token_123", save_dir)
        assert Path(downloaded_path).exists()
        assert Path(downloaded_path).name == "工资表.xlsx"

        # Step 6: Verify is_ready_for_print returns True (all mandatory roles approved)
        assert is_ready_for_print(detail) is True

        # Step 7: Verify approver role mapping
        approvers = get_approvers_with_roles(detail)
        roles = {a["role"] for a in approvers if a["role"]}
        assert "总经理签字" in roles
        assert "部长签字" in roles
        assert "财务审核" in roles

    def test_full_workflow_approval_passed(
        self, requests_mock, mock_token, tmp_path, temp_cache_file
    ):
        """
        Full workflow with APPROVED status, verify is_approval_passed returns True.
        """
        # Mock token
        requests_mock.post(
            TOKEN_URL,
            json={
                "code": 0,
                "msg": "success",
                "tenant_access_token": mock_token,
                "expire": 7200,
            },
        )

        # Mock query instances (POST endpoint)
        requests_mock.post(
            QUERY_URL,
            json={
                "code": 0,
                "msg": "success",
                "data": {
                    "instance_list": [
                        {"instance_code": "INST-002", "status": "APPROVED"}
                    ],
                    "has_more": False,
                    "page_token": "",
                },
            },
        )

        # Mock instance detail
        requests_mock.get(
            INSTANCE_DETAIL_URL.format(instance_code="INST-002"),
            json={
                "code": 0,
                "msg": "success",
                "data": {
                    "instance_code": "INST-002",
                    "status": "APPROVED",
                    "title": "报销审批",
                    "form": "[]",
                    "approver_list": [
                        {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                    ],
                },
            },
        )

        # Step 1: Get token
        token = get_tenant_token("app_id", "app_secret")

        # Step 2: Query instances by status
        result = query_instances(token, "APPROVAL-CODE", instance_status="APPROVED")
        assert result["instance_list"][0]["instance_code"] == "INST-002"

        # Step 3: Get detail
        detail = get_instance_detail(token, "INST-002")

        # Step 4: Verify is_approval_passed
        assert is_approval_passed(detail) is True
        assert detail["status"] == "APPROVED"

    def test_full_workflow_with_cache(
        self, requests_mock, mock_token, tmp_path, temp_cache_file
    ):
        """
        First call hits API, second call uses cache. Verify cache stats.
        """
        cache_dir = str(tmp_path / "cache")
        detail_cache = InstanceDetailCache(cache_dir=cache_dir)
        instance_code = "INST-003"

        detail_data = {
            "instance_code": instance_code,
            "status": "APPROVED",
            "title": "测试审批",
            "form": "[]",
            "approver_list": [
                {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                {"approver_name": "部门负责人", "status": "APPROVED", "comment": "同意"},
                {"approver_name": "财务", "status": "APPROVED", "comment": "同意"},
            ],
        }

        # Mock instance detail endpoint
        requests_mock.get(
            INSTANCE_DETAIL_URL.format(instance_code=instance_code),
            json={
                "code": 0,
                "msg": "success",
                "data": detail_data,
            },
        )

        # First call: cache miss → hits API
        cached = detail_cache.get(instance_code)
        assert cached is None
        assert detail_cache.misses == 1

        # Fetch from API and store in cache
        detail = get_instance_detail(mock_token, instance_code)
        detail_cache.set(instance_code, detail)
        assert detail_cache.hits == 0  # No hit yet

        # Second call: cache hit
        cached_detail = detail_cache.get(instance_code)
        assert cached_detail is not None
        assert detail_cache.hits == 1
        assert detail_cache.misses == 1
        assert cached_detail["instance_code"] == instance_code
        assert cached_detail["status"] == "APPROVED"

    def test_edge_case_no_attachments(
        self, requests_mock, mock_token
    ):
        """
        Instance with no attachments — verify extract_attachments returns [].
        """
        # Mock instance detail with form containing only text widgets
        requests_mock.get(
            INSTANCE_DETAIL_URL.format(instance_code="INST-004"),
            json={
                "code": 0,
                "msg": "success",
                "data": {
                    "instance_code": "INST-004",
                    "status": "APPROVED",
                    "title": "无附件审批",
                    "form": json.dumps([
                        {"id": "w1", "type": "text", "name": "申请人", "value": "张三"},
                        {"id": "w2", "type": "number", "name": "金额", "value": "1000"},
                    ]),
                    "approver_list": [
                        {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                    ],
                },
            },
        )

        detail = get_instance_detail(mock_token, "INST-004")
        form_widgets = parse_form(detail)
        attachments = extract_attachments(form_widgets)

        assert attachments == []
        assert is_approval_passed(detail) is True

    def test_edge_case_empty_approver_list(
        self, requests_mock, mock_token
    ):
        """
        Instance with empty approver_list — verify is_ready_for_print returns False.
        """
        requests_mock.get(
            INSTANCE_DETAIL_URL.format(instance_code="INST-005"),
            json={
                "code": 0,
                "msg": "success",
                "data": {
                    "instance_code": "INST-005",
                    "status": "APPROVED",
                    "title": "空审批人列表",
                    "form": json.dumps([
                        {"id": "w1", "type": "attachmentV2", "name": "附件", "value": ["tok123"]},
                    ]),
                    "approver_list": [],
                },
            },
        )

        detail = get_instance_detail(mock_token, "INST-005")

        # Empty approver_list means no approved roles
        assert is_ready_for_print(detail) is False

        # But overall approval can still be passed
        assert is_approval_passed(detail) is True

        # Extract attachments should still work
        form_widgets = parse_form(detail)
        attachments = extract_attachments(form_widgets)
        assert len(attachments) == 1
