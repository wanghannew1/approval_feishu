"""
Tests for approval status logic in batch_processor.

TDD tests for Feishu approver_list based approval status functions.
"""

import json
from pathlib import Path

import pytest

from app.batch_processor import (
    get_approvers_with_roles,
    is_approval_passed,
    is_ready_for_print,
)


# Paths to config files used by the functions under test
ROLE_MAPPING_PATH = Path(__file__).parent.parent / "app" / "role_mapping.json"
PAYROLL_CONFIG_PATH = Path(__file__).parent.parent / "app" / "payroll_sheet_config.json"


class TestIsReadyForPrint:
    """Test suite for is_ready_for_print function."""

    def test_is_ready_for_print_all_approved(self):
        """All mandatory roles APPROVED -> True."""
        details = {
            "approver_list": [
                {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                {"approver_name": "部门负责人", "status": "APPROVED", "comment": "同意"},
                {"approver_name": "财务", "status": "APPROVED", "comment": "同意"},
            ]
        }
        assert is_ready_for_print(details) is True

    def test_is_ready_for_print_partial(self):
        """Some roles pending -> False."""
        details = {
            "approver_list": [
                {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                {"approver_name": "部门负责人", "status": "PENDING", "comment": ""},
                {"approver_name": "财务", "status": "APPROVED", "comment": "同意"},
            ]
        }
        assert is_ready_for_print(details) is False

    def test_is_ready_for_print_empty(self):
        """Empty approver_list -> False."""
        details = {"approver_list": []}
        assert is_ready_for_print(details) is False

    def test_is_ready_for_print_missing_role(self):
        """Missing role mapping -> False."""
        # Only 2 of 3 mandatory roles are present and approved
        details = {
            "approver_list": [
                {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                {"approver_name": "部门负责人", "status": "APPROVED", "comment": "同意"},
                # "财务" role is missing entirely from approver_list
            ]
        }
        assert is_ready_for_print(details) is False


class TestIsApprovalPassed:
    """Test suite for is_approval_passed function."""

    def test_is_approval_passed_approved(self):
        """status APPROVED -> True."""
        details = {"status": "APPROVED"}
        assert is_approval_passed(details) is True

    def test_is_approval_passed_pending(self):
        """status PENDING -> False."""
        details = {"status": "PENDING"}
        assert is_approval_passed(details) is False

    def test_is_approval_passed_rejected(self):
        """status REJECTED -> False."""
        details = {"status": "REJECTED"}
        assert is_approval_passed(details) is False


class TestGetApproversWithRoles:
    """Test suite for get_approvers_with_roles function."""

    def test_get_approvers_with_roles(self):
        """Verify role mapping from approver_name."""
        details = {
            "approver_list": [
                {"approver_name": "总经理", "status": "APPROVED", "comment": "同意"},
                {"approver_name": "部门负责人", "status": "PENDING", "comment": ""},
                {"approver_name": "未知人员", "status": "APPROVED", "comment": "同意"},
            ]
        }
        result = get_approvers_with_roles(details, ROLE_MAPPING_PATH)

        assert len(result) == 3
        assert result[0] == {
            "approver_name": "总经理",
            "role": "总经理签字",
            "status": "APPROVED",
        }
        assert result[1] == {
            "approver_name": "部门负责人",
            "role": "部长签字",
            "status": "PENDING",
        }
        assert result[2] == {
            "approver_name": "未知人员",
            "role": None,
            "status": "APPROVED",
        }

    def test_get_approvers_no_id(self):
        """Verify works with only approver_name (no user_id)."""
        details = {
            "approver_list": [
                {"approver_name": "财务", "status": "APPROVED"},
            ]
        }
        result = get_approvers_with_roles(details, ROLE_MAPPING_PATH)

        assert len(result) == 1
        assert result[0] == {
            "approver_name": "财务",
            "role": "财务审核",
            "status": "APPROVED",
        }
