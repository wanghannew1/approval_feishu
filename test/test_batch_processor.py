"""
Tests for batch_processor module.

TDD test suite covering is_payroll_sheet, get_signature_path,
is_ready_for_print edge cases, and process_single_approval.
"""

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from app.batch_processor import (
    get_signature_path,
    is_payroll_sheet,
    is_ready_for_print,
    process_single_approval,
)


class TestIsPayrollSheet:
    """Test suite for is_payroll_sheet function."""

    def _make_mock_worksheet(self, row_data):
        """Create a mock worksheet with given row data.

        Args:
            row_data: Dict mapping (row, col) to cell value.
        """
        ws = MagicMock()
        ws.max_row = max(r for r, c in row_data.keys()) if row_data else 0
        ws.max_column = max(c for r, c in row_data.keys()) if row_data else 0

        def get_cell(row, column=None):
            cell = MagicMock()
            cell.value = row_data.get((row, column))
            return cell

        ws.cell = get_cell
        return ws

    def test_payroll_sheet_with_all_required_elements(self):
        """Test detection of valid payroll sheet with all required elements."""
        row_data = {
            (1, 1): "2026年派遣员工5月工资明细表（林下参）",
            (2, 1): "序号",
            (2, 2): "姓名",
            (2, 3): "基本工资",
            (2, 4): "应发工资",
            (2, 5): "转款合计",
            (2, 6): "实发工资",
            (3, 1): "养老",
            (10, 1): "   总经理签字：",
            (10, 11): "财务审核：",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is True

    def test_payroll_sheet_missing_title_keyword(self):
        """Test rejection when row 1 lacks required keyword."""
        row_data = {
            (1, 1): "普通表格",
            (2, 1): "序号",
            (2, 2): "应发工资",
            (3, 1): "养老",
            (10, 1): "   总经理签字：",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_missing_org_keyword(self):
        """Test rejection when row 2 lacks required keyword."""
        row_data = {
            (1, 1): "2026年派遣员工5月工资明细表（林下参）",
            (2, 1): "部门信息",
            (3, 1): "养老",
            (10, 1): "   总经理签字：",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_missing_headers(self):
        """Test rejection when row 3 lacks required headers."""
        row_data = {
            (1, 1): "2026年派遣员工5月工资明细表（林下参）",
            (2, 1): "序号",
            (2, 2): "应发工资",
            (3, 1): "转账合计",
            (10, 1): "   总经理签字：",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_missing_signature_keywords(self):
        """Test rejection when signature keywords not found."""
        row_data = {
            (1, 1): "2026年派遣员工5月工资明细表（林下参）",
            (2, 1): "序号",
            (2, 2): "应发工资",
            (3, 1): "养老",
            (10, 1): "其他内容",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_with_custom_config(self):
        """Test is_payroll_sheet with custom config."""
        row_data = {
            (1, 1): "自定义工资表",
            (2, 1): "公司名",
            (3, 1): "转账合计",
            (3, 2): "应发工资",
            (3, 3): "实发工资",
            (3, 4): "实发合计",
            (5, 1): "manager",
        }
        ws = self._make_mock_worksheet(row_data)
        custom_config = {
            "sheet_filter": {
                "row1_title": {"required_keyword": "自定义工资表"},
                "row2_org": {"required_keyword": "公司名"},
                "row3_headers": {"required": ["转账合计", "应发工资", "实发工资", "实发合计"]},
                "signatures": {
                    "mandatory": {
                        "manager": ["manager"],
                    },
                    "optional": {},
                },
            }
        }
        assert is_payroll_sheet(ws, custom_config) is True


class TestGetSignaturePath:
    """Test suite for get_signature_path function."""

    def test_signature_path_exact_match(self, tmp_path):
        """Test finding signature with exact name match."""
        sig_dir = tmp_path
        (sig_dir / "张三.png").touch()

        result = get_signature_path("张三", sig_dir)
        assert result == sig_dir / "张三.png"

    def test_signature_path_glob_match(self, tmp_path):
        """Test finding signature using glob when exact match fails."""
        sig_dir = tmp_path
        (sig_dir / "李四_签名.png").touch()

        result = get_signature_path("李四", sig_dir)
        assert result == sig_dir / "李四_签名.png"

    def test_signature_path_not_found(self, tmp_path):
        """Test None when signature not found."""
        sig_dir = tmp_path
        (sig_dir / "其他人.png").touch()

        result = get_signature_path("不存在", sig_dir)
        assert result is None

    def test_signature_path_empty_name(self, tmp_path):
        """Test None when approver name is empty."""
        sig_dir = tmp_path
        (sig_dir / "test.png").touch()

        result = get_signature_path("", sig_dir)
        assert result is None

    def test_signature_path_no_files_in_dir(self, tmp_path):
        """Test None when directory is empty."""
        result = get_signature_path("任何人", tmp_path)
        assert result is None


class TestIsReadyForPrint:
    """Test suite for is_ready_for_print edge cases."""

    @patch("app.batch_processor._get_role_mapping")
    @patch("app.batch_processor._get_mandatory_roles")
    def test_ready_when_all_mandatory_roles_approved(self, mock_mandatory, mock_mapping):
        """Test True when all mandatory roles have APPROVED status."""
        mock_mandatory.return_value = {"总经理签字", "部长签字"}
        mock_mapping.return_value = {"张三": "总经理签字", "李四": "部长签字"}

        details = {
            "approver_list": [
                {"approver_name": "张三", "status": "APPROVED"},
                {"approver_name": "李四", "status": "APPROVED"},
            ],
            "task_list": [
                {"node_name": "出纳办理", "status": "PENDING"},
            ]
        }
        assert is_ready_for_print(details) is True

    @patch("app.batch_processor._get_role_mapping")
    @patch("app.batch_processor._get_mandatory_roles")
    def test_not_ready_when_mandatory_role_missing(self, mock_mandatory, mock_mapping):
        """Test False when a mandatory role is missing or not approved."""
        mock_mandatory.return_value = {"总经理签字", "部长签字"}
        mock_mapping.return_value = {"张三": "总经理签字", "李四": "部长签字"}

        details = {
            "approver_list": [
                {"approver_name": "张三", "status": "APPROVED"},
                {"approver_name": "李四", "status": "PENDING"},
            ]
        }
        assert is_ready_for_print(details) is False

    @patch("app.batch_processor._get_role_mapping")
    @patch("app.batch_processor._get_mandatory_roles")
    def test_not_ready_when_no_approvers(self, mock_mandatory, mock_mapping):
        """Test False when approver_list is empty."""
        mock_mandatory.return_value = {"总经理签字"}
        mock_mapping.return_value = {}

        details = {"approver_list": []}
        assert is_ready_for_print(details) is False

    @patch("app.batch_processor._get_role_mapping")
    @patch("app.batch_processor._get_mandatory_roles")
    def test_false_when_mandatory_roles_empty(self, mock_mandatory, mock_mapping):
        """Test False when mandatory_roles is empty (no config)."""
        mock_mandatory.return_value = set()
        mock_mapping.return_value = {}

        details = {"approver_list": [{"approver_name": "张三", "status": "APPROVED"}]}
        assert is_ready_for_print(details) is False

    @patch("app.batch_processor._get_role_mapping")
    @patch("app.batch_processor._get_mandatory_roles")
    def test_ignores_non_approved_approvers(self, mock_mandatory, mock_mapping):
        """Test that only APPROVED status counts toward ready."""
        mock_mandatory.return_value = {"总经理签字"}
        mock_mapping.return_value = {"张三": "总经理签字"}

        details = {
            "approver_list": [
                {"approver_name": "张三", "status": "REJECTED"},
            ]
        }
        assert is_ready_for_print(details) is False


class TestProcessSingleApproval:
    """Test suite for process_single_approval function."""

    @patch("app.batch_processor.extract_attachments")
    @patch("app.batch_processor.parse_form")
    @patch("app.batch_processor.get_instance_detail")
    def test_process_downloads_non_approved_instance(
        self, mock_detail, mock_parse, mock_extract
    ):
        """Test that non-APPROVED instances still download but skip signing."""
        mock_detail.return_value = {
            "instance_code": "test123",
            "approval_name": "测试审批",
            "status": "PENDING",
        }
        mock_parse.return_value = []
        mock_extract.return_value = [{"field_name": "工资表", "value": ["url"]}]

        result = process_single_approval("test123", "fake_token", {})

        assert result["skipped"] is False
        assert "审批未通过" not in result.get("message", "")

    @patch("app.batch_processor.get_approvers_with_roles")
    @patch("app.batch_processor.extract_attachments")
    @patch("app.batch_processor.parse_form")
    @patch("app.batch_processor.get_instance_detail")
    def test_process_downloads_without_signers(
        self, mock_detail, mock_parse, mock_extract, mock_approvers
    ):
        """Test that instances without valid signers still download."""
        mock_detail.return_value = {
            "instance_code": "test123",
            "approval_name": "测试审批",
            "status": "APPROVED",
            "approver_list": [],
        }
        mock_approvers.return_value = []
        mock_parse.return_value = []
        mock_extract.return_value = [{"field_name": "工资表", "value": ["url"]}]

        result = process_single_approval("test123", "fake_token", {})

        assert result["skipped"] is False

    @patch("app.batch_processor.get_approvers_with_roles")
    @patch("app.batch_processor.extract_attachments")
    @patch("app.batch_processor.parse_form")
    @patch("app.batch_processor.get_instance_detail")
    def test_process_skips_when_no_attachments(
        self, mock_detail, mock_parse, mock_extract, mock_approvers
    ):
        """Test when instance has no attachments - returns failure."""
        mock_detail.return_value = {
            "instance_code": "test123",
            "title": "测试审批",
            "status": "APPROVED",
            "approver_list": [
                {"approver_name": "张三", "status": "APPROVED", "role": "总经理签字"}
            ],
        }
        mock_approvers.return_value = [
            {"approver_name": "张三", "status": "APPROVED", "role": "总经理签字"}
        ]
        mock_parse.return_value = []
        mock_extract.return_value = []

        result = process_single_approval("test123", "fake_token", {})

        assert result["success"] is False
        assert result["message"] == "无附件"

    @patch("app.batch_processor.get_approvers_with_roles")
    @patch("app.batch_processor.extract_attachments")
    @patch("app.batch_processor.parse_form")
    @patch("app.batch_processor.get_instance_detail")
    @patch("app.batch_processor.download_file")
    def test_process_success_with_download_only(
        self, mock_download, mock_detail, mock_parse, mock_extract, mock_approvers, tmp_path
    ):
        """Test successful processing with attachment download (no signing)."""
        mock_detail.return_value = {
            "instance_code": "test123",
            "title": "测试审批",
            "status": "APPROVED",
            "approver_list": [
                {"approver_name": "张三", "status": "APPROVED", "role": "总经理签字"}
            ],
        }
        mock_approvers.return_value = [
            {"approver_name": "张三", "status": "APPROVED", "role": "总经理签字"}
        ]
        mock_parse.return_value = []
        mock_extract.return_value = [
            {"field_name": "附件", "value": ["file_token_123"]}
        ]
        mock_download.return_value = str(tmp_path / "test.txt")

        config = {
            "save_dir": str(tmp_path),
            "signatures_dir": str(tmp_path / "sigs"),
        }

        result = process_single_approval("test123", "fake_token", config)

        assert result["success"] is True
        assert "test.txt" in result["downloaded"]

    @patch("app.batch_processor.get_approvers_with_roles")
    @patch("app.batch_processor.extract_attachments")
    @patch("app.batch_processor.parse_form")
    @patch("app.batch_processor.get_instance_detail")
    def test_process_returns_correct_instance_code(
        self, mock_detail, mock_parse, mock_extract, mock_approvers, tmp_path
    ):
        """Test that result contains the correct instance_code."""
        mock_detail.return_value = {
            "instance_code": "instance_abc",
            "approval_name": "标题",
            "status": "APPROVED",
        }
        mock_approvers.return_value = []
        mock_parse.return_value = []
        mock_extract.return_value = []

        result = process_single_approval("instance_abc", "fake_token", {})

        assert result["instance_code"] == "instance_abc"
        assert result["title"] == "标题"
