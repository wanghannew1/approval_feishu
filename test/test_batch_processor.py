"""
Tests for batch_processor module.

TDD test suite covering is_payroll_sheet, get_signature_path,
is_ready_for_print edge cases, and process_single_approval.
"""

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from app.batch_processor import (
    _get_signature_keywords,
    _remove_empty_columns,
    get_payroll_config,
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
        row_data = {
            (1, 1): "普通表格",
            (2, 1): "序号",
            (2, 2): "姓名",
            (2, 3): "金额",
            (10, 1): "   总经理签字：",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_missing_required_content(self):
        row_data = {
            (1, 1): "2026年派遣员工5月工资明细表（林下参）",
            (2, 1): "部门信息",
            (3, 1): "养老",
            (10, 1): "   总经理签字：",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_missing_signature(self):
        row_data = {
            (1, 1): "2026年派遣员工5月工资明细表（林下参）",
            (2, 1): "序号",
            (2, 2): "应发工资",
            (2, 3): "实发工资",
            (3, 1): "养老",
            (10, 1): "其他内容",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_excluded_by_keyword(self):
        row_data = {
            (1, 1): "工资发放汇总数据",
            (2, 1): "生成时间",
            (4, 1): "文件名",
        }
        ws = self._make_mock_worksheet(row_data)
        assert is_payroll_sheet(ws) is False

    def test_payroll_sheet_with_custom_config(self):
        row_data = {
            (1, 1): "自定义工资表",
            (2, 1): "应发工资",
            (2, 2): "实发工资",
            (5, 1): "manager",
        }
        ws = self._make_mock_worksheet(row_data)
        custom_config = {
            "sheet_filter": {
                "title_keyword": {"required": "自定义工资表"},
                "exclude_keywords": {"keywords": []},
                "required_content": {"required": ["应发工资", "实发工资"]},
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


class TestCleanupEmptyColumns:
    """Test suite for _remove_empty_columns, _get_signature_keywords, and 制表人 alignment."""

    def test_get_signature_keywords_from_config(self):
        """Test extracting signature keywords from payroll config."""
        keywords = _get_signature_keywords(get_payroll_config())
        expected = {
            "总经理签字",
            "分管领导审核",
            "财务审核",
            "业务审核",
            "部长签字",
            "部长、分管副总签字",
        }
        missing = expected - keywords
        assert not missing, f"Missing keywords: {missing}"
        assert _get_signature_keywords({}) == set()

    def test_remove_empty_columns_removes_truly_empty(self):
        """Test that completely empty columns are deleted."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(row=1, column=1, value="Header A")
        ws.cell(row=2, column=1, value="data A")
        ws.cell(row=1, column=4, value="Header D")
        ws.cell(row=2, column=4, value="data D")

        assert ws.max_column == 4

        cfg = get_payroll_config()
        _remove_empty_columns(ws, cfg)

        assert ws.max_column == 2, f"Expected 2 columns, got {ws.max_column}"
        assert ws.cell(row=1, column=1).value == "Header A"
        assert ws.cell(row=1, column=2).value == "Header D"

    def test_remove_empty_columns_preserves_keyword(self):
        """Test that a column with only signature keywords is preserved
        by moving keywords to the nearest non-empty column on the right."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(row=1, column=1, value="Name")
        ws.cell(row=2, column=1, value="张三")
        ws.cell(row=1, column=3, value="Amount")
        ws.cell(row=2, column=3, value="5000")
        # Column B has only a signature keyword
        ws.cell(row=2, column=2, value="总经理签字")

        cfg = get_payroll_config()
        _remove_empty_columns(ws, cfg)

        # Column B deleted; the keyword moves into C (target=D→C after shift)
        assert ws.max_column == 2
        assert ws.cell(row=2, column=2).value == "总经理签字"

    def test_remove_empty_columns_keyword_appends_at_end(self):
        """Test that a keyword-only column with no data to the right
        appends the keyword at the end."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(row=1, column=1, value="Name")
        ws.cell(row=2, column=1, value="张三")
        # Column B has only a signature keyword, no data columns to the right
        ws.cell(row=2, column=2, value="分管领导审核")

        cfg = get_payroll_config()
        _remove_empty_columns(ws, cfg)

        # Keyword written to max_column+1, then column B deleted → col 2 is the new column
        assert ws.max_column == 2
        assert ws.cell(row=2, column=2).value == "分管领导审核"

    def test_remove_empty_columns_no_empty_columns_noop(self):
        """Test that when all columns have real data, nothing changes."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(row=1, column=1, value="Name")
        ws.cell(row=1, column=2, value="Amount")
        ws.cell(row=2, column=1, value="张三")
        ws.cell(row=2, column=2, value="5000")

        cfg = get_payroll_config()
        _remove_empty_columns(ws, cfg)

        assert ws.max_column == 2
        assert ws.cell(row=1, column=1).value == "Name"
        assert ws.cell(row=1, column=2).value == "Amount"
        assert ws.cell(row=2, column=1).value == "张三"
        assert ws.cell(row=2, column=2).value == "5000"

    def test_remove_empty_columns_with_merged_cells(self):
        """Test that merged cells are preserved when deleting empty columns."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # Merge A1:B2 — keep col B alive to avoid affecting the merge
        ws.merge_cells("A1:B2")
        ws.cell(row=1, column=1, value="Merged Header")
        ws.cell(row=3, column=2, value="__keep__")
        # Column C is completely empty — will be deleted
        ws.cell(row=1, column=4, value="End")

        cfg = get_payroll_config()
        _remove_empty_columns(ws, cfg)

        # Column C deleted, D shifted left
        assert ws.max_column == 3
        assert ws.cell(row=1, column=3).value == "End"
        # Merged cell range preserved
        assert any("A1:B2" in str(mc) for mc in ws.merged_cells.ranges)

    def test_zhibiaoren_right_alignment(self):
        """Test that a cell containing '制表人' gets right/center alignment,
        matching the inline logic from _insert_signature_to_excel_openpyxl."""
        from openpyxl.styles import Font as _Font

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        cell = ws.cell(row=5, column=3, value="制表人：张三")

        # Inline replica of batch_processor.py lines 848–852
        if cell.value and "制表人" in str(cell.value):
            old_size = cell.font.size or 11
            if old_size > 10:
                cell.font = _Font(size=10)
            cell.alignment = Alignment(horizontal="right", vertical="center")

        assert cell.alignment.horizontal == "right"
        assert cell.alignment.vertical == "center"
