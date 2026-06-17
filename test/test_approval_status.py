"""
Tests for approval status querying.
"""

import pytest
from app.feishu_api import query_approval_instances


class TestApprovalStatus:
    """Test suite for approval status functionality."""

    def test_query_instances_returns_dict(self, mock_auth_headers: dict):
        """Test that query returns a dictionary."""
        # TODO: Implement with requests-mock
        pass

    def test_instance_detail_parsing(self):
        """Test parsing of instance detail response."""
        # TODO: Implement
        pass

    def test_status_filtering(self):
        """Test filtering instances by status."""
        # TODO: Implement
        pass
