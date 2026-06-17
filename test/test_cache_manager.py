"""
Tests for cache manager module.

TDD test suite covering cache hit, miss, expiry, force refresh,
and Feishu-specific TTL requirements (12-hour download URLs).
"""

import json
import time
from pathlib import Path

import pytest

from app.cache_manager import (
    BaseFileCache,
    DownloadURLCache,
    InstanceDetailCache,
)


class TestBaseFileCache:
    """Test suite for BaseFileCache shared behaviour."""

    def test_cache_hit(self, tmp_path):
        """Cached value returned, no HTTP request needed."""
        cache = BaseFileCache(cache_dir=str(tmp_path), default_ttl=3600)
        cache.set("key1", {"data": "value1"})
        result = cache.get("key1")
        assert result == {"data": "value1"}
        assert cache.hits == 1
        assert cache.misses == 0

    def test_cache_miss(self, tmp_path):
        """No cache entry → fetch from API (simulated by returning None)."""
        cache = BaseFileCache(cache_dir=str(tmp_path), default_ttl=3600)
        result = cache.get("missing_key")
        assert result is None
        assert cache.hits == 0
        assert cache.misses == 1

    def test_cache_expired(self, tmp_path):
        """Expired cache entry → fetch from API (simulated by returning None)."""
        cache = BaseFileCache(cache_dir=str(tmp_path), default_ttl=3600)
        cache.set("key1", {"data": "value1"}, ttl=-1)
        result = cache.get("key1")
        assert result is None
        assert cache.hits == 0
        assert cache.misses == 1

    def test_force_refresh(self, tmp_path):
        """force=True ignores cache and acts as a miss."""
        cache = BaseFileCache(cache_dir=str(tmp_path), default_ttl=3600)
        cache.set("key1", {"data": "value1"})
        result = cache.get("key1", force=True)
        assert result is None
        assert cache.hits == 0
        assert cache.misses == 1

    def test_clear_removes_entries_and_resets_stats(self, tmp_path):
        """Test clear removes all entries and resets statistics."""
        cache = BaseFileCache(cache_dir=str(tmp_path), default_ttl=3600)
        cache.set("key1", {"data": "value1"})
        cache.get("key1")
        assert cache.hits == 1
        cache.clear()
        assert cache.hits == 0
        assert cache.misses == 0
        assert cache.get("key1") is None
        assert cache.misses == 1

    def test_default_ttl(self):
        """Test default TTL is passed correctly."""
        cache = BaseFileCache(cache_dir=".cache", default_ttl=7200)
        assert cache.default_ttl == 7200


class TestDownloadURLCache:
    """Test suite for DownloadURLCache (Feishu temporary URLs)."""

    def test_download_url_ttl_12h(self, tmp_path):
        """Verify DownloadURLCache default TTL is 43200 seconds (12 hours)."""
        cache = DownloadURLCache(cache_dir=str(tmp_path))
        assert cache.default_ttl == 43200

    def test_cache_hit_uses_url_as_key(self, tmp_path):
        """Cache hit when URL is already cached."""
        cache = DownloadURLCache(cache_dir=str(tmp_path))
        url = "https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/authcode/?code=abc123"
        cache.set(url, {"filename": "report.xlsx"})
        result = cache.get(url)
        assert result == {"filename": "report.xlsx"}
        assert cache.hits == 1

    def test_cache_miss_for_new_url(self, tmp_path):
        """Cache miss for uncached URL."""
        cache = DownloadURLCache(cache_dir=str(tmp_path))
        url = "https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/authcode/?code=newcode"
        result = cache.get(url)
        assert result is None
        assert cache.misses == 1

    def test_force_refresh_bypasses_cache(self, tmp_path):
        """force=True ignores cached download URL."""
        cache = DownloadURLCache(cache_dir=str(tmp_path))
        url = "https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/authcode/?code=oldcode"
        cache.set(url, {"filename": "old.xlsx"})
        result = cache.get(url, force=True)
        assert result is None
        assert cache.misses == 1


class TestInstanceDetailCache:
    """Test suite for InstanceDetailCache (Feishu approval instances)."""

    def test_cache_hit_by_instance_code(self, tmp_path):
        """Cache hit when instance_code is already cached."""
        cache = InstanceDetailCache(cache_dir=str(tmp_path))
        instance_code = "7650052824440671208"
        detail = {
            "status": "APPROVED",
            "form": [
                {"type": "attachmentV2", "name": "附件", "value": ["https://example.com"]}
            ],
        }
        cache.set(instance_code, detail)
        result = cache.get(instance_code)
        assert result == detail
        assert cache.hits == 1

    def test_cache_miss_for_unknown_instance(self, tmp_path):
        """Cache miss for unknown instance_code."""
        cache = InstanceDetailCache(cache_dir=str(tmp_path))
        result = cache.get("unknown_instance")
        assert result is None
        assert cache.misses == 1

    def test_cache_expired_instance_detail(self, tmp_path):
        """Expired instance detail cache returns None."""
        cache = InstanceDetailCache(cache_dir=str(tmp_path))
        cache.set("inst1", {"status": "PENDING"}, ttl=-1)
        result = cache.get("inst1")
        assert result is None
        assert cache.misses == 1

    def test_force_refresh_instance_detail(self, tmp_path):
        """force=True ignores cached instance detail."""
        cache = InstanceDetailCache(cache_dir=str(tmp_path))
        cache.set("inst1", {"status": "APPROVED"})
        result = cache.get("inst1", force=True)
        assert result is None
        assert cache.misses == 1

    def test_clear_removes_instance_entries(self, tmp_path):
        """Test clear removes instance detail entries."""
        cache = InstanceDetailCache(cache_dir=str(tmp_path))
        cache.set("inst1", {"status": "APPROVED"})
        cache.clear()
        assert cache.get("inst1") is None
        assert cache.hits == 0

    def test_json_file_storage_format(self, tmp_path):
        """Verify cache stores data as JSON with metadata."""
        cache = InstanceDetailCache(cache_dir=str(tmp_path))
        cache.set("inst1", {"status": "APPROVED"})
        # Find the cached file
        cache_file = list(tmp_path.glob("*.json"))[0]
        with open(cache_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert "value" in data
        assert "expires_at" in data
        assert "cached_at" in data
        assert data["value"]["status"] == "APPROVED"
