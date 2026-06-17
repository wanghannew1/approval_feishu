"""
Tests for cache manager module.
"""

import pytest
from app.cache_manager import TokenCache


class TestCacheManager:
    """Test suite for cache_manager module."""

    def test_token_cache_init(self):
        """Test TokenCache initialization."""
        cache = TokenCache()
        assert cache.cache_path.name == ".token_cache.json"

    def test_token_cache_custom_path(self):
        """Test TokenCache with custom path."""
        cache = TokenCache(cache_path="/tmp/custom_cache.json")
        assert str(cache.cache_path) == "/tmp/custom_cache.json"

    def test_load_returns_none_when_empty(self):
        """Test that load returns None for missing/expired cache."""
        cache = TokenCache()
        result = cache.load()
        assert result is None

    def test_clear_does_not_raise(self):
        """Test that clear handles missing cache gracefully."""
        cache = TokenCache()
        cache.clear()  # Should not raise
