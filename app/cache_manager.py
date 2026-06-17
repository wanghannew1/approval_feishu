"""
Cache manager for Feishu API operations.

Handles caching of:
- Tenant access tokens (persistent file storage with expiry)
- Download URLs (12-hour TTL matching Feishu temporary URL validity)
- Approval instance details (form JSON parsing with configurable TTL)

All caches use JSON file storage and support force refresh plus
hit/miss statistics.
"""

import json
import time
from pathlib import Path
from typing import Any, Optional


class BaseFileCache:
    """
    Base class for file-based JSON caches with TTL and statistics.
    """

    def __init__(self, cache_dir: str = ".cache", default_ttl: int = 3600):
        """
        Initialize the file cache.

        Args:
            cache_dir: Directory to store cache JSON files.
            default_ttl: Default time-to-live in seconds.
        """
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.default_ttl = default_ttl
        self._hits = 0
        self._misses = 0

    def _cache_file(self, key: str) -> Path:
        """
        Derive a filesystem-safe cache file path from a key.

        Args:
            key: Cache key (e.g. instance_code or download URL).

        Returns:
            Path to the JSON cache file.
        """
        # Use MD5 hash for keys that may contain long URLs or special chars
        import hashlib

        safe_key = hashlib.md5(key.encode("utf-8")).hexdigest()
        return self.cache_dir / f"{safe_key}.json"

    def get(self, key: str, force: bool = False) -> Optional[Any]:
        """
        Retrieve a cached value if present and not expired.

        Args:
            key: Cache key.
            force: If True, bypass cache and treat as a miss.

        Returns:
            Cached value or None on miss/force/expiry.
        """
        if force:
            self._misses += 1
            return None

        cache_file = self._cache_file(key)
        if not cache_file.exists():
            self._misses += 1
            return None

        try:
            with cache_file.open("r", encoding="utf-8") as f:
                data = json.load(f)
            expires_at = data.get("expires_at", 0)
            if expires_at < time.time():
                self._misses += 1
                return None
            self._hits += 1
            return data.get("value")
        except (json.JSONDecodeError, OSError):
            self._misses += 1
            return None

    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> None:
        """
        Store a value in the cache with an optional TTL.

        Args:
            key: Cache key.
            value: JSON-serialisable value to cache.
            ttl: Time-to-live in seconds (defaults to ``default_ttl``).
        """
        cache_file = self._cache_file(key)
        data = {
            "value": value,
            "expires_at": time.time() + (ttl if ttl is not None else self.default_ttl),
            "cached_at": time.time(),
        }
        try:
            with cache_file.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def clear(self) -> None:
        """Remove all cached entries and reset statistics."""
        for cache_file in self.cache_dir.glob("*.json"):
            try:
                cache_file.unlink()
            except OSError:
                pass
        self._hits = 0
        self._misses = 0

    @property
    def hits(self) -> int:
        """Number of cache hits since initialisation or last ``clear``."""
        return self._hits

    @property
    def misses(self) -> int:
        """Number of cache misses since initialisation or last ``clear``."""
        return self._misses


class DownloadURLCache(BaseFileCache):
    """
    Cache for Feishu temporary download URLs.

    Feishu download URLs are valid for 12 hours (43_200 seconds).
    """

    def __init__(self, cache_dir: str = ".cache/download_urls"):
        super().__init__(cache_dir, default_ttl=43200)


class InstanceDetailCache(BaseFileCache):
    """
    Cache for Feishu approval instance details (form JSON, attachments, etc.).
    """

    def __init__(self, cache_dir: str = ".cache/instance_details"):
        super().__init__(cache_dir, default_ttl=3600)
