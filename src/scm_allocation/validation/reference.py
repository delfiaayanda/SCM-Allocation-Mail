"""Reference data provider abstractions and routing implementations."""

from __future__ import annotations

from abc import ABC, abstractmethod
import json
from pathlib import Path
from typing import Any, Dict, List, Optional
import urllib.error
import urllib.request

from scm_allocation.models.reference import AllocationContext, RoutingRow, RoutingSnapshot


class ReferenceDataProvider(ABC):
    """Abstract base class for master data and reference providers."""

    @abstractmethod
    def get_data(self) -> Any:
        """Retrieve reference dataset."""
        pass


class RoutingProvider(ReferenceDataProvider):
    """Abstract base class for allocation routing rule providers."""

    @abstractmethod
    def get_routing_snapshot(self, force_refresh: bool = False) -> RoutingSnapshot:
        """Obtain a current snapshot of corporate routing rules."""
        pass

    def get_data(self) -> RoutingSnapshot:
        return self.get_routing_snapshot()

    def resolve_sender_wh(
        self,
        plant_code: str,
        context: AllocationContext = AllocationContext.UNKNOWN,
    ) -> Optional[str]:
        """Convenience method to query sender warehouse for a plant."""
        snapshot = self.get_routing_snapshot()
        return snapshot.get_sender_wh(plant_code, context)


class LocalSnapshotRoutingProvider(RoutingProvider):
    """Routing provider that serves pre-recorded snapshots from local JSON files or memory."""

    def __init__(self, snapshot_data: Optional[Dict[str, Any]] = None, filepath: Optional[Path] = None):
        self._snapshot_data = snapshot_data
        self._filepath = filepath
        self._cached_snapshot: Optional[RoutingSnapshot] = None

    def get_routing_snapshot(self, force_refresh: bool = False) -> RoutingSnapshot:
        if self._cached_snapshot and not force_refresh:
            return self._cached_snapshot

        raw: Dict[str, Any] = {}
        if self._snapshot_data is not None:
            raw = self._snapshot_data
        elif self._filepath is not None and self._filepath.exists():
            with open(self._filepath, "r", encoding="utf-8") as f:
                raw = json.load(f)
        else:
            raw = {"rows": [], "fetchedAt": "", "hash": None, "log": None}

        rows = [RoutingRow.from_dict(item) for item in raw.get("rows", [])]
        self._cached_snapshot = RoutingSnapshot(
            fetched_at=raw.get("fetchedAt", ""),
            source=f"local_snapshot:{self._filepath.name if self._filepath else 'memory'}",
            rows=rows,
            hash_code=raw.get("hash"),
            log=raw.get("log"),
        )
        return self._cached_snapshot


class ApexRoutingProvider(RoutingProvider):
    """Routing provider that synchronizes with the APEX RICO API endpoint."""

    DEFAULT_URL = "https://api.apexrico.space/api/routing/sync"

    def __init__(
        self,
        api_url: str = DEFAULT_URL,
        cache_path: Optional[Path] = None,
        timeout: float = 30.0,
        auth_token: Optional[str] = None,
    ):
        self.api_url = api_url
        self.cache_path = cache_path
        self.timeout = timeout
        self.auth_token = auth_token
        self._memory_cache: Optional[RoutingSnapshot] = None

    def get_routing_snapshot(self, force_refresh: bool = False) -> RoutingSnapshot:
        if self._memory_cache and not force_refresh:
            return self._memory_cache

        # If cache file exists and no force_refresh, try loading from cache
        if not force_refresh and self.cache_path and self.cache_path.exists():
            try:
                local_provider = LocalSnapshotRoutingProvider(filepath=self.cache_path)
                self._memory_cache = local_provider.get_routing_snapshot()
                return self._memory_cache
            except Exception:
                pass

        # Otherwise perform sync
        raw_payload = self.fetch_remote_sync()
        rows = [RoutingRow.from_dict(item) for item in raw_payload.get("rows", [])]

        snapshot = RoutingSnapshot(
            fetched_at=raw_payload.get("fetchedAt", ""),
            source="apex_api",
            rows=rows,
            hash_code=raw_payload.get("hash"),
            log=raw_payload.get("log"),
        )

        # Write to cache file if configured
        if self.cache_path:
            try:
                self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                with open(self.cache_path, "w", encoding="utf-8") as f:
                    json.dump(raw_payload, f, indent=2, ensure_ascii=False)
            except Exception:
                pass

        self._memory_cache = snapshot
        return snapshot

    def fetch_remote_sync(self) -> Dict[str, Any]:
        """Execute HTTP POST request against APEX RICO endpoint."""
        req = urllib.request.Request(
            self.api_url,
            data=b"{}",
            headers={
                "Content-Type": "application/json",
                "User-Agent": "SCM-Allocation-Mail/0.1.0",
                **({"Authorization": f"Bearer {self.auth_token}"} if self.auth_token else {}),
            },
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as response:
                payload = json.loads(response.read().decode("utf-8"))
                return payload
        except urllib.error.URLError as err:
            raise RuntimeError(f"Failed connecting to APEX Routing API at {self.api_url}: {err}") from err
