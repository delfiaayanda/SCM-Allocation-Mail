"""Unit tests for routing reference models and provider abstractions."""

from pathlib import Path
from unittest.mock import MagicMock
import pytest

from scm_allocation.models.reference import AllocationContext, RoutingRow, RoutingSnapshot
from scm_allocation.validation.reference import ApexRoutingProvider, LocalSnapshotRoutingProvider


def test_allocation_context_resolution():
    assert AllocationContext.from_str("Accessories") == AllocationContext.ACCESSORIES
    assert AllocationContext.from_str("Aksesoris") == AllocationContext.ACCESSORIES
    assert AllocationContext.from_str("Device") == AllocationContext.DEVICE
    assert AllocationContext.from_str("handphone") == AllocationContext.DEVICE
    assert AllocationContext.from_str("NPI") == AllocationContext.NPI
    assert AllocationContext.from_str("TV") == AllocationContext.TV
    assert AllocationContext.from_str("Unknown") == AllocationContext.UNKNOWN


def test_routing_row_from_dict():
    raw = {
        "PT": "EAR",
        "Plant code": "X071",
        "Plant Desc": "IBOX PLAZA MEDAN FAIR",
        "WH Sender Code (Alokasi Device)": "BIC1",
        "Plant Code (Alokasi Device)": "BIC1",
        "WH Sender Code (Alokasi Accs)": "BI04",
        "Plant Code (Alokasi Accs)": "BI04",
        "JADWAL Device": "RABU, JUMAT",
    }
    row = RoutingRow.from_dict(raw)
    assert row.pt == "EAR"
    assert row.plant_code == "X071"
    assert row.plant_desc == "IBOX PLAZA MEDAN FAIR"
    assert row.get_sender_wh(AllocationContext.DEVICE) == "BIC1"
    assert row.get_sender_wh(AllocationContext.ACCESSORIES) == "BI04"


def test_routing_snapshot_indexing_and_context_lookup():
    rows = [
        RoutingRow(
            plant_code="X071",
            plant_desc="IBOX MEDAN FAIR",
            wh_sender_device="BIC1",
            wh_sender_accs="BI04",
        ),
        RoutingRow(
            plant_code="XE05",
            plant_desc="BCA EXPO IBOX",
            wh_sender_device="BIC1",
            wh_sender_accs="BIC1",
        ),
    ]
    snapshot = RoutingSnapshot(
        fetched_at="2026-09-07T12:00:00Z",
        source="test_fixture",
        rows=rows,
        hash_code="test_hash_123",
    )

    assert snapshot.total_rows == 2
    # Case-insensitive plant lookup
    matches = snapshot.find_by_plant_code("x071")
    assert len(matches) == 1
    assert matches[0].plant_desc == "IBOX MEDAN FAIR"

    # Context-dependent routing validation
    assert snapshot.get_sender_wh("X071", AllocationContext.DEVICE) == "BIC1"
    assert snapshot.get_sender_wh("X071", AllocationContext.ACCESSORIES) == "BI04"
    assert snapshot.get_sender_wh("XE05", AllocationContext.ACCESSORIES) == "BIC1"
    assert snapshot.get_sender_wh("NON_EXISTENT") is None


def test_local_snapshot_provider(tmp_path: Path):
    sample_payload = {
        "fetchedAt": "2026-09-07T10:00:00Z",
        "hash": "abc456",
        "rows": [
            {
                "Plant code": "X115",
                "Plant Desc": "IBOX YOS SUDARSO TARAKAN",
                "WH Sender Code (Alokasi Device)": "BIC1",
            }
        ],
    }
    provider = LocalSnapshotRoutingProvider(snapshot_data=sample_payload)
    snapshot = provider.get_routing_snapshot()

    assert snapshot.total_rows == 1
    assert provider.resolve_sender_wh("X115", AllocationContext.DEVICE) == "BIC1"


def test_apex_routing_provider_with_mocked_sync(tmp_path: Path):
    cache_file = tmp_path / "cache" / "routing.json"
    provider = ApexRoutingProvider(cache_path=cache_file)

    mock_remote_data = {
        "fetchedAt": "2026-09-07T14:30:00Z",
        "hash": "remote_hash_999",
        "rows": [
            {
                "Plant code": "XE05",
                "Plant Desc": "BCA Expo",
                "WH Sender Code (Alokasi Accs)": "BIC1",
            }
        ],
    }
    # Mock network call
    provider.fetch_remote_sync = MagicMock(return_value=mock_remote_data)

    snapshot = provider.get_routing_snapshot()
    assert snapshot.total_rows == 1
    assert snapshot.hash_code == "remote_hash_999"
    assert cache_file.exists()

    # Second call should load from memory without calling fetch_remote_sync again
    cached_snapshot = provider.get_routing_snapshot()
    assert cached_snapshot.total_rows == 1
    provider.fetch_remote_sync.assert_called_once()
