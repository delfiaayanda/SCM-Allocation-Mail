"""Canonical data models for reference data, routing tables, and snapshots."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


class AllocationContext(str, Enum):
    """Business context for allocation routing (different products route to different WHs)."""
    DEVICE = "device"
    ACCESSORIES = "accessories"
    NPI = "npi"
    TV = "tv"
    UNKNOWN = "unknown"

    @classmethod
    def from_str(cls, value: str) -> AllocationContext:
        norm = (value or "").strip().lower()
        if "acc" in norm or "accessories" in norm or "aksesoris" in norm:
            return cls.ACCESSORIES
        if "device" in norm or "hp" in norm or "handphone" in norm:
            return cls.DEVICE
        if "npi" in norm:
            return cls.NPI
        if "tv" in norm:
            return cls.TV
        return cls.UNKNOWN


@dataclass(frozen=True)
class RoutingRow:
    """Represents a single master routing rule row from SAP / APEX RICO."""

    pt: str = ""
    region_sales: str = ""
    region: str = ""
    area: str = ""
    city: str = ""
    sloc: str = ""
    plant_code: str = ""
    plant_desc: str = ""

    # Forward routing senders by context
    wh_sender_device: str = ""
    plant_code_device: str = ""
    plant_desc_device: str = ""
    jadwal_device: str = ""

    wh_sender_accs: str = ""
    plant_code_accs: str = ""
    plant_desc_accs: str = ""
    jadwal_accs: str = ""

    wh_sender_npi: str = ""
    plant_code_npi: str = ""
    plant_desc_npi: str = ""

    wh_sender_tv: str = ""
    plant_code_tv: str = ""
    plant_desc_tv: str = ""

    # Returns
    sloc_retur_device: str = ""
    plant_code_retur_device: str = ""
    plant_desc_retur_device: str = ""
    sloc_retur_accs: str = ""
    plant_code_retur_accs: str = ""
    plant_desc_retur_accs: str = ""

    raw_data: Dict[str, Any] = field(default_factory=dict, repr=False)

    def get_sender_wh(self, context: AllocationContext) -> Optional[str]:
        """Resolve expected warehouse sender code based on allocation context."""
        if context == AllocationContext.DEVICE:
            return self.wh_sender_device or None
        if context == AllocationContext.ACCESSORIES:
            return self.wh_sender_accs or None
        if context == AllocationContext.NPI:
            return self.wh_sender_npi or None
        if context == AllocationContext.TV:
            return self.wh_sender_tv or None
        # Default fallback: return device sender if available, else accs
        return self.wh_sender_device or self.wh_sender_accs or None

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> RoutingRow:
        """Instantiate RoutingRow from raw API dictionary keys."""
        return cls(
            pt=str(data.get("PT", "")).strip(),
            region_sales=str(data.get("REGION SALES", "")).strip(),
            region=str(data.get("Region", "")).strip(),
            area=str(data.get("Area", "")).strip(),
            city=str(data.get("City", "")).strip(),
            sloc=str(data.get("Sloc", "")).strip(),
            plant_code=str(data.get("Plant code", "")).strip(),
            plant_desc=str(data.get("Plant Desc", "")).strip(),
            wh_sender_device=str(data.get("WH Sender Code (Alokasi Device)", "")).strip(),
            plant_code_device=str(data.get("Plant Code (Alokasi Device)", "")).strip(),
            plant_desc_device=str(data.get("Plant Desc (Alokasi Device)", "")).strip(),
            jadwal_device=str(data.get("JADWAL Device", "")).strip(),
            wh_sender_accs=str(data.get("WH Sender Code (Alokasi Accs)", "")).strip(),
            plant_code_accs=str(data.get("Plant Code (Alokasi Accs)", "")).strip(),
            plant_desc_accs=str(data.get("Plant Desc (Alokasi Accs)", "")).strip(),
            jadwal_accs=str(data.get("JADWAL Accs", "")).strip(),
            wh_sender_npi=str(data.get("WH Sender Code (Alokasi NPI)", "")).strip(),
            plant_code_npi=str(data.get("Plant Code (Alokasi NPI)", "")).strip(),
            plant_desc_npi=str(data.get("Plant Desc (Alokasi NPI)", "")).strip(),
            wh_sender_tv=str(data.get("WH Sender Code (Alokasi TV)", "")).strip(),
            plant_code_tv=str(data.get("Plant Code (Alokasi TV)", "")).strip(),
            plant_desc_tv=str(data.get("Plant Desc (Alokasi TV)", "")).strip(),
            sloc_retur_device=str(data.get("Sloc Retur (Device)", "")).strip(),
            plant_code_retur_device=str(data.get("Plant Code Retur (Device)", "")).strip(),
            plant_desc_retur_device=str(data.get("Plant Desc Retur (Device)", "")).strip(),
            sloc_retur_accs=str(data.get("Sloc Retur (Accs)", "")).strip(),
            plant_code_retur_accs=str(data.get("Plant Code Retur (Accs)", "")).strip(),
            plant_desc_retur_accs=str(data.get("Plant Desc Retur (Accs)", "")).strip(),
            raw_data=data,
        )


@dataclass
class RoutingSnapshot:
    """Immutable in-memory snapshot of corporate allocation routing rules."""

    fetched_at: str
    source: str
    rows: List[RoutingRow]
    hash_code: Optional[str] = None
    log: Optional[str] = None
    _index_by_plant: Dict[str, List[RoutingRow]] = field(default_factory=dict, init=False, repr=False)

    def __post_init__(self):
        # Index rows by uppercase plant_code for O(1) lookup
        self._index_by_plant = {}
        for r in self.rows:
            key = r.plant_code.upper()
            if key not in self._index_by_plant:
                self._index_by_plant[key] = []
            self._index_by_plant[key].append(r)

    @property
    def total_rows(self) -> int:
        return len(self.rows)

    def find_by_plant_code(self, plant_code: str) -> List[RoutingRow]:
        """Find all routing entries for a specific plant code."""
        return self._index_by_plant.get(plant_code.strip().upper(), [])

    def get_sender_wh(
        self,
        plant_code: str,
        context: AllocationContext = AllocationContext.UNKNOWN,
    ) -> Optional[str]:
        """Lookup designated warehouse sender for a given plant code and allocation context."""
        matches = self.find_by_plant_code(plant_code)
        if not matches:
            return None
        # Return first matching rule with defined sender for context
        for row in matches:
            sender = row.get_sender_wh(context)
            if sender:
                return sender
        return None
