"""Canonical allocation extraction records and parser results."""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import Any, Dict, List, Optional

from scm_allocation.models.reference import AllocationContext


class ExtractionStatus(str, Enum):
    """Status for an extracted record or an extraction operation."""

    VALID = "valid"
    PARTIAL = "partial"
    INVALID = "invalid"
    SKIPPED = "skipped"
    UNSUPPORTED = "unsupported"


@dataclass(frozen=True)
class AllocationRecord:
    """A normalized allocation row with source traceability."""

    material_code: Optional[str]
    material_description: Optional[str]
    quantity: Optional[Decimal]
    destination_plant_code: Optional[str]
    destination_plant_description: Optional[str]
    issuing_warehouse_code: Optional[str]
    issuing_warehouse_description: Optional[str]
    allocation_context: AllocationContext
    source_email_id: Optional[str]
    source_filename: Optional[str]
    source_row: Optional[int]
    source_location: Optional[str]
    parser_type: str
    extraction_status: ExtractionStatus
    extraction_confidence: float
    errors: tuple[str, ...] = field(default_factory=tuple)
    source_sheet: Optional[str] = None


@dataclass
class ExtractionResult:
    """Result of parsing one source, including controlled parser errors."""

    records: List[AllocationRecord]
    status: ExtractionStatus
    parser_type: str
    source_email_id: Optional[str] = None
    errors: tuple[str, ...] = field(default_factory=tuple)
    metadata: Dict[str, Any] = field(default_factory=dict)