"""Central, conservative defaults for allocation storage locations."""

from __future__ import annotations

from typing import Any, Optional

from scm_allocation.models.reference import AllocationContext
from scm_allocation.normalization.fields import normalize_text


CENTRAL_SOURCE_CODES = frozenset({"BGC1", "BHC1", "BIC1", "BFC1"})
DEVICE_SOURCE_SLOC = "1005"
STANDARD_SLOC = "1001"


def resolve_storage_locations(
    *,
    issuing_warehouse_code: Any,
    allocation_context: AllocationContext,
    explicit_issuing_sloc: Any = None,
    explicit_destination_sloc: Any = None,
    destination_plant_code: Any = None,
) -> tuple[Optional[str], Optional[str]]:
    """Preserve explicit values; otherwise apply documented central/store defaults."""
    source_sloc = normalize_text(explicit_issuing_sloc)
    destination_sloc = normalize_text(explicit_destination_sloc)
    source_code = (normalize_text(issuing_warehouse_code) or "").upper()

    if source_sloc is None and source_code in CENTRAL_SOURCE_CODES:
        source_sloc = DEVICE_SOURCE_SLOC if allocation_context is AllocationContext.DEVICE else STANDARD_SLOC
    if destination_sloc is None and normalize_text(destination_plant_code):
        destination_sloc = STANDARD_SLOC
    return source_sloc, destination_sloc
