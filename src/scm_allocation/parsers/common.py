"""Shared construction helpers for allocation parsers."""

from __future__ import annotations

from typing import Any, Optional

from scm_allocation.models.allocation import AllocationRecord, ExtractionStatus
from scm_allocation.models.reference import AllocationContext
from scm_allocation.normalization import normalize_material_code, normalize_quantity, normalize_text


def infer_allocation_context(*values: Any) -> AllocationContext:
    """Infer context only from unambiguous product/category text."""
    text = " ".join(filter(None, (normalize_text(value) for value in values))).casefold()
    if any(term in text for term in ("logitech", "belkin", "charger", "cable", "case", "accessor")):
        return AllocationContext.ACCESSORIES
    if any(term in text for term in ("iphone", "ipad", "macbook", "mba ", "mac ")):
        return AllocationContext.DEVICE
    return AllocationContext.UNKNOWN


def make_record(
    *,
    material_code: Any,
    material_description: Any,
    quantity: Any,
    destination_plant_code: Any,
    destination_plant_description: Any,
    issuing_warehouse_code: Any,
    issuing_warehouse_description: Any,
    allocation_context: AllocationContext,
    source_email_id: Optional[str],
    source_filename: Optional[str],
    source_sheet: Optional[str],
    source_row: Optional[int],
    source_location: Optional[str],
    parser_type: str,
    confidence: float,
    extra_errors: tuple[str, ...] = (),
) -> AllocationRecord:
    normalized_material = normalize_material_code(material_code)
    normalized_description = normalize_text(material_description)
    normalized_quantity = normalize_quantity(quantity)
    normalized_plant = normalize_text(destination_plant_code)
    normalized_plant_description = normalize_text(destination_plant_description)
    normalized_warehouse = normalize_text(issuing_warehouse_code)
    normalized_warehouse_description = normalize_text(issuing_warehouse_description)

    errors = list(extra_errors)
    if normalized_material is None:
        errors.append("missing material_code")
    if normalized_quantity is None:
        errors.append("missing or invalid quantity")
    if normalized_plant is None:
        errors.append("missing destination_plant_code")

    status = ExtractionStatus.VALID if not errors else ExtractionStatus.PARTIAL
    if normalized_material is None and normalized_quantity is None:
        status = ExtractionStatus.INVALID

    return AllocationRecord(
        material_code=normalized_material,
        material_description=normalized_description,
        quantity=normalized_quantity,
        destination_plant_code=normalized_plant,
        destination_plant_description=normalized_plant_description,
        issuing_warehouse_code=normalized_warehouse,
        issuing_warehouse_description=normalized_warehouse_description,
        allocation_context=allocation_context,
        source_email_id=source_email_id,
        source_filename=source_filename,
        source_sheet=source_sheet,
        source_row=source_row,
        source_location=source_location,
        parser_type=parser_type,
        extraction_status=status,
        extraction_confidence=max(0.0, min(1.0, confidence if not errors else confidence - 0.2)),
        errors=tuple(errors),
    )