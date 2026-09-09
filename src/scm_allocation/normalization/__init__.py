"""Shared normalization helpers for allocation extraction."""

from scm_allocation.normalization.fields import (
	FIELD_ALIASES,
	canonical_field_name,
	find_alias_index,
	looks_like_warehouse_code,
	normalize_material_code,
	normalize_quantity,
	normalize_text,
)
from scm_allocation.normalization.storage_locations import resolve_storage_locations

__all__ = [
	"FIELD_ALIASES",
	"canonical_field_name",
	"find_alias_index",
	"looks_like_warehouse_code",
	"normalize_material_code",
	"normalize_quantity",
	"normalize_text",
	"resolve_storage_locations",
]
