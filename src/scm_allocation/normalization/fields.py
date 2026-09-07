"""Deterministic normalization helpers shared by allocation parsers."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re
from typing import Any, Iterable, Mapping, Optional


FIELD_ALIASES: Mapping[str, tuple[str, ...]] = {
    "material_code": (
        "article code",
        "article sap",
        "material",
        "material code",
        "material sap",
        "sku",
    ),
    "material_description": (
        "article description",
        "article description sap",
        "article desc",
        "material desc",
        "material desc.",
        "row labels",
        "brand type 2",
        "material description",
        "desc",
    ),
    "destination_plant_code": ("to store site code", "site code", "code store", "plan", "plant", "plant code"),
    "destination_plant_description": ("to store site desc", "site desc", "store", "plant_name", "plant desc"),
    "issuing_warehouse_code": ("from site code", "from warehouse", "warehouse code"),
    "issuing_warehouse_description": ("from site desc", "warehouse description"),
    "quantity": ("alokasi", "po", "request", "qty", "qty request", "quantity"),
}


def normalize_text(value: Any) -> Optional[str]:
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def normalize_material_code(value: Any) -> Optional[str]:
    text = normalize_text(value)
    if text is None or text.lower() in {"grand total", "total"}:
        return None
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def normalize_quantity(value: Any) -> Optional[Decimal]:
    text = normalize_text(value)
    if text is None:
        return None
    text = text.replace(",", "")
    try:
        quantity = Decimal(text)
    except (InvalidOperation, ValueError):
        return None
    return quantity if quantity >= 0 else None


def canonical_field_name(value: Any) -> Optional[str]:
    normalized = normalize_text(value)
    if normalized is None:
        return None
    normalized = normalized.casefold()
    for field_name, aliases in FIELD_ALIASES.items():
        if normalized in aliases:
            return field_name
    return None


def find_alias_index(headers: Iterable[Any], field_name: str) -> Optional[int]:
    aliases = {alias.casefold() for alias in FIELD_ALIASES[field_name]}
    for index, header in enumerate(headers):
        normalized = normalize_text(header)
        if normalized and normalized.casefold() in aliases:
            return index
    return None


def looks_like_warehouse_code(value: Any) -> bool:
    text = normalize_text(value)
    return bool(text and re.fullmatch(r"[A-Za-z]{2,5}\d{1,3}", text))