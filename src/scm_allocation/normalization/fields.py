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
        "item code",
        "item_code",
        "kode item",
        "kode barang",
        "item",
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
        "item description",
        "item desc",
        "nama barang",
        "deskripsi item",
        "description",
    ),
    "destination_plant_code": (
        "to store site code",
        "site code",
        "code store",
        "destinasi code",
        "plan",
        "plant",
        "plant code",
        "store code",
        "kode store",
        "kode toko",
    ),
    "destination_plant_description": (
        "to store site desc",
        "site desc",
        "destinasi desc",
        "store",
        "plant_name",
        "plant desc",
        "store name",
        "store desc",
        "nama store",
        "nama toko",
        "store description",
    ),
    "issuing_warehouse_code": (
        "from site code",
        "from warehouse",
        "warehouse code",
        "wh code",
        "plant asal",
        "issuing warehouse",
    ),
    "issuing_warehouse_description": (
        "from site desc",
        "warehouse description",
        "wh desc",
    ),
    "quantity": (
        "alokasi",
        "po",
        "request",
        "qty",
        "qty request",
        "quantity",
        "sum of qty",
        "jumlah",
    ),
}


def normalize_text(value: Any) -> Optional[str]:
    if value is None:
        return None

    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def normalize_material_code(value: Any) -> Optional[str]:
    if value is None:
        return None

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.0f}"

    if isinstance(value, int):
        return str(value)

    text = normalize_text(value)
    if text is None:
        return None

    lowered = text.casefold()
    if lowered in {"grand total", "total", "item code", "article code", "material", "material code", "sku"}:
        return None

    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]

    if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", text):
        try:
            val = float(text)
            if val.is_integer():
                return str(int(val))
            return f"{val:.0f}"
        except (ValueError, OverflowError):
            pass

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


def match_semantic_field(header: Any) -> Optional[str]:
    text = normalize_text(header)
    if text is None:
        return None

    canonical = canonical_field_name(text)
    if canonical:
        return canonical

    normalized = re.sub(r"[^\w\s]", " ", text.casefold())
    normalized = re.sub(r"\s+", " ", normalized).strip()
    words = set(normalized.split())

    has_word = lambda *terms: any(bool(re.search(rf"\b{re.escape(term)}\b", normalized)) for term in terms)

    is_source = has_word("from", "asal", "source", "issuing", "supplying", "sender", "origin")

    # 1. Quantity
    if has_word("quantity", "qty", "alokasi", "allocation", "request", "requested", "allocated", "units", "jumlah", "vol", "volume", "pcs", "count"):
        if not has_word("date", "time", "status", "by", "user", "type", "code", "id", "no", "number"):
            return "quantity"

    # 2. Material / SKU / Product Code
    if not is_source and has_word("sku", "material", "article", "product", "item", "barang", "part"):
        if not has_word("desc", "description", "deskripsi", "nama", "name", "spec", "type", "details"):
            if (
                has_word("code", "number", "no", "num", "id", "sap", "cd", "sku")
                or "sku" in words
                or "material" in words
                or "article" in words
                or "item" in words
                or "product" in words
            ):
                return "material_code"

    # 3. Material / Item Description
    if has_word("desc", "description", "deskripsi", "nama", "name", "spec", "label") and has_word("material", "article", "item", "product", "sku", "barang", "brand", "row"):
        return "material_description"

    # 4. Destination Plant / Store / Site Code
    if not is_source and has_word("store", "site", "plant", "outlet", "location", "destinasi", "destination", "branch", "toko", "pos"):
        if not has_word("desc", "description", "deskripsi", "nama", "name"):
            return "destination_plant_code"

    # 5. Destination Plant / Store / Site Description
    if not is_source and has_word("store", "site", "plant", "outlet", "location", "destinasi", "destination", "branch", "toko") and has_word("desc", "description", "deskripsi", "nama", "name"):
        return "destination_plant_description"

    # 6. Issuing Warehouse Code
    if is_source and has_word("site", "plant", "store", "warehouse", "gudang", "wh", "code", "id"):
        if not has_word("desc", "description", "deskripsi", "nama", "name"):
            return "issuing_warehouse_code"

    return None


def canonical_field_name(value: Any) -> Optional[str]:
    normalized = normalize_text(value)

    if normalized is None:
        return None

    normalized = normalized.casefold()

    for field_name, aliases in FIELD_ALIASES.items():
        if normalized in aliases:
            return field_name

    return None


def find_alias_index(
    headers: Iterable[Any],
    field_name: str,
) -> Optional[int]:
    aliases = {
        alias.casefold()
        for alias in FIELD_ALIASES.get(field_name, ())
    }

    for index, header in enumerate(headers):
        normalized = normalize_text(header)

        if normalized and normalized.casefold() in aliases:
            return index

    for index, header in enumerate(headers):
        if match_semantic_field(header) == field_name:
            return index

    return None


def looks_like_warehouse_code(value: Any) -> bool:
    text = normalize_text(value)

    return bool(
        text and re.fullmatch(r"[A-Za-z]{2,5}\d{1,3}", text)
    )
