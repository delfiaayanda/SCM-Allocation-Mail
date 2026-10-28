"""Permintaan alokasi yang ditulis sebagai teks di isi email (bukan tabel/lampiran).

Audit 28 Okt 2026: 88 email "tak terbaca" berisi kode material di teks, mis.

    Material : 8100261708
    Item Desc : MBA 15 MDN/10C GPU/16GB/1TB M5
    Qty : 1unit

    8100225999 JBL PartyBox Club 120 BLACK
    Qty: 1 unit
    E807 Erafone ruko negara

Parser terakhir — hanya dipakai kalau lampiran dan tabel HTML tidak menghasilkan
apa pun. Qty WAJIB tertulis eksplisit ("Qty 2", "2 unit"); tanpa itu baris tidak
dibuat (tidak menebak). Semua rekamannya ditandai untuk dicek (kategori Review).
"""

from __future__ import annotations

import re
from typing import Optional

from scm_allocation.allocation_subject import MATERIAL_CODE_RE, own_text
from scm_allocation.models.allocation import AllocationRecord
from scm_allocation.parsers.common import infer_allocation_context, make_record
from scm_allocation.parsers.generic import PLANT_RE, WAREHOUSE_RE

_QTY_RE = re.compile(
    r"(?:\bqty|\bquantity|\bjumlah|\bjml|\bqt)\s*[:=.]?\s*(\d{1,4})(?!\d)"
    r"|(?<![\d/.,])(\d{1,4})\s*(?:unit|units|pcs|pc|ea|buah)\b",
    re.IGNORECASE,
)
# Akhir isi permintaan: salam penutup / tanda tangan / disclaimer.
_SIGN_OFF_RE = re.compile(
    r"^\s*(?:thanks|thank you|terima ?kasih|terimakasih|thx|regards|best regards|br|salam|hormat kami"
    r"|important warning|this message and any attachments)\b",
    re.IGNORECASE | re.MULTILINE,
)
_DESC_LABEL_RE = re.compile(r"^\s*(?:item\s*desc(?:ription)?|desc(?:ription)?|deskripsi|nama\s*barang|type|tipe)\s*[:=]\s*(.+)$",
                            re.IGNORECASE | re.MULTILINE)


def _codes(text: str) -> tuple[list[str], list[str]]:
    tokens = {t for t in re.split(r"[^A-Za-z0-9]+", text) if t and PLANT_RE.fullmatch(t.upper()) and not t.isdigit()}
    upper = sorted({t.upper() for t in tokens if any(ch.isupper() for ch in t) or t.isupper()})
    stores = [t for t in upper if not WAREHOUSE_RE.fullmatch(t)]
    warehouses = [t for t in upper if WAREHOUSE_RE.fullmatch(t)]
    return stores, warehouses


def parse_text_request(
    body: object,
    subject: object,
    *,
    source_email_id: Optional[str],
    destination_plant_code: Optional[str] = None,
) -> list[AllocationRecord]:
    text = own_text(body)
    sign_off = _SIGN_OFF_RE.search(text)
    if sign_off:
        text = text[: sign_off.start()]
    matches = list(MATERIAL_CODE_RE.finditer(text))
    if not matches:
        return []

    stores, warehouses = _codes(text)
    subject_stores, subject_warehouses = _codes(str(subject or ""))
    email_store = (stores[0] if len(stores) == 1 else None) or (subject_stores[0] if len(subject_stores) == 1 else None) \
        or destination_plant_code
    email_source = warehouses[0] if len(warehouses) == 1 else (subject_warehouses[0] if len(subject_warehouses) == 1 else None)

    records: list[AllocationRecord] = []
    seen: set[str] = set()
    for number, match in enumerate(matches):
        material = match.group(0)
        if material in seen:
            continue
        end = matches[number + 1].start() if number + 1 < len(matches) else len(text)
        segment = text[match.end():end]
        qty_match = _QTY_RE.search(segment)
        if not qty_match:
            continue
        quantity = qty_match.group(1) or qty_match.group(2)
        line = segment.split("\n", 1)[0].strip(" :-\t|")
        label = _DESC_LABEL_RE.search(segment)
        description = (label.group(1).strip() if label else line) or None
        seg_stores, seg_warehouses = _codes(segment)
        destination = (seg_stores[0] if len(seg_stores) == 1 else None) or email_store
        source = (seg_warehouses[0] if len(seg_warehouses) == 1 else None) or email_source
        seen.add(material)
        records.append(make_record(
            material_code=material,
            material_description=description[:120] if description else None,
            quantity=quantity,
            destination_plant_code=destination,
            destination_plant_description=None,
            issuing_warehouse_code=source,
            issuing_warehouse_description=None,
            allocation_context=infer_allocation_context(str(subject or ""), description),
            source_email_id=source_email_id,
            source_filename=None,
            source_sheet=None,
            source_row=None,
            source_location=f"body text #{number + 1}",
            parser_type="text_body",
            confidence=0.6,
            extra_errors=("dibaca dari teks isi email — cek ulang qty & toko",),
        ))
    return records
