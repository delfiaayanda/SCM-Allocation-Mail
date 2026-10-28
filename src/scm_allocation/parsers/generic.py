"""Parser cadangan untuk tabel alokasi berformat bebas (sheet Excel atau tabel HTML).

Parser khusus (compound grid, matrix, new store, dst.) tetap dicoba lebih dulu;
modul ini baru dipakai kalau semuanya tidak menemukan apa-apa. Audit 28 Okt 2026:
ratusan email alokasi gagal hanya karena nama header tidak ada di daftar alias
("Mat Code", "REQ", "Po Tambahan", "Alokasi Sleeve STI"), atau kode toko langsung
jadi judul kolom (X219, S234). Di sini kolom dikenali dari NAMA header sekaligus
dari ISI datanya: kode material 8–12 digit berawalan 81/22/…, kode plant
huruf+angka 4 karakter (X249, BGC1, BC34), dan angka qty.

Tiga bentuk yang dikenali:
  • tabel panjang  — satu baris = satu material (+ toko/qty di kolom)
  • matriks toko   — kode toko sebagai judul kolom, qty di selnya
  • matriks material — kode material sebagai judul kolom, toko per baris
"""

from __future__ import annotations

import re
from decimal import Decimal
from typing import Any, Optional, Sequence

from scm_allocation.models.allocation import AllocationRecord
from scm_allocation.normalization import normalize_quantity, normalize_text
from scm_allocation.parsers.common import infer_allocation_context, make_record


PLANT_RE = re.compile(r"[A-Z][A-Z0-9]{2}\d")
# Gudang (MDC/RDC/DC): BGC1, BIC1, BC34, BHC1 — selain itu toko.
WAREHOUSE_RE = re.compile(r"B[A-Z][A-Z0-9]\d")
MATERIAL_DATA_RE = re.compile(r"(?:81|82|83|84|21|22|23)\d{6,10}")
MATERIAL_ANY_RE = re.compile(r"\d{6,12}|[A-Z0-9][A-Z0-9-]{4,24}")

_HEADER_ROWS = 40
_SAMPLE_ROWS = 2000
_MAX_QTY = Decimal(1_000_000)


def _text(value: Any) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return normalize_text(value) or ""


def _words(header: str) -> str:
    return " " + re.sub(r"[^a-z0-9]+", " ", header.casefold()).strip() + " "


def _has(header: str, *terms: str) -> bool:
    w = _words(header)
    return any(f" {term} " in w for term in terms)


def _plant(value: Any) -> Optional[str]:
    text = _text(value).upper()
    return text if PLANT_RE.fullmatch(text) else None


def _material_data(value: Any) -> bool:
    return bool(MATERIAL_DATA_RE.fullmatch(_text(value)))


def _qty(value: Any) -> Optional[Decimal]:
    if isinstance(value, bool):
        return None
    q = normalize_quantity(_text(value))
    if q is None or q <= 0 or q >= _MAX_QTY or q != q.to_integral_value():
        return None
    return q


def _is_total(row: Sequence[Any]) -> bool:
    for value in row:
        text = _text(value).casefold()
        if text:
            return bool(re.match(r"(grand\s*)?(sub\s*)?total\b|jumlah\b|sum\b", text))
    return True


# ── Tipe kolom dari header ────────────────────────────────────────────────────

# Tidak pernah kolom qty alokasi, apa pun kata lainnya.
_QTY_HARD_EXCLUDE = (
    "date", "tgl", "tanggal", "week", "month", "bulan", "percent", "persen", "rate", "code", "kode",
    "no", "nomor", "number", "telp", "phone", "hp", "id", "sloc", "loc", "stor", "price", "harga", "srp",
    "store", "toko", "name", "nama", "desc", "reason", "ket", "keterangan",
)
# Kolom pendukung (stok, penjualan, DOS) — kalah dari kata permintaan/alokasi.
_QTY_SOFT_EXCLUDE = (
    "stock", "stok", "sales", "sale", "dos", "avg", "average", "est", "min", "max", "unrestricted",
    "atp", "sr", "pusat", "hub", "sell", "out", "in", "mtd", "ytd", "total", "sum", "type",
)


def _qty_header_score(header: str) -> int:
    if not header:
        return 0
    # "BGE1 SLOC 1001": judul berawal kode gudang = qty yang dialokasikan dari
    # gudang/SLoc itu (kata "sloc" di sini bukan kolom nomor SLoc).
    first = header.split()[0].upper()
    if WAREHOUSE_RE.fullmatch(first) and not _has(header, "stock", "stok", "sales", "unrestricted", "atp"):
        return 8
    if _has(header, *_QTY_HARD_EXCLUDE):
        return 0
    if _has(header, "alokasi", "allocation", "alocation", "alloc", "allocated") or header.casefold().startswith(("alokasi", "allocation")):
        return 10
    if _has(header, "req", "request", "reques", "requset", "requested", "permintaan", "order", "po", "kebutuhan", "fulfill", "fullfill",
            "fulfillment", "tambahan", "additional", "suggest", "proposal", "refill", "replenish", "replenishment"):
        return 9
    # Pivot hasil PM ("Sum of Quantity") — sah, tapi kalah dari kolom permintaan
    # bernama tegas di sheet lain (lihat excel._drop_weak_generic_sheets).
    if _has(header, "sum") and _has(header, "qty", "quantity", "alokasi", "jumlah"):
        return 7
    if _has(header, *_QTY_SOFT_EXCLUDE):
        return 0
    if _has(header, "qty", "quantity", "jumlah", "pcs", "unit", "units", "sto", "kirim", "transfer", "pengiriman", "plan", "tdr"):
        return 8
    # Kolom berjudul gudang asal ("DC CIKUPA", "MDC", "Gudang BGC1") berisi qty yang
    # dialokasikan dari gudang itu (di sampingnya SALES/STOCK sebagai data pendukung).
    if _has(header, "dc", "mdc", "rdc", "gudang", "warehouse", "wh", "depo") or WAREHOUSE_RE.fullmatch(header.strip().upper()):
        return 8
    return 0


def _is_material_header(header: str) -> bool:
    if _has(header, "desc", "description", "deskripsi", "nama", "name", "descr", "putus") and not _has(header, "code", "kode"):
        return False
    return _has(header, "material", "article", "articel", "artikel", "sku", "mat", "item", "kode barang", "sap", "barang") and not _has(header, "group", "grp", "type", "brand", "category", "cat")


def _is_desc_header(header: str) -> bool:
    if _has(header, "site", "store", "toko", "plant", "destinasi", "destination", "bu", "outlet", "region", "customer", "custemer"):
        return False
    return _has(header, "desc", "description", "deskripsi", "descr", "nama barang", "item name", "product", "produk", "type 2", "tipe 2", "model", "variant", "article name", "material name") or (
        _has(header, "nama", "name") and not _has(header, "store", "toko", "site", "plant", "customer", "custemer", "pelanggan", "1"))


def _source_header(header: str) -> bool:
    return _has(header, "from", "asal", "dari", "supply", "supplying", "source", "issuing", "dc", "gudang", "wh", "warehouse", "pengirim")


def _dest_header(header: str) -> bool:
    return _has(header, "to", "tujuan", "dest", "destinasi", "destination", "store", "toko", "site", "outlet", "plant", "era", "bu", "penerima", "row labels", "code store", "store code", "sap code")


# ── Pemilihan baris header & kolom ────────────────────────────────────────────

_STATS_CACHE: dict[tuple[int, int, int], list[dict[str, float]]] = {}


def _column_stats(rows: Sequence[Sequence[Any]], start: int, width: int) -> list[dict[str, float]]:
    # Tiga detektor menilai baris header yang sama; statistik kolom (sampai 2.000
    # baris — kolom REQ bisa baru terisi jauh di bawah) cukup dihitung sekali.
    key = (id(rows), start, width)
    if key not in _STATS_CACHE:
        _STATS_CACHE[key] = _compute_column_stats(rows, start, width)
    return _STATS_CACHE[key]


def _compute_column_stats(rows: Sequence[Sequence[Any]], start: int, width: int) -> list[dict[str, float]]:
    sample = [row for row in rows[start:start + _SAMPLE_ROWS] if any(_text(v) for v in row)]
    stats = []
    for c in range(width):
        values = [row[c] if c < len(row) else None for row in sample]
        filled = [v for v in values if _text(v)]
        n = len(filled) or 1
        stats.append({
            "filled": len(filled),
            "material": sum(_material_data(v) for v in filled) / n,
            "plant": sum(bool(_plant(v)) for v in filled) / n,
            "warehouse": sum(bool(WAREHOUSE_RE.fullmatch(_text(v).upper())) for v in filled) / n,
            "qty": sum(_qty(v) is not None for v in filled) / n,
            "numeric": sum(normalize_quantity(_text(v)) is not None for v in filled) / n,
        })
    return stats


def _long_layout(rows: Sequence[Sequence[Any]], h: int, width: int) -> Optional[dict[str, Any]]:
    headers = [_text(rows[h][c]) if c < len(rows[h]) else "" for c in range(width)]
    stats = _column_stats(rows, h + 1, width)
    if not any(s["filled"] for s in stats):
        return None

    material = next((c for c in range(width) if _is_material_header(headers[c]) and stats[c]["material"] >= 0.5), None)
    if material is None:
        material = next((c for c in range(width) if stats[c]["material"] >= 0.6 and stats[c]["filled"] >= 1), None)
    description = next((c for c in range(width) if c != material and _is_desc_header(headers[c]) and stats[c]["numeric"] < 0.5), None)
    if material is None and description is None:
        return None

    # Kolom berjudul tegas boleh bercampur catatan teks ("push alokasi" di kolom
    # DC CIKUPA, 28 Okt 2026: 38% angka) — baris teksnya dilewati saat ekstraksi.
    qty_candidates = [
        (score, c) for c in range(width)
        if c not in (material, description) and stats[c]["filled"] and stats[c]["material"] < 0.3
        for score in (_qty_header_score(headers[c]),)
        if stats[c]["numeric"] >= (0.3 if score >= 8 else 0.6) and stats[c]["qty"] > 0
    ]
    qty_candidates = [(score, c) for score, c in qty_candidates if score]
    header_row = h
    if not qty_candidates and any(_has(x, "alokasi", "allocation", "alocation") and _has(x, "to", "ke", "tujuan") for x in headers):
        # "Stock | Alocation To | Name": ada kolom tujuan alokasi, jadi satu-satunya
        # kolom angka kecil (walau berjudul Stock) adalah qty yang dialokasikan.
        numeric = [c for c in range(width) if c not in (material, description) and stats[c]["qty"] >= 0.8
                   and stats[c]["material"] < 0.3 and stats[c]["plant"] < 0.3]
        if len(numeric) == 1:
            qty_candidates = [(8, numeric[0])]
    if not qty_candidates and material is not None and not _material_data(headers[material]):
        # Tabel ringkas "Article Code | Article Descr | Total" (atau Ket / nama toko /
        # nama depo sebagai judul): tepat SATU kolom angka yang bukan data pendukung
        # (stok, penjualan, DOS). Laporan stok berkolom angka banyak tetap ditolak.
        support = ("stock", "stok", "sales", "sale", "dos", "avg", "average", "est", "min", "max", "unrestricted",
                   "atp", "sr", "mtd", "ytd", "sell", "penjualan", "price", "harga", "srp", "inventory", "soh")
        numeric = [c for c in range(width) if c not in (material, description) and stats[c]["filled"]
                   and stats[c]["numeric"] >= 0.6 and stats[c]["material"] < 0.3 and stats[c]["plant"] < 0.3]
        if len(numeric) == 1 and not _has(headers[numeric[0]], *support) and stats[numeric[0]]["qty"] >= 0.5:
            qty_candidates = [(6, numeric[0])]
    if not qty_candidates:
        # Tabel tanpa header (mis. "8100304352 | SKZ S610 OpenRun | 3"): baris h
        # sendiri sudah data, dan qty = satu-satunya kolom angka kecil.
        if material is None or not _material_data(headers[material]):
            return None
        # Ada judul kolom di baris atasnya (mis. pivot stok "Sum of QTY" yang
        # sengaja ditolak) — berarti tabelnya ber-header, bukan tanpa header.
        # Begitu juga kalau kolom material sudah terisi di baris mana pun di atasnya
        # (baris ini cuma kelanjutan tabel lain, mis. sesudah baris "Total" pivot).
        if any(_text(row[material] if material < len(row) else None) for row in rows[:h]):
            return None
        numeric = [c for c in range(width) if c != material and stats[c]["qty"] >= 0.8 and _qty(headers[c]) is not None]
        if len(numeric) != 1:
            return None
        quantity = numeric[0]
        header_row = h - 1
        description = description if description is not None else next(
            (c for c in range(width) if c not in (material, quantity) and stats[c]["numeric"] < 0.3 and stats[c]["filled"]), None)
        best = 0
    else:
        best = max(score for score, _ in qty_candidates)
        quantity = min(c for score, c in qty_candidates if score == best)

    plant_columns = [c for c in range(width) if c not in (material, quantity) and stats[c]["plant"] >= 0.6]
    is_wh = {c: stats[c]["warehouse"] >= 0.6 for c in plant_columns}
    # 1) Judul kolom yang tegas: "From/Asal/Supply/DC" = asal, "Dest/Tujuan/To" = tujuan.
    source = next((c for c in plant_columns if _source_header(headers[c]) and not _has(headers[c], "dest", "destinasi", "tujuan", "to")), None)
    destination = next((c for c in plant_columns if c != source and _has(headers[c], "dest", "destinasi", "destination", "tujuan", "to", "penerima")), None)
    # 2) Sisanya ("Plant", "Site", "Store Code"): kode gudang = asal, kode toko = tujuan.
    rest = [c for c in plant_columns if c not in (source, destination)]
    if destination is None:
        destination = next((c for c in rest if not is_wh[c]), None)
        if destination is None and rest and (source is not None or len(rest) == 1):
            destination = next(c for c in rest if c != source)  # mis. STO ke RDC: tujuannya gudang
    if source is None:
        source = next((c for c in rest if c != destination and is_wh[c]), None)
    # Format blok per toko ("Material | … | E595 | BGC1 | PO", header berulang
    # tiap toko): kode toko ada di JUDUL kolom stok, qty di kolom PO/REQ.
    header_store = header_source = None
    if destination is None:
        codes = {c: _plant(headers[c]) for c in range(width) if c != quantity and _plant(headers[c])}
        stores = [code for code in codes.values() if not WAREHOUSE_RE.fullmatch(code)]
        if len(stores) == 1:
            header_store = stores[0]
            warehouses = [code for code in codes.values() if WAREHOUSE_RE.fullmatch(code)]
            header_source = warehouses[0] if len(warehouses) == 1 else None
    dest_desc = next((c for c in range(width) if c not in (material, description, quantity, source, destination)
                      and _has(headers[c], "site desc", "store name", "nama toko", "nama store", "name 1", "plant desc", "store desc", "destinasi desc", "bu description", "store")
                      and stats[c]["numeric"] < 0.5), None)
    return {"kind": "long", "header": header_row, "qty_score": best or 8, "material": material, "description": description, "quantity": quantity,
            "source": source, "destination": destination, "dest_desc": dest_desc,
            "header_store": header_store, "header_source": header_source,
            # Kolom permintaan bernama tegas (REQ/PO/Alokasi) mengalahkan matriks
            # toko yang kebetulan punya kode toko di judul kolom stok.
            "score": (2 if material is not None else 1) + (1 if destination is not None or header_store else 0)
                     + (1.5 if best >= 9 else 0)}


def _store_matrix_layout(rows: Sequence[Sequence[Any]], h: int, width: int) -> Optional[dict[str, Any]]:
    headers = [_text(rows[h][c]) if c < len(rows[h]) else "" for c in range(width)]
    stats = _column_stats(rows, h + 1, width)
    stores = {c: _plant(headers[c]) for c in range(width)
              if _plant(headers[c]) and stats[c]["numeric"] >= 0.5 and stats[c]["plant"] < 0.3 and stats[c]["material"] < 0.3}
    # Satu kode plant di baris data (bukan judul kolom) juga cocok pola ini —
    # baris header matriks sungguhan tidak berisi kode material.
    if any(_material_data(headers[c]) for c in range(width)):
        return None
    if not stores:
        return None
    material = next((c for c in range(width) if c not in stores and (_is_material_header(headers[c]) or stats[c]["material"] >= 0.6) and stats[c]["material"] >= 0.5), None)
    description = next((c for c in range(width) if c not in stores and c != material and (_is_desc_header(headers[c]) or (stats[c]["numeric"] < 0.3 and stats[c]["filled"] and stats[c]["plant"] < 0.3))), None)
    if material is None and description is None:
        return None
    warehouses = {c: code for c, code in stores.items() if WAREHOUSE_RE.fullmatch(code)}
    # Kolom gudang di samping kolom toko = stok asal, bukan tujuan.
    if warehouses and len(warehouses) < len(stores):
        for c in warehouses:
            stores.pop(c)
    source = next(iter(warehouses.values())) if len(warehouses) == 1 else None
    return {"kind": "store_matrix", "header": h, "qty_score": 8, "material": material, "description": description,
            "stores": stores, "source_code": source, "score": 3 + len(stores) / 1000}


def _material_matrix_layout(rows: Sequence[Sequence[Any]], h: int, width: int) -> Optional[dict[str, Any]]:
    headers = [_text(rows[h][c]) if c < len(rows[h]) else "" for c in range(width)]
    above = [_text(rows[h - 1][c]) if h > 0 and c < len(rows[h - 1]) else "" for c in range(width)]
    stats = _column_stats(rows, h + 1, width)
    materials = {}
    for c in range(width):
        # Isi kolom harus qty kecil — bukan kode material/plant lagi (itu tanda
        # baris "header" ini sebenarnya baris data pertama tabel panjang).
        # (Nol tetap angka: kolom varian yang banyak 0-nya tetap kolom qty.)
        if stats[c]["numeric"] < 0.6 or stats[c]["qty"] == 0 or stats[c]["material"] >= 0.3 or stats[c]["plant"] >= 0.3:
            continue
        for code in (headers[c], above[c]):
            if _material_data(code):
                materials[c] = code
                break
    if not materials:
        return None
    destination = next((c for c in range(width) if c not in materials and stats[c]["plant"] >= 0.6), None)
    if destination is None:
        return None
    descriptions = {c: (headers[c] if not _material_data(headers[c]) else above[c] if not _material_data(above[c]) else None) for c in materials}
    dest_desc = next((c for c in range(width) if c not in materials and c != destination
                      and _has(headers[c], "store name", "site desc", "nama toko", "store", "name") and stats[c]["numeric"] < 0.5), None)
    return {"kind": "material_matrix", "header": h, "qty_score": 8, "materials": materials, "descriptions": descriptions,
            "destination": destination, "dest_desc": dest_desc, "score": 3 + len(materials) / 1000}


def detect_layout(rows: Sequence[Sequence[Any]]) -> Optional[dict[str, Any]]:
    width = max((len(row) for row in rows[:_HEADER_ROWS + _SAMPLE_ROWS]), default=0)
    if width < 2 or len(rows) < 2:
        return None
    best = None
    _STATS_CACHE.clear()
    for h in range(min(len(rows) - 1, _HEADER_ROWS)):
        if not any(_text(v) for v in rows[h]):
            continue
        # Semua kandidat baris header dinilai (tanpa berhenti di temuan pertama):
        # matriks yang kode materialnya di baris atas & deskripsinya di baris bawah
        # baru utuh kalau header = baris deskripsi.
        for finder in (_material_matrix_layout, _store_matrix_layout, _long_layout):
            layout = finder(rows, h, width)
            if layout and (best is None or layout["score"] > best["score"]):
                best = layout
    _STATS_CACHE.clear()
    return best


# ── Baris → rekaman ───────────────────────────────────────────────────────────

def extract_generic(
    rows: Sequence[Sequence[Any]],
    *,
    source_email_id: Optional[str],
    source_filename: Optional[str],
    source_sheet: Optional[str],
    destination_plant_code: Optional[str],
    destination_plant_description: Optional[str],
    location_prefix: str,
) -> list[AllocationRecord]:
    return extract_generic_layout(
        rows, source_email_id=source_email_id, source_filename=source_filename, source_sheet=source_sheet,
        destination_plant_code=destination_plant_code, destination_plant_description=destination_plant_description,
        location_prefix=location_prefix,
    )[0]


def extract_generic_layout(
    rows: Sequence[Sequence[Any]],
    *,
    source_email_id: Optional[str],
    source_filename: Optional[str],
    source_sheet: Optional[str],
    destination_plant_code: Optional[str],
    destination_plant_description: Optional[str],
    location_prefix: str,
) -> tuple[list[AllocationRecord], Optional[dict[str, Any]]]:
    """Seperti extract_generic, plus layout terpilih (qty_score dipakai excel.py
    untuk memilih sheet permintaan di atas sheet pivot/pendukung)."""
    layout = detect_layout(rows)
    if layout is None:
        return [], None
    # Kode toko dari subjek ("… (Xe14)") ditulis sembarang huruf besar/kecil.
    destination_plant_code = _plant(destination_plant_code) or destination_plant_code
    records: list[AllocationRecord] = []

    def cell(row: Sequence[Any], c: Optional[int]) -> Any:
        return row[c] if c is not None and c < len(row) else None

    def add(row_number: int, column: int, *, material: Any, description: Any, quantity: Decimal,
            destination: Any, destination_description: Any, source: Any) -> None:
        records.append(make_record(
            material_code=material if _text(material) and MATERIAL_ANY_RE.fullmatch(_text(material)) else None,
            material_description=description,
            quantity=quantity,
            destination_plant_code=destination or destination_plant_code,
            destination_plant_description=destination_description or (None if destination else destination_plant_description),
            issuing_warehouse_code=source,
            issuing_warehouse_description=None,
            allocation_context=infer_allocation_context(source_sheet, description),
            source_email_id=source_email_id,
            source_filename=source_filename,
            source_sheet=source_sheet,
            source_row=row_number,
            source_location=f"{location_prefix}r{row_number}c{column + 1}",
            parser_type=f"generic_{layout['kind']}",
            confidence=0.8,
        ))

    last_dest = None
    block_store, block_source = layout.get("header_store"), layout.get("header_source")
    header_material = _text(cell(rows[layout["header"]], layout.get("material"))) if layout["header"] >= 0 else ""
    for i in range(layout["header"] + 1, len(rows)):
        row = rows[i]
        if _is_total(row):
            continue
        if layout["kind"] == "long" and block_store and header_material and _text(cell(row, layout["material"])) == header_material:
            # Header blok berikutnya: ambil kode toko & gudang blok ini.
            codes = [_plant(v) for v in row if _plant(v)]
            stores = [code for code in codes if not WAREHOUSE_RE.fullmatch(code)]
            if stores:
                block_store = stores[0]
                warehouses = [code for code in codes if WAREHOUSE_RE.fullmatch(code)]
                block_source = warehouses[0] if warehouses else None
            continue
        if layout["kind"] == "long":
            material = cell(row, layout["material"])
            description = cell(row, layout["description"])
            if not _text(material) and not _text(description):
                continue
            quantity = _qty(cell(row, layout["quantity"]))
            if quantity is None:
                continue
            destination = _plant(cell(row, layout["destination"])) if layout["destination"] is not None else None
            # Sel toko yang digabung (rowspan/merge) hanya terisi di baris pertama.
            if layout["destination"] is not None:
                destination = destination or last_dest
                last_dest = destination
            source = _plant(cell(row, layout["source"])) if layout["source"] is not None else None
            add(i + 1, layout["quantity"], material=material, description=description, quantity=quantity,
                destination=destination or block_store, destination_description=cell(row, layout["dest_desc"]),
                source=source or (block_source if block_store else None))
        elif layout["kind"] == "store_matrix":
            material = cell(row, layout["material"])
            description = cell(row, layout["description"])
            if not _text(material) and not _text(description):
                continue
            for c, store in layout["stores"].items():
                quantity = _qty(cell(row, c))
                if quantity is not None:
                    add(i + 1, c, material=material, description=description, quantity=quantity,
                        destination=store, destination_description=None, source=layout["source_code"])
        else:
            destination = _plant(cell(row, layout["destination"]))
            if not destination:
                continue
            for c, material in layout["materials"].items():
                quantity = _qty(cell(row, c))
                if quantity is not None:
                    add(i + 1, c, material=material, description=layout["descriptions"].get(c), quantity=quantity,
                        destination=destination, destination_description=cell(row, layout["dest_desc"]), source=None)
    return records, layout
