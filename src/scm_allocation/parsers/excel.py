"""Parser for allocation worksheets in multi-sheet Excel workbooks."""

from __future__ import annotations

from pathlib import Path
import re
from typing import Any, Optional

import openpyxl

from scm_allocation.models.allocation import AllocationRecord, ExtractionResult, ExtractionStatus
from scm_allocation.normalization import find_alias_index, normalize_quantity, normalize_text
from scm_allocation.parsers.common import infer_allocation_context, make_record
from scm_allocation.parsers.generic import extract_generic_layout


def _category_bands(ws: Any, row_number: int) -> list[Optional[str]]:
    bands: list[Optional[str]] = []
    current: Optional[str] = None
    for column in range(1, ws.max_column + 1):
        value = normalize_text(ws.cell(row_number, column).value)
        if value:
            current = value.casefold()
        bands.append(current)
    return bands


def _headers(ws: Any, row_number: int) -> list[Optional[str]]:
    return [normalize_text(ws.cell(row_number, column).value) for column in range(1, ws.max_column + 1)]


def _find_compound_header(ws: Any) -> Optional[int]:
    for row_number in range(1, min(ws.max_row, 20) + 1):
        headers = _headers(ws, row_number)
        if any(value and value.casefold() == "article code" for value in headers):
            return row_number
    return None


def _find_simple_header(ws: Any) -> Optional[int]:
    for row_number in range(1, min(ws.max_row, 20) + 1):
        headers = _headers(ws, row_number)
        if (
            find_alias_index(headers, "material_code") is not None
            and find_alias_index(headers, "quantity") is not None
        ):
            return row_number
    return None


def _find_site_allocation_header(ws: Any) -> Optional[int]:
    for row_number in range(1, min(ws.max_row, 20) + 1):
        headers = _headers(ws, row_number)
        header_names = {value.casefold() for value in headers if value}
        if {"site code", "site desc", "allocation"}.issubset(header_names):
            return row_number
    return None


def _find_text_index(headers: list[Optional[str]], *terms: str) -> Optional[int]:
    normalized_terms = tuple(term.casefold() for term in terms)
    for index, header in enumerate(headers):
        value = (header or "").casefold()
        if any(term in value for term in normalized_terms):
            return index
    return None


def _compound_columns(ws: Any, header_row: int) -> Optional[dict[str, int]]:
    headers = _headers(ws, header_row)
    bands = _category_bands(ws, max(1, header_row - 1))
    columns: dict[str, Optional[int]] = {
        "material_code": next((i for i, value in enumerate(headers) if value and value.casefold() == "article code"), None),
        "material_description": next((i for i, value in enumerate(headers) if value and value.casefold() == "article description"), None),
        "quantity": next((i for i, value in enumerate(headers) if value and value.casefold() == "alokasi"), None),
    }
    for field_name, band_name, header_name in (
        ("from_code", "from", "site code"),
        ("from_description", "from", "site desc"),
        ("to_code", "to store", "site code"),
        ("to_description", "to store", "site desc"),
    ):
        columns[field_name] = next(
            (i for i, value in enumerate(headers) if value and value.casefold() == header_name and bands[i] == band_name),
            None,
        )
    if any(index is None for index in columns.values()):
        return None
    return {name: index for name, index in columns.items() if index is not None}


def _simple_columns(ws: Any, header_row: int) -> Optional[dict[str, Optional[int]]]:
    headers = _headers(ws, header_row)
    material_index = find_alias_index(headers, "material_code")
    quantity_index = find_alias_index(headers, "quantity")
    if material_index is None or quantity_index is None:
        return None

    description_index = find_alias_index(headers, "material_description")
    destination_code = find_alias_index(headers, "destination_plant_code")
    destination_description = find_alias_index(headers, "destination_plant_description")
    issuing_code = find_alias_index(headers, "issuing_warehouse_code")
    issuing_description = find_alias_index(headers, "issuing_warehouse_description")

    marker_headers = {value.casefold() for value in headers if value}
    has_allocation_marker = bool(
        destination_code is not None
        or {"request", "site code", "site desc", "from", "wh code", "warehouse code", "store code", "store name", "plant code", "code store", "destinasi code", "alokasi", "qty", "quantity"} & marker_headers
        or any(value and (value.casefold().startswith("gudang") or value.casefold().startswith("ur ") or value.casefold().startswith("store") or value.casefold().startswith("outlet") or value.casefold().startswith("location")) for value in headers)
    )
    if not has_allocation_marker:
        return None

    if destination_code is None:
        destination_code = next(
            (index for index, value in enumerate(headers) if value and value.casefold() in {"wh code", "store code", "plant code", "outlet id", "location"}),
            None,
        )
    if destination_description is None:
        destination_description = _find_text_index(headers, "bu desc", "plant desc", "ur ", "store name", "outlet name")
    if issuing_code is None:
        issuing_code = _find_text_index(headers, "from site code", "from", "warehouse code", "plant asal")
    if issuing_description is None:
        issuing_description = _find_text_index(headers, "from site desc", "gudang")

    return {
        "material_code": material_index,
        "material_description": description_index,
        "quantity": quantity_index,
        "destination_code": destination_code,
        "destination_description": destination_description,
        "issuing_code": issuing_code,
        "issuing_description": issuing_description,
    }


def _find_destination_matrix(ws: Any) -> Optional[dict[str, Any]]:
    """Detect a wide allocation matrix with destination codes above columns."""
    for header_row in range(2, min(ws.max_row, 20) + 1):
        headers = _headers(ws, header_row)
        material_index = find_alias_index(headers, "material_code")
        description_index = find_alias_index(headers, "material_description")
        if material_index is None or description_index is None:
            continue
        code_row = header_row - 1
        code_columns = {
            column - 1: normalize_text(ws.cell(code_row, column).value)
            for column in range(1, ws.max_column + 1)
            if normalize_text(ws.cell(code_row, column).value)
            and re.fullmatch(r"[A-Za-z]{1,5}\d{1,4}", normalize_text(ws.cell(code_row, column).value) or "")
        }
        if not code_columns:
            continue
        has_quantity = any(
            normalize_quantity(ws.cell(row, column + 1).value) is not None
            for row in range(header_row + 1, min(ws.max_row, header_row + 20) + 1)
            for column in code_columns
        )
        if has_quantity:
            return {
                "header_row": header_row,
                "material_code": material_index,
                "material_description": description_index,
                "code_columns": code_columns,
            }
    return None


def _find_new_store_header(ws: Any) -> Optional[int]:
    for row_number in range(1, min(ws.max_row, 20) + 1):
        headers = _headers(ws, row_number)
        names = {(header or "").casefold() for header in headers}
        if {"article number (sap code)", "material description", "qty", "plant asal", "plant code"}.issubset(names):
            return row_number
    return None


def _find_site_material_matrix(ws: Any) -> Optional[dict[str, Any]]:
    """Detect a wide matrix: material codes row 1, descriptions row 2, sites below."""
    if ws.max_row < 3 or ws.max_column < 3:
        return None
    headers = _headers(ws, 2)
    if not (headers[0] and headers[0].casefold() == "site code" and headers[1] and headers[1].casefold() == "site desc"):
        return None
    columns = {
        index: normalize_text(ws.cell(1, index + 1).value)
        for index in range(2, ws.max_column)
        if normalize_text(ws.cell(1, index + 1).value)
        and re.fullmatch(r"\d{6,}", normalize_text(ws.cell(1, index + 1).value) or "")
        and headers[index]
    }
    return {"header_row": 2, "material_columns": columns} if columns else None


def _site_material_matrix_records(
    worksheet: Any,
    matrix: dict[str, Any],
    *,
    filepath: Path,
    source_email_id: Optional[str],
) -> list[AllocationRecord]:
    records: list[AllocationRecord] = []
    headers = _headers(worksheet, matrix["header_row"])
    for row_number in range(matrix["header_row"] + 1, worksheet.max_row + 1):
        values = _row_values(worksheet, row_number)
        destination_code = values[0] if values else None
        if _is_total_or_empty(destination_code):
            continue
        for column, material_code in matrix["material_columns"].items():
            quantity = values[column] if column < len(values) else None
            if normalize_quantity(quantity) in {None, 0}:
                continue
            records.append(make_record(
                material_code=material_code,
                material_description=headers[column],
                quantity=quantity,
                destination_plant_code=destination_code,
                destination_plant_description=values[1] if len(values) > 1 else None,
                issuing_warehouse_code=None,
                issuing_warehouse_description=None,
                allocation_context=infer_allocation_context(worksheet.title, headers[column]),
                source_email_id=source_email_id,
                source_filename=filepath.name,
                source_sheet=worksheet.title,
                source_row=row_number,
                source_location=f"{worksheet.title}!{column + 1}{row_number}",
                parser_type="excel_site_material_matrix",
                confidence=0.95,
            ))
    return records


def _new_store_columns(ws: Any, header_row: int) -> Optional[dict[str, int]]:
    """Detect the IT/LOOPS new-store table with source and destination SLoc columns."""
    headers = _headers(ws, header_row)
    def exact(value: str) -> Optional[int]:
        return next((index for index, header in enumerate(headers) if header and header.casefold() == value), None)
    material = exact("article number (sap code)")
    description = exact("material description")
    quantity = exact("qty")
    source = exact("plant asal")
    destination = exact("plant code")
    if None in {material, description, quantity, source, destination}:
        return None
    source_sloc = source + 1 if source + 1 < len(headers) and headers[source + 1] and "storage location" in headers[source + 1].casefold() else None
    destination_sloc = destination + 1 if destination + 1 < len(headers) and headers[destination + 1] and "storage location" in headers[destination + 1].casefold() else None
    if source_sloc is None or destination_sloc is None:
        return None
    return {"material_code": material, "material_description": description, "quantity": quantity, "issuing_code": source, "destination_code": destination, "issuing_sloc": source_sloc, "destination_sloc": destination_sloc}


def _new_store_records(
    worksheet: Any,
    header_row: int,
    columns: dict[str, int],
    *,
    filepath: Path,
    source_email_id: Optional[str],
    destination_names: dict[str, str],
) -> list[AllocationRecord]:
    records: list[AllocationRecord] = []
    for row_number in range(header_row + 1, worksheet.max_row + 1):
        values = _row_values(worksheet, row_number)
        material = values[columns["material_code"]] if columns["material_code"] < len(values) else None
        if _is_total_or_empty(material):
            continue
        destination = values[columns["destination_code"]] if columns["destination_code"] < len(values) else None
        records.append(make_record(
            material_code=material,
            material_description=values[columns["material_description"]],
            quantity=values[columns["quantity"]],
            destination_plant_code=destination,
            destination_plant_description=destination_names.get(normalize_text(destination) or ""),
            issuing_warehouse_code=values[columns["issuing_code"]],
            issuing_warehouse_description=None,
            issuing_warehouse_sloc=values[columns["issuing_sloc"]],
            destination_sloc=values[columns["destination_sloc"]],
            allocation_context=infer_allocation_context(worksheet.title, values[columns["material_description"]]),
            source_email_id=source_email_id,
            source_filename=filepath.name,
            source_sheet=worksheet.title,
            source_row=row_number,
            source_location=f"{worksheet.title}!A{row_number}",
            parser_type="excel_new_store_table",
            confidence=0.95,
        ))
    return records


def _new_store_destination_names(workbook: Any) -> dict[str, str]:
    names: dict[str, str] = {}
    for worksheet in workbook.worksheets:
        if (normalize_text(worksheet.cell(1, 1).value) or "").casefold() != "list new store":
            continue
        for row_number in range(2, worksheet.max_row + 1):
            code = normalize_text(worksheet.cell(row_number, 1).value)
            description = normalize_text(worksheet.cell(row_number, 2).value)
            if code and description:
                names[code] = description
    return names


def _trim_phantom_cells(workbook: Any) -> None:
    """Buang sel TANPA nilai di luar area yang benar-benar berisi.

    Sheet yang pernah diformat sampai bawah (mis. "Alokasi Tecno 12 Oct'26.xlsx",
    12 Okt 2026: max_row 1.048.213, baris berisi terakhir 181) membuat semua loop
    `range(..., ws.max_row + 1)` berjalan sejuta kali; tiap putaran `_row_values`
    membuat sel baru lalu `ws.max_column` menelusuri semua sel → kuadratik, scan
    tidak pernah selesai dan proses menumpuk tiap jam. Dengan sel kosong di ujung
    dibuang, max_row/max_column = ukuran data nyata dan loop lama tetap benar.
    `_cells` = penyimpanan sel openpyxl (bukan API publik, tapi stabil); hanya
    sel bernilai None yang dihapus, jadi tidak ada data yang hilang.
    """
    for ws in workbook.worksheets:
        cells = getattr(ws, "_cells", None)
        if not cells:
            continue
        filled = [key for key, cell in cells.items() if cell.value is not None]
        last_row = max((r for r, _ in filled), default=0)
        last_col = max((c for _, c in filled), default=0)
        for key in [k for k, cell in cells.items()
                    if cell.value is None and (k[0] > last_row or k[1] > last_col)]:
            del cells[key]


_REFERENCE_TITLE_RE = re.compile(r"^B[A-Z][A-Z0-9]\d$|stock|stok|sales|\bdata\b|master|pvt|pivot|soh|inventory", re.IGNORECASE)
_REQUEST_TITLE_RE = re.compile(r"req|\bpo\b|alo[ck]asi|alloc|order|\bsto\b|npi|new\s*store|kebutuhan|permintaan", re.IGNORECASE)
_DEAD_TITLE_RE = re.compile(r"database|\bdb\b|status|\blookup\b|\bref\b|referensi|\blow\b|\bmb52\b", re.IGNORECASE)
_MONTHS = {
    "jan": 1, "januari": 1, "january": 1, "feb": 2, "februari": 2, "february": 2, "mar": 3, "maret": 3, "march": 3,
    "apr": 4, "april": 4, "mei": 5, "may": 5, "jun": 6, "juni": 6, "june": 6, "jul": 7, "juli": 7, "july": 7,
    "agu": 8, "agt": 8, "agustus": 8, "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "okt": 10, "oct": 10, "oktober": 10, "october": 10, "nov": 11, "november": 11, "des": 12, "dec": 12,
    "desember": 12, "december": 12,
}


def _stale_month_title(title: str, received: Optional[str]) -> bool:
    """Judul sheet menyebut bulan yang sudah lewat ("STO MARSHALL APRIL" di email
    September) — sisa alokasi lama yang ikut tersimpan di workbook."""
    match = re.match(r"\s*(\d{4})-(\d{2})", received or "")
    if not match:
        return False
    current = int(match.group(2))
    months = {_MONTHS[w] for w in re.findall(r"[a-z]+", (title or "").casefold()) if w in _MONTHS}
    return bool(months) and not months & {current, (current - 2) % 12 + 1}


_REFERENCE_HEADERS = {"stor. loc.", "storage location", "unrestricted", "material group", "sales stock", "base unit of measure", "mch 1"}


def _is_reference_sheet(title: str, headers: list[Any]) -> bool:
    """Sheet dump stok/penjualan yang ikut dilampirkan, bukan isi permintaan."""
    if _REQUEST_TITLE_RE.search(title or ""):
        return False
    names = {(normalize_text(h) or "").casefold() for h in headers}
    return bool(_REFERENCE_TITLE_RE.search(title or "")) or bool(names & _REFERENCE_HEADERS)


_SHEET_DATE_RE = re.compile(r"(?<!\d)(\d{2})(\d{2})(\d{2}|\d{4})(?!\d)")


def _sheet_date(text: Optional[str]) -> Optional[tuple[int, int, int]]:
    """Tanggal ddmmyy/ddmmyyyy dari judul sheet atau subjek ("211026", "R2 (23012026)")."""
    for match in _SHEET_DATE_RE.finditer(text or ""):
        day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 1 <= day <= 31 and 1 <= month <= 12:
            return (year % 100, month, day)
    return None


def _keep_current_log_sheets(records: list[AllocationRecord], workbook: Any, subject: Optional[str],
                             diagnostics: list[dict[str, object]]) -> list[AllocationRecord]:
    """Workbook "log berjalan" (satu sheet per tanggal alokasi, mis. Lamina: 77 sheet
    sejak Januari) dulu terbaca SELURUHNYA tiap email — alokasi lama ikut lagi.
    Kalau ≥3 sheet bertanggal: ambil sheet bertanggal sama dengan subjek email,
    atau kalau tidak ada yang cocok, sheet bertanggal paling baru."""
    dated = {ws.title: _sheet_date(ws.title) for ws in workbook.worksheets}
    dated = {title: date for title, date in dated.items() if date}
    if len(dated) < 3:
        return records
    with_records = {r.source_sheet for r in records}
    candidates = {title: date for title, date in dated.items() if title in with_records}
    if not candidates:
        return records
    target = _sheet_date(subject)
    if target not in candidates.values():
        target = max(candidates.values())
    keep = {title for title, date in candidates.items() if date == target}
    kept = [r for r in records if r.source_sheet in keep or r.source_sheet not in dated]
    if len(kept) != len(records):
        for d in diagnostics:
            if d["sheet"] in dated and d["sheet"] not in keep and d.get("records"):
                d["status"], d["reason"] = "dropped_old_log", "sheet tanggal lama di workbook log berjalan"
    return kept


def _keep_named_sheets(records: list[AllocationRecord], body_text: Optional[str],
                       diagnostics: list[dict[str, object]]) -> list[AllocationRecord]:
    """Isi email menunjuk sheetnya ("lihat sheet PvT_Summary", "di sheet Final"):
    sheet lain di workbook itu (draf, "Coret", salinan "Batch 1 (2)") adalah
    kertas kerja — dulu ikut terbaca sehingga satu alokasi muncul 3-4 kali."""
    body = (body_text or "").casefold()
    if "sheet" not in body:
        return records
    sheets = {r.source_sheet for r in records if r.source_sheet}
    named = {
        title for title in sheets
        if len(title) >= 3 and re.search(r"sheet\s*[\"'(\[]?\s*" + re.escape(title.casefold()) + r"(?![\w(])", body)
    }
    if not named or named == sheets:
        return records
    for d in diagnostics:
        if d.get("sheet") in sheets - named and d.get("records"):
            d["status"], d["reason"] = "dropped_unnamed", "isi email menunjuk sheet lain"
    return [r for r in records if r.source_sheet in named]


def _freeze_dimensions(workbook: Any) -> None:
    """Bekukan max_row/max_column tiap sheet setelah _trim_phantom_cells.

    Di openpyxl dua properti itu menyisir SEMUA sel tiap kali dibaca, padahal
    parser ini membacanya per baris (`_row_values`, `_headers` di dalam loop).
    Sheet 20 rb baris × 165 rb sel ("PO Multi1 28 Sept.xlsx") jadi miliaran
    langkah — scan menggantung >5 menit dan RAM membengkak. Ukurannya dihitung
    sekali di sini; ws.cell() yang membaca sel kosong memang membuat sel baru,
    tapi tidak lagi menggeser ukuran sheet.
    """
    for ws in workbook.worksheets:
        max_row, max_column = ws.max_row, ws.max_column
        ws.__class__ = type(
            "FrozenWorksheet",
            (ws.__class__,),
            {"max_row": property(lambda self, v=max_row: v), "max_column": property(lambda self, v=max_column: v)},
        )


def _row_values(ws: Any, row_number: int) -> list[Any]:
    return [ws.cell(row_number, column).value for column in range(1, ws.max_column + 1)]


def _is_total_or_empty(value: Any) -> bool:
    text = normalize_text(value)
    if text is None:
        return True
    lowered = text.casefold()
    if lowered in {"grand total", "total", "subtotal"}:
        return True
    from scm_allocation.normalization.fields import canonical_field_name
    if canonical_field_name(lowered) == "material_code":
        return True
    return False


def _site_allocation_records(
    worksheet: Any,
    header_row: int,
    *,
    filepath: Path,
    source_email_id: Optional[str],
) -> list[AllocationRecord]:
    headers = _headers(worksheet, header_row)
    site_code_index = next(index for index, value in enumerate(headers) if value and value.casefold() == "site code")
    site_desc_index = next(index for index, value in enumerate(headers) if value and value.casefold() == "site desc")
    quantity_index = next(index for index, value in enumerate(headers) if value and value.casefold() == "allocation")
    # Alokasi Apple UR: kode barang ada di kolom "Article Code" — dulu tidak
    # dibaca, jadi semua baris tanpa material (20 dari 20 email, Okt 2026).
    material_index = next((i for i, v in enumerate(headers) if v and v.casefold() in {"article code", "article"}), None)
    if material_index is None:
        material_index = find_alias_index(headers, "material_code")
    description_index = next((i for i, v in enumerate(headers) if v and v.casefold() in {"article desc", "article description"}), None)
    if description_index is None:
        description_index = find_alias_index(headers, "material_description")

    def at(values: list[Any], index: Optional[int]) -> Any:
        return values[index] if index is not None and index < len(values) else None

    records: list[AllocationRecord] = []
    for row_number in range(header_row + 1, worksheet.max_row + 1):
        values = _row_values(worksheet, row_number)
        quantity = normalize_quantity(values[quantity_index] if quantity_index < len(values) else None)
        if quantity is None or quantity == 0:
            continue
        material = at(values, material_index)
        records.append(
            make_record(
                material_code=material,
                material_description=at(values, description_index),
                quantity=quantity,
                destination_plant_code=values[site_code_index] if site_code_index < len(values) else None,
                destination_plant_description=values[site_desc_index] if site_desc_index < len(values) else None,
                issuing_warehouse_code=None,
                issuing_warehouse_description=None,
                allocation_context=infer_allocation_context(worksheet.title),
                source_email_id=source_email_id,
                source_filename=filepath.name,
                source_sheet=worksheet.title,
                source_row=row_number,
                source_location=f"{worksheet.title}!A{row_number}",
                parser_type="excel_site_allocation",
                confidence=0.9 if normalize_text(material) else 0.7,
                extra_errors=() if normalize_text(material) else ("source has no material code or material description",),
            )
        )
    return records


def _matrix_records(
    worksheet: Any,
    matrix: dict[str, Any],
    *,
    filepath: Path,
    source_email_id: Optional[str],
    destination_plant_code: Optional[str],
    destination_plant_description: Optional[str],
) -> list[AllocationRecord]:
    header_row = matrix["header_row"]
    material_index = matrix["material_code"]
    description_index = matrix["material_description"]
    code_columns: dict[int, str] = matrix["code_columns"]
    destination_codes = set(code_columns.values())
    if destination_plant_code and destination_plant_code in destination_codes:
        selected_destinations = {destination_plant_code}
        source_candidates = [code for code in code_columns.values() if code != destination_plant_code]
        issuing_warehouse_code = source_candidates[0] if len(source_candidates) == 1 else None
    else:
        selected_destinations = destination_codes
        issuing_warehouse_code = None

    records: list[AllocationRecord] = []
    for row_number in range(header_row + 1, worksheet.max_row + 1):
        values = _row_values(worksheet, row_number)
        material_value = values[material_index] if material_index < len(values) else None
        if _is_total_or_empty(material_value):
            continue
        description = values[description_index] if description_index < len(values) else None
        for column, code in code_columns.items():
            if code not in selected_destinations:
                continue
            quantity = normalize_quantity(values[column] if column < len(values) else None)
            if quantity is None or quantity == 0:
                continue
            headers = _headers(worksheet, header_row)
            record = make_record(
                material_code=material_value,
                material_description=description,
                quantity=quantity,
                destination_plant_code=code,
                destination_plant_description=headers[column] or destination_plant_description,
                issuing_warehouse_code=issuing_warehouse_code,
                issuing_warehouse_description=None,
                allocation_context=infer_allocation_context(worksheet.title, description),
                source_email_id=source_email_id,
                source_filename=filepath.name,
                source_sheet=worksheet.title,
                source_row=row_number,
                source_location=f"{worksheet.title}!{column + 1}{row_number}",
                parser_type="excel_destination_matrix",
                confidence=0.9,
            )
            records.append(record)
    return records


def _workbook_from_rows(sheets: list[tuple[str, list[list[Any]]]]) -> Any:
    """Bangun Workbook openpyxl dari baris mentah, supaya semua detektor sheet di
    bawah (yang memakai API openpyxl) berlaku juga untuk .xlsb/.xls/.csv."""
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets:
        ws = workbook.create_sheet(re.sub(r"[\\/*?:\[\]]", " ", str(title))[:31] or "Sheet")
        for r, row in enumerate(rows, start=1):
            for c, value in enumerate(row, start=1):
                if value not in (None, ""):
                    ws.cell(r, c, value)
    return workbook


def load_any_workbook(filepath: Path) -> Any:
    """Buka lampiran Excel apa pun formatnya: .xlsx/.xlsm/.xltx (openpyxl),
    .xlsb (pyxlsb), .xls (xlrd), .csv. Ekstensi yang menipu (file .xls yang
    isinya xlsx, dsb.) dicoba ulang dengan pembaca lain."""
    suffix = filepath.suffix.casefold()

    def read_xlsb() -> Any:
        from pyxlsb import open_workbook
        sheets = []
        with open_workbook(str(filepath)) as wb:
            for name in wb.sheets:
                with wb.get_sheet(name) as sheet:
                    sheets.append((name, [[cell.v for cell in row] for row in sheet.rows(sparse=True)]))
        return _workbook_from_rows(sheets)

    def read_xls() -> Any:
        import xlrd
        book = xlrd.open_workbook(str(filepath))
        sheets = []
        for sheet in book.sheets():
            rows = []
            for r in range(sheet.nrows):
                row = []
                for c in range(sheet.ncols):
                    cell = sheet.cell(r, c)
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            value = xlrd.xldate.xldate_as_datetime(value, book.datemode)
                        except Exception:
                            pass
                    row.append(value)
                rows.append(row)
            sheets.append((sheet.name, rows))
        return _workbook_from_rows(sheets)

    def read_csv() -> Any:
        import csv
        raw = filepath.read_bytes()
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        return _workbook_from_rows([(filepath.stem, list(csv.reader(text.splitlines(), dialect)))])

    def read_openpyxl() -> Any:
        # read_only = streaming, tanpa memuat pivot cache/format: workbook dengan
        # pivot besar ("REQUEST STOCK DEVICE … X049.xlsx", 7 MB) dulu butuh ~2 menit
        # hanya untuk dibuka. Nilainya disalin ke Workbook biasa supaya detektor
        # di bawah tetap bisa memakai ws.cell(); baris kosong beruntun di ujung
        # sheet yang pernah diformat sampai bawah dihentikan lebih awal.
        source = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        try:
            sheets, hidden = [], set()
            for ws in source.worksheets:
                rows, blank_run = [], 0
                for row in ws.iter_rows(values_only=True):
                    if any(value is not None for value in row):
                        rows.extend([[]] * blank_run)
                        rows.append(list(row))
                        blank_run = 0
                    else:
                        blank_run += 1
                        if blank_run > 5000:
                            break
                sheets.append((ws.title, rows))
                if getattr(ws, "sheet_state", "visible") != "visible":
                    hidden.add(ws.title)
        finally:
            source.close()
        workbook = _workbook_from_rows(sheets)
        for ws in workbook.worksheets:
            if ws.title in hidden:
                ws.sheet_state = "hidden"
        return workbook

    order = {
        ".xlsb": (read_xlsb, read_openpyxl),
        ".xls": (read_xls, read_openpyxl),
        ".csv": (read_csv,),
    }.get(suffix, (read_openpyxl, read_xlsb, read_xls))
    last_error: Optional[Exception] = None
    for reader in order:
        try:
            return reader()
        except Exception as err:  # format salah / file rusak — coba pembaca berikutnya
            last_error = err
    raise last_error or ValueError(f"cannot open {filepath.name}")


def parse_excel(
    filepath: Path,
    *,
    source_email_id: Optional[str] = None,
    destination_plant_code: Optional[str] = None,
    destination_plant_description: Optional[str] = None,
    subject: Optional[str] = None,
    body_text: Optional[str] = None,
    received: Optional[str] = None,
) -> ExtractionResult:
    """Inspect every worksheet and extract records from all supported allocation sheets."""
    parser_type = "excel_compound_grid"
    try:
        workbook = load_any_workbook(filepath)
    except Exception as err:
        return ExtractionResult([], ExtractionStatus.INVALID, parser_type, source_email_id, (str(err),))
    _trim_phantom_cells(workbook)
    _freeze_dimensions(workbook)

    records: list[AllocationRecord] = []
    seen_records: set[tuple[object, ...]] = set()
    parser_types: set[str] = set()
    diagnostics: list[dict[str, object]] = []
    matched_sheet_count = 0
    destination_names = _new_store_destination_names(workbook)

    for worksheet in workbook.worksheets:
        compound_header = _find_compound_header(worksheet)
        compound_columns = _compound_columns(worksheet, compound_header) if compound_header else None
        matrix = None if compound_columns else _find_destination_matrix(worksheet)
        site_material_matrix = None if compound_columns or matrix else _find_site_material_matrix(worksheet)
        site_allocation_header = None if compound_columns or matrix or site_material_matrix else _find_site_allocation_header(worksheet)
        new_store_header = None if compound_columns or matrix or site_material_matrix or site_allocation_header else _find_new_store_header(worksheet)
        simple_header = None if new_store_header is not None or compound_columns or matrix or site_material_matrix or site_allocation_header else _find_simple_header(worksheet)
        simple_columns = _simple_columns(worksheet, simple_header) if simple_header is not None else None
        new_store_columns = _new_store_columns(worksheet, new_store_header) if new_store_header is not None else None

        if compound_columns and compound_header is not None:
            format_name = "excel_compound_grid"
            header_row = compound_header
            columns: dict[str, Optional[int]] = compound_columns
        elif matrix:
            format_name = "excel_destination_matrix"
            header_row = matrix["header_row"]
            matrix_sheet_records = _matrix_records(
                worksheet,
                matrix,
                filepath=filepath,
                source_email_id=source_email_id,
                destination_plant_code=destination_plant_code,
                destination_plant_description=destination_plant_description,
            )
            matched_sheet_count += 1
            parser_types.add(format_name)
            sheet_records = 0
            duplicate_records = 0
            for record in matrix_sheet_records:
                identity = (
                    record.material_code,
                    record.material_description,
                    record.quantity,
                    record.destination_plant_code,
                    record.issuing_warehouse_code,
                )
                if identity in seen_records:
                    duplicate_records += 1
                    continue
                seen_records.add(identity)
                records.append(record)
                sheet_records += 1
            headers = _headers(worksheet, header_row)
            diagnostics.append({
                "workbook": filepath.name,
                "sheet": worksheet.title,
                "status": "matched",
                "format": format_name,
                "header_row": header_row,
                "dimensions": worksheet.dimensions,
                "headers": headers[:16],
                "sample_rows": [_row_values(worksheet, row)[:16] for row in range(max(1, header_row - 1), min(header_row + 2, worksheet.max_row) + 1)],
                "records": sheet_records,
                "duplicates_skipped": duplicate_records,
            })
            continue
        elif site_material_matrix:
            format_name = "excel_site_material_matrix"
            header_row = site_material_matrix["header_row"]
            source_records = _site_material_matrix_records(
                worksheet, site_material_matrix, filepath=filepath, source_email_id=source_email_id
            )
        elif new_store_columns and new_store_header is not None:
            format_name = "excel_new_store_table"
            header_row = new_store_header
            source_records = _new_store_records(
                worksheet, header_row, new_store_columns, filepath=filepath,
                source_email_id=source_email_id, destination_names=destination_names,
            )
        elif site_allocation_header is not None:
            format_name = "excel_site_allocation"
            header_row = site_allocation_header
            matched_sheet_count += 1
            parser_types.add(format_name)
            site_records = _site_allocation_records(
                worksheet,
                header_row,
                filepath=filepath,
                source_email_id=source_email_id,
            )
            sheet_records = 0
            duplicate_records = 0
            for record in site_records:
                identity = (
                    record.material_code,
                    record.material_description,
                    record.quantity,
                    record.destination_plant_code,
                    record.issuing_warehouse_code,
                )
                if identity in seen_records:
                    duplicate_records += 1
                    continue
                seen_records.add(identity)
                records.append(record)
                sheet_records += 1
            diagnostics.append({
                "workbook": filepath.name,
                "sheet": worksheet.title,
                "status": "matched" if sheet_records else "matched_empty",
                "format": format_name,
                "header_row": header_row,
                "dimensions": worksheet.dimensions,
                "headers": _headers(worksheet, header_row)[:16],
                "sample_rows": [_row_values(worksheet, row)[:16] for row in range(header_row, min(header_row + 3, worksheet.max_row) + 1)],
                "records": sheet_records,
                "duplicates_skipped": duplicate_records,
                "reason": "material identity missing from source" if sheet_records else "no nonblank allocation quantities",
            })
            continue
        elif simple_columns and simple_header is not None:
            format_name = "simple_excel_table"
            header_row = simple_header
            columns = simple_columns
        else:
            diagnostics.append({
                "workbook": filepath.name,
                "sheet": worksheet.title,
                "status": "skipped",
                "reason": "no supported allocation structure",
                "dimensions": worksheet.dimensions,
                    "headers": _headers(worksheet, 1)[:16] if worksheet.max_row else [],
                    "sample_rows": [_row_values(worksheet, row)[:16] for row in range(1, min(4, worksheet.max_row) + 1)],
            })
            continue

        if site_material_matrix or new_store_columns:
            matched_sheet_count += 1
            parser_types.add(format_name)
            sheet_records = 0
            duplicate_records = 0
            for record in source_records:
                identity = (record.material_code, record.material_description, record.quantity, record.destination_plant_code, record.issuing_warehouse_code)
                if identity in seen_records:
                    duplicate_records += 1
                    continue
                seen_records.add(identity)
                records.append(record)
                sheet_records += 1
            diagnostics.append({
                "workbook": filepath.name, "sheet": worksheet.title,
                "status": "matched" if sheet_records else "matched_empty", "format": format_name,
                "header_row": header_row, "dimensions": worksheet.dimensions,
                "headers": _headers(worksheet, header_row)[:16],
                "sample_rows": [_row_values(worksheet, row)[:16] for row in range(max(1, header_row - 1), min(header_row + 2, worksheet.max_row) + 1)],
                "records": sheet_records, "duplicates_skipped": duplicate_records,
                "reason": "no nonblank allocation quantities" if not sheet_records else None,
            })
            continue

        matched_sheet_count += 1
        parser_types.add(format_name)
        sheet_records = 0
        duplicate_records = 0
        for row_number in range(header_row + 1, worksheet.max_row + 1):
            values = _row_values(worksheet, row_number)
            material_index = columns["material_code"]
            material_value = values[material_index] if isinstance(material_index, int) and material_index < len(values) else None
            if _is_total_or_empty(material_value):
                continue

            description_index = columns["material_description"]
            quantity_index = columns["quantity"]
            quantity_header = _headers(worksheet, header_row)[quantity_index]
            if (
                quantity_header
                and quantity_header.casefold() in {"request", "alokasi"}
                and normalize_quantity(values[quantity_index]) is None
            ):
                continue
            destination_index = columns.get("to_code", columns.get("destination_code"))
            destination_description_index = columns.get("to_description", columns.get("destination_description"))
            issuing_index = columns.get("from_code", columns.get("issuing_code"))
            issuing_description_index = columns.get("from_description", columns.get("issuing_description"))
            row_destination = values[destination_index] if isinstance(destination_index, int) and destination_index < len(values) else destination_plant_code
            row_destination_description = (
                values[destination_description_index]
                if isinstance(destination_description_index, int) and destination_description_index < len(values)
                else destination_plant_description
            )
            row_issuing = values[issuing_index] if isinstance(issuing_index, int) and issuing_index < len(values) else None
            row_issuing_description = (
                values[issuing_description_index]
                if isinstance(issuing_description_index, int) and issuing_description_index < len(values)
                else None
            )
            description = values[description_index] if isinstance(description_index, int) and description_index < len(values) else None
            quantity = values[quantity_index] if isinstance(quantity_index, int) and quantity_index < len(values) else None
            record = make_record(
                material_code=material_value,
                material_description=description,
                quantity=quantity,
                destination_plant_code=row_destination,
                destination_plant_description=row_destination_description,
                issuing_warehouse_code=row_issuing,
                issuing_warehouse_description=row_issuing_description,
                allocation_context=infer_allocation_context(
                    values[0] if format_name == "excel_compound_grid" and values else worksheet.title,
                    description,
                ),
                source_email_id=source_email_id,
                source_filename=filepath.name,
                source_sheet=worksheet.title,
                source_row=row_number,
                source_location=f"{worksheet.title}!A{row_number}",
                parser_type=format_name,
                    confidence=1.0 if format_name == "excel_compound_grid" else 0.9,
            )
            identity = (
                record.material_code,
                record.material_description,
                record.quantity,
                record.destination_plant_code,
                record.issuing_warehouse_code,
            )
            if identity in seen_records:
                duplicate_records += 1
                continue
            seen_records.add(identity)
            records.append(record)
            sheet_records += 1

        diagnostics.append({
            "workbook": filepath.name,
            "sheet": worksheet.title,
            "status": "matched" if sheet_records else "matched_empty",
            "format": format_name,
            "header_row": header_row,
            "dimensions": worksheet.dimensions,
                "headers": _headers(worksheet, header_row)[:16],
                "sample_rows": [_row_values(worksheet, row)[:16] for row in range(header_row, min(header_row + 3, worksheet.max_row) + 1)],
            "records": sheet_records,
            "duplicates_skipped": duplicate_records,
            "reason": "no nonblank request quantities" if not sheet_records else None,
        })

    # Sheet yang tidak menghasilkan satu rekaman pun dari parser khusus di atas
    # dicoba parser generik (header bebas / kode toko sebagai judul kolom).
    productive = {d["sheet"] for d in diagnostics if d.get("records")}
    generic_strength: dict[str, int] = {}
    # Sheet berjudul referensi (MASTER, Stock, …) dibaca TERAKHIR: isinya sering
    # salinan baris sheet permintaan, dan kalau dibaca duluan baris sheet
    # permintaan terbuang sebagai duplikat lalu sheet referensinya ikut dibuang
    # (PO STOCK Samarinda 19 Okt: yang tersisa malah dump stok 2023).
    for worksheet in sorted(workbook.worksheets, key=lambda ws: _is_reference_sheet(ws.title, [])):
        if worksheet.title in productive or worksheet.sheet_state != "visible":
            continue
        rows = [list(row) for row in worksheet.iter_rows(max_row=min(worksheet.max_row, 20000), values_only=True)]
        generic_records, layout = extract_generic_layout(
            rows,
            source_email_id=source_email_id,
            source_filename=filepath.name,
            source_sheet=worksheet.title,
            destination_plant_code=destination_plant_code,
            destination_plant_description=destination_plant_description,
            location_prefix=f"{worksheet.title}!",
        )
        generic_strength[worksheet.title] = layout["qty_score"] if layout else 0
        added = 0
        for record in generic_records:
            identity = (record.material_code, record.material_description, record.quantity,
                        record.destination_plant_code, record.issuing_warehouse_code)
            if identity in seen_records:
                continue
            seen_records.add(identity)
            records.append(record)
            added += 1
        if added:
            matched_sheet_count += 1
            parser_types.add(generic_records[0].parser_type)
            diagnostics.append({"workbook": filepath.name, "sheet": worksheet.title, "status": "matched",
                                "format": generic_records[0].parser_type, "records": added})

    # Sheet generik berkolom qty lemah (pivot "Sum of Quantity", kolom gudang)
    # dibuang kalau ada sheet generik lain dengan kolom permintaan bernama tegas
    # (REQ/PO/Alokasi) — itu sheet permintaannya, sisanya data pendukung.
    generic_sheets = {r.source_sheet for r in records if r.parser_type.startswith("generic_")}
    strong = {t for t in generic_sheets if generic_strength.get(t, 0) >= 9}
    weak = {t for t in generic_sheets if generic_strength.get(t, 0) < 9}
    if strong and weak:
        records = [r for r in records if r.source_sheet not in weak]
        for d in diagnostics:
            if d["sheet"] in weak and str(d.get("format", "")).startswith("generic_"):
                d["status"], d["reason"] = "dropped_weak", "kolom qty lemah, ada sheet permintaan lain"
        parser_types = {r.parser_type for r in records}

    # Workbook permintaan sering menyertakan sheet referensi: dump stok SAP
    # ("Sales Stock", sheet bernama kode gudang BGC1/BFC1, kolom Stor. Loc./
    # Material Group). Parser tabel sederhana membacanya sebagai alokasi ribuan
    # baris (28 Okt 2026: 11.373 baris dari satu email permintaan 1 toko). Kalau
    # ada sheet lain yang menghasilkan rekaman, sheet referensi itu dibuang.
    reference = {d["sheet"] for d in diagnostics if d.get("records") and _is_reference_sheet(d["sheet"], d.get("headers") or [])}
    if reference and any(r.source_sheet not in reference for r in records):
        records = [r for r in records if r.source_sheet not in reference]
        for d in diagnostics:
            if d["sheet"] in reference and d.get("records"):
                d["status"], d["reason"] = "dropped_reference", "sheet referensi stok, bukan permintaan"
        parser_types = {r.parser_type for r in records}

    # Sheet tersembunyi yang jelas data mati (PO STOCK Samarinda: "DATABASE",
    # dump stok 2023, 11 rb baris terbaca sebagai alokasi) dibuang kalau ada
    # sheet terlihat yang menghasilkan rekaman. TIDAK semua sheet tersembunyi:
    # alokasi per toko sering justru di sheet yang disembunyikan PM (Samsung GO:
    # "GAP"; iBox Pick & Win: Sheet1 berisi STO ke XE18; Lamina: alokasi ke
    # RDC) — dicek ke STO SAP. Yang dibuang: judul data mati / bulan lama.
    visible_count = sum(1 for r in records if r.source_sheet not in
                        {ws.title for ws in workbook.worksheets if ws.sheet_state != "visible"})
    hidden = {
        ws.title for ws in workbook.worksheets
        if ws.sheet_state != "visible" and (
            _REFERENCE_TITLE_RE.search(ws.title) or _DEAD_TITLE_RE.search(ws.title)
            or _stale_month_title(ws.title, received)
        )
    }
    if hidden and visible_count:
        records = [r for r in records if r.source_sheet not in hidden]
        for d in diagnostics:
            if d["sheet"] in hidden and d.get("records"):
                d["status"], d["reason"] = "dropped_hidden", "sheet tersembunyi"
        parser_types = {r.parser_type for r in records}

    records = _keep_current_log_sheets(records, workbook, subject, diagnostics)
    records = _keep_named_sheets(records, body_text, diagnostics)
    parser_types = {r.parser_type for r in records} or parser_types

    metadata = {"worksheet_diagnostics": diagnostics}
    if not records:
        errors = (
            ("allocation worksheet(s) found but no records extracted",)
            if matched_sheet_count
            else ("no allocation worksheet found",)
        )
        return ExtractionResult(
            [],
            ExtractionStatus.INVALID,
            "+".join(sorted(parser_types)) or parser_type,
            source_email_id,
            errors,
            metadata=metadata,
        )

    status = ExtractionStatus.PARTIAL if any(record.extraction_status != ExtractionStatus.VALID for record in records) else ExtractionStatus.VALID
    result_parser_type = next(iter(parser_types)) if len(parser_types) == 1 else "excel_multi_sheet"
    return ExtractionResult(records, status, result_parser_type, source_email_id, metadata=metadata)
