"""Parsers for flat and pivot allocation tables embedded in HTML email bodies."""

from __future__ import annotations

from html.parser import HTMLParser
from typing import Optional

from scm_allocation.models.allocation import AllocationRecord, ExtractionResult, ExtractionStatus
from scm_allocation.normalization import find_alias_index, looks_like_warehouse_code, normalize_text
from scm_allocation.parsers.common import infer_allocation_context, make_record
from scm_allocation.parsers.generic import MATERIAL_DATA_RE, extract_generic


class _TableParser(HTMLParser):
    """Kumpulkan semua <table>, termasuk yang bersarang.

    Email Outlook membungkus isi dengan tabel layout; dulu tabel dalam menimpa
    status tabel luar, jadi baris sesudahnya hilang. Sekarang tiap tabel punya
    tumpukan sendiri.

    `expand_spans=True` (dipakai parser generik): colspan direntang dan sel
    rowspan diulang di baris bawahnya, supaya kolom data tetap sejajar header
    (email "Permintaan Alokasi … X249": kolom toko di-merge, kolom sesudahnya
    bergeser). Parser lama tetap menerima baris apa adanya (False).
    """

    def __init__(self, expand_spans: bool = False) -> None:
        super().__init__(convert_charrefs=True)
        self.expand_spans = expand_spans
        self.tables: list[list[list[str]]] = []
        # Tiap entri: {"rows", "row": [(teks, colspan, rowspan)] | None,
        #              "cell": [potongan teks] | None, "span": (colspan, rowspan),
        #              "pending": {kolom: [sisa baris, teks]}}
        self._stack: list[dict] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        tag = tag.casefold()
        if tag == "table":
            self._stack.append({"rows": [], "row": None, "cell": None, "span": (1, 1), "pending": {}})
        elif not self._stack:
            return
        elif tag == "tr":
            self._close_row()
            self._stack[-1]["row"] = []
        elif tag in {"td", "th"}:
            top = self._stack[-1]
            if top["row"] is None:
                top["row"] = []
            self._close_cell()
            values = dict(attrs)

            def span(name: str) -> int:
                try:
                    return max(1, min(int(values.get(name) or 1), 50))
                except ValueError:
                    return 1
            top["cell"], top["span"] = [], (span("colspan"), span("rowspan"))
        elif tag == "br" and self._stack[-1]["cell"] is not None:
            self._stack[-1]["cell"].append(" ")

    def handle_data(self, data: str) -> None:
        if self._stack and self._stack[-1]["cell"] is not None:
            self._stack[-1]["cell"].append(data)

    def _close_cell(self) -> None:
        top = self._stack[-1]
        if top["cell"] is not None and top["row"] is not None:
            top["row"].append((" ".join("".join(top["cell"]).split()), *top["span"]))
        top["cell"] = None

    def _layout(self, top: dict, cells: list[tuple[str, int, int]]) -> list[str]:
        if not self.expand_spans:
            return [text for text, _, _ in cells]
        pending, out, col = top["pending"], [], 0

        def take_pending() -> None:
            nonlocal col
            while col in pending:
                remaining, text = pending[col]
                out.append(text)
                if remaining <= 1:
                    del pending[col]
                else:
                    pending[col][0] = remaining - 1
                col += 1
        for text, colspan, rowspan in cells:
            take_pending()
            for k in range(colspan):
                out.append(text if k == 0 else "")
                if rowspan > 1:
                    pending[col] = [rowspan - 1, text if k == 0 else ""]
                col += 1
        while pending and col <= max(pending):
            if col in pending:
                take_pending()
            else:
                out.append("")
                col += 1
        return out

    def _close_row(self) -> None:
        self._close_cell()
        top = self._stack[-1]
        if top["row"]:
            row = self._layout(top, top["row"])
            if any(value for value in row):
                top["rows"].append(row)
        top["row"] = None

    def handle_endtag(self, tag: str) -> None:
        tag = tag.casefold()
        if not self._stack:
            return
        if tag in {"td", "th"}:
            self._close_cell()
        elif tag == "tr":
            self._close_row()
        elif tag == "table":
            self._close_row()
            rows = self._stack.pop()["rows"]
            if rows:
                self.tables.append(rows)

    def close(self) -> None:
        super().close()
        while self._stack:
            self.handle_endtag("table")


def parse_html_tables(html: str, *, expand_spans: bool = False) -> list[list[list[str]]]:
    parser = _TableParser(expand_spans)
    parser.feed(html or "")
    parser.close()
    return parser.tables


def _is_total_row(row: list[str]) -> bool:
    first = normalize_text(row[0]) if row else None
    return first is None or first.casefold() in {"grand total", "total"}


def _parse_flat_table(
    table: list[list[str]],
    *,
    source_email_id: Optional[str],
    table_number: int,
) -> list[AllocationRecord]:
    headers = table[0]
    indices = {
        "material_code": find_alias_index(headers, "material_code"),
        "material_description": find_alias_index(headers, "material_description"),
        "destination_plant_code": find_alias_index(headers, "destination_plant_code"),
        "destination_plant_description": find_alias_index(headers, "destination_plant_description"),
        "quantity": find_alias_index(headers, "quantity"),
    }
    if any(index is None for index in indices.values()):
        return []

    warehouse_indices = [index for index, value in enumerate(headers) if looks_like_warehouse_code(value)]
    warehouse_index = warehouse_indices[0] if len(warehouse_indices) == 1 else None
    records: list[AllocationRecord] = []
    for row_number, row in enumerate(table[1:], start=2):
        if _is_total_row(row):
            continue
        values = {name: row[index] if index < len(row) else None for name, index in indices.items()}
        records.append(
            make_record(
                **values,
                issuing_warehouse_code=headers[warehouse_index] if warehouse_index is not None else None,
                issuing_warehouse_description=None,
                allocation_context=infer_allocation_context(row[4] if len(row) > 4 else None, values["material_description"]),
                source_email_id=source_email_id,
                source_filename=None,
                source_sheet=None,
                source_row=row_number,
                source_location=f"table[{table_number}] row[{row_number}]",
                parser_type="html_flat_table",
                confidence=0.95 if warehouse_index is not None else 0.9,
                extra_errors=("issuing warehouse is ambiguous",) if warehouse_index is None else (),
            )
        )
    return records


def _parse_single_request_table(
    table: list[list[str]],
    *,
    source_email_id: Optional[str],
    table_number: int,
    destination_plant_code: Optional[str],
    destination_plant_description: Optional[str],
) -> list[AllocationRecord]:
    """Parse a one-row request matrix with explicit warehouse columns."""
    headers = table[0]
    material_index = find_alias_index(headers, "material_code")
    description_index = find_alias_index(headers, "material_description")
    quantity_index = find_alias_index(headers, "quantity")
    warehouse_indices = [index for index, value in enumerate(headers) if looks_like_warehouse_code(value)]
    if material_index is None or description_index is None or quantity_index is None or len(warehouse_indices) != 1:
        return []

    warehouse_index = warehouse_indices[0]
    records: list[AllocationRecord] = []
    for row_number, row in enumerate(table[1:], start=2):
        if _is_total_row(row):
            continue
        records.append(
            make_record(
                material_code=row[material_index] if material_index < len(row) else None,
                material_description=row[description_index] if description_index < len(row) else None,
                quantity=row[quantity_index] if quantity_index < len(row) else None,
                destination_plant_code=destination_plant_code,
                destination_plant_description=destination_plant_description,
                issuing_warehouse_code=headers[warehouse_index],
                issuing_warehouse_description=None,
                allocation_context=infer_allocation_context(row[description_index] if description_index < len(row) else None),
                source_email_id=source_email_id,
                source_filename=None,
                source_sheet=None,
                source_row=row_number,
                source_location=f"table[{table_number}] row[{row_number}]",
                parser_type="html_request_matrix",
                confidence=0.95 if destination_plant_code else 0.75,
                extra_errors=() if destination_plant_code else ("missing destination plant code",),
            )
        )
    return records


def _parse_pivot_table(
    table: list[list[str]],
    *,
    source_email_id: Optional[str],
    table_number: int,
) -> list[AllocationRecord]:
    if len(table) < 2:
        return []
    section_headers = table[0]
    subheaders = table[1]
    request_section = next((index for index, value in enumerate(section_headers) if (normalize_text(value) or "").casefold() == "request"), None)
    material_index = find_alias_index(section_headers, "material_code")
    description_index = find_alias_index(section_headers, "material_description")
    if request_section is None or material_index is None or description_index is None:
        return []

    request_subheader_index = request_section
    # Tabel cuma dua baris (judul + subjudul, tanpa data) dulu meledak di table[2]
    # ("list index out of range") dan seluruh email jatuh INVALID.
    if request_subheader_index >= len(subheaders) or len(table) < 3:
        return []
    request_data_index = len(table[2]) - len(subheaders) + request_subheader_index
    destination_plant_code = subheaders[request_subheader_index]
    records: list[AllocationRecord] = []
    for row_number, row in enumerate(table[2:], start=3):
        if _is_total_row(row) or request_data_index >= len(row):
            continue
        records.append(
            make_record(
                material_code=row[material_index] if material_index < len(row) else None,
                material_description=row[description_index] if description_index < len(row) else None,
                quantity=row[request_data_index],
                destination_plant_code=destination_plant_code,
                destination_plant_description=None,
                issuing_warehouse_code=None,
                issuing_warehouse_description=None,
                allocation_context=infer_allocation_context(row[description_index] if description_index < len(row) else None),
                source_email_id=source_email_id,
                source_filename=None,
                source_sheet=None,
                source_row=row_number,
                source_location=f"table[{table_number}] row[{row_number}]",
                parser_type="html_pivot_table",
                confidence=0.9,
                extra_errors=("issuing warehouse is not explicit; source has multiple stock columns",),
            )
        )
    return records


def parse_html(
    html: str,
    *,
    source_email_id: Optional[str] = None,
    destination_plant_code: Optional[str] = None,
    destination_plant_description: Optional[str] = None,
) -> ExtractionResult:
    """Parse every supported allocation table in an HTML email body."""
    parser_tables = parse_html_tables(html)
    records: list[AllocationRecord] = []
    parser_types: set[str] = set()
    for table_number, table in enumerate(parser_tables, start=1):
        if not table:
            continue
        table_records: list[AllocationRecord] = []
        parser_type_used: Optional[str] = None
        if find_alias_index(table[0], "destination_plant_code") is not None and find_alias_index(table[0], "quantity") is not None:
            table_records = _parse_flat_table(table, source_email_id=source_email_id, table_number=table_number)
            parser_type_used = "html_flat_table"
        # A generic "plant"/"store" header can also match a SOURCE-side column
        # (e.g. a redundant "Plant" column alongside an explicit warehouse-code
        # header) — when that false-positive heuristic leaves _parse_flat_table
        # with no rows (it additionally requires a plant DESCRIPTION column),
        # fall through to the matrix/pivot shapes instead of discarding the
        # table outright.
        if not table_records and len(table) > 1:
            table_records = _parse_single_request_table(
                table,
                source_email_id=source_email_id,
                table_number=table_number,
                destination_plant_code=destination_plant_code,
                destination_plant_description=destination_plant_description,
            )
            parser_type_used = "html_request_matrix"
        if not table_records and len(table) > 1:
            table_records = _parse_pivot_table(table, source_email_id=source_email_id, table_number=table_number)
            parser_type_used = "html_pivot_table"
        if table_records:
            parser_types.add(parser_type_used)
            records.extend(table_records)

    if not records:
        # Parser generik: header bebas, kode toko sebagai judul kolom, sel merge.
        # Hanya kalau parser tabel lama tidak menemukan apa pun di email ini.
        seen: set[tuple[object, ...]] = set()
        grids = parse_html_tables(html, expand_spans=True)
        # Outlook kadang memecah satu tabel jadi dua: baris judul di tabel sendiri,
        # baris data di tabel berikutnya (lebar sama). Pasangan itu dicoba digabung
        # lebih dulu; tabel yang sudah terbaca lewat gabungan tidak dibaca ulang.
        candidates: list[tuple[str, list[list[str]], tuple[int, ...]]] = []
        for i in range(len(grids) - 1):
            head, body = grids[i], grids[i + 1]
            if (head and body and len({len(r) for r in head + body}) == 1
                    and not any(MATERIAL_DATA_RE.fullmatch(c.strip()) for r in head for c in r)
                    and any(MATERIAL_DATA_RE.fullmatch(c.strip()) for r in body for c in r)):
                candidates.append((f"table[{i + 1}+{i + 2}] ", head + body, (i, i + 1)))
        candidates += [(f"table[{i + 1}] ", grid, (i,)) for i, grid in enumerate(grids)]
        used: set[int] = set()
        for prefix, grid, members in candidates:
            if used & set(members):
                continue
            found = extract_generic(
                grid,
                source_email_id=source_email_id,
                source_filename=None,
                source_sheet=None,
                destination_plant_code=destination_plant_code,
                destination_plant_description=destination_plant_description,
                location_prefix=prefix,
            )
            if found:
                used |= set(members)
            for record in found:
                identity = (record.material_code, record.material_description, record.quantity,
                            record.destination_plant_code, record.issuing_warehouse_code)
                if identity in seen:
                    continue
                seen.add(identity)
                records.append(record)
                parser_types.add(record.parser_type)

    if not records:
        return ExtractionResult([], ExtractionStatus.INVALID, "html_table", source_email_id, ("no supported allocation table found",))
    status = ExtractionStatus.PARTIAL if any(record.extraction_status != ExtractionStatus.VALID for record in records) else ExtractionStatus.VALID
    parser_type = "+".join(sorted(parser_types))
    return ExtractionResult(records, status, parser_type, source_email_id, metadata={"table_count": len(parser_tables)})