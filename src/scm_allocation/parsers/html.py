"""Parsers for flat and pivot allocation tables embedded in HTML email bodies."""

from __future__ import annotations

from html.parser import HTMLParser
from typing import Optional

from scm_allocation.models.allocation import AllocationRecord, ExtractionResult, ExtractionStatus
from scm_allocation.normalization import find_alias_index, looks_like_warehouse_code, normalize_text
from scm_allocation.parsers.common import infer_allocation_context, make_record


class _TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._table: Optional[list[list[str]]] = None
        self._row: Optional[list[str]] = None
        self._cell: Optional[list[str]] = None

    def handle_starttag(self, tag: str, _attrs: list[tuple[str, Optional[str]]]) -> None:
        del _attrs
        tag = tag.casefold()
        if tag == "table":
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in {"td", "th"} and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.casefold()
        if tag in {"td", "th"} and self._row is not None and self._cell is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None
        elif tag == "tr" and self._table is not None and self._row:
            self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table:
            self.tables.append(self._table)
            self._table = None


def parse_html_tables(html: str) -> list[list[list[str]]]:
    parser = _TableParser()
    parser.feed(html or "")
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
    if request_subheader_index >= len(subheaders):
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
        if find_alias_index(table[0], "destination_plant_code") is not None and find_alias_index(table[0], "quantity") is not None:
            table_records = _parse_flat_table(table, source_email_id=source_email_id, table_number=table_number)
            if table_records:
                parser_types.add("html_flat_table")
                records.extend(table_records)
        elif len(table) > 1:
            table_records = _parse_single_request_table(
                table,
                source_email_id=source_email_id,
                table_number=table_number,
                destination_plant_code=destination_plant_code,
                destination_plant_description=destination_plant_description,
            )
            if table_records:
                parser_types.add("html_request_matrix")
                records.extend(table_records)
            else:
                table_records = _parse_pivot_table(table, source_email_id=source_email_id, table_number=table_number)
                if table_records:
                    parser_types.add("html_pivot_table")
                    records.extend(table_records)

    if not records:
        return ExtractionResult([], ExtractionStatus.INVALID, "html_table", source_email_id, ("no supported allocation table found",))
    status = ExtractionStatus.PARTIAL if any(record.extraction_status != ExtractionStatus.VALID for record in records) else ExtractionStatus.VALID
    parser_type = "+".join(sorted(parser_types))
    return ExtractionResult(records, status, parser_type, source_email_id, metadata={"table_count": len(parser_tables)})