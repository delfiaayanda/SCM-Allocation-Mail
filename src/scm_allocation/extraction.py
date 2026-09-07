"""Format detection and orchestration for the extraction MVP."""

from __future__ import annotations

from pathlib import Path
import re
from typing import Mapping, Optional

from scm_allocation.models.allocation import ExtractionResult, ExtractionStatus
from scm_allocation.normalization import normalize_text
from scm_allocation.parsers import parse_excel, parse_html


OUT_OF_SCOPE_BATAM_SUBJECT = "re: sto device alokasi batam asia brand wk36 2026"


def is_out_of_scope_email(subject: object) -> bool:
    normalized = normalize_text(subject)
    if not normalized:
        return False
    normalized = normalized.casefold()
    return normalized == OUT_OF_SCOPE_BATAM_SUBJECT or (
        "batam" in normalized and ("alokasi" in normalized or "ppbj" in normalized or "sto" in normalized)
    )


def _extract_email_plant(email: Mapping[str, object]) -> tuple[Optional[str], Optional[str]]:
    """Read explicit plant labels from the email body for tables without plant columns."""
    body = str(email.get("body_plain_text") or "")
    for label in ("store desc", "plant", "store code"):
        match = re.search(rf"\b{re.escape(label)}\s*:\s*([^\r\n]+)", body, flags=re.IGNORECASE)
        if not match:
            continue
        value = normalize_text(match.group(1))
        if value and re.fullmatch(r"[A-Za-z]{1,4}\d{2,4}", value):
            return value, None
    subject = normalize_text(email.get("subject")) or ""
    match = re.search(r"\b([A-Za-z]{1,4}\d{2,4})\b", subject)
    if match:
        return match.group(1), None
    return None, None


def extract_email(
    email: Mapping[str, object],
    *,
    attachment_paths: Optional[Mapping[str, Path]] = None,
) -> ExtractionResult:
    """Extract supported allocation formats from one inspected email record."""
    source_email_id = normalize_text(email.get("entry_id") or email.get("source_email_id"))
    subject = email.get("subject")
    if is_out_of_scope_email(subject):
        return ExtractionResult(
            [],
            ExtractionStatus.SKIPPED,
            "out_of_scope",
            source_email_id,
            ("Batam PPBJ email is outside the allocation extraction scope",),
        )

    attachments = email.get("attachments") or []
    excel_results: list[ExtractionResult] = []
    for attachment in attachments:
        if not isinstance(attachment, Mapping):
            continue
        filename = normalize_text(attachment.get("filename"))
        if not filename or not filename.casefold().endswith(".xlsx"):
            continue
        filepath = attachment_paths.get(filename) if attachment_paths else None
        if filepath is None or not filepath.exists():
            excel_results.append(
                ExtractionResult(
                    [],
                    ExtractionStatus.INVALID,
                    "excel_compound_grid",
                    source_email_id,
                    (f"Excel attachment is unavailable locally: {filename}",),
                )
            )
            continue
        plant_code, plant_description = _extract_email_plant(email)
        excel_results.append(
            parse_excel(
                filepath,
                source_email_id=source_email_id,
                destination_plant_code=plant_code,
                destination_plant_description=plant_description,
            )
        )

    if excel_results:
        records = []
        seen_records: set[tuple[object, ...]] = set()
        errors: list[str] = []
        diagnostics: list[object] = []
        parser_types: set[str] = set()
        for result in excel_results:
            parser_types.add(result.parser_type)
            errors.extend(result.errors)
            attachment_diagnostics = result.metadata.get("worksheet_diagnostics", [])
            diagnostics.extend(attachment_diagnostics)
            for record in result.records:
                identity = (
                    record.material_code,
                    record.material_description,
                    record.quantity,
                    record.destination_plant_code,
                    record.issuing_warehouse_code,
                )
                if identity not in seen_records:
                    seen_records.add(identity)
                    records.append(record)
        if records:
            status = ExtractionStatus.PARTIAL if any(record.extraction_status != ExtractionStatus.VALID for record in records) else ExtractionStatus.VALID
            parser_type = next(iter(parser_types)) if len(parser_types) == 1 else "excel_multi_attachment"
            return ExtractionResult(
                records,
                status,
                parser_type,
                source_email_id,
                metadata={"worksheet_diagnostics": diagnostics},
            )
        return ExtractionResult(
            [],
            ExtractionStatus.INVALID,
            next(iter(parser_types)) if len(parser_types) == 1 else "excel_multi_attachment",
            source_email_id,
            tuple(errors) or ("no allocation worksheet found",),
            metadata={"worksheet_diagnostics": diagnostics},
        )

    html_body = email.get("body_html")
    if normalize_text(html_body):
        plant_code, plant_description = _extract_email_plant(email)
        return parse_html(
            str(html_body),
            source_email_id=source_email_id,
            destination_plant_code=plant_code,
            destination_plant_description=plant_description,
        )

    return ExtractionResult(
        [],
        ExtractionStatus.UNSUPPORTED,
        "format_detection",
        source_email_id,
        ("email has no supported Excel attachment or HTML body",),
    )