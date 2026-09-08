"""Append-only CSV audit logging for Outlook allocation scans."""

from __future__ import annotations

import csv
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Mapping
from uuid import uuid4

if TYPE_CHECKING:
    from scm_allocation.processing import ProcessingSummary


HISTORY_FIELDS = (
    "run_id", "timestamp", "message_id", "subject", "sender", "received_time",
    "extraction_status", "parser_type", "planned_category", "actual_category",
    "category_written", "dry_run", "attachment_count", "record_count", "reason_error",
)
RUN_FIELDS = (
    "run_id", "timestamp", "folder", "dry_run", "scanned_count", "candidate_count",
    "processed_count", "skipped_already_processed", "scraped_count", "review_count",
    "invalid_or_excluded_count", "write_failure_count", "error_count",
)


def _timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


class CsvProcessingLogger:
    """Append-only audit log for successful live category persistence only."""

    def __init__(self, log_dir: Path = Path("logs"), run_id: str | None = None) -> None:
        self.log_dir = log_dir
        self.history_path = log_dir / "allocation_bot_history.csv"
        self.runs_path = log_dir / "allocation_bot_runs.csv"
        self.run_id = run_id or uuid4().hex

    def _append(self, path: Path, fields: tuple[str, ...], row: Mapping[str, object]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self._migrate_legacy_schema(path, fields)
        needs_header = not path.exists() or path.stat().st_size == 0
        with path.open("a", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
            if needs_header:
                writer.writeheader()
            writer.writerow({field: row.get(field, "") for field in fields})

    def _migrate_legacy_schema(self, path: Path, fields: tuple[str, ...]) -> None:
        """Preserve prior audit rows before using the current, explicit schema."""
        if not path.exists() or path.stat().st_size == 0:
            return
        with path.open(newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            if tuple(reader.fieldnames or ()) == fields:
                return
            legacy_rows = list(reader)

        def value(legacy: Mapping[str, str], *names: str) -> str:
            return next((legacy.get(name, "") for name in names if legacy.get(name, "")), "")

        migrated: list[dict[str, str]] = []
        for legacy in legacy_rows:
            if path == self.history_path:
                written = value(legacy, "category_written", "written", "category_write_success")
                migrated.append({
                    "run_id": legacy.get("run_id", ""),
                    "timestamp": legacy.get("timestamp", ""),
                    "message_id": legacy.get("message_id", ""),
                    "subject": legacy.get("subject", ""),
                    "sender": legacy.get("sender", ""),
                    "received_time": value(legacy, "received_time", "email_received_time"),
                    "extraction_status": legacy.get("extraction_status", ""),
                    "parser_type": legacy.get("parser_type", ""),
                    "planned_category": value(legacy, "planned_category", "category"),
                    "actual_category": value(legacy, "actual_category", "category") if written == "True" else "",
                    "category_written": written,
                    "dry_run": legacy.get("dry_run", ""),
                    "attachment_count": legacy.get("attachment_count", ""),
                    "record_count": value(legacy, "record_count", "records_extracted"),
                    "reason_error": value(legacy, "reason_error", "error"),
                })
            else:
                migrated.append({
                    "run_id": legacy.get("run_id", ""),
                    "timestamp": legacy.get("timestamp", ""),
                    "folder": legacy.get("folder", ""),
                    "dry_run": legacy.get("dry_run", ""),
                    "scanned_count": value(legacy, "scanned_count", "scanned"),
                    "candidate_count": value(legacy, "candidate_count", "candidates"),
                    "processed_count": legacy.get("processed_count", ""),
                    "skipped_already_processed": legacy.get("skipped_already_processed", ""),
                    "scraped_count": value(legacy, "scraped_count", "valid"),
                    "review_count": legacy.get("review_count", "review"),
                    "invalid_or_excluded_count": value(legacy, "invalid_or_excluded_count", "invalid"),
                    "write_failure_count": value(legacy, "write_failure_count", "category_write_failure"),
                    "error_count": value(legacy, "error_count", "errors"),
                })
        with path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            writer.writerows(migrated)

    def record_successful_processing(self, summary: ProcessingSummary, attachment_count: int) -> None:
        """Record a row only after a valid/partial email was categorized live."""
        self._append(
            self.history_path,
            HISTORY_FIELDS,
            {
                "run_id": self.run_id,
                "timestamp": _timestamp(),
                "message_id": summary.email_id,
                "subject": summary.subject,
                "sender": summary.sender,
                "received_time": summary.received_at,
                "extraction_status": summary.extraction_status.value,
                "parser_type": summary.parser_type,
                "planned_category": summary.category or "",
                "actual_category": summary.category or "",
                "category_written": summary.category_written,
                "dry_run": False,
                "attachment_count": attachment_count,
                "record_count": summary.record_count,
                "reason_error": summary.metadata.get("category_write_reason", ""),
            },
        )

    def record_run(self, folder: str, counters: Mapping[str, object], dry_run: bool) -> None:
        self._append(
            self.runs_path,
            RUN_FIELDS,
            {
                "run_id": self.run_id,
                "timestamp": _timestamp(),
                "folder": folder,
                "dry_run": dry_run,
                **{field: counters.get(field, 0) for field in RUN_FIELDS if field.endswith("_count")},
                "skipped_already_processed": counters.get("skipped_already_processed", 0),
            },
        )
