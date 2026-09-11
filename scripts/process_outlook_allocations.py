"""Run the bounded Outlook allocation scanner."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from scm_allocation.ingestion.category_writer import OutlookCategoryWriter
from scm_allocation.ingestion.history import CsvProcessingLogger
from scm_allocation.models.allocation import ExtractionStatus
from scm_allocation.output.apex import ApexPushError, load_env_file, push_records_to_apex
from scm_allocation.processing import OutlookEmailScanner, OutlookScanConfig


def _push_and_categorize(summaries, *, outlook, write_categories: bool,
                         history_logger: CsvProcessingLogger) -> None:
    """Push each candidate's records to APEX, then — ONLY for pushes that
    succeeded, and only if requested — apply the Outlook category ourselves.

    This is deliberately a SEPARATE step from scanner.scan() (which is always
    called with dry_run=True when --push-to-apex is set — see main()). If
    scan() wrote categories itself before we knew whether the push to APEX
    succeeded, a failed push would leave the email marked "already processed"
    forever, silently dropping its data. Categorizing only after a confirmed
    push keeps a failed push retryable on the next run.
    """
    api_base    = os.getenv("APEX_API_BASE", "").strip()
    machine_key = os.getenv("APEX_MACHINE_KEY", "").strip()
    source_id_raw = os.getenv("APEX_SOURCE_ID", "").strip()
    if not api_base or not machine_key or not source_id_raw:
        print(
            "[SCM] ERROR: --push-to-apex butuh APEX_API_BASE, APEX_MACHINE_KEY, "
            "APEX_SOURCE_ID (set di .env atau environment). Tidak ada yang di-push.",
            flush=True,
        )
        return
    source_id = int(source_id_raw)

    namespace = None
    category_writer = None
    if write_categories:
        namespace = outlook.GetNamespace("MAPI")
        category_writer = OutlookCategoryWriter(namespace.Categories)

    for summary in summaries:
        records = summary.metadata.get("records", [])
        if not records or summary.extraction_status not in (ExtractionStatus.VALID, ExtractionStatus.PARTIAL):
            continue
        try:
            result = push_records_to_apex(
                api_base=api_base, machine_key=machine_key, source_id=source_id,
                email_id=summary.email_id, sender=summary.sender, subject=summary.subject,
                records=records,
            )
        except ApexPushError as err:
            print(f"[SCM] APEX PUSH FAILED: {summary.subject} | {err}", flush=True)
            continue

        dup = result.get("duplicate_of")
        dup_note = f" (duplikat dari submission #{dup['id']})" if dup else ""
        print(f"[SCM] APEX submission #{result['id']} dibuat untuk: {summary.subject}{dup_note}", flush=True)

        if not write_categories or not summary.category:
            continue
        try:
            item = namespace.GetItemFromID(summary.email_id)
            outcome = category_writer.apply_category(item, summary.category)
        except Exception as err:
            print(f"[SCM] WARN: gagal menandai kategori Outlook untuk {summary.subject}: {err}", flush=True)
            continue
        summary.category_written = outcome.written
        if not outcome.written and outcome.reason != "category already present":
            print(f"[SCM] WARN: kategori tidak tertulis untuk {summary.subject}: {outcome.reason}", flush=True)
        elif history_logger and outcome.written:
            try:
                attachment_count = int(getattr(getattr(item, "Attachments", []), "Count", 0))
            except (TypeError, ValueError):
                attachment_count = 0
            history_logger.record_successful_processing(summary, attachment_count)


def main() -> int:
    parser = argparse.ArgumentParser(description="Scan Outlook allocation emails safely.")
    parser.add_argument("--folder", default="Inbox", help="Outlook folder name; default: Inbox")
    parser.add_argument("--max-emails", type=int, default=None, help="Maximum newest items to inspect; omitted scans the full folder")
    parser.add_argument(
        "--write-categories",
        action="store_true",
        help="Enable SCM Bot category writes; omitted means dry-run",
    )
    parser.add_argument(
        "--preview-records",
        action="store_true",
        help="Show a preview of extracted AllocationRecord data after the scan summary",
    )
    parser.add_argument(
        "--push-to-apex",
        action="store_true",
        help="Push extracted records to APEX's Ingest Alokasi PM (/api/submissions/from-outlook). "
             "Requires APEX_API_BASE, APEX_MACHINE_KEY, APEX_SOURCE_ID (.env or environment). "
             "When set, --write-categories only marks emails whose push to APEX actually succeeded.",
    )
    args = parser.parse_args()

    load_env_file(str(Path(__file__).resolve().parent.parent / ".env"))

    want_categories = args.write_categories
    # Dengan --push-to-apex, scan() SENDIRI selalu dijalankan dry-run: kategori
    # ditulis belakangan di _push_and_categorize, cuma untuk email yang push-nya
    # sukses. Tanpa flag ini perilakunya sama persis seperti sebelumnya.
    scan_dry_run = (not want_categories) if not args.push_to_apex else True
    print(
        f"[SCM] Starting Outlook allocation scan (folder={args.folder}, "
        f"max_emails={args.max_emails}, dry_run={scan_dry_run}).",
        flush=True,
    )
    print("[SCM] Connecting to Outlook...", flush=True)
    config = OutlookScanConfig(
        folder_name=args.folder,
        max_emails=args.max_emails,
        dry_run=scan_dry_run,
    )
    history_logger = CsvProcessingLogger(Path("logs"))
    scanner = OutlookEmailScanner(
        config=config,
        history_logger=history_logger,
        progress_callback=lambda message: print(f"[SCM] {message}", flush=True),
    )
    outlook_application = None
    if args.push_to_apex:
        # Koneksi dipegang sendiri (bukan diserahkan ke scan() lalu dibuang)
        # supaya bisa dipakai lagi untuk GetItemFromID setelah push ke APEX.
        if sys.platform != "win32":
            print("[SCM] ERROR: Outlook COM automation cuma jalan di Windows.", flush=True)
            return 1
        import win32com.client
        try:
            outlook_application = win32com.client.Dispatch("Outlook.Application")
        except Exception as err:
            print(f"[SCM] ERROR: Failed to connect to Outlook desktop: {err}", flush=True)
            return 1
    try:
        summaries = scanner.scan(outlook_application=outlook_application)
    except Exception as err:
        counters = scanner.run_counters()
        counters["error_count"] += 1
        try:
            history_logger.record_run(args.folder, counters, not want_categories)
        except Exception as audit_err:
            print(f"[SCM] AUDIT ERROR: {audit_err}", flush=True)
        print(f"[SCM] ERROR: {err}", flush=True)
        return 1

    counters = scanner.run_counters()
    history_logger.record_run(args.folder, counters, not want_categories)

    print(
        f"[SCM] Scan complete: inspected={scanner.last_scan_inspected}, "
        f"candidates={scanner.last_scan_candidates}.",
        flush=True,
    )

    for summary in summaries:
        details = summary.error_information or summary.review_information
        detail_text = f" | details={' ; '.join(details)}" if details else ""
        status = "SKIPPED_ALREADY_PROCESSED" if summary.skipped else summary.extraction_status.value.upper()

        print(
            f"Subject: {summary.subject} | Email ID: {summary.email_id} | Status: {status} | "
            f"Parser: {summary.parser_type} | records={summary.record_count} | "
            f"Planned category: {summary.category or 'None'} | Written: {summary.category_written}"
            f"{detail_text}",
            flush=True,
        )

        for diagnostic in summary.metadata.get("worksheet_diagnostics", []):
            workbook = diagnostic.get("workbook", "<unknown workbook>")
            sheet = diagnostic.get("sheet", "<unknown>")
            status = diagnostic.get("status", "unknown")
            detail = diagnostic.get("format") or diagnostic.get("reason", "")
            records = diagnostic.get("records")
            record_text = f" records={records}" if records is not None else ""

            print(
                f"  workbook={workbook} sheet={sheet} -> {status}: {detail}{record_text}",
                flush=True,
            )

        for record in summary.metadata.get("records", []):
            context_val = (
                record.allocation_context.value
                if hasattr(record.allocation_context, "value")
                else str(record.allocation_context)
            )

            print(
                f"  -> AllocationRecord: material={record.material_code or '<none>'} ({record.material_description or ''}) "
                f"qty={record.quantity} from={record.issuing_warehouse_code or '<none>'} (sloc={record.issuing_warehouse_sloc or '<none>'}) "
                f"to={record.destination_plant_code or '<none>'} (sloc={record.destination_sloc or '<none>'}) "
                f"context={context_val} src={record.source_filename or ''}/{record.source_sheet or ''} row={record.source_row}",
                flush=True,
            )

    if args.push_to_apex:
        print("", flush=True)
        print("=" * 60, flush=True)
        print("APEX PUSH (Ingest Alokasi PM)", flush=True)
        print("=" * 60, flush=True)
        _push_and_categorize(
            summaries, outlook=outlook_application,
            write_categories=want_categories, history_logger=history_logger,
        )

    print(f"Scanned: {counters['scanned_count']}", flush=True)
    print(f"Candidates: {counters['candidate_count']}", flush=True)
    print(f"Processed: {counters['processed_count']}", flush=True)
    print(f"Skipped already processed: {counters['skipped_already_processed']}", flush=True)
    print(f"Scraped: {counters['scraped_count']}", flush=True)
    print(f"Review: {counters['review_count']}", flush=True)
    print(f"Invalid/Excluded: {counters['invalid_or_excluded_count']}", flush=True)
    print(f"Write failures: {counters['write_failure_count']}", flush=True)
    print(f"Errors: {counters['error_count']}", flush=True)

    if args.preview_records:
        print("", flush=True)
        print("=" * 60, flush=True)
        print("SCRAPED DATA PREVIEW", flush=True)
        print("=" * 60, flush=True)

        for summary in summaries:
            if summary.extraction_status.value.upper() != "VALID":
                continue

            records = summary.metadata.get("records", [])
            if not records:
                continue

            print("", flush=True)
            print(f"Email: {summary.subject}", flush=True)
            print(f"Parser: {summary.parser_type}", flush=True)
            print(f"Records: {len(records)}", flush=True)

            preview_records = records[:5]

            for record in preview_records:
                print(
                    f"  -> material={record.material_code or '<none>'} | "
                    f"description={record.material_description or '<none>'} | "
                    f"destination={record.destination_plant_code or '<none>'} | "
                    f"qty={record.quantity} | "
                    f"sloc={record.destination_sloc or '<none>'}",
                    flush=True,
                )

            if len(records) > 5:
                print(
                    f"  ... showing 5 of {len(records)} records",
                    flush=True,
                )

    print(f"Dry-run: {not want_categories}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())