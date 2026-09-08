"""Run the bounded Outlook allocation scanner."""

from __future__ import annotations

import argparse
from pathlib import Path

from scm_allocation.ingestion.history import CsvProcessingLogger
from scm_allocation.processing import OutlookEmailScanner, OutlookScanConfig


def main() -> int:
    parser = argparse.ArgumentParser(description="Scan Outlook allocation emails safely.")
    parser.add_argument("--folder", default="Inbox", help="Outlook folder name; default: Inbox")
    parser.add_argument("--max-emails", type=int, default=None, help="Maximum newest items to inspect; omitted scans the full folder")
    parser.add_argument(
        "--write-categories",
        action="store_true",
        help="Enable SCM Bot category writes; omitted means dry-run",
    )
    args = parser.parse_args()

    dry_run = not args.write_categories
    print(
        f"[SCM] Starting Outlook allocation scan (folder={args.folder}, "
        f"max_emails={args.max_emails}, dry_run={dry_run}).",
        flush=True,
    )
    print("[SCM] Connecting to Outlook...", flush=True)
    config = OutlookScanConfig(
        folder_name=args.folder,
        max_emails=args.max_emails,
        dry_run=dry_run,
    )
    history_logger = CsvProcessingLogger(Path("logs"))
    scanner = OutlookEmailScanner(
        config=config,
        history_logger=history_logger,
        progress_callback=lambda message: print(f"[SCM] {message}", flush=True),
    )
    try:
        summaries = scanner.scan()
    except Exception as err:
        counters = scanner.run_counters()
        counters["error_count"] += 1
        try:
            history_logger.record_run(args.folder, counters, dry_run)
        except Exception as audit_err:
            print(f"[SCM] AUDIT ERROR: {audit_err}", flush=True)
        print(f"[SCM] ERROR: {err}", flush=True)
        return 1
    counters = scanner.run_counters()
    history_logger.record_run(args.folder, counters, dry_run)
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
            f"Subject: {summary.subject} | Status: {status} | "
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
            print(f"  workbook={workbook} sheet={sheet} -> {status}: {detail}{record_text}", flush=True)
    print(f"Scanned: {counters['scanned_count']}", flush=True)
    print(f"Candidates: {counters['candidate_count']}", flush=True)
    print(f"Processed: {counters['processed_count']}", flush=True)
    print(f"Skipped already processed: {counters['skipped_already_processed']}", flush=True)
    print(f"Scraped: {counters['scraped_count']}", flush=True)
    print(f"Review: {counters['review_count']}", flush=True)
    print(f"Invalid/Excluded: {counters['invalid_or_excluded_count']}", flush=True)
    print(f"Write failures: {counters['write_failure_count']}", flush=True)
    print(f"Errors: {counters['error_count']}", flush=True)
    print(f"Dry-run: {dry_run}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
