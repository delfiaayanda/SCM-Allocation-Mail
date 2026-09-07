"""Run the bounded Outlook allocation scanner."""

from __future__ import annotations

import argparse

from scm_allocation.processing import OutlookEmailScanner, OutlookScanConfig


def main() -> int:
    parser = argparse.ArgumentParser(description="Scan Outlook allocation emails safely.")
    parser.add_argument("--folder", default="Inbox", help="Outlook folder name; default: Inbox")
    parser.add_argument("--max-emails", type=int, default=100, help="Maximum newest items to inspect")
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
    scanner = OutlookEmailScanner(config=config, progress_callback=lambda message: print(f"[SCM] {message}", flush=True))
    try:
        summaries = scanner.scan()
    except Exception as err:
        print(f"[SCM] ERROR: {err}", flush=True)
        return 1
    print(
        f"[SCM] Scan complete: inspected={scanner.last_scan_inspected}, "
        f"candidates={scanner.last_scan_candidates}.",
        flush=True,
    )
    for summary in summaries:
        details = summary.error_information or summary.review_information
        detail_text = f" | details={' ; '.join(details)}" if details else ""
        print(
            f"Subject: {summary.subject} | Status: {summary.extraction_status.value.upper()} | "
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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())