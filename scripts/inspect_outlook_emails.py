"""Script to run read-only email content inspection on top 5 Outlook messages.

Saves raw data into data/inspection/ for local reconnaissance without modifying mailbox state.
"""

from __future__ import annotations

from pathlib import Path
import sys

from scm_allocation.ingestion.inspector import inspect_email_item, save_inspection_result


def main() -> int:
    print("=" * 70)
    print("READ-ONLY OUTLOOK EMAIL CONTENT INSPECTION (TOP 5)")
    print("=" * 70)

    try:
        import win32com.client
    except ImportError:
        print("[ERROR] pywin32 is not installed.")
        return 1

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox = 6
    except Exception as err:
        print(f"[ERROR] Failed connecting to Outlook: {err}")
        return 1

    items = inbox.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    output_dir = Path("data/inspection")
    output_dir.mkdir(parents=True, exist_ok=True)

    inspected_records = []
    total_count = items.Count
    fetch_count = min(5, total_count)

    print(f"Total Inbox items: {total_count}. Inspecting top {fetch_count} emails...\n")

    for i in range(1, fetch_count + 1):
        item = items[i]
        record = inspect_email_item(item, index=i)
        saved = save_inspection_result(record, output_dir)
        inspected_records.append(record)

        print(f"[{record['index']}] Subject: {record['subject']}")
        print(f"    Sender:        {record['sender']}")
        print(f"    Received:      {record['received_time']}")
        print(f"    Format:        {record['body_format']}")
        print(f"    Plain Text:    {len(record['body_plain_text'])} chars")
        print(f"    HTML Body:     {len(record['body_html'])} chars")
        print(f"    HTML Table:    {'YES' if record['has_html_table'] else 'NO'}")
        print(f"    Attachments:   {record['attachment_count']}")
        for att in record['attachments']:
            print(f"      - {att['filename']} ({att['extension']}, {att['size_bytes']} bytes, {att['content_type']})")
        print(f"    UnRead State:  {'PRESERVED' if record['unread_status_preserved'] else 'CHANGED'}")
        print(f"    Saved Files:   {saved['json']}")
        print("-" * 70)

    print("\n[SUCCESS] All 5 emails inspected and saved to data/inspection/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
