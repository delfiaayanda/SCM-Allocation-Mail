"""Attachment collector module for exporting attachments from Outlook MailItems."""

from __future__ import annotations

import json
from pathlib import Path
import sys
from typing import Any, Dict, List

from scm_allocation.ingestion.inspector import get_sender_address


def collect_attachments(
    target_dir: Path = Path("tests/fixtures/attachments"),
    max_emails: int = 5,
) -> List[Dict[str, Any]]:
    """Download attachments from the top inspected emails into target_dir."""
    if sys.platform != "win32":
        raise RuntimeError("Outlook COM automation is only supported on Windows.")

    try:
        import win32com.client
    except ImportError as err:
        raise RuntimeError("pywin32 is required for Outlook automation.") from err

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox = 6
    except Exception as err:
        raise RuntimeError(f"Failed to connect to Outlook desktop: {err}") from err

    items = inbox.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    target_dir.mkdir(parents=True, exist_ok=True)
    manifest: List[Dict[str, Any]] = []

    try:
        total_count = int(getattr(items, "Count", 0))
    except (TypeError, ValueError):
        total_count = 0
    fetch_count = min(max_emails, total_count)

    for i in range(1, fetch_count + 1):
        item = items[i]
        unread_before = getattr(item, "UnRead", None)

        subject = getattr(item, "Subject", "<No Subject>")
        sender = get_sender_address(item)
        received_time = str(getattr(item, "ReceivedTime", "<Unknown>"))
        attachments = getattr(item, "Attachments", [])
        try:
            att_count = getattr(attachments, "Count", len(attachments))
        except Exception:
            att_count = len(attachments)

        if att_count == 0:
            unread_after = getattr(item, "UnRead", None)
            assert unread_before == unread_after, f"UnRead status mutated for email {i}"
            continue

        for att in attachments:
            filename = getattr(att, "FileName", "")
            if not filename:
                continue

            ext = Path(filename).suffix.lower()
            size = getattr(att, "Size", 0)

            # Retrieve MIME tag if available
            mime = "<Unknown>"
            try:
                pr_mime = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
                prop_accessor = getattr(att, "PropertyAccessor", None)
                if prop_accessor:
                    mime = str(prop_accessor.GetProperty(pr_mime))
            except Exception:
                pass

            # Destination path: preserve original filename
            dest_path = target_dir / filename
            att.SaveAsFile(str(dest_path.resolve()))

            actual_size = dest_path.stat().st_size

            manifest_entry = {
                "source_email_number": i,
                "original_email_subject": subject,
                "original_sender": sender,
                "received_time": received_time,
                "original_attachment_filename": filename,
                "extension": ext,
                "reported_size_bytes": size,
                "actual_size_bytes": actual_size,
                "local_fixture_path": str(dest_path.as_posix()),
                "mime_type": mime,
            }
            manifest.append(manifest_entry)

        unread_after = getattr(item, "UnRead", None)
        assert unread_before == unread_after, f"UnRead status mutated for email {i} after copying attachments"

    # Save manifest.json
    manifest_file = target_dir / "manifest.json"
    with open(manifest_file, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    return manifest
