"""Email inspection module for analyzing raw Outlook allocation request emails.

Strict safety constraints:
- READ-ONLY access to Outlook items.
- No modifications, deletions, replies, forwards, or state transitions (UnRead status is preserved).
- Attachments are inspected for metadata only; files are NOT downloaded.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import sys
from typing import Any, Dict, List, Optional


def get_sender_address(item: Any) -> str:
    """Safely extract sender address, resolving Exchange addresses when possible."""
    try:
        sender_type = getattr(item, "SenderEmailType", "")
        if sender_type == "EX":
            sender = getattr(item, "Sender", None)
            if sender:
                try:
                    ex_user = sender.GetExchangeUser()
                    if ex_user and ex_user.PrimarySmtpAddress:
                        return str(ex_user.PrimarySmtpAddress)
                except Exception:
                    pass
            try:
                pr_smtp_address = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                prop_accessor = getattr(item, "PropertyAccessor", None)
                if prop_accessor:
                    smtp = prop_accessor.GetProperty(pr_smtp_address)
                    if smtp:
                        return str(smtp)
            except Exception:
                pass

        email = getattr(item, "SenderEmailAddress", None)
        if email:
            return str(email)
        name = getattr(item, "SenderName", None)
        if name:
            return str(name)
    except Exception:
        pass
    return "<Unknown Sender>"


def inspect_attachments(item: Any) -> List[Dict[str, Any]]:
    """Inspect attachment metadata without downloading contents."""
    attachments_meta: List[Dict[str, Any]] = []
    try:
        for att in item.Attachments:
            filename = getattr(att, "FileName", "")
            ext = Path(filename).suffix.lower() if filename else ""
            size = getattr(att, "Size", 0)

            # Attempt to read MIME tag via MAPI property tag 0x370E001F
            content_type = ""
            try:
                pr_mime = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
                prop_accessor = getattr(att, "PropertyAccessor", None)
                if prop_accessor:
                    content_type = str(prop_accessor.GetProperty(pr_mime))
            except Exception:
                content_type = ""

            attachments_meta.append({
                "filename": filename,
                "extension": ext,
                "size_bytes": size,
                "content_type": content_type or "<Unknown>",
            })
    except Exception as err:
        attachments_meta.append({
            "error": f"Failed inspecting attachments: {err}",
        })
    return attachments_meta


def inspect_email_item(item: Any, index: int) -> Dict[str, Any]:
    """Extract metadata and raw bodies from a single Outlook MailItem in read-only mode."""
    unread_before = getattr(item, "UnRead", None)

    entry_id = getattr(item, "EntryID", None)
    subject = getattr(item, "Subject", "<No Subject>")
    received_time = getattr(item, "ReceivedTime", None)
    sender = get_sender_address(item)

    body_text = getattr(item, "Body", "") or ""
    html_body = getattr(item, "HTMLBody", "") or ""

    attachments = inspect_attachments(item)

    unread_after = getattr(item, "UnRead", None)

    # Detect body formats and presence of HTML tables
    has_plain_text = bool(body_text.strip())
    has_html = bool(html_body.strip())
    has_html_table = "<table" in html_body.lower() if has_html else False

    if has_plain_text and has_html_table:
        format_type = "mixed (text + html table)"
    elif has_html_table:
        format_type = "html table"
    elif has_html:
        format_type = "html"
    elif has_plain_text:
        format_type = "plain text"
    else:
        format_type = "empty"

    return {
        "index": index,
        "entry_id": entry_id,
        "subject": subject,
        "sender": sender,
        "received_time": str(received_time) if received_time else "<Unknown>",
        "body_format": format_type,
        "has_plain_text": has_plain_text,
        "has_html": has_html,
        "has_html_table": has_html_table,
        "attachment_count": len(attachments),
        "attachments": attachments,
        "unread_status_preserved": (unread_before == unread_after),
        "body_plain_text": body_text,
        "body_html": html_body,
    }


def save_inspection_result(
    data: Dict[str, Any],
    output_dir: Path,
) -> Dict[str, str]:
    """Save raw inspection data into JSON, plain text, and HTML files."""
    output_dir.mkdir(parents=True, exist_ok=True)
    idx_str = f"{data['index']:03d}"

    json_filename = f"email_{idx_str}.json"
    txt_filename = f"email_{idx_str}_body.txt"
    html_filename = f"email_{idx_str}_body.html"

    json_path = output_dir / json_filename
    txt_path = output_dir / txt_filename
    html_path = output_dir / html_filename

    # Save plain text body
    if data.get("body_plain_text"):
        with open(txt_path, "w", encoding="utf-8", errors="replace") as f:
            f.write(data["body_plain_text"])

    # Save HTML body
    if data.get("body_html"):
        with open(html_path, "w", encoding="utf-8", errors="replace") as f:
            f.write(data["body_html"])

    # Prepare JSON structure (preserving body excerpts and references to text/html files)
    json_record = {
        "index": data["index"],
        "entry_id": data["entry_id"],
        "subject": data["subject"],
        "sender": data["sender"],
        "received_time": data["received_time"],
        "body_format": data["body_format"],
        "has_plain_text": data["has_plain_text"],
        "has_html": data["has_html"],
        "has_html_table": data["has_html_table"],
        "attachment_count": data["attachment_count"],
        "attachments": data["attachments"],
        "unread_status_preserved": data["unread_status_preserved"],
        "saved_files": {
            "json": json_filename,
            "plain_text": txt_filename if data.get("body_plain_text") else None,
            "html": html_filename if data.get("body_html") else None,
        },
        "body_plain_text": data["body_plain_text"],
        "body_html": data["body_html"],
    }

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(json_record, f, indent=2, ensure_ascii=False)

    return {
        "json": str(json_path),
        "txt": str(txt_path) if data.get("body_plain_text") else "",
        "html": str(html_path) if data.get("body_html") else "",
    }
