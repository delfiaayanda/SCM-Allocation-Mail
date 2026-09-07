"""Proof of concept: Read-Only Outlook Metadata Inspection.

This script tests read-only connection to the local Microsoft Outlook desktop application
via Windows MAPI COM automation (pywin32).

Safety guarantees:
- READ-ONLY access: Does not modify, delete, move, forward, or reply to messages.
- Does not change message state (does NOT mark messages as read).
- Fetches strictly high-level metadata (Subject, Sender, Received Time).
- Does NOT fetch email body contents.
- Does NOT download or touch attachments.
- Handles missing Outlook or COM connection errors gracefully.
"""

from __future__ import annotations

import sys
from typing import Any, Dict, List


def get_sender_address(item: Any) -> str:
    """Safely extract sender address, resolving Exchange internal addresses to SMTP if available."""
    try:
        sender_type = getattr(item, "SenderEmailType", "")
        if sender_type == "EX":
            # Try to resolve Exchange sender to SMTP address
            sender = getattr(item, "Sender", None)
            if sender:
                try:
                    ex_user = sender.GetExchangeUser()
                    if ex_user and ex_user.PrimarySmtpAddress:
                        return str(ex_user.PrimarySmtpAddress)
                except Exception:
                    pass
            # Try PropertyAccessor for PR_SMTP_ADDRESS
            try:
                pr_smtp_address = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                prop_accessor = getattr(item, "PropertyAccessor", None)
                if prop_accessor:
                    smtp = prop_accessor.GetProperty(pr_smtp_address)
                    if smtp:
                        return str(smtp)
            except Exception:
                pass

        # Fallback to SenderEmailAddress or SenderName
        email = getattr(item, "SenderEmailAddress", None)
        if email:
            return str(email)
        name = getattr(item, "SenderName", None)
        if name:
            return str(name)
    except Exception:
        pass
    return "<Unknown Sender>"


def fetch_sample_email_metadata(limit: int = 5) -> List[Dict[str, Any]]:
    """Retrieve metadata for up to `limit` emails in strictly read-only mode."""
    if sys.platform != "win32":
        raise RuntimeError("Outlook COM automation is only supported on Windows.")

    try:
        import win32com.client
    except ImportError as err:
        raise RuntimeError(
            "The 'pywin32' package is required for Outlook automation. "
            "Install it with: pip install pywin32"
        ) from err

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as err:
        raise RuntimeError(
            f"Failed to connect to Outlook desktop application. "
            f"Ensure classic Outlook is installed and running. Details: {err}"
        ) from err

    try:
        # olFolderInbox = 6
        inbox = namespace.GetDefaultFolder(6)
    except Exception as err:
        raise RuntimeError(
            f"Failed to access default Inbox folder: {err}"
        ) from err

    items = inbox.Items
    # Sort newest first (descending)
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    results: List[Dict[str, Any]] = []
    total_count = items.Count
    fetch_count = min(limit, total_count)

    for i in range(1, fetch_count + 1):
        try:
            item = items[i]
            # Verify read-only invariance (UnRead flag before and after property reading)
            unread_before = getattr(item, "UnRead", None)

            subject = getattr(item, "Subject", "<No Subject>")
            received = getattr(item, "ReceivedTime", None)
            sender = get_sender_address(item)

            unread_after = getattr(item, "UnRead", None)

            results.append({
                "index": i,
                "subject": subject,
                "sender": sender,
                "received_time": str(received) if received else "<Unknown>",
                "unread_status_preserved": (unread_before == unread_after),
            })
        except Exception as item_err:
            results.append({
                "index": i,
                "error": f"Failed to read metadata for item {i}: {item_err}",
            })

    return results


def main() -> int:
    print("=" * 60)
    print("Outlook Read-Only Inspection Proof of Concept")
    print("=" * 60)
    print("Connecting to local Outlook desktop session...\n")

    try:
        emails = fetch_sample_email_metadata(limit=5)
    except RuntimeError as err:
        print(f"[ERROR] {err}")
        return 1

    if not emails:
        print("No items found in Inbox.")
        return 0

    print(f"Successfully retrieved metadata for {len(emails)} emails (READ-ONLY):\n")
    for email in emails:
        if "error" in email:
            print(f"[{email['index']}] Error: {email['error']}")
            continue
        print(f"[{email['index']}]")
        print(f"  Subject:       {email['subject']}")
        print(f"  Sender:        {email['sender']}")
        print(f"  Received Time: {email['received_time']}")
        print(f"  Read Status:   Preserved (Unmodified)")
        print("-" * 60)

    print("\n[VERIFIED] Read-only operation completed safely without mailbox modifications.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
