"""Unit tests for the email inspector module."""

from pathlib import Path
from unittest.mock import MagicMock
import json
import pytest

from scm_allocation.ingestion.inspector import (
    get_sender_address,
    inspect_attachments,
    inspect_email_item,
    save_inspection_result,
)


def test_get_sender_address_direct_email():
    item = MagicMock()
    item.SenderEmailType = "SMTP"
    item.SenderEmailAddress = "pm@example.com"
    item.SenderName = "Product Manager"
    assert get_sender_address(item) == "pm@example.com"


def test_get_sender_address_exchange_user_fallback():
    item = MagicMock()
    item.SenderEmailType = "EX"
    item.Sender.GetExchangeUser.return_value.PrimarySmtpAddress = "user@company.com"
    assert get_sender_address(item) == "user@company.com"


def test_inspect_attachments_metadata():
    item = MagicMock()
    mock_att = MagicMock()
    mock_att.FileName = "Allocation_Sept.xlsx"
    mock_att.Size = 24000
    mock_att.PropertyAccessor.GetProperty.return_value = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    item.Attachments = [mock_att]

    attachments = inspect_attachments(item)
    assert len(attachments) == 1
    assert attachments[0]["filename"] == "Allocation_Sept.xlsx"
    assert attachments[0]["extension"] == ".xlsx"
    assert attachments[0]["size_bytes"] == 24000
    assert "spreadsheetml" in attachments[0]["content_type"]


def test_inspect_email_item_mixed_format():
    item = MagicMock()
    item.UnRead = True
    item.EntryID = "00000000TEST123"
    item.Subject = "Allocation Request Table"
    item.ReceivedTime = "2026-09-07 15:00:00+00:00"
    item.SenderEmailType = "SMTP"
    item.SenderEmailAddress = "sender@company.com"
    item.Body = "Please find table below."
    item.HTMLBody = "<html><body><table><tr><th>MATERIAL</th></tr></table></body></html>"
    item.Attachments = []

    result = inspect_email_item(item, index=1)
    assert result["index"] == 1
    assert result["subject"] == "Allocation Request Table"
    assert result["body_format"] == "mixed (text + html table)"
    assert result["has_html_table"] is True
    assert result["unread_status_preserved"] is True


def test_save_inspection_result(tmp_path: Path):
    mock_data = {
        "index": 1,
        "entry_id": "0000123",
        "subject": "Test Email",
        "sender": "test@domain.com",
        "received_time": "2026-09-07 12:00:00",
        "body_format": "plain text",
        "has_plain_text": True,
        "has_html": False,
        "has_html_table": False,
        "attachment_count": 0,
        "attachments": [],
        "unread_status_preserved": True,
        "body_plain_text": "Sample text body content",
        "body_html": "",
    }

    saved = save_inspection_result(mock_data, tmp_path)
    assert Path(saved["json"]).exists()
    assert Path(saved["txt"]).exists()

    with open(saved["json"], "r", encoding="utf-8") as f:
        loaded = json.load(f)
    assert loaded["subject"] == "Test Email"
    assert loaded["body_plain_text"] == "Sample text body content"
