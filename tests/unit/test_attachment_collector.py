"""Unit tests for attachment collection and manifest generation."""

from pathlib import Path
import sys
from unittest.mock import MagicMock
import json
import pytest

from scm_allocation.ingestion.collector import collect_attachments


@pytest.mark.skipif(sys.platform != "win32", reason="Requires Windows OS with Outlook")
def test_collect_attachments_mocked(tmp_path: Path, monkeypatch):
    """Verify attachment copying and manifest structure using mocked COM objects."""
    # Create mock attachment
    mock_att = MagicMock()
    mock_att.FileName = "test_sheet.xlsx"
    mock_att.Size = 1024
    mock_att.PropertyAccessor.GetProperty.return_value = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def fake_save(path_str):
        Path(path_str).write_bytes(b"PK\x03\x04fake_excel_bytes")
    mock_att.SaveAsFile.side_effect = fake_save

    # Create mock email item
    mock_item = MagicMock()
    mock_item.UnRead = True
    mock_item.Subject = "Alokasi Item Sample"
    mock_item.SenderEmailType = "SMTP"
    mock_item.SenderEmailAddress = "pm@company.com"
    mock_item.ReceivedTime = "2026-09-07 15:00:00+00:00"
    mock_item.Attachments = [mock_att]

    # Mock Items collection
    mock_items = MagicMock()
    mock_items.Count = 1
    mock_items.__getitem__.return_value = mock_item

    mock_inbox = MagicMock()
    mock_inbox.Items = mock_items

    mock_ns = MagicMock()
    mock_ns.GetDefaultFolder.return_value = mock_inbox

    mock_app = MagicMock()
    mock_app.GetNamespace.return_value = mock_ns

    # Monkeypatch win32com.client.Dispatch
    import win32com.client
    monkeypatch.setattr(win32com.client, "Dispatch", MagicMock(return_value=mock_app))

    manifest = collect_attachments(target_dir=tmp_path, max_emails=1)

    # Verify manifest output
    assert len(manifest) == 1
    entry = manifest[0]
    assert entry["source_email_number"] == 1
    assert entry["original_email_subject"] == "Alokasi Item Sample"
    assert entry["original_attachment_filename"] == "test_sheet.xlsx"
    assert entry["extension"] == ".xlsx"
    assert entry["mime_type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert entry["actual_size_bytes"] > 0

    # Verify saved file and manifest.json file on disk
    saved_file = tmp_path / "test_sheet.xlsx"
    assert saved_file.exists()
    assert saved_file.read_bytes() == b"PK\x03\x04fake_excel_bytes"

    manifest_file = tmp_path / "manifest.json"
    assert manifest_file.exists()
    loaded_manifest = json.loads(manifest_file.read_text(encoding="utf-8"))
    assert len(loaded_manifest) == 1
    assert loaded_manifest[0]["original_attachment_filename"] == "test_sheet.xlsx"
