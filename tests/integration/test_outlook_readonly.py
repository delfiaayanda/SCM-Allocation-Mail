"""Integration test for read-only Outlook inspection."""

import sys
from pathlib import Path
import pytest

# Ensure scripts directory is in sys.path for test execution
SCRIPTS_DIR = Path(__file__).resolve().parent.parent.parent / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

try:
    from test_outlook_readonly import fetch_sample_email_metadata
except ImportError:
    fetch_sample_email_metadata = None


@pytest.mark.skipif(sys.platform != "win32", reason="Requires Windows OS with Outlook")
def test_outlook_readonly_metadata_fetch():
    """Verify that Outlook metadata can be retrieved safely in read-only mode."""
    assert fetch_sample_email_metadata is not None, "Could not import fetch_sample_email_metadata"

    try:
        emails = fetch_sample_email_metadata(limit=3)
    except RuntimeError as err:
        pytest.skip(f"Outlook desktop session not accessible in current environment: {err}")

    assert isinstance(emails, list)
    for email in emails:
        assert "subject" in email
        assert "sender" in email
        assert "received_time" in email
        assert email.get("unread_status_preserved") is True, (
            "Read-only invariant violated: message read status changed!"
        )
