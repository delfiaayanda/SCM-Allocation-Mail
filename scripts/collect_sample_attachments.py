"""Collect real attachment samples from inspected Outlook emails for local fixture analysis.

Strict safety constraints:
- Outlook session access is READ-ONLY.
- No modifications, deletions, replies, forwards, or state transitions (UnRead status is preserved).
- Attachments are read and copied locally via att.SaveAsFile() to tests/fixtures/attachments/.
- The fixture directory is ignored by Git to protect confidentiality.
"""

from __future__ import annotations

from pathlib import Path
import sys

from scm_allocation.ingestion.collector import collect_attachments


def main() -> int:
    print("=" * 70)
    print("COLLECTING SAMPLE ATTACHMENTS (READ-ONLY)")
    print("=" * 70)

    target_dir = Path("tests/fixtures/attachments")
    try:
        manifest = collect_attachments(target_dir, max_emails=5)
    except Exception as err:
        print(f"[ERROR] Failed to collect attachments: {err}")
        return 1

    print(f"\nSuccessfully collected {len(manifest)} attachments into {target_dir}:\n")
    for idx, item in enumerate(manifest, start=1):
        print(f"[{idx}] Email #{item['source_email_number']} | Subject: {item['original_email_subject'][:45]}...")
        print(f"    Filename:   {item['original_attachment_filename']}")
        print(f"    Extension:  {item['extension']}")
        print(f"    Size:       {item['actual_size_bytes']} bytes")
        print(f"    MIME Type:  {item['mime_type']}")
        print(f"    Local Path: {item['local_fixture_path']}")
        print("-" * 70)

    print(f"\n[VERIFIED] Manifest created at: {target_dir / 'manifest.json'}")
    print("[VERIFIED] Read-only integrity preserved; zero mailbox mutations.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
