"""Ringkasan antrian review — cek cepat email mana yang butuh perhatian.

`allocation_bot_history.csv` cuma menyimpan email yang berhasil ditandai
kategori (VALID -> "SCM Bot - Scraped", PARTIAL -> "SCM Bot - Review").
INVALID/UNSUPPORTED tidak dapat kategori jadi tidak pernah masuk history.csv
(lihat README "Local audit files") — makanya angka invalid di sini cuma level
agregat per-run dari runs.csv, bukan per-email. Kalau butuh detail per-email
untuk INVALID, jalankan ulang scanner-nya dan baca output console (diagnostik
per-worksheet dicetak di sana, tidak disimpan ke file manapun).

Pemakaian:
    .\\.venv\\Scripts\\python.exe scripts\\review_queue_summary.py
    .\\.venv\\Scripts\\python.exe scripts\\review_queue_summary.py --days 7 --limit 20
"""

from __future__ import annotations

import argparse
import csv
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path


def _read_csv(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def _parse_ts(value: str) -> datetime | None:
    try:
        return datetime.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Ringkasan antrian review SCM allocation bot.")
    parser.add_argument("--history", default="logs/allocation_bot_history.csv")
    parser.add_argument("--runs", default="logs/allocation_bot_runs.csv")
    parser.add_argument("--days", type=int, default=14, help="Cuma tampilkan N hari terakhir (default 14)")
    parser.add_argument("--limit", type=int, default=15, help="Maks item ditampilkan per grup parser (default 15)")
    args = parser.parse_args()

    cutoff = datetime.now(timezone.utc) - timedelta(days=args.days)

    history = _read_csv(Path(args.history))
    runs = _read_csv(Path(args.runs))

    print("=" * 72)
    print(f"TREN PER-RUN (runs.csv, {len(runs)} run tercatat)")
    print("=" * 72)
    if not runs:
        print("(belum ada run tercatat)")
    for r in runs:
        ts = _parse_ts(r.get("timestamp", ""))
        if ts and ts < cutoff:
            continue
        candidate = int(r.get("candidate_count") or 0)
        review = int(r.get("review_count") or 0)
        invalid = int(r.get("invalid_or_excluded_count") or 0)
        scraped = int(r.get("scraped_count") or 0)
        ratio = f"{(review / candidate * 100):.0f}%" if candidate else "-"
        flag = "  ⚠ review > 15% kandidat" if candidate and review / candidate > 0.15 else ""
        print(
            f"{r.get('timestamp','')[:19]}  folder={r.get('folder')}  "
            f"scanned={r.get('scanned_count')} candidate={candidate}  "
            f"scraped={scraped} review={review} ({ratio}) invalid/excluded={invalid}{flag}"
        )

    partial = [
        h for h in history
        if h.get("extraction_status", "").strip().lower() == "partial"
        and (_parse_ts(h.get("timestamp", "")) or cutoff) >= cutoff
    ]
    partial.sort(key=lambda h: h.get("timestamp", ""), reverse=True)

    print()
    print("=" * 72)
    print(f"ANTRIAN REVIEW — {len(partial)} email PARTIAL dalam {args.days} hari terakhir")
    print("(dikelompokkan per parser — parser yang sering PARTIAL = kandidat kuat")
    print(" buat ditambah alias/kasus baru di normalization/fields.py atau parsers/)")
    print("=" * 72)

    by_parser: dict[str, list[dict]] = defaultdict(list)
    for h in partial:
        by_parser[h.get("parser_type") or "(tidak diketahui)"].append(h)

    for parser_type, items in sorted(by_parser.items(), key=lambda kv: -len(kv[1])):
        print(f"\n--- {parser_type}  ({len(items)} email) ---")
        for h in items[: args.limit]:
            reason = h.get("reason_error") or "(tidak ada detail — lihat console log run terkait)"
            print(
                f"  [{h.get('timestamp','')[:19]}] {h.get('subject','')!r}"
                f"\n      dari={h.get('sender','')} records={h.get('record_count','?')} "
                f"| alasan: {reason}"
            )
        if len(items) > args.limit:
            print(f"  ... {len(items) - args.limit} lagi (naikkan --limit untuk lihat semua)")

    if not partial:
        print("(kosong — tidak ada email PARTIAL dalam rentang ini)")

    # invalid/excluded cuma agregat, tidak ada rincian per-email di CSV manapun
    recent_runs = [r for r in runs if (_parse_ts(r.get("timestamp", "")) or cutoff) >= cutoff]
    total_invalid = sum(int(r.get("invalid_or_excluded_count") or 0) for r in recent_runs)
    if total_invalid:
        print()
        print(
            f"Catatan: {total_invalid} email berstatus invalid/excluded dalam {args.days} hari "
            "terakhir tidak punya rincian per-email tersimpan (bukan email alokasi, atau format "
            "gagal total — dua-duanya tidak dibedakan di log). Kalau perlu rincian, jalankan ulang "
            "scanner-nya dan simpan output console-nya, mis.:\n"
            "  .\\.venv\\Scripts\\python.exe scripts\\process_outlook_allocations.py "
            "--folder Inbox --preview-records | Tee-Object logs\\scan_console.log"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
