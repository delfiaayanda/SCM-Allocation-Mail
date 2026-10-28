"""Push extracted AllocationRecord data to APEX's Ingest Alokasi PM API.

Bridges this tool's output to the existing `/api/submissions/from-outlook`
endpoint (apex-rico-accessories backend) instead of building a second data
store. Authenticated with a static machine key (same pattern APEX already
uses for its n8n ingest endpoints), not a user login — this is a bot, not a
person with an APEX account.
"""

from __future__ import annotations

import json
import os
from decimal import Decimal
from typing import Any
from urllib import request as urlrequest
from urllib.error import HTTPError, URLError

from scm_allocation.models.allocation import AllocationRecord


class ApexPushError(RuntimeError):
    """Raised when a push to APEX fails — network, auth, or server-side reject."""


def load_env_file(path: str) -> None:
    """Minimal KEY=VALUE loader — only sets vars not already in the environment.

    No new dependency (python-dotenv) for three config values; mirrors just
    enough of its behavior (skip blanks/comments, don't override real env).
    """
    if not os.path.isfile(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            if key and key not in os.environ:
                os.environ[key] = value.strip().strip('"').strip("'")


def _json_safe(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)
    if hasattr(value, "value"):     # enum (ExtractionStatus, AllocationContext)
        return value.value
    return value


def record_to_dict(record: AllocationRecord) -> dict[str, Any]:
    """AllocationRecord -> the field names APEX's OutlookAllocationRecord expects."""
    return {
        "material_code":   record.material_code,
        "plant_source":    record.issuing_warehouse_code,
        "plant_dest":      record.destination_plant_code,
        "qty":             _json_safe(record.quantity),
        "source_sloc":     record.issuing_warehouse_sloc,
        "dest_sloc":       record.destination_sloc,
        "context":         _json_safe(record.allocation_context),
        "confidence":      record.extraction_confidence,
        "parser_type":     record.parser_type,
        "source_filename": record.source_filename,
        "source_location": record.source_location,
    }


def fetch_pengirim_wajib(*, api_base: str, machine_key: str, cache_path: str,
                         timeout: float = 15.0) -> tuple[frozenset[str], str]:
    """Daftar email pengirim WAJIB (sheet MNP/CM -> PIC RICO di APEX) —
    email mereka diambil walau tidak berbau alokasi.

    Berhasil -> disimpan ke `cache_path`. Gagal (APEX mati, key salah) -> pakai
    cache terakhir; tanpa cache -> daftar kosong (perilaku lama: hanya email
    berbau alokasi). Return (daftar, keterangan sumber untuk log).
    """
    try:
        if not api_base or not machine_key:
            raise ApexPushError("APEX_API_BASE / APEX_MACHINE_KEY belum di-set")
        req = urlrequest.Request(
            url=f"{api_base.rstrip('/')}/api/submissions/pengirim-wajib",
            method="GET", headers={"x-machine-key": machine_key},
        )
        with urlrequest.urlopen(req, timeout=timeout) as resp:
            emails = json.loads(resp.read().decode("utf-8")).get("emails") or []
        daftar = frozenset(str(e).strip().casefold() for e in emails if str(e).strip())
        try:
            os.makedirs(os.path.dirname(cache_path) or ".", exist_ok=True)
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(sorted(daftar), f, ensure_ascii=False, indent=1)
        except OSError:
            pass
        return daftar, f"APEX ({len(daftar)} email)"
    except (ApexPushError, HTTPError, URLError, OSError, ValueError) as err:
        try:
            with open(cache_path, encoding="utf-8") as f:
                daftar = frozenset(str(e).strip().casefold() for e in json.load(f) if str(e).strip())
            return daftar, f"cache {cache_path} ({len(daftar)} email; APEX gagal: {err})"
        except (OSError, ValueError):
            return frozenset(), f"kosong (APEX gagal: {err}; belum ada cache)"


def push_records_to_apex(
    *, api_base: str, machine_key: str, source_id: int,
    email_id: str, sender: str | None, subject: str | None,
    records: list[AllocationRecord], timeout: float = 30.0,
    body_preview: str | None = None,
    tak_terbaca: str | None = None,
    catatan_bot: str | None = None,
) -> dict:
    """POST one email's records to /api/submissions/from-outlook.

    Raises ApexPushError on any failure (network, 4xx/5xx) — caller decides
    what that means for Outlook category writes (see process_outlook_allocations.py:
    a failed push must NOT be followed by a category write, or the email would
    be silently skipped on every future scan with its data never having landed).
    """
    body = json.dumps({
        "source_id": source_id,
        "email_id":  email_id,
        "sender":    sender,
        "subject":   subject,
        "records":   [record_to_dict(r) for r in records],
        # records kosong + body_preview = email info dari pengirim di daftar
        # wajib; APEX menolak records kosong dari pengirim lain — kecuali
        # tak_terbaca terisi (email alokasi yang isinya gagal diparse).
        "body_preview": body_preview,
        "tak_terbaca": tak_terbaca,
        # Email bersubjek alokasi yang bukan permintaan baru (balasan thread,
        # memo): records kosong, masuk APEX sebagai info dengan catatan ini.
        "catatan_bot": catatan_bot,
    }).encode("utf-8")

    req = urlrequest.Request(
        url=f"{api_base.rstrip('/')}/api/submissions/from-outlook",
        data=body,
        method="POST",
        headers={"Content-Type": "application/json", "x-machine-key": machine_key},
    )
    try:
        with urlrequest.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except HTTPError as err:
        detail = err.read().decode("utf-8", errors="replace")
        raise ApexPushError(f"HTTP {err.code} dari APEX: {detail}") from err
    except URLError as err:
        raise ApexPushError(f"Tidak bisa menghubungi APEX ({api_base}): {err.reason}") from err
