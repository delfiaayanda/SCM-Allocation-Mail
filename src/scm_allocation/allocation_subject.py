"""Satu definisi "email alokasi" dari subjek — dipakai saringan kandidat dan
jaring pengaman push APEX.

Latar (audit 28 Okt 2026): saringan lama mensyaratkan domain erajaya.com DAN
kata "alokasi"/"allocation", dan email yang isinya gagal diparse dibuang diam-diam.
Akibatnya ±100 thread alokasi tidak pernah tampil di APEX: permintaan dari akun
Gmail toko ("Req Alokasi Device iBox Senayan City"), subjek tanpa kata alokasi
("REQUEST STOCK LDU … NEW STORE", "Permohonan Fullfillment Stok"), serta
lampiran/gambar yang tidak terbaca. Sekarang subjek yang cocok di sini selalu
jadi kandidat, dan kalau isinya tetap tidak terbaca email itu tetap dikirim ke
APEX sebagai jenis 'tak_terbaca' untuk dikerjakan manual.
"""

from __future__ import annotations

import re

from scm_allocation.normalization import normalize_text


# Alokasi biasa, NPI, New Store, Event, permintaan/fulfilment stok toko, STO.
ALLOCATION_SUBJECT_RE = re.compile(
    r"alo[ck]a[st]i|allocation|\balloc\b"
    r"|\bnpi\b|new\s*store|grand\s*opening|\bopening\b|\bnso\b"
    r"|\bevent\b|bazaa?r|pameran|\bexpo\b"
    r"|\bre?q(?:uest|uset)?\.?\s+(?:stock|stok|barang|unit|ldu)"
    r"|permintaan\s+(?:stock|stok|barang|unit|ldu)|permohonan\s+(?:stock|stok|alokasi|fulfil|fullfil)"
    r"|kebutuhan\s+(?:stock|stok|unit|ldu)"
    r"|ful+fil+(?:l?ment)?|refill|replenish"
    r"|\binde?n(?:t)?\b"
    r"|\bpo\s+(?:stock|stok|khusus|barang|demo)"
    r"|\bsto\b",
    re.IGNORECASE,
)

# Subjek yang memang bukan permintaan alokasi walau mengandung kata di atas.
NON_ALLOCATION_SUBJECT_RE = re.compile(
    r"\bretur|\breturn|\bdelete\s+sto\b|\bcancel\s+sto\b|notifikasi|undeliverable|automatic\s+reply"
    r"|out\s+of\s+office|daftar\s+barang\s+masuk|pre-?alert|\bincoming\b|\brecall:",
    re.IGNORECASE,
)

# Kiriman mesin (laporan BY, upload SAP), bukan orang yang minta alokasi.
SYSTEM_SENDER_RE = re.compile(
    r"^(?:batchadm|by_support|noreply|no-reply|mailer-daemon|postmaster)@",
    re.IGNORECASE,
)


def is_allocation_subject(subject: object, sender: object = None) -> bool:
    """True kalau subjek email menandakan permintaan/kiriman alokasi."""
    text = normalize_text(subject) or ""
    if not text or not ALLOCATION_SUBJECT_RE.search(text) or NON_ALLOCATION_SUBJECT_RE.search(text):
        return False
    return not SYSTEM_SENDER_RE.search(str(sender or "").strip())


# Balasan: "RE:", "Re:", "Bls:" (bisa berulang). Terusan (FW/Fwd) TIDAK — itu
# biasanya permintaan toko yang diteruskan ke RICO.
REPLY_SUBJECT_RE = re.compile(r"^\s*(?:(?:re|bls|balas|aw)\s*:\s*)+", re.IGNORECASE)
# Balasan yang membawa permintaan sendiri: "tolong bantu request terlampir",
# "alokasi untuk article di bawah ini".
OWN_REQUEST_RE = re.compile(
    r"terlampir|attached|\blampir|di\s?bawah\s+ini|berikut\s+(?:data|list|detail|request|permintaan|kebutuhan)",
    re.IGNORECASE,
)
# Awal kutipan email sebelumnya di isi balasan (Outlook, Gmail, iOS, Android).
QUOTED_REPLY_RE = re.compile(
    r"^\s*(?:From|Dari|Sent|Dikirim)\s*:|^-{2,}\s*Original Message|^On .{5,200}wrote:"
    r"|^Pada .{5,200}menulis:|^_{8,}|^-{5,}",
    re.IGNORECASE | re.MULTILINE,
)
MATERIAL_CODE_RE = re.compile(r"\b8[01]\d{8}\b")
MEMO_SUBJECT_RE = re.compile(r"\bmemo\b", re.IGNORECASE)


def own_text(body: object) -> str:
    """Isi balasan yang ditulis pengirim sendiri — sebelum kutipan email lama."""
    text = str(body or "").replace("\r\n", "\n")
    match = QUOTED_REPLY_RE.search(text)
    return text[: match.start()] if match else text


def non_request_reason(subject: object, body: object, has_excel_attachment: bool) -> str | None:
    """Email bersubjek alokasi yang BUKAN permintaan baru — kalimat untuk APEX
    (masuk sebagai info), atau None kalau email ini memang perlu dikerjakan.

    Audit 28 Okt 2026: separuh "tak terbaca" (267 dari 511) adalah balasan dalam
    thread — "ok", "Berikut no STO 930…", "Approved", "stoknya kosong" — yang
    mengantre sebagai kerja manual padahal permintaannya ada di email awal.
    """
    text = normalize_text(subject) or ""
    if MEMO_SUBJECT_RE.search(text):
        return "Bot: memo (soft opening / internal), bukan permintaan alokasi."
    if not REPLY_SUBJECT_RE.match(text) or has_excel_attachment:
        return None
    body_text = str(body or "")
    if not QUOTED_REPLY_RE.search(body_text.replace("\r\n", "\n")):
        return None
    own = own_text(body_text)
    if MATERIAL_CODE_RE.search(own) or OWN_REQUEST_RE.search(own):
        return None
    return "Bot: balasan dalam thread alokasi tanpa permintaan baru (permintaannya ada di email awal thread)."
