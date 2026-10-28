"""Pengirim di daftar wajib APEX (sheet MNP/CM -> PIC RICO): emailnya diambil
walau tidak berbau alokasi, dan yang tanpa alokasi terbaca jadi email info."""

from __future__ import annotations

import json

from scm_allocation.models.allocation import ExtractionResult, ExtractionStatus
from scm_allocation.output import apex
from scm_allocation.processing import (
    INFO_CATEGORY,
    SCRAPED_CATEGORY,
    CandidateFilter,
    OutlookEmailScanner,
    OutlookScanConfig,
)

from tests.unit.test_outlook_processing import FakeFolder, FakeItem, FakeOutlook, result

WAJIB = frozenset({"mnp.uji@erajaya.com"})


def _item(sender: str, subject: str = "Info stok minggu ini", body: str = "Halo,\r\n\r\n\r\n\r\nMohon dicek   stok X071.") -> FakeItem:
    item = FakeItem(subject=subject, categories="")
    item.SenderEmailAddress = sender
    item.Body = body
    item.HTMLBody = ""
    return item


def _scanner(item: FakeItem, status: ExtractionStatus, *, records: list | None = None, dry_run: bool = True):
    hasil = result(status)
    if records:
        hasil = ExtractionResult(records, status, hasil.parser_type, hasil.source_email_id, hasil.errors)
    scanner = OutlookEmailScanner(
        config=OutlookScanConfig(dry_run=dry_run, max_emails=1,
                                 candidate_filter=CandidateFilter(wajib_senders=WAJIB)),
        extractor=lambda email, **kw: hasil,
    )
    return scanner, FakeOutlook(FakeFolder(item))


def test_pengirim_wajib_lolos_tanpa_kata_alokasi_dan_beda_domain():
    f = CandidateFilter(wajib_senders=frozenset({"vendor@gmail.com", *WAJIB}))
    assert f.matches(_item("MNP.Uji@erajaya.com"))
    assert f.matches(_item("vendor@gmail.com"))
    assert not f.matches(_item("orang.lain@erajaya.com"))           # bukan wajib, tanpa "alokasi"
    assert f.matches(_item("orang.lain@erajaya.com", subject="Alokasi X071"))
    assert not f.matches(_item("MNP.Uji@erajaya.com", subject="Alokasi STO Batam"))  # excluded tetap menang


def test_email_wajib_tanpa_rekaman_jadi_info_dengan_isi():
    scanner, outlook = _scanner(_item("mnp.uji@erajaya.com"), ExtractionStatus.INVALID)
    [summary] = scanner.scan(outlook)
    assert summary.category == INFO_CATEGORY
    assert summary.metadata["wajib"] is True
    assert summary.metadata["body_preview"] == "Halo,\n\nMohon dicek stok X071."


def test_email_wajib_berrekaman_tetap_alokasi_biasa():
    scanner, outlook = _scanner(_item("mnp.uji@erajaya.com"), ExtractionStatus.VALID, records=["rekaman"])
    [summary] = scanner.scan(outlook)
    assert summary.category == SCRAPED_CATEGORY and "wajib" not in summary.metadata


def test_email_berkategori_info_tidak_diproses_ulang():
    item = _item("mnp.uji@erajaya.com")
    item.Categories = INFO_CATEGORY
    scanner, outlook = _scanner(item, ExtractionStatus.INVALID)
    [summary] = scanner.scan(outlook)
    assert summary.skipped


def test_fetch_pengirim_wajib_jatuh_ke_cache(tmp_path, monkeypatch):
    cache = tmp_path / "pengirim_wajib.json"

    class Resp:
        def __enter__(self): return self
        def __exit__(self, *a): return False
        def read(self): return json.dumps({"emails": ["MNP.Uji@erajaya.com"]}).encode()

    monkeypatch.setattr(apex.urlrequest, "urlopen", lambda req, timeout: Resp())
    daftar, _ = apex.fetch_pengirim_wajib(api_base="http://x", machine_key="k", cache_path=str(cache))
    assert daftar == WAJIB and json.loads(cache.read_text()) == ["mnp.uji@erajaya.com"]

    def gagal(req, timeout):
        raise apex.URLError("mati")
    monkeypatch.setattr(apex.urlrequest, "urlopen", gagal)
    daftar, sumber = apex.fetch_pengirim_wajib(api_base="http://x", machine_key="k", cache_path=str(cache))
    assert daftar == WAJIB and "cache" in sumber

    daftar, sumber = apex.fetch_pengirim_wajib(api_base="http://x", machine_key="k",
                                               cache_path=str(tmp_path / "tidak-ada.json"))
    assert daftar == frozenset() and "kosong" in sumber
