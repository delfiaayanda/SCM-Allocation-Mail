"""Audit akurasi 28 Okt 2026: hasil parse dicocokkan ke STO SAP.

Tiap test mewakili pola nyata yang dulu terbaca SALAH (material kosong, sheet
dump stok jadi alokasi, toko tujuan tertukar dengan gudang asal) atau masuk
antrean "tak terbaca" padahal bisa dibaca / memang bukan permintaan.
"""

from __future__ import annotations

import openpyxl

from scm_allocation.allocation_subject import non_request_reason
from scm_allocation.extraction import extract_email
from scm_allocation.models.allocation import ExtractionResult, ExtractionStatus
from scm_allocation.parsers.excel import parse_excel
from scm_allocation.parsers.html import parse_html
from scm_allocation.parsers.text import parse_text_request
from scm_allocation.processing import INFO_CATEGORY, CandidateFilter, OutlookEmailScanner, OutlookScanConfig

from tests.unit.test_cakupan_email_alokasi import _Item, _Outlook, _ScanItem, _sig, _xlsx


# ── Excel ─────────────────────────────────────────────────────────────────────

def test_alokasi_apple_site_allocation_membaca_article_code(tmp_path):
    path = _xlsx(tmp_path / "Alokasi Accs.xlsx", {"Batch 1": [
        ["Article Code", "Article Desc", "Site Code", "Site Desc", "Sales", "Stock", "Allocation"],
        [8100227168, "APP AIRPODS PRO (3RD GEN)", "U008", "UR SUMMARECON MALL", 5, 0, 2],
        [8100227168, "APP AIRPODS PRO (3RD GEN)", "U047", "UR BINTARO", 2, 2, None],
    ]})
    result = parse_excel(path)
    assert result.status is ExtractionStatus.VALID
    assert _sig(result.records) == [("8100227168", 2, "U008", None)]


def test_sheet_tersembunyi_dan_master_tidak_mengalahkan_sheet_permintaan(tmp_path):
    header = ["Area", "City", "MATERIAL", "BRAND TYPE 2", "STOCK", "SALES", "PO", "BGC1"]
    row = ["KALIMANTAN 5", "ERAFONE ALL SAMARINDA", 8100122759, "iPhone 15 128GB", 0, 4, 5, 17]
    path = _xlsx(tmp_path / "PO STOCK.xlsx", {
        "DATABASE": [["plant", "material", "Qty"], ["E018", 8100018451, 1], ["E018", 8100018531, 1]],
        "MASTER": [header, row],
        "APPLE": [header, row],
    })
    wb = openpyxl.load_workbook(path)
    wb["DATABASE"].sheet_state = "hidden"
    wb.save(path)
    result = parse_excel(path)
    assert [(r.material_code, r.quantity, r.source_sheet) for r in result.records] == [("8100122759", 5, "APPLE")]


def test_isi_email_menunjuk_sheet(tmp_path):
    header = ["Material", "Description", "Store Code", "Qty"]
    path = _xlsx(tmp_path / "Alokasi.xlsx", {
        "Draft": [header, [8100000001, "A", "E101", 9]],
        "Final": [header, [8100000001, "A", "E101", 3]],
    })
    result = parse_excel(path, body_text="Dear tim, mohon alokasi sesuai sheet Final ya")
    assert _sig(result.records) == [("8100000001", 3, "E101", None)]


# ── HTML ──────────────────────────────────────────────────────────────────────

def test_html_toko_tujuan_di_judul_kolom_qty():
    html = ("<table><tr><td>Material</td><td>Desc</td><td>Stock BGC1</td><td>Inden M215</td></tr>"
            "<tr><td>8100267705</td><td>DJI Lito X1</td><td>5</td><td>1</td></tr></table>")
    assert _sig(parse_html(html).records) == [("8100267705", 1, "M215", "BGC1")]


def test_html_judul_dan_data_terpecah_dua_tabel():
    html = ("<table><tr><td>Article Code</td><td>Article Description</td><td>BF02/BFC1</td><td>Alokasi XU06</td></tr></table>"
            "<table><tr><td>8100016597</td><td>KIC JETFAN MINI BLUE</td><td>BFC1</td><td>2</td></tr>"
            "<tr><td>8100114619</td><td>KIC Jet Fan 2 Turbo</td><td>BFC1</td><td>4</td></tr></table>")
    assert _sig(parse_html(html).records) == [("8100016597", 2, "XU06", "BFC1"), ("8100114619", 4, "XU06", "BFC1")]


# ── Teks isi email ────────────────────────────────────────────────────────────

def test_permintaan_berbentuk_teks_berlabel():
    body = ("Dear Team,\r\nMohon dibantu pembuatan sto:\r\nMaterial : 8100261708\r\n"
            "Item Desc : MBA 15 M5\r\nQty : 1unit\r\nThanks & Regards.\r\nElsa\r\n0812 3456 7890")
    records = parse_text_request(body, "PEMBUATAN STO IBOX PATTIMURA PADANG X147", source_email_id="e")
    assert _sig(records) == [("8100261708", 1, "X147", None)]
    assert records[0].material_description == "MBA 15 M5"
    assert records[0].extraction_status is ExtractionStatus.PARTIAL


def test_teks_tanpa_qty_eksplisit_tidak_ditebak():
    assert parse_text_request("Mohon cek stok 8100261708 di gudang ya", "Alokasi X147", source_email_id="e") == []


def test_extract_email_jatuh_ke_teks_kalau_tabel_kosong():
    email = {"entry_id": "e", "subject": "RE: INDEN STOK SPEAKER JBL", "attachments": [], "body_html": "<p>x</p>",
             "body_plain_text": "Dear Delfia,\nMohon indent:\n8100225999 JBL PartyBox Club 120\nQty: 1 unit\nE807 Erafone ruko negara\n"
                                "Thanks\nFrom: Farel\nSent: Monday"}
    result = extract_email(email)
    assert result.parser_type == "text_body" and _sig(result.records) == [("8100225999", 1, "E807", None)]


# ── Bukan permintaan baru → info ─────────────────────────────────────────────

def test_balasan_thread_tanpa_isi_baru_jadi_info():
    body = "ok\r\n\r\nFrom: Kevin <kevin@erajaya.com>\r\nSent: Monday\r\nMohon alokasi 8100000001 qty 2"
    assert non_request_reason("RE: ALOKASI TECNO 26 OCT", body, False)
    assert non_request_reason("Re: Memo Soft Opening New Store MI-STORE SUMEDANG (N462)", "Oke", False)


def test_balasan_yang_membawa_permintaan_tetap_dikerjakan():
    quoted = "\r\nFrom: A\r\nSent: B\r\nisi lama"
    assert non_request_reason("RE: Indent Accs Tab S10", "Mohon proses 8100019995 qty 1" + quoted, False) is None
    assert non_request_reason("Re: Request Alokasi iPhone", "Tolong bantu request terlampir" + quoted, False) is None
    assert non_request_reason("FW: Request Alokasi iPhone", "ok" + quoted, False) is None
    assert non_request_reason("RE: Request Alokasi iPhone", "ok" + quoted, True) is None


def test_scanner_balasan_thread_dikirim_sebagai_info():
    item = _ScanItem("RE: ALOKASI TECNO 26 OCT")
    item.Body = "ok\r\nFrom: Kevin <kevin@erajaya.com>\r\nSent: Monday"
    scanner = OutlookEmailScanner(
        config=OutlookScanConfig(dry_run=True),
        extractor=lambda email, **_: ExtractionResult([], ExtractionStatus.INVALID, "html_table", "id-1",
                                                      ("no supported allocation table found",)),
    )
    [summary] = scanner.scan(_Outlook(item))
    assert summary.category == INFO_CATEGORY
    assert "balasan" in summary.metadata["catatan_bot"] and "tak_terbaca" not in summary.metadata


def test_email_retur_bukan_kandidat_walau_isinya_menyebut_alokasi():
    item = _Item("Request Returan Terpusat W39 2026", "pm@erajaya.com")
    item.Body = "Mohon retur barang sisa alokasi event"
    assert not CandidateFilter().matches(item)


def test_sheet_tersembunyi_berisi_alokasi_per_toko_tetap_dibaca(tmp_path):
    path = _xlsx(tmp_path / "Alokasi GO.xlsx", {
        "RICO": [["Material", "Description", "Qty"], [8100002156, "Fit3", 16]],
        "GAP": [["Material", "Description", "Store Code", "Alokasi"], [8100002155, "Fit3", "F218", 3]],
    })
    wb = openpyxl.load_workbook(path)
    wb["GAP"].sheet_state = "hidden"
    wb.save(path)
    assert ("8100002155", 3, "F218", None) in _sig(parse_excel(path).records)


def test_sheet_tersembunyi_bulan_lama_dibuang(tmp_path):
    header = ["Material", "Description", "Store Code", "Qty"]
    path = _xlsx(tmp_path / "STO NEW STORE.xlsx", {
        "STO MARSHALL APRIL": [header, [8100028957, "Emberton", "X109", 6]],
        "MHL": [header, [8100028957, "Emberton", "X253", 1]],
    })
    wb = openpyxl.load_workbook(path)
    wb["STO MARSHALL APRIL"].sheet_state = "hidden"
    wb.save(path)
    assert _sig(parse_excel(path, received="2026-09-18 10:00:00").records) == [("8100028957", 1, "X253", None)]


def test_toko_di_judul_kolom_kalah_dari_toko_di_subjek(tmp_path):
    path = _xlsx(tmp_path / "Request Alokasi Olike - M183.xlsx", {"Sheet1": [
        ["Article Code", "Article Description", "Alokasi N183"], [8100233169, "OLI CAR CHARGER R8", 7]]})
    email = {"entry_id": "e", "subject": "Request Alokasi: Refill Stock - Olike M183 (9 Oct 2026)",
             "body_plain_text": "", "attachments": [{"filename": path.name}]}
    assert _sig(extract_email(email, attachment_paths={path.name: path}).records) == [("8100233169", 7, "M183", None)]
