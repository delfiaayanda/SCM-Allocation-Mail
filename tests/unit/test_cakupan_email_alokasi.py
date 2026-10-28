"""Audit cakupan 28 Okt 2026: email alokasi tidak boleh terlewat.

Tiap test di sini mewakili satu pola nyata yang dulu membuat email alokasi
dibuang diam-diam (data sintetis, bentuknya sama dengan lampiran aslinya).
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from scm_allocation.allocation_subject import is_allocation_subject
from scm_allocation.extraction import extract_email, is_allocation_attachment_filename
from scm_allocation.models.allocation import ExtractionResult, ExtractionStatus
from scm_allocation.parsers.excel import parse_excel
from scm_allocation.parsers.generic import extract_generic
from scm_allocation.parsers.html import parse_html, parse_html_tables
from scm_allocation.processing import UNREADABLE_CATEGORY, CandidateFilter, OutlookEmailScanner, OutlookScanConfig


def _xlsx(path: Path, sheets: dict[str, list[list[object]]]) -> Path:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return path


def _generic(rows, plant=None):
    return extract_generic(rows, source_email_id="e", source_filename=None, source_sheet=None,
                           destination_plant_code=plant, destination_plant_description=None, location_prefix="")


def _sig(records):
    return [(r.material_code, r.quantity, r.destination_plant_code, r.issuing_warehouse_code) for r in records]


# ── Definisi email alokasi dari subjek ────────────────────────────────────────

@pytest.mark.parametrize("subject", [
    "Req Alokasi Device iBox Senayan City Week 44",
    "REQUEST STOCK LDU DAN DEVICE XIAOMI NEW STORE ERAFONE OPTIMIZED RUKO PINANG",
    "PERMOHONAN FULLFILLMENT STOK REALME STORE ERAFONE WLINGI BLITAR (F141)",
    "Huawei NPI Freebuds Neo 1st Alloc 06 Oct",
    "Draft Alokasi Event UR - Mandiri Livin Markt (XU01)",
    "PERMINTAAN ALOKASI STOK DARI DC LOMBOK (BG39) KE ERAFONE SOEKARNO HATTA",
    "PO KHUSUS CUSTOMER VIP IBOX A YANI SERANG X142",
])
def test_subjek_alokasi_dikenali(subject):
    assert is_allocation_subject(subject, "toko@gmail.com")


@pytest.mark.parametrize("subject,sender", [
    ("Retur Terpusat Onpoint & Jisulife - Return to BFC1", "pm@erajaya.com"),
    ("Daftar barang masuk (27.10.2026)(Accs)", "pm@erajaya.com"),
    ("Notifikasi PO/STO SAP from BY", "batchadm@aws.erajaya.com"),
    ("Erajaya BY DNF PROD | UDT Upload STO Outbound", "BY_Support@erajaya.com"),
    ("Weekly team announcement", "pm@erajaya.com"),
])
def test_subjek_bukan_alokasi(subject, sender):
    assert not is_allocation_subject(subject, sender)


class _Item:
    def __init__(self, subject, sender):
        self.Subject, self.SenderEmailType, self.SenderEmailAddress, self.SenderName = subject, "SMTP", sender, "x"
        self.Body = self.HTMLBody = ""


def test_akun_gmail_toko_dengan_subjek_alokasi_jadi_kandidat():
    assert CandidateFilter().matches(_Item("Req Alokasi Device iBox Senayan City Week 44", "iboxsc@gmail.com"))
    assert not CandidateFilter().matches(_Item("Promo akhir pekan", "iboxsc@gmail.com"))


# ── Lampiran ──────────────────────────────────────────────────────────────────

def test_nama_lampiran_berspasi_ganda_tetap_ketemu(tmp_path):
    raw = "STO  Alokasi - EVENT BAZAAR (XU11).xlsx"
    path = _xlsx(tmp_path / "a.xlsx", {"Sheet1": [["Article Code", "Article Description", "Qty", "Site Code"],
                                                  [8100000001, "Barang A", 3, "XU11"]]})
    email = {"entry_id": "e", "subject": "Draft Alokasi Event", "attachments": [{"filename": raw}]}
    result = extract_email(email, attachment_paths={raw: path})
    assert result.records, result.errors
    assert not any("unavailable" in e for e in result.errors)


@pytest.mark.parametrize("name", ["a.xlsx", "a.XLSX", "a.xlsm", "a.xlsb", "a.xls", "a.csv"])
def test_format_lampiran_excel_lain_didukung(name):
    assert is_allocation_attachment_filename(name)


def test_csv_terbaca(tmp_path):
    path = tmp_path / "req.csv"
    path.write_text("Material;Material Desc;Plant;Qty\n8100000001;Barang A;E013;4\n", encoding="utf-8")
    assert _sig(parse_excel(path).records) == [("8100000001", 4, "E013", None)]


def test_sheet_referensi_stok_dibuang_kalau_ada_sheet_permintaan(tmp_path):
    path = _xlsx(tmp_path / "po.xlsx", {
        "Sales Stock": [["Plant", "Article", "Article Description", "Quantity"]]
                       + [["E272", 8100000000 + i, f"Barang {i}", 5] for i in range(50)],
        "PO Device": [["Article", "Article Description", "E272"], [8100009999, "Barang Z", 2]],
    })
    records = parse_excel(path).records
    assert _sig(records) == [("8100009999", 2, "E272", None)]


def test_workbook_besar_tidak_menggantung(tmp_path):
    import time
    rows = [["Plant", "Article", "Article Description", "Quantity", "Brand", "CAT", "X", "Y"]]
    rows += [["E272", 8100000000 + i, "Barang", i % 7, "B", "C", 1, 2] for i in range(15000)]
    path = _xlsx(tmp_path / "big.xlsx", {"Data": rows})
    start = time.monotonic()
    parse_excel(path)
    assert time.monotonic() - start < 60


# ── Parser generik ────────────────────────────────────────────────────────────

def test_tabel_panjang_header_tidak_baku():
    rows = [["DC", "Plant", "Store", "Mat Code", "Mat Desc", "Qty"],
            ["BGC1", "F259", "ERAFONE AR HAKIM", 8100033996, "LMA TG", 3],
            ["BGC1", "F259", "ERAFONE AR HAKIM", 8100033994, "LMA TG 2", 5],
            ["Grand Total", None, None, None, None, 8]]
    assert _sig(_generic(rows)) == [("8100033996", 3, "F259", "BGC1"), ("8100033994", 5, "F259", "BGC1")]


def test_kolom_request_stock_dan_kolom_pendukung():
    rows = [["PLANT", "Site Desc", "Type 2", "SALES", "STOCK", "DOS", "Request stock"],
            ["E013", "ERAFONE SUMMARECON", "Realme 16P 12/256GB", 8, 0, 0, 10]]
    records = _generic(rows)
    assert [(r.material_description, r.quantity, r.destination_plant_code) for r in records] == [("Realme 16P 12/256GB", 10, "E013")]


def test_dua_kolom_plant_gudang_jadi_asal():
    rows = [["Plant", "Name 1", "Article", "Article Description", "Qty", "Plant", "Name 1"],
            ["BGC1", "MDC CIKUPA", 8100226119, "VIVO V60", 205, "M018", "ERAFONE AND MORE"]]
    assert _sig(_generic(rows)) == [("8100226119", 205, "M018", "BGC1")]


def test_matriks_kode_toko_sebagai_judul_kolom():
    rows = [["article_code", "article_description", "Total", "X219", "X217", "BIC1"],
            ["", "", "", "IBOX A", "IBOX B", ""],
            [8100196951, "CMM TG", 45, 35, 10, 999]]
    assert _sig(_generic(rows)) == [("8100196951", 35, "X219", "BIC1"), ("8100196951", 10, "X217", "BIC1")]


def test_matriks_kode_material_di_atas_header():
    rows = [["", "", "", 8100302039, 8100302037],
            ["Era Code", "Region", "Store Name", "Freebuds Neo Blue", "Freebuds Neo White"],
            ["N128", "Jakarta 1", "HES A", 9, 22],
            ["N119", "Jakarta 1", "HES B", 0, 20]]
    assert _sig(_generic(rows)) == [("8100302039", 9, "N128", None), ("8100302037", 22, "N128", None),
                                    ("8100302037", 20, "N119", None)]


def test_tabel_tanpa_header_pakai_toko_dari_subjek():
    rows = [["8100304352", "SKZ S610 OpenRun", "3"], ["8100304353", "SKZ S611 OpenRun", "2"]]
    assert _sig(_generic(rows, plant="izd1")) == [("8100304352", 3, "IZD1", None), ("8100304353", 2, "IZD1", None)]


def test_blok_per_toko_dengan_header_berulang():
    rows = [[None, "Material", "Material desc", "Brand Name", "E595", "BGC1", "PO", "REMAKS"],
            [None, 8100226818, "iPhone 17 Pro", "APPLE", 0, 2, 1, "Stock langsung"],
            [None, None, None, None, None, None, None, None],
            [None, "Material", "Material desc", "Brand Name", "F208", "BGC1", "PO", "REMAKS"],
            [None, 8100226814, "iPhone 17 Pro Max", "APPLE", 0, 18, 1, "Need support"]]
    assert _sig(_generic(rows)) == [("8100226818", 1, "E595", "BGC1"), ("8100226814", 1, "F208", "BGC1")]


def test_kolom_qty_baru_terisi_jauh_di_bawah():
    rows = [["Material", "Material descr", "NASA MDC CIKUPA", "SES SENAYAN CITY", "REQ"]]
    rows += [[8100000000 + i, f"Barang {i}", 5, 1, None] for i in range(300)]
    rows += [[8100099999, "Barang diminta", 5, 0, 4]]
    assert _sig(_generic(rows, plant="S018")) == [("8100099999", 4, "S018", None)]


def test_pivot_sum_of_qty_kalah_dari_sheet_req(tmp_path):
    # Pivot stok gudang ("Sum of QTY") di samping sheet permintaan ber-kolom REQ:
    # yang diambil hanya sheet REQ. Baris sesudah "Total" pivot juga tidak boleh
    # dibaca sebagai tabel tanpa header.
    path = _xlsx(tmp_path / "req.xlsx", {
        "Sheet2": [[None] * 5, ["plant_name", "CAT", "material", "material_descr", "Sum of QTY"],
                   ["DCM DC ROXY", "AW S11", 8100227179, "APP WATCH 11", 3],
                   [None, None, 8100227180, "APP WATCH 11 B", 1],
                   ["DCM DC ROXY Total", None, None, None, 4],
                   ["DCM DC CIKUPA", "AW SE", 8100227169, "APP WATCH SE", 48]],
        "Sheet3": [["MATERIAL", "ITEM DESC", "STOCK", "SALES", "DOS", "REQ"],
                   [8100227169, "APP WATCH SE", 2, 5, 12, 3]],
    })
    assert _sig(parse_excel(path, destination_plant_code="X021").records) == [("8100227169", 3, "X021", None)]


def test_pivot_sum_of_quantity_sendirian_tetap_terbaca():
    rows = [[None, "Article Code", "Article Description", "Sum of Quantity"],
            [None, 8100254863, "VKM ISTAR", 3], [None, 8100254864, "VKM ITWILL", 2]]
    assert _sig(_generic(rows, plant="X250")) == [("8100254863", 3, "X250", None), ("8100254864", 2, "X250", None)]


def test_tabel_ringkas_kolom_total_atau_ket():
    rows = [["Article Code", "Article Descr", "Total"], ["8100147072", "DJI Pocket 3", "2"], ["8100228761", "DJI Neo 2", "3"]]
    assert _sig(_generic(rows, plant="E530")) == [("8100147072", 2, "E530", None), ("8100228761", 3, "E530", None)]


def test_laporan_stok_berkolom_angka_banyak_tidak_jadi_alokasi():
    rows = [["MATERIAL", "ITEM DESCRIPTION", "STOK", "SALES", "DOS"], ["8100268388", "REALME C100I", "1", "2", "15"]]
    assert _generic(rows, plant="M094") == []


def test_workbook_log_berjalan_ambil_sheet_tanggal_subjek(tmp_path):
    sheet = lambda qty: [["DC", "BU Code", "MATERIAL CODE", "MATERIAL DESC", "ALOKASI"], ["BGC1", "E013", 8100000001, "Gel", qty]]
    path = _xlsx(tmp_path / "lamina.xlsx", {"140126": sheet(1), "121026": sheet(2), "121026 onhold": sheet(3), "211026": sheet(4)})
    assert [r.quantity for r in parse_excel(path, subject="Alokasi LAMINA - 211026").records] == [4]
    assert [r.quantity for r in parse_excel(path, subject="Alokasi LAMINA").records] == [4]
    assert sorted(r.quantity for r in parse_excel(path, subject="Alokasi LAMINA - 121026").records) == [2, 3]


def test_laporan_stok_tanpa_kolom_permintaan_tidak_jadi_alokasi():
    rows = [["Material", "Material description", "Plant", "Unrestricted", "Actual Inventory"],
            [2200000082, "Tas", "BHC1", 46, 46]]
    assert _generic(rows) == []


# ── HTML ──────────────────────────────────────────────────────────────────────

def test_html_rowspan_tidak_menggeser_kolom():
    html = """<table>
      <tr><td>Article</td><td>Article Description</td><td>Plant</td><td>STOCK</td><td>STORE CODE</td><td>REQ</td></tr>
      <tr><td>8100122787</td><td>iPhone 15 Plus</td><td rowspan=2>BIC1</td><td>10</td><td rowspan=2>X249</td><td>1</td></tr>
      <tr><td>8100122784</td><td>iPhone 15 Plus B</td><td>7</td><td>2</td></tr>
    </table>"""
    assert _sig(parse_html(html).records) == [("8100122787", 1, "X249", "BIC1"), ("8100122784", 2, "X249", "BIC1")]


def test_html_tabel_bersarang_tidak_hilang():
    html = ("<table><tr><td>Halo tim</td></tr><tr><td><table>"
            "<tr><td>SUPPLY</td><td>SITE</td><td>MATERIAL</td><td>ARTICLE DESC</td><td>REQ</td></tr>"
            "<tr><td>BIC1</td><td>X069</td><td>8100102404</td><td>iPhone 16 Plus</td><td>2</td></tr>"
            "</table></td></tr><tr><td>Terima kasih</td></tr></table>")
    assert len(parse_html_tables(html)) == 2
    assert _sig(parse_html(html).records) == [("8100102404", 2, "X069", "BIC1")]


def test_html_pivot_dua_baris_tidak_crash():
    html = "<table><tr><td>Article Code</td><td>Article Description</td><td>Request</td></tr><tr><td>X</td></tr></table>"
    assert parse_html(html).status is ExtractionStatus.INVALID


# ── Jaring pengaman: tetap sampai ke APEX ─────────────────────────────────────

class _ScanItem(_Item):
    def __init__(self, subject):
        super().__init__(subject, "store@erajaya.com")
        self.EntryID, self.ReceivedTime, self.UnRead, self.Categories = "id-1", "2026-10-28", True, ""
        self.Attachments = []
        self.Body = "Mohon dibantu stok untuk new store, terlampir foto rak."


class _Items(list):
    Count = property(lambda self: len(self))

    def Item(self, i):
        return self[i - 1]

    def Sort(self, *_):
        pass


class _Outlook:
    def __init__(self, item):
        folder = type("F", (), {"Items": _Items([item])})()
        self.ns = type("N", (), {"GetDefaultFolder": lambda s, i: folder, "Categories": None})()

    def GetNamespace(self, _):
        return self.ns


def test_email_alokasi_tak_terbaca_dikirim_dengan_alasan():
    item = _ScanItem("REQUEST STOCK LDU NEW STORE ERAFONE PINANG RANTI")
    scanner = OutlookEmailScanner(
        config=OutlookScanConfig(dry_run=True),
        extractor=lambda email, **_: ExtractionResult([], ExtractionStatus.INVALID, "html_table", "id-1",
                                                      ("no supported allocation table found",)),
    )
    [summary] = scanner.scan(_Outlook(item))
    assert summary.category == UNREADABLE_CATEGORY
    assert "tabel alokasi" in summary.metadata["tak_terbaca"]
    assert "new store" in summary.metadata["body_preview"]
