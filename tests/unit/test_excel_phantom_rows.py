"""Regresi 14 Okt 2026: sheet yang diformat sampai ~1 juta baris membuat parse_excel
berputar kuadratik (scan Outlook tiap jam menggantung, 44 proses menumpuk)."""
from __future__ import annotations

import time

import openpyxl
from openpyxl.styles import Font

from scm_allocation.parsers.excel import _trim_phantom_cells, parse_excel

HEADER = ["GAB", "SISTEM", "BRAND", "FROM", "DEST. CODE", "DESTINASI DESC", "REGION", "SKU",
          "Item Deskripsi", "Type", "SLOC", "Qty", "Area", "STO", "Keterangan", "Check", "Tanggal"]


def _book(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(HEADER)
    ws.append(["BC588100249939", 1010, "TECNO", "BCC1", "BC58", "TAM DC PALANGKARAYA", "REGION 4",
               8100249939, "Tecno Spark Go 3 4/64 Grey", "Tecno Spark Go 3 4/64", 1001, 30, "KALIMANTAN"])
    ws.append(["BC588100294274", 1010, "TECNO", "BCC1", "BC58", "TAM DC PALANGKARAYA", "REGION 4",
               8100294274, "Tecno Megapad 2 8/256 Green", "Tecno Megapad 2", 1001, 5, "KALIMANTAN"])
    # Sel berformat tanpa nilai jauh di bawah & di kanan — persis pola file Tecno.
    ws.cell(1_048_213, 17).font = Font(bold=True)
    ws.cell(5, 40).font = Font(bold=True)
    wb.save(path)


def test_trim_phantom_cells(tmp_path):
    p = tmp_path / "phantom.xlsx"
    _book(p)
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb["Sheet1"]
    assert ws.max_row == 1_048_213
    _trim_phantom_cells(wb)
    assert ws.max_row == 3
    assert ws.max_column == 17   # kolom 40 (berformat, kosong) terbuang
    assert ws.cell(2, 12).value == 30


def test_parse_excel_tidak_menggantung_di_baris_kosong(tmp_path):
    p = tmp_path / "phantom.xlsx"
    _book(p)
    t = time.monotonic()
    res = parse_excel(p, source_email_id="X")
    assert time.monotonic() - t < 10
    assert [r.quantity for r in res.records] == [30, 5]
