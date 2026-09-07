"""Unit tests for attachment inspection utilities."""

from io import BytesIO
from pathlib import Path
import openpyxl
from PIL import Image
import pytest


def test_excel_compound_header_detection():
    """Verify detecting and flattening 2-row compound headers in openpyxl."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Allocation"

    # Row 2 compound headers
    ws["D2"] = "FROM "
    ws["E2"] = None
    ws.merge_cells("D2:E2")

    ws["F2"] = "TO STORE "
    ws["G2"] = None
    ws.merge_cells("F2:G2")

    # Row 3 column headers
    ws["A3"] = "Brand"
    ws["B3"] = "Article Code"
    ws["C3"] = "Article Description"
    ws["D3"] = "Site Code"
    ws["E3"] = "Site Desc"
    ws["F3"] = "Site Code"
    ws["G3"] = "Site Desc"
    ws["H3"] = "Alokasi "

    # Check merged ranges
    merged_ranges = [str(r) for r in ws.merged_cells.ranges]
    assert "D2:E2" in merged_ranges
    assert "F2:G2" in merged_ranges

    # Verify unmerged lookup
    from_header = ws["D2"].value.strip()
    to_header = ws["F2"].value.strip()
    assert from_header == "FROM"
    assert to_header == "TO STORE"


def test_image_inspection_utility(tmp_path: Path):
    """Verify extracting image dimensions and format with PIL."""
    img_path = tmp_path / "test_signature.png"
    img = Image.new("RGBA", (200, 100), color=(255, 255, 255, 255))
    img.save(img_path)

    with Image.open(img_path) as loaded_img:
        assert loaded_img.format == "PNG"
        assert loaded_img.size == (200, 100)
        assert loaded_img.mode == "RGBA"
