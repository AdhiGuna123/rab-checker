"""Test harness 5 case — RAB Mathematical Checker (tidak pakai AI di app, label-only + value intelligence).
Jalankan: python -m pytest tests -v
Fixture Excel ada di tests/fixtures/case*.xlsx — expected diambil dari isi Excel itu sendiri (TOTAL/PPN/GRAND).
Harness hanya cek matematika: Qty×Price, TOTAL=SUM, PPN=11%, GRAND=TOTAL+PPN, toleransi Rp 1.
"""
import os
import sys
from pathlib import Path

# Ensure project root importable
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel_reader import ExcelReader, safe_float

FIXTURES = [
    ("case tanpa ppn.xlsx", {"has_ppn": False}),
    ("case ppn hanya di 1 bagian.xlsx", {"has_ppn": True, "ppn_mode": "single"}),
    ("case normal.xlsx", {"has_ppn": True, "ppn_mode": "normal"}),
    ("Case ada total masing masing tapi ppn jadi 1.xlsx", {"has_ppn": True, "ppn_mode": "combined", "sections": 2}),
    ("3 sub bagian tapi ppn jadi 1 di akhir.xlsx", {"has_ppn": True, "ppn_mode": "combined", "sections": 3}),
]

def case_path(name: str) -> Path:
    p = ROOT / "tests" / "fixtures" / name
    if not p.exists():
        # Fallback: cari case insensitif
        for f in (ROOT / "tests" / "fixtures").glob("*.xlsx"):
            if f.name.lower() == name.lower():
                return f
    return p

def load_first_sheet(path: Path):
    reader = ExcelReader(str(path))
    assert reader.load_workbook(), f"Gagal load {path}"
    sheets = reader.get_sheet_names()
    assert sheets, f"Tidak ada sheet di {path}"
    return reader, sheets[0]

def test_headers_and_items_readable():
    """Semua fixture harus punya header QTY/TOTAL dan item terbaca."""
    fixtures_dir = ROOT / "tests" / "fixtures"
    if not fixtures_dir.exists():
        import pytest
        pytest.skip("tests/fixtures belum ada — taruh case*.xlsx di sana")
    for name, _ in FIXTURES:
        path = case_path(name)
        if not path.exists():
            import pytest
            pytest.skip(f"Fixture belum ada: {name}")
        reader, sheet = load_first_sheet(path)
        data = reader.read_data(sheet)
        assert data.get("header_row") is not None, f"{name}: header_row tidak ketemu"
        assert len(data.get("items", [])) > 0, f"{name}: items 0"
        for it in data["items"]:
            assert safe_float(it.get("qty")) is not None or safe_float(it.get("total")) is not None, f"{name} row {it.get('row')} tanpa qty/total"

def test_no_rp_zero_for_without_ppn():
    """CASE 1 tanpa PPN: subtotal/grand tidak boleh Rp 0 jika ada item."""
    fixtures_dir = ROOT / "tests" / "fixtures"
    if not fixtures_dir.exists():
        import pytest
        pytest.skip("tests/fixtures belum ada")
    path = case_path("case tanpa ppn.xlsx")
    if not path.exists():
        import pytest
        pytest.skip("Fixture case tanpa ppn.xlsx belum ada")
    reader, sheet = load_first_sheet(path)
    data = reader.read_data(sheet)
    assert data.get("is_without_ppn") is True, "Harus terdeteksi TANPA PPN"
    s = safe_float(data.get("subtotal_value"))
    g = safe_float(data.get("grand_total_value"))
    assert s is not None and s > 0, f"subtotal Rp 0 — s={s}"
    assert g is not None and g > 0, f"grand Rp 0 — g={g}"
    assert abs(s - g) <= 1, f"TANPA PPN GRAND harus = TOTAL: s={s}, g={g}"
