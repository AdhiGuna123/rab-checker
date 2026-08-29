# Plan Teliti — 5 Case RAB + Beyond (tanpa AI di app)

## Prinsip (sesuai catatan)
- Tidak pakai AI di dalam aplikasi (AI hanya bantu koding).
- Tidak ubah file asli, hanya laporan (RAB AUDIT REPORT).
- Auto Detection label: TOTAL/JUMLAH/SUBTOTAL/PPN/GRAND TOTAL (label-only), pencarian bukan posisi baris tetap.
- Hitungan sumber kebenaran adalah ANGKA (Qty×Price, SUM, PPN×11%, GRAND), label hanya petunjuk.
- Sebelum ubah backend lagi, buat fixture + test harness.

## 5 Case + Expected

| Case | File | PPN | TOTAL | PPN | GRAND | Catatan |
|------|------|-----|-------|-----|-------|---------|
| 1 | case tanpa ppn.xlsx | tidak ada | SUM item | - | GRAND = TOTAL |  |
| 2 | case ppn hanya di 1 bagian.xlsx | hanya di 1 bagian (B) | Total A + Jumlah B | Jumlah B×11% | Total A + Total B |  |
| 3 | case normal.xlsx | di akhir | SUM item | TOTAL×11% | TOTAL+PPN |  |
| 4 | Case ada total masing masing tapi ppn jadi 1.xlsx | PPN 1 di akhir (gabungan A+B) | TOTAL(A+B)=Total A+Total B | TOTAL×11% | TOTAL+PPN |  |
| 5 | 3 sub bagian tapi ppn jadi 1 di akhir.xlsx | PPN 1 di akhir (A+B+C) | TOTAL(A+B+C) dinamis | TOTAL×11% | TOTAL+PPN | deteksi jumlah sub bagian dinamis |

## Urutan Pemeriksaan (checker.py)
1. Qty×Unit Price per item
2. TOTAL/SUBTOTAL per section: SUM item vs Excel
3. Jumlah Global: SUM TOTAL section vs Jumlah Global Excel (hanya jika ada baris JUMLAH global)
4. PPN Global (gabungan) vs PPN 1 Bagian vs Per-section (mode-aware)
5. GRAND TOTAL (sudah termasuk PPN) — skip jika is_without_ppn (sudah di 1)

## Rencana Perbaikan (tanpa ubah UI backend besar)
1. Taruh 5 file Excel ke `tests/fixtures/` (user).
2. Jalankan `python -m pytest tests -v` — harus hijau sebelum ubah excel_reader/checker lagi.
3. Fallback `Rp 0` untuk Case 1: jika tanpa PPN dan subtotal/grand masih None, pakai SUM item (sudah 072a44a, butuh test).
4. Lock backend: jangan ubah excel_reader lagi sampai test 5 case hijau.

## Beyond 5 case
- Toleran typo label (JML/JumlahA/TOTAL(A+B)) via rapidfuzz + normalisasi, tapi validasi tetap angka.
- Jumlah section dinamis: tidak hardcode A/B/C, loop detected_sections.
- Laporan: RAB AUDIT REPORT dengan Sheet/Section/PPN mode + detail lokasi/nilai Excel/seharusnya/selisih.
