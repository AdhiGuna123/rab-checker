#!/usr/bin/env python3
"""Fix broken emoji: replace literal ???? with HTML entities"""
import re

with open('app.py', 'r', encoding='utf-8-sig') as f:
    content = f.read()

original = content

# Title
content = content.replace('???? RAB Checker ??? Cek Hitungan Otomatis',
    '&#128202; RAB Checker &#8212; Cek Hitungan Otomatis')

# Chips
content = content.replace('??? Tanpa langganan AI',
    '&#10003; Tanpa langganan AI')
content = content.replace('???? Toleran typo (JML / JumlahA)',
    '&#128161; Toleran typo (JML / Jumlah)')
content = content.replace('???? Kategori &amp; Section fleksibel',
    '&#128295; Kategori &amp; Section fleksibel')

# Hero illustration
content = content.replace('<div class="pic">????</div>',
    '<div class="pic">&#128269;</div>')

# Langkah 1
content = content.replace('???? Langkah 1 ??? Upload File Excel',
    '&#128194; Langkah 1 &#8212; Upload File Excel')
content = content.replace('??? sistem toleran typo',
    '&#8212; sistem toleran typo')

# Upload icon
content = content.replace('<div style="font-size:1.6rem;">????</div>',
    '<div style="font-size:1.6rem;">&#128194;</div>')

# Settings header
content = content.replace('?????? Pengaturan Pemeriksaan',
    '&#128295; Pengaturan Pemeriksaan')

# Sheet picker
content = content.replace('**???? Sheet**',
    '**&#128196; Sheet**')
content = content.replace('["??? Semua", "???? Pilih"]',
    '["&#128194; Semua", "&#128269; Pilih"]')
content = content.replace('if sheet_mode == "???? Pilih":',
    'if sheet_mode == "&#128269; Pilih":')

# Model Case
content = content.replace('**???? Model Case**',
    '**&#128295; Model Case**')

# Advanced settings expander
content = content.replace('???? Pengaturan Lanjutan',
    '&#128295; Pengaturan Lanjutan')

# Preview items
content = content.replace("???? <b>{_items_preview} item</b> terdeteksi di sheet <b>{preview_sheet}</b> ??? klik",
    "&#128269; <b>{_items_preview} item</b> terdeteksi di sheet <b>{preview_sheet}</b> &#8212; klik")

# START CHECK button
content = content.replace('???? START CHECK',
    '&#128640; START CHECK')

# Langkah 3 header
content = content.replace('???? Langkah 3 ??? Hasil Pemeriksaan',
    '&#128202; Langkah 3 &#8212; Hasil Pemeriksaan')
content = content.replace('Biru = Jumlah (sebelum PPN) ??? Oranye = PPN 11% ??? Hijau = Grand Total ??? Merah = Selisih',
    '&#128308; Merah = Selisih &#8226; &#128994; Oranye = PPN &#8226; &#128994; Hijau = Total')

# Sheet checked info
content = content.replace('???? Sheet yang diperiksa:',
    '&#128196; Sheet yang diperiksa:')

# Status banner icon
content = content.replace('<span style="font-size:1.1rem;">??????</span>',
    '<span style="font-size:1.1rem;">&#128203;</span>')

# KPI labels
content = content.replace("???? Sheet Dicek", "&#128196; Sheet Dicek")
content = content.replace("???? Jumlah Item", "&#128230; Jumlah Item")
content = content.replace("???? Perlu Dicek", "&#9888;&#65039; Perlu Dicek")
content = content.replace("???? Status", "&#128203; Status")

# Status values
content = content.replace("??? COCOK", "&#10004;&#65039; COCOK")
content = content.replace("???? CEK LAGI", "&#128308; CEK LAGI")

# Item table header
content = content.replace("???? Daftar Item (Qty ?? Harga = Jumlah)",
    "&#128203; Daftar Item (Qty x Harga = Jumlah)")

# Sheet card
content = content.replace("<b>???? Sheet:</b>",
    "<b>&#128196; Sheet:</b>")

# Show more/less
content = content.replace('f"???? Tampilkan semua {_total_rows} baris"',
    'f"&#128194; Tampilkan semua {_total_rows} baris"')
content = content.replace('f"?????? Sembunyikan, tampilkan 10 baris saja"',
    'f"&#128065; Sembunyikan, tampilkan 10 baris saja"')

# Debug section
content = content.replace('???? DEBUG ??? copy teks ini ke chat jika masih salah',
    '&#128269; DEBUG &#8212; copy teks ini ke chat jika masih salah')
content = content.replace('???? Select semua (Ctrl+A) di tabel atas ??? Ctrl+C ??? paste ke chat',
    '&#128269; Select semua (Ctrl+A) di tabel atas &#8212; Ctrl+C &#8212; paste ke chat')
content = content.replace('???? Copy debug sebagai teks',
    '&#128203; Copy debug sebagai teks')

# Skipped items
content = content.replace('?????? {len(skipped)} baris ter-skip',
    '&#9888;&#65039; {len(skipped)} baris ter-skip')

# Classification
content = content.replace("???? Klasifikasi (tulisan ??? tipe, toleran typo): jika typo, cek ?????? tapi tidak bikin error hitungan; hanya angka yang divalidasi.",
    "&#128269; Klasifikasi: jika typo, cek warning tapi tidak bikin error hitungan; hanya angka yang divalidasi.")

# Value not read
content = content.replace('value=?????? tidak kebaca',
    'value=&#10060; tidak kebaca')
content = content.replace('?????? typo', '&#9888;&#65039; typo')

# No summary
content = content.replace("???? Klasifikasi: tidak ada baris ringkasan terdeteksi (Jumlah/Total/PPN).",
    "&#128269; Klasifikasi: tidak ada baris ringkasan terdeteksi.")

# Ringkasan header
content = content.replace('???? RINGKASAN ??? Bandingkan DIHITUNG vs DI EXCEL',
    '&#128202; RINGKASAN &#8212; Bandingkan DIHITUNG vs DI EXCEL')

# Section headers
content = content.replace('???? {"KATEGORI" if is_cat else "SECTION"}',
    '&#128196; {"KATEGORI" if is_cat else "SECTION"}')

# DIHITUNG / DI EXCEL labels
content = content.replace('???? JUMLAH (DIHITUNG)', '&#128200; JUMLAH (DIHITUNG)')
content = content.replace('???? JUMLAH (DI EXCEL)', '&#128196; JUMLAH (DI EXCEL)')
content = content.replace('???? PPN 11% (DIHITUNG)', '&#128200; PPN 11% (DIHITUNG)')
content = content.replace('???? PPN (DI EXCEL)', '&#128196; PPN (DI EXCEL)')
content = content.replace('???? DISKON (DI EXCEL)', '&#128196; DISKON (DI EXCEL)')
content = content.replace('???? TOTAL SECTION (DIHITUNG)', '&#128200; TOTAL SECTION (DIHITUNG)')
content = content.replace('???? TOTAL SECTION (DI EXCEL)', '&#128196; TOTAL SECTION (DI EXCEL)')

# No grand total
content = content.replace('?????? Tidak ada data Grand Total',
    '&#128196; Tidak ada data Grand Total')

# Audit report
content = content.replace('???? RAB AUDIT REPORT ??? DETAIL TEMUAN',
    '&#128202; RAB AUDIT REPORT &#8212; DETAIL TEMUAN')

# Detail item
content = content.replace('???? Detail Item yang Terbaca:',
    '&#128203; Detail Item yang Terbaca:')

# Warning expander
content = content.replace('f"?????? Warning {i}:',
    'f"&#9888;&#65039; Warning {i}:')

# Export button
content = content.replace('???? Export Report',
    '&#128230; Export Report')

# Download button
content = content.replace('?????? Download Report',
    '&#128229; Download Report')

# Langkah 2 reference (check if exists)
content = content.replace('???? Langkah 2',
    '&#128202; Langkah 2')

# Catch any remaining ???? or ????
content = re.sub(r'\?{4}(?![a-zA-Z])', '&#128269;', content)
content = re.sub(r'\?{3}(?![a-zA-Z])', '&#8212;', content)

with open('app.py', 'w', encoding='utf-8-sig') as f:
    f.write(content)

changes = sum(1 for a, b in zip(original, content) if a != b)
print(f'Done - {changes} character changes')
