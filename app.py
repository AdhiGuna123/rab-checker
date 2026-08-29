import streamlit as st
import pandas as pd
from datetime import datetime
import os
import tempfile

from excel_reader import ExcelReader
from checker import RABChecker
from report import ReportGenerator

# Konfigurasi page
st.set_page_config(
    page_title="LOCAL RAB MATHEMATICAL CHECKER",
    page_icon="checklist",
    layout="wide"
)

# RAB Checker — Human Readable UI (backend tidak diubah)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Nunito:wght@600;700;800&display=swap');

:root{
  --bg:#f4f6ff;
  --card:#ffffff;
  --ink:#1e293b;
  --muted:#64748b;
  --line:#e2e8f0;
  --blue:#2563eb; --blue2:#3b82f6;
  --orange:#f59e0b; --orange2:#fb923c;
  --green:#059669; --green2:#10b981;
  --red:#dc2626; --red2:#ef4444;
  --purple:#7c3aed;
  --radius:20px;
}
.stApp{
  font-family:'Inter',sans-serif;
  background: linear-gradient(180deg, #eef2ff 0%, #f8fafc 40%, #ffffff 100%);
  color: var(--ink);
}
.main .block-container{
  background: transparent;
  padding-top: 1rem;
  max-width: 1100px;
}
.hero{
  background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 45%, #a855f7 100%);
  color:white; border-radius: 28px;
  padding: 2.2rem 1.6rem;
  text-align:center;
  box-shadow: 0 18px 50px rgba(79,70,229,.35);
  margin-bottom: 1.2rem;
}
.hero h1{font-family:'Nunito',sans-serif; font-size:2.15rem; font-weight:800; margin:0; letter-spacing:.5px;}
.hero p{margin:.5rem 0 0 0; font-size:1rem; opacity:.95;}
.hero .chips{margin-top:1rem; display:flex; gap:.5rem; justify-content:center; flex-wrap:wrap;}
.chip{ background: rgba(255,255,255,.22); border:1px solid rgba(255,255,255,.35); color:white; padding:.35rem .75rem; border-radius:999px; font-size:.8rem; font-weight:600; }
.stepper{ display:flex; gap:.6rem; justify-content:center; margin-top: .9rem; }
.step{ display:flex; align-items:center; gap:.5rem; background: white; border:2px solid #e0e7ff; border-radius:999px; padding:.45rem .8rem; font-size:.82rem; font-weight:700; color:#4338ca; }
.step.active{ background:#4338ca; color:white; border-color:#4338ca; }
.step.done{ background:#ecfdf5; color:#065f46; border-color:#a7f3d0;}
.step .num{ width:22px; height:22px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:.75rem; font-weight:800; background:#e0e7ff; color:#4338ca; }
.step.active .num{ background:white; color:#4338ca; }
.step.done .num{ background:#10b981; color:white; }
.card{ background: var(--card); border:1px solid var(--line); border-radius: var(--radius); padding: 1.1rem 1.2rem; box-shadow: 0 10px 30px rgba(15,23,42,.06); margin: 1rem 0; }
.card h3{ font-family:'Nunito',sans-serif; font-size:1.05rem; font-weight:800; margin:0 0 .35rem 0; color:#1e293b;}
.card .hint{ color: var(--muted); font-size:.86rem; margin:0; line-height:1.5;}
 .kpi{ background:white; border:1px solid #e2e8f0; border-radius:16px; padding:.9rem; text-align:center; box-shadow: 0 6px 20px rgba(15,23,42,.05);}
 .kpi .label{ font-size:.72rem; letter-spacing:.04em; color:#64748b; font-weight:700; text-transform:uppercase;}
 .kpi .value{ font-size:1.25rem; font-weight:800; margin-top:.25rem;}
 .kpi.ok{ border-color:#a7f3d0; background: linear-gradient(180deg, #ecfdf5, #ffffff); }
 .kpi.bad{ border-color:#fecaca; background: linear-gradient(180deg, #fef2f2, #ffffff); }
.how{ display:grid; grid-template-columns: repeat(3,1fr); gap:.8rem; margin: .6rem 0; }
.how .how-card{ background:white; border:1px solid #e2e8f0; border-radius:16px; padding:.9rem; display:flex; gap:.75rem; align-items:flex-start; }
.how .how-card .icon{ width:36px; height:36px; border-radius:10px; display:flex; align-items:center; justify-content:center; font-size:1.1rem; }
.illust{ background: linear-gradient(135deg, #eef2ff, #f0fdfa); border:1px dashed #c7d2fe; border-radius:16px; padding:1rem; display:flex; gap:.9rem; align-items:center; }
.illust .pic{ width:64px; height:64px; border-radius:14px; background:white; display:flex; align-items:center; justify-content:center; font-size:2rem; box-shadow: 0 6px 18px rgba(15,23,42,.08);}

.flow{ display:flex; align-items:center; gap:.5rem; flex-wrap:wrap; justify-content:center; margin:.6rem 0; }
.flow .node{ background:white; border:1px solid #e2e8f0; border-radius:14px; padding:.6rem .75rem; min-width:140px; text-align:center; }
.flow .arrow{ font-size:1.1rem; color:#94a3b8;}

.money{ font-variant-numeric: tabular-nums; }
.badge{ display:inline-block; padding:.2rem .55rem; border-radius:999px; font-size:.72rem; font-weight:700;}
.badge.ok{ background:#ecfdf5; color:#065f46; border:1px solid #a7f3d0;}
.badge.bad{ background:#fef2f2; color:#991b1b; border:1px solid #fecaca;}
.badge.neutral{ background:#f1f5f9; color:#475569; border:1px solid #e2e8f0;}

.box{ border-radius:16px; padding:1rem; text-align:center; color:white; box-shadow: 0 10px 24px rgba(15,23,42,.10);}
.box, .subtotal-box, .ppn-box, .grandtotal-box, .selisih-box, .sesuai-box{ padding: 1.1rem 1.4rem !important; margin: .35rem 0 !important; border-radius: 14px !important; }
.subtotal-box, .ppn-box, .grandtotal-box{ min-height: 88px; display:flex; flex-direction:column; align-items:center; justify-content:center; }
.selisih-box, .sesuai-box{ min-height: 88px; display:flex; flex-direction:column; align-items:center; justify-content:center; }
.box.blue{ background: linear-gradient(135deg, #1e40af, #3b82f6); }
.box.orange{ background: linear-gradient(135deg, #c2410c, #f97316); }
.box.green{ background: linear-gradient(135deg, #065f46, #10b981); }
.box.red{ background: linear-gradient(135deg, #991b1b, #ef4444); }
.box.dark{ background: #1e293b; }
.center{ text-align:center;}

.stButton > button{ background: linear-gradient(135deg, #4f46e5, #7c3aed); color:white; border:none; border-radius:14px; padding:.9rem 1.4rem; font-weight:800; letter-spacing:.02em; box-shadow: 0 10px 24px rgba(79,70,229,.30); }
.stButton > button:hover{ transform: translateY(-1px); box-shadow: 0 14px 28px rgba(79,70,229,.35); }
/* Uploader — latar PUTIH terang (bukan hitam) sebelum klik, teks abu gelap */
.stFileUploader{ border: 2px dashed #c7d2fe; background: #ffffff !important; border-radius:16px; padding:.2rem; }
.stFileUploader [data-testid="stFileUploaderDropzone"]{ background: #ffffff !important; border: 2px dashed #a5b4fc !important; border-radius:12px; }
.stFileUploader [data-testid="stFileUploaderDropzone"]:hover{ border-color: #6366f1 !important; background: #eef2ff !important; }
.stFileUploader [data-testid="stFileUploaderDropzone"] span,
.stFileUploader [data-testid="stFileUploaderDropzone"] p,
.stFileUploader [data-testid="stFileUploaderDropzone"] div,
.stFileUploader [data-testid="stFileUploaderDropzone"] label{ color:#ffffff !important; }
.stFileUploader [data-testid="stFileUploaderDropzone"] *{ color:#ffffff !important; }
.stFileUploader [data-testid="stFileUploaderDropzone"] button{ background:white !important; color:#4338ca !important; border:none !important; border-radius:10px !important; font-weight:800 !important; }
.stFileUploader [data-testid="stFileUploaderDropzone"] button p,
.stFileUploader [data-testid="stFileUploaderDropzone"] button span{ color:#4338ca !important; }
.stFileUploader small{ color:#111827 !important; }
[data-testid="stFileUploaderDropzoneInstructions"]{ color:#ffffff !important; }
[data-testid="stFileUploaderDropzoneInstructions"] small{ color:#111827 !important; }
[data-testid="stFileDropzoneInstructions"]{ color:#ffffff !important; }
[data-testid="stFileDropzoneInstructions"] small{ color:#111827 !important; }
[data-testid="stFileDropzoneInstructions"] div small{ color:#111827 !important; }
.stExpander{ border:1px solid #e2e8f0; border-radius:16px; background:white;}
/* Sheet picker — huruf gelap terlihat, terpilih ungu kontras */
.stRadio > div{ background:white; border:2px solid #c7d2fe; border-radius:14px; padding:.7rem; }
.stRadio [role="radiogroup"] label{ background:#ffffff; border:2px solid #cbd5e1; border-radius:12px; padding:.55rem .8rem; font-weight:700; color:#1e293b !important;}
.stRadio [role="radiogroup"] label p{ color:#1e293b !important; font-weight:700 !important;}
.stRadio [role="radiogroup"] label:has(input:checked){ background:#4338ca !important; color:white !important; border-color:#4338ca !important;}
.stRadio [role="radiogroup"] label:has(input:checked) p{ color:white !important;}
[data-testid="stMultiSelect"]{ background:white; border-radius:12px;}
[data-testid="stMultiSelect"] span{ color:#1e293b !important;}
/* Tabel — paksa PUTIH di theme gelap Streamlit */
[data-testid="stDataFrame"], [data-testid="stDataFrame"] > div, [data-testid="stDataFrame"] div{ background:#ffffff !important; }
[data-testid="stDataFrame"]{ border-radius:16px !important; overflow:hidden !important; border:1.5px solid #cbd5e1 !important; box-shadow: 0 8px 24px rgba(15,23,42,.08) !important; }
[data-testid="stDataFrame"] div[data-testid="stDataFrameResizable"]{ background:#ffffff !important;}
[data-testid="stDataFrame"] [role="grid"]{ background:#ffffff !important; text-align:center !important;}
[data-testid="stDataFrame"] [role="columnheader"]{ background:#e0e7ff !important; color:#3730a3 !important; font-weight:800 !important; font-size:.82rem !important; border-bottom: 2px solid #a5b4fc !important; border-right:1px solid #e0e7ff !important; text-align:center !important; justify-content:center !important;}
[data-testid="stDataFrame"] [role="columnheader"] div{ justify-content:center !important; text-align:center !important; color:#3730a3 !important;}
[data-testid="stDataFrame"] [role="gridcell"]{ color:#1e293b !important; font-weight:500 !important; border-bottom:1px solid #eef2ff !important; border-right:1px solid #f1f5f9 !important; background:#ffffff !important; text-align:center !important; justify-content:center !important;}
[data-testid="stDataFrame"] [role="gridcell"] div, [data-testid="stDataFrame"] [role="gridcell"] span{ color:#1e293b !important; text-align:center !important; justify-content:center !important;}
[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"]{ background:#f8fafc !important;}
[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] div, [data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] span{ background:#f8fafc !important;}
[data-testid="stDataFrame"] [role="row"]:hover [role="gridcell"]{ background:#eef2ff !important; }
[data-testid="stDataFrame"] canvas{ background:#ffffff !important;}
/* Fallback: jika dataframe masih hitam (glide-data-grid), paksa table putih */
[data-testid="stTable"]{ background:white !important; }
[data-testid="stTable"] th{ background:#eef2ff !important; color:#3730a3 !important; text-align:center !important; border:1px solid #cbd5e1 !important;}
[data-testid="stTable"] td{ background:white !important; color:#1e293b !important; text-align:center !important; border:1px solid #e2e8f0 !important;}
[data-testid="stTable"] tr:nth-child(even) td{ background:#f8fafc !important;}
hr{ border:none; height:1px; background: linear-gradient(90deg, transparent, #c7d2fe, transparent); margin:1.2rem 0;}
@media (max-width: 768px){ .how{ grid-template-columns: 1fr; } .hero h1{ font-size:1.6rem;} }
</style>
""", unsafe_allow_html=True)

def safe_float(value):
    """Convert value to float safely, handling Indonesian formatting"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace('Rp', '').replace('Rp.', '').replace('Rp ', '').strip()
        if ',' in cleaned and '.' in cleaned:
            cleaned = cleaned.replace('.', '').replace(',', '.')
        elif ',' in cleaned:
            parts = cleaned.split(',')
            if len(parts) == 2 and len(parts[1]) == 3 and parts[1].isdigit() and parts[0].replace('-','').isdigit():
                cleaned = cleaned.replace(',', '')
            elif len(parts) == 2 and len(parts[1]) <= 2:
                cleaned = cleaned.replace(',', '.')
            else:
                cleaned = cleaned.replace(',', '').replace('.', '')
        else:
            cleaned = cleaned.replace('.', '')
        cleaned = cleaned.strip()
        if cleaned == '' or cleaned == '-':
            return None
        try:
            return float(cleaned)
        except:
            return None
    return None

def format_currency(value):
    """Format angka sebagai mata uang Rupiah"""
    if value is None:
        return "Rp 0"
    try:
        num = safe_float(value)
        if num is None:
            return str(value)
        return f"Rp {num:,.0f}".replace(',', '.')
    except:
        return str(value)

def main():
    st.markdown("""
    <div class="hero">
      <h1>🔍 RAB Checker — Cek Hitungan Otomatis</h1>
      <p>Upload Excel RAB → kami hitung ulang <b>Qty × Harga</b>, <b>Jumlah</b>, <b>PPN 11%</b> & <b>Grand Total</b>. Salah hitung langsung terlihat.</p>
      <div class="chips"><span class="chip">✅ Tanpa langganan AI</span><span class="chip">🧠 Toleran typo (JML / JumlahA)</span><span class="chip">📊 Kategori & Section fleksibel</span></div>
      <div class="illust" style="margin-top:1rem; text-align:left;">
        <div class="pic">📊</div>
        <div><b style="color:#1e293b;">Gimana bacanya?</b><br><span style="color:#475569; font-size:.9rem;">Biru = <b>Jumlah</b> (sebelum PPN) &nbsp;•&nbsp; Oranye = <b>PPN 11%</b> &nbsp;•&nbsp; Hijau = <b>Grand Total</b> &nbsp;•&nbsp; Merah = <b>Selisih</b></span></div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Inisialisasi session state
    if 'check_results' not in st.session_state:
        st.session_state.check_results = None
    if 'errors' not in st.session_state:
        st.session_state.errors = None
    if 'warnings' not in st.session_state:
        st.session_state.warnings = None
    if 'file_name' not in st.session_state:
        st.session_state.file_name = None
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    
    # Stepper (awam: 3 langkah)
    cur_step = 1
    if st.session_state.get('check_results'): cur_step = 3
    elif st.session_state.get('file_name'): cur_step = 2
    st.markdown(f"""
    <div class="stepper">
      <div class="step {'active' if cur_step==1 else 'done' if cur_step>1 else ''}"><span class="num">1</span> Upload Excel</div>
      <div class="step {'active' if cur_step==2 else 'done' if cur_step>2 else ''}"><span class="num">2</span> Atur & Cek</div>
      <div class="step {'active' if cur_step==3 else ''}"><span class="num">3</span> Lihat Hasil</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("""
    <div class="how">
      <div class="how-card"><div class="icon" style="background:#e0e7ff;">📤</div><div><b>1. Upload</b><br><span style="color:#64748b; font-size:.85rem;">Pilih file .xlsx RAB</span></div></div>
      <div class="how-card"><div class="icon" style="background:#ffedd5;">⚙️</div><div><b>2. Cek</b><br><span style="color:#64748b; font-size:.85rem;">Sistem hitung ulang otomatis</span></div></div>
      <div class="how-card"><div class="icon" style="background:#dcfce7;">✅</div><div><b>3. Hasil</b><br><span style="color:#64748b; font-size:.85rem;">Selisih langsung terlihat</span></div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="card"><h3>📤 Langkah 1 — Upload File Excel</h3><p class="hint">Pilih file RAB/Quotation (.xlsx). Tidak perlu ubah format — sistem toleran typo <i>Jumlah/TOTAL/JML</i>.</p></div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Tarik file Excel ke sini atau klik Browse  •  .xlsx / .xls",
        type=['xlsx', 'xls'],
        help="RAB/Quotation Excel. Garis putus-putus = area upload. Tabel di bawah akan terang (bukan hitam).",
        label_visibility="collapsed"
    )
    
    if uploaded_file is not None:
        # Simpan file sementara
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        st.session_state.file_name = uploaded_file.name
        st.session_state.tmp_file_path = tmp_file_path
        
        st.markdown(f"""
        <div class="card" style="display:flex; gap:1rem; align-items:center;">
          <div style="font-size:1.6rem;">📄</div>
          <div style="flex:1;"><b>File terpilih:</b> {uploaded_file.name}<br><span style="color:#64748b; font-size:.85rem;">Ukuran {uploaded_file.size/1024:.1f} KB • Siap dicek</span></div>
          <span class="badge ok">Siap</span>
        </div>
        """, unsafe_allow_html=True)
        
        reader = ExcelReader(tmp_file_path)
        if reader.load_workbook():
            sheet_names = reader.get_sheet_names()
            st.markdown('<div class="card" style="border:2px solid #c7d2fe;"><h3>📑 Pilih Sheet yang dicek</h3><p class="hint">Pilih <b>Semua</b> untuk cek sekaligus, atau <b>Pilih sheet</b> untuk cek 1 / beberapa sheet.</p></div>', unsafe_allow_html=True)
            sheet_mode = st.radio("Mode sheet", ["✅ Semua sheet", "📄 Pilih sheet"], horizontal=True, label_visibility="collapsed")
            if sheet_mode == "📄 Pilih sheet":
                picks = st.multiselect("Pilih sheet (bisa lebih dari satu)", sheet_names, default=[sheet_names[0]] if sheet_names else [])
                sheets_to_check = picks if picks else sheet_names
                st.markdown(f"<div class='badge ok'>Dicek: {', '.join(sheets_to_check) if sheets_to_check else '-'}</div>", unsafe_allow_html=True)
            else:
                sheets_to_check = sheet_names
                st.markdown(f"<div class='badge neutral'>Semua sheet: {', '.join(sheets_to_check)}</div>", unsafe_allow_html=True)
            
            st.markdown('<div class="card"><h3>📋 Langkah 2 — Pilih Sheet & Mulai Cek</h3><p class="hint">Default paling cepat: <b>Auto</b>. Upload → Pilih Sheet → <b>START CHECK</b> langsung jadi tanpa pilih. Jika 1 sheet membingungkan, baru pilih <b>Model Case</b> di bawah — sheet lain tetap Auto, tidak ikut kena.</p></div>', unsafe_allow_html=True)
            with st.expander("⚙️ Pengaturan Lanjutan — Model Case per-sheet (opsional, tetap cepat)", expanded=False):
                st.caption("Kosongkan = auto-detect. Isi hanya jika hasil auto salah (mis. PPN gabungan vs per-section). Tetap berjalan lokal tanpa langganan.")
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    adv_header = st.text_input("Baris Header (angka, kosong=auto)", key="adv_header", placeholder="auto")
                with c2:
                    st.markdown("**Model Case (pilih agar tidak bingung PPN/TOTAL satu vs beda)**")
                    model_case = st.selectbox("Model Case", ["Auto — deteksi otomatis", "1 — Tanpa PPN", "2 — PPN hanya di 1 bagian", "3 — Normal (PPN di akhir)", "4 — PPN 1 di akhir (2 bagian: Total A+Total B)", "5 — PPN 1 di akhir (3+ bagian: dinamis)"], key="adv_model_case")
                    # Mapping Model Case -> ppn_mode + total_mode (agar case lain tidak ikut kena)
                    model_map = {
                        "Auto — deteksi otomatis": ("auto", "auto"),
                        "1 — Tanpa PPN": ("none", "auto"),
                        "2 — PPN hanya di 1 bagian": ("single", "auto"),
                        "3 — Normal (PPN di akhir)": ("auto", "auto"),
                        "4 — PPN 1 di akhir (2 bagian: Total A+Total B)": ("combined", "auto"),
                        "5 — PPN 1 di akhir (3+ bagian: dinamis)": ("combined", "auto"),
                    }
                    adv_ppn_mapped, adv_total_mapped = model_map[model_case]
                    st.markdown("**Pengaturan Akurasi (kolom, opsional)**")
                    adv_qty = st.text_input("Kolom Qty (huruf A=1, B=2...)", key="adv_qty", placeholder="auto")
                with c3:
                    adv_price = st.text_input("Kolom Harga Satuan (huruf)", key="adv_price", placeholder="auto")
                with c4:
                    adv_total = st.text_input("Kolom Jumlah/Total (huruf)", key="adv_total", placeholder="auto")
                # Mode detail tetap tapi tidak mengganggu cepat — sudah mapping dari Model Case di atas
                adv_ppn = f"Auto ({adv_ppn_mapped})"
                adv_total_mode = "Gabungan (1 Grand Total)" if model_case.startswith("4") or model_case.startswith("5") else "Auto (deteksi)"
                adv_ppn = f"Auto ({adv_ppn_mapped})"
                adv_total_mode = "Auto (deteksi)"
                st.caption(f"Mode PPN/Total dari Model Case: PPN={adv_ppn} • Total={adv_total_mode}")
                st.caption("Tips: pilih Model Case jika Auto membingungkan PPN/TOTAL satu vs beda. Case lain tidak ikut kena karena mapping langsung ke Mode PPN/Total.")
                # Store for START CHECK
                def _col_letter_to_num(s: str):
                    s = s.strip().upper()
                    if not s: return None
                    if s.isdigit(): return int(s)
                    n = 0
                    for ch in s:
                        if 'A' <= ch <= 'Z': n = n*26 + (ord(ch)-64)
                        else: return None
                    return n if n else None
                # Mapping Model Case -> override (agar tidak bingung kapan total satu vs beda)
                if 'model_case' in locals() and model_case != "Auto — deteksi otomatis":
                    ppn_mode_final = adv_ppn_mapped
                    total_mode_final = adv_total_mapped
                else:
                    ppn_mode_final = {'Auto (deteksi)':'auto','Auto (combined)':'auto','Per-section (masing-masing)':'per_section','Gabungan (1 PPN A+B)':'combined','Hanya 1 section (A atau B)':'single','Tanpa PPN':'none'}.get(adv_ppn, 'auto')
                    total_mode_final = {'Auto (deteksi)':'auto','Per-section (Total A & Total B)':'per_section','Gabungan (1 Grand Total)':'combined'}.get(adv_total_mode, 'auto')
                st.session_state['adv_overrides_preview'] = {
                    'header_row': int(adv_header) if adv_header.strip().isdigit() else None,
                    'qty_col': _col_letter_to_num(adv_qty),
                    'unit_price_col': _col_letter_to_num(adv_price),
                    'total_col': _col_letter_to_num(adv_total),
                    'ppn_mode': ppn_mode_final,
                    'total_mode': total_mode_final,
                    'model_case': model_case
                }
                st.session_state['adv_ai_preview'] = {'provider': 'none', 'gemini_key': ""}
                st.caption(f"Preview override: header={st.session_state['adv_overrides_preview']['header_row'] or 'auto'} qty={adv_qty or 'auto'} price={adv_price or 'auto'} total={adv_total or 'auto'} | PPN={st.session_state['adv_overrides_preview']['ppn_mode']} TOTAL={st.session_state['adv_overrides_preview']['total_mode']}")

            st.markdown("<br>", unsafe_allow_html=True)
            
            # Tombol mulai check
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("🔍 START CHECK", type="primary", use_container_width=True):
                    with st.spinner("⏳ Sedang memeriksa file..."):
                        all_errors = []
                        all_warnings = []
                        total_items = 0
                        st.session_state.all_items = []
                        
                        # Progress bar
                        progress = st.progress(0)
                        status_text = st.empty()
                        
                        for idx, sheet_name in enumerate(sheets_to_check):
                            # Update progress
                            progress.progress((idx + 1) / len(sheets_to_check))
                            status_text.text(f"Memeriksa sheet: {sheet_name}...")
                            
                            # Baca data (Value Intelligence + Label + AI gratis opsional)
                            overrides = st.session_state.get('adv_overrides_preview', {})
                            overrides = {k: v for k, v in overrides.items() if v is not None}
                            ai_overrides = st.session_state.get('adv_ai_preview', {})
                            ai_overrides = {k: v for k, v in ai_overrides.items() if v}
                            try:
                                data = reader.read_data(sheet_name, overrides=overrides if overrides else None, ai_overrides=ai_overrides if ai_overrides else None)
                            except Exception as _e:
                                import traceback
                                data = {'sheet_name': sheet_name, 'subtotal_value': None, 'ppn_value': None, 'grand_total_value': None, 'sections': {}, 'skipped_rows': [{'row': '?', 'dump': [f'read_data error: {_e}', traceback.format_exc()[:600]] }], 'classifications': [], 'summary_rows_debug': [], 'columns': {}, 'overrides_applied': {}, 'header_values_debug': [], 'items': []}
                            if not isinstance(data, dict):
                                data = {'sheet_name': sheet_name, 'subtotal_value': None, 'ppn_value': None, 'grand_total_value': None, 'sections': {}, 'skipped_rows': [{'row': '?', 'dump': [f'data type {type(data)} value={repr(data)[:400]}'] }], 'classifications': [], 'summary_rows_debug': [], 'columns': {}, 'overrides_applied': {}, 'header_values_debug': [], 'items': []}
                            if 'excel_sheets_data' not in st.session_state:
                                st.session_state.excel_sheets_data = {}
                            st.session_state.excel_sheets_data[sheet_name] = {
                                'subtotal_value': data.get('subtotal_value'),
                                'ppn_value': data.get('ppn_value'),
                                'grand_total_value': data.get('grand_total_value'),
                                'sections': data.get('sections', {}),
                                'skipped_rows': data.get('skipped_rows', []),
                                'classifications': data.get('classifications', []),
                                'summary_rows_debug': data.get('summary_rows_debug', []),
                                'columns': data.get('columns', {}),
                                'overrides_applied': data.get('overrides_applied', {}),
                                'header_values_debug': data.get('header_values_debug', [])
                            }
                            
                            # Lakukan pemeriksaan
                            checker = RABChecker(reader)
                            results = checker.check_all(data)
                            
                            # Kumpulkan hasil
                            all_errors.extend(results.get('errors', []))
                            all_warnings.extend(results.get('warnings', []))
                            total_items += results.get('total_items', 0)
                            
                            # Simpan data item untuk preview
                            if 'all_items' not in st.session_state:
                                st.session_state.all_items = []
                            for item in results.get('items', []):
                                item['sheet'] = sheet_name
                                st.session_state.all_items.append(item)
                        
                        # Selesai
                        progress.progress(1.0)
                        status_text.text("Pemeriksaan selesai!")
                        
                        # Gabungkan hasil
                        combined_results = {
                            'total_items': total_items,
                            'total_errors': len(all_errors),
                            'total_warnings': len(all_warnings),
                            'errors': all_errors,
                            'warnings': all_warnings
                        }
                        
                        # Simpan hasil
                        st.session_state.check_results = combined_results
                        st.session_state.errors = all_errors
                        st.session_state.warnings = all_warnings
                        st.session_state.sheets_checked = sheets_to_check
                        
                        st.success(f"✅ Pemeriksaan selesai! {len(sheets_to_check)} sheet diperiksa.")
        
        # Tampilkan hasil jika ada
        if st.session_state.check_results:
            display_results()
    
    # Cleanup
    if 'tmp_file_path' in st.session_state:
        try:
            os.unlink(st.session_state.tmp_file_path)
        except:
            pass

def display_results():
    """Tampilkan hasil pemeriksaan — ramah awam: angka besar, warna jelas, bahasa sederhana."""
    results = st.session_state.check_results
    errors = st.session_state.errors
    warnings = st.session_state.warnings
    sheets_checked = st.session_state.get('sheets_checked', [])
    all_items = st.session_state.get('all_items', [])
    
    st.markdown('<div class="card" style="text-align:center; max-width:760px; margin:1rem auto; border: 2px solid #c7d2fe;"><h3 style="margin:0; text-align:center;">📊 Langkah 3 — Hasil Pemeriksaan</h3><p class="hint" style="margin:.25rem 0 0 0; text-align:center;">Biru = Jumlah (sebelum PPN) • Oranye = PPN 11% • Hijau = Grand Total • Merah = Selisih</p></div>', unsafe_allow_html=True)
    
    # Tampilkan sheet yang diperiksa
    if sheets_checked and len(sheets_checked) > 1:
        st.info(f"📋 Sheet yang diperiksa: {', '.join(sheets_checked)}")
    
    # Status — center
    if results['total_errors'] == 0:
        st.markdown("""
        <div class="card" style="background: linear-gradient(135deg, #ecfdf5, #d1fae5); border:2px solid #6ee7b7; text-align:center; padding:1.4rem; max-width:760px; margin:1rem auto;">
          <div style="font-size:2rem; text-align:center;">✅</div>
          <div style="font-weight:800; font-size:1.2rem; color:#065f46; text-align:center;">Semua hitungan COCOK</div>
          <div style="color:#047857; font-size:.9rem; text-align:center;">Qty × Harga, Jumlah, PPN & Grand Total sudah benar</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="card" style="background: linear-gradient(135deg, #fef2f2, #fee2e2); border:2px solid #fca5a5; text-align:center; padding:1.4rem; max-width:760px; margin:1rem auto;">
          <div style="font-size:2rem; text-align:center;">⚠️</div>
          <div style="font-weight:800; font-size:1.2rem; color:#991b1b; text-align:center;">Ditemukan {results["total_errors"]} yang perlu dicek</div>
          <div style="color:#b91c1c; font-size:.9rem; text-align:center;">Lihat kotak <b>SELISIH</b> merah di bawah — nilai yang benar ada di kolom <b>DIHITUNG</b></div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # KPI awam
    k1,k2,k3,k4 = st.columns(4)
    with k1: st.markdown(f"<div class='kpi'><div class='label'>Sheet dicek</div><div class='value'>{len(sheets_checked) if sheets_checked else 1}</div></div>", unsafe_allow_html=True)
    with k2: st.markdown(f"<div class='kpi'><div class='label'>Jumlah item</div><div class='value'>{results['total_items']}</div></div>", unsafe_allow_html=True)
    with k3: st.markdown(f"<div class='kpi {'bad' if results['total_errors'] else 'ok'}'><div class='label'>Perlu cek</div><div class='value' style=\"color:{'#dc2626' if results['total_errors'] else '#059669'};\">{results['total_errors']}</div></div>", unsafe_allow_html=True)
    with k4:
        ok = results['total_errors']==0
        st.markdown(f"<div class='kpi { 'ok' if ok else 'bad'}'><div class='label'>Status</div><div class='value' style=\"color:{'#059669' if ok else '#dc2626'};\">{'✅ COCOK' if ok else '🔴 CEK LAGI'}</div></div>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Preview Items — ramah awam
    if all_items or st.session_state.get('excel_sheets_data'):
        st.markdown('<div class="card"><h3>🧾 Daftar Item (Qty × Harga = Jumlah)</h3><p class="hint">Ini yang dibaca dari Excel. <b>Total</b> dihitung ulang <b>Qty × Harga Satuan</b>.</p></div>', unsafe_allow_html=True)
        
        # Group items by sheet
        sheets_data = {}
        for item in all_items:
            sheet = item.get('sheet', 'Unknown')
            if sheet not in sheets_data:
                sheets_data[sheet] = []
            sheets_data[sheet].append(item)
        # Ensure sheets with 0 items still render (debug + summary visible)
        for sn in st.session_state.get('excel_sheets_data', {}).keys():
            if sn not in sheets_data:
                sheets_data[sn] = []
        
        # Tampilkan per sheet
        for sheet_name, items in sheets_data.items():
            st.markdown(f"<div class='card' style='border-left: 6px solid #6366f1;'><b>📄 Sheet:</b> {sheet_name} &nbsp;<span class='badge neutral'>{len(items)} item</span></div>", unsafe_allow_html=True)
            
            # Buat DataFrame
            item_data = []
            sheet_total_calculated = 0
            
            for item in items:
                total_val = safe_float(item.get('total'))
                if total_val is not None:
                    sheet_total_calculated += total_val
                
                # Ambil nama item dari kolom yang benar
                item_name = item.get('item_name', '-')
                if item_name == 'ganti kompres the creator' or item_name == 'ganti cdangan kompres atas puyo':
                    # Jika nama item adalah nama sheet, coba ambil dari kolom lain
                    item_name = item.get('description', item.get('name', '-'))
                
                item_data.append({
                    'No': item.get('row', ''),
                    'Item': item_name,
                    'Qty': item.get('qty', '-'),
                    'Unit Price': item.get('unit_price', '-'),
                    'Total': item.get('total', '-')
                })
            
            # Tabel: Rupiah hanya Unit Price & Total, Qty = angka biasa; latar PUTIH bukan hitam
            if item_data:
                for r in item_data:
                    for k in ('Unit Price','Total'):
                        v = r.get(k)
                        try:
                            if v not in ('-','', None) and safe_float(v) is not None:
                                r[k] = format_currency(v)
                        except:
                            pass
                    # Qty: angka biasa, bukan Rp
                    vq = r.get('Qty')
                    try:
                        fv = safe_float(vq) if vq not in ('-','', None) else None
                        if fv is not None:
                            r['Qty'] = f"{fv:,.0f}".replace(',', '.') if fv == int(fv) else f"{fv:,.2f}".replace(',', '.')
                    except:
                        pass
            df_items = pd.DataFrame(item_data)
            # Pakai st.table (putih native, bukan glide-data-grid hitam) + center
            try:
                st.table(df_items.style.set_properties(**{'text-align': 'center'}).set_table_styles([{'selector': 'th', 'props': [('text-align','center'),('background','#eef2ff'),('color','#3730a3')]}, {'selector': 'td', 'props': [('text-align','center')]}]))
            except Exception:
                st.dataframe(df_items, use_container_width=True, hide_index=True, height=min(520, 44+len(df_items)*36))

            # === PANEL DEBUG (tanpa perlu kirim gambar/file) ===
            # Simpan raw values di display debug — copy-paste teks ini ke chat
            with st.expander("🐛 DEBUG — copy teks ini ke chat jika masih salah", expanded=False):
                st.caption("Fungsinya supaya saya bisa lihat nilai mentah Excel tanpa perlu foto.")
                
                # Kolom mapping yang terdeteksi
                excel_sheets_data_dbg = st.session_state.get('excel_sheets_data', {})
                sheet_dbg = excel_sheets_data_dbg.get(sheet_name, {})
                cols_dbg = {}
                # Coba baca dari item pertama atau simpan di excel_sheets_data kalau ada
                # Fallback: tampilkan kolom yang dipakai per item
                st.write("**Kolom terdeteksi (header → kolom):**")
                # Ambil dari excel_reader yang terakhir dipakai — simpan di session
                # Kita tampilkan dari items raw
                dbg_cols = []
                for it in items[:1]:
                    dbg_cols.append(f"qty_raw={it.get('qty_raw','?')} | unit_price_raw={it.get('unit_price_raw','?')} | total_raw={it.get('total_raw','?')}")
                if dbg_cols:
                    st.code("\n".join(dbg_cols), language="text")

                # Tabel debug per item (qty_raw, unit_price_raw, total_raw, excel vs calc)
                debug_rows = []
                for it in items:
                    qty_raw = it.get('qty_raw', it.get('qty','-'))
                    up_raw = it.get('unit_price_raw', it.get('unit_price','-'))
                    total_raw = it.get('total_raw', it.get('total','-'))
                    qty = safe_float(it.get('qty'))
                    up = safe_float(it.get('unit_price'))
                    total = safe_float(it.get('total'))
                    calc = qty * up if qty is not None and up is not None else None
                    debug_rows.append({
                        'Row': it.get('row',''),
                        'Item': str(it.get('item_name','-'))[:18],
                        'qty_raw': str(qty_raw),
                        'qty': qty,
                        'unit_price_raw': str(up_raw),
                        'unit_price': up,
                        'total_raw': str(total_raw),
                        'total(hasil calc)': total,
                        'qty*price': calc,
                        'mismatch': it.get('calc_mismatch', False)
                    })
                st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)
                st.caption("👉 Select semua (Ctrl+A) di tabel atas → Ctrl+C → paste ke chat. Atau screenshot panel ini (lebih mudah dari foto Excel).")

                if st.button("📋 Copy debug sebagai teks", key=f"debug_copy_{sheet_name}"):
                    lines = []
                    lines.append(f"Sheet: {sheet_name} | Items: {len(items)}")
                    for r in debug_rows:
                        lines.append(f"Row {r['Row']:>3} | {r['Item']:<18} | qty_raw={r['qty_raw']} -> {r['qty']} | up_raw={r['unit_price_raw']} -> {r['unit_price']} | total_raw={r['total_raw']} -> {r['total(hasil calc)']} | qty*price={r['qty*price']} | mismatch={r['mismatch']}")
                    st.code("\n".join(lines), language="text")

                # Show skipped rows if any
                skipped = sheet_dbg.get('skipped_rows', [])
                if skipped:
                    st.warning(f"⚠️ {len(skipped)} baris ter-skip (mungkin terdeteksi section/PPN). Detail:")
                    st.code("\n".join([f"Row {s['row']}: {s['dump']}" for s in skipped]), language="text")

                # Klasifikasi baris ringkasan (toleran typo) — angka tetap sumber kebenaran, tulisan hanya petunjuk
                klass = sheet_dbg.get('classifications', [])
                summary_dbg = sheet_dbg.get('summary_rows_debug', [])
                klass_map = {k['row']: k for k in summary_dbg} if summary_dbg else {}
                if klass:
                    st.caption("🧭 Klasifikasi (tulisan → tipe, toleran typo): jika typo, cek ⚠️ tapi tidak bikin error hitungan; hanya angka yang divalidasi.")
                    def _fmt_k(k):
                        v = klass_map.get(k['row'], {}).get('value', None)
                        v_str = f" | value={v:,.0f}" if v is not None else " | value=⚠️ tidak kebaca"
                        return f"Row {k['row']}: '{k['raw']}' → {k['normalized']} → {k['type']}{(' ⚠️ typo' if k.get('fuzzy') else '')}{v_str}"
                    st.code("\n".join([_fmt_k(k) for k in klass]), language="text")
                else:
                    st.caption("🧭 Klasifikasi: tidak ada baris ringkasan terdeteksi (Jumlah/Total/PPN).")

            # Tombol alternatif: override kolom manual
            
            # Ambil nilai dari Excel (dari data per sheet)
            excel_sheets_data = st.session_state.get('excel_sheets_data', {})
            sheet_data = excel_sheets_data.get(sheet_name, {})
            excel_subtotal = sheet_data.get('subtotal_value')
            excel_ppn = sheet_data.get('ppn_value')
            excel_grand_total = sheet_data.get('grand_total_value')
            sections = sheet_data.get('sections', {})
            
            # Hitung total items yang dibaca
            calculated_total_items = sheet_total_calculated
            
            # Alur hitungan awam
            st.markdown("""
            <div class="flow">
              <div class="node"><b>Jumlah</b><br><span style="color:#64748b; font-size:.8rem;">Qty × Harga</span></div>
              <div class="arrow">→</div>
              <div class="node"><b>TOTAL</b><br><span style="color:#64748b; font-size:.8rem;">Jumlah A+B</span></div>
              <div class="arrow">→</div>
              <div class="node"><b>PPN 11%</b><br><span style="color:#64748b; font-size:.8rem;">TOTAL × 11%</span></div>
              <div class="arrow">→</div>
              <div class="node" style="border-color:#86efac; background:#ecfdf5;"><b>GRAND TOTAL</b><br><span style="color:#065f46; font-size:.8rem;">TOTAL + PPN</span></div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("<div style='text-align:center;'><span class='badge neutral'>RINGKASAN — DIHITUNG vs DI EXCEL</span></div>", unsafe_allow_html=True)
            # Header Ringkasan — center
            st.markdown("""
            <div style="background: linear-gradient(90deg, #2563eb 0%, #7c3aed 50%, #059669 100%); 
                        color: white; 
                        padding: 1rem; 
                        border-radius: 16px; 
                        text-align: center;
                        margin: 1rem auto;
                        max-width: 760px;
                        box-shadow: 0 10px 30px rgba(37,99,235,.25);">
                <h3 style="margin: 0; color: white; font-weight: 800; font-size:1.05rem; text-align:center;">📊 RINGKASAN — Bandingkan DIHITUNG vs DI EXCEL</h3>
                <div style="font-size:.82rem; opacity:.95; margin-top:.2rem; text-align:center;">Kiri = hitungan sistem &nbsp;•&nbsp; Tengah = cocok/selisih &nbsp;•&nbsp; Kanan = angka di Excel</div>
            </div>
            """, unsafe_allow_html=True)
            
            # CEK APAKAH ADA MULTIPLE SECTIONS
            has_multiple_sections = len(sections) > 1
            
            if has_multiple_sections:
                # TAMPILKAN PER SECTION
                for section_letter in sorted(sections.keys()):
                    section_data = sections[section_letter]
                    section_items = section_data.get('items', [])
                    section_subtotal_excel = section_data.get('subtotal_value')
                    section_ppn_excel = section_data.get('ppn_value')
                    section_discount_excel = section_data.get('discount_value')
                    section_total_excel = section_data.get('total_value')
                    
                    # Hitung subtotal dari items section ini
                    section_calculated = 0
                    for item in section_items:
                        total_val = safe_float(item.get('total'))
                        if total_val is not None:
                            section_calculated += total_val
                    
                    # Header Section — kategori vs section dijumlah terpisah
                    is_cat = section_data.get('is_category', False)
                    label_suffix = " (Kategori)" if is_cat else ""
                    hdr_color = "linear-gradient(135deg, #475569 0%, #64748b 100%)" if is_cat else "linear-gradient(135deg, #059669 0%, #10b981 100%)"
                    st.markdown(f"""
                    <div style="background: {hdr_color}; 
                                color: white; 
                                padding: 0.8rem; 
                                border-radius: 12px; 
                                text-align: center;
                                margin: 1rem 0;">
                        <h4 style="margin: 0; color: white;">📁 {"KATEGORI" if is_cat else "SECTION"} {section_letter}{label_suffix}</h4>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Subtotal Section (Jumlah X) — selalu tampil DIHITUNG (sebelum PPN), DI EXCEL tampil jika ada atau flag calculated
                    is_calc = section_data.get('subtotal_is_calculated', False)
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 JUMLAH (DIHITUNG)</div>
                            <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(section_calculated)), unsafe_allow_html=True)
                    with col2:
                        if section_subtotal_excel is not None:
                            try:
                                excel_val = float(section_subtotal_excel)
                                difference = section_calculated - excel_val
                                if abs(difference) > 1:
                                    st.markdown("""
                                    <div class="selisih-box" style="display:flex; flex-direction:column; align-items:center; justify-content:center; min-height:88px;">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                    </div>
                                    """.format(format_currency(difference)), unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box" style="display:flex; flex-direction:column; align-items:center; justify-content:center; min-height:88px;">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            except:
                                pass
                        elif is_calc:
                            st.markdown("""
                            <div class="comparison-box" style="text-align: center; padding: 1rem; border: 1px dashed rgba(59,130,246,0.5); display:flex; flex-direction:column; align-items:center; justify-content:center;">
                                <div style="color: #93c5fd; font-size: 0.85rem;">Tidak ada Jumlah di Excel — pakai hitungan</div>
                                <div style="color: #60a5fa; font-weight: 700; margin-top: 0.3rem;">{}</div>
                            </div>
                            """.format(format_currency(section_calculated)), unsafe_allow_html=True)
                        else:
                            st.markdown("""
                            <div style="background: rgba(107, 114, 128, 0.3); border-radius: 12px; padding: 1rem; text-align: center; display:flex; flex-direction:column; align-items:center; justify-content:center;">
                                <div style="color: #9ca3af;">Tidak ada subtotal</div>
                            </div>
                            """, unsafe_allow_html=True)
                    with col3:
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 JUMLAH (DI EXCEL)</div>
                            <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(section_subtotal_excel if section_subtotal_excel is not None else section_calculated) if (section_subtotal_excel is not None or is_calc) else "-"), unsafe_allow_html=True)
                    
                    # PPN Section (fleksibel: tampilkan hanya jika section ini punya PPN)
                    # Jika tidak ada ppn section, skip (PPN mungkin global gabungan)
                    has_section_ppn = section_ppn_excel is not None
                    if has_section_ppn:
                        section_calculated_ppn = section_calculated * 0.11
                        col1, col2, col3 = st.columns([2, 1, 2])
                        with col1:
                            st.markdown("""
                            <div class="ppn-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 PPN 11% (DIHITUNG)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(section_calculated_ppn)), unsafe_allow_html=True)
                        with col2:
                            try:
                                excel_val = float(section_ppn_excel)
                                difference = section_calculated_ppn - excel_val
                                if abs(difference) > 1:
                                    st.markdown("""
                                    <div class="selisih-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                    </div>
                                    """.format(format_currency(difference)), unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            except:
                                pass
                        with col3:
                            st.markdown("""
                            <div class="ppn-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 PPN (DI EXCEL)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(section_ppn_excel)), unsafe_allow_html=True)
                    
                    # Diskon Section (jika ada)
                    if section_discount_excel is not None:
                        col1, col2, col3 = st.columns([2, 1, 2])
                        with col1:
                            st.markdown("""
                            <div style="background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%); 
                                        border-radius: 16px; padding: 1.5rem; text-align: center;
                                        box-shadow: 0 10px 30px rgba(239, 68, 68, 0.3);
                                        border: 1px solid rgba(255, 255, 255, 0.2); color: white;">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 DISKON (DI EXCEL)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(section_discount_excel)), unsafe_allow_html=True)
                        with col2:
                            st.markdown("""
                            <div style="background: rgba(107, 114, 128, 0.3); border-radius: 12px; padding: 1rem; text-align: center; height: 100%;">
                                <div style="color: #9ca3af;">Diskon</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with col3:
                            st.markdown("""
                            <div style="background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%); 
                                        border-radius: 16px; padding: 1.5rem; text-align: center;
                                        box-shadow: 0 10px 30px rgba(239, 68, 68, 0.3);
                                        border: 1px solid rgba(255, 255, 255, 0.2); color: white;">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 DISKON (DI EXCEL)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(section_discount_excel)), unsafe_allow_html=True)
                    
                    # Total Section (Total X = Jumlah + PPN - Diskon)
                    if section_total_excel is not None:
                        # Hitung total seharusnya
                        base = section_calculated
                        ppn = section_calculated * 0.11 if section_ppn_excel is not None else 0
                        discount = float(section_discount_excel) if section_discount_excel else 0
                        section_calculated_total = base + ppn - discount
                        
                        col1, col2, col3 = st.columns([2, 1, 2])
                        with col1:
                            st.markdown("""
                            <div class="grandtotal-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 TOTAL SECTION (DIHITUNG)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(section_calculated_total)), unsafe_allow_html=True)
                        with col2:
                            try:
                                excel_val = float(section_total_excel)
                                difference = section_calculated_total - excel_val
                                if abs(difference) > 1:
                                    st.markdown("""
                                    <div class="selisih-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                    </div>
                                    """.format(format_currency(difference)), unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            except:
                                pass
                        with col3:
                            st.markdown("""
                            <div class="grandtotal-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 TOTAL SECTION (DI EXCEL)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(section_total_excel)), unsafe_allow_html=True)
                
                # === TOTAL KATEGORI (hanya jika perlu) — simpan untuk ringkasan tapi tidak tampil jika 1 PPN ===
                sheet_dbg_global = excel_sheets_data.get(sheet_name, {}) or sheet_data
                excel_ppn_global = sheet_dbg_global.get('ppn_value')
                is_combined_global = sheet_dbg_global.get('ppn_is_combined', False)
                has_any_section_ppn = any(sd.get('ppn_value') is not None for sd in sections.values())
                # TOTAL kategori: Jumlah Global sebelum PPN — jangan tampil jika 1 PPN (sudah ada per-section)
                _is_single_ppn = sum(1 for sd in sections.values() if sd.get('ppn_value') is not None) == 1 and len(sections) > 1
                has_total_kategori = not _is_single_ppn and (sheet_dbg_global.get('subtotal_value') is not None or sheet_dbg_global.get('jumlah_global_excel') is not None)
                total_kategori_excel = sheet_dbg_global.get('jumlah_global_excel') if sheet_dbg_global.get('jumlah_global_excel') is not None else sheet_dbg_global.get('subtotal_value')
                if has_total_kategori and len(sections) > 1:
                    sum_sub_for_total = sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">TOTAL — DIHITUNG</div>
                            <div style="font-size: 0.8rem; opacity:.85; margin-bottom:.3rem;">Jumlah A + Jumlah B</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(sum_sub_for_total)), unsafe_allow_html=True)
                    with col2:
                        try:
                            excel_val_t = float(total_kategori_excel)
                            diff_t = sum_sub_for_total - excel_val_t
                            if abs(diff_t) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌ SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(diff_t)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅ COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">TOTAL benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    with col3:
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">TOTAL — DI EXCEL (sebelum PPN)</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(total_kategori_excel)), unsafe_allow_html=True)

                # === Mode badge (auto-detect 5 case) ===
                _is_without = sheet_dbg_global.get('is_without_ppn', False) if isinstance(sheet_dbg_global, dict) else False
                _mode = "TANPA PPN" if _is_without else ("PPN GABUNGAN" if is_combined_global else ("PPN 1 BAGIAN" if sum(1 for sd in sections.values() if sd.get('ppn_value') is not None)==1 and len(sections)>1 else ("NORMAL" if len(sections)<=1 else "MULTI")))
                st.markdown(f"<div style='text-align:center; margin:.4rem 0;'><span class='badge neutral'>Mode: {_mode}</span></div>", unsafe_allow_html=True)
                if not _is_without:
                    st.markdown("""
                    <div style="display:flex; align-items:center; justify-content:center; gap:.5rem; margin:1rem 0; flex-wrap:wrap;">
                      <span class="badge ok">1. Jumlah A+B = TOTAL</span><span style="color:#94a3b8;">→</span>
                      <span class="badge neutral">2. TOTAL × 11% = PPN</span><span style="color:#94a3b8;">→</span>
                      <span class="badge ok" style="background:#dcfce7; border-color:#86efac;">3. TOTAL + PPN = GRAND</span>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div style="display:flex; align-items:center; justify-content:center; gap:.5rem; margin:1rem 0; flex-wrap:wrap;">
                      <span class="badge ok">1. Σ Item = TOTAL</span><span style="color:#94a3b8;">→</span>
                      <span class="badge ok" style="background:#dcfce7; border-color:#86efac;">GRAND = TOTAL (tanpa PPN)</span>
                    </div>
                    """, unsafe_allow_html=True)

                # Langkah 2 — selalu tampil jika ada PPN (global gabungan ATAU per-section)
                is_single = len(sections) == 1
                has_ppn_section = has_any_section_ppn
                show_global_ppn = not _is_without and (excel_ppn_global is not None and (is_combined_global or not has_ppn_section) or (is_single and has_ppn_section) or (not is_single and has_ppn_section and not is_combined_global))
                if show_global_ppn:
                    if is_single and has_ppn_section:
                        sd0 = list(sections.values())[0]
                        sum_sub_for_ppn = safe_float(sd0.get('subtotal_value')) or safe_float(sheet_dbg_global.get('jumlah_global_excel')) or 0
                        excel_ppn_global = safe_float(sd0.get('ppn_value')) or excel_ppn_global
                        is_combined_global = False
                    elif has_ppn_section and not is_combined_global:
                        # MULTI tapi PPN per-section (NORMAL single-section ganda): pakai TOTAL global vs PPN section pertama
                        sum_sub_for_ppn = safe_float(sheet_dbg_global.get('jumlah_global_excel')) or sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                        # Jika gap global, fallback ke PPN section pertama
                        if excel_ppn_global is None:
                            first_ppn = next((safe_float(sd.get('ppn_value')) for sd in sections.values() if sd.get('ppn_value') is not None), None)
                            excel_ppn_global = first_ppn
                    else:
                        sum_sub_for_ppn = safe_float(sheet_dbg_global.get('jumlah_global_excel')) or sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                    calc_ppn_global = sum_sub_for_ppn * 0.11 if sum_sub_for_ppn else 0
                    st.markdown("""
                    <div class="card" style="border:2px solid #fed7aa; background: linear-gradient(180deg, #fffbeb, #ffffff);">
                      <div style="display:flex; align-items:center; gap:.6rem; margin-bottom:.6rem;">
                        <span style="background:#f97316; color:white; border-radius:8px; padding:.3rem .6rem; font-weight:800;">Langkah 2</span>
                        <b style="font-size:1.05rem;">PPN 11% — dari TOTAL kategori</b>
                        <span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: TOTAL × 11%</span>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="ppn-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">PPN (DIHITUNG)</div>
                            <div style="font-size: 0.8rem; opacity:.85; margin-bottom:.3rem;">TOTAL × 11%</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(calc_ppn_global)), unsafe_allow_html=True)
                    with col2:
                        try:
                            excel_val2 = float(excel_ppn_global)
                            diff2 = calc_ppn_global - excel_val2
                            if abs(diff2) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌ SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(diff2)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅ COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">PPN benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    with col3:
                        st.markdown("""
                        <div class="ppn-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">PPN (DI EXCEL)</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(excel_ppn_global)), unsafe_allow_html=True)

                # Fallback Normal: PPN Section A → isi global agar Langkah 2 tampil (check sudah pakai ini)
                if not show_global_ppn and len(sections) == 1 and not sheet_dbg_global.get('is_without_ppn', False):
                    sd0 = list(sections.values())[0]
                    if sd0.get('ppn_value') is not None and sd0.get('subtotal_value') is not None and excel_ppn_global is None:
                        excel_ppn_global = safe_float(sd0.get('ppn_value'))
                        show_global_ppn = True
                        # Override sum untuk NORMAL single
                        sum_sub_for_ppn = safe_float(sd0.get('subtotal_value')) or 0
                        calc_ppn_global = sum_sub_for_ppn * 0.11
                        st.markdown("""
                        <div class="card" style="border:2px solid #fed7aa; background: linear-gradient(180deg, #fffbeb, #ffffff);">
                          <div style="display:flex; align-items:center; gap:.6rem; margin-bottom:.6rem;">
                            <span style="background:#f97316; color:white; border-radius:8px; padding:.3rem .6rem; font-weight:800;">Langkah 2</span>
                            <b style="font-size:1.05rem;">PPN 11% — dari Jumlah</b>
                            <span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: Jumlah × 11%</span>
                          </div>
                        </div>
                        """, unsafe_allow_html=True)
                        col1, col2, col3 = st.columns([2, 1, 2])
                        with col1:
                            st.markdown(f"""
                            <div class="ppn-box">
                                <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">PPN (DIHITUNG)</div>
                                <div style="font-size: 0.8rem; opacity:.85; margin-bottom:.3rem;">Jumlah × 11%</div>
                                <div style="font-size: 1.7rem; font-weight: 800;">{format_currency(calc_ppn_global)}</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with col2:
                            try:
                                ev = float(excel_ppn_global)
                                d = calc_ppn_global - ev
                                if abs(d) > 1:
                                    st.markdown(f"""
                                    <div class="selisih-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌ SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{format_currency(d)}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅ COCOK</div>
                                        <div style="font-weight: 700; font-size: .85rem; opacity:.9;">PPN benar</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            except: pass
                        with col3:
                            st.markdown(f"""
                            <div class="ppn-box">
                                <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">PPN (DI EXCEL)</div>
                                <div style="font-size: 1.7rem; font-weight: 800;">{format_currency(excel_ppn_global)}</div>
                            </div>
                            """, unsafe_allow_html=True)

                # Judul Langkah 3
                if excel_grand_total is not None:
                    st.markdown("""
                    <div class="card" style="border:2px solid #86efac; background: linear-gradient(180deg, #ecfdf5, #ffffff);">
                      <div style="display:flex; gap:.6rem; align-items:center;"><span style="background:#059669; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 3</span><b>Grand Total — sudah termasuk PPN</b><span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: TOTAL + PPN</span></div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Hitung grand total fleksibel - HONOR apa adanya di Excel, sudah termasuk PPN:
                    # - Jika section punya total_value (sudah termasuk PPN/diskon) -> pakai itu
                    # - Jika tidak, hitung sub + ppn_section (jika ada) + PPN global (jika gabungan) - discount
                    calculated_grand_total = 0
                    for sl in sorted(sections.keys()):
                        sd = sections[sl]
                        sec_total = sd.get('total_value')
                        if sec_total is None:
                            sub = safe_float(sd.get('subtotal_value')) or 0
                            sec_ppn = safe_float(sd.get('ppn_value')) or 0
                            disc = safe_float(sd.get('discount_value')) or 0
                            sec_total = sub + sec_ppn - disc if (sub or sec_ppn or disc) else None
                            if sec_total is None:
                                sec_total = sd.get('subtotal_value')
                        if sec_total is not None:
                            calculated_grand_total += float(sec_total)
                    # Jika ada PPN global gabungan, tambahkan ke grand total (karena per-section tidak ada PPN)
                    if show_global_ppn:
                        sum_sub_for_gt = sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                        calculated_grand_total = sum_sub_for_gt + (sum_sub_for_gt * 0.11)
                        # jika ada diskon global (jarang), kurangi
                        # Note: grand total sudah termasuk PPN global, jadi pakai rumus ini
                    elif has_any_section_ppn:
                        # grand total sudah sum dari section yang sudah termasuk PPN per-section, tidak perlu tambah lagi
                        pass
                    
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="grandtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total — DIHITUNG</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                            <div style="font-size:.75rem; opacity:.85; margin-top:.2rem;">TOTAL + PPN</div>
                        </div>
                        """.format(format_currency(calculated_grand_total)), unsafe_allow_html=True)
                    with col2:
                        try:
                            excel_val = float(excel_grand_total)
                            difference = calculated_grand_total - excel_val
                            if abs(difference) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌ SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅ COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">Grand benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    with col3:
                        st.markdown("""
                        <div class="grandtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total — DI EXCEL</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(excel_grand_total)), unsafe_allow_html=True)
                    st.markdown("<div style='height:.6rem;'></div>", unsafe_allow_html=True)
            
            else:
                has_ppn_single = excel_ppn is not None
                calculated_ppn = calculated_total_items * 0.11 if has_ppn_single else 0
                calculated_grand_total = calculated_total_items + calculated_ppn

                # Judul langkah single
                st.markdown("""
                <div class="card" style="border:2px solid #bfdbfe;">
                  <div style="display:flex; gap:.6rem; align-items:center; flex-wrap:wrap;">
                    <span style="background:#2563eb; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 1</span>
                    <b>Jumlah (sebelum PPN)</b> <span style="color:#64748b; font-size:.85rem;">— penjumlahan semua item</span>
                    <span style="margin-left:auto; background:#f1f5f9; border-radius:999px; padding:.25rem .6rem; font-size:.78rem; color:#475569;"><b>Rumus:</b> Σ Qty × Harga</span>
                  </div>
                </div>
                """, unsafe_allow_html=True)
                 # Subtotal
                col1, col2, col3 = st.columns([2, 1, 2])
                with col1:
                    st.markdown("""
                    <div class="subtotal-box">
                        <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Jumlah (DIHITUNG)</div>
                        <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                    </div>
                    """.format(format_currency(calculated_total_items)), unsafe_allow_html=True)
                with col2:
                    if excel_subtotal is not None:
                        try:
                            excel_val = float(excel_subtotal)
                            difference = calculated_total_items - excel_val
                            if abs(difference) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌ SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅ COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">Jumlah benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                with col3:
                    if excel_subtotal is not None:
                        try:
                            excel_val = float(excel_subtotal)
                            st.markdown("""
                            <div class="subtotal-box">
                                <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Jumlah (DI EXCEL)</div>
                                <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(excel_val)), unsafe_allow_html=True)
                        except:
                            pass
                
                # PPN single — kartu Langkah 2
                if has_ppn_single:
                    st.markdown("""
                    <div class="card" style="border:2px solid #fed7aa; background: linear-gradient(180deg, #fffbeb, #ffffff);">
                      <div style="display:flex; gap:.6rem; align-items:center;"><span style="background:#f97316; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 2</span><b>PPN 11%</b><span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: Jumlah × 11%</span></div>
                    </div>
                    """, unsafe_allow_html=True)
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="ppn-box">
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 PPN 11% (DIHITUNG)</div>
                            <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(calculated_ppn)), unsafe_allow_html=True)
                    with col2:
                        try:
                            excel_val = float(excel_ppn)
                            difference = calculated_ppn - excel_val
                            if abs(difference) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌</div>
                                    <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅</div>
                                    <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    with col3:
                        try:
                            excel_val = float(excel_ppn)
                            st.markdown("""
                            <div class="ppn-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 PPN (DI EXCEL)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(excel_val)), unsafe_allow_html=True)
                        except:
                            pass
                
                st.markdown("""
                <div class="card" style="border:2px solid #86efac; background: linear-gradient(180deg, #ecfdf5, #ffffff);">
                  <div style="display:flex; gap:.6rem; align-items:center;"><span style="background:#059669; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 3</span><b>Grand Total — sudah termasuk PPN</b><span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: Jumlah + PPN</span></div>
                </div>
                """, unsafe_allow_html=True)
                # Grand Total
                col1, col2, col3 = st.columns([2, 1, 2])
                with col1:
                    st.markdown("""
                    <div class="grandtotal-box">
                        <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total — DIHITUNG</div>
                        <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        <div style="font-size:.75rem; opacity:.85; margin-top:.2rem;">Jumlah + PPN</div>
                    </div>
                    """.format(format_currency(calculated_grand_total)), unsafe_allow_html=True)
                with col2:
                    if excel_grand_total is not None:
                        try:
                            excel_val = float(excel_grand_total)
                            difference = calculated_grand_total - excel_val
                            if abs(difference) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">❌ SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">✅ COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">Grand benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    else:
                        st.markdown("""
                        <div class="comparison-box" style="text-align: center; padding: 1.5rem;">
                            <div style="font-size: 1.2rem; color: #9ca3af;">⚠️ Tidak ada data Grand Total</div>
                        </div>
                        """, unsafe_allow_html=True)
                with col3:
                    if excel_grand_total is not None:
                        try:
                            excel_val = float(excel_grand_total)
                            st.markdown("""
                            <div class="grandtotal-box">
                                <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total — DI EXCEL</div>
                                <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(excel_val)), unsafe_allow_html=True)
                        except:
                            pass
            
            st.markdown("---")
    
    # Download RAB AUDIT REPORT (Excel) — tidak mengubah file asli, hanya laporan
    try:
        rg = ReportGenerator()
        rpt_bytes = rg.build_bytes(file_name=st.session_state.get('file_name','RAB.xlsx'), check_results=results, errors=errors, warnings=warnings, sheets_data=st.session_state.get('excel_sheets_data',{}))
        st.download_button("📥 Download RAB AUDIT REPORT (.xlsx)", data=rpt_bytes, file_name=f"{(st.session_state.get('file_name','RAB').rsplit('.',1)[0])}_RAB_AUDIT_REPORT.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    except Exception as _e:
        st.caption(f"Gagal buat laporan: {_e}")

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Detail Errors — RAB AUDIT REPORT style (lokasi, Excel, Seharusnya, Selisih)
    if errors:
        st.markdown('<div class="section-header">📋 RAB AUDIT REPORT — DETAIL TEMUAN</div>', unsafe_allow_html=True)
        
        for i, error in enumerate(errors, 1):
            item_name = error.get('item_name', 'Unknown')
            row = error.get('row', '?')
            error_type = error.get('type', '')
            
            with st.expander(f"❌ Error {i}: {item_name} - Baris {row}", expanded=True):
                # Create a nice layout using columns
                loc_col, status_col = st.columns([3, 1])
                
                with loc_col:
                    st.markdown(f"""
                    <div class="error-card">
                        <div style="margin-bottom: 10px;">
                            <span class="result-label">Lokasi:</span><br>
                            <span class="result-value">{error.get('sheet', '')} Baris {row}</span>
                        </div>
                        <div style="margin-bottom: 10px;">
                            <span class="result-label">Masalah:</span><br>
                            <span class="result-value">{error.get('detail', '')}</span>
                        </div>
                        <div style="margin-bottom: 10px;">
                            <span class="result-label">Perhitungan:</span><br>
                            <span class="result-value">{error.get('calculation', '')}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with status_col:
                    st.markdown(f"""
                    <div class="error-card">
                        <div style="margin-bottom: 10px;">
                            <span class="result-label">Nilai Excel:</span><br>
                            <span class="diff-negative">{format_currency(error.get('actual', 0))}</span>
                        </div>
                        <div style="margin-bottom: 10px;">
                            <span class="result-label">Seharusnya:</span><br>
                            <span class="diff-negative">{format_currency(error.get('expected', 0))}</span>
                        </div>
                        <div style="margin-bottom: 10px;">
                            <span class="result-label">Selisih:</span><br>
                            <span class="diff-negative">{format_currency(error.get('difference', 0))}</span>
                        </div>
                        <div>
                            <span class="result-label">Status:</span><br>
                            <span style="color: #ff5722; font-weight: bold; font-size: 1.1rem;">{error.get('status', '')}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Tampilkan detail item jika ada (untuk SUBTOTAL_ERROR)
                if error_type == 'SUBTOTAL_ERROR' and error.get('items_summary'):
                    st.markdown("---")
                    st.markdown("**📋 Detail Item yang Terbaca:**")
                    st.code(error.get('items_summary', ''), language=None)
    
    # Detail Warnings
    if warnings:
        st.markdown('<div class="section-header">PERINGATAN</div>', unsafe_allow_html=True)
        
        for i, warning in enumerate(warnings, 1):
            item_name = warning.get('item_name', 'Unknown')
            row = warning.get('row', '?')
            
            with st.expander(f"⚠️ Warning {i}: {item_name} - Baris {row}"):
                st.markdown(f"""
                <div class="warning-card">
                    <div style="margin-bottom: 10px;">
                        <span class="result-label">Lokasi:</span><br>
                        <span class="result-value">{warning.get('sheet', '')} Baris {row}</span>
                    </div>
                    <div style="margin-bottom: 10px;">
                        <span class="result-label">Masalah:</span><br>
                        <span class="result-value">{warning.get('detail', '')}</span>
                    </div>
                    <div style="margin-bottom: 10px;">
                        <span class="result-label">Nilai:</span><br>
                        <span class="result-value">{warning.get('value', '')}</span>
                    </div>
                    <div>
                        <span class="result-label">Status:</span><br>
                        <span style="color: #ff9800; font-weight: bold; font-size: 1.1rem;">{warning.get('status', '')}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    
    # Export Report
    st.divider()
    st.markdown('<div class="section-header">EXPORT REPORT</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("📥 Export Report (RAB_CHECK_REPORT.xlsx)", type="primary", use_container_width=True):
            reporter = ReportGenerator()
            
            # Generate report
            df_summary, df_errors, df_warnings = reporter.generate_report(
                st.session_state.file_name,
                results,
                errors,
                warnings
            )
            
            # Simpan ke buffer
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_summary.to_excel(writer, sheet_name='Summary', index=False)
                df_errors.to_excel(writer, sheet_name='Errors', index=False)
                df_warnings.to_excel(writer, sheet_name='Warnings', index=False)
                
                # Format kolom
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
            
            buffer.seek(0)
            
            # Download button
            st.download_button(
                label="⬇️ Download Report",
                data=buffer,
                file_name="RAB_CHECK_REPORT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
