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

# Custom CSS - Modern Dark Theme Premium
st.markdown("""
<style>
/* Import Google Fonts */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Poppins:wght@400;500;600;700&display=swap');

/* Global Styles */
.stApp {
    font-family: 'Inter', sans-serif;
    background: linear-gradient(135deg, #0a0a1a 0%, #1a1a2e 50%, #0f0f23 100%);
    color: #e0e0e0;
}

/* Main Container - Glassmorphism */
.main .block-container {
    background: rgba(20, 20, 40, 0.8);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border: 1px solid rgba(102, 126, 234, 0.2);
    border-radius: 24px;
    padding: 2.5rem;
    box-shadow: 
        0 25px 50px rgba(0, 0, 0, 0.5),
        0 0 100px rgba(102, 126, 234, 0.1),
        inset 0 1px 0 rgba(255, 255, 255, 0.1);
    margin-top: 1.5rem;
    margin-bottom: 2rem;
}

/* Main Header - Premium Gradient */
.main-header {
    font-family: 'Poppins', sans-serif;
    font-size: 3rem;
    font-weight: 800;
    text-align: center;
    margin-bottom: 2.5rem;
    padding: 2.5rem 2rem;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 25%, #a855f7 50%, #d946ef 75%, #ec4899 100%);
    background-size: 200% 200%;
    animation: gradientShift 8s ease infinite;
    color: white;
    border-radius: 24px;
    box-shadow: 
        0 20px 60px rgba(139, 92, 246, 0.4),
        0 0 40px rgba(168, 85, 247, 0.3),
        inset 0 2px 0 rgba(255, 255, 255, 0.2);
    text-shadow: 0 2px 10px rgba(0, 0, 0, 0.3);
    letter-spacing: 2px;
    position: relative;
    overflow: hidden;
}

.main-header::before {
    content: '';
    position: absolute;
    top: -50%;
    left: -50%;
    width: 200%;
    height: 200%;
    background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 60%);
    animation: shimmer 3s ease-in-out infinite;
}

@keyframes gradientShift {
    0%, 100% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
}

@keyframes shimmer {
    0%, 100% { transform: translateX(-50%) translateY(-50%); }
    50% { transform: translateX(50%) translateY(50%); }
}

/* Section Header */
.section-header {
    font-family: 'Poppins', sans-serif;
    font-size: 1.3rem;
    font-weight: 600;
    color: white;
    margin: 2rem 0 1.2rem 0;
    padding: 1rem 1.5rem;
    background: linear-gradient(90deg, #6366f1 0%, #8b5cf6 50%, #a855f7 100%);
    border-radius: 16px;
    box-shadow: 
        0 8px 25px rgba(99, 102, 241, 0.35),
        inset 0 1px 0 rgba(255, 255, 255, 0.2);
    border: 1px solid rgba(255, 255, 255, 0.1);
}

/* Info Box - Glass Effect */
.info-box {
    background: rgba(99, 102, 241, 0.15);
    backdrop-filter: blur(10px);
    color: #a5b4fc;
    padding: 1.2rem 1.5rem;
    border-radius: 16px;
    margin: 0.5rem 0;
    border-left: 5px solid #6366f1;
    font-size: 1rem;
    box-shadow: 
        0 8px 25px rgba(99, 102, 241, 0.2),
        inset 0 1px 0 rgba(255, 255, 255, 0.1);
    border: 1px solid rgba(99, 102, 241, 0.3);
}

/* Status OK - Premium Green */
.status-ok {
    background: linear-gradient(135deg, #10b981 0%, #059669 50%, #047857 100%);
    color: white;
    padding: 2rem;
    border-radius: 20px;
    text-align: center;
    font-size: 1.5rem;
    font-weight: 700;
    box-shadow: 
        0 15px 40px rgba(16, 185, 129, 0.4),
        0 0 30px rgba(16, 185, 129, 0.2),
        inset 0 2px 0 rgba(255, 255, 255, 0.2);
    text-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
}

/* Status Error - Premium Red */
.status-error {
    background: linear-gradient(135deg, #ef4444 0%, #dc2626 50%, #b91c1c 100%);
    color: white;
    padding: 2rem;
    border-radius: 20px;
    text-align: center;
    font-size: 1.5rem;
    font-weight: 700;
    box-shadow: 
        0 15px 40px rgba(239, 68, 68, 0.4),
        0 0 30px rgba(239, 68, 68, 0.2),
        inset 0 2px 0 rgba(255, 255, 255, 0.2);
    text-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
}

/* Error Card - Glass Orange */
.error-card {
    background: rgba(249, 115, 22, 0.12);
    backdrop-filter: blur(10px);
    border-left: 5px solid #f97316;
    padding: 1.5rem;
    margin: 0.8rem 0;
    border-radius: 16px;
    box-shadow: 
        0 8px 25px rgba(249, 115, 22, 0.2),
        inset 0 1px 0 rgba(255, 255, 255, 0.1);
    border: 1px solid rgba(249, 115, 22, 0.3);
}

.error-card strong {
    color: #fb923c;
    font-size: 1rem;
    font-weight: 600;
}

.error-card span {
    color: #e0e0e0;
    font-size: 0.95rem;
}

/* Warning Card - Glass Green */
.warning-card {
    background: rgba(34, 197, 94, 0.12);
    backdrop-filter: blur(10px);
    border-left: 5px solid #22c55e;
    padding: 1.5rem;
    margin: 0.8rem 0;
    border-radius: 16px;
    box-shadow: 
        0 8px 25px rgba(34, 197, 94, 0.2),
        inset 0 1px 0 rgba(255, 255, 255, 0.1);
    border: 1px solid rgba(34, 197, 94, 0.3);
}

.warning-card strong {
    color: #4ade80;
    font-size: 1rem;
    font-weight: 600;
}

.warning-card span {
    color: #e0e0e0;
    font-size: 0.95rem;
}

/* Result Labels */
.result-label {
    font-weight: 600;
    color: #c4b5fd;
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.result-value {
    color: #a5b4fc;
    font-size: 1.1rem;
    font-weight: 700;
}

/* Difference Colors */
.diff-positive {
    color: #fca5a5;
    font-weight: 700;
}

.diff-negative {
    color: #86efac;
    font-weight: 700;
}

/* Buttons - Premium */
.stButton > button {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #a855f7 100%);
    color: white;
    border: none;
    padding: 1rem 2.5rem;
    border-radius: 14px;
    font-weight: 600;
    font-size: 1rem;
    box-shadow: 
        0 8px 25px rgba(99, 102, 241, 0.4),
        inset 0 2px 0 rgba(255, 255, 255, 0.2);
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border: 1px solid rgba(255, 255, 255, 0.1);
}

.stButton > button:hover {
    transform: translateY(-3px);
    box-shadow: 
        0 12px 35px rgba(99, 102, 241, 0.5),
        0 0 20px rgba(139, 92, 246, 0.3);
}

.stButton > button:active {
    transform: translateY(-1px);
}

/* File Uploader - Glass */
.stFileUploader {
    border: 2px dashed rgba(99, 102, 241, 0.5);
    border-radius: 16px;
    padding: 1.5rem;
    background: rgba(99, 102, 241, 0.08);
    backdrop-filter: blur(10px);
}

/* Expander - Glass */
.stExpander {
    border: 1px solid rgba(99, 102, 241, 0.3);
    border-radius: 16px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
    background: rgba(30, 30, 50, 0.6);
    backdrop-filter: blur(10px);
}

/* Messages - Glass */
.stSuccess {
    background: rgba(16, 185, 129, 0.15) !important;
    border-left: 5px solid #10b981 !important;
    border-radius: 12px !important;
    color: #6ee7b7 !important;
}

.stWarning {
    background: rgba(245, 158, 11, 0.15) !important;
    border-left: 5px solid #f59e0b !important;
    border-radius: 12px !important;
    color: #fcd34d !important;
}

.stError {
    background: rgba(239, 68, 68, 0.15) !important;
    border-left: 5px solid #ef4444 !important;
    border-radius: 12px !important;
    color: #fca5a5 !important;
}

.stInfo {
    background: rgba(99, 102, 241, 0.15) !important;
    border-left: 5px solid #6366f1 !important;
    border-radius: 12px !important;
    color: #a5b4fc !important;
}

/* Divider */
hr {
    border: none;
    height: 2px;
    background: linear-gradient(90deg, transparent 0%, #6366f1 50%, transparent 100%);
    margin: 2rem 0;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #12121f 0%, #1a1a2e 100%) !important;
    border-right: 1px solid rgba(99, 102, 241, 0.2);
}

[data-testid="stSidebar"] .stMarkdown {
    color: #c4b5fd;
}

/* Progress Bar */
.stProgress > div > div {
    background: linear-gradient(90deg, #6366f1 0%, #a855f7 100%);
    border-radius: 10px;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: rgba(30, 30, 50, 0.5);
    padding: 8px;
    border-radius: 12px;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 10px;
    padding: 12px 24px;
    font-weight: 500;
    color: #a5b4fc;
    border: 1px solid transparent;
}

.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
    color: white !important;
    box-shadow: 0 4px 15px rgba(99, 102, 241, 0.4);
    border: 1px solid rgba(255, 255, 255, 0.2);
}

/* Subtotal Box - Clean Blue */
.subtotal-box {
    background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
    border-radius: 16px;
    padding: 1.5rem;
    text-align: center;
    box-shadow: 0 10px 30px rgba(59, 130, 246, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
    color: white;
}

/* PPN Box - Clean Orange */
.ppn-box {
    background: linear-gradient(135deg, #c2410c 0%, #f97316 100%);
    border-radius: 16px;
    padding: 1.5rem;
    text-align: center;
    box-shadow: 0 10px 30px rgba(249, 115, 22, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
    color: white;
}

/* Grand Total Box - Clean Green */
.grandtotal-box {
    background: linear-gradient(135deg, #15803d 0%, #22c55e 100%);
    border-radius: 16px;
    padding: 1.5rem;
    text-align: center;
    box-shadow: 0 10px 30px rgba(34, 197, 94, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
    color: white;
}

/* Comparison Box - Clean Dark */
.comparison-box {
    background: linear-gradient(135deg, #1f2937 0%, #374151 100%);
    border-radius: 12px;
    padding: 1rem;
    margin: 0.5rem 0;
    border: 1px solid rgba(255, 255, 255, 0.1);
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    color: white;
}

/* Selisih Box - Clean Red */
.selisih-box {
    background: linear-gradient(135deg, #991b1b 0%, #ef4444 100%);
    border-radius: 12px;
    padding: 1rem;
    text-align: center;
    box-shadow: 0 8px 25px rgba(239, 68, 68, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
    color: white;
}

/* Sesuai Box - Clean Green */
.sesuai-box {
    background: linear-gradient(135deg, #166534 0%, #22c55e 100%);
    border-radius: 12px;
    padding: 1rem;
    text-align: center;
    box-shadow: 0 8px 25px rgba(34, 197, 94, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
    color: white;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 8px 30px rgba(0, 0, 0, 0.3);
    border: 1px solid rgba(99, 102, 241, 0.2);
}

/* Input Fields */
.stTextInput > div > div > input {
    border-radius: 12px;
    border: 2px solid rgba(99, 102, 241, 0.3);
    padding: 0.75rem 1rem;
    background: rgba(30, 30, 50, 0.6);
    color: white;
    backdrop-filter: blur(10px);
}

.stTextInput > div > div > input:focus {
    border-color: #8b5cf6;
    box-shadow: 0 0 0 4px rgba(139, 92, 246, 0.2);
}

/* Radio */
.stRadio > div {
    background: rgba(30, 30, 50, 0.6);
    padding: 0.75rem;
    border-radius: 12px;
    border: 1px solid rgba(99, 102, 241, 0.2);
    color: #c4b5fd;
}

/* Selectbox */
.stSelectbox > div > div {
    background: rgba(30, 30, 50, 0.6);
    border-radius: 12px;
    border: 1px solid rgba(99, 102, 241, 0.3);
}

/* Text Colors */
.stMarkdown {
    color: #d1d5db;
}

h1, h2, h3, h4, h5, h6 {
    color: white !important;
    font-family: 'Poppins', sans-serif;
}

p {
    color: #d1d5db;
}

/* Metric */
[data-testid="stMetricValue"] {
    color: white !important;
    font-weight: 700;
}

[data-testid="stMetricLabel"] {
    color: #a5b4fc !important;
}

/* Spinner */
.stSpinner > div {
    border-color: #8b5cf6 transparent transparent transparent;
}

/* Glow Animation */
@keyframes pulse {
    0%, 100% { box-shadow: 0 0 20px rgba(99, 102, 241, 0.3); }
    50% { box-shadow: 0 0 40px rgba(139, 92, 246, 0.5); }
}

.glow {
    animation: pulse 2s ease-in-out infinite;
}

/* Badge */
.badge-ok {
    display: inline-block;
    background: linear-gradient(135deg, #10b981, #059669);
    color: white;
    padding: 0.4rem 1rem;
    border-radius: 20px;
    font-size: 0.85rem;
    font-weight: 600;
    box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4);
}

.badge-error {
    display: inline-block;
    background: linear-gradient(135deg, #ef4444, #dc2626);
    color: white;
    padding: 0.4rem 1rem;
    border-radius: 20px;
    font-size: 0.85rem;
    font-weight: 600;
    box-shadow: 0 4px 15px rgba(239, 68, 68, 0.4);
}

/* Fade In Animation */
@keyframes fadeInUp {
    from {
        opacity: 0;
        transform: translateY(30px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.fade-in {
    animation: fadeInUp 0.6s ease-out;
}

/* Scrollbar */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-track {
    background: rgba(30, 30, 50, 0.5);
    border-radius: 10px;
}

::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    border-radius: 10px;
}

::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(135deg, #8b5cf6, #a855f7);
}

</style>
""", unsafe_allow_html=True)

def format_currency(value):
    """Format angka sebagai mata uang Rupiah"""
    if value is None:
        return "Rp 0"
    try:
        return f"Rp {float(value):,.0f}".replace(',', '.')
    except:
        return str(value)

def main():
    # Header
    st.markdown('<div class="main-header">LOCAL RAB MATHEMATICAL CHECKER</div>', unsafe_allow_html=True)
    
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
    
    # Upload file
    st.markdown('<div class="section-header">UPLOAD FILE</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Choose Excel File",
        type=['xlsx', 'xls'],
        help="Upload file Excel RAB/Quotation yang ingin dicek"
    )
    
    if uploaded_file is not None:
        # Simpan file sementara
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        st.session_state.file_name = uploaded_file.name
        st.session_state.tmp_file_path = tmp_file_path
        
        # Tampilkan info file
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
            <div class="info-box">
                <strong>📁 Nama File:</strong> {uploaded_file.name}
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="info-box">
                <strong>💾 Ukuran:</strong> {uploaded_file.size / 1024:.2f} KB
            </div>
            """, unsafe_allow_html=True)
        
        # Pilih sheet
        reader = ExcelReader(tmp_file_path)
        if reader.load_workbook():
            sheet_names = reader.get_sheet_names()
            
            # Opsi pilihan sheet
            sheet_option = st.radio(
                "📋 Pilih Sheet:",
                ["Sheet Tertentu", "Semua Sheet"],
                horizontal=True
            )
            
            if sheet_option == "Sheet Tertentu":
                selected_sheet = st.selectbox("📋 Pilih Sheet:", sheet_names)
                sheets_to_check = [selected_sheet]
            else:
                sheets_to_check = sheet_names
                st.info(f"📋 Akan mengecek {len(sheets_to_check)} sheet: {', '.join(sheets_to_check)}")
            
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
                            
                            # Baca data
                            data = reader.read_data(sheet_name)
                            
                            # Simpan data Excel untuk perbandingan
                            if 'excel_sheets_data' not in st.session_state:
                                st.session_state.excel_sheets_data = {}
                            st.session_state.excel_sheets_data[sheet_name] = {
                                'subtotal_value': data.get('subtotal_value'),
                                'ppn_value': data.get('ppn_value'),
                                'grand_total_value': data.get('grand_total_value')
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
    """Tampilkan hasil pemeriksaan"""
    results = st.session_state.check_results
    errors = st.session_state.errors
    warnings = st.session_state.warnings
    sheets_checked = st.session_state.get('sheets_checked', [])
    all_items = st.session_state.get('all_items', [])
    
    st.divider()
    st.markdown('<div class="section-header">HASIL PEMERIKSAAN</div>', unsafe_allow_html=True)
    
    # Tampilkan sheet yang diperiksa
    if sheets_checked and len(sheets_checked) > 1:
        st.info(f"📋 Sheet yang diperiksa: {', '.join(sheets_checked)}")
    
    # Status
    if results['total_errors'] == 0:
        st.markdown(
            '<div class="status-ok">✓ SEMUA PERHITUNGAN SESUAI</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f'<div class="status-error">⚠ DITEMUKAN {results["total_errors"]} KESALAHAN</div>',
            unsafe_allow_html=True
        )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Summary Cards
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric(
            label="📋 Sheet",
            value=len(sheets_checked) if sheets_checked else 1,
            help="Jumlah sheet yang diperiksa"
        )
    with col2:
        st.metric(
            label="📦 Total Item",
            value=results['total_items'],
            help="Jumlah item yang diperiksa"
        )
    with col3:
        st.metric(
            label="❌ Error",
            value=results['total_errors'],
            help="Jumlah kesalahan ditemukan"
        )
    with col4:
        st.metric(
            label="⚠️ Warning",
            value=results['total_warnings'],
            help="Jumlah peringatan"
        )
    with col5:
        status = "✅ OK" if results['total_errors'] == 0 else "🔴 PERLU CEK"
        st.metric(
            label="📊 Status",
            value=status,
            help="Status akhir pemeriksaan"
        )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Preview Items - Group by Sheet
    if all_items:
        st.markdown('<div class="section-header">📋 ITEM YANG TERBACA</div>', unsafe_allow_html=True)
        st.caption("Berikut adalah semua item yang berhasil dibaca dari Excel, dikelompokkan per sheet.")
        
        # Group items by sheet
        sheets_data = {}
        for item in all_items:
            sheet = item.get('sheet', 'Unknown')
            if sheet not in sheets_data:
                sheets_data[sheet] = []
            sheets_data[sheet].append(item)
        
        # Tampilkan per sheet
        for sheet_name, items in sheets_data.items():
            st.markdown(f"### 📄 Sheet: {sheet_name}")
            
            # Buat DataFrame
            item_data = []
            sheet_total_calculated = 0
            
            for item in items:
                total_val = item.get('total')
                if total_val is not None:
                    try:
                        sheet_total_calculated += float(total_val)
                    except:
                        pass
                
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
            
            df_items = pd.DataFrame(item_data)
            st.dataframe(df_items, use_container_width=True)
            
            # Ambil nilai dari Excel (dari data per sheet)
            excel_sheets_data = st.session_state.get('excel_sheets_data', {})
            sheet_data = excel_sheets_data.get(sheet_name, {})
            excel_subtotal = sheet_data.get('subtotal_value')
            excel_ppn = sheet_data.get('ppn_value')
            excel_grand_total = sheet_data.get('grand_total_value')
            
            # Hitung PPN dan Grand Total
            calculated_ppn = sheet_total_calculated * 0.11
            calculated_grand_total = sheet_total_calculated + calculated_ppn
            
            # Tampilkan perbandingan dalam box menarik
            st.markdown("""
            <div style="background: linear-gradient(90deg, #1e40af 0%, #7c3aed 50%, #db2777 100%); 
                        color: white; 
                        padding: 1.2rem; 
                        border-radius: 16px; 
                        text-align: center;
                        margin: 1.5rem 0;
                        box-shadow: 0 10px 30px rgba(124, 58, 237, 0.4);
                        border: 1px solid rgba(255, 255, 255, 0.2);">
                <h3 style="margin: 0; color: white; font-weight: 700;">📊 RINGKASAN PERHITUNGAN</h3>
            </div>
            """, unsafe_allow_html=True)
            
            # Tampilkan perbandingan Subtotal
            col1, col2, col3 = st.columns([2, 1, 2])
            with col1:
                st.markdown("""
                <div class="subtotal-box">
                    <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 SUBTOTAL (DIHITUNG)</div>
                    <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                </div>
                """.format(format_currency(sheet_total_calculated)), unsafe_allow_html=True)
            with col2:
                if excel_subtotal is not None:
                    try:
                        excel_val = float(excel_subtotal)
                        difference = sheet_total_calculated - excel_val
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
                if excel_subtotal is not None:
                    try:
                        excel_val = float(excel_subtotal)
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 SUBTOTAL (DI EXCEL)</div>
                            <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(excel_val)), unsafe_allow_html=True)
                    except:
                        pass
            
            # Tampilkan perbandingan PPN
            col1, col2, col3 = st.columns([2, 1, 2])
            with col1:
                st.markdown("""
                <div class="ppn-box">
                    <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 PPN 11% (DIHITUNG)</div>
                    <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                </div>
                """.format(format_currency(calculated_ppn)), unsafe_allow_html=True)
            with col2:
                if excel_ppn is not None:
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
                else:
                    st.markdown("""
                    <div class="comparison-box" style="text-align: center; padding: 1.5rem;">
                        <div style="font-size: 1.2rem; color: #9ca3af;">⚠️ Tidak ada data PPN</div>
                    </div>
                    """, unsafe_allow_html=True)
            with col3:
                if excel_ppn is not None:
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
            
            # Tampilkan perbandingan Grand Total
            col1, col2, col3 = st.columns([2, 1, 2])
            with col1:
                st.markdown("""
                <div class="grandtotal-box">
                    <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📥 GRAND TOTAL (DIHITUNG)</div>
                    <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
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
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 GRAND TOTAL (DI EXCEL)</div>
                            <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(excel_val)), unsafe_allow_html=True)
                    except:
                        pass
                if excel_grand_total is not None:
                    try:
                        excel_val = float(excel_grand_total)
                        st.markdown("""
                        <div class="grandtotal-box">
                            <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem;">📤 Grand Total (di Excel)</div>
                            <div style="font-size: 1.5rem; font-weight: 700; color: #2e7d32;">{}</div>
                        </div>
                        """.format(format_currency(excel_val)), unsafe_allow_html=True)
                    except:
                        pass
            
            st.markdown("---")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Detail Errors
    if errors:
        st.markdown('<div class="section-header">DETAIL KESALAHAN</div>', unsafe_allow_html=True)
        
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
