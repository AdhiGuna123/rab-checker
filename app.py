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

# RAB Checker &#8212; Human Readable UI (backend tidak diubah)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Nunito:wght@600;700;800&display=swap');

:root{
  --bg:#f6f8f5;
  --card:#ffffff;
  --ink:#1e293b;
  --muted:#64748b;
  --line:#e2e8f0;
  --sage:#8EA58C; --sage2:#a3b8a1;
  --orange:#f59e0b; --orange2:#fb923c;
  --emerald:#10b981; --emerald2:#34d399;
  --red:#dc2626; --red2:#ef4444;
  
  --radius:20px;
}
.stApp{
  font-family:'Inter',sans-serif;
  background: linear-gradient(180deg, #f0f4ef 0%, #f8faf8 40%, #ffffff 100%);
  color: var(--ink);
  color-scheme: light;
}
.main .block-container{
  background: transparent;
  padding-top: 1rem;
  max-width: 1100px;
}
.hero{
  background: linear-gradient(135deg, #6b8a68 0%, #8EA58C 40%, #10b981 100%);
  color:white; border-radius: 28px;
  padding: 2.2rem 1.6rem;
  text-align:center;
  box-shadow: 0 18px 50px rgba(142,165,140,.35);
  margin-bottom: 1.2rem;
}
.hero h1{font-family:'Nunito',sans-serif; font-size:2.15rem; font-weight:800; margin:0; letter-spacing:.5px;}
.hero p{margin:.5rem 0 0 0; font-size:1rem; opacity:.95;}
.hero .chips{margin-top:1rem; display:flex; gap:.5rem; justify-content:center; flex-wrap:wrap;}
.chip{ background: rgba(255,255,255,.22); border:1px solid rgba(255,255,255,.35); color:white; padding:.35rem .75rem; border-radius:999px; font-size:.8rem; font-weight:600; }
.stepper{ display:flex; gap:.6rem; justify-content:center; margin-top: .9rem; }
.step{ display:flex; align-items:center; gap:.5rem; background: white; border:2px solid #e8ede7; border-radius:999px; padding:.45rem .8rem; font-size:.82rem; font-weight:700; color:#5a7a57; }
.step.active{ background:#5a7a57; color:white; border-color:#5a7a57; }
.step.done{ background:#ecfdf5; color:#065f46; border-color:#a7f3d0;}
.step .num{ width:22px; height:22px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:.75rem; font-weight:800; background:#e8ede7; color:#5a7a57; }
.step.active .num{ background:white; color:#5a7a57; }
.step.done .num{ background:#10b981; color:white; }
.card{ background: var(--card); border:1px solid var(--line); border-radius: var(--radius); padding: .9rem 1.1rem; box-shadow: 0 8px 24px rgba(15,23,42,.05); margin: .6rem 0; }
.card h3{ font-family:'Nunito',sans-serif; font-size:1.05rem; font-weight:800; margin:0 0 .35rem 0; color:#1e293b;}
.card .hint{ color: var(--muted); font-size:.86rem; margin:0; line-height:1.5;}
 .kpi{ background:white; border:1px solid #e2e8f0; border-radius:14px; padding:.7rem .5rem; text-align:center; box-shadow: 0 4px 14px rgba(15,23,42,.04); transition: transform .15s; }
 .kpi:hover{ transform: translateY(-2px); box-shadow: 0 8px 20px rgba(15,23,42,.08); }
 .kpi .label{ font-size:.68rem; letter-spacing:.05em; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:.2rem; }
 .kpi .value{ font-size:1.15rem; font-weight:800; margin-top:.15rem; color:#1e293b; }
 .kpi.ok{ border-color:#a7f3d0; background: linear-gradient(180deg, #ecfdf5, #ffffff); }
 .kpi.bad{ border-color:#fecaca; background: linear-gradient(180deg, #fef2f2, #ffffff); }
.how{ display:grid; grid-template-columns: repeat(3,1fr); gap:.8rem; margin: .6rem 0; }
.how .how-card{ background:white; border:1px solid #e2e8f0; border-radius:16px; padding:.9rem; display:flex; gap:.75rem; align-items:flex-start; }
.how .how-card .icon{ width:36px; height:36px; border-radius:10px; display:flex; align-items:center; justify-content:center; font-size:1.1rem; }
.illust{ background: linear-gradient(135deg, #eef2ff, #f0fdfa); border:1px dashed #c7d9c5; border-radius:16px; padding:1rem; display:flex; gap:.9rem; align-items:center; }
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
.box.blue{ background: linear-gradient(135deg, #1e40af, #10b981); }
.box.orange{ background: linear-gradient(135deg, #c2410c, #f97316); }
.box.green{ background: linear-gradient(135deg, #065f46, #10b981); }
.box.red{ background: linear-gradient(135deg, #991b1b, #ef4444); }
.box.dark{ background: #1e293b; }
.center{ text-align:center;}

.stButton > button{ background: linear-gradient(135deg, #6b8a68, #6b8a68); color:white; border:none; border-radius:14px; padding:.9rem 1.4rem; font-weight:800; letter-spacing:.02em; box-shadow: 0 10px 24px rgba(142,165,140,.30); }
.stButton > button:hover{ transform: translateY(-1px); box-shadow: 0 14px 28px rgba(142,165,140,.35); }
/* Uploader &#8212; latar PUTIH terang sebelum upload, teks gelap */
.stFileUploader{ border: 2px dashed #c7d9c5; background: #ffffff !important; border-radius:16px; padding:.2rem; }
.stFileUploader [data-testid="stFileUploaderDropzone"]{ background: #ffffff !important; border: 2px dashed #8EA58C !important; border-radius:12px; }
.stFileUploader [data-testid="stFileUploaderDropzone"]:hover{ border-color: #8EA58C !important; background: #eef2ff !important; }
/* Dropzone teks instruksi &#8212; gelap terbaca */
.stFileUploader [data-testid="stFileUploaderDropzone"] p,
.stFileUploader [data-testid="stFileUploaderDropzone"] label,
.stFileUploader [data-testid="stFileUploaderDropzone"] span{ color:#374151 !important; }
.stFileUploader [data-testid="stFileUploaderDropzone"] button{ background:#5a7a57 !important; color:white !important; border:none !important; border-radius:10px !important; font-weight:800 !important; }
.stFileUploader [data-testid="stFileUploaderDropzone"] button p,
.stFileUploader [data-testid="stFileUploaderDropzone"] button span{ color:white !important; }
.stFileUploader small{ color:#6b7280 !important; }
/* Uploaded file chip &#8212; visible dark text on light bg */
[data-testid="stFileUploader"] [data-testid="stBaseUploader-header"],
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"],
[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
[data-testid="stFileUploader"] [data-testid="stBaseUpoader-file"],
[data-testid="stFileUploader"] section[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] section[data-testid="stFileUploader"] div{ color:#1e293b !important; }
[data-testid="stFileUploaderDropzoneInstructions"]{ color:#374151 !important; }
[data-testid="stFileDropzoneInstructions"]{ color:#374151 !important; }
[data-testid="stFileDropzoneInstructions"] div small{ color:#6b7280 !important; }
.stExpander{ border:1px solid #e2e8f0; border-radius:16px; background:white;}
/* Sheet picker &#8212; huruf gelap terlihat, terpilih ungu kontras */
.stRadio > div{ background:white; border:2px solid #c7d9c5; border-radius:14px; padding:.7rem; }
.stRadio [role="radiogroup"] label{ background:#ffffff; border:2px solid #d1dcd0; border-radius:12px; padding:.55rem .8rem; font-weight:700; color:#1e293b !important;}
.stRadio [role="radiogroup"] label p{ color:#1e293b !important; font-weight:700 !important;}
.stRadio [role="radiogroup"] label:has(input:checked){ background:#5a7a57 !important; color:white !important; border-color:#5a7a57 !important;}
.stRadio [role="radiogroup"] label:has(input:checked) p{ color:white !important;}
[data-testid="stMultiSelect"]{ background:white; border-radius:12px;}
[data-testid="stMultiSelect"] span{ color:#1e293b !important;}
/* Selectbox & Multiselect &#8212; paksa light */
[data-testid="stSelectbox"],
[data-testid="stMultiSelect"]{
  background: white !important;
}
[data-testid="stSelectbox"] [data-baseweb="select"],
[data-testid="stMultiSelect"] [data-baseweb="select"]{
  background: #ffffff !important;
  border-color: #c7d9c5 !important;
  color: #1e293b !important;
}
[data-testid="stSelectbox"] [data-baseweb="select"] *,
[data-testid="stMultiSelect"] [data-baseweb="select"] *{
  color: #1e293b !important;
  background: transparent !important;
}
[data-testid="stSelectbox"] [data-baseweb="select"] svg,
[data-testid="stMultiSelect"] [data-baseweb="select"] svg{
  color: #8EA58C !important;
}
/* Multiselect chips &#8212; light */
[data-testid="stMultiSelect"] [data-baseweb="tag"]{ background:#eef2ff !important; color:#5a7a57 !important; border:1px solid #c7d9c5 !important; }
[data-testid="stMultiSelect"] [data-baseweb="tag"] span{ color:#5a7a57 !important; }
[data-testid="stMultiSelect"] [data-baseweb="input"]{ color:#1e293b !important; }
[data-testid="stMultiSelect"] [data-baseweb="input"]::placeholder{ color:#94a3b8 !important; }
[data-testid="stMultiSelect"] div[role="listbox"]{ background:white !important; }
[data-testid="stMultiSelect"] div[role="option"]{ color:#1e293b !important; }
/* Dropdown menu &#8212; light */
div[data-baseweb="menu"]{ background:#ffffff !important; border:1px solid #e2e8f0 !important; box-shadow: 0 8px 24px rgba(0,0,0,.12) !important; }
div[data-baseweb="menu"] div[role="option"]{ color:#1e293b !important; }
div[data-baseweb="menu"] div[role="option"]:hover,
div[data-baseweb="menu"] div[role="option"]:focus,
div[data-baseweb="menu"] div[role="option"][aria-selected="true"]{ background:#eef2ff !important; color:#5a7a57 !important; }
/* Force light on all baseweb internals */
[data-baseweb="select"]{ background:#ffffff !important; }
[data-baseweb="select"] *{ color-scheme: light !important; }
[data-baseweb="input"]{ color-scheme: light !important; }
[data-baseweb="tag"]{ color-scheme: light !important; }
[data-baseweb="menu"]{ color-scheme: light !important; }
/* Tabel &#8212; paksa PUTIH di theme gelap Streamlit */
[data-testid="stDataFrame"], [data-testid="stDataFrame"] > div, [data-testid="stDataFrame"] div{ background:#ffffff !important; }
[data-testid="stDataFrame"]{ border-radius:16px !important; overflow:hidden !important; border:1.5px solid #d1dcd0 !important; box-shadow: 0 8px 24px rgba(15,23,42,.08) !important; }
[data-testid="stDataFrame"] div[data-testid="stDataFrameResizable"]{ background:#ffffff !important;}
[data-testid="stDataFrame"] [role="grid"]{ background:#ffffff !important; text-align:center !important;}
[data-testid="stDataFrame"] [role="columnheader"]{ background:#e8ede7 !important; color:#3d5a3a !important; font-weight:800 !important; font-size:.82rem !important; border-bottom: 2px solid #8EA58C !important; border-right:1px solid #e8ede7 !important; text-align:center !important; justify-content:center !important;}
[data-testid="stDataFrame"] [role="columnheader"] div{ justify-content:center !important; text-align:center !important; color:#3d5a3a !important;}
[data-testid="stDataFrame"] [role="gridcell"]{ color:#1e293b !important; font-weight:500 !important; border-bottom:1px solid #eef2ff !important; border-right:1px solid #f1f5f9 !important; background:#ffffff !important; text-align:center !important; justify-content:center !important;}
[data-testid="stDataFrame"] [role="gridcell"] div, [data-testid="stDataFrame"] [role="gridcell"] span{ color:#1e293b !important; text-align:center !important; justify-content:center !important;}
[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"]{ background:#f8fafc !important;}
[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] div, [data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] span{ background:#f8fafc !important;}
[data-testid="stDataFrame"] [role="row"]:hover [role="gridcell"]{ background:#eef2ff !important; }
[data-testid="stDataFrame"] canvas{ background:#ffffff !important;}
/* Fallback: jika dataframe masih hitam (glide-data-grid), paksa table putih */
[data-testid="stTable"]{ background:white !important; }
[data-testid="stTable"] th{ background:#eef2ff !important; color:#3d5a3a !important; text-align:center !important; border:1px solid #d1dcd0 !important;}
[data-testid="stTable"] td{ background:white !important; color:#1e293b !important; text-align:center !important; border:1px solid #e2e8f0 !important;}
[data-testid="stTable"] tr:nth-child(even) td{ background:#f8fafc !important;}
hr{ border:none; height:1px; background: linear-gradient(90deg, transparent, #c7d9c5, transparent); margin:1.2rem 0;}
@media (max-width: 768px){ .how{ grid-template-columns: 1fr; } .hero h1{ font-size:1.6rem;} }
/* Lotus Splash Screen */
@keyframes lotusBloom{
  0%{ transform: scale(0) rotate(-20deg); opacity:0; filter:blur(4px); }
  50%{ transform: scale(1.15) rotate(5deg); opacity:1; filter:blur(0); }
  70%{ transform: scale(0.95) rotate(-2deg); opacity:1; }
  100%{ transform: scale(1) rotate(0deg); opacity:1; }
}
@keyframes lotusFadeOut{
  0%{ opacity:1; visibility:visible; }
  99%{ opacity:0; visibility:visible; }
  100%{ opacity:0; visibility:hidden; pointer-events:none; }
}
@keyframes petalL{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(-18deg) scale(1); opacity:1; } }
@keyframes petalR{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(18deg) scale(1); opacity:1; } }
@keyframes petalL2{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(-38deg) scale(1); opacity:1; } }
@keyframes petalR2{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(38deg) scale(1); opacity:1; } }
@keyframes petalL3{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(-55deg) scale(1); opacity:1; } }
@keyframes petalR3{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(55deg) scale(1); opacity:1; } }
@keyframes leafL{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(-20deg) scale(1); opacity:0.7; } }
@keyframes leafR{ 0%{ transform: rotate(0deg) scale(0); opacity:0; } 100%{ transform: rotate(20deg) scale(1); opacity:0.7; } }
@keyframes glowPulse{ 0%,100%{ filter: drop-shadow(0 0 8px rgba(236,72,153,.3)); } 50%{ filter: drop-shadow(0 0 18px rgba(236,72,153,.5)); } }
@keyframes lotusText{ 0%{ opacity:0; transform:translateY(18px); } 100%{ opacity:1; transform:translateY(0); } }
@keyframes lotusWelcome{ 0%{ opacity:0; transform:translateY(10px) scale(0.95); } 100%{ opacity:1; transform:translateY(0) scale(1); } }
.splash{ position:fixed; top:0; left:0; width:100vw; height:100vh; z-index:99999; display:flex; flex-direction:column; align-items:center; justify-content:center; background: linear-gradient(180deg, #f8f4f0 0%, #fdf8f5 30%, #ffffff 70%); animation: lotusFadeOut 0.8s ease-in 3.5s forwards; }
.splash-lotus{ position:relative; width:160px; height:140px; margin-bottom:1rem; animation: lotusBloom 2s ease-out forwards, glowPulse 2s ease-in-out 2s infinite; }
.splash-lotus svg{ width:160px; height:140px; overflow:visible; }
.splash-welcome{ font-family:'Nunito',sans-serif; font-size:.9rem; font-weight:600; color:#be185d; letter-spacing:.3px; animation: lotusWelcome 0.8s ease-out 1.4s both; }
.splash-title{ font-family:'Nunito',sans-serif; font-size:2.6rem; font-weight:800; color:#1e293b; animation: lotusText 0.8s ease-out 1.8s both; }
.splash-sub{ font-size:.95rem; color:#64748b; margin-top:.3rem; animation: lotusText 0.8s ease-out 2.2s both; }
.splash-hint{ font-size:.8rem; color:#94a3b8; margin-top:1.8rem; animation: lotusText 0.8s ease-out 2.8s both; }
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
    if 'splash_done' not in st.session_state:
        st.session_state.splash_done = False
    if not st.session_state.splash_done:
        st.markdown("""
        <div class="splash">
          <div class="splash-lotus">
            <svg viewBox="0 0 200 180" xmlns="http://www.w3.org/2000/svg">
              <defs>
                <radialGradient id="petalGrad" cx="50%" cy="30%" r="70%">
                  <stop offset="0%" style="stop-color:#fce7f3;stop-opacity:1"/>
                  <stop offset="40%" style="stop-color:#f9a8d4;stop-opacity:0.9"/>
                  <stop offset="100%" style="stop-color:#ec4899;stop-opacity:0.7"/>
                </radialGradient>
                <radialGradient id="petalGrad2" cx="50%" cy="30%" r="70%">
                  <stop offset="0%" style="stop-color:#fdf2f8;stop-opacity:1"/>
                  <stop offset="40%" style="stop-color:#fbcfe8;stop-opacity:0.85"/>
                  <stop offset="100%" style="stop-color:#f472b6;stop-opacity:0.65"/>
                </radialGradient>
                <radialGradient id="centerGrad" cx="50%" cy="50%" r="50%">
                  <stop offset="0%" style="stop-color:#fbbf24;stop-opacity:1"/>
                  <stop offset="100%" style="stop-color:#f59e0b;stop-opacity:0.8"/>
                </radialGradient>
                <radialGradient id="leafGrad" cx="50%" cy="70%" r="60%">
                  <stop offset="0%" style="stop-color:#86efac;stop-opacity:0.9"/>
                  <stop offset="100%" style="stop-color:#22c55e;stop-opacity:0.7"/>
                </radialGradient>
                <filter id="softShadow">
                  <feDropShadow dx="0" dy="2" stdDeviation="3" flood-color="#be185d" flood-opacity="0.15"/>
                </filter>
              </defs>
              <ellipse cx="70" cy="145" rx="30" ry="10" fill="url(#leafGrad)" style="animation: leafL 0.8s ease-out 0.2s both; transform-origin: 100px 145px;"/>
              <ellipse cx="130" cy="145" rx="30" ry="10" fill="url(#leafGrad)" style="animation: leafR 0.8s ease-out 0.2s both; transform-origin: 100px 145px;"/>
              <ellipse cx="100" cy="70" rx="16" ry="42" fill="url(#petalGrad2)" filter="url(#softShadow)" style="animation: petalL3 0.9s ease-out 0.3s both; transform-origin: 100px 112px;"/>
              <ellipse cx="100" cy="70" rx="16" ry="42" fill="url(#petalGrad2)" filter="url(#softShadow)" style="animation: petalR3 0.9s ease-out 0.3s both; transform-origin: 100px 112px;"/>
              <ellipse cx="100" cy="65" rx="17" ry="44" fill="url(#petalGrad)" filter="url(#softShadow)" style="animation: petalL2 0.9s ease-out 0.5s both; transform-origin: 100px 109px;"/>
              <ellipse cx="100" cy="65" rx="17" ry="44" fill="url(#petalGrad)" filter="url(#softShadow)" style="animation: petalR2 0.9s ease-out 0.5s both; transform-origin: 100px 109px;"/>
              <ellipse cx="100" cy="62" rx="15" ry="40" fill="url(#petalGrad)" filter="url(#softShadow)" style="animation: petalL 0.9s ease-out 0.7s both; transform-origin: 100px 102px;"/>
              <ellipse cx="100" cy="62" rx="15" ry="40" fill="url(#petalGrad)" filter="url(#softShadow)" style="animation: petalR 0.9s ease-out 0.7s both; transform-origin: 100px 102px;"/>
              <ellipse cx="100" cy="58" rx="12" ry="38" fill="url(#petalGrad)" style="animation: petalL 0.8s ease-out 0.9s both; transform-origin: 100px 96px;"/>
              <ellipse cx="100" cy="58" rx="12" ry="38" fill="url(#petalGrad)" style="animation: petalR 0.8s ease-out 0.9s both; transform-origin: 100px 96px;"/>
              <ellipse cx="100" cy="65" rx="8" ry="20" fill="#fce7f3" opacity="0.9"/>
              <circle cx="100" cy="95" r="6" fill="url(#centerGrad)"/>
              <circle cx="100" cy="95" r="3" fill="#fbbf24" opacity="0.8"/>
            </svg>
          </div>
          <div class="splash-welcome">Selamat Datang Kak Wintari</div>
          <div class="splash-title">RAB Checker</div>
          <div class="splash-sub">Sistem Pemeriksaan Hitungan Otomatis</div>
          <div class="splash-hint">&#127800; Sedang menyiapkan sistem...</div>
        </div>
        """, unsafe_allow_html=True)
        st.session_state.splash_done = True

    st.markdown("""
    <div class="hero">
      <h1>&#128202; RAB Checker &#8212; Cek Hitungan Otomatis</h1>
      <p>Upload Excel RAB &#8212; kami hitung ulang <b>Qty ?? Harga</b>, <b>Jumlah</b>, <b>PPN 11%</b> & <b>Grand Total</b>. Salah hitung langsung terlihat.</p>
      <div class="chips"><span class="chip">&#10003; Tanpa langganan AI</span><span class="chip">&#128161; Toleran typo (JML / Jumlah)</span><span class="chip">&#128269; Kategori & Section fleksibel</span></div>
      <div class="illust" style="margin-top:1rem; text-align:left;">
        <div class="pic">&#128269;</div>
        <div><b style="color:#1e293b;">Gimana bacanya?</b><br><span style="color:#475569; font-size:.9rem;">Biru = <b>Jumlah</b> (sebelum PPN) &nbsp;&#8212;&nbsp; Oranye = <b>PPN 11%</b> &nbsp;&#8212;&nbsp; Hijau = <b>Grand Total</b> &nbsp;&#8212;&nbsp; Merah = <b>Selisih</b></span></div>
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
      <div class="how-card"><div class="icon" style="background:#e8ede7;">&#128194;</div><div><b>1. Upload</b><br><span style="color:#64748b; font-size:.85rem;">Pilih file .xlsx RAB</span></div></div>
      <div class="how-card"><div class="icon" style="background:#ffedd5;">&#128269;</div><div><b>2. Cek</b><br><span style="color:#64748b; font-size:.85rem;">Sistem hitung ulang otomatis</span></div></div>
      <div class="how-card"><div class="icon" style="background:#dcfce7;">&#128202;</div><div><b>3. Hasil</b><br><span style="color:#64748b; font-size:.85rem;">Selisih langsung terlihat</span></div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="card"><h3>&#128194; Langkah 1 &#8212; Upload File Excel</h3><p class="hint">Pilih file RAB/Quotation (.xlsx). Tidak perlu ubah format &#8212; sistem toleran typo <i>Jumlah/TOTAL/JML</i>.</p></div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Tarik file Excel ke sini atau klik Browse  &#8212;  .xlsx / .xls",
        type=['xlsx', 'xls'],
        help="RAB/Quotation Excel.",
        label_visibility="collapsed"
    )
    
    if uploaded_file is not None:
        _prev_file = st.session_state.get('file_name', '')
        if _prev_file and _prev_file != uploaded_file.name:
            for _k in ['excel_sheets_data', 'all_items', 'check_results', 'errors', 'warnings', 'sheets_checked', 'sheet_total_calculated']:
                st.session_state.pop(_k, None)
            for _sk in list(st.session_state.keys()):
                if _sk.startswith('_show_all_'):
                    st.session_state.pop(_sk, None)

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        st.session_state.file_name = uploaded_file.name
        st.session_state.tmp_file_path = tmp_file_path
        
        st.markdown(f"""
        <div class="card" style="display:flex; gap:1rem; align-items:center; border-left:4px solid #059669;">
          <div style="font-size:1.6rem;">&#128194;</div>
          <div style="flex:1; color:#1e293b;"><b>{uploaded_file.name}</b><br><span style="color:#64748b; font-size:.85rem;">{uploaded_file.size/1024:.1f} KB &#8212; Siap dicek</span></div>
          <span class="badge ok">Siap</span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<div id='start_check'></div>", unsafe_allow_html=True)
        
        reader = ExcelReader(tmp_file_path)
        if reader.load_workbook():
            sheet_names = reader.get_sheet_names()

            st.markdown("""
            <div class="card" style="border-left:4px solid #6b8a68;">
              <h3>&#128295; Pengaturan Pemeriksaan</h3>
            </div>
            """, unsafe_allow_html=True)

            col_sheet, col_model = st.columns([3, 5])
            with col_sheet:
                st.markdown("**&#128196; Sheet**")
                sheet_mode = st.radio("Sheet", ["&#128194; Semua", "&#128269; Pilih"], horizontal=True, label_visibility="collapsed")
                if sheet_mode == "&#128269; Pilih":
                    picks = st.multiselect("Pilih sheet", sheet_names, default=[sheet_names[0]] if sheet_names else [], label_visibility="collapsed")
                    sheets_to_check = picks if picks else sheet_names
                else:
                    sheets_to_check = sheet_names

            with col_model:
                st.markdown("**&#128295; Model Case**")
                st.caption("Pilih jika auto-deteksi salah. Biasanya biarkan Auto saja.")
                model_case = st.selectbox("Model Case", ["Auto &#8212; deteksi otomatis", "1 &#8212; Tanpa PPN", "2 &#8212; PPN hanya di 1 bagian", "3 &#8212; Normal (PPN di akhir)", "4 &#8212; PPN 1 di akhir (2 bagian: Total A+Total B)", "5 &#8212; PPN 1 di akhir (3+ bagian: dinamis)"], key="adv_model_case", label_visibility="collapsed")

            model_map = {
                "Auto &#8212; deteksi otomatis": ("auto", "auto"),
                "1 &#8212; Tanpa PPN": ("none", "auto"),
                "2 &#8212; PPN hanya di 1 bagian": ("single", "auto"),
                "3 &#8212; Normal (PPN di akhir)": ("auto", "auto"),
                "4 &#8212; PPN 1 di akhir (2 bagian: Total A+Total B)": ("combined", "auto"),
                "5 &#8212; PPN 1 di akhir (3+ bagian: dinamis)": ("combined", "auto"),
            }
            adv_ppn_mapped, adv_total_mapped = model_map[model_case]

            with st.expander("&#128295; Pengaturan Lanjutan (Kolom & Header)", expanded=False):
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    adv_header = st.text_input("Baris Header", key="adv_header", placeholder="auto")
                with c2:
                    adv_qty = st.text_input("Kolom Qty", key="adv_qty", placeholder="auto")
                with c3:
                    adv_price = st.text_input("Kolom Harga", key="adv_price", placeholder="auto")
                with c4:
                    adv_total = st.text_input("Kolom Total", key="adv_total", placeholder="auto")

            adv_ppn = f"Auto ({adv_ppn_mapped})"
            adv_total_mode = "Auto (deteksi)"
            def _col_letter_to_num(s: str):
                s = s.strip().upper()
                if not s: return None
                if s.isdigit(): return int(s)
                n = 0
                for ch in s:
                    if 'A' <= ch <= 'Z': n = n*26 + (ord(ch)-64)
                    else: return None
                return n if n else None
            if model_case != "Auto &#8212; deteksi otomatis":
                ppn_mode_final = adv_ppn_mapped
                total_mode_final = adv_total_mapped
            else:
                ppn_mode_final = 'auto'
                total_mode_final = 'auto'
            st.session_state['adv_overrides_preview'] = {
                'header_row': int(adv_header) if adv_header and adv_header.strip().isdigit() else None,
                'qty_col': _col_letter_to_num(adv_qty) if adv_qty else None,
                'unit_price_col': _col_letter_to_num(adv_price) if adv_price else None,
                'total_col': _col_letter_to_num(adv_total) if adv_total else None,
                'ppn_mode': ppn_mode_final,
                'total_mode': total_mode_final,
                'model_case': model_case
            }
            st.session_state['adv_ai_preview'] = {'provider': 'none', 'gemini_key': ""}

            try:
                preview_reader = ExcelReader(tmp_file_path)
                if preview_reader.load_workbook():
                    preview_sheet = sheets_to_check[0] if sheets_to_check else None
                    if preview_sheet:
                        preview_reader.select_sheet(preview_sheet)
                        hr = preview_reader.find_header_row()
                        if hr:
                            cols = preview_reader.find_data_columns(hr)
                            _items_preview = 0
                            for _r in range(hr+1, min(hr+60, preview_reader.ws.max_row+1)):
                                _v = preview_reader.ws.cell(row=_r, column=cols.get('qty',4)).value
                                if safe_float(_v) is not None:
                                    _items_preview += 1
                            st.markdown(f"<div style='text-align:center; color:#64748b; font-size:.85rem; margin:.3rem 0;'>&#128269; <b>{_items_preview} item</b> terdeteksi di sheet <b>{preview_sheet}</b> &#8212; klik <b>START CHECK</b> untuk mulai</div>", unsafe_allow_html=True)
            except: pass
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("&#128640; START CHECK", type="primary", use_container_width=True):
                    with st.spinner("&#8212; Sedang memeriksa file..."):
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
                                'header_values_debug': data.get('header_values_debug', []),
                                'ppn_is_combined': data.get('ppn_is_combined', False),
                                'jumlah_global_excel': data.get('jumlah_global_excel'),
                                'is_without_ppn': data.get('is_without_ppn', False),
                                'items': data.get('items', [])
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
                        
                        st.success(f"&#8212; Pemeriksaan selesai! {len(sheets_to_check)} sheet diperiksa.")
        
        # Tampilkan hasil jika ada
        if st.session_state.get('check_results'):
            display_results()
    
    # Cleanup
    if 'tmp_file_path' in st.session_state:
        try:
            os.unlink(st.session_state.tmp_file_path)
        except:
            pass

def display_results():
    """Tampilkan hasil pemeriksaan &#8212; ramah awam: angka besar, warna jelas, bahasa sederhana."""
    results = st.session_state.get('check_results') or {}
    errors = st.session_state.get('errors') or []
    warnings = st.session_state.get('warnings') or []
    sheets_checked = st.session_state.get('sheets_checked') or []
    all_items = st.session_state.get('all_items') or []
    
    if not results:
        return
    
    st.markdown('<div class="card" style="text-align:center; max-width:760px; margin:1rem auto; border: 2px solid #c7d9c5;"><h3 style="margin:0; text-align:center;">&#128202; Langkah 3 &#8212; Hasil Pemeriksaan</h3><p class="hint" style="margin:.25rem 0 0 0; text-align:center;">&#128308; Merah = Selisih &#8226; &#128994; Oranye = PPN &#8226; &#128994; Hijau = Total</p></div>', unsafe_allow_html=True)
    
    # Tampilkan sheet yang diperiksa
    if sheets_checked and len(sheets_checked) > 1:
        st.info(f"&#128196; Sheet yang diperiksa: {', '.join(sheets_checked)}")
    
    _total_errors = results.get('total_errors', 0)
    _total_items = results.get('total_items', 0)
    
    # Status &#8212; center
    if _total_errors == 0:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #ecfdf5, #d1fae5); border:1px solid #6ee7b7; border-radius:12px; text-align:center; padding:.6rem 1rem; margin:.5rem 0; display:flex; align-items:center; justify-content:center; gap:.6rem;">
          <span style="font-size:1.1rem;">&#8212;</span>
          <span style="font-weight:700; font-size:.9rem; color:#065f46;">Semua hitungan COCOK</span>
          <span style="color:#047857; font-size:.78rem;">&#8212; Qty ?? Harga, Jumlah, PPN & Grand Total sudah benar</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #fef2f2, #fee2e2); border:1px solid #fca5a5; border-radius:12px; text-align:center; padding:.6rem 1rem; margin:.5rem 0; display:flex; align-items:center; justify-content:center; gap:.6rem;">
          <span style="font-size:1.1rem;">&#128203;</span>
          <span style="font-weight:700; font-size:.9rem; color:#991b1b;">Ditemukan {_total_errors} yang perlu dicek</span>
          <span style="color:#b91c1c; font-size:.78rem;">&#8212; Lihat kotak <b>SELISIH</b> merah di bawah</span>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<div style='height:.3rem;'></div>", unsafe_allow_html=True)
    
    # KPI awam
    k1,k2,k3,k4 = st.columns(4)
    with k1: st.markdown(f"<div class='kpi'><div class='label'>&#128196; Sheet Dicek</div><div class='value'>{len(sheets_checked) if sheets_checked else 1}</div></div>", unsafe_allow_html=True)
    with k2: st.markdown(f"<div class='kpi'><div class='label'>&#128230; Jumlah Item</div><div class='value'>{_total_items}</div></div>", unsafe_allow_html=True)
    with k3: st.markdown(f"<div class='kpi {'bad' if _total_errors else 'ok'}'><div class='label'>&#9888;&#65039; Perlu Dicek</div><div class='value' style=\"color:{'#dc2226' if _total_errors else '#059669'};\">{_total_errors}</div></div>", unsafe_allow_html=True)
    with k4:
        ok = _total_errors==0
        st.markdown(f"<div class='kpi { 'ok' if ok else 'bad'}'><div class='label'>&#128203; Status</div><div class='value' style=\"color:{'#059669' if ok else '#dc2226'};\">{'&#10004;&#65039; COCOK' if ok else '&#128308; CEK LAGI'}</div></div>", unsafe_allow_html=True)
    
    st.markdown("<div style='height:.3rem;'></div>", unsafe_allow_html=True)
    
    # Preview Items &#8212; ramah awam
    if all_items or st.session_state.get('excel_sheets_data'):
        st.markdown('<div class="card"><h3>&#128203; Daftar Item (Qty x Harga = Jumlah)</h3><p class="hint">Ini yang dibaca dari Excel. <b>Total</b> dihitung ulang <b>Qty ?? Harga Satuan</b>.</p></div>', unsafe_allow_html=True)
        
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
            st.markdown(f"<div class='card' style='border-left: 6px solid #8EA58C;'><b>&#128196; Sheet:</b> {sheet_name} &nbsp;<span class='badge neutral'>{len(items)} item</span></div>", unsafe_allow_html=True)
            
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
            _show_key = f"_show_all_{sheet_name}"
            _total_rows = len(df_items)
            _show_all = st.session_state.get(_show_key, False)
            if _total_rows > 10 and not _show_all:
                df_display = df_items.head(10)
            else:
                df_display = df_items
            try:
                st.table(df_display.style.set_properties(**{'text-align': 'center'}).set_table_styles([{'selector': 'th', 'props': [('text-align','center'),('background','#eef2ff'),('color','#3d5a3a')]}, {'selector': 'td', 'props': [('text-align','center')]}]))
            except Exception:
                st.dataframe(df_display, use_container_width=True, hide_index=True, height=min(520, 44+len(df_display)*36))
            if _total_rows > 10 and not _show_all:
                if st.button(f"&#128194; Tampilkan semua {_total_rows} baris", key=f"show_all_{sheet_name}"):
                    st.session_state[_show_key] = True
                    st.rerun()
            elif _total_rows > 10 and _show_all:
                if st.button(f"&#128065; Sembunyikan, tampilkan 10 baris saja", key=f"hide_{sheet_name}"):
                    st.session_state[_show_key] = False
                    st.rerun()

            # === PANEL DEBUG (tanpa perlu kirim gambar/file) ===
            # Simpan raw values di display debug &#8212; copy-paste teks ini ke chat
            with st.expander("&#128269; DEBUG &#8212; copy teks ini ke chat jika masih salah", expanded=False):
                st.caption("Fungsinya supaya saya bisa lihat nilai mentah Excel tanpa perlu foto.")
                
                # Kolom mapping yang terdeteksi
                excel_sheets_data_dbg = st.session_state.get('excel_sheets_data', {})
                sheet_dbg = excel_sheets_data_dbg.get(sheet_name, {})
                cols_dbg = {}
                # Coba baca dari item pertama atau simpan di excel_sheets_data kalau ada
                # Fallback: tampilkan kolom yang dipakai per item
                st.write("**Kolom terdeteksi (header &#8212; kolom):**")
                # Ambil dari excel_reader yang terakhir dipakai &#8212; simpan di session
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
                st.caption("&#128269; Select semua (Ctrl+A) di tabel atas &#8212; Ctrl+C &#8212; paste ke chat. Atau screenshot panel ini (lebih mudah dari foto Excel).")

                if st.button("&#128203; Copy debug sebagai teks", key=f"debug_copy_{sheet_name}"):
                    lines = []
                    lines.append(f"Sheet: {sheet_name} | Items: {len(items)}")
                    for r in debug_rows:
                        lines.append(f"Row {r['Row']:>3} | {r['Item']:<18} | qty_raw={r['qty_raw']} -> {r['qty']} | up_raw={r['unit_price_raw']} -> {r['unit_price']} | total_raw={r['total_raw']} -> {r['total(hasil calc)']} | qty*price={r['qty*price']} | mismatch={r['mismatch']}")
                    st.code("\n".join(lines), language="text")

                # Show skipped rows if any
                skipped = sheet_dbg.get('skipped_rows', [])
                if skipped:
                    st.warning(f"&#9888;&#65039; {len(skipped)} baris ter-skip (mungkin terdeteksi section/PPN). Detail:")
                    st.code("\n".join([f"Row {s['row']}: {s['dump']}" for s in skipped]), language="text")

                # Klasifikasi baris ringkasan (toleran typo) &#8212; angka tetap sumber kebenaran, tulisan hanya petunjuk
                klass = sheet_dbg.get('classifications', [])
                summary_dbg = sheet_dbg.get('summary_rows_debug', [])
                klass_map = {k['row']: k for k in summary_dbg} if summary_dbg else {}
                if klass:
                    st.caption("&#128269; Klasifikasi: jika typo, cek warning tapi tidak bikin error hitungan; hanya angka yang divalidasi.")
                    def _fmt_k(k):
                        v = klass_map.get(k['row'], {}).get('value', None)
                        try:
                            v_str = f" | value={float(v):,.0f}" if v is not None and safe_float(v) is not None else " | value=&#10060; tidak kebaca"
                        except:
                            v_str = " | value=&#10060; tidak kebaca"
                        return f"Row {k['row']}: '{k['raw']}' &#8212; {k['normalized']} &#8212; {k['type']}{(' &#9888;&#65039; typo' if k.get('fuzzy') else '')}{v_str}"
                    st.code("\n".join([_fmt_k(k) for k in klass]), language="text")
                else:
                    st.caption("&#128269; Klasifikasi: tidak ada baris ringkasan terdeteksi.")

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
              <div class="node"><b>Jumlah</b><br><span style="color:#64748b; font-size:.8rem;">Qty ?? Harga</span></div>
              <div class="arrow">&#8212;</div>
              <div class="node"><b>TOTAL</b><br><span style="color:#64748b; font-size:.8rem;">Jumlah A+B</span></div>
              <div class="arrow">&#8212;</div>
              <div class="node"><b>PPN 11%</b><br><span style="color:#64748b; font-size:.8rem;">TOTAL ?? 11%</span></div>
              <div class="arrow">&#8212;</div>
              <div class="node" style="border-color:#86efac; background:#ecfdf5;"><b>GRAND TOTAL</b><br><span style="color:#065f46; font-size:.8rem;">TOTAL + PPN</span></div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("<div style='text-align:center;'><span class='badge neutral'>RINGKASAN &#8212; DIHITUNG vs DI EXCEL</span></div>", unsafe_allow_html=True)
            # Header Ringkasan &#8212; center
            st.markdown("""
            <div style="background: linear-gradient(90deg, #10b981 0%, #6b8a68 50%, #059669 100%); 
                        color: white; 
                        padding: 1rem; 
                        border-radius: 16px; 
                        text-align: center;
                        margin: 1rem auto;
                        max-width: 760px;
                        box-shadow: 0 10px 30px rgba(37,99,235,.25);">
                <h3 style="margin: 0; color: white; font-weight: 800; font-size:1.05rem; text-align:center;">&#128202; RINGKASAN &#8212; Bandingkan DIHITUNG vs DI EXCEL</h3>
                <div style="font-size:.82rem; opacity:.95; margin-top:.2rem; text-align:center;">Kiri = hitungan sistem &nbsp;&#8212;&nbsp; Tengah = cocok/selisih &nbsp;&#8212;&nbsp; Kanan = angka di Excel</div>
            </div>
            """, unsafe_allow_html=True)
            
            # CEK APAKAH ADA MULTIPLE SECTIONS
            has_multiple_sections = len(sections) > 1
            _is_single_ppn_mode = sum(1 for sd in sections.values() if safe_float(sd.get('ppn_value')) is not None) == 1 and len(sections) > 1
            
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
                    
                    # Header Section &#8212; kategori vs section dijumlah terpisah
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
                        <h4 style="margin: 0; color: white;">&#128196; {"KATEGORI" if is_cat else "SECTION"} {section_letter}{label_suffix}</h4>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Subtotal Section (Jumlah X) &#8212; selalu tampil DIHITUNG (sebelum PPN), DI EXCEL tampil jika ada atau flag calculated
                    is_calc = section_data.get('subtotal_is_calculated', False)
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128200; JUMLAH (DIHITUNG)</div>
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
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212;</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                    </div>
                                    """.format(format_currency(difference)), unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039;</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            except:
                                pass
                        elif is_calc:
                            st.markdown("""
                            <div class="comparison-box" style="text-align: center; padding: 1rem; border: 1px dashed rgba(59,130,246,0.5); display:flex; flex-direction:column; align-items:center; justify-content:center;">
                                <div style="color: #93c5fd; font-size: 0.85rem;">Tidak ada Jumlah di Excel &#8212; pakai hitungan</div>
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
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128196; JUMLAH (DI EXCEL)</div>
                            <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(section_subtotal_excel if section_subtotal_excel is not None else section_calculated) if (section_subtotal_excel is not None or is_calc) else "-"), unsafe_allow_html=True)
                    
                    # PPN Section &#8212; skip jika PPN 1 bagian (sudah ditampilkan di Langkah 2 global)
                    has_section_ppn = section_ppn_excel is not None and not _is_single_ppn_mode
                    if has_section_ppn:
                        section_calculated_ppn = section_calculated * 0.11
                        col1, col2, col3 = st.columns([2, 1, 2])
                        with col1:
                            st.markdown("""
                            <div class="ppn-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128200; PPN 11% (DIHITUNG)</div>
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
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212;</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                    </div>
                                    """.format(format_currency(difference)), unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039;</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            except:
                                pass
                        with col3:
                            st.markdown("""
                            <div class="ppn-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128196; PPN (DI EXCEL)</div>
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
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128196; DISKON (DI EXCEL)</div>
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
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128196; DISKON (DI EXCEL)</div>
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
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128200; TOTAL SECTION (DIHITUNG)</div>
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
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212;</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                    </div>
                                    """.format(format_currency(difference)), unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039;</div>
                                        <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            except:
                                pass
                        with col3:
                            st.markdown("""
                            <div class="grandtotal-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128196; TOTAL SECTION (DI EXCEL)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(section_total_excel)), unsafe_allow_html=True)
                
                # === TOTAL KATEGORI &#8212; 3+ section: TOTAL (A+B+C) = sum sections, bukan Total A ===
                sheet_dbg_global = excel_sheets_data.get(sheet_name, {}) or sheet_data
                excel_ppn_global = sheet_dbg_global.get('ppn_value')
                is_combined_global = sheet_dbg_global.get('ppn_is_combined', False)
                has_any_section_ppn = any(sd.get('ppn_value') is not None for sd in sections.values())
                _is_single_ppn = sum(1 for sd in sections.values() if sd.get('ppn_value') is not None) == 1 and len(sections) > 1
                letters = "+".join(sorted(sections.keys()))
                if is_combined_global and len(sections) >= 2:
                    # PPN 1 di akhir (2 atau 3+): TOTAL &#8212; DI EXCEL harus TOTAL (A+B[+C]) dari Excel
                    # Priority: grand-PPN (kasus case 3sub) tapi jika blank combo, sum sections (case ada total masing) juga benar
                    g = safe_float(sheet_dbg_global.get('grand_total_value'))
                    pg = safe_float(sheet_dbg_global.get('ppn_value'))
                    s_sum = sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                    if g is not None and pg is not None and g > pg and abs((g-pg)-s_sum) > 1:
                        # Jika g-pg != sum, tetap pakai grand-PPN (yang benar untuk 3 sub 6.5jt)
                        total_kategori_excel = g - pg
                    else:
                        total_kategori_excel = s_sum
                    has_total_kategori = True
                    total_label = f"Jumlah {' + '.join(sorted(sections.keys()))}"
                else:
                    total_label = f"Jumlah {' + '.join(sorted(sections.keys()))}" if len(sections) <= 4 else f"Jumlah {len(sections)} bagian"
                    has_total_kategori = not _is_single_ppn and (sheet_dbg_global.get('jumlah_global_excel') is not None or sheet_dbg_global.get('subtotal_value') is not None or sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values()) > 0)
                    if sheet_dbg_global.get('jumlah_global_excel') is not None:
                        total_kategori_excel = safe_float(sheet_dbg_global.get('jumlah_global_excel'))
                    elif sheet_dbg_global.get('is_without_ppn') and sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values()) > 0:
                        total_kategori_excel = sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                    else:
                        total_kategori_excel = safe_float(sheet_dbg_global.get('subtotal_value'))
                if has_total_kategori and len(sections) > 1:
                    sum_sub_for_total = sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        letters_plus = "+".join(sorted(sections.keys()))
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">TOTAL &#8212; DIHITUNG</div>
                            <div style="font-size: 0.8rem; opacity:.85; margin-bottom:.3rem;">{}</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(total_label, format_currency(sum_sub_for_total)), unsafe_allow_html=True)
                    with col2:
                        try:
                            excel_val_t = float(total_kategori_excel)
                            diff_t = sum_sub_for_total - excel_val_t
                            if abs(diff_t) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212; SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(diff_t)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039; COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">TOTAL benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    with col3:
                        st.markdown("""
                        <div class="subtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">TOTAL &#8212; DI EXCEL (sebelum PPN)</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(total_kategori_excel)), unsafe_allow_html=True)

                # === Mode badge (auto-detect 5 case) ===
                _is_without = sheet_dbg_global.get('is_without_ppn', False) if isinstance(sheet_dbg_global, dict) else False
                _mode = "TANPA PPN" if _is_without else ("PPN GABUNGAN" if is_combined_global else ("PPN 1 BAGIAN" if sum(1 for sd in sections.values() if sd.get('ppn_value') is not None)==1 and len(sections)>1 else ("NORMAL" if len(sections)<=1 else "MULTI")))
                st.markdown(f"<div style='text-align:center; margin:.4rem 0;'><span class='badge neutral'>Mode: {_mode}</span></div>", unsafe_allow_html=True)
                _ppn_sec_count = sum(1 for sd in sections.values() if safe_float(sd.get('ppn_value')) is not None)
                _ppn_1bagian = _ppn_sec_count == 1 and len(sections) > 1 and not _is_without
                if not _is_without:
                    if _ppn_1bagian:
                        _ppn_letter = next((k for k, sd in sections.items() if safe_float(sd.get('ppn_value')) is not None), "?")
                        st.markdown(f"""
                        <div style="display:flex; align-items:center; justify-content:center; gap:.5rem; margin:1rem 0; flex-wrap:wrap;">
                          <span class="badge ok">1. Total A + Total B = TOTAL</span><span style="color:#94a3b8;">&#8212;</span>
                          <span class="badge neutral">2. Jumlah {_ppn_letter} ?? 11% = PPN</span><span style="color:#94a3b8;">&#8212;</span>
                          <span class="badge ok" style="background:#dcfce7; border-color:#86efac;">3. TOTAL (sudah termasuk PPN)</span>
                        </div>
                        """, unsafe_allow_html=True)
                    elif is_combined_global:
                        st.markdown("""
                        <div style="display:flex; align-items:center; justify-content:center; gap:.5rem; margin:1rem 0; flex-wrap:wrap;">
                          <span class="badge ok">1. Jumlah A+B = TOTAL</span><span style="color:#94a3b8;">&#8212;</span>
                          <span class="badge neutral">2. TOTAL ?? 11% = PPN</span><span style="color:#94a3b8;">&#8212;</span>
                          <span class="badge ok" style="background:#dcfce7; border-color:#86efac;">3. TOTAL + PPN = GRAND</span>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown("""
                        <div style="display:flex; align-items:center; justify-content:center; gap:.5rem; margin:1rem 0; flex-wrap:wrap;">
                          <span class="badge ok">1. Jumlah A+B = TOTAL</span><span style="color:#94a3b8;">&#8212;</span>
                          <span class="badge neutral">2. TOTAL ?? 11% = PPN</span><span style="color:#94a3b8;">&#8212;</span>
                          <span class="badge ok" style="background:#dcfce7; border-color:#86efac;">3. TOTAL + PPN = GRAND</span>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div style="display:flex; align-items:center; justify-content:center; gap:.5rem; margin:1rem 0; flex-wrap:wrap;">
                      <span class="badge ok">1. ?? Item = TOTAL</span><span style="color:#94a3b8;">&#8212;</span>
                      <span class="badge ok" style="background:#dcfce7; border-color:#86efac;">GRAND = TOTAL (tanpa PPN)</span>
                    </div>
                    """, unsafe_allow_html=True)

                is_single = len(sections) == 1
                has_ppn_section = has_any_section_ppn
                # Langkah 1 TOTAL Kategori mungkin belum ke-render tapi kita tetap butuh TOTAL untuk PPN
                # Jika ada TOTAL &#8212; DI EXCEL (sebelum PPN) yang tadi kelewat, hitung ulang di sini untuk PPN
                _total_for_ppn = safe_float(sheet_dbg_global.get('jumlah_global_excel'))
                if _total_for_ppn is None: _total_for_ppn = sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                show_global_ppn = not _is_without and (excel_ppn_global is not None and (is_combined_global or not has_ppn_section) or (is_single and has_ppn_section) or (not is_single and has_ppn_section and not is_combined_global) or (_total_for_ppn and _total_for_ppn > 0))
                _ppn_section_count = 0
                if show_global_ppn:
                    if is_single and has_ppn_section:
                        sd0 = list(sections.values())[0]
                        sum_sub_for_ppn = safe_float(sd0.get('subtotal_value')) or safe_float(sheet_dbg_global.get('jumlah_global_excel')) or 0
                        excel_ppn_global = safe_float(sd0.get('ppn_value')) or excel_ppn_global
                        is_combined_global = False
                    elif has_ppn_section and not is_combined_global:
                        # Hitung berapa section yang punya PPN
                        _ppn_section_count = sum(1 for sd in sections.values() if safe_float(sd.get('ppn_value')) is not None)
                        if _ppn_section_count == 1 and len(sections) > 1:
                            # PPN hanya di 1 section: gunakan subtotal section itu saja
                            _ppn_sec = next(sd for sd in sections.values() if safe_float(sd.get('ppn_value')) is not None)
                            sum_sub_for_ppn = safe_float(_ppn_sec.get('subtotal_value')) or 0
                            if excel_ppn_global is None:
                                excel_ppn_global = safe_float(_ppn_sec.get('ppn_value'))
                        else:
                            sum_sub_for_ppn = safe_float(sheet_dbg_global.get('jumlah_global_excel')) or sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                            if excel_ppn_global is None:
                                first_ppn = next((safe_float(sd.get('ppn_value')) for sd in sections.values() if sd.get('ppn_value') is not None), None)
                                excel_ppn_global = first_ppn
                    else:
                        sum_sub_for_ppn = safe_float(sheet_dbg_global.get('jumlah_global_excel')) or sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                    calc_ppn_global = sum_sub_for_ppn * 0.11 if sum_sub_for_ppn else 0
                    _ppn_from_section = ""
                    if _ppn_section_count == 1 and len(sections) > 1:
                        _ppn_sec_letter = next((k for k, sd in sections.items() if safe_float(sd.get('ppn_value')) is not None), "?")
                        _ppn_from_section = f" &#8212; dari Jumlah {_ppn_sec_letter}"
                    st.markdown(f"""
                    <div class="card" style="border:2px solid #fed7aa; background: linear-gradient(180deg, #fffbeb, #ffffff);">
                      <div style="display:flex; align-items:center; gap:.6rem; margin-bottom:.6rem;">
                        <span style="background:#f97316; color:white; border-radius:8px; padding:.3rem .6rem; font-weight:800;">Langkah 2</span>
                        <b style="font-size:1.05rem;">PPN 11%{_ppn_from_section}</b>
                        <span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: Jumlah ?? 11%</span>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="ppn-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">PPN (DIHITUNG)</div>
                            <div style="font-size: 0.8rem; opacity:.85; margin-bottom:.3rem;">TOTAL ?? 11%</div>
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
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212; SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(diff2)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039; COCOK</div>
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
                    # Sinkron untuk Grand Total
                    excel_grand_total = sheet_dbg_global.get('grand_total_value')

                # Fallback Normal umum: jika ada PPN di section (PPN 11% 1,344,420 di A) tapi show_global_ppn False, tampilkan Langkah 2
                _sheet_fallback_ppn = False
                if not show_global_ppn and has_any_section_ppn and not sheet_dbg_global.get('is_without_ppn', False):
                    # Ambil PPN pertama yang ada (NORMAL / PPN 1 BAGIAN)
                    first_ppn = next((safe_float(sd.get('ppn_value')) for sd in sections.values() if sd.get('ppn_value') is not None), None)
                    first_sub = next((safe_float(sd.get('subtotal_value')) for sd in sections.values() if sd.get('subtotal_value') is not None), 0)
                    if first_ppn is not None:
                        excel_ppn_global = first_ppn
                        sum_sub_for_ppn = first_sub
                        calc_ppn_global = sum_sub_for_ppn * 0.11 if sum_sub_for_ppn else first_ppn
                        show_global_ppn = True
                        _sheet_fallback_ppn = True
                        st.markdown("""
                        <div class="card" style="border:2px solid #fed7aa; background: linear-gradient(180deg, #fffbeb, #ffffff);">
                          <div style="display:flex; align-items:center; gap:.6rem; margin-bottom:.6rem;">
                            <span style="background:#f97316; color:white; border-radius:8px; padding:.3rem .6rem; font-weight:800;">Langkah 2</span>
                            <b style="font-size:1.05rem;">PPN 11% &#8212; dari Jumlah</b>
                            <span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: Jumlah ?? 11%</span>
                          </div>
                        </div>
                        """, unsafe_allow_html=True)
                        col1, col2, col3 = st.columns([2, 1, 2])
                        with col1:
                            st.markdown(f"""
                            <div class="ppn-box">
                                <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">PPN (DIHITUNG)</div>
                                <div style="font-size: 0.8rem; opacity:.85; margin-bottom:.3rem;">Jumlah ?? 11%</div>
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
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212; SELISIH</div>
                                        <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{format_currency(d)}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="sesuai-box">
                                        <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039; COCOK</div>
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
                      <div style="display:flex; gap:.6rem; align-items:center;"><span style="background:#059669; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 3</span><b>Grand Total &#8212; sudah termasuk PPN</b><span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: TOTAL + PPN</span></div>
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
                    # Jika PPN gabungan (bukan per-section), tambahkan PPN ke grand total
                    # Untuk PPN 1 bagian: section total sudah termasuk PPN, jangan tambah lagi
                    _ppn_sec_cnt_gt = sum(1 for sd in sections.values() if safe_float(sd.get('ppn_value')) is not None)
                    _is_ppn_1bagian = _ppn_sec_cnt_gt == 1 and len(sections) > 1
                    if show_global_ppn and not _is_ppn_1bagian:
                        sum_sub_for_gt = sum(safe_float(sd.get('subtotal_value')) or 0 for sd in sections.values())
                        calculated_grand_total = sum_sub_for_gt + (sum_sub_for_gt * 0.11)
                    elif has_any_section_ppn:
                        # grand total sudah sum dari section yang sudah termasuk PPN per-section
                        pass
                    
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="grandtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total &#8212; DIHITUNG</div>
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
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212; SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039; COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">Grand benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    with col3:
                        st.markdown("""
                        <div class="grandtotal-box">
                            <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total &#8212; DI EXCEL</div>
                            <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(excel_grand_total)), unsafe_allow_html=True)
                    st.markdown("<div style='height:.6rem;'></div>", unsafe_allow_html=True)
            
            else:
                _sec_ppn = None
                if len(sections) == 1:
                    for _sd in sections.values():
                        _sp = safe_float(_sd.get('ppn_value'))
                        if _sp is not None and _sp != 0:
                            _sec_ppn = _sp
                            break
                has_ppn_single = excel_ppn is not None or _sec_ppn is not None
                _effective_ppn = excel_ppn if excel_ppn is not None else _sec_ppn
                calculated_ppn = calculated_total_items * 0.11 if has_ppn_single else 0
                calculated_grand_total = calculated_total_items + calculated_ppn

                # Judul langkah single
                st.markdown("""
                <div class="card" style="border:2px solid #bfdbfe;">
                  <div style="display:flex; gap:.6rem; align-items:center; flex-wrap:wrap;">
                    <span style="background:#10b981; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 1</span>
                    <b>Jumlah (sebelum PPN)</b> <span style="color:#64748b; font-size:.85rem;">&#8212; penjumlahan semua item</span>
                    <span style="margin-left:auto; background:#f1f5f9; border-radius:999px; padding:.25rem .6rem; font-size:.78rem; color:#475569;"><b>Rumus:</b> ?? Qty ?? Harga</span>
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
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212; SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039; COCOK</div>
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
                
                # PPN single &#8212; kartu Langkah 2
                if has_ppn_single:
                    st.markdown("""
                    <div class="card" style="border:2px solid #fed7aa; background: linear-gradient(180deg, #fffbeb, #ffffff);">
                      <div style="display:flex; gap:.6rem; align-items:center;"><span style="background:#f97316; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 2</span><b>PPN 11%</b><span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: Jumlah ?? 11%</span></div>
                    </div>
                    """, unsafe_allow_html=True)
                    col1, col2, col3 = st.columns([2, 1, 2])
                    with col1:
                        st.markdown("""
                        <div class="ppn-box">
                            <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128200; PPN 11% (DIHITUNG)</div>
                            <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                        </div>
                        """.format(format_currency(calculated_ppn)), unsafe_allow_html=True)
                    with col2:
                        try:
                            excel_val = float(_effective_ppn)
                            difference = calculated_ppn - excel_val
                            if abs(difference) > 1:
                                st.markdown("""
                                <div class="selisih-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212;</div>
                                    <div style="font-weight: 700; font-size: 0.9rem;">SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.2rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039;</div>
                                    <div style="font-weight: 700; font-size: 0.9rem;">SESUAI</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    with col3:
                        try:
                            excel_val = float(_effective_ppn)
                            st.markdown("""
                            <div class="ppn-box">
                                <div style="font-size: 0.85rem; opacity: 0.9; margin-bottom: 0.5rem;">&#128196; PPN (DI EXCEL)</div>
                                <div style="font-size: 1.6rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(excel_val)), unsafe_allow_html=True)
                        except:
                            pass
                else:
                    st.markdown("""
                    <div class="card" style="border:2px solid #e5e7eb; background: linear-gradient(180deg, #f9fafb, #ffffff);">
                      <div style="display:flex; gap:.6rem; align-items:center;"><span style="background:#9ca3af; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 2</span><b>Tidak ada PPN</b><span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">File ini tidak mengandung PPN</span></div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("""
                <div class="card" style="border:2px solid #86efac; background: linear-gradient(180deg, #ecfdf5, #ffffff);">
                  <div style="display:flex; gap:.6rem; align-items:center;"><span style="background:#059669; color:white; border-radius:8px; padding:.35rem .7rem; font-weight:800;">Langkah 3</span><b>Grand Total &#8212; sudah termasuk PPN</b><span style="margin-left:auto; color:#9ca3af; font-size:.8rem;">Rumus: Jumlah + PPN</span></div>
                </div>
                """, unsafe_allow_html=True)
                # Grand Total
                col1, col2, col3 = st.columns([2, 1, 2])
                with col1:
                    st.markdown("""
                    <div class="grandtotal-box">
                        <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total &#8212; DIHITUNG</div>
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
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#8212; SELISIH</div>
                                    <div style="font-weight: 800; font-size: 1.1rem; margin-top: 0.3rem;">{}</div>
                                </div>
                                """.format(format_currency(difference)), unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="sesuai-box">
                                    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">&#10004;&#65039; COCOK</div>
                                    <div style="font-weight: 700; font-size: .85rem; opacity:.9;">Grand benar</div>
                                </div>
                                """, unsafe_allow_html=True)
                        except:
                            pass
                    else:
                        st.markdown("""
                        <div class="comparison-box" style="text-align: center; padding: 1.5rem;">
                            <div style="font-size: 1.2rem; color: #9ca3af;">&#128196; Tidak ada data Grand Total</div>
                        </div>
                        """, unsafe_allow_html=True)
                with col3:
                    if excel_grand_total is not None:
                        try:
                            excel_val = float(excel_grand_total)
                            st.markdown("""
                            <div class="grandtotal-box">
                                <div style="font-size: 0.95rem; font-weight:700; opacity: 0.95; margin-bottom: 0.4rem;">Grand Total &#8212; DI EXCEL</div>
                                <div style="font-size: 1.7rem; font-weight: 800;">{}</div>
                            </div>
                            """.format(format_currency(excel_val)), unsafe_allow_html=True)
                        except:
                            pass
            
            st.markdown("---")
    
    pass  # Download RAB AUDIT REPORT removed per user request

    st.markdown("<div style='height:.3rem;'></div>", unsafe_allow_html=True)
    
    # Detail Errors &#8212; RAB AUDIT REPORT style (lokasi, Excel, Seharusnya, Selisih)
    if errors:
        st.markdown('<div class="section-header">&#128202; RAB AUDIT REPORT &#8212; DETAIL TEMUAN</div>', unsafe_allow_html=True)
        
        for i, error in enumerate(errors, 1):
            item_name = error.get('item_name', 'Unknown')
            row = error.get('row', '?')
            error_type = error.get('type', '')
            
            with st.expander(f"&#8212; Error {i}: {item_name} - Baris {row}", expanded=True):
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
                    st.markdown("**&#128203; Detail Item yang Terbaca:**")
                    st.code(error.get('items_summary', ''), language=None)
    
    # Detail Warnings
    if warnings:
        st.markdown('<div class="section-header">PERINGATAN</div>', unsafe_allow_html=True)
        
        for i, warning in enumerate(warnings, 1):
            item_name = warning.get('item_name', 'Unknown')
            row = warning.get('row', '?')
            
            with st.expander(f"&#9888;&#65039; Warning {i}: {item_name} - Baris {row}"):
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
        if st.button("&#128230; Export Report (RAB_CHECK_REPORT.xlsx)", type="primary", use_container_width=True):
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
                label="&#128229; Download Report",
                data=buffer,
                file_name="RAB_CHECK_REPORT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
