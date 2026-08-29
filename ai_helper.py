"""AI untuk 5 CASE + beyond — gratis selamanya (Gemini free tier 1.500/hari via GOOGLE_API_KEY).
AI HANYA klasifikasi label baris ringkasan (TOTAL/JUMLAH/PPN/GRAND), tidak hitung.
Hitungan tetap checker.py Qty×Price & TOTAL/PPN/GRAND (deterministik, tidak ubah file, hanya laporan).
Fallback Value Intelligence gratis jika tanpa key.
"""
import os
import json
from typing import List, Dict, Optional

SYSTEM_PROMPT = (
    "Kamu adalah classifier baris ringkasan RAB (Rencana Anggaran Biaya). "
    "Klasifikasikan setiap baris ke salah satu:\n"
    "- 'jumlah_global': JUMLAH/TOTAL/SUBTOTAL tanpa huruf section, sebelum PPN (mis. TOTAL, JUMLAH, TOTAL A+B+C adalah gabungan jumlah)\n"
    "- 'subtotal': Jumlah A, Total B, Total C, Subtotal A, dsb — dengan huruf section A/B/C/...\n"
    "- 'ppn': PPN 11%, Pajak, PPN Global, PPN A, PPN B\n"
    "- 'grand_total': GRAND TOTAL, Grand Total (A+B), TOTAL AKHIR\n"
    "- 'discount': Diskon, Potongan\n"
    "- 'section_header': A, B, C (kategori/pemisah)\n"
    "- 'unknown': notes, syarat, alamat, bukan ringkasan\n"
    "Aturan penting:\n"
    "- 'TOTAL' polos sebelum PPN dan GRAND TOTAL adalah 'jumlah_global', bukan grand_total\n"
    "- 'TOTAL (A+B)' sebelum PPN adalah 'jumlah_global', 'GRAND TOTAL' setelah PPN adalah 'grand_total'\n"
    "- Abaikan notes panjang seperti '2. Penawaran sudah termasuk PPN 11%' -> 'unknown'\n"
    "- Keluaran JSON array: [{\"row\": int, \"type\": str}] hanya type di daftar di atas.\n"
    "- Harus menangani 5 case dan beyond: tanpa PPN, PPN 1 bagian, normal, PPN gabungan 2/3/n bagian, dinamis.\n"
    "Input rows berupa: Row N: 'raw text' | value=angka | normalized=xxx\n"
)

def classify_with_gemini_free(rows: List[Dict], api_key: Optional[str] = None) -> Optional[Dict[int, str]]:
    """Gemini 1.5 Flash free tier selamanya."""
    key = api_key or os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
    if not key:
        return None
    try:
        import requests
        payload_rows = "\n".join([f"Row {r['row']}: '{r['raw']}' | normalized={r.get('normalized','')} | value={r.get('value')}" for r in rows])
        body = {
            "contents": [{"parts": [{"text": SYSTEM_PROMPT + "\n\nRows:\n" + payload_rows}]}],
            "generationConfig": {"temperature": 0.1, "maxOutputTokens": 1024}
        }
        resp = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}",
            json=body, timeout=12
        )
        if resp.status_code != 200:
            return None
        data = resp.json()
        text = data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
        start = text.find("[")
        end = text.rfind("]") + 1
        if start == -1 or end == 0:
            return None
        arr = json.loads(text[start:end])
        allowed = {'jumlah_global','ppn','grand_total','subtotal','discount','unknown','section_header'}
        return {int(x["row"]): x["type"] for x in arr if "row" in x and "type" in x and x["type"] in allowed}
    except Exception:
        return None

def classify_with_groq_free(rows: List[Dict], api_key: Optional[str] = None) -> Optional[Dict[int, str]]:
    key = api_key or os.environ.get("GROQ_API_KEY")
    if not key:
        return None
    try:
        import requests
        payload_rows = "\n".join([f"Row {r['row']}: '{r['raw']}' | value={r.get('value')}" for r in rows])
        body = {
            "model": "llama-3.1-8b-instant",
            "messages": [{"role": "user", "content": SYSTEM_PROMPT + "\n\nRows:\n" + payload_rows}],
            "temperature": 0.1, "max_tokens": 1024
        }
        resp = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            json=body, headers={"Authorization": f"Bearer {key}"}, timeout=12
        )
        if resp.status_code != 200:
            return None
        text = resp.json()["choices"][0]["message"]["content"]
        start = text.find("[")
        end = text.rfind("]") + 1
        if start == -1 or end == 0:
            return None
        arr = json.loads(text[start:end])
        allowed = {'jumlah_global','ppn','grand_total','subtotal','discount','unknown','section_header'}
        return {int(x["row"]): x["type"] for x in arr if "row" in x and "type" in x and x["type"] in allowed}
    except Exception:
        return None
