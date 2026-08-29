"""AI helper gratis opsional — Value Intelligence gratis default, AI hanya fallback klasifikasi label.
- Gemini 1.5 Flash free tier (1.500 req/hari, tanpa kartu untuk free tier) via GOOGLE_API_KEY env atau input di UI.
- Groq free tier opsional via GROQ_API_KEY.
- 100% aman tanpa AI: jika tidak ada key, hanya pakai value+label intelligence (gratis, offline).
AI TIDAK menghitung — hanya klasifikasi baris ('JUMLAH', 'TOTAL', 'PPN', 'GRAND TOTAL') agar TOTAL sebelum PPN tidak miss.
Hitungan tetap di checker.py qty*price & subtotal (deterministik).
"""
import os
import json
from typing import List, Dict, Optional

SYSTEM_PROMPT = (
    "Klasifikasikan setiap baris ringkasan RAB ke salah satu: "
    "'jumlah_global' (JUMLAH/TOTAL sebelum PPN, tanpa huruf section), "
    "'ppn' (PPN 11% / PAJAK), 'grand_total' (GRAND TOTAL / TOTAL A+B), "
    "'subtotal' (Jumlah A, Total B, dsb dengan huruf section), 'discount', 'unknown'. "
    "Output JSON array: [{\"row\": int, \"type\": str}]. "
    "Aturan: 'TOTAL' polos sebelum PPN dan GRAND TOTAL adalah 'jumlah_global', bukan grand_total. "
    "Abaikan notes panjang seperti '2. Penawaran sudah termasuk PPN' -> unknown."
)

def classify_with_gemini_free(rows: List[Dict], api_key: Optional[str] = None) -> Optional[Dict[int, str]]:
    """Gemini 1.5 Flash free tier. Return dict row->type atau None jika gagal."""
    key = api_key or os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
    if not key:
        return None
    try:
        # Lazy import agar tidak wajib install jika tidak pakai AI
        import requests
        payload_rows = "\n".join([f"Row {r['row']}: '{r['raw']}' | value={r.get('value')}" for r in rows])
        body = {
            "contents": [{"parts": [{"text": SYSTEM_PROMPT + "\n\nRows:\n" + payload_rows}]}],
            "generationConfig": {"temperature": 0.1, "maxOutputTokens": 512}
        }
        resp = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}",
            json=body, timeout=10
        )
        if resp.status_code != 200:
            return None
        data = resp.json()
        text = data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
        # Extract JSON array
        start = text.find("[")
        end = text.rfind("]") + 1
        if start == -1 or end == 0:
            return None
        arr = json.loads(text[start:end])
        return {int(x["row"]): x["type"] for x in arr if "row" in x and "type" in x}
    except Exception:
        return None

def classify_with_groq_free(rows: List[Dict], api_key: Optional[str] = None) -> Optional[Dict[int, str]]:
    key = api_key or os.environ.get("GROQ_API_KEY")
    if not key:
        return None
    try:
        import requests
        payload_rows = "\n".join([f"Row {r['row']}: '{r['raw']}'" for r in rows])
        body = {
            "model": "llama-3.1-8b-instant",
            "messages": [{"role": "user", "content": SYSTEM_PROMPT + "\n\nRows:\n" + payload_rows}],
            "temperature": 0.1, "max_tokens": 512
        }
        resp = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            json=body, headers={"Authorization": f"Bearer {key}"}, timeout=10
        )
        if resp.status_code != 200:
            return None
        text = resp.json()["choices"][0]["message"]["content"]
        start = text.find("[")
        end = text.rfind("]") + 1
        if start == -1 or end == 0:
            return None
        arr = json.loads(text[start:end])
        return {int(x["row"]): x["type"] for x in arr if "row" in x and "type" in x}
    except Exception:
        return None
