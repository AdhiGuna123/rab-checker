import pandas as pd
from datetime import datetime
from typing import Dict, List, Any, Optional
import os
import io

class ReportGenerator:
    """Class untuk generate laporan hasil pemeriksaan — RAB AUDIT REPORT (hanya laporan, tidak ubah file)"""
    
    def __init__(self):
        self.report_data = []
    
    def _fmt_rp(self, v: Any) -> str:
        if v is None or v == '':
            return "-"
        try:
            n = float(v)
            return f"Rp {n:,.0f}".replace(',', '.')
        except:
            return str(v)

    def build_bytes(self, file_name: str, check_results: Dict[str, Any], errors: List[Dict[str, Any]], warnings: List[Dict[str, Any]], sheets_data: Optional[Dict[str, Any]] = None) -> bytes:
        """Build RAB AUDIT REPORT xlsx sebagai bytes untuk st.download_button."""
        sheets_data = sheets_data or {}
        summary_data = {
            'Item': ['Nama File', 'Waktu Pemeriksaan', 'Jumlah Sheet', 'Jumlah Bagian/Sub Bagian', 'Jumlah Item', 'Temuan (Error)', 'Warning', 'Status'],
            'Value': [
                file_name,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                len(sheets_data) if sheets_data else (1 if check_results.get('total_items', 0) else 0),
                ", ".join([f"{k}: {', '.join(sorted(v.get('sections',{}).keys())) if v.get('sections') else '-'}" for k, v in sheets_data.items()]) if sheets_data else "-",
                check_results.get('total_items', 0),
                check_results.get('total_errors', 0),
                check_results.get('total_warnings', 0),
                'OK' if check_results.get('total_errors', 0) == 0 else 'PERLU CEK'
            ]
        }
        df_summary = pd.DataFrame(summary_data)
        lock_note = pd.DataFrame([{'Catatan': 'Sistem TIDAK mengubah file asli. Hanya laporan pengecekan matematika & logika Excel (Qty×Price, TOTAL, PPN, GRAND TOTAL).'}])
        error_rows = []
        for i, error in enumerate(errors, 1):
            error_rows.append({
                'No': i,
                'Jenis': error.get('type', ''),
                'Lokasi': f"{error.get('sheet','')} Baris {error.get('row','')}",
                'Item': error.get('item_name', ''),
                'Masalah': error.get('detail', ''),
                'Perhitungan': error.get('calculation', ''),
                'Nilai Excel': self._fmt_rp(error.get('actual')),
                'Seharusnya': self._fmt_rp(error.get('expected')),
                'Selisih': self._fmt_rp(error.get('difference')),
                'Status': error.get('status', '')
            })
        df_errors = pd.DataFrame(error_rows) if error_rows else pd.DataFrame(columns=['No','Jenis','Lokasi','Item','Masalah','Perhitungan','Nilai Excel','Seharusnya','Selisih','Status'])
        warning_rows = []
        for i, warning in enumerate(warnings, 1):
            warning_rows.append({
                'No': i, 'Jenis': warning.get('type',''), 'Sheet': warning.get('sheet',''),
                'Baris': warning.get('row',''), 'Item': warning.get('item_name',''),
                'Masalah': warning.get('detail',''), 'Nilai': warning.get('value',''), 'Status': warning.get('status','')
            })
        df_warnings = pd.DataFrame(warning_rows) if warning_rows else pd.DataFrame(columns=['No','Jenis','Sheet','Baris','Item','Masalah','Nilai','Status'])
        # Sheet struktur (auto detection)
        struct_rows = []
        for sn, sd in (sheets_data or {}).items():
            secs = sorted(sd.get('sections',{}).keys()) if isinstance(sd.get('sections'), dict) else []
            struct_rows.append({'Sheet': sn, 'Jumlah Bagian': len(secs), 'Bagian': ", ".join(secs) if secs else "-", 'Kata Kunci Terdeteksi': ", ".join([k for k in ['TOTAL','JUMLAH','SUBTOTAL','PPN','GRAND TOTAL'] if any(k.lower() in str(c).lower() for c in sd.get('classifications',[]) if isinstance(c, dict))]) or "-"})
        df_struct = pd.DataFrame(struct_rows) if struct_rows else pd.DataFrame(columns=['Sheet','Jumlah Bagian','Bagian','Kata Kunci Terdeteksi'])
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='RAB AUDIT REPORT', startrow=0, index=False)
            lock_note.to_excel(writer, sheet_name='RAB AUDIT REPORT', startrow=len(df_summary)+3, index=False)
            df_struct.to_excel(writer, sheet_name='Struktur', index=False)
            df_errors.to_excel(writer, sheet_name='Temuan', index=False)
            df_warnings.to_excel(writer, sheet_name='Warnings', index=False)
            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                for column in ws.columns:
                    max_len = 0
                    col_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value or "")) > max_len:
                                max_len = len(str(cell.value))
                        except:
                            pass
                    ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
        return bio.getvalue()

    def generate_report(self, file_name: str, check_results: Dict[str, Any], errors: List[Dict[str, Any]], warnings: List[Dict[str, Any]]) -> pd.ExcelWriter:
        summary_data = {'Item': ['Nama File', 'Waktu Pemeriksaan', 'Jumlah Item', 'Jumlah Error', 'Jumlah Warning', 'Status'], 'Value': [file_name, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), check_results.get('total_items', 0), check_results.get('total_errors', 0), check_results.get('total_warnings', 0), 'OK' if check_results.get('total_errors', 0) == 0 else 'PERLU CEK']}
        df_summary = pd.DataFrame(summary_data)
        error_rows = []
        for i, error in enumerate(errors, 1):
            error_rows.append({'No': i, 'Tipe': error.get('type', ''), 'Sheet': error.get('sheet', ''), 'Baris': error.get('row', ''), 'Item': error.get('item_name', ''), 'Masalah': error.get('detail', ''), 'Perhitungan': error.get('calculation', ''), 'Nilai Excel': error.get('actual', ''), 'Seharusnya': error.get('expected', ''), 'Selisih': error.get('difference', ''), 'Status': error.get('status', '')})
        df_errors = pd.DataFrame(error_rows) if error_rows else pd.DataFrame(columns=['No', 'Tipe', 'Sheet', 'Baris', 'Item', 'Masalah', 'Perhitungan', 'Nilai Excel', 'Seharusnya', 'Selisih', 'Status'])
        warning_rows = []
        for i, warning in enumerate(warnings, 1):
            warning_rows.append({'No': i, 'Tipe': warning.get('type', ''), 'Sheet': warning.get('sheet', ''), 'Baris': warning.get('row', ''), 'Item': warning.get('item_name', ''), 'Masalah': warning.get('detail', ''), 'Nilai': warning.get('value', ''), 'Status': warning.get('status', '')})
        df_warnings = pd.DataFrame(warning_rows) if warning_rows else pd.DataFrame(columns=['No', 'Tipe', 'Sheet', 'Baris', 'Item', 'Masalah', 'Nilai', 'Status'])
        return df_summary, df_errors, df_warnings
    
    def save_report(self, file_name: str, check_results: Dict[str, Any], errors: List[Dict[str, Any]], warnings: List[Dict[str, Any]], output_path: str = None) -> str:
        if output_path is None:
            base_name = os.path.splitext(file_name)[0]
            output_path = f"{base_name}_RAB_CHECK_REPORT.xlsx"
        df_summary, df_errors, df_warnings = self.generate_report(file_name, check_results, errors, warnings)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            df_errors.to_excel(writer, sheet_name='Errors', index=False)
            df_warnings.to_excel(writer, sheet_name='Warnings', index=False)
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
        return output_path
    
    def format_currency(self, value: float) -> str:
        if value is None:
            return "Rp 0"
        return f"Rp {value:,.0f}".replace(',', '.')
