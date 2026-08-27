import pandas as pd
from datetime import datetime
from typing import Dict, List, Any
import os

class ReportGenerator:
    """Class untuk generate laporan hasil pemeriksaan"""
    
    def __init__(self):
        self.report_data = []
        
    def generate_report(self, 
                       file_name: str,
                       check_results: Dict[str, Any],
                       errors: List[Dict[str, Any]],
                       warnings: List[Dict[str, Any]]) -> pd.ExcelWriter:
        """Generate laporan dalam format Excel"""
        
        # Buat DataFrame untuk summary
        summary_data = {
            'Item': ['Nama File', 'Waktu Pemeriksaan', 'Jumlah Item', 'Jumlah Error', 'Jumlah Warning', 'Status'],
            'Value': [
                file_name,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                check_results.get('total_items', 0),
                check_results.get('total_errors', 0),
                check_results.get('total_warnings', 0),
                'OK' if check_results.get('total_errors', 0) == 0 else 'PERLU CEK'
            ]
        }
        df_summary = pd.DataFrame(summary_data)
        
        # Buat DataFrame untuk detail errors
        error_rows = []
        for i, error in enumerate(errors, 1):
            error_rows.append({
                'No': i,
                'Tipe': error.get('type', ''),
                'Sheet': error.get('sheet', ''),
                'Baris': error.get('row', ''),
                'Item': error.get('item_name', ''),
                'Masalah': error.get('detail', ''),
                'Perhitungan': error.get('calculation', ''),
                'Nilai Excel': error.get('actual', ''),
                'Seharusnya': error.get('expected', ''),
                'Selisih': error.get('difference', ''),
                'Status': error.get('status', '')
            })
        df_errors = pd.DataFrame(error_rows) if error_rows else pd.DataFrame(columns=[
            'No', 'Tipe', 'Sheet', 'Baris', 'Item', 'Masalah', 
            'Perhitungan', 'Nilai Excel', 'Seharusnya', 'Selisih', 'Status'
        ])
        
        # Buat DataFrame untuk warnings
        warning_rows = []
        for i, warning in enumerate(warnings, 1):
            warning_rows.append({
                'No': i,
                'Tipe': warning.get('type', ''),
                'Sheet': warning.get('sheet', ''),
                'Baris': warning.get('row', ''),
                'Item': warning.get('item_name', ''),
                'Masalah': warning.get('detail', ''),
                'Nilai': warning.get('value', ''),
                'Status': warning.get('status', '')
            })
        df_warnings = pd.DataFrame(warning_rows) if warning_rows else pd.DataFrame(columns=[
            'No', 'Tipe', 'Sheet', 'Baris', 'Item', 'Masalah', 'Nilai', 'Status'
        ])
        
        return df_summary, df_errors, df_warnings
    
    def save_report(self, 
                   file_name: str,
                   check_results: Dict[str, Any],
                   errors: List[Dict[str, Any]],
                   warnings: List[Dict[str, Any]],
                   output_path: str = None) -> str:
        """Simpan laporan ke file Excel"""
        
        if output_path is None:
            # Generate output filename
            base_name = os.path.splitext(file_name)[0]
            output_path = f"{base_name}_RAB_CHECK_REPORT.xlsx"
        
        df_summary, df_errors, df_warnings = self.generate_report(
            file_name, check_results, errors, warnings
        )
        
        # Simpan ke Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
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
        
        return output_path
    
    def format_currency(self, value: float) -> str:
        """Format angka sebagai mata uang Rupiah"""
        if value is None:
            return "Rp 0"
        return f"Rp {value:,.0f}".replace(',', '.')
