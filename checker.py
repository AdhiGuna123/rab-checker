from typing import Dict, List, Any, Optional
from excel_reader import ExcelReader

class RABChecker:
    """Class untuk melakukan pemeriksaan RAB"""
    
    def __init__(self, excel_reader: ExcelReader):
        self.reader = excel_reader
        self.errors = []
        self.warnings = []
        
    def check_all(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Melakukan semua pemeriksaan"""
        self.errors = []
        self.warnings = []
        
        # 1. Cek perkalian setiap item
        self.check_multiplication(data)
        
        # 2. Cek formula Excel
        self.check_formulas(data)
        
        # 3. Cek data kosong
        self.check_empty_data(data)
        
        # 4. Cek subtotal
        self.check_subtotal(data)
        
        # 5. Cek PPN
        self.check_ppn(data)
        
        # 6. Cek grand total
        self.check_grand_total(data)
        
        return {
            'errors': self.errors,
            'warnings': self.warnings,
            'items': data.get('items', []),
            'total_items': len(data.get('items', [])),
            'total_errors': len(self.errors),
            'total_warnings': len(self.warnings)
        }
    
    def check_multiplication(self, data: Dict[str, Any]) -> None:
        """Cek apakah Qty × Unit Price = Total"""
        items = data.get('items', [])
        columns = data.get('columns', {})
        total_col = columns.get('total', 10)
        
        for item in items:
            row = item.get('row')
            qty = item.get('qty')
            unit_price = item.get('unit_price')
            total_excel = item.get('total')
            
            # Skip jika ada formula (sudah dicek di check_formulas)
            if item.get('has_formula'):
                continue
            
            # Validasi nilai
            if qty is None or unit_price is None or total_excel is None:
                continue
                
            # Pastikan nilai numerik
            try:
                qty_num = float(qty) if qty else 0
                unit_price_num = float(unit_price) if unit_price else 0
                total_num = float(total_excel) if total_excel else 0
            except (ValueError, TypeError):
                continue
            
            # Hitung total seharusnya
            expected_total = qty_num * unit_price_num
            
            # Bandingkan dengan toleransi
            tolerance = 1  # toleransi 1 rupiah
            if abs(expected_total - total_num) > tolerance:
                self.errors.append({
                    'type': 'MULTIPLICATION_ERROR',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': f'Total tidak sesuai',
                    'calculation': f'Qty ({qty_num}) × Unit Price ({unit_price_num})',
                    'expected': expected_total,
                    'actual': total_num,
                    'difference': expected_total - total_num,
                    'status': 'PERLU CEK'
                })
    
    def check_formulas(self, data: Dict[str, Any]) -> None:
        """Cek apakah kolom Total menggunakan formula"""
        items = data.get('items', [])
        columns = data.get('columns', {})
        total_col = columns.get('total', 10)
        
        for item in items:
            row = item.get('row')
            has_formula = item.get('has_formula', False)
            total_value = item.get('total')
            
            # Jika total ada tapi tidak ada formula
            if total_value is not None and not has_formula:
                self.warnings.append({
                    'type': 'FORMULA_MISSING',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': f'Formula hilang pada kolom Total',
                    'value': total_value,
                    'status': 'PERLU CEK'
                })
    
    def check_empty_data(self, data: Dict[str, Any]) -> None:
        """Cek data kosong atau tidak valid"""
        items = data.get('items', [])
        
        for item in items:
            row = item.get('row')
            qty = item.get('qty')
            unit_price = item.get('unit_price')
            total = item.get('total')
            
            # Cek qty kosong
            if qty is None or (isinstance(qty, str) and qty.strip() == ''):
                self.errors.append({
                    'type': 'EMPTY_QTY',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': 'Qty kosong',
                    'status': 'PERLU CEK'
                })
            
            # Cek unit price kosong
            if unit_price is None or (isinstance(unit_price, str) and unit_price.strip() == ''):
                self.errors.append({
                    'type': 'EMPTY_UNIT_PRICE',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': 'Unit Price kosong',
                    'status': 'PERLU CEK'
                })
            
            # Cek total kosong
            if total is None or (isinstance(total, str) and total.strip() == ''):
                self.errors.append({
                    'type': 'EMPTY_TOTAL',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': 'Total kosong',
                    'status': 'PERLU CEK'
                })
            
            # Cek angka terbaca sebagai teks
            if isinstance(qty, str) and qty.strip().isdigit():
                self.warnings.append({
                    'type': 'TEXT_AS_NUMBER',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': 'Qty terbaca sebagai teks',
                    'value': qty,
                    'status': 'PERLU CEK'
                })
    
    def check_subtotal(self, data: Dict[str, Any]) -> None:
        """Cek subtotal"""
        items = data.get('items', [])
        subtotal_value = data.get('subtotal_value')
        subtotal_row = data.get('subtotal_row')
        
        # Hitung total dari semua item
        calculated_total = 0
        item_count = 0
        item_details = []
        
        for item in items:
            total = item.get('total')
            if total is not None:
                try:
                    total_num = float(total)
                    calculated_total += total_num
                    item_count += 1
                    item_details.append(f"Baris {item.get('row')}: {total_num:,.0f}")
                except (ValueError, TypeError):
                    continue
        
        # Jika tidak ada subtotal di Excel, tidak perlu dicek
        if subtotal_value is None:
            return
        
        # Bandingkan
        try:
            subtotal_num = float(subtotal_value)
        except (ValueError, TypeError):
            return
        
        tolerance = 1
        if abs(calculated_total - subtotal_num) > tolerance:
            # Hitung selisih
            difference = calculated_total - subtotal_num
            
            # Buat detail item untuk debugging
            items_summary = "\n".join(item_details[:10])  # Tampilkan 10 item pertama
            if len(item_details) > 10:
                items_summary += f"\n... dan {len(item_details) - 10} item lainnya"
            
            self.errors.append({
                'type': 'SUBTOTAL_ERROR',
                'sheet': data.get('sheet_name'),
                'row': subtotal_row,
                'item_name': 'Subtotal / TOTAL',
                'detail': f'TOTAL tidak sesuai! Ditemukan {item_count} item',
                'calculation': f'Jumlah {item_count} item = {calculated_total:,.0f}',
                'expected': calculated_total,
                'actual': subtotal_num,
                'difference': difference,
                'status': 'PERLU CEK',
                'item_count': item_count,
                'items_summary': items_summary
            })
    
    def check_ppn(self, data: Dict[str, Any]) -> None:
        """Cek PPN (11%)"""
        subtotal_value = data.get('subtotal_value')
        ppn_value = data.get('ppn_value')
        ppn_row = data.get('ppn_row')
        
        if subtotal_value is None or ppn_value is None:
            return
        
        try:
            subtotal_num = float(subtotal_value)
            ppn_num = float(ppn_value)
        except (ValueError, TypeError):
            return
        
        # Hitung PPN 11%
        expected_ppn = subtotal_num * 0.11
        
        tolerance = 1
        if abs(expected_ppn - ppn_num) > tolerance:
            self.errors.append({
                'type': 'PPN_ERROR',
                'sheet': data.get('sheet_name'),
                'row': ppn_row,
                'item_name': 'PPN (11%)',
                'detail': 'PPN tidak sesuai',
                'calculation': f'Subtotal ({subtotal_num}) × 11%',
                'expected': expected_ppn,
                'actual': ppn_num,
                'difference': expected_ppn - ppn_num,
                'status': 'PERLU CEK'
            })
    
    def check_grand_total(self, data: Dict[str, Any]) -> None:
        """Cek Grand Total"""
        subtotal_value = data.get('subtotal_value')
        ppn_value = data.get('ppn_value')
        grand_total_value = data.get('grand_total_value')
        grand_total_row = data.get('grand_total_row')
        
        if subtotal_value is None or grand_total_value is None:
            return
        
        try:
            subtotal_num = float(subtotal_value)
            grand_total_num = float(grand_total_value)
            
            # Hitung grand total
            ppn_num = 0
            if ppn_value is not None:
                ppn_num = float(ppn_value)
            
            expected_grand_total = subtotal_num + ppn_num
        except (ValueError, TypeError):
            return
        
        tolerance = 1
        if abs(expected_grand_total - grand_total_num) > tolerance:
            self.errors.append({
                'type': 'GRAND_TOTAL_ERROR',
                'sheet': data.get('sheet_name'),
                'row': grand_total_row,
                'item_name': 'Grand Total',
                'detail': 'Grand Total tidak sesuai',
                'calculation': f'Subtotal ({subtotal_num}) + PPN ({ppn_num})',
                'expected': expected_grand_total,
                'actual': grand_total_num,
                'difference': expected_grand_total - grand_total_num,
                'status': 'PERLU CEK'
            })
