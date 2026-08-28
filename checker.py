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
        
        # 4. Cek per section
        self.check_sections(data)
        
        # 5. Cek grand total global
        self.check_grand_total_global(data)
        
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
        
        for item in items:
            row = item.get('row')
            qty = item.get('qty')
            unit_price = item.get('unit_price')
            total_excel = item.get('total')
            
            if item.get('has_formula'):
                continue
            
            if qty is None or unit_price is None or total_excel is None:
                continue
                
            try:
                qty_num = float(qty) if qty else 0
                unit_price_num = float(unit_price) if unit_price else 0
                total_num = float(total_excel) if total_excel else 0
            except (ValueError, TypeError):
                continue
            
            expected_total = qty_num * unit_price_num
            
            tolerance = 1
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
        
        for item in items:
            row = item.get('row')
            has_formula = item.get('has_formula', False)
            total_value = item.get('total')
            
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
            
            if qty is None or (isinstance(qty, str) and qty.strip() == ''):
                self.errors.append({
                    'type': 'EMPTY_QTY',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': 'Qty kosong',
                    'status': 'PERLU CEK'
                })
            
            if unit_price is None or (isinstance(unit_price, str) and unit_price.strip() == ''):
                self.errors.append({
                    'type': 'EMPTY_UNIT_PRICE',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': 'Unit Price kosong',
                    'status': 'PERLU CEK'
                })
            
            if total is None or (isinstance(total, str) and total.strip() == ''):
                self.errors.append({
                    'type': 'EMPTY_TOTAL',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': 'Total kosong',
                    'status': 'PERLU CEK'
                })
    
    def check_sections(self, data: Dict[str, Any]) -> None:
        """Cek perhitungan per section"""
        sections = data.get('sections', {})
        
        for section_letter, section_data in sections.items():
            items = section_data.get('items', [])
            subtotal_excel = section_data.get('subtotal_value')
            ppn_excel = section_data.get('ppn_value')
            discount_excel = section_data.get('discount_value')
            total_excel = section_data.get('total_value')
            
            # Hitung subtotal dari items
            calculated_subtotal = 0
            item_count = 0
            for item in items:
                total = item.get('total')
                if total is not None:
                    try:
                        calculated_subtotal += float(total)
                        item_count += 1
                    except:
                        pass
            
            # 1. Cek Subtotal (Jumlah X)
            if subtotal_excel is not None:
                try:
                    subtotal_num = float(subtotal_excel)
                    tolerance = 1
                    if abs(calculated_subtotal - subtotal_num) > tolerance:
                        self.errors.append({
                            'type': 'SECTION_SUBTOTAL_ERROR',
                            'sheet': data.get('sheet_name'),
                            'row': section_data.get('subtotal_row'),
                            'item_name': f'Subtotal Section {section_letter}',
                            'detail': f'Subtotal Section {section_letter} tidak sesuai! {item_count} item',
                            'calculation': f'Jumlah {item_count} item',
                            'expected': calculated_subtotal,
                            'actual': subtotal_num,
                            'difference': calculated_subtotal - subtotal_num,
                            'status': 'PERLU CEK',
                            'section': section_letter
                        })
                except:
                    pass
            
            # 2. Cek PPN per section
            if ppn_excel is not None and subtotal_excel is not None:
                try:
                    subtotal_num = float(subtotal_excel)
                    ppn_num = float(ppn_excel)
                    expected_ppn = subtotal_num * 0.11
                    tolerance = 1
                    if abs(expected_ppn - ppn_num) > tolerance:
                        self.errors.append({
                            'type': 'SECTION_PPN_ERROR',
                            'sheet': data.get('sheet_name'),
                            'row': section_data.get('ppn_row'),
                            'item_name': f'PPN Section {section_letter}',
                            'detail': f'PPN Section {section_letter} tidak sesuai',
                            'calculation': f'Subtotal ({subtotal_num}) × 11%',
                            'expected': expected_ppn,
                            'actual': ppn_num,
                            'difference': expected_ppn - ppn_num,
                            'status': 'PERLU CEK',
                            'section': section_letter
                        })
                except:
                    pass
            
            # 3. Cek Total Section (Subtotal + PPN - Diskon)
            if total_excel is not None:
                try:
                    total_num = float(total_excel)
                    base = float(subtotal_excel) if subtotal_excel else calculated_subtotal
                    ppn = float(ppn_excel) if ppn_excel else 0
                    discount = float(discount_excel) if discount_excel else 0
                    
                    expected_total = base + ppn - discount
                    tolerance = 1
                    if abs(expected_total - total_num) > tolerance:
                        detail_parts = [f'Subtotal ({base:,.0f})']
                        if ppn > 0:
                            detail_parts.append(f'+ PPN ({ppn:,.0f})')
                        if discount > 0:
                            detail_parts.append(f'- Diskon ({discount:,.0f})')
                        
                        self.errors.append({
                            'type': 'SECTION_TOTAL_ERROR',
                            'sheet': data.get('sheet_name'),
                            'row': section_data.get('total_row'),
                            'item_name': f'Total Section {section_letter}',
                            'detail': f'Total Section {section_letter} tidak sesuai',
                            'calculation': ' + '.join(detail_parts),
                            'expected': expected_total,
                            'actual': total_num,
                            'difference': expected_total - total_num,
                            'status': 'PERLU CEK',
                            'section': section_letter
                        })
                except:
                    pass
    
    def check_grand_total_global(self, data: Dict[str, Any]) -> None:
        """Cek Grand Total global"""
        sections = data.get('sections', {})
        grand_total_excel = data.get('grand_total_value')
        
        if grand_total_excel is None or len(sections) == 0:
            return
        
        try:
            grand_total_num = float(grand_total_excel)
            
            # Hitung total dari semua section
            total_from_sections = 0
            for section_letter, section_data in sections.items():
                # Prioritas: total_value > subtotal_value
                section_total = section_data.get('total_value')
                if section_total is None:
                    section_total = section_data.get('subtotal_value')
                if section_total is not None:
                    total_from_sections += float(section_total)
            
            tolerance = 1
            if abs(total_from_sections - grand_total_num) > tolerance:
                self.errors.append({
                    'type': 'GRAND_TOTAL_ERROR',
                    'sheet': data.get('sheet_name'),
                    'row': data.get('grand_total_row'),
                    'item_name': 'Grand Total',
                    'detail': 'Grand Total tidak sesuai dengan jumlah semua section',
                    'calculation': f'Jumlah {len(sections)} section',
                    'expected': total_from_sections,
                    'actual': grand_total_num,
                    'difference': total_from_sections - grand_total_num,
                    'status': 'PERLU CEK'
                })
        except:
            pass
