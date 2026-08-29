import pandas as pd
from openpyxl import load_workbook
from typing import Dict, List, Any, Optional
import os
import re

def safe_float(value):
    """Convert value to float safely, handling commas and formatting"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        # Remove currency symbols and whitespace
        cleaned = value.replace('Rp', '').replace('.', '').replace(',', '').strip()
        if cleaned == '' or cleaned == '-':
            return None
        try:
            return float(cleaned)
        except:
            return None
    return None

class ExcelReader:
    """Class untuk membaca file Excel RAB"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.wb = None
        self.wb_data = None
        self.ws = None
        self.ws_data = None
        self.df = None
        
    def load_workbook(self) -> bool:
        """Load workbook - 2 versi: formula dan data"""
        try:
            self.wb = load_workbook(self.file_path, data_only=False)
            self.wb_data = load_workbook(self.file_path, data_only=True)
            return True
        except Exception as e:
            print(f"Error loading workbook: {e}")
            return False
    
    def get_sheet_names(self) -> List[str]:
        """Mendapatkan nama-nama sheet"""
        if self.wb is None:
            if not self.load_workbook():
                return []
        return self.wb.sheetnames
    
    def select_sheet(self, sheet_name: str) -> bool:
        """Pilih sheet yang akan dianalisis"""
        try:
            if self.wb is None:
                if not self.load_workbook():
                    return False
            self.ws = self.wb[sheet_name]
            self.ws_data = self.wb_data[sheet_name]
            return True
        except Exception as e:
            print(f"Error selecting sheet: {e}")
            return False
    
    def find_header_row(self) -> Optional[int]:
        """Mencari baris header (kolom QTY, Unit Price, TOTAL)"""
        if self.ws is None:
            return None
        
        for row in range(1, self.ws.max_row + 1):
            row_values = []
            for col in range(1, self.ws.max_column + 1):
                cell_value = self.ws.cell(row=row, column=col).value
                if cell_value:
                    row_values.append(str(cell_value).upper().strip())
            
            has_qty = any('QTY' in val or 'QUANTITY' in val or 'JUMLAH' in val for val in row_values)
            has_total = any('TOTAL' in val or 'HARGA TOTAL' in val for val in row_values)
            
            if has_qty and has_total:
                return row
        return None
    
    def find_data_columns(self, header_row: int) -> Dict[str, int]:
        """Mencari posisi kolom data"""
        columns = {
            'item_name': None,
            'qty': None,
            'harga_awal': None,
            'mark_up': None,
            'unit_price': None,
            'total': None
        }
        
        header_values = {}
        for col in range(1, self.ws.max_column + 1):
            cell_value = self.ws.cell(row=header_row, column=col).value
            if cell_value:
                header_values[str(cell_value).upper().strip()] = col
        
        for key, value in header_values.items():
            if 'QTY' in key or 'QUANTITY' in key or 'JUMLAH' in key:
                columns['qty'] = value
            elif 'UNIT PRICE' in key or 'HARGA SATUAN' in key:
                columns['unit_price'] = value
            elif 'TOTAL' in key:
                columns['total'] = value
            elif 'HARGA AWAL' in key or 'PRICE' in key:
                columns['harga_awal'] = value
            elif 'MARK UP' in key or 'MARKUP' in key:
                columns['mark_up'] = value
            elif 'ITEM' in key or 'NAME' in key or 'NAMA' in key or 'DESKRIPSI' in key or 'DESCRIPTION' in key or 'KETERANGAN' in key or 'BARANG' in key:
                columns['item_name'] = value
        
        if columns['qty'] is None:
            columns['qty'] = 6
        if columns['unit_price'] is None:
            columns['unit_price'] = 9
        if columns['total'] is None:
            columns['total'] = 10
        
        if columns['item_name'] is None:
            for col in range(1, columns['qty']):
                has_text = False
                for row in range(header_row + 1, min(header_row + 5, self.ws.max_row + 1)):
                    cell_value = self.ws.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and len(cell_value) > 2:
                        has_text = True
                        break
                if has_text:
                    columns['item_name'] = col
                    break
        
        return columns
    
    def _get_total_value(self, row: int, total_col: int) -> Any:
        """Ambil nilai dari kolom total"""
        total_cell_data = self.ws_data.cell(row=row, column=total_col)
        total_cell_formula = self.ws.cell(row=row, column=total_col)
        
        raw_val = None
        if total_cell_data.value is not None:
            raw_val = total_cell_data.value
        elif total_cell_formula.value is not None:
            if isinstance(total_cell_formula.value, str) and total_cell_formula.value.startswith('='):
                return None
            else:
                raw_val = total_cell_formula.value
        
        if raw_val is not None:
            converted = safe_float(raw_val)
            return converted if converted is not None else raw_val
        return None
    
    def _detect_section_letter(self, text: str) -> Optional[str]:
        """Deteksi huruf section dari text seperti 'Total A', 'Jumlah B'"""
        text_upper = text.upper().strip()
        
        # Pattern: "Total A", "Jumlah B", "Subtotal C", dll
        match = re.search(r'(?:TOTAL|JUMLAH|SUBTOTAL)\s*([A-Z])\b', text_upper)
        if match:
            return match.group(1)
        
        # Pattern: "TOTAL (A+B)" - ini grand total
        if re.search(r'TOTAL\s*\([A-Z]\+[A-Z]\)', text_upper):
            return 'GRAND'
        
        return None
    
    def _detect_row_type(self, cell_str: str) -> str:
        """Deteksi jenis baris: section_header, subtotal, ppn, discount, total, grand_total"""
        cell_upper = cell_str.upper().strip()
        
        # Grand Total
        if 'GRAND TOTAL' in cell_upper or re.search(r'TOTAL\s*\([A-Z]\+[A-Z]\)', cell_upper):
            return 'grand_total'
        
        # Section header standalone (A, B, C)
        if len(cell_upper) <= 2 and cell_upper.isalpha() and len(cell_upper) == 1:
            return 'section_header'
        
        # PPN/Tax
        if 'PPN' in cell_upper or 'TAX' in cell_upper or 'PAJAK' in cell_upper:
            return 'ppn'
        
        # Discount/Diskon
        if 'DISKON' in cell_upper or 'DISCOUNT' in cell_upper or 'POTONGAN' in cell_upper:
            return 'discount'
        
        # Subtotal/Total section
        if 'TOTAL' in cell_upper or 'SUBTOTAL' in cell_upper or 'JUMLAH' in cell_upper:
            return 'subtotal'
        
        return 'unknown'
    
    def read_data(self, sheet_name: str) -> Dict[str, Any]:
        """Membaca data dari sheet dengan support multiple sections dan flexible fields"""
        result = {
            'sheet_name': sheet_name,
            'header_row': None,
            'columns': None,
            'items': [],
            'sections': {},
            'subtotal_row': None,
            'subtotal_value': None,
            'ppn_row': None,
            'ppn_value': None,
            'grand_total_row': None,
            'grand_total_value': None,
            'formulas': {}
        }
        
        if not self.select_sheet(sheet_name):
            return result
        
        header_row = self.find_header_row()
        if header_row is None:
            return result
        result['header_row'] = header_row
        
        columns = self.find_data_columns(header_row)
        result['columns'] = columns
        
        total_col = columns.get('total', 10)
        current_section = None
        pending_items = []
        section_order = []  # Track urutan section
        
        for row in range(header_row + 1, self.ws.max_row + 1):
            is_summary_row = False
            
            # Cek semua kolom untuk mencari label
            for col in range(1, min(self.ws.max_column + 1, 15)):
                cell_value = self.ws.cell(row=row, column=col).value
                if cell_value:
                    cell_str = str(cell_value).strip()
                    row_type = self._detect_row_type(cell_str)
                    
                    if row_type == 'grand_total':
                        # Grand Total
                        grand_total_value = self._get_total_value(row, total_col)
                        if grand_total_value is not None:
                            result['grand_total_row'] = row
                            result['grand_total_value'] = grand_total_value
                            is_summary_row = True
                            break
                    
                    elif row_type == 'section_header':
                        # Section header (A, B, C)
                        section_letter = cell_str.upper()
                        current_section = section_letter
                        
                        if section_letter not in result['sections']:
                            result['sections'][section_letter] = {
                                'items': [],
                                'subtotal_row': None,
                                'subtotal_value': None,
                                'ppn_row': None,
                                'ppn_value': None,
                                'discount_row': None,
                                'discount_value': None,
                                'total_row': None,
                                'total_value': None
                            }
                            section_order.append(section_letter)
                        
                        # Pindahkan pending items ke section ini
                        if pending_items:
                            result['sections'][section_letter]['items'] = pending_items
                            pending_items = []
                        
                        is_summary_row = True
                        break
                    
                    elif row_type == 'ppn':
                        # PPN - simpan ke section yang aktif
                        ppn_value = self._get_total_value(row, total_col)
                        if ppn_value is not None:
                            if current_section and current_section in result['sections']:
                                result['sections'][current_section]['ppn_value'] = ppn_value
                                result['sections'][current_section]['ppn_row'] = row
                            # Global PPN fallback
                            if result['ppn_value'] is None:
                                result['ppn_row'] = row
                                result['ppn_value'] = ppn_value
                            is_summary_row = True
                            break
                    
                    elif row_type == 'discount':
                        # Diskon - simpan ke section yang aktif
                        discount_value = self._get_total_value(row, total_col)
                        if discount_value is not None:
                            if current_section and current_section in result['sections']:
                                result['sections'][current_section]['discount_value'] = discount_value
                                result['sections'][current_section]['discount_row'] = row
                            is_summary_row = True
                            break
                    
                    elif row_type == 'subtotal':
                        # Subtotal/Total section
                        total_value = self._get_total_value(row, total_col)
                        if total_value is not None:
                            section_letter = self._detect_section_letter(cell_str)
                            
                            if section_letter == 'GRAND':
                                result['grand_total_row'] = row
                                result['grand_total_value'] = total_value
                            elif section_letter:
                                # Total dengan label section (Total A, Jumlah B)
                                if section_letter not in result['sections']:
                                    result['sections'][section_letter] = {
                                        'items': [],
                                        'subtotal_row': None,
                                        'subtotal_value': None,
                                        'ppn_row': None,
                                        'ppn_value': None,
                                        'discount_row': None,
                                        'discount_value': None,
                                        'total_row': row,
                                        'total_value': total_value
                                    }
                                    section_order.append(section_letter)
                                else:
                                    # Check apakah ini "Total X" (sudah termasuk PPN) atau "Jumlah X" (belum PPN)
                                    cell_upper = cell_str.upper()
                                    if 'TOTAL' in cell_upper and 'JUMLAH' not in cell_upper:
                                        # Ini Total Section (sudah termasuk PPN/diskon)
                                        result['sections'][section_letter]['total_row'] = row
                                        result['sections'][section_letter]['total_value'] = total_value
                                    else:
                                        # Ini Jumlah/Subtotal (belum PPN)
                                        result['sections'][section_letter]['subtotal_row'] = row
                                        result['sections'][section_letter]['subtotal_value'] = total_value
                                
                                current_section = section_letter
                                
                                if result['subtotal_value'] is None:
                                    result['subtotal_row'] = row
                                    result['subtotal_value'] = total_value
                            else:
                                # Generic total tanpa label section
                                if current_section:
                                    # Check Total vs Jumlah
                                    cell_upper = cell_str.upper()
                                    if 'TOTAL' in cell_upper and 'JUMLAH' not in cell_upper:
                                        result['sections'][current_section]['total_row'] = row
                                        result['sections'][current_section]['total_value'] = total_value
                                    else:
                                        result['sections'][current_section]['subtotal_row'] = row
                                        result['sections'][current_section]['subtotal_value'] = total_value
                                else:
                                    # Belum ada section, buat default
                                    result['sections']['A'] = {
                                        'items': [],
                                        'subtotal_row': row,
                                        'subtotal_value': total_value,
                                        'ppn_row': None,
                                        'ppn_value': None,
                                        'discount_row': None,
                                        'discount_value': None,
                                        'total_row': None,
                                        'total_value': None
                                    }
                                    current_section = 'A'
                                    section_order.append('A')
                                
                                if result['subtotal_value'] is None:
                                    result['subtotal_row'] = row
                                    result['subtotal_value'] = total_value
                            
                            is_summary_row = True
                            break
            
            # Jika bukan baris summary, baca sebagai item
            if not is_summary_row:
                item = self.read_item(row, columns)
                
                unit_price = item.get('unit_price')
                total = item.get('total')
                
                if isinstance(unit_price, str) and ('Rp' in unit_price or 'Rp.' in unit_price):
                    continue
                if isinstance(total, str) and ('Rp' in total or 'Rp.' in total):
                    continue
                
                has_valid_data = False
                
                # Konversi total ke angka
                if item.get('total') is not None:
                    converted = safe_float(item.get('total'))
                    if converted is not None:
                        item['total'] = converted
                        has_valid_data = True
                
                # Pastikan total selalu ada: hitung dari qty × unit_price jika perlu
                qty = safe_float(item.get('qty'))
                up = safe_float(item.get('unit_price'))
                
                if item.get('total') is None or (isinstance(item.get('total'), str)):
                    if qty is not None and up is not None:
                        item['total'] = qty * up
                        has_valid_data = True
                elif qty is not None and up is not None:
                    expected = qty * up
                    current_total = safe_float(item.get('total'))
                    if current_total is None:
                        item['total'] = expected
                    # Biarkan total dari Excel, jangan ditimpa
                
                if not has_valid_data:
                    if qty is not None:
                        has_valid_data = True
                    elif item.get('total') is not None:
                        converted = safe_float(item.get('total'))
                        if converted is not None:
                            has_valid_data = True
                
                if has_valid_data:
                    result['items'].append(item)
                    
                    # Selalu tambahkan ke section aktif atau pending
                    if current_section and current_section in result['sections']:
                        result['sections'][current_section]['items'].append(item)
                    else:
                        pending_items.append(item)
        
        # Handle pending items - PENTING: pastikan semua item masuk ke section
        if pending_items:
            if not result['sections']:
                # Belum ada section, buat section A
                result['sections']['A'] = {
                    'items': pending_items,
                    'subtotal_row': None,
                    'subtotal_value': None,
                    'ppn_row': None,
                    'ppn_value': None,
                    'discount_row': None,
                    'discount_value': None,
                    'total_row': None,
                    'total_value': None
                }
            else:
                # Sudah ada section, masukkan ke section pertama
                first_section = min(result['sections'].keys())
                result['sections'][first_section]['items'].extend(pending_items)
        
        # Hitung subtotal per section jika belum ada
        for section_letter, section_data in result['sections'].items():
            if section_data['subtotal_value'] is None:
                calc_subtotal = 0
                for i in section_data['items']:
                    val = safe_float(i.get('total'))
                    if val is not None:
                        calc_subtotal += val
                section_data['subtotal_value'] = calc_subtotal
        
        # Hitung total items global
        if result['subtotal_value'] is None:
            result['subtotal_value'] = 0
            for i in result['items']:
                val = safe_float(i.get('total'))
                if val is not None:
                    result['subtotal_value'] += val
        
        # Hitung grand total dari sections jika belum ada
        if result['grand_total_value'] is None and len(result['sections']) > 1:
            total_all = 0
            for sl, sd in result['sections'].items():
                section_total = safe_float(sd.get('total_value')) or safe_float(sd.get('subtotal_value')) or 0
                total_all += section_total or 0
            if total_all > 0:
                result['grand_total_value'] = total_all
        
        return result
    
    def read_item(self, row: int, columns: Dict[str, int]) -> Dict[str, Any]:
        """Membaca satu baris item dengan formula dan nilai"""
        item = {
            'row': row,
            'item_name': None,
            'description': None,
            'qty': None,
            'harga_awal': None,
            'mark_up': None,
            'unit_price': None,
            'total': None,
            'total_formula': None,
            'has_formula': False
        }
        
        if columns.get('item_name'):
            cell = self.ws.cell(row=row, column=columns['item_name'])
            item['item_name'] = cell.value
            item['description'] = cell.value
        
        if not item['item_name'] or (isinstance(item['item_name'], str) and len(item['item_name']) > 30):
            for col in range(1, columns.get('qty', 6)):
                cell = self.ws.cell(row=row, column=col)
                cell_value = cell.value
                if cell_value and isinstance(cell_value, str):
                    if cell_value.isdigit() or cell_value.startswith('=') or len(cell_value) < 2:
                        continue
                    if cell_value.upper() in ['RP.', 'RP', 'NO.', 'NO', 'UNIT', 'QTY', 'DESCRIPTION', 'ITEM', 'NAME']:
                        continue
                    item['item_name'] = cell_value
                    item['description'] = cell_value
                    break
        
        if columns.get('qty'):
            cell_data = self.ws_data.cell(row=row, column=columns['qty'])
            cell_formula = self.ws.cell(row=row, column=columns['qty'])
            raw_val = cell_data.value if cell_data.value is not None else cell_formula.value
            item['qty'] = safe_float(raw_val)
        
        if columns.get('harga_awal'):
            cell_data = self.ws_data.cell(row=row, column=columns['harga_awal'])
            cell_formula = self.ws.cell(row=row, column=columns['harga_awal'])
            item['harga_awal'] = cell_data.value if cell_data.value is not None else cell_formula.value
        
        if columns.get('mark_up'):
            cell_data = self.ws_data.cell(row=row, column=columns['mark_up'])
            cell_formula = self.ws.cell(row=row, column=columns['mark_up'])
            item['mark_up'] = cell_data.value if cell_data.value is not None else cell_formula.value
        
        if columns.get('unit_price'):
            cell_data = self.ws_data.cell(row=row, column=columns['unit_price'])
            cell_formula = self.ws.cell(row=row, column=columns['unit_price'])
            raw_val = cell_data.value if cell_data.value is not None else cell_formula.value
            if raw_val is not None:
                converted = safe_float(raw_val)
                if converted is not None:
                    item['unit_price'] = converted
                elif cell_formula.data_type == 'f':
                    item['has_formula'] = True
                    item['unit_price'] = None
                else:
                    item['unit_price'] = raw_val
        
        if columns.get('total'):
            cell_data = self.ws_data.cell(row=row, column=columns['total'])
            cell_formula = self.ws.cell(row=row, column=columns['total'])
            
            raw_val = cell_data.value if cell_data.value is not None else cell_formula.value
            if raw_val is not None:
                converted = safe_float(raw_val)
                if converted is not None:
                    item['total'] = converted
                elif cell_formula.data_type == 'f':
                    item['has_formula'] = True
                    item['total_formula'] = cell_formula.value
                    if item.get('qty') is not None and item.get('unit_price') is not None:
                        item['total'] = item['qty'] * item['unit_price']
                else:
                    item['total'] = raw_val
        
        return item
    
    def get_cell_value(self, row: int, col: int) -> Any:
        """Mendapatkan nilai cell"""
        if self.ws is None:
            return None
        return self.ws.cell(row=row, column=col).value
    
    def get_cell_formula(self, row: int, col: int) -> Optional[str]:
        """Mendapatkan formula cell"""
        if self.ws is None:
            return None
        cell = self.ws.cell(row=row, column=col)
        if cell.data_type == 'f':
            return cell.value
        return None
