import pandas as pd
from openpyxl import load_workbook
from typing import Dict, List, Any, Optional
import os
import re

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
        
        if total_cell_data.value is not None:
            return total_cell_data.value
        elif total_cell_formula.value is not None:
            if isinstance(total_cell_formula.value, str) and total_cell_formula.value.startswith('='):
                return None  # Formula without cached value
            else:
                return total_cell_formula.value
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
    
    def read_data(self, sheet_name: str) -> Dict[str, Any]:
        """Membaca data dari sheet dengan support multiple sections"""
        result = {
            'sheet_name': sheet_name,
            'header_row': None,
            'columns': None,
            'items': [],
            'sections': {},  # {section_letter: {subtotal, ppn, total, items}}
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
        current_section = None  # Track section aktif
        pending_items = []  # Items yang belum ditentukan sectionnya
        
        for row in range(header_row + 1, self.ws.max_row + 1):
            is_summary_row = False
            is_section_header = False
            
            # Cek semua kolom untuk mencari label
            for col in range(1, min(self.ws.max_column + 1, 15)):
                cell_value = self.ws.cell(row=row, column=col).value
                if cell_value:
                    cell_str = str(cell_value).upper().strip()
                    
                    # 1. Cek GRAND TOTAL dulu (paling prioritas)
                    if 'GRAND TOTAL' in cell_str or re.search(r'TOTAL\s*\([A-Z]\+[A-Z]\)', cell_str):
                        grand_total_value = self._get_total_value(row, total_col)
                        if grand_total_value is not None:
                            result['grand_total_row'] = row
                            result['grand_total_value'] = grand_total_value
                            is_summary_row = True
                            break
                    
                    # 2. Cek standalone section header (A, B, C, dll) - hanya 1-2 karakter
                    elif len(cell_str) <= 2 and cell_str.isalpha() and len(cell_str) == 1:
                        # Ini adalah section header seperti "A", "B"
                        section_letter = cell_str
                        current_section = section_letter
                        
                        # Buat section baru jika belum ada
                        if section_letter not in result['sections']:
                            result['sections'][section_letter] = {
                                'subtotal_row': None,
                                'subtotal_value': None,
                                'ppn_row': None,
                                'ppn_value': None,
                                'total_row': None,
                                'total_value': None,
                                'items': []
                            }
                        
                        # Pindahkan pending items ke section ini
                        if pending_items:
                            result['sections'][section_letter]['items'] = pending_items
                            pending_items = []
                        
                        is_section_header = True
                        is_summary_row = True
                        break
                    
                    # 3. Cek PPN
                    elif 'PPN' in cell_str or 'TAX' in cell_str:
                        ppn_value = self._get_total_value(row, total_col)
                        if ppn_value is not None:
                            # Simpan PPN ke section yang sedang aktif
                            if current_section and current_section in result['sections']:
                                result['sections'][current_section]['ppn_value'] = ppn_value
                                result['sections'][current_section]['ppn_row'] = row
                            # Juga simpan sebagai global PPN
                            if result['ppn_value'] is None:
                                result['ppn_row'] = row
                                result['ppn_value'] = ppn_value
                            is_summary_row = True
                            break
                    
                    # 4. Cek TOTAL/Subtotal - bisa multiple sections
                    elif 'TOTAL' in cell_str or 'SUBTOTAL' in cell_str or 'JUMLAH' in cell_str:
                        total_value = self._get_total_value(row, total_col)
                        if total_value is not None:
                            # Deteksi section letter
                            section_letter = self._detect_section_letter(cell_str)
                            
                            if section_letter == 'GRAND':
                                # Ini grand total
                                result['grand_total_row'] = row
                                result['grand_total_value'] = total_value
                            elif section_letter:
                                # Ini subtotal section tertentu (A, B, C, dst)
                                if section_letter not in result['sections']:
                                    result['sections'][section_letter] = {
                                        'subtotal_row': row,
                                        'subtotal_value': total_value,
                                        'ppn_row': None,
                                        'ppn_value': None,
                                        'total_row': None,
                                        'total_value': None,
                                        'items': []
                                    }
                                else:
                                    result['sections'][section_letter]['subtotal_row'] = row
                                    result['sections'][section_letter]['subtotal_value'] = total_value
                                
                                current_section = section_letter
                                
                                # Simpan sebagai global pertama
                                if result['subtotal_value'] is None:
                                    result['subtotal_row'] = row
                                    result['subtotal_value'] = total_value
                            else:
                                # Generic subtotal tanpa huruf - mungkin "Jumlah B"
                                # Simpan ke section yang aktif atau buat default
                                if current_section:
                                    result['sections'][current_section]['total_row'] = row
                                    result['sections'][current_section]['total_value'] = total_value
                                else:
                                    # Buat section default
                                    result['sections']['A'] = {
                                        'subtotal_row': row,
                                        'subtotal_value': total_value,
                                        'ppn_row': None,
                                        'ppn_value': None,
                                        'total_row': row,
                                        'total_value': total_value,
                                        'items': []
                                    }
                                    current_section = 'A'
                                
                                if result['subtotal_value'] is None:
                                    result['subtotal_row'] = row
                                    result['subtotal_value'] = total_value
                            
                            is_summary_row = True
                            break
            
            # Jika bukan baris summary dan bukan section header, baca sebagai item
            if not is_summary_row and not is_section_header:
                item = self.read_item(row, columns)
                
                unit_price = item.get('unit_price')
                total = item.get('total')
                
                if isinstance(unit_price, str) and ('Rp' in unit_price or 'Rp.' in unit_price):
                    continue
                if isinstance(total, str) and ('Rp' in total or 'Rp.' in total):
                    continue
                
                has_valid_data = False
                if item.get('total') is not None:
                    try:
                        float(item.get('total'))
                        has_valid_data = True
                    except (ValueError, TypeError):
                        pass
                
                if item.get('qty') is not None:
                    try:
                        float(item.get('qty'))
                        has_valid_data = True
                    except (ValueError, TypeError):
                        pass
                
                if has_valid_data:
                    # Tambahkan ke items global
                    result['items'].append(item)
                    
                    # Tambahkan ke section yang sedang aktif atau pending
                    if current_section:
                        if current_section in result['sections']:
                            result['sections'][current_section]['items'].append(item)
                        else:
                            pending_items.append(item)
                    else:
                        # Belum ada section, simpan ke pending
                        pending_items.append(item)
        
        # Jika masih ada pending items, buat section default
        if pending_items:
            if not result['sections']:
                result['sections']['A'] = {
                    'subtotal_row': None,
                    'subtotal_value': None,
                    'ppn_row': None,
                    'ppn_value': None,
                    'total_row': None,
                    'total_value': None,
                    'items': pending_items
                }
            else:
                # Masukkan ke section pertama
                first_section = min(result['sections'].keys())
                result['sections'][first_section]['items'].extend(pending_items)
        
        # Hitung total items sebagai fallback
        if result['subtotal_value'] is None:
            total_items = sum(float(i.get('total', 0) or 0) for i in result['items'])
            result['subtotal_value'] = total_items
        
        # Jika ada multiple sections tapi belum ada grand total, hitung dari total sections
        if result['grand_total_value'] is None and len(result['sections']) > 1:
            total_all_sections = 0
            for section_letter, section_data in result['sections'].items():
                section_total = section_data.get('total_value') or section_data.get('subtotal_value') or 0
                total_all_sections += float(section_total or 0)
            if total_all_sections > 0:
                result['grand_total_value'] = total_all_sections
        
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
            item['qty'] = cell_data.value if cell_data.value is not None else cell_formula.value
        
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
            if cell_data.value is not None:
                item['unit_price'] = cell_data.value
            elif cell_formula.data_type == 'f':
                item['has_formula'] = True
                item['unit_price'] = None
            else:
                item['unit_price'] = cell_formula.value
        
        if columns.get('total'):
            cell_data = self.ws_data.cell(row=row, column=columns['total'])
            cell_formula = self.ws.cell(row=row, column=columns['total'])
            
            if cell_data.value is not None:
                item['total'] = cell_data.value
            elif cell_formula.data_type == 'f':
                item['has_formula'] = True
                item['total_formula'] = cell_formula.value
                if item.get('qty') is not None and item.get('unit_price') is not None:
                    try:
                        item['total'] = float(item['qty']) * float(item['unit_price'])
                    except:
                        item['total'] = None
            else:
                item['total'] = cell_formula.value
        
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
