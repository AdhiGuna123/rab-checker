import pandas as pd
from openpyxl import load_workbook
from typing import Dict, List, Any, Optional
import os

class ExcelReader:
    """Class untuk membaca file Excel RAB"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.wb = None
        self.wb_data = None  # Workbook dengan data_only=True
        self.ws = None
        self.ws_data = None
        self.df = None
        
    def load_workbook(self) -> bool:
        """Load workbook - 2 versi: formula dan data"""
        try:
            # Load untuk membaca formula
            self.wb = load_workbook(self.file_path, data_only=False)
            # Load untuk membaca nilai (cached)
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
            
            # Cek apakah ada kolom QTY dan TOTAL
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
        
        # Mapping kolom berdasarkan posisi (F=6, G=7, H=8, I=9, J=10)
        # atau berdasarkan header
        header_values = {}
        for col in range(1, self.ws.max_column + 1):
            cell_value = self.ws.cell(row=header_row, column=col).value
            if cell_value:
                header_values[str(cell_value).upper().strip()] = col
        
        # Cari kolom berdasarkan header
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
        
        # Default berdasarkan posisi jika tidak ditemukan
        if columns['qty'] is None:
            columns['qty'] = 6  # Kolom F
        if columns['unit_price'] is None:
            columns['unit_price'] = 9  # Kolom I
        if columns['total'] is None:
            columns['total'] = 10  # Kolom J
        
        # Jika item_name belum ketemu, cari kolom yang bukan angka
        if columns['item_name'] is None:
            for col in range(1, columns['qty']):
                # Cek beberapa baris untuk pastikan ini kolom deskripsi
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
    
    def read_data(self, sheet_name: str) -> Dict[str, Any]:
        """Membaca data dari sheet"""
        result = {
            'sheet_name': sheet_name,
            'header_row': None,
            'columns': None,
            'items': [],
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
        
        # Cari header row
        header_row = self.find_header_row()
        if header_row is None:
            return result
        result['header_row'] = header_row
        
        # Cari kolom
        columns = self.find_data_columns(header_row)
        result['columns'] = columns
        
        # Track subtotal yang sudah ditemukan
        found_subtotal = False
        found_ppn = False
        
        # Baca data item
        for row in range(header_row + 1, self.ws.max_row + 1):
            # Cek semua kolom untuk mencari TOTAL/PPN/GRAND TOTAL
            is_summary_row = False
            
            for col in range(1, min(self.ws.max_column + 1, 15)):  # Cek 14 kolom pertama
                cell_value = self.ws.cell(row=row, column=col).value
                if cell_value:
                    cell_str = str(cell_value).upper().strip()
                    
                    # Cek TOTAL/Subtotal
                    if ('SUBTOTAL' in cell_str or 'TOTAL' in cell_str) and 'GRAND' not in cell_str:
                        # Ambil nilai dari kolom total (gunakan data workbook untuk cached value)
                        total_col = columns.get('total', 10)
                        total_cell_data = self.ws_data.cell(row=row, column=total_col)
                        total_cell_formula = self.ws.cell(row=row, column=total_col)
                        
                        # Prioritas: cached value > formula
                        total_value = None
                        if total_cell_data.value is not None:
                            total_value = total_cell_data.value
                        elif total_cell_formula.value is not None:
                            # Jika formula, coba hitung sendiri
                            if isinstance(total_cell_formula.value, str) and total_cell_formula.value.startswith('='):
                                # Hitung subtotal dari semua item
                                calculated_subtotal = 0
                                for item_row in range(header_row + 1, row):
                                    item_total_cell = self.ws_data.cell(row=item_row, column=total_col)
                                    if item_total_cell.value is not None:
                                        try:
                                            calculated_subtotal += float(item_total_cell.value)
                                        except:
                                            pass
                                total_value = calculated_subtotal
                            else:
                                total_value = total_cell_formula.value
                        
                        # Cek apakah ada nilai di kolom total
                        if total_value is not None:
                            if not found_subtotal:
                                result['subtotal_row'] = row
                                result['subtotal_value'] = total_value
                                found_subtotal = True
                            is_summary_row = True
                            break
                    
                    # Cek PPN
                    elif 'PPN' in cell_str or 'TAX' in cell_str:
                        # Ambil nilai dari kolom total (gunakan data workbook untuk cached value)
                        total_col = columns.get('total', 10)
                        total_cell_data = self.ws_data.cell(row=row, column=total_col)
                        total_cell_formula = self.ws.cell(row=row, column=total_col)
                        
                        # Prioritas: cached value > formula
                        ppn_value = None
                        if total_cell_data.value is not None:
                            ppn_value = total_cell_data.value
                        elif total_cell_formula.value is not None:
                            # Jika formula, hitung 11% dari subtotal
                            if isinstance(total_cell_formula.value, str) and total_cell_formula.value.startswith('='):
                                if result.get('subtotal_value') is not None:
                                    try:
                                        ppn_value = float(result['subtotal_value']) * 0.11
                                    except:
                                        pass
                            else:
                                ppn_value = total_cell_formula.value
                        
                        if ppn_value is not None:
                            if not found_ppn:
                                result['ppn_row'] = row
                                result['ppn_value'] = ppn_value
                                found_ppn = True
                            is_summary_row = True
                            break
                    
                    # Cek GRAND TOTAL
                    elif 'GRAND TOTAL' in cell_str:
                        # Ambil nilai dari kolom total (gunakan data workbook untuk cached value)
                        total_col = columns.get('total', 10)
                        total_cell_data = self.ws_data.cell(row=row, column=total_col)
                        total_cell_formula = self.ws.cell(row=row, column=total_col)
                        
                        # Prioritas: cached value > formula
                        grand_total_value = None
                        if total_cell_data.value is not None:
                            grand_total_value = total_cell_data.value
                        elif total_cell_formula.value is not None:
                            # Jika formula, hitung subtotal + ppn
                            if isinstance(total_cell_formula.value, str) and total_cell_formula.value.startswith('='):
                                subtotal_val = result.get('subtotal_value', 0) or 0
                                ppn_val = result.get('ppn_value', 0) or 0
                                try:
                                    grand_total_value = float(subtotal_val) + float(ppn_val)
                                except:
                                    pass
                            else:
                                grand_total_value = total_cell_formula.value
                        
                        if grand_total_value is not None:
                            result['grand_total_row'] = row
                            result['grand_total_value'] = grand_total_value
                            is_summary_row = True
                            break
            
            # Jika bukan baris summary, baca sebagai item
            if not is_summary_row:
                item = self.read_item(row, columns)
                
                # Filter: skip baris header (yang punya "Rp." atau text lain di kolom numerik)
                unit_price = item.get('unit_price')
                total = item.get('total')
                
                # Cek apakah unit_price adalah text seperti "Rp."
                if isinstance(unit_price, str) and ('Rp' in unit_price or 'Rp.' in unit_price):
                    continue  # Skip baris ini
                
                # Cek apakah total adalah text seperti "Rp."
                if isinstance(total, str) and ('Rp' in total or 'Rp.' in total):
                    continue  # Skip baris ini
                
                # Tambahkan item jika ada total atau qty (harus numerik)
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
                    result['items'].append(item)
        
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
        
        # Baca nama item dari kolom item_name
        if columns.get('item_name'):
            cell = self.ws.cell(row=row, column=columns['item_name'])
            item['item_name'] = cell.value
            item['description'] = cell.value  # Simpan juga sebagai description
        
        # Jika item_name kosong atau adalah nama sheet, coba cari kolom deskripsi
        if not item['item_name'] or (isinstance(item['item_name'], str) and len(item['item_name']) > 30):
            # Cari kolom yang berisi text deskripsi (bukan angka, bukan formula)
            for col in range(1, columns.get('qty', 6)):
                cell = self.ws.cell(row=row, column=col)
                cell_value = cell.value
                if cell_value and isinstance(cell_value, str):
                    # Skip jika ini angka, formula, atau text pendek
                    if cell_value.isdigit() or cell_value.startswith('=') or len(cell_value) < 2:
                        continue
                    # Skip jika ini header seperti "Rp.", "NO.", "UNIT", dll
                    if cell_value.upper() in ['RP.', 'RP', 'NO.', 'NO', 'UNIT', 'QTY', 'DESCRIPTION', 'ITEM', 'NAME']:
                        continue
                    # Ini mungkin deskripsi item
                    item['item_name'] = cell_value
                    item['description'] = cell_value
                    break
        
        # Baca qty - ambil nilai dari data workbook
        if columns.get('qty'):
            cell_data = self.ws_data.cell(row=row, column=columns['qty'])
            cell_formula = self.ws.cell(row=row, column=columns['qty'])
            item['qty'] = cell_data.value if cell_data.value is not None else cell_formula.value
        
        # Baca harga awal
        if columns.get('harga_awal'):
            cell_data = self.ws_data.cell(row=row, column=columns['harga_awal'])
            cell_formula = self.ws.cell(row=row, column=columns['harga_awal'])
            item['harga_awal'] = cell_data.value if cell_data.value is not None else cell_formula.value
        
        # Baca mark up
        if columns.get('mark_up'):
            cell_data = self.ws_data.cell(row=row, column=columns['mark_up'])
            cell_formula = self.ws.cell(row=row, column=columns['mark_up'])
            item['mark_up'] = cell_data.value if cell_data.value is not None else cell_formula.value
        
        # Baca unit price - ambil nilai dari data workbook
        if columns.get('unit_price'):
            cell_data = self.ws_data.cell(row=row, column=columns['unit_price'])
            cell_formula = self.ws.cell(row=row, column=columns['unit_price'])
            # Jika ada formula tapi tidak ada cached value, coba hitung dari harga awal + markup
            if cell_data.value is not None:
                item['unit_price'] = cell_data.value
            elif cell_formula.data_type == 'f':
                # Ada formula tapi tidak ada cached value
                item['has_formula'] = True
                item['unit_price'] = None  # Tidak bisa dihitung tanpa nilai
            else:
                item['unit_price'] = cell_formula.value
        
        # Baca total - ambil nilai dari data workbook
        if columns.get('total'):
            cell_data = self.ws_data.cell(row=row, column=columns['total'])
            cell_formula = self.ws.cell(row=row, column=columns['total'])
            
            if cell_data.value is not None:
                item['total'] = cell_data.value
            elif cell_formula.data_type == 'f':
                # Ada formula tapi tidak ada cached value
                item['has_formula'] = True
                item['total_formula'] = cell_formula.value
                # Coba hitung dari qty × unit_price
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
