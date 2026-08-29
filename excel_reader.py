import pandas as pd
from openpyxl import load_workbook
from typing import Dict, List, Any, Optional
import os
import re

def safe_float(value):
    """Convert value to float safely, handling commas and formatting (Indonesian)"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace('Rp', '').replace('Rp.', '').replace('Rp ', '').strip()
        # Handle Indonesian format: 1.234.567,89 -> 1234567.89  or 280,000 -> 280000
        # If contains ',' as decimal separator, normalize
        if ',' in cleaned and '.' in cleaned:
            # 1.234,50
            cleaned = cleaned.replace('.', '').replace(',', '.')
        elif ',' in cleaned:
            # Check if comma is thousands separator (3 digits after comma) or decimal
            parts = cleaned.split(',')
            if len(parts) == 2 and len(parts[1]) == 3 and parts[1].isdigit() and parts[0].replace('-','').isdigit():
                # Likely thousands separator: 280,000
                cleaned = cleaned.replace(',', '')
            elif len(parts) == 2 and len(parts[1]) <= 2:
                # Decimal: 1,5
                cleaned = cleaned.replace(',', '.')
            else:
                # Multiple commas or other -> treat as thousands
                cleaned = cleaned.replace(',', '').replace('.', '')
        else:
            # No comma, just remove dots (thousands)
            cleaned = cleaned.replace('.', '')
        cleaned = cleaned.strip()
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
        
        # Fallback: infer columns from numeric patterns if header detection failed
        if columns['qty'] is None or columns['unit_price'] is None or columns['total'] is None:
            numeric_cols = []
            for col in range(1, self.ws.max_column + 1):
                numeric_count = 0
                for r in range(header_row + 1, min(header_row + 8, self.ws.max_row + 1)):
                    v = self.ws_data.cell(row=r, column=col).value
                    if v is None:
                        v = self.ws.cell(row=r, column=col).value
                    if isinstance(v, (int, float)) and v > 0:
                        numeric_count += 1
                    elif isinstance(v, str) and v.strip().replace(',','').replace('.','').strip().isdigit():
                        # string numeric like "280,000"
                        numeric_count += 1
                if numeric_count >= 2:
                    numeric_cols.append(col)
            numeric_cols = sorted(set(numeric_cols))
            # Heuristic: qty is smallest/col leftmost among numeric, total is rightmost
            if columns['qty'] is None and numeric_cols:
                columns['qty'] = numeric_cols[0]
            if columns['total'] is None and numeric_cols:
                columns['total'] = numeric_cols[-1]
            if columns['unit_price'] is None and len(numeric_cols) >= 2:
                # pick middle numeric col
                if len(numeric_cols) >= 3:
                    columns['unit_price'] = numeric_cols[-2]
                else:
                    columns['unit_price'] = numeric_cols[1] if numeric_cols[1] != columns['total'] else numeric_cols[0]
        
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
    
    def _normalize_label(self, text: str) -> str:
        """Normalisasi label: ringkas spasi, hapus .:()- dan uppercase. Untuk toleran typo."""
        s = text.upper().strip()
        # Hapus separator umum yang sering bikin typo spasi
        s = s.replace('.', ' ').replace(':', ' ').replace('(', ' ').replace(')', ' ').replace('-', ' ').replace('/', ' ')
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    def _detect_section_letter(self, text: str) -> Optional[str]:
        """Deteksi huruf section dari text seperti 'Total A', 'Jumlah B' — toleran typo."""
        norm = self._normalize_label(text)
        # Normalisasi JML -> JUMLAH
        norm_jml = re.sub(r'\bJML\b', 'JUMLAH', norm)
        # Pattern: "Total A", "Jumlah B", "Subtotal C", "Jml. A", "JumlahA"
        # Support tanpa spasi: JUMLAHA -> JUMLAH A
        m2 = re.search(r'(?:TOTAL|JUMLAH|SUBTOTAL)[\s\.]*([A-Z])\b', norm_jml)
        if m2:
            return m2.group(1)
        # Tanpa spasi sama sekali: JUMLAHA
        m3 = re.search(r'(?:TOTAL|JUMLAH|SUBTOTAL)([A-Z])\b', norm_jml.replace(' ', ''))
        if m3:
            return m3.group(1)
        # Pattern: "TOTAL (A+B)" - ini grand total
        if re.search(r'TOTAL\s*[\(]?\s*[A-Z]\s*[\+]\s*[A-Z]\s*[\)]?', norm):
            return 'GRAND'
        # Fuzzy fallback via rapidfuzz untuk label panjang
        try:
            from rapidfuzz import fuzz
            for target in ['TOTAL', 'JUMLAH', 'SUBTOTAL']:
                # Ambil token pertama sebagai kandidat
                token = norm_jml.split()[0] if norm_jml.split() else ''
                if token and fuzz.ratio(token, target) >= 85 and re.search(r'[A-Z]\b', norm_jml):
                    ml = re.search(r'([A-Z])\b', norm_jml)
                    if ml:
                        return ml.group(1)
        except ImportError:
            pass
        return None

    def _detect_row_type(self, cell_str: str) -> str:
        """Deteksi jenis baris: section_header, subtotal, ppn, discount, total, grand_total — toleran typo."""
        norm = self._normalize_label(cell_str)
        norm_jml = re.sub(r'\bJML\b', 'JUMLAH', norm)
        
        # Grand Total — toleran: GRAND TOTAL / TOTAL A+B / GRANDTOTAL / TOTAL(A+B)
        if 'GRAND TOTAL' in norm or 'GRANDTOTAL' in norm.replace(' ', '') or re.search(r'TOTAL\s*[\(]?\s*[A-Z]\s*[\+]\s*[A-Z]', norm):
            return 'grand_total'
        # Fuzzy GRAND TOTAL
        try:
            from rapidfuzz import fuzz
            if fuzz.partial_ratio(norm, 'GRAND TOTAL') >= 85 and 'TOTAL' in norm:
                return 'grand_total'
        except ImportError:
            pass
        
        # Section header standalone (A, B, C) — hanya kalau benar-benar 1 huruf
        # Jangan match satuan 'm' yang sudah di-filter di loop (col==1 + cek qty)
        stripped = norm.strip()
        if len(stripped) == 1 and stripped.isalpha() and stripped in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
            return 'section_header'
        
        # PPN/Tax — toleran: PPN / P P N / PAJAK / TAX / PPN 11% / PPN 11 % / PAJAK PPN
        ppn_norm = norm.replace(' ', '')
        if 'PPN' in norm or 'PPN' in ppn_norm or 'PAJAK' in norm or norm_jml.replace(' ', '') == 'TAX':
            # Pastikan bukan substring random: harus mengandung PPN atau PAJAK atau TAX exact
            if re.search(r'\bPPN\b', norm) or 'PPN' in ppn_norm or re.search(r'\bPAJAK\b', norm) or re.search(r'\bTAX\b', norm):
                return 'ppn'
        try:
            from rapidfuzz import fuzz
            for tok in norm.split():
                if fuzz.ratio(tok, 'PPN') >= 80 or fuzz.ratio(tok, 'PAJAK') >= 80:
                    return 'ppn'
        except ImportError:
            pass
        if 'PPN' in cell_str.upper() or 'TAX' in cell_str.upper() or 'PAJAK' in cell_str.upper():
            return 'ppn'
        # Jika tidak match via norm tapi match via raw upper (fallback)
        if 'PPN' in norm or 'PAJAK' in norm:
            return 'ppn'
        
        # Discount/Diskon
        if 'DISKON' in norm or 'DISCOUNT' in norm or 'POTONGAN' in norm:
            return 'discount'
        try:
            from rapidfuzz import fuzz
            for tok in norm.split():
                if fuzz.ratio(tok, 'DISKON') >= 85 or fuzz.ratio(tok, 'DISCOUNT') >= 85:
                    return 'discount'
        except ImportError:
            pass
        
        # Subtotal/Total section — toleran: TOTAL / JUMLAH / JML / SUBTOTAL / SUB TOTAL
        # Rapidfuzz + JML sudah di-normalisasi
        if ('TOTAL' in norm_jml or 'JUMLAH' in norm_jml or 'SUBTOTAL' in norm_jml or 'SUB TOTAL' in norm):
            # Avoid false positive on single "m" etc
            if len(norm.strip()) <= 2 and norm.strip().isalpha():
                return 'unknown'
            return 'subtotal'
        # Fuzzy fallback untuk TOTAL/JUMLAH typo 1 huruf
        try:
            from rapidfuzz import fuzz
            first_tok = norm_jml.split()[0] if norm_jml.split() else ''
            if first_tok and (fuzz.ratio(first_tok, 'TOTAL') >= 85 or fuzz.ratio(first_tok, 'JUMLAH') >= 85 or fuzz.ratio(first_tok, 'SUBTOTAL') >= 85):
                return 'subtotal'
        except ImportError:
            pass
        
        return 'unknown'
    
    def read_data(self, sheet_name: str, overrides: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Membaca data dari sheet dengan support multiple sections dan flexible fields.
        
        overrides (opsional, tanpa AI, 100% lokal):
            header_row: int | None — override baris header (1-indexed)
            qty_col, unit_price_col, total_col: int | None — huruf kolom A=1, B=2, ...
            ppn_mode: 'auto' | 'per_section' | 'combined' | 'single' | 'none'
            total_mode: 'auto' | 'per_section' | 'combined'  — combined = 1 Total global
        """
        overrides = overrides or {}
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
            'formulas': {},
            'overrides_applied': overrides
        }
        
        if not self.select_sheet(sheet_name):
            return result
        
        header_row = overrides.get('header_row') or self.find_header_row()
        if header_row is None:
            return result
        result['header_row'] = header_row
        
        columns = self.find_data_columns(header_row)
        # Apply column overrides
        if overrides.get('qty_col'):
            columns['qty'] = int(overrides['qty_col'])
        if overrides.get('unit_price_col'):
            columns['unit_price'] = int(overrides['unit_price_col'])
        if overrides.get('total_col'):
            columns['total'] = int(overrides['total_col'])
        result['columns'] = columns
        
        total_col = columns.get('total', 10)
        current_section = None
        pending_items = []
        section_order = []  # Track urutan section
        skipped_rows = []  # For debug
        ppn_candidates = []  # Kumpulkan semua baris PPN, klasifikasi setelah loop (lebih akurat)
        classifications = []  # Untuk DEBUG: log klasifikasi dengan normalisasi + flag typo
        
        for row in range(header_row + 1, self.ws.max_row + 1):
            is_summary_row = False
            
            # Cek semua kolom untuk mencari label
            row_classified = False
            for col in range(1, min(self.ws.max_column + 1, 15)):
                cell_value = self.ws.cell(row=row, column=col).value
                if cell_value:
                    cell_str = str(cell_value).strip()
                    norm_dbg = self._normalize_label(cell_str)
                    row_type = self._detect_row_type(cell_str)
                    # Log klasifikasi ringkasan untuk DEBUG (skip item text biasa)
                    if row_type != 'unknown' and not row_classified:
                        is_fuzzy = False
                        try:
                            from rapidfuzz import fuzz
                            # Tandai typo jika mengandung JML/TOTAl typo tapi norm mengandunya
                            if norm_dbg != cell_str.upper().strip():
                                # Ada normalisasi (.:()- dihilangkan) -> kemungkinan typo minor
                                pass
                            # Fuzzy flag: jika first token fuzzy match tapi bukan exact
                            tok = norm_dbg.split()[0] if norm_dbg.split() else ''
                            for target in ['TOTAL','JUMLAH','SUBTOTAL','GRAND TOTAL','PPN','DISKON']:
                                if tok and fuzz.ratio(tok, target.split()[0]) >= 85 and tok != target.split()[0]:
                                    is_fuzzy = True
                                    break
                        except ImportError:
                            pass
                        if cell_str.upper().strip() != norm_dbg and re.search(r'[.:()\-/]', cell_str):
                            is_fuzzy = True
                        classifications.append({'row': row, 'raw': cell_str[:40], 'normalized': norm_dbg[:40], 'type': row_type, 'fuzzy': is_fuzzy})
                        row_classified = True
                    
                    if row_type == 'grand_total':
                        # Grand Total
                        grand_total_value = self._get_total_value(row, total_col)
                        if grand_total_value is not None:
                            result['grand_total_row'] = row
                            result['grand_total_value'] = grand_total_value
                            is_summary_row = True
                            break
                    
                    elif row_type == 'section_header':
                        # Section header (A, B, C) — HANYA jika di kolom 1
                        # Bug sebelumnya: satuan "m" di kolom unit terdeteksi sebagai section "M"
                        if col != 1:
                            continue
                        # Pastikan bukan baris item yang kebetulan punya satuan "m"/"l"
                        # Jika baris punya qty numerik, ini bukan section header
                        _qty_check = None
                        _up_check = None
                        if columns.get('qty'):
                            _v = self.ws_data.cell(row=row, column=columns['qty']).value
                            if _v is None:
                                _v = self.ws.cell(row=row, column=columns['qty']).value
                            _qty_check = safe_float(_v)
                        if columns.get('unit_price'):
                            _v2 = self.ws_data.cell(row=row, column=columns['unit_price']).value
                            if _v2 is None:
                                _v2 = self.ws.cell(row=row, column=columns['unit_price']).value
                            _up_check = safe_float(_v2)
                        if _qty_check is not None or _up_check is not None:
                            continue
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
                        # Kumpulkan dulu, klasifikasi setelah loop (saat semua subtotal sudah kebaca)
                        ppn_value = self._get_total_value(row, total_col)
                        if ppn_value is not None:
                            ppn_candidates.append({
                                'row': row, 'value': ppn_value, 'label': cell_str,
                                'current_section': current_section
                            })
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
                row_dump = None
                
                # STEP 1: always compute total as qty * unit_price as source of truth
                qty = safe_float(item.get('qty'))
                up = safe_float(item.get('unit_price'))
                excel_total = safe_float(item.get('total'))
                
                calc_total = None
                if qty is not None and up is not None:
                    calc_total = qty * up
                
                # STEP 2: decide item total = calc_total if available, else excel_total
                if calc_total is not None:
                    item['total'] = calc_total
                    # Keep raw excel total for error detection if different
                    if excel_total is not None and abs(excel_total - calc_total) > 0.01:
                        item['excel_total_raw'] = excel_total
                        item['calc_mismatch'] = True
                    has_valid_data = True
                elif excel_total is not None:
                    item['total'] = excel_total
                    has_valid_data = True
                elif qty is not None:
                    # Qty exists but no price — keep item so debug is visible
                    has_valid_data = True
                else:
                    # No valid data — log for debug why row skipped
                    row_dump = [str(self.ws.cell(row=row, column=c).value)[:12] if self.ws.cell(row=row, column=c).value else "" for c in range(1, 8)]
                    skipped_rows.append({'row': row, 'dump': row_dump})
                
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
        
        # === OVERRIDE TOTAL MODE (manual) ===
        total_mode = overrides.get('total_mode', 'auto')
        if total_mode == 'combined' and len(result['sections']) > 1 and result['grand_total_value'] is None and section_order:
            # User memaksa: satu Total di bawah = Grand Total A+B, hapus Total per-section
            last = section_order[-1]
            last_data = result['sections'][last]
            cand = last_data.get('total_value')
            cand_row = last_data.get('total_row')
            cand_kind = 'total'
            if cand is None:
                cand = last_data.get('subtotal_value')
                cand_row = last_data.get('subtotal_row')
                cand_kind = 'subtotal'
            if cand is not None:
                result['grand_total_value'] = cand
                result['grand_total_row'] = cand_row
                if cand_kind == 'total':
                    last_data['total_value'] = None
                    last_data['total_row'] = None
                else:
                    last_data['subtotal_value'] = None
                    last_data['subtotal_row'] = None
        elif total_mode == 'auto':
            # === TOTAL FLEKSIBEL OTOMATIS ===
            # Kasus user: Section A tidak isi Total, hanya 1 Total di bawah (seharusnya Grand Total A+B) tapi terbaca sebagai Total B -> salah.
            if len(result['sections']) > 1 and result['grand_total_value'] is None and section_order:
                # Hitung berapa section yang punya nilai dari Excel (sebelum auto-calc subtotal)
                excel_counts = sum(1 for v in result['sections'].values() if v.get('total_value') is not None or v.get('subtotal_value') is not None)
                last = section_order[-1]
                last_data = result['sections'][last]
                cand = last_data.get('total_value')
                cand_row = last_data.get('total_row')
                cand_kind = 'total'
                if cand is None:
                    cand = last_data.get('subtotal_value')
                    cand_row = last_data.get('subtotal_row')
                    cand_kind = 'subtotal'
                # Rule 1: jika HANYA 1 section yang punya Total/Jumlah dari Excel dan itu di section terakhir -> pasti Grand Total
                if excel_counts == 1 and cand is not None:
                    result['grand_total_value'] = cand
                    result['grand_total_row'] = cand_row
                    if cand_kind == 'total':
                        last_data['total_value'] = None
                        last_data['total_row'] = None
                    else:
                        last_data['subtotal_value'] = None
                        last_data['subtotal_row'] = None
                # Rule 2: nilai mendekati sum semua item tapi jauh dari sum section terakhir -> juga Grand Total
                elif cand is not None:
                    cand_f = safe_float(cand)
                    sum_all_items = sum(safe_float(i.get('total')) or 0 for i in result['items'])
                    sum_last_items = sum(safe_float(i.get('total')) or 0 for i in last_data.get('items', []))
                    sum_all_sub = sum(safe_float(v.get('subtotal_value')) or 0 for v in result['sections'].values() if safe_float(v.get('subtotal_value')) is not None)
                    if cand_f is not None:
                        is_global = False
                        for base in [sum_all_items, sum_all_sub, sum_all_items * 1.11, sum_all_sub * 1.11]:
                            if base > 0 and abs(cand_f - base) <= max(2, base * 0.008) and abs(cand_f - sum_last_items) > max(2, cand_f * 0.008):
                                is_global = True
                                break
                        if is_global:
                            result['grand_total_value'] = cand
                            result['grand_total_row'] = cand_row
                            if cand_kind == 'total':
                                last_data['total_value'] = None
                                last_data['total_row'] = None
                            else:
                                last_data['subtotal_value'] = None
                                last_data['subtotal_row'] = None
        elif total_mode == 'per_section':
            # User memaksa: jangan promosikan, biarkan per-section apa adanya
            pass

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
        
        # === KLASIFIKASI PPN setelah semua subtotal/total terbaca ===
        ppn_mode = overrides.get('ppn_mode', 'auto')
        if ppn_candidates:
            if ppn_mode == 'combined':
                # User memaksa: satu PPN gabungan untuk semua section
                sum_sub_all = sum(safe_float(v.get('subtotal_value')) or 0 for v in result['sections'].values())
                best = max(ppn_candidates, key=lambda c: c['row'])  # ambil yang paling bawah (biasanya gabungan)
                result['ppn_value'] = best['value']
                result['ppn_row'] = best['row']
                result['ppn_is_combined'] = True
            elif ppn_mode == 'per_section':
                # User memaksa: semua generik adalah per-section
                section_order_sorted = sorted(result['sections'].keys())
                for cand in ppn_candidates:
                    sec_letter = self._detect_section_letter(cand['label'])
                    if sec_letter and sec_letter != 'GRAND' and sec_letter in result['sections']:
                        result['sections'][sec_letter]['ppn_value'] = cand['value']
                        result['sections'][sec_letter]['ppn_row'] = cand['row']
                    else:
                        for sl in section_order_sorted:
                            if result['sections'][sl]['ppn_value'] is None:
                                result['sections'][sl]['ppn_value'] = cand['value']
                                result['sections'][sl]['ppn_row'] = cand['row']
                                break
                # also set global as sum for display
                s = sum(safe_float(v.get('ppn_value')) or 0 for v in result['sections'].values())
                if s > 0:
                    result['ppn_value'] = s
                    result['ppn_is_combined'] = False
            elif ppn_mode == 'single':
                # Hanya section tertentu (paling bawah / yang punya PPN saja)
                best = max(ppn_candidates, key=lambda c: c['row'])
                pref = best['current_section']
                target = pref if pref and pref in result['sections'] else sorted(result['sections'].keys())[-1]
                result['sections'][target]['ppn_value'] = best['value']
                result['sections'][target]['ppn_row'] = best['row']
                result['ppn_value'] = best['value']
                result['ppn_row'] = best['row']
                result['ppn_is_combined'] = False
            elif ppn_mode == 'none':
                pass  # jangan isi PPN sama sekali
            else:  # auto — heuristik lama yang flexible
                remaining = []
                for cand in ppn_candidates:
                    sec_letter = self._detect_section_letter(cand['label'])
                    if sec_letter and sec_letter != 'GRAND' and sec_letter in result['sections']:
                        result['sections'][sec_letter]['ppn_value'] = cand['value']
                        result['sections'][sec_letter]['ppn_row'] = cand['row']
                    else:
                        remaining.append(cand)
                if remaining:
                    sum_sub_all = sum(safe_float(v.get('subtotal_value')) or 0 for v in result['sections'].values())
                    combined_candidate = None
                    if len(remaining) >= 1 and sum_sub_all > 0:
                        for cand in remaining:
                            expected_combined = sum_sub_all * 0.11
                            if abs(cand['value'] - expected_combined) <= max(2, expected_combined * 0.008):
                                combined_candidate = cand
                                break
                    if combined_candidate is not None:
                        result['ppn_value'] = combined_candidate['value']
                        result['ppn_row'] = combined_candidate['row']
                        result['ppn_is_combined'] = True
                        remaining = [c for c in remaining if c is not combined_candidate]
                    section_order_sorted = sorted(result['sections'].keys())
                    for cand in remaining:
                        placed = False
                        pref = cand['current_section']
                        if pref and pref in result['sections'] and result['sections'][pref]['ppn_value'] is None:
                            result['sections'][pref]['ppn_value'] = cand['value']
                            result['sections'][pref]['ppn_row'] = cand['row']
                            placed = True
                        else:
                            for sl in section_order_sorted:
                                if result['sections'][sl]['ppn_value'] is None:
                                    result['sections'][sl]['ppn_value'] = cand['value']
                                    result['sections'][sl]['ppn_row'] = cand['row']
                                    placed = True
                                    break
                        if not placed:
                            if result.get('ppn_value') is None:
                                result['ppn_value'] = cand['value']
                                result['ppn_row'] = cand['row']
                                result['ppn_is_combined'] = True
                    if len(result['sections']) == 1 and result.get('ppn_value') is None:
                        only = list(result['sections'].values())[0]
                        if only.get('ppn_value') is not None:
                            result['ppn_value'] = only['ppn_value']
                            result['ppn_row'] = only['ppn_row']
        
        # Hitung grand total dari sections jika belum ada (fleksibel)
        if result['grand_total_value'] is None and len(result['sections']) > 1:
            total_all = 0
            for sl, sd in result['sections'].items():
                sec_total = safe_float(sd.get('total_value'))
                if sec_total is None:
                    sub = safe_float(sd.get('subtotal_value')) or 0
                    ppn = safe_float(sd.get('ppn_value')) or 0
                    disc = safe_float(sd.get('discount_value')) or 0
                    sec_total = sub + ppn - disc if (sub or ppn or disc) else None
                if sec_total is not None:
                    total_all += sec_total
            if total_all > 0:
                result['grand_total_value'] = total_all
        # Grand total sudah ada tapi PPN gabungan belum terdeteksi: jangan timpa
        
        # PPN global fallback ringkasan: jika tidak ada global tapi sections ada PPN, global = sum PPN sections (bukan combined)
        if result.get('ppn_value') is None and len(result['sections']) > 1:
            ppn_sum = sum(safe_float(v.get('ppn_value')) or 0 for v in result['sections'].values())
            if ppn_sum > 0:
                result['ppn_value'] = ppn_sum
                result['ppn_is_combined'] = False
        result.setdefault('ppn_is_combined', False)
        
        result['skipped_rows'] = skipped_rows
        result['classifications'] = classifications
        result['columns'] = columns
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
            item['qty_raw'] = raw_val
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
            item['unit_price_raw'] = raw_val
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
            item['total_raw'] = raw_val
            item['total_formula_exposed'] = cell_formula.value if cell_formula.data_type == 'f' else None
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
