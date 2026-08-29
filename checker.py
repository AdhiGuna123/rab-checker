from typing import Dict, List, Any, Optional
from excel_reader import ExcelReader, safe_float

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
        
        # 4. Cek per section (sudah termasuk cek Jumlah/subtotal per section sebelum PPN)
        self.check_sections(data)
        
        # 5. Cek Jumlah Global sebelum PPN (sum semua Jumlah section vs Jumlah/Total global)
        self.check_global_subtotal(data)
        
        # 6. Cek PPN Global gabungan (11% * Jumlah Global)
        self.check_global_ppn(data)
        
        # 7. Cek tanpa PPN & PPN hanya 1 bagian (case khusus)
        self.check_without_ppn(data)
        self.check_case2_single_ppn_section(data)

        # 8. Cek grand total global (sudah termasuk PPN)
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
        """Cek apakah Qty × Unit Price = Total (pakai nilai Excel asli bila ada)"""
        items = data.get('items', [])
        
        for item in items:
            row = item.get('row')
            qty = safe_float(item.get('qty'))
            unit_price = safe_float(item.get('unit_price'))
            # Prefer raw Excel total if we overwrote total with calc
            total_excel = safe_float(item.get('excel_total_raw'))
            if total_excel is None:
                total_excel = safe_float(item.get('total'))
                # If total was already overwritten to calc, skip duplicate check
                if item.get('excel_total_raw') is None and item.get('calc_mismatch'):
                    total_excel = safe_float(item.get('excel_total_raw'))
            
            if qty is None or unit_price is None:
                continue
            expected_total = qty * unit_price
            # If we have a stored excel raw, compare against it
            actual_excel = safe_float(item.get('excel_total_raw'))
            if actual_excel is not None:
                total_excel = actual_excel
            elif item.get('has_formula'):
                continue
            
            if total_excel is None:
                continue
            
            tolerance = 1
            if abs(expected_total - total_excel) > tolerance:
                self.errors.append({
                    'type': 'MULTIPLICATION_ERROR',
                    'sheet': data.get('sheet_name'),
                    'row': row,
                    'item_name': item.get('item_name', 'Unknown'),
                    'detail': f'Total tidak sesuai',
                    'calculation': f'Qty ({qty}) × Unit Price ({unit_price})',
                    'expected': expected_total,
                    'actual': total_excel,
                    'difference': expected_total - total_excel,
                    'status': 'PERLU CEK'
                })
    
    def check_formulas(self, data: Dict[str, Any]) -> None:
        """CASE: Pemeriksaan Formula Excel — TOTAL harus pakai rumus, SUM harus cover semua item, bukan angka manual."""
        items = data.get('items', [])
        for it in items:
            # excel_reader sudah tandai has_formula / total_formula_exposed
            if not it.get('has_formula') and safe_float(it.get('total')) is not None:
                # Jika ada pattern SUM di sheet tapi item ini angka manual, beri warning (bukan error keras — Excel baru kadang cache kosong)
                raw = str(it.get('total_formula_exposed') or '')
                if raw.startswith('=') and 'SUM' in raw.upper():
                    continue
                # Hanya warning ringan — jangan spam untuk semua item tanpa rumus (template manual)
                # Warning hanya jika ada indikasi file seharusnya pakai formula (ada minimal 1 item berformula di sheet)
                has_any_formula = any(x.get('has_formula') for x in items)
                if has_any_formula:
                    self.warnings.append({
                        'type': 'FORMULA_MISSING',
                        'sheet': data.get('sheet_name'),
                        'row': it.get('row'),
                        'item_name': it.get('item_name', 'Unknown'),
                        'detail': 'Formula hilang atau diganti nilai manual pada TOTAL',
                        'calculation': 'Seharusnya =Qty*UnitPrice atau SUM',
                        'expected': None,
                        'actual': it.get('total'),
                        'status': 'WARNING'
                    })
    
    def check_empty_data(self, data: Dict[str, Any]) -> None:
        """Cek data kosong — hanya jika benar-benar tidak ada total (bukan dari qty*price)"""
        # Dinonaktifkan: total sekarang selalu dihitung dari qty*unit_price, jadi EMPTY_TOTAL/EMPTY_* jadi noise
        return
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
            subtotal_excel = safe_float(section_data.get('subtotal_value'))
            ppn_excel = safe_float(section_data.get('ppn_value'))
            discount_excel = safe_float(section_data.get('discount_value'))
            total_excel = safe_float(section_data.get('total_value'))
            
            # Hitung subtotal dari items
            calculated_subtotal = 0
            item_count = 0
            for item in items:
                total = safe_float(item.get('total'))
                if total is not None:
                    calculated_subtotal += total
                    item_count += 1
            
            # 1. Cek Subtotal (Jumlah X) — hanya jika ada baris Jumlah/Total kategori/section dari Excel
            if subtotal_excel is not None:
                is_calc = section_data.get('subtotal_is_calculated', False)
                is_category = section_data.get('is_category', False)
                if is_calc and is_category:
                    # Kategori pemisah (Sparepart/Instalasi) tanpa baris Jumlah di Excel -> jangan error palsu
                    pass
                elif is_calc:
                    pass
                else:
                    tolerance = 1
                    if abs(calculated_subtotal - subtotal_excel) > tolerance:
                        self.errors.append({
                            'type': 'SECTION_SUBTOTAL_ERROR',
                            'sheet': data.get('sheet_name'),
                            'row': section_data.get('subtotal_row'),
                            'item_name': f'Subtotal Section {section_letter}',
                            'detail': f'Subtotal Section {section_letter} tidak sesuai! {item_count} item',
                            'calculation': f'Jumlah {item_count} item',
                            'expected': calculated_subtotal,
                            'actual': subtotal_excel,
                            'difference': calculated_subtotal - subtotal_excel,
                            'status': 'PERLU CEK',
                            'section': section_letter
                        })
            
            # 2. Cek PPN per section (fleksibel: jika ada PPN global gabungan, jangan validasi PPN section kosong sebagai error)
            if ppn_excel is not None and subtotal_excel is not None:
                expected_ppn = subtotal_excel * 0.11
                tolerance = 1
                if abs(expected_ppn - ppn_excel) > tolerance:
                    self.errors.append({
                        'type': 'SECTION_PPN_ERROR',
                        'sheet': data.get('sheet_name'),
                        'row': section_data.get('ppn_row'),
                        'item_name': f'PPN Section {section_letter}',
                        'detail': f'PPN Section {section_letter} tidak sesuai',
                        'calculation': f'Subtotal ({subtotal_excel}) × 11%',
                        'expected': expected_ppn,
                        'actual': ppn_excel,
                        'difference': expected_ppn - ppn_excel,
                        'status': 'PERLU CEK',
                        'section': section_letter
                    })
            # Jika PPN tidak ada di section tapi ada global, jangan anggap salah — kasus PPN gabungan A+B
            
            # 3. Cek Total Section — fleksibel: skip jika ini kasus 1 Total global yang sudah dipromosikan jadi Grand Total
            # Jangan cek total section jika nilainya memang adalah grand_total (sudah dipindahkan)
            if total_excel is not None:
                base = subtotal_excel if subtotal_excel is not None else calculated_subtotal
                ppn = ppn_excel if ppn_excel else 0
                discount = discount_excel if discount_excel else 0
                
                expected_total = base + ppn - discount
                tolerance = 1
                if abs(expected_total - total_excel) > tolerance:
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
                        'actual': total_excel,
                        'difference': expected_total - total_excel,
                        'status': 'PERLU CEK',
                        'section': section_letter
                    })

    def check_global_subtotal(self, data: Dict[str, Any]) -> None:
        """CASE 4/5: TOTAL (A+B) / (A+B+C) = sum Total A/B/C (dinamis, bukan hardcode A/B/C)."""
        sections = data.get('sections', {})
        if len(sections) <= 1:
            return
        subtotal_global_excel = safe_float(data.get('jumlah_global_excel'))
        if subtotal_global_excel is None:
            subtotal_global_excel = safe_float(data.get('subtotal_value'))
        subtotal_row = data.get('jumlah_global_row') or data.get('subtotal_row')
        if subtotal_global_excel is None:
            return
        # 5-CASE: jumlah sub bagian dinamis (A/B/C atau lebih) — hitung via loop
        sum_sub = 0
        for sd in sections.values():
            v = safe_float(sd.get('subtotal_value'))
            if v is None:
                calc = sum(safe_float(it.get('total')) or 0 for it in sd.get('items', []))
                v = calc if calc else 0
            sum_sub += v or 0
        if sum_sub == 0:
            return
        if abs(sum_sub - subtotal_global_excel) > 1:
            letters = "+".join(sorted(sections.keys()))
            self.errors.append({
                'type': 'GLOBAL_SUBTOTAL_ERROR',
                'sheet': data.get('sheet_name'),
                'row': subtotal_row,
                'item_name': f'TOTAL ({letters})' if len(letters) <= 7 else 'TOTAL Gabungan',
                'detail': f'TOTAL ({letters}) tidak sesuai — SUM semua Total section salah' if len(letters) <= 7 else 'TOTAL gabungan tidak sesuai',
                'calculation': f'{" + ".join(sorted(sections.keys()))}',
                'expected': sum_sub,
                'actual': subtotal_global_excel,
                'difference': sum_sub - subtotal_global_excel,
                'status': 'PERLU CEK'
            })

    def check_global_ppn(self, data: Dict[str, Any]) -> None:
        """PPN global — pakai TOTAL (Jumlah Global) × 11% sesuai tulisan 1) Jumlah A+B=TOTAL → 2) TOTAL×11%=PPN."""
        sections = data.get('sections', {})
        if not sections:
            return
        global_ppn = safe_float(data.get('ppn_value'))
        if global_ppn is None:
            return
        # Prioritas: jumlah_global_excel (TOTAL) × 11% sesuai tulisan; fallback sum sections
        total_global = safe_float(data.get('jumlah_global_excel'))
        if total_global is None:
            total_global = safe_float(data.get('subtotal_value'))
        if total_global is None:
            total_global = sum(safe_float(sd.get('subtotal_value')) or sum(safe_float(it.get('total')) or 0 for it in sd.get('items', [])) or 0 for sd in sections.values())
        if not total_global:
            return
        has_any_section_ppn = any(safe_float(v.get('ppn_value')) is not None for v in sections.values())
        if len(sections) > 1 and (data.get('ppn_is_combined') or not has_any_section_ppn):
            expected = total_global * 0.11
            if abs(expected - global_ppn) > 1:
                self.errors.append({
                    'type': 'GLOBAL_PPN_ERROR',
                    'sheet': data.get('sheet_name'),
                    'row': data.get('ppn_row'),
                    'item_name': 'PPN Global',
                    'detail': 'PPN tidak sesuai — harus TOTAL × 11% (Jumlah A+B × 11%)',
                    'calculation': f'TOTAL ({total_global:,.0f}) × 11%',
                    'expected': expected,
                    'actual': global_ppn,
                    'difference': expected - global_ppn,
                    'status': 'PERLU CEK'
                })

    def check_without_ppn(self, data: Dict[str, Any]) -> None:
        """CASE 1: RAB TANPA PPN — GRAND TOTAL harus = TOTAL (SUM item), tanpa PPN."""
        if not data.get('is_without_ppn'):
            return
        grand = safe_float(data.get('grand_total_value'))
        if grand is None:
            return
        sections = data.get('sections', {})
        # TOTAL = SUM item (jumlah_global_excel atau subtotal_value atau sum section)
        total_excel = safe_float(data.get('jumlah_global_excel'))
        if total_excel is None:
            total_excel = safe_float(data.get('subtotal_value'))
        if total_excel is None and sections:
            total_excel = sum(safe_float(v.get('subtotal_value')) or 0 for v in sections.values())
        if total_excel is None:
            total_excel = sum(safe_float(it.get('total')) or 0 for it in data.get('items', []))
        if total_excel is None:
            return
        if abs(grand - total_excel) > 1:
            self.errors.append({
                'type': 'WITHOUT_PPN_GRAND_ERROR',
                'sheet': data.get('sheet_name'),
                'row': data.get('grand_total_row') or data.get('subtotal_row'),
                'item_name': 'GRAND TOTAL (Tanpa PPN)',
                'detail': 'GRAND TOTAL harus sama dengan TOTAL (tanpa PPN)',
                'calculation': 'GRAND = TOTAL',
                'expected': total_excel,
                'actual': grand,
                'difference': total_excel - grand,
                'status': 'PERLU CEK'
            })
        # Cek jangan ada PPN tiba-tiba (jika file tanpa PPN tapi keisi)
        if safe_float(data.get('ppn_value')) is not None:
            self.warnings.append({
                'type': 'WITHOUT_PPN_HAS_PPN',
                'sheet': data.get('sheet_name'),
                'row': data.get('ppn_row'),
                'item_name': 'PPN',
                'detail': 'File tanpa PPN tapi ditemukan nilai PPN',
                'status': 'WARNING'
            })

    def check_case2_single_ppn_section(self, data: Dict[str, Any]) -> None:
        """CASE 2: PPN hanya di 1 bagian (mis. Bagian B) — TOTAL A+B = Total A + Total B (B sudah termasuk PPN)."""
        sections = data.get('sections', {})
        if len(sections) < 2 or data.get('is_without_ppn') or data.get('ppn_is_combined'):
            return
        ppn_sections = [k for k, v in sections.items() if safe_float(v.get('ppn_value')) is not None]
        if len(ppn_sections) != 1:
            return
        grand = safe_float(data.get('grand_total_value'))
        if grand is None:
            return
        sum_total = 0
        for sl, sd in sections.items():
            t = safe_float(sd.get('total_value'))
            if t is None:
                sub = safe_float(sd.get('subtotal_value'))
                if sub is None:
                    sub = sum(safe_float(it.get('total')) or 0 for it in sd.get('items', []))
                ppn = safe_float(sd.get('ppn_value')) or 0
                if sub is not None:
                    t = sub + ppn
            if t is not None:
                sum_total += t
        if abs(sum_total - grand) > 1:
            self.errors.append({
                'type': 'CASE2_GRAND_ERROR',
                'sheet': data.get('sheet_name'),
                'row': data.get('grand_total_row'),
                'item_name': 'TOTAL A+B',
                'detail': 'TOTAL A+B tidak sesuai — Total A + Total B (salah satu sudah PPN)',
                'calculation': ' + '.join(sorted(sections.keys())),
                'expected': sum_total,
                'actual': grand,
                'difference': sum_total - grand,
                'status': 'PERLU CEK'
            })

    def _audit_mode_label(self, data: Dict[str, Any]) -> str:
        sections = data.get('sections', {})
        if data.get('is_without_ppn'):
            return 'TANPA PPN'
        if data.get('ppn_is_combined') and len(sections) > 1:
            return 'PPN GABUNGAN'
        if len([k for k, v in sections.items() if safe_float(v.get('ppn_value')) is not None]) == 1 and len(sections) > 1:
            return 'PPN 1 BAGIAN'
        if len(sections) > 1:
            return 'MULTI-SECTION'
        return 'NORMAL'

    def check_grand_total_global(self, data: Dict[str, Any]) -> None:
        """CASE 3/4/5 fleksibel: dukung PPN per-section, gabungan, tanpa PPN."""
        sections = data.get('sections', {})
        grand_total_excel = safe_float(data.get('grand_total_value'))
        
        if grand_total_excel is None or len(sections) == 0:
            return
        
        # Hitung total fleksibel dari sections
        total_from_sections = 0
        sum_subtotals = 0
        for section_letter, section_data in sections.items():
            section_total = safe_float(section_data.get('total_value'))
            if section_total is None:
                sub = safe_float(section_data.get('subtotal_value'))
                if sub is None:
                    # fallback hitung dari items
                    calc = 0
                    for it in section_data.get('items', []):
                        v = safe_float(it.get('total'))
                        if v is not None: calc += v
                    sub = calc if calc else None
                ppn = safe_float(section_data.get('ppn_value')) or 0
                disc = safe_float(section_data.get('discount_value')) or 0
                if sub is not None:
                    section_total = sub + ppn - disc
                    sum_subtotals += sub
                else:
                    section_total = None
            else:
                # total sudah termasuk PPN, tetap hitung sum_sub untuk opsi gabungan
                sub = safe_float(section_data.get('subtotal_value'))
                if sub is not None: sum_subtotals += sub
            if section_total is not None:
                total_from_sections += section_total
        
        # Case 1 tanpa PPN sudah di-handle di check_without_ppn — skip di sini agar tidak dobel
        if data.get('is_without_ppn'):
            return
        global_ppn = safe_float(data.get('ppn_value'))
        is_combined = data.get('ppn_is_combined', False)
        has_any_section_ppn = any(safe_float(v.get('ppn_value')) is not None for v in sections.values())
        
        candidates = [total_from_sections]
        if global_ppn is not None:
            if is_combined or not has_any_section_ppn:
                candidates.append(sum_subtotals + global_ppn)
            if not has_any_section_ppn:
                candidates.append(total_from_sections + global_ppn)
        
        tolerance = 1
        if any(abs(c - grand_total_excel) <= tolerance for c in candidates):
            return
        expected = (sum_subtotals + global_ppn) if (is_combined and global_ppn is not None) else total_from_sections
        self.errors.append({
            'type': 'GRAND_TOTAL_ERROR',
            'sheet': data.get('sheet_name'),
            'row': data.get('grand_total_row'),
            'item_name': 'Grand Total',
            'detail': 'Grand Total tidak sesuai dengan jumlah semua section',
            'calculation': f'Jumlah {len(sections)} section' + (f' + PPN global' if global_ppn and not has_any_section_ppn else ''),
            'expected': expected,
            'actual': grand_total_excel,
            'difference': expected - grand_total_excel,
            'status': 'PERLU CEK'
        })
