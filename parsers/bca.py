import re
import sys
import pdfplumber
from .common import write_xlsx, apply_universal_fields, ACCOUNT_CODES, month_name, build_filename

MONTH_MAP = {
    'JANUARI': '01', 'FEBRUARI': '02', 'MARET': '03', 'APRIL': '04',
    'MEI': '05', 'JUNI': '06', 'JULI': '07', 'AGUSTUS': '08',
    'SEPTEMBER': '09', 'OKTOBER': '10', 'NOVEMBER': '11', 'DESEMBER': '12',
}

HEADER_RE = re.compile(
    r'^(\d{2}/\d{2})\s+(.+?)\s+([\d.,]+\.\d{2})(\s+DB)?(?:\s+([\d.,]+\.\d{2}))?\s*$'
)
AMOUNT_ECHO_RE = re.compile(r'^\d+\.\d{2}$')          # e.g. 15000000.00
VA_CODE_RE = re.compile(r'^(\d+)/([A-Z .]+)$')         # e.g. 12608/SHOPEE
VA_NUMBER_RE = re.compile(r'^\d{6,}$')                 # long digit-only line
QR_MERCHANT_RE = re.compile(r'^[\d.]+(.*)$')           # 00000.00MADAM BAHA

KNOWN_TYPES = [
    'TRSF E-BANKING CR', 'TRSF E-BANKING DB',
    'TRANSAKSI DEBIT', 'TRANSAKSI KREDIT',
    'BI-FAST DB', 'BI-FAST CR',
    'SWITCHING DB', 'SWITCHING CR',
    'KR OTOMATIS',
    'SETORAN VIA CDM', 'KARTU DEBIT',
    'BIAYA ADM', 'BUNGA', 'PAJAK BUNGA',
]

DISPLAY_LABEL_OVERRIDES = {
    'BIAYA ADM': 'Biaya Admin',
}

# rekening BCA milik Stoa yang sudah dikenal -> kode singkat untuk kolom
# Subjek/Objek (dipetakan dari NO. REKENING di header statement)
ACCOUNT_NUMBER_MAP = {
    '0561864887': ACCOUNT_CODES['bca_887'],
    '7257104417': ACCOUNT_CODES['bca_417'],
    '7255353292': ACCOUNT_CODES['bca_giro'],
}
# entitas lain yang dikenali sebagai rekening Stoa sendiri (untuk kolom
# Subjek/Objek saat entitas itu muncul sebagai lawan transaksi)
ENTITY_CODE_MAP = {
    'STOASPACE RASA': ACCOUNT_CODES['bca_giro'],
}


def canonical_type(raw_label):
    for t in KNOWN_TYPES:
        if raw_label.startswith(t):
            return t, raw_label[len(t):].strip()
    return raw_label.strip(), ''


def display_label(canon, header_extra):
    if canon.startswith('BI-FAST') and header_extra.startswith('BIF BIAYA TXN KE'):
        return 'Biaya Admin'
    return DISPLAY_LABEL_OVERRIDES.get(canon, canon)


def to_float(s):
    return float(s.replace(',', ''))


def extract_lines(pdf_path):
    lines = []
    periode_year, periode_month = None, None
    account_number, account_holder = None, None
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ''
            raw_lines = [l.strip() for l in text.split('\n') if l.strip()]

            for i, line in enumerate(raw_lines):
                if line.startswith('PERIODE'):
                    m = re.search(r'PERIODE\s*:\s*([A-Z]+)\s+(\d{4})', line)
                    if m:
                        periode_month = MONTH_MAP.get(m.group(1))
                        periode_year = m.group(2)
                if 'NO. REKENING' in line and account_number is None:
                    m = re.search(r'NO\.\s*REKENING\s*:\s*(\d+)', line)
                    if m:
                        account_number = m.group(1)
                if page_idx == 0 and account_holder is None and line.startswith('KCP') and i + 1 < len(raw_lines):
                    account_holder = raw_lines[i + 1].strip()

            try:
                start = next(i for i, l in enumerate(raw_lines) if l.startswith('TANGGAL KETERANGAN'))
            except StopIteration:
                continue
            end = len(raw_lines)
            for i, l in enumerate(raw_lines):
                if l.startswith('Bersambung ke halaman berikut'):
                    end = i
                    break
            lines.extend(raw_lines[start + 1:end])
    return lines, periode_month, periode_year, account_number, account_holder


def parse_blocks(lines):
    blocks = []
    current = None
    for line in lines:
        m = HEADER_RE.match(line)
        if m:
            if current:
                blocks.append(current)
            current = {'header': m, 'details': []}
        elif line.startswith('SALDO AWAL') or line.startswith('MUTASI CR') or \
             line.startswith('MUTASI DB') or line.startswith('SALDO AKHIR'):
            if current:
                blocks.append(current)
                current = None
            continue
        else:
            if current is not None:
                current['details'].append(line)
    if current:
        blocks.append(current)
    return blocks


def extract_object_and_note(canon_type, header_extra, details):
    """Return (objek, keterangan_tambahan, is_qris) using per-type heuristics.
    is_qris is only meaningful for 'KR OTOMATIS' (True = confirmed QRIS/EDC
    merchant settlement via the MID:/QR:/DDR: fields, False = some other kind
    of automatic BCA credit, e.g. 'AutoCr-PL', that should NOT be relabeled
    Penjualan QRIS)."""
    clean = []
    va_code = None
    va_merchant = None
    va_number = None
    for d in details:
        if AMOUNT_ECHO_RE.match(d):
            continue
        if d == '-':
            continue
        m_va = VA_CODE_RE.match(d)
        if m_va:
            va_code, va_merchant = m_va.groups()
            continue
        if va_merchant and VA_NUMBER_RE.match(d):
            va_number = d
            continue
        clean.append(d)

    if va_merchant:
        objek = va_merchant.strip()
        note_parts = [f'kode {va_code}']
        if va_number:
            note_parts.append(f'No. VA/ref: {va_number}')
        note_parts.extend(clean)
        return objek, '; '.join(note_parts), None

    if canon_type in ('TRANSAKSI DEBIT', 'TRANSAKSI KREDIT'):
        objek, note = '', []
        for d in clean:
            m = QR_MERCHANT_RE.match(d)
            if m and m.group(1).strip():
                objek = m.group(1).strip()
            else:
                note.append(d)
        return objek, '; '.join(note), None

    if canon_type.startswith('BI-FAST'):
        codes = [d for d in clean if re.match(r'^\d{3}$', d)]
        rest = [d for d in clean if not re.match(r'^\d{3}$', d)]
        objek = rest[0] if rest else ''
        note = codes + rest[1:]
        if header_extra:
            note.insert(0, header_extra)
        return objek, '; '.join(note), None

    if canon_type == 'KR OTOMATIS':
        m = re.match(r'MID\s*:\s*(\d+)\s*(\d{3,4})?$', header_extra.strip())
        mid_no, cbg = (m.group(1), m.group(2) or '') if m else ('', '')
        merchant, qty_line, ddr_line, tgl_settle = '', '', '', ''
        extra = []
        for d in clean:
            mt = re.match(r'^TANGGAL\s*:(\d{2}/\d{2})\s*(.*)$', d)
            if mt:
                tgl_settle = mt.group(1)
                if mt.group(2).strip():
                    merchant = mt.group(2).strip()
                continue
            if d.startswith('DDR'):
                ddr_line = d
            elif d.startswith('QR') or d.startswith('TGH'):
                qty_line = d
            elif not merchant:
                merchant = d
            else:
                extra.append(d)
        note_parts = []
        if header_extra and not m:
            note_parts.append(header_extra)
        if mid_no:
            note_parts.append(f'MID: {mid_no}')
        if cbg:
            note_parts.append(f'CBG: {cbg}')
        if tgl_settle:
            note_parts.append(f'Tgl QRIS: {tgl_settle}')
        if qty_line:
            note_parts.append(qty_line)
        if ddr_line:
            note_parts.append(ddr_line)
        note_parts.extend(extra)
        return merchant, '; '.join(note_parts), bool(m)

    if canon_type.startswith('SWITCHING'):
        objek = ' '.join(clean) if clean else ''
        return objek, header_extra, None

    if canon_type == 'SETORAN VIA CDM':
        objek = clean[-1] if clean else ''
        note = ([header_extra] if header_extra else []) + clean[:-1]
        return objek, '; '.join(note), None

    if canon_type == 'KARTU DEBIT':
        # header_extra holds the merchant name (e.g. "SINAR BAHAGIA PANC")
        return header_extra, '; '.join(clean), None

    if canon_type.startswith('TRSF E-BANKING'):
        if not clean:
            return '', header_extra, None
        objek = clean[-1]
        note = clean[:-1]
        if header_extra:
            note = [header_extra] + note
        return objek, '; '.join(note), None

    note = list(clean)
    if header_extra:
        note.insert(0, header_extra)
    return '', '; '.join(note), None


def build_rows(pdf_path):
    lines, month, year, account_number, account_holder = extract_lines(pdf_path)
    blocks = parse_blocks(lines)
    self_code = ACCOUNT_NUMBER_MAP.get(account_number, account_number or 'BCA')

    rows = []
    running_balance = None
    saldo_awal = None
    warnings = []

    for b in blocks:
        m = b['header']
        ddmm, label_raw, amount_s, db_flag, saldo_s = m.groups()
        amount = to_float(amount_s)
        saldo_printed = to_float(saldo_s) if saldo_s else None

        if label_raw.strip() == 'SALDO AWAL':
            running_balance = amount
            saldo_awal = amount
            continue

        canon, header_extra = canonical_type(label_raw)
        objek, catatan, is_qris = extract_object_and_note(canon, header_extra, b['details'])
        if canon == 'KR OTOMATIS':
            keterangan_label = 'Penjualan QRIS' if is_qris else 'KR OTOMATIS'
        else:
            keterangan_label = display_label(canon, header_extra)

        is_debit = bool(db_flag)
        debit = -amount if is_debit else None
        kredit = None if is_debit else amount

        if running_balance is not None:
            running_balance = running_balance + (kredit or 0) - (-debit if debit else 0)
            running_balance = round(running_balance, 2)
            if saldo_printed is not None and abs(running_balance - saldo_printed) > 0.01:
                warnings.append(
                    f"{ddmm} {canon}: saldo hitung {running_balance:,.2f} != saldo cetak {saldo_printed:,.2f}"
                )
                running_balance = saldo_printed
        else:
            running_balance = saldo_printed

        rows.append({
            'tanggal': f'{ddmm}/{year}' if year else ddmm,
            'keterangan': keterangan_label,
            'debit': debit, 'kredit': kredit,
            'saldo': running_balance,
            'objek': objek, 'catatan': catatan,
        })

    saldo_akhir = running_balance
    apply_universal_fields(rows, self_code, ENTITY_CODE_MAP)
    meta = {'self_code': self_code, 'bulan': month_name(month), 'tahun': year}
    return rows, warnings, saldo_awal, saldo_akhir, self_code, meta


if __name__ == '__main__':
    pdf_path = sys.argv[1]
    out_path = sys.argv[2]
    rows, warnings, saldo_awal, saldo_akhir, self_code, meta = build_rows(pdf_path)
    write_xlsx(rows, out_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)
    print(f'Total baris: {len(rows)}')
    print('Nama file disarankan:', build_filename(meta['self_code'], meta['bulan'], meta['tahun']))
    if warnings:
        print(f'PERINGATAN ({len(warnings)}):')
        for w in warnings:
            print(' -', w)
    else:
        print('Semua saldo kumulatif cocok dengan saldo cetak di statement.')
