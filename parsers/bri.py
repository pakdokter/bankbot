import re
import sys
import pdfplumber
from .common import write_xlsx, apply_universal_fields, ACCOUNT_CODES

HEADER_RE = re.compile(
    r'^(\d{2}/\d{2}/\d{2})\s+(\d{2}:\d{2}:\d{2})\s+(.+?)\s+'
    r'([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$'
)
TELLER_RE = re.compile(r'^[A-Z0-9]{5,12}$')

SKIP_STARTS = (
    'LAPORAN TRANSAKSI', 'STATEMENT OF FINANCIAL', 'Halaman', 'Page',
    'Tanggal Laporan', 'Statement Date', 'Kepada Yth', 'Periode Transaksi',
    'Transaction Period', 'No. Rekening', 'Account No', 'Nama Produk',
    'Product Name', 'Alamat Unit Kerja', 'Business Unit Address', 'Valuta',
    'Currency', 'Unit Kerja', 'Business Unit', 'Tanggal Transaksi',
    'Transaction Date',
    'DUSUN', 'KAB.', 'PATTIMURA', 'PR,OVINSI',
)

ACCOUNT_NUMBER_MAP = {
    '015701106040507': ACCOUNT_CODES['bri_personal'],
    '015701001903567': ACCOUNT_CODES['bri_business'],
}

KNOWN_ENTITY_LABELS = {
    '473501000343538': 'Rekening Keluarga/Owner (...538)',
}


def to_float(s):
    return float(s.replace(',', ''))


def extract_account_number(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ''
    m = re.search(r'No\.\s*Rekening[\s\S]{0,40}?:\s*(\d+)', text)
    return m.group(1) if m else None


def extract_lines(pdf_path):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            for raw in text.split('\n'):
                line = raw.strip()
                if not line:
                    continue
                if line.startswith(SKIP_STARTS):
                    continue
                if line.startswith(('AEAV', 'Created By', 'Statement')):
                    continue
                if re.match(r'^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}$', line):
                    continue
                if re.match(r'^\d{10,}$', line):
                    continue
                lines.append(line)

    saldo_awal, saldo_akhir = None, None
    for idx, l in enumerate(lines):
        if l.startswith('Saldo Awal'):
            # the figures are on a line shortly after this header (past the
            # English translation line), e.g.
            # "70,787.00 4,448,094.00 4,415,700.00 38,393.00"
            for lookahead in lines[idx:idx + 4]:
                m = re.match(r'^([\d,]+\.\d{2})\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})\s*$', lookahead)
                if m:
                    saldo_awal = to_float(m.group(1))
                    saldo_akhir = to_float(m.group(2))
                    break
            return lines[:idx], saldo_awal, saldo_akhir
    return lines, saldo_awal, saldo_akhir


FEE_EXACT = {'Admin Fee', 'Monthly Fee ATM'}


def is_fee_description(desc):
    d = desc.strip()
    if d in FEE_EXACT:
        return True
    dl = d.lower()
    return dl.startswith('biaya sms') or 'biaya transfer' in dl or dl.startswith('biaya admin') or dl.startswith('biaya bulanan')


def parse_description(desc):
    """Return (jenis, objek, catatan) from a BRI Uraian Transaksi string.
    objek is None for the "X TO Y" internal-transfer pattern — the caller
    (build_rows) must resolve it using transaction direction, since the
    counterparty is X for a credit and Y for a debit."""
    desc = desc.strip()
    if is_fee_description(desc):
        return 'Biaya Admin', '', desc
    if ' - ' in desc:
        parts = [p.strip() for p in desc.split(' - ') if p.strip()]
        jenis = parts[0]
        objek = parts[-1] if len(parts) > 1 else ''
        extra = parts[1:-1]
        return jenis, objek, '; '.join(extra)
    m = re.match(r'^(Pembayaran QRIS)\s+(.*)$', desc, re.I)
    if m and m.group(2).strip():
        return m.group(1), m.group(2).strip(), ''
    if desc.upper().startswith('QRIS'):
        return 'QRIS', '', desc
    if re.match(r'^(.+?)\s+TO\s+(.+)$', desc):
        return desc, None, ''
    # generic "... ke X" / "... dari X [via Y]" — X is the recipient/sender
    m3 = re.match(r'^.*?\b(?:ke|dari)\b\s+(.+?)(?:\s+via\s+.+)?$', desc, re.I)
    if m3 and m3.group(1).strip():
        return 'Transfer', m3.group(1).strip(), ''
    m2 = re.match(r'^(Top Up [A-Za-z]+|Pembelian [A-Za-z]+ [A-Za-z]+)\s+(.*)$', desc)
    if m2 and m2.group(2).strip():
        return m2.group(1).strip(), '', m2.group(2).strip()
    return desc, '', ''


def build_rows(pdf_path):
    account_number = extract_account_number(pdf_path)
    self_code = ACCOUNT_NUMBER_MAP.get(account_number, account_number or 'BRI')

    # The counterpart in BRI's own "X TO Y" internal-transfer labels is
    # bank-context-dependent: from the personal account (507) the business
    # account shows up as "IBIZ PT STOASPACE ...", from the business account
    # (567) the personal account shows up as "... TO AHMAD ROZIYAN ...".
    entity_code_map = dict(KNOWN_ENTITY_LABELS)
    if self_code == ACCOUNT_CODES['bri_personal']:
        entity_code_map['IBIZ'] = ACCOUNT_CODES['bri_business']
        entity_code_map['STOASPACE'] = ACCOUNT_CODES['bri_business']
    elif self_code == ACCOUNT_CODES['bri_business']:
        entity_code_map['AHMAD ROZIYAN'] = ACCOUNT_CODES['bri_personal']

    lines, saldo_awal, saldo_akhir = extract_lines(pdf_path)
    rows = []
    i = 0
    n = len(lines)
    while i < n:
        m = HEADER_RE.match(lines[i])
        if not m:
            i += 1
            continue
        date_s, time_s, mid, debet_s, kredit_s, saldo_s = m.groups()
        j = i + 1
        extra_desc = []
        while j < n and not HEADER_RE.match(lines[j]):
            extra_desc.append(lines[j])
            j += 1

        tokens = mid.rsplit(' ', 1)
        teller = ''
        desc_full = mid
        if len(tokens) == 2 and TELLER_RE.match(tokens[1]):
            desc_full, teller = tokens[0], tokens[1]

        if extra_desc:
            desc_full = desc_full + ' ' + ' '.join(extra_desc)

        jenis, objek, catatan = parse_description(desc_full)

        debet = to_float(debet_s)
        kredit = to_float(kredit_s)
        saldo = to_float(saldo_s)

        if objek is None:
            # "X TO Y" — counterparty is the sender X for a credit (money
            # came from X into this account) or the recipient Y for a debit
            # (money left this account and went to Y)
            m_to = re.match(r'^(.+?)\s+TO\s+(.+)$', desc_full)
            objek = m_to.group(1).strip() if kredit > 0 else m_to.group(2).strip()

        if (rows and debet > 0 and debet <= 6500 and
                rows[-1].get('_raw_desc') == desc_full):
            jenis, objek = 'Biaya Admin', ''
            catatan = f'Biaya terkait transaksi: {desc_full}'

        desc_upper_compact = desc_full.upper().replace(' ', '')
        if kredit > 0 and ('QRIS' in desc_full.upper() or 'ONUS' in desc_upper_compact or 'OFFUS' in desc_upper_compact):
            jenis = 'Penjualan QRIS'

        if teller:
            catatan = (catatan + f'; Teller/User ID: {teller}').strip('; ')
        catatan = (f'Jam {time_s}; ' + catatan).strip('; ')

        tanggal_full = date_s[:6] + '20' + date_s[6:]  # dd/mm/yy -> dd/mm/yyyy
        rows.append({
            'tanggal': tanggal_full,
            'keterangan': jenis,
            'debit': -debet if debet > 0 else None,
            'kredit': kredit if kredit > 0 else None,
            'saldo': saldo,
            'objek': objek,
            'catatan': catatan,
            '_raw_desc': desc_full,
        })
        i = j
    for r in rows:
        r.pop('_raw_desc', None)

    apply_universal_fields(rows, self_code, entity_code_map)
    return rows, saldo_awal, saldo_akhir, self_code


if __name__ == '__main__':
    pdf_path = sys.argv[1]
    out_path = sys.argv[2]
    rows, saldo_awal, saldo_akhir, self_code = build_rows(pdf_path)
    write_xlsx(rows, out_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)
    print(f'Total baris: {len(rows)}')
