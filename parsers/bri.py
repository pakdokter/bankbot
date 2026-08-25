import re
import sys
import pdfplumber
from .common import write_xlsx

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


def to_float(s):
    return float(s.replace(',', ''))


def extract_lines(pdf_path):
    lines = []
    header_name, header_periode = None, None
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            for raw in text.split('\n'):
                line = raw.strip()
                if not line:
                    continue
                if line.startswith(SKIP_STARTS):
                    continue
                # per-page footer watermark/stamp lines (file id, "Created
                # By ...", print timestamp, trailing numeric doc code) —
                # these repeat at the bottom of every page and must never
                # be absorbed as a continuation of the last row on a page
                if line.startswith(('AEAV', 'Created By', 'Statement')):
                    continue
                if re.match(r'^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}$', line):
                    continue
                if re.match(r'^\d{10,}$', line):
                    continue
                # skip the person/company name + address block right after
                # "Kepada Yth." — heuristically these have no digits and are
                # short lines already excluded by prefix checks above; the
                # remaining letterhead lines are filtered by SKIP_STARTS.
                lines.append(line)
    # Everything from the closing summary table onward (Saldo Awal / Total
    # Transaksi Debet / Terbilang / footer stamps) is not a transaction row.
    # It always appears once, on the last page, after every real row.
    for idx, l in enumerate(lines):
        if l.startswith('Saldo Awal'):
            return lines[:idx]
    return lines


FEE_EXACT = {'Admin Fee', 'Monthly Fee ATM'}


def is_fee_description(desc):
    d = desc.strip()
    if d in FEE_EXACT:
        return True
    dl = d.lower()
    return dl.startswith('biaya sms') or 'biaya transfer' in dl or dl.startswith('biaya admin') or dl.startswith('biaya bulanan')


def parse_description(desc):
    """Return (jenis, objek, catatan) from a BRI Uraian Transaksi string."""
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
    # BRI's own internal-transfer labels, e.g. "IBIZ PT STOASPACE R TO
    # AHMAD ROZIYAN HID" or "NBMB AHMAD ROZIYAN TO FLIPTECH LENTERA" —
    # keep the original description as jenis, but pull the recipient
    # (the part after "TO") into objek.
    m_to = re.match(r'^(.+?)\s+TO\s+(.+)$', desc)
    if m_to and m_to.group(2).strip():
        return desc, m_to.group(2).strip(), ''
    # generic "... ke X" / "... dari X [via Y]" — X is the recipient/sender
    m3 = re.match(r'^.*?\b(?:ke|dari)\b\s+(.+?)(?:\s+via\s+.+)?$', desc, re.I)
    if m3 and m3.group(1).strip():
        return 'Transfer', m3.group(1).strip(), ''
    m2 = re.match(r'^(Top Up [A-Za-z]+|Pembelian [A-Za-z]+ [A-Za-z]+)\s+(.*)$', desc)
    if m2 and m2.group(2).strip():
        return m2.group(1).strip(), '', m2.group(2).strip()
    return desc, '', ''


def build_rows(pdf_path):
    lines = extract_lines(pdf_path)
    rows = []
    i = 0
    n = len(lines)
    while i < n:
        m = HEADER_RE.match(lines[i])
        if not m:
            i += 1
            continue
        date_s, time_s, mid, debet_s, kredit_s, saldo_s = m.groups()
        # absorb continuation lines (wrapped description) until next header
        j = i + 1
        extra_desc = []
        while j < n and not HEADER_RE.match(lines[j]):
            extra_desc.append(lines[j])
            j += 1

        # last whitespace token of `mid` may be a teller/user id
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

        # BRI prints a transaction fee as its own row with the exact same
        # Uraian Transaksi as the transfer/top-up it belongs to, just with a
        # small nominal (Rp 1.000 - Rp 6.500) right after it.
        if (rows and debet > 0 and debet <= 6500 and
                rows[-1].get('_raw_desc') == desc_full):
            jenis, objek = 'Biaya Admin', ''
            catatan = f'Biaya terkait transaksi: {desc_full}'

        # incoming money via QRIS is sales revenue for Stoa, not a generic
        # "QRIS" label (also catches the rare rows where BRI's own PDF
        # drops the "QRIS" prefix but it's still an OnUs/OffUs QRIS credit)
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
            'debit': debet if debet > 0 else None,
            'kredit': kredit if kredit > 0 else None,
            'saldo': saldo,
            'objek': objek,
            'catatan': catatan,
            '_raw_desc': desc_full,
        })
        i = j
    for r in rows:
        r.pop('_raw_desc', None)
    return rows


if __name__ == '__main__':
    pdf_path = sys.argv[1]
    out_path = sys.argv[2]
    rows = build_rows(pdf_path)
    write_xlsx(rows, out_path)
    print(f'Total baris: {len(rows)}')
