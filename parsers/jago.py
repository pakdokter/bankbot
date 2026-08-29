import re
import sys
import pdfplumber
from .common import write_xlsx, apply_universal_fields, ACCOUNT_CODES, month_name, build_filename

MONTH_MAP = {
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'Mei': '05', 'May': '05',
    'Jun': '06', 'Jul': '07', 'Agu': '08', 'Aug': '08', 'Sep': '09',
    'Okt': '10', 'Oct': '10', 'Nov': '11', 'Des': '12', 'Dec': '12',
}

KNOWN_LABELS = sorted([
    'Isi Saldo Dompet Digital', 'Tarik Uang Kantong', 'Tambah Uang Kantong',
    'Transfer Masuk', 'Transfer Keluar', 'Pembayaran QRIS', 'Transaksi POS',
    'Pajak Bunga', 'Bunga',
], key=len, reverse=True)

INTERNAL_LABELS = {'Tarik Uang Kantong', 'Tambah Uang Kantong'}

BANK_KEYWORDS = ('BCA', 'BRI', 'Bank ', 'GoPay', 'OVO', 'DANA', 'Mandiri', 'Pindah uang antar Kantong')

DATE_LINE_RE = re.compile(r'^(\d{2})\s([A-Za-z]{3})\s(\d{4})\s+(.*?)\s+([+-][\d.,]+)\s+([\d.,]+)$')
KANTONG_HDR_RE = re.compile(r'^(.+?)\s+Saldo Sebelumnya\s+[\d.,]+$')
TIME_RE = re.compile(r'^(\d{2}\.\d{2})\s+(.*)$')
ID_RE = re.compile(r'ID#\s*(\d+)')

SKIP_PREFIXES = (
    'Laporan Keuangan Bulanan', 'PT Bank Jago Tbk', 'www.jago.com',
    'merupakan peserta', 'Tanggal & Waktu',
    'RINGKASAN SALDO', 'SOROTAN', 'KANTONG PERSONAL', 'KANTONG BERSAMA',
    'Nama Kantong', 'Total Saldo Bersama',
)


def id_to_float(s):
    """Indonesian number format: '.' thousands sep, ',' decimal sep."""
    s = s.strip()
    sign = 1
    if s.startswith('+'):
        s = s[1:]
    elif s.startswith('-'):
        sign = -1
        s = s[1:]
    s = s.replace('.', '').replace(',', '.')
    return sign * float(s) if s else 0.0


def looks_like_bank_or_channel(s):
    if not s:
        return False
    if any(ch.isdigit() for ch in s):
        return True
    return any(k in s for k in BANK_KEYWORDS)


def extract_raw_lines(pdf_path):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            for raw in text.split('\n'):
                line = raw.strip()
                if not line or line.startswith(SKIP_PREFIXES):
                    continue
                if line.startswith('INFO PENTING'):
                    return lines  # everything after this is footer disclaimer text
                lines.append(line)
    return lines


def find_saldo_anchor(lines):
    """Consolidated Saldo Sebelumnya / Saldo Akhir across all personal kantong."""
    for i, l in enumerate(lines):
        if l == 'Total Saldo Personal IDR' and i > 0:
            m = re.match(r'^([\d.,]+)\s+([\d.,]+)$', lines[i - 1])
            if m:
                return id_to_float(m.group(1)), id_to_float(m.group(2))
    return None, None


def split_label(middle):
    for label in KNOWN_LABELS:
        pos = middle.rfind(label)
        if pos == -1:
            continue
        end = pos + len(label)
        if end == len(middle) or middle[end] == ' ':
            return middle[:pos].strip(), label, middle[end:].strip()
    return middle.strip(), '', ''


def parse_blocks(lines):
    blocks = []
    current_kantong = None
    cur = None
    for line in lines:
        m_k = KANTONG_HDR_RE.match(line)
        if m_k:
            if cur:
                blocks.append(cur)
                cur = None
            current_kantong = m_k.group(1).strip()
            continue
        m_d = DATE_LINE_RE.match(line)
        if m_d:
            if cur:
                blocks.append(cur)
            cur = {'kantong': current_kantong, 'line1': m_d, 'extra': []}
        else:
            if cur is not None:
                cur['extra'].append(line)
    if cur:
        blocks.append(cur)
    return blocks


def build_transactions(pdf_path):
    lines = extract_raw_lines(pdf_path)
    saldo_awal, saldo_akhir = find_saldo_anchor(lines)
    blocks = parse_blocks(lines)

    txns = []
    for b in blocks:
        dd, mon, yyyy, middle, jumlah_s, saldo_s = b['line1'].groups()
        month = MONTH_MAP.get(mon, '01')
        tanggal = f'{dd}/{month}/{yyyy}'
        name_part, label, catatan_line1 = split_label(middle)

        time_s = ''
        name_cont = ''
        bank_lines = []
        tx_id = ''
        catatan_extra = ''

        for idx, ex in enumerate(b['extra']):
            m_t = TIME_RE.match(ex)
            rest = ex
            if m_t:
                time_s = m_t.group(1)
                rest = m_t.group(2)
            m_id = ID_RE.search(rest)
            if m_id:
                tx_id = m_id.group(1)
                before = rest[:m_id.start()].strip()
                after = rest[m_id.end():].strip()
                if before:
                    if looks_like_bank_or_channel(before):
                        bank_lines.append(before)
                    else:
                        name_cont = (name_cont + ' ' + before).strip()
                if after:
                    catatan_extra = (catatan_extra + ' ' + after).strip()
            else:
                if rest:
                    bank_lines.append(rest)

        objek = (name_part + ' ' + name_cont).strip()
        if label in ('Bunga', 'Pajak Bunga'):
            objek = ''  # Sumber/Tujuan is just self-referential to the kantong
        jumlah = id_to_float(jumlah_s)
        saldo_kantong = id_to_float(saldo_s)
        catatan = '; '.join(p for p in [catatan_line1, catatan_extra] if p)
        bank_info = '; '.join(bank_lines)

        txns.append({
            'kantong': b['kantong'],
            'tanggal': tanggal,
            'sort_key': (yyyy, month, dd, time_s.replace('.', ':')),
            'time': time_s,
            'label': label,
            'jumlah': jumlah,
            'saldo_kantong': saldo_kantong,
            'objek': objek,
            'tx_id': tx_id,
            'catatan': catatan,
            'bank_info': bank_info,
        })

    return txns, saldo_awal, saldo_akhir


def build_rows(pdf_path):
    txns, saldo_awal, saldo_akhir = build_transactions(pdf_path)

    external = [t for t in txns if t['label'] not in INTERNAL_LABELS]
    external.sort(key=lambda t: t['sort_key'])

    rows = []
    running = saldo_awal if saldo_awal is not None else 0.0
    for t in external:
        debit = t['jumlah'] if t['jumlah'] < 0 else None
        kredit = t['jumlah'] if t['jumlah'] > 0 else None
        running = round(running + t['jumlah'], 2)

        note_parts = []
        if t['bank_info']:
            note_parts.append(t['bank_info'])
        if t['catatan']:
            note_parts.append(f"Catatan: {t['catatan']}")
        if t['tx_id']:
            note_parts.append(f"ID#: {t['tx_id']}")
        note_parts.append(f"Kantong: {t['kantong']}")
        if t['time']:
            note_parts.append(f"Jam: {t['time'].replace('.', ':')}")

        rows.append({
            'tanggal': t['tanggal'],
            'keterangan': t['label'],
            'debit': debit,
            'kredit': kredit,
            'saldo': running,
            'objek': t['objek'],
            'catatan': '; '.join(note_parts),
        })

    # Saldo Awal/Akhir yang DIPAKAI di output selalu anchor resmi dari
    # "Total Saldo Personal IDR" di statement -- bukan hasil rekonstruksi
    # `running` di atas. Selisih kecil (recehan) antara keduanya wajar
    # terjadi karena bunga di kantong-kantong kecil (GoPay/Stockbit dst)
    # sering dibulatkan jadi "+0"/"-0" duluan oleh Jago sendiri di PDF-nya,
    # jadi tidak bisa direkonstruksi presisi. Cuma selisih yang genuinely
    # besar yang perlu diperingatkan ke user.
    warning = None
    if saldo_akhir is not None and abs(running - saldo_akhir) > 5:
        warning = f'Saldo akhir hasil konsolidasi {running:,.2f} != saldo akhir statement {saldo_akhir:,.2f}'

    apply_universal_fields(rows, ACCOUNT_CODES['jago'], {})

    bulan, tahun = '', ''
    ref = rows[0] if rows else (txns[0] if txns else None)
    if ref:
        d, m, y = ref['tanggal'].split('/')
        bulan, tahun = month_name(m), y
    meta = {'self_code': ACCOUNT_CODES['jago'], 'bulan': bulan, 'tahun': tahun}
    return rows, saldo_awal, saldo_akhir, warning, meta


if __name__ == '__main__':
    pdf_path = sys.argv[1]
    out_path = sys.argv[2]
    rows, saldo_awal, saldo_akhir, warning, meta = build_rows(pdf_path)
    write_xlsx(rows, out_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)
    print(f'Total baris (transaksi eksternal saja): {len(rows)}')
    print(f'Saldo awal konsolidasi: {saldo_awal}')
    print(f'Saldo akhir statement (anchor): {saldo_akhir}')
    print('Nama file disarankan:', build_filename(meta['self_code'], meta['bulan'], meta['tahun']))
    if warning:
        print('PERINGATAN:', warning)
    else:
        print('Saldo akhir hasil konsolidasi cocok dengan statement.')
