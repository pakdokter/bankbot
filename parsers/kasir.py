import re
import sys
import openpyxl
from .common import write_xlsx, month_name, build_filename

MONTH_MAP = {
    'JANUARI': 1, 'FEBRUARI': 2, 'MARET': 3, 'APRIL': 4, 'MEI': 5, 'JUNI': 6,
    'JULI': 7, 'AGUSTUS': 8, 'SEPTEMBER': 9, 'OKTOBER': 10, 'NOVEMBER': 11, 'DESEMBER': 12,
}

# Nama toko/tenant yang sudah dikenal (dari data kasir yang berulang dan
# cross-reference dengan merchant yang muncul di statement bank Stoa).
# Cocok sebagai prefix ATAU substring, tidak case-sensitive. Kalau ada toko
# baru yang sering muncul tapi belum masuk sini, tambahkan ke daftar ini --
# ini bukan daftar resmi dari stoabot, cuma yang bisa dikenali dari data.
KNOWN_TOKO = {
    'ALFAMART': 'Alfamart',
    'MAK OPIK': 'Mak Opik',
    'MAH OPIK': 'Mak Opik',
    'MIRA': 'Mira Laundry',
    'SB': 'SB (Sinar Bahagia)',
    'PASAR': 'Pasar',
    'PRIMER': 'Primer',
    'DINDA FOOD': 'Dinda Food',
    'TOKO YOGA': 'Toko Yoga',
    'MADAM': 'Madam Baha',
    'TOKO SURYA': 'Toko Surya',
    'WAROENG': 'Waroeng',
    'CV HARAPAN': 'CV Harapan Kita',
    'ORIEL CHICKEN': 'Oriel Chicken',
    'DECO COFFEE': 'Deco Coffee',
    'NIDAUL JIHAD': 'Nidaul Jihad',
}
TENANT_LAIN = 'Tenant Lain'

OWNER_NAME_MARKERS = ('OJAN', 'IYAN', 'ROZIYAN')
GENERIC_MONEY_WORDS = {'LEBIH', 'KURANG', 'MINUS', 'KEMBALI', 'CUSTOMER'}


def match_toko(text):
    up = (text or '').upper().strip()
    for needle, canonical in KNOWN_TOKO.items():
        if up.startswith(needle) or needle in up:
            return canonical
    return None


def split_store_item(desc):
    """Return (store_or_None, item_text) from a kasir 'Keterangan' string."""
    for sep in (' – ', ' - ', '-'):
        if sep in desc:
            left, right = desc.split(sep, 1)
            store = match_toko(left)
            if store:
                return store, right.strip()
            return None, desc.strip()
    store = match_toko(desc)
    if store:
        # strip the matched prefix off, if there's more text after it
        up = desc.upper()
        for needle in KNOWN_TOKO:
            if up.startswith(needle):
                rest = desc[len(needle):].strip(' -–')
                return store, (rest if rest else desc.strip())
        return store, desc.strip()
    return None, desc.strip()


def categorize_kasir(kategori_kasir, item_text, person_name=None):
    it = (item_text or '').upper()
    if kategori_kasir == 'Penjualan':
        return 'Penjualan'
    if kategori_kasir == 'Penarikan':
        if person_name and any(m in person_name.upper() for m in OWNER_NAME_MARKERS):
            return 'Modal & Setoran Pemilik'
        return 'Gaji & Tenaga Kerja'
    if kategori_kasir in ('Penerimaan', 'Koreksi'):
        return 'Lainnya / Perlu Verifikasi'
    if any(k in it for k in ('LAUNDRY', 'PARKIR', 'TISU', 'OJEK', 'ONGKIR', 'LAP ')):
        return 'Belanja Operasional'
    if 'LISTRIK' in it:
        return 'Sewa & Utilitas'
    return 'Belanja Bahan'


def to_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def build_rows(xlsx_path, sheet_name=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    title = ws['A1'].value or ''
    ym = re.search(r'(\d{4})', title)
    year = int(ym.group(1)) if ym else None

    rows = []
    saldo_awal = None
    running = None
    matched_count, tenant_lain_count = 0, 0

    for row in ws.iter_rows(min_row=3, values_only=True):
        if row is None or all(c is None for c in row):
            continue
        bulan, tanggal, sesi, keterangan, kategori_kasir, debit_k, kredit_k, saldo_k, flag = (list(row) + [None] * 9)[:9]
        if not keterangan or not kategori_kasir:
            continue

        month = MONTH_MAP.get((bulan or '').strip().upper())
        tgl_str = f'{int(tanggal):02d}/{month:02d}/{year}' if (month and year and tanggal) else None

        debit_k = to_float(debit_k)
        kredit_k = to_float(kredit_k)
        saldo_k = to_float(saldo_k)

        if kategori_kasir == 'Saldo Awal':
            saldo_awal = debit_k if debit_k is not None else saldo_k
            running = saldo_awal
            continue
        if kategori_kasir == 'Saldo':
            # session checkpoint / signature line, not a real transaction
            continue

        # kasir sheet uses the asset-ledger convention (Debit = uang masuk,
        # Kredit = uang keluar) -- flip to this bot's bank-statement
        # convention (Debit = keluar/negatif, Kredit = masuk/positif)
        my_debit = -kredit_k if kredit_k else None
        my_kredit = debit_k if debit_k else None

        person_name = None
        if kategori_kasir == 'Penarikan':
            m = re.search(r'Tarik Tunai\s+(.+)', keterangan, re.I)
            person_name = m.group(1).strip() if m else None
        elif kategori_kasir == 'Penerimaan':
            m = re.search(r'Uang\s+(.+)', keterangan, re.I)
            if m and m.group(1).strip().upper() not in GENERIC_MONEY_WORDS:
                person_name = m.group(1).strip()

        store, item_text = split_store_item(keterangan)
        if store:
            matched_count += 1
            toko = store
        else:
            toko = TENANT_LAIN
            if kategori_kasir == 'Pengeluaran':
                tenant_lain_count += 1

        kategori = categorize_kasir(kategori_kasir, item_text, person_name)

        if my_debit is not None:
            subjek_field, objek_field = 'Kasir', toko
        else:
            if person_name:
                src = person_name
            elif kategori_kasir == 'Penjualan':
                src = 'Penjualan'
            elif kategori_kasir in ('Penerimaan', 'Koreksi'):
                src = '-'
            else:
                src = toko
            subjek_field, objek_field = src, 'Kasir'

        catatan_parts = []
        if flag:
            catatan_parts.append(str(flag))
        if sesi:
            catatan_parts.append(f'Sesi: {sesi}')

        if running is not None:
            running = round(running + (my_kredit or 0) + (my_debit or 0), 2)
        elif saldo_k is not None:
            running = saldo_k

        rows.append({
            'tanggal': tgl_str,
            'keterangan': item_text or keterangan,
            'debit': my_debit,
            'kredit': my_kredit,
            'saldo': running if running is not None else saldo_k,
            'subjek': subjek_field,
            'objek': objek_field,
            'catatan': '; '.join(catatan_parts),
            'kategori': kategori,
        })

    saldo_akhir = running

    info = {
        'toko_dikenali': matched_count,
        'tenant_lain': tenant_lain_count,
    }
    bulan, tahun = '', ''
    if rows and rows[0]['tanggal']:
        d, m, y = rows[0]['tanggal'].split('/')
        bulan, tahun = month_name(m), y
    meta = {'self_code': 'Kasir', 'bulan': bulan, 'tahun': tahun}
    return rows, saldo_awal, saldo_akhir, info, meta


if __name__ == '__main__':
    xlsx_path = sys.argv[1]
    out_path = sys.argv[2]
    rows, saldo_awal, saldo_akhir, info, meta = build_rows(xlsx_path)
    write_xlsx(rows, out_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)
    print(f'Total baris: {len(rows)}')
    print(f'Toko dikenali: {info["toko_dikenali"]}, Tenant Lain (pengeluaran): {info["tenant_lain"]}')
    print(f'Saldo awal: {saldo_awal}, Saldo akhir: {saldo_akhir}')
    print('Nama file disarankan:', build_filename(meta['self_code'], meta['bulan'], meta['tahun']))
