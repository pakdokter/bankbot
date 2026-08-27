"""Parser for XLSX files that are already in this bot's own 9-column output
format (e.g. a manually reconstructed old statement, or a previous bot
output the user edited/annotated). Column order is detected by header name
so extra analysis columns (e.g. a manual "Selisih vs Saldo Tercatat" check)
can sit alongside ours without breaking anything."""
import re
import sys
from collections import Counter
import openpyxl

from .common import HEADERS, write_xlsx, build_filename, month_name

FIELD_KEYS = ['tanggal', 'keterangan', 'kategori', 'debit', 'kredit', 'saldo', 'subjek', 'objek', 'catatan']
HEADER_NAME_MAP = {h.upper(): key for h, key in zip(HEADERS, FIELD_KEYS)}
SELISIH_NAMES = {'SELISIH', 'SELISIH VS SALDO TERCATAT'}


def _find_columns(ws, max_scan=3):
    for r in range(1, max_scan + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        cols, selisih_col = {}, None
        for i, v in enumerate(row):
            name = str(v or '').strip().upper()
            if name in HEADER_NAME_MAP:
                cols[HEADER_NAME_MAP[name]] = i
            elif name in SELISIH_NAMES:
                selisih_col = i
        if 'tanggal' in cols and 'keterangan' in cols and 'debit' in cols and 'kredit' in cols:
            return r, cols, selisih_col
    return None, {}, None


def is_preformatted(xlsx_path, sheet_name=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
    header_row, cols, _ = _find_columns(ws)
    return header_row is not None


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

    header_row, cols, selisih_col = _find_columns(ws)
    if header_row is None:
        raise ValueError(
            'File ini bukan format bank yang sudah aku kenali maupun format '
            'output bot sendiri (kolom Tanggal/Keterangan/Debit/Kredit tidak '
            'ketemu). Kirim contoh strukturnya biar disesuaikan.'
        )

    def get(row, key):
        idx = cols.get(key)
        return row[idx] if idx is not None and idx < len(row) else None

    rows = []
    saldo_awal, saldo_akhir = None, None
    party_counter = Counter()
    month_year_counter = Counter()
    selisih_flags = 0

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if row is None or all(c is None for c in row):
            continue
        keterangan = get(row, 'keterangan')
        if not keterangan:
            continue
        keterangan = str(keterangan).strip()
        if keterangan in ('Total Debit (Uang Keluar)', 'Total Kredit (Uang Masuk)', 'Saldo Akhir'):
            saldo_val = to_float(get(row, 'saldo'))
            if keterangan == 'Saldo Akhir' and saldo_val is not None:
                saldo_akhir = saldo_val
            continue

        tanggal = get(row, 'tanggal')
        kategori = get(row, 'kategori') or ''
        debit = to_float(get(row, 'debit'))
        kredit = to_float(get(row, 'kredit'))
        saldo = to_float(get(row, 'saldo'))
        subjek = get(row, 'subjek') or ''
        objek = get(row, 'objek') or ''
        catatan = get(row, 'catatan') or ''

        if str(kategori).strip() == 'Saldo Awal' or keterangan == 'Saldo Awal':
            saldo_awal = saldo if saldo is not None else saldo_awal
            continue

        if hasattr(tanggal, 'strftime'):
            tgl_str = tanggal.strftime('%d/%m/%Y')
            month_year_counter[(tanggal.month, tanggal.year)] += 1
        elif isinstance(tanggal, str) and re.match(r'^\d{2}/\d{2}/\d{4}$', tanggal):
            tgl_str = tanggal
            d, m, y = tanggal.split('/')
            month_year_counter[(int(m), int(y))] += 1
        else:
            tgl_str = tanggal if isinstance(tanggal, str) else None

        for p in (subjek, objek):
            p = str(p).strip()
            if p and p != '-':
                party_counter[p] += 1

        if selisih_col is not None:
            raw_selisih = row[selisih_col] if selisih_col < len(row) else None
            sv = to_float(raw_selisih)
            if sv and abs(sv) > 0.01:
                selisih_flags += 1

        rows.append({
            'tanggal': tgl_str,
            'keterangan': keterangan,
            'kategori': str(kategori).strip(),
            'debit': debit,
            'kredit': kredit,
            'saldo': saldo,
            'subjek': subjek,
            'objek': objek,
            'catatan': catatan,
        })

    if saldo_akhir is None:
        saldo_akhir = rows[-1]['saldo'] if rows else saldo_awal

    self_code = party_counter.most_common(1)[0][0] if party_counter else 'Rekening'
    if month_year_counter:
        (m, y), _ = month_year_counter.most_common(1)[0]
        bulan, tahun = month_name(m), y
    else:
        bulan, tahun = '', ''

    info = {'selisih_flags': selisih_flags}
    meta = {'self_code': self_code, 'bulan': bulan, 'tahun': tahun}
    return rows, saldo_awal, saldo_akhir, info, meta


if __name__ == '__main__':
    xlsx_path = sys.argv[1]
    out_path = sys.argv[2]
    rows, saldo_awal, saldo_akhir, info, meta = build_rows(xlsx_path)
    write_xlsx(rows, out_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)
    print(f'Total baris: {len(rows)}')
    print(f'Saldo awal: {saldo_awal}, Saldo akhir: {saldo_akhir}')
    print('Meta:', meta)
    if info['selisih_flags']:
        print(f'PERINGATAN: {info["selisih_flags"]} baris punya selisih != 0 di kolom Selisih.')
    print('Nama file disarankan:', build_filename(meta['self_code'], meta['bulan'], meta['tahun']))
