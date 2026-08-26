import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

HEADERS = [
    'Tanggal', 'Keterangan Transaksi', 'Kategori Transaksi', 'Debit', 'Kredit',
    'Saldo Kumulatif', 'Subjek Transaksi', 'Objek Transaksi', 'Keterangan Tambahan',
]

COL_WIDTHS = [12, 24, 24, 16, 16, 16, 20, 24, 40]

INDO_MONTHS = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
]


def month_name(month_num):
    """month_num: int or numeric string (1-12) -> Indonesian month name."""
    try:
        return INDO_MONTHS[int(month_num) - 1]
    except (TypeError, ValueError, IndexError):
        return ''


def build_filename(self_code, bulan, tahun, ext='xlsx'):
    """'BCA-887', 'Januari', '2025' -> 'BCA-887 Januari 2025.xlsx'"""
    parts = [p for p in (self_code or 'Rekening', bulan or '', str(tahun or '')) if p]
    name = ' '.join(parts).strip()
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    return f'{name}.{ext}'


def sheet_title_from_meta(meta):
    parts = [p for p in (meta.get('self_code') or 'Rekening', meta.get('bulan') or '', str(meta.get('tahun') or '')) if p]
    return ' '.join(parts).strip()[:31]

# Konvensi tanda: Debit = uang KELUAR (disimpan NEGATIF, tampil dalam kurung
# lewat number_format akuntansi), Kredit = uang MASUK (positif). Ini mengikuti
# label DB/CR asli tiap bank, bukan konvensi buku besar aset (yang membalik
# Debit/Kredit). Kalau ternyata maunya dibalik, gampang ditukar di satu
# tempat ini (ROW_BUILDER di bawah).
DEBIT_NUMBER_FORMAT = '#,##0.00;(#,##0.00)'
KREDIT_NUMBER_FORMAT = '#,##0.00;(#,##0.00)'
SALDO_NUMBER_FORMAT = '#,##0.00'

# --- kode singkat rekening Stoa sendiri, untuk kolom Subjek/Objek saat ------
# transaksinya adalah pemindahan uang ANTAR rekening milik Stoa sendiri
# (bukan ke pihak luar) -- memudahkan rekonsiliasi lintas rekening.
ACCOUNT_CODES = {
    'bri_personal': 'BRI-507',
    'bri_business': 'BRI-567(Biz)',
    'bca_887': 'BCA-887',
    'bca_417': 'BCA-417',
    'bca_giro': 'BCA-292(Biz)',
    'jago': 'Jago',
}
FLIPTECH_LABEL = 'Rekening Lain (via Fliptech)'

# --- kategori transaksi untuk kebutuhan akuntansi operasional Stoa -------
# Heuristik berbasis kata kunci pada objek/catatan, plus lookup merchant dan
# nomor rekening spesifik yang sudah dikonfirmasi user lewat feedback.
# Tetap belum 100% akurat untuk semua kasus (lihat catatan di bawah).

OWNER_KEYWORDS = ('AHMAD ROZIYAN', 'ROZIYAN HIDAYAT', 'OJAN')
DEBT_KEYWORDS = ('CICILAN', 'ANGSURAN', 'BAYAR SB', 'PINJAM', 'UTANG')
UTILITY_KEYWORDS = ('SEWA', 'LISTRIK', ' PLN', 'AIR STO', 'UTILITAS')
WALLET_KEYWORDS = ('SHOPEE', 'OVO ', ' OVO', 'GOPAY', 'DANA ', 'TELKOMSEL', 'TOP UP', 'ISI SALDO', 'PULSA', 'BRIVA')
PURCHASE_KEYWORDS = ('BELANJA', 'BELI ', 'ONGKIR', 'SUPPLIER', 'GANTI UANG BELANJA')
PAYROLL_KEYWORDS = ('GAJI',)
TRANSFER_TYPES = ('TRANSFER', 'TRSF', 'BI-FAST', 'SWITCHING', 'KIRIM')

MERCHANT_CATEGORY_MAP = {
    'CV HARAPAN KITA': 'Belanja Operasional',
    'CV HARAPAN': 'Belanja Operasional',
    'MIRA LAUNDRY': 'Belanja Operasional',
    'ORIEL CHICKEN': 'Belanja Konsumsi',
    'DECO COFFEE': 'Belanja Konsumsi',
    'TRANSFERPAY': 'Penjualan',
    'NIDAUL JIHAD': 'Belanja Bahan',
    'DEWA AYU EKA FERR': 'Belanja Konsumsi',
    'ADOBE': 'Belanja Operasional',
    'VISIONET': 'Penjualan',
}

KNOWN_OWNER_ACCOUNTS = {
    '473501000343538',
}

CATEGORIES_REFERENCE = [
    'Saldo Awal',
    'Penjualan',
    'Belanja Bahan',
    'Belanja Konsumsi',
    'Belanja Operasional',
    'Gaji & Tenaga Kerja',
    'Sewa & Utilitas',
    'Cicilan & Utang',
    'Modal & Setoran Pemilik',
    'Transfer Internal (Pindah Rekening)',
    'Transfer Antar Rekening (Fliptech)',
    'Biaya Admin & Pajak Bank',
    'Bunga Bank (Pendapatan)',
    'Transfer Lainnya',
    'Lainnya / Perlu Verifikasi',
]

# --- keterangan universal ---------------------------------------------------
# Menormalkan label mentah tiap bank ke kosakata yang sama, supaya baris di
# BCA/BRI/Jago bisa dibandingkan langsung saat rekonsiliasi antar rekening.

UNIVERSAL_KETERANGAN_EXACT = {
    'BUNGA': 'Bunga Bank',
    'PAJAK BUNGA': 'Pajak Bunga',
    'BIAYA ADMIN': 'Biaya Admin',
    'BIAYA ADM': 'Biaya Admin',
    'SALDO AWAL': 'Saldo Awal',
    'PENJUALAN QRIS': 'Penjualan QRIS',
    'TRANSAKSI DEBIT': 'Pembayaran QRIS',
    'TRANSAKSI KREDIT': 'Penjualan QRIS',
    'PEMBAYARAN QRIS': 'Pembayaran QRIS',
    'QRIS': 'Pembayaran QRIS',
}


def normalize_keterangan(keterangan, debit, kredit):
    ket = (keterangan or '').strip()
    ket_upper = ket.upper()
    if ket_upper in UNIVERSAL_KETERANGAN_EXACT:
        return UNIVERSAL_KETERANGAN_EXACT[ket_upper]
    if 'FLIPTECH' in ket_upper:
        return 'Transfer Internal (Pindah Rekening) via Fliptech'
    if ' TO ' in ket_upper or ket_upper.startswith(('TRANSFER', 'TRSF', 'BI-FAST', 'SWITCHING')):
        return 'Transfer Masuk' if kredit else 'Transfer Keluar'
    return ket


def _merchant_match(objek):
    ob = (objek or '').upper()
    for merchant, cat in MERCHANT_CATEGORY_MAP.items():
        if merchant in ob:
            return cat
    return None


def categorize(keterangan, objek, catatan, debit, kredit):
    """keterangan is expected to already be the *normalized* universal label
    (see normalize_keterangan) — categorize() is always called after that
    step in write_xlsx."""
    ket = (keterangan or '').upper()
    ob = (objek or '').upper()
    text = f"{keterangan or ''} {objek or ''} {catatan or ''}".upper()

    if ket == 'SALDO AWAL':
        return 'Saldo Awal'
    if ket in ('BUNGA', 'BUNGA BANK'):
        return 'Bunga Bank (Pendapatan)'
    if ket in ('PAJAK BUNGA', 'BIAYA ADMIN', 'BIAYA ADM'):
        return 'Biaya Admin & Pajak Bank'
    if 'FLIPTECH' in text:
        return 'Transfer Antar Rekening (Fliptech)'
    if ' TO ' in (keterangan or '').upper() or 'IBIZ' in text or 'NBMB' in text:
        return 'Transfer Internal (Pindah Rekening)'

    merchant_hit = _merchant_match(objek)
    if merchant_hit:
        return merchant_hit

    if ket == 'PENJUALAN QRIS':
        return 'Penjualan'
    if kredit and 'QRIS' in text:
        return 'Penjualan'

    if any(k in text for k in PAYROLL_KEYWORDS):
        return 'Gaji & Tenaga Kerja'

    if any(acc in ob for acc in KNOWN_OWNER_ACCOUNTS):
        return 'Modal & Setoran Pemilik'
    if any(k in ob for k in OWNER_KEYWORDS):
        return 'Modal & Setoran Pemilik'
    if 'TARIK TUNAI' in text or ('SETOR' in text and 'SETORAN' not in ket):
        return 'Modal & Setoran Pemilik'

    if any(k in text for k in DEBT_KEYWORDS):
        return 'Cicilan & Utang'
    if any(k in text for k in UTILITY_KEYWORDS):
        return 'Sewa & Utilitas'
    if any(k in text for k in WALLET_KEYWORDS):
        return 'Belanja Operasional'
    if any(k in text for k in PURCHASE_KEYWORDS):
        return 'Belanja Operasional'
    if any(k in ket for k in TRANSFER_TYPES):
        return 'Belanja Operasional' if debit else 'Transfer Lainnya'
    return 'Lainnya / Perlu Verifikasi'


def resolve_party(name, self_code, entity_code_map):
    """Map a raw counterparty name to a short account code if it's one of
    Stoa's own known accounts; otherwise return the name unchanged."""
    if not name:
        return name
    up = name.upper()
    if 'FLIPTECH' in up:
        return FLIPTECH_LABEL
    for needle, code in entity_code_map.items():
        if needle in up:
            return code
    return name


def apply_universal_fields(rows, self_code='', entity_code_map=None):
    """Normalizes keterangan, fills kategori, and derives Subjek/Objek from
    transaction direction + the account holder's own short code. Call this
    once per statement, from each parser's build_rows (or let write_xlsx do
    it lazily with self_code='')."""
    entity_code_map = entity_code_map or {}
    for r in rows:
        debit, kredit = r.get('debit'), r.get('kredit')
        raw_ket = r.get('keterangan')
        # categorize() looks for bank-specific raw markers (" TO ", "IBIZ",
        # "FLIPTECH", ...) that normalize_keterangan() already collapses
        # away, so category detection must run on the RAW label.
        r['kategori'] = categorize(raw_ket, r.get('objek'), r.get('catatan'), debit, kredit)
        r['keterangan'] = normalize_keterangan(raw_ket, debit, kredit)

        counterparty = resolve_party(r.get('objek', ''), self_code, entity_code_map)
        if r.get('_is_opening_balance'):
            r['subjek'], r['objek'] = '-', '-'
        elif debit:
            r['subjek'], r['objek'] = (self_code or 'Rekening Ini'), counterparty
        elif kredit:
            r['subjek'], r['objek'] = counterparty, (self_code or 'Rekening Ini')
        else:
            r['subjek'], r['objek'] = '', counterparty
    return rows


def populate_sheet(ws, rows, self_code='', entity_code_map=None, saldo_awal=None, saldo_akhir=None):
    """Fills one worksheet with the standard 9-column layout: header, an
    optional opening-balance row, all transaction rows, and a closing
    summary block. Shared by write_xlsx (single account) and
    write_recon_xlsx (multiple accounts, one sheet each)."""
    if rows and 'subjek' not in rows[0]:
        apply_universal_fields(rows, self_code, entity_code_map)

    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(name='Arial', bold=True)
        c.alignment = Alignment(horizontal='center')

    if saldo_awal is not None:
        ws.append(['', 'Saldo Awal', 'Saldo Awal', None, None, saldo_awal, '-', '-', None])

    total_debit, total_kredit = 0.0, 0.0
    for r in rows:
        debit = r.get('debit')
        kredit = r.get('kredit')
        if debit:
            total_debit += debit
        if kredit:
            total_kredit += kredit
        ws.append([
            r['tanggal'], r['keterangan'], r.get('kategori', ''),
            debit, kredit, r.get('saldo'),
            r.get('subjek', ''), r.get('objek', ''), r.get('catatan', ''),
        ])

    data_last_row = ws.max_row
    ws.append([])
    summary_start = ws.max_row + 1
    ws.append(['', 'Saldo Awal', '', None, None, saldo_awal, '', '', ''])
    ws.append(['', 'Total Debit (Uang Keluar)', '', total_debit or None, None, None, '', '', ''])
    ws.append(['', 'Total Kredit (Uang Masuk)', '', None, total_kredit or None, None, '', '', ''])
    final_saldo = saldo_akhir if saldo_akhir is not None else (rows[-1]['saldo'] if rows else saldo_awal)
    ws.append(['', 'Saldo Akhir', '', None, None, final_saldo, '', '', ''])
    for row in ws.iter_rows(min_row=summary_start, max_row=ws.max_row):
        for cell in row:
            if cell.column_letter == 'B':
                cell.font = Font(name='Arial', bold=True)

    for row in ws.iter_rows(min_row=2, max_row=data_last_row, min_col=4, max_col=4):
        for cell in row:
            if cell.value is not None:
                cell.number_format = DEBIT_NUMBER_FORMAT
    for row in ws.iter_rows(min_row=2, max_row=data_last_row, min_col=5, max_col=5):
        for cell in row:
            if cell.value is not None:
                cell.number_format = KREDIT_NUMBER_FORMAT
    for row in ws.iter_rows(min_row=2, max_row=data_last_row, min_col=6, max_col=6):
        for cell in row:
            if cell.value is not None:
                cell.number_format = SALDO_NUMBER_FORMAT
    for row in ws.iter_rows(min_row=summary_start, max_row=ws.max_row, min_col=4, max_col=6):
        for cell in row:
            if cell.value is not None:
                cell.number_format = DEBIT_NUMBER_FORMAT if cell.column_letter == 'D' else (
                    KREDIT_NUMBER_FORMAT if cell.column_letter == 'E' else SALDO_NUMBER_FORMAT)

    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if not cell.font.bold:
                cell.font = Font(name='Arial')


def write_xlsx(rows, out_path, sheet_title='Mutasi', self_code='', entity_code_map=None,
               saldo_awal=None, saldo_akhir=None):
    """rows: list of dicts with keys tanggal, keterangan, debit (negative or
    None), kredit (positive or None), saldo, objek, catatan. This function:
    normalizes keterangan/kategori/subjek/objek uniformly, prepends a Saldo
    Awal row (if saldo_awal is given), and appends a closing summary block
    (Saldo Awal / Total Debit / Total Kredit / Saldo Akhir)."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    populate_sheet(ws, rows, self_code, entity_code_map, saldo_awal, saldo_akhir)
    wb.save(out_path)
    return out_path


def write_recon_xlsx(entries, out_path, include_blank_recon_tab=True):
    """entries: list of dicts, one per account/kantong, each already fully
    processed (rows carry subjek/objek/kategori already), with keys:
    sheet_title, rows, saldo_awal, saldo_akhir. Writes one combined workbook
    with one sheet per entry (each in the same 9-column layout), plus an
    empty 'Rekonsiliasi' tab at the end as a starting point for manual
    cross-referencing."""
    wb = Workbook()
    wb.remove(wb.active)
    used_titles = set()
    for e in entries:
        title = (e['sheet_title'] or 'Sheet')[:31]
        base, n = title, 2
        while title in used_titles:
            suffix = f' ({n})'
            title = base[:31 - len(suffix)] + suffix
            n += 1
        used_titles.add(title)
        ws = wb.create_sheet(title=title)
        populate_sheet(ws, e['rows'], saldo_awal=e.get('saldo_awal'), saldo_akhir=e.get('saldo_akhir'))
    if include_blank_recon_tab:
        wb.create_sheet(title='Rekonsiliasi')
    wb.save(out_path)
    return out_path
