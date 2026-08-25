from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

HEADERS = [
    'Tanggal', 'Keterangan Transaksi', 'Kategori Transaksi', 'Debit', 'Kredit',
    'Saldo Kumulatif', 'Objek Transaksi', 'Keterangan Tambahan',
]

COL_WIDTHS = [12, 20, 26, 14, 14, 16, 26, 45]

# --- kategori transaksi untuk kebutuhan akuntansi operasional Stoa -------
# Heuristik berbasis kata kunci pada jenis transaksi/objek/catatan. Tidak
# akan 100% akurat untuk semua kasus -- anggap ini kategori awal yang masih
# perlu dicek ulang untuk baris yang jatuh ke "Lainnya / Perlu Verifikasi".

OWNER_KEYWORDS = ('AHMAD ROZIYAN', 'STOASPACE', 'STOA COFFEE', 'STOA SPACE', 'PT STOASPACE')
DEBT_KEYWORDS = ('CICILAN', 'ANGSURAN', 'BAYAR SB', 'PINJAM', 'UTANG')
UTILITY_KEYWORDS = ('SEWA', 'LISTRIK', ' PLN', 'AIR STO', 'UTILITAS')
WALLET_KEYWORDS = ('SHOPEE', 'OVO ', ' OVO', 'GOPAY', 'DANA ', 'TELKOMSEL', 'TOP UP', 'ISI SALDO', 'PULSA')
PURCHASE_KEYWORDS = ('BELANJA', 'BELI ', 'ONGKIR', 'BAHAN', 'SUPPLIER', 'GANTI UANG BELANJA')
TRANSFER_TYPES = ('TRANSFER', 'TRSF', 'BI-FAST', 'SWITCHING', 'KIRIM')

CATEGORIES_REFERENCE = [
    'Penjualan / Pendapatan Usaha',
    'Pembelian Bahan & Operasional',
    'Sewa & Utilitas',
    'Cicilan & Utang',
    'Modal & Setoran Pemilik',
    'Biaya Admin & Pajak Bank',
    'Top Up Dompet Digital',
    'Bunga Bank (Pendapatan)',
    'Transfer Lainnya',
    'Lainnya / Perlu Verifikasi',
]


def categorize(keterangan, objek, catatan, debit, kredit):
    ket = (keterangan or '').upper()
    text = f"{keterangan or ''} {objek or ''} {catatan or ''}".upper()

    if ket == 'BUNGA':
        return 'Bunga Bank (Pendapatan)'
    if ket == 'SALDO AWAL':
        return 'Saldo Awal (Bukan Transaksi)'
    if ket in ('PAJAK BUNGA', 'BIAYA ADMIN', 'BIAYA ADM'):
        return 'Biaya Admin & Pajak Bank'
    if ket in ('PENJUALAN QRIS', 'KR OTOMATIS'):
        return 'Penjualan / Pendapatan Usaha'
    if kredit and 'QRIS' in text:
        return 'Penjualan / Pendapatan Usaha'
    if any(k in text for k in DEBT_KEYWORDS):
        return 'Cicilan & Utang'
    if any(k in text for k in UTILITY_KEYWORDS):
        return 'Sewa & Utilitas'
    if any(k in text for k in OWNER_KEYWORDS) or 'TARIK TUNAI' in text or 'SETOR' in text:
        return 'Modal & Setoran Pemilik'
    if any(k in text for k in WALLET_KEYWORDS):
        return 'Top Up Dompet Digital'
    if any(k in text for k in PURCHASE_KEYWORDS):
        return 'Pembelian Bahan & Operasional'
    if any(k in ket for k in TRANSFER_TYPES):
        return 'Transfer Lainnya'
    return 'Lainnya / Perlu Verifikasi'


def apply_categories(rows):
    for r in rows:
        r['kategori'] = categorize(
            r.get('keterangan'), r.get('objek'), r.get('catatan'),
            r.get('debit'), r.get('kredit'),
        )
    return rows


def write_xlsx(rows, out_path, sheet_title='Mutasi'):
    """rows: list of dicts with keys tanggal, keterangan, debit, kredit,
    saldo, objek, catatan (kategori is filled in by apply_categories if not
    already present). Shared by every bank parser."""
    if rows and 'kategori' not in rows[0]:
        apply_categories(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(name='Arial', bold=True)
        c.alignment = Alignment(horizontal='center')

    for r in rows:
        ws.append([
            r['tanggal'], r['keterangan'], r.get('kategori', ''),
            r.get('debit'), r.get('kredit'), r.get('saldo'),
            r.get('objek', ''), r.get('catatan', ''),
        ])

    for row in ws.iter_rows(min_row=2, min_col=4, max_col=6):
        for cell in row:
            if cell.value is not None:
                cell.number_format = '#,##0.00'

    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name='Arial')

    wb.save(out_path)
    return out_path
