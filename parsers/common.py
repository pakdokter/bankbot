from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

HEADERS = [
    'Tanggal', 'Keterangan Transaksi', 'Debit', 'Kredit',
    'Saldo Kumulatif', 'Objek Transaksi', 'Keterangan Tambahan',
]

COL_WIDTHS = [12, 22, 14, 14, 16, 26, 45]


def write_xlsx(rows, out_path, sheet_title='Mutasi'):
    """rows: list of dicts with keys tanggal, keterangan, debit, kredit,
    saldo, objek, catatan. Shared by every bank parser so the Telegram bot
    only needs one writer regardless of which parser produced the rows."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(name='Arial', bold=True)
        c.alignment = Alignment(horizontal='center')

    for r in rows:
        ws.append([
            r['tanggal'], r['keterangan'],
            r.get('debit'), r.get('kredit'), r.get('saldo'),
            r.get('objek', ''), r.get('catatan', ''),
        ])

    for row in ws.iter_rows(min_row=2, min_col=3, max_col=5):
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
