import pdfplumber

from . import bca, bri, jago


def sniff_bank(pdf_path):
    """Look at the first page's text and return one of:
    'bca', 'bri', 'jago', or None if unrecognised."""
    with pdfplumber.open(pdf_path) as pdf:
        text = (pdf.pages[0].extract_text() or '').upper()

    if 'REKENING TAHAPAN' in text or 'REKENING GIRO' in text or ('BCA' in text and 'MUTASI' in text):
        return 'bca'
    if 'LAPORAN TRANSAKSI FINANSIAL' in text or 'BRIMO' in text or 'IBBIZ' in text or 'BANK BRI' in text:
        return 'bri'
    if 'BANK JAGO' in text or 'KANTONG' in text:
        return 'jago'
    return None


PARSERS = {
    'bca': bca,
    'bri': bri,
    'jago': jago,
}

BANK_LABELS = {
    'bca': 'BCA',
    'bri': 'BRI',
    'jago': 'Bank Jago',
}


def parse_statement(pdf_path, out_path):
    """Detect the bank, run the matching parser, write the XLSX.
    Returns (bank_key, extra_info) where extra_info is a dict with any
    warnings/validation notes the parser produced (varies per bank)."""
    bank = sniff_bank(pdf_path)
    if bank is None:
        raise ValueError(
            'Format PDF ini belum aku kenali (bukan BCA/BRI/Jago yang sudah '
            'didukung). Kirim contoh PDF-nya biar aku bikinkan parser baru.'
        )

    mod = PARSERS[bank]
    info = {}
    saldo_awal = saldo_akhir = None

    if bank == 'bca':
        rows, warnings, saldo_awal, saldo_akhir, self_code = mod.build_rows(pdf_path)
        info['warnings'] = warnings
    elif bank == 'bri':
        rows, saldo_awal, saldo_akhir, self_code = mod.build_rows(pdf_path)
    elif bank == 'jago':
        rows, saldo_awal, saldo_akhir, warning = mod.build_rows(pdf_path)
        info['warning'] = warning

    info['saldo_awal'] = saldo_awal
    info['saldo_akhir'] = saldo_akhir

    from .common import write_xlsx
    write_xlsx(rows, out_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)
    info['jumlah_baris'] = len(rows)
    return bank, info
