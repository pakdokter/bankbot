import os
import pdfplumber

from . import bca, bri, jago
from .common import write_xlsx, build_filename


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


def parse_statement(pdf_path, out_dir):
    """Detect the bank, run the matching parser, write the XLSX into out_dir
    using the "<Nama Kantong> <Bulan> <Tahun>.xlsx" naming convention (e.g.
    "BCA-887 Januari 2025.xlsx"). Returns (bank_key, extra_info, out_path,
    rows, meta) — rows/meta are handy for callers that want to combine
    several statements into one workbook (see common.write_recon_xlsx)
    without re-parsing the PDF."""
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
        rows, warnings, saldo_awal, saldo_akhir, self_code, meta = mod.build_rows(pdf_path)
        info['warnings'] = warnings
    elif bank == 'bri':
        rows, saldo_awal, saldo_akhir, self_code, meta = mod.build_rows(pdf_path)
    elif bank == 'jago':
        rows, saldo_awal, saldo_akhir, warning, meta = mod.build_rows(pdf_path)
        info['warning'] = warning

    info['saldo_awal'] = saldo_awal
    info['saldo_akhir'] = saldo_akhir

    filename = build_filename(meta['self_code'], meta['bulan'], meta['tahun'])
    out_path = os.path.join(out_dir, filename)
    write_xlsx(rows, out_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)
    info['jumlah_baris'] = len(rows)
    return bank, info, out_path, rows, meta
