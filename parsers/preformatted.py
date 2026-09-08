"""Parser for XLSX files that are already in this bot's own 9-column output
format (e.g. a manually reconstructed old statement, or a previous bot
output the user edited/annotated). Column order is detected by header name
so extra analysis columns (e.g. a manual "Selisih vs Saldo Tercatat" check)
can sit alongside ours without breaking anything."""
import re
import sys
from collections import Counter
import openpyxl

from .common import HEADERS, write_xlsx, build_filename, month_name, match_gaji

FIELD_KEYS = ['tanggal', 'keterangan', 'kategori', 'debit', 'kredit', 'saldo', 'subjek', 'objek', 'catatan']
HEADER_NAME_MAP = {h.upper(): key for h, key in zip(HEADERS, FIELD_KEYS)}
SELISIH_NAMES = {'SELISIH', 'SELISIH VS SALDO TERCATAT'}

# --- aturan yang sudah dikonfirmasi lewat feedback -- berlaku untuk semua
# dokumen "sudah diolah" berikutnya, bukan cuma satu bulan tertentu ---

# nama yang selalu berarti "Modal Masuk" (uang masuk dari pemilik/keluarga,
# bukan penjualan)
MODAL_MASUK_NAMES = ('AHMAD RIZAN HENDRA',)

# alias merchant/tenant tunggal (nama panjang -> nama pendek kanonik)
TENANT_ALIASES = {
    'PRIMER RAYA': 'Primer',
    'PRIMER': 'Primer',
}
OWNER_MARKERS = ('OWNER', 'ROZIYAN HIDAYAT', 'OJAN', 'KAK OJAN', 'AHMAD ROZIYAN HIDAYAT')
GENERIC_UNRESOLVED_CATEGORIES = {'TRANSFER KELUAR', 'PENGELUARAN', 'TRANSFER LAINNYA'}

# kata kunci -> (kategori, objek_atau_None). Dicek pada gabungan teks
# Keterangan + Keterangan Tambahan, tidak case-sensitive, dengan word
# boundary supaya tidak salah tangkap ("Web" tidak match "Website" dst
# kalau memang perlu lebih ketat -- di sini cukup longgar karena datanya
# manual/singkat).
KEYWORD_RULES = [
    (r'MADAM', 'Belanja Bahan', 'Madam'),
    (r'BEANS', 'Belanja Bahan', None),
    (r'SHOPEE', 'Belanja Bahan', 'Shopee'),
    (r'SINAR BAHAGIA', 'Belanja Bahan', 'Sinar Bahagia'),
    (r'KONSUMSI', 'OpEx', None),
    (r'\bWEB\b', 'Overhead', None),
    (r'UTILITIES', 'Overhead', None),
    (r'SPOTIFY', 'Overhead', None),
    (r'TELKOM', 'Overhead', None),
    (r'MR\s*DIY', None, 'MR DIY'),
    (r'PELATIHAN', 'Riset dan Pengembangan', None),
    (r'TARIKAN?\s*ATM', 'OpEx', None),
    (r'INDOMARET', None, 'Indomaret'),
    (r'MASUYA', None, 'Masuya'),
    (r'ANUGERAH', None, 'Anugerah'),
    (r'AMANAH', 'Belanja Bahan', 'Amanah'),
    (r'APOTEK\s*SURYA\s*FARMA|SURYA\s*FARMA', None, 'Apotek Surya Farma'),
    (r'APOTEK\s*YODAN|\bYODAN\b', None, 'Apotek Yodan'),
    (r'BINTANG\s*PLASTIK', None, 'Bintang Plastik'),
    (r'FADHILAH', 'Belanja Bahan', 'Fadhilah'),
    (r'MAK\s*OPIK|MAH\s*OPIK', 'Belanja Bahan', 'Mak Opik'),
    (r'PASAR\s*PANCOR|\bPASAR\b', 'Belanja Bahan', 'Pasar'),
]
KEYWORD_RULES = [(re.compile(pat, re.I), kat, obj) for pat, kat, obj in KEYWORD_RULES]

KONSUMSI_RE = re.compile(r'KONSUMSI', re.I)
PENJUALAN_RE = re.compile(r'PENJUALAN', re.I)
BUNGA_RE = re.compile(r'BUNGA', re.I)
BIAYA_ADMIN_RE = re.compile(r'BIAYA\s*ADMIN|PAJAK|ADMIN\s*TRANSFER', re.I)
FLIPTECH_RE = re.compile(r'FLIPTECH', re.I)
PRIMER_RE = re.compile(r'PRIMER(?:\s*RAYA)?', re.I)
MODAL_MASUK_KETERANGAN_RE = re.compile(r'^MODAL\s*MASUK$', re.I)
KOREKSI_RE = re.compile(r'KOREKSI', re.I)
# "Setoran Via CDM", "Setoran Tunai", "Setor tunai -> BCA", "Setoran ke BCA", dst
SETORAN_RE = re.compile(r'SETOR(?:AN)?\s*TUNAI|SETORAN', re.I)
BANK_NAME_RE = re.compile(r'\b(BCA|BRI|MANDIRI|JAGO|BSI|BNI|CIMB)\b', re.I)

# --- kas buku: pisahkan parkir dari belanja, dan tarik nama tenant dari
# pola "Vendor – Item" (dikonfirmasi lewat revisi Kas Buku Februari 2025) --
REFUND_KETERANGAN_RE = re.compile(r'^REFUND\b', re.I)
COD_RE = re.compile(r'^COD\b', re.I)
PARKIR_EXACT_RE = re.compile(r'^PARKIR$', re.I)
PLUS_PARKIR_RE = re.compile(r'^(.*?)\s*\+\s*Parkir\s*$', re.I)
VENDOR_ITEM_RE = re.compile(r'^(.*?)\s*[–-]\s*(.+)$')
PARKIR_AMOUNT = 2000.0
ES_BATU_ESTIMATE = 12000.0
VENDOR_NAME_ALIASES = {'MR. DIY': 'Mr. DIY', 'MR DIY': 'Mr. DIY', 'AMANAH': 'Amanah'}


def _vendor_from_belanja(desc):
    """desc tanpa 'Belanja ' di depan, mis. 'Kiki – Mak Opik Bawang',
    'Sekar Dinda Frozen – Kentang', 'Gia Abadi – Es Batu', 'Luna (Abadi) – Es Batu x2',
    'Upi (AMANAH) – Creamch.', 'Kurnia – Plastik'. Returns (employee, vendor_or_None, item)."""
    m = re.match(r'^(\S+)\s*(?:\(([^)]+)\))?\s*(.*)$', desc)
    if not m:
        return desc, None, ''
    employee, paren_vendor, rest = m.groups()
    rest = rest.strip()
    if paren_vendor:
        item = re.sub(r'^[–-]\s*', '', rest)
        return employee, paren_vendor.strip(), item
    m2 = VENDOR_ITEM_RE.match(rest)
    if m2:
        maybe_vendor, item = m2.groups()
        maybe_vendor = maybe_vendor.strip()
        if maybe_vendor:
            return employee, maybe_vendor, item.strip()
        return employee, None, item.strip()
    return employee, None, rest


def _apply_keyword_overrides(keterangan, kategori, objek, catatan, is_kredit=False):
    """Returns (keterangan, kategori, objek) after applying every keyword
    rule confirmed via feedback. Order matters: more specific rules first."""
    text = f'{keterangan} {catatan}'.upper()
    ob_upper = (objek or '').strip().upper()

    if ob_upper in MODAL_MASUK_NAMES or MODAL_MASUK_KETERANGAN_RE.match(keterangan.strip()):
        return 'Modal Masuk', 'Modal & Setoran Pemilik', objek

    if 'BELANJA PRIBADI' in text or 'PEMBAYARAN PRIBADI' in text:
        return 'Belanja Pribadi', 'Belanja Pribadi', objek

    if KOREKSI_RE.search(keterangan):
        return 'Tip/Minus', 'Tip/Minus/Lebih', objek

    if REFUND_KETERANGAN_RE.match(keterangan.strip()):
        return keterangan, 'Penjualan', objek

    if COD_RE.match(keterangan.strip()):
        return 'Belanja Shopee', 'Belanja Bahan', 'Shopee'

    if PARKIR_EXACT_RE.match(keterangan.strip()):
        # parkir tidak pernah terkait tenant/vendor transaksi sebelumnya --
        # objek selalu dinetralkan, apa pun yang kebetulan ada di kolom itu
        return keterangan, 'OpEx', 'Tenant Lain'

    gaji = match_gaji(keterangan)
    if gaji:
        employee_name, _bulan_gaji = gaji
        # gaji milik owner yang MASUK (bukan Stoa menggaji dia) sebenarnya
        # penghasilan luar yang disuntikkan sebagai modal, bukan beban gaji
        if is_kredit and any(m in employee_name.upper() for m in OWNER_MARKERS):
            return 'Modal Masuk', 'Modal & Setoran Pemilik', employee_name
        return keterangan, 'Gaji Pegawai', employee_name

    if PENJUALAN_RE.search(kategori) or PENJUALAN_RE.search(keterangan):
        return 'Penjualan', 'Penjualan', objek

    # cek juga kategori ASLI dari sumber -- beberapa file manual sudah
    # menulis "Biaya Admin"/"Pajak Bank"/dst di kolom kategori sendiri
    # walau keterangannya tidak secara harfiah menyebut kata itu
    admin_bunga_text = f'{text} {kategori}'.upper()
    if BIAYA_ADMIN_RE.search(admin_bunga_text):
        return 'Biaya Admin', 'Biaya Admin & Pajak Bank', objek
    if BUNGA_RE.search(admin_bunga_text):
        return 'Bunga Bank', 'Biaya Admin & Pajak Bank', objek

    if SETORAN_RE.search(text):
        # setoran tunai kasir <-> bank = perpindahan antar "kantong" Stoa
        # sendiri. Kalau nama bank tujuannya kesebut eksplisit, itu jadi
        # objek (rekening ini yang mengirim); kalau tidak, rekening ini
        # dianggap sisi penerima (bank), objek jadi self_code -- diselesaikan
        # di build_rows setelah self_code diketahui.
        m_bank = BANK_NAME_RE.search(text)
        target = m_bank.group(1).upper() if m_bank else objek
        return keterangan, 'Transaksi Internal', target

    if PRIMER_RE.search(text):
        return 'Belanja Bahan', 'Belanja Bahan', 'Primer'

    new_keterangan = 'Belanja Konsumsi' if KONSUMSI_RE.search(text) else keterangan
    new_kategori, new_objek = kategori, objek
    for pattern, kat, obj in KEYWORD_RULES:
        if pattern.search(text):
            if kat:
                new_kategori = kat
            if obj and (not new_objek or new_objek == '-'):
                new_objek = obj

    # kategori generik/tidak jelas yang belum kena aturan spesifik apa pun di
    # atas -- selama uangnya keluar, anggap sebagai belanja operasional
    # biasa daripada dibiarkan sebagai label transfer mentah
    if new_kategori.strip().upper() in GENERIC_UNRESOLVED_CATEGORIES and not is_kredit:
        new_kategori = 'OpEx'

    # tarik nama tenant dari pola "Belanja <Karyawan> [Vendor] – Item" kalau
    # belum kena aturan spesifik apa pun di atas (mis. Dinda Frozen, Abadi --
    # vendor yang belum masuk KEYWORD_RULES, atau memang tidak ada vendornya
    # sama sekali sehingga nama karyawan dipakai)
    objek_unresolved = not new_objek or new_objek in ('-', 'Tenant Lain')
    keterangan_masih_asli = new_keterangan == keterangan
    if objek_unresolved and keterangan_masih_asli and re.match(r'^BELANJA\s+', keterangan, re.I) \
            and not re.match(r'^BELANJA\s+(BAHAN|OPERASIONAL|KONSUMSI)\b', keterangan, re.I):
        employee, vendor, item = _vendor_from_belanja(re.sub(r'^BELANJA\s+', '', keterangan, flags=re.I).strip())
        if vendor:
            vendor_norm = VENDOR_NAME_ALIASES.get(vendor.upper(), vendor)
            new_objek = vendor_norm
            new_keterangan = item if item else vendor_norm
        else:
            new_objek = employee
            if item:
                new_keterangan = item
    elif objek_unresolved and keterangan_masih_asli:
        # fallback umum: pola "Vendor – Item" biasa (Bintang, Abadi, Istana
        # Sosis, Toko Buah, Qia Mart, dst -- vendor apa pun yang belum
        # dikenal secara eksplisit lewat KEYWORD_RULES di atas)
        m_vi = VENDOR_ITEM_RE.match(keterangan)
        if m_vi:
            vendor, item = m_vi.groups()
            vendor, item = vendor.strip(), item.strip()
            if vendor and not COD_RE.match(vendor):
                new_objek = vendor
                new_keterangan = item if item else vendor

    # kalaupun vendornya sudah dikenal lewat KEYWORD_RULES (mis. FADHILAH,
    # MAK OPIK) di atas, keterangannya masih baris utuh "Vendor – Item" --
    # sederhanakan jadi item saja karena nama tenantnya sudah pindah ke objek.
    # Hanya kalau vendor yang ketarik dari tanda pisah itu benar-benar cocok
    # sama objek yang sudah ditentukan -- supaya "V-Soy"/"Lap – 8.500" (yang
    # kebetulan ada tanda pisah tapi bukan pola Vendor-Item) tidak ketimpa.
    if keterangan_masih_asli and new_keterangan == keterangan and new_objek and new_objek not in ('-', 'Tenant Lain'):
        m_vi = VENDOR_ITEM_RE.match(keterangan)
        if m_vi:
            vendor, item = m_vi.groups()
            vendor, item = vendor.strip(), item.strip()
            vendor_matches_objek = vendor.upper() in new_objek.upper() or new_objek.upper() in vendor.upper()
            if item and vendor.upper() != 'COD' and vendor_matches_objek:
                new_keterangan = item

    if (new_kategori.strip().lower().startswith('belanja') or new_kategori in ('Overhead', 'OpEx')) and (not new_objek or new_objek == '-'):
        new_objek = 'Tenant Lain'

    return new_keterangan, new_kategori, new_objek


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

    # --- pass 1: baca semua baris mentah + tentukan self_code lebih dulu,
    # dari nilai Subjek/Objek asli (sebelum override apa pun) --------------
    raw_rows = []
    party_counter = Counter()
    month_year_counter = Counter()
    saldo_awal, saldo_akhir = None, None
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
        kategori = str(get(row, 'kategori') or '').strip()
        debit = to_float(get(row, 'debit'))
        kredit = to_float(get(row, 'kredit'))
        saldo = to_float(get(row, 'saldo'))
        subjek = str(get(row, 'subjek') or '')
        objek = str(get(row, 'objek') or '')
        catatan = str(get(row, 'catatan') or '')
        # "Kasir" adalah nama lama sebelum diseragamkan jadi "Kas/Buku" --
        # normalisasi di level baris juga, bukan cuma di self_code, supaya
        # tidak ada sisa label lama yang lolos ke output.
        if subjek.strip().upper() == 'KASIR':
            subjek = 'Kas/Buku'
        if objek.strip().upper() == 'KASIR':
            objek = 'Kas/Buku'

        if kategori.strip().upper().startswith('SALDO AWAL') or keterangan.strip().upper().startswith('SALDO AWAL'):
            # beberapa file manual cuma isi Kredit (atau Debit) untuk baris
            # ini dan biarkan Saldo Kumulatif kosong -- pakai itu sebagai
            # fallback kalau kolom Saldo-nya sendiri tidak terisi
            fallback = kredit if kredit is not None else debit
            saldo_awal = saldo if saldo is not None else (fallback if fallback is not None else saldo_awal)
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

        if selisih_col is not None:
            raw_selisih = row[selisih_col] if selisih_col < len(row) else None
            sv = to_float(raw_selisih)
            if sv and abs(sv) > 0.01:
                selisih_flags += 1

        for p in (subjek, objek):
            p = str(p).strip()
            if p and p != '-':
                party_counter[p] += 1

        raw_rows.append({
            'tanggal': tgl_str, 'keterangan': keterangan, 'kategori': kategori,
            'debit': debit, 'kredit': kredit, 'saldo': saldo,
            'subjek': subjek, 'objek': objek, 'catatan': catatan,
        })

    self_code = party_counter.most_common(1)[0][0] if party_counter else 'Rekening'
    # "Kasir" adalah nama lama sebelum diseragamkan jadi "Kas/Buku" -- file
    # lama yang diupload ulang (atau hasil proses sebelum penyeragaman itu)
    # masih bisa punya label ini persis di datanya, jadi dinormalisasi di
    # sini supaya tidak lolos ke rekonsiliasi dengan nama yang beda sendiri.
    if self_code.strip().upper() == 'KASIR':
        self_code = 'Kas/Buku'
    if month_year_counter:
        (m, y), _ = month_year_counter.most_common(1)[0]
        bulan, tahun = month_name(m), y
    else:
        bulan, tahun = '', ''

    # --- pass 2: terapkan aturan kata kunci + pemecahan Fliptech, sekarang
    # self_code sudah diketahui untuk baris Penjualan/Biaya Admin/Bunga Bank
    rows = []
    running = saldo_awal

    def _emit(row_dict):
        """Tambahkan baris ke rows, dan hitung ulang Saldo Kumulatif kalau
        sumbernya tidak mengisi kolom itu (cuma isi Debit/Kredit). Kalau
        sumbernya MEMANG mengisi Saldo, itu dipercaya sebagai checkpoint dan
        running balance disinkronkan ke situ."""
        nonlocal running
        if running is not None:
            delta = (row_dict['kredit'] or 0) + (row_dict['debit'] or 0)
            running = round(running + delta, 2)
            if row_dict['saldo'] is None:
                row_dict['saldo'] = running
            else:
                running = row_dict['saldo']
        rows.append(row_dict)

    for r in raw_rows:
        tgl_str, keterangan, kategori = r['tanggal'], r['keterangan'], r['kategori']
        debit, kredit, saldo = r['debit'], r['kredit'], r['saldo']
        subjek, objek, catatan = r['subjek'], r['objek'], r['catatan']

        if FLIPTECH_RE.search(keterangan) or FLIPTECH_RE.search(catatan):
            total = debit if debit is not None else kredit
            is_debit = debit is not None
            if total is not None:
                abs_total = abs(total)
                main = float((int(abs_total) // 1000) * 1000)
                remainder = round(abs_total - main, 2)
                _emit({
                    'tanggal': tgl_str, 'keterangan': 'Transfer Internal',
                    'kategori': 'Transaksi Internal',
                    'debit': -main if is_debit else None,
                    'kredit': None if is_debit else main,
                    'saldo': None, 'subjek': subjek,
                    'objek': objek or 'Fliptech',
                    'catatan': catatan,
                })
                if remainder:
                    fee_label = 'Biaya Admin' if is_debit else 'Bunga Bank'
                    _emit({
                        'tanggal': tgl_str, 'keterangan': fee_label, 'kategori': 'Biaya Admin & Pajak Bank',
                        'debit': -remainder if is_debit else None,
                        'kredit': None if is_debit else remainder,
                        'saldo': saldo, 'subjek': '-', 'objek': self_code,
                        'catatan': f'Bagian dari transaksi Fliptech: {keterangan}',
                    })
                continue

        # "Vendor – Item1 & Es Batu + Parkir" atau "Item + Parkir" biasa --
        # pisahkan ongkos parkir (dan Es Batu kalau tergabung) dari belanja
        # utamanya supaya masing-masing kelihatan sendiri-sendiri.
        m_parkir = PLUS_PARKIR_RE.match(keterangan) if debit is not None else None
        if m_parkir:
            base_desc = m_parkir.group(1).strip().rstrip('–- ').strip()
            has_es_batu = bool(re.search(r'&\s*Es\s*Batu', base_desc, re.I))
            if has_es_batu:
                main_item = re.sub(r'\s*&\s*Es\s*Batu', '', base_desc, flags=re.I).strip().rstrip('–- ').strip()
                main_amt = debit + PARKIR_AMOUNT + ES_BATU_ESTIMATE
                m_v = VENDOR_ITEM_RE.match(main_item)
                vendor_guess = m_v.group(1).strip() if m_v else (objek if objek and objek != '-' else None)
                item_only = m_v.group(2).strip() if m_v else main_item
                ket1, kat1, obj1 = _apply_keyword_overrides(main_item, kategori, objek, catatan, is_kredit=False)
                _emit({
                    'tanggal': tgl_str, 'keterangan': ket1, 'kategori': kat1,
                    'debit': main_amt, 'kredit': None, 'saldo': None,
                    'subjek': subjek, 'objek': obj1,
                    'catatan': f'Dipecah dari: {keterangan}',
                })
                es_batu_vendor = obj1 if obj1 and obj1 != 'Tenant Lain' else (vendor_guess or 'Tenant Lain')
                _emit({
                    'tanggal': tgl_str, 'keterangan': 'Es Batu', 'kategori': 'Belanja Bahan',
                    'debit': -ES_BATU_ESTIMATE, 'kredit': None, 'saldo': None,
                    'subjek': subjek, 'objek': es_batu_vendor,
                    'catatan': f'Estimasi harga Es Batu (~Rp{ES_BATU_ESTIMATE:,.0f}), dipecah dari: {keterangan}',
                })
                _emit({
                    'tanggal': tgl_str, 'keterangan': 'Parkir', 'kategori': 'OpEx',
                    'debit': -PARKIR_AMOUNT, 'kredit': None, 'saldo': saldo,
                    'subjek': subjek, 'objek': 'Tenant Lain',
                    'catatan': f'Dipecah dari: {keterangan}',
                })
            else:
                base_amt = debit + PARKIR_AMOUNT
                ket1, kat1, obj1 = _apply_keyword_overrides(base_desc, kategori, objek, catatan, is_kredit=False)
                if obj1 == 'Tenant Lain' and not VENDOR_ITEM_RE.match(base_desc) and not re.match(r'^BELANJA\s+', base_desc, re.I):
                    # base_desc cuma nama vendor polos tanpa rincian item, mis. "Fadhilah"
                    obj1 = base_desc
                _emit({
                    'tanggal': tgl_str, 'keterangan': ket1, 'kategori': kat1,
                    'debit': base_amt, 'kredit': None, 'saldo': None,
                    'subjek': subjek, 'objek': obj1,
                    'catatan': f'Dipecah dari: {keterangan}',
                })
                _emit({
                    'tanggal': tgl_str, 'keterangan': 'Parkir', 'kategori': 'OpEx',
                    'debit': -PARKIR_AMOUNT, 'kredit': None, 'saldo': saldo,
                    'subjek': subjek, 'objek': 'Tenant Lain',
                    'catatan': f'Dipecah dari: {keterangan}',
                })
            continue

        keterangan, kategori, objek = _apply_keyword_overrides(
            keterangan, kategori, objek, catatan, is_kredit=(kredit is not None)
        )

        if kategori == 'Penjualan':
            subjek, objek = 'Penjualan', self_code
        elif kategori == 'Biaya Admin & Pajak Bank':
            subjek, objek = '-', self_code
        elif kategori == 'Transaksi Internal':
            # dari aturan SETORAN_RE di _apply_keyword_overrides: kalau nama
            # bank tujuan disebut eksplisit di teks, rekening ini yang
            # mengirim (objek = bank tujuan); kalau tidak, rekening ini
            # dianggap sisi bank yang menerima setoran dari kas kasir.
            if objek and objek not in ('-', ''):
                subjek = self_code
            else:
                subjek, objek = 'Kas Kasir', self_code

        _emit({
            'tanggal': tgl_str,
            'keterangan': keterangan,
            'kategori': kategori,
            'debit': debit,
            'kredit': kredit,
            'saldo': saldo,
            'subjek': subjek,
            'objek': objek,
            'catatan': catatan,
        })

    if saldo_akhir is None:
        saldo_akhir = rows[-1]['saldo'] if rows else saldo_awal

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
