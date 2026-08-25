# Bank Statement → XLSX Telegram Bot

Bot Telegram terpisah (bukan bagian dari stoabot) yang menerima upload PDF
rekening koran dan membalas dengan file XLSX 7 kolom:
Tanggal, Keterangan Transaksi, Debit, Kredit, Saldo Kumulatif,
Objek Transaksi, Keterangan Tambahan.

## Bank yang didukung saat ini
- BCA (Rekening Tahapan & Rekening Giro)
- BRI (BRImo personal & IBIZ business)
- Bank Jago (kantong-kantong virtual otomatis dikonsolidasi jadi satu akun,
  transfer internal antar kantong dibuang)

Deteksi bank otomatis dari isi PDF (lihat `parsers/detect.py`).

## Struktur project
```
bot.py                  # entrypoint Telegram bot
parsers/
  common.py              # writer XLSX bersama
  detect.py              # deteksi bank + dispatch ke parser yang sesuai
  bca.py                  # parser BCA (Tahapan & Giro)
  bri.py                  # parser BRI (BRImo & IBIZ)
  jago.py                 # parser Bank Jago (konsolidasi kantong)
requirements.txt
Procfile
```

## Deploy ke Railway
1. Buat project Railway baru (terpisah dari project stoabot yang sudah ada).
2. Push folder ini ke repo Git baru, hubungkan ke Railway.
3. Buat bot Telegram baru lewat @BotFather, salin token-nya.
4. Di Railway → Variables, set `BOT_TOKEN` = token dari BotFather.
5. Railway otomatis detect `Procfile` dan jalankan `python bot.py` sebagai worker.
6. Selesai — chat bot-nya di Telegram, kirim PDF rekening koran.

## Menambah bank baru
1. Kirim contoh PDF rekening koran bank tersebut.
2. Parser baru dibuat di `parsers/<nama_bank>.py` dengan fungsi
   `build_rows(pdf_path)` yang mengembalikan list dict dengan key:
   `tanggal, keterangan, debit, kredit, saldo, objek, catatan`.
3. Tambahkan marker deteksi di `parsers/detect.py` (`sniff_bank`) dan
   daftarkan di `PARSERS`/`BANK_LABELS`.

## Catatan akurasi parsing
- Kolom Objek Transaksi & Keterangan Tambahan hasil heuristik dari teks PDF
  (bukan kolom terstruktur asli), jadi sesekali perlu dicek manual —
  terutama untuk format transaksi yang belum pernah muncul di contoh.
- Saldo Kumulatif untuk BCA dihitung ulang transaksi-demi-transaksi lalu
  dicocokkan ke checkpoint saldo yang dicetak PDF; kalau tidak cocok akan
  muncul warning di balasan bot.
- Untuk Bank Jago, saldo hasil konsolidasi dicocokkan ke saldo akhir
  statement; selisih kecil (recehan) bisa muncul karena pembulatan bunga
  di kantong-kantong kecil yang sudah dibulatkan duluan oleh Jago di PDF.
