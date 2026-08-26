import logging
import os
import tempfile

from telegram import Update
from telegram.constants import ChatAction
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

from parsers.detect import parse_statement, BANK_LABELS
from parsers import kasir as kasir_parser
from parsers.common import write_xlsx, build_filename

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get('BOT_TOKEN', '')

WELCOME = (
    "Halo! Kirim aku:\n"
    "- PDF rekening koran (BCA, BRI, atau Bank Jago), atau\n"
    "- XLSX rekap kasir (format: Bulan, Tanggal, Sesi, Keterangan, Kategori, "
    "Debit, Kredit, Saldo, Flag)\n\n"
    "Nanti aku ubah jadi XLSX format seragam: Tanggal, Keterangan Transaksi, "
    "Kategori Transaksi, Debit, Kredit, Saldo Kumulatif, Subjek Transaksi, "
    "Objek Transaksi, Keterangan Tambahan."
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(WELCOME)


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE, doc, tmp):
    pdf_path = os.path.join(tmp, doc.file_name)

    tg_file = await doc.get_file()
    await tg_file.download_to_drive(pdf_path)

    bank, info, xlsx_path = parse_statement(pdf_path, tmp)
    caption_lines = [
        f"Bank terdeteksi: {BANK_LABELS.get(bank, bank)}",
        f"Total baris transaksi: {info.get('jumlah_baris')}",
    ]
    if info.get('warnings'):
        caption_lines.append(f"⚠️ {len(info['warnings'])} baris saldo tidak cocok checkpoint, cek manual.")
    if info.get('warning'):
        caption_lines.append(f"⚠️ {info['warning']}")
    return xlsx_path, caption_lines


async def handle_kasir(update: Update, context: ContextTypes.DEFAULT_TYPE, doc, tmp):
    src_path = os.path.join(tmp, doc.file_name)

    tg_file = await doc.get_file()
    await tg_file.download_to_drive(src_path)

    rows, saldo_awal, saldo_akhir, info, meta = kasir_parser.build_rows(src_path)
    filename = build_filename(meta['self_code'], meta['bulan'], meta['tahun'])
    xlsx_path = os.path.join(tmp, filename)
    write_xlsx(rows, xlsx_path, saldo_awal=saldo_awal, saldo_akhir=saldo_akhir)

    caption_lines = [
        "Kasir Stoa Space (dikonversi ke format seragam)",
        f"Total baris transaksi: {len(rows)}",
        f"Toko dikenali: {info['toko_dikenali']}, Tenant Lain: {info['tenant_lain']}",
    ]
    return xlsx_path, caption_lines


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return
    name = doc.file_name.lower()
    is_pdf = name.endswith('.pdf')
    is_xlsx = name.endswith('.xlsx')
    if not (is_pdf or is_xlsx):
        await update.message.reply_text('Kirim file PDF (rekening koran) atau XLSX (rekap kasir) ya.')
        return

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action=ChatAction.TYPING)
    status_msg = await update.message.reply_text('Lagi diproses...')

    with tempfile.TemporaryDirectory() as tmp:
        try:
            if is_pdf:
                xlsx_path, caption_lines = await handle_pdf(update, context, doc, tmp)
            else:
                xlsx_path, caption_lines = await handle_kasir(update, context, doc, tmp)
        except ValueError as e:
            await status_msg.edit_text(str(e))
            return
        except Exception:
            logger.exception('Gagal memproses %s', doc.file_name)
            await status_msg.edit_text(
                'Gagal memproses file ini. Kemungkinan formatnya sedikit beda dari '
                'yang sudah aku pelajari — kirim ke admin untuk dicek.'
            )
            return

        await status_msg.edit_text('\n'.join(caption_lines))
        with open(xlsx_path, 'rb') as f:
            await update.message.reply_document(document=f, filename=os.path.basename(xlsx_path))


def main():
    if not BOT_TOKEN:
        raise SystemExit('BOT_TOKEN env var belum diset.')

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(MessageHandler(filters.Document.PDF | filters.Document.FileExtension('xlsx'), handle_document))

    logger.info('Bot jalan...')
    app.run_polling()


if __name__ == '__main__':
    main()
