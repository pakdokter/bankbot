import logging
import os
import tempfile

from telegram import Update
from telegram.constants import ChatAction
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

from parsers.detect import parse_statement, BANK_LABELS

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get('BOT_TOKEN', '')

WELCOME = (
    "Halo! Kirim aku PDF rekening koran (BCA, BRI, atau Bank Jago), "
    "nanti aku ubah jadi file XLSX dengan kolom:\n"
    "Tanggal, Keterangan Transaksi, Debit, Kredit, Saldo Kumulatif, "
    "Objek Transaksi, Keterangan Tambahan."
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(WELCOME)


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.file_name.lower().endswith('.pdf'):
        await update.message.reply_text('Kirim file PDF ya, format lain belum didukung.')
        return

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action=ChatAction.TYPING)
    status_msg = await update.message.reply_text('Lagi diproses...')

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, doc.file_name)
        xlsx_path = os.path.join(tmp, os.path.splitext(doc.file_name)[0] + '.xlsx')

        tg_file = await doc.get_file()
        await tg_file.download_to_drive(pdf_path)

        try:
            bank, info = parse_statement(pdf_path, xlsx_path)
        except ValueError as e:
            await status_msg.edit_text(str(e))
            return
        except Exception:
            logger.exception('Parsing gagal untuk %s', doc.file_name)
            await status_msg.edit_text(
                'Gagal parsing PDF ini. Kemungkinan formatnya sedikit beda dari '
                'yang sudah aku pelajari — kirim ke admin untuk dicek.'
            )
            return

        caption_lines = [
            f"Bank terdeteksi: {BANK_LABELS.get(bank, bank)}",
            f"Total baris transaksi: {info.get('jumlah_baris')}",
        ]
        if info.get('warnings'):
            caption_lines.append(f"⚠️ {len(info['warnings'])} baris saldo tidak cocok checkpoint, cek manual.")
        if info.get('warning'):
            caption_lines.append(f"⚠️ {info['warning']}")

        await status_msg.edit_text('\n'.join(caption_lines))
        with open(xlsx_path, 'rb') as f:
            await update.message.reply_document(document=f, filename=os.path.basename(xlsx_path))


def main():
    if not BOT_TOKEN:
        raise SystemExit('BOT_TOKEN env var belum diset.')

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_document))

    logger.info('Bot jalan...')
    app.run_polling()


if __name__ == '__main__':
    main()
