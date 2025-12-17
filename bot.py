import logging
from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters,
)

from config import BOT_TOKEN
from parser import parse_message
from excel_builder import (
    calculate_fields,
    choose_sheet_name,
    _row_from_data,
    build_excel_from_sheet_dict,
)
from db import init_db, insert_row, fetch_all_rows_by_sheet

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXAMPLE_TEXT = (
    "--- product 1 ---\n"
    "Date: 24.11.2025\n"
    "Addresss: ចំការគ\n"
    "Outlet-Type: WS\n"
    "Category: Oil\n"
    "Sub-Category: Soybean\n"
    "Brand: Simply\n"
    "Packaging: Bottle\n"
    "Size: 1000ml\n"
    "Packs: 12\n"
    "Weight per Ctn: 12L\n"
    "Buy-in: 15.90$\n"
    "Scheme(base): 4\n"
    "FOC: 1\n"
    "Discount(%): 0\n"
    "Discount($): 0\n"
    "Direct Disc.(%): 0\n"
    "Direct Disc($): 0\n"
    "Mark - up: 0.5\n"
    "Sell Out ($): 14.5$\n"
    "Exchange Rate: 4000\n"
    "Price Unit: 1000\n"
)

# ---------------- Handlers ----------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Send one or many products.\n"
        "Separate products with '--- product N ---' lines.\n\n"
        "Example:\n\n" + EXAMPLE_TEXT
    )

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    try:
        parts = text.split('---')
        blocks = [b for b in parts if 'Date:' in b]

        if not blocks:
            await update.message.reply_text("No valid product data found.")
            return

        new_count = 0
        for block in blocks:
            parsed = parse_message(block)
            calc = calculate_fields(parsed)
            sheet_name = choose_sheet_name(calc)
            insert_row(sheet_name, calc)
            new_count += 1

        sheet_rows_raw = fetch_all_rows_by_sheet()
        sheet_rows = {s: [_row_from_data(r) for r in rows] for s, rows in sheet_rows_raw.items()}

        excel_bytes = build_excel_from_sheet_dict(sheet_rows)
        total_rows = sum(len(v) for v in sheet_rows.values())

        await update.message.reply_document(
            document=InputFile(excel_bytes, filename="calculation_result.xlsx"),
            caption=f"Saved {new_count} new product(s). Excel now has {total_rows} product(s) split by style and Outlet-Type (WS/RT)."
        )
    except Exception as e:
        logger.exception("Error processing message")
        await update.message.reply_text(f"Error: {e}")

# ---------------- Main ----------------
def main():
    init_db()
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.run_polling()

if __name__ == "__main__":
    main()
