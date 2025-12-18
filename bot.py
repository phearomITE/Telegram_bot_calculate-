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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# In‑memory storage: {sheet_name: [row_dict, ...]}
SHEET_ROWS: dict[str, list[dict]] = {}

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
    if not update.message or not update.message.text:
        return

    text = update.message.text
    logger.info("User %s sent: %s", update.effective_user.id, text)

    # Simple greeting example
    if text.lower().strip() in {"hi", "hello", "hey"}:
        await update.message.reply_text(
            "Hi! Send product data in the template format shown in /start."
        )
        return

    # Try to detect product blocks
    parts = text.split('---')
    blocks = [b for b in parts if 'Date:' in b]

    # If the message is not in product format, just ignore silently
    # (no 'No valid product data found' message)
    if not blocks:
        return

    try:
        new_count = 0
        global SHEET_ROWS

        for block in blocks:
            parsed = parse_message(block)
            calc = calculate_fields(parsed)
            sheet_name = choose_sheet_name(calc)

            SHEET_ROWS.setdefault(sheet_name, [])
            row_dict = _row_from_data(calc)
            SHEET_ROWS[sheet_name].append(row_dict)
            new_count += 1

        excel_bytes = build_excel_from_sheet_dict(SHEET_ROWS)
        total_rows = sum(len(v) for v in SHEET_ROWS.values())

        await update.message.reply_document(
            document=InputFile(excel_bytes, filename="calculation_result.xlsx"),
            caption=(
                f"Saved {new_count} new product(s). "
                f"Excel now has {total_rows} product(s) split by style and Outlet-Type (WS/RT)."
            ),
        )
    except Exception as e:
        logger.exception("Error processing message")
        await update.message.reply_text(f"Error: {e}")

# ---------------- Main ----------------
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))

    # Track all text messages
    app.add_handler(MessageHandler(filters.TEXT, handle_text))

    app.run_polling()

if __name__ == "__main__":
    main()
