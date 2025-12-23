import logging
from collections import defaultdict

from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters,
)

from config import BOT_TOKEN, EXCHANGE_RATE_DEFAULT  # make sure EXCHANGE_RATE_DEFAULT exists
from parser import parse_message
from excel_builder import (
    calculate_fields,
    choose_sheet_name,
    _row_from_data,
    build_excel_from_sheet_dict,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# per-sheet rows for Excel (rebuilt from ALL_PRODUCTS)
SHEET_ROWS: dict[str, list[dict]] = {}
# flat list of all parsed products in input order
ALL_PRODUCTS: list[dict] = []

# per-user settings (simple in‑memory example)
USER_SETTINGS: dict[int, dict] = {}

EXAMPLE_TEXT = (
    "--- product 1 ---\n"
    "Date: 24.11.2025\n"
    "Address: ចំការគ\n"
    "Category: Oil\n"
    "Sub-Category: Soybean\n"
    "Brand: Health Pro\n"
    "Packaging: Bottle\n"
    "Size: 1000ml\n"
    "Packs: 12\n"
    "Buy-in: 22.50$                 # ត្រូវតែបំពេញ\n"
    "Scheme(base): 4\n"
    "FOC: 0\n"
    "Direct Disc.(%): 0.0%          # បំពេញក៏បាន អត់ក៏បាន\n"
    "Mark - up: 0.50$               # ត្រូវតែបំពេញ\n"
    "Price Unit: 9000               # ត្រូវតែបំពេញ\n"
    "\n"
    "--- product 2 ---\n"
    "Date: 24.11.2025\n"
    "Address: ចំការគ\n"
    "Category: Milk\n"
    "Sub-Category: Condensed\n"
    "Brand: Phka Chhouk\n"
    "Packaging: Can\n"
    "Size: 390g\n"
    "Packs: 48\n"
    "Buy-in: 28.60$                 # ត្រូវតែបំពេញ\n"
    "Scheme(base): 1\n"
    "FOC: 0\n"
    "Direct Disc.(%): 0.0%          # បំពេញក៏បាន អត់ក៏បាន\n"
    "Mark - up: 1.00$               # ត្រូវតែបំពេញ\n"
    "Price Unit: 3000               # ត្រូវតែបំពេញ\n"
)

def normalize_sheet(name: str) -> str:
    """Lowercase + strip for robust sheet comparison."""
    return (name or "").strip().lower()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Welcome to Price Calculator Bot.\n\n"
        "Send one or many products.\n"
        "Separate products with '--- product N ---' lines.\n\n"
        "Example:\n\n" + EXAMPLE_TEXT + "\n\n"
        "Use /help to see all features."
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "📘 Help – Price Calculator Bot\n\n"
        "Commands:\n"
        "/start – Show example format and how to start.\n"
        "/help – Show this help message.\n"
        "/settings – Change language, default exchange rate, default outlet type, etc.\n"
        "/clear – Clear current session data (Excel will start from 0 product again).\n"
        "/about – Show bot information.\n"
        "/list – Show all products saved in memory with Ids.\n"
        "/delete <Sheet> <Id> – Delete one row from a sheet.\n"
        "/delete_sheet <Sheet> – Delete all rows from a sheet.\n\n"
        "Input format (one product):\n"
        "Date: 24.11.2025\n"
        "Address: ចំការគ\n"
        "Category: Detergent\n"
        "Sub-Category: Powder\n"
        "Brand: Viso\n"
        "Packaging: Pouch\n"
        "Size: 3700 g\n"
        "Packs: 4\n"
        "Buy-in: 21.68$\n"
        "Scheme(base): 1\n"
        "FOC: 0\n"
        "Direct Disc.(%): 12.00%\n"
        "Mark - up: 1.00$\n"
        "Price Unit: 22000\n\n"
        "Main formulas (inside Excel):\n"
        "• Weight per Ctn = (Size × Packs) / 1000 (kg or L).\n"
        "• Discount(%) = FOC / (Scheme(base) + FOC).\n"
        "• Discount($) = Discount(%) × Buy-in.\n"
        "• Direct Disc($) = Direct Disc(%) × Buy-in.\n"
        "• Net Buy-in = Buy-in − (Discount($) + Direct Disc($)).\n"
        "• Price / 100g or 100ml uses Size scale.\n"
        "• Sell Out($) = Net Buy-in + Mark-up.\n"
        "• Sell Out(KHR), Margin/Unit, Price Ctn, Margin/Ctn are calculated from Sell Out($), Packs, and Price Unit.\n\n"
        "Rounding:\n"
        "All calculated values use your rule: look at 3rd decimal; 6–9 round up, 1–5 round down, keep 2 decimals."
    )
    await update.message.reply_text(text)

async def settings_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Very simple /settings: show and allow basic changes with arguments."""
    user_id = update.effective_user.id
    settings = USER_SETTINGS.setdefault(
        user_id,
        {
            "language": "km",
            "default_exchange_rate": EXCHANGE_RATE_DEFAULT,
            "default_outlet_type": "WS",
            "rounding_mode": "custom",  # your 3rd-decimal rule
        },
    )

    # If user sends arguments, allow quick updates, e.g.
    # /settings outlet=RT rate=4100 lang=en
    for arg in context.args:
        if arg.startswith("outlet="):
            settings["default_outlet_type"] = arg.split("=", 1)[1].upper()
        elif arg.startswith("rate="):
            try:
                settings["default_exchange_rate"] = float(arg.split("=", 1)[1])
            except ValueError:
                pass
        elif arg.startswith("lang="):
            settings["language"] = arg.split("=", 1)[1].lower()

    await update.message.reply_text(
        "⚙️ Settings:\n"
        f"Language: {settings['language']}\n"
        f"Default exchange rate: {settings['default_exchange_rate']}\n"
        f"Default outlet type: {settings['default_outlet_type']}\n"
        f"Rounding mode: {settings['rounding_mode']}\n\n"
        "Change values with, for example:\n"
        "/settings outlet=RT rate=4100 lang=en"
    )

async def clear_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Clear current in-memory products (session data)."""
    global ALL_PRODUCTS, SHEET_ROWS
    ALL_PRODUCTS = []
    SHEET_ROWS = {}
    await update.message.reply_text(
        "🧹 Cleared current session data.\n"
        "Next Excel will start from 0 product (only new messages)."
    )

async def about_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📦 Price Calculator Bot v1.0\n"
        "For sales / pricing calculations of WS/RT items.\n"
        "• Parses text to Excel rows\n"
        "• Calculates Net Buy-in, Sell Out, margins with custom rounding\n"
        "• Groups products by sheet (Oil, Detergent, Milk, etc.)\n\n"
        "For support, contact: your-email@example.com"
    )

def _rebuild_sheet_rows() -> dict[str, list[dict]]:
    """Rebuild SHEET_ROWS from ALL_PRODUCTS (used after delete or clear)."""
    sheet_rows: dict[str, list[dict]] = {}
    for parsed in ALL_PRODUCTS:
        calc = calculate_fields(parsed)
        sheet_name = choose_sheet_name(calc)
        sheet_rows.setdefault(sheet_name, [])
        row_dict = _row_from_data(calc)
        sheet_rows[sheet_name].append(row_dict)
    return sheet_rows

def _build_index_by_sheet() -> dict[str, list[tuple[int, int]]]:
    """
    Build mapping: {norm_sheet_name: [(sheet_id, global_index_in_ALL_PRODUCTS), ...]}
    sheet_id is 1..N inside each sheet (after Date sort).
    """
    items = []
    for idx, parsed in enumerate(ALL_PRODUCTS):
        calc = calculate_fields(parsed)
        sheet_name = choose_sheet_name(calc)
        items.append((idx, sheet_name, calc))

    by_sheet: defaultdict[str, list[tuple[int, dict]]] = defaultdict(list)
    for global_idx, sheet_name, calc in items:
        by_sheet[normalize_sheet(sheet_name)].append((global_idx, calc))

    # sort each sheet by Date
    for sheet_name in by_sheet:
        by_sheet[sheet_name].sort(key=lambda t: t[1].get("date") or "")

    index_map: dict[str, list[tuple[int, int]]] = {}
    for sheet_name, rows in by_sheet.items():
        index_map[sheet_name] = [
            (i + 1, global_idx) for i, (global_idx, _) in enumerate(rows)
        ]
    logger.info("Sheets in index_map: %s", list(index_map.keys()))
    return index_map

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return

    text = update.message.text
    logger.info("User %s sent: %s", update.effective_user.id, text)

    if text.lower().strip() in {"hi", "hello", "hey", "/hi", "/hello", "/hey"}:
        await update.message.reply_text(
            "Hi! Send product data in the template format shown in /start."
        )
        return

    parts = text.split("---")
    blocks = [b for b in parts if "Date:" in b]

    if not blocks:
        return

    try:
        global SHEET_ROWS, ALL_PRODUCTS

        for block in blocks:
            parsed = parse_message(block)
            ALL_PRODUCTS.append(parsed)

        SHEET_ROWS = _rebuild_sheet_rows()
        new_count = len(blocks)

        excel_bytes = build_excel_from_sheet_dict(SHEET_ROWS)
        total_rows = sum(len(v) for v in SHEET_ROWS.values())

        await update.message.reply_document(
            document=InputFile(excel_bytes, filename="calculation_result.xlsx"),
            caption=(
                f"Saved {new_count} new product(s). "
                f"Excel now has {total_rows} product(s). "
                f"Use /list to see Ids, /delete <Sheet> <Id> to delete one "
                f"(example: /delete Milk 2), /delete_sheet <Sheet> to "
                f"delete all in a sheet, or /clear to clear all session data."
            ),
        )
    except Exception as e:
        logger.exception("Error processing message")
        await update.message.reply_text(f"Error: {e}")

async def list_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show list with global Ids and sheet names."""
    if not ALL_PRODUCTS:
        await update.message.reply_text("No products saved yet.")
        return

    lines = []
    for idx, parsed in enumerate(ALL_PRODUCTS, start=1):
        calc = calculate_fields(parsed)
        sheet_name = choose_sheet_name(calc)
        date = parsed.get("date", "?")
        cat = parsed.get("category", "?")
        brand = parsed.get("brand", "?")
        lines.append(f"{idx}. [{sheet_name}] {date} | {cat} | {brand}")

    await update.message.reply_text(
        "Current products (Global Id – [Sheet] Date | Category | Brand):\n\n"
        + "\n".join(lines)
    )

async def delete_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /delete <SheetName> <Id>
    Id = row number inside that sheet after sorting by Date.
    Example: /delete Data 1 , /delete Milk 2
    """
    if len(context.args) < 2:
        await update.message.reply_text(
            "Usage: /delete <SheetName> <Id>\nExample: /delete Milk 2"
        )
        return

    global ALL_PRODUCTS, SHEET_ROWS

    sheet_name_input = context.args[0]
    try:
        sheet_id = int(context.args[1])
    except ValueError:
        await update.message.reply_text(
            "Id must be a number. Example: /delete Milk 2"
        )
        return

    sheet_key = normalize_sheet(sheet_name_input)

    index_map = _build_index_by_sheet()
    if sheet_key not in index_map:
        await update.message.reply_text(
            f"Sheet '{sheet_name_input}' not found. "
            "Check the sheet name in Excel (Oil, Milk, Data, Toilet, etc.)."
        )
        return

    entries = index_map[sheet_key]  # list[(sheet_id, global_idx)]
    match_global = next((g for sid, g in entries if sid == sheet_id), None)
    if match_global is None:
        await update.message.reply_text(
            f"Id {sheet_id} not found in sheet '{sheet_name_input}'."
        )
        return

    removed = ALL_PRODUCTS.pop(match_global)

    SHEET_ROWS = _rebuild_sheet_rows()
    excel_bytes = build_excel_from_sheet_dict(SHEET_ROWS)
    total_rows = sum(len(v) for v in SHEET_ROWS.values())

    await update.message.reply_document(
        document=InputFile(excel_bytes, filename="calculation_result.xlsx"),
        caption=(
            f"Deleted from sheet '{sheet_name_input}' Id {sheet_id} "
            f"({removed.get('date','?')} | {removed.get('category','?')} | "
            f"{removed.get('brand','?')}). "
            f"Excel now has {total_rows} product(s)."
        ),
    )

async def delete_sheet_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /delete_sheet <SheetName>
    Deletes ALL products that belong to that sheet.
    Example: /delete_sheet Milk
    """
    if not context.args:
        await update.message.reply_text(
            "Usage: /delete_sheet <SheetName>\nExample: /delete_sheet Milk"
        )
        return

    global ALL_PRODUCTS, SHEET_ROWS

    sheet_name_input = context.args[0]
    sheet_key = normalize_sheet(sheet_name_input)

    remaining = []
    removed = []
    for parsed in ALL_PRODUCTS:
        calc = calculate_fields(parsed)
        sheet_name = choose_sheet_name(calc)
        if normalize_sheet(sheet_name) == sheet_key:
            removed.append(parsed)
        else:
            remaining.append(parsed)

    if not removed:
        await update.message.reply_text(
            f"No products found in sheet '{sheet_name_input}'."
        )
        return

    ALL_PRODUCTS = remaining
    SHEET_ROWS = _rebuild_sheet_rows()
    excel_bytes = build_excel_from_sheet_dict(SHEET_ROWS)
    total_rows = sum(len(v) for v in SHEET_ROWS.values())

    await update.message.reply_document(
        document=InputFile(excel_bytes, filename="calculation_result.xlsx"),
        caption=(
            f"Deleted {len(removed)} product(s) from sheet '{sheet_name_input}'. "
            f"Excel now has {total_rows} product(s)."
        ),
    )

def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN is not set")

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("settings", settings_command))
    app.add_handler(CommandHandler("clear", clear_command))
    app.add_handler(CommandHandler("about", about_command))

    app.add_handler(CommandHandler("list", list_products))
    app.add_handler(CommandHandler("delete", delete_command))
    app.add_handler(CommandHandler("delete_sheet", delete_sheet_command))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    app.run_polling()

if __name__ == "__main__":
    main()
