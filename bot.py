import logging
from collections import defaultdict


from telegram import (
    Update,
    InputFile,
    ReplyKeyboardMarkup,
    KeyboardButton,
)
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters,
)


from config import BOT_TOKEN, EXCHANGE_RATE_DEFAULT
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




def main_menu_keyboard() -> ReplyKeyboardMarkup:
    keyboard = [
        [KeyboardButton("🆕 New calculation (/start)")],
        [KeyboardButton("📄 Show products (/list)")],
        [
            KeyboardButton("ℹ️ Help (/help)"),
            KeyboardButton("⚙️ Settings (/settings)"),
        ],
        [
            KeyboardButton("📦 About (/about)"),
            KeyboardButton("📊 Summary (/summary)"),
        ],
        [KeyboardButton("🔄 Restart Bot (/restart)")],
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)




async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Welcome to Price Calculator Bot.\n\n"
        "Send one or many products.\n"
        "Separate products with '--- product N ---' lines.\n\n"
        "Example:\n\n" + EXAMPLE_TEXT + "\n\n"
        "Use /help to see all features.",
        reply_markup=main_menu_keyboard(),
    )




async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "📘 Help – Price Calculator Bot\n\n"
        "Commands:\n"
        "/start – Show example format and how to start.\n"
        "/help – Show this help message.\n"
        "/restart – Delete ALL products and start fresh.\n"
        "/settings – Change language, rate, etc.\n"
        "/about – Show bot information.\n"
        "/list – Show all products with Ids.\n"
        "/delete <Sheet> <Id> – Delete one row (Ex: /delete Milk 1).\n"
        "/delete_sheet <Sheet> – Delete all in a sheet.\n"
        "/summary – Show counts per sheet.\n\n"
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
    )
    await update.message.reply_text(text, reply_markup=main_menu_keyboard())




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
        "/settings outlet=RT rate=4100 lang=en",
        reply_markup=main_menu_keyboard(),
    )




async def about_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📦 Price Calculator Bot v2.0\n"
        "For sales / pricing calculations of WS/RT items.\n"
        "• Parses text to Excel rows\n"
        "• Calculates Net Buy-in, Sell Out, margins with custom rounding\n"
        "• Groups products by sheet (Oil, Detergent, Milk, etc.)\n\n"
        "For support, contact: raphearom077@gmail.com or https://t.me/Phearom252005",
        reply_markup=main_menu_keyboard(),
    )




def _rebuild_sheet_rows() -> dict[str, list[dict]]:
    """Rebuild SHEET_ROWS from ALL_PRODUCTS (used after delete)."""
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




async def summary_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show count of products per sheet."""
    if not ALL_PRODUCTS:
        await update.message.reply_text(
            "No products saved yet.\nSend some products first.",
            reply_markup=main_menu_keyboard(),
        )
        return



    counts: dict[str, int] = {}
    for parsed in ALL_PRODUCTS:
        calc = calculate_fields(parsed)
        sheet = choose_sheet_name(calc)
        counts[sheet] = counts.get(sheet, 0) + 1



    total = sum(counts.values())
    lines = [f"• {sheet}: {count} product(s)" for sheet, count in counts.items()]



    await update.message.reply_text(
        "📊 Summary – products per sheet:\n\n"
        + "\n".join(lines)
        + f"\n\nTotal: {total} product(s).",
        reply_markup=main_menu_keyboard(),
    )




async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return



    text = update.message.text
    logger.info("User %s sent: %s", update.effective_user.id, text)



    lower = text.lower().strip()
    # Button shortcuts
    if "new calculation" in lower:
        await start(update, context)
        return
    if "restart bot" in lower:
        await restart_command(update, context)
        return
    if lower in {"hi", "hello", "hey", "/hi", "/hello", "/hey"}:
        await update.message.reply_text(
            "Hi! Send product data in the template format shown in /start.",
            reply_markup=main_menu_keyboard(),
        )
        return
    if "show products" in lower:
        await list_products(update, context)
        return
    if "help" in lower and not lower.startswith("/"):
        await help_command(update, context)
        return
    if "settings" in lower and not lower.startswith("/"):
        await settings_command(update, context)
        return
    if "about" in lower and not lower.startswith("/"):
        await about_command(update, context)
        return
    if "summary" in lower and not lower.startswith("/"):
        await summary_command(update, context)
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
                f"delete all in a sheet."
            ),
            reply_markup=main_menu_keyboard(),
        )
    except Exception as e:
        logger.exception("Error processing message")
        await update.message.reply_text(f"Error: {e}")




async def list_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not ALL_PRODUCTS:
        await update.message.reply_text(
            "No products saved yet.",
            reply_markup=main_menu_keyboard(),
        )
        return



    grouped: dict[str, list[dict]] = {}
    for parsed in ALL_PRODUCTS:
        calc = calculate_fields(parsed)
        sheet = choose_sheet_name(calc)
        grouped.setdefault(sheet, []).append(parsed)



    for sheet in grouped:
        grouped[sheet].sort(key=lambda p: p.get("date") or "")



    lines: list[str] = [
        "Current products\n[Sheet]\n(Id – Date | Category | Brand):\n"
    ]
    for sheet, rows in grouped.items():
        lines.append(f"[{sheet}]")
        for i, parsed in enumerate(rows, start=1):
            date = parsed.get("date", "?")
            cat = parsed.get("category", "?")
            brand = parsed.get("brand", "?")
            lines.append(f"{i}. {date} | {cat} | {brand}")
        lines.append("")



    await update.message.reply_text(
        "\n".join(lines),
        reply_markup=main_menu_keyboard(),
    )



async def restart_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    NEW FEATURE: Clears all stored products and resets the Excel state.
    """
    global ALL_PRODUCTS, SHEET_ROWS
    count = len(ALL_PRODUCTS)
    ALL_PRODUCTS = []
    SHEET_ROWS = {}
    
    await update.message.reply_text(
        f"🔄 Bot Restarted!\nAll {count} products have been cleared.\nYou can start a new calculation now.",
        reply_markup=main_menu_keyboard()
    )




async def delete_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    UPDATED: Handles multi-word sheet names and case-insensitivity.
    /delete <SheetName> <Id>
    """
    if len(context.args) < 2:
        await update.message.reply_text(
            "Usage: /delete <SheetName> <Id>\nExample: /delete Powder detergent 1",
            reply_markup=main_menu_keyboard(),
        )
        return


    global ALL_PRODUCTS, SHEET_ROWS


    sheet_name_input = " ".join(context.args[:-1]).strip()
    try:
        sheet_row_id = int(context.args[-1])
    except ValueError:
        await update.message.reply_text(
            "Id must be a number at the end. Example: /delete Powder detergent 1",
            reply_markup=main_menu_keyboard(),
        )
        return


    sheet_key = normalize_sheet(sheet_name_input)


    items = []
    for idx, parsed in enumerate(ALL_PRODUCTS):
        calc = calculate_fields(parsed)
        sheet_name = choose_sheet_name(calc)
        items.append((idx, sheet_name, calc))


    by_sheet: defaultdict[str, list[tuple[int, dict]]] = defaultdict(list)
    for global_idx, sheet_name, calc in items:
        by_sheet[normalize_sheet(sheet_name)].append((global_idx, calc))


    for s in by_sheet:
        by_sheet[s].sort(key=lambda t: t[1].get("date") or "")


    if sheet_key not in by_sheet:
        await update.message.reply_text(
            f"Sheet '{sheet_name_input}' not found. Check /list.",
            reply_markup=main_menu_keyboard(),
        )
        return


    rows = by_sheet[sheet_key]
    if sheet_row_id < 1 or sheet_row_id > len(rows):
        await update.message.reply_text(
            f"Id {sheet_row_id} not found in sheet '{sheet_name_input}'.",
            reply_markup=main_menu_keyboard(),
        )
        return


    global_idx, _ = rows[sheet_row_id - 1]
    removed = ALL_PRODUCTS.pop(global_idx)


    SHEET_ROWS = _rebuild_sheet_rows()
    excel_bytes = build_excel_from_sheet_dict(SHEET_ROWS)
    total_rows = sum(len(v) for v in SHEET_ROWS.values())


    await update.message.reply_document(
        document=InputFile(excel_bytes, filename="calculation_result.xlsx"),
        caption=(
            f"Deleted from sheet '{sheet_name_input}' Id {sheet_row_id}.\n"
            f"Excel now has {total_rows} product(s)."
        ),
        reply_markup=main_menu_keyboard(),
    )



async def delete_sheet_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    UPDATED: Handles multi-word sheet names and case-insensitivity.
    /delete_sheet <SheetName>
    """
    if not context.args:
        await update.message.reply_text(
            "Usage: /delete_sheet <SheetName>\nExample: /delete_sheet Powder detergent",
            reply_markup=main_menu_keyboard(),
        )
        return


    global ALL_PRODUCTS, SHEET_ROWS


    sheet_name_input = " ".join(context.args).strip()
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
            f"No products found in sheet '{sheet_name_input}'.",
            reply_markup=main_menu_keyboard(),
        )
        return


    ALL_PRODUCTS = remaining
    SHEET_ROWS = _rebuild_sheet_rows()
    excel_bytes = build_excel_from_sheet_dict(SHEET_ROWS)
    total_rows = sum(len(v) for v in SHEET_ROWS.values())


    await update.message.reply_document(
        document=InputFile(excel_bytes, filename="calculation_result.xlsx"),
        caption=(
            f"Deleted {len(removed)} product(s) from sheet '{sheet_name_input}'.\n"
            f"Excel now has {total_rows} product(s)."
        ),
        reply_markup=main_menu_keyboard(),
    )



def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN is not set")


    app = ApplicationBuilder().token(BOT_TOKEN).build()


    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("restart", restart_command))
    app.add_handler(CommandHandler("settings", settings_command))
    app.add_handler(CommandHandler("about", about_command))
    app.add_handler(CommandHandler("summary", summary_command))


    app.add_handler(CommandHandler("list", list_products))
    app.add_handler(CommandHandler("delete", delete_command))
    app.add_handler(CommandHandler("delete_sheet", delete_sheet_command))


    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))


    app.run_polling()



if __name__ == "__main__":
    main()
