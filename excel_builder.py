# excel_builder.py
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from config import EXCHANGE_RATE_DEFAULT

# ------------- Sheet tracking rules -------------
#
# Oil sheet           : Category = "Cooking Oil" or "Oil"
# Powder Detergent    : Category = "Detergent" AND Sub-Category = "Powder"
# Liquid Detergent    : Category = "Detergent" AND Sub-Category = "Liquid"
# Milk                : Category = "Milk"
# Dishwash            : Category = "Dishwash"
# Fabric Softener     : Category = "Fabric Softener"
# Eco Dishwash        : Category = "Eco Dishwash"
# Toilet              : Category = "Toilet"
# Otherwise           : Data
#

def choose_sheet_name(data: dict) -> str:
    """Decide Excel sheet name from Category / Sub-Category."""
    category = (data.get("category") or "").strip().lower()
    sub_cat = (data.get("sub_category") or "").strip().lower()

    if category in ("cooking oil", "oil"):
        return "Oil"

    if category == "detergent" and sub_cat == "powder":
        return "Powder Detergent"

    if category == "detergent" and sub_cat == "liquid":
        return "Liquid Detergent"

    if category == "milk":
        return "Milk"

    if category == "dishwash":
        return "Dishwash"

    if category == "fabric softener":
        return "Fabric Softener"

    if category == "eco dishwash":
        return "Eco Dishwash"

    if category == "toilet":
        return "Toilet"

    return "Data"

# ------------- Calculations -------------

def calculate_fields(data: dict) -> dict:
    """
    WHOLESALE BUY-IN
      Discount(%)      = FOC / (Scheme(Base) + FOC)
      Discount($)      = Discount(%) * Buy-in
      Direct Disc.($)  = Direct Disc.(%) * Buy-in
      Net Buy-in       = Buy-in - (Discount($) + Direct Disc.($))
      Price/100ml      = Net Buy-in / ((Size * Packs) / 100)

    WHOLESALE SELL-OUT
      Sell-out($)      = Net Buy-in + Mark-up
      Sell-out(KHR)    = Sell-out($) * Exchange Rate

    RETAIL
      Margin/Unit      = Price/Unit - (Sell-out(KHR) / Packs)
      Price/Ctn        = Price/Unit * Packs
      Margin/Ctn       = Price/Ctn - Sell-out(KHR)
    """
    buy_in = data.get("buy_in")
    scheme = data.get("scheme_base")
    foc = data.get("foc")
    discount_pct = data.get("discount_pct")
    discount_value = data.get("discount_value")
    direct_disc_pct = data.get("direct_disc_pct")
    direct_disc_value = data.get("direct_disc_value")
    mark_up = data.get("mark_up") or 0
    sell_out_usd = data.get("sell_out_usd")
    size_ml = data.get("size_ml") or 0
    packs = data.get("packs") or 1
    price_unit_khr = data.get("price_unit_khr")
    exchange_rate = data.get("exchange_rate") or EXCHANGE_RATE_DEFAULT

    if buy_in is None:
        raise ValueError("Buy-in is required")
    if price_unit_khr is None:
        raise ValueError("Price Unit (KHR) is required")

    # WHOLESALE BUY-IN
    if scheme and foc and (scheme + foc) != 0:
        discount_pct = (foc / (scheme + foc)) * 100.0
    else:
        discount_pct = discount_pct or 0

    discount_value = (discount_pct / 100.0) * buy_in
    direct_disc_pct = direct_disc_pct or 0
    direct_disc_value = (direct_disc_pct / 100.0) * buy_in
    net_buy_in = buy_in - (discount_value + direct_disc_value)

    total_unit = size_ml * packs if size_ml and packs else 0
    price_100ml = net_buy_in / (total_unit / 100.0) if total_unit else None

    # WHOLESALE SELL-OUT
    if not sell_out_usd:
        sell_out_usd = net_buy_in + mark_up
    sell_out_khr = sell_out_usd * exchange_rate

    # RETAIL
    margin_unit_khr = price_unit_khr - (sell_out_khr / packs)
    price_ctn_khr = price_unit_khr * packs
    margin_ctn_khr = price_ctn_khr - sell_out_khr

    result = dict(data)
    result.update(
        discount_pct=round(discount_pct, 4),
        discount_value=round(discount_value, 4),
        direct_disc_pct=round(direct_disc_pct, 4),
        direct_disc_value=round(direct_disc_value, 4),
        net_buy_in=round(net_buy_in, 4),
        price_100ml=round(price_100ml, 4) if price_100ml is not None else None,
        sell_out_usd=round(sell_out_usd, 4),
        exchange_rate=exchange_rate,
        sell_out_khr=round(sell_out_khr, 2),
        margin_unit_khr=round(margin_unit_khr, 2),
        price_ctn_khr=round(price_ctn_khr, 2),
        margin_ctn_khr=round(margin_ctn_khr, 2),
    )
    return result

# ------------- Row mapping with units in values -------------

def _fmt(value, unit: str | None = None):
    """Format numeric value with unit as text, e.g. 22.5 -> '22.5 $'."""
    if value is None:
        return None
    if unit:
        return f"{value} {unit}"
    return value

def _row_from_data(data: dict) -> dict:
    """Convert internal data dict to Excel row, embedding units in cell values."""
    return {
        "Date": data.get("date"),
        "Address": data.get("address"),
        "Category": data.get("category"),
        "Sub-Category": data.get("sub_category"),
        "Brand": data.get("brand"),
        "Packaging": data.get("packaging"),

        # PRODUCT INFO
        "Size": _fmt(data.get("size_ml"), "ml"),
        "Packs": data.get("packs"),
        "Weight per Ctn": _fmt(data.get("weight_ctn_l"), "L"),

        # WHOLESALE BUY-IN
        "Buy-in": _fmt(data.get("buy_in"), "$"),
        "Scheme(base)": data.get("scheme_base"),
        "FOC": data.get("foc"),
        "Discount(%)": _fmt(data.get("discount_pct"), "%"),
        "Discount($)": _fmt(data.get("discount_value"), "$"),
        "Direct Disc.(%)": _fmt(data.get("direct_disc_pct"), "%"),
        "Direct Disc($)": _fmt(data.get("direct_disc_value"), "$"),
        "Net Buy-in": _fmt(data.get("net_buy_in"), "$"),
        "Price / 100ml": _fmt(data.get("price_100ml"), "$"),

        # WHOLESALE SELL-OUT
        "Mark - up": _fmt(data.get("mark_up"), "$"),
        "Sell Out ($)": _fmt(data.get("sell_out_usd"), "$"),
        "Exchange Rate": data.get("exchange_rate"),
        "Sell Out (KHR)": _fmt(data.get("sell_out_khr"), "KHR"),

        # RETAIL
        "Price Unit (KHR)": _fmt(data.get("price_unit_khr"), "KHR"),
        "Margin/Unit (KHR)": _fmt(data.get("margin_unit_khr"), "KHR"),
        "Price Ctn (KHR)": _fmt(data.get("price_ctn_khr"), "KHR"),
        "Margin/Ctn (KHR)": _fmt(data.get("margin_ctn_khr"), "KHR"),
    }

# ------------- Excel builder -------------

def build_excel_from_sheet_dict(sheet_rows: dict) -> bytes:
    """
    Build Excel with:
      - Row 1: PRODUCT INFO | WHOLESALE BUY-IN | WHOLESALE SELL-OUT | RETAIL
      - Row 2: Column headers (no units, units are inside values)
      - Row 3+: Data rows sorted by Date, values like '22.5 $', '12 L'
    """
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, rows in sheet_rows.items():
        if not rows:
            continue

        ws = wb.create_sheet(title=sheet_name)

        # Row 1: colored section headers
        section_headers = [
            ("A", "I", "PRODUCT INFO", "FF105437"),      # dark green
            ("J", "R", "WHOLESALE BUY-IN", "FF0070C0"),  # blue
            ("S", "V", "WHOLESALE SELL-OUT", "FF7030A0"),# purple
            ("W", "Z", "RETAIL", "FFED7D31"),            # orange
        ]
        for col_start, col_end, label, color in section_headers:
            start_row = 1
            ws.merge_cells(f"{col_start}{start_row}:{col_end}{start_row}")
            cell = ws[f"{col_start}{start_row}"]
            cell.value = label
            cell.font = Font(bold=True, size=11, color="FFFFFF")
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Row 2: column headers (plain text)
        headers = [
            "Date", "Address", "Category", "Sub-Category", "Brand", "Packaging",
            "Size", "Packs", "Weight per Ctn",
            "Buy-in", "Scheme(base)", "FOC", "Discount(%)", "Discount($)",
            "Direct Disc.(%)", "Direct Disc($)", "Net Buy-in", "Price / 100ml",
            "Mark - up", "Sell Out ($)", "Exchange Rate", "Sell Out (KHR)",
            "Price Unit (KHR)", "Margin/Unit (KHR)", "Price Ctn (KHR)", "Margin/Ctn (KHR)",
        ]

        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="FF404040", end_color="FF404040", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

        # Data rows
        df = pd.DataFrame(rows)
        if "Date" in df.columns:
            df = df.sort_values(by=["Date"], ascending=True, na_position="last")

        for row_idx, (_, row_data) in enumerate(df.iterrows(), start=3):
            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = row_data.get(header)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                )

        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 30
        for col_num in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col_num)].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
