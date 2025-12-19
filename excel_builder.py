import io
from decimal import Decimal, ROUND_FLOOR

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from config import EXCHANGE_RATE_DEFAULT


# ---------- rounding helper: 6–9 up, 1–5 down, always 2 decimals ----------
def round2(x):
    if x is None:
        return None

    d = Decimal(str(x))

    scaled = (d * 1000).quantize(Decimal("1"), rounding=ROUND_FLOOR)
    third = int(scaled % 10)

    if third >= 6:
        q = Decimal("0.01")
        d2 = (-d).quantize(q, rounding=ROUND_FLOOR)
        d2 = -d2
    else:
        d2 = d.quantize(Decimal("0.01"), rounding=ROUND_FLOOR)

    return f"{d2:.2f}"


def size_is_gram(data: dict) -> bool:
    raw = (data.get("size_raw") or "").lower()
    return "g" in raw and "ml" not in raw


def size_is_ml(data: dict) -> bool:
    raw = (data.get("size_raw") or "").lower()
    return "ml" in raw


def choose_sheet_name(data: dict) -> str:
    """
    Map category + sub_category to sheet names with many allowed user inputs.
    """
    category = (data.get("category") or "").strip().lower()
    sub_cat = (data.get("sub_category") or "").strip().lower()

    # Oil
    oil_keywords = {
        "cooking oil", "oil", "palm oil", "vegetable oil",
        "coconut oil", "sunflower oil"
    }
    if category in oil_keywords:
        return "Oil"

    # Powder Detergent
    if category in {"detergent", "powder detergent", "washing powder"} and \
       sub_cat in {"powder", "powdered", "dry", ""}:
        return "Powder Detergent"

    # Liquid Detergent
    if category in {"detergent", "liquid detergent", "laundry liquid"} and \
       sub_cat in {"liquid", "fluid", ""}:
        return "Liquid Detergent"

    # Milk
    milk_keywords = {
        "milk", "dairy milk", "fresh milk",
        "evaporated milk", "condensed milk", "uht milk"
    }
    if category in milk_keywords:
        return "Milk"

    # Dishwash (non‑eco)
    dishwash_keywords = {
        "dishwash", "dish wash", "dishwashing liquid",
        "washing liquid", "dishwashing"
    }
    if category in dishwash_keywords and sub_cat not in {"eco", "eco-friendly"}:
        return "Dishwash"

    # Fabric Softener
    fabric_keywords = {
        "fabric softener", "fabric conditioner", "softener",
        "conditioner", "fabric softner"
    }
    if category in fabric_keywords:
        return "Fabric Softener"

    # Eco Dishwash
    eco_keywords = {
        "eco dishwash", "eco dishwashing", "eco-dishwash",
        "biodegradable dishwash", "green dishwash"
    }
    if category in eco_keywords or (
        category in dishwash_keywords and sub_cat in {"eco", "eco-friendly"}
    ):
        return "Eco Dishwash"

    # Toilet
    toilet_keywords = {
        "toilet", "toilet cleaner", "toilet bowl cleaner",
        "wc cleaner", "toilet liquid"
    }
    if category in toilet_keywords:
        return "Toilet"

    # Fallback
    return "Data"


def calculate_fields(data: dict) -> dict:
    buy_in = data.get("buy_in")
    scheme = data.get("scheme_base")
    foc = data.get("foc")
    direct_disc_pct_input = data.get("direct_disc_pct")
    mark_up = data.get("mark_up") or 0
    size_val = data.get("size_ml") or 0
    packs = data.get("packs") or 1
    price_unit_khr_input = data.get("price_unit_khr")

    exchange_rate = EXCHANGE_RATE_DEFAULT

    if buy_in is None:
        raise ValueError("Buy-in is required")
    if price_unit_khr_input is None:
        raise ValueError("Price Unit (KHR) is required")

    # Discount(%) and Discount($) from Scheme & FOC
    if scheme and foc and (scheme + foc) != 0:
        discount_pct_val = (foc / (scheme + foc)) * 100.0
    else:
        discount_pct_val = 0.0

    discount_value_val = (discount_pct_val / 100.0) * buy_in

    # Direct Disc($) from Direct Disc(%)
    direct_disc_pct_val = direct_disc_pct_input or 0.0
    direct_disc_value_val = (direct_disc_pct_val / 100.0) * buy_in

    net_buy_in_val = buy_in - (discount_value_val + direct_disc_value_val)

    # size‑based
    total_size = size_val * packs if size_val and packs else 0
    price_100_unit_val = (
        net_buy_in_val / (total_size / 100.0) if total_size else None
    )

    # Weight per Ctn raw in kg/L, then integer (no decimal)
    weight_ctn_base_raw = (total_size / 1000.0) if total_size else None
    if weight_ctn_base_raw is not None:
        weight_ctn_int = int(round(weight_ctn_base_raw))
        weight_ctn_base = str(weight_ctn_int)
    else:
        weight_ctn_base = None

    sell_out_usd_val = net_buy_in_val + mark_up

    # KHR as integers
    sell_out_khr_val = sell_out_usd_val * exchange_rate
    sell_out_khr_int = int(round(sell_out_khr_val))

    price_unit_khr_int = int(round(price_unit_khr_input))
    margin_unit_khr_val = price_unit_khr_int - (sell_out_khr_int / packs)
    price_ctn_khr_val = price_unit_khr_int * packs
    margin_ctn_khr_val = price_ctn_khr_val - sell_out_khr_int

    result = dict(data)
    result.update(
        size_ml=str(int(round(size_val))) if size_val is not None else None,
        buy_in=round2(buy_in),
        discount_pct=round2(discount_pct_val),
        discount_value=round2(discount_value_val),
        direct_disc_pct=round2(direct_disc_pct_val),
        direct_disc_value=round2(direct_disc_value_val),
        net_buy_in=round2(net_buy_in_val),
        price_100ml=(
            round2(price_100_unit_val)
            if price_100_unit_val is not None
            else None
        ),
        mark_up=round2(mark_up),
        sell_out_usd=round2(sell_out_usd_val),
        exchange_rate=f"KHR {exchange_rate}",
        sell_out_khr=str(sell_out_khr_int),
        price_unit_khr=str(price_unit_khr_int),
        margin_unit_khr=str(int(round(margin_unit_khr_val))),
        price_ctn_khr=str(int(round(price_ctn_khr_val))),
        margin_ctn_khr=str(int(round(margin_ctn_khr_val))),
        weight_ctn_l=weight_ctn_base,
    )
    return result


def _fmt(value, unit: str | None = None):
    if value is None:
        return None
    if unit:
        # for money, keep unit first: KHR 92,000 ; $ 13.80
        return f"{unit} {value}"
    return value


def _row_from_data(data: dict) -> dict:
    if size_is_gram(data):
        size_unit = "g"
        weight_unit = "kg"
        price_header = "Price / 100g"
    elif size_is_ml(data):
        size_unit = "ml"
        weight_unit = "L"
        price_header = "Price / 100ml"
    else:
        size_unit = ""
        weight_unit = ""
        price_header = "Price / 100 unit"

    # ----- value then unit for 4 columns -----
    # Size  -> "230 ml" / "400 g"
    size_val_raw = data.get("size_ml")
    if size_val_raw is not None:
        try:
            size_num = int(float(size_val_raw))
        except (TypeError, ValueError):
            size_num = size_val_raw
        size_display = f"{size_num} {size_unit}".strip()
    else:
        size_display = None

    # Weight per Ctn -> "11 L" / "10 kg"
    weight_val_raw = data.get("weight_ctn_l")
    if weight_val_raw is not None:
        weight_display = f"{weight_val_raw} {weight_unit}".strip()
    else:
        weight_display = None

    # Discount(%) -> "10.00%"
    disc_pct_raw = data.get("discount_pct")
    disc_pct_display = f"{disc_pct_raw}%" if disc_pct_raw is not None else None

    # Direct Disc.(%) -> "5.00%"
    direct_disc_pct_raw = data.get("direct_disc_pct")
    direct_disc_pct_display = (
        f"{direct_disc_pct_raw}%" if direct_disc_pct_raw is not None else None
    )
    # ----------------------------------------

    row = {
        "Date": data.get("date"),
        "Address": data.get("address"),
        "Category": data.get("category"),
        "Sub-Category": data.get("sub_category"),
        "Brand": data.get("brand"),
        "Packaging": data.get("packaging"),

        # 1) Size (value then unit)
        "Size": size_display,
        "Packs": data.get("packs"),

        # 2) Weight per Ctn (value then unit)
        "Weight per Ctn": weight_display,

        "Buy-in": _fmt(data.get("buy_in"), "$"),
        "Scheme(base)": data.get("scheme_base"),
        "FOC": data.get("foc"),

        # 3) Discount(%) (value then %)
        "Discount(%)": disc_pct_display,

        "Discount($)": _fmt(data.get("discount_value"), "$"),

        # 4) Direct Disc.(%) (value then %)
        "Direct Disc.(%)": direct_disc_pct_display,

        "Direct Disc($)": _fmt(data.get("direct_disc_value"), "$"),
        "Net Buy-in": _fmt(data.get("net_buy_in"), "$"),
        "Price / 100 unit": _fmt(data.get("price_100ml"), "$"),
        "PriceHeader": price_header,
        "Mark - up": _fmt(data.get("mark_up"), "$"),
        "Sell Out ($)": _fmt(data.get("sell_out_usd"), "$"),
        "Exchange Rate": data.get("exchange_rate"),
        "Sell Out (KHR)": _fmt(data.get("sell_out_khr"), "KHR"),
        "Price Unit (KHR)": _fmt(data.get("price_unit_khr"), "KHR"),
        "Margin/Unit (KHR)": _fmt(data.get("margin_unit_khr"), "KHR"),
        "Price Ctn (KHR)": _fmt(data.get("price_ctn_khr"), "KHR"),
        "Margin/Ctn (KHR)": _fmt(data.get("margin_ctn_khr"), "KHR"),
    }
    return row


def build_excel_from_sheet_dict(sheet_rows: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, rows in sheet_rows.items():
        if not rows:
            continue

        ws = wb.create_sheet(title=sheet_name)

        section_headers = [
            ("A", "I", "PRODUCT INFO", "FF105437"),
            ("J", "R", "WHOLESALE BUY-IN", "FF0070C0"),
            ("S", "V", "WHOLESALE SELL-OUT", "FF7030A0"),
            ("W", "Z", "RETAIL", "FFED7D31"),
        ]
        for col_start, col_end, label, color in section_headers:
            ws.merge_cells(f"{col_start}1:{col_end}1")
            cell = ws[f"{col_start}1"]
            cell.value = label
            cell.font = Font(bold=True, size=11, color="FFFFFF")
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        headers = [
            "Date", "Address", "Category", "Sub-Category", "Brand", "Packaging",
            "Size", "Packs", "Weight per Ctn",
            "Buy-in", "Scheme(base)", "FOC", "Discount(%)", "Discount($)",
            "Direct Disc.(%)", "Direct Disc($)", "Net Buy-in", "Price / 100 unit",
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

        df = pd.DataFrame(rows)
        if "Date" in df.columns:
            df = df.sort_values(by=["Date"], ascending=True, na_position="last")

        price_col_idx = headers.index("Price / 100 unit") + 1

        for row_idx, (_, row_data) in enumerate(df.iterrows(), start=3):
            price_header = row_data.get("PriceHeader")
            if price_header and ws.cell(row=2, column=price_col_idx).value != price_header:
                ws.cell(row=2, column=price_col_idx).value = price_header

            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                value = row_data.get(header)

                # force 4 columns to be plain strings (value then unit)
                if header in ["Size", "Weight per Ctn", "Discount(%)", "Direct Disc.(%)"]:
                    cell.value = "" if value is None else str(value)
                    cell.number_format = "General"
                else:
                    cell.value = value

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
