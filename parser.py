import re
from dateutil import parser as dateparser


def _extract_value(text, key_pattern):
    pattern = rf"^\s*({key_pattern})\s*:\s*(.*)$"
    for line in text.splitlines():
        m = re.match(pattern, line, re.IGNORECASE)
        if m:
            value = m.group(2)
            if value is None:
                return None
            value = value.strip()
            return value if value != "" else None
    return None


def num_or_none(value):
    """
    Convert strings like '15.90$', '1,000 KHR', ' 0.5 ' to float.
    Returns None if not a valid number.
    """
    if not value:
        return None

    v = value.strip()
    # support comma decimal as well
    v = v.replace(",", ".")
    # remove currency symbols and letters, keep digits . and -
    v = re.sub(r"[^0-9.\-]", "", v)

    if v in {"", "-", ".", "-.", ".-"}:
        return None

    try:
        return float(v)
    except ValueError:
        return None


def parse_message(text: str) -> dict:
    d = {}

    # Date
    d["date_raw"] = _extract_value(text, r"Date")
    if d["date_raw"]:
        try:
            d["date"] = dateparser.parse(
                d["date_raw"],
                dayfirst=True,
            ).date()
        except Exception:
            d["date"] = None
    else:
        d["date"] = None

    # Address / Addresss
    d["address"] = _extract_value(text, r"Addresss|Address")

    d["outlet_type"] = _extract_value(text, r"Outlet-Type")
    d["category"] = _extract_value(text, r"Category")
    d["sub_category"] = _extract_value(text, r"Sub-Category")
    d["brand"] = _extract_value(text, r"Brand")
    d["packaging"] = _extract_value(text, r"Packaging")

    d["size_raw"] = _extract_value(text, r"Size")
    d["packs_raw"] = _extract_value(text, r"Packs")
    d["weight_raw"] = _extract_value(text, r"Weight per Ctn")

    d["size_ml"] = num_or_none(d["size_raw"])
    packs_val = num_or_none(d["packs_raw"])
    d["packs"] = int(packs_val) if packs_val is not None else None
    d["weight_ctn_l"] = num_or_none(d["weight_raw"])

    d["buy_in"] = num_or_none(_extract_value(text, r"Buy-in"))

    d["scheme_base_raw"] = _extract_value(
        text,
        r"Scheme\(base\)|Scheme\(Base\)|Scheme",
    )
    d["scheme_base"] = num_or_none(d["scheme_base_raw"])

    d["foc_raw"] = _extract_value(text, r"FOC")
    d["foc"] = num_or_none(d["foc_raw"])

    d["discount_pct"] = num_or_none(
        _extract_value(text, r"Discount\(%\)"),
    )
    d["discount_value"] = num_or_none(
        _extract_value(text, r"Discount\(\$\)"),
    )

    d["direct_disc_pct"] = num_or_none(
        _extract_value(text, r"Direct Disc\.\(%\)"),
    )
    d["direct_disc_value"] = num_or_none(
        _extract_value(text, r"Direct Disc\(\$\)"),
    )

    d["mark_up"] = num_or_none(
        _extract_value(text, r"Mark\s*-\s*up|Mark\s*up"),
    )

    sell_out_raw = _extract_value(text, r"Sell Out \(\$\)")
    d["sell_out_usd"] = num_or_none(sell_out_raw)

    price_unit_raw = _extract_value(text, r"Price Unit")
    d["price_unit_khr"] = num_or_none(price_unit_raw)

    # Exchange rate always default in calculations
    d["exchange_rate"] = None

    return d
