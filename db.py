# db.py
import psycopg2
from config import PG_HOST, PG_PORT, PG_DB, PG_USER, PG_PASSWORD, PG_TABLE_BASE

COMMON_COLUMNS = [
    "date DATE",
    "address TEXT",
    "outlet_type TEXT",
    "category TEXT",
    "sub_category TEXT",
    "brand TEXT",
    "packaging TEXT",
    "size_ml NUMERIC",
    "packs INTEGER",
    "weight_ctn_l NUMERIC",
    "buy_in NUMERIC",
    "scheme_base NUMERIC",
    "foc NUMERIC",
    "discount_pct NUMERIC",
    "discount_value NUMERIC",
    "direct_disc_pct NUMERIC",
    "direct_disc_value NUMERIC",
    "net_buy_in NUMERIC",
    "price_100ml NUMERIC",
    "mark_up NUMERIC",
    "sell_out_usd NUMERIC",
    "exchange_rate NUMERIC",
    "sell_out_khr NUMERIC",
    "price_unit_khr NUMERIC",
    "margin_unit_khr NUMERIC",
    "price_ctn_khr NUMERIC",
    "margin_ctn_khr NUMERIC",
]

SHEET_NAMES = [
    "Oil",
    "Powder Detergent",
    "Liquid Detergent",
    "Milk",
    "Dishwash",
    "Fabric Softener",
    "Eco Dishwash",
    "Toilet",
    "Data",
]

def table_name_for_sheet(sheet_name: str) -> str:
    suffix = sheet_name.lower().replace(" ", "_")
    return f"{PG_TABLE_BASE}_{suffix}"

def get_conn():
    return psycopg2.connect(
        host=PG_HOST,
        port=PG_PORT,
        dbname=PG_DB,
        user=PG_USER,
        password=PG_PASSWORD,
    )

def ensure_table(sheet_name: str):
    """Create table for this sheet if not exists."""
    tbl = table_name_for_sheet(sheet_name)
    cols_sql = ",\n    ".join(COMMON_COLUMNS)
    ddl = f"""
    CREATE TABLE IF NOT EXISTS {tbl} (
        id SERIAL PRIMARY KEY,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        {cols_sql}
    );
    """
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(ddl)
    conn.commit()
    cur.close()
    conn.close()

def init_db():
    for s in SHEET_NAMES:
        ensure_table(s)

def insert_row(sheet_name: str, data: dict):
    """Insert one row into the table corresponding to sheet_name."""
    ensure_table(sheet_name)
    tbl = table_name_for_sheet(sheet_name)
    conn = get_conn()
    cur = conn.cursor()

    cols = [
        "date","address","outlet_type","category","sub_category","brand",
        "packaging","size_ml","packs","weight_ctn_l","buy_in","scheme_base",
        "foc","discount_pct","discount_value","direct_disc_pct",
        "direct_disc_value","net_buy_in","price_100ml","mark_up",
        "sell_out_usd","exchange_rate","sell_out_khr","price_unit_khr",
        "margin_unit_khr","price_ctn_khr","margin_ctn_khr"
    ]
    values = [data.get(c) for c in cols]
    placeholders = ",".join(["%s"] * len(cols))
    sql = f"INSERT INTO {tbl} ({','.join(cols)}) VALUES ({placeholders})"
    cur.execute(sql, values)
    conn.commit()
    cur.close()
    conn.close()

def fetch_all_rows_by_sheet() -> dict:
    """
    Return { sheet_name: [row_dict, ...] } for all known sheets.
    """
    result = {}
    cols = [
        "date","address","outlet_type","category","sub_category","brand",
        "packaging","size_ml","packs","weight_ctn_l","buy_in","scheme_base",
        "foc","discount_pct","discount_value","direct_disc_pct",
        "direct_disc_value","net_buy_in","price_100ml","mark_up",
        "sell_out_usd","exchange_rate","sell_out_khr","price_unit_khr",
        "margin_unit_khr","price_ctn_khr","margin_ctn_khr"
    ]

    conn = get_conn()
    cur = conn.cursor()

    for sheet in SHEET_NAMES:
        tbl = table_name_for_sheet(sheet)
        ensure_table(sheet)
        cur.execute(f"SELECT {','.join(cols)} FROM {tbl}")
        rows = cur.fetchall()
        sheet_rows = []
        for r in rows:
            sheet_rows.append(dict(zip(cols, r)))
        result[sheet] = sheet_rows

    cur.close()
    conn.close()
    return result
