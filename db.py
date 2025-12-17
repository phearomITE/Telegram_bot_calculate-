# db.py
import pymysql
from config import (
    MYSQL_HOST,
    MYSQL_PORT,
    MYSQL_DB,
    MYSQL_USER,
    MYSQL_PASSWORD,
    PG_TABLE_BASE,
)

# Common columns shared by all sheet tables
COMMON_COLUMNS = [
    "date DATE",
    "address TEXT",
    "outlet_type TEXT",
    "category TEXT",
    "sub_category TEXT",
    "brand TEXT",
    "packaging TEXT",
    "size_ml DOUBLE",
    "packs INT",
    "weight_ctn_l DOUBLE",
    "buy_in DOUBLE",
    "scheme_base DOUBLE",
    "foc DOUBLE",
    "discount_pct DOUBLE",
    "discount_value DOUBLE",
    "direct_disc_pct DOUBLE",
    "direct_disc_value DOUBLE",
    "net_buy_in DOUBLE",
    "price_100ml DOUBLE",
    "mark_up DOUBLE",
    "sell_out_usd DOUBLE",
    "exchange_rate DOUBLE",
    "sell_out_khr DOUBLE",
    "price_unit_khr DOUBLE",
    "margin_unit_khr DOUBLE",
    "price_ctn_khr DOUBLE",
    "margin_ctn_khr DOUBLE",
]

# All sheet names used in Excel and DB
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
    """Generate table name from base + normalized sheet name."""
    suffix = sheet_name.lower().replace(" ", "_")
    return f"{PG_TABLE_BASE}_{suffix}"

def get_conn():
    """Open a MySQL connection."""
    return pymysql.connect(
        host=MYSQL_HOST,
        port=MYSQL_PORT,
        user=MYSQL_USER,
        password=MYSQL_PASSWORD,
        database=MYSQL_DB,
        cursorclass=pymysql.cursors.Cursor,
    )

def ensure_table(sheet_name: str):
    """Create table for a sheet if it does not exist."""
    tbl = table_name_for_sheet(sheet_name)
    cols_sql = ",\n    ".join(COMMON_COLUMNS)
    ddl = f"""
    CREATE TABLE IF NOT EXISTS {tbl} (
        id INT AUTO_INCREMENT PRIMARY KEY,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        {cols_sql}
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(ddl)
        conn.commit()
    finally:
        conn.close()

def init_db():
    """Ensure all sheet tables exist."""
    for s in SHEET_NAMES:
        ensure_table(s)

def insert_row(sheet_name: str, data: dict):
    """Insert a single row of data into the sheet's table."""
    ensure_table(sheet_name)
    tbl = table_name_for_sheet(sheet_name)

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

    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, values)
        conn.commit()
    finally:
        conn.close()

def fetch_all_rows_by_sheet() -> dict:
    """
    Return { sheet_name: [row_dict, ...] } for all known sheets.
    Used when building the Excel workbook.
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
    try:
        for sheet in SHEET_NAMES:
            tbl = table_name_for_sheet(sheet)
            ensure_table(sheet)
            with conn.cursor() as cur:
                cur.execute(f"SELECT {','.join(cols)} FROM {tbl}")
                rows = cur.fetchall()
            sheet_rows = [dict(zip(cols, r)) for r in rows]
            result[sheet] = sheet_rows
    finally:
        conn.close()

    return result
