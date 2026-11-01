# -*- coding: utf-8 -*-
"""
工程點工與材料損耗資料庫（SQLite）。

執行：python construction_db.py
會建立旁邊的 construction.db，並用暫存資料庫跑完自我測試。
"""
import math
import re
import sqlite3
import sys
import tempfile
from contextlib import contextmanager
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation, Overflow
from pathlib import Path

try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:  # Jupyter / 互動環境沒有 __file__
    BASE_DIR = Path.cwd()

DB_NAME = BASE_DIR / "construction.db"
SQLITE_INT_MAX = 2**63 - 1

# 日期本體必填；後面只接受時間，避免 "2026-02-28abc" 被當成合法日期。
_DATE_RE = re.compile(
    r"\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})"
    r"(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?(?:\.\d+)?"
    r"(?:Z|[+-]\d{2}:?\d{2})?)?\s*"
)

MATERIAL_FIELDS = (
    "id", "date", "item_code", "location", "design_qty", "actual_qty",
    "loss_qty", "loss_rate", "warning_status",
)


# =====================================================
# 1. 連線。交易要自己 BEGIN，否則 DDL 會先被提交，例外也還原不了。
# =====================================================
def get_connection(db_path=DB_NAME):
    """開啟連線並啟用外鍵。呼叫端要自己 commit，或改用 db_session。"""
    conn = sqlite3.connect(db_path, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def db_session(db_path=DB_NAME):
    """成功才提交；發生例外就還原。一開始就取得寫入鎖，連線最後一定關閉。"""
    conn = get_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        yield conn
        conn.commit()
    except Exception:
        try:
            if conn.in_transaction:
                conn.rollback()
        except sqlite3.Error:
            pass
        raise
    finally:
        conn.close()


def _execute_script(conn, script):
    """逐句執行。不用 executescript()，那個方法會先 COMMIT，把外層交易切掉。"""
    for statement in script.split(";"):
        statement = statement.strip()
        if statement:
            conn.execute(statement)


# =====================================================
# 2. 資料表。loss_rate 可為 NULL：設計用量 0 且實際用量 > 0 時算不出損耗率。
# =====================================================
def _material_columns(loss_rate_not_null=False):
    loss_rate = "loss_rate REAL NOT NULL" if loss_rate_not_null else "loss_rate REAL"
    return f"""
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,
    item_code TEXT NOT NULL,
    location TEXT NOT NULL,
    design_qty REAL NOT NULL,
    actual_qty REAL NOT NULL,
    loss_qty REAL NOT NULL,
    {loss_rate},
    warning_status TEXT NOT NULL,
    FOREIGN KEY (item_code) REFERENCES contract_items(item_code)
        ON UPDATE CASCADE ON DELETE RESTRICT
    """


def _schema_sql(loss_rate_not_null=False):
    return f"""
CREATE TABLE IF NOT EXISTS contract_items (
    item_code TEXT PRIMARY KEY,
    item_name TEXT NOT NULL,
    unit TEXT NOT NULL,
    contract_price REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS labor_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,
    vendor_name TEXT NOT NULL,
    worker_name TEXT NOT NULL,
    work_type TEXT NOT NULL,
    reported_hours REAL NOT NULL,
    system_hours REAL NOT NULL,
    approved_hours REAL NOT NULL,
    saved_unit_price REAL NOT NULL,
    total_amount INTEGER NOT NULL,
    status_flag TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS material_loss_records ({_material_columns(loss_rate_not_null)});
"""


INDEX_SQL = """
CREATE INDEX IF NOT EXISTS idx_labor_date ON labor_records(date);
CREATE INDEX IF NOT EXISTS idx_labor_vendor ON labor_records(vendor_name);
CREATE INDEX IF NOT EXISTS idx_material_item ON material_loss_records(item_code);
CREATE INDEX IF NOT EXISTS idx_material_date ON material_loss_records(date);
"""


def _migrate_loss_rate_nullable(conn):
    """舊版 loss_rate 是 NOT NULL，設計用量為 0 的資料會寫不進去。"""
    columns = {
        row["name"]: row["notnull"]
        for row in conn.execute("PRAGMA table_info(material_loss_records)")
    }
    if "loss_rate" not in columns or not columns["loss_rate"]:
        return
    if set(columns) != set(MATERIAL_FIELDS):
        raise ValueError("material_loss_records 欄位與程式不符，已停止自動升級")

    fields = ", ".join(MATERIAL_FIELDS)
    conn.execute("DROP INDEX IF EXISTS idx_material_item")
    conn.execute("DROP INDEX IF EXISTS idx_material_date")
    conn.execute("ALTER TABLE material_loss_records RENAME TO material_loss_old")
    conn.execute(f"CREATE TABLE material_loss_records ({_material_columns()})")
    conn.execute(
        f"INSERT INTO material_loss_records ({fields}) "
        f"SELECT {fields} FROM material_loss_old"
    )
    conn.execute("DROP TABLE material_loss_old")


def init_db(db_path=DB_NAME):
    """建立資料表（已存在則略過），並自動升級舊版材料表。"""
    with db_session(db_path) as conn:
        _execute_script(conn, _schema_sql())
        _migrate_loss_rate_nullable(conn)
        _execute_script(conn, INDEX_SQL)


# =====================================================
# 3. 輸入驗證
# =====================================================
def to_decimal(value, name):
    """轉成有限、非負的 Decimal。空白、文字、NaN、無限大、負數都是 ValueError。"""
    try:
        number = Decimal(str(value).strip().replace(",", ""))
    except (InvalidOperation, Overflow, ValueError, TypeError):
        raise ValueError(f"{name}必須是數字：{value!r}") from None
    if not number.is_finite():
        raise ValueError(f"{name}不可為 NaN 或無限大")
    if number < 0:
        raise ValueError(f"{name}不可為負數")
    return number


def _as_real(value, name):
    """SQLite 的 REAL 會把 NaN 存成 NULL、把溢位存成無限大，所以這裡先擋下來。"""
    try:
        number = float(value)
    except (OverflowError, ValueError):
        raise ValueError(f"{name}超出可儲存範圍") from None
    if not math.isfinite(number):
        raise ValueError(f"{name}超出可儲存範圍")
    return number


def _as_money(value, name="金額"):
    """四捨五入到整數元，且要放得進 SQLite 的 64 位元整數。"""
    try:
        if value >= SQLITE_INT_MAX + 1:
            raise ValueError(f"{name}超出可儲存範圍")
        amount = int(value.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, Overflow):
        raise ValueError(f"{name}超出可計算範圍") from None
    if amount > SQLITE_INT_MAX:
        raise ValueError(f"{name}超出可儲存範圍")
    return amount


def require_text(value, name):
    text = "" if value is None else str(value).strip()
    if not text:
        raise ValueError(f"{name}不可空白")
    return text


def normalize_date(value):
    """統一成 YYYY-MM-DD，日期才能排序，也不會存成 '2026-10-01 00:00:00'。"""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()

    matched = _DATE_RE.fullmatch(str(value))
    if not matched:
        raise ValueError(f"日期格式錯誤：{value!r}（請用 YYYY-MM-DD）")
    year, month, day, hour, minute, second = matched.groups()
    if hour is not None:
        second = 0 if second is None else int(second)
        if not (0 <= int(hour) <= 23 and 0 <= int(minute) <= 59 and 0 <= second <= 59):
            raise ValueError(f"時間格式錯誤：{value!r}")
    try:
        return date(int(year), int(month), int(day)).isoformat()
    except ValueError:
        raise ValueError(f"日期不存在：{value!r}") from None


# =====================================================
# 4. 計算與寫入
# =====================================================
def calculate_labor_amount(reported_hours, system_hours, unit_price):
    """核定工時 = min(申報, 門禁)；金額 = 核定工時 × 單價，四捨五入到整數元。"""
    reported = to_decimal(reported_hours, "申報工時")
    system = to_decimal(system_hours, "門禁工時")
    price = to_decimal(unit_price, "單價")
    approved = min(reported, system)
    try:
        product = approved * price
    except (InvalidOperation, Overflow):
        raise ValueError("金額超出可計算範圍") from None
    status = "正常" if reported <= system else "超額攔截"
    return _as_real(approved, "核定工時"), _as_money(product), status


def calculate_material_loss(design_qty, actual_qty, limit_rate=5.0):
    """損耗數量 = 實際 − 設計；損耗率 = 損耗數量 ÷ 設計 × 100。

    設計用量為 0 時損耗率是 None；設計與實際都是 0 時損耗率記 0。
    """
    design = to_decimal(design_qty, "設計用量")
    actual = to_decimal(actual_qty, "實際用量")
    limit = to_decimal(limit_rate, "損耗上限率")
    loss_qty = actual - design
    if design == 0:
        loss_rate = 0.0 if actual == 0 else None
        warning = "正常" if actual == 0 else "無設計用量，需人工確認"
        return _as_real(loss_qty, "損耗數量"), loss_rate, warning

    raw_rate = (loss_qty / design) * 100
    loss_rate = _as_real(raw_rate, "損耗率")
    warning = "異常超損" if Decimal(str(loss_rate)) > limit else "正常"
    return _as_real(loss_qty, "損耗數量"), loss_rate, warning


def insert_contract_item(conn, item_code, item_name, unit, contract_price):
    conn.execute(
        "INSERT OR REPLACE INTO contract_items (item_code, item_name, unit, contract_price) VALUES (?, ?, ?, ?)",
        (require_text(item_code, "工項編號"), require_text(item_name, "工項名稱"),
         require_text(unit, "單位"), float(to_decimal(contract_price, "合約單價"))),
    )


def insert_labor_record(conn, work_date, vendor_name, worker_name, work_type,
                        reported_hours, system_hours, unit_price):
    approved, amount, status = calculate_labor_amount(reported_hours, system_hours, unit_price)
    cur = conn.execute(
        """INSERT INTO labor_records (date, vendor_name, worker_name, work_type, reported_hours,
               system_hours, approved_hours, saved_unit_price, total_amount, status_flag)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (normalize_date(work_date), require_text(vendor_name, "廠商名稱"),
         require_text(worker_name, "工人姓名"), require_text(work_type, "工種"),
         float(to_decimal(reported_hours, "申報工時")), float(to_decimal(system_hours, "門禁工時")),
         approved, float(to_decimal(unit_price, "單價")), amount, status),
    )
    return cur.lastrowid


def insert_material_record(conn, work_date, item_code, location,
                           design_qty, actual_qty, limit_rate=5.0):
    loss_qty, loss_rate, warning = calculate_material_loss(design_qty, actual_qty, limit_rate)
    cur = conn.execute(
        """INSERT INTO material_loss_records (date, item_code, location, design_qty, actual_qty,
               loss_qty, loss_rate, warning_status)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (normalize_date(work_date), require_text(item_code, "工項編號"), require_text(location, "位置"),
         float(to_decimal(design_qty, "設計用量")), float(to_decimal(actual_qty, "實際用量")),
         loss_qty, loss_rate, warning),
    )
    return cur.lastrowid


if __name__ == "__main__":
    init_db()
    print("✅ 資料庫基本初始化與自我測試邏輯載入完畢！")