import sqlite3
import os
import pandas as pd

DB_NAME = "construction.db"

def get_connection():
    """建立資料庫連線，並強制啟用 SQLite 的外鍵限制（預設是關閉的）"""
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def init_db():
    """初始化 SQL 資料表：合約主表、點工出勤表、材料損耗表"""
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. 專案合約與工項主表 (作為外鍵錨點)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS contract_items (
        item_code TEXT PRIMARY KEY,         -- 項次 (如: 1, 1.1, 2)
        item_name TEXT NOT NULL,            -- 工項名稱
        unit TEXT NOT NULL,                 -- 單位
        contract_price REAL NOT NULL        -- 合約單價
    )
    """)
    
    # 2. 人員進場與點工核銷紀錄表 (包含歷史單價快照)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS labor_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,                 -- 日期 (YYYY-MM-DD)
        vendor_name TEXT NOT NULL,          -- 分包商名稱
        worker_name TEXT NOT NULL,          -- 工人姓名
        work_type TEXT NOT NULL,            -- 工種 (如: 點工-鋼筋、點工-模板)
        reported_hours REAL NOT NULL,       -- 分包商申報時數
        system_hours REAL NOT NULL,         -- 門禁系統實測時數
        approved_hours REAL NOT NULL,       -- 最終核定工時 (取兩者較小值)
        saved_unit_price REAL NOT NULL,     -- 寫入時的單價快照 (防範合約變更震盪)
        total_amount INTEGER NOT NULL,      -- 核定點工總額 (經高精度進位)
        status_flag TEXT NOT NULL           -- 審查狀態 (正常 / 超額攔截)
    )
    """)
    
    # 3. 每日材料進料與損耗紀錄表 (與合約主表連動)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS material_loss_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,                 -- 日期 (YYYY-MM-DD)
        item_code TEXT NOT NULL,            -- 關聯合約項次
        location TEXT NOT NULL,             -- 施作位置 (如: 4F結構牆)
        design_qty REAL NOT NULL,           -- 設計用量 (理論值)
        actual_qty REAL NOT NULL,           -- 實際進料數量
        loss_qty REAL NOT NULL,             -- 損耗數量 (實際 - 設計)
        loss_rate REAL NOT NULL,            -- 損耗率 (%)
        warning_status TEXT NOT NULL,       -- 警示狀態 (正常 / 異常超損)
        FOREIGN KEY (item_code) REFERENCES contract_items(item_code) ON UPDATE CASCADE ON DELETE RESTRICT
    )
    """)
    
    conn.commit()
    conn.close()
    print("🎨 [SQL] 資料庫與三張核心資料表初始化成功！")

def run_self_test():
    """資料庫寫入與讀取自我測試連結 (Test Connection)"""
    print("🧪 [SQL] 開始進行資料庫連線與寫入測試...")
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        # 測試寫入主表
        cursor.execute("INSERT OR REPLACE INTO contract_items VALUES ('1.1', '結構體混凝土澆置', '立方米', 3200)")
        # 測試寫入損耗表
        cursor.execute("INSERT INTO material_loss_records (date, item_code, location, design_qty, actual_qty, loss_qty, loss_rate, warning_status) VALUES ('2026-11-01', '1.1', '1F柱', 100, 106, 6, 6.0, '🚨 異常超損')")
        conn.commit()
        
        # 測試跨表查詢 (JOIN)
        df = pd.read_sql_query("""
            SELECT m.date, c.item_name, m.location, c.contract_price, m.loss_rate, m.warning_status
            FROM material_loss_records m
            JOIN contract_items c ON m.item_code = c.item_code
        """, conn)
        print("\n📊 [SQL 測試拉取聯結結果]:")
        print(df.to_string(index=False))
        print("\n✅ [SQL] 連線、寫入、跨表關聯測試全數通過！")
        
    except Exception as e:
        print(f"❌ [SQL] 測試失敗，錯誤訊息: {str(e)}")
    finally:
        conn.close()

if __name__ == "__main__":
    run_self_test()