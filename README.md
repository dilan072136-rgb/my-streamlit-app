# 工程點工與材料損耗資料庫

用 SQLite 記錄點工與材料損耗。標準函式庫即可執行，不需要另外安裝套件。

核定工時取申報工時和門禁工時的較小值。金額是核定工時乘上單價，四捨五入到整數元。申報工時大於門禁工時時，狀態記為「超額攔截」。

材料損耗數量是實際用量減設計用量，損耗率是損耗數量除以設計用量再乘上 100。超過上限（預設 5%）記為「異常超損」。設計用量為 0 時損耗率留空，狀態為「無設計用量，需人工確認」；設計和實際都是 0 時，損耗率記 0。

## 執行

需要 Python 3.9 以上。

```bash
python construction_db.py
```

這會在程式旁邊建立 `construction.db`，並用暫存資料庫跑自我測試。正式檔不會被測試資料寫入。

在程式裡寫入：

```python
from construction_db import (
    db_session,
    init_db,
    insert_contract_item,
    insert_labor_record,
    insert_material_record,
)

init_db()
with db_session() as conn:
    insert_contract_item(conn, "1.1", "結構體混凝土澆置", "立方米", 3200)
    insert_labor_record(conn, "2026-11-02", "A工程行", "王小明", "點工-鋼筋", 10, 8, 350)
    insert_material_record(conn, "2026-11-02", "1.1", "1F柱", 100, 106)
```

材料紀錄的工項編號必須先存在於合約工項，否則這筆寫入會失敗，同一筆交易裡還沒提交的變更也會一起還原。

日期接受 `YYYY-MM-DD`、`YYYY/MM/DD`、`YYYY.MM.DD`，也可以帶時間，例如 `2026-10-01 00:00:00`。存進資料庫時只留日期。
