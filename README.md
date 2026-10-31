# 🏗️ 營建估驗保留款與稅額智慧計算系統 (Engineering Valuation Smart Calculator)

這是一個專為營建工程之**工程估驗、保留款計算與加值型營業稅申報**設計的 Python 智慧計算系統，採用 Streamlit 提供直觀且高互動性的網頁介面。

## 🌟 核心解決痛點與功能

1. **高精度財務四捨五入（防制浮點數累積誤差）**：
   - 採用 Python `decimal` 模組的 `ROUND_HALF_UP` 演算法，取代傳統浮點數及 Python 預設之「銀行家捨入法」（ROUND_HALF_EVEN），確保估驗金額、保留款與營業稅額計算完全符合台灣稅務申報與財務審核標準，精確至整數或指定小數位。
2. **複雜雙列表頭自動辨識與扁平化**：
   - 解決營建工程 Excel 明細表常見的多層級、合併儲存格複雜表頭（例如：上方「本期估驗」，下方「數量」、「金額」）。系統利用 Forward Fill 演算法自動進行表頭扁平化與欄位對齊。
3. **安全防護機制（Excel Formula Injection 防範）**：
   - 防範惡意使用者上傳含有 Excel 注入公式（如以 `=`, `+`, `-`, `@` 開頭之儲存格內容）的表格。系統自動進行跳脫與安全過濾，確保匯出的報表對內部審核人員與財務端 100% 安全。
4. **精美報表導出與會計格式化**：
   - 使用 `openpyxl` 生成具備企業級樣式、自適應欄寬、會計學雙下底線、及 Excel 自訂加總公式之 Excel 明細，免去人工重製報表的繁雜手續。

---

## 🛠️ 技術棧 (Tech Stack)

* **Web UI Framework**: Streamlit
* **Data Processing**: Pandas, NumPy
* **Excel Engine**: Openpyxl
* **Precision Math**: Decimal (Python Standard Library)
* **Testing Suite**: Unittest

---

## 🚀 快速開始 (Getting Started)

### 1. 安裝依賴環境
請確保本機已安裝 Python 3.9+。執行以下指令安裝所需套件：
```bash
pip install -r requirements.txt
```

### 2. 啟動 Streamlit 網頁服務
```bash
streamlit run app.py
```
啟動後即可在瀏覽器開啟本地端服務，並於側邊欄點選 **「下載範例 Excel 估驗表」** 取得標準雙列結構範本進行功能測試。

### 3. 執行單元測試 (Unit Tests)
為確保高精度計算、雙表頭解析、與公式注入防護之邏輯穩定，專案內建完整的單元測試：
```bash
python3 -m unittest test_app.py
```

---

## 📝 專案目錄結構

```
├── app.py              # Streamlit 應用程式主程式與核心邏輯
├── test_app.py         # 涵蓋核心算法、安全防護與解析器之單元測試
├── requirements.txt    # 專案依賴管理文件
└── README.md           # 專案說明文件
```
