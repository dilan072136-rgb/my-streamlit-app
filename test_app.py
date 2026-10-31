import unittest
import pandas as pd
import numpy as np
from decimal import Decimal
from app import precise_round, sanitize_value, parse_double_row_headers, calculate_valuation_row

class TestEngineeringValuationApp(unittest.TestCase):
    
    def test_precise_rounding(self):
        """
        測試 Decimal 高精度四捨五入（ROUND_HALF_UP）是否運作正常。
        驗證傳統浮點數造成的誤差，在 precise_round 中是否能得到財務與稅法要求的精確結果。
        """
        # 傳統浮點數 2.5 在 Python 的 round(2.5) 為 2 （銀行家捨入）
        # 但在營建估驗實務與台灣稅額計算中，必須四捨五入為 3
        self.assertEqual(precise_round(2.5), 3)
        self.assertEqual(precise_round(3.5), 4)
        
        # 浮點數微小偏差測試 (例如 1.0005 四捨五入至小數第三位)
        self.assertEqual(precise_round(1.0005, 3), 1.001)
        self.assertEqual(precise_round(1.0004, 3), 1.000)
        
        # 處理 None / NaN 與異常值
        self.assertEqual(precise_round(None), 0)
        self.assertEqual(precise_round(np.nan), 0)
        self.assertEqual(precise_round("invalid"), 0)

    def test_sanitize_formula_injection(self):
        """
        測試防範 Excel 公式注入（Excel Formula Injection）的安全防護邏輯。
        任何以 =, +, -, @ 開頭的字串輸入都應加上單引號 "'" 跳脫。
        """
        self.assertEqual(sanitize_value("=SUM(A1:A10)"), "'=SUM(A1:A10)")
        self.assertEqual(sanitize_value("+100"), "'+100")
        self.assertEqual(sanitize_value("-50"), "'-50")
        self.assertEqual(sanitize_value("@SUM"), "'@SUM")
        
        # 正常輸入不應受影響
        self.assertEqual(sanitize_value("普通工項名稱"), "普通工項名稱")
        self.assertEqual(sanitize_value(12345), 12345)
        self.assertEqual(sanitize_value(12.345), 12.345)
        self.assertEqual(sanitize_value("   =SUM(A1)"), "'=SUM(A1)") # 測試帶有空格的開頭

    def test_parse_double_row_headers(self):
        """
        測試自動辨識並扁平化雙列複雜表頭的邏輯。
        """
        # 模擬一個帶有雙列表頭的 Excel 資料結構
        data = {
            "項目": [np.nan, "1", "2"],
            "工項名稱": [np.nan, "土方開挖", "鋼筋組立"],
            "合約": ["數量", 100, 200],
            "Unnamed: 3": ["單價", 500, 1500],
            "本期估驗": ["數量", 20, 50],
            "Unnamed: 5": ["金額", 10000, 75000]
        }
        df = pd.DataFrame(data)
        
        # 進行扁平化解析
        cleaned_df, flat_headers = parse_double_row_headers(df)
        
        # 預期欄位
        expected_headers = [
            "項目",
            "工項名稱",
            "合約_數量",
            "合約_單價",
            "本期估驗_數量",
            "本期估驗_金額"
        ]
        
        self.assertEqual(flat_headers, expected_headers)
        self.assertEqual(cleaned_df.columns.tolist(), expected_headers)
        self.assertEqual(len(cleaned_df), 2)
        # 檢查第一筆數據
        self.assertEqual(cleaned_df.iloc[0]["項目"], "1")
        self.assertEqual(float(cleaned_df.iloc[0]["合約_數量"]), 100)

    def test_calculate_valuation_row(self):
        """
        測試單一工項之估驗及保留款的核心計算公式。
        公式：
          - 本期估驗_金額 = 本期估驗_數量 * 合約_單價
          - 累計估驗_數量 = 歷史累計_數量 + 本期估驗_數量
          - 累計估驗_金額 = 累計估驗_數量 * 合約_單價
          - 本期保留款 = 本期估驗_金額 * 保留款比例 (10%)
        """
        row = {
            "項目": "1",
            "工項名稱": "鋼筋混凝土澆置",
            "合約_單價": 3200,
            "歷史累計_數量": 50,
            "本期估驗_數量": 10
        }
        
        # 設定保留款比例 10% (0.1)
        result = calculate_valuation_row(row, retention_rate=0.1)
        
        # 10 * 3200 = 32000
        self.assertEqual(result['本期估驗_金額'], 32000)
        # 50 + 10 = 60
        self.assertEqual(result['累計估驗_數量'], 60)
        # 60 * 3200 = 192000
        self.assertEqual(result['累計估驗_金額'], 192000)
        # 32000 * 0.10 = 3200
        self.assertEqual(result['本期保留款'], 3200)

if __name__ == '__main__':
    unittest.main()