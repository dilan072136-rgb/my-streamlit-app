"""工程估驗彙整的計算、讀檔與報表測試。"""
import unittest
from io import BytesIO

import pandas as pd
from openpyxl import Workbook, load_workbook

import app


def frame_from_rows(rows, sheet="甲工區"):
    columns = list(rows[0])
    data = pd.DataFrame(rows[1:], columns=columns)
    data["來源工作表"] = sheet
    data["來源列"] = range(2, len(data) + 2)
    return data


def run_clean(rows, **kwargs):
    data = frame_from_rows(rows)
    mapping = {field: field if field in data.columns else app.NONE for field in app.CANONICAL}
    params = dict(retention_rate=0.05, tax_rate=0.05, tolerance=1, anomaly_limit=0,
                  ffill_vendor=True, tax_on_net=False)
    params.update(kwargs)
    return app.build_clean_data(data, mapping, **params)


def workbook_bytes(builder, name="t.xlsx"):
    wb = Workbook()
    ws = wb.active
    builder(wb, ws)
    buf = BytesIO()
    wb.save(buf)
    buf.name = name
    buf.seek(0)
    return buf


class MoneyTests(unittest.TestCase):
    def test_half_up_not_bankers(self):
        self.assertEqual(app.round_half_up(pd.Series([0.025, 0.035, 1.225, 2.5, -1.5]), 2).tolist(),
                         [0.03, 0.04, 1.23, 2.5, -1.5])
        self.assertEqual(app.round_half_up(0.5, 0), 1.0)
        self.assertEqual(app.round_half_up(-0.5, 0), -1.0)
        # 0.70 * 5% 的二進位是 0.034999999999999996，銀行家捨入會得到 0.03。
        self.assertEqual(app.round_half_up(app._mul_scalar(pd.Series([0.7]), 0.05), 2).iloc[0], 0.04)

    def test_line_retention_uses_half_up(self):
        clean, _ = run_clean([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "挖土", 0.5],
            ["甲", "回填", 0.7],
        ])
        by_item = dict(zip(clean["工程項目"], clean["保留款"]))
        self.assertEqual(by_item["挖土"], 0.03)
        self.assertEqual(by_item["回填"], 0.04)

    def test_vendor_rounds_half_up_to_integer(self):
        clean, _ = run_clean([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "挖土", 10],
        ])
        vendor = app.summarize_vendors(clean, 0.05, 0.05, tax_on_net=False)
        self.assertEqual(vendor.loc[0, "本期保留款"], 1.0)
        self.assertEqual(vendor.loc[0, "本期稅額"], 1.0)
        self.assertEqual(vendor.loc[0, "實際應付金額"], 10.0)
        self.assertEqual(clean.loc[0, "保留款"], 0.5)

    def test_tax_on_net_rounds_47_5_up(self):
        clean, _ = run_clean([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "挖土", 1000],
        ], tax_on_net=True)
        vendor = app.summarize_vendors(clean, 0.05, 0.05, tax_on_net=True)
        self.assertEqual(vendor.loc[0, "本期保留款"], 50.0)
        self.assertEqual(vendor.loc[0, "本期稅額"], 48.0)
        self.assertEqual(vendor.loc[0, "實際應付金額"], 998.0)

    def test_decimal_sum_of_tenths(self):
        clean, _ = run_clean([
            ["分包商", "工程項目", "本期金額"],
            *[["甲", f"項{i}", 0.1] for i in range(10)],
        ])
        vendor = app.summarize_vendors(clean, 0, 0)
        self.assertEqual(vendor.loc[0, "本期未稅估驗"], 1.0)

    def test_money_display_half_up(self):
        self.assertEqual(app.money(1.225), "NT$ 1.23")
        self.assertEqual(app.money(2.5), "NT$ 2.50")


class ParseTests(unittest.TestCase):
    def test_fullwidth_parentheses_and_currency(self):
        raw = pd.Series(["１，２３４", "（50）", "1,000元", "NT$1,200", "100-", "-", "", "NA"])
        out = app.to_number(raw)
        self.assertEqual(out.iloc[0], 1234.0)
        self.assertEqual(out.iloc[1], -50.0)
        self.assertEqual(out.iloc[2], 1000.0)
        self.assertEqual(out.iloc[3], 1200.0)
        self.assertEqual(out.iloc[4], -100.0)
        self.assertTrue(out.iloc[5:].isna().all())

    def test_numeric_vendor_and_fullwidth_name(self):
        text = app.clean_text(pd.Series([101.0, "Ａ廠", "  甲  ", "NA"]))
        self.assertEqual(text.tolist(), ["101", "A廠", "甲", "NA"])

    def test_short_headers_with_suffix(self):
        cols = ["單價(元)", "金額(未稅)", "未稅金額", "前期金額", "本期完成數量(m)",
                "單位(m2)", "契約數量(m2)", "金額合計", "分包商", "工程項目"]
        rename = app.auto_rename(cols)
        self.assertEqual(rename["單價(元)"], "單價")
        self.assertEqual(rename["金額(未稅)"], "本期金額")
        self.assertNotIn("未稅金額", rename)
        self.assertEqual(rename["前期金額"], "前期累計金額")
        self.assertEqual(rename["本期完成數量(m)"], "本期數量")
        self.assertEqual(rename["單位(m2)"], "單位")
        self.assertEqual(rename["契約數量(m2)"], "契約數量")
        self.assertNotIn("金額合計", rename)

    def test_summary_does_not_eat_real_items(self):
        self.assertTrue(app.is_summary_label("合計（未稅）"))
        self.assertTrue(app.is_summary_label("小計1"))
        self.assertTrue(app.is_summary_label("稅額5%"))
        self.assertFalse(app.is_summary_label("說明牌"))
        self.assertFalse(app.is_summary_label("累計土方"))
        self.assertFalse(app.is_summary_label("應付土方"))

    def test_na_text_is_kept(self):
        def build(wb, ws):
            ws.append(["分包商", "工程項目", "本期金額"])
            ws.append(["NA", "代碼列", 10])
        bio = workbook_bytes(build)
        source, _, _, _ = app.read_workbook(bio)
        self.assertEqual(source.loc[0, "分包商"], "NA")


class CleanTests(unittest.TestCase):
    def test_summary_vendor_does_not_bleed(self):
        clean, anomalies = run_clean([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "挖土", 10],
            ["合計商", "合計", 99],
            [None, "回填", 20],
        ])
        self.assertEqual(set(clean["工程項目"]), {"挖土", "回填"})
        self.assertTrue((clean["分包商"] == "甲").all())
        self.assertIn("疑似合計／小計列（已排除，未列入計算）", set(anomalies["異常原因"]))

    def test_real_items_kept(self):
        clean, anomalies = run_clean([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "說明牌", 10],
            ["甲", "累計土方", 20],
            ["甲", "應付土方", 30],
            ["甲", "合計（未稅）", 999],
        ])
        self.assertEqual(set(clean["工程項目"]), {"說明牌", "累計土方", "應付土方"})
        self.assertEqual(len(anomalies), 1)

    def test_blank_sheet_uses_sheet_name(self):
        clean, _ = run_clean([
            ["分包商", "工程項目", "本期金額"],
            [None, "挖土", 10],
        ])
        self.assertEqual(clean.loc[0, "分包商"], "甲工區")

    def test_ffill_does_not_cross_sheets(self):
        rows = frame_from_rows([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "挖土", 10],
        ], sheet="甲")
        other = frame_from_rows([
            ["分包商", "工程項目", "本期金額"],
            [None, "回填", 20],
        ], sheet="乙")
        data = pd.concat([rows, other], ignore_index=True)
        mapping = {field: field if field in data.columns else app.NONE for field in app.CANONICAL}
        clean, _ = app.build_clean_data(data, mapping, 0.05, 0.05, 1, 0, True, False)
        vendors = dict(zip(clean["工程項目"], clean["分包商"]))
        self.assertEqual(vendors["挖土"], "甲")
        self.assertEqual(vendors["回填"], "乙")

    def test_only_summary_is_empty(self):
        clean, anomalies = run_clean([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "合計", 10],
        ])
        self.assertTrue(clean.empty)
        self.assertEqual(len(anomalies), 1)

    def test_duplicate_negative_and_rate(self):
        clean, anomalies = run_clean([
            ["分包商", "工程項目", "單位", "契約數量", "單價", "本期數量", "本期金額"],
            ["甲", "挖土", "m", 10, 9, 10, 100],
            ["甲", "挖土", "m", 10, 9, 10, 100],
            ["甲", "退項", "式", 1, 5, 1, -5],
            ["甲", "超量", "m", 10, 1, 12, 12],
        ])
        reasons = set(anomalies["異常原因"])
        self.assertTrue(any(r.startswith("本期金額與") for r in reasons))
        self.assertTrue(any("負數" in r for r in reasons))
        self.assertTrue(any("100%" in r for r in reasons))
        self.assertTrue(any(r.startswith("疑似重複列") for r in reasons))

    def test_same_length_files_do_not_share_token(self):
        self.assertNotEqual(app.file_token(b"aaaa"), app.file_token(b"bbbb"))


class ReadTests(unittest.TestCase):
    def test_title_row_does_not_eat_data(self):
        def build(wb, ws):
            ws.append(["工程名稱", "測試工程"])
            ws.append(["分包商", "工程項目", "單位", "契約數量", "單價", "前期累計數量", "本期數量", "本期金額"])
            ws.append(["甲", "挖土", "m", 10, 100, 1, 2, 200])
        source, notes, _, warnings = app.read_workbook(workbook_bytes(build))
        self.assertEqual(len(source), 1)
        self.assertEqual(source.loc[0, "分包商"], "甲")
        self.assertEqual(source.loc[0, "本期金額"], 200)
        self.assertEqual(warnings, [])
        self.assertIn("第 2 列", notes[0])

    def test_two_row_header(self):
        def build(wb, ws):
            ws.append(["估驗表", None, None, None, None, None])
            ws.append([None, None, "前期累計", None, "本期", None])
            ws.append(["分包商", "工程項目", "數量", "金額", "數量", "金額"])
            ws.append(["甲", "挖土", 1, 10, 2, 20])
        source, notes, _, _ = app.read_workbook(workbook_bytes(build))
        self.assertEqual(source.loc[0, "分包商"], "甲")
        self.assertEqual(source.loc[0, "前期累計數量"], 1)
        self.assertEqual(source.loc[0, "前期累計金額"], 10)
        self.assertEqual(source.loc[0, "本期數量"], 2)
        self.assertEqual(source.loc[0, "本期金額"], 20)
        self.assertIn("第 2–3 列", notes[0])

    def test_formula_warning_and_injection(self):
        def build(wb, ws):
            ws.append(["分包商", "工程項目", "本期金額"])
            ws["A2"] = "甲"
            ws["B2"] = "=1+1"
            ws["C2"] = 10
        source, _, _, warnings = app.read_workbook(workbook_bytes(build))
        self.assertEqual(warnings, [app.FORMULA_WARN])
        clean, _ = run_clean([
            ["分包商", "工程項目", "本期金額"],
            ["甲", "=1+1", 10],
        ])
        vendor = app.summarize_vendors(clean, 0.05, 0.05)
        data = app.make_excel(clean, vendor, pd.DataFrame(columns=app.ANOM_COLS),
                              "=1+1", "C", "第1期", 0.05, 0.05)
        wb = load_workbook(BytesIO(data))
        detail = wb[wb.sheetnames[1]]
        found = [cell for row in detail.iter_rows() for cell in row if cell.value == "=1+1"]
        self.assertTrue(found)
        self.assertTrue(all(cell.data_type == "s" for cell in found))

    def test_sheet_names_ignore_case(self):
        rows = []
        for name, amount in (("ABC", 10), ("abc", 20), ("估驗總表", 30)):
            part = frame_from_rows([
                ["分包商", "工程項目", "本期金額"],
                [name, "挖土", amount],
            ], sheet=name)
            rows.append(part)
        data = pd.concat(rows, ignore_index=True)
        mapping = {field: field if field in data.columns else app.NONE for field in app.CANONICAL}
        clean, anomalies = app.build_clean_data(data, mapping, 0.05, 0.05, 1, 0, True, False)
        vendor = app.summarize_vendors(clean, 0.05, 0.05)
        blob = app.make_excel(clean, vendor, anomalies, "工程", "C", "第1期", 0.05, 0.05)
        names = load_workbook(BytesIO(blob)).sheetnames
        self.assertEqual(names[0], "估驗總表")
        folded = [n.casefold() for n in names]
        self.assertEqual(len(folded), len(set(folded)))
        self.assertIn("abc", folded)
        self.assertIn("abc2", folded)

    def test_contract_qty_in_detail(self):
        clean, anomalies = run_clean([
            ["分包商", "工程項目", "契約數量", "本期金額"],
            ["甲", "挖土", 10, 100],
        ])
        vendor = app.summarize_vendors(clean, 0.05, 0.05)
        blob = app.make_excel(clean, vendor, anomalies, "工程", "", "", 0.05, 0.05)
        wb = load_workbook(BytesIO(blob))
        headers = [cell.value for cell in next(wb[wb.sheetnames[1]].iter_rows(min_row=1, max_row=20))]
        # 標題列可能不在第一列，改掃前 15 列找表頭
        found = False
        for row in wb[wb.sheetnames[1]].iter_rows(max_row=15, max_col=20):
            values = [cell.value for cell in row]
            if "契約數量" in values:
                found = True
        self.assertTrue(found)

    def test_rejects_oversized_zip(self):
        def build(wb, ws):
            ws.append(["分包商", "本期金額"])
            ws.append(["甲", 1])
        data = workbook_bytes(build).getvalue()
        with self.assertRaises(ValueError):
            app.assert_safe_workbook(data, "a.xlsx", max_total=50)

    def test_xls(self):
        import xlwt
        book = xlwt.Workbook()
        sheet = book.add_sheet("S")
        sheet.write(0, 0, "分包商")
        sheet.write(0, 1, "本期金額")
        sheet.write(1, 0, "甲")
        sheet.write(1, 1, 100.5)
        buf = BytesIO()
        book.save(buf)
        buf.name = "a.xls"
        buf.seek(0)
        source, _, _, _ = app.read_workbook(buf)
        self.assertEqual(source.loc[0, "分包商"], "甲")
        self.assertEqual(source.loc[0, "本期金額"], 100.5)

    def test_empty_workbook(self):
        def build(wb, ws):
            ws.append([None, None])
        with self.assertRaises(ValueError):
            app.read_workbook(workbook_bytes(build))


def _sample_xlsx():
    def build(wb, ws):
        ws.title = "甲工區"
        ws.append(["工程名稱", "測試工程"])
        ws.append(["分包商", "工程項目", "單位(m)", "契約數量", "單價(元)", "前期累計數量", "本期數量", "金額(未稅)"])
        ws.append(["甲廠", "挖土", "m3", 100, 50, 10, 5, None])
        ws.append([None, "回填", "m3", 80, 40, 0, 4, 160])
        ws.append([None, "合計（未稅）", None, None, None, None, None, 999])
        ws.append(["合計商", "說明牌", "式", 1, 20, 0, 1, 20])
        other = wb.create_sheet("乙工區")
        other.append(["廠商", "品名", "本期完成數量(m)", "單價"])
        other.append(["乙廠", "累計土方", 3, 15])
        other.append(["乙廠", "應付土方", 2, 30])
    return workbook_bytes(build).getvalue()


class UiTests(unittest.TestCase):
    def test_app_boots(self):
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file("app.py", default_timeout=30)
        at.run()
        self.assertFalse(at.exception)

    def test_upload_and_calculate(self):
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file("app.py", default_timeout=40)
        at.run()
        at.file_uploader[0].upload(
            "估驗.xlsx", _sample_xlsx(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        at.run()
        self.assertFalse(at.exception)
        self.assertTrue(at.button)
        at.button[0].click().run()
        self.assertFalse(at.exception)
        labels = [m.label for m in at.metric]
        self.assertIn("實際應付金額", labels)
        self.assertTrue(at.download_button)
        # 換一個欄位不同的檔，選單不應因舊鍵當掉。
        def other(wb, ws):
            ws.append(["協力廠商", "工作項目", "契約單價", "數量"])
            ws.append(["丙", "模板", 8, 4])
        at.file_uploader[0].upload(
            "other.xlsx", workbook_bytes(other).getvalue(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        at.run()
        self.assertFalse(at.exception)


if __name__ == "__main__":
    unittest.main()
