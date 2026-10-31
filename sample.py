"""做出一份故意不整齊的範例，用來核對閉合、排除小計與兩期差異。"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
SAMPLE_PATH = ROOT / "examples" / "sample.xlsx"
PREVIOUS_PATH = ROOT / "examples" / "sample_previous.xlsx"

TOP = ["分包商", "工程項目", "單位", "契約數量", "單價", "前期累計", "前期累計", "本期", "本期", "累計", "累計"]
BOTTOM = [None, None, None, None, None, "數量", "金額", "數量", "金額", "數量", "金額"]
FULL = [
    "分包商",
    "工程項目",
    "單位",
    "契約數量",
    "單價",
    "前期累計數量",
    "前期累計金額",
    "本期數量",
    "本期金額",
    "累計數量",
    "累計金額",
]


def write_examples(folder: Path | None = None) -> None:
    target = folder or SAMPLE_PATH.parent
    target.mkdir(parents=True, exist_ok=True)
    _save(_current_book(), target / "sample.xlsx")
    _save(_previous_book(), target / "sample_previous.xlsx")


def _save(book: Workbook, path: Path) -> None:
    book.save(path)


def _current_book() -> Workbook:
    book = Workbook()
    cover = book.active
    cover.title = "封面"
    cover["A1"] = "這張封面沒有欄位標題，計算時應略過。"
    _two_row(
        book,
        "水電",
        ["水電分包估驗表"],
        [
            ["水電行", "配管", "m", 100, 1000, 40, None, 10, None, None, None],
            [None, "說明牌", "式", 1, 5000, 0, 0, 1, None, None, None],
            [None, "累計土方", "M3", 10, 100, 0, 0, 2, None, None, None],
            [None, "小計", None, None, None, None, None, None, 99999, None, None],
        ],
        merges=["A5:A7"],
    )
    _two_row(
        book,
        "進位示範",
        ["兩列各 10 元，用來核對合計後進位"],
        [
            [None, "工項甲", "式", None, None, None, None, None, 10, None, None],
            [None, "工項乙", "式", None, None, None, None, None, 10, None, None],
        ],
    )
    _two_row(
        book,
        "稅額進位",
        ["1,000 元用來核對 47.5 進位"],
        [
            ["稅額示範", "工項", "式", None, None, None, None, None, 1000, None, None],
        ],
    )
    _single(
        book,
        "異常",
        [
            ["旗標", "不閉合", "m", 10, 100, 10, 1000, 10, 1000, 30, 3000],
            ["旗標", "超做", "m", 10, 100, 0, 0, 12, 1200, 12, 1200],
            ["旗標", "退項", "式", 1, 100, 0, 0, -1, -100, -1, -100],
            ["旗標", "單價不符", "m", 5, 100, 0, 0, 2, 250, 2, 250],
            ["旗標", "重複項", "式", 1, 50, 0, 0, 1, 50, 1, 50],
            ["旗標", "重複項", "式", 1, 50, 0, 0, 1, 50, 1, 50],
        ],
    )
    return book


def _previous_book() -> Workbook:
    book = Workbook()
    cover = book.active
    cover.title = "封面"
    cover["A1"] = "上期封面"
    _alias_sheet(
        book,
        "水電",
        "上期水電，表頭用不同說法",
        [
            ["水電行", "配管", "m", 100, 1000, 30, 30000, 8, 8000, 38, 38000],
            ["水電行", "累計土方", "M3", 10, 90, 0, 0, 2, 180, 2, 180],
            ["水電行", "臨時抽水", "式", 1, 300, 0, 0, 1, 300, 1, 300],
        ],
    )
    _two_row(
        book,
        "進位示範",
        ["上期進位"],
        [
            [None, "工項甲", "式", None, None, None, None, None, 10, None, None],
            [None, "工項乙", "式", None, None, None, None, None, 10, None, None],
        ],
    )
    _two_row(
        book,
        "稅額進位",
        ["上期稅額"],
        [
            ["稅額示範", "工項", "式", None, None, None, None, None, 1000, None, None],
        ],
    )
    _single(
        book,
        "異常",
        [
            ["旗標", "不閉合", "m", 10, 100, 10, 1000, 10, 1000, 30, 3000],
            ["旗標", "超做", "m", 10, 100, 0, 0, 12, 1200, 12, 1200],
            ["旗標", "退項", "式", 1, 100, 0, 0, -1, -100, -1, -100],
            ["旗標", "單價不符", "m", 5, 100, 0, 0, 2, 250, 2, 250],
            ["旗標", "重複項", "式", 1, 50, 0, 0, 1, 50, 1, 50],
            ["旗標", "重複項", "式", 1, 50, 0, 0, 1, 50, 1, 50],
        ],
    )
    return book


def _two_row(book: Workbook, name: str, title: list, data: list[list], merges: list[str] | None = None) -> None:
    sheet = book.create_sheet(name)
    sheet.append(title)
    sheet.append([])
    sheet.append(TOP)
    sheet.append(BOTTOM)
    for row in data:
        sheet.append(row)
    for merge in merges or []:
        sheet.merge_cells(merge)


def _alias_sheet(book: Workbook, name: str, title: str, data: list[list]) -> None:
    headers = ["廠商", "工項", "單位", "契約數量", "單價", "前期數量", "前期金額", "本月數量", "本月金額", "累計數量", "累計金額"]
    sheet = book.create_sheet(name)
    sheet.append([title])
    sheet.append([])
    sheet.append(headers)
    for row in data:
        sheet.append(row)


def _single(book: Workbook, name: str, data: list[list]) -> None:
    sheet = book.create_sheet(name)
    sheet.append(["異常案例"])
    sheet.append([])
    sheet.append(FULL)
    for row in data:
        sheet.append(row)


if __name__ == "__main__":
    write_examples()
    print(SAMPLE_PATH)
    print(PREVIOUS_PATH)
