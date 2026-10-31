"""計算說明與可下載的 Excel。正式數字只引用合計結果。"""

from __future__ import annotations

from decimal import Decimal
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from mep_app.constants import TAX_GROSS, TAX_NET
from mep_app.money import format_yuan, percent_text, round_half_up

HEADER_FILL = PatternFill("solid", fgColor="1C1917")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TOTAL_FILL = PatternFill("solid", fgColor="F3E6D8")
THIN = Border(
    left=Side(style="thin", color="E7E0D6"),
    right=Side(style="thin", color="E7E0D6"),
    top=Side(style="thin", color="E7E0D6"),
    bottom=Side(style="thin", color="E7E0D6"),
)
TITLE_FONT = Font(bold=True, size=16, color="1C1917")


def build_notes(job: dict, summary: list[dict], detail: list[dict]) -> str:
    basis = job["basis"]
    if basis == TAX_NET:
        tax_rule = "先扣保留款再計稅：稅額＝（本期未稅－保留款）×稅率。"
    else:
        tax_rule = "未稅計稅：稅額＝本期未稅×稅率。"
    parts = [
        "## 實際應付怎麼來",
        "實際應付＝本期未稅＋稅額－保留款。稅額與保留款用進位後的數字回減，應付不再另外進位。",
        tax_rule,
        f"這一期保留款比率 {percent_text(job['retention_rate'])}，營業稅率 {percent_text(job['tax_rate'])}，稅基是{basis}。",
        "同一筆本期未稅 10,000 元、保留款 5%、稅率 5%：未稅計稅的應付是 10,000 元；先扣保留款再計稅的應付是 9,975 元。相差 25 元，是保留款 500 元沒有再課稅。",
        "## 進位",
        "明細上的保留款與稅額四捨五入到小數二位，只供核對。總表把同一分包商的本期金額先加總，保留款與稅額再四捨五入到整數元。",
        "若每一列先乘 5% 再進到整數，兩列各 10 元會得到 1＋1＝2 元。先加總成 20 元再乘 5%，得到 1 元。估驗總表採用 1 元。",
        _gap_sentence(summary, detail),
        "## 三欄閉合",
        "累計應等於前期加本期。只有一欄空白時才回推，三欄都有值但對不上時列入異常，原始數字不改。",
        "契約數量小於或等於 0 時，完成率留空。超過 100% 只警示，不把數量改小。",
        "## 合約依據",
        _clause_sentence(job),
        "條文摘句只說明比率從哪一條來，不取代分包商合計的計算。",
        "## 誰看哪一張表",
        "承辦對總表的閉合與進位。工地主任對異常清單的超額、重複與退項。專案經理、業主與監造對實際應付，以及上面的稅基。",
        "數量橫跨公尺、立方公尺、式、噸，不能加在一起當結論。對帳看本期未稅、保留款、實際應付。",
    ]
    return "\n\n".join(part for part in parts if part)


def make_excel(job: dict, summary: list[dict], detail: list[dict], anomalies: list[dict], diff: dict | None, notes: str) -> bytes:
    book = Workbook()
    _write_summary(book.active, job, summary)
    _write_detail(book.create_sheet("明細"), detail, summary)
    _write_table_sheet(
        book.create_sheet("異常清單"),
        ["來源工作表", "來源列", "分包商", "工程項目", "原金額", "原因", "下一步"],
        anomalies,
        money_cols={"原金額"},
        integer_cols=set(),
    )
    _write_diff(book.create_sheet("兩期差異"), diff)
    sheet = book.create_sheet("計算說明")
    sheet["A1"] = notes
    sheet["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.column_dimensions["A"].width = 96
    sheet.row_dimensions[1].height = 320
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def _gap_sentence(summary: list[dict], detail: list[dict]) -> str:
    official = sum((row["保留款"] for row in summary), Decimal("0"))
    reference = sum((row.get("參考保留款") or Decimal("0") for row in detail if row.get("納入合計")), Decimal("0"))
    gap = reference - official
    if gap == 0:
        return "這份檔案的明細參考保留款加總，與總表整數保留款相同。"
    direction = "多" if gap > 0 else "少"
    return (
        f"這份檔案若把明細參考保留款直接加總，會得到 {format_yuan(reference)} 元，"
        f"比總表的 {format_yuan(official)} 元{direction} {format_yuan(abs(gap))} 元。對帳請用總表。"
    )


def _clause_sentence(job: dict) -> str:
    number = (job.get("clause_no") or "").strip()
    text = (job.get("clause_text") or "").strip()
    if not number and not text:
        return "保留款比率與稅率目前由畫面輸入，尚未附上合約條號或條文摘句。"
    quote = f"{number} {text}".strip()
    return f"合約依據：{quote}"


def _write_summary(sheet, job: dict, summary: list[dict]) -> None:
    sheet.title = "估驗總表"
    sheet["A1"] = "多分包商估驗彙整"
    sheet["A1"].font = TITLE_FONT
    meta = [
        ("工程名稱", job.get("project") or ""),
        ("合約編號", job.get("contract_no") or ""),
        ("期別", job.get("period") or ""),
        ("稅額計算基礎", job.get("basis") or TAX_GROSS),
        ("保留款比率", percent_text(job["retention_rate"])),
        ("營業稅率", percent_text(job["tax_rate"])),
        ("合約條號", job.get("clause_no") or "（未填）"),
        ("條文摘句", job.get("clause_text") or "（未填）"),
        ("製作時間", job.get("generated_at") or ""),
    ]
    for index, (label, value) in enumerate(meta, start=3):
        sheet.cell(index, 1, label).font = Font(bold=True, color="57534E")
        sheet.cell(index, 2, value)
    header_row = 13
    headers = ["分包商", "項次數", "本期未稅", "保留款", "稅額", "實際應付"]
    _paint_header(sheet, header_row, headers)
    for offset, row in enumerate(summary):
        values = [row["分包商"], row["項次數"], row["本期未稅"], row["保留款"], row["稅額"], row["實際應付"]]
        _paint_values(sheet, header_row + 1 + offset, values, money={3, 6}, integer={4, 5}, qty=set())
    total_row = header_row + 1 + len(summary)
    totals = ["合計", sum(row["項次數"] for row in summary)]
    for key in ("本期未稅", "保留款", "稅額", "實際應付"):
        totals.append(sum((row[key] for row in summary), Decimal("0")))
    _paint_values(sheet, total_row, totals, money={3, 6}, integer={4, 5}, qty=set(), fill=TOTAL_FILL)
    sign_row = total_row + 2
    sheet.cell(sign_row, 1, "簽核").font = Font(bold=True)
    signs = [
        ("承辦人", "對總表的數量金額是否閉合，以及保留款、稅額是否在合計後進位"),
        ("工地主任", "對異常清單的超額、重複與可能退項"),
        ("專案經理", "對本期實際應付"),
        ("業主／監造", "對實際應付與稅基、合約條文"),
    ]
    _paint_header(sheet, sign_row + 1, ["角色", "看哪一張結果", "簽名"])
    for offset, (role, duty) in enumerate(signs):
        _paint_values(sheet, sign_row + 2 + offset, [role, duty, ""], money=set(), integer=set(), qty=set())
    sheet.column_dimensions["A"].width = 28
    sheet.column_dimensions["B"].width = 62
    for column in range(3, 7):
        sheet.column_dimensions[get_column_letter(column)].width = 16
    sheet.freeze_panes = "A14"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    _note_tax_name(sheet, job)


def _note_tax_name(sheet, job: dict) -> None:
    if job.get("basis") not in (TAX_GROSS, TAX_NET):
        sheet["B6"] = TAX_GROSS


def _write_detail(sheet, detail: list[dict], summary: list[dict]) -> None:
    headers = [
        "來源工作表",
        "來源列",
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
        "累計完成率",
        "參考保留款",
        "參考稅額",
        "實際應付",
        "補值",
    ]
    _paint_header(sheet, 1, headers)
    official = {row["分包商"]: row for row in summary}
    excel_row = 2
    grouped: dict[str, list[dict]] = {}
    order: list[str] = []
    for record in detail:
        vendor = record.get("分包商") or ""
        if vendor not in grouped:
            grouped[vendor] = []
            order.append(vendor)
        grouped[vendor].append(record)
    for vendor in order:
        for record in grouped[vendor]:
            values = [_detail_value(record, header) for header in headers]
            _paint_values(
                sheet,
                excel_row,
                values,
                money={7, 9, 11, 13, 15, 16},
                integer=set(),
                qty={6, 8, 10, 12},
            )
            excel_row += 1
        total = official.get(vendor)
        if total:
            footer = [""] * len(headers)
            footer[2] = vendor
            footer[3] = "分包商合計（正式數字）"
            footer[10] = total["本期未稅"]
            footer[14] = total["保留款"]
            footer[15] = total["稅額"]
            footer[16] = total["實際應付"]
            _paint_values(
                sheet,
                excel_row,
                footer,
                money={11, 17},
                integer={15, 16},
                qty=set(),
                fill=TOTAL_FILL,
            )
            excel_row += 1
    for index, header in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = 18 if len(header) > 4 else 14
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(excel_row - 1, 1)}"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.sheet_view.view = "pageBreakPreview"
    sheet.sheet_view.zoomScale = 90
    sheet.oddHeader.left.text = "明細"
    sheet.oddFooter.right.text = "第 &P 頁"


def _detail_value(record: dict, header: str):
    if header == "實際應付":
        return None
    if header == "累計完成率":
        rate = record.get(header)
        if rate is None:
            return None
        return f"{(rate * Decimal('100')):.2f}%"
    return record.get(header)


def _write_diff(sheet, diff: dict | None) -> None:
    sheet["A1"] = "跟上期的差異。這裡只列出工項變化，不重算稅額與保留款。"
    sheet["A1"].font = Font(bold=True)
    if not diff:
        sheet["A3"] = "這次沒有上傳上期檔。"
        sheet.column_dimensions["A"].width = 72
        return
    cursor = 3
    sections = [
        ("新增工項", ["分包商", "工程項目", "單位", "本期數量", "本期金額", "單價"], diff["added"]),
        (
            "數量變動",
            ["分包商", "工程項目", "單位", "上期本期數量", "本期數量", "上期累計數量", "本期累計數量"],
            diff["quantity_changes"],
        ),
        ("單價變動", ["分包商", "工程項目", "單位", "上期單價", "本期單價"], diff["price_changes"]),
        ("本期未列", ["分包商", "工程項目", "單位", "本期數量", "本期金額", "單價"], diff["missing"]),
        ("無法對齊", ["出現在", "分包商", "工程項目", "單位", "筆數", "本期金額合計"], diff["unaligned"]),
    ]
    for title, headers, rows in sections:
        sheet.cell(cursor, 1, title).font = Font(bold=True, color="9A3412")
        cursor += 1
        _paint_header(sheet, cursor, headers)
        cursor += 1
        if not rows:
            sheet.cell(cursor, 1, "無")
            cursor += 2
            continue
        for row in rows:
            values = [row.get(header) for header in headers]
            money = {index + 1 for index, header in enumerate(headers) if "金額" in header or header.endswith("單價")}
            qty = {index + 1 for index, header in enumerate(headers) if "數量" in header}
            _paint_values(sheet, cursor, values, money=money, integer=set(), qty=qty)
            cursor += 1
        cursor += 1
    sheet.column_dimensions["A"].width = 24
    for column in range(2, 8):
        sheet.column_dimensions[get_column_letter(column)].width = 18


def _write_table_sheet(sheet, headers: list[str], rows: list[dict], money_cols: set[str], integer_cols: set[str]) -> None:
    _paint_header(sheet, 1, headers)
    for offset, row in enumerate(rows):
        values = [row.get(header) for header in headers]
        money = {index + 1 for index, header in enumerate(headers) if header in money_cols}
        integer = {index + 1 for index, header in enumerate(headers) if header in integer_cols}
        _paint_values(sheet, offset + 2, values, money=money, integer=integer, qty=set())
    for index, header in enumerate(headers, start=1):
        width = 42 if header == "下一步" else 18
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    if headers:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(rows) + 1, 1)}"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True


def _paint_header(sheet, row_no: int, headers: list[str]) -> None:
    for index, header in enumerate(headers, start=1):
        cell = sheet.cell(row_no, index, header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
        cell.border = THIN
    sheet.row_dimensions[row_no].height = 22
    sheet.sheet_view.showGridLines = False
    sheet.page_setup.horizontalCentered = True
    sheet.page_margins.left = 0.4
    sheet.page_margins.right = 0.4
    sheet.page_margins.top = 0.6
    sheet.page_margins.bottom = 0.6
    sheet.sheet_properties.tabColor = "9A3412"
    sheet.oddHeader.center.text = sheet.title
    sheet.oddFooter.center.text = "正式數字以分包商合計為準"


def _paint_values(sheet, row_no, values, money: set[int], integer: set[int], qty: set[int], fill=None) -> None:
    for index, value in enumerate(values, start=1):
        cell = sheet.cell(row_no, index, _excel_value(value, index in money, index in integer, index in qty))
        cell.border = THIN
        cell.alignment = Alignment(vertical="center", wrap_text=index == 2 and sheet.title == "估驗總表")
        if fill is not None:
            cell.fill = fill
            cell.font = Font(bold=True)
        if index in integer:
            cell.number_format = "#,##0"
        elif index in money:
            cell.number_format = "#,##0.00"
        elif index in qty:
            cell.number_format = "#,##0.####"


def _excel_value(value, money: bool, integer: bool, qty: bool):
    if value is None:
        return None
    if isinstance(value, Decimal):
        if integer:
            return int(round_half_up(value, 0))
        if money:
            return float(round_half_up(value, 2))
        if qty:
            return float(round_half_up(value, 4))
        return float(value)
    return value
