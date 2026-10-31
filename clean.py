"""去掉合計列、沿用分包商，並把讀不懂的儲存格留在異常清單。"""

from __future__ import annotations

import unicodedata
from decimal import Decimal

from mep_app.columns import usable
from mep_app.constants import NUMERIC_FIELDS, SIX_FIELDS, SUMMARY_KEYWORDS
from mep_app.money import format_yuan, to_decimal
from mep_app.workbook import DataRow, SheetTable


def is_summary_label(text: object) -> bool:
    """合計、小計、總計、營業稅、保留款不進金額。說明牌、累計土方要留下。"""
    compact = _compact(text)
    if not compact:
        return False
    for keyword in SUMMARY_KEYWORDS:
        if compact == keyword:
            return True
        if compact.startswith(keyword):
            rest = compact[len(keyword) :]
            digits = rest.replace("%", "").replace(".", "")
            if rest[:1] in "(（:：-" or rest in {"列", "項"} or _only_decoration(rest) or digits.isdigit():
                return True
        if compact.endswith(keyword) and len(compact) <= len(keyword) + 4:
            return True
    return False


def _compact(text: object) -> str:
    if text is None:
        return ""
    raw = unicodedata.normalize("NFKC", str(text))
    if raw.strip().lower() == "nan":
        return ""
    return "".join(raw.split())


def _only_decoration(rest: str) -> bool:
    return rest.strip(" :：()（）.·-—_／/、") == ""


def as_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value != value:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def make_anomaly(record: dict, reason: str, action: str) -> dict:
    """句子只用這一列的原金額，不另外乘稅率或保留款。"""
    amount = record.get("原本期金額")
    sheet = record.get("來源工作表", "")
    row_no = record.get("來源列", "")
    if isinstance(amount, Decimal):
        where = f"工作表「{sheet}」第 {row_no} 列，原本期金額 {format_yuan(amount)} 元。"
    else:
        where = f"工作表「{sheet}」第 {row_no} 列，原本期金額空白。"
    return {
        "來源工作表": sheet,
        "來源列": row_no,
        "分包商": record.get("分包商") or "",
        "工程項目": record.get("工程項目") or "",
        "原金額": amount,
        "原因": reason,
        "下一步": where + action,
    }


def build_clean_data(sheet: SheetTable, selection: dict[str, str]) -> tuple[list[dict], list[dict]]:
    """合計列先排除，才沿用上一列分包商，避免合計列上的廠商名稱流到下一筆。"""
    pending: list[dict] = []
    anomalies: list[dict] = []
    for data_row in sheet.rows:
        record, row_flags, keep = _read_row(sheet.name, data_row, selection)
        anomalies.extend(row_flags)
        if keep:
            pending.append(record)

    _assign_vendors(pending, sheet.name)
    return pending, anomalies


def _assign_vendors(pending: list[dict], sheet_name: str) -> None:
    """空白沿用上一列。整張都沒有時用工表名稱。開頭的空白列改跟下一個有名字的廠商。"""
    if not pending:
        return
    if not any(row["分包商"] for row in pending):
        for row in pending:
            row["分包商"] = sheet_name
        return
    last = ""
    for row in pending:
        if row["分包商"]:
            last = row["分包商"]
        elif last:
            row["分包商"] = last
    first_named = next(row["分包商"] for row in pending if row["分包商"])
    for row in pending:
        if not row["分包商"]:
            row["分包商"] = first_named


def _read_row(sheet_name: str, data_row: DataRow, selection: dict[str, str]) -> tuple[dict, list[dict], bool]:
    vendor = as_text(_mapped(data_row, selection, "分包商"))
    item = as_text(_mapped(data_row, selection, "工程項目"))
    unit = as_text(_mapped(data_row, selection, "單位"))
    numbers: dict[str, object] = {}
    invalid_fields: list[str] = []
    for field in NUMERIC_FIELDS:
        number, bad = to_decimal(_mapped(data_row, selection, field))
        numbers[field] = number
        if bad:
            invalid_fields.append(field)
    base = {
        "來源工作表": sheet_name,
        "來源列": data_row.excel_row,
        "分包商": vendor,
        "工程項目": item,
        "單位": unit,
        "原本期金額": numbers.get("本期金額"),
    }
    flags = [
        make_anomaly(
            base,
            "數字無法讀取",
            f"「{field}」不是數字。請改成數字後重新計算，這格不會被當成 0。",
        )
        for field in invalid_fields
    ]
    invalid = bool(invalid_fields)
    if any(is_summary_label(text) for text in (vendor, item, unit)):
        flags.append(
            make_anomaly(
                base,
                "合計、小計、營業稅或保留款列，已排除不加總",
                "請勿把這列再加進本期未稅。",
            )
        )
        return base, flags, False
    if all(numbers[field] is None for field in SIX_FIELDS) and not invalid:
        return base, flags, False
    record = {**base, **numbers}
    return record, flags, True


def _mapped(data_row: DataRow, selection: dict[str, str], field: str) -> object:
    column = selection.get(field)
    if not usable(column):
        return None
    return data_row.cells.get(column or "")
