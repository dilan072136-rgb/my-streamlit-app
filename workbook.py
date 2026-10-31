"""讀入多工作表 Excel，並找出不在第 1 列的標題。"""

from __future__ import annotations

import math
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException

from mep_app.columns import Header, prepare_headers, score_headers
from mep_app.constants import HEADER_SCAN_ROWS, MAX_ROWS, MAX_UPLOAD_BYTES


@dataclass
class DataRow:
    excel_row: int
    cells: dict[str, object]


@dataclass
class SheetTable:
    name: str
    headers: list[Header]
    rows: list[DataRow]
    header_excel_row: int
    header_height: int


@dataclass
class WorkbookData:
    sheets: list[SheetTable] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)


def read_workbook(source: bytes | bytearray | str | Path) -> WorkbookData:
    if isinstance(source, (bytes, bytearray)):
        if len(source) > MAX_UPLOAD_BYTES:
            raise ValueError("檔案超過 20MB。請拆成較小的活頁簿再上傳。")
        handle: object = BytesIO(source)
    else:
        handle = source
    try:
        frames = pd.read_excel(
            handle,
            sheet_name=None,
            header=None,
            dtype=object,
            engine="openpyxl",
        )
    except (zipfile.BadZipFile, InvalidFileException, ValueError) as exc:
        raise ValueError("這個檔案不是可讀的 xlsx。請用 Excel 另存為 .xlsx 後再上傳。") from exc
    except OSError as exc:
        raise ValueError(f"讀取 Excel 失敗：{exc}") from exc

    if not frames:
        raise ValueError("檔案裡沒有工作表。")

    total_rows = sum(len(frame.index) for frame in frames.values())
    if total_rows > MAX_ROWS:
        raise ValueError(f"資料超過 {MAX_ROWS:,} 列。請拆開後再計算，避免一次吃掉整份不相關的活頁簿。")

    names = _unique_sheet_names([str(name) for name in frames])
    book = WorkbookData()
    for name, frame in zip(names, frames.values()):
        sheet = _parse_sheet(name, frame)
        if sheet is None:
            book.skipped.append(name)
        else:
            book.sheets.append(sheet)
    if not book.sheets:
        raise ValueError("前 30 列裡找不到欄位標題。請確認表頭名稱，或把標題移到比較上面的列。")
    return book


def _unique_sheet_names(names: list[str]) -> list[str]:
    totals = Counter(names)
    seen: Counter[str] = Counter()
    unique: list[str] = []
    for name in names:
        seen[name] += 1
        unique.append(name if totals[name] == 1 else f"{name} #{seen[name]}")
    return unique


def _cell_str(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).replace("\n", "").strip()
    if text.lower() == "nan":
        return ""
    return text


def _combine(top: list[str], bottom: list[str]) -> list[str]:
    width = max(len(top), len(bottom), 0)
    labels: list[str] = []
    for index in range(width):
        above = top[index].strip() if index < len(top) else ""
        below = bottom[index].strip() if index < len(bottom) else ""
        if above and below and above != below:
            labels.append(f"{above}{below}")
        else:
            labels.append(above or below)
    return labels


def _parse_sheet(name: str, frame: pd.DataFrame) -> SheetTable | None:
    raw = frame.where(frame.notna(), None).values.tolist()
    if not raw:
        return None
    choice = _choose_header(raw)
    if choice is None:
        return None
    start, height, labels = choice
    headers = prepare_headers(labels)
    rows: list[DataRow] = []
    for offset, values in enumerate(raw[start + height :], start=start + height + 1):
        padded = list(values) + [None] * max(0, len(headers) - len(values))
        cells = {header.label: padded[index] for index, header in enumerate(headers)}
        if any(_cell_str(cells[header.label]) for header in headers if header.norm):
            rows.append(DataRow(offset, cells))
    headers = [header for header in headers if header.norm]
    if not headers:
        return None
    return SheetTable(name, headers, rows, start + 1, height)


def _choose_header(raw: list[list[object]]) -> tuple[int, int, list[str]] | None:
    """回傳 (起始列索引, 標題佔幾列, 欄名)。雙列標題只在合併後更像欄名時採用。"""
    best: tuple[int, int, int, list[str]] | None = None
    limit = min(len(raw), HEADER_SCAN_ROWS)
    for index in range(limit):
        top = [_cell_str(value) for value in raw[index]]
        single = score_headers(top)
        combined_score = -1
        combined: list[str] = []
        if index + 1 < len(raw):
            bottom = [_cell_str(value) for value in raw[index + 1]]
            combined = _combine(top, bottom)
            combined_score = score_headers(combined)
        if combined_score > single and combined_score >= 6:
            candidate = (combined_score, index, 2, combined)
        elif single >= 6:
            candidate = (single, index, 1, top)
        else:
            continue
        if best is None or candidate[0] > best[0]:
            best = candidate
    if best is None:
        return None
    _, start, height, labels = best
    return start, height, labels
