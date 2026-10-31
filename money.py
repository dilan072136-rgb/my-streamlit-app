"""工程款進位。不用內建 round，也不用二進位浮點數相乘。"""

from __future__ import annotations

import math
import numbers
import unicodedata
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Optional, Tuple

DASHES = {"-", "—", "–", "－", "N/A", "n/a", "無", "—"}


def round_half_up(value: Decimal, digits: int = 0) -> Decimal:
    """四捨五入，0.5 進位。digits 是小數位數，0 代表整數元。"""
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    quant = Decimal("1").scaleb(-digits)
    return value.quantize(quant, rounding=ROUND_HALF_UP)


def format_yuan(value: Decimal) -> str:
    """整數不帶小數，有角分才顯示兩位。"""
    quantized = round_half_up(value, 2)
    if quantized == quantized.to_integral_value():
        return f"{int(quantized):,}"
    return f"{quantized:,.2f}"


def percent_text(rate: Decimal) -> str:
    percent = (rate * Decimal("100")).normalize()
    return f"{percent}%"


def to_decimal(value: object) -> Tuple[Optional[Decimal], bool]:
    """把儲存格變成小數。

    回傳 (數字, 是否為讀不懂的內容)。空白與橫線是缺值，不是錯誤。
    """
    if value is None:
        return None, False
    if isinstance(value, Decimal):
        return value, False
    if isinstance(value, str):
        text = unicodedata.normalize("NFKC", value).strip()
        text = text.replace(",", "").replace("，", "").replace(" ", "").replace("\u3000", "")
        if not text or text in DASHES:
            return None, False
        if text.startswith("(") and text.endswith(")"):
            text = "-" + text[1:-1]
        try:
            return Decimal(text), False
        except InvalidOperation:
            return None, True
    if isinstance(value, bool):
        return None, True
    if isinstance(value, numbers.Integral):
        return Decimal(int(value)), False
    if isinstance(value, numbers.Real):
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            return None, True
        return Decimal(str(number)), False
    return None, True
