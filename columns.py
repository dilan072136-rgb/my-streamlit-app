"""表頭對到標準欄。沒有把握時不擅自採用。"""

from __future__ import annotations

import unicodedata
from collections import Counter
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from mep_app.constants import ALIASES, CANONICAL, NEGATIVE, PLACEHOLDER, UNUSED


@dataclass(frozen=True)
class Header:
    label: str
    norm: str


@dataclass(frozen=True)
class Guess:
    field: str
    column: Optional[str]
    alternatives: Tuple[str, ...]
    status: str  # unique, ambiguous, missing


def normalize(text: object) -> str:
    """全形轉半形，去掉空白與括號，供欄名比對。"""
    raw = unicodedata.normalize("NFKC", "" if text is None else str(text))
    drop = " \t\r\n\u3000()（）[]【】{}「」『』:：,，.．、"
    return "".join(ch for ch in raw if ch not in drop).strip().casefold()


def prepare_headers(raw_labels: List[object]) -> List[Header]:
    norms = [normalize(label) for label in raw_labels]
    totals = Counter(norm for norm in norms if norm)
    seen: Counter[str] = Counter()
    headers: List[Header] = []
    for index, (raw, norm) in enumerate(zip(raw_labels, norms), start=1):
        if not norm:
            headers.append(Header(f"空白欄{index}", ""))
            continue
        seen[norm] += 1
        text = str(raw).strip()
        label = text if totals[norm] == 1 else f"{text} #{seen[norm]}"
        headers.append(Header(label, norm))
    return headers


def guess_column(headers: List[Header], field: str) -> Guess:
    """完全相同優先。多欄都像、或完全對不到時，不回傳採用欄。"""
    exact: List[str] = []
    for alias in ALIASES[field]:
        alias_norm = normalize(alias)
        for header in headers:
            if header.norm and header.norm == alias_norm and header.label not in exact:
                exact.append(header.label)
    if len(exact) == 1:
        return Guess(field, exact[0], (), "unique")
    if len(exact) > 1:
        return Guess(field, None, tuple(exact), "ambiguous")

    contained: List[str] = []
    blocked = tuple(normalize(word) for word in NEGATIVE.get(field, ()))
    for alias in ALIASES[field]:
        alias_norm = normalize(alias)
        if len(alias_norm) < 3:
            continue
        for header in headers:
            if not header.norm or alias_norm not in header.norm:
                continue
            if any(word and word in header.norm for word in blocked):
                continue
            if header.label not in contained:
                contained.append(header.label)
    if len(contained) == 1:
        return Guess(field, contained[0], (), "unique")
    if len(contained) > 1:
        return Guess(field, None, tuple(contained), "ambiguous")
    return Guess(field, None, (), "missing")


def guess_columns(headers: List[Header]) -> List[Guess]:
    return [guess_column(headers, field) for field in CANONICAL]


def score_headers(raw_labels: List[object]) -> int:
    """標題列要比資料列更像欄名。唯一對上的欄權重高於多個可能。"""
    guesses = guess_columns(prepare_headers(raw_labels))
    score = 0
    for guess in guesses:
        if guess.status == "unique":
            score += 3
        elif guess.status == "ambiguous":
            score += 1
    return score


def initial_selection(guesses: List[Guess]) -> Dict[str, str]:
    selection: Dict[str, str] = {}
    for guess in guesses:
        if guess.status == "unique" and guess.column:
            selection[guess.field] = guess.column
        elif guess.status == "ambiguous":
            selection[guess.field] = PLACEHOLDER
        else:
            selection[guess.field] = UNUSED
    return selection


def usable(value: Optional[str]) -> bool:
    return value not in (None, "", PLACEHOLDER, UNUSED)


def sanitize_selection(
    selection: Dict[str, str],
    headers: List[Header],
    guesses: List[Guess],
) -> Dict[str, str]:
    labels = {header.label for header in headers if header.norm}
    base = initial_selection(guesses)
    clean: Dict[str, str] = {}
    for field in CANONICAL:
        chosen = selection.get(field, base[field])
        if chosen in labels or chosen in (PLACEHOLDER, UNUSED):
            clean[field] = chosen
        else:
            clean[field] = base[field]
    return clean


def mapping_problems(selection: Dict[str, str], guesses: List[Guess]) -> List[str]:
    """多個欄都像、或缺了算得出錢的欄時，擋下計算。"""
    problems: List[str] = []
    for guess in guesses:
        chosen = selection.get(guess.field)
        if guess.status == "ambiguous" and chosen in (None, "", PLACEHOLDER):
            options = "、".join(guess.alternatives)
            problems.append(f"{guess.field}有多個可能欄位（{options}），請指定")
    if not usable(selection.get("工程項目")):
        problems.append("請指定工程項目")
    has_amount = usable(selection.get("本期金額"))
    has_qty_price = usable(selection.get("本期數量")) and usable(selection.get("單價"))
    if not has_amount and not has_qty_price:
        problems.append("請指定本期金額，或同時指定本期數量與單價")
    return problems


def duplicate_columns(selection: Dict[str, str]) -> List[str]:
    chosen = [value for value in selection.values() if usable(value)]
    return sorted({value for value in chosen if chosen.count(value) > 1})
