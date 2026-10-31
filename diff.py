"""跟上期比對新增、數量變動與單價變動。金額仍用本期合計規則，這裡不重算稅。"""

from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from mep_app.money import round_half_up

TOLERANCE = Decimal("0.0001")


def compare_periods(current: list[dict], previous: list[dict]) -> dict[str, list[dict]]:
    current_groups = _group(current)
    previous_groups = _group(previous)
    added: list[dict] = []
    missing: list[dict] = []
    quantity_changes: list[dict] = []
    price_changes: list[dict] = []
    unaligned: list[dict] = []

    for key in list(current_groups) + list(previous_groups):
        if key in current_groups and len(current_groups[key]) > 1:
            unaligned.append(_unaligned("本期", key, current_groups[key]))
        if key in previous_groups and len(previous_groups[key]) > 1:
            unaligned.append(_unaligned("上期", key, previous_groups[key]))

    blocked = {(row["分包商"], row["工程項目"], row["單位"]) for row in unaligned}
    for key, rows in current_groups.items():
        if key in blocked or len(rows) != 1:
            continue
        if key not in previous_groups or len(previous_groups[key]) != 1:
            if key not in previous_groups:
                added.append(_snapshot(rows[0]))
            continue
        now = rows[0]
        before = previous_groups[key][0]
        if not _same(now.get("本期數量"), before.get("本期數量")) or not _same(
            now.get("累計數量"), before.get("累計數量")
        ):
            quantity_changes.append(
                {
                    "分包商": key[0],
                    "工程項目": key[1],
                    "單位": key[2],
                    "上期本期數量": before.get("本期數量"),
                    "本期數量": now.get("本期數量"),
                    "上期累計數量": before.get("累計數量"),
                    "本期累計數量": now.get("累計數量"),
                }
            )
        if not _same(now.get("單價"), before.get("單價")):
            price_changes.append(
                {
                    "分包商": key[0],
                    "工程項目": key[1],
                    "單位": key[2],
                    "上期單價": before.get("單價"),
                    "本期單價": now.get("單價"),
                }
            )
    for key, rows in previous_groups.items():
        if key in blocked or len(rows) != 1:
            continue
        if key not in current_groups:
            missing.append(_snapshot(rows[0]))
    return {
        "added": added,
        "quantity_changes": quantity_changes,
        "price_changes": price_changes,
        "missing": missing,
        "unaligned": _dedupe(unaligned),
    }


def _group(records: list[dict]) -> dict[tuple[str, str, str], list[dict]]:
    groups: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
    for record in records:
        item = record.get("工程項目") or ""
        if not item:
            continue
        key = (record.get("分包商") or "", item, record.get("單位") or "")
        groups[key].append(record)
    return groups


def _same(left: Decimal | None, right: Decimal | None) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return abs(Decimal(left) - Decimal(right)) <= TOLERANCE


def _snapshot(record: dict) -> dict:
    return {
        "分包商": record.get("分包商") or "",
        "工程項目": record.get("工程項目") or "",
        "單位": record.get("單位") or "",
        "本期數量": record.get("本期數量"),
        "本期金額": record.get("本期金額"),
        "單價": record.get("單價"),
    }


def _unaligned(side: str, key: tuple[str, str, str], rows: list[dict]) -> dict:
    amount = sum((row.get("本期金額") or Decimal("0") for row in rows), Decimal("0"))
    return {
        "出現在": side,
        "分包商": key[0],
        "工程項目": key[1],
        "單位": key[2],
        "筆數": len(rows),
        "本期金額合計": round_half_up(amount, 2),
    }


def _dedupe(rows: list[dict]) -> list[dict]:
    seen: set[tuple] = set()
    unique: list[dict] = []
    for row in rows:
        identity = (row["出現在"], row["分包商"], row["工程項目"], row["單位"])
        if identity in seen:
            continue
        seen.add(identity)
        unique.append(row)
    return unique
