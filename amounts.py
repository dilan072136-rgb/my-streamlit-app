"""補空白的數量與金額，並在分包商合計後把保留款、稅額進到整數元。"""

from __future__ import annotations

from collections import Counter
from decimal import Decimal

from mep_app.clean import make_anomaly
from mep_app.constants import TAX_NET
from mep_app.money import round_half_up

QTY_TOLERANCE = Decimal("0.0001")
MONEY_TOLERANCE = Decimal("0.01")


def compute_amounts(
    records: list[dict],
    retention_rate: Decimal,
    tax_rate: Decimal,
    basis: str,
) -> tuple[list[dict], list[dict]]:
    """已填的數字保持原樣。空白才用另外兩欄回推，或用數量乘單價補上。"""
    computed: list[dict] = []
    anomalies: list[dict] = []
    for source in records:
        record = dict(source)
        flags = _fill_record(record, retention_rate, tax_rate, basis)
        computed.append(record)
        anomalies.extend(flags)
    return computed, anomalies


def summarize_vendors(
    records: list[dict],
    retention_rate: Decimal,
    tax_rate: Decimal,
    basis: str,
) -> list[dict]:
    """先加總本期未稅，再把保留款與稅額四捨五入到整數元。應付不再多進一次位。"""
    groups: dict[str, list[dict]] = {}
    order: list[str] = []
    for record in records:
        if not record.get("納入合計"):
            continue
        vendor = record.get("分包商") or "（未填分包商）"
        if vendor not in groups:
            groups[vendor] = []
            order.append(vendor)
        groups[vendor].append(record)

    summary: list[dict] = []
    for vendor in order:
        rows = groups[vendor]
        total = sum((row["本期金額"] for row in rows), Decimal("0"))
        retention = round_half_up(total * retention_rate, 0)
        if basis == TAX_NET:
            tax = round_half_up((total - retention) * tax_rate, 0)
        else:
            tax = round_half_up(total * tax_rate, 0)
        summary.append(
            {
                "分包商": vendor,
                "項次數": len(rows),
                "本期未稅": total,
                "保留款": retention,
                "稅額": tax,
                "實際應付": total + tax - retention,
            }
        )
    return summary


def mark_duplicates(records: list[dict]) -> list[dict]:
    counts: Counter[tuple[str, str, str]] = Counter()
    keys: list[tuple[str, str, str]] = []
    for record in records:
        key = (record.get("分包商") or "", record.get("工程項目") or "", record.get("單位") or "")
        keys.append(key)
        if key[1]:
            counts[key] += 1
    anomalies: list[dict] = []
    for record, key in zip(records, keys):
        if key[1] and counts[key] > 1:
            anomalies.append(make_anomaly(record, "同一分包商、工程項目與單位重複", "請確認是否重複計價。"))
    return anomalies


def _fill_record(record: dict, retention_rate: Decimal, tax_rate: Decimal, basis: str) -> list[dict]:
    anomalies: list[dict] = []
    original_qty = (record.get("前期累計數量"), record.get("本期數量"), record.get("累計數量"))
    original_amt = (record.get("前期累計金額"), record.get("本期金額"), record.get("累計金額"))
    original_amount = record.get("本期金額")
    original_qty_now = record.get("本期數量")
    original_price = record.get("單價")

    prev_qty, this_qty, cum_qty, qty_mismatch = _close_triplet(original_qty, QTY_TOLERANCE, _q_qty)
    prev_amt, this_amt, cum_amt, amt_mismatch = _close_triplet(original_amt, MONEY_TOLERANCE, _q_money)
    price = record.get("單價")
    if prev_amt is None and prev_qty is not None and price is not None:
        prev_amt = _q_money(prev_qty * price)
    if this_amt is None and this_qty is not None and price is not None:
        this_amt = _q_money(this_qty * price)
    prev_amt, this_amt, cum_amt, _ = _close_triplet((prev_amt, this_amt, cum_amt), MONEY_TOLERANCE, _q_money)
    record["補值"] = _filled_fields(
        original_qty,
        original_amt,
        (prev_qty, this_qty, cum_qty),
        (prev_amt, this_amt, cum_amt),
    )

    record["前期累計數量"] = prev_qty
    record["本期數量"] = this_qty
    record["累計數量"] = cum_qty
    record["前期累計金額"] = prev_amt
    record["本期金額"] = this_amt
    record["累計金額"] = cum_amt

    contract = record.get("契約數量")
    if contract is not None and contract > 0 and cum_qty is not None:
        record["累計完成率"] = cum_qty / contract
        over = _q_qty(cum_qty) > _q_qty(contract)
    else:
        record["累計完成率"] = None
        over = False

    if this_amt is not None:
        ref_retention = round_half_up(this_amt * retention_rate, 2)
        if basis == TAX_NET:
            ref_tax = round_half_up((this_amt - ref_retention) * tax_rate, 2)
        else:
            ref_tax = round_half_up(this_amt * tax_rate, 2)
        record["參考保留款"] = ref_retention
        record["參考稅額"] = ref_tax
        record["納入合計"] = True
    else:
        record["參考保留款"] = None
        record["參考稅額"] = None
        record["納入合計"] = False
        anomalies.append(
            make_anomaly(
                record,
                "缺少本期金額",
                "這列沒有本期金額，也無法用本期數量乘單價補上，所以沒有進合計。",
            )
        )

    if qty_mismatch:
        anomalies.append(
            make_anomaly(
                record,
                "累計數量 ≠ 前期累計＋本期",
                "請退回核對。三欄都有數字時程式不會改寫。",
            )
        )
    if amt_mismatch:
        anomalies.append(
            make_anomaly(
                record,
                "累計金額 ≠ 前期累計＋本期",
                "請退回核對。三欄都有數字時程式不會改寫。",
            )
        )
    if (this_amt is not None and this_amt < 0) or (this_qty is not None and this_qty < 0):
        anomalies.append(
            make_anomaly(
                record,
                "本期金額或數量為負",
                "請確認是否為退項或扣款。這列仍保留，也計入合計。",
            )
        )
    if over:
        anomalies.append(
            make_anomaly(
                record,
                "累計完成率超過 100%",
                "請確認是否超做或已變更設計。完成率只警示，不會把數量改小。",
            )
        )
    if original_qty_now is not None and original_price is not None and original_amount is not None:
        expected = _q_money(original_qty_now * original_price)
        if abs(expected - original_amount) > MONEY_TOLERANCE:
            anomalies.append(
                make_anomaly(
                    record,
                    "本期金額 ≠ 本期數量 × 單價",
                    "請核對數量、單價與本期金額。已填的數字保持原樣。",
                )
            )
    elif original_amount is None and this_amt is not None and this_qty is not None and price is not None:
        expected = _q_money(this_qty * price)
        if abs(expected - this_amt) > MONEY_TOLERANCE:
            anomalies.append(
                make_anomaly(
                    record,
                    "回推的本期金額 ≠ 本期數量 × 單價",
                    "本期金額是用另外兩欄補上的，和數量乘單價不同。程式沒有改這兩個結果，請指定要以哪一個為準。",
                )
            )
    return anomalies


def _filled_fields(original_qty, original_amt, qty, amt) -> str:
    labels = ("前期累計數量", "本期數量", "累計數量", "前期累計金額", "本期金額", "累計金額")
    originals = (*original_qty, *original_amt)
    current = (*qty, *amt)
    filled = [label for label, before, after in zip(labels, originals, current) if before is None and after is not None]
    return "、".join(filled)


def _close_triplet(values: tuple, tolerance: Decimal, quant) -> tuple:
    prev, current, cumulative = values
    present = [prev is not None, current is not None, cumulative is not None]
    if present == [False, True, True]:
        return quant(cumulative - current), current, cumulative, False
    if present == [True, False, True]:
        return prev, quant(cumulative - prev), cumulative, False
    if present == [True, True, False]:
        return prev, current, quant(prev + current), False
    if all(present) and abs(prev + current - cumulative) > tolerance:
        return prev, current, cumulative, True
    return prev, current, cumulative, False


def _q_qty(value: Decimal) -> Decimal:
    return round_half_up(value, 4)


def _q_money(value: Decimal) -> Decimal:
    return round_half_up(value, 2)
