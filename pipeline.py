"""把讀檔、清洗、計價、兩期差異與報表串起來。"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from typing import Dict, List, Optional, Set
from zoneinfo import ZoneInfo

from mep_app.amounts import compute_amounts, mark_duplicates, summarize_vendors
from mep_app.clean import build_clean_data
from mep_app.columns import guess_columns, mapping_problems, sanitize_selection
from mep_app.diff import compare_periods
from mep_app.report import build_notes, make_excel
from mep_app.workbook import WorkbookData


@dataclass
class Outcome:
    detail: List[dict]
    summary: List[dict]
    anomalies: List[dict]
    diff: Optional[dict]
    notes: str
    excel: bytes
    skipped: List[str]

    @property
    def totals(self) -> Dict[str, Decimal]:
        totals: Dict[str, Decimal] = {}
        for key in ("本期未稅", "保留款", "稅額", "實際應付"):
            totals[key] = sum((row[key] for row in self.summary), Decimal("0"))
        return totals


def run(
    book: WorkbookData,
    selections: Dict[str, Dict[str, str]],
    skips: Set[str],
    job: dict,
    previous: Optional[WorkbookData] = None,
    previous_selections: Optional[Dict[str, Dict[str, str]]] = None,
    previous_skips: Optional[Set[str]] = None,
) -> Outcome:
    detail, anomalies = _assemble(book, selections, skips, job)
    summary = summarize_vendors(detail, job["retention_rate"], job["tax_rate"], job["basis"])
    diff = None
    if previous is not None:
        previous_detail, _ = _assemble(previous, previous_selections or {}, previous_skips or set(), job)
        diff = compare_periods(detail, previous_detail)
    notes = build_notes(job, summary, detail)
    excel = make_excel(job, summary, detail, anomalies, diff, notes)
    return Outcome(
        detail=detail,
        summary=summary,
        anomalies=anomalies,
        diff=diff,
        notes=notes,
        excel=excel,
        skipped=list(book.skipped),
    )


def sheet_problems(book: WorkbookData, selections: Dict[str, Dict[str, str]], skips: Set[str]) -> List[str]:
    problems: List[str] = []
    active = 0
    for sheet in book.sheets:
        if sheet.name in skips:
            continue
        active += 1
        guesses = guess_columns(sheet.headers)
        selection = sanitize_selection(selections.get(sheet.name, {}), sheet.headers, guesses)
        for message in mapping_problems(selection, guesses):
            problems.append(f"{sheet.name}：{message}")
    if active == 0:
        problems.append("所有工作表都被略過了")
    return problems


def file_token(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()[:12]


def result_fingerprint(
    token: str,
    job: dict,
    selections: dict,
    skips: Set[str],
    previous_token: str,
    previous_selections: dict,
    previous_skips: Set[str],
) -> str:
    payload = {
        "token": token,
        "project": job.get("project", ""),
        "contract_no": job.get("contract_no", ""),
        "period": job.get("period", ""),
        "retention_rate": str(job["retention_rate"]),
        "tax_rate": str(job["tax_rate"]),
        "basis": job["basis"],
        "clause_no": job.get("clause_no", ""),
        "clause_text": job.get("clause_text", ""),
        "selections": selections,
        "skips": sorted(skips),
        "previous_token": previous_token,
        "previous_selections": previous_selections,
        "previous_skips": sorted(previous_skips),
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode()).hexdigest()


def taipei_now() -> str:
    return datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y-%m-%d %H:%M")


def _assemble(book: WorkbookData, selections: Dict[str, Dict[str, str]], skips: Set[str], job: dict):
    detail: List[dict] = []
    anomalies: List[dict] = []
    for sheet in book.sheets:
        if sheet.name in skips:
            continue
        guesses = guess_columns(sheet.headers)
        selection = sanitize_selection(selections.get(sheet.name, {}), sheet.headers, guesses)
        rows, flags = build_clean_data(sheet, selection)
        detail.extend(rows)
        anomalies.extend(flags)
    computed, more = compute_amounts(detail, job["retention_rate"], job["tax_rate"], job["basis"])
    anomalies.extend(more)
    anomalies.extend(mark_duplicates(computed))
    anomalies.sort(key=_anomaly_key)
    return computed, anomalies


def _anomaly_key(row: dict) -> tuple:
    try:
        number = int(row.get("來源列") or 0)
    except (TypeError, ValueError):
        number = 0
    return (str(row.get("來源工作表") or ""), number, str(row.get("原因") or ""))
