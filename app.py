"""多分包商估驗彙整。用 streamlit run app.py 開啟。"""

from __future__ import annotations

from decimal import Decimal

import pandas as pd
import streamlit as st

from mep_app.columns import (
    duplicate_columns,
    guess_columns,
    initial_selection,
    mapping_problems,
    sanitize_selection,
)
from mep_app.constants import CANONICAL, PLACEHOLDER, TAX_CHOICES, TAX_GROSS, TAX_NET, UNUSED
from mep_app.money import format_yuan, round_half_up
from mep_app.pipeline import file_token, result_fingerprint, run, sheet_problems, taipei_now
from mep_app.sample import PREVIOUS_PATH, SAMPLE_PATH, write_examples
from mep_app.workbook import WorkbookData, read_workbook

MONEY_WORDS = ("金額", "單價", "保留款", "稅額", "應付", "未稅", "原金額")
QTY_WORDS = ("數量", "契約數量")


def main() -> None:
    st.set_page_config(page_title="多分包商估驗彙整", layout="wide")
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.4rem; max-width: 1180px; }
        div[data-testid="stMetric"] {
            background: #fffdf9;
            border: 1px solid #e7e0d6;
            border-radius: 10px;
            padding: 0.7rem 0.85rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("多分包商估驗彙整")
    st.caption("把各家不整齊的估驗表收成同一份可簽核的數字。原始儲存格不會被改寫，異常只列出來讓人確認。")
    try:
        _render()
    except Exception as exc:  # 最後一道，避免整頁白屏
        st.error("畫面沒有完成這次計算，已填的數字沒有被改寫。")
        st.exception(exc)


def _render() -> None:
    job = _sidebar()
    current_bytes, current_name, previous_bytes = _resolve_files()
    if current_bytes is None:
        _empty_state()
        return

    try:
        book = _load_book(current_bytes)
    except ValueError as exc:
        st.error(str(exc))
        return

    previous_book = None
    if previous_bytes:
        try:
            previous_book = _load_book(previous_bytes)
        except ValueError as exc:
            st.error(f"上期檔讀取失敗，這次只算本期：{exc}")

    if book.skipped:
        st.warning("這些工作表在前 30 列找不到欄位標題，已略過：" + "、".join(book.skipped))

    st.caption(f"本期檔案：{current_name}")
    token = file_token(current_bytes)
    _drop_stale_result(token)
    selections, skips = _mapping_block("本期欄位對應", token, book)
    previous_selections: dict = {}
    previous_skips: set[str] = set()
    previous_token = ""
    if previous_book is not None:
        previous_token = file_token(previous_bytes or b"")
        if previous_book.skipped:
            st.caption("上期略過：" + "、".join(previous_book.skipped))
        previous_selections, previous_skips = _mapping_block("上期欄位對應", "prev" + previous_token, previous_book)

    problems = sheet_problems(book, selections, skips)
    if previous_book is not None:
        problems += [f"上期／{item}" for item in sheet_problems(previous_book, previous_selections, previous_skips)]
    if problems:
        st.error("還沒能計算：\n\n" + "\n\n".join(f"- {item}" for item in problems))

    clicked = st.button("計算", type="primary", disabled=bool(problems))
    if st.session_state.get("pending_calc") and not problems:
        clicked = True
        st.session_state["pending_calc"] = False

    fingerprint = result_fingerprint(
        token,
        job,
        selections,
        skips,
        previous_token,
        previous_selections,
        previous_skips,
    )
    if clicked:
        with st.spinner("正在彙整本期金額"):
            st.session_state["result"] = run(
                book,
                selections,
                skips,
                job,
                previous_book,
                previous_selections,
                previous_skips,
            )
            st.session_state["result_fp"] = fingerprint

    outcome = st.session_state.get("result")
    if outcome is None:
        st.info("欄位確認後按計算。總表的保留款與稅額會在分包商合計後進到整數元。")
        return
    stale = st.session_state.get("result_fp") != fingerprint
    if stale:
        st.warning("工程名稱、合約、期別、稅率、欄位或檔案已變更。下面仍是上一次的結果，請重新計算後再下載。")
    _show_outcome(outcome, job, stale)


def _sidebar() -> dict:
    with st.sidebar:
        st.header("本期條件")
        project = st.text_input("工程名稱", value="未命名工程")
        contract_no = st.text_input("合約編號", value="")
        period = st.text_input("期別", value="第1期")
        retention_pct = st.number_input("保留款比率（%）", min_value=0.0, max_value=100.0, value=5.0, step=0.1)
        tax_pct = st.number_input("營業稅率（%）", min_value=0.0, max_value=100.0, value=5.0, step=0.1)
        basis = st.radio("稅額計算基礎", TAX_CHOICES, index=0, help="兩種稅基的應付不能混著對帳。")
        clause_no = st.text_input("合約條號", value="", help="例如第 12 條。沒有就留空。")
        clause_text = st.text_area("條文摘句", value="", height=88, help="摘句只作依據，不取代合計。")
        st.divider()
        st.file_uploader("本期估驗 Excel", type=["xlsx", "xlsm"], key="current_upload")
        st.file_uploader("上期估驗 Excel（選填）", type=["xlsx", "xlsm"], key="previous_upload")
        if st.button("載入內建範例", width="stretch"):
            current, previous = _example_bytes()
            st.session_state["use_sample"] = True
            st.session_state["sample_bytes"] = current
            st.session_state["previous_sample_bytes"] = previous
            st.session_state["pending_calc"] = True
            st.session_state.pop("result", None)
            st.rerun()
        if st.session_state.get("use_sample") and st.button("清除範例", width="stretch"):
            for key in ("use_sample", "sample_bytes", "previous_sample_bytes", "result", "result_fp", "pending_calc"):
                st.session_state.pop(key, None)
            st.rerun()
        if basis == TAX_NET:
            st.caption("先扣保留款再乘稅率。10,000 元、保留 5%、稅 5% 的應付是 9,975 元。")
        else:
            st.caption("以本期未稅直接乘稅率。同一筆的應付是 10,000 元。")
    return {
        "project": project.strip() or "未命名工程",
        "contract_no": contract_no.strip(),
        "period": period.strip() or "本期",
        "retention_rate": Decimal(f"{retention_pct:.4f}") / Decimal("100"),
        "tax_rate": Decimal(f"{tax_pct:.4f}") / Decimal("100"),
        "basis": basis,
        "clause_no": clause_no.strip(),
        "clause_text": clause_text.strip(),
        "generated_at": taipei_now() + " 台北",
    }


def _resolve_files() -> tuple[bytes | None, str, bytes | None]:
    current_upload = st.session_state.get("current_upload")
    previous_upload = st.session_state.get("previous_upload")
    if current_upload is not None:
        current_bytes = current_upload.getvalue()
        current_name = current_upload.name
    elif st.session_state.get("use_sample"):
        current_bytes = st.session_state.get("sample_bytes")
        current_name = "sample.xlsx"
    else:
        return None, "", None
    if previous_upload is not None:
        previous_bytes = previous_upload.getvalue()
    elif st.session_state.get("use_sample"):
        previous_bytes = st.session_state.get("previous_sample_bytes")
    else:
        previous_bytes = None
    return current_bytes, current_name, previous_bytes


def _empty_state() -> None:
    st.subheader("先給一份本期估驗表")
    st.write("側欄可以上傳 xlsx，或先載入內建範例。範例含雙列標題、只填在第一列的分包商、小計列，以及說明牌、累計土方。")
    left, right = st.columns(2)
    left.markdown("**計算會守住的規則**\n\n- 累計等於前期加本期；只有空白才回推\n- 有本期金額就能算；沒有就用本期數量乘單價\n- 保留款與稅額在分包商合計後進到整數元")
    right.markdown("**對帳只看三個金額**\n\n- 本期未稅\n- 本期保留款\n- 實際應付\n\n數量單位不同，不能加總當結論。")


def _mapping_block(title: str, token: str, book: WorkbookData):
    st.subheader(title)
    prepared = []
    forced = False
    for sheet in book.sheets:
        guesses = guess_columns(sheet.headers)
        suggested = initial_selection(guesses)
        labels = [header.label for header in sheet.headers if header.norm]
        options = [PLACEHOLDER, UNUSED, *labels]
        stored = _stored_selection(token, sheet.name, suggested, options)
        problems = mapping_problems(stored, guesses)
        duplicates = duplicate_columns(stored)
        if problems:
            forced = True
        prepared.append((sheet, guesses, suggested, options, stored, problems, duplicates))

    if forced:
        show_editors = True
        st.caption("有欄位對到一個以上，或還缺算得出錢的欄。指定完成後才能計算。")
    else:
        show_editors = st.checkbox("調整欄位對應", key=f"edit:{token}:{title}")
        st.caption(f"{len(prepared)} 個工作表已有單一建議，可直接計算。")

    selections: dict[str, dict[str, str]] = {}
    skips: set[str] = set()
    for sheet, guesses, suggested, options, stored, problems, duplicates in prepared:
        skip_key = f"skip:{token}:{sheet.name}"
        if show_editors:
            with st.expander(f"{sheet.name}　{len(sheet.rows)} 列資料", expanded=bool(problems or duplicates)):
                if skip_key not in st.session_state:
                    st.session_state[skip_key] = False
                skipped = st.checkbox("略過此工作表", key=skip_key)
                selection = _render_selectors(token, sheet.name, guesses, suggested, options)
                selection = sanitize_selection(selection, sheet.headers, guesses)
                if problems:
                    st.caption("要先指定：" + "；".join(mapping_problems(selection, guesses)))
                elif duplicates:
                    st.caption("同一欄被選了兩次：" + "、".join(duplicate_columns(selection)) + "。計算仍會照你選的欄讀取。")
        else:
            skipped = bool(st.session_state.get(skip_key))
            selection = sanitize_selection(stored, sheet.headers, guesses)
        if skipped:
            skips.add(sheet.name)
        selections[sheet.name] = selection
    return selections, skips


def _stored_selection(token: str, sheet_name: str, suggested: dict[str, str], options: list[str]) -> dict[str, str]:
    allowed = set(options)
    selection = {}
    for field, default in suggested.items():
        chosen = st.session_state.get(f"map:{token}:{sheet_name}:{field}", default)
        selection[field] = chosen if chosen in allowed else default
    return selection


def _render_selectors(token: str, sheet_name: str, guesses, suggested: dict[str, str], options: list[str]) -> dict[str, str]:
    selection: dict[str, str] = {}
    columns = st.columns(2)
    for index, field in enumerate(CANONICAL):
        guess = next(item for item in guesses if item.field == field)
        key = f"map:{token}:{sheet_name}:{field}"
        default = suggested[field] if suggested[field] in options else PLACEHOLDER
        if key not in st.session_state or st.session_state[key] not in options:
            st.session_state[key] = default
        with columns[index % 2]:
            selection[field] = st.selectbox(field, options, key=key, help=_guess_help(guess))
    return selection


def _guess_help(guess) -> str:
    if guess.status == "unique":
        return "建議欄。只有一個欄名對得上，仍可以改。"
    if guess.status == "ambiguous":
        return "其他可能：" + "、".join(guess.alternatives) + "。請指定後才計算。"
    return "對不到欄名。請指定，或選不使用。"


def _show_outcome(outcome, job: dict, stale: bool) -> None:
    totals = outcome.totals
    st.subheader("本期實際應付")
    top_left, top_right = st.columns(2)
    bottom_left, bottom_right = st.columns(2)
    top_left.metric("本期未稅", format_yuan(totals["本期未稅"]))
    top_right.metric("保留款", format_yuan(totals["保留款"]))
    bottom_left.metric("稅額", format_yuan(totals["稅額"]))
    bottom_right.metric("實際應付", format_yuan(totals["實際應付"]))
    basis_text = "未稅直接乘稅率" if job["basis"] == TAX_GROSS else "先扣保留款再乘稅率"
    st.caption(f"稅基：{job['basis']}（{basis_text}）。對帳用這三個金額：本期未稅、保留款、實際應付。數量合計不是結論。")
    _gap_caption(outcome)
    filename = f"{_safe_name(job['project'])}_{_safe_name(job['period'])}_估驗彙整.xlsx"
    st.download_button(
        "下載估驗報表",
        data=outcome.excel,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=stale,
        type="primary",
    )
    tabs = st.tabs(["估驗總表", "明細", "異常清單", "與上期差異", "計算說明"])
    with tabs[0]:
        if outcome.summary:
            st.dataframe(_display(outcome.summary), width="stretch", hide_index=True)
        else:
            st.info("沒有可計價的工項。")
        st.markdown("簽核看總表裡的角色說明：承辦對閉合與進位，工地主任對異常，專案經理與業主對實際應付。")
    with tabs[1]:
        st.caption("「補值」是原本空白、由另外兩欄或數量乘單價補上的欄。已填的數字不會被改。參考保留款只供核對，正式數字在分包商合計。")
        st.dataframe(_display(outcome.detail), width="stretch", hide_index=True)
    with tabs[2]:
        if outcome.anomalies:
            st.dataframe(_display(outcome.anomalies), width="stretch", hide_index=True)
        else:
            st.success("這次沒有要確認的異常。")
    with tabs[3]:
        _show_diff(outcome.diff)
    with tabs[4]:
        st.markdown(outcome.notes)


def _gap_caption(outcome) -> None:
    official = outcome.totals["保留款"]
    reference = sum(
        (
            row["參考保留款"]
            for row in outcome.detail
            if row.get("納入合計") and row.get("參考保留款") is not None
        ),
        Decimal("0"),
    )
    gap = reference - official
    if gap == 0:
        st.caption("明細參考保留款加總與總表整數保留款相同。")
        return
    direction = "多" if gap > 0 else "少"
    st.info(
        f"明細參考保留款加總是 {format_yuan(reference)} 元，比總表的 {format_yuan(official)} 元{direction} {format_yuan(abs(gap))} 元。"
        "總表採用先加總再進到整數元。"
    )


def _show_diff(diff: dict | None) -> None:
    if not diff:
        st.info("上傳上期估驗表後，這裡會列出新增工項、數量變動、單價變動，以及上期有但本期沒有的項目。差異表不重算稅。")
        return
    sections = [
        ("新增工項", diff["added"]),
        ("數量變動", diff["quantity_changes"]),
        ("單價變動", diff["price_changes"]),
        ("本期未列", diff["missing"]),
        ("無法對齊", diff["unaligned"]),
    ]
    for title, rows in sections:
        st.markdown(f"**{title}**　{len(rows)} 筆")
        if rows:
            st.dataframe(_display(rows), width="stretch", hide_index=True)
        else:
            st.caption("無")


def _display(rows: list[dict]) -> pd.DataFrame:
    shown = []
    for row in rows:
        item = {}
        for key, value in row.items():
            if str(key).startswith("_"):
                continue
            item[key] = _show_value(key, value)
        shown.append(item)
    return pd.DataFrame(shown)


def _show_value(key: str, value):
    if value is None:
        return ""
    if key == "納入合計":
        return "是" if value else "否"
    if isinstance(value, Decimal):
        if "完成率" in key:
            return f"{(value * Decimal('100')):.2f}%"
        if any(word in key for word in MONEY_WORDS):
            return format_yuan(value)
        if any(word in key for word in QTY_WORDS):
            quantized = round_half_up(value, 4)
            text = f"{quantized:,.4f}".rstrip("0").rstrip(".")
            return text or "0"
        return format_yuan(value)
    return value


def _safe_name(text: str) -> str:
    cleaned = "".join(char for char in text if char not in '\\/:*?"<>|').strip()
    return cleaned or "估驗"


def _drop_stale_result(token: str) -> None:
    if st.session_state.get("file_token") == token:
        return
    st.session_state["file_token"] = token
    st.session_state.pop("result", None)
    st.session_state.pop("result_fp", None)


def _example_bytes() -> tuple[bytes, bytes]:
    if not SAMPLE_PATH.exists() or not PREVIOUS_PATH.exists():
        write_examples()
    return SAMPLE_PATH.read_bytes(), PREVIOUS_PATH.read_bytes()


@st.cache_data(show_spinner=False)
def _load_book(content: bytes) -> WorkbookData:
    return read_workbook(content)


main()
