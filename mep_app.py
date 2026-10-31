# -*- coding: utf-8 -*-
"""
機電工程專案管理暨估驗計價自動化系統（修正整合版）
Python 3.10+
安裝：  pip install streamlit pandas plotly openpyxl
執行：  streamlit run mep_app.py
"""
import inspect
import re
from datetime import date, datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# 1. 欄位定義與別名
# =========================================================
SCHEMAS = {
    "schedule": {
        "工作項目": ["工作項目", "工程項目", "施工項目", "項目", "task", "activity"],
        "分包商": ["分包商", "施工廠商", "承包商", "廠商", "subcontractor"],
        "預計開始": ["預計開始", "預定開始", "計畫開始", "預計開工", "start_plan", "planned_start"],
        "預計結束": ["預計結束", "預定結束", "計畫完成", "預計完工", "end_plan", "planned_end"],
        "實際開始": ["實際開始", "實際開工", "開始日期", "start_actual", "actual_start"],
        "實際結束": ["實際結束", "實際完工", "完成日期", "end_actual", "actual_end"],
        "完成率": ["完成率", "施工進度", "進度百分比", "progress", "completion"],
    },
    "manpower": {
        "日期": ["日期", "施工日期", "出工日期", "工作日期", "date", "work_date"],
        "分包商": ["分包商", "施工廠商", "承包商", "廠商", "subcontractor"],
        "工作項目": ["工作項目", "工程項目", "施工項目", "項目", "task", "activity"],
        "出工人數": ["出工人數", "實際人數", "施工人數", "人數", "workers", "actual_workers"],
        "工作時數": ["工作時數", "出工時數", "工時", "hours", "work_hours"],
        "實際人天": ["實際人天", "人天", "工數", "man_days", "actual_man_days"],
    },
    "estimate": {
        "工程項目": ["工程項目", "項目名稱", "施工項目", "品名", "項目", "item"],
        "規格": ["規格", "尺寸", "spec", "specification"],
        "單位": ["單位", "unit"],
        "分包商": ["分包商", "施工廠商", "承包商", "廠商", "subcontractor"],
        "前期累計數量": ["前期累計數量", "上期累計數量", "前期數量", "previous_qty"],
        "本期數量": ["本期數量", "本期估驗數量", "本期完成數量", "數量", "quantity", "qty"],
        "單價": ["單價", "unit_price", "price"],
        "前期累計金額": ["前期累計金額", "上期累計金額", "前期金額", "previous_amount"],
        "本期金額": ["本期金額", "本期估驗金額", "金額", "amount", "current_amount"],
    },
}
DEFAULT_COLUMNS = {k: list(v.keys()) for k, v in SCHEMAS.items()}
ESTIMATE_DERIVED = ["累計數量", "累計金額"]

DISTINCT = {
    "schedule": ["預計開始", "預計結束", "實際開始", "實際結束", "完成率"],
    "manpower": ["日期", "出工人數", "工作時數", "實際人天"],
    "estimate": ["單價", "本期數量", "本期金額", "前期累計數量", "前期累計金額"],
}
KIND_LABEL = {"schedule": "施工進度資料", "manpower": "每日出工資料", "estimate": "估驗計價資料"}
DATE_COLS = ["預計開始", "預計結束", "實際開始", "實際結束", "日期"]
EMPTY_DTYPES = {
    "預計開始": "datetime64[ns]", "預計結束": "datetime64[ns]",
    "實際開始": "datetime64[ns]", "實際結束": "datetime64[ns]", "日期": "datetime64[ns]",
    "完成率": "float64", "出工人數": "float64", "工作時數": "float64", "實際人天": "float64",
    "前期累計數量": "float64", "本期數量": "float64", "單價": "float64",
    "前期累計金額": "float64", "本期金額": "float64", "累計數量": "float64", "累計金額": "float64",
}


def empty_frame(kind):
    cols = DEFAULT_COLUMNS[kind] + (ESTIMATE_DERIVED if kind == "estimate" else [])
    return pd.DataFrame({c: pd.Series(dtype=EMPTY_DTYPES.get(c, "object")) for c in cols})


# =========================================================
# 2. 共用資料處理
# =========================================================
def normalize_text(value):
    text = str(value).strip().lower()
    return re.sub(r"[\s_\-（）()【】\[\]]+", "", text)


def unique_names(names):
    seen, out = {}, []
    for i, n in enumerate(names, start=1):
        n = "" if n is None or (not isinstance(n, str) and pd.isna(n)) else str(n).strip()
        n = n or f"欄位{i}"
        seen[n] = seen.get(n, 0) + 1
        out.append(n if seen[n] == 1 else f"{n}_{seen[n]}")
    return out


def round_half_up(x):
    return np.sign(x) * np.floor(np.abs(x) + 0.5)


def safe_filename(s):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", str(s or "").strip()) or "工程"


def parse_date(value):
    if value is None or (not isinstance(value, str) and pd.isna(value)) or str(value).strip() == "":
        return pd.NaT
    if isinstance(value, (pd.Timestamp, date, datetime)):
        return pd.to_datetime(value, errors="coerce")
    if isinstance(value, (int, float, np.number)) and 20000 < float(value) < 80000:
        return pd.to_datetime(float(value), unit="D", origin="1899-12-30", errors="coerce")

    text = str(value).strip()
    m = re.match(r"^(\d{2,3})[/\-.](\d{1,2})[/\-.](\d{1,2})$", text)
    if m:
        y = int(m.group(1))
        y = y + 1911 if y < 1911 else y
        return pd.to_datetime(f"{y}-{int(m.group(2)):02d}-{int(m.group(3)):02d}", errors="coerce")
    m = re.match(r"^(\d{2,3})年(\d{1,2})月(\d{1,2})日?$", text)
    if m:
        y = int(m.group(1)) + 1911
        return pd.to_datetime(f"{y}-{int(m.group(2)):02d}-{int(m.group(3)):02d}", errors="coerce")
    return pd.to_datetime(text, errors="coerce")


def to_datetime_series(series):
    return pd.to_datetime(series.map(parse_date), errors="coerce")


def parse_number(value):
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return np.nan
    text = str(value).strip()
    if text == "":
        return np.nan
    negative = text.startswith("(") and text.endswith(")")
    text = text.replace(",", "").replace("，", "").replace("(", "").replace(")", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        number = float(text)
        return -abs(number) if negative else number
    except ValueError:
        return np.nan


def to_number_series(series):
    return series.map(parse_number).astype("float64")


def scale_percent(series):
    p = to_number_series(series)
    if p.notna().any() and p.max() <= 1:
        p = p * 100
    return p


# =========================================================
# 3. 欄位配對與資料表判斷
# =========================================================
def find_header_row(raw, schema, scan_rows=20):
    aliases = set()
    for field, names in schema.items():
        aliases.add(normalize_text(field))
        aliases.update(normalize_text(x) for x in names)
    best_row, best_score = 0, -1
    for i in range(min(scan_rows, len(raw))):
        score = 0
        for v in raw.iloc[i].tolist():
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                continue
            n = normalize_text(v)
            if len(n) < 2:
                continue
            if n in aliases:
                score += 2
            elif any(len(a) >= 3 and a in n for a in aliases):
                score += 1
        if score > best_score:
            best_row, best_score = i, score
    return best_row, best_score


def map_columns(df, schema):
    result, used = {}, set()
    columns = list(df.columns)
    for field, aliases in schema.items():
        cands = [normalize_text(x) for x in [field] + aliases]
        match = None
        for col in columns:
            if col not in used and normalize_text(col) in cands:
                match = col
                break
        if match is None:
            for col in columns:
                if col in used:
                    continue
                n = normalize_text(col)
                if any(len(c) >= 3 and c in n for c in cands):
                    match = col
                    break
        if match is not None:
            result[match] = field
            used.add(match)
    return df.rename(columns=result), result


def detect_kind(raw):
    best_kind, best_hits, best_row = None, 0, 0
    for kind, schema in SCHEMAS.items():
        row, _ = find_header_row(raw, schema)
        cols = unique_names(list(raw.iloc[row]))
        # 建立臨時 df 進行欄位比對評分
        temp_df = pd.DataFrame(columns=cols)
        _, mapping = map_columns(temp_df, schema)
        hits = sum(1 for f in mapping.values() if f in DISTINCT[kind])
        if hits > best_hits:
            best_kind, best_hits, best_row = kind, hits, row
    if best_hits < 2:
        return None, 0, 0
    return best_kind, best_hits, best_row


def prepare_dataframe(df, kind):
    """【修正與完整補全】補全遺失邏輯，並對欄位進行轉型與自動衍生計算。"""
    df = df.copy()
    target_cols = DEFAULT_COLUMNS[kind]
    
    # 確保所有目標欄位都存在（若無則補空欄）
    for col in target_cols:
        if col not in df.columns:
            df[col] = np.nan
            
    # 依資料型態進行轉換
    for col in target_cols:
        dtype = EMPTY_DTYPES.get(col)
        if dtype == "datetime64[ns]":
            df[col] = to_datetime_series(df[col])
        elif dtype == "float64":
            if col == "完成率":
                df[col] = scale_percent(df[col])
            else:
                df[col] = to_number_series(df[col])
        else:
            df[col] = df[col].astype(str).replace(["nan", "None", "<NA>"], np.nan)

    # 針對估驗計價（estimate）自動計算衍生欄位
    if kind == "estimate":
        df["前期累計數量"] = df["前期累計數量"].fillna(0)
        df["本期數量"] = df["本期數量"].fillna(0)
        df["單價"] = df["單價"].fillna(0)
        
        # 衍生欄位計算
        df["累計數量"] = df["前期累計數量"] + df["本期數量"]
        df["本期金額"] = round_half_up(df["本期數量"] * df["單價"])
        df["前期累計金額"] = round_half_up(df["前期累計數量"] * df["單價"])
        df["累計金額"] = df["前期累計金額"] + df["本期金額"]
        
    # 只保留標準欄位與衍生欄位
    final_cols = target_cols + (ESTIMATE_DERIVED if kind == "estimate" else [])
    return df[final_cols]


# =========================================================
# 4. Streamlit 主程式 UI 介面
# =========================================================
def main():
    st.set_page_config(page_title="機電工程自動化系統", layout="wide")
    st.title("🏗️ 機電工程專案管理暨估驗計價自動化系統")
    
    uploaded_file = st.file_uploader("請上報 Excel 檔案 (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            # 讀取 Excel 內的所有 工作表
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            st.info(f"成功讀取檔案，共偵測到 {len(sheet_names)} 個工作表。")
            
            for sheet in sheet_names:
                # 先讀取原始矩陣，不用 header
                raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                if raw.empty or len(raw) < 2:
                    continue
                
                # 自動辨識資料表類型