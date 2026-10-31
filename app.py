import streamlit as pd_st # Avoid name conflict
import streamlit as st
import pandas as pd
import numpy as np
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Set page config for professional look
st.set_page_config(
    page_title="營建估驗保留款與稅額智慧計算系統",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# CORE ENGINEERING LOGIC & UTILITIES
# -----------------------------------------------------------------------------

def precise_round(val, decimals=0):
    """
    使用 Decimal 進行精確的四捨五入（ROUND_HALF_UP），避免 Python 預設的銀行家捨入法（ROUND_HALF_EVEN）。
    在財務與稅額計算中，精確的四捨五入至整數或特定小數位數是法規合規與合約審核的關鍵。
    """
    if pd.isna(val) or val is None:
        return 0
    try:
        # 轉換為字串以防浮點數精度遺失
        d = Decimal(str(val))
        q = Decimal('10') ** -decimals
        res = d.quantize(q, rounding=ROUND_HALF_UP)
        return int(res) if decimals == 0 else float(res)
    except (ValueError, TypeError, InvalidOperation):
        return 0


def sanitize_value(val) -> str:
    """
    防範 Excel 公式注入（Excel Formula Injection）。
    如果輸入的字串以 '=', '+', '-', '@' 開頭，則在前方加上單引號 "'" 進行跳脫，
    防止惡意公式在 Excel 開啟時被自動執行，確保資料匯出的安全性。
    """
    if isinstance(val, str):
        val_strip = val.strip()
        if val_strip.startswith(('=', '+', '-', '@')):
            return f"'{val_strip}"
    return val


def parse_double_row_headers(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """
    自動辨識並合併雙列複雜表頭（Double-row complex headers）。
    營建工程報表常以 Excel 合併儲存格呈現雙列甚至多列設計（例如：上方「合約」、下方「數量」「單價」「複價」）。
    本函數會：
    1. 偵測前兩列是否有 NaN 或未命名欄位，判斷是否為雙列結構。
    2. 將第一列的分類名稱（如「合約」）向右填補（Forward Fill）到合併欄位。
    3. 合併第一列與第二列名稱，產出具備唯一性、語意清晰的扁平化欄位清單。
    """
    if df.shape[0] < 2:
        return df, list(df.columns)
    
    row1 = df.columns.to_list()
    row2 = df.iloc[0].to_list()
    
    # 判斷是否為雙列表頭結構：第二列在前幾個必填欄位（項目、工項名稱、單位）通常為 NaN，
    # 且後面數值欄位（如數量、單價、複價）在第一列為大類別名稱，第二列為子類別名稱。
    # 我們檢查第一列與第二列的特徵：
    is_double_row = False
    unnamed_count = sum(1 for col in row1 if 'Unnamed' in str(col) or pd.isna(col))
    nan_in_row2 = sum(1 for val in row2 if pd.isna(val))
    
    # 如果第一列有許多 Unnamed，且第二列有值，極可能是合併表頭
    if unnamed_count > 0 and nan_in_row2 > 0:
        is_double_row = True
        
    if not is_double_row:
        # 如果不是雙列表頭，直接返回
        return df, [str(col).strip() for col in df.columns]
    
    # 開始扁平化雙列表頭
    flat_headers = []
    current_parent = ""
    
    for col1, col2 in zip(row1, row2):
        c1_str = str(col1).strip() if not pd.isna(col1) else ""
        c2_str = str(col2).strip() if not pd.isna(col2) else ""
        
        # 處理 Pandas 自動生成的 Unnamed
        if "Unnamed:" in c1_str:
            c1_str = ""
            
        # 決定父級類別（Forward Fill 邏輯）
        if c1_str != "":
            current_parent = c1_str
            
        if current_parent != "" and c2_str != "":
            flat_name = f"{current_parent}_{c2_str}"
        elif c2_str != "":
            flat_name = c2_str
        elif c1_str != "":
            flat_name = c1_str
        else:
            flat_name = "未命名欄位"
            
        flat_headers.append(flat_name)
        
    # 去除作為表頭的第二列數據，並重新設定扁平化表頭
    cleaned_df = df.iloc[1:].copy()
    cleaned_df.columns = flat_headers
    cleaned_df = cleaned_df.reset_index(drop=True)
    
    return cleaned_df, flat_headers


def calculate_valuation_row(row, retention_rate: float) -> dict:
    """
    單筆估驗工項之核心計算邏輯。
    輸入工項合約資訊與本期估驗數量，計算：
    1. 本期估驗金額 = 本期估驗數量 * 合約單價
    2. 累計估驗數量 = 上期累計數量 + 本期估驗數量
    3. 累計估驗金額 = 累計估驗數量 * 合約單價
    4. 保留款 = 本期估驗金額 * 保留款比例
    """
    try:
        unit_price = float(row.get('合約_單價', 0))
    except (ValueError, TypeError):
        unit_price = 0.0
        
    try:
        current_qty = float(row.get('本期估驗_數量', 0))
    except (ValueError, TypeError):
        current_qty = 0.0
        
    try:
        prev_accum_qty = float(row.get('歷史累計_數量', 0))
    except (ValueError, TypeError):
        prev_accum_qty = 0.0
        
    # 計算金額與累計
    current_amount = precise_round(current_qty * unit_price)
    accum_qty = prev_accum_qty + current_qty
    accum_amount = precise_round(accum_qty * unit_price)
    retention = precise_round(current_amount * retention_rate)
    
    return {
        '本期估驗_金額': current_amount,
        '累計估驗_數量': accum_qty,
        '累計估驗_金額': accum_amount,
        '本期保留款': retention
    }


def generate_sample_excel() -> bytes:
    """
    生成符合雙列複雜表頭規範的範例 Excel 檔案，供使用者下載測試。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "估驗明細表"
    
    # 啟用網格線
    ws.views.sheetView[0].showGridLines = True
    
    # 設置雙列表頭資料
    headers_r1 = ["項目", "工項名稱", "單位", "合約", None, None, "歷史累計", None, "本期估驗", None]
    headers_r2 = [None, None, None, "數量", "單價", "複價", "數量", "金額", "數量", "金額"]
    
    ws.append(headers_r1)
    ws.append(headers_r2)
    
    # 合併儲存格 (A1:A2, B1:B2, C1:C2, D1:F1, G1:H1, I1:J1)
    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    ws.merge_cells("C1:C2")
    ws.merge_cells("D1:F1")
    ws.merge_cells("G1:H1")
    ws.merge_cells("I1:J1")
    
    # 範例數據
    data = [
        ["1", "基礎土方開挖與清運工程", "立方米", 1500, 450, 675000, 1000, 450000, 500, 225000],
        ["2", "強固結構混凝土澆置 (3000psi)", "立方米", 800, 3200, 2560000, 400, 1280000, 200, 640000],
        ["3", "竹節鋼筋綁紮與組立 (SD280)", "噸", 50, 28000, 1400000, 20, 560000, 15, 420000],
        ["4", "外牆防水分合層施工", "平方公尺", 1200, 350, 420000, 600, 210000, 300, 105000],
        ["5", "臨時假設工程及環境維護費", "式", 1, 150000, 150000, 0.5, 75000, 0.2, 30000],
    ]
    
    for row in data:
        ws.append(row)
        
    # 套用基本樣式
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # 格式化表頭
    for r in [1, 2]:
        for col_idx in range(1, 11):
            cell = ws.cell(row=r, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            
    # 設定內容樣式
    thin_side = Side(border_style="thin", color="D3D3D3")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    data_font = Font(name="微軟正黑體", size=10)
    
    for row in range(3, 8):
        for col in range(1, 11):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = border_all
            if col in [1, 3]:
                cell.alignment = align_center
            elif col == 2:
                cell.alignment = align_left
            else:
                cell.alignment = align_right
                if col in [4, 5, 6, 7, 8, 9, 10]:
                    cell.number_format = "#,##0"
                    
    # 設定欄寬
    column_widths = {1: 8, 2: 30, 3: 10, 4: 12, 5: 12, 6: 15, 7: 12, 8: 15, 9: 12, 10: 15}
    for col_idx, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        
    # 輸出成 bytes
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# -----------------------------------------------------------------------------
# STREAMLIT UI IMPLEMENTATION
# -----------------------------------------------------------------------------

def check_password() -> bool:
    """
    驗證使用者輸入的系統密碼是否正確，用以保護營建財務等敏感性商業數據。
    解鎖密碼預設為: construction888
    """
    CORRECT_PASSWORD = "construction888"

    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    # 渲染精美登入頁面
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("""
        <div style="background-color:#1F497D;padding:20px;border-radius:10px 10px 0px 0px;text-align:center;">
            <h2 style="color:white;margin:0;font-family:'微軟正黑體';">🔒 系統安全鎖定</h2>
            <p style="color:#D3D3D3;margin:5px 0 0 0;font-size:14px;font-family:'微軟正黑體';">營建估驗保留款與稅額智慧計算系統</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            password_input = st.text_input("請輸入系統密碼：", type="password", placeholder="請輸入安全存取密碼")
            submit_button = st.form_submit_button("🔑 驗證並解鎖系統", use_container_width=True)
            
            if submit_button:
                if password_input == CORRECT_PASSWORD:
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("😕 密碼錯誤，請重新輸入！")
                    
        st.markdown("""
        <div style="text-align:center;margin-top:15px;">
            <small style="color:gray;font-family:'微軟正黑體';">預設體驗解鎖密碼為 <code>construction888</code></small>
        </div>
        """, unsafe_allow_html=True)
        
    return False


def main():
    # 密碼安全鎖定
    if not check_password():
        st.stop()

    st.title("🏗️ 營建估驗保留款與稅額智慧計算系統")
    st.markdown("""
    本系統專為營建工程之**工程估驗、保留款計算與營業稅申報**設計。
    支援自動辨識 Excel 雙列合併表頭、高精確度財務四捨五入（防制浮點數精度誤差），並內建 Excel 公式注入防制安全機制。
    """)
    
    # Sidebar: 參數配置
    st.sidebar.header("⚙️ 計算參數設定")
    retention_rate_pct = st.sidebar.number_input(
        "工程保留款比例 (%)",
        min_value=0.0, max_value=100.0, value=10.0, step=0.5,
        help="通常合約規定每期估驗款扣留之成數，例如 10%"
    )
    retention_rate = retention_rate_pct / 100.0
    
    tax_rate_pct = st.sidebar.number_input(
        "加值型營業稅率 (%)",
        min_value=0.0, max_value=100.0, value=5.0, step=0.5,
        help="台灣法定營業稅率為 5%"
    )
    tax_rate = tax_rate_pct / 100.0
    
    rounding_decimals = st.sidebar.selectbox(
        "金額四捨五入位數",
        options=[0, 1, 2],
        index=0,
        format_func=lambda x: f"整數 (小數點後 {x} 位)" if x == 0 else f"小數點後 {x} 位",
        help="營建估驗實務上，台幣款項通常四捨五入至整數"
    )
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 範例檔案下載")
    st.sidebar.write("請下載標準雙列複雜表頭 Excel 範本進行測試：")
    
    sample_data = generate_sample_excel()
    st.sidebar.download_button(
        label="💾 下載範例 Excel 估驗表",
        data=sample_data,
        file_name="engineering_valuation_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.sidebar.info("💡 提示：本系統具備 Excel Formula Injection 防護，所有導出字串格式將自動檢測並跳脫公式關鍵字。")
    
    # Main Section: 檔案上傳與編輯
    uploaded_file = st.file_uploader(
        "👉 請上傳工程估驗明細表 (Excel 格式 .xlsx)",
        type=["xlsx"],
        help="必須包含項目、工項名稱、單位、合約單價、歷史累計數量、本期估驗數量等資訊。"
    )
    
    # 預設數據載入（若使用者未上傳，則使用範例數據）
    df_raw = None
    if uploaded_file is not None:
        try:
            # 讀取不指定 header，由我們的 parse_double_row_headers 自行解析
            df_raw = pd.read_excel(uploaded_file, header=None)
            st.success("🎉 檔案上傳成功！系統已載入您的數據。")
        except Exception as e:
            st.error(f"檔案解析失敗，請確認檔案格式是否正確。錯誤訊息：{str(e)}")
    else:
        st.info("ℹ️ 尚未上傳檔案。下方將自動以預設範例數據進行演示，您亦可於左側下載範例檔案。")
        # 載入預設範例 bytes
        df_raw = pd.read_excel(io.BytesIO(sample_data), header=None)
        
    if df_raw is not None:
        # 第一步：處理 Pandas 的 Header
        # 因 read_excel(header=None) 讀入，第一列為 columns 屬性，我們重新轉化
        df_raw.columns = df_raw.iloc[0]
        df_processed_raw = df_raw.iloc[1:].reset_index(drop=True)
        
        # 解析雙列表頭
        df_cleaned, flat_headers = parse_double_row_headers(df_processed_raw)
        
        st.subheader("📊 雙列複雜表頭自動解析與扁平化結果")
        st.markdown(f"系統已自動將多層級表頭扁平化為一維欄位。解析出之欄位結構：")
        st.code(" , ".join(flat_headers))
        
        # 檢查必備欄位
        required_keywords = ['項目', '工項名稱', '單位', '合約_單價', '歷史累計_數量', '本期估驗_數量']
        missing_fields = [f for f in required_keywords if f not in flat_headers]
        
        if missing_fields:
            st.warning(f"⚠️ 注意！上傳的表格缺少關鍵欄位：{', '.join(missing_fields)}。系統將為您建立虛擬數據補齊，以維持計算。")
            # 自動修復/填充缺失欄位以防出錯
            for field in missing_fields:
                if field == '項目':
                    df_cleaned['項目'] = [str(i+1) for i in range(len(df_cleaned))]
                elif field == '工項名稱':
                    df_cleaned['工項名稱'] = [f"新增工項 {i+1}" for i in range(len(df_cleaned))]
                elif field == '單位':
                    df_cleaned['單位'] = "項"
                elif field in ['合約_單價', '歷史累計_數量', '本期估驗_數量']:
                    df_cleaned[field] = 0.0
                    
        # 進行數據清潔
        # 對於數值型欄位，進行空值填充與型態轉換
        for num_col in ['合約_單價', '歷史累計_數量', '本期估驗_數量']:
            df_cleaned[num_col] = pd.to_numeric(df_cleaned[num_col], errors='coerce').fillna(0.0)
            
        # 讓使用者可在 Streamlit 介面上直接編輯「本期估驗_數量」
        st.subheader("✏️ 互動式資料編輯區")
        st.markdown("您可以直接在下方表格中修改 **[本期估驗_數量]**，系統將即時重新計算估驗額、保留款與稅金。")
        
        # 為了安全起見，在將資料傳給 Data Editor 前先進行 Formula Injection 防護過濾（展示作用）
        for col in df_cleaned.columns:
            df_cleaned[col] = df_cleaned[col].apply(sanitize_value)
            
        # 設定 data_editor 參數，僅允許編輯本期估驗數量
        edited_df = st.data_editor(
            df_cleaned,
            column_config={
                "本期估驗_數量": st.column_config.NumberColumn(
                    "本期估驗數量",
                    help="請輸入本期實際施工完成並欲估驗的數量",
                    min_value=0.0,
                    step=0.01,
                    required=True,
                ),
                "合約_單價": st.column_config.NumberColumn("合約單價 (ReadOnly)", disabled=True),
                "項目": st.column_config.TextColumn("項目 (ReadOnly)", disabled=True),
                "工項名稱": st.column_config.TextColumn("工項名稱 (ReadOnly)", disabled=True),
                "單位": st.column_config.TextColumn("單位 (ReadOnly)", disabled=True),
            },
            disabled=[col for col in df_cleaned.columns if col != "本期估驗_數量"],
            use_container_width=True
        )
        
        # 重新計算估驗各列結果
        calculated_rows = []
        for idx, row in edited_df.iterrows():
            calc_res = calculate_valuation_row(row, retention_rate)
            calculated_rows.append(calc_res)
            
        df_calc = pd.DataFrame(calculated_rows)
        
        # 合併原欄位與計算結果
        # 原欄位中與計算結果重複的（如本期估驗_金額、累計估驗_數量、累計估驗_金額）進行更新
        for col in df_calc.columns:
            edited_df[col] = df_calc[col]
            
        # 計算合約複價
        edited_df['合約_複價'] = (edited_df['合約_數量'].astype(float, errors='ignore') * edited_df['合約_單價'].astype(float)).apply(precise_round)
        
        # 重新排列欄位順序以求精美
        cols_ordered = ['項目', '工項名稱', '單位', '合約_數量', '合約_單價', '合約_複價', '歷史累計_數量', '本期估驗_數量', '本期估驗_金額', '累計估驗_數量', '累計估驗_金額', '本期保留款']
        existing_ordered_cols = [c for c in cols_ordered if c in edited_df.columns]
        # 把沒列入的其他欄位拼在後面
        remaining_cols = [c for c in edited_df.columns if c not in existing_ordered_cols]
        final_df = edited_df[existing_ordered_cols + remaining_cols]
        
        st.subheader("📋 重新計算後之估驗明細")
        st.dataframe(final_df, use_container_width=True)
        
        # -----------------------------------------------------------------------------
        # SUMMARY FINANCIAL COMPUTATIONS (DECIMAL PRECISION)
        # -----------------------------------------------------------------------------
        # 使用 Decimal 進行全域加總，防範浮點數累加誤差
        sum_current_valuation = Decimal('0')
        sum_retention = Decimal('0')
        
        for idx, row in final_df.iterrows():
            sum_current_valuation += Decimal(str(row['本期估驗_金額']))
            sum_retention += Decimal(str(row['本期保留款']))
            
        # 應稅金額 (通常為 本期估驗總額 扣除 保留款)
        # 根據業界合約不同，有時營業稅基是本期估驗總額，本處採用最具工程代表性的「扣除保留款後之淨額為稅基」或「估驗總額為稅基」可由公式決定
        # 在此我們採用：「本期應稅金額 = 本期估驗總額」作為標準發票稅基計算（台灣常見實務）
        taxable_amount = sum_current_valuation
        
        # 加值型營業稅 = 應稅金額 * 5% (高精度四捨五入)
        vat_tax = Decimal(str(taxable_amount)) * Decimal(str(tax_rate))
        vat_tax_rounded = Decimal(str(precise_round(float(vat_tax), rounding_decimals)))
        
        # 實付金額 = 本期估驗總額 - 本期保留款 + 營業稅
        net_payable = sum_current_valuation - sum_retention + vat_tax_rounded
        
        # UI 統計看板
        st.subheader("💰 本期估驗財務資訊總結")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric(
                label="🏢 本期估驗總額 (A)",
                value=f"${int(sum_current_valuation):,}",
                help="本期所有完工工項估驗金額之加總"
            )
        with col2:
            st.metric(
                label="🔒 本期扣留保留款 (B)",
                value=f"${int(sum_retention):,}",
                help=f"依保留款比例 {retention_rate_pct}% 計算之扣留款"
            )
        with col3:
            st.metric(
                label="📝 應稅金額基數 (C = A)",
                value=f"${int(taxable_amount):,}",
                help="申報加值型營業稅之應稅基數金額"
            )
        with col4:
            st.metric(
                label="📈 營業稅額 (D)",
                value=f"${int(vat_tax_rounded):,}",
                help=f"加值型營業稅 (應稅金額 * {tax_rate_pct}%)，採精確四捨五入"
            )
        with col5:
            st.metric(
                label="💵 本期實付淨額 (E = A - B + D)",
                value=f"${int(net_payable):,}",
                help="承包商本期可實際領取之票匯金額"
            )
            
        # -----------------------------------------------------------------------------
        # EXPORT TO EXCEL WITH FORMATTING (OPENPYXL)
        # -----------------------------------------------------------------------------
        # 導出包含格式與彙總公式的精緻 Excel
        def export_calculated_excel(df: pd.DataFrame, sum_val, sum_ret, tax_base, tax, net_pay) -> bytes:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "估驗請款計算單"
            
            # 啟用網格
            ws.views.sheetView[0].showGridLines = True
            
            # 標題區
            ws.merge_cells("A1:L1")
            title_cell = ws["A1"]
            title_cell.value = "工程估驗請款計算明細單 (第 01 期)"
            title_cell.font = Font(name="微軟正黑體", size=16, bold=True, color="1F497D")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 40
            
            # 設定參數說明區
            ws["A2"] = "保留款比例:"
            ws["B2"] = f"{retention_rate_pct}%"
            ws["D2"] = "營業稅率:"
            ws["E2"] = f"{tax_rate_pct}%"
            ws["G2"] = "製表日期:"
            import datetime
            ws["H2"] = datetime.date.today().strftime("%Y-%m-%d")
            
            # 參數區樣式
            param_font = Font(name="微軟正黑體", size=10, bold=True)
            for r in [2]:
                for c in range(1, 13):
                    cell = ws.cell(row=r, column=c)
                    cell.font = param_font
                    
            ws.row_dimensions[2].height = 20
            
            # 扁平化表頭寫入 (Row 4)
            headers = df.columns.tolist()
            ws.append([]) # Row 3 empty
            ws.append(headers) # Row 4
            
            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            header_font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[4].height = 28
            
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=4, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center
                
            # 寫入內容數據
            thin_side = Side(border_style="thin", color="D3D3D3")
            border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            data_font = Font(name="微軟正黑體", size=10)
            align_left = Alignment(horizontal="left", vertical="center")
            align_right = Alignment(horizontal="right", vertical="center")
            
            current_row = 5
            for idx, row in df.iterrows():
                # 安全跳脫防 formula injection
                row_vals = [sanitize_value(v) for v in row.values]
                ws.append(row_vals)
                
                # 設定格式
                ws.row_dimensions[current_row].height = 22
                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=current_row, column=col_idx)
                    cell.font = data_font
                    cell.border = border_all
                    
                    # 對齊與數值格式化
                    val = row_vals[col_idx - 1]
                    if col_idx in [1, 3]: # 項目、單位
                        cell.alignment = align_center
                    elif col_idx == 2: # 工項名稱
                        cell.alignment = align_left
                    else: # 數值欄位
                        cell.alignment = align_right
                        if isinstance(val, (int, float, Decimal)):
                            cell.number_format = "#,##0"
                current_row += 1
                
            # 彙總區樣式 (Row start = current_row + 1)
            ws.append([]) # 留空一列
            current_row += 1
            
            # 本期估驗總額 彙總
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
            ws.cell(row=current_row, column=1).value = "本期估驗金額總計 (A)"
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right")
            val_cell = ws.cell(row=current_row, column=9)
            # 使用 Excel 公式進行驗算，同時也寫入 Decimal 預算結果
            val_cell.value = f"=SUM(I5:I{current_row-2})"
            val_cell.number_format = "#,##0"
            val_cell.font = Font(name="微軟正黑體", size=10, bold=True)
            
            # 保留款 彙總
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
            ws.cell(row=current_row, column=1).value = f"扣除工程保留款 (B = A * {retention_rate_pct}%)"
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right")
            ret_cell = ws.cell(row=current_row, column=12)
            ret_cell.value = f"=SUM(L5:L{current_row-3})"
            ret_cell.number_format = "#,##0"
            ret_cell.font = Font(name="微軟正黑體", size=10, bold=True)
            
            # 營業稅
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
            ws.cell(row=current_row, column=1).value = f"加計 5% 營業稅額 (D = A * {tax_rate_pct}%)"
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right")
            tax_cell = ws.cell(row=current_row, column=12)
            # 營業稅公式：本期估驗總額(I) * 5%
            tax_cell.value = f"=ROUND(I{current_row-2}*{tax_rate}, 0)"
            tax_cell.number_format = "#,##0"
            tax_cell.font = Font(name="微軟正黑體", size=10, bold=True)
            
            # 實付總額 (雙線下底線，代表會計學最終金額)
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
            ws.cell(row=current_row, column=1).value = "本期實際實付金額 (E = A - B + D)"
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right")
            ws.cell(row=current_row, column=1).font = Font(name="微軟正黑體", size=11, bold=True, color="9C0006")
            
            net_cell = ws.cell(row=current_row, column=12)
            # 實付 = A - B + D = I(current_row-3) - L(current_row-2) + L(current_row-1)
            net_cell.value = f"=I{current_row-3}-L{current_row-2}+L{current_row-1}"
            net_cell.number_format = "#,##0"
            net_cell.font = Font(name="微軟正黑體", size=11, bold=True, color="9C0006")
            
            # 套用彙總欄位框線 (上細線、下雙線)
            double_bottom = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))
            for col_idx in range(1, 13):
                ws.cell(row=current_row, column=col_idx).border = double_bottom
                
            # 自動適應調整其餘欄寬
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                if col_letter in ['A', 'B', 'C']:
                    # 避免合併儲存格造成欄寬暴增
                    continue
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
                
            # 個別指定 A, B, C 欄寬
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 10
            
            output = io.BytesIO()
            wb.save(output)
            return output.getvalue()
            
        st.subheader("📤 導出精緻請款報表")
        st.markdown("您可以將本期重算完畢之數據，導出為符合企業申報、審查標準之 Excel 請款報表。")
        
        excel_data = export_calculated_excel(
            final_df,
            sum_current_valuation,
            sum_retention,
            taxable_amount,
            vat_tax_rounded,
            net_payable
        )
        
        st.download_button(
            label="📥 匯出精確估驗計算單 (Excel)",
            data=excel_data,
            file_name=f"valuation_report_calculated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.balloons()

if __name__ == "__main__":
    main()
