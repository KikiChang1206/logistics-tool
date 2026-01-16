import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime

# 1. 網頁基本設定
st.set_page_config(page_title="信天翁系統", layout="centered")

# CSS 修復顏色與版面
st.markdown("""
    <style>
    .stApp { background-color: #F5F5F5; }
    .big-title { font-size: 36px !important; font-weight: bold; color: #000000 !important; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #333333 !important; margin-bottom: 25px; }
    p, span, label, div { color: #000000 !important; }

    .stFileUploader section {
        background-color: #FFFFFF !important;
        padding: 50px !important;
        border: 2px dashed #333333 !important;
        border-radius: 10px;
    }
    .stFileUploader [data-testid='stMarkdownContainer'] p { color: #000000 !important; }

    div.stButton > button:first-child {
        background-color: #004A99 !important;
        color: #FFFFFF !important;
        border: none;
        height: 50px;
        font-size: 18px;
        font-weight: bold;
        border-radius: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換器</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">請將「一般」與「聯郵」檔案拖入上方大框框中</p>', unsafe_allow_html=True)

# 2. 檔案上傳區
uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True)

has_general = has_lian_yu = False
gen_file_data = lian_file_data = None

if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name:
            has_general, gen_file_data = True, f
        elif "聯郵" in f.name:
            has_lian_yu, lian_file_data = True, f

# 3. 狀態檢查
st.write("### 📁 檔案狀態確認")
c1, c2 = st.columns(2)
with c1: st.success("✅ 一般文件：已就緒") if has_general else st.warning("⬜ 一般文件：待上傳")
with c2: st.success("✅ 聯郵文件：已就緒") if has_lian_yu else st.warning("⬜ 聯郵文件：待上傳")

# 4. 產出邏輯
if has_general and has_lian_yu:
    st.write("---")
    if st.button("🚀 信天翁文件產出", use_container_width=True):
        try:
            with st.spinner('正在按報關種類分組並插入間隔行...'):
                # --- 處理資料 ---
                df_gen_raw = pd.read_excel(gen_file_data, dtype=str).fillna('')
                gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"]
                df_gen_raw.columns = gen_headers[:len(df_gen_raw.columns)]
                df_gen_raw['HAWB / CN'] = df_gen_raw['HAWB / CN'].str.strip()
                search_db = df_gen_raw.set_index('HAWB / CN')

                dtype_dict = {'單價(TWD)': str, '寄件公司/統編': str, '統計方式': str, '提單號碼': str}
                df_customs = pd.read_excel(lian_file_data, sheet_name='報關明細', dtype=dtype_dict).fillna('')
                df_no_customs = pd.read_excel(lian_file_data, sheet_name='不報關-X7明細', dtype=dtype_dict).fillna('')
                
                # 補上「不報關」標記邏輯
                df_no_customs['報關'] = "不報關"
                
                # 合併資料
                combined = pd.concat([df_customs, df_no_customs], ignore_index=True)
                combined['提單號碼'] = combined['提單號碼'].str.strip()
                combined = combined[combined['提單號碼'] != '']

                # 比對一般文件資料
                def lookup_info(row):
                    hawb = row['提單號碼']
                    if hawb in search_db.index:
                        info = search_db.loc[hawb]
                        if isinstance(info, pd.DataFrame): info = info.iloc[0]
                        return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                    return pd.Series(["", "", "", ""])

                combined[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined.apply(lookup_info, axis=1)

                # --- 關鍵修正：按報關種類插入空行 ---
                final_cols_x = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
                
                spaced_data = []
                last_type = None
                
                for _, row in combined.iterrows():
                    current_type = str(row['報關']).strip()
                    # 如果這行的報關種類跟上一行不一樣，且不是第一行，就插入一個空行
                    if last_type is not None and current_type != last_type:
                        spaced_data.append(pd.Series([None] * len(final_cols_x), index=final_cols_x))
                    
                    # 處理「不報關」只顯示在該類第一行的邏輯
                    display_row = row.copy()
                    if current_type == "不報關":
                        if last_type == "不報關":
                            display_row['報關'] = "" # 同類型的後續行不顯示字樣
                    
                    spaced_data.append(display_row)
                    last_type = current_type

                df_final_export = pd.DataFrame(spaced_data).fillna('')[final_cols_x]

                # --- 輸出 Excel ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final_export.to_excel(writer, sheet_name='出口總明細', index=False)
                    df_gen_raw.to_excel(writer, sheet_name='一般', index=False)
                    for sheetname in ['出口總明細', '一般']:
                        ws = writer.sheets[sheetname]
                        for r_idx, row in enumerate(ws.iter_rows()):
                            for cell in row:
                                cell.font = Font(name='Arial', size=10)
                                cell.border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
                                if r_idx == 0: cell.alignment = Alignment(horizontal='left')

                today_str = datetime.now().strftime("%Y%m%d")
                st.balloons()
                st.success("✅ 處理完成！")
                st.download_button(label="📥 下載檔案", data=output.getvalue(), file_name=f"{today_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"發生錯誤: {e}")
else:
    st.info("💡 貼心提醒：請將兩個檔案同時選取並拖入上方框框中。")
