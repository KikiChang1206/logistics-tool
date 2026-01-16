import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime

# 1. 網頁基本設定
st.set_page_config(page_title="信天翁系統", layout="centered")

# 自定義 CSS：換成黑底風格，並修復按鈕顏色
st.markdown("""
    <style>
    /* 全域背景：深黑色 */
    .stApp { background-color: #0E1117; }
    
    /* 標題與標題文字：白色 */
    .big-title { font-size: 36px !important; font-weight: bold; color: #FFFFFF !important; }
    .sub-title { font-size: 16px; color: #CCCCCC !important; margin-bottom: 25px; }

    /* 上傳框內部：純白背景 */
    .stFileUploader section {
        background-color: #FFFFFF !important;
        padding: 40px !important;
        border: 2px dashed #FFFFFF !important;
        border-radius: 10px;
    }
    
    /* 修正「Browse Files」按鈕：白底、黑字、黑框 */
    button[data-testid="baseButton-secondary"] {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #000000 !important;
    }
    button[data-testid="baseButton-secondary"] p {
        color: #000000 !important;
    }

    /* 狀態框樣式：徹底透明背景，文字改為白色 */
    div[data-testid="stNotification"], div[data-testid="stNotificationV2"] {
        background-color: transparent !important;
        background: none !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0px !important;
    }
    
    /* 狀態文字：白色 (配合黑底) */
    div[data-testid="stNotification"] p, 
    div[data-testid="stNotificationV2"] p {
        color: #FFFFFF !important;
        font-size: 16px !important;
    }

    /* 「信天翁文件產出」與「下載」按鈕：白底、黑字、黑邊 */
    div.stButton > button {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 2px solid #000000 !important;
        height: 50px;
        font-size: 18px;
        font-weight: bold;
        width: 100%;
    }
    /* 按鈕內的文字強制黑色 */
    div.stButton > button p {
        color: #000000 !important;
    }
    div.stButton > button:hover {
        background-color: #EEEEEE !important;
    }

    /* 檔案確認標題：白色 */
    h3 { color: #FFFFFF !important; }

    /* 強制所有 Markdown 文字為白色 */
    .stMarkdown p, .stMarkdown span, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換器</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">請將「一般」與「聯郵」檔案拖入上方白色區域中</p>', unsafe_allow_html=True)

# 2. 檔案上傳區
uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True)

has_general = False
has_lian_yu = False
gen_file_data = None
lian_file_data = None

if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name:
            has_general = True
            gen_file_data = f
        elif "聯郵" in f.name:
            has_lian_yu = True
            lian_file_data = f

# 3. 狀態檢查清單
st.write("### 📁 檔案狀態確認")
col1, col2 = st.columns(2)
with col1:
    if has_general:
        st.success("✅ 一般文件：已就緒")
    else:
        st.info("⬜ 一般文件：待上傳")

with col2:
    if has_lian_yu:
        st.success("✅ 聯郵文件：已就緒")
    else:
        st.info("⬜ 聯郵文件：待上傳")

# 4. 產出邏輯
if has_general and has_lian_yu:
    st.write("---")
    if st.button("🚀 信天翁文件產出", use_container_width=True):
        try:
            with st.spinner('正在分析報關種類並進行精確對位...'):
                # --- A. 讀取一般文件 ---
                df_gen_raw = pd.read_excel(gen_file_data, dtype=str).fillna('')
                gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"]
                df_gen_raw.columns = gen_headers[:len(df_gen_raw.columns)]
                df_gen_raw['HAWB / CN'] = df_gen_raw['HAWB / CN'].str.strip()
                search_db = df_gen_raw.set_index('HAWB / CN')

                # --- B. 讀取聯郵文件 ---
                dtype_dict = {'單價(TWD)': str, '寄件公司/統編': str, '統計方式': str, '提單號碼': str}
                df_customs = pd.read_excel(lian_file_data, sheet_name='報關明細', dtype=dtype_dict).fillna('')
                df_no_customs = pd.read_excel(lian_file_data, sheet_name='不報關-X7明細', dtype=dtype_dict).fillna('')
                
                df_no_customs['報關'] = "不報關"
                combined = pd.concat([df_customs, df_no_customs], ignore_index=True)
                combined['提單號碼'] = combined['提單號碼'].str.strip()
                combined = combined[(combined['提單號碼'] != '') & (combined['提單號碼'] != 'nan')]

                def format_price(val):
                    try: return "{:.2f}".format(float(val)) if val and val != 'nan' else ""
                    except: return val
                combined['單價(TWD)'] = combined['單價(TWD)'].apply(format_price)

                # --- C. 跨表比對 ---
                def lookup_info(row):
                    hawb = row['提單號碼']
                    if hawb in search_db.index:
                        info = search_db.loc[hawb]
                        if isinstance(info, pd.DataFrame): info = info.iloc[0]
                        return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                    return pd.Series(["", "", "", ""])

                combined[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined.apply(lookup_info, axis=1)

                # --- D. 報關種類插入空行邏輯 ---
                final_cols_x = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
                
                spaced_rows = []
                last_type = None
                for _, row in combined.iterrows():
                    current_type = str(row['報關']).strip()
                    if last_type is not None and current_type != last_type and current_type != "":
                        spaced_rows.append(pd.Series([None] * len(final_cols_x), index=final_cols_x))
                    
                    display_row = row.copy()
                    if current_type == "不報關" and last_type == "不報關":
                        display_row['報關'] = ""
                    
                    spaced_rows.append(display_row)
                    last_type = current_type

                df_final_export = pd.DataFrame(spaced_rows).fillna('')[final_cols_x]

                # --- E. 產出單一頁籤 Excel ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final_export.to_excel(writer, sheet_name='出口總明細', index=False)
                    ws = writer.sheets['出口總明細']
                    for r_idx, row in enumerate(ws.iter_rows()):
                        for cell in row:
                            cell.font = Font(name='Arial', size=10)
                            cell.border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
                            if r_idx == 0:
                                cell.alignment = Alignment(horizontal='left')

                today_str = datetime.now().strftime("%Y%m%d")
                st.balloons()
                st.success("✅ 處理完成！")
                st.download_button(label="📥 下載轉換後的信天翁檔案", data=output.getvalue(), file_name=f"{today_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"發生錯誤: {e}")
else:
    st.info("💡 提示：請將含有『一般』與『聯郵』字樣的兩個檔案一起拖入白色框框中。")
