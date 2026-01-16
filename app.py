import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime
import urllib.parse

# 1. 網頁基本設定 (設定為暗黑風格與置中版面)
st.set_page_config(page_title="信天翁系統", layout="centered")

# 自定義 CSS：黑底背景、白色標題、白底黑字按鈕與上傳框
st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 36px !important; font-weight: bold; color: #FFFFFF !important; }
    .sub-title { font-size: 16px; color: #CCCCCC !important; margin-bottom: 25px; }
    
    /* 上傳框：背景白色，邊框虛線 */
    .stFileUploader section {
        background-color: #FFFFFF !important;
        padding: 40px !important;
        border: 2px dashed #333333 !important;
        border-radius: 10px;
    }
    .stFileUploader [data-testid='stMarkdownContainer'] p { color: #000000 !important; }
    
    /* 狀態顯示：透明背景，文字白色 */
    div[data-testid="stNotification"], div[data-testid="stNotificationV2"] {
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }
    div[data-testid="stNotification"] p, div[data-testid="stNotificationV2"] p {
        color: #FFFFFF !important;
    }

    /* 按鈕：白底、黑字、黑邊框 */
    div.stButton > button {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 2px solid #000000 !important;
        height: 50px;
        font-size: 18px;
        font-weight: bold;
        width: 100%;
    }
    div.stButton > button p { color: #000000 !important; }
    
    /* 寄信按鈕：白底、黑字、綠邊框區別 */
    .email-btn {
        display: inline-block;
        width: 100%;
        text-align: center;
        background-color: #FFFFFF;
        color: #000000 !important;
        border: 2px solid #28A745;
        padding: 12px;
        font-weight: bold;
        text-decoration: none;
        border-radius: 5px;
        margin-top: 10px;
    }
    h3 { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換器</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">請將「一般」與「聯郵」檔案拖入上方區域中</p>', unsafe_allow_html=True)

# 2. 檔案上傳區
uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True)

has_gen = has_lian = False
gen_file = lian_file = None

if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name:
            has_gen, gen_file = True, f
        elif "聯郵" in f.name:
            has_lian, lian_file = True, f

# 3. 狀態檢查清單
st.write("### 📁 檔案狀態確認")
c1, c2 = st.columns(2)
with c1: st.success("✅ 一般文件：已就緒") if has_gen else st.info("⬜ 一般文件：待上傳")
with c2: st.success("✅ 聯郵文件：已就緒") if has_lian else st.info("⬜ 聯郵文件：待上傳")

# 4. 處理邏輯
if has_gen and has_lian:
    st.write("---")
    if st.button("🚀 信天翁文件產出", use_container_width=True):
        try:
            with st.spinner('正在分析報關人與件數...'):
                # --- A. 讀取一般文件 (資料庫) ---
                df_gen_raw = pd.read_excel(gen_file, dtype=str).fillna('')
                gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"]
                df_gen_raw.columns = gen_headers[:len(df_gen_raw.columns)]
                search_db = df_gen_raw.set_index('HAWB / CN')

                # --- B. 讀取聯郵文件 ---
                dtype_dict = {'單價(TWD)': str, '寄件公司/統編': str, '統計方式': str, '提單號碼': str}
                df_cust = pd.read_excel(lian_file, sheet_name='報關明細', dtype=str).fillna('')
                df_no_customs = pd.read_excel(lian_file, sheet_name='不報關-X7明細', dtype=str).fillna('')
                df_no_customs['報關'] = "不報關"

                # --- C. 核心邏輯：判斷正報寄件人與統計件數 ---
                combined_raw = pd.concat([df_cust, df_no_customs], ignore_index=True)
                combined_raw['提單號碼'] = combined_raw['提單號碼'].str.strip()
                combined_raw = combined_raw[combined_raw['提單號碼'] != '']

                stats_positive = {} # 儲存格式: { "寄件人名稱": 件數 }
                current_sender = None
                
                for _, row in combined_raw.iterrows():
                    val_a = str(row['報關']).strip()
                    val_p = str(row['寄件人']).strip()
                    
                    # 判斷是否為新的正報區塊開始
                    if "正式報關" in val_a or "合併正報" in val_a:
                        current_sender = val_p if val_p != "" else "未知"
                        if current_sender not in stats_positive:
                            stats_positive[current_sender] = 0
                    
                    # 如果處於正報區間，且沒進入不報關區間，則累計
                    if current_sender and "不報關" not in val_a:
                        stats_positive[current_sender] += 1
                    elif "不報關" in val_a:
                        current_sender = None

                # 組合成 Gmail 用的正報字串
                pos_string = "、".join([f"{name} {count} 件" for name, count in stats_positive.items()])
                count_no = len(df_no_customs)
                total_count = len(combined_raw)

                # --- D. 跨表比對收件人資料 ---
                def lookup_info(row):
                    hawb = str(row['提單號碼']).strip()
                    if hawb in search_db.index:
                        info = search_db.loc[hawb]
                        if isinstance(info, pd.DataFrame): info = info.iloc[0]
                        return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                    return pd.Series(["", "", "", ""])

                combined_raw[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined_raw.apply(lookup_info, axis=1)

                # --- E. 插入分組空行邏輯 ---
                final_cols = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
                spaced_rows = []
                last_type = None
                for _, row in combined_raw.iterrows():
                    curr_type = str(row['報關']).strip()
                    if last_type is not None and curr_type != last_type and curr_type != "":
                        spaced_rows.append(pd.Series([None] * len(final_cols), index=final_cols))
                    
                    display_row = row.copy()
                    if curr_type == "不報關" and last_type == "不報關":
                        display_row['報關'] = ""
                    spaced_rows.append(display_row)
                    last_type = curr_type

                df_final = pd.DataFrame(spaced_rows).fillna('')[final_cols]

                # --- F. 產出單一頁籤 Excel ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, sheet_name='出口總明細', index=False)
                    ws = writer.sheets['出口總明細']
                    for r_idx, row in enumerate(ws.iter_rows()):
                        for cell in row:
                            cell.font = Font(name='Arial', size=10)
                            if r_idx == 0: cell.alignment = Alignment(horizontal='left')

                today_str = datetime.now().strftime("%Y%m%d")
                st.balloons()
                st.success("✅ 處理完成！已自動完成件數統計與格式整理。")
                st.download_button(label="📥 下載轉換後的信天翁檔案", data=output.getvalue(), file_name=f"{today_str}_信天翁_Manifest.xlsx", use_container_width=True)

                # --- G. Gmail 範本與連結 ---
                recipient = "請輸入窗口信箱@gmail.com"
                subject = f"{today_str} 信天翁報關資料"
                email_body = (
                    f"Dears\n\n"
                    f"今日出口明細如附檔，共 {total_count} 件\n"
                    f"請再協助申報，並安排出口，謝謝\n\n"
                    f"正報：{pos_string}\n"
                    f"不報關： {count_no} 件"
                )
                
                mailto_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={recipient}&su={urllib.parse.quote(subject)}&body={urllib.parse.quote(email_body)}"
                st.markdown(f'<a href="{mailto_url}" target="_blank" class="email-btn">📧 開啟 Gmail (自動填寫：{pos_string})</a>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"發生錯誤: {e}")
else:
    st.info("💡 提示：請將含有『一般』與『聯郵』字樣的兩個檔案同時拖入上方白色區域。")
