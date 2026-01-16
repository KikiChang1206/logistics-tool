import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment
from datetime import datetime
import urllib.parse

# 1. 網頁基本設定
st.set_page_config(page_title="信天翁系統", layout="centered")

# 自定義 CSS
st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 36px !important; font-weight: bold; color: #FFFFFF !important; }
    .sub-title { font-size: 16px; color: #CCCCCC !important; margin-bottom: 25px; }
    .stFileUploader section { background-color: #FFFFFF !important; padding: 40px !important; border: 2px dashed #333 !important; border-radius: 10px; }
    button[data-testid="baseButton-secondary"] { background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #000000 !important; }
    button[data-testid="baseButton-secondary"] p { color: #000000 !important; }
    div[data-testid="stNotification"], div[data-testid="stNotificationV2"] { background-color: transparent !important; border: none !important; }
    div[data-testid="stNotification"] p, div[data-testid="stNotificationV2"] p { color: #FFFFFF !important; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #000000 !important; height: 50px; font-weight: bold; width: 100%; }
    div.stButton > button p { color: #000000 !important; }
    .email-btn { display: inline-block; width: 100%; text-align: center; background-color: #FFFFFF; color: #000000 !important; border: 2px solid #28A745; padding: 12px; font-weight: bold; text-decoration: none; border-radius: 5px; margin-top: 10px; }
    h3 { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換器</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">請上傳「一般」與「聯郵」檔案</p>', unsafe_allow_html=True)

uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True)

has_gen = has_lian = False
gen_file = lian_file = None

if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name: has_gen, gen_file = True, f
        elif "聯郵" in f.name: has_lian, lian_file = True, f

st.write("### 📁 檔案狀態確認")
c1, c2 = st.columns(2)
with c1: st.success("✅ 一般文件：就緒") if has_gen else st.info("⬜ 一般文件：待上傳")
with c2: st.success("✅ 聯郵文件：就緒") if has_lian else st.info("⬜ 聯郵文件：待上傳")

if has_gen and has_lian:
    st.write("---")
    if st.button("🚀 信天翁文件產出", use_container_width=True):
        try:
            # --- A. 讀取一般文件作為資料庫 ---
            df_gen = pd.read_excel(gen_file, dtype=str).fillna('')
            gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"]
            df_gen.columns = gen_headers[:len(df_gen.columns)]
            search_db = df_gen.set_index('HAWB / CN')

            # --- B. 讀取聯郵檔案 ---
            df_cust = pd.read_excel(lian_file, sheet_name='報關明細', dtype=str).fillna('')
            df_no = pd.read_excel(lian_file, sheet_name='不報關-X7明細', dtype=str).fillna('')
            df_no['報關'] = "不報關"
            
            # 合併處理
            combined = pd.concat([df_cust, df_no], ignore_index=True)
            combined['提單號碼'] = combined['提單號碼'].str.strip()
            combined = combined[combined['提單號碼'] != '']

            # --- C. 關鍵：動態判斷「正報人」與「件數」 ---
            # 建立統計字典
            stats_positive = {} # { "名稱": 件數 }
            current_sender = None
            
            # 遍歷資料（模擬 Excel A欄與P欄關係）
            for i, row in combined.iterrows():
                val_a = str(row['報關']).strip()
                val_p = str(row['寄件人']).strip()
                
                # 如果 A 欄出現關鍵字，更新當前寄件人
                if "正式報關" in val_a or "合併正報" in val_a:
                    current_sender = val_p if val_p != "" else "未知寄件人"
                    if current_sender not in stats_positive:
                        stats_positive[current_sender] = 0
                
                # 如果目前在正報區間，且 A 欄不是「不報關」
                if current_sender and "不報關" not in val_a:
                    stats_positive[current_sender] += 1
                elif "不報關" in val_a:
                    current_sender = None # 進入不報關區間，重設

            # 格式化輸出字串
            pos_list = [f"{name} {count} 件" for name, count in stats_positive.items()]
            pos_string = "、".join(pos_list)
            
            # 其他統計
            count_no = len(df_no)
            total_count = len(combined)

            # --- D. 跨表比對收件人資料 ---
            def lookup_info(row):
                hawb = str(row['提單號碼']).strip()
                if hawb in search_db.index:
                    info = search_db.loc[hawb]
                    if isinstance(info, pd.DataFrame): info = info.iloc[0]
                    return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                return pd.Series(["", "", "", ""])
            
            combined[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined.apply(lookup_info, axis=1)

            # --- E. 產出檔案 ---
            final_cols = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
            
            # 插入視覺空行邏輯 (略...)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined[final_cols].to_excel(writer, sheet_name='出口總明細', index=False)
            
            today_str = datetime.now().strftime("%Y%m%d")
            st.balloons()
            st.success("✅ 處理完成！已根據 A 欄與 P 欄邏輯自動分析。")
            st.download_button(label="📥 下載檔案", data=output.getvalue(), file_name=f"{today_str}_信天翁_Manifest.xlsx", use_container_width=True)

            # --- F. Gmail 範本產出 ---
            recipient = "窗口信箱@gmail.com"
            subject = f"{today_str} 信天翁報關資料"
            email_body = (
                f"Dears\n\n"
                f"今日出口明細如附檔，共 {total_count} 件\n"
                f"請再協助申報，並安排出口，謝謝\n\n"
                f"正報：{pos_string}\n"
                f"不報關： {count_no} 件"
            )
            
            mailto_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={recipient}&su={urllib.parse.quote(subject)}&body={urllib.parse.quote(email_body)}"
            st.markdown(f'<a href="{mailto_url}" target="_blank" class="email-btn">📧 開啟 Gmail (已自動判斷：{pos_string})</a>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"發生錯誤: {e}")
