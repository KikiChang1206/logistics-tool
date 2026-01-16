import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment
from datetime import datetime
import urllib.parse

# 1. 網頁基本設定
st.set_page_config(page_title="信天翁系統", layout="centered")

# 自定義 CSS (保持黑底風格)
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

# 檔案狀態確認
st.write("### 📁 檔案狀態確認")
c1, c2 = st.columns(2)
with c1:
    if has_gen: st.success("✅ 一般文件：就緒")
    else: st.info("⬜ 一般文件：待上傳")
with c2:
    if has_lian: st.success("✅ 聯郵文件：就緒")
    else: st.info("⬜ 聯郵文件：待上傳")

if has_gen and has_lian:
    st.write("---")
    if st.button("🚀 信天翁文件產出", use_container_width=True):
        try:
            # A. 讀取與處理資料 (邏輯維持先前版本)
            df_gen = pd.read_excel(gen_file, dtype=str).fillna('')
            gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"]
            df_gen.columns = gen_headers[:len(df_gen.columns)]
            search_db = df_gen.set_index('HAWB / CN')

            df_cust = pd.read_excel(lian_file, sheet_name='報關明細', dtype=str).fillna('')
            df_no_sheet = pd.read_excel(lian_file, sheet_name='不報關-X7明細', dtype=str).fillna('')
            df_no_sheet['報關'] = "不報關"
            
            combined = pd.concat([df_cust, df_no_sheet], ignore_index=True)
            combined['提單號碼'] = combined['提單號碼'].str.strip()
            combined = combined[combined['提單號碼'] != '']

            # B. 自動統計正報人件數
            stats_positive = {}
            current_sender = None
            for i, row in combined.iterrows():
                val_a, val_p = str(row['報關']).strip(), str(row['寄件人']).strip()
                if "正式報關" in val_a or "合併正報" in val_a:
                    current_sender = val_p if val_p != "" else "未知"
                    if current_sender not in stats_positive: stats_positive[current_sender] = 0
                if current_sender and "不報關" not in val_a: stats_positive[current_sender] += 1
                elif "不報關" in val_a: current_sender = None

            pos_string = "、".join([f"{name} {count} 件" for name, count in stats_positive.items()])
            count_no = len(df_no_sheet)
            total_count = len(combined)

            # C. 跨表比對收件人
            def lookup_info(row):
                hawb = str(row['提單號碼']).strip()
                if hawb in search_db.index:
                    info = search_db.loc[hawb]
                    if isinstance(info, pd.DataFrame): info = info.iloc[0]
                    return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                return pd.Series(["", "", "", ""])
            combined[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined.apply(lookup_info, axis=1)

            # D. 下載檔案產出
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_cols = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
                combined[final_cols].to_excel(writer, sheet_name='出口總明細', index=False)
            
            today_str = datetime.now().strftime("%Y%m%d")
            st.balloons()
            st.success("✅ 處理完成！")
            st.download_button(label="📥 下載轉換後的信天翁檔案", data=output.getvalue(), file_name=f"{today_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

            # --- E. Gmail 固定格式設定 ---
            recipients = "twnalex2009@gmail.com,twnalex24471640.01@gmail.com"
            cc_list = "gmcs@goodmaji.com,gmop@goodmaji.com,gmfa@goodmaji.com,bdm@goodmaji.com"
            subject = f"{today_str} 信天翁 to MO (出口明細)"
            
            email_body = (
                f"Dears\n\n"
                f"今日出口明細如附檔，共 {total_count} 件\n"
                f"請再協助申報，並安排出口，謝謝\n\n"
                f"正報：{pos_string}\n"
                f"不報關： {count_no} 件"
            )
            
            # 使用 urllib.parse.quote 處理特殊字元
            mailto_url = f"https://mail.google.com/mail/?view=cm&fs=1" \
                         f"&to={recipients}" \
                         f"&cc={cc_list}" \
                         f"&su={urllib.parse.quote(subject)}" \
                         f"&body={urllib.parse.quote(email_body)}"
            
            st.markdown(f'<a href="{mailto_url}" target="_blank" class="email-btn">📧 開啟 Gmail (已自動填妥收件人與內容)</a>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"發生錯誤: {e}")
