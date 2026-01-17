import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime, timedelta
import urllib.parse

# 1. 網頁基本設定
st.set_page_config(page_title="信天翁系統", layout="centered")

# 自定義 CSS (維持你要求的黑底風格與白底黑字按鈕)
st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 36px !important; font-weight: bold; color: #FFFFFF !important; }
    .sub-title { font-size: 16px; color: #CCCCCC !important; margin-bottom: 25px; }
    .stFileUploader section { background-color: #FFFFFF !important; padding: 40px !important; border: 2px dashed #FFFFFF !important; border-radius: 10px; }
    button[data-testid="baseButton-secondary"] { background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #000000 !important; }
    button[data-testid="baseButton-secondary"] p { color: #000000 !important; }
    div[data-testid="stNotification"], div[data-testid="stNotificationV2"] { background-color: transparent !important; background: none !important; border: none !important; box-shadow: none !important; padding: 0px !important; }
    div[data-testid="stNotification"] p, div[data-testid="stNotificationV2"] p { color: #FFFFFF !important; font-size: 16px !important; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #000000 !important; height: 50px; font-size: 18px; font-weight: bold; width: 100%; }
    div.stButton > button p { color: #000000 !important; }
    div.stButton > button:hover { background-color: #EEEEEE !important; }
    .email-btn { display: inline-block; width: 100%; text-align: center; background-color: #FFFFFF; color: #000000 !important; border: 2px solid #28A745; padding: 12px; font-weight: bold; text-decoration: none; border-radius: 5px; margin-top: 10px; }
    h3 { color: #FFFFFF !important; }
    .stMarkdown p, .stMarkdown span, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換器</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">請將「一般」與「聯郵」檔案拖入上方白色區域中</p>', unsafe_allow_html=True)

# 2. 檔案上傳區
uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True)

has_general = has_lian_yu = False
gen_file_data = lian_file_data = None

if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name: has_general, gen_file_data = True, f
        elif "聯郵" in f.name: has_lian_yu, lian_file_data = True, f

# 3. 狀態檢查清單
st.write("### 📁 檔案狀態確認")
col1, col2 = st.columns(2)
with col1:
    if has_general: st.success("✅ 一般文件：已就緒")
    else: st.info("⬜ 一般文件：待上傳")
with col2:
    if has_lian_yu: st.success("✅ 聯郵文件：已就緒")
    else: st.info("⬜ 聯郵文件：待上傳")

# 4. 產出邏輯
if has_general and has_lian_yu:
    st.write("---")
    # 使用 Session State 確保產出後 Gmail 按鈕不消失
    if 'processed' not in st.session_state: st.session_state.processed = False

    if st.button("🚀 信天翁文件產出", use_container_width=True) or st.session_state.processed:
        try:
            with st.spinner('正在分析報關種類並進行精確對位...'):
                # 強制台灣日期
                tw_now = datetime.utcnow() + timedelta(hours=8)
                today_str = tw_now.strftime("%Y%m%d")

                # --- A. 讀取一般文件 ---
                df_gen_raw = pd.read_excel(gen_file_data, dtype=str).fillna('')
                gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"]
                df_gen_raw.columns = gen_headers[:len(df_gen_raw.columns)]
                df_gen_raw['HAWB / CN'] = df_gen_raw['HAWB / CN'].str.strip()
                search_db = df_gen_raw.set_index('HAWB / CN')

                # --- B. 讀取聯郵文件 ---
                dtype_dict = {'單價(TWD)': str, '寄件公司/統編': str, '統計方式': str, '提單號碼': str}
                df_customs = pd.read_excel(lian_file_data, sheet_name='報關明細', dtype=dtype_dict)
                df_no_customs = pd.read_excel(lian_file_data, sheet_name='不報關-X7明細', dtype=dtype_dict).fillna('')
                
                # --- C. 精確統計件數 (修正大研問題) ---
                def get_stats(df, pos_keys, sim_keys):
                    temp = df.copy()
                    # 只有在提單號碼不為空時才填充，避免算到空白行
                    temp.loc[temp['提單號碼'].astype(str).str.strip() == '', ['報關', '寄件人']] = pd.NA
                    temp['報關'] = temp['報關'].ffill()
                    temp['寄件人'] = temp['寄件人'].ffill()
                    pos_c, sim_c = {}, {}
                    # 過濾有效行
                    valid_df = temp[temp['提單號碼'].astype(str).str.strip() != ''].copy()
                    for _, r in valid_df.iterrows():
                        a, p = str(r['報關']).strip(), str(r['寄件人']).strip().replace("有限公司","").replace("股份有限公司","")
                        if any(k in a for k in pos_keys): pos_c[p] = pos_c.get(p, 0) + 1
                        elif any(k in a for k in sim_keys): sim_c[p] = sim_c.get(p, 0) + 1
                    return pos_c, sim_c

                stats_pos, stats_sim = get_stats(df_customs, ["正式報關", "合併正報"], ["簡易報關", "合併簡報"])
                pos_text = "、".join([f"{n} {c}件" for n, c in stats_pos.items()]) if stats_pos else "無"
                sim_text = "、".join([f"{n} {c}件" for n, c in stats_sim.items()]) if stats_sim else "無"

                # --- D. 跨表比對 ---
                df_no_customs['報關'] = "不報關"
                combined = pd.concat([df_customs.fillna(''), df_no_customs], ignore_index=True)
                combined['提單號碼'] = combined['提單號碼'].str.strip()
                combined = combined[(combined['提單號碼'] != '') & (combined['提單號碼'] != 'nan')]

                def format_price(val):
                    try: return "{:.2f}".format(float(val)) if val and val != 'nan' else ""
                    except: return val
                combined['單價(TWD)'] = combined['單價(TWD)'].apply(format_price)

                def lookup_info(row):
                    hawb = row['提單號碼']
                    if hawb in search_db.index:
                        info = search_db.loc[hawb]
                        if isinstance(info, pd.DataFrame): info = info.iloc[0]
                        return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                    return pd.Series(["", "", "", ""])

                combined[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined.apply(lookup_info, axis=1)

                # --- E. 報關種類插入空行邏輯 (原始正確格式) ---
                final_cols_x = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
                spaced_rows, last_type = [], None
                for _, row in combined.iterrows():
                    curr_type = str(row['報關']).strip()
                    if last_type is not None and curr_type != last_type and curr_type != "":
                        spaced_rows.append(pd.Series([None] * len(final_cols_x), index=final_cols_x))
                    display_row = row.copy()
                    if curr_type == "不報關" and last_type == "不報關": display_row['報關'] = ""
                    spaced_rows.append(display_row)
                    last_type = curr_type

                df_final_export = pd.DataFrame(spaced_rows).fillna('')[final_cols_x]

                # --- F. 產出 Excel ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final_export.to_excel(writer, sheet_name='出口總明細', index=False)
                    ws = writer.sheets['出口總明細']
                    no_border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
                    for r_idx, row in enumerate(ws.iter_rows()):
                        for cell in row:
                            cell.font = Font(name='Arial', size=10)
                            cell.border = no_border
                            if r_idx == 0: cell.alignment = Alignment(horizontal='left')

                if not st.session_state.processed:
                    st.balloons()
                    st.session_state.processed = True

                st.success(f"✅ 處理完成！今日日期：{today_str}")
                st.download_button(label="📥 下載轉換後的信天翁檔案", data=output.getvalue(), file_name=f"{today_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

                # --- G. Gmail 範本 (自動填妥收件人與內容) ---
                to = "twnalex2009@gmail.com,twnalex24471640.01@gmail.com"
                cc = "gmcs@goodmaji.com,gmop@goodmaji.com,gmfa@goodmaji.com,bdm@goodmaji.com"
                sub = f"{today_str} 信天翁 to MO (出口明細)"
                total_valid = len(combined)
                body = f"Dears\n\n今日出口明細如附檔，共 {total_valid} 件\n請再協助申報，並安排出口，謝謝\n\n正報：{pos_text}\n簡報：{sim_text}\n不報關：{len(df_no_customs[df_no_customs['提單號碼'].astype(str).str.strip() != ''])} 件"
                url = f"https://mail.google.com/mail/?view=cm&fs=1&to={to}&cc={cc}&su={urllib.parse.quote(sub)}&body={urllib.parse.quote(body)}"
                st.markdown(f'<a href="{url}" target="_blank" class="email-btn">📧 開啟 Gmail (自動帶入日期與正確件數)</a>', unsafe_allow_html=True)

        except Exception as e: st.error(f"發生錯誤: {e}")
