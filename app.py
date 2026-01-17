import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, Border, Side
from datetime import datetime
import urllib.parse

# 1. 網頁基本設定
st.set_page_config(page_title="信天翁系統", layout="centered")

# 自定義 CSS
st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 36px !important; font-weight: bold; color: #FFFFFF !important; }
    .stFileUploader section { background-color: #FFFFFF !important; padding: 40px !important; border: 2px dashed #333 !important; border-radius: 10px; }
    button[data-testid="baseButton-secondary"] { background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #000000 !important; }
    button[data-testid="baseButton-secondary"] p { color: #000000 !important; }
    div[data-testid="stNotification"], div[data-testid="stNotificationV2"] { background-color: transparent !important; border: none !important; }
    div[data-testid="stNotification"] p, div[data-testid="stNotificationV2"] p { color: #FFFFFF !important; font-size: 16px !important; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #000000 !important; height: 50px; font-weight: bold; width: 100%; }
    div.stButton > button p { color: #000000 !important; }
    .email-btn { display: inline-block; width: 100%; text-align: center; background-color: #FFFFFF; color: #000000 !important; border: 2px solid #28A745; padding: 12px; font-weight: bold; text-decoration: none; border-radius: 5px; margin-top: 10px; }
    h3 { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換器</p>', unsafe_allow_html=True)

# 2. 檔案上傳區
uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True)

has_gen = has_lian = False
gen_f = lian_f = None
if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name: has_gen, gen_f = True, f
        elif "聯郵" in f.name: has_lian, lian_f = True, f

st.write("### 📁 檔案狀態確認")
c1, c2 = st.columns(2)
with c1:
    if has_gen: st.success("✅ 一般文件：已就緒")
    else: st.info("⬜ 一般文件：待上傳")
with c2:
    if has_lian: st.success("✅ 聯郵文件：已就緒")
    else: st.info("⬜ 聯郵文件：待上傳")

# 4. 處理邏輯
if has_gen and has_lian:
    st.write("---")
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    
    if st.button("🚀 信天翁文件產出", use_container_width=True) or st.session_state.processed:
        try:
            with st.spinner('分析中...'):
                # 取得當前日期 (20260118)
                t_str = datetime.now().strftime("%Y%m%d")

                # A. 讀取資料
                df_g = pd.read_excel(gen_f, dtype=str).fillna('')
                df_g.columns = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"][:len(df_g.columns)]
                db = df_g.set_index('HAWB / CN')

                df_c = pd.read_excel(lian_f, sheet_name='報關明細', dtype=str).fillna('')
                df_n = pd.read_excel(lian_f, sheet_name='不報關-X7明細', dtype=str).fillna('')
                
                # B. 統計邏輯 (在合併與插入空行前統計，避免算錯)
                def count_by_category(df, keywords):
                    counts = {}
                    current_s = None
                    # 只處理提單號碼不為空的有效行
                    valid_df = df[df['提單號碼'].str.strip() != ''].copy()
                    for _, r in valid_df.iterrows():
                        a, p = str(r['報關']).strip(), str(r['寄件人']).strip()
                        if any(k in a for k in keywords):
                            current_s = p if p != "" else "未知"
                            counts[current_s] = counts.get(current_s, 0) + 1
                    return counts

                # 正報關鍵字：正式報關、合併正報
                stats_pos = count_by_category(df_c, ["正式報關", "合併正報"])
                # 簡報關鍵字：簡易報關、合併簡報
                stats_sim = count_by_category(df_c, ["簡易報關", "合併簡報"])
                
                pos_text = "、".join([f"{n} {c}件" for n, c in stats_pos.items()]) if stats_pos else "無"
                sim_text = "、".join([f"{n} {c}件" for n, c in stats_sim.items()]) if stats_sim else "無"
                
                # C. 合併與比對
                df_n['報關'] = "不報關"
                comb = pd.concat([df_c, df_n], ignore_index=True)
                comb['提單號碼'] = comb['提單號碼'].str.strip()
                comb = comb[comb['提單號碼'] != '']

                def lookup(r):
                    h = str(r['提單號碼']).strip()
                    if h in db.index:
                        i = db.loc[h]
                        if isinstance(i, pd.DataFrame): i = i.iloc[0]
                        return pd.Series([i["CONSIGNEE'S NAME"], i["CONSIGNEE'S ADDRESS"], i["PostCode"], i["CONSIGNEE'S TEL"]])
                    return pd.Series([""]*4)
                
                comb[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = comb.apply(lookup, axis=1)

                # D. 插入分組空行
                final_cols = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
                spaced_rows = []
                last_type = None
                for _, row in comb.iterrows():
                    curr_type = str(row['報關']).strip()
                    if last_type is not None and curr_type != last_type and curr_type != "":
                        spaced_rows.append(pd.Series([None] * len(final_cols), index=final_cols))
                    display_row = row.copy()
                    if curr_type == "不報關" and last_type == "不報關": display_row['報關'] = ""
                    spaced_rows.append(display_row)
                    last_type = curr_type

                df_final = pd.DataFrame(spaced_rows).fillna('')[final_cols]

                # E. 產出 Excel (移除所有框線)
                no_border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
                out = BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as writer:
                    df_final.to_excel(writer, sheet_name='出口總明細', index=False)
                    ws = writer.sheets['出口總明細']
                    for row in ws.iter_rows():
                        for cell in row:
                            cell.font = Font(name='Arial', size=10)
                            cell.border = no_border
                            cell.alignment = Alignment(horizontal='left')

                if not st.session_state.processed:
                    st.balloons()
                    st.session_state.processed = True
                
                st.success(f"✅ 處理完成！日期：{t_str}")
                st.download_button("📥 下載轉換後的信天翁檔案", out.getvalue(), f"{t_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

                # F. Gmail 範本 (自動統計)
                to = "twnalex2009@gmail.com,twnalex24471640.01@gmail.com"
                cc = "gmcs@goodmaji.com,gmop@goodmaji.com,gmfa@goodmaji.com,bdm@goodmaji.com"
                sub = f"{t_str} 信天翁 to MO (出口明細)"
                total_valid = len(df_c[df_c['提單號碼'].str.strip() != '']) + len(df_n[df_n['提單號碼'].str.strip() != ''])
                
                body = (f"Dears\n\n今日出口明細如附檔，共 {total_count if 'total_count' in locals() else total_valid} 件\n"
                        f"請再協助申報，並安排出口，謝謝\n\n"
                        f"正報：{pos_text}\n"
                        f"簡報：{sim_text}\n"
                        f"不報關：{len(df_n[df_n['提單號碼'].str.strip() != ''])} 件")
                
                url = f"https://mail.google.com/mail/?view=cm&fs=1&to={to}&cc={cc}&su={urllib.parse.quote(sub)}&body={urllib.parse.quote(body)}"
                st.markdown(f'<a href="{url}" target="_blank" class="email-btn">📧 開啟 Gmail (自動填妥日期與件數)</a>', unsafe_allow_html=True)

        except Exception as e: st.error(f"發生錯誤: {e}")
