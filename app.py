import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime, timedelta
import urllib.parse

# 1. 網頁基本設定 (維持黑底風格)
st.set_page_config(page_title="信天翁系統", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 36px !important; font-weight: bold; color: #FFFFFF !important; }
    .stFileUploader section { background-color: #FFFFFF !important; padding: 40px !important; border: 2px dashed #FFFFFF !important; border-radius: 10px; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #000000 !important; height: 50px; font-weight: bold; width: 100%; }
    .email-btn { display: inline-block; width: 100%; text-align: center; background-color: #FFFFFF; color: #000000 !important; border: 2px solid #28A745; padding: 12px; font-weight: bold; text-decoration: none; border-radius: 5px; margin-top: 10px; }
    h3 { color: #FFFFFF !important; }
    .stMarkdown p, .stMarkdown span, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換器</p>', unsafe_allow_html=True)

# 2. 檔案上傳區
uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True)

has_gen = has_lian = False
gen_file = lian_file = None
if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name: has_gen, gen_file = True, f
        elif "聯郵" in f.name: has_lian, lian_file = True, f

# 3. 處理邏輯
if has_gen and has_lian:
    if 'processed' not in st.session_state: st.session_state.processed = False

    if st.button("🚀 信天翁文件產出", use_container_width=True) or st.session_state.processed:
        try:
            with st.spinner('正在精確統計件數...'):
                # 取得台灣時間 (UTC+8)
                tw_now = datetime.utcnow() + timedelta(hours=8)
                today_str = tw_now.strftime("%Y%m%d")

                # A. 讀取一般文件
                df_gen = pd.read_excel(gen_file, dtype=str).fillna('')
                df_gen.columns = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"][:len(df_gen.columns)]
                search_db = df_gen.set_index('HAWB / CN')

                # B. 讀取聯郵檔案 (原始處理，不使用 ffill)
                df_c = pd.read_excel(lian_file, sheet_name='報關明細', dtype=str).fillna('')
                df_n = pd.read_excel(lian_file, sheet_name='不報關-X7明細', dtype=str).fillna('')

                # C. 【精確邏輯修正】：到空格前停止統計
                def get_stats_v2(df, pos_keys, sim_keys):
                    pos_counts, sim_counts = {}, {}
                    i = 0
                    max_len = len(df)
                    while i < max_len:
                        val_a = str(df.iloc[i]['報關']).strip()
                        # 檢查是否為 正報 或 簡報 的起始點
                        is_pos = any(k in val_a for k in pos_keys)
                        is_sim = any(k in val_a for k in sim_keys)
                        
                        if is_pos or is_sim:
                            # 鎖定當前這一組的寄件人名稱 (P欄)
                            sender = str(df.iloc[i]['寄件人']).strip()
                            if sender == "": # 如果第一行沒寫名字，往下找第一個有名字的
                                for j in range(i, max_len):
                                    if str(df.iloc[j]['寄件人']).strip() != "":
                                        sender = str(df.iloc[j]['寄件人']).strip()
                                        break
                            
                            short_sender = sender.replace("有限公司","").replace("股份有限公司","").replace("生醫國際","")
                            count = 0
                            
                            # 從當前行往下數，直到遇到第一個「提單號碼為空」的行(空格)為止
                            while i < max_len:
                                if str(df.iloc[i]['提單號碼']).strip() == "": # 遇到空格停止
                                    break
                                count += 1
                                i += 1
                            
                            # 存入統計
                            if is_pos: pos_counts[short_sender] = pos_counts.get(short_sender, 0) + count
                            if is_sim: sim_counts[short_sender] = sim_counts.get(short_sender, 0) + count
                        else:
                            i += 1
                    return pos_counts, sim_counts

                stats_pos, stats_sim = get_stats_v2(df_c, ["正式報關", "合併正報"], ["簡易報關", "合併簡報"])
                pos_text = "、".join([f"{n} {c}件" for n, c in stats_pos.items()]) if stats_pos else "無"
                sim_text = "、".join([f"{n} {c}件" for n, c in stats_sim.items()]) if stats_sim else "無"

                # D. 產出原始正確格式之 Excel
                df_n['報關'] = "不報關"
                combined = pd.concat([df_c, df_n], ignore_index=True)
                combined = combined[combined['提單號碼'].str.strip() != '']

                def lookup(r):
                    h = str(r['提單號碼']).strip()
                    if h in search_db.index:
                        info = search_db.loc[h]
                        if isinstance(info, pd.DataFrame): info = info.iloc[0]
                        return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                    return pd.Series([""]*4)
                combined[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined.apply(lookup, axis=1)

                # 插入分組空行邏輯 (原始格式)
                final_cols = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
                spaced_rows, last_type = [], None
                for _, row in combined.iterrows():
                    curr = str(row['報關']).strip()
                    if last_type is not None and curr != last_type and curr != "":
                        spaced_rows.append(pd.Series([None] * len(final_cols), index=final_cols))
                    disp = row.copy()
                    if curr == "不報關" and last_type == "不報關": disp['報關'] = ""
                    spaced_rows.append(disp)
                    last_type = curr
                df_final = pd.DataFrame(spaced_rows).fillna('')[final_cols]

                out = BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as writer:
                    df_final.to_excel(writer, sheet_name='出口總明細', index=False)
                    ws = writer.sheets['出口總明細']
                    for r_idx, row in enumerate(ws.iter_rows()):
                        for cell in row:
                            cell.font = Font(name='Arial', size=10)
                            cell.border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
                            if r_idx == 0: cell.alignment = Alignment(horizontal='left')

                st.session_state.processed = True
                st.success(f"✅ 處理完成！日期：{today_str}")
                st.download_button("📥 下載轉換後的信天翁檔案", out.getvalue(), f"{today_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

                # E. Gmail 範本 (修正件數)
                to = "twnalex2009@gmail.com,twnalex24471640.01@gmail.com"
                cc = "gmcs@goodmaji.com,gmop@goodmaji.com,gmfa@goodmaji.com,bdm@goodmaji.com"
                sub = f"{today_str} 信天翁 to MO (出口明細)"
                total_count = len(combined)
                body = f"Dears\n\n今日出口明細如附檔，共 {total_count} 件\n請再協助申報，並安排出口，謝謝\n\n正報：{pos_text}\n簡報：{sim_text}\n不報關：{len(df_n[df_n['提單號碼'].str.strip() != ''])} 件"
                url = f"https://mail.google.com/mail/?view=cm&fs=1&to={to}&cc={cc}&su={urllib.parse.quote(sub)}&body={urllib.parse.quote(body)}"
                st.markdown(f'<a href="{url}" target="_blank" class="email-btn">📧 開啟 Gmail (大研 {len(stats_sim.get("大研", [0])) if isinstance(stats_sim.get("大研"), list) else stats_sim.get("大研", 0)}件)</a>', unsafe_allow_html=True)

        except Exception as e: st.error(f"錯誤: {e}")
