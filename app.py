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
    .email-btn-sub { display: inline-block; width: 100%; text-align: center; background-color: #E3F2FD; color: #1565C0 !important; border: 1px solid #1565C0; padding: 8px; font-weight: bold; text-decoration: none; border-radius: 5px; margin-top: 5px; font-size: 14px; }
    h3, h4 { color: #FFFFFF !important; }
    .stMarkdown p, .stMarkdown span, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🐦 信天翁 自動轉換</p>', unsafe_allow_html=True)

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
            with st.spinner('正在產出文件...'):
                tw_now = datetime.utcnow() + timedelta(hours=8)
                today_str = tw_now.strftime("%Y%m%d")

                # A. 讀取一般文件
                df_gen = pd.read_excel(gen_file, dtype=str).fillna('')
                df_gen.columns = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", "VALUE (USD)", "BAG NO.", "SHORT NAME"][:len(df_gen.columns)]
                search_db = df_gen.set_index('HAWB / CN')

                # B. 讀取聯郵檔案
                df_c = pd.read_excel(lian_file, sheet_name='報關明細', dtype=str).fillna('')
                df_n = pd.read_excel(lian_file, sheet_name='不報關-X7明細', dtype=str).fillna('')

                # C. 統計與首筆提單號碼抓取邏輯
                def get_stats_v2(df, pos_keys, sim_keys):
                    pos_info, sim_info = {}, {}
                    i, max_len = 0, len(df)
                    while i < max_len:
                        val_a = str(df.iloc[i]['報關']).strip()
                        is_pos = any(k in val_a for k in pos_keys)
                        is_sim = any(k in val_a for k in sim_keys)
                        
                        if is_pos or is_sim:
                            sender = str(df.iloc[i]['寄件人']).strip()
                            if sender == "":
                                for j in range(i, max_len):
                                    if str(df.iloc[j]['寄件人']).strip() != "":
                                        sender = str(df.iloc[j]['寄件人']).strip()
                                        break
                            short_sender = sender.replace("有限公司","").replace("股份有限公司","").replace("生醫國際","")
                            
                            # 關鍵：抓取該區塊的第一筆提單號碼
                            first_hawb = str(df.iloc[i]['提單號碼']).strip()
                            
                            count = 0
                            while i < max_len:
                                if str(df.iloc[i]['提單號碼']).strip() == "": break
                                count += 1; i += 1
                            
                            target_dict = pos_info if is_pos else sim_info
                            if short_sender not in target_dict:
                                target_dict[short_sender] = {"count": 0, "first_hawb": first_hawb}
                            target_dict[short_sender]["count"] += count
                        else: i += 1
                    return pos_info, sim_info

                stats_pos, stats_sim = get_stats_v2(df_c, ["正式報關", "合併正報"], ["簡易報關", "合併簡報"])
                pos_sum_text = "、".join([f"{n} {d['count']}件" for n, d in stats_pos.items()]) if stats_pos else "無"
                sim_sum_text = "、".join([f"{n} {d['count']}件" for n, d in stats_sim.items()]) if stats_sim else "無"

                # D. 產出 Excel
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
                            # --- 核心修改：徹底移除框線 ---
                            cell.border = Border() 
                            # ---------------------------
                            if r_idx == 0:
                                cell.alignment = Alignment(horizontal='left')

                st.session_state.processed = True
                st.success(f"✅ 處理完成！日期：{today_str}")
                st.download_button("📥 下載檔案 ", out.getvalue(), f"{today_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

                # E. Gmail 範本產出區
                st.write("---")
                st.write("### 📧 Gmail 草稿清單")
                
                to_all = "twnalex2009@gmail.com,twnalex24471640.01@gmail.com"
                cc_all = "gmcs@goodmaji.com,gmop@goodmaji.com,gmfa@goodmaji.com,bdm@goodmaji.com"
                
                # 1. 總出口明細草稿
                sub_main = f"{today_str} 信天翁 to MO (出口明細)"
                total_count = len(combined[combined['提單號碼'].str.strip() != ''])
                body_main = f"Dears\n\n今日出口明細如附檔，共 {total_count} 件\n請再協助申報，並安排出口，謝謝\n\n正報：{pos_sum_text}\n簡報：{sim_sum_text}\n不報關：{len(df_n[df_n['提單號碼'].str.strip() != ''])} 件"
                url_main = f"https://mail.google.com/mail/?view=cm&fs=1&to={to_all}&cc={cc_all}&su={urllib.parse.quote(sub_main)}&body={urllib.parse.quote(body_main)}"
                st.markdown(f'<a href="{url_main}" target="_blank" class="email-btn">📧 1. 總出口明細草稿</a>', unsafe_allow_html=True)

                # 2. 個別廠商報關草稿 (含首筆單號)
                all_brand_info = {**stats_pos, **stats_sim}
                if all_brand_info:
                    st.write("#### 報關文件草稿：")
                    cc_sub = "gmop@goodmaji.com"
                    for idx, (brand, data) in enumerate(sorted(all_brand_info.items()), 2):
                        sub_brand = f"{today_str} 信天翁 to MO ( {brand} 文件)"
                        # 依照需求：內文第一行放首筆單號
                        body_brand = f"Dears,\n\n{data['first_hawb']}\n{brand}報關文件如附檔，請您協助申報，感恩"
                        url_brand = f"https://mail.google.com/mail/?view=cm&fs=1&to={to_all}&cc={cc_sub}&su={urllib.parse.quote(sub_brand)}&body={urllib.parse.quote(body_brand)}"
                        st.markdown(f'<a href="{url_brand}" target="_blank" class="email-btn-sub">📩 {idx}. 報關草稿：{brand}</a>', unsafe_allow_html=True)

        except Exception as e: st.error(f"錯誤: {e}")
