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

                # ★★★ 修正處 1：欄位名稱去除前後空白，避免因空白差異抓不到欄位 ★★★
                df_c.columns = df_c.columns.str.strip()
                df_n.columns = df_n.columns.str.strip()

                # ★★★ 修正處 2：若「不報關-X7明細」分頁是空的、或缺少必要欄位，
                # 就補上一個空的「提單號碼」欄位，讓後面程式碼可以正常運作，而不是報錯 ★★★
                required_cols_n = ['提單號碼', '報關', '寄件人']
                for col in required_cols_n:
                    if col not in df_n.columns:
                        df_n[col] = ''

                # C. 全新精確統計邏輯 (逐行嚴格檢查，解決吃行問題)
                def get_stats_v2(df, pos_keys, sim_keys):
                    pos_info, sim_info = {}, {}
                    current_dict = None
                    current_sender = None

                    for i in range(len(df)):
                        row = df.iloc[i]
                        hawb = str(row['提單號碼']).strip()
                        if hawb == "":
                            continue  # 略過完全空白行

                        type_val = str(row['報關']).strip()
                        sender_val = str(row['寄件人']).strip()

                        if type_val != "":
                            # 判斷是正報還簡報
                            is_pos = any(k in type_val for k in pos_keys)
                            is_sim = any(k in type_val for k in sim_keys)

                            if is_pos or is_sim:
                                current_dict = pos_info if is_pos else sim_info

                        # 若第一行寄件人空白，往下找到同區塊的第一個寄件人
                        if sender_val == "" and current_dict is not None:
                            for j in range(i, len(df)):
                                if str(df.iloc[j]['提單號碼']).strip() == "": break
                                if str(df.iloc[j]['寄件人']).strip() != "":
                                    sender_val = str(df.iloc[j]['寄件人']).strip()
                                    break

                        # 更新寄件人 (過濾多餘字眼)
                        if sender_val != "":
                            current_sender = sender_val.replace("股份有限公司","").replace("有限公司","").replace("生醫國際","").replace("國際開發股份","").replace("開發股份","").replace("國際","").strip()

                        # 累加件數與紀錄首筆單號
                        if current_dict is not None and current_sender is not None:
                            if current_sender not in current_dict:
                                current_dict[current_sender] = {"count": 0, "first_hawb": hawb}
                            current_dict[current_sender]["count"] += 1

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

                # ★★★ 修正處 3：若聯郵檔案裡缺少 final_cols 中的某些欄位，先補空欄，避免選取欄位時出錯 ★★★
                for col in final_cols:
                    if col not in combined.columns:
                        combined[col] = ''

                # ★★★ 新功能 1：檢查「品名」「中文品名」是否有缺漏 ★★★
                missing_mask = (combined['品名'].astype(str).str.strip() == '') | (combined['中文品名'].astype(str).str.strip() == '')
                missing_hawb_list = combined.loc[missing_mask, '提單號碼'].astype(str).str.strip().tolist()
                missing_hawb_list = [h for h in missing_hawb_list if h != '']

                spaced_rows = []
                last_type = None
                last_sender = None

                # 全新斷行邏輯：寄件人改變，或是報關類型改變，就強制斷行
                for _, row in combined.iterrows():
                    curr_type = str(row['報關']).strip()
                    curr_sender = str(row['寄件人']).strip()
                    if curr_sender == "" and last_sender is not None:
                        curr_sender = last_sender # 繼承同區塊寄件人

                    if len(spaced_rows) > 0:
                        is_new_group = False
                        if curr_type != "" and last_type is not None and curr_type != last_type:
                            is_new_group = True
                        if curr_sender != "" and last_sender is not None and curr_sender != last_sender:
                            is_new_group = True

                        if is_new_group:
                            if curr_type == "不報關" and last_type == "不報關":
                                pass # 連續的不報關不強制斷行
                            else:
                                spaced_rows.append(pd.Series([None] * len(final_cols), index=final_cols))

                    disp = row.copy()
                    if curr_type == "不報關" and last_type == "不報關":
                        disp['報關'] = ""

                    spaced_rows.append(disp)

                    if curr_type != "": last_type = curr_type
                    if curr_sender != "": last_sender = curr_sender

                df_final = pd.DataFrame(spaced_rows).fillna('')[final_cols]

                out = BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as writer:
                    df_final.to_excel(writer, sheet_name='出口總明細', index=False)
                    ws = writer.sheets['出口總明細']
                    for r_idx, row in enumerate(ws.iter_rows()):
                        for cell in row:
                            cell.font = Font(name='Arial', size=10)
                            cell.border = Border() # 徹底移除框線
                            if r_idx == 0:
                                cell.alignment = Alignment(horizontal='left')

                st.session_state.processed = True
                st.success(f"✅ 處理完成！日期：{today_str}")
                st.download_button("📥 下載檔案 (無框線版)", out.getvalue(), f"{today_str}_信天翁 TO MO_Manifest.xlsx", use_container_width=True)

                # ★★★ 新功能 1：品名 / 中文品名 缺漏提醒 ★★★
                if missing_hawb_list:
                    st.warning(f"⚠️ 品名資料有缺少，請補資料\n\n缺漏的提單號碼：{'、'.join(missing_hawb_list)}")

                # E. Gmail 範本產出區
                st.write("---")
                st.write("### 📧 Gmail 草稿清單")

                to_all = "twnalex2009@gmail.com,twnalex24471640.01@gmail.com"
                cc_all = "gmcs@goodmaji.com,gmop@goodmaji.com,gmfa@goodmaji.com,bdm@goodmaji.com"

                # 1. 總出口明細草稿
                sub_main = f"{today_str} 信天翁 to MO (出口明細)"
                total_count = len(combined[combined['提單號碼'].str.strip() != ''])
                # ★★★ 修正處 4：不報關件數改用「提單號碼」不為空白的行數計算，
                # 這樣即使 df_n 原本是空的，也會正確算出 0 件，而不是報錯 ★★★
                no_declare_count = len(df_n[df_n['提單號碼'].astype(str).str.strip() != ''])
                body_main = f"Dears\n\n今日出口明細如附檔，共 {total_count} 件\n請再協助申報，並安排出口，謝謝\n\n正報：{pos_sum_text}\n簡報：{sim_sum_text}\n不報關：{no_declare_count} 件"
                url_main = f"https://mail.google.com/mail/?view=cm&fs=1&to={to_all}&cc={cc_all}&su={urllib.parse.quote(sub_main)}&body={urllib.parse.quote(body_main)}"
                st.markdown(f'<a href="{url_main}" target="_blank" class="email-btn">📧 1. 總出口明細草稿</a>', unsafe_allow_html=True)

                # 2. 個別廠商報關草稿 (含首筆單號)
                all_brand_info = {**stats_pos, **stats_sim}
                if all_brand_info:
                    st.write("#### 報關文件草稿：")
                    cc_sub = "gmop@goodmaji.com"
                    for idx, (brand, data) in enumerate(sorted(all_brand_info.items()), 2):
                        sub_brand = f"{today_str} 信天翁 to MO ( {brand} 文件)"
                        body_brand = f"Dears,\n\n{data['first_hawb']}\n{brand}報關文件如附檔，請您協助申報，感恩"

                        # ★★★ 新功能 2 & 3：特定廠商自動加註 ★★★
                        if any(k in brand for k in ["蜜凱", "綺麗絲"]):
                            body_brand += "\n1元做銷售02"
                        elif "大研" in brand:
                            body_brand += "\n1元做贈品04"

                        url_brand = f"https://mail.google.com/mail/?view=cm&fs=1&to={to_all}&cc={cc_sub}&su={urllib.parse.quote(sub_brand)}&body={urllib.parse.quote(body_brand)}"
                        st.markdown(f'<a href="{url_brand}" target="_blank" class="email-btn-sub">📩 {idx}. 報關草稿：{brand}</a>', unsafe_allow_html=True)

        except Exception as e: st.error(f"錯誤: {e}")
