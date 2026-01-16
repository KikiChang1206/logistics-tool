import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime

# 設定網頁小縮圖與版面
st.set_page_config(page_title="信天翁", layout="centered")

# 使用自定義 CSS 縮小標題字體與美化版面
st.markdown("""
    <style>
    .main-title { font-size: 24px; font-weight: bold; margin-bottom: 20px; }
    .status-box { padding: 10px; border-radius: 5px; margin-bottom: 5px; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="main-title">信天翁 自動轉換器</p>', unsafe_allow_html=True)

# 檔案上傳區
uploaded_files = st.file_uploader("請上傳文件", accept_multiple_files=True, help="請同時選取『一般』與『聯郵』檔案上傳")

# 狀態變數
has_general = None
has_lian_yu = None
gen_file_data = None
lian_file_data = None

# 檔案辨識邏輯
if uploaded_files:
    for f in uploaded_files:
        if "一般" in f.name:
            has_general = True
            gen_file_data = f
        elif "聯郵" in f.name:
            has_lian_yu = True
            lian_file_data = f

# 顯示檢查清單
st.write("---")
c1, c2 = st.columns(2)
with c1:
    if has_general:
        st.write("✅ 一般文件 已就緒")
    else:
        st.write("⬜ 待上傳一般文件")

with c2:
    if has_lian_yu:
        st.write("✅ 聯郵文件 已就緒")
    else:
        st.write("⬜ 待上傳聯郵文件")

st.write("---")

# 只有當兩者都存在時才顯示轉換按鈕
if has_general and has_lian_yu:
    if st.button("信天翁文件產出", type="primary"):
        try:
            # --- 1. 處理【一般】文件 ---
            df_gen_raw = pd.read_excel(gen_file_data, dtype=str)
            gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", 
                           "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", 
                           "VALUE (USD)", "BAG NO.", "SHORT NAME"]
            df_gen_raw.columns = gen_headers[:len(df_gen_raw.columns)]
            df_gen_raw['HAWB / CN'] = df_gen_raw['HAWB / CN'].astype(str).str.strip()
            df_gen_raw = df_gen_raw[df_gen_raw['HAWB / CN'] != 'nan']
            search_db = df_gen_raw.set_index('HAWB / CN')

            # --- 2. 處理【聯郵】文件 ---
            dtype_dict = {'單價(TWD)': str, '寄件公司/統編': str, '統計方式': str, '提單號碼': str}
            df_customs = pd.read_excel(lian_file_data, sheet_name='報關明細', dtype=dtype_dict)
            df_no_customs = pd.read_excel(lian_file_data, sheet_name='不報關-X7明細', dtype=dtype_dict)
            
            def format_price(val):
                try: return "{:.2f}".format(float(val))
                except: return val

            df_customs['單價(TWD)'] = df_customs['單價(TWD)'].apply(format_price)
            df_no_customs['單價(TWD)'] = df_no_customs['單價(TWD)'].apply(format_price)

            df_no_customs.iloc[:, 0] = "" 
            if not df_no_customs.empty:
                df_no_customs.iloc[0, 0] = "不報關"
            
            combined_lian_yu = pd.concat([df_customs, df_no_customs], ignore_index=True)
            combined_lian_yu['提單號碼'] = combined_lian_yu['提單號碼'].astype(str).str.strip()
            combined_lian_yu = combined_lian_yu[combined_lian_yu['提單號碼'] != 'nan']

            # --- 3. 比對邏輯 ---
            def lookup_info(row):
                hawb = row['提單號碼']
                if hawb in search_db.index:
                    info = search_db.loc[hawb]
                    if isinstance(info, pd.DataFrame): info = info.iloc[0]
                    return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
                return pd.Series(["", "", "", ""])

            combined_lian_yu[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined_lian_yu.apply(lookup_info, axis=1)

            # --- 4. 格式化輸出 A~X ---
            final_cols_x = ['報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]
            df_final_export = pd.DataFrame(columns=final_cols_x)
            df_final_export = pd.concat([df_final_export, combined_lian_yu], join='outer').fillna('')[final_cols_x]

            # --- 5. 輸出設定 ---
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
            st.success("✅ 處理完成！")
            st.download_button(label="📥 下載轉換結果", data=output.getvalue(), file_name=f"{today_str}_信天翁 TO MO_Manifest.xlsx")

        except Exception as e:
            st.error(f"發生錯誤: {e}")
else:
    st.info("請同時上傳含有『一般』與『聯郵』字樣的檔案以解鎖按鈕。")
