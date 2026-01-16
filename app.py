import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="信天翁物流轉換器", layout="wide")
st.title("📦 信天翁 TO MO_Manifest 自動轉換系統")

lian_yu = st.file_uploader("1. 上傳【聯郵】檔案", type=["xls", "xlsx"])
general = st.file_uploader("2. 上傳【一般】檔案", type=["xls", "xlsx"])

if lian_yu and general:
    try:
        # --- 處理【聯郵】邏輯 ---
        # 讀取標籤
        df_customs = pd.read_excel(lian_yu, sheet_name='報關明細')
        df_no_customs = pd.read_excel(lian_yu, sheet_name='不報關-X7明細')
        
        # A欄強制補上「不報關」
        # 假設 A 欄是第一欄，我們強制覆蓋這一欄的資料
        df_no_customs.iloc[:, 0] = "不報關"
        
        # 合併兩份資料
        combined_lian_yu = pd.concat([df_customs, df_no_customs], ignore_index=True)
        
        # 定義信天翁要求的 A~X 標題 (共 24 欄)
        final_cols_x = [
            '報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', 
            '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', 
            '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', 
            "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"
        ]
        
        # 建立一個空的表格，確保所有 A~X 欄位都存在
        df_final_export = pd.DataFrame(columns=final_cols_x)
        
        # 將合併後的資料填入對應的欄位 (這裡會根據欄位名稱自動對齊)
        # 如果原始資料的欄位名稱與 final_cols_x 一模一樣，就會自動對上
        df_final_export = pd.concat([df_final_export, combined_lian_yu], join='outer').fillna('')
        df_final_export = df_final_export[final_cols_x] # 確保順序是 A 到 X

        # --- 處理【一般】邏輯 ---
        df_gen_raw = pd.read_excel(general)
        # 定義一般文件的 A~N 標題
        gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", 
                       "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", 
                       "VALUE (USD)", "BAG NO.", "SHORT NAME"]
        
        # 強制覆蓋標題列
        df_gen_final = df_gen_raw.copy()
        df_gen_final.columns = gen_headers[:len(df_gen_final.columns)]

        # --- 產出檔案 ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_export.to_excel(writer, sheet_name='出口總明細', index=False)
            df_gen_final.to_excel(writer, sheet_name='一般', index=False)
        
        st.success("✅ 格式整理完成！")
        st.download_button(label="📥 下載 整理後的信天翁檔案", data=output.getvalue(), file_name="信天翁_整理完成.xlsx")

    except Exception as e:
        st.error(f"發生錯誤: {e}")
