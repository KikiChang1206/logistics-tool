import streamlit as st
import pandas as pd
from io import BytesIO

# 設定網頁標題
st.set_page_config(page_title="信天翁物流轉換器", layout="wide")

st.title("📦 信天翁 TO MO_Manifest 自動轉換系統")

# 檔案上傳區
col1, col2 = st.columns(2)
with col1:
    lian_yu = st.file_uploader("請上傳【聯郵】檔案", type=["xls", "xlsx"])
with col2:
    general = st.file_uploader("請上傳【一般】檔案", type=["xls", "xlsx"])

if lian_yu and general:
    try:
        # 處理 聯郵
        df_customs = pd.read_excel(lian_yu, sheet_name='報關明細')
        df_no_customs = pd.read_excel(lian_yu, sheet_name='不報關-X7明細')
        df_no_customs.iloc[:, 0] = "不報關"
        export_total = pd.concat([df_customs, df_no_customs], ignore_index=True)
        
        # 處理 一般
        df_gen = pd.read_excel(general)
        gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", 
                       "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", 
                       "VALUE (USD)", "BAG NO.", "SHORT NAME"]
        df_gen.columns = gen_headers[:len(df_gen.columns)]

        # 產生 Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            export_total.to_excel(writer, sheet_name='出口總明細', index=False)
            df_gen.to_excel(writer, sheet_name='一般', index=False)
        
        st.success("✅ 檔案處理完成！")
        st.download_button(label="📥 點我下載整理後的檔案", data=output.getvalue(), file_name="信天翁 TO MO_Manifest(資料整理).xlsx")

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
