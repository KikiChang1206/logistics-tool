import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime

st.set_page_config(page_title="信天翁物流轉換器", layout="wide")
st.title("📦 信天翁 TO MO_Manifest 自動轉換系統")

lian_yu = st.file_uploader("1. 上傳【聯郵】檔案", type=["xls", "xlsx"])
general = st.file_uploader("2. 上傳【一般】檔案", type=["xls", "xlsx"])

if lian_yu and general:
    try:
        # --- 1. 處理【一般】文件 ---
        # 指定特定欄位為字串，防止數字變形
        df_gen_raw = pd.read_excel(general, dtype={'PostCode': str, 'CONSIGNEE\'S TEL': str})
        gen_headers = ["NO.", "HAWB / CN", "Marking", "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", 
                       "PostCode", "COD", "CONSIGNEE'S TEL", "PCS", "WT (KG)", "DESCRIPTION", 
                       "VALUE (USD)", "BAG NO.", "SHORT NAME"]
        df_gen_raw.columns = gen_headers[:len(df_gen_raw.columns)]
        df_gen_raw['HAWB / CN'] = df_gen_raw['HAWB / CN'].astype(str).str.strip()
        search_db = df_gen_raw.set_index('HAWB / CN')

        # --- 2. 處理【聯郵】文件 ---
        # 強制指定 N欄(單價) 和 S欄(統計方式) 為字串
        df_customs = pd.read_excel(lian_yu, sheet_name='報關明細', dtype={'單價(TWD)': str, '統計方式': str})
        df_no_customs = pd.read_excel(lian_yu, sheet_name='不報關-X7明細', dtype={'單價(TWD)': str, '統計方式': str})
        
        df_no_customs.iloc[:, 0] = "" 
        if not df_no_customs.empty:
            df_no_customs.iloc[0, 0] = "不報關"
        
        combined_lian_yu = pd.concat([df_customs, df_no_customs], ignore_index=True)
        combined_lian_yu['提單號碼'] = combined_lian_yu['提單號碼'].astype(str).str.strip()

        # --- 3. 執行「跨表比對」 ---
        def lookup_info(row):
            hawb = row['提單號碼']
            if hawb in search_db.index:
                info = search_db.loc[hawb]
                if isinstance(info, pd.DataFrame):
                    info = info.iloc[0]
                return pd.Series([info["CONSIGNEE'S NAME"], info["CONSIGNEE'S ADDRESS"], info["PostCode"], info["CONSIGNEE'S TEL"]])
            return pd.Series(["", "", "", ""])

        combined_lian_yu[["CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"]] = combined_lian_yu.apply(lookup_info, axis=1)

        # --- 4. 格式化輸出 A~X ---
        final_cols_x = [
            '報關', '好馬吉袋號', '袋號', '編號', '提單號碼', '發票號碼', '件數', 
            '提單重量(KG)', '品名', '中文品名', '數量', '單位', '產地', '單價(TWD)', 
            '寄件公司/統編', '寄件人', '電話', '寄件人地址', '統計方式', '商標', 
            "CONSIGNEE'S NAME", "CONSIGNEE'S ADDRESS", "PostCode", "CONSIGNEE'S TEL"
        ]
        
        df_final_export = pd.DataFrame(columns=final_cols_x)
        df_final_export = pd.concat([df_final_export, combined_lian_yu], join='outer').fillna('')
        df_final_export = df_final_export[final_cols_x]

        # --- 5. 產出檔案與樣式設定 ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_export.to_excel(writer, sheet_name='出口總明細', index=False)
            df_gen_raw.to_excel(writer, sheet_name='一般', index=False)
            
            for sheetname in ['出口總明細', '一般']:
                worksheet = writer.sheets[sheetname]
                no_border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
                arial_font = Font(name='Arial', size=10)
                right_align = Alignment(horizontal='right')
                
                for r_idx, row in enumerate(worksheet.iter_rows()):
                    for cell in row:
                        cell.font = arial_font
                        cell.border = no_border
                        # 如果是第一列(標題)，設定靠右對齊
                        if r_idx == 0:
                            cell.alignment = right_align

        # 取得今天日期作為檔名
        today_str = datetime.now().strftime("%Y%m%d")
        final_filename = f"{today_str}_信天翁 TO MO_Manifest.xlsx"

        st.success("✅ 處理完成！")
        st.download_button(label="📥 下載轉換後的檔案", data=output.getvalue(), file_name=final_filename)

    except Exception as e:
        st.error(f"發生錯誤: {e}")
