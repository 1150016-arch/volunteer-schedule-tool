import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="義工排班自動化系統",
    page_icon="📅",
    layout="wide"
)

st.title("📅 義工排班自動化轉換工具")
st.markdown("將 **表格 A（雲端預排檔）** 與 **表格 B（登記發心名單）** 放入，即可一鍵生成完整的 **表格 C（班表與列印檔）**。")

# 側邊欄設定
st.sidebar.header("⚙️ 參數設定")
target_date = st.sidebar.text_input("排班日期 / Sheet 名稱", value="1021", help="例如 1021，需對應表格 A 中的工作表名稱")

# 檔案上傳區
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上傳表格 A (雲端預排檔)")
    file_a = st.file_uploader("選擇表格 A (.xlsx, .xls)", type=["xlsx", "xls"], key="file_a")

with col2:
    st.subheader("2. 上傳表格 B (登記發心名單)")
    file_b = st.file_uploader("選擇表格 B (.xlsx, .xls)", type=["xlsx", "xls"], key="file_b")

def load_clean_b(file):
    """讀取表格 B 並自動處理標題在第二列的問題"""
    df = pd.read_excel(file)
    if '義工編號' not in df.columns and '編號' not in df.columns:
        if len(df) > 0 and any(df.iloc[0].astype(str).str.contains('義工編號|編號|姓名')):
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
    return df

def process_schedule(df_a, df_b):
    # 1. 自動尋找表格 A 與 B 的編號欄位名稱
    a_id_col = '編號' if '編號' in df_a.columns else ('義工編號' if '義工編號' in df_a.columns else df_a.columns[1])
    b_id_col = '義工編號' if '義工編號' in df_b.columns else ('編號' if '編號' in df_b.columns else df_b.columns[0])

    # 2. 清洗表格 A：過濾掉空列，並確保編號唯一不重複
    df_a_clean = df_a.dropna(subset=[a_id_col]).copy()
    df_a_unique = df_a_clean.drop_duplicates(subset=[a_id_col]).set_index(a_id_col)

    # 3. 清洗表格 B：移除 SOP 指定不需留存的欄位
    drop_cols = [
        '大隊別', '小隊別', '防疫證書證號', '任務', '地點', 
        '檢定資格', '活動名稱', '護持開始時間', '護持結束時間'
    ]
    df_b_cleaned = df_b.drop(columns=[c for c in drop_cols if c in df_b.columns]).copy()

    # 4. VLOOKUP 比對：帶入 館、哨、任、特展
    df_b_cleaned['館'] = df_b_cleaned[b_id_col].map(df_a_unique['館別']) if '館別' in df_a_unique.columns else ''
    df_b_cleaned['哨'] = df_b_cleaned[b_id_col].map(df_a_unique['哨別']) if '哨別' in df_a_unique.columns else ''
    df_b_cleaned['任'] = df_b_cleaned[b_id_col].map(df_a_unique['任務']) if '任務' in df_a_unique.columns else ''
    df_b_cleaned['特展'] = df_b_cleaned[b_id_col].map(df_a_unique['特展定導']) if '特展定導' in df_a_unique.columns else ''

    # 5. 調整欄位順序：在「精舍」後加入 館、哨、任、特展
    base_cols = list(df_b_cleaned.columns)
    if '精舍' in base_cols:
        idx = base_cols.index('精舍') + 1
        for c in ['館', '哨', '任', '特展']:
            if c in base_cols:
                base_cols.remove(c)
        new_cols = base_cols[:idx] + ['館', '哨', '任', '特展'] + base_cols[idx:]
        df_b_cleaned = df_b_cleaned[new_cols]

    # 6. 隔離發心時間異常者（置底）
    if '發心起始時間' in df_b_cleaned.columns:
        df_normal = df_b_cleaned[df_b_cleaned['發心起始時間'].notna()]
        df_abnormal = df_b_cleaned[df_b_cleaned['發心起始時間'].isna()]
        df_b_cleaned = pd.concat([df_normal, df_abnormal], ignore_index=True)

    # 7. 建立「館員用」公式連動 Dataframe
    df_staff = pd.DataFrame(columns=df_b_cleaned.columns)
    
    # 填入 Excel 連動公式：例如 ='1021'!A2，代表自動連動日期 Sheet 的該儲存格
    for row_idx in range(len(df_b_cleaned)):
        row_formulas = []
        excel_row_num = row_idx + 2  # Excel 標題是第 1 列，資料從第 2 列開始
        for col_idx in range(len(df_b_cleaned.columns)):
            col_letter = get_column_letter(col_idx + 1)
            formula = f"='{target_date}'!{col_letter}{excel_row_num}"
            row_formulas.append(formula)
        df_staff.loc[row_idx] = row_formulas

    # 8. 輸出 Excel 檔
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: 排班日期 Sheet（靜態初始資料，供手動修改）
        df_b_cleaned.to_excel(writer, sheet_name=target_date, index=False)

        # Sheet 2: 館員用 Sheet（帶入動態連動公式）
        df_staff.to_excel(writer, sheet_name="館員用", index=False)

    output.seek(0)
    return output

if file_a and file_b:
    try:
        xls_a = pd.ExcelFile(file_a)
        
        if target_date not in xls_a.sheet_names:
            st.error(f"⚠️ 在表格 A 中找不到名稱為 **{target_date}** 的工作表！請確認側邊欄的日期設定。")
            st.info(f"表格 A 包含的工作表：{', '.join(xls_a.sheet_names)}")
        else:
            df_a = pd.read_excel(xls_a, sheet_name=target_date)
            df_b = load_clean_b(file_b)

            st.success("✅ 檔案讀取成功！點擊下方按鈕開始排班轉換。")

            if st.button("🚀 開始自動排班並轉換表格 C", type="primary"):
                with st.spinner("正在處理中..."):
                    result_excel = process_schedule(df_a, df_b)
                
                st.balloons()
                st.download_button(
                    label="📥 下載完成的表格 C (.xlsx)",
                    data=result_excel,
                    file_name=f"表格C_排班完成_{target_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 處理過程發生錯誤：{str(e)}")
else:
    st.info("💡 請在上方同時上傳表格 A 與表格 B 進行轉換。")
