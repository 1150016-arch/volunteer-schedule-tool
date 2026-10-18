import streamlit as st
import pandas as pd
import io

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

def process_schedule(df_a, df_b):
    # 1. 依據 SOP 清洗表格 B，移除不需要的欄位
    drop_cols = [
        '大隊別', '小隊別', '防疫證書證號', '任務', '地點', 
        '檢定資格', '活動名稱', '護持開始時間', '護持結束時間'
    ]
    df_b_cleaned = df_b.drop(columns=[c for c in drop_cols if c in df_b.columns])

    # 2. 自動判斷欄位名稱 (相容 '編號' 與 '義工編號')
    a_id_col = '編號' if '編號' in df_a.columns else ('義工編號' if '義工編號' in df_a.columns else df_a.columns[1])
    b_id_col = '義工編號' if '義工編號' in df_b_cleaned.columns else ('編號' if '編號' in df_b_cleaned.columns else df_b_cleaned.columns[0])

    # VLOOKUP 匯入 <A> 資料
    a_indexed = df_a.set_index(a_id_col)
    
    # 欄位對應（防呆檢查）
    df_b_cleaned['館'] = df_b_cleaned[b_id_col].map(a_indexed['館別']) if '館別' in a_indexed.columns else ''
    df_b_cleaned['哨'] = df_b_cleaned[b_id_col].map(a_indexed['哨別']) if '哨別' in a_indexed.columns else ''
    df_b_cleaned['任'] = df_b_cleaned[b_id_col].map(a_indexed['任務']) if '任務' in a_indexed.columns else ''
    df_b_cleaned['特展'] = df_b_cleaned[b_id_col].map(a_indexed['特展定導']) if '特展定導' in a_indexed.columns else ''

    # 3. 調整欄位順序：在精舍後加入 館、哨、任、特
    base_cols = list(df_b_cleaned.columns)
    if '精舍' in base_cols:
        idx = base_cols.index('精舍') + 1
        for c in ['館', '哨', '任', '特展']:
            if c in base_cols:
                base_cols.remove(c)
        new_cols = base_cols[:idx] + ['館', '哨', '任', '特展'] + base_cols[idx:]
        df_b_cleaned = df_b_cleaned[new_cols]

    # 4. 隔離發心日期/時間異常者（置底）
    if '發心起始時間' in df_b_cleaned.columns:
        df_normal = df_b_cleaned[df_b_cleaned['發心起始時間'].notna()]
        df_abnormal = df_b_cleaned[df_b_cleaned['發心起始時間'].isna()]
        df_b_cleaned = pd.concat([df_normal, df_abnormal], ignore_index=True)

    # 5. 產生多個 Sheet 工作表
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_b_cleaned.to_excel(writer, sheet_name=target_date, index=False)
        
        cols_guardian = [b_id_col, '姓名', '法名', '性別', '館', '任']
        guardian_cols_exist = [c for c in cols_guardian if c in df_b_cleaned.columns]
        df_guardian = df_b_cleaned[guardian_cols_exist].copy()
        if '館' in df_guardian.columns:
            df_guardian.rename(columns={'館': '地點'}, inplace=True)
        df_guardian.to_excel(writer, sheet_name="貼後上傳守護網用", index=False)

        cols_meal = [b_id_col, '精舍', '姓名', '法名', '館', '哨']
        meal_cols_exist = [c for c in cols_meal if c in df_b_cleaned.columns]
        df_meal = df_b_cleaned[meal_cols_exist].copy()
        df_meal['午齋'] = 1
        df_meal.to_excel(writer, sheet_name="用齋用", index=False)

        df_b_cleaned.to_excel(writer, sheet_name="館員用", index=False)

        df_sub = df_b_cleaned[df_b_cleaned['館'] == '分館'] if '館' in df_b_cleaned.columns else df_b_cleaned
        df_sub.to_excel(writer, sheet_name="分館上哨白板組長用", index=False)

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
            df_b = pd.read_excel(file_b)

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
