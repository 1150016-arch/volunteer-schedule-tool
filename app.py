import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="義工排班自動化系統",
    page_icon="📅",
    layout="wide"
)

st.title("📅 義工排班與列印檔自動生成系統")
st.markdown("本系統採 **兩階段處理**：先生成待確認班表，確認後一鍵生成所有列印頁籤（含自動彙整票口導覽資訊）。")

tab1, tab2 = st.tabs(["1️⃣ 階段一：初步 VLOOKUP 比對與車牌著色", "2️⃣ 階段二：生成最終表格 C 與列印檔"])

PASTEL_COLORS = [
    "FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC", "EFE5D9",
    "E1D5E7", "FFE599", "B6D7A8", "A2C4C9", "F1C232"
]
TIME_ABNORMAL_FILL = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")

def load_clean_b(file):
    """讀取表格 B 並自動處理標題在第二列的問題"""
    df = pd.read_excel(file)
    if '義工編號' not in df.columns and '編號' not in df.columns:
        if len(df) > 0 and any(df.iloc[0].astype(str).str.contains('義工編號|編號|姓名')):
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
    return df

def setup_sheet_print(ws, orientation='landscape'):
    """設定 Excel 列印屬性：自動縮放為 1 頁寬"""
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if orientation == 'landscape' else ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

def process_bulletin_phone(df_in):
    """大班表電話處理：僅保留任務含『組長』者的電話，一般組員清空"""
    df_out = df_in.copy()
    phone_cols = [c for c in df_out.columns if '電話' in str(c) or '手機' in str(c) or '聯絡' in str(c)]
    task_col = '任' if '任' in df_out.columns else ('任務' if '任務' in df_out.columns else None)

    if phone_cols and task_col:
        for p_col in phone_cols:
            new_phones = []
            for _, row in df_out.iterrows():
                task_val = str(row.get(task_col, ''))
                phone_val = str(row.get(p_col, ''))
                if '組長' in task_val and pd.notna(row.get(p_col)):
                    new_phones.append(phone_val)
                else:
                    new_phones.append('')
            df_out[p_col] = new_phones
    return df_out

# ==========================================
# 階段一：初步 VLOOKUP 比對與車牌著色
# ==========================================
with tab1:
    st.header("1️⃣ 匯入原始檔案，生成帶有【車牌色彩與時間提醒】的待確認班表")
    target_date_1 = st.text_input("排班日期 / Sheet 名稱", value="1021", key="t1_date")
    
    col1, col2 = st.columns(2)
    with col1:
        file_a = st.file_uploader("上傳表格 A (雲端預排檔)", type=["xlsx", "xls"], key="f_a")
    with col2:
        file_b = st.file_uploader("上傳表格 B (登記發心名單)", type=["xlsx", "xls"], key="f_b")

    if file_a and file_b:
        try:
            xls_a = pd.ExcelFile(file_a)
            if target_date_1 not in xls_a.sheet_names:
                st.error(f"⚠️ 在表格 A 中找不到名稱為 **{target_date_1}** 的工作表！")
            else:
                df_a = pd.read_excel(xls_a, sheet_name=target_date_1)
                df_b = load_clean_b(file_b)

                a_id_col = '編號' if '編號' in df_a.columns else ('義工編號' if '義工編號' in df_a.columns else df_a.columns[1])
                b_id_col = '義工編號' if '義工編號' in df_b.columns else ('編號' if '編號' in df_b.columns else df_b.columns[0])

                df_a_clean = df_a.dropna(subset=[a_id_col]).copy()
                df_a_unique = df_a_clean.drop_duplicates(subset=[a_id_col]).set_index(a_id_col)

                drop_cols = ['大隊別', '小隊別', '防疫證書證號', '任務', '地點', '檢定資格', '活動名稱', '護持開始時間', '護持結束時間']
                df_b_cleaned = df_b.drop(columns=[c for c in drop_cols if c in df_b.columns]).copy()

                df_b_cleaned['館'] = df_b_cleaned[b_id_col].map(df_a_unique['館別']) if '館別' in df_a_unique.columns else ''
                df_b_cleaned['哨'] = df_b_cleaned[b_id_col].map(df_a_unique['哨別']) if '哨別' in df_a_unique.columns else ''
                df_b_cleaned['任'] = df_b_cleaned[b_id_col].map(df_a_unique['任務']) if '任務' in df_a_unique.columns else ''
                df_b_cleaned['特展'] = df_b_cleaned[b_id_col].map(df_a_unique['特展定導']) if '特展定導' in df_a_unique.columns else ''

                car_col_a = None
                for c in ['車牌號碼', '車牌', '車號']:
                    if c in df_a_unique.columns:
                        car_col_a = c
                        break
                if car_col_a:
                    df_b_cleaned['車牌號碼'] = df_b_cleaned[b_id_col].map(df_a_unique[car_col_a])

                remark_col_a = None
                for c in ['備註', '備註1', '常態備註', '個人備註']:
                    if c in df_a_unique.columns:
                        remark_col_a = c
                        break
                if remark_col_a:
                    df_b_cleaned['常態備註(表格A)'] = df_b_cleaned[b_id_col].map(df_a_unique[remark_col_a])

                rename_dict = {}
                for c in df_b_cleaned.columns:
                    if str(c) in ['備註', '備註1'] and c != '常態備註(表格A)':
                        rename_dict[c] = '動態備註(表格B)'
                if rename_dict:
                    df_b_cleaned.rename(columns=rename_dict, inplace=True)

                base_cols = list(df_b_cleaned.columns)
                if '精舍' in base_cols:
                    idx = base_cols.index('精舍') + 1
                    for c in ['館', '哨', '任', '特展']:
                        if c in base_cols:
                            base_cols.remove(c)
                    new_cols = base_cols[:idx] + ['館', '哨', '任', '特展'] + base_cols[idx:]
                    df_b_cleaned = df_b_cleaned[new_cols]

                no_time_count = 0
                if '發心起始時間' in df_b_cleaned.columns:
                    df_normal = df_b_cleaned[df_b_cleaned['發心起始時間'].notna()].copy()
                    df_abnormal = df_b_cleaned[df_b_cleaned['發心起始時間'].isna()].copy()
                    no_time_count = len(df_abnormal)
                    df_b_cleaned = pd.concat([df_normal, df_abnormal], ignore_index=True)

                if no_time_count > 0:
                    st.warning(f"⚠️ 偵測到有 **{no_time_count}** 位義工無發心時間（僅報名上課/接駁），已自動移至表格最下方！如為本館上課請留意訂午齋需求。")

                out1 = io.BytesIO()
                with pd.ExcelWriter(out1, engine='openpyxl') as writer:
                    df_b_cleaned.to_excel(writer, sheet_name=target_date_1, index=False)
                    ws = writer.sheets[target_date_1]

                    car_col_b = None
                    name_col_b = None
                    start_time_col_b = None
                    end_time_col_b = None

                    for col_idx, col_name in enumerate(df_b_cleaned.columns, 1):
                        c_str = str(col_name)
                        if c_str in ['車牌號碼', '車牌', '車號']:
                            car_col_b = col_idx
                        if c_str in ['姓名', '義工姓名']:
                            name_col_b = col_idx
                        if c_str in ['發心起始時間', '發心開始時間']:
                            start_time_col_b = col_idx
                        if c_str in ['發心迄止時間', '發心結束時間', '發心截止時間']:
                            end_time_col_b = col_idx

                    if car_col_b:
                        car_values = df_b_cleaned.iloc[:, car_col_b - 1].dropna().astype(str).str.strip()
                        car_counts = car_values.value_counts()
                        valid_cars = car_counts[car_counts > 1].index.tolist()

                        car_color_map = {}
                        for idx, car in enumerate(valid_cars):
                            car_color_map[car] = PASTEL_COLORS[idx % len(PASTEL_COLORS)]

                        for row_idx, row in df_b_cleaned.iterrows():
                            c_val = str(row.iloc[car_col_b - 1]).strip() if pd.notna(row.iloc[car_col_b - 1]) else ''
                            if c_val in car_color_map:
                                color_hex = car_color_map[c_val]
                                fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                                excel_row = row_idx + 2
                                if name_col_b:
                                    ws.cell(row=excel_row, column=name_col_b).fill = fill
                                ws.cell(row=excel_row, column=car_col_b).fill = fill

                    if start_time_col_b or end_time_col_b:
                        for row_idx, row in df_b_cleaned.iterrows():
                            excel_row = row_idx + 2
                            st_val = str(row.iloc[start_time_col_b - 1]).strip() if (start_time_col_b and pd.notna(row.iloc[start_time_col_b - 1])) else ''
                            et_val = str(row.iloc[end_time_col_b - 1]).strip() if (end_time_col_b and pd.notna(row.iloc[end_time_col_b - 1])) else ''
                            
                            is_start_abnormal = (st_val and st_val != 'nan' and '08:00' not in st_val and '8:00' not in st_val)
                            is_end_abnormal = (et_val and et_val != 'nan' and '17:00' not in et_val and '5:00' not in et_val)

                            if is_start_abnormal or is_end_abnormal:
                                if start_time_col_b:
                                    ws.cell(row=excel_row, column=start_time_col_b).fill = TIME_ABNORMAL_FILL
                                if end_time_col_b:
                                    ws.cell(row=excel_row, column=end_time_col_b).fill = TIME_ABNORMAL_FILL

                out1.seek(0)

                st.success("✅ 初步比對完成！同車牌姓名已上色、發心時間異常（遲到/早退）者已塗上淺橘色提醒。")
                st.download_button(
                    label="📥 下載待確認班表 (.xlsx)",
                    data=out1,
                    file_name=f"工作班表_{target_date_1}_待確認.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ 處理失敗：{str(e)}")

# ==========================================
# 階段二：生成最終表格 C 與所有列印頁籤 (自動抽取票口導覽)
# ==========================================
with tab2:
    st.header("2️⃣ 上傳最終班表，生成表格 C 與所有列印 Sheet")
    file_final = st.file_uploader("上傳您已排好/確認好的工作班表 (.xlsx)", type=["xlsx", "xls"], key="f_final")
    target_date_2 = st.text_input("排班日期 / Sheet 名稱", value="1021", key="t2_date")

    if file_final:
        try:
            xls_f = pd.ExcelFile(file_final)
            sheet_use = target_date_2 if target_date_2 in xls_f.sheet_names else xls_f.sheet_names[0]
            df_final = pd.read_excel(xls_f, sheet_name=sheet_use)

            st.subheader("⚙️ 午齋單參數設定")
            cat_col = None
            for c in ['錄取類別', '類別', '單位', '動態備註(表格B)', '備註']:
                if c in df_final.columns:
                    cat_col = c
                    break

            if cat_col:
                all_cats = [str(x) for x in df_final[cat_col].dropna().unique() if str(x).strip() != '']
                default_selected = [c for c in all_cats if '分館' not in c]
                selected_meal_cats = st.multiselect("請選擇【需要列入午齋單】的類別/單位：", options=all_cats, default=default_selected)
            else:
                selected_meal_cats = None
                st.info("ℹ️ 未偵測到明確類別欄位，將預設依據『館』欄位不為分館者列入午齋單。")

            if st.button("🚀 一鍵生成完整表格 C (含自動抽取的票口導覽資訊)", type="primary"):
                b_id_col = '義工編號' if '義工編號' in df_final.columns else ('編號' if '編號' in df_final.columns else df_final.columns[0])

                out2 = io.BytesIO()
                with pd.ExcelWriter(out2, engine='openpyxl') as writer:
                    # 1. 總班表
                    df_final.to_excel(writer, sheet_name=target_date_2, index=False)
                    setup_sheet_print(writer.sheets[target_date_2])

                    # 2. 館員用 (連動公式)
                    df_staff = pd.DataFrame(columns=df_final.columns)
                    for row_idx in range(len(df_final)):
                        row_formulas = []
                        excel_row = row_idx + 2
                        for col_idx in range(len(df_final.columns)):
                            col_letter = get_column_letter(col_idx + 1)
                            row_formulas.append(f"='{target_date_2}'!{col_letter}{excel_row}")
                        df_staff.loc[row_idx] = row_formulas
                    df_staff.to_excel(writer, sheet_name="館員用", index=False)
                    setup_sheet_print(writer.sheets["館員用"])

                    # 3. 兩館接駁乘車名單
                    shuttle_cols = [c for c in df_final.columns if '主體到大博' in str(c) or '大博到主體' in str(c)]
                    if shuttle_cols:
                        mask = False
                        for sc in shuttle_cols:
                            mask = mask | (df_final[sc].astype(str).str.strip() == '1')
                        df_shuttle = df_final[mask]
                    else:
                        df_shuttle = pd.DataFrame(columns=df_final.columns)
                    df_shuttle.to_excel(writer, sheet_name="兩館接駁乘車名單", index=False)
                    setup_sheet_print(writer.sheets["兩館接駁乘車名單"])

                    # 4. 用齋用
                    if cat_col and selected_meal_cats is not None:
                        df_meal_src = df_final[df_final[cat_col].astype(str).isin(selected_meal_cats)].copy()
                    else:
                        df_meal_src = (df_final[df_final['館'] != '分館'] if '館' in df_final.columns else df_final).copy()
                    
                    df_meal_out = pd.DataFrame()
                    df_meal_out['NO.'] = range(1, len(df_meal_src) + 1)
                    df_meal_out['編號'] = df_meal_src[b_id_col].values if b_id_col in df_meal_src.columns else ''
                    df_meal_out['精舍'] = df_meal_src['精舍'].values if '精舍' in df_meal_src.columns else ''
                    
                    name_list = []
                    for _, r in df_meal_src.iterrows():
                        nm = str(r.get('姓名', '')) if pd.notna(r.get('姓名', '')) else ''
                        fm = str(r.get('法名', '')) if pd.notna(r.get('法名', '')) else ''
                        if fm and fm != 'nan':
                            name_list.append(f"{nm}({fm})")
                        else:
                            name_list.append(nm)
                    df_meal_out['姓名(法名)'] = name_list
                    df_meal_out['午齋'] = 1
                    
                    if cat_col and cat_col in df_meal_src.columns:
                        df_meal_out['備註'] = df_meal_src[cat_col].values
                    else:
                        df_meal_out['備註'] = ''

                    df_meal_out.to_excel(writer, sheet_name="用齋用", index=False)
                    setup_sheet_print(writer.sheets["用齋用"], orientation='portrait')

                    # 5. 團導 SHEET
                    df_tuandao = pd.DataFrame([
                        ["參訪日期", "星期", "時間", "主題", "團體名稱", "人數", "預排導覽員", "備註"],
                        ["", "", "", "", "", "", "", ""],
                        ["分館_團體導覽表", "", "", "", "", "", "", ""],
                        ["參訪日期", "星期", "時間", "主題", "團體名稱", "人數", "預排導覽員", "備註"]
                    ], columns=["本館_團體導覽表", "Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 4", "Unnamed: 5", "Unnamed: 6", "Unnamed: 7"])
                    df_tuandao.to_excel(writer, sheet_name="團導", index=False)
                    setup_sheet_print(writer.sheets["團導"])

                    # 6. 自動抽取並生成「票口導覽資訊」SHEET
                    # 搜尋導覽相關欄位 (第一哨、第二哨、第三哨)
                    s1_col = next((c for c in df_final.columns if '第一哨' in str(c)), None)
                    s2_col = next((c for c in df_final.columns if '第二哨' in str(c)), None)
                    s3_col = next((c for c in df_final.columns if '第三哨' in str(c)), None)
                    task_col = next((c for c in df_final.columns if '任' in str(c) or '任務' in str(c)), None)
                    hall_col = next((c for c in df_final.columns if '館' in str(c)), None)
                    shao_col = next((c for c in df_final.columns if '哨' in str(c)), None)

                    piaokou_rows = [
                        ["參訪日期", "星期", "時間", "主題", "團體名稱", "人數", "預排導覽員"],
                        ["", "", "", "", "", "", ""],
                        ["", f"{target_date_2[:2]}/{target_date_2[2:]}", "本館", "定導", "哨點", "第一哨\nA:09:00-10:30\nB:10:30-12:10", "第二哨\nA:12:10-13:10\nB:13:10-14:00", "第三哨\nA:14:00-15:15\nB:15:15-16:30"]
                    ]

                    # 抓取導覽員資料填入票口導覽表
                    if task_col and hall_col and shao_col:
                        # 篩選導覽任務（展廳/特展定導）
                        guide_df = df_final[df_final[task_col].astype(str).str.contains('導覽|定導')]
                        for _, grow in guide_df.iterrows():
                            g_hall = str(grow.get(hall_col, ''))
                            g_shao = str(grow.get(shao_col, ''))
                            g_s1 = str(grow.get(s1_col, '')) if s1_col else ''
                            g_s2 = str(grow.get(s2_col, '')) if s2_col else ''
                            g_s3 = str(grow.get(s3_col, '')) if s3_col else ''
                            piaokou_rows.append(["", "", g_hall, "定導", g_shao, g_s1, g_s2, g_s3])
                    else:
                        piaokou_rows.append(["", "", "A", "定導", "", "", "", ""])
                        piaokou_rows.append(["", "", "B", "定導", "", "", "", ""])

                    df_piaokou = pd.DataFrame(piaokou_rows)
                    df_piaokou.to_excel(writer, sheet_name="票口導覽資訊", index=False, header=False)
                    setup_sheet_print(writer.sheets["票口導覽資訊"])

                    # 7. 佈告欄大班表
                    df_bulletin = process_bulletin_phone(df_final)
                    for t_col in ['發心起始時間', '發心開始時間', '發心結束時間', '發心截止時間']:
                        if t_col in df_bulletin.columns:
                            df_bulletin[t_col] = ''
                    df_bulletin.to_excel(writer, sheet_name="佈告欄大班表", index=False)
                    setup_sheet_print(writer.sheets["佈告欄大班表"])

                    # 8. 分館佈告欄班表
                    if '館' in df_final.columns:
                        df_sub = df_final[df_final['館'] == '分館'].copy()
                    else:
                        df_sub = df_final.copy()
                    df_sub = process_bulletin_phone(df_sub)
                    for t_col in ['發心起始時間', '發心開始時間', '發心結束時間', '發心截止時間']:
                        if t_col in df_sub.columns:
                            df_sub[t_col] = ''
                    df_sub.to_excel(writer, sheet_name="分館佈告欄班表", index=False)
                    setup_sheet_print(writer.sheets["分館佈告欄班表"])

                    # 9. 組長班表 (動態判斷)
                    if task_col:
                        df_lead_main_a = df_final[(df_final[task_col].astype(str).str.contains('組長')) & (df_final['館'] != '分館') & (df_final['哨'] == 'A')]
                        if len(df_lead_main_a) > 0:
                            df_lead_main_a.to_excel(writer, sheet_name="本館A哨組長表", index=False)
                            setup_sheet_print(writer.sheets["本館A哨組長表"])

                        df_lead_main_b = df_final[(df_final[task_col].astype(str).str.contains('組長')) & (df_final['館'] != '分館') & (df_final['哨'] == 'B')]
                        if len(df_lead_main_b) > 0:
                            df_lead_main_b.to_excel(writer, sheet_name="本館B哨組長表", index=False)
                            setup_sheet_print(writer.sheets["本館B哨組長表"])

                        df_lead_sub = df_final[(df_final[task_col].astype(str).str.contains('組長')) & (df_final['館'] == '分館')]
                        if len(df_lead_sub) > 0:
                            df_lead_sub.to_excel(writer, sheet_name="分館組長表", index=False)
                            setup_sheet_print(writer.sheets["分館組長表"])

                out2.seek(0)
                st.balloons()
                st.success("🎉 所有 Sheet 生成完畢（含自動抽取的票口導覽資訊）！")
                st.download_button(
                    label="📥 下載完整表格 C (.xlsx)",
                    data=out2,
                    file_name=f"表格C_完整排班與列印檔_{target_date_2}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ 處理失敗：{str(e)}")
