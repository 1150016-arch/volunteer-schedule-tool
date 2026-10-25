import io
import re
import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="義工排班資料自動比對系統", page_icon="📅", layout="wide"
)

st.title("📅 義工排班資料自動比對系統")
st.markdown(
    "匯入雲端預排檔（表格 A）與登記綜合名單（表格 B），自動進行 VLOOKUP 比對、剔除非核心欄位，並**自動標註共乘色彩與時間異常提醒**。"
)

# 32 色精選柔和粉彩配色池 (專為百人大型排班設計，高對比度不刺眼)
PASTEL_COLORS = [
    "FFF2CC",
    "D9EAD3",
    "D0E0E3",
    "F4CCCC",
    "EFE5D9",
    "E1D5E7",
    "FFE599",
    "B6D7A8",
    "A2C4C9",
    "F1C232",
    "FCE5CD",
    "D0E0E3",
    "EAD1DC",
    "D9D2E9",
    "C9DAF8",
    "FFF2CC",
    "F9CB9C",
    "EA9999",
    "FFE599",
    "B6D7A8",
    "A2C4C9",
    "9FC5E8",
    "B4A7D6",
    "D5A6BD",
    "E06666",
    "F6B26B",
    "FFD966",
    "93C47D",
    "76A5AF",
    "6FA8DC",
    "8E7CC3",
    "C27BA0",
]

# 時間異常警告底色 (淺橘色)
TIME_ABNORMAL_FILL = PatternFill(
    start_color="FFE6CC", end_color="FFE6CC", fill_type="solid"
)


def load_clean_b(file):
    """讀取新版表格 B：自動辨識標題列位置（相容第一列與第二列標題）"""
    df = pd.read_excel(file)
    if "義工編號" not in df.columns and "編號" not in df.columns:
        if len(df) > 0 and any(
            df.iloc[0].astype(str).str.contains("義工編號|編號|姓名")
        ):
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
    return df


def setup_sheet_style(ws):
    """自動欄寬與列印縮放設定"""
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                lines = str(cell.value).split("\n")
                line_len = max(len(l) for l in lines)
                if line_len > max_len:
                    max_len = line_len
        ws.column_dimensions[col_letter].width = min(max(max_len * 2.2, 10), 38)


# ==========================================
# 核心功能：VLOOKUP 資料比對與自動帶入 + 著色
# ==========================================
target_date = st.text_input(
    "排班日期 / Sheet 名稱 (如: 1022 或 1028)", value="1028", key="t_date"
)

col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader(
        "上傳表格 A (雲端預排檔)", type=["xlsx", "xls"], key="f_a"
    )
with col2:
    file_b = st.file_uploader(
        "上傳表格 B (登記綜合名單)", type=["xlsx", "xls"], key="f_b"
    )

if file_a and file_b:
    try:
        xls_a = pd.ExcelFile(file_a)
        if target_date not in xls_a.sheet_names:
            st.error(f"⚠️ 在表格 A 中找不到名稱為 **{target_date}** 的工作表！")
        else:
            df_a = pd.read_excel(xls_a, sheet_name=target_date)
            df_b = load_clean_b(file_b)

            a_id_col = (
                "編號"
                if "編號" in df_a.columns
                else ("義工編號" if "義工編號" in df_a.columns else df_a.columns[1])
            )
            b_id_col = (
                "義工編號"
                if "義工編號" in df_b.columns
                else ("編號" if "編號" in df_b.columns else df_b.columns[0])
            )

            df_a_clean = df_a.dropna(subset=[a_id_col]).copy()
            df_a_unique = df_a_clean.drop_duplicates(subset=[a_id_col]).set_index(
                a_id_col
            )

            # 剔除不需要的非排班欄位
            drop_cols = [
                "大隊別", "小隊別", "防疫證書證號", "任務", "地點", "檢定資格", "活動名稱", 
                "護持起始時間", "護持迄止時間", "護持開始時間", "護持結束時間", 
                "異動備註", "動態備註(表格B)", "備註2", "備註3", "備註4", "是否納保", 
                "保險註記", "精舍確認"
            ]
            df_b_cleaned = df_b.drop(columns=[c for c in drop_cols if c in df_b.columns]).copy()

            # 【徹底刪除所有重複的手機欄位】只保留第一個出現的「手機」，後續不論是 手機.1、手機.2 或是獨立出現的第二個 手機 欄位通通剔除
            phone_cols = [c for c in df_b_cleaned.columns if str(c).startswith("手機")]
            if len(phone_cols) > 1:
                df_b_cleaned = df_b_cleaned.drop(columns=phone_cols[1:])

            # VLOOKUP 帶入表格 A 核心欄位
            df_b_cleaned["館"] = (
                df_b_cleaned[b_id_col].map(df_a_unique["館別"])
                if "館別" in df_a_unique.columns
                else ""
            )
            df_b_cleaned["哨"] = (
                df_b_cleaned[b_id_col].map(df_a_unique["哨別"])
                if "哨別" in df_a_unique.columns
                else ""
            )
            df_b_cleaned["任"] = (
                df_b_cleaned[b_id_col].map(df_a_unique["任務"])
                if "任務" in df_a_unique.columns
                else ""
            )
            df_b_cleaned["特展"] = (
                df_b_cleaned[b_id_col].map(df_a_unique["特展定導"])
                if "特展定導" in df_a_unique.columns
                else ""
            )

            zongfan_col_a = next(
                (c for c in df_a_unique.columns if "仰懷宗範" in str(c)), None
            )
            if zongfan_col_a:
                df_b_cleaned["仰懷宗範"] = df_b_cleaned[b_id_col].map(
                    df_a_unique[zongfan_col_a]
                )

            car_col_a = next(
                (
                    c
                    for c in ["同館哨/親眷或共乘", "車牌號碼", "車牌", "車號"]
                    if c in df_a_unique.columns
                ),
                None,
            )
            if car_col_a:
                df_b_cleaned["同館哨/親眷或共乘"] = df_b_cleaned[b_id_col].map(
                    df_a_unique[car_col_a]
                )

            remark_col_a = next(
                (
                    c
                    for c in ["特殊需求", "備註", "備註1", "常態備註", "個人備註"]
                    if c in df_a_unique.columns
                ),
                None,
            )
            if remark_col_a:
                df_b_cleaned["常態備註(表格A)"] = df_b_cleaned[b_id_col].map(
                    df_a_unique[remark_col_a]
                )

            # 調整欄位順序（放在精舍後方）
            base_cols = list(df_b_cleaned.columns)
            if "精舍" in base_cols:
                idx = base_cols.index("精舍") + 1
                add_cols = [
                    c for c in ["館", "哨", "任", "特展", "仰懷宗範"] if c in base_cols
                ]
                for c in add_cols:
                    base_cols.remove(c)
                new_cols = base_cols[:idx] + add_cols + base_cols[idx:]
                df_b_cleaned = df_b_cleaned[new_cols]

            # 無發心時間者自動置底
            start_time_col_name = next(
                (c for c in df_b_cleaned.columns if "發心" in str(c) and "起" in str(c)),
                None,
            )
            if start_time_col_name:
                df_normal = df_b_cleaned[df_b_cleaned[start_time_col_name].notna()].copy()
                df_abnormal = df_b_cleaned[
                    df_b_cleaned[start_time_col_name].isna()
                ].copy()
                df_b_cleaned = pd.concat([df_normal, df_abnormal], ignore_index=True)

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df_b_cleaned.to_excel(writer, sheet_name=target_date, index=False)
                ws = writer.sheets[target_date]

                # ----------------------------------------
                # 自動著色處理：共乘對應與時間提醒
                # ----------------------------------------
                car_col_b = None
                name_col_b = None
                start_time_col_b = None
                end_time_col_b = None

                for col_idx, col_name in enumerate(df_b_cleaned.columns, 1):
                    c_str = str(col_name)
                    if c_str in ["同館哨/親眷或共乘", "車牌號碼", "車牌", "車號"]:
                        car_col_b = col_idx
                    if c_str in ["姓名", "義工姓名"]:
                        name_col_b = col_idx
                    if "發心" in c_str and ("起" in c_str or "始" in c_str):
                        start_time_col_b = col_idx
                    if "發心" in c_str and ("迄" in c_str or "止" in c_str or "終" in c_str):
                        end_time_col_b = col_idx

                # 1. 共乘色彩著色 (支援百人 32 種粉彩輪替)
                if car_col_b:
                    car_series = (
                        df_b_cleaned.iloc[:, car_col_b - 1]
                        .dropna()
                        .astype(str)
                        .str.strip()
                    )
                    car_series = car_series[
                        ~car_series.isin(["0", "0.0", "nan", "None", ""])
                    ]
                    car_counts = car_series.value_counts()
                    valid_cars = car_counts[car_counts > 1].index.tolist()

                    car_color_map = {}
                    for idx_c, car in enumerate(valid_cars):
                        car_color_map[car] = PASTEL_COLORS[idx_c % len(PASTEL_COLORS)]

                    for row_idx, row in df_b_cleaned.iterrows():
                        c_val = (
                            str(row.iloc[car_col_b - 1]).strip()
                            if pd.notna(row.iloc[car_col_b - 1])
                            else ""
                        )
                        if c_val in car_color_map:
                            color_hex = car_color_map[c_val]
                            fill = PatternFill(
                                start_color=color_hex, end_color=color_hex, fill_type="solid"
                            )
                            excel_row = row_idx + 2
                            if name_col_b:
                                ws.cell(row=excel_row, column=name_col_b).fill = fill
                            ws.cell(row=excel_row, column=car_col_b).fill = fill

                # 2. 發心時間異常提醒 (橘底)
                if start_time_col_b or end_time_col_b:
                    for row_idx, row in df_b_cleaned.iterrows():
                        excel_row = row_idx + 2
                        st_val = (
                            str(row.iloc[start_time_col_b - 1]).strip()
                            if (
                                start_time_col_b
                                and pd.notna(row.iloc[start_time_col_b - 1])
                            )
                            else ""
                        )
                        et_val = (
                            str(row.iloc[end_time_col_b - 1]).strip()
                            if (end_time_col_b and pd.notna(row.iloc[end_time_col_b - 1]))
                            else ""
                        )

                        is_start_abnormal = (
                            st_val
                            and st_val != "nan"
                            and "08:00" not in st_val
                            and "8:00" not in st_val
                        )
                        is_end_abnormal = (
                            et_val
                            and et_val != "nan"
                            and "17:00" not in et_val
                            and "5:00" not in et_val
                        )

                        if is_start_abnormal or is_end_abnormal:
                            if start_time_col_b:
                                ws.cell(
                                    row=excel_row, column=start_time_col_b
                                ).fill = TIME_ABNORMAL_FILL
                            if end_time_col_b:
                                ws.cell(
                                    row=excel_row, column=end_time_col_b
                                ).fill = TIME_ABNORMAL_FILL

                setup_sheet_style(ws)

            out.seek(0)
            st.success("✅ 自動比對與色彩標註完成！")
            st.download_button(
                label="📥 下載待確認班表 (.xlsx)",
                data=out,
                file_name=f"工作班表_{target_date}_待確認.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                type="primary",
            )
    except Exception as e:
        st.error(f"❌ 處理失敗：{str(e)}")
