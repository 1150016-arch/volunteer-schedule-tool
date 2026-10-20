from datetime import datetime
import io
import re
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="義工排班自動化系統", page_icon="📅", layout="wide"
)

st.title("📅 義工排班與列印檔自動生成系統")
st.markdown(
    "本系統採 **兩階段處理**：先生成待確認班表，確認後一鍵生成所有列印頁籤。"
)

tab1, tab2 = st.tabs([
    "1️⃣ 階段一：初步 VLOOKUP 比對與共乘/需求著色",
    "2️⃣ 階段二：生成最終表格 C 與列印檔",
])

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
]
TIME_ABNORMAL_FILL = PatternFill(
    start_color="FFE6CC", end_color="FFE6CC", fill_type="solid"
)


def load_clean_b(file):
  """讀取表格 B 並自動處理標題在第二列的問題"""
  df = pd.read_excel(file)
  if "義工編號" not in df.columns and "編號" not in df.columns:
    if len(df) > 0 and any(
        df.iloc[0].astype(str).str.contains("義工編號|編號|姓名")
    ):
      df.columns = df.iloc[0]
      df = df[1:].reset_index(drop=True)
  return df


def setup_sheet_print(ws, orientation="landscape"):
  """設定 Excel 列印屬性：自動縮放為 1 頁寬，並自動適應欄寬"""
  ws.page_setup.orientation = (
      ws.ORIENTATION_LANDSCAPE
      if orientation == "landscape"
      else ws.ORIENTATION_PORTRAIT
  )
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


def process_bulletin_phone(df_in):
  """大班表電話處理：僅保留任務含『組長』者的電話，一般組員清空"""
  df_out = df_in.copy()
  phone_cols = [
      c
      for c in df_out.columns
      if "電話" in str(c) or "手機" in str(c) or "聯絡" in str(c)
  ]
  task_col = (
      "任"
      if "任" in df_out.columns
      else ("任務" if "任務" in df_out.columns else None)
  )

  if phone_cols and task_col:
    for p_col in phone_cols:
      new_phones = []
      for _, row in df_out.iterrows():
        task_val = str(row.get(task_col, ""))
        phone_val = str(row.get(p_col, ""))
        if (
            "組長" in task_val
            and pd.notna(row.get(p_col))
            and phone_val != "nan"
        ):
          new_phones.append(phone_val)
        else:
          new_phones.append("")
      df_out[p_col] = new_phones
  return df_out


def parse_m_d(date_str):
  """剖析如 '10/18', '1018', '10-18' 等日期格式，回傳 (月, 日)"""
  if pd.isna(date_str):
    return None
  s = str(date_str).strip()
  m = re.search(r"(\d{1,2})[/\-_](\d{1,2})", s)
  if m:
    return (int(m.group(1)), int(m.group(2)))
  if len(s) == 4 and s.isdigit():
    return (int(s[:2]), int(s[2:]))
  return None


# ==========================================
# 階段一：初步 VLOOKUP 比對與欄位清理
# ==========================================
with tab1:
  st.header("1️⃣ 匯入原始檔案，生成帶有【共乘色彩與時間提醒】的待確認班表")
  target_date_1 = st.text_input(
      "排班日期 / Sheet 名稱 (如: 1022 或 10/22)", value="1022", key="t1_date"
  )

  col1, col2 = st.columns(2)
  with col1:
    file_a = st.file_uploader(
        "上傳表格 A (雲端預排檔)", type=["xlsx", "xls"], key="f_a"
    )
  with col2:
    file_b = st.file_uploader(
        "上傳表格 B (登記發心名單)", type=["xlsx", "xls"], key="f_b"
    )

  if file_a and file_b:
    try:
      xls_a = pd.ExcelFile(file_a)
      if target_date_1 not in xls_a.sheet_names:
        st.error(
            f"⚠️ 在表格 A 中找不到名稱為 **{target_date_1}** 的工作表！"
        )
      else:
        df_a = pd.read_excel(xls_a, sheet_name=target_date_1)
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

        # 清除不需要的欄位
        drop_cols = [
            "大隊別",
            "小隊別",
            "防疫證書證號",
            "任務",
            "地點",
            "檢定資格",
            "活動名稱",
            "護持開始時間",
            "護持結束時間",
            "異動備註",
            "動態備註(表格B)",
            "備註2",
            "備註3",
            "備註4",
            "是否納保",
            "保險註記",
            "精舍確認",
        ]
        df_b_cleaned = df_b.drop(
            columns=[c for c in drop_cols if c in df_b.columns]
        ).copy()

        # VLOOKUP 表格 A (包含新增的「仰懷宗範」)
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

        # 自動 VLOOKUP「仰懷宗範」
        zongfan_col_a = next(
            (c for c in df_a_unique.columns if "仰懷宗範" in str(c)), None
        )
        if zongfan_col_a:
          df_b_cleaned["仰懷宗範"] = df_b_cleaned[b_id_col].map(
              df_a_unique[zongfan_col_a]
          )

        car_col_a = None
        for c in ["同館哨/親眷或共乘", "車牌號碼", "車牌", "車號"]:
          if c in df_a_unique.columns:
            car_col_a = c
            break
        if car_col_a:
          df_b_cleaned["同館哨/親眷或共乘"] = df_b_cleaned[b_id_col].map(
              df_a_unique[car_col_a]
          )

        remark_col_a = None
        for c in ["特殊需求", "備註", "備註1", "常態備註", "個人備註"]:
          if c in df_a_unique.columns:
            remark_col_a = c
            break
        if remark_col_a:
          df_b_cleaned["常態備註(表格A)"] = df_b_cleaned[b_id_col].map(
              df_a_unique[remark_col_a]
          )

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

        no_time_count = 0
        if "發心起始時間" in df_b_cleaned.columns:
          df_normal = df_b_cleaned[
              df_b_cleaned["發心起始時間"].notna()
          ].copy()
          df_abnormal = df_b_cleaned[
              df_b_cleaned["發心起始時間"].isna()
          ].copy()
          no_time_count = len(df_abnormal)
          df_b_cleaned = pd.concat(
              [df_normal, df_abnormal], ignore_index=True
          )

        if no_time_count > 0:
          st.warning(
              f"⚠️ 偵測到有 **{no_time_count}**"
              " 位義工無發心時間，已自動移至表格最下方！"
          )

        out1 = io.BytesIO()
        with pd.ExcelWriter(out1, engine="openpyxl") as writer:
          df_b_cleaned.to_excel(
              writer, sheet_name=target_date_1, index=False
          )
          ws = writer.sheets[target_date_1]

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
            if c_str in ["發心起始時間", "發心開始時間"]:
              start_time_col_b = col_idx
            if c_str in ["發心迄止時間", "發心結束時間", "發心截止時間"]:
              end_time_col_b = col_idx

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
            for idx, car in enumerate(valid_cars):
              car_color_map[car] = PASTEL_COLORS[idx % len(PASTEL_COLORS)]

            for row_idx, row in df_b_cleaned.iterrows():
              c_val = (
                  str(row.iloc[car_col_b - 1]).strip()
                  if pd.notna(row.iloc[car_col_b - 1])
                  else ""
              )
              if c_val in car_color_map:
                color_hex = car_color_map[c_val]
                fill = PatternFill(
                    start_color=color_hex,
                    end_color=color_hex,
                    fill_type="solid",
                )
                excel_row = row_idx + 2
                if name_col_b:
                  ws.cell(row=excel_row, column=name_col_b).fill = fill
                ws.cell(row=excel_row, column=car_col_b).fill = fill

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

          setup_sheet_print(ws)

        out1.seek(0)
        st.success(
            "✅"
            " 初步比對完成！已成功對照『仰懷宗範』導覽資格、去除無用欄位並自動標記共乘/時間。"
        )
        st.download_button(
            label="📥 下載待確認班表 (.xlsx)",
            data=out1,
            file_name=f"工作班表_{target_date_1}_待確認.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
    except Exception as e:
      st.error(f"❌ 處理失敗：{str(e)}")

# ==========================================
# 階段二：生成最終表格 C 與所有列印頁籤 (含掛退單接駁判斷)
# ==========================================
with tab2:
  st.header("2️⃣ 上傳最終班表，生成表格 C 與所有列印 Sheet")
  file_final = st.file_uploader(
      "上傳您已排好/確認好的工作班表 (.xlsx)",
      type=["xlsx", "xls"],
      key="f_final",
  )
  target_date_2 = st.text_input(
      "排班日期 / Sheet 名稱 (如: 1022)", value="1022", key="t2_date"
  )

  if file_final:
    try:
      xls_f = pd.ExcelFile(file_final)
      sheet_use = (
          target_date_2
          if target_date_2 in xls_f.sheet_names
          else xls_f.sheet_names[0]
      )
      df_final = pd.read_excel(xls_f, sheet_name=sheet_use)

      if st.button("🚀 一鍵生成完整列印 Sheet", type="primary"):
        b_id_col = (
            "義工編號"
            if "義工編號" in df_final.columns
            else (
                "編號" if "編號" in df_final.columns else df_final.columns[0]
            )
        )
        task_col = next(
            (c for c in df_final.columns if "任" in str(c) or "任務" in str(c)),
            None,
        )
        hall_col = next(
            (c for c in df_final.columns if "館" in str(c)), None
        )
        shao_col = next(
            (c for c in df_final.columns if "哨" in str(c)), None
        )

        out2 = io.BytesIO()
        with pd.ExcelWriter(out2, engine="openpyxl") as writer:

          # 1. 大班表
          df_final.to_excel(writer, sheet_name=target_date_2, index=False)
          setup_sheet_print(writer.sheets[target_date_2])

          # 2. 館員用 (連動公式，包含 仰懷宗範、常設定導、特展定導、菩薩券、午齋)
          df_staff = df_final.copy()
          for col_needed in [
              "常設定導",
              "特展定導",
              "仰懷宗範",
              "菩薩券",
              "午齋",
          ]:
            if col_needed not in df_staff.columns:
              df_staff[col_needed] = 1 if col_needed == "午齋" else ""

          for row_idx in range(len(df_staff)):
            excel_row = row_idx + 2
            for col_idx, col_name in enumerate(df_staff.columns):
              if col_name in df_final.columns:
                col_letter = get_column_letter(
                    list(df_final.columns).index(col_name) + 1
                )
                df_staff.iloc[row_idx, col_idx] = (
                    f"='{target_date_2}'!{col_letter}{excel_row}"
                )

          df_staff.to_excel(writer, sheet_name="館員用", index=False)
          setup_sheet_print(writer.sheets["館員用"])

          # 3. 本館義工用齋登記
          if hall_col:
            df_meal_src = df_final[
                df_final[hall_col].astype(str).str.contains("本")
            ].copy()
          else:
            df_meal_src = df_final.copy()

          df_meal_out = pd.DataFrame()
          df_meal_out["序號"] = range(1, len(df_meal_src) + 1)
          df_meal_out["編號"] = (
              df_meal_src[b_id_col].values
              if b_id_col in df_meal_src.columns
              else ""
          )
          df_meal_out["精舍"] = (
              df_meal_src["精舍"].values
              if "精舍" in df_meal_src.columns
              else ""
          )

          name_list = []
          for _, r in df_meal_src.iterrows():
            nm = (
                str(r.get("姓名", ""))
                if pd.notna(r.get("姓名", ""))
                else ""
            )
            fm = (
                str(r.get("法名", ""))
                if pd.notna(r.get("法名", ""))
                else ""
            )
            full_n = f"{nm}({fm})" if fm and fm != "nan" else nm
            name_list.append(full_n)
          df_meal_out["姓名(法名)"] = name_list
          df_meal_out["午齋"] = 1

          remarks = []
          for _, r in df_meal_src.iterrows():
            tk = str(r.get(task_col, "")) if task_col else ""
            sh = str(r.get(shao_col, "")) if shao_col else ""
            if (
                any(
                    k in tk
                    for k in ["禪茶", "校外教學", "導覽", "特展", "特護"]
                )
                and "組長" not in tk
            ):
              remarks.append(tk)
            else:
              remarks.append(sh)
          df_meal_out["備註"] = remarks

          df_meal_out.to_excel(
              writer, sheet_name="本館義工用齋登記", index=False
          )
          ws_meal = writer.sheets["本館義工用齋登記"]
          last_r = len(df_meal_out) + 2
          ws_meal.cell(row=last_r, column=4, value="合計").alignment = Alignment(
              horizontal="center"
          )
          ws_meal.cell(
              row=last_r, column=5, value=f"=SUM(E2:E{last_r-1})"
          ).alignment = Alignment(horizontal="center")
          setup_sheet_print(ws_meal, orientation="portrait")

          # 4. 分館義工輪值表
          s1_col = next(
              (c for c in df_final.columns if "第一哨" in str(c)), None
          )
          s2_col = next(
              (c for c in df_final.columns if "第二哨" in str(c)), None
          )
          s3_col = next(
              (c for c in df_final.columns if "第三哨" in str(c)), None
          )

          if hall_col:
            df_sub_src = df_final[
                df_final[hall_col].astype(str).str.contains("分")
            ].copy()
          else:
            df_sub_src = df_final.copy()

          sub_rows = []
          for idx, (_, r) in enumerate(df_sub_src.iterrows(), 1):
            nm = str(r.get("姓名", "")) if pd.notna(r.get("姓名", "")) else ""
            fm = str(r.get("法名", "")) if pd.notna(r.get("法名", "")) else ""
            full_nm = f"{nm}({fm})" if fm and fm != "nan" else nm

            sub_rows.append({
                "序號": idx,
                "館別": "分館",
                "哨別": r.get(shao_col, "") if shao_col else "",
                "任務": r.get(task_col, "") if task_col else "",
                "編號": r.get(b_id_col, ""),
                "精舍": r.get("精舍", ""),
                "姓名(法名)": full_nm,
                "第一哨\nA:09:00-10:30\nB:10:30-12:00": (
                    r.get(s1_col, "") if s1_col else ""
                ),
                "第二哨\nA:12:00-13:00\nB:13:00-14:00": (
                    r.get(s2_col, "") if s2_col else ""
                ),
                "第三哨\nA:14:00-15:30\nB:15:30-17:00": (
                    r.get(s3_col, "") if s3_col else ""
                ),
            })

          df_sub_out = pd.DataFrame(sub_rows)
          df_sub_out.to_excel(
              writer, sheet_name="分館義工輪值表", index=False
          )
          ws_sub = writer.sheets["分館義工輪值表"]
          for row in ws_sub.iter_rows():
            for cell in row:
              cell.alignment = Alignment(
                  wrap_text=True, vertical="center", horizontal="center"
              )
          setup_sheet_print(ws_sub)

          # 5. 組長班表
          if hall_col and shao_col:
            df_lead_a = df_final[
                (df_final[hall_col].astype(str).str.contains("本"))
                & (df_final[shao_col].astype(str) == "A")
            ].copy()
            if len(df_lead_a) > 0:
              df_lead_a.to_excel(
                  writer, sheet_name="本館A哨組長表", index=False
              )
              setup_sheet_print(writer.sheets["本館A哨組長表"])

            df_lead_b = df_final[
                (df_final[hall_col].astype(str).str.contains("本"))
                & (df_final[shao_col].astype(str) == "B")
            ].copy()
            if len(df_lead_b) > 0:
              df_lead_b.to_excel(
                  writer, sheet_name="本館B哨組長表", index=False
              )
              setup_sheet_print(writer.sheets["本館B哨組長表"])

          # 6. 兩館接駁乘車名單
          gua_col = next(
              (c for c in df_final.columns if "掛單" in str(c)), None
          )
          tui_col = next(
              (c for c in df_final.columns if "退單" in str(c)), None
          )

          cur_md = parse_m_d(target_date_2)
          shuttle_rows = []

          for _, r in df_final.iterrows():
            has_shuttle = False
            to_sub = ""
            to_main = ""

            if gua_col and tui_col and cur_md:
              g_md = parse_m_d(r.get(gua_col))
              t_md = parse_m_d(r.get(tui_col))

              if g_md and t_md:
                c_val = cur_md[0] * 100 + cur_md[1]
                g_val = g_md[0] * 100 + g_md[1]
                t_val = t_md[0] * 100 + t_md[1]

                if g_val <= c_val <= t_val:
                  has_shuttle = True
                  if c_val == g_val and c_val == t_val:
                    to_main = "1"
                    to_sub = "1"
                  elif c_val == g_val:
                    to_main = "1"
                  elif c_val == t_val:
                    to_sub = "1"
                  else:
                    to_main = "1"
                    to_sub = "1"

            if not has_shuttle:
              m_to_d = str(r.get("主體到大博", "")).strip()
              d_to_m = str(r.get("大博到主體", "")).strip()
              if m_to_d == "1" or d_to_m == "1":
                has_shuttle = True
                to_sub = "1" if m_to_d == "1" else ""
                to_main = "1" if d_to_m == "1" else ""

            if has_shuttle:
              r_dict = r.to_dict()
              r_dict["主體到大博"] = to_sub
              r_dict["大博到主體"] = to_main
              shuttle_rows.append(r_dict)

          df_shuttle = pd.DataFrame(shuttle_rows)
          if len(df_shuttle) == 0:
            df_shuttle = pd.DataFrame(columns=df_final.columns)

          df_shuttle.to_excel(
              writer, sheet_name="兩館接駁乘車名單", index=False
          )
          setup_sheet_print(writer.sheets["兩館接駁乘車名單"])

          # 7. 佈告欄大班表
          df_bulletin = process_bulletin_phone(df_final)
          for t_col in [
              "發心起始時間",
              "發心開始時間",
              "發心結束時間",
              "發心截止時間",
          ]:
            if t_col in df_bulletin.columns:
              df_bulletin[t_col] = ""
          df_bulletin.to_excel(writer, sheet_name="佈告欄大班表", index=False)
          setup_sheet_print(writer.sheets["佈告欄大班表"])

          # 8. 票口導覽資訊
          date_str = (
              f"{target_date_2[:2]}/{target_date_2[2:]}"
              if len(target_date_2) >= 4
              else target_date_2
          )
          piaokou_rows = [
              [
                  "參訪日期",
                  "星期",
                  "時間",
                  "主題",
                  "團體名稱",
                  "人數",
                  "預排導覽員",
              ],
              ["", "", "", "", "", "", ""],
              [
                  "",
                  date_str,
                  "本館",
                  "定導",
                  "哨點",
                  "第一哨\nA:09:00-10:30\nB:10:30-12:10",
                  "第二哨\nA:12:10-13:10\nB:13:10-14:00",
                  "第三哨\nA:14:00-15:15\nB:15:15-16:30",
              ],
          ]
          if task_col and hall_col and shao_col:
            guide_df = df_final[
                df_final[task_col].astype(str).str.contains("導覽|定導")
            ]
            for _, grow in guide_df.iterrows():
              piaokou_rows.append([
                  "",
                  "",
                  str(grow.get(hall_col, "")),
                  "定導",
                  str(grow.get(shao_col, "")),
                  str(grow.get(s1_col, "")) if s1_col else "",
                  str(grow.get(s2_col, "")) if s2_col else "",
                  str(grow.get(s3_col, "")) if s3_col else "",
              ])

          df_pk = pd.DataFrame(piaokou_rows)
          df_pk.to_excel(
              writer, sheet_name="票口導覽資訊", index=False, header=False
          )
          ws_pk = writer.sheets["票口導覽資訊"]
          for row in ws_pk.iter_rows():
            for cell in row:
              cell.alignment = Alignment(
                  wrap_text=True, vertical="center", horizontal="center"
              )
          setup_sheet_print(ws_pk)

        out2.seek(0)
        st.balloons()
        st.success("🎉 列印檔生成成功！已自動包含仰懷宗範與智慧接駁判斷。")
        st.download_button(
            label="📥 下載表格 C 列印檔 (.xlsx)",
            data=out2,
            file_name=f"表格C_完整排班與列印檔_{target_date_2}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    except Exception as e:
      st.error(f"❌ 處理失敗：{str(e)}")
