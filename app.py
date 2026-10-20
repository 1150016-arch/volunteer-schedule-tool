from datetime import datetime
import io
import re
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="義工排班與列印檔自動生成系統",
    page_icon="📅",
    layout="wide",
)

st.title("📅 義工排班與列印檔自動生成系統")
st.markdown(
    "本系統支援 **「模板排班後自動補齊與雙檔產出」**："
    "上傳完成排班的模板檔，自動更新相關頁籤，並同步生成影印專用 Excel！"
)

tab1, tab2 = st.tabs([
    "1️⃣ 階段一：初步 VLOOKUP 比對與帶入",
    "2️⃣ 階段二：上傳排班模板 ➔ 自動補齊 & 生成影印專用檔",
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
  df = pd.read_excel(file)
  if "義工編號" not in df.columns and "編號" not in df.columns:
    if len(df) > 0 and any(
        df.iloc[0].astype(str).str.contains("義工編號|編號|姓名")
    ):
      df.columns = df.iloc[0]
      df = df[1:].reset_index(drop=True)
  return df


def setup_sheet_print(ws, orientation="landscape"):
  """設定 Excel 列印頁面屬性：Fit to 1 Page Wide (A4)"""
  ws.page_setup.orientation = (
      ws.ORIENTATION_LANDSCAPE
      if orientation == "landscape"
      else ws.ORIENTATION_PORTRAIT
  )
  ws.page_setup.paperSize = ws.PAPERSIZE_A4
  ws.sheet_properties.pageSetUpPr.fitToPage = True
  ws.page_setup.fitToWidth = 1
  ws.page_setup.fitToHeight = 0

  thin_border = Border(
      left=Side(style="thin", color="000000"),
      right=Side(style="thin", color="000000"),
      top=Side(style="thin", color="000000"),
      bottom=Side(style="thin", color="000000"),
  )

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


def parse_m_d(date_str):
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
# 階段一：初步 VLOOKUP 比對與帶入
# ==========================================
with tab1:
  st.header("1️⃣ 匯入原始檔案，生成待確認班表")
  target_date_1 = st.text_input(
      "排班日期 / Sheet 名稱 (如: 1021)", value="1021", key="t1_date"
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
        st.error(f"⚠️ 在表格 A 中找不到名稱為 {target_date_1} 的工作表！")
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
                for c in [
                    "特殊需求",
                    "備註",
                    "備註1",
                    "常態備註",
                    "個人備註",
                ]
                if c in df_a_unique.columns
            ),
            None,
        )
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

        out1 = io.BytesIO()
        with pd.ExcelWriter(out1, engine="openpyxl") as writer:
          df_b_cleaned.to_excel(
              writer, sheet_name=target_date_1, index=False
          )
          setup_sheet_print(writer.sheets[target_date_1])

        out1.seek(0)
        st.success("✅ 第一階段待確認班表生成成功！")
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
# 階段二：模板排班後自動補齊與雙檔生成
# ==========================================
with tab2:
  st.header(
      "2️⃣ 上傳排班完成之模板 Excel ➔ 自動更新後續 Sheet 並生成影印檔"
  )
  file_template = st.file_uploader(
      "上傳排好班的模板檔案 (如: 1021練習.xls / .xlsx)",
      type=["xlsx", "xls"],
      key="f_tpl",
  )
  date_code = st.text_input(
      "排班日期代碼 (如: 1021)", value="1021", key="t2_date"
  )

  if file_template:
    try:
      # 讀取完整 Excel 頁籤
      xls = pd.ExcelFile(file_template)
      sheet_names = xls.sheet_names

      # 找出主數據 Sheet
      main_sheet = (
          date_code
          if date_code in sheet_names
          else ("1021" if "1021" in sheet_names else sheet_names[0])
      )
      df_main = pd.read_excel(xls, sheet_name=main_sheet)

      # 讀取館員用 Sheet
      df_staff_raw = (
          pd.read_excel(xls, sheet_name="館員用")
          if "館員用" in sheet_names
          else None
      )

      if st.button("🚀 一鍵自動更新模板頁籤並生成雙檔", type="primary"):
        b_id_col = next(
            (
                c
                for c in ["義工編號", "編號"]
                if c in df_main.columns or c in df_main.iloc[0].values
            ),
            df_main.columns[0],
        )

        # ----------------------------------------
        # 1. 整理「館員用」與數據提取
        # ----------------------------------------
        if df_staff_raw is not None:
          # 若第一列是標題
          if "姓名(法名)" in list(df_staff_raw.columns) or "義工輪值表" in str(
              df_staff_raw.columns[0]
          ):
            header_idx = 1 if len(df_staff_raw) > 0 else 0
            df_staff = df_staff_raw.copy()
            if "館別" not in df_staff.columns and len(df_staff) > 1:
              df_staff.columns = df_staff.iloc[0]
              df_staff = df_staff[1:].reset_index(drop=True)
          else:
            df_staff = df_staff_raw.copy()
        else:
          df_staff = df_main.copy()

        # ----------------------------------------
        # 2. 自動生成/更新「用齋用」Sheet
        # ----------------------------------------
        # 本館用齋名單
        main_hall_mask = (
            df_staff["館別"].astype(str).str.contains("本")
            if "館別" in df_staff.columns
            else pd.Series([True] * len(df_staff))
        )
        df_meal_main_src = df_staff[main_hall_mask].copy()

        meal_main_rows = []
        for idx, (_, r) in enumerate(df_meal_main_src.iterrows(), 1):
          nm = str(r.get("姓名(法名)", r.get("姓名", "")))
          num = str(r.get("編號", r.get("義工編號", "")))
          js = str(r.get("精舍", ""))
          shao = str(r.get("哨\n別", r.get("哨別", r.get("哨", ""))))
          task = str(r.get("任務", ""))

          # 備註優先判定特殊任務，否則帶入哨別
          if any(
              k in task for k in ["禪茶", "校外教學", "導覽", "特展", "團★"]
          ) and ("組長" not in task):
            rem = task
          else:
            rem = shao if shao else "A"

          meal_main_rows.append({
              "序號": idx,
              "編號": num,
              "精舍": js,
              "姓名（法名）": nm,
              "午齋": 1,
              "備註": rem,
          })
        df_meal_main = pd.DataFrame(meal_main_rows)

        # 分館用齋名單
        sub_hall_mask = (
            df_staff["館別"].astype(str).str.contains("分")
            if "館別" in df_staff.columns
            else pd.Series([False] * len(df_staff))
        )
        df_meal_sub_src = df_staff[sub_hall_mask].copy()

        meal_sub_rows = []
        for idx, (_, r) in enumerate(df_meal_sub_src.iterrows(), 1):
          nm = str(r.get("姓名(法名)", r.get("姓名", "")))
          num = str(r.get("編號", r.get("義工編號", "")))
          js = str(r.get("精舍", ""))
          shao = str(r.get("哨\n別", r.get("哨別", r.get("哨", ""))))

          meal_sub_rows.append({
              "序號": idx,
              "編號": num,
              "精舍": js,
              "姓名（法名）": nm,
              "午齋": 1,
              "備註": shao if shao else "A",
          })
        df_meal_sub = pd.DataFrame(meal_sub_rows)

        # ----------------------------------------
        # 3. 自動更新「貼後上傳守護網用(先糾錯，再安任務)」
        # ----------------------------------------
        guard_rows = []
        for _, r in df_main.iterrows():
          num = r.get("義工編號", r.get("編號", ""))
          nm = r.get("姓名", "")
          fm = r.get("法名", "")
          sex = r.get("性別", "")
          hall = r.get("館", r.get("館別", ""))
          hall_str = "本館" if "本" in str(hall) else ("分館" if "分" in str(hall) else "")
          task = r.get("任", r.get("任務", ""))

          guard_rows.append({
              "義工編號": num,
              "姓名": nm,
              "法名": fm,
              "性別": sex,
              "地點": hall_str,
              "任務": task if pd.notna(task) else "",
          })
        df_guard = pd.DataFrame(guard_rows)

        # ----------------------------------------
        # 4. 自動更新「兩館接駁」
        # ----------------------------------------
        shuttle_rows = []
        cur_md = parse_m_d(date_code)
        gua_col = next(
            (c for c in df_main.columns if "掛單" in str(c)), None
        )
        tui_col = next(
            (c for c in df_main.columns if "退單" in str(c)), None
        )

        for idx, (_, r) in enumerate(df_main.iterrows(), 1):
          has_shuttle = False
          if gua_col and tui_col and cur_md:
            g_md = parse_m_d(r.get(gua_col))
            t_md = parse_m_d(r.get(tui_col))
            if g_md and t_md:
              c_val = cur_md[0] * 100 + cur_md[1]
              g_val = g_md[0] * 100 + g_md[1]
              t_val = t_md[0] * 100 + t_md[1]
              if g_val <= c_val <= t_val:
                has_shuttle = True

          if has_shuttle or str(r.get("主體到大博", "")) == "1" or str(r.get("大博到主體", "")) == "1":
            nm = r.get("姓名", "")
            fm = r.get("法名", "")
            full_nm = f"{nm}({fm})" if pd.notna(fm) and str(fm) != "nan" else nm
            shuttle_rows.append({
                "NO.": idx,
                "義工編號": r.get("義工編號", r.get("編號", "")),
                "精舍": r.get("精舍", ""),
                "姓名(法名)": full_nm,
                "電話": r.get("手機", r.get("電話", "")),
                "館別": r.get("館", r.get("館別", "")),
            })
        df_shuttle = pd.DataFrame(shuttle_rows)

        # ----------------------------------------
        # 5. 自動更新「票口導覽資訊」
        # ----------------------------------------
        date_display = (
            f"{date_code[:2]}/{date_code[2:]}"
            if len(date_code) >= 4
            else date_code
        )
        pk_rows = [
            ["參訪日期", "星期", "時間", "主題", "團體名稱", "人數", "預排導覽員"],
            ["", "", "", "", "", "", ""],
            [
                "",
                f"{date_display} 定導 哨點",
                "第一哨\nA:09:00-10:30\nB:10:30-12:10",
                "第二哨\nA:12:10-13:10\nB:13:10-14:00",
                "第三哨\nA:14:00-15:15\nB:15:15-16:30",
                "",
                "",
            ],
        ]

        if "第一哨\nA:09:00-10:30\nB:10:30-12:10" in df_staff.columns:
          guide_rows = df_staff[
              df_staff["任務"].astype(str).str.contains("導覽|定導")
          ]
          for _, grow in guide_rows.iterrows():
            pk_rows.append([
                "",
                "",
                str(grow.get("館別", "")),
                "定導",
                str(grow.get("哨\n別", grow.get("哨別", ""))),
                str(grow.get("第一哨\nA:09:00-10:30\nB:10:30-12:10", "")),
                str(grow.get("第二哨\nA:12:10-13:10\nB:13:10-14:00", "")),
                str(grow.get("第三哨\nA:14:00-15:15\nB:15:15-16:30", "")),
            ])
        df_pk = pd.DataFrame(pk_rows)

        # --------------------------------------------------
        # 匯出 1：更新後的完整排班備份檔 (.xlsx)
        # --------------------------------------------------
        out_updated = io.BytesIO()
        with pd.ExcelWriter(out_updated, engine="openpyxl") as writer:
          # 寫入原本主要工作頁籤
          for sn in sheet_names:
            df_curr = pd.read_excel(xls, sheet_name=sn)
            df_curr.to_excel(writer, sheet_name=sn, index=False)

          # 更新/覆蓋自動計算頁籤
          df_meal_main.to_excel(writer, sheet_name="用齋用", index=False)
          df_guard.to_excel(
              writer,
              sheet_name="貼後上傳守護網用(先糾錯，再安任務)",
              index=False,
          )
          df_shuttle.to_excel(writer, sheet_name="兩館接駁", index=False)
          df_pk.to_excel(
              writer, sheet_name="票口導覽資訊", index=False, header=False
          )

        out_updated.seek(0)

        # --------------------------------------------------
        # 匯出 2：影印專用 Excel 檔 (完全縮放與專業樣式)
        # --------------------------------------------------
        out_print = io.BytesIO()
        with pd.ExcelWriter(out_print, engine="openpyxl") as writer:

          # 1. 館員用
          df_staff.to_excel(writer, sheet_name="館員用", index=False)
          setup_sheet_print(writer.sheets["館員用"])

          # 2. 佈告欄大班表 (隱藏組員電話供簽到)
          df_bulletin = df_staff.copy()
          phone_col = next(
              (c for c in df_bulletin.columns if "電話" in str(c)), None
          )
          if phone_col:
            new_p = []
            for _, r in df_bulletin.iterrows():
              tk = str(r.get("任務", ""))
              if "組長" in tk:
                new_p.append(r.get(phone_col, ""))
              else:
                new_p.append("")
            df_bulletin[phone_col] = new_p

          df_bulletin.to_excel(writer, sheet_name="佈告欄大班表", index=False)
          setup_sheet_print(writer.sheets["佈告欄大班表"])

          # 3. 本館 A 哨組長表 / B 哨組長表
          if "哨\n別" in df_staff.columns or "哨別" in df_staff.columns:
            s_col = "哨\n別" if "哨\n別" in df_staff.columns else "哨別"
            df_lead_a = df_staff[
                (df_staff["館別"].astype(str).str.contains("本"))
                & (df_staff[s_col].astype(str) == "A")
            ]
            if len(df_lead_a) > 0:
              df_lead_a.to_excel(
                  writer, sheet_name="本館A哨組長表", index=False
              )
              setup_sheet_print(writer.sheets["本館A哨組長表"])

            df_lead_b = df_staff[
                (df_staff["館別"].astype(str).str.contains("本"))
                & (df_staff[s_col].astype(str) == "B")
            ]
            if len(df_lead_b) > 0:
              df_lead_b.to_excel(
                  writer, sheet_name="本館B哨組長表", index=False
              )
              setup_sheet_print(writer.sheets["本館B哨組長表"])

          # 4. 團導表
          if "團導" in sheet_names:
            df_tuandao = pd.read_excel(xls, sheet_name="團導")
            df_tuandao.to_excel(writer, sheet_name="團導表", index=False)
            setup_sheet_print(writer.sheets["團導表"])

          # 5. 午齋單 (本館與分館併列或單列)
          df_meal_main.to_excel(
              writer, sheet_name="本館用齋單", index=False
          )
          ws_m = writer.sheets["本館用齋單"]
          lr = len(df_meal_main) + 2
          ws_m.cell(row=lr, column=4, value="合計").alignment = Alignment(
              horizontal="center"
          )
          ws_m.cell(
              row=lr, column=5, value=f"=SUM(E2:E{lr-1})"
          ).alignment = Alignment(horizontal="center")
          setup_sheet_print(ws_m, orientation="portrait")

          # 6. 票口導覽資訊
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

        out_print.seek(0)

        st.balloons()
        st.success("🎉 自動更新與影印檔生成完成！請點選下方按鈕下載：")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
          st.download_button(
              label="📥 下載 1：更新後完整備份檔 (.xlsx)",
              data=out_updated,
              file_name=f"排班檔_{date_code}_已自動更新.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )
        with col_dl2:
          st.download_button(
              label="🖨️ 下載 2：影印專用工作簿 (.xlsx)",
              data=out_print,
              file_name=f"影印專用班表_{date_code}.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

    except Exception as e:
      st.error(f"❌ 處理失敗：{str(e)}")
