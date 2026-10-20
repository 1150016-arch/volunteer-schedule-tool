import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="義工排班與自動更新系統", page_icon="📅", layout="wide"
)

st.title("📅 義工排班與自動更新系統")
st.markdown("本系統採 **兩階段處理**：初步 VLOOKUP 比對 ➔ 模板排班後直接原檔更新。")

tab1, tab2 = st.tabs(["1️⃣ 階段一：初步 VLOOKUP 比對與帶入", "2️⃣ 階段二：上傳模板原檔，直接更新並下載"])


def load_clean_b(file):
  df = pd.read_excel(file)
  if "義工編號" not in df.columns and "編號" not in df.columns:
    if len(df) > 0 and any(
        df.iloc[0].astype(str).str.contains("義工編號|編號|姓名")
    ):
      df.columns = df.iloc[0]
      df = df[1:].reset_index(drop=True)
  return df


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
        st.error(f"⚠️ 在表格 A 中找不到名稱為 **{target_date_1}** 的工作表！")
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
# 階段二：直接開啟模板原檔進行更新（完全保留格式）
# ==========================================
with tab2:
  st.header("2️⃣ 上傳排好班的模板檔，直接原檔更新")
  file_template = st.file_uploader(
      "上傳排好班的模板檔案 (.xlsx 或 .xls)",
      type=["xlsx", "xls"],  # <--- 已放寬允許 .xls 上傳
      key="f_tpl",
  )

  if file_template:
    if st.button("🚀 在原模板檔上直接自動更新並下載", type="primary"):
      try:
        # 直接讀取 Excel Workbook (完全保留原始排版與格式)
        wb = openpyxl.load_workbook(file_template)

        # 1. 讀取「館員用」資料（從第 3 列開始讀取真實數據）
        ws_staff = (
            wb["館員用"]
            if "館員用" in wb.sheetnames
            else wb[wb.sheetnames[0]]
        )

        staff_data = []
        for r in range(3, ws_staff.max_row + 1):
          hall_val = ws_staff.cell(row=r, column=2).value
          shao_val = ws_staff.cell(row=r, column=3).value
          task_val = ws_staff.cell(row=r, column=4).value
          id_val = ws_staff.cell(row=r, column=5).value
          js_val = ws_staff.cell(row=r, column=6).value
          name_val = ws_staff.cell(row=r, column=7).value

          if id_val or name_val:
            staff_data.append({
                "館別": str(hall_val or ""),
                "哨別": str(shao_val or ""),
                "任務": str(task_val or ""),
                "編號": str(id_val or ""),
                "精舍": str(js_val or ""),
                "姓名(法名)": str(name_val or ""),
            })

        # 2. 更新「用齋用」Sheet
        if "用齋用" in wb.sheetnames:
          ws_meal = wb["用齋用"]
          main_meal = [
              d for d in staff_data if "本" in d["館別"] or d["館別"] == ""
          ]
          sub_meal = [d for d in staff_data if "分" in d["館別"]]

          # 本館用齋（寫入 A-F 欄）
          for idx, item in enumerate(main_meal, start=3):
            ws_meal.cell(row=idx, column=1, value=idx - 2)
            ws_meal.cell(row=idx, column=2, value=item["編號"])
            ws_meal.cell(row=idx, column=3, value=item["精舍"])
            ws_meal.cell(row=idx, column=4, value=item["姓名(法名)"])
            ws_meal.cell(row=idx, column=5, value=1)

            tk = item["任務"]
            if any(
                k in tk for k in ["禪茶", "校外教學", "導覽", "特展", "團★"]
            ) and ("組長" not in tk):
              ws_meal.cell(row=idx, column=6, value=tk)
            else:
              ws_meal.cell(row=idx, column=6, value=item["哨別"])

          # 分館用齋（寫入 I-N 欄）
          for idx, item in enumerate(sub_meal, start=3):
            ws_meal.cell(row=idx, column=9, value=idx - 2)
            ws_meal.cell(row=idx, column=10, value=item["編號"])
            ws_meal.cell(row=idx, column=11, value=item["精舍"])
            ws_meal.cell(row=idx, column=12, value=item["姓名(法名)"])
            ws_meal.cell(row=idx, column=13, value=1)
            ws_meal.cell(row=idx, column=14, value=item["哨別"])

        # 3. 更新「貼後上傳守護網用」Sheet
        if "貼後上傳守護網用(先糾錯，再安任務)" in wb.sheetnames:
          ws_guard = wb["貼後上傳守護網用(先糾錯，再安任務)"]
          for idx, item in enumerate(staff_data, start=2):
            ws_guard.cell(row=idx, column=1, value=item["編號"])
            full_nm = item["姓名(法名)"]
            nm_only = full_nm.split("(")[0] if "(" in full_nm else full_nm
            fm_only = (
                full_nm.split("(")[1].replace(")", "")
                if "(" in full_nm
                else ""
            )

            ws_guard.cell(row=idx, column=2, value=nm_only)
            ws_guard.cell(row=idx, column=3, value=fm_only)
            ws_guard.cell(
                row=idx,
                column=5,
                value="本館" if "本" in item["館別"] else "分館",
            )
            ws_guard.cell(row=idx, column=6, value=item["任務"])

        # 存檔
        out_saved = io.BytesIO()
        wb.save(out_saved)
        out_saved.seek(0)

        st.session_state["updated_excel"] = out_saved.getvalue()
        st.balloons()
        st.success("🎉 已成功在您原本的模板上更新完畢！格式 100% 完美保留！")

      except Exception as e:
        st.error(f"❌ 處理失敗：{str(e)}")

    if "updated_excel" in st.session_state:
      st.download_button(
          label="📥 下載更新後的原模板班表 (.xlsx)",
          data=st.session_state["updated_excel"],
          file_name="排班檔_已原檔更新.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )
