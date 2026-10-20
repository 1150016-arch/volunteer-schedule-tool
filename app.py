import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="義工排班自動化系統", page_icon="📅", layout="wide"
)

st.title("📅 義工排班自動化系統")
st.markdown(
    "本系統協助完成 **階段一：初步 VLOOKUP 比對與資料帶入**，方便後續複製貼上至排班模板。"
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


# ==========================================
# 階段一：初步 VLOOKUP 比對與帶入
# ==========================================
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

      # 剔除不需要的欄位
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

      out1 = io.BytesIO()
      with pd.ExcelWriter(out1, engine="openpyxl") as writer:
        df_b_cleaned.to_excel(writer, sheet_name=target_date_1, index=False)

      out1.seek(0)
      st.success("✅ 第一階段待確認班表生成成功！")
      st.download_button(
          label="📥 下載待確認班表 (.xlsx)",
          data=out1,
          file_name=f"工作班表_{target_date_1}_待確認.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )
  except Exception as e:
    st.error(f"❌ 處理失敗：{str(e)}")
