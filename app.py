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
target_date = st
