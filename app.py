import streamlit as st
import openpyxl
from datetime import datetime
from io import BytesIO
import base64
from decimal import Decimal, ROUND_HALF_UP

# === 設定 ===
TEMPLATE = "検査報告書_(株)広島フォーマット.xlsx"  # 同じフォルダにテンプレートExcelを置く

# --- ExcelのROUNDと同じ四捨五入（0.5は常に切り上げ）---
def excel_round(value, digits):
    q = '1.' + '0' * digits
    return float(Decimal(str(value)).quantize(Decimal(q), rounding=ROUND_HALF_UP))

st.title("📘 気密試験記録 入力フォーム")

# --- 入力項目 ---
系統名 = st.text_input("系統名")
試験圧力 = st.text_input("試験圧力", placeholder="例：0.8MPa以上")
試験範囲 = st.text_input("試験範囲")
試験媒体 = st.text_input("試験媒体")
放置時間 = st.text_input("放置時間", placeholder="例：30min以上")
使用機器No = st.text_input("使用圧力計機器No.")
測定場所 = st.text_input("測定場所")

# --- 開始日時 ---
st.subheader("開始日時")
col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    開始日 = st.date_input("日付", key="start_date")
with col2:
    開始時 = st.text_input("時", value="", key="start_hour", placeholder="例：9")
with col3:
    開始分 = st.text_input("分", value="", key="start_minute", placeholder="例：30")

# --- 終了日時 ---
st.subheader("終了日時")
col4, col5, col6 = st.columns([2, 1, 1])
with col4:
    終了日 = st.date_input("日付", key="end_date")
with col5:
    終了時 = st.text_input("時", value="", key="end_hour", placeholder="例：10")
with col6:
    終了分 = st.text_input("分", value="", key="end_minute", placeholder="例：15")

# --- 測定値入力 ---
st.subheader("測定値入力")
col5, col6 = st.columns(2)
with col5:
    P1 = st.text_input("開始圧力 (MPa)", placeholder="例：0.8760")
with col6:
    T1 = st.text_input("開始温度 (℃)", placeholder="例：20.1")

col7, col8 = st.columns(2)
with col7:
    P2p = st.text_input("終了圧力 (MPa)", placeholder="例：0.8756")
with col8:
    T2 = st.text_input("終了温度 (℃)", placeholder="例：19.3")

試験実施者 = st.text_input("試験実施者")

# --- 数値変換 ---
def safe_float(v):
    try:
        return float(v.strip()) if v else None
    except:
        return None

P1 = safe_float(P1)
T1 = safe_float(T1)
P2p = safe_float(P2p)
T2 = safe_float(T2)

# --- 判定・保存 ---
if st.button("判定・保存"):
    if None in (P1, T1, P2p, T2):
        st.warning("⚠ 圧力・温度のすべてを入力してください。")
    else:
        try:
            # --- 日時生成（未入力時は00:00扱い）---
            try:
                開始日時 = datetime.combine(
                    開始日,
                    datetime.strptime(f"{int(開始時 or 0):02d}:{int(開始分 or 0):02d}", "%H:%M").time()
                )
                終了日時 = datetime.combine(
                    終了日,
                    datetime.strptime(f"{int(終了時 or 0):02d}:{int(終了分 or 0):02d}", "%H:%M").time()
                )
            except:
                開始日時 = datetime.combine(開始日, datetime.strptime("00:00", "%H:%M").time())
                終了日時 = datetime.combine(終了日, datetime.strptime("00:00", "%H:%M").time())

            # --- 補正後圧力（Excel式と同一）---
            T1_K = T1 + 273.15
            T2_K = T2 + 273.15
            P2_corr_raw = ((P1 + 0.1013) * (T2_K / T1_K)) - 0.1013
            P2_corr = excel_round(P2_corr_raw, 3)

            # --- Excelと同じ丸め処理での判定 ---
            ΔP_dec = Decimal(str(P1)) - Decimal(str(P2_corr))
            ΔP = float(Decimal(ΔP_dec).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP))
            判定範囲 = float(Decimal(str(P1 * 0.01)).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP))
            合否 = "合格" if abs(ΔP) <= 判定範囲 else "不合格"
            色 = "green" if 合否 == "合格" else "red"

            # --- 結果表示 ---
            st.markdown("## 📊 計算結果（Excelと完全一致）")
            st.write(f"- 補正後終了圧力 P2_corr: **{P2_corr:.3f} MPa**")
            st.write(f"- 圧力変化量 ΔP（開始−補正後）: **{ΔP:.3f} MPa**")
            st.write(f"- 判定範囲: ±**{判定範囲:.3f} MPa**")
            st.markdown(f"### <span style='color:{色};'>判定結果: {合否}</span>", unsafe_allow_html=True)

            # --- Excel出力 ---
            wb = openpyxl.load_workbook(TEMPLATE)
            ws = wb["気密試験記録"]

            def write(ws, cell, value):
                """結合セル対応"""
                try:
                    ws[cell].value = value
                except AttributeError:
                    r = ws[cell].row
                    c = ws[cell].column
                    ws.cell(row=r, column=c, value=value)

            # --- Excel書き込み ---
            write(ws, "D3", 系統名)
            write(ws, "D4", 試験圧力)
            write(ws, "M4", 試験範囲)
            write(ws, "D5", 試験媒体)
            write(ws, "M5", 放置時間)
            write(ws, "D6", 使用機器No)
            write(ws, "M6", 測定場所)
            write(ws, "D8", 開始日時.strftime("%Y/%m/%d %H:%M"))
            write(ws, "M8", 終了日時.strftime("%Y/%m/%d %H:%M"))

            write(ws, "A10", f"{P1:.4f}")
            write(ws, "C10", f"{T1:.1f}")
            write(ws, "E10", f"{P2p:.4f}")
            write(ws, "G10", f"{T2:.1f}")
            write(ws, "J10", f"{P2_corr:.3f}MPa")
            write(ws, "M10", f"{ΔP:.3f}MPa")
            write(ws, "O10", f"±{判定範囲:.3f}MPa")
            write(ws, "M11", 合否)
            write(ws, "E11", 試験実施者)

            # --- ダウンロード処理 ---
            output = BytesIO()
            wb.save(output)
            excel_data = output.getvalue()
            filename = f"気密検査報告書_{datetime.now().strftime('%Y%m%d')}.xlsx"
            b64 = base64.b64encode(excel_data).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Excelをダウンロード</a>'
            st.markdown(href, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"⚠ エラーが発生しました: {e}")
