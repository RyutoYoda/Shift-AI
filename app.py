# -*- coding: utf-8 -*-
"""
Streamlit アプリ：シフト調整ツール
入力 : 既存シフトが入った Excel
出力 : ルール準拠に自動調整した Excel
---------------------------------------
※ シート構成や行列の配置が月によって変わる場合は、
   SECTION ▶ 「設定」 の定数を編集してください。
"""

import pandas as pd
import numpy as np
from io import BytesIO
import streamlit as st

# ╭──────────────────────────────────────────────╮
# │ SECTION ▶ 設定                              │
# ╰──────────────────────────────────────────────╯
SHEET_NAME = 0               # 読み込むシート（0 なら 1 枚目）
NIGHT_ROWS = list(range(4, 16))   # 夜間支援員：Excel 行番号-1 (E5:AI16)
CARE_ROWS  = list(range(19, 30))  # 世話人　　：Excel 行番号-1 (E20:AI30)
DATE_ROW   = 3               # 日付が入っている行番号-1 (E4:AI4)
DATE_COL_START = "E"         # 開始列
DATE_COL_END   = "AI"        # 終了列
LIMIT_TABLE_CELL = "B34"     # 上限表左上セル（名前下に時間が縦配置と想定）

# ──────────────────────────────────────────────
def col_range(df: pd.DataFrame, start_col: str, end_col: str):
    """列番号範囲をリストで返す（Excel 列記法 → 位置）"""
    cols = df.columns.tolist()
    start_idx = cols.index(start_col)
    end_idx   = cols.index(end_col)
    return cols[start_idx : end_idx + 1]

def extract_limits(df: pd.DataFrame, anchor_cell: str) -> dict:
    """'上限(時間)' 表から {氏名: 時間} を取得（縦 2 列構成想定）"""
    anchor = df.columns.get_loc(anchor_cell[0])  # 列番号
    row    = int(anchor_cell[1:]) - 1           # 行番号-1
    names  = df.iloc[row+1 :, anchor].dropna()
    hours  = df.iloc[row+1 :, anchor+1].dropna()
    return dict(zip(names.astype(str), hours.astype(float)))

def hours_by_staff(df: pd.DataFrame, rows: list, date_cols: list):
    """各 staff 行の合計時間 Series を返す"""
    sub = df.loc[rows, date_cols].replace("", 0).fillna(0)
    return sub.sum(axis=1)

def enforce_one_per_day(df: pd.DataFrame, rows: list, date_cols: list,
                        totals: pd.Series, limits: dict):
    """各日につき 1 名だけ残し、それ以外クリア"""
    for col in date_cols:
        non_zeros = [r for r in rows if df.at[r, col] not in [0, "", np.nan]]
        # 0 はロック扱い
        non_zeros = [r for r in non_zeros if df.at[r, col] != 0]
        if len(non_zeros) <= 1:
            continue
        # 候補のうち「(上限 - 現在) が最大」の人を残す
        def slack(r):
            name = df.at[r, df.columns[1]]
            return limits.get(name, 1e9) - totals.get(r, 0)
        keep = max(non_zeros, key=slack)
        for r in non_zeros:
            if r != keep:
                df.at[r, col] = ""
                totals[r] -= df.at[r, col] if isinstance(df.at[r, col], (int, float)) else 0

def remove_excess(df: pd.DataFrame, rows: list, date_cols: list,
                  totals: pd.Series, limits: dict):
    """上限を超えたスタッフから勤務を削除（世話人 → 夜勤 の順）"""
    for r in rows:
        name = df.at[r, df.columns[1]]
        limit = limits.get(name, np.inf)
        if totals[r] <= limit:
            continue
        # 1) CARE_ROWS → 2) NIGHT_ROWS の順で後方から削除
        for col in reversed(date_cols):
            if totals[r] <= limit:
                break
            if df.at[r, col] not in [0, "", np.nan]:
                df.at[r, col] = ""
                totals[r] -= df.at[r, col] if isinstance(df.at[r, col], (int, float)) else 0

def block_after_night(df: pd.DataFrame, night_rows: list, care_rows: list,
                      date_cols: list):
    """夜勤翌日・翌々日の自分の世話人シフトを削除"""
    col_idx = {c: i for i, c in enumerate(date_cols)}
    for r in night_rows:
        for col in date_cols:
            if df.at[r, col] not in [0, "", np.nan]:
                name = df.at[r, df.columns[1]]
                i = col_idx[col]
                for offset in [1, 2]:
                    if i + offset >= len(date_cols):
                        continue
                    tgt_col = date_cols[i + offset]
                    # 該当 staff の care_rows を探す
                    for care_r in care_rows:
                        if df.at[care_r, df.columns[1]] == name:
                            df.at[care_r, tgt_col] = ""

def optimize(df: pd.DataFrame):
    """全工程をまとめて実行し DataFrame を返す"""
    date_cols = col_range(df, DATE_COL_START, DATE_COL_END)
    limits = extract_limits(df, LIMIT_TABLE_CELL)
    total_night = hours_by_staff(df, NIGHT_ROWS, date_cols)
    total_care  = hours_by_staff(df, CARE_ROWS, date_cols)
    totals = pd.concat([total_night, total_care])
    # ① 各日 1 名ずつ
    enforce_one_per_day(df, NIGHT_ROWS, date_cols, totals, limits)
    enforce_one_per_day(df, CARE_ROWS , date_cols, totals, limits)
    # ② 夜勤後インターバル
    block_after_night(df, NIGHT_ROWS, CARE_ROWS, date_cols)
    # ③ 上限超過の削減
    totals = hours_by_staff(df, NIGHT_ROWS+CARE_ROWS, date_cols)
    remove_excess(df, CARE_ROWS+NIGHT_ROWS, date_cols, totals, limits)
    return df, totals, limits

# ╭──────────────────────────────────────────────╮
# │ SECTION ▶ Streamlit UI                      │
# ╰──────────────────────────────────────────────╯
st.set_page_config(page_title="シフト調整ツール", layout="wide")
st.title("🗓️ シフト自動調整ツール")

show_help = st.toggle("使い方を表示する", value=False)
if show_help:
    st.markdown("""
**手順**  
1. 「ファイルを選択」からシフト案（Excel）をアップロード  
2. 自動調整後のプレビューが表に表示されます  
3. 問題なければ「完成版をダウンロード」で保存  

**前提**  
- 夜間支援員 (E5:AI16)・世話人 (E20:AI30) が入力対象  
- 0 が入ったセルには変更を加えません  
- 夜勤は 1 日 1 名、世話人も 1 日 1 名  
- 夜勤後は 2 日空けて同一人物の世話人シフト不可  
- 「上限(時間)」表の値を厳守（表はシート左下付近を想定）  
""")


uploaded = st.file_uploader("Excel ファイルを選択してください", type=["xlsx"])
if uploaded:
    df = pd.read_excel(uploaded, header=None, sheet_name=SHEET_NAME)
    df_opt, totals, limits = optimize(df.copy())
    st.success("✅ 自動調整が完了しました")
    # 仕上がりプレビュー
    st.subheader("プレビュー")
    st.dataframe(df_opt.replace(np.nan, ""), use_container_width=True)
    # 各人の時間サマリ
    summ = pd.DataFrame({
        "氏名": [df_opt.at[r, df_opt.columns[1]] for r in (NIGHT_ROWS+CARE_ROWS)],
        "合計時間": totals.values,
        "上限": [limits.get(df_opt.at[r, df_opt.columns[1]], np.nan)
                 for r in (NIGHT_ROWS+CARE_ROWS)],
    }).drop_duplicates("氏名").set_index("氏名")
    st.subheader("各人の労働時間サマリ")
    st.dataframe(summ)
    # Excel 書き出し
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_opt.to_excel(writer, index=False, header=False)
    st.download_button(
        label="📥 完成版をダウンロード",
        data=output.getvalue(),
        file_name="シフト完成版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("↑ まずは Excel ファイルをアップロードしてください")
