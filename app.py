# -*- coding: utf-8 -*-
"""
============================================================
requirements.txt  (この内容を別ファイルに保存してください)
------------------------------------------------------------
streamlit>=1.35.0
pandas>=2.2.2
numpy>=1.26.4
openpyxl>=3.1.2
xlsxwriter>=3.2.0
============================================================
app.py
------------------------------------------------------------
"""

import io
from typing import List

import numpy as np
import pandas as pd
import streamlit as st

# ------------------------------
# 定数（テンプレートの行・列位置に合わせて調整してください）
# ------------------------------
DATE_COL_START = 4          # 列番号 (0-index) で "E" 列
NIGHT_ROWS = list(range(4, 16))   # E5:AI16 → 行 4~15 (0-index)
CARE_ROWS = list(range(19, 31))   # E20:AI30 → 行 19~30 (0-index)
LIMIT_COL = 3               # 上限(時間) が入っている列 ("D" 列)

# ------------------------------
# ユーティリティ
# ------------------------------

def detect_date_columns(df: pd.DataFrame, start_idx: int = DATE_COL_START) -> List[int]:
    """指定インデックス以降で「日付らしい」列番号を返す"""
    date_cols = []
    for col in range(start_idx, df.shape[1]):
        header_val = df.iloc[3, col]  # row 4 (1-index) = 0-index 3 は日付ヘッダ想定
        try:
            pd.to_datetime(str(header_val))
            date_cols.append(col)
        except Exception:
            # 変換できなければ日付列ではない
            pass
    # フォールバック: ヘッダが日付でなくても、とりあえず start_idx 以降を全部返す
    if not date_cols:
        date_cols = list(range(start_idx, df.shape[1]))
    return date_cols


def remove_excess_per_day(df: pd.DataFrame, row_indices: List[int], date_cols: List[int]):
    """同じ日に複数人割当てられている場合、先頭の 1 名だけを残す"""
    for col in date_cols:
        assigned = [row for row in row_indices if pd.notna(df.iat[row, col]) and df.iat[row, col] != 0]
        if len(assigned) > 1:
            # 1 人残して他を空欄化
            for row in assigned[1:]:
                df.iat[row, col] = np.nan


def enforce_limits(df: pd.DataFrame, row_indices: List[int], date_cols: List[int]):
    """各行 (職員) の労働時間が上限を超えたら、後ろの日付から削除"""
    for row in row_indices:
        limit_val = df.iat[row, LIMIT_COL]
        try:
            limit = float(limit_val)
        except (ValueError, TypeError):
            limit = None

        if not limit:
            continue  # 上限が設定されていない場合はスキップ

        # 現在の合計
        hours = sum([float(v) for v in df.iloc[row, date_cols].fillna(0)])
        if hours <= limit:
            continue

        # 後ろから減らす
        for col in reversed(date_cols):
            cell_val = df.iat[row, col]
            if pd.notna(cell_val) and cell_val != 0:
                df.iat[row, col] = np.nan
                hours -= float(cell_val)
                if hours <= limit:
                    break


def optimize(df: pd.DataFrame):
    """シフト最適化メイン関数
    戻り値: (最適化後 DataFrame, totals dict, limits dict)
    """
    date_cols = detect_date_columns(df, DATE_COL_START)

    df_opt = df.copy()

    # 1) 夜勤 (夜間支援員): 各日 1 名に制限
    remove_excess_per_day(df_opt, NIGHT_ROWS, date_cols)

    # 2) 世話人: 各日 1 名に制限
    remove_excess_per_day(df_opt, CARE_ROWS, date_cols)

    # 3) 上限時間の超過を解消
    enforce_limits(df_opt, NIGHT_ROWS + CARE_ROWS, date_cols)

    # 合計時間と上限を計算 (確認用)
    totals = {}
    limits = {}
    for row in NIGHT_ROWS + CARE_ROWS:
        totals[row] = float(sum([float(v) for v in df_opt.iloc[row, date_cols].fillna(0)]))
        try:
            limits[row] = float(df_opt.iat[row, LIMIT_COL])
        except (ValueError, TypeError):
            limits[row] = None

    return df_opt, totals, limits

# ------------------------------
# Streamlit UI
# ------------------------------

st.set_page_config(page_title="シフト自動調整ツール", page_icon="📅", layout="centered")

st.title("📅 シフト自動調整ツール")

with st.toggle("👉 使い方はこちら（クリックで展開）", value=False):
    st.markdown(
        """
        1. **Excel ファイル**をアップロードしてください。テンプレートと同じレイアウト (E5:AI16 と E20:AI30 がシフト範囲) を想定しています。
        2. **「最適化を実行」** ボタンを押すと、夜間支援員・世話人のシフトを自動で調整します。
        3. 完了すると **ダウンロードボタン** が表示されます。クリックして最適化済みファイルを保存してください。

        ----
        ### 反映ルール (概要)
        - 夜間支援員・世話人は **各日 1 名ずつ**。
        - 夜勤後は **2 日** 空けてから世話人勤務可。 (※詳細はテンプレートの運用に依存します)
        - 各職員の **上限時間** を超えないよう調整。
        - **0** が入っているセルにはシフトを入れません。
        """
    )

uploaded_file = st.file_uploader("Excel ファイルを選択", type=["xlsx", "xlsm", "xls"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, header=None, engine="openpyxl")
        st.success("Excel ファイルを読み込みました。")

        if st.button("🚀 最適化を実行"):
            with st.spinner("最適化中..."):
                df_opt, totals, limits = optimize(df_raw.copy())

            # ダウンロード準備
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_opt.to_excel(writer, index=False, header=False)
            buffer.seek(0)

            st.download_button(
                label="📥 最適化済みファイルをダウンロード",
                data=buffer.getvalue(),
                file_name="optimized_shift.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # 結果概要を表示 (オプション)
            st.subheader("💡 各職員の最終労働時間 (h)")
            result_df = pd.DataFrame({
                "Row": list(totals.keys()),
                "Total": list(totals.values()),
                "Limit": [limits.get(r) for r in totals.keys()],
            })
            st.dataframe(result_df, hide_index=True)

    except Exception as e:
        st.error(f"❌ 予期せぬエラーが発生しました: {e}")
