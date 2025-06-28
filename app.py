import io
from typing import List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# -------------------- 定数 --------------------
# テンプレート構造が変わった場合はここを調整
# Excel は 1 行目 = index 0 扱い（pandas のヘッダ無し読込を想定）
NIGHT_ROWS = list(range(4, 16))   # E5:AI16 → 0‑index 行 4‑15
CARE_ROWS  = list(range(19, 31))  # E20:AI30 → 0‑index 行 19‑30
DATE_HEADER_ROW = 3               # E4 行 (0‑index 3)

# -------------------- 関数群 --------------------

def detect_date_columns(df: pd.DataFrame) -> List[str]:
    """ヘッダーから日付列を推定し、連続する範囲（列名リスト）を返す"""
    date_cols = []
    for col in df.columns:
        header = str(df.at[DATE_HEADER_ROW, col])
        try:
            pd.to_datetime(header, errors="raise")
            date_cols.append(col)
        except (ValueError, TypeError):
            pass
    if not date_cols:
        raise ValueError("日付列を検出できませんでした。ヘッダー行と列番号を確認してください。")
    # 最初と最後の連続ブロックだけ抽出
    first_idx = df.columns.get_loc(date_cols[0])
    last_idx  = df.columns.get_loc(date_cols[-1]) + 1
    return list(df.columns[first_idx:last_idx])


def optimize(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    """シフト最適化ロジック（簡易版）
    - 指定行ブロックを全消去して空欄化（0 は残す）
    - 本格的な最適化アルゴリズムは必要に応じて実装してください
    """
    date_cols = detect_date_columns(df)

    # -------------------- 勤務時間上限の取得 --------------------
    # Excel テンプレート左端付近に「上限」列があると仮定
    # 文字列を float 化しようとして失敗するケースに備え、数値以外は NaN にする
    try:
        limits_raw = (
            df.iloc[:, :4]  # 氏名列を含むブロック（列 0‑3 想定）
            .set_index(df.columns[0])
            .iloc[:, -1]   # そのブロックの一番右列に "上限" がある想定
        )
        limits = pd.to_numeric(limits_raw, errors="coerce")  # ← 文字列は NaN
        limits = limits.dropna()
    except Exception:
        limits = pd.Series(dtype=float)

    # totals は後段のアルゴリズム実装用のダミー（現状は全 0）
    totals = pd.Series(0.0, index=limits.index, dtype=float)

    # -------------------- 指定ブロックのクリア --------------------
    def clear_block(rows: List[int]):
        for r in rows:
            for c in date_cols:
                if df.at[r, c] != 0:  # 0 は "固定で不可" の意味なので維持
                    df.at[r, c] = np.nan  # 空セル化

    clear_block(NIGHT_ROWS)
    clear_block(CARE_ROWS)

    # -------------------- TODO: 割当アルゴリズム --------------------
    # 夜勤 1 名／世話人 1 名のロジックをここに実装してください

    return df, totals, limits

# -------------------- Streamlit UI --------------------

st.set_page_config(page_title="シフト自動最適化", layout="wide")
st.title("📅 シフト自動最適化ツール")

with st.expander("👉 使い方はこちら（クリックで展開）", expanded=False):
    st.markdown(
        """
        **▼ 手順**
        1. 左サイドバーで **テンプレート形式** の Excel ファイル (.xlsx) を選択してアップロード。
        2. **「🚀 最適化を実行」** ボタンを押す。
        3. 右側に最適化後のシフトプレビューが表示される。
        4. **「📥 ダウンロード」** ボタンで Excel を取得。

        *行・列の位置がテンプレートと異なる場合は、ソースコード冒頭の定数を調整してください。*
        """
    )

st.sidebar.header("📂 入力ファイル")
uploaded = st.sidebar.file_uploader("Excel ファイル (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        df_input = pd.read_excel(uploaded, header=None, engine="openpyxl")
        st.subheader("アップロードされたシフト表")
        st.dataframe(df_input, use_container_width=True)

        if st.sidebar.button("🚀 最適化を実行"):
            df_opt, totals, limits = optimize(df_input.copy())
            st.success("最適化が完了しました 🎉")

            st.subheader("最適化後のシフト表")
            st.dataframe(df_opt, use_container_width=True)

            if not limits.empty:
                st.subheader("勤務時間の合計と上限")
                st.dataframe(
                    pd.DataFrame({"合計時間": totals, "上限時間": limits})
                )

            # Excel 出力
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_opt.to_excel(writer, index=False, header=False)
            st.download_button(
                label="📥 最適化シフトをダウンロード",
                data=buffer.getvalue(),
                file_name="optimized_shift.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"ファイルの読み込みまたは最適化中にエラーが発生しました: {e}")
else:
    st.info("左のサイドバーからテンプレート形式の Excel ファイルをアップロードしてください。")
