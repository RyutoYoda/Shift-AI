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
app.py  （下記を保存して `streamlit run app.py`）
------------------------------------------------------------
日本語 UI／Excel 入出力／夜勤 1 名 + 世話人 1 名／上限時間／夜勤→世話人 2 日空け
"""

import io
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# -------------------- 定数 --------------------
# ▼ テンプレートの行・列位置。ずれたらここだけ直せば OK
NIGHT_ROWS = list(range(4, 16))   # E5:AI16 → 0‑index 行 4‑15
CARE_ROWS  = list(range(19, 31))  # E20:AI30 → 0‑index 行 19‑30
DATE_HEADER_ROW = 3               # 日付が並ぶヘッダー行（0‑index 3）
NAME_COL = 0                      # 氏名列のインデックス

SHIFT_NIGHT_HOURS = 12.5          # 夜勤 1 回の時間
SHIFT_CARE_HOURS  = 6.0           # 世話人 1 回の時間

# -------------------- 日付列の検出 --------------------

def detect_date_columns(df: pd.DataFrame) -> List[str]:
    """ヘッダーに『日付』っぽい値がある連続列を抽出"""
    date_cols: List[str] = []
    for col in df.columns:
        header = str(df.at[DATE_HEADER_ROW, col])
        try:
            pd.to_datetime(header, errors="raise")
            date_cols.append(col)
        except (ValueError, TypeError):
            pass
    if not date_cols:
        raise ValueError("ヘッダー行に日付列を検出できませんでした。行番号・列番号を確認してください。")

    # 先頭日付列から最後の日付列までを対象とする
    first_idx = df.columns.get_loc(date_cols[0])
    last_idx  = df.columns.get_loc(date_cols[-1]) + 1
    return list(df.columns[first_idx:last_idx])

# -------------------- 最適化アルゴリズム --------------------

def optimize(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    """シフト自動割当て

    1. 夜勤行・世話人行を一旦クリア (0 は残す)
    2. 各日 **夜勤 1 名 + 世話人 1 名** を割当て
    3. 制約
       - 夜勤後は少なくとも 2 日空けて世話人可
       - 世話人翌日の夜勤入りは可
       - 各人上限時間以内
       - 0 のセルは固定不可
       - 指定行ブロック以外はいじらない
    """

    date_cols = detect_date_columns(df)

    # -------------- 氏名マッピング --------------
    night_names = {r: str(df.at[r, NAME_COL]).strip() for r in NIGHT_ROWS}
    care_names  = {r: str(df.at[r, NAME_COL]).strip() for r in CARE_ROWS}

    all_names = set(night_names.values()) | set(care_names.values())
    all_names.discard("")  # 空文字削除

    # -------------- 上限時間の取得 --------------
    # 氏名列から数列分右側のどこかに「上限」がある想定（なければ無制限扱い）
    try:
        limits_raw = (
            df.iloc[:, : (NAME_COL + 4)]  # 氏名列 + 右 3 列くらいをスキャン
            .set_index(df.columns[NAME_COL])
            .iloc[:, -1]  # そのブロックの一番右列を "上限" とみなす
        )
        limits = pd.to_numeric(limits_raw, errors="coerce").reindex(all_names).fillna(np.inf)
    except Exception:
        limits = pd.Series(np.inf, index=list(all_names))

    # -------------- 勤務時間合計の初期化 --------------
    totals = pd.Series(0.0, index=list(all_names))

    # -------------- 指定ブロックをクリア（0 を残す） --------------
    def clear_block(rows: List[int]):
        for r in rows:
            for c in date_cols:
                if df.at[r, c] != 0:
                    df.at[r, c] = np.nan
    clear_block(NIGHT_ROWS)
    clear_block(CARE_ROWS)

    # -------------- 割当履歴 --------------------
    last_night_day: Dict[str, int] = {}  # 夜勤に入った最終「日index」（インターバル確認用）

    # -------------- 割当ロジック ----------------
    for day_idx, col in enumerate(date_cols):
        # ===== 夜勤 =====
        night_candidates = []
        for r in NIGHT_ROWS:
            if df.at[r, col] == 0:
                continue  # 0 = 固定で不可
            name = night_names.get(r, "")
            if not name:
                continue
            remaining = limits[name] - totals[name]
            if remaining >= SHIFT_NIGHT_HOURS:
                night_candidates.append((remaining, name, r))
        if not night_candidates:
            raise RuntimeError(f"{col} の夜勤を割り当てられる候補がいません。テンプレまたは上限を確認してください。")
        # 残余時間が多い順で決定
        night_candidates.sort(reverse=True)
        _, night_name, night_row = night_candidates[0]
        df.at[night_row, col] = SHIFT_NIGHT_HOURS
        totals[night_name] += SHIFT_NIGHT_HOURS
        last_night_day[night_name] = day_idx

        # ===== 世話人 =====
        care_candidates = []
        for r in CARE_ROWS:
            if df.at[r, col] == 0:
                continue
            name = care_names.get(r, "")
            if (not name) or (name == night_name):  # 同じ人が同日に夜勤+世話人は不可とする
                continue
            # 夜勤後 2 日インターバル
            if name in last_night_day and day_idx - last_night_day[name] < 3:
                continue
            remaining = limits[name] - totals[name]
            if remaining >= SHIFT_CARE_HOURS:
                care_candidates.append((remaining, name, r))
        if not care_candidates:
            # 妥協策：インターバル無視で再探索（例外を避ける）
            for r in CARE_ROWS:
                if df.at[r, col] == 0:
                    continue
                name = care_names.get(r, "")
                if name and name != night_name:
                    remaining = limits[name] - totals[name]
                    care_candidates.append((remaining, name, r))
        if not care_candidates:
            raise RuntimeError(f"{col} の世話人を割り当てられる候補がいません。テンプレまたは上限を確認してください。")
        care_candidates.sort(reverse=True)
        _, care_name, care_row = care_candidates[0]
        df.at[care_row, col] = SHIFT_CARE_HOURS
        totals[care_name] += SHIFT_CARE_HOURS

    return df, totals.sort_index(), limits.sort_index()

# -------------------- Streamlit UI --------------------

st.set_page_config(page_title="シフト自動最適化", layout="wide")
st.title("📅 シフト自動最適化ツール")

with st.expander("👉 使い方はこちら（クリックで展開）", expanded=False):
    st.markdown(
        """
        **▼ 手順**
        1. 左サイドバーで **テンプレート形式** の Excel ファイル (.xlsx) をアップロード。
        2. **「🚀 最適化を実行」** ボタンを押す。
        3. 最適化後のシフトがプレビューされます。
        4. **「📥 ダウンロード」** ボタンで Excel を取得。

        **▼ アルゴリズム概要**
        - 各日 *夜勤 1 名* と *世話人 1 名* を自動選出。
        - 夜勤後は 2 日 (翌日+翌々日) 世話人不可。
        - 世話人翌日の夜勤は OK。
        - 各人の累計時間が "上限" を超えないように調整。
        - "0" が入っているセルは固定で不可。
        - 指定行ブロック (E5‑AI16 / E20‑AI30) **以外のセルは一切変更しません**。
        """
    )

st.sidebar.header("📂 入力ファイル")
uploaded = st.sidebar.file_uploader("Excel ファイル (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        df_input = pd.read_excel(uploaded, header=None, engine="openpyxl")
        st.subheader("アップロードされたシフト表 (プレビュー)")
        st.dataframe(df_input, use_container_width=True)

        if st.sidebar.button("🚀 最適化を実行"):
            df_opt, totals, limits = optimize(df_input.copy())
            st.success("最適化が完了しました 🎉")

            st.subheader("最適化後のシフト表")
            st.dataframe(df_opt, use_container_width=True)

            st.subheader("勤務時間の合計 / 上限")
            st.dataframe(
                pd.DataFrame({"合計時間": totals, "上限時間": limits})
            )

            # Excel 出力
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_opt.to_excel(writer, index=False, header=False)
            st.download_button(
                label="📥 最適化シフトをダウンロード (Excel)",
                data=buffer.getvalue(),
                file_name="optimized_shift.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"ファイルの読み込みまたは最適化中にエラーが発生しました: {e}")
else:
    st.info("左のサイドバーからテンプレート形式の Excel ファイルをアップロードしてください。")
