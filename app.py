# -*- coding: utf-8 -*-
"""
ルール完全
------------------------------------------------------------
- **夜勤 1 名 + 世話人 1 名 / 日（両ホーム合算）**
- **夜勤後 2 日インターバル**（どうしても割当て不能な場合のみ 1 日）
- **0 のセルは固定**
- **上限(時間) を厳守**（シート下部「上限(時間)」表から自動取得）
- **指定セル (E5‑AI16, E20‑AI30) 以外は一切変更しない**
- **出力は元のブックを保持** (openpyxl で該当セルだけ更新)
------------------------------------------------------------
"""

import io
from typing import List, Tuple, Dict

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

# -------------------- 主要定数 --------------------
HEADER_ROW = 3      # 日付が並ぶ行 (0-index)
START_ROW  = 4      # シフトが始まる最上行 (0-index)
NAME_COL   = 1      # 氏名列 (0-index)

# シフトを書き換えて良い範囲
EDIT_BLOCKS = [
    (4, 15),   # E5 〜 AI16  (0‑index: rows 4‑15)
    (19, 29),  # E20 〜 AI30 (0‑index: rows 19‑29)
]

SHIFT_NIGHT_HOURS = 12.5  # 夜勤 1 回
SHIFT_CARE_HOURS  = 6.0   # 世話人 1 回

# -------------------- ユーティリティ --------------------

def detect_date_columns(df: pd.DataFrame) -> List[int]:
    """1〜31 の整数が入っているヘッダー列を日付列とみなす"""
    date_cols: List[int] = []
    for c in df.columns:
        val = df.iat[HEADER_ROW, c]
        try:
            v = int(float(val))
            if 1 <= v <= 31:
                date_cols.append(c)
        except (ValueError, TypeError):
            continue
    if not date_cols:
        raise ValueError("ヘッダー行に 1〜31 の日付が見つかりません。行番号・列番号を確認してください。")
    return date_cols


def detect_row_indices(df: pd.DataFrame) -> Tuple[List[int], List[int]]:
    """1 列目のラベルで『夜間支援員』『世話人』を判定 (氏名が空でない行のみ)"""
    night_rows, care_rows = [], []
    for r in range(START_ROW, df.shape[0]):
        role = df.iat[r, 0]
        name = df.iat[r, NAME_COL]
        if not isinstance(role, str) or not isinstance(name, str) or not name.strip():
            continue
        role_flat = role.replace("\n", "")
        if "夜間" in role_flat and "支援員" in role_flat:
            night_rows.append(r)
        elif "世話人" in role_flat:
            care_rows.append(r)
    if not night_rows or not care_rows:
        raise ValueError("夜間支援員 / 世話人 の行が検出できませんでした。行ラベルを確認してください。")
    return night_rows, care_rows


def get_limits(df: pd.DataFrame) -> pd.Series:
    """下部にある『上限(時間)』テーブルを自動抽出"""
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            if str(df.iat[r, c]).startswith("上限"):
                name_col, limit_col = c - 1, c
                limits = {}
                rr = r + 1
                while rr < df.shape[0]:
                    name = df.iat[rr, name_col]
                    if not isinstance(name, str) or not name.strip():
                        break
                    try:
                        limit = float(df.iat[rr, limit_col])
                    except (ValueError, TypeError):
                        limit = np.inf
                    limits[name.strip()] = limit
                    rr += 1
                return pd.Series(limits)
    raise ValueError("『上限(時間)』テーブルが見つかりませんでした。シート最下部に配置してください。")


def in_edit_blocks(r: int) -> bool:
    """行 r が編集可能ブロックに含まれるか"""
    for start, end in EDIT_BLOCKS:
        if start <= r <= end:
            return True
    return False

# -------------------- 割当てアルゴリズム --------------------

def optimize(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    date_cols = detect_date_columns(df)
    night_rows, care_rows = detect_row_indices(df)

    # 行チェック: 編集ブロック外が混ざっていないか警告
    for r in night_rows + care_rows:
        if not in_edit_blocks(r):
            raise ValueError("夜間支援員/世話人 の行が EDIT_BLOCKS から外れています。定数 EDIT_BLOCKS を調整してください。")

    night_names = {r: df.iat[r, NAME_COL].strip() for r in night_rows}
    care_names  = {r: df.iat[r, NAME_COL].strip() for r in care_rows}
    all_names = set(night_names.values()) | set(care_names.values())

    limits = get_limits(df).reindex(all_names).fillna(np.inf)
    totals = pd.Series(0.0, index=all_names)

    # ---------- 既存シフトを削除 (0 は保持) ----------
    for r in night_rows + care_rows:
        for c in date_cols:
            if df.iat[r, c] != 0 and not pd.isna(df.iat[r, c]):
                df.iat[r, c] = np.nan

    # ---------- 割当て状態 ----------
    last_night_day: Dict[str, int] = {}

    # ---------- 各日ループ ----------
    for d_idx, c in enumerate(date_cols):
        # ===== 夜勤候補 =====
        night_cand = [
            (limits[night_names[r]] - totals[night_names[r]], night_names[r], r)
            for r in night_rows
            if pd.isna(df.iat[r, c])  # 空欄のみ
            and totals[night_names[r]] + SHIFT_NIGHT_HOURS <= limits[night_names[r]]
        ]
        if not night_cand:
            raise RuntimeError(f"{d_idx+1} 日目の夜勤を割り当てられません。上限・0 セルを確認してください。")
        night_cand.sort(key=lambda x: (x[0], x[1]))  # 残余が少ない人 → 氏名順
        _, night_name, night_row = night_cand[0]
        df.iat[night_row, c] = SHIFT_NIGHT_HOURS
        totals[night_name] += SHIFT_NIGHT_HOURS
        last_night_day[night_name] = d_idx

        # ===== 世話人候補 =====
        care_cand = [
            (limits[care_names[r]] - totals[care_names[r]], care_names[r], r)
            for r in care_rows
            if pd.isna(df.iat[r, c])
            and care_names[r] != night_name
            and (care_names[r] not in last_night_day or d_idx - last_night_day[care_names[r]] >= 3)
            and totals[care_names[r]] + SHIFT_CARE_HOURS <= limits[care_names[r]]
        ]
        if not care_cand:
            # どうしても空く場合はインターバル 1 日で緩和
            care_cand = [
                (limits[care_names[r]] - totals[care_names[r]], care_names[r], r)
                for r in care_rows
                if pd.isna(df.iat[r, c])
                and care_names[r] != night_name
                and (care_names[r] not in last_night_day or d_idx - last_night_day[care_names[r]] >= 2)
                and totals[care_names[r]] + SHIFT_CARE_HOURS <= limits[care_names[r]]
            ]
        if not care_cand:
            raise RuntimeError(f"{d_idx+1} 日目の世話人を割り当てられません。上限・0 セルを確認してください。")
        care_cand.sort(key=lambda x: (x[0], x[1]))
        _, care_name, care_row = care_cand[0]
        df.iat[care_row, c] = SHIFT_CARE_HOURS
        totals[care_name] += SHIFT_CARE_HOURS

    # ------ 完了 ------
    return df, totals.sort_index(), limits.sort_index()

# -------------------- セル更新 (openpyxl) --------------------

def write_back(original_stream: io.BytesIO, df_opt: pd.DataFrame) -> bytes:
    """元ブックに対して変更セルだけ上書きし、bytes を返す"""
    original_stream.seek(0)
    wb: Workbook = load_workbook(original_stream, data_only=False)
    ws = wb.active

    for r in range(df_opt.shape[0]):
        if not in_edit_blocks(r):
            continue  # 編集許可外
        for c in range(df_opt.shape[1]):
            # 列が日付列かどうかはヘッダー行で判定
            header_val = df_opt.iat[HEADER_ROW, c]
            try:
                _ = int(float(header_val))
            except (ValueError, TypeError):
                continue  # 日付列でない
            new_val = df_opt.iat[r, c]
            if pd.isna(new_val):  # Nothing to write
                new_val = None
            ws.cell(row=r + 1, column=c + 1, value=new_val)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# -------------------- Streamlit UI --------------------

st.set_page_config(page_title="シフト自動最適化", layout="wide")
st.title("📅 シフト自動最適化ツール (ルール完全版・セル最小更新)")

with st.expander("👉 使い方はこちら", expanded=False):
    st.markdown(
        """
        **手順**
        1. 左サイドバーで Excel ファイル (.xlsx) をアップロード。
        2. **「🚀 最適化を実行」** をクリック。
        3. 右側にプレビューが表示されます。
        4. **「📥 ダウンロード」** で最適化済み Excel を取得。

        **割当てロジック**
        - 日毎に *夜勤 1 名* + *世話人 1 名*（両ホーム合算）。
        - 夜勤後 2 日は世話人不可（やむを得ない場合は 1 日）。
        - 0 セルは固定で不可。
        - 下部『上限(時間)』表の値を厳守。
        - 指定セル (E5‑AI16, E20‑AI30) 以外は一切変更しません。元の数式も保持します。
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
            st.success("✅ 最適化が完了しました")

            st.subheader("最適化後のシフト表 (プレビュー)")
            st.dataframe(df_opt, use_container_width=True)

            st.subheader("勤務時間の合計 / 上限")
            st.dataframe(pd.DataFrame({"合計時間": totals, "上限時間": limits}))

            # ------- Excel 出力 (元ブックに書き戻し) -------
            optimized_bytes = write_back(uploaded, df_opt)
            st.download_button(
                "📥 最適化シフトをダウンロード (Excel)",
                data=optimized_bytes,
                file_name="optimized_shift.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"ファイルの読み込みまたは最適化中にエラーが発生しました: {e}")
else:
    st.info("左のサイドバーからテンプレート Excel をアップロードしてください。")
