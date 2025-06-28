# -*- coding: utf-8 -*-
"""
- **グループホーム1（GH1）/ グループホーム2（GH2）** それぞれに *夜勤 1 名 + 世話人 1 名 / 日* を必ず充当
- 夜勤 → 世話人インターバル = 2 日（足りない場合 1→0 に自動緩和）
- 同一人物が同一ロールで **連続日** に入るのを禁止
- 同一人物が同日に複数シフトに入るのを禁止
- 0 セル厳守 / 上限時間厳守（限界まで使用）
- 一旦 E5:AI16 (GH1) / E20:AI30 (GH2) をクリアして再配置
- その他セル＆列 C の集計式は保持
- 出力は `.xlsx`
------------------------------------------------------------
"""

import io
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

# -------------------- 定数 --------------------
HEADER_ROW = 3      # 日付が並ぶ行 (0-index)
START_ROW  = 4      # シフト行開始 (0-index)
NAME_COL   = 1      # 氏名列 (0-index)

# グループホームごとの編集ブロック (0‑index)
HOME_BLOCKS = {
    1: (4, 15),   # GH1: E5‑AI16
    2: (19, 29),  # GH2: E20‑AI30
}

SHIFT_NIGHT_HOURS = 12.5  # 夜勤 1 回
SHIFT_CARE_HOURS  = 6.0   # 世話人 1 回

# -------------------- ユーティリティ --------------------

def detect_date_columns(df: pd.DataFrame) -> List[int]:
    date_cols = []
    for c in df.columns:
        val = df.iat[HEADER_ROW, c]
        try:
            v = int(float(val))
            if 1 <= v <= 31:
                date_cols.append(c)
        except (ValueError, TypeError):
            continue
    if not date_cols:
        raise ValueError("ヘッダー行に日付が見つかりません。行番号・列番号を確認してください。")
    return date_cols


def detect_rows_by_home(df: pd.DataFrame) -> Tuple[Dict[int, List[int]], Dict[int, List[int]]]:
    night_rows: Dict[int, List[int]] = {1: [], 2: []}
    care_rows: Dict[int, List[int]] = {1: [], 2: []}
    for home, (start, end) in HOME_BLOCKS.items():
        for r in range(start, end + 1):
            role = df.iat[r, 0]
            name = df.iat[r, NAME_COL]
            if not isinstance(role, str) or not isinstance(name, str) or not name.strip():
                continue
            role_flat = role.replace("\n", "")
            if "夜間" in role_flat and "支援員" in role_flat:
                night_rows[home].append(r)
            elif "世話人" in role_flat:
                care_rows[home].append(r)
    if any(len(v) == 0 for v in night_rows.values()) or any(len(v) == 0 for v in care_rows.values()):
        raise ValueError("夜間支援員 / 世話人 の行が検出できません。行ラベル・ブロックを確認してください。")
    return night_rows, care_rows


def get_limits(df: pd.DataFrame) -> pd.Series:
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
                    limit_val = pd.to_numeric(df.iat[rr, limit_col], errors="coerce")
                    limits[name.strip()] = float(limit_val) if not np.isnan(limit_val) else np.inf
                    rr += 1
                return pd.Series(limits)
    raise ValueError("『上限(時間)』テーブルが見つかりません。")


# -------------------- 割当てロジック --------------------

def clear_blocks(df: pd.DataFrame):
    """編集ブロックをクリア。ただし 0 は残す。"""
    for start, end in HOME_BLOCKS.values():
        for r in range(start, end + 1):
            for c in detect_date_columns(df):
                if df.iat[r, c] != 0 and not pd.isna(df.iat[r, c]):
                    df.iat[r, c] = np.nan


def choose_candidate(cands):
    """(残余時間, 連続回避, name, row) のタプルを残余時間が少ない順→連続回避→名前順でソート"""
    return sorted(cands, key=lambda x: (x[0], x[1], x[2]))[0]


def optimize(df: pd.DataFrame):
    date_cols = detect_date_columns(df)
    night_rows, care_rows = detect_rows_by_home(df)

    limits = get_limits(df)
    names = limits.index.tolist()

    # クリア
    clear_blocks(df)

    # 各人物の累積時間
    totals = pd.Series(0.0, index=names)

    # 各人物が夜勤に入った直近の日 (全ホーム共有)
    last_night_day: Dict[str, int] = {}
    # 各人物が各ロールを担当した直近の日 (consecutive 回避)
    last_role_day: Dict[Tuple[str, str], int] = {}

    # -------------- メインループ --------------
    for d_idx, c in enumerate(date_cols):
        assigned_today: set[str] = set()
        for home in (1, 2):
            # ---------- 夜勤 ----------
            night_cand = []
            for r in night_rows[home]:
                name = df.iat[r, NAME_COL].strip()
                if not pd.isna(df.iat[r, c]):
                    continue  # 既に何か入っている（0 のみ残っているはず）
                if name in assigned_today:
                    continue  # 同日複数シフト禁止
                if totals[name] + SHIFT_NIGHT_HOURS > limits.get(name, np.inf):
                    continue  # 上限超える
                # 連続回避
                consec_penalty = 0
                if last_role_day.get((name, f"night_{home}"), -99) == d_idx - 1:
                    consec_penalty = 1
                night_cand.append((limits[name] - totals[name], consec_penalty, name, r))
            if not night_cand:
                raise RuntimeError(f"{d_idx+1} 日目 GH{home} の夜勤が充当できません。")
            _, _, night_name, night_row = choose_candidate(night_cand)
            df.iat[night_row, c] = SHIFT_NIGHT_HOURS
            totals[night_name] += SHIFT_NIGHT_HOURS
            last_night_day[night_name] = d_idx
            last_role_day[(night_name, f"night_{home}")] = d_idx
            assigned_today.add(night_name)

            # ---------- 世話人 ----------
            care_cand = []
            for r in care_rows[home]:
                name = df.iat[r, NAME_COL].strip()
                if not pd.isna(df.iat[r, c]):
                    continue
                if name in assigned_today:
                    continue
                if totals[name] + SHIFT_CARE_HOURS > limits.get(name, np.inf):
                    continue
                # 夜勤→世話人インターバル
                interval_ok = False
                for interval in (2, 1, 0):
                    if name not in last_night_day or d_idx - last_night_day[name] >= interval + 1:
                        interval_ok = True
                        break
                if not interval_ok:
                    continue
                # 連続回避
                consec_penalty = 0
                if last_role_day.get((name, f"care_{home}"), -99) == d_idx - 1:
                    consec_penalty = 1
                care_cand.append((limits[name] - totals[name], consec_penalty, name, r))
            if not care_cand:
                raise RuntimeError(f"{d_idx+1} 日目 GH{home} の世話人が充当できません。")
            _, _, care_name, care_row = choose_candidate(care_cand)
            df.iat[care_row, c] = SHIFT_CARE_HOURS
            totals[care_name] += SHIFT_CARE_HOURS
            last_role_day[(care_name, f"care_{home}")] = d_idx
            assigned_today.add(care_name)

    # ---------- 完成チェック ----------
    for home in (1, 2):
        for c in date_cols:
            if all(pd.isna(df.iat[r, c]) or df.iat[r, c] == 0 for r in night_rows[home]):
                raise RuntimeError(f"日 {df.iat[HEADER_ROW, c]} GH{home} の夜勤が空欄です")
            if all(pd.isna(df.iat[r, c]) or df.iat[r, c] == 0 for r in care_rows[home]):
                raise RuntimeError(f"日 {df.iat[HEADER_ROW, c]} GH{home} の世話人が空欄です")
    return df, totals.sort_index(), limits.sort_index()

# -------------------- 書き戻し --------------------

def write_back(original_stream: io.BytesIO, df_opt: pd.DataFrame) -> bytes:
    original_stream.seek(0)
    wb: Workbook = load_workbook(original_stream, data_only=False)
    ws = wb.active

    date_cols = detect_date_columns(df_opt)
    for home, (start, end) in HOME_BLOCKS.items():
        for r in range(start, end + 1):
            for c in date_cols:
                new_val = df_opt.iat[r, c]
                if pd.isna(new_val):
                    new_val = None
                ws.cell(row=r + 1, column=c + 1, value=new_val)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# -------------------- Streamlit UI --------------------

st.set_page_config(page_title="シフト自動最適化 (GH1・GH2 完全版)", layout="wide")
st.title("🏠 シフト自動最適化ツール ‑ GH1 & GH2")

with st.expander("👉 使い方はこちら", expanded=False):
    st.markdown(
        """
        **手順**
        1. 左サイドバーで Excel ファイル (.xlsx) をアップロード。
        2. **「🚀 最適化を実行」** をクリック。
        3. 右側にプレビューが表示されます（**GH1/GH2 共に毎日 2 枠充足**）。
        4. **「📥 ダウンロード」** で最適化済み Excel を取得。

        **割当てロジック概要**
        - 各日 *各ホーム* に **夜勤 1 名 + 世話人 1 名** を必ず充当。
        - 夜勤→世話人インターバルは 2 日を原則とし、どうしても足りなければ 1→0 日へ自動緩和。
        - 同一人物が同ロールで連続日は不可。同日に複数シフトも不可。
        - 0 セルは固定で上書きしません。『上限(時間)』を超えない範囲で限界まで時間を使用します。
        - 編集ブロック: GH1 = E5‑AI16, GH2 = E20‑AI30 のみを書き換え。他セル (列 C の数式など) は保持します。
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
            st.success("✅ GH1・GH2 含め全日充足しました")

            st.subheader("最適化後のシフト表 (プレビュー)")
            st.dataframe(df_opt, use_container_width=True)

            st.subheader("勤務時間の合計 / 上限")
            st.dataframe(pd.DataFrame({"合計時間": totals, "上限時間": limits}))

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
