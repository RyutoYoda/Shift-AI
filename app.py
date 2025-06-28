import io
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

# -------------------- 定数 --------------------
HEADER_ROW = 3      # 日付が並ぶ行 (0-index)
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
    date_cols = detect_date_columns(df)
    for start, end in HOME_BLOCKS.values():
        for r in range(start, end + 1):
            for c in date_cols:
                if df.iat[r, c] != 0 and not pd.isna(df.iat[r, c]):
                    df.iat[r, c] = np.nan


def choose_candidate(cands):
    """(残余時間, 連続回避, name, row) でソート"""
    return sorted(cands, key=lambda x: (x[1], x[0], x[2]))[0]


def optimize(df: pd.DataFrame):
    date_cols = detect_date_columns(df)
    night_rows, care_rows = detect_rows_by_home(df)

    limits = get_limits(df)
    names = limits.index.tolist()

    clear_blocks(df)

    totals = pd.Series(0.0, index=names)

    last_night_day: Dict[str, int] = {}
    last_any_day: Dict[str, int] = {}
    last_role_day: Dict[Tuple[str, str], int] = {}

    for d_idx, c in enumerate(date_cols):
        assigned_today: set[str] = set()
        for home in (1, 2):
            # ---------- 夜勤 ----------
            night_cands_strict, night_cands_relax = [], []
            for r in night_rows[home]:
                name = df.iat[r, NAME_COL].strip()
                if not pd.isna(df.iat[r, c]):
                    continue
                if name in assigned_today:
                    continue
                if totals[name] + SHIFT_NIGHT_HOURS > limits.get(name, np.inf):
                    continue
                # 連続禁止 (前日 / 翌日) 判定
                consecutive = last_any_day.get(name, -99) == d_idx - 1
                cand = (limits[name] - totals[name], consecutive, name, r)
                if consecutive:
                    night_cands_relax.append(cand)
                else:
                    night_cands_strict.append(cand)
            chosen = None
            for pool in (night_cands_strict, night_cands_relax):
                if pool:
                    chosen = choose_candidate(pool)
                    break
            if not chosen:
                raise RuntimeError(f"{d_idx+1} 日目 GH{home} の夜勤が充当できません。")
            _, _, night_name, night_row = chosen
            df.iat[night_row, c] = SHIFT_NIGHT_HOURS
            totals[night_name] += SHIFT_NIGHT_HOURS
            last_night_day[night_name] = d_idx
            last_any_day[night_name] = d_idx
            last_role_day[(night_name, f"night_{home}")] = d_idx
            assigned_today.add(night_name)

            # ---------- 世話人 ----------
            care_cands_strict, care_cands_relax = [], []
            for r in care_rows[home]:
                name = df.iat[r, NAME_COL].strip()
                if not pd.isna(df.iat[r, c]):
                    continue
                if name in assigned_today:
                    continue
                if totals[name] + SHIFT_CARE_HOURS > limits.get(name, np.inf):
                    continue
                # 夜勤→世話人インターバル
                interval_days = d_idx - last_night_day.get(name, -99)
                if interval_days <= 2:
                    continue  # 原則 2 日離す
                # 連続禁止
                consecutive = last_any_day.get(name, -99) == d_idx - 1
                cand = (limits[name] - totals[name], consecutive, name, r)
                if consecutive:
                    care_cands_relax.append(cand)
                else:
                    care_cands_strict.append(cand)
            # 緩和ステップ: (A)連続許可 → (B)夜勤→世話人 1 日 → (C)夜勤→世話人 0 日
            chosen = None
            for pool in (care_cands_strict, care_cands_relax):
                if pool:
                    chosen = choose_candidate(pool)
                    break
            if not chosen:
                # (B) interval 1 日
                care_cands = []
                for r in care_rows[home]:
                    name = df.iat[r, NAME_COL].strip()
                    if name in assigned_today:
                        continue
                    if totals[name] + SHIFT_CARE_HOURS > limits.get(name, np.inf):
                        continue
                    if d_idx - last_night_day.get(name, -99) <= 1:
                        continue  # 1 日も空いていない
                    consecutive = last_any_day.get(name, -99) == d_idx - 1
                    care_cands.append((limits[name] - totals[name], consecutive, name, r))
                if care_cands:
                    chosen = choose_candidate(care_cands)
            if not chosen:
                # (C) interval 0 日 (最終手段)
                care_cands = []
                for r in care_rows[home]:
                    name = df.iat[r, NAME_COL].strip()
                    if name in assigned_today:
                        continue
                    if totals[name] + SHIFT_CARE_HOURS > limits.get(name, np.inf):
                        continue
                    consecutive = last_any_day.get(name, -99) == d_idx - 1
                    care_cands.append((limits[name] - totals[name], consecutive, name, r))
                if care_cands:
                    chosen = choose_candidate(care_cands)
            if not chosen:
                raise RuntimeError(f"{d_idx+1} 日目 GH{home} の世話人が充当できません。")

            _, _, care_name, care_row = chosen
            df.iat[care_row, c] = SHIFT_CARE_HOURS
            totals[care_name] += SHIFT_CARE_HOURS
            last_any_day[care_name] = d_idx
            last_role_day[(care_name, f"care_{home}")] = d_idx
            assigned_today.add(care_name)

    # 完成チェック
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
