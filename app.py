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
ortools>=9.9
============================================================
app.py  （"streamlit run app.py" で実行）
------------------------------------------------------------
【2025‑06‑29 INT 修正版】
------------------------------------------------------------
- GH1 / GH2 **毎日**: 夜勤 1 名 + 世話人 1 名
- 夜勤 → 世話人 **2 日以上** インターバル
- 同一人物の連続勤務禁止 (前後 1 日)
- 0 セル維持／C 列数式保持／上限内
- **OR‑Tools CP‑SAT** を整数モデルで利用（時間を 0.5h 単位の整数にスケール）
  → 浮動小数点による `le()` 例外を解消
"""

import io
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from ortools.sat.python import cp_model

# -------------------- 定数 --------------------
HEADER_ROW = 3      # 0-index で日付ヘッダー行 (E4=列4) → 行 3
NAME_COL   = 1      # 氏名列 (B 列)
ROLE_COL   = 0      # 役割列 (A 列)

HOME_BLOCKS = {
    1: (4, 15),   # GH1: E5:AI16 → 行 4‑15
    2: (19, 29),  # GH2: E20:AI30 → 行 19‑29
}

# 時間を整数で扱うため 0.5 時間単位でスケール
SCALE = 2               # 0.5h = 1 point
SHIFT_HOURS = {         # 実数
    "night": 12.5,
    "care": 6.0,
}
SHIFT_HOURS_INT = {k: int(v * SCALE + 0.5) for k, v in SHIFT_HOURS.items()}  # {night:25, care:12}
INTERVAL_N2C = 2  # night→care のインターバル（日）
BIG_M = 1_000_000

# -------------------- ユーティリティ --------------------

def detect_date_cols(df: pd.DataFrame) -> List[int]:
    cols = []
    for c in range(df.shape[1]):
        v = df.iat[HEADER_ROW, c]
        try:
            day = int(float(v))
            if 1 <= day <= 31:
                cols.append(c)
        except (ValueError, TypeError):
            continue
    if not cols:
        raise ValueError("ヘッダー行に 1‑31 の日付が見つかりません。")
    return cols


def detect_rows(df: pd.DataFrame) -> Tuple[Dict[int, Dict[str, List[int]]], Dict[str, List[int]]]:
    """戻り値: home→{'night':[row], 'care':[row]},  name→[rows]"""
    home_rows: Dict[int, Dict[str, List[int]]] = {1: {"night": [], "care": []}, 2: {"night": [], "care": []}}
    name_rows: Dict[str, List[int]] = {}
    for home, (rs, re) in HOME_BLOCKS.items():
        for r in range(rs, re + 1):
            role_raw = str(df.iat[r, ROLE_COL])
            name = str(df.iat[r, NAME_COL]).strip()
            if not name:
                continue
            role_flat = role_raw.replace("\n", "")
            if "夜間" in role_flat and "支援員" in role_flat:
                home_rows[home]["night"].append(r)
                name_rows.setdefault(name, []).append(r)
            elif "世話人" in role_flat:
                home_rows[home]["care"].append(r)
                name_rows.setdefault(name, []).append(r)
    return home_rows, name_rows


def get_limits(df: pd.DataFrame) -> Dict[str, float]:
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            if str(df.iat[r, c]).startswith("上限"):
                name_col, lim_col = c - 1, c
                limits = {}
                rr = r + 1
                while rr < df.shape[0]:
                    name = str(df.iat[rr, name_col]).strip()
                    if not name:
                        break
                    val = pd.to_numeric(df.iat[rr, lim_col], errors="coerce")
                    limits[name] = float(val) if not np.isnan(val) else np.inf
                    rr += 1
                return limits
    raise ValueError("『上限(時間)』 テーブルが見つかりません。")

# -------------------- CP‑SAT モデル --------------------

def build_model(df: pd.DataFrame):
    date_cols = detect_date_cols(df)
    n_days = len(date_cols)

    home_rows, name_rows = detect_rows(df)
    limits_real = get_limits(df)  # float 時間

    # availability map: True if シフト可
    avail: Dict[Tuple[int, int], bool] = {}
    for r in range(df.shape[0]):
        for d_idx, c in enumerate(date_cols):
            val = df.iat[r, c]
            avail[(r, d_idx)] = not (val == 0)

    model = cp_model.CpModel()

    # decision vars x[(r,d)] ∈ {0,1}
    x: Dict[Tuple[int, int], cp_model.IntVar] = {}
    for (r, d), ok in avail.items():
        if not ok:
            continue
        x[(r, d)] = model.NewBoolVar(f"x_r{r}_d{d}")

    # row → role dict
    row_role: Dict[int, str] = {}
    for home, roles in home_rows.items():
        for role, rows in roles.items():
            for r in rows:
                row_role[r] = role

    # --- 制約 ---
    # (1) 各ホーム・各日・各役 で 1 名
    for home, roles in home_rows.items():
        for role, rows in roles.items():
            for d in range(n_days):
                vars_ = [x[(r, d)] for r in rows if (r, d) in x]
                model.Add(sum(vars_) == 1)

    # (2) 個人の上限
    for name, rows in name_rows.items():
        expr = []
        for r in rows:
            role = row_role[r]
            h_int = SHIFT_HOURS_INT[role]
            for d in range(n_days):
                if (r, d) in x:
                    expr.append(h_int * x[(r, d)])
        if not expr:
            continue
        lim_int = int(limits_real.get(name, np.inf) * SCALE + 0.5)
        if lim_int >= BIG_M:
            continue
        model.Add(sum(expr) <= lim_int)

    # (3) 同一人物の同日複数禁止 & 連続禁止
    for name, rows in name_rows.items():
        for d in range(n_days):
            vars_day = [x[(r, d)] for r in rows if (r, d) in x]
            if len(vars_day) > 1:
                model.Add(sum(vars_day) <= 1)
        for d in range(n_days - 1):
            v1 = [x[(r, d)] for r in rows if (r, d) in x]
            v2 = [x[(r, d + 1)] for r in rows if (r, d + 1) in x]
            if v1 and v2:
                model.Add(sum(v1 + v2) <= 1)

    # (4) 夜勤→世話人 インターバル 2 日
    for name, rows in name_rows.items():
        night_rows = [r for r in rows if row_role[r] == "night"]
        care_rows  = [r for r in rows if row_role[r] == "care"]
        for d in range(n_days):
            for r_n in night_rows:
                if (r_n, d) not in x:
                    continue
                for dt in range(1, INTERVAL_N2C + 1):
                    if d + dt >= n_days:
                        continue
                    for r_c in care_rows:
                        if (r_c, d + dt) in x:
                            model.Add(x[(r_n, d)] + x[(r_c, d + dt)] <= 1)

    # 目的: 最大労働時間の最小化
    max_hrs = model.NewIntVar(0, BIG_M, "max_hrs")
    for name, rows in name_rows.items():
        expr = []
        for r in rows:
            role = row_role[r]
            h_int = SHIFT_HOURS_INT[role]
            for d in range(n_days):
                if (r, d) in x:
                    expr.append(h_int * x[(r, d)])
        if not expr:
            continue
        tot = model.NewIntVar(0, BIG_M, f"tot_{name}")
        model.Add(tot == sum(expr))
        model.Add(tot <= max_hrs)
    model.Minimize(max_hrs)

    return model, x, date_cols, row_role

# -------------------- 解いて書き戻し --------------------

def solve_and_write(file_bytes: bytes) -> bytes:
    df = pd.read_excel(io.BytesIO(file_bytes), header=None).fillna(np.nan)

    model, x, date_cols, row_role = build_model(df)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120
    result = solver.Solve(model)
    if result not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("制約を満たすシフトが見つかりません。人員または上限を見直してください。")

    # apply solution
    for (r, d), var in x.items():
        if solver.Value(var):
            val = SHIFT_HOURS[row_role[r]]
            df.iat[r, date_cols[d]] = val
        else:
            # 未選択セルは空白に戻す
            if pd.notna(df.iat[r, date_cols[d]]) and df.iat[r, date_cols[d]] in SHIFT_HOURS.values():
                df.iat[r, date_cols[d]] = np.nan

    # save back via openpyxl to preserve formulas
    wb: Workbook = load_workbook(io.BytesIO(file_bytes), data_only=False)
    ws = wb.active
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            ws.cell(row=r+1, column=c+1, value=None if pd.isna(val) else val)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# -------------------- UI --------------------

st.set_page_config(page_title="シフト自動最適化", layout="centered")
st.title("📅 グループホーム シフト最適化ツール (INTモデル)")

with st.expander("👉 使い方を見る", expanded=False):
    st.markdown(
        """
1. **テンプレートの Excel (.xlsx)** をアップロードしてください。\
   - GH1:E5‑AI16, GH2:E20‑AI30 が対象セルです。\
   - 0 が入っているセルは固定されます。\
2. **最適化を実行** をクリックすると、毎日**夜勤 1 + 世話人 1**が各ホームに割り当てられます。\
3. 完了すると **📥 ダウンロード** ボタンが現れ、修正済みファイルを取得できます。
"""
    )

uploaded = st.file_uploader("Excel テンプレートを選択 (.xlsx)", type=["xlsx"])

if uploaded is not None:
    if st.button("🚀 最適化を実行", type="primary"):
        try:
            data = uploaded.getvalue()
            result_bytes = solve_and_write(data)
            st.success("最適化が完了しました！")
            st.download_button("📥 ダウンロード", data=result_bytes, file_name="optimized_shift.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"エラー: {e}")
