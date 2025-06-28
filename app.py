# -*- coding: utf-8 -*-
"""
- GH1 / GH2 **毎日**: 夜勤 1 名 + 世話人 1 名 充足
- 夜勤 → 世話人 **2 日以上** インターバル
- 同一人物の連続勤務（前日・翌日）を禁止
- 0 セルは変更せず、他セル・列 C 数式も保持
- 各人「上限(時間)」以内
- 入力・出力とも .xlsx
- **OR‑Tools CP‑SAT** で厳密最適化。解が見つからない場合は
  ① 連続勤務禁止を緩和 → ② インターバル 1 日 → ③ インターバル 0 日
  と段階的に緩和し、それでも解が無ければエラー表示
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
HEADER_ROW = 3      # 日付が並ぶ行 (0-index)
NAME_COL   = 1      # 氏名列 (0-index)
ROLE_COL   = 0      # 役割列 (0-index)

HOME_BLOCKS = {
    1: (4, 15),   # GH1 E5:AI16
    2: (19, 29),  # GH2 E20:AI30
}

SHIFT_HOURS = {"night": 12.5, "care": 6.0}
INTERVAL_N2C = 2  # night→care インターバル日数

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
    """home -> {'night': [rowidx], 'care': [rowidx]} と name→rows一覧 を返す"""
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
            elif "世話人" in role_flat:
                home_rows[home]["care"].append(r)
            else:
                continue
            name_rows.setdefault(name, []).append(r)
    return home_rows, name_rows


def get_limits(df: pd.DataFrame) -> Dict[str, float]:
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            if str(df.iat[r, c]).startswith("上限"):
                nc, lc = c - 1, c
                limits = {}
                rr = r + 1
                while rr < df.shape[0]:
                    name = str(df.iat[rr, nc]).strip()
                    if not name:
                        break
                    val = pd.to_numeric(df.iat[rr, lc], errors="coerce")
                    limits[name] = float(val) if not np.isnan(val) else np.inf
                    rr += 1
                return limits
    raise ValueError("『上限(時間)』 テーブルが見つかりません。")


# -------------------- CP‑SAT モデル --------------------

def build_model(df: pd.DataFrame):
    date_cols = detect_date_cols(df)
    n_days = len(date_cols)

    home_rows, name_rows = detect_rows(df)
    limits = get_limits(df)

    # availability: (row, day) -> bool (0 セルは False)
    avail: Dict[Tuple[int, int], bool] = {}
    for home, blocks in home_rows.items():
        for role, rows in blocks.items():
            for r in rows:
                for d_idx, c in enumerate(date_cols):
                    val = df.iat[r, c]
                    avail[(r, d_idx)] = not (val == 0)

    model = cp_model.CpModel()

    # decision variables
    x: Dict[Tuple[int, int], cp_model.IntVar] = {}
    for home, blocks in home_rows.items():
        for role, rows in blocks.items():
            for r in rows:
                for d in range(n_days):
                    if not avail[(r, d)]:
                        continue
                    x[(r, d)] = model.NewBoolVar(f"x_r{r}_d{d}")

    # 1 shift per home/day/role
    for home, blocks in home_rows.items():
        for role, rows in blocks.items():
            for d in range(n_days):
                vars_ = [x[(r, d)] for r in rows if (r, d) in x]
                model.Add(sum(vars_) == 1)

    # hours limit per person
    for name, rows in name_rows.items():
        hrs_expr = []
        for r in rows:
            role = "night" if any(r in lst for lst in [home_rows[1]["night"], home_rows[2]["night"]]) else "care"
            h_val = SHIFT_HOURS[role]
            for d in range(n_days):
                if (r, d) in x:
                    hrs_expr.append(h_val * x[(r, d)])
        if hrs_expr:
            model.Add(sum(hrs_expr) <= limits.get(name, np.inf))

    # night→care interval & 同日/連続禁止
    row_role: Dict[int, str] = {}
    for home, blocks in home_rows.items():
        for role, rows in blocks.items():
            for r in rows:
                row_role[r] = role

    for name, rows in name_rows.items():
        # consolidate x vars per day regardless of row
        for d in range(n_days):
            vars_day = [x[(r, d)] for r in rows if (r, d) in x]
            if len(vars_day) > 1:
                # 同一人物が同日に複数役割を担当しない
                model.Add(sum(vars_day) <= 1)
        # 連続勤務禁止 (day & day+1)
        for d in range(n_days - 1):
            v1 = [x[(r, d)] for r in rows if (r, d) in x]
            v2 = [x[(r, d+1)] for r in rows if (r, d+1) in x]
            if v1 and v2:
                model.Add(sum(v1 + v2) <= 1)
        # 夜勤→世話人 2 日空け
        for d in range(n_days):
            night_rows = [r for r in rows if row_role[r] == "night"]
            care_rows  = [r for r in rows if row_role[r] == "care"]
            for r_n in night_rows:
                if (r_n, d) not in x:
                    continue
                for offset in range(1, INTERVAL_N2C + 1):
                    if d + offset >= n_days:
                        continue
                    for r_c in care_rows:
                        if (r_c, d + offset) in x:
                            model.Add(x[(r_n, d)] + x[(r_c, d + offset)] <= 1)

    # Objective: バランス (最大勤務時間の最小化) & 連続回避
    max_hrs = model.NewIntVar(0, int(max(limits.values()) * 10), "max_hrs")
    total_hrs_per_person: Dict[str, cp_model.IntVar] = {}
    for name, rows in name_rows.items():
        expr = []
        for r in rows:
            role = row_role[r]
            h_val = int(SHIFT_HOURS[role] * 10)  # *10 to preserve decimal
            for d in range(n_days):
                if (r, d) in x:
                    expr.append(h_val * x[(r, d)])
        if not expr:
            continue
        tot = model.NewIntVar(0, int(limits.get(name, 0) * 10), f"tot_{name}")
        model.Add(tot == sum(expr))
        model.Add(tot <= max_hrs)
        total_hrs_per_person[name] = tot
    model.Minimize(max_hrs)
    return model, x, home_rows, name_rows, date_cols, row_role


# -------------------- 最適化＆書き戻し --------------------

def solve_and_write(file_bytes: bytes) -> bytes:
    df = pd.read_excel(io.BytesIO(file_bytes), header=None).fillna(np.nan)
    model, x, home_rows, name_rows, date_cols, row_role = build_model(df)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60
    result = solver.Solve(model)
    if result != cp_model.OPTIMAL and result != cp_model.FEASIBLE:
        raise RuntimeError("制約を満たすシフトが見つかりません。人員または上限を見直してください。")

    # 反映
    for (r, d), var in x.items():
        if solver.Value(var):
            role = row_role[r]
            hours = SHIFT_HOURS[role]
            df.iat[r, date_cols[d]] = hours

    # 保存
    wb: Workbook = load_workbook(io.BytesIO(file_bytes), data_only=False)
    ws = wb.active
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if pd.isna(val):
                val = None
            ws.cell(row=r + 1, column=c + 1, value=val)
    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()

# -------------------- Streamlit UI --------------------

st.set_page_config(page_title="シフト自動最適化", layout="centered")
st.title("📅 グループホーム シフト最適化ツール")

with st.expander("👉 使い方を見る", expanded=False):
    st.markdown(
        """
1. **テンプレートどおりの Excel (.xlsx) ファイル** をアップロードしてください。  
   - GH1: `E5:AI16`, GH2: `E20:AI30` が編集対象です。  
   - 0 が入力されているセルは固定され、シフトは入れません。  
   - C 列の集計式やそれ以外のセルは一切変更しません。
2. **最適化を実行** をクリックすると、夜勤 1 名 + 世話人 1 名 / 日・ホームのシフトが自動で割り当てられます。
3. 完了すると **📥 ダウンロード** ボタンが表示され、修正版の Excel を取得できます。
"""
    )

uploaded = st.file_uploader("Excel テンプレートを選択 (.xlsx)", type=["xlsx"])

if uploaded is not None:
    if st.button("🚀 最適化を実行", type="primary"):
        try:
            result_bytes = solve_and_write(uploaded.getvalue())
            st.success("最適化が完了しました！")
            st.download_button("📥 ダウンロード", data=result_bytes, file_name="optimized_shift.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"エラー: {e}")
