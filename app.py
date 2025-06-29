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
from typing import List, Tuple, Dict, Set
import random

import numpy as np
import pandas as pd
import streamlit as st

# -------------------- 定数 --------------------
# CSVファイルの構造に基づく調整
NIGHT_ROWS_GH1 = list(range(10, 16))  # グループホーム①夜勤（11-16行目、0-indexedで10-15）
CARE_ROWS_GH1  = list(range(5, 10))   # グループホーム①世話人（6-10行目、0-indexedで5-9）
NIGHT_ROWS_GH2 = list(range(25, 31))  # グループホーム②夜勤（26-31行目、0-indexedで25-30）
CARE_ROWS_GH2  = list(range(20, 25))  # グループホーム②世話人（21-25行目、0-indexedで20-24）
DATE_HEADER_ROW = 4                   # 5行目（0-index 4）
DATE_START_COL = 5                    # 日付データは6列目以降（0-indexedで5以降）

# -------------------- 関数群 --------------------

def detect_date_columns(df: pd.DataFrame) -> List[str]:
    """ヘッダーから日付列を推定し、連続する範囲（列名リスト）を返す"""
    date_cols = []
    
    # 5列目以降を日付列として扱う
    for col_idx in range(DATE_START_COL, len(df.columns)):
        col = df.columns[col_idx]
        try:
            header = df.iloc[DATE_HEADER_ROW, col_idx]
            if pd.notna(header):
                day = int(float(header))
                if 1 <= day <= 31:
                    date_cols.append(col)
        except (ValueError, TypeError):
            pass
    
    if not date_cols:
        raise ValueError("日付列を検出できませんでした。")
    
    return date_cols


def get_staff_info(df: pd.DataFrame) -> Tuple[List[Tuple[str, int]], List[Tuple[str, int]], List[Tuple[str, int]], List[Tuple[str, int]], Dict[str, float]]:
    """スタッフ情報を取得する"""
    night_staff_gh1 = []
    care_staff_gh1 = []
    night_staff_gh2 = []
    care_staff_gh2 = []
    limits = {}
    
    # グループホーム①の夜勤スタッフ
    for row in NIGHT_ROWS_GH1:
        if row < len(df):
            role = str(df.iloc[row, 1]).strip()  # B列: 役職
            name = str(df.iloc[row, 2]).strip()  # C列: 名前
            limit_val = df.iloc[row, 3] if pd.notna(df.iloc[row, 3]) else 0  # D列: 上限
            
            if role and name and role != 'nan' and name != 'nan' and '夜間' in role:
                night_staff_gh1.append((name, row))
                limits[f"{name}_GH1_夜勤"] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    # グループホーム①の世話人スタッフ
    for row in CARE_ROWS_GH1:
        if row < len(df):
            role = str(df.iloc[row, 1]).strip()  # B列: 役職
            name = str(df.iloc[row, 2]).strip()  # C列: 名前
            limit_val = df.iloc[row, 3] if pd.notna(df.iloc[row, 3]) else 0  # D列: 上限
            
            if role and name and role != 'nan' and name != 'nan' and '世話人' in role:
                care_staff_gh1.append((name, row))
                limits[f"{name}_GH1_世話人"] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    # グループホーム②の夜勤スタッフ
    for row in NIGHT_ROWS_GH2:
        if row < len(df):
            role = str(df.iloc[row, 1]).strip()  # B列: 役職
            name = str(df.iloc[row, 2]).strip()  # C列: 名前
            limit_val = df.iloc[row, 3] if pd.notna(df.iloc[row, 3]) else 0  # D列: 上限
            
            if role and name and role != 'nan' and name != 'nan' and '夜間' in role:
                night_staff_gh2.append((name, row))
                limits[f"{name}_GH2_夜勤"] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    # グループホーム②の世話人スタッフ
    for row in CARE_ROWS_GH2:
        if row < len(df):
            role = str(df.iloc[row, 1]).strip()  # B列: 役職
            name = str(df.iloc[row, 2]).strip()  # C列: 名前
            limit_val = df.iloc[row, 3] if pd.notna(df.iloc[row, 3]) else 0  # D列: 上限
            
            if role and name and role != 'nan' and name != 'nan' and '世話人' in role:
                care_staff_gh2.append((name, row))
                limits[f"{name}_GH2_世話人"] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    return night_staff_gh1, care_staff_gh1, night_staff_gh2, care_staff_gh2, limits


def parse_constraints(df: pd.DataFrame, staff_list: List[Tuple[str, int]]) -> Dict[str, str]:
    """D列の制約を解析"""
    constraints = {}
    for name, row in staff_list:
        constraint = str(df.iloc[row, 4]).strip() if pd.notna(df.iloc[row, 4]) else ""  # E列: 制約
        constraints[name] = constraint
    return constraints


def can_work_on_day(constraint: str, day: int, day_of_week: str) -> bool:
    """制約に基づいて指定日に勤務可能かチェック"""
    if not constraint or constraint == "条件なし" or str(constraint) == "nan" or constraint in ["0.5"]:
        return True
    
    constraint = str(constraint).lower()
    
    # 曜日制約
    if "日曜" in constraint and day_of_week != "日":
        return False
    if "月曜" in constraint and day_of_week != "月":
        return False
    if "火曜" in constraint and day_of_week != "火":
        return False
    if "水曜" in constraint and day_of_week != "水":
        return False
    if "木曜" in constraint and day_of_week != "木":
        return False
    if "金曜" in constraint and day_of_week != "金":
        return False
    if "土曜" in constraint and day_of_week != "土":
        return False
    
    # 特定日制約
    if "日" in constraint and not any(wd in constraint for wd in ["月", "火", "水", "木", "金", "土", "日"]):
        import re
        days = re.findall(r'(\d+)日', constraint)
        if days and str(day) not in days:
            return False
    
    return True


def can_assign_shift(df: pd.DataFrame, staff_name: str, staff_row: int, day_col: str, 
                    date_cols: List[str], constraints: Dict[str, str],
                    assignment_history: Dict[str, List[int]]) -> bool:
    """指定したスタッフが指定日にシフトに入れるかチェック"""
    
    day_col_idx = df.columns.get_loc(day_col)
    current_day_idx = date_cols.index(day_col)
    
    # 当日が0（勤務不可）でないかチェック
    current_value = df.iloc[staff_row, day_col_idx]
    if current_value == 0:
        return False
    
    # 制約チェック
    try:
        day_num = int(float(df.iloc[DATE_HEADER_ROW, day_col_idx]))
    except (ValueError, TypeError):
        day_num = 1
    
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    day_of_week = weekdays[(day_num - 1) % 7]
    
    if not can_work_on_day(constraints.get(staff_name, ""), day_num, day_of_week):
        return False
    
    # 連続勤務の制限
    staff_history = assignment_history.get(staff_name, [])
    if staff_history and current_day_idx - 1 in staff_history:
        return False
    
    return True


def count_staff_hours(assignment_history: Dict[str, List[int]], staff_name: str, 
                     is_night_shift: bool) -> float:
    """スタッフの総勤務時間を計算"""
    if staff_name not in assignment_history:
        return 0
    
    total_hours = 0
    for day_idx in assignment_history[staff_name]:
        if is_night_shift:
            total_hours += 12.5
        else:
            total_hours += 6
    
    return total_hours


def assign_shifts_for_group(df: pd.DataFrame, date_cols: List[str], 
                           night_staff: List[Tuple[str, int]], care_staff: List[Tuple[str, int]],
                           night_constraints: Dict[str, str], care_constraints: Dict[str, str],
                           assignment_history: Dict[str, List[int]], limits: Dict[str, float],
                           group_name: str) -> None:
    """特定のグループのシフト割り当て"""
    
    for day_col in date_cols:
        day_col_idx = df.columns.get_loc(day_col)
        current_day_idx = date_cols.index(day_col)
        
        # 夜勤シフト割当（1名のみ残す）
        available_night_staff = []
        for name, row in night_staff:
            if can_assign_shift(df, name, row, day_col, date_cols, night_constraints, assignment_history):
                current_hours = count_staff_hours(assignment_history, name, True)
                limit_key = f"{name}_{group_name}_夜勤"
                if limit_key in limits and current_hours + 12.5 <= limits[limit_key]:
                    current_value = df.iloc[row, day_col_idx]
                    if current_value != 0 and pd.notna(current_value):
                        available_night_staff.append((name, row, current_hours))
        
        # 勤務時間が少ないスタッフを優先
        if available_night_staff:
            available_night_staff.sort(key=lambda x: x[2])
            selected_name, selected_row, _ = available_night_staff[0]
            assignment_history[selected_name].append(current_day_idx)
            
            # 選ばれなかったスタッフのセルを空白にする
            for name, row in night_staff:
                if row != selected_row:
                    df.iloc[row, day_col_idx] = ""
        else:
            # 誰も選ばれなかった場合は全て空白
            for name, row in night_staff:
                df.iloc[row, day_col_idx] = ""
        
        # 世話人シフト割当（1名のみ残す）
        available_care_staff = []
        for name, row in care_staff:
            if can_assign_shift(df, name, row, day_col, date_cols, care_constraints, assignment_history):
                current_hours = count_staff_hours(assignment_history, name, False)
                limit_key = f"{name}_{group_name}_世話人"
                if limit_key in limits and current_hours + 6 <= limits[limit_key]:
                    current_value = df.iloc[row, day_col_idx]
                    if current_value != 0 and pd.notna(current_value):
                        available_care_staff.append((name, row, current_hours))
        
        # 勤務時間が少ないスタッフを優先
        if available_care_staff:
            available_care_staff.sort(key=lambda x: x[2])
            selected_name, selected_row, _ = available_care_staff[0]
            assignment_history[selected_name].append(current_day_idx)
            
            # 選ばれなかったスタッフのセルを空白にする
            for name, row in care_staff:
                if row != selected_row:
                    df.iloc[row, day_col_idx] = ""
        else:
            # 誰も選ばれなかった場合は全て空白
            for name, row in care_staff:
                df.iloc[row, day_col_idx] = ""


def optimize_shifts(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    """シフト最適化ロジック"""
    date_cols = detect_date_columns(df)
    night_staff_gh1, care_staff_gh1, night_staff_gh2, care_staff_gh2, limits = get_staff_info(df)
    
    # 制約情報を取得
    night_constraints_gh1 = parse_constraints(df, night_staff_gh1)
    care_constraints_gh1 = parse_constraints(df, care_staff_gh1)
    night_constraints_gh2 = parse_constraints(df, night_staff_gh2)
    care_constraints_gh2 = parse_constraints(df, care_staff_gh2)
    
    # 割り当て履歴を追跡
    all_staff = [name for name, _ in night_staff_gh1 + care_staff_gh1 + night_staff_gh2 + care_staff_gh2]
    assignment_history = {name: [] for name in all_staff}
    
    # グループホーム①のシフト割り当て
    assign_shifts_for_group(df, date_cols, night_staff_gh1, care_staff_gh1,
                           night_constraints_gh1, care_constraints_gh1,
                           assignment_history, limits, "GH1")
    
    # グループホーム②のシフト割り当て
    assign_shifts_for_group(df, date_cols, night_staff_gh2, care_staff_gh2,
                           night_constraints_gh2, care_constraints_gh2,
                           assignment_history, limits, "GH2")
    
    # -------------------- 結果の集計 --------------------
    staff_totals = {}
    staff_limits = {}
    
    # グループホーム①
    for name, row in night_staff_gh1:
        key = f"{name}(GH1夜勤)"
        night_hours = count_staff_hours(assignment_history, name, True)
        staff_totals[key] = night_hours
        staff_limits[key] = limits.get(f"{name}_GH1_夜勤", 0)
    
    for name, row in care_staff_gh1:
        key = f"{name}(GH1世話人)"
        care_hours = count_staff_hours(assignment_history, name, False)
        staff_totals[key] = care_hours
        staff_limits[key] = limits.get(f"{name}_GH1_世話人", 0)
    
    # グループホーム②
    for name, row in night_staff_gh2:
        key = f"{name}(GH2夜勤)"
        night_hours = count_staff_hours(assignment_history, name, True)
        staff_totals[key] = night_hours
        staff_limits[key] = limits.get(f"{name}_GH2_夜勤", 0)
    
    for name, row in care_staff_gh2:
        key = f"{name}(GH2世話人)"
        care_hours = count_staff_hours(assignment_history, name, False)
        staff_totals[key] = care_hours
        staff_limits[key] = limits.get(f"{name}_GH2_世話人", 0)
    
    totals = pd.Series(staff_totals, dtype=float)
    limits_series = pd.Series(staff_limits, dtype=float)
    
    return df, totals, limits_series


# -------------------- Streamlit UI --------------------

st.set_page_config(page_title="シフト自動最適化", layout="wide")
st.title("📅 シフト自動最適化ツール")

with st.expander("👉 使い方はこちら（クリックで展開）", expanded=False):
    st.markdown(
        """
        **▼ 手順**
        1. 左サイドバーで **CSV/Excel ファイル** を選択してアップロード。
        2. **「🚀 最適化を実行」** ボタンを押す。
        3. 右側に最適化後のシフトプレビューが表示される。
        4. **「📥 ダウンロード」** ボタンで Excel を取得。

        **▼ 処理方式**
        - グループホーム①とグループホーム②のそれぞれで、**各日1人だけ**数値を残す
        - **選ばれなかった人のセルは空白**になります
        - **0が入っているセルは勤務不可**として維持
        - **それ以外のセル（スタッフ名、上限時間、制約等）は一切変更しません**

        **▼ 実装されたルール**
        - 夜勤・世話人は1日に1人ずつ選択（各グループで）
        - E列の制約（曜日、特定日など）を考慮
        - 連続勤務は避ける
        - 各スタッフの上限時間を考慮（夜勤12.5h、世話人6h）
        - 勤務時間が少ないスタッフを優先的に割当

        *元のテンプレート構造を完全に保持します。*
        """
    )

st.sidebar.header("📂 入力ファイル")
uploaded = st.sidebar.file_uploader("CSV/Excel ファイル", type=["csv", "xlsx"])

if uploaded is not None:
    try:
        # ファイル形式に応じて読み込み
        if uploaded.name.endswith('.csv'):
            df_input = pd.read_csv(uploaded, header=None, encoding='utf-8')
        else:
            df_input = pd.read_excel(uploaded, header=None, engine="openpyxl")
        
        # スタッフ情報を事前に取得して表示
        try:
            night_staff_gh1, care_staff_gh1, night_staff_gh2, care_staff_gh2, limits = get_staff_info(df_input)
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("グループホーム① スタッフ")
                st.write("**夜勤:**")
                for name, row in night_staff_gh1:
                    constraint = df_input.iloc[row, 4] if pd.notna(df_input.iloc[row, 4]) else "条件なし"
                    st.write(f"• {name} (行{row+1}) - {constraint}")
                st.write("**世話人:**")
                for name, row in care_staff_gh1:
                    constraint = df_input.iloc[row, 4] if pd.notna(df_input.iloc[row, 4]) else "条件なし"
                    st.write(f"• {name} (行{row+1}) - {constraint}")
            
            with col2:
                st.subheader("グループホーム② スタッフ")
                st.write("**夜勤:**")
                for name, row in night_staff_gh2:
                    constraint = df_input.iloc[row, 4] if pd.notna(df_input.iloc[row, 4]) else "条件なし"
                    st.write(f"• {name} (行{row+1}) - {constraint}")
                st.write("**世話人:**")
                for name, row in care_staff_gh2:
                    constraint = df_input.iloc[row, 4] if pd.notna(df_input.iloc[row, 4]) else "条件なし"
                    st.write(f"• {name} (行{row+1}) - {constraint}")
        
        except Exception as e:
            st.warning(f"スタッフ情報の取得でエラー: {e}")
        
        st.subheader("アップロードされたシフト表")
        st.dataframe(df_input, use_container_width=True)

        if st.sidebar.button("🚀 最適化を実行"):
            with st.spinner("シフトを最適化中..."):
                df_opt, totals, limits_series = optimize_shifts(df_input.copy())
            
            st.success("最適化が完了しました 🎉")

            st.subheader("最適化後のシフト表")
            st.dataframe(df_opt, use_container_width=True)

            if not limits_series.empty:
                st.subheader("勤務時間の合計と上限")
                
                comparison_data = []
                for staff_key in totals.index:
                    comparison_data.append({
                        "スタッフ": staff_key,
                        "合計時間": totals[staff_key],
                        "上限時間": limits_series[staff_key],
                        "残り時間": limits_series[staff_key] - totals[staff_key]
                    })
                
                comparison_df = pd.DataFrame(comparison_data)
                
                def highlight_over_limit(row):
                    color = 'background-color: red' if row['残り時間'] < 0 else ''
                    return [color] * len(row)
                
                if len(comparison_df) > 0:
                    styled_df = comparison_df.style.apply(highlight_over_limit, axis=1)
                    st.dataframe(styled_df, use_container_width=True)
                    
                    over_limit_staff = comparison_df[comparison_df['残り時間'] < 0]['スタッフ'].tolist()
                    if over_limit_staff:
                        st.warning(f"⚠️ 上限時間を超過しているスタッフ: {', '.join(over_limit_staff)}")
                else:
                    st.dataframe(comparison_df, use_container_width=True)

            # Excel 出力
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_opt.to_excel(writer, index=False, header=False, sheet_name="最適化シフト")
                
                if not limits_series.empty:
                    comparison_df.to_excel(writer, sheet_name="勤務時間統計", index=False)
                
            st.download_button(
                label="📥 最適化シフトをダウンロード(.xlsx)",
                data=buffer.getvalue(),
                file_name="optimized_shift.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"ファイルの読み込みまたは最適化中にエラーが発生しました: {e}")
        st.error("詳細エラー情報:")
        st.exception(e)
else:
    st.info("左のサイドバーからCSVまたはExcelファイルをアップロードしてください。")
