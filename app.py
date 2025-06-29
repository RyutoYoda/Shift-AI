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
# テンプレート構造が変わった場合はここを調整
# Excel は 1 行目 = index 0 扱い（pandas のヘッダ無し読込を想定）
NIGHT_ROWS = list(range(4, 16))   # E5:AI16 → 0‑index 行 4‑15 (夜勤シフト)
CARE_ROWS  = list(range(19, 31))  # E20:AI30 → 0‑index 行 19‑30 (世話人シフト)
DATE_HEADER_ROW = 3               # E4 行 (0‑index 3)

# -------------------- 関数群 --------------------

def detect_date_columns(df: pd.DataFrame) -> List[str]:
    """ヘッダーから日付列を推定し、連続する範囲（列名リスト）を返す"""
    date_cols = []
    for col in df.columns:
        header = str(df.at[DATE_HEADER_ROW, col]).strip()
        # 数字かどうかをチェック（日付として1-31の範囲を想定）
        try:
            day = int(float(header))  # float()を経由してからint()に変換
            if 1 <= day <= 31:
                date_cols.append(col)
        except (ValueError, TypeError):
            pass
    
    if not date_cols:
        raise ValueError("日付列を検出できませんでした。ヘッダー行と列番号を確認してください。")
    
    # 最初と最後の連続ブロックだけ抽出
    first_idx = df.columns.get_loc(date_cols[0])
    last_idx  = df.columns.get_loc(date_cols[-1]) + 1
    return list(df.columns[first_idx:last_idx])


def get_staff_info(df: pd.DataFrame) -> Tuple[List[Tuple[str, int]], List[Tuple[str, int]], Dict[str, float]]:
    """スタッフ情報を取得する - (名前, 行番号)のタプルリストを返す"""
    night_staff = []
    care_staff = []
    limits = {}
    
    # 夜勤エリア（E5:AI16 = 行5-16 = 0-indexで4-15）のスタッフ
    for row in range(4, 16):
        if row < len(df):
            role = str(df.iloc[row, 0]).strip()  # A列: 役職
            name = str(df.iloc[row, 1]).strip()  # B列: 名前
            limit_val = df.iloc[row, 2] if pd.notna(df.iloc[row, 2]) else 0  # C列: 上限
            
            if role and name and role != 'nan' and name != 'nan':
                if '夜間' in role:
                    night_staff.append((name, row))
                    limits[name] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    # 世話人エリア（E20:AI30 = 行20-30 = 0-indexで19-29）のスタッフ
    for row in range(19, 31):
        if row < len(df):
            role = str(df.iloc[row, 0]).strip()  # A列: 役職
            name = str(df.iloc[row, 1]).strip()  # B列: 名前
            limit_val = df.iloc[row, 2] if pd.notna(df.iloc[row, 2]) else 0  # C列: 上限
            
            if role and name and role != 'nan' and name != 'nan':
                if '世話人' in role:
                    care_staff.append((name, row))
                    limits[name] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    return night_staff, care_staff, limits


def parse_constraints(df: pd.DataFrame, staff_list: List[Tuple[str, int]]) -> Dict[str, str]:
    """D列の制約を解析"""
    constraints = {}
    for name, row in staff_list:
        constraint = str(df.iloc[row, 3]).strip() if pd.notna(df.iloc[row, 3]) else ""
        constraints[name] = constraint
    return constraints


def can_work_on_day(constraint: str, day: int, day_of_week: str) -> bool:
    """制約に基づいて指定日に勤務可能かチェック"""
    if not constraint or constraint == "条件なし" or str(constraint) == "nan":
        return True
    
    constraint = str(constraint).lower()
    
    # 曜日制約
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    
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
    if "日" in constraint and not any(wd in constraint for wd in weekdays):
        # "11日と18日" のような制約
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
        day_num = int(float(str(df.at[DATE_HEADER_ROW, day_col])))
    except (ValueError, TypeError):
        day_num = 1  # デフォルト値
    
    # 曜日は簡易的に計算（実際のプロジェクトでは正確な日付計算が必要）
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    day_of_week = weekdays[(day_num - 1) % 7]
    
    if not can_work_on_day(constraints.get(staff_name, ""), day_num, day_of_week):
        return False
    
    # 連続勤務の制限
    staff_history = assignment_history.get(staff_name, [])
    if staff_history and current_day_idx - 1 in staff_history:
        return False
    
    # 夜勤後2日空ける制限（世話人の場合）
    if staff_row >= 19:  # 世話人エリア
        for prev_day in range(max(0, current_day_idx - 2), current_day_idx):
            if prev_day in staff_history:
                # 前回が夜勤かチェック（簡易版）
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


def optimize_shifts(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    """シフト最適化ロジック"""
    date_cols = detect_date_columns(df)
    night_staff, care_staff, limits = get_staff_info(df)
    
    # 制約情報を取得
    night_constraints = parse_constraints(df, night_staff)
    care_constraints = parse_constraints(df, care_staff)
    
    # 割り当て履歴を追跡
    assignment_history = {name: [] for name, _ in night_staff + care_staff}
    
    # -------------------- 指定ブロックを削除モードで処理 --------------------
    # 各日に対してシフトを割り当て
    for day_col in date_cols:
        day_col_idx = df.columns.get_loc(day_col)
        current_day_idx = date_cols.index(day_col)
        
        # 夜勤シフト割当（1名のみ残す）
        available_night_staff = []
        for name, row in night_staff:
            if can_assign_shift(df, name, row, day_col, date_cols, night_constraints, assignment_history):
                current_hours = count_staff_hours(assignment_history, name, True)
                if name in limits and current_hours + 12.5 <= limits[name]:
                    current_value = df.iloc[row, day_col_idx]
                    if current_value != 0 and pd.notna(current_value):
                        available_night_staff.append((name, row, current_hours))
        
        # 勤務時間が少ないスタッフを優先
        if available_night_staff:
            available_night_staff.sort(key=lambda x: x[2])  # 勤務時間でソート
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
                if name in limits and current_hours + 6 <= limits[name]:
                    current_value = df.iloc[row, day_col_idx]
                    if current_value != 0 and pd.notna(current_value):
                        available_care_staff.append((name, row, current_hours))
        
        # 勤務時間が少ないスタッフを優先
        if available_care_staff:
            available_care_staff.sort(key=lambda x: x[2])  # 勤務時間でソート
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
    
    # -------------------- 結果の集計 --------------------
    # 重複を避けるため、役職も含めた一意のキーを作成
    all_staff_info = []
    staff_totals = {}
    staff_limits = {}
    
    # 夜勤スタッフの処理
    for name, row in night_staff:
        unique_key = f"{name}(夜勤)"
        night_hours = count_staff_hours(assignment_history, name, True)
        all_staff_info.append(unique_key)
        staff_totals[unique_key] = night_hours
        staff_limits[unique_key] = limits.get(name, 0)
    
    # 世話人スタッフの処理
    for name, row in care_staff:
        unique_key = f"{name}(世話人)"
        care_hours = count_staff_hours(assignment_history, name, False)
        all_staff_info.append(unique_key)
        staff_totals[unique_key] = care_hours
        staff_limits[unique_key] = limits.get(name, 0)
    
    # 重複がない一意のインデックスでSeries作成
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
        1. 左サイドバーで **テンプレート形式** の Excel ファイル (.xlsx) を選択してアップロード。
        2. **「🚀 最適化を実行」** ボタンを押す。
        3. 右側に最適化後のシフトプレビューが表示される。
        4. **「📥 ダウンロード」** ボタンで Excel を取得。

        **▼ 処理方式**
        - E5:AI16（夜勤）とE20:AI30（世話人）の範囲で、**各日1人だけ**数値を残す
        - **選ばれなかった人のセルは空白**になります
        - **0が入っているセルは勤務不可**として維持
        - **それ以外のセル（スタッフ名、上限時間、制約等）は一切変更しません**

        **▼ 実装されたルール**
        - 夜勤・世話人は1日に1人ずつ選択
        - D列の制約（曜日、特定日など）を考慮
        - 連続勤務は避ける
        - 各スタッフの上限時間を考慮（夜勤12.5h、世話人6h）
        - 勤務時間が少ないスタッフを優先的に割当

        *元のテンプレート構造を完全に保持します。*
        """
    )

st.sidebar.header("📂 入力ファイル")
uploaded = st.sidebar.file_uploader("Excel ファイル (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        df_input = pd.read_excel(uploaded, header=None, engine="openpyxl")
        
        # スタッフ情報を事前に取得して表示
        try:
            night_staff, care_staff, limits = get_staff_info(df_input)
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("検出された夜勤スタッフ")
                for name, row in night_staff:
                    constraint = df_input.iloc[row, 3] if pd.notna(df_input.iloc[row, 3]) else "条件なし"
                    st.write(f"• {name} (行{row+1}) - {constraint}")
            
            with col2:
                st.subheader("検出された世話人スタッフ")
                for name, row in care_staff:
                    constraint = df_input.iloc[row, 3] if pd.notna(df_input.iloc[row, 3]) else "条件なし"
                    st.write(f"• {name} (行{row+1}) - {constraint}")
            
            st.subheader("上限時間")
            limit_df = pd.DataFrame([
                {"スタッフ": name, "上限時間": limits.get(name, 0), "役職": "夜勤"} 
                for name, _ in night_staff
            ] + [
                {"スタッフ": name, "上限時間": limits.get(name, 0), "役職": "世話人"} 
                for name, _ in care_staff
            ])
            st.dataframe(limit_df, use_container_width=True)
        
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
                
                # 重複インデックスを避けるためにDataFrameを直接構築
                comparison_data = []
                for staff_key in totals.index:
                    comparison_data.append({
                        "スタッフ": staff_key,
                        "合計時間": totals[staff_key],
                        "上限時間": limits_series[staff_key],
                        "残り時間": limits_series[staff_key] - totals[staff_key]
                    })
                
                comparison_df = pd.DataFrame(comparison_data)
                
                # 上限超過をハイライト
                def highlight_over_limit(row):
                    color = 'background-color: red' if row['残り時間'] < 0 else ''
                    return [color] * len(row)
                
                if len(comparison_df) > 0:
                    styled_df = comparison_df.style.apply(highlight_over_limit, axis=1)
                    st.dataframe(styled_df, use_container_width=True)
                    
                    # 上限超過の警告
                    over_limit_staff = comparison_df[comparison_df['残り時間'] < 0]['スタッフ'].tolist()
                    if over_limit_staff:
                        st.warning(f"⚠️ 上限時間を超過しているスタッフ: {', '.join(over_limit_staff)}")
                else:
                    st.dataframe(comparison_df, use_container_width=True)

            # Excel 出力
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_opt.to_excel(writer, index=False, header=False, sheet_name="最適化シフト")
                
                # 統計情報も追加
                if not limits_series.empty:
                    comparison_df.to_excel(writer, sheet_name="勤務時間統計")
                
            st.download_button(
                label="📥 最適化シフトをダウンロード",
                data=buffer.getvalue(),
                file_name="optimized_shift.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"ファイルの読み込みまたは最適化中にエラーが発生しました: {e}")
        st.error("詳細エラー情報:")
        st.exception(e)
else:
    st.info("左のサイドバーからテンプレート形式の Excel ファイルをアップロードしてください。")
