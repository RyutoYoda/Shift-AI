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
from typing import List, Tuple, Dict
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

# シフト種別
NIGHT_SHIFT = "夜勤"
CARE_SHIFT = "世話人"

# -------------------- 関数群 --------------------

def detect_date_columns(df: pd.DataFrame) -> List[str]:
    """ヘッダーから日付列を推定し、連続する範囲（列名リスト）を返す"""
    date_cols = []
    for col in df.columns:
        header = str(df.at[DATE_HEADER_ROW, col]).strip()
        # 数字かどうかをチェック（日付として1-31の範囲を想定）
        try:
            day = int(float(header))
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


def get_staff_info(df: pd.DataFrame) -> Tuple[List[str], List[str], Dict[str, float]]:
    """スタッフ情報を取得する"""
    night_staff = []
    care_staff = []
    limits = {}
    
    # 夜勤スタッフ（行10-16）- 「夜間支援員」の人たち
    for row in range(9, 16):  # 0-indexedで9-15
        if row < len(df):
            name = str(df.iloc[row, 0]).strip()
            limit_val = df.iloc[row, 2] if pd.notna(df.iloc[row, 2]) else 0
            if name and name != 'nan' and '夜間' in name:
                night_staff.append(name)
                limits[name] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    # 世話人スタッフ（行5-9）- 「世話人」の人たち
    for row in range(4, 9):  # 0-indexedで4-8
        if row < len(df):
            name = str(df.iloc[row, 0]).strip()
            limit_val = df.iloc[row, 2] if pd.notna(df.iloc[row, 2]) else 0
            if name and name != 'nan' and '世話人' in name:
                care_staff.append(name)
                limits[name] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    return night_staff, care_staff, limits


def get_staff_row_mapping(df: pd.DataFrame, night_staff: List[str], care_staff: List[str]) -> Dict[str, Tuple[int, int]]:
    """スタッフ名と対応する行番号のマッピングを作成"""
    staff_rows = {}
    
    # 夜勤スタッフの行マッピング
    night_row_idx = 0
    for row in range(9, 16):  # 夜勤エリア（10-16行目）
        if row < len(df):
            name = str(df.iloc[row, 0]).strip()
            if name and name != 'nan' and '夜間' in name:
                if night_row_idx < len(night_staff):
                    staff_rows[night_staff[night_row_idx]] = (row, NIGHT_ROWS[night_row_idx])
                    night_row_idx += 1
    
    # 世話人スタッフの行マッピング
    care_row_idx = 0
    for row in range(4, 9):  # 世話人エリア（5-9行目）
        if row < len(df):
            name = str(df.iloc[row, 0]).strip()
            if name and name != 'nan' and '世話人' in name:
                if care_row_idx < len(care_staff):
                    staff_rows[care_staff[care_row_idx]] = (row, CARE_ROWS[care_row_idx])
                    care_row_idx += 1
    
    return staff_rows


def can_assign_shift(df: pd.DataFrame, staff_name: str, day_col: str, shift_type: str, 
                    date_cols: List[str], staff_rows: Dict[str, Tuple[int, int]], 
                    assignment_history: Dict[str, List[str]]) -> bool:
    """指定したスタッフが指定日にシフトに入れるかチェック"""
    
    if staff_name not in staff_rows:
        return False
    
    staff_row = staff_rows[staff_name][0]  # 実際の行番号
    current_day_idx = date_cols.index(day_col)
    day_col_idx = df.columns.get_loc(day_col)
    
    # 当日が0（勤務不可）でないかチェック
    current_value = df.iloc[staff_row, day_col_idx]
    if current_value == 0:
        return False
    
    # 共通ルールのチェック
    staff_history = assignment_history.get(staff_name, [])
    
    # 1. 夜勤後は2日空けて世話人勤務可
    if shift_type == CARE_SHIFT:
        for prev_day_idx in range(max(0, current_day_idx - 2), current_day_idx):
            prev_day = date_cols[prev_day_idx]
            if prev_day in staff_history and assignment_history[staff_name][-1] == NIGHT_SHIFT:
                return False
    
    # 2. 連続勤務の制限（同じスタッフが連続して入らないように）
    if current_day_idx > 0:
        prev_day = date_cols[current_day_idx - 1]
        if prev_day in staff_history:
            return False
    
    return True


def count_staff_hours(assignment_history: Dict[str, List[str]], staff_name: str) -> float:
    """スタッフの総勤務時間を計算"""
    if staff_name not in assignment_history:
        return 0
    
    total_hours = 0
    for shift_type in assignment_history[staff_name]:
        if shift_type == NIGHT_SHIFT:
            total_hours += 12.5  # 夜勤時間
        elif shift_type == CARE_SHIFT:
            total_hours += 6     # 世話人時間
    
    return total_hours


def optimize_shifts(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    """シフト最適化ロジック"""
    date_cols = detect_date_columns(df)
    night_staff, care_staff, limits = get_staff_info(df)
    staff_rows = get_staff_row_mapping(df, night_staff, care_staff)
    
    # 割り当て履歴を追跡
    assignment_history = {staff: [] for staff in night_staff + care_staff}
    
    # -------------------- 指定ブロックのクリア --------------------
    def clear_block(rows: List[int]):
        for r in rows:
            if r < len(df):
                for c in date_cols:
                    col_idx = df.columns.get_loc(c)
                    if df.iloc[r, col_idx] != 0:  # 0 は "固定で不可" の意味なので維持
                        df.iloc[r, col_idx] = ""  # 空セル化
    
    clear_block(NIGHT_ROWS)
    clear_block(CARE_ROWS)
    
    # -------------------- シフト割当アルゴリズム --------------------
    # 各日に対してシフトを割り当て
    for day_col in date_cols:
        day_col_idx = df.columns.get_loc(day_col)
        
        # 夜勤シフト割当（1名）
        available_night_staff = []
        for staff in night_staff:
            if can_assign_shift(df, staff, day_col, NIGHT_SHIFT, date_cols, staff_rows, assignment_history):
                current_hours = count_staff_hours(assignment_history, staff)
                if staff in limits and current_hours + 12.5 <= limits[staff]:
                    available_night_staff.append(staff)
        
        if available_night_staff:
            # 勤務時間が少ないスタッフを優先
            available_night_staff.sort(key=lambda s: count_staff_hours(assignment_history, s))
            selected_night_staff = available_night_staff[0]
            
            # 夜勤スタッフの行に名前を入力
            staff_row = staff_rows[selected_night_staff][0]
            df.iloc[staff_row, day_col_idx] = selected_night_staff
            assignment_history[selected_night_staff].append(NIGHT_SHIFT)
        
        # 世話人シフト割当（1名）
        available_care_staff = []
        for staff in care_staff:
            if can_assign_shift(df, staff, day_col, CARE_SHIFT, date_cols, staff_rows, assignment_history):
                current_hours = count_staff_hours(assignment_history, staff)
                if staff in limits and current_hours + 6 <= limits[staff]:
                    available_care_staff.append(staff)
        
        if available_care_staff:
            # 勤務時間が少ないスタッフを優先
            available_care_staff.sort(key=lambda s: count_staff_hours(assignment_history, s))
            selected_care_staff = available_care_staff[0]
            
            # 世話人スタッフの行に名前を入力
            staff_row = staff_rows[selected_care_staff][0]
            df.iloc[staff_row, day_col_idx] = selected_care_staff
            assignment_history[selected_care_staff].append(CARE_SHIFT)
    
    # -------------------- 結果の集計 --------------------
    all_staff = list(set(night_staff + care_staff))
    totals = pd.Series(dtype=float, index=all_staff)
    limits_series = pd.Series(dtype=float, index=all_staff)
    
    for staff in all_staff:
        totals[staff] = count_staff_hours(assignment_history, staff)
        limits_series[staff] = limits.get(staff, 0)
    
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

        **▼ 実装されたルール**
        - 夜勤・世話人は1日に1人ずつ選択
        - 夜勤後は2日空けて世話人勤務可
        - 連続勤務は避ける
        - 0が入っているセルは勤務不可として維持
        - 各スタッフの上限時間を考慮（夜勤12.5h、世話人6h）
        - 勤務時間が少ないスタッフを優先的に割当

        **▼ 処理内容**
        - E5:AI16とE20:AI30の範囲を一旦クリア（0は除く）
        - 各日1名ずつスタッフ名を割り当て
        - 選ばれなかったスタッフのセルは空白になります

        *行・列の位置がテンプレートと異なる場合は、ソースコード冒頭の定数を調整してください。*
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
                st.subheader("検出されたスタッフ情報")
                st.write("**夜勤スタッフ:**", night_staff)
                st.write("**世話人スタッフ:**", care_staff)
            
            with col2:
                st.subheader("上限時間")
                for staff in night_staff + care_staff:
                    st.write(f"{staff}: {limits.get(staff, 0)}時間")
        
        except Exception as e:
            st.warning(f"スタッフ情報の取得でエラー: {e}")
        
        st.subheader("アップロードされたシフト表")
        st.dataframe(df_input, use_container_width=True)

        if st.sidebar.button("🚀 最適化を実行"):
            with st.spinner("シフトを最適化中..."):
                df_opt, totals, limits = optimize_shifts(df_input.copy())
            
            st.success("最適化が完了しました 🎉")

            st.subheader("最適化後のシフト表")
            st.dataframe(df_opt, use_container_width=True)

            if not limits.empty:
                st.subheader("勤務時間の合計と上限")
                comparison_df = pd.DataFrame({
                    "合計時間": totals, 
                    "上限時間": limits,
                    "残り時間": limits - totals
                })
                # 上限超過をハイライト
                def highlight_over_limit(val):
                    return 'background-color: red' if val < 0 else ''
                
                styled_df = comparison_df.style.applymap(highlight_over_limit, subset=['残り時間'])
                st.dataframe(styled_df, use_container_width=True)
                
                # 上限超過の警告
                over_limit_staff = comparison_df[comparison_df['残り時間'] < 0].index.tolist()
                if over_limit_staff:
                    st.warning(f"⚠️ 上限時間を超過しているスタッフ: {', '.join(over_limit_staff)}")

            # Excel 出力
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_opt.to_excel(writer, index=False, header=False, sheet_name="最適化シフト")
                
                # 統計情報も追加
                if not limits.empty:
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
