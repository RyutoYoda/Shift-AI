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
    
    # 夜勤スタッフ（行10-16）
    for row in range(9, 16):  # 0-indexedで9-15
        if row < len(df):
            name = str(df.iloc[row, 0]).strip()
            limit_val = df.iloc[row, 2] if pd.notna(df.iloc[row, 2]) else 0
            if name and name != 'nan':
                night_staff.append(name)
                limits[name] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    # 世話人スタッフ（行5-9）
    for row in range(4, 9):  # 0-indexedで4-8
        if row < len(df):
            name = str(df.iloc[row, 0]).strip()
            limit_val = df.iloc[row, 2] if pd.notna(df.iloc[row, 2]) else 0
            if name and name != 'nan':
                care_staff.append(name)
                limits[name] = float(limit_val) if isinstance(limit_val, (int, float)) else 0
    
    return night_staff, care_staff, limits


def can_assign_shift(df: pd.DataFrame, staff_name: str, day_col: str, shift_type: str, 
                    date_cols: List[str], night_staff: List[str], care_staff: List[str]) -> bool:
    """指定したスタッフが指定日にシフトに入れるかチェック"""
    
    current_day_idx = date_cols.index(day_col)
    
    # スタッフの行を特定
    if shift_type == NIGHT_SHIFT:
        staff_rows = NIGHT_ROWS
        staff_list = night_staff
    else:
        staff_rows = CARE_ROWS
        staff_list = care_staff
    
    if staff_name not in staff_list:
        return False
    
    staff_idx = staff_list.index(staff_name)
    staff_row = staff_rows[staff_idx]
    
    # 当日が0（勤務不可）でないかチェック
    current_value = df.iloc[staff_row, df.columns.get_loc(day_col)]
    if current_value == 0:
        return False
    
    # 共通ルールのチェック
    # 1. 夜勤後は2日空けて世話人勤務可
    # 2. 世話人から翌日夜勤入り可能
    
    # 前日と前々日をチェック
    for offset in [1, 2]:
        if current_day_idx - offset >= 0:
            prev_day_col = date_cols[current_day_idx - offset]
            prev_day_idx = df.columns.get_loc(prev_day_col)
            
            # 前日・前々日に他のシフトに入っているかチェック
            if shift_type == CARE_SHIFT:
                # 世話人の場合、夜勤後2日空ける必要がある
                for night_row in NIGHT_ROWS:
                    if night_row < len(df) and df.iloc[night_row, prev_day_idx] == staff_name:
                        if offset <= 2:  # 2日以内
                            return False
            
            # 同日に他のシフトに既に入っているかチェック
            current_day_idx_in_df = df.columns.get_loc(day_col)
            
            if shift_type == NIGHT_SHIFT:
                # 夜勤の場合、同日の世話人シフトをチェック
                for care_row in CARE_ROWS:
                    if care_row < len(df) and df.iloc[care_row, current_day_idx_in_df] == staff_name:
                        return False
            else:
                # 世話人の場合、同日の夜勤シフトをチェック
                for night_row in NIGHT_ROWS:
                    if night_row < len(df) and df.iloc[night_row, current_day_idx_in_df] == staff_name:
                        return False
    
    return True


def count_staff_hours(df: pd.DataFrame, staff_name: str, date_cols: List[str], 
                     night_staff: List[str], care_staff: List[str]) -> float:
    """スタッフの総勤務時間を計算"""
    total_hours = 0
    
    # 夜勤時間をカウント
    if staff_name in night_staff:
        staff_idx = night_staff.index(staff_name)
        staff_row = NIGHT_ROWS[staff_idx]
        for col in date_cols:
            col_idx = df.columns.get_loc(col)
            if df.iloc[staff_row, col_idx] == staff_name:
                # 夜勤は通常12時間と仮定
                total_hours += 12
    
    # 世話人時間をカウント
    if staff_name in care_staff:
        staff_idx = care_staff.index(staff_name)
        staff_row = CARE_ROWS[staff_idx]
        for col in date_cols:
            col_idx = df.columns.get_loc(col)
            if df.iloc[staff_row, col_idx] == staff_name:
                # 世話人は通常8時間と仮定
                total_hours += 8
    
    return total_hours


def optimize_shifts(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series, pd.Series]:
    """シフト最適化ロジック"""
    date_cols = detect_date_columns(df)
    night_staff, care_staff, limits = get_staff_info(df)
    
    # -------------------- 指定ブロックのクリア --------------------
    def clear_block(rows: List[int]):
        for r in rows:
            if r < len(df):
                for c in date_cols:
                    col_idx = df.columns.get_loc(c)
                    if df.iloc[r, col_idx] != 0:  # 0 は "固定で不可" の意味なので維持
                        df.iloc[r, col_idx] = np.nan  # 空セル化

    clear_block(NIGHT_ROWS)
    clear_block(CARE_ROWS)
    
    # -------------------- シフト割当アルゴリズム --------------------
    # 各日に対してシフトを割り当て
    for day_col in date_cols:
        day_idx = df.columns.get_loc(day_col)
        
        # 夜勤シフト割当（1名）
        available_night_staff = []
        for staff in night_staff:
            if can_assign_shift(df, staff, day_col, NIGHT_SHIFT, date_cols, night_staff, care_staff):
                current_hours = count_staff_hours(df, staff, date_cols[:date_cols.index(day_col)], 
                                                night_staff, care_staff)
                if staff in limits and current_hours + 12 <= limits[staff]:
                    available_night_staff.append(staff)
        
        if available_night_staff:
            # 勤務時間が少ないスタッフを優先
            available_night_staff.sort(key=lambda s: count_staff_hours(df, s, date_cols[:date_cols.index(day_col)], 
                                                                       night_staff, care_staff))
            selected_night_staff = available_night_staff[0]
            
            # 夜勤スタッフの行に名前を入力
            staff_idx = night_staff.index(selected_night_staff)
            staff_row = NIGHT_ROWS[staff_idx]
            df.iloc[staff_row, day_idx] = selected_night_staff
        
        # 世話人シフト割当（1名）
        available_care_staff = []
        for staff in care_staff:
            if can_assign_shift(df, staff, day_col, CARE_SHIFT, date_cols, night_staff, care_staff):
                current_hours = count_staff_hours(df, staff, date_cols[:date_cols.index(day_col)+1], 
                                                night_staff, care_staff)
                if staff in limits and current_hours + 8 <= limits[staff]:
                    available_care_staff.append(staff)
        
        if available_care_staff:
            # 勤務時間が少ないスタッフを優先
            available_care_staff.sort(key=lambda s: count_staff_hours(df, s, date_cols[:date_cols.index(day_col)], 
                                                                     night_staff, care_staff))
            selected_care_staff = available_care_staff[0]
            
            # 世話人スタッフの行に名前を入力
            staff_idx = care_staff.index(selected_care_staff)
            staff_row = CARE_ROWS[staff_idx]
            df.iloc[staff_row, day_idx] = selected_care_staff
    
    # -------------------- 結果の集計 --------------------
    all_staff = list(set(night_staff + care_staff))
    totals = pd.Series(dtype=float, index=all_staff)
    limits_series = pd.Series(dtype=float, index=all_staff)
    
    for staff in all_staff:
        totals[staff] = count_staff_hours(df, staff, date_cols, night_staff, care_staff)
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
        - 夜勤・世話人は1日に1人ずつ
        - 夜勤後は2日空けて世話人勤務可
        - 世話人から翌日夜勤入り可能
        - 0が入っているセルは勤務不可として維持
        - 各スタッフの上限時間を考慮
        - 勤務時間が少ないスタッフを優先的に割当

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
        st.error("詳細エラー情報:", e)
else:
    st.info("左のサイドバーからテンプレート形式の Excel ファイルをアップロードしてください。")
