import streamlit as st
import pandas as pd
import docx
from io import BytesIO

# ルール抽出関数
def extract_text_from_docx(file):
    doc = docx.Document(file)
    full_text = "\n".join([para.text for para in doc.paragraphs if para.text.strip() != ""])
    return full_text

# Excel出力関数
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Streamlit アプリ開始
st.title("スマイル シフト自動作成アプリ")

# --- サイドバー：ルールファイルアップロード ---
st.sidebar.header("1. ルールファイルのアップロード (.docx)")
rule_file = st.sidebar.file_uploader("Google Docsから保存した.docxを選択", type=["docx"])

rules_text = ""
if rule_file:
    rules_text = extract_text_from_docx(rule_file)
    st.sidebar.success("ルール読み込み成功！")

# --- メイン：ルール表示 ---
if rules_text:
    with st.expander("📄 読み込んだルール内容", expanded=False):
        st.text_area("ルール", rules_text, height=400)

# --- メイン：Excelアップロード ---
st.header("2. 勤務表テンプレート（Excel）をアップロード")
uploaded_excel = st.file_uploader("勤務表テンプレートをアップロード（.xlsx）", type=["xlsx"])

if uploaded_excel:
    try:
        shift_df = pd.read_excel(uploaded_excel)
        st.success("勤務表読み込み成功！")
        st.dataframe(shift_df)
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

# --- シフト作成ボタン ---
if st.button("3. シフトを自動作成"):
    if not rules_text or not uploaded_excel:
        st.warning("ルールファイルと勤務表を両方アップロードしてください。")
    else:
        # ここに◯を埋めるロジックを追加（仮：全セルに"◯"を入れる）
        filled_df = shift_df.copy()
        for col in filled_df.columns[1:]:
            filled_df[col] = "◯"
        
        st.success("✅ シフト自動作成完了")
        st.dataframe(filled_df)

        excel_bytes = convert_df_to_excel(filled_df)
        st.download_button("📥 シフト表をダウンロード", data=excel_bytes, file_name="generated_shift.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
