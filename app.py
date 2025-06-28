# --- ライブラリ読み込み ---
import pandas as pd
import random
import docx
import openai
import streamlit as st
from io import BytesIO

# --- GPT API設定（サイドバー） ---
st.sidebar.title("設定")
api_key = st.sidebar.text_input("OpenAI APIキーを入力", type="password")
if not api_key:
    st.warning("APIキーをサイドバーに入力してください")
    st.stop()
client = openai.OpenAI(api_key=api_key)

# --- Wordから勤務ルールを抽出する関数 ---
def extract_text_from_docx(file):
    doc = docx.Document(file)
    return "\n".join([para.text for para in doc.paragraphs if para.text.strip() != ""])

def parse_rules_from_gpt(doc_text):
    system_prompt = """
    以下の勤務ルールから、各スタッフの勤務可能条件をJSON形式で出力してください。
    スタッフ名をキーとして、home（①/②）、世話人可否、夜勤可否、上限時間（単位:時間）、曜日指定、特記事項などをまとめてください。
    """
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": doc_text}
    ]
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=messages,
        temperature=0
    )
    return eval(response.choices[0].message.content)  # JSON文字列がPython dictとして出力される前提

# --- Streamlit UI ---
st.title("すまいるシフト自動作成アプリ（GPT連携版）")

excel_file = st.file_uploader("勤務表（Excel）をアップロード", type=["xlsx"])
docx_file = st.file_uploader("スタッフルール（Word）をアップロード", type=["docx"])

if excel_file and docx_file:
    df = pd.read_excel(excel_file, sheet_name=0, header=None)
    doc_text = extract_text_from_docx(docx_file)

    with st.spinner("GPTでスタッフルール解析中..."):
        staff_rules = parse_rules_from_gpt(doc_text)
        st.success("スタッフルールを解析しました")
        st.json(staff_rules)

    # 初期設定
    start_row = 4
    start_col = 2
    num_days = df.shape[1] - start_col
    roles = df.iloc[start_row:, 0].tolist()
    rows = list(range(start_row, df.shape[0]))

    night_rows = [r for r in rows if '夜間' in str(df.iat[r, 0])]
    day_rows = [r for r in rows if '世話人' in str(df.iat[r, 0])]

    legal_work_limit = 14
    work_count = {r: 0 for r in rows}
    last_night_work_day = {r: -3 for r in night_rows}

    # 日ごとにシフト割当
    for day in range(num_days):
        col = start_col + day

        # 夜間支援員
        candidates = [r for r in night_rows if
                      work_count[r] < legal_work_limit and
                      day - last_night_work_day[r] >= 3 and
                      str(df.iat[r, col]).strip() != '×']
        random.shuffle(candidates)
        for r in candidates:
            if any(str(df.iat[rr, col]).strip() == '◯' for rr in night_rows):
                break
            df.iat[r, col] = '◯'
            work_count[r] += 1
            last_night_work_day[r] = day
            for d_off in [1, 2]:
                if col + d_off < df.shape[1]:
                    df.iat[r, col + d_off] = ''
            break

        # 世話人
        candidates = [r for r in day_rows if
                      work_count[r] < legal_work_limit and
                      str(df.iat[r, col]).strip() != '×']
        random.shuffle(candidates)
        for r in candidates:
            if any(str(df.iat[rr, col]).strip() == '◯' for rr in day_rows):
                break
            df.iat[r, col] = '◯'
            work_count[r] += 1
            break

    # 出力
    st.markdown("### 自動生成されたシフト表")
    st.dataframe(df)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, header=False)
    st.download_button("シフト表をダウンロード", output.getvalue(), file_name="自動生成_シフト表.xlsx")
