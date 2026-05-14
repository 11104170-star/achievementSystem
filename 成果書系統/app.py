import streamlit as st
import pandas as pd
from collections import Counter
from docx import Document
from io import BytesIO
import os

# =========================
# 網頁標題
# =========================

st.set_page_config(
    page_title="成果書生成器",
    page_icon="📄",
    layout="wide"
)

st.title("📄 茶道社成果書生成器")

# =========================
# 上傳檔案
# =========================

template_file = st.file_uploader(
    "上傳 Word 模板",
    type=["docx"]
)

excel_file = st.file_uploader(
    "上傳問卷 Excel / CSV",
    type=["xlsx", "csv"]
)

# =========================
# 活動資訊
# =========================

st.header("活動資訊")

activity_name = st.text_input("活動名稱")
activity_date = st.text_input("活動日期")
activity_place = st.text_input("活動地點")
activity_people = st.text_input("參加人數")
activity_leader = st.text_input("活動負責人")
phone = st.text_input("連絡電話")

# =========================
# 生成按鈕
# =========================

if st.button("生成成果書"):

    if template_file is None:

        st.error("請上傳 Word 模板")
        st.stop()

    if excel_file is None:

        st.error("請上傳問卷")
        st.stop()

    # =========================
    # 讀取 Excel
    # =========================

    ext = os.path.splitext(
        excel_file.name
    )[1].lower()

    try:

        if ext == ".xlsx":

            df = pd.read_excel(
                excel_file
            )

        else:

            encodings = [
                "utf-8-sig",
                "utf-8",
                "cp950",
                "big5"
            ]

            success = False

            for enc in encodings:

                try:

                    excel_file.seek(0)

                    df = pd.read_csv(
                        excel_file,
                        encoding=enc
                    )

                    success = True
                    break

                except:
                    pass

            if not success:

                st.error("CSV 編碼錯誤")
                st.stop()

    except Exception as e:

        st.error("讀取問卷失敗")
        st.exception(e)
        st.stop()

    # =========================
    # 問卷分析
    # =========================

    rows = df.values.tolist()
    headers = list(df.columns)

    total = len(rows)

    score_questions = []
    text_questions = []

    for col in headers:

        values = (
            df[col]
            .dropna()
            .astype(str)
        )

        is_score = True

        for v in values:

            v = v.strip()

            if v not in [
                "1",
                "2",
                "3",
                "4",
                "5"
            ]:

                is_score = False
                break

        if is_score:

            score_questions.append(col)

        else:

            text_questions.append(col)

    # =========================
    # 生成分析文字
    # =========================

    result_text = ""

    result_text += (
        f"填寫回饋表單人數：{total}人\n\n"
    )

    if len(score_questions) > 0:

        result_text += (
            "題目（1-5分，1分最低、5分最高）\n"
        )

        for i, question in enumerate(
            score_questions,
            start=1
        ):

            counter = Counter()

            values = (
                df[question]
                .dropna()
                .astype(str)
            )

            for v in values:

                counter[v.strip()] += 1

            result_text += (
                f"{i}.{question}\n"
            )

            result = []

            for score in sorted(
                counter.keys(),
                reverse=True
            ):

                count = counter[score]

                percent = (
                    count / total * 100
                )

                result.append(
                    f"{count}人{score}分（{percent:.2f}%）"
                )

            result_text += (
                "、".join(result)
            )

            result_text += "\n"

    # =========================
    # 文字題
    # =========================

    start_index = (
        len(score_questions) + 1
    )

    for idx, question in enumerate(
        text_questions,
        start=start_index
    ):

        result_text += (
            f"\n{idx}.{question}\n"
        )

        counter = Counter()

        blank = 0

        values = df[question]

        for v in values:

            if pd.isna(v):

                blank += 1
                continue

            text = str(v).strip()

            if text == "":

                blank += 1

            else:

                counter[text] += 1

        for text, count in counter.items():

            result_text += (
                f"● {text}（{count}）\n"
            )

        result_text += (
            f"● 空白（{blank}）\n"
        )

    # =========================
    # 讀取 Word
    # =========================

    doc = Document(template_file)

    replace_map = {

        "{{活動名稱}}": activity_name,
        "{{活動日期}}": activity_date,
        "{{活動地點}}": activity_place,
        "{{參加人數}}": activity_people,
        "{{活動負責人}}": activity_leader,
        "{{連絡電話}}": phone,
        "{{問卷分析結果}}": result_text,

    }

    # =========================
    # 替換文字
    # =========================

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for para in cell.paragraphs:

                    for key, value in replace_map.items():

                        if key in para.text:

                            para.text = para.text.replace(
                                key,
                                value
                            )

    # =========================
    # 輸出 Word
    # =========================

    output = BytesIO()

    doc.save(output)

    output.seek(0)

    st.success("成果書生成完成！")

    st.download_button(
        label="📥 下載成果書",
        data=output,
        file_name="成果書.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    )