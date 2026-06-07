import streamlit as st
import pandas as pd
import pdfplumber
import re
import base64
from io import BytesIO

st.set_page_config(
    page_title="個人目標評価システム",
    layout="wide"
)

st.title("個人目標評価システム")

uploaded_files = st.file_uploader(
    "評価シートPDFを選択してください",
    type=["pdf"],
    accept_multiple_files=True
)

score_map = {
    "◎": 100,
    "○": 80,
    "〇": 80,
    "△": 50,
    "×": 0
}


def extract_text(pdf_file):
    text = ""

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

    except Exception as e:
        st.error(f"PDF読込エラー: {e}")

    return text


def parse_pdf(text, filename):
    result = {
        "ファイル名": filename,
        "部門": "",
        "目標1評価": "",
        "目標2評価": "",
        "目標1点": 0,
        "目標2点": 0,
        "平均点": 0,

        # "セルフチェック合計": 0,
        # "上司評価合計": 0,

        # "自己評価": 80,
        # "上司評価": 70,
        # "評価差異": 10,
        # "評価ランク": "B"
    }
    if "受付" in filename:
        result["部門"] = "受付"

    elif "培養" in filename:
        result["部門"] = "培養"

    elif "看護" in filename:
        result["部門"] = "看護"

    elif "MA" in filename:
        result["部門"] = "MA"
    
    elif "事務" in filename:
        result["部門"] = "事務"

    elif "看護助手" in filename:
        result["部門"] = "看護助手"

    else:
        result["部門"] = "検査"

    name_match = re.search(
        r"名前[　\s]*([^\n]+)",
        text
    )

    if name_match:
        result["氏名"] = name_match.group(1).strip()

    evaluations = re.findall(
        r"[◎○〇△×]",
        text
    )

    if len(evaluations) >= 2:
        result["目標1評価"] = evaluations[0]
        result["目標2評価"] = evaluations[1]

        result["目標1点"] = score_map.get(evaluations[0], 0)
        result["目標2点"] = score_map.get(evaluations[1], 0)

        result["平均点"] = (
            result["目標1点"] + result["目標2点"]
        ) / 2

    score = result["平均点"]

    if score >= 90:
        result["評価ランク"] = "S"
    elif score >= 80:
        result["評価ランク"] = "A"
    elif score >= 70:
        result["評価ランク"] = "B"
    elif score >= 60:
        result["評価ランク"] = "C"
    else:
        result["評価ランク"] = "D"

    return result


def display_pdf(pdf_bytes):
    base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

    pdf_display = f"""
    <iframe
        src="data:application/pdf;base64,{base64_pdf}"
        width="100%"
        height="1100"
        type="application/pdf">
    </iframe>
    """

    st.markdown(
        pdf_display,
        unsafe_allow_html=True
    )


if uploaded_files:
    results = []
    pdf_dict = {}

    progress = st.progress(0)

    for i, file in enumerate(uploaded_files):
        pdf_bytes = file.getvalue()
        pdf_dict[file.name] = pdf_bytes

        text = extract_text(BytesIO(pdf_bytes))
        data = parse_pdf(
            text,
            file.name
        )

        results.append(data)

        progress.progress(
            (i + 1) / len(uploaded_files)
        )

    df = pd.DataFrame(results)

    # 部門選択

    selected_dept = st.selectbox(
        "部門選択",
        ["全部門"] +
        sorted(df["部門"].unique())
    )

    if selected_dept == "全部門":
        filtered_df = df
    else:
        filtered_df = df[
            df["部門"] == selected_dept
        ]

    tab1, tab2 = st.tabs(
        ["評価一覧", "個人評価"]
    )

    with tab1:
        # st.subheader("評価差異ランキング")

        # diff_df = df[
        #     [
        #         "氏名",
        #         "自己評価",
        #         "上司評価",
        #         "評価差異"
        #     ]
        # ].sort_values(
        #     by="評価差異",
        #     ascending=False
        # )

        # st.dataframe(
        #     diff_df,
        #     use_container_width=True
        # )

        st.subheader("評価一覧")

        display_df = filtered_df.drop(
            columns=["氏名"],
            errors="ignore"
        )

        st.dataframe(
            display_df,
            use_container_width=True,
            height=800
        )
        
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="評価一覧"
            )

        st.download_button(
            "Excelダウンロード",
            data=output.getvalue(),
            file_name="評価一覧.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with tab2:
        st.subheader("個人評価")

        selected_name = st.selectbox(
            "職員選択",
            df["ファイル名"]
        )

        selected_row = df[
            df["ファイル名"] == selected_name
        ].iloc[0]

        with st.sidebar:
            st.subheader("評価情報")

            st.metric(
                "平均点",
                selected_row["平均点"]
            )

            st.metric(
                "評価ランク",
                selected_row["評価ランク"]
            )

            # st.metric(
            #     "自己評価",
            #     selected_row["自己評価"]
            # )

            # st.metric(
            #     "上司評価",
            #     selected_row["上司評価"]
            # )

            # diff = selected_row["評価差異"]

            # if diff > 0:
            #     st.error(f"評価差異 +{diff}")
            # elif diff < 0:
            #     st.warning(f"評価差異 {diff}")
            # else:
            #     st.success("評価差異 0")

            st.write("### 基本情報")
            st.write(f"氏名：{selected_row['氏名']}")
            st.write(f"目標①：{selected_row['目標1評価']}")
            st.write(f"目標②：{selected_row['目標2評価']}")

        st.write("### 評価シート")

        file_name = selected_row["ファイル名"]

        display_pdf(
            pdf_dict[file_name]
        )

else:
    st.info(
        "評価シートPDFをアップロードしてください"
    )