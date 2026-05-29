import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

st.set_page_config(
    page_title="Open End Report Generator",
    layout="wide"
)

st.title("Open End Word Report Generator")

uploaded_file = st.file_uploader(
    "Upload Excel or CSV",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:

    # Read file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.success(f"Loaded {len(df)} records")

    st.subheader("Select Open End Questions")

    selected_questions = st.multiselect(
        "Choose OE Variables",
        options=df.columns
    )

    report_title = st.text_input(
        "Report Title",
        value="Open End Comments Report"
    )

    if st.button("Generate Word Report"):

        doc = Document()

        doc.add_heading(report_title, level=1)

        for question in selected_questions:

            responses = (
                df[question]
                .dropna()
                .astype(str)
                .str.strip()
            )

            responses = responses[
                responses != ""
            ]

            if len(responses) == 0:
                continue

            doc.add_heading(
                f"{question} (n={len(responses)})",
                level=2
            )

            for response in responses:

                doc.add_paragraph(
                    response,
                    style="List Bullet"
                )

            doc.add_paragraph()

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="Download Word Report",
            data=buffer,
            file_name="Open_End_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )