import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO

st.set_page_config(
    page_title="Open End Report Generator",
    layout="wide"
)

st.title("Open End Report Generator")

uploaded_file = st.file_uploader(
    "Upload Excel / CSV",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.success(f"{len(df):,} records loaded")

    st.subheader("Cover Page")

    client_name = st.text_input(
        "Client Name",
        "BWH HOTELS"
    )

    study_name = st.text_input(
        "Study Name",
        "2025 Member Survey"
    )

    report_title = st.text_input(
        "Report Title",
        "Open End Comments"
    )

    segment_name = st.text_input(
        "Brand / Segment",
        "Surestay and Collections"
    )

    region_name = st.text_input(
        "Region",
        "North America"
    )

    footnote = st.text_input(
        "Footnote",
        "*SureStay Responses Highlighted in Blue"
    )

    st.subheader("Question Mapping")

    oe_column = st.selectbox(
        "Open End Response Column",
        df.columns
    )

    attribute_column = st.selectbox(
        "Attribute / Service Column (optional)",
        ["None"] + list(df.columns)
    )

    question_text = st.text_area(
        "Question Text",
        value="Q7 - Open End Comments"
    )

    if st.button("Generate Report"):

        doc = Document()

        # COVER PAGE

        def center_text(text, bold=False):
            p = doc.add_paragraph()
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(12)

        center_text(client_name, True)

        doc.add_paragraph()

        center_text(study_name)

        doc.add_paragraph()

        center_text(report_title)

        doc.add_paragraph()

        center_text(segment_name)

        doc.add_paragraph()

        center_text(region_name)

        doc.add_paragraph()
        doc.add_paragraph()

        center_text(footnote)

        doc.add_page_break()

        # ATTRIBUTE GROUPING

        if attribute_column != "None":

            grouped = df[
                [attribute_column, oe_column]
            ].copy()

            grouped = grouped.dropna(
                subset=[oe_column]
            )

            grouped[oe_column] = (
                grouped[oe_column]
                .astype(str)
                .str.strip()
            )

            grouped = grouped[
                grouped[oe_column] != ""
            ]

            for attr in grouped[attribute_column].dropna().unique():

                subset = grouped[
                    grouped[attribute_column] == attr
                ]

                p = doc.add_paragraph()

                run = p.add_run(
                    f"{question_text} {attr}"
                )

                run.bold = True
                run.font.size = Pt(11)

                for response in subset[
                    oe_column
                ]:

                    p = doc.add_paragraph()

                    run = p.add_run(
                        f"➢ {response}"
                    )

                    run.font.size = Pt(10)

                doc.add_paragraph()

        else:

            p = doc.add_paragraph()

            run = p.add_run(question_text)

            run.bold = True

            responses = (
                df[oe_column]
                .dropna()
                .astype(str)
                .str.strip()
            )

            responses = responses[
                responses != ""
            ]

            for response in responses:

                p = doc.add_paragraph()

                run = p.add_run(
                    f"➢ {response}"
                )

                run.font.size = Pt(10)

        buffer = BytesIO()

        doc.save(buffer)

        buffer.seek(0)

        st.download_button(
            "Download Word Report",
            buffer,
            file_name="Open_End_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )