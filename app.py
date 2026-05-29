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

    st.success(f"Loaded {len(df):,} records")

    st.subheader("Cover Page Details")

    client_name = st.text_input(
        "Client Name",
        value="BWH HOTELS"
    )

    study_name = st.text_input(
        "Study Name",
        value="2025 Member Survey"
    )

    report_title = st.text_input(
        "Report Title",
        value="Open End Comments"
    )

    brand_name = st.text_input(
        "Brand / Segment",
        value="Surestay and Collections"
    )

    region = st.text_input(
        "Region",
        value="North America"
    )

    footnote = st.text_input(
        "Footnote",
        value="*SureStay Responses Highlighted in Blue"
    )

    # Auto detect likely OE columns
    suggested_cols = [
        col for col in df.columns
        if (
            "comment" in col.lower()
            or "open" in col.lower()
            or "specify" in col.lower()
            or col.upper().startswith("Q")
        )
    ]

    st.subheader("Select Open-End Questions")

    selected_questions = st.multiselect(
        "Questions",
        options=df.columns.tolist(),
        default=suggested_cols
    )

    if st.button("Generate Word Report"):

        doc = Document()

        # ----------------------
        # COVER PAGE
        # ----------------------

        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        run = p.add_run(client_name)
        run.bold = True
        run.font.size = Pt(18)

        doc.add_paragraph()

        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run(study_name)

        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run(report_title)

        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run(brand_name)

        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run(region)

        doc.add_paragraph()
        doc.add_paragraph()

        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        run = p.add_run(footnote)
        run.italic = True

        doc.add_page_break()

        # ----------------------
        # OPEN ENDS
        # ----------------------

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

            doc.add_heading(question, level=2)

            for response in responses:

                doc.add_paragraph(
                    response,
                    style="List Bullet"
                )

            doc.add_paragraph()

        # ----------------------
        # DOWNLOAD
        # ----------------------

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("Report Generated Successfully")

        st.download_button(
            label="Download Word Report",
            data=buffer,
            file_name="Open_End_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )