import streamlit as st
import pandas as pd
from io import BytesIO

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

st.set_page_config(
page_title="Open End Report Generator",
layout="wide"
)

st.title("Open End Comments Report Generator")

uploaded_file = st.file_uploader(
"Upload Excel or CSV",
type=["xlsx", "xls", "csv"]
)

if uploaded_file:

```
# -------------------------
# READ FILE
# -------------------------

try:

    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.success(f"{len(df):,} records loaded")

except Exception as e:
    st.error(f"Unable to read file: {e}")
    st.stop()

# -------------------------
# COVER PAGE DETAILS
# -------------------------

st.subheader("Cover Page")

col1, col2 = st.columns(2)

with col1:

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

with col2:

    segment_name = st.text_input(
        "Brand / Segment",
        value="Surestay and Collections"
    )

    region_name = st.text_input(
        "Region",
        value="North America"
    )

    footnote = st.text_input(
        "Footnote",
        value="*SureStay Responses Highlighted in Blue"
    )

# -------------------------
# QUESTION SELECTION
# -------------------------

st.subheader("Select Open End Questions")

selected_questions = st.multiselect(
    "Choose Open End Columns",
    options=df.columns.tolist()
)

# -------------------------
# GENERATE REPORT
# -------------------------

if st.button("Generate Word Report"):

    if len(selected_questions) == 0:

        st.warning(
            "Please select at least one open end question."
        )

        st.stop()

    doc = Document()

    # ====================================
    # COVER PAGE
    # ====================================

    def center_line(text, bold=False):

        p = doc.add_paragraph()

        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        run = p.add_run(str(text))

        run.bold = bold
        run.font.size = Pt(12)

    center_line(client_name, True)

    doc.add_paragraph()

    center_line(study_name)

    doc.add_paragraph()

    center_line(report_title)

    doc.add_paragraph()

    center_line(segment_name)

    doc.add_paragraph()

    center_line(region_name)

    doc.add_paragraph()
    doc.add_paragraph()

    center_line(footnote)

    doc.add_page_break()

    # ====================================
    # OPEN END REPORT
    # ====================================

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

        # Question Heading

        p = doc.add_paragraph()

        run = p.add_run(question)

        run.bold = True
        run.font.size = Pt(11)

        # Responses

        for response in responses:

            p = doc.add_paragraph()

            run = p.add_run(
                f"➢ {response}"
            )

            run.font.size = Pt(10)

        # Page Break

        doc.add_page_break()

    # ====================================
    # DOWNLOAD
    # ====================================

    buffer = BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    st.success("Report generated successfully.")

    st.download_button(
        label="Download Word Report",
        data=buffer,
        file_name="Open_End_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
```