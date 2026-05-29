import streamlit as st
import pandas as pd
from io import BytesIO

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

st.set_page_config(
page_title="Open End Report Generator",
layout="wide"
)

st.title("Open End Comments Report Generator")

uploaded_file = st.file_uploader(
"Upload CSV or Excel File",
type=["csv", "xlsx", "xls"]
)

if uploaded_file is not None:

```
# Read File

try:

    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.success(f"{len(df):,} records loaded successfully")

except Exception as e:

    st.error(f"Error reading file: {e}")
    st.stop()

# Cover Page

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

brand_group = st.text_input(
    "Brand Group",
    value="Surestay and Collections"
)

market = st.text_input(
    "Market",
    value="North America"
)

footnote = st.text_input(
    "Footnote",
    value="*SureStay Responses Highlighted in Blue"
)

st.subheader("Question Selection")

# Find likely OE questions

oe_columns = []

for col in df.columns:

    col_upper = str(col).upper()

    if col_upper.startswith("Q7") or col_upper.startswith("Q8"):
        oe_columns.append(col)

if len(oe_columns) == 0:
    oe_columns = list(df.columns)

selected_questions = st.multiselect(
    "Select Open End Questions",
    options=oe_columns
)

highlight_blue = st.checkbox(
    "Highlight SureStay Responses Blue",
    value=True
)

if st.button("Generate Word Report"):

    if len(selected_questions) == 0:

        st.warning(
            "Please select at least one open-end question."
        )

        st.stop()

    doc = Document()

    # Cover Page

    def add_center_line(text, bold=False):

        p = doc.add_paragraph()

        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        run = p.add_run(str(text))

        run.bold = bold
        run.font.name = "Arial"
        run.font.size = Pt(12)

    add_center_line(client_name, True)

    doc.add_paragraph()

    add_center_line(study_name)

    doc.add_paragraph()

    add_center_line(report_title)

    doc.add_paragraph()

    add_center_line(brand_group)

    doc.add_paragraph()

    add_center_line(market)

    doc.add_paragraph()
    doc.add_paragraph()

    add_center_line(footnote)

    doc.add_page_break()

    # SureStay Brands

    surestay_brands = [
        "SSH",
        "SSC",
        "SSPL",
        "SSP"
    ]

    # Report Sections

    for question in selected_questions:

        if question not in df.columns:
            continue

        if "Brand" in df.columns:

            temp = df[
                ["Brand", question]
            ].copy()

        else:

            temp = df[
                [question]
            ].copy()

            temp["Brand"] = ""

        temp = temp.dropna(
            subset=[question]
        )

        temp[question] = (
            temp[question]
            .astype(str)
            .str.strip()
        )

        temp = temp[
            temp[question] != ""
        ]

        if len(temp) == 0:
            continue

        # Question Heading

        p = doc.add_paragraph()

        run = p.add_run(question)

        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(11)

        doc.add_paragraph()

        # Responses

        for _, row in temp.iterrows():

            response = str(row[question])

            brand = str(row["Brand"])

            p = doc.add_paragraph()

            run = p.add_run(
                f"➢ {response}"
            )

            run.font.name = "Arial"
            run.font.size = Pt(10)

            if (
                highlight_blue
                and brand in surestay_brands
            ):
                run.font.color.rgb = RGBColor(
                    0,
                    0,
                    255
                )

        doc.add_page_break()

    # Download

    output = BytesIO()

    doc.save(output)

    output.seek(0)

    st.download_button(
        label="Download Word Report",
        data=output,
        file_name="Open_End_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    st.success(
        "Word report generated successfully."
    )
```