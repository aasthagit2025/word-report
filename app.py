import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# --- Helper Functions for Word Styling ---
def set_cell_background(cell, color_hex):
    """Fills a table cell background with a specific HEX color."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def process_survey_to_docx(csv_file):
    """Processes the uploaded CSV and returns a styled Word document in memory."""
    # Read the uploaded CSV file
    df = pd.read_csv(csv_file)
    
    # Initialize the Word document
    doc = Document()
    
    # Set standard 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # 1. Add Header Content
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run1 = title_p.add_run("BWH Hotels\n")
    run1.font.name = 'Arial'
    run1.font.size = Pt(16)
    run1.font.bold = True
    
    run2 = title_p.add_run("2025 Member Survey, Open End Comments\n")
    run2.font.name = 'Arial'
    run2.font.size = Pt(14)
    run2.font.bold = True
    
    run3 = title_p.add_run("Surestay and Collections, North America")
    run3.font.name = 'Arial'
    run3.font.size = Pt(12)
    run3.font.italic = True
    
    # Note tag
    note_p = doc.add_paragraph()
    run_note = note_p.add_run("\n*SureStay Responses Highlighted in Blue")
    run_note.font.name = 'Arial'
    run_note.font.size = Pt(10)
    run_note.font.italic = True
    
    # 2. Identify Target Columns dynamically
    q7_columns = [col for col in df.columns if col.startswith("Q7 -")]
    q8_column = "Q8 - Is there anything else you would like us to know? Any additional thoughts you would like to share are appreciated."
    target_questions = q7_columns + ([q8_column] if q8_column in df.columns else [])
    
    # 3. Extract and build Word tables
    for q_full_name in target_questions:
        valid_responses = df[[q_full_name, 'Brand']].dropna(subset=[q_full_name])
        valid_responses = valid_responses[valid_responses[q_full_name].astype(str).str.strip() != ""]
        
        if valid_responses.empty:
            continue
            
        doc.add_paragraph("\n") 
        q_p = doc.add_paragraph()
        q_run = q_p.add_run(q_full_name)
        q_run.font.name = 'Arial'
        q_run.font.size = Pt(11)
        q_run.font.bold = True
        
        # Create single-column table box
        table = doc.add_table(rows=0, cols=1)
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.autofit = False
        
        # Subtle horizontal dividers
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            '<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
            '</w:tblBorders>'
        )
        tblPr.append(borders)

        for _, row in valid_responses.iterrows():
            text_comment = str(row[q_full_name]).strip()
            brand_type = str(row['Brand']).upper()
            
            row_cells = table.add_row().cells
            cell = row_cells[0]
            cell.width = Inches(6.5)
            
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15
            
            text_run = p.add_run(text_comment)
            text_run.font.name = 'Arial'
            text_run.font.size = Pt(10)
            
            # Match Brand format (SureStay highlight logic)
            if brand_type.startswith("SS"):
                set_cell_background(cell, "E6F2FF") 
                
    # Save document to a BytesIO byte stream instead of a hard-drive path
    docx_buffer = io.BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer

# --- Streamlit Web Interface UI ---
st.set_page_config(page_title="Survey Response Converter", page_icon="📝", layout="centered")

st.title("📝 Open-End Survey Comment Extractor")
st.write("Upload your raw survey CSV file below to generate a formatted Word document formatted with standard spacing and brand highlights.")

uploaded_file = st.file_uploader("Choose a CSV file", type=["csv"])

if uploaded_file is not None:
    st.success("File uploaded successfully!")
    
    # Process button
    if st.button("Generate Word Document", type="primary"):
        with st.spinner("Parsing data and rendering layout..."):
            try:
                # Run background generation
                output_buffer = process_survey_to_docx(uploaded_file)
                
                st.balloons()
                st.success("Your Word Document is ready!")
                
                # Expose file download interface to user browser
                st.download_button(
                    label="📥 Download Word Document (.docx)",
                    data=output_buffer,
                    file_name="BWH_2025_Member_Survey_Open_End_Comments.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"An error occurred while compiling the document: {e}")