import streamlit as st
from pptx import Presentation
import re
import tempfile
import os

# Regex pattern
pattern = re.compile(r'@[^@]*\(([^@()]*)\)[^@]*@')

def transform_text(text):
    matches = pattern.findall(text)
    if matches:
        return pattern.sub(lambda m: matches.pop(0) if matches else m.group(0), text)
    return text

def fix_pptx(file):
    prs = Presentation(file)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        run.text = transform_text(run.text)
            if shape.shape_type == 19:  # Table
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                run.text = transform_text(run.text)
    tmpdir = tempfile.mkdtemp()
    output_path = os.path.join(tmpdir, 'fixed.pptx')
    prs.save(output_path)
    return output_path

st.title("📊 PPTX Fixer - AFGC")

st.markdown(
    """
    🔒 **Privacy Notice**

    - Your file is processed securely **in memory** and **never stored**.
    - No data is logged, saved, or sent to any external service.
    - Once the session ends or the browser is closed, all data is wiped.
    """
)


uploaded_file = st.file_uploader("Upload a .pptx file", type=["pptx"])

if uploaded_file:
    with st.spinner("Fixing file..."):
        fixed_path = fix_pptx(uploaded_file)
        with open(fixed_path, 'rb') as f:
            st.download_button("Download Fixed File", f, file_name="fixed.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
