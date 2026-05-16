import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Mm, Cm
from google import genai
import io

st.set_page_config(page_title="Doc Maker", layout="wide")

st.title("📄 AutoScribe")

# --- UI Toggles ---
mode = st.radio("Select Mode", ["Generate New Document", "Clean Existing Document"], horizontal=True)
st.divider()

# Sidebar for Settings
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input("Gemini API Key", type="password", help="Get a free key from aistudio.google.com")
    st.divider()
    
    if mode == "Generate New Document":
        st.subheader("Formatting Overrides")
        st.caption("Use this if your reference document doesn't have clean styles, or if you aren't uploading one.")
        override_formatting = st.checkbox("Override Reference Styles", value=False)
        font_family = "Arial"
        body_size = 12
        h1_size = 18
        selected_color = (0,0,0)

        if override_formatting:
            font_family = st.selectbox("Font Family", ["Arial", "Times New Roman", "Calibri", "Helvetica"])
            body_size = st.number_input("Body Text Size (pt)", min_value=8, max_value=24, value=12)
            h1_size = st.number_input("Heading 1 Size (pt)", min_value=12, max_value=36, value=18)
            color_choice = st.selectbox("Heading Color", ["Black", "Dark Blue", "Dark Red", "Dark Green"])
            color_map = {"Black": (0, 0, 0), "Dark Blue": (0, 0, 139), "Dark Red": (139, 0, 0), "Dark Green": (0, 100, 0)}
            selected_color = color_map[color_choice]
    else:
        st.subheader("Cleaner Settings")
        st.caption("Settings for formatting the messy document.")
        font_family = st.selectbox("Normalize Body Font to:", ["Arial", "Times New Roman", "Calibri", "Helvetica"])
        body_size = st.number_input("Normalize Body Size to (pt):", min_value=8, max_value=24, value=12)


def generate_content(title, outline, api_key):
    if not api_key:
        return f"# {title}\n\n## Introduction\n\nThis is generated dummy text because no API key was provided. In a real scenario, the AI would generate a full introduction about {title} based on your outline.\n\n## Body Section\n\nLorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.\n\n## Conclusion\n\nThis concludes the mock document. If you want real text, paste a free Gemini API key in the sidebar."
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"Write a comprehensive document titled '{title}'.\n\nHere is the outline to follow:\n{outline}\n\nWrite the content section by section. Use Markdown formatting: ## for Headings and standard text for paragraphs. Do not include a title at the very top as it will be added separately."
        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
        return response.text
    except Exception as e:
        st.error(f"Error generating content: {e}")
        return None

def create_document(title, content_md, template_file, apply_overrides):
    if template_file:
        doc = Document(template_file)
        doc.add_page_break()
    else:
        doc = Document()
        # Apply standard margins if no template
        for section in doc.sections:
            section.page_width = Mm(210)
            section.page_height = Mm(297)
            section.top_margin = Cm(2.54)
            section.bottom_margin = Cm(2.54)
            section.left_margin = Cm(2.54)
            section.right_margin = Cm(2.54)
            
    title_para = doc.add_paragraph(title, style='Title')
    title_para.alignment = 1 
    
    lines = content_md.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith('## '):
            p = doc.add_paragraph(line.replace('## ', ''), style='Heading 1')
            if apply_overrides:
                run = p.runs[0] if p.runs else p.add_run(p.text)
                if p.runs and len(p.runs) > 1:
                   run.text = p.text
                   for r in p.runs[1:]: r.text = ''
                run.font.name = font_family
                run.font.size = Pt(h1_size)
                run.font.color.rgb = RGBColor(*selected_color)
        elif line.startswith('### '):
            p = doc.add_paragraph(line.replace('### ', ''), style='Heading 2')
            if apply_overrides:
                run = p.runs[0] if p.runs else p.add_run(p.text)
                if p.runs and len(p.runs) > 1:
                   run.text = p.text
                   for r in p.runs[1:]: r.text = ''
                run.font.name = font_family
                run.font.size = Pt(max(12, h1_size - 2))
                run.font.color.rgb = RGBColor(*selected_color)
        elif line.startswith('# '):
            p = doc.add_paragraph(line.replace('# ', ''), style='Heading 1')
            if apply_overrides:
                run = p.runs[0] if p.runs else p.add_run(p.text)
                if p.runs and len(p.runs) > 1:
                   run.text = p.text
                   for r in p.runs[1:]: r.text = ''
                run.font.name = font_family
                run.font.size = Pt(h1_size)
                run.font.color.rgb = RGBColor(*selected_color)
        else:
            p = doc.add_paragraph(line, style='Normal')
            if apply_overrides:
                if p.runs:
                    for run in p.runs:
                        run.font.name = font_family
                        run.font.size = Pt(body_size)
                else:
                    run = p.add_run(line)
                    p.text = ""
                    run.font.name = font_family
                    run.font.size = Pt(body_size)
    return doc

def clean_document(uploaded_messy_file, font_family, body_size):
    doc = Document(uploaded_messy_file)
    clean_doc = Document()
    
    # 1. Enforce Print-Ready Canvas (A4 & Margins)
    for section in clean_doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # 2. Clean Paragraphs
    blank_count = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            blank_count += 1
            if blank_count > 1:
                continue # Skip excessive blank lines
        else:
            blank_count = 0
            
        try:
            new_p = clean_doc.add_paragraph(text, style=p.style.name)
        except KeyError:
            new_p = clean_doc.add_paragraph(text, style='Normal')
            
        new_p.alignment = p.alignment
        
        # Normalize font for body text
        if new_p.style.name == 'Normal':
            if new_p.runs:
                for run in new_p.runs:
                    run.font.name = font_family
                    run.font.size = Pt(body_size)
            else:
                run = new_p.add_run(text)
                new_p.text = ""
                run.font.name = font_family
                run.font.size = Pt(body_size)
                
    return clean_doc

# --- Main App Logic ---
if mode == "Generate New Document":
    st.write("Upload a reference document, or use the settings panel to generate a perfectly formatted, print-ready document.")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.subheader("Document Content")
        doc_title = st.text_input("Document Title", placeholder="e.g., The History of Artificial Intelligence")
        doc_outline = st.text_area("Outline / Table of Contents", height=150, placeholder="1. Introduction\n2. Early History\n3. Modern Era\n4. Conclusion")
    with col2:
        st.subheader("Reference Template")
        uploaded_file = st.file_uploader("Upload Reference Document (.docx)", type=["docx"], key="gen_up")

    if st.button("Generate Document", type="primary", use_container_width=True):
        if not doc_title:
            st.warning("Please provide a Document Title.")
        else:
            with st.spinner("Generating document content..."):
                markdown_content = generate_content(doc_title, doc_outline, api_key)
            if markdown_content:
                with st.spinner("Applying formatting and creating Word document..."):
                    try:
                        doc = create_document(doc_title, markdown_content, uploaded_file, override_formatting)
                        buffer = io.BytesIO()
                        doc.save(buffer)
                        buffer.seek(0)
                        st.success("Document generated successfully!")
                        st.download_button(
                            label="📥 Download Generated .docx",
                            data=buffer,
                            file_name=f"{doc_title.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    except Exception as e:
                        st.error(f"Error creating document: {e}")

else:
    # Cleaner Mode
    st.write("Upload a messy document. The app will force A4 sizing, 1-inch margins, delete random blank pages, and normalize fonts for printing.")
    st.subheader("Upload Messy Document")
    messy_file = st.file_uploader("Upload Document (.docx)", type=["docx"], key="clean_up")
    
    if st.button(" Clean and Format Document", type="primary", use_container_width=True):
        if not messy_file:
            st.warning("Please upload a document to clean.")
        else:
            with st.spinner("Cleaning document layout..."):
                try:
                    cleaned_doc = clean_document(messy_file, font_family, body_size)
                    buffer = io.BytesIO()
                    cleaned_doc.save(buffer)
                    buffer.seek(0)
                    st.success("Document cleaned successfully! (A4 Size, 1-inch Margins, No blank pages)")
                    st.download_button(
                        label="📥 Download Cleaned .docx",
                        data=buffer,
                        file_name=f"Cleaned_{messy_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"Error cleaning document: {e}")
