import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_COLOR_INDEX
import io

# Function to add text with formatting to Word
def add_formatted_paragraph(doc, text):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    # Set text to bold
    run.bold = True

    # highlight text, code from Gemini
    font = run.font
    font.highlight_color = WD_COLOR_INDEX.YELLOW

# Function to process input text and generate a Word document
def process_text(input_text):
    doc = Document()
    for line in input_text.strip().split("\n"):
        if line.startswith("*"):
            add_formatted_paragraph(doc, line[1:].strip())  # Remove the asterisk
        else:
            doc.add_paragraph(line)
    return doc

# Streamlit app
st.title("Text Formatter and Word File Generator")
st.write("Upload a Word or text file. Lines starting with an asterisk (*) will be bolded and highlighted in yellow.")

# File upload
uploaded_file = st.file_uploader("Upload your file", type=["txt", "docx"])
filename, ext = str(uploaded_file.name).split('.')

if uploaded_file:
    if uploaded_file.name.endswith(".txt"):
        # Process text file
        input_text = uploaded_file.read().decode("utf-8")
    elif uploaded_file.name.endswith(".docx"):
        # Process Word file
        doc = Document(uploaded_file)
        input_text = "\n".join([para.text for para in doc.paragraphs])
    else:
        input_text = ""

    if input_text:
        # Process the text and generate the Word document
        formatted_doc = process_text(input_text)
        # Save the Word file in memory
        output = io.BytesIO()
        formatted_doc.save(output)
        output.seek(0)
        # Provide a download link
        st.download_button(label="Download Formatted Word File", data=output, file_name="{filename}_formatted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.warning("The uploaded file is empty. Please try again.")

