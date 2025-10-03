import streamlit as st
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import re, io

def highlight_docx(file, terms):
    doc = Document(file)
    patterns = [re.compile(rf"({re.escape(term)})", re.IGNORECASE) for term in terms]

    for para in doc.paragraphs:
        for pattern in patterns:
            if re.search(pattern, para.text):
                inline = para.runs
                for i in range(len(inline)):
                    text = inline[i].text
                    new_text, last_end = [], 0
                    for match in re.finditer(pattern, text):
                        if match.start() > last_end:
                            new_text.append((text[last_end:match.start()], False))
                        new_text.append((match.group(), True))
                        last_end = match.end()
                    if last_end < len(text):
                        new_text.append((text[last_end:], False))
                    if new_text:
                        inline[i].text = ""
                        for t, highlight in new_text:
                            run = para.add_run(t)
                            if highlight:
                                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    output = io.BytesIO()
    doc.save(output)
    return output

# --- Streamlit App ---
st.title("📄 Word Highlighter")
st.write("Upload a Word file, enter words, and download a highlighted copy.")

file = st.file_uploader("Upload a Word document", type="docx")
terms = st.text_input("Enter words to highlight (comma-separated)")

if file and terms:
    highlighted = highlight_docx(file, [t.strip() for t in terms.split(",")])
    st.download_button(
        "⬇️ Download Highlighted File",
        highlighted.getvalue(),
        file_name="highlighted.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
