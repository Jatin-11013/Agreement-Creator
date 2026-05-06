import streamlit as st
from docx import Document
from datetime import date
import io
import os

# -----------------------------
# HELPER FUNCTION (KEY FIX)
# -----------------------------
def replace_text_in_paragraph(paragraph, replacements):
    full_text = "".join(run.text for run in paragraph.runs)

    for key, value in replacements.items():
        if key in full_text:
            full_text = full_text.replace(key, value if value else "")

    for run in paragraph.runs:
        run.text = ""

    if paragraph.runs:
        paragraph.runs[0].text = full_text


# -----------------------------
# UI
# -----------------------------
st.set_page_config(layout="wide")
st.title("Agreement Generator")

col1, col2 = st.columns(2)

with col1:
    org_name = st.text_input("Organisation Name")
    doc_date = st.date_input("Date", value=date.today())
    doc_type = st.text_input("Document Type")
    doc_number = st.text_input("Document Number")
    email = st.text_input("Email")

with col2:
    address = st.text_area("Address")
    org_sign_name = st.text_input("Organisation Signing Name")
    aer_sign_name = st.text_input("Aertrip Signing Name")

designation_options = [
    "Director", "Partner", "Managing Partner",
    "Proprietor", "Vice President - Operations", "Other"
]

org_designation = st.selectbox("Organisation Signing Designation", designation_options)
if org_designation == "Other":
    org_designation = st.text_input("Enter Organisation Designation")

aer_designation = st.selectbox("Aertrip Signing Designation", designation_options)
if aer_designation == "Other":
    aer_designation = st.text_input("Enter Aertrip Designation")

col3, col4 = st.columns(2)

with col3:
    annexure_a = st.selectbox("Annexure A: Fees", ["Yes", "No"])

with col4:
    annexure_b = st.selectbox("Annexure B: Related Parties", ["Yes", "No"])

party_names = []

if annexure_b == "Yes":
    num_parties = st.number_input("Total Number of Parties", min_value=1, step=1)
    for i in range(int(num_parties)):
        name = st.text_input(f"Related Party {i+1}", key=f"party_{i}")
        party_names.append(name)

# -----------------------------
# GENERATE
# -----------------------------
if st.button("Generate Agreement"):

    try:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(BASE_DIR, "template.docx")
        annex_path = os.path.join(BASE_DIR, "annexure_a.docx")

        doc = Document(template_path)

        # -----------------------------
        # MAIN REPLACEMENTS
        # -----------------------------
        replacements = {
            "{{org_name}}": org_name,
            "{{date}}": str(doc_date),
            "{{document_type}}": doc_type,
            "{{document_number}}": doc_number,
            "{{address}}": address,
            "{{email}}": email,
            "{{org_sign_name}}": org_sign_name,
            "{{org_sign_designation}}": org_designation,
            "{{aer_sign_name}}": aer_sign_name,
            "{{aer_sign_designation}}": aer_designation,
        }

        for para in doc.paragraphs:
            replace_text_in_paragraph(para, replacements)

        # -----------------------------
        # ANNEXURE LINES
        # -----------------------------
        for para in doc.paragraphs:

            if "{{ANNEXURE_A_LINE}}" in para.text:
                replace_text_in_paragraph(para, {
                    "{{ANNEXURE_A_LINE}}":
                    "The implications of Aertrip Fees are outlined in Annexure A"
                    if annexure_a == "Yes" else ""
                })

            if "{{ANNEXURE_B_LINE}}" in para.text:
                replace_text_in_paragraph(para, {
                    "{{ANNEXURE_B_LINE}}":
                    "and its related entities as mentioned in Annexure B"
                    if annexure_b == "Yes" else ""
                })

        # -----------------------------
        # ANNEXURE A INSERT
        # -----------------------------
        for para in doc.paragraphs:
            if "{{ANNEXURE_A_SECTION}}" in para.text:

                parent = para._element.getparent()
                index = parent.index(para._element)

                if annexure_a == "Yes":
                    annex_doc = Document(annex_path)

                    for element in annex_doc.element.body:
                        index += 1
                        parent.insert(index, element)

                parent.remove(para._element)
                break

        # -----------------------------
        # ANNEXURE B INSERT
        # -----------------------------
        from docx.oxml import OxmlElement
        from docx.text.paragraph import Paragraph

        for para in doc.paragraphs:
            if "{{ANNEXURE_B_SECTION}}" in para.text:

                parent = para._element.getparent()
                index = parent.index(para._element)

                if annexure_b == "Yes":

                    # Heading
                    new_p = OxmlElement("w:p")
                    parent.insert(index + 1, new_p)
                    p = Paragraph(new_p, doc)
                    p.add_run("ANNEXURE B - CLIENT’S ENTITIES")

                    index += 1

                    # List
                    for i, name in enumerate(party_names, 1):
                        if name:
                            new_p = OxmlElement("w:p")
                            parent.insert(index + 1, new_p)
                            p = Paragraph(new_p, doc)
                            p.add_run(f"{i}. {name}")
                            index += 1

                parent.remove(para._element)
                break

        # -----------------------------
        # SAVE
        # -----------------------------
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            "Download Agreement",
            buffer,
            file_name=f"{org_name}_Agreement.docx"
        )

        st.success("✅ Agreement Generated Successfully")

    except Exception as e:
        st.error(f"❌ Error: {e}")
