import streamlit as st
from docxtpl import DocxTemplate
from docx import Document
import io
import os
from datetime import date

# Page configuration
st.set_page_config(page_title="Agreement Generator", layout="wide")
st.title("📄 Professional Agreement Generator")

# -----------------------------
# 1. INPUT SECTION
# -----------------------------
col1, col2 = st.columns(2)

with col1:
    org_name = st.text_input("Organisation Name")
    doc_date = st.date_input("Agreement Date", value=date.today())
    doc_type = st.text_input("Document Type (e.g. Private Limited)")
    doc_number = st.text_input("GSTIN / Document Number")
    email = st.text_input("Client Email")

with col2:
    address = st.text_area("Registered Address")
    org_sign_name = st.text_input("Organisation Signatory Name")
    aer_sign_name = st.text_input("Aertrip Signatory Name", value="Luvkesh")

designation_options = ["Director", "Partner", "Proprietor", "Vice President - Operations", "Other"]
st.write("---")
c3, c4 = st.columns(2)
with c3:
    org_designation = st.selectbox("Organisation Signatory Designation", designation_options)
    if org_designation == "Other": org_designation = st.text_input("Enter Custom Org Designation")
with c4:
    aer_designation = st.selectbox("Aertrip Signatory Designation", designation_options, index=4)
    if aer_designation == "Other": aer_designation = st.text_input("Enter Custom Aertrip Designation", value="Manager")

# -----------------------------
# 2. ANNEXURE SETTINGS
# -----------------------------
st.write("---")
st.subheader("Annexure Settings")
col5, col6 = st.columns(2)

with col5:
    annexure_a_choice = st.radio("Include Annexure A (Fees)?", ["Yes", "No"], horizontal=True)

with col6:
    annexure_b_choice = st.radio("Include Annexure B (Related Parties)?", ["Yes", "No"], horizontal=True)

party_names = []
if annexure_b_choice == "Yes":
    num_parties = st.number_input("How many related parties?", min_value=1, step=1)
    for i in range(int(num_parties)):
        name = st.text_input(f"Related Party {i+1} Name", key=f"party_{i}")
        if name: party_names.append(name)

# -----------------------------
# 3. GENERATION ENGINE
# -----------------------------
if st.button("🚀 Generate & Download Agreement", type="primary"):
    if not org_name or not doc_number:
        st.error("Please fill Organisation Name and Document Number!")
    else:
        try:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(BASE_DIR, "template.docx")
            annexure_a_file_path = os.path.join(BASE_DIR, "annexure_a.docx") # Fixed Name

            doc = DocxTemplate(template_path)

            # Preamble lines logic [cite: 109, 20]
            ann_a_line = "The implications of Aertrip Fees are outlined in Annexure A" if annexure_a_choice == "Yes" else "" [cite: 109]
            ann_b_line = "and its related entities as mentioned in Annexure B" if annexure_b_choice == "Yes" else "" [cite: 20]

            context = {
                "org_name": org_name, [cite: 13, 266, 269]
                "date": doc_date.strftime('%d %B %Y'), [cite: 18]
                "document_type": doc_type, [cite: 20]
                "document_number": doc_number, [cite: 20]
                "address": address, [cite: 20, 273]
                "email": email, [cite: 272]
                "org_sign_name": org_sign_name, [cite: 270]
                "org_sign_designation": org_designation, [cite: 266, 271]
                "aer_sign_name": aer_sign_name, [cite: 277]
                "aer_sign_designation": aer_designation, [cite: 278]
                "annexure_a_line": ann_a_line, [cite: 109]
                "annexure_b_line": ann_b_line [cite: 20]
            }

            # --- Annexure A Logic ---
            if annexure_a_choice == "Yes" and os.path.exists(annexure_a_file_path):
                sub_doc_a = doc.new_subdoc(annexure_a_file_path)
                context["annexure_a_section"] = sub_doc_a [cite: 1, 262]
            else:
                context["annexure_a_section"] = "" [cite: 1, 262]

            # --- Annexure B Logic ---
            if annexure_b_choice == "Yes" and party_names:
                b_buffer = io.BytesIO()
                temp_b_doc = Document()
                temp_b_doc.add_page_break() # Adds break only if Annexure B exists
                temp_b_doc.add_heading("ANNEXURE B - CLIENT’S ENTITIES", level=1)
                for i, name in enumerate(party_names, 1):
                    temp_b_doc.add_paragraph(f"{i}. {name}")
                temp_b_doc.save(b_buffer)
                b_buffer.seek(0)
                context["annexure_b_section"] = doc.new_subdoc(b_buffer) [cite: 263]
            else:
                context["annexure_b_section"] = "" [cite: 263]

            doc.render(context)

            final_buffer = io.BytesIO()
            doc.save(final_buffer)
            final_buffer.seek(0)

            st.success("✅ Agreement generated!")
            st.download_button(
                label="📥 Download Word File",
                data=final_buffer,
                file_name=f"Agreement_{org_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
