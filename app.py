import streamlit as st
from docxtpl import DocxTemplate
from docx import Document
import io
import os
from datetime import date

# Page configuration
st.set_page_config(page_title="Agreement Generator", layout="wide")
st.title("📄 Professional Agreement Generator")
st.markdown("Fill in the details below to generate your formatted agreement.")

# -----------------------------
# 1. INPUT SECTION
# -----------------------------
col1, col2 = st.columns(2)

with col1:
    org_name = st.text_input("Organisation Name", placeholder="e.g. ABC Pvt Ltd")
    doc_date = st.date_input("Agreement Date", value=date.today())
    doc_type = st.text_input("Document Type", placeholder="e.g. Private Limited Company")
    doc_number = st.text_input("GSTIN / Document Number")
    email = st.text_input("Client Email")

with col2:
    address = st.text_area("Registered Address")
    org_sign_name = st.text_input("Organisation Signatory Name")
    aer_sign_name = st.text_input("Aertrip Signatory Name", value="Luvkesh")

# Designation Logic
designation_options = ["Director", "Partner", "Proprietor", "Vice President - Operations", "Other"]

st.write("---")
c3, c4 = st.columns(2)
with c3:
    org_designation = st.selectbox("Organisation Signatory Designation", designation_options)
    if org_designation == "Other":
        org_designation = st.text_input("Enter Custom Org Designation")

with c4:
    aer_designation = st.selectbox("Aertrip Signatory Designation", designation_options, index=4)
    if aer_designation == "Other":
        aer_designation = st.text_input("Enter Custom Aertrip Designation", value="Manager")

# -----------------------------
# 2. ANNEXURE LOGIC
# -----------------------------
st.write("---")
st.subheader("Annexure Settings")
col5, col6 = st.columns(2)

with col5:
    annexure_a = st.radio("Include Annexure A (Fees)?", ["Yes", "No"], horizontal=True)

with col6:
    annexure_b = st.radio("Include Annexure B (Related Parties)?", ["Yes", "No"], horizontal=True)

party_names = []
if annexure_b == "Yes":
    num_parties = st.number_input("How many related parties?", min_value=1, step=1)
    for i in range(int(num_parties)):
        name = st.text_input(f"Related Party {i+1} Name", key=f"party_{i}")
        if name:
            party_names.append(name)

# -----------------------------
# 3. GENERATION ENGINE
# -----------------------------
if st.button("🚀 Generate & Download Agreement", type="primary"):
    if not org_name or not doc_number:
        st.error("Please fill Organisation Name and Document Number!")
    else:
        try:
            # Paths
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(BASE_DIR, "template.docx")
            annex_a_path = os.path.join(BASE_DIR, "annexure_a.docx")

            # Initialize DocxTemplate
            doc = DocxTemplate(template_path)

            # Prepare Conditional Lines for Preamble
            # [cite: 20, 109]
            ann_a_line = "The implications of Aertrip Fees are outlined in Annexure A" if annexure_a == "Yes" else ""
            ann_b_line = "and its related entities as mentioned in Annexure B" if annexure_b == "Yes" else ""

            # Prepare Context (Placeholders in your Word files)
            context = {
                "org_name": org_name,
                "date": doc_date.strftime('%d %B %Y'),
                "document_type": doc_type,
                "document_number": doc_number,
                "address": address,
                "email": email,
                "org_sign_name": org_sign_name,
                "org_sign_designation": org_designation,
                "aer_sign_name": aer_sign_name,
                "aer_sign_designation": aer_designation,
                "annexure_a_line": ann_a_line,
                "annexure_b_line": ann_b_line
            }

            # --- Handle Annexure A Insertion (Style Preserved) ---
            if annexure_a == "Yes":
                if os.path.exists(ann_a_path):
                    # This inserts the full content of annexure_a.docx into the tag
                    sub_doc_a = doc.new_subdoc(ann_a_path)
                    context["annexure_a_section"] = sub_doc_a
                else:
                    st.warning("annexure_a.docx not found!")
                    context["annexure_a_section"] = ""
            else:
                context["annexure_a_section"] = ""

            # --- Handle Annexure B Insertion (Dynamic List) ---
            if annexure_b == "Yes" and party_names:
                # Create a temporary sub-document for Annexure B
                b_buffer = io.BytesIO()
                temp_b_doc = Document()
                temp_b_doc.add_heading("ANNEXURE B - CLIENT’S ENTITIES", level=1)
                for i, name in enumerate(party_names, 1):
                    temp_b_doc.add_paragraph(f"{i}. {name}")
                temp_b_doc.save(b_buffer)
                b_buffer.seek(0)
                
                sub_doc_b = doc.new_subdoc(b_buffer)
                context["annexure_b_section"] = sub_doc_b
            else:
                context["annexure_b_section"] = ""

            # 4. Final Render
            doc.render(context)

            # 5. Save to memory and Download
            final_buffer = io.BytesIO()
            doc.save(final_buffer)
            final_buffer.seek(0)

            st.success("✅ Agreement generated successfully!")
            st.download_button(
                label="📥 Download Word File",
                data=final_buffer,
                file_name=f"Agreement_{org_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
