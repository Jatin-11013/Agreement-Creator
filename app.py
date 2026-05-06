from docxtpl import DocxTemplate
import streamlit as st
import io

doc = DocxTemplate("template.docx")

context = {
    "org_name": org_name,
    "date": doc_date,
    "document_type": doc_type,
    "document_number": doc_number,
    "address": address,
    "email": email,
    "org_sign_name": org_sign_name,
    "org_sign_designation": org_designation,
    "aer_sign_name": aer_sign_name,
    "aer_sign_designation": aer_designation,
    "annexure_a": annexure_a == "Yes",
    "annexure_b": annexure_b == "Yes",
    "party_names": party_names
}

doc.render(context)

buffer = io.BytesIO()
doc.save(buffer)
buffer.seek(0)

st.download_button("Download Agreement", buffer, "Agreement.docx")
