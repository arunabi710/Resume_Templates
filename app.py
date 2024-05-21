import streamlit as st
import io
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from openpyxl import Workbook

# Azure Form Recognizer credentials
endpoint = "https://document-i-testing.cognitiveservices.azure.com/"
credential = AzureKeyCredential("fca7074b4c814fe3a6f6942fb873ff2b")
document_analysis_client = DocumentAnalysisClient(endpoint, credential)
model_id = "Rental-Agreement-Processing"

def process_pdf(file):
    # Process the uploaded PDF
    document = file.read()
    
    poller = document_analysis_client.begin_analyze_document(model_id, document)
    result = poller.result()

    # Create Excel workbook and write data
    workbook = Workbook()
    sheet = workbook.active
    headers = ["Landlord", "Tenant", "Rent", "Rental Agreement Start Date", "Rental Agreement End Date"]
    sheet.append(headers)
    row_data = [
        result.documents[0].fields.get("Landlord", {}).value,
        result.documents[0].fields.get("Tenant", {}).value,
        result.documents[0].fields.get("Rent", {}).value,
        result.documents[0].fields.get("Rental Agreement Start Date", {}).value,
        result.documents[0].fields.get("Rental Agreement End Date", {}).value
    ]
    sheet.append(row_data)

    # Save the workbook to a bytes buffer
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    
    return buffer

st.title('Document Intelligence Tool')
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    st.write('PDF file uploaded successfully.')
    st.write('Processing...')
    excel_buffer = process_pdf(uploaded_file)
    
    if excel_buffer:
        st.success('Excel file generated successfully!')
        st.download_button(
            label="Download Excel file",
            data=excel_buffer,
            file_name="extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
