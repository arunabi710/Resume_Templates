import streamlit as st
import os
import tempfile
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from openpyxl import Workbook

# Azure Form Recognizer credentials
endpoint = "https://document-i-testing.cognitiveservices.azure.com/"
credential = AzureKeyCredential("your-form-recognizer-api-key")
document_analysis_client = DocumentAnalysisClient(endpoint, credential)
model_id = "Rental-Agreement-Processing"

def process_pdf(file):
    # Save the uploaded PDF to a temporary location
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        temp_file.write(file.read())
        temp_file_path = temp_file.name

    # Process the temporary PDF file
    with open(temp_file_path, "rb") as fd:
        document = fd.read()
    
    try:
        poller = document_analysis_client.begin_analyze_document(model_id, document)
        result = poller.result()
    except Exception as e:
        st.error(f"An error occurred: {e}")
        os.unlink(temp_file_path)
        return None

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

    # Save Excel file
    excel_file_path = 'extracted_data.xlsx'
    workbook.save(excel_file_path)

    # Remove the temporary file
    os.unlink(temp_file_path)

    return excel_file_path

st.title('PDF to Excel Converter')
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    st.write('PDF file uploaded successfully.')
    st.write('Processing...')
    excel_file_path = process_pdf(uploaded_file)
    if excel_file_path:
        st.success('Excel file generated successfully!')

        # Add a download button for the Excel file
        with open(excel_file_path, "rb") as excel_file:
            excel_bytes = excel_file.read()
        st.download_button(label="Download Excel file", data=excel_bytes, file_name="extracted_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
