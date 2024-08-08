import streamlit as st
import io
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from openpyxl import Workbook

# Azure Form Recognizer credentials
endpoint = "https://arun-document-ai.cognitiveservices.azure.com/"
credential = AzureKeyCredential("d3abb1fb970e41d8b7f3330e202f342a")
document_analysis_client = DocumentAnalysisClient(endpoint, credential)
model_id_personal_details = "resume-template"

def process_document(file):
    # Process the uploaded document using the Personal-Details-Model
    document_data = file.read()
    
    poller = document_analysis_client.begin_analyze_document(model_id_personal_details, document_data)
    result = poller.result()
    
    # Extract data
    row_data = [
        result.documents[0].fields.get("Name", {}).value,
        result.documents[0].fields.get("Skills", {}).value,
        result.documents[0].fields.get("Education", {}).value,
        result.documents[0].fields.get("Professional Experience", {}).value
    ]
    
    return row_data

st.title('Personal Details Extraction Tool')

# Use session state to keep track of processed documents
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = []

uploaded_file = st.file_uploader("Choose a file", type=["pdf", "jpg", "jpeg", "png", "tiff"], key="file_uploader")

if uploaded_file is not None:
    st.write(f'{uploaded_file.type} file uploaded successfully.')
    st.write('Processing...')
    
    # Process the uploaded file
    row_data = process_document(uploaded_file)
    
    # Check if this data is not already in the processed_data list
    if row_data not in st.session_state.processed_data:
        st.session_state.processed_data.append(row_data)
    
    st.success('Document processed successfully!')
    
    # Offer options to the user
    option = st.radio("Choose an option:", ("Upload another document", "Generate Excel"))
    
    if option == "Upload another document":
        st.write("Please upload another document using the file uploader above.")
    
    elif option == "Generate Excel":
        # Create Excel workbook and write data
        workbook = Workbook()
        sheet = workbook.active
        
        # Define headers
        headers = ["Name", "Skills", "Education", "Professional Experience"]
        sheet.append(headers)
        
        # Add all processed data to the sheet
        for row in st.session_state.processed_data:
            sheet.append(row)
        
        # Save the workbook to a bytes buffer
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)
        
        st.success('Excel file generated successfully!')
        st.download_button(
            label="Download Excel file",
            data=buffer,
            file_name="extracted_personal_details.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Clear processed data after generating Excel
        st.session_state.processed_data = []

# Display currently processed documents
if st.session_state.processed_data:
    st.write("Currently processed documents:")
    for i, data in enumerate(st.session_state.processed_data, 1):
        st.write(f"{i}. {data[0]}")  # Assuming the first element is the Name
