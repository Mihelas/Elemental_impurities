import streamlit as st
import pandas as pd
from datetime import datetime
import io
from docx import Document

# Set page config
st.set_page_config(page_title="Analysis Request System", layout="wide")

# Initialize session state
if 'submitted_requests' not in st.session_state:
    st.session_state.submitted_requests = []

# Create tabs
tab1, tab2 = st.tabs(["Request Form", "Request Status"])

# Function to create Word document
def create_word_document(form_data):
    doc = Document()
    doc.add_heading('Inorganic Analysis Request', 0)
    
    # Add sections
    doc.add_paragraph('BioA-Elemental Analysis')
    doc.add_paragraph('(Elemental Analysis Laboratory – Vitry Lavoisier Building L304)')
    doc.add_paragraph('Jean-francois.rameau@sanofi.com / Sylvie.monget@sanofi.com')
    
    # Requestor Information
    doc.add_heading('REQUESTOR INFORMATION', level=1)
    doc.add_paragraph(f"Requestor Site: {form_data['requestor_site']}")
    doc.add_paragraph(f"Requestor Name/Phone/E.mail: {form_data['requestor_info']}")
    doc.add_paragraph(f"Request Date: {form_data['request_date']}")
    
    # Sample Information
    doc.add_heading('SAMPLE INFORMATION', level=1)
    doc.add_paragraph(f"PRODUCT Name: {form_data['product_name']}")
    doc.add_paragraph(f"Actime Code: {form_data['actime_code']}")
    doc.add_paragraph(f"PRODUCT Form: {form_data['product_form']}")
    doc.add_paragraph(f"Batch number: {form_data['batch_number']}")
    doc.add_paragraph(f"Sample quantity: {form_data['sample_quantity']}")
    doc.add_paragraph(f"Number of vials: {form_data['number_of_vials']}")
    doc.add_paragraph(f"Safety risk: {form_data['safety_risk']}")
    doc.add_paragraph(f"Shipment conditions: {form_data['shipment_conditions']}")
    doc.add_paragraph(f"Storage conditions: {form_data['storage_conditions']}")
    
    # Analysis Information
    doc.add_heading('ANALYSIS INFORMATION', level=1)
    doc.add_paragraph(f"GMP Analysis: {form_data['gmp_analysis']}")
    if form_data['gmp_analysis'] == 'Yes':
        doc.add_paragraph(f"Purpose: {form_data['gmp_purpose']}")
    doc.add_paragraph(f"Analysis Type: {form_data['analysis_type']}")
    doc.add_paragraph(f"Element(s) to be determined: {form_data['elements_to_determine']}")
    doc.add_paragraph(f"ICHQ3D Analysis: {'Yes' if form_data['ichq3d_analysis'] else 'No'}")
    doc.add_paragraph(f"Method reference: {form_data['method_reference']}")
    
    # Request reference
    doc.add_paragraph("(Completed by the BioA/AE Laboratory)")
    doc.add_paragraph("Request reference (Steel or iLab): _____________________")
    
    # Save to BytesIO
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# Request Form Tab
with tab1:
    st.title("Inorganic Analysis Request Form")
    
    with st.form("analysis_request_form"):
        st.subheader("Requestor Information")
        requestor_site = st.text_input("Requestor Site")
        requestor_info = st.text_input("Requestor Name/Phone/Email")
        request_date = st.date_input("Request Date", datetime.now())
        
        st.subheader("Sample Information")
        product_name = st.text_input("Product Name")
        actime_code = st.text_input("Actime Code")
        product_form = st.selectbox("Product Form", ["Drug Product", "Drug Substance", "Other"])
        batch_number = st.text_area("Batch Number(s)")
        sample_quantity = st.text_input("Sample Quantity (volume or weight)")
        number_of_vials = st.number_input("Number of Vials", min_value=1, step=1)
        safety_risk = st.text_area("Safety Risk")
        shipment_conditions = st.text_input("Shipment Conditions")
        storage_conditions = st.text_input("Storage Conditions")
        
        st.subheader("Analysis Information")
        gmp_analysis = st.radio("GMP Analysis", ["Yes", "No"])
        gmp_purpose = st.radio("Purpose", ["For Release", "For Information"]) if gmp_analysis == "Yes" else "N/A"
        analysis_type = st.radio("Analysis Type", ["Quantitative Analysis", "Qualitative Analysis (Screening)"])
        elements_to_determine = st.text_area("Element(s) to be determined")
        ichq3d_analysis = st.checkbox("ICHQ3D Analysis")
        method_reference = st.text_area("Method reference")
        
        submitted = st.form_submit_button("Generate Document")
    
    if submitted:
        try:
            # Create form data dictionary
            form_data = {
                "requestor_site": requestor_site,
                "requestor_info": requestor_info,
                "request_date": request_date.strftime("%Y-%m-%d"),
                "product_name": product_name,
                "actime_code": actime_code,
                "product_form": product_form,
                "batch_number": batch_number,
                "sample_quantity": sample_quantity,
                "number_of_vials": str(number_of_vials),
                "safety_risk": safety_risk,
                "shipment_conditions": shipment_conditions,
                "storage_conditions": storage_conditions,
                "gmp_analysis": gmp_analysis,
                "gmp_purpose": gmp_purpose,
                "analysis_type": analysis_type,
                "elements_to_determine": elements_to_determine,
                "ichq3d_analysis": ichq3d_analysis,
                "method_reference": method_reference,
            }
            
            # Generate document
            doc_io = create_word_document(form_data)
            
            # Success message
            st.success("Document generated successfully!")
            
            # Download button (outside the form)
            filename = f"Analysis_Request_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            st.download_button(
                label="Download Filled Template",
                data=doc_io,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Store in session state
            st.session_state.submitted_requests.append({
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "product": product_name,
                "requestor": requestor_info,
                "status": "Submitted"
            })
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

# Request Status Tab
with tab2:
    st.title("Request Status")
    
    if not st.session_state.submitted_requests:
        st.info("No requests have been submitted yet.")
    else:
        df = pd.DataFrame(st.session_state.submitted_requests)
        st.dataframe(df)
        
        if st.button("Clear All Requests"):
            st.session_state.submitted_requests = []
            st.experimental_rerun()