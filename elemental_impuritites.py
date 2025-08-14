import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os
from docx import Document

st.set_page_config(page_title="Analysis Request System", layout="wide")

# Initialize session state variables if they don't exist
if 'current_tab' not in st.session_state:
    st.session_state.current_tab = "Request Form"
if 'submitted_requests' not in st.session_state:
    st.session_state.submitted_requests = []

# Create tabs
tabs = ["Request Form", "Request Status"]
st.sidebar.title("Navigation")
selected_tab = st.sidebar.radio("Go to", tabs, index=tabs.index(st.session_state.current_tab))
st.session_state.current_tab = selected_tab

def create_word_document(form_data):
    """Create a Word document with the form data"""
    doc = Document()
    
    # Add title
    doc.add_heading('Inorganic Analysis Request', 0)
    
    # Add BioA-Elemental Analysis section
    doc.add_heading('BioA-Elemental Analysis', level=1)
    p = doc.add_paragraph('(Elemental Analysis Laboratory – Vitry Lavoisier Building L304)')
    doc.add_paragraph('Jean-francois.rameau@sanofi.com / Sylvie.monget@sanofi.com')
    
    # Add Requestor Information section
    doc.add_heading('REQUESTOR INFORMATION', level=1)
    p = doc.add_paragraph()
    p.add_run('Requestor Site: ').bold = True
    p.add_run(form_data['requestor_site'])
    
    p = doc.add_paragraph()
    p.add_run('Requestor Name/Phone/E.mail: ').bold = True
    p.add_run(form_data['requestor_info'])
    
    p = doc.add_paragraph()
    p.add_run('Request Date: ').bold = True
    p.add_run(form_data['request_date'])
    
    # Add Sample Information section
    doc.add_heading('SAMPLE INFORMATION', level=1)
    
    p = doc.add_paragraph()
    p.add_run('PRODUCT Name: ').bold = True
    p.add_run(form_data['product_name'])
    
    p = doc.add_paragraph()
    p.add_run('Actime Code: ').bold = True
    p.add_run(form_data['actime_code'])
    
    p = doc.add_paragraph()
    p.add_run('PRODUCT Form: ').bold = True
    p.add_run(form_data['product_form'])
    
    p = doc.add_paragraph()
    p.add_run('Batch number: ').bold = True
    p.add_run(form_data['batch_number'])
    
    p = doc.add_paragraph()
    p.add_run('Sample quantity: ').bold = True
    p.add_run(form_data['sample_quantity'])
    
    p = doc.add_paragraph()
    p.add_run('Number of vials: ').bold = True
    p.add_run(str(form_data['number_of_vials']))
    
    p = doc.add_paragraph()
    p.add_run('Safety risk: ').bold = True
    p.add_run(form_data['safety_risk'])
    
    p = doc.add_paragraph()
    p.add_run('Shipment conditions: ').bold = True
    p.add_run(form_data['shipment_conditions'])
    
    p = doc.add_paragraph()
    p.add_run('Storage conditions: ').bold = True
    p.add_run(form_data['storage_conditions'])
    
    # Add Analysis Information section
    doc.add_heading('ANALYSIS INFORMATION', level=1)
    
    p = doc.add_paragraph()
    p.add_run('GMP Analysis: ').bold = True
    p.add_run(form_data['gmp_analysis'])
    
    if form_data['gmp_analysis'] == 'Yes':
        p = doc.add_paragraph()
        p.add_run('Purpose: ').bold = True
        p.add_run(form_data['gmp_purpose'])
    
    p = doc.add_paragraph()
    p.add_run('Analysis Type: ').bold = True
    p.add_run(form_data['analysis_type'])
    
    p = doc.add_paragraph()
    p.add_run('Element(s) to be determined: ').bold = True
    p.add_run(form_data['elements_to_determine'])
    
    p = doc.add_paragraph()
    p.add_run('ICHQ3D Analysis: ').bold = True
    p.add_run('Yes' if form_data['ichq3d_analysis'] else 'No')
    
    if form_data['ichq3d_analysis']:
        p = doc.add_paragraph('For ICHQ3D request, documents to be provided:')
        doc.add_paragraph('* Phase 1 and 2: R&D Medicinal product ID Card (SD-000133)', style='List Bullet')
        doc.add_paragraph('* Phase 3: Medicinal Product ID Card (SD-000134) and Risk Assessment (SD-000131)', style='List Bullet')
    
    p = doc.add_paragraph()
    p.add_run('Method reference and/or specification: ').bold = True
    p.add_run(form_data['method_reference'])
    
    # Add Request reference section
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run('(Completed by the BioA/AE Laboratory)').italic = True
    p = doc.add_paragraph()
    p.add_run('Request reference (Steel or iLab): ').bold = True
    p.add_run('_____________________')
    
    # Save to a BytesIO object
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    
    return doc_io

if selected_tab == "Request Form":
    st.title("Inorganic Analysis Request Form")
    
    # Create form
    with st.form("analysis_request_form"):
        st.subheader("Requestor Information")
        col1, col2 = st.columns(2)
        with col1:
            requestor_site = st.text_input("Requestor Site")
        with col2:
            requestor_info = st.text_input("Requestor Name/Phone/Email")
        
        request_date = st.date_input("Request Date", datetime.now())
        
        st.subheader("Sample Information")
        col1, col2 = st.columns(2)
        with col1:
            product_name = st.text_input("Product Name")
        with col2:
            actime_code = st.text_input("Actime Code")
        
        product_form = st.selectbox("Product Form", 
                                   ["Drug Product", "Drug Substance", "Other"])
        
        batch_number = st.text_area("Batch Number(s)", 
                                   help="Provide a list in case of several samples")
        
        col1, col2 = st.columns(2)
        with col1:
            sample_quantity = st.text_input("Sample Quantity (volume or weight)")
        with col2:
            number_of_vials = st.number_input("Number of Vials", min_value=1, step=1)
        
        safety_risk = st.text_area("Safety Risk", 
                                  help="Safety data sheet to be provided by the requestor")
        
        col1, col2 = st.columns(2)
        with col1:
            shipment_conditions = st.text_input("Shipment Conditions")
        with col2:
            storage_conditions = st.text_input("Storage Conditions")
        
        st.subheader("Analysis Information")
        
        col1, col2 = st.columns(2)
        with col1:
            gmp_analysis = st.radio("GMP Analysis", ["Yes", "No"])
        with col2:
            if gmp_analysis == "Yes":
                gmp_purpose = st.radio("Purpose", ["For Release", "For Information"])
            else:
                gmp_purpose = "N/A"
        
        analysis_type = st.radio("Analysis Type", 
                               ["Quantitative Analysis", "Qualitative Analysis (Screening)"])
        
        elements_to_determine = st.text_area("Element(s) to be determined (quantitative analysis)")
        
        ichq3d_analysis = st.checkbox("ICHQ3D Analysis")
        
        if ichq3d_analysis:
            st.info("For ICHQ3D request, documents to be provided:\n"
                   "* Phase 1 and 2: R&D Medicinal product ID Card (SD-000133)\n"
                   "* Phase 3: Medicinal Product ID Card (SD-000134) and Risk Assessment (SD-000131)")
        
        method_reference = st.text_area("Method reference and/or specification to be applied if relevant")
        
        submitted = st.form_submit_button("Generate Document")
        
        if submitted:
            try:
                # Create a dictionary with all the form data
                form_data = {
                    "requestor_site": requestor_site,
                    "requestor_info": requestor_info,
                    "request_date": request_date.strftime("%Y-%m-%d"),
                    "product_name": product_name,
                    "actime_code": actime_code,
                    "product_form": product_form,
                    "batch_number": batch_number,
                    "sample_quantity": sample_quantity,
                    "number_of_vials": number_of_vials,
                    "safety_risk": safety_risk,
                    "shipment_conditions": shipment_conditions,
                    "storage_conditions": storage_conditions,
                    "gmp_analysis": gmp_analysis,
                    "gmp_purpose": gmp_purpose if gmp_analysis == "Yes" else "",
                    "analysis_type": analysis_type,
                    "elements_to_determine": elements_to_determine,
                    "ichq3d_analysis": ichq3d_analysis,
                    "method_reference": method_reference,
                }
                
                # Generate the Word document
                doc_io = create_word_document(form_data)
                
                # Create a download button for the filled template
                st.success("Document generated successfully!")
                
                # Generate a filename with timestamp
                filename = f"Analysis_Request_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                
                # Create download button
                st.download_button(
                    label="Download Filled Template",
                    data=doc_io,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Store in session state for tracking
                st.session_state.submitted_requests.append({
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "product": product_name,
                    "requestor": requestor_info,
                    "status": "Submitted"
                })
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.exception(e)

elif selected_tab == "Request Status":
    st.title("Request Status")
    
    if not st.session_state.submitted_requests:
        st.info("No requests have been submitted yet.")
    else:
        # Convert the list of dictionaries to a DataFrame
        df = pd.DataFrame(st.session_state.submitted_requests)
        
        # Display the DataFrame
        st.dataframe(df)
        
        # Add a button to clear all requests (for testing purposes)
        if st.button("Clear All Requests"):
            st.session_state.submitted_requests = []
            st.experimental_rerun()