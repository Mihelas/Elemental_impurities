import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import base64
from datetime import datetime

st.set_page_config(page_title="Analysis Request System", layout="wide")

# Initialize session state variables if they don't exist
if 'current_tab' not in st.session_state:
    st.session_state.current_tab = "Request Form"

# Create tabs
tabs = ["Request Form", "Request Status"]
st.sidebar.title("Navigation")
selected_tab = st.sidebar.radio("Go to", tabs, index=tabs.index(st.session_state.current_tab))
st.session_state.current_tab = selected_tab

if selected_tab == "Request Form":
    st.title("Inorganic Analysis Request Form")
    
    # Load the template
    template_path = "Analysis_Request.docx"  # Make sure this file is in the same directory as your app
    
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
                context = {
                    "Requestor_Site": requestor_site,
                    "Requestor_Name_Phone_Email": requestor_info,
                    "Request_Date": request_date.strftime("%Y-%m-%d"),
                    "PRODUCT_Name": product_name,
                    "Actime_Code": actime_code,
                    "PRODUCT_Form": product_form,
                    "Batch_number": batch_number,
                    "Sample_quantity": sample_quantity,
                    "Number_of_vials": number_of_vials,
                    "Safety_risk": safety_risk,
                    "Shipment_conditions": shipment_conditions,
                    "Storage_conditions": storage_conditions,
                    "GMP_Analysis": gmp_analysis,
                    "GMP_Purpose": gmp_purpose if gmp_analysis == "Yes" else "",
                    "Analysis_Type": analysis_type,
                    "Elements_to_determine": elements_to_determine,
                    "ICHQ3D_Analysis": "Yes" if ichq3d_analysis else "No",
                    "Method_reference": method_reference,
                }
                
                # Load the template
                doc = DocxTemplate(template_path)
                
                # Render the template with the context
                doc.render(context)
                
                # Save the document to a BytesIO object
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
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
                if 'submitted_requests' not in st.session_state:
                    st.session_state.submitted_requests = []
                
                st.session_state.submitted_requests.append({
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "product": product_name,
                    "requestor": requestor_info,
                    "status": "Submitted"
                })
                
            except Exception as e:
                st.error(f"An error occurred: {e}")

elif selected_tab == "Request Status":
    st.title("Request Status")
    
    if 'submitted_requests' not in st.session_state or not st.session_state.submitted_requests:
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