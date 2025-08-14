import streamlit as st
import pandas as pd
from datetime import datetime
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Set page config
st.set_page_config(page_title="Analysis Request System", layout="wide")

# Initialize session state
if 'submitted_requests' not in st.session_state:
    st.session_state.submitted_requests = []

# Create tabs
tab1, tab2 = st.tabs(["Request Form", "Request Status"])

# Predefined elements table
elements_table = {
    "Cd": {"Class": "1", "If intentionally added": True, "If not intentionally added": True},
    "Pb": {"Class": "1", "If intentionally added": True, "If not intentionally added": True},
    "As": {"Class": "1", "If intentionally added": True, "If not intentionally added": True},
    "Hg": {"Class": "1", "If intentionally added": True, "If not intentionally added": True},
    "Co": {"Class": "2A", "If intentionally added": True, "If not intentionally added": True},
    "V": {"Class": "2A", "If intentionally added": True, "If not intentionally added": True},
    "Ni": {"Class": "2A", "If intentionally added": True, "If not intentionally added": True},
    "Tl": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Au": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Pd": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Ir": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Os": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Rh": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Ru": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Se": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Ag": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Pt": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False},
    "Li": {"Class": "3", "If intentionally added": True, "If not intentionally added": True},
    "Sb": {"Class": "3", "If intentionally added": True, "If not intentionally added": True},
    "Ba": {"Class": "3", "If intentionally added": True, "If not intentionally added": False},
    "Mo": {"Class": "3", "If intentionally added": True, "If not intentionally added": False},
    "Cu": {"Class": "3", "If intentionally added": True, "If not intentionally added": True},
    "Sn": {"Class": "3", "If intentionally added": True, "If not intentionally added": False},
    "Cr": {"Class": "3", "If intentionally added": True, "If not intentionally added": False},
}

# Function to create Word document
def create_word_document(form_data):
    doc = Document()
    
    # Set margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    
    # Title
    title = doc.add_heading('Inorganic Analysis Request', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # BioA-Elemental Analysis section
    subtitle = doc.add_heading('BioA-Elemental Analysis', level=1)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph('(Elemental Analysis Laboratory – Vitry Lavoisier Building L304)')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph('Jean-francois.rameau@sanofi.com / Sylvie.monget@sanofi.com')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add some space
    doc.add_paragraph()
    
    # Requestor Information section
    doc.add_heading('REQUESTOR INFORMATION', level=1)
    
    # Create a table for requestor info
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Table Grid'
    
    # Row 1
    cell = table.cell(0, 0)
    cell.text = "Requestor Site:"
    cell = table.cell(0, 1)
    cell.text = form_data['requestor_site']
    
    # Row 2
    cell = table.cell(1, 0)
    cell.text = "Requestor Name/Phone/E.mail:"
    cell = table.cell(1, 1)
    cell.text = f"{form_data['requestor_name']} / {form_data['requestor_phone']} / {form_data['requestor_email']}"
    
    # Row 3
    cell = table.cell(2, 0)
    cell.text = "Request Date:"
    cell = table.cell(2, 1)
    cell.text = form_data['request_date']
    
    # Add some space
    doc.add_paragraph()
    
    # Sample Information section
    doc.add_heading('SAMPLE INFORMATION', level=1)
    
    # Create a table for sample info
    table = doc.add_table(rows=9, cols=2)
    table.style = 'Table Grid'
    
    # Row 1
    cell = table.cell(0, 0)
    cell.text = "PRODUCT Name:"
    cell = table.cell(0, 1)
    cell.text = form_data['product_name']
    
    # Row 2
    cell = table.cell(1, 0)
    cell.text = "Actime Code:"
    cell = table.cell(1, 1)
    cell.text = form_data['actime_code']
    
    # Row 3
    cell = table.cell(2, 0)
    cell.text = "PRODUCT Form (Drug Product, Drug substance, other):"
    cell = table.cell(2, 1)
    cell.text = form_data['product_form']
    
    # Row 4
    cell = table.cell(3, 0)
    cell.text = "Batch number (provide a list in attachment in case of several samples):"
    cell = table.cell(3, 1)
    cell.text = form_data['batch_number']
    
    # Row 5
    cell = table.cell(4, 0)
    cell.text = "Sample quantity (volume or weight):"
    cell = table.cell(4, 1)
    cell.text = f"{form_data['sample_quantity']} {form_data['sample_unit']}"
    
    # Row 6
    cell = table.cell(5, 0)
    cell.text = "Number of vials:"
    cell = table.cell(5, 1)
    cell.text = str(form_data['number_of_vials'])
    
    # Row 7
    cell = table.cell(6, 0)
    cell.text = "Safety risk (Safety data sheet to be provided by the requestor):"
    cell = table.cell(6, 1)
    cell.text = form_data['safety_risk']
    
    # Row 8
    cell = table.cell(7, 0)
    cell.text = "Shipment conditions:"
    cell = table.cell(7, 1)
    cell.text = form_data['shipment_conditions']
    
    # Row 9
    cell = table.cell(8, 0)
    cell.text = "Storage conditions:"
    cell = table.cell(8, 1)
    cell.text = form_data['storage_conditions']
    
    # Add some space
    doc.add_paragraph()
    
    # Analysis Information section
    doc.add_heading('ANALYSIS INFORMATION', level=1)
    
    # GMP Analysis
    p = doc.add_paragraph()
    p.add_run("GMP Analysis: ").bold = True
    p.add_run(f"{form_data['gmp_analysis']}")
    
    if form_data['gmp_analysis'] == 'Yes':
        p.add_run("  For release ☐  For information ☐")
        # Replace the appropriate checkbox with an X
        if form_data['gmp_purpose'] == "For Release":
            p.text = p.text.replace("For release ☐", "For release ☒")
        else:
            p.text = p.text.replace("For information ☐", "For information ☒")
    
    # Analysis Type
    p = doc.add_paragraph()
    p.add_run("Quantitative Analysis ☐  Qualitative Analysis (Screening) ☐")
    # Replace the appropriate checkbox with an X
    if form_data['analysis_type'] == "Quantitative Analysis":
        p.text = p.text.replace("Quantitative Analysis ☐", "Quantitative Analysis ☒")
    else:
        p.text = p.text.replace("Qualitative Analysis (Screening) ☐", "Qualitative Analysis (Screening) ☒")
    
    # Elements to be determined
    p = doc.add_paragraph()
    p.add_run("Element(s) to be determined (quantitative analysis):").bold = True
    
    # Create a list of selected elements
    selected_elements = [element for element, checked in form_data['elements'].items() if checked]
    p = doc.add_paragraph(", ".join(selected_elements))
    
    # ICHQ3D Analysis - separate section
    p = doc.add_paragraph()
    p.add_run("ICHQ3D Analysis:").bold = True
    p.add_run(" " + ("Yes" if form_data['ichq3d_analysis'] else "No"))
    
    if form_data['ichq3d_analysis']:
        p = doc.add_paragraph()
        p.add_run("For ICHQ3D request, documents to be provided:").bold = True
        
        p = doc.add_paragraph("Phase 1 and 2: R&D Medecinal product ID Card (SD-000133)")
        p.paragraph_format.left_indent = Inches(0.5)
        
        p = doc.add_paragraph("Phase 3: Medicinal Product ID Card (SD-000134) and Risk Assessment (SD-000131)")
        p.paragraph_format.left_indent = Inches(0.5)
    
    # Method reference
    p = doc.add_paragraph()
    p.add_run("Method reference and/or specification to be applied if relevant (Veeva Vault or Pharmacopoeia reference):").bold = True
    p = doc.add_paragraph(form_data['method_reference'])
    
    # Add some space
    doc.add_paragraph()
    
    # Request reference section
    p = doc.add_paragraph()
    p.add_run("(Completed by the BioA/AE Laboratory)").italic = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.add_run("Request reference (Steel or iLab): ").bold = True
    p.add_run("_____________________")
    
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
        requestor_site = st.selectbox("Requestor Site", ["Frankfurt", "Vitry", "Framingham"])
        
        col1, col2, col3 = st.columns(3)
        with col1:
            requestor_name = st.text_input("Requestor Name")
        with col2:
            requestor_phone = st.text_input("Requestor Phone")
        with col3:
            requestor_email = st.text_input("Requestor Email")
        
        request_date = st.date_input("Request Date", datetime.now())
        
        st.subheader("Sample Information")
        product_name = st.text_input("Product Name")
        actime_code = st.text_input("Actime Code")
        product_form = st.selectbox("Product Form", ["Drug Product", "Drug Substance", "Other"])
        batch_number = st.text_area("Batch Number(s)")
        
        # Modified sample quantity input
        col1, col2 = st.columns(2)
        with col1:
            sample_unit = st.selectbox("Sample Quantity Unit", ["mg", "ml"])
        with col2:
            sample_quantity = st.number_input(f"Sample Quantity ({sample_unit})", min_value=0.0, step=0.1)
        
        number_of_vials = st.number_input("Number of Vials", min_value=1, step=1)
        safety_risk = st.text_area("Safety Risk")
        shipment_conditions = st.text_input("Shipment Conditions")
        storage_conditions = st.text_input("Storage Conditions")
        
        st.subheader("Analysis Information")
        gmp_analysis = st.radio("GMP Analysis", ["Yes", "No"])
        gmp_purpose = st.radio("Purpose", ["For Release", "For Information"]) if gmp_analysis == "Yes" else "N/A"
        analysis_type = st.radio("Analysis Type", ["Quantitative Analysis", "Qualitative Analysis (Screening)"])
        
        # Elements to be determined
        st.subheader("Elements to be determined")
        st.write("Uncheck elements that are not needed:")
        
        # Calculate number of rows needed (5 elements per row)
        elements_list = list(elements_table.items())
        num_rows = (len(elements_list) + 4) // 5  # Ceiling division by 5
        
        elements_selected = {}
        # Create grid layout for elements
        for row in range(num_rows):
            cols = st.columns(5)
            for col in range(5):
                idx = row * 5 + col
                if idx < len(elements_list):
                    element, properties = elements_list[idx]
                    with cols[col]:
                        # Pre-select all checkboxes (value=True)
                        elements_selected[element] = st.checkbox(
                            f"{element} (Class {properties['Class']})",
                            value=True,  # Pre-select all checkboxes
                            key=f"element_{element}"
                        )
        
        # Separate ICHQ3D Analysis section with some space
        st.markdown("---")
        st.subheader("ICHQ3D Analysis")
        ichq3d_analysis = st.checkbox("Request ICHQ3D Analysis")
        
        if ichq3d_analysis:
            st.info("""
            For ICHQ3D request, documents to be provided:
            * Phase 1 and 2: R&D Medecinal product ID Card (SD-000133)
            * Phase 3: Medicinal Product ID Card (SD-000134) and Risk Assessment (SD-000131)
            """)
        
        method_reference = st.text_area("Method reference and/or specification to be applied if relevant")
        submitted = st.form_submit_button("Generate Document")
    
    if submitted:
        try:
            # Create form data dictionary
            form_data = {
                "requestor_site": requestor_site,
                "requestor_name": requestor_name,
                "requestor_phone": requestor_phone,
                "requestor_email": requestor_email,
                "request_date": request_date.strftime("%Y-%m-%d"),
                "product_name": product_name,
                "actime_code": actime_code,
                "product_form": product_form,
                "batch_number": batch_number,
                "sample_quantity": sample_quantity,
                "sample_unit": sample_unit,
                "number_of_vials": str(number_of_vials),
                "safety_risk": safety_risk,
                "shipment_conditions": shipment_conditions,
                "storage_conditions": storage_conditions,
                "gmp_analysis": gmp_analysis,
                "gmp_purpose": gmp_purpose,
                "analysis_type": analysis_type,
                "elements": elements_selected,
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
                "requestor": requestor_name,
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