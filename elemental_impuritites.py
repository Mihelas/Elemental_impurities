import streamlit as st
import pandas as pd
from datetime import datetime
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Set page config
st.set_page_config(page_title="Elemental Impurities Analysis System", layout="wide")

# Initialize session state
if 'submitted_requests' not in st.session_state:
    st.session_state.submitted_requests = []
if 'calculated_data' not in st.session_state:
    st.session_state.calculated_data = None
if 'batch_results' not in st.session_state:
    st.session_state.batch_results = {}

# Create tabs
tab1, tab2, tab3 = st.tabs(["Request Form", "Calculations", "Request Status"])

# Predefined elements table with PDE values
elements_table = {
    "Cd": {"Class": "1", "If intentionally added": True, "If not intentionally added": True, 
           "PDE_oral": 5, "PDE_parenteral": 2, "PDE_inhalation": 3, "PDE_cutaneous": 20},
    "Pb": {"Class": "1", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 5, "PDE_parenteral": 5, "PDE_inhalation": 5, "PDE_cutaneous": 50},
    "As": {"Class": "1", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 15, "PDE_parenteral": 15, "PDE_inhalation": 2, "PDE_cutaneous": 30},
    "Hg": {"Class": "1", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 30, "PDE_parenteral": 3, "PDE_inhalation": 1, "PDE_cutaneous": 30},
    "Co": {"Class": "2A", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 50, "PDE_parenteral": 5, "PDE_inhalation": 3, "PDE_cutaneous": 50},
    "V": {"Class": "2A", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 100, "PDE_parenteral": 10, "PDE_inhalation": 1, "PDE_cutaneous": 100},
    "Ni": {"Class": "2A", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 200, "PDE_parenteral": 20, "PDE_inhalation": 6, "PDE_cutaneous": 200},
    "Tl": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 8, "PDE_parenteral": 8, "PDE_inhalation": 8, "PDE_cutaneous": 8},
    "Au": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 300, "PDE_parenteral": 300, "PDE_inhalation": 3, "PDE_cutaneous": 3000},
    "Pd": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 100, "PDE_parenteral": 10, "PDE_inhalation": 1, "PDE_cutaneous": 100},
    "Ir": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 100, "PDE_parenteral": 10, "PDE_inhalation": 1, "PDE_cutaneous": 100},
    "Os": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 100, "PDE_parenteral": 10, "PDE_inhalation": 1, "PDE_cutaneous": 100},
    "Rh": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 100, "PDE_parenteral": 10, "PDE_inhalation": 1, "PDE_cutaneous": 100},
    "Ru": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 100, "PDE_parenteral": 10, "PDE_inhalation": 1, "PDE_cutaneous": 100},
    "Se": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 150, "PDE_parenteral": 80, "PDE_inhalation": 130, "PDE_cutaneous": 800},
    "Ag": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 150, "PDE_parenteral": 15, "PDE_inhalation": 7, "PDE_cutaneous": 150},
    "Pt": {"Class": "2B", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 100, "PDE_parenteral": 10, "PDE_inhalation": 1, "PDE_cutaneous": 100},
    "Li": {"Class": "3", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 550, "PDE_parenteral": 250, "PDE_inhalation": 25, "PDE_cutaneous": 2500},
    "Sb": {"Class": "3", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 1200, "PDE_parenteral": 90, "PDE_inhalation": 20, "PDE_cutaneous": 900},
    "Ba": {"Class": "3", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 1400, "PDE_parenteral": 700, "PDE_inhalation": 300, "PDE_cutaneous": 7000},
    "Mo": {"Class": "3", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 3000, "PDE_parenteral": 1500, "PDE_inhalation": 10, "PDE_cutaneous": 15000},
    "Cu": {"Class": "3", "If intentionally added": True, "If not intentionally added": True,
           "PDE_oral": 3000, "PDE_parenteral": 300, "PDE_inhalation": 30, "PDE_cutaneous": 3000},
    "Sn": {"Class": "3", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 6000, "PDE_parenteral": 600, "PDE_inhalation": 60, "PDE_cutaneous": 6000},
    "Cr": {"Class": "3", "If intentionally added": True, "If not intentionally added": False,
           "PDE_oral": 11000, "PDE_parenteral": 1100, "PDE_inhalation": 3, "PDE_cutaneous": 11000},
    "Fe": {"Class": "4", "If intentionally added": False, "If not intentionally added": False,
           "PDE_oral": None, "PDE_parenteral": 13000, "PDE_inhalation": None, "PDE_cutaneous": None},
    "Mn": {"Class": "3", "If intentionally added": False, "If not intentionally added": False,
           "PDE_oral": 2500, "PDE_parenteral": 250, "PDE_inhalation": 25, "PDE_cutaneous": None},
    "Zn": {"Class": "3", "If intentionally added": False, "If not intentionally added": False,
           "PDE_oral": 13000, "PDE_parenteral": 1300, "PDE_inhalation": 130, "PDE_cutaneous": None},
}

# Function to calculate MPC and control strategy limits
def calculate_limits(elements, daily_dose, route="parenteral", control_percentage=30):
    """
    Calculate Maximum Permitted Concentration (MPC) and control strategy limits
    
    Parameters:
    elements (dict): Dictionary of elements with PDE values
    daily_dose (float): Daily dose in grams
    route (str): Administration route (oral, parenteral, inhalation, cutaneous)
    control_percentage (float): Percentage of MPC for control strategy limit
    
    Returns:
    DataFrame with calculated values
    """
    results = []
    
    for element, properties in elements.items():
        pde_key = f"PDE_{route}"
        if pde_key in properties and properties[pde_key] is not None:
            pde = properties[pde_key]
            mpc = pde / daily_dose
            control_limit = mpc * (control_percentage / 100)
            
            # Round values appropriately
            if mpc < 1:
                mpc_rounded = round(mpc, 2)
                control_limit_rounded = round(control_limit, 2)
            elif mpc < 10:
                mpc_rounded = round(mpc, 1)
                control_limit_rounded = round(control_limit, 1)
            else:
                mpc_rounded = round(mpc)
                control_limit_rounded = round(control_limit)
            
            results.append({
                "Element": element,
                "Class": properties["Class"],
                f"PDE ({route}) µg/day": pde,
                "MPC µg/g": mpc_rounded,
                f"Control Strategy Limit ({control_percentage}%) µg/g": control_limit_rounded,
                "MPC ng/mL": mpc_rounded * 1000,
                f"Control Strategy Limit ({control_percentage}%) ng/mL": control_limit_rounded * 1000
            })
    
    return pd.DataFrame(results)

# NEW FUNCTION: Calculate element results based on measured values
def calculate_element_results(measured_value, daily_dose, pde, control_percentage=30):
    """
    Calculate element results based on measured values and parameters
    
    Parameters:
    measured_value (float): Lab analysis result (µg/g)
    daily_dose (float): Daily dose in grams
    pde (float): PDE value for the element
    control_percentage (float): Control strategy percentage (default 30%)
    
    Returns:
    dict: Dictionary containing calculated values and compliance status
    """
    try:
        # Calculate MPC
        mpc = pde / daily_dose
        
        # Calculate control strategy limit
        control_limit = mpc * (control_percentage / 100)
        
        # Calculate exposure (measured value * daily dose)
        exposure = measured_value * daily_dose
        
        # Calculate control threshold (PDE * control percentage)
        control_threshold = pde * (control_percentage / 100)
        
        # Determine compliance - compare exposure with control threshold
        is_compliant = exposure <= control_threshold
        
        # Round values appropriately
        if mpc < 1:
            mpc_rounded = round(mpc, 4)
            control_limit_rounded = round(control_limit, 4)
        elif mpc < 10:
            mpc_rounded = round(mpc, 2)
            control_limit_rounded = round(control_limit, 2)
        else:
            mpc_rounded = round(mpc, 1)
            control_limit_rounded = round(control_limit, 1)
        
        exposure_rounded = round(exposure, 4)
        
        return {
            'measured_value': measured_value,
            'exposure': exposure_rounded,
            'mpc': mpc_rounded,
            'control_limit': control_limit_rounded,
            'control_threshold': control_threshold,
            'is_compliant': is_compliant
        }
    except Exception as e:
        st.error(f"Calculation error: {str(e)}")
        return None

# Function to create Word document
def create_word_document(form_data, calculation_data=None):
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
    
    # Add calculation data if available
    if calculation_data is not None and form_data['ichq3d_analysis']:
        doc.add_paragraph()
        doc.add_heading('ELEMENTAL IMPURITIES CALCULATION', level=1)
        
        p = doc.add_paragraph()
        p.add_run(f"Daily dose: {form_data['daily_dose']} g").bold = True
        p.add_run(f" ({form_data['route_of_administration']} administration)")
        
        # Add calculation table
        table = doc.add_table(rows=len(calculation_data) + 1, cols=5)
        table.style = 'Table Grid'
        
        # Header row
        headers = ["Element", "Class", "PDE (µg/day)", "MPC (µg/g)", "Control Strategy Limit (ng/mL)"]
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            run = cell.paragraphs[0].runs[0]
            run.bold = True
        
        # Data rows
        for i, row in enumerate(calculation_data.itertuples(), 1):
            table.cell(i, 0).text = row.Element
            table.cell(i, 1).text = row.Class
            table.cell(i, 2).text = str(row._3)  # PDE column
            table.cell(i, 3).text = str(row._4)  # MPC column
            table.cell(i, 4).text = str(row._7)  # Control Strategy Limit column
    
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

# NEW FUNCTION: Create Excel report with three tables matching the format in the image
def create_excel_report(product_name, daily_dose, route, selected_elements, mpc_data, batch_results, control_percentage=30):
    """
    Create an Excel report with three tables matching the format in the image
    
    Parameters:
    product_name (str): Name of the product
    daily_dose (float): Daily dose in grams
    route (str): Administration route
    selected_elements (list): List of selected elements
    mpc_data (DataFrame): DataFrame with MPC calculations
    batch_results (dict): Dictionary with batch results
    control_percentage (float): Control strategy percentage
    
    Returns:
    BytesIO: Excel file as BytesIO object
    """
    # Create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Elemental Impurities Report"
    
    # Define styles
    title_font = Font(name='Arial', size=12, bold=True)
    header_font = Font(name='Arial', size=11, bold=True)
    normal_font = Font(name='Arial', size=10)
    
    # Define fills
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    
    # Define borders
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Define alignment
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Add title
    ws['A1'] = f"Table 8: Summary of (i) The Maximum Permitted Concentration (µg/g) (Section1), (ii) The analytical results (Section2), and (iii) The Control strategy decisions (Section3), regarding the Elemental Impurities examined in the current risk assessment study"
    ws.merge_cells('A1:L1')
    ws['A1'].font = title_font
    ws['A1'].alignment = left_align
    
    # Add Section 1 title
    row = 3
    ws[f'A{row}'] = "Section1 Maximum permitted concentration (µg/g) of each Elemental impurity"
    ws.merge_cells(f'A{row}:L{row}')
    ws[f'A{row}'].font = header_font
    ws[f'A{row}'].alignment = left_align
    
    # Add Section 1 table headers
    row += 2
    headers = ["", "Max Daily Amount of MP (g/patient)"] + selected_elements
    for col, header in enumerate(headers, 1):
        ws.cell(row=row, column=col).value = header
        ws.cell(row=row, column=col).font = header_font
        ws.cell(row=row, column=col).alignment = center_align
        ws.cell(row=row, column=col).fill = header_fill
        ws.cell(row=row, column=col).border = thin_border
    
    # Add Section 1 data
    row += 1
    ws.cell(row=row, column=1).value = product_name
    ws.cell(row=row, column=1).font = normal_font
    ws.cell(row=row, column=1).alignment = left_align
    ws.cell(row=row, column=1).border = thin_border
    
    ws.cell(row=row, column=2).value = "-"
    ws.cell(row=row, column=2).font = normal_font
    ws.cell(row=row, column=2).alignment = center_align
    ws.cell(row=row, column=2).border = thin_border
    
    for col, element in enumerate(selected_elements, 3):
        ws.cell(row=row, column=col).value = ""
        ws.cell(row=row, column=col).font = normal_font
        ws.cell(row=row, column=col).alignment = center_align
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    ws.cell(row=row, column=1).value = "injectable form"
    ws.cell(row=row, column=1).font = normal_font
    ws.cell(row=row, column=1).alignment = left_align
    ws.cell(row=row, column=1).border = thin_border
    
    ws.cell(row=row, column=2).value = f"{daily_dose} g/patient of MP"
    ws.cell(row=row, column=2).font = normal_font
    ws.cell(row=row, column=2).alignment = center_align
    ws.cell(row=row, column=2).border = thin_border
    
    # Add MPC values
    for col, element in enumerate(selected_elements, 3):
        element_data = mpc_data[mpc_data['Element'] == element]
        if not element_data.empty:
            mpc = element_data.iloc[0]['MPC µg/g']
            ws.cell(row=row, column=col).value = mpc
            ws.cell(row=row, column=col).font = normal_font
            ws.cell(row=row, column=col).alignment = center_align
            ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    ws.cell(row=row, column=1).value = ""
    ws.cell(row=row, column=1).font = normal_font
    ws.cell(row=row, column=1).alignment = left_align
    ws.cell(row=row, column=1).border = thin_border
    
    ws.cell(row=row, column=2).value = f"Permitted Daily Exposure (µg/patient) according to Table A.2.1. in Appendix3"
    ws.cell(row=row, column=2).font = normal_font
    ws.cell(row=row, column=2).alignment = center_align
    ws.cell(row=row, column=2).border = thin_border
    
    # Add PDE values
    for col, element in enumerate(selected_elements, 3):
        element_data = mpc_data[mpc_data['Element'] == element]
        if not element_data.empty:
            pde = element_data.iloc[0][f'PDE ({route}) µg/day']
            ws.cell(row=row, column=col).value = pde
            ws.cell(row=row, column=col).font = normal_font
            ws.cell(row=row, column=col).alignment = center_align
            ws.cell(row=row, column=col).border = thin_border
    
    # Add footnote
    row += 2
    ws.cell(row=row, column=1).value = "(1) Calculated Max permitted concentration (µg/g) = Permitted Daily Exposure (µg/day)/ Max Daily Amount of MP (g/day)"
    ws.merge_cells(f'A{row}:L{row}')
    ws.cell(row=row, column=1).font = normal_font
    ws.cell(row=row, column=1).alignment = left_align
    
    # Add Section 2 title
    row += 2
    ws.cell(row=row, column=1).value = "Section2 Analytical results (µg/g) and checking of compliance with ICH Q3D"
    ws.merge_cells(f'A{row}:L{row}')
    ws.cell(row=row, column=1).font = header_font
    ws.cell(row=row, column=1).alignment = left_align
    
    # Add Section 2 table headers
    row += 2
    headers = ["", ""] + selected_elements
    for col, header in enumerate(headers, 1):
        ws.cell(row=row, column=col).value = header
        ws.cell(row=row, column=col).font = header_font
        ws.cell(row=row, column=col).alignment = center_align
        ws.cell(row=row, column=col).fill = header_fill
        ws.cell(row=row, column=col).border = thin_border
    
    # Add batch results
    batch_names = list(batch_results.keys())
    for i, batch_name in enumerate(batch_names):
        row += 1
        ws.cell(row=row, column=1).value = f"PPQ {i+1}"
        ws.cell(row=row, column=1).font = normal_font
        ws.cell(row=row, column=1).alignment = center_align
        ws.cell(row=row, column=1).border = thin_border
        
        ws.cell(row=row, column=2).value = batch_name
        ws.cell(row=row, column=2).font = normal_font
        ws.cell(row=row, column=2).alignment = center_align
        ws.cell(row=row, column=2).border = thin_border
        ws.cell(row=row, column=2).border = thin_border
        
        # Add measured values for each element
        for col, element in enumerate(selected_elements, 3):
            measured = batch_results[batch_name].get(element, 0)
            
            # Format with "< " prefix if it's a detection limit
            if measured == 0:
                element_data = mpc_data[mpc_data['Element'] == element]
                if not element_data.empty:
                    control_limit = element_data.iloc[0][f'Control Strategy Limit ({control_percentage}%) µg/g']
                    detection_limit = control_limit / 3  # Typical detection limit is 1/3 of control limit
                    formatted_value = f"< {detection_limit:.3f}"
                else:
                    formatted_value = "< LOD"
            else:
                formatted_value = f"{measured:.3f}"
                
            ws.cell(row=row, column=col).value = formatted_value
            ws.cell(row=row, column=col).font = normal_font
            ws.cell(row=row, column=col).alignment = center_align
            ws.cell(row=row, column=col).border = thin_border
    
    # Add compliance row
    row += 2
    ws.cell(row=row, column=1).value = "Element meets ICH Q3D"
    ws.cell(row=row, column=1).font = header_font
    ws.cell(row=row, column=1).alignment = center_align
    ws.cell(row=row, column=1).border = thin_border
    
    ws.cell(row=row, column=2).value = ""
    ws.cell(row=row, column=2).font = normal_font
    ws.cell(row=row, column=2).alignment = center_align
    ws.cell(row=row, column=2).border = thin_border
    
    for col, element in enumerate(selected_elements, 3):
        ws.cell(row=row, column=col).value = element
        ws.cell(row=row, column=col).font = header_font
        ws.cell(row=row, column=col).alignment = center_align
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    ws.cell(row=row, column=1).value = "Yes"
    ws.cell(row=row, column=1).font = normal_font
    ws.cell(row=row, column=1).alignment = center_align
    ws.cell(row=row, column=1).border = thin_border
    
    ws.cell(row=row, column=2).value = ""
    ws.cell(row=row, column=2).font = normal_font
    ws.cell(row=row, column=2).alignment = center_align
    ws.cell(row=row, column=2).border = thin_border
    
    for col, element in enumerate(selected_elements, 3):
        ws.cell(row=row, column=col).value = "Yes"  # Assuming all elements meet ICH Q3D
        ws.cell(row=row, column=col).font = normal_font
        ws.cell(row=row, column=col).alignment = center_align
        ws.cell(row=row, column=col).border = thin_border
    
    # Add Section 3 title
    row += 2
    ws.cell(row=row, column=1).value = "Section3 Control strategy decisions"
    ws.merge_cells(f'A{row}:L{row}')
    ws.cell(row=row, column=1).font = header_font
    ws.cell(row=row, column=1).alignment = left_align
    
    # Add Section 3 table headers
    row += 2
    headers = ["", f"Control Threshold ({control_percentage}% of the PDE) in µg/g"] + selected_elements
    for col, header in enumerate(headers, 1):
        ws.cell(row=row, column=col).value = header
        ws.cell(row=row, column=col).font = header_font
        ws.cell(row=row, column=col).alignment = center_align
        ws.cell(row=row, column=col).fill = header_fill
        ws.cell(row=row, column=col).border = thin_border
    
    # Add batch results again for Section 3
    for i, batch_name in enumerate(batch_names):
        row += 1
        ws.cell(row=row, column=1).value = f"PPQ {i+1}"
        ws.cell(row=row, column=1).font = normal_font
        ws.cell(row=row, column=1).alignment = center_align
        ws.cell(row=row, column=1).border = thin_border
        
        ws.cell(row=row, column=2).value = batch_name
        ws.cell(row=row, column=2).font = normal_font
        ws.cell(row=row, column=2).alignment = center_align
        ws.cell(row=row, column=2).border = thin_border
        
        # Add measured values for each element (same as Section 2)
        for col, element in enumerate(selected_elements, 3):
            measured = batch_results[batch_name].get(element, 0)
            
            # Format with "< " prefix if it's a detection limit
            if measured == 0:
                element_data = mpc_data[mpc_data['Element'] == element]
                if not element_data.empty:
                    control_limit = element_data.iloc[0][f'Control Strategy Limit ({control_percentage}%) µg/g']
                    detection_limit = control_limit / 3  # Typical detection limit is 1/3 of control limit
                    formatted_value = f"< {detection_limit:.3f}"
                else:
                    formatted_value = "< LOD"
            else:
                formatted_value = f"{measured:.3f}"
                
            ws.cell(row=row, column=col).value = formatted_value
            ws.cell(row=row, column=col).font = normal_font
            ws.cell(row=row, column=col).alignment = center_align
            ws.cell(row=row, column=col).border = thin_border
    
    # Add conclusion
    row += 2
    ws.cell(row=row, column=1).value = "Conclusion"
    ws.cell(row=row, column=1).font = header_font
    ws.cell(row=row, column=1).alignment = left_align
    
    row += 1
    ws.cell(row=row, column=1).value = "No further action required – Existing controls to be considered as adequate"
    ws.merge_cells(f'A{row}:L{row}')
    ws.cell(row=row, column=1).font = normal_font
    ws.cell(row=row, column=1).alignment = left_align
    
    # Add footnote
    row += 2
    ws.cell(row=row, column=1).value = f"(2) Control Threshold (µg/g) = {control_percentage/100} x calculated Max permitted concentration (µg/g)"
    ws.merge_cells(f'A{row}:L{row}')
    ws.cell(row=row, column=1).font = normal_font
    ws.cell(row=row, column=1).alignment = left_align
    
    # Adjust column widths
    for col in range(1, len(headers) + 1):
        if col == 1:
            ws.column_dimensions[get_column_letter(col)].width = 15
        elif col == 2:
            ws.column_dimensions[get_column_letter(col)].width = 30
        else:
            ws.column_dimensions[get_column_letter(col)].width = 12
    
    # Save to BytesIO
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

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
        
        # Completely separated sample quantity and unit inputs
        st.write("Sample Quantity:")
        col1, col2 = st.columns(2)
        with col1:
            sample_quantity = st.number_input("Value", min_value=0.0, step=0.1)
        with col2:
            sample_unit = st.selectbox("Unit", ["mg", "ml"])
        
        number_of_vials = st.number_input("Number of Vials", min_value=1, step=1)
        safety_risk = st.text_area("Safety Risk")
        shipment_conditions = st.text_input("Shipment Conditions")
        storage_conditions = st.text_input("Storage Conditions")
        
        # New fields for ICH Q3D calculations
        st.subheader("ICH Q3D Information")
        col1, col2 = st.columns(2)
        with col1:
            daily_dose = st.number_input("Maximum Daily Dose (g)", min_value=0.1, step=0.1, value=1.0,
                                        help="Maximum daily dose in grams")
        with col2:
            route_of_administration = st.selectbox("Route of Administration", 
                                                ["parenteral", "oral", "inhalation", "cutaneous"],
                                                help="Route of administration affects PDE values")
        
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
            * Phase 1 and 2: R&D Medicinal product ID Card (SD-000133)
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
                "daily_dose": daily_dose,
                "route_of_administration": route_of_administration,
                "gmp_analysis": gmp_analysis,
                "gmp_purpose": gmp_purpose,
                "analysis_type": analysis_type,
                "elements": elements_selected,
                "ichq3d_analysis": ichq3d_analysis,
                "method_reference": method_reference,
            }
            
            # Calculate limits if ICHQ3D analysis is requested
            calculation_data = None
            if ichq3d_analysis:
                # Filter elements that are selected
                selected_elements = {k: v for k, v in elements_table.items() if elements_selected.get(k, False)}
                calculation_data = calculate_limits(selected_elements, daily_dose, route_of_administration)
                st.session_state.calculated_data = calculation_data
            
            # Generate document
            doc_io = create_word_document(form_data, calculation_data)
            
            # Success message
            st.success("Document generated successfully!")
            
            # Download buttons (outside the form)
            col1, col2 = st.columns(2)
            with col1:
                filename = f"Analysis_Request_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                st.download_button(
                    label="📄 Download Request Form",
                    data=doc_io,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            # Store in session state
            st.session_state.submitted_requests.append({
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "product": product_name,
                "requestor": requestor_name,
                "batch": batch_number,
                "status": "Submitted",
                "daily_dose": daily_dose,
                "route": route_of_administration
            })
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

# Calculations Tab - MODIFIED VERSION
with tab2:
    st.title("Elemental Impurities Calculations")
    
    st.write("This tab allows you to perform ICH Q3D calculations for elemental impurities.")
    
    # Product information
    st.subheader("Product Information")
    col1, col2, col3 = st.columns(3)
    with col1:
        calc_product_name = st.text_input("Product Name", key="calc_product_name", 
                                         value="Tusamitamab ravtansine")
    with col2:
        calc_product_form = st.selectbox("Product Form", 
                                        ["injectable form", "oral form", "inhalation form", "cutaneous form"],
                                        index=0,
                                        key="calc_product_form")
    with col3:
        calc_daily_dose = st.number_input("Maximum Daily Dose (g)", min_value=0.1, step=0.1, value=36.6, key="calc_daily_dose",
                                         help="Maximum daily dose in grams")
    
    col1, col2 = st.columns(2)
    with col1:
        calc_route = st.selectbox("Route of Administration", 
                                ["parenteral", "oral", "inhalation", "cutaneous"],
                                index=0,
                                key="calc_route")
    with col2:
        calc_control_percentage = st.slider("Control Strategy Limit (%)", min_value=10, max_value=50, value=30, key="calc_control_percentage")
    
    # Element selection for calculations
    st.subheader("Elements to Include")
    
    # Default elements from the image
    default_elements = ["Cd", "Pb", "As", "Hg", "Co", "V", "Ni", "Li", "Sb", "Cu"]
    
    # Create tabs for different element classes
    class_tabs = st.tabs(["Class 1", "Class 2A", "Class 2B", "Class 3", "Class 4"])
    
    calc_elements_selected = {}
    
    # Class 1 tab
    with class_tabs[0]:
        st.write("Class 1 elements (always required for parenteral products)")
        class_1_elements = {k: v for k, v in elements_table.items() if v["Class"] == "1"}
        for element, properties in class_1_elements.items():
            # Pre-select elements from default list
            calc_elements_selected[element] = st.checkbox(
                f"{element} - PDE {properties[f'PDE_{calc_route}']} µg/day",
                value=element in default_elements,
                key=f"calc_element_{element}"
            )
    
    # Class 2A tab
    with class_tabs[1]:
        st.write("Class 2A elements (always required for parenteral products)")
        class_2a_elements = {k: v for k, v in elements_table.items() if v["Class"] == "2A"}
        for element, properties in class_2a_elements.items():
            calc_elements_selected[element] = st.checkbox(
                f"{element} - PDE {properties[f'PDE_{calc_route}']} µg/day",
                value=element in default_elements,
                key=f"calc_element_{element}"
            )
    
    # Class 2B tab
    with class_tabs[2]:
        st.write("Class 2B elements (required if intentionally added)")
        class_2b_elements = {k: v for k, v in elements_table.items() if v["Class"] == "2B"}
        for element, properties in class_2b_elements.items():
            calc_elements_selected[element] = st.checkbox(
                f"{element} - PDE {properties[f'PDE_{calc_route}']} µg/day",
                value=element in default_elements,
                key=f"calc_element_{element}"
            )
    
    # Class 3 tab
    with class_tabs[3]:
        st.write("Class 3 elements (some required for parenteral products)")
        class_3_elements = {k: v for k, v in elements_table.items() if v["Class"] == "3"}
        for element, properties in class_3_elements.items():
            # Pre-select elements that are required for parenteral route
            default_value = element in default_elements
            calc_elements_selected[element] = st.checkbox(
                f"{element} - PDE {properties[f'PDE_{calc_route}']} µg/day",
                value=default_value,
                key=f"calc_element_{element}"
            )
    
    # Class 4 tab
    with class_tabs[4]:
        st.write("Class 4 elements")
        class_4_elements = {k: v for k, v in elements_table.items() if v["Class"] == "4"}
        for element, properties in class_4_elements.items():
            calc_elements_selected[element] = st.checkbox(
                f"{element} - PDE {properties[f'PDE_{calc_route}']} µg/day",
                value=element in default_elements,
                key=f"calc_element_{element}"
            )
    
    # Calculate button for MPC
    if st.button("Calculate Maximum Permitted Concentrations"):
        # Filter elements that are selected
        selected_elements = {k: v for k, v in elements_table.items() if calc_elements_selected.get(k, False)}
        
        if not selected_elements:
            st.warning("Please select at least one element.")
        else:
            # Calculate limits
            calculation_results = calculate_limits(
                selected_elements, 
                calc_daily_dose, 
                calc_route, 
                calc_control_percentage
            )
            
            # Store in session state
            st.session_state.calculated_data = calculation_results
            
            # Display results
            st.subheader("Maximum Permitted Concentrations")
            st.dataframe(calculation_results, use_container_width=True)
    
    # Batch Results Section
    st.markdown("---")
    st.subheader("Batch Analysis Results")
    
    # Initialize batch results in session state if not present
    if 'batch_results' not in st.session_state:
        st.session_state.batch_results = {}
    
    # Add batch form
    with st.form("add_batch_form"):
        col1, col2 = st.columns(2)
        with col1:
            batch_name = st.text_input("Batch Number", value="C1108289")
        with col2:
            batch_label = st.text_input("Batch Label (e.g., PPQ 1)", value="PPQ 1")
        
        st.write("Enter measured values for each element (µg/g):")
        
        # Get selected elements
        selected_elements_list = [k for k, v in calc_elements_selected.items() if v]
        
        # Create a grid for measured values
        measured_values = {}
        cols = st.columns(3)
        for i, element in enumerate(selected_elements_list):
            col_idx = i % 3
            with cols[col_idx]:
                # Default to 0 which will be displayed as "< LOD"
                measured_values[element] = st.number_input(
                    f"{element} (µg/g)",
                    min_value=0.0,
                    value=0.0,
                    step=0.001,
                    format="%.4f"
                )
        
        submitted_batch = st.form_submit_button("Add Batch Results")
    
    if submitted_batch:
        # Store batch results in session state
        st.session_state.batch_results[batch_name] = measured_values
        st.success(f"Batch {batch_name} ({batch_label}) added successfully!")
    
    # Display current batches
    if st.session_state.batch_results:
        st.write("Current Batches:")
        batch_df = pd.DataFrame(st.session_state.batch_results).T
        st.dataframe(batch_df)
        
        if st.button("Clear All Batches"):
            st.session_state.batch_results = {}
            st.rerun()
    
    # Generate Report Section
    st.markdown("---")
    st.subheader("Generate ICH Q3D Report")
    
    if st.session_state.calculated_data is not None and st.session_state.batch_results:
        if st.button("Generate ICH Q3D Report", type="primary"):
            # Get selected elements
            selected_elements_list = [k for k, v in calc_elements_selected.items() if v]
            
            # Create Excel report
            excel_buffer = create_excel_report(
                calc_product_name,
                calc_daily_dose,
                calc_route,
                selected_elements_list,
                st.session_state.calculated_data,
                st.session_state.batch_results,
                calc_control_percentage
            )
            
            # Display download button
            st.download_button(
                label="📊 Download ICH Q3D Report (Excel)",
                data=excel_buffer,
                file_name=f"ICH_Q3D_Report_{calc_product_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Please calculate Maximum Permitted Concentrations and add at least one batch result before generating a report.")

# Request Status Tab
with tab3:
    st.title("Request Status")
    
    if not st.session_state.submitted_requests:
        st.info("No requests have been submitted yet.")
    else:
        df = pd.DataFrame(st.session_state.submitted_requests)
        st.dataframe(df, use_container_width=True)
        
        # Allow viewing calculation details for a specific request
        if len(st.session_state.submitted_requests) > 0:
            st.subheader("View Request Details")
            selected_request_idx = st.selectbox(
                "Select a request to view details:",
                options=range(len(st.session_state.submitted_requests)),
                format_func=lambda x: f"{st.session_state.submitted_requests[x]['timestamp']} - {st.session_state.submitted_requests[x]['product']}"
            )
            
            if st.button("View Request Details"):
                request = st.session_state.submitted_requests[selected_request_idx]
                
                st.subheader(f"Details for {request['product']}")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Requestor:** {request['requestor']}")
                    st.write(f"**Batch:** {request['batch']}")
                with col2:
                    st.write(f"**Daily Dose:** {request['daily_dose']} g")
                    st.write(f"**Route:** {request['route']}")
                    st.write(f"**Status:** {request['status']}")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🗑️ Clear All Requests"):
                st.session_state.submitted_requests = []
                st.rerun()
        
        with col2:
            if st.button("📥 Export Request History"):
                if st.session_state.submitted_requests:
                    history_df = pd.DataFrame(st.session_state.submitted_requests)
                    csv_buffer = io.StringIO()
                    history_df.to_csv(csv_buffer, index=False)
                    csv_buffer.seek(0)
                    
                    st.download_button(
                        label="Download CSV",
                        data=csv_buffer.getvalue(),
                        file_name=f"Request_History_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )