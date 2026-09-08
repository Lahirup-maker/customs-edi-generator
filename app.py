import streamlit as st
import pandas as pd
import csv
import io
import re
from datetime import datetime

# Pull official lookup data blocks from separate customs_data file
import customs_data

# 1. Page Configuration & Setup
st.set_page_config(
    page_title="Dubai Customs EDI Generator",
    page_icon="🇦🇪",
    layout="wide"
)

# Sidebar Layout with Developer Info
with st.sidebar:
    st.title("👨‍💻 System Administrator")
    st.info("💡 **Developed by: Lahiru**")
    st.markdown("---")
    st.markdown("""
    ### 🎨 How to Change Themes:
    1. Click the **three dots (⋮)** in the top-right corner of the webpage.
    2. Go to **Settings** → **Theme**.
    3. Switch between **Light** or **Dark** templates.
    """)

# 2. Secure Login Authentication Control Layer
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown("<h2 style='text-align: center;'>🔐 Customs Portal Authentication</h2>", unsafe_allow_html=True)
    
    _, col2, _ = st.columns(3)
    with col2:
        with st.form("Login Form"):
            username = st.text_input("Username", value="admin")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Access Workspace")
            if submit:
                if username == "admin" and password == "DubaiCustoms2026":
                    st.session_state["authenticated"] = True
                    st.success("Access Granted!")
                    st.rerun()
                else:
                    st.error("❌ Incorrect credentials.")
    return False

if not check_password():
    st.stop()

# --- APPLICATION WORKSPACE ---
st.title("🇦🇪 Dubai Customs EDI Flat File Platform")
st.subheader("Automated Spreadsheet-to-EDI Translation Studio")
st.caption("🔒 Secured Workspace Session")
st.markdown("---")


# 3. Generate Built-in Template Function
def generate_sample_template():
    """
    Generate a properly formatted Excel template with sample data
    """
    # Create sample data for Invoices & Spare Parts sheet
    parts_data = {
        "Invoice Number": ["INV-2024-001", "INV-2024-001", "INV-2024-002"],
        "Invoice Date (YYYY-MM-DD)": ["2024-01-15", "2024-01-15", "2024-01-20"],
        "Line Number": [1, 2, 1],
        "HS Code": ["87082990", "40117020", "85044020"],
        "Goods Description": ["VEHICLE SPARE PARTS", "RUBBER TYRES", "AUTOMOTIVE WIRING"],
        "Goods Condition (N/U)": ["N", "N", "N"],
        "Quantity": [5, 10, 20],
        "Net Weight (kg)": [150.5, 280.0, 45.25],
        "Line Total Value": [5000.00, 3500.00, 2200.00],
        "Country of Origin (2 Letter)": ["JP", "IN", "DE"],
        "Seller Name": ["HONDA PARTS CO", "APOLLO TYRES", "BOSCH GMBH"],
        "Invoice Currency": ["AED", "AED", "AED"],
        "Total Invoice Value": [8500.00, 3500.00, 2200.00],
        "INCO Terms": ["CIF", "FOB", "CIF"],
    }
    
    parts_df = pd.DataFrame(parts_data)
    
    # Create sample data for Vehicle Details sheet
    vehicle_data = {
        "Invoice Number Link": ["INV-2024-001", "INV-2024-001"],
        "Invoice Line Number Link": [1, 2],
        "Vehicle Chassis Number": ["WBADT43452G915123", "IFFIN2S15Y5X25894"],
        "Vehicle Brand Code": ["5", "8"],
        "Vehicle Model": ["BMW 320I", "MARUTI SWIFT"],
        "Vehicle Engine Number": ["SN123456", "EN987654"],
        "Engine Capacity (Liters)": [2.0, 1.2],
        "Passenger Capacity": [5, 5],
        "Carriage Capacity (Tons)": [0.0, 0.0],
        "Year of Built (YYYY)": [2023, 2023],
        "Vehicle Color": ["BLACK", "WHITE"],
        "Vehicle Condition (New/Used)": ["New", "New"],
        "Vehicle Type": ["CAR", "CAR"],
        "Vehicle Drive (L/R)": ["L", "R"],
        "Specification (1=GCC, 2=Non-GCC)": [2, 1],
    }
    
    vehicle_df = pd.DataFrame(vehicle_data)
    
    # Create Excel file with multiple sheets
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        parts_df.to_excel(writer, sheet_name="Invoices & Spare Parts", index=False, startrow=1)
        vehicle_df.to_excel(writer, sheet_name="Vehicle Details", index=False, startrow=1)
        
        # Add header row with instructions
        workbook = writer.book
        
        # Invoices & Spare Parts sheet
        ws1 = writer.sheets["Invoices & Spare Parts"]
        ws1['A1'] = "INSTRUCTIONS: Fill in the data below. All fields marked with * are required. HS Codes must be 8+ digits. Country codes must be 2 letters."
        ws1['A1'].font = ws1['A1'].font.copy()
        ws1.row_dimensions[1].height = 30
        
        # Vehicle Details sheet
        ws2 = writer.sheets["Vehicle Details"]
        ws2['A1'] = "INSTRUCTIONS: Fill in vehicle details. Chassis Number must be exactly 17 characters (VIN). Leave blank if no vehicles."
        ws2['A1'].font = ws2['A1'].font.copy()
        ws2.row_dimensions[1].height = 25
    
    output.seek(0)
    return output.getvalue()


# 4. Live Diagnostic Scanner Logic
def validate_customs_data(parts_df, vehicles_df):
    errors = []
    for idx, row in parts_df.iterrows():
        row_num = idx + 3
        inv = str(row.get("Invoice Number", "Unknown")).upper()
        
        hs = str(row.get("HS Code", "")).replace('.', '').strip()
        if pd.isna(row.get("HS Code")) or hs == "":
            errors.append(f"⚠️ Row {row_num} (Inv: {inv}): Missing HS Commodity Code.")
        elif not hs.isdigit() or len(hs) < 8:
            errors.append(f"⚠️ Row {row_num} (Inv: {inv}): Code '{hs}' is invalid. Needs at least 8 digits.")
            
        coo = str(row.get("Country of Origin (2 Letter)", "")).strip()
        if pd.isna(row.get("Country of Origin (2 Letter)")) or len(coo) != 2:
            errors.append(f"⚠️ Row {row_num} (Inv: {inv}): Country code '{coo}' should be 2 letters.")

    if not vehicles_df.empty:
        for idx, row in vehicles_df.iterrows():
            row_num = idx + 3
            v_inv = str(row.get("Invoice Number Link", "Unknown")).upper()
            chassis = str(row.get("Vehicle Chassis Number", "")).strip()
            
            if pd.isna(row.get("Vehicle Chassis Number")) or chassis == "":
                errors.append(f"❌ Row {row_num} (Vehicle Sheet): Missing Chassis / VIN.")
            elif len(chassis) != 17:
                errors.append(f"❌ Row {row_num} (Vehicle Sheet, Inv: {v_inv}): Chassis string '{chassis}' is {len(chassis)} digits. Standard VINs are exactly 17 characters.")
    return errors

# 5. Document Compiler Engine with Forced Capitalization
def convert_excel_to_edi_dict(excel_file):
    try:
        parts_df = pd.read_excel(excel_file, sheet_name="Invoices & Spare Parts", skiprows=1)
    except Exception as e:
        st.error(f"❌ Failed to parse 'Invoices & Spare Parts' tab. Error: {str(e)}")
        return None, None

    try:
        vehicles_df = pd.read_excel(excel_file, sheet_name="Vehicle Details", skiprows=1)
    except Exception:
        vehicles_df = pd.DataFrame()

    parts_df['Invoice Number'] = parts_df['Invoice Number'].astype(str).str.strip().str.upper()
    if not vehicles_df.empty:
        vehicles_df['Invoice Number Link'] = vehicles_df['Invoice Number Link'].astype(str).str.strip().str.upper()

    validation_logs = validate_customs_data(parts_df, vehicles_df)
    unique_invoices = [inv for inv in parts_df['Invoice Number'].dropna().unique() if str(inv).lower() != 'nan' and str(inv).strip() != '']
    edi_outputs = {}

    for inv_no in unique_invoices:
        safe_inv_name = re.sub(r'[\\/*?:"<>|]', "", inv_no).upper()
        output_filename = f"Invoice_{safe_inv_name}_Declaration.txt"
        current_items = parts_df[parts_df['Invoice Number'] == inv_no]
        first_row = current_items.iloc[0]

        string_buffer = io.StringIO()
        writer = csv.writer(string_buffer, delimiter=',', quoting=csv.QUOTE_ALL)

        try:
            inv_val = f"{float(first_row.get('Total Invoice Value', 0)):.2f}"
        except Exception:
            inv_val = "0.00"

        # Compilation Header row (IH)
        ih_row = [
            "IH", 
            inv_no.upper(), 
            str(first_row.get("Invoice Date (YYYY-MM-DD)", "")).split()[0] if pd.notna(first_row.get("Invoice Date (YYYY-MM-DD)")) else "",
            "1", "1", 
            str(first_row.get("Seller Name", "")).upper(), 
            "1", "1",
            str(first_row.get("Invoice Currency", "AED")).upper(), 
            inv_val,
            str(first_row.get("INCO Terms", "CIF")).upper(), 
            "", "", "", ""
        ]
        writer.writerow(ih_row)

        # Compilation Detail line row (ID)
        for idx, (_, item) in enumerate(current_items.iterrows(), start=1):
            line_no = item.get("Line Number", idx)
            clean_item_hs = str(item.get("HS Code", "")).replace('.', '').strip()
            try:
                net_wt = f"{float(item.get('Net Weight (kg)', 0)):.4f}"
            except Exception:
                net_wt = "0.0000"
            try:
                line_val = f"{float(item.get('Line Total Value', 0)):.2f}"
            except Exception:
                line_val = "0.00"

            id_row = [
                "ID", str(int(line_no)), clean_item_hs,
                str(item.get("Goods Description", "")).upper(), 
                str(item.get("Goods Condition (N/U)", "N")).upper(),
                "kg", 
                str(item.get("Quantity", "1")),
                "kg", net_wt, "", "", line_val,
                str(item.get("Country of Origin (2 Letter)", "JP")).strip().upper(),
                "", "", "", "", ""
            ]
            writer.writerow(id_row)

            # Compilation Nested Vehicle sub-row (VD)
            if not vehicles_df.empty:
                matching_vds = vehicles_df[(vehicles_df['Invoice Number Link'] == inv_no) & (vehicles_df['Invoice Line Number Link'] == line_no)]
                for _, v in matching_vds.iterrows():
                    brand_raw = str(v.get("Vehicle Brand Code", 5)).strip()
                    brand_val = str(int(float(brand_raw))) if brand_raw.replace('.','').replace('-','').isdigit() else "5"
                    type_val = str(v.get("Vehicle Type", "CAR")).strip().upper()
                    
                    vd_row = [
                        "VD", 
                        str(v.get("Vehicle Chassis Number", "")).strip().upper(), 
                        brand_val,
                        str(v.get("Vehicle Model", "")).upper(), 
                        str(v.get("Vehicle Engine Number", "")).upper(),
                        f"{float(v.get('Engine Capacity (Liters)', 0)):.2f}" if pd.notna(v.get('Engine Capacity (Liters)')) else "",
                        str(int(v.get("Passenger Capacity", 0))) if pd.notna(v.get('Passenger Capacity')) else "",
                        f"{float(v.get('Carriage Capacity (Tons)', 0)):.2f}" if pd.notna(v.get('Carriage Capacity (Tons)')) else "",
                        str(int(v.get("Year of Built (YYYY)", 2024))) if pd.notna(v.get("Year of Built (YYYY)")) else "", 
                        str(v.get("Vehicle Color", "")).upper(), 
                        str(v.get("Vehicle Condition (New/Used)", "New")).upper(), 
                        type_val, 
                        str(v.get("Vehicle Drive (L/R)", "L")).upper(), 
                        str(int(v.get("Specification (1=GCC, 2=Non-GCC)", 1))) if pd.notna(v.get("Specification (1=GCC, 2=Non-GCC)")) else "1"
                    ]
                    vd_row = [str(x) for x in vd_row]
                    writer.writerow(vd_row)

        edi_outputs[output_filename] = string_buffer.getvalue()
    return edi_outputs, validation_logs


# --- TEMPLATE DOWNLOAD SECTION ---
st.markdown("## 📋 Built-in Template Download")
col_template1, col_template2 = st.columns(2)

with col_template1:
    st.info("📥 **Download Sample Template**")
    st.caption("Get a pre-formatted Excel file with all required columns and sample data.")
    template_data = generate_sample_template()
    st.download_button(
        label="⬇️ Download Customs_Template.xlsx",
        data=template_data,
        file_name=f"Customs_Template_Sample_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="template_download"
    )

with col_template2:
    st.success("📝 **Template Instructions**")
    st.markdown("""
    ✅ **Required Columns** (Invoices Sheet):
    - Invoice Number, Date, HS Code, Country Code
    - Quantity, Weight, Line Value
    
    ✅ **Vehicle Details** (Optional Sheet):
    - Chassis Number (17 chars), Brand Code, Model
    - Engine details, Color, Condition
    
    ⚠️ **Data Rules**:
    - HS Codes: 8+ digits (dots removed automatically)
    - Country: Exactly 2 letters (e.g., JP, DE, IN)
    - Dates: YYYY-MM-DD format
    """)

st.markdown("---")

# --- MOVED INTERFACE TO THE TOP FOR MAXIMUM VISIBILITY ---
st.markdown("## ⚙️ Upload & Process Declaration")
w1, w2 = st.columns(2)

with w1:
    st.info("💡 **Step 1: Document Upload**")
    uploaded_file = st.file_uploader("Upload your Customs_Template.xlsx file here", type=["xlsx"])

with w2:
    st.success("⚡ **Step 2: Analysis & File Generation**")
    if uploaded_file is not None:
        with st.spinner("Analyzing spreadsheet arrays and forcing uppercase formatting..."):
            output_files, logs = convert_excel_to_edi_dict(uploaded_file)
            
        if logs:
            st.error("🚨 **Spreadsheet Validation Warning Logs**")
            for log in logs:
                st.write(log)
            st.warning("Please verify data warnings prior to clearing submissions.")
        else:
            st.success("✅ **Data Quality Scan Passed! All values have been processed successfully.**")
            
        if output_files:
            st.markdown(f"### 📂 Split Invoice Batches Available: `{len(output_files)}`")
            for filename, text_data in output_files.items():
                with st.expander(f"📁 {filename}", expanded=True):
                    st.text_area("File content preview (All Capitalized)", text_data[:400], height=120, disabled=True)
                    # Complete Download Button structure configuration
                    st.download_button(
                        label=f"⬇️ Download File Stream", 
                        data=text_data,
                        file_name=filename, 
                        mime="text/plain", 
                        key=filename
                    )
    else:
        st.info("System operational. Drop your tracking spreadsheets above to execute structural compilation blocks.")

st.markdown("<br><br><hr>", unsafe_allow_html=True)


# 5. Interactive Code Search desk (Moved to the bottom)
st.markdown("### 🔍 Live Customs Definition Lookup Desk")
tab1, tab2, tab3 = st.tabs(["📦 HS Code & Rules", "🚘 Vehicle Specification Codes", "🌍 Payment Terms & Currencies"])

with tab1:
    c1, col_hs = st.columns(2)
    with c1:
        st.write("**HS Code Rules:**")
        st.caption("• Must be exactly 8 digits or longer.\n• Period characters are stripped dynamically by our engine.")
    with col_hs:
        search_hs = st.text_input("Test formatting rules for an HS code:", placeholder="e.g., 8708.29.90")
        if search_hs:
            clean_hs = re.sub(r'[^0-9]', '', search_hs)
            if len(clean_hs) >= 8:
                st.success(f"✅ Valid Format Pattern ({len(clean_hs)} digits generated).")
            else:
                st.error(f"❌ Error: Extracted code is only {len(clean_hs)} digits. Target needs 8 digits.")

with tab2:
    v_col1, v_col2 = st.columns(2)
    with v_col1:
        st.write("**Official Vehicle Brand Codes (451 Records):**")
        search_brand = st.text_input("Search Brand Name (e.g. TOYOTA, BENTLEY):").upper()
        brand_data = [{"Code": k, "Name": v} for k, v in customs_data.VEHICLE_BRANDS.items() if search_brand in v]
        st.dataframe(pd.DataFrame(brand_data), hide_index=True, use_container_width=True, height=200)
    with v_col2:
        st.write("**Official Vehicle Type Classifications:**")
        st.dataframe(pd.DataFrame(list(customs_data.VEHICLE_TYPES.items()), columns=["Type Code", "Description"]), hide_index=True, use_container_width=True, height=200)

with tab3:
    s_col1, s_col2 = st.columns(2)
    with s_col1:
        st.write("**Payment Instrument Codes:**")
        st.dataframe(pd.DataFrame(list(customs_data.PAYMENT_METHODS.items()), columns=["ID Code", "Method Name"]), hide_index=True, use_container_width=True)
    with s_col2:
        st.write("**INCOTERMS Codes:**")
        st.dataframe(pd.DataFrame(list(customs_data.INCOTERMS.items()), columns=["ID", "Incoterm Code"]), hide_index=True, use_container_width=True)
