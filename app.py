import streamlit as st
import pandas as pd
import csv
import io
import re

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


# 3. Live Diagnostic Scanner Logic
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

# 4. Document Compiler Engine with Forced Capitalization
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


# --- MOVED INTERFACE TO THE TOP FOR MAXIMUM VISIBILITY ---
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
                    st.download_button(
