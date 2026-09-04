import streamlit as st
import pandas as pd
import csv
import io
import re

# 1. Page Configuration & Setup
st.set_page_config(
    page_title="Dubai Customs EDI Generator",
    page_icon="🇦🇪",
    layout="wide"
)

# Static Reference Data Dictionary based on Customs Guides
CUSTOMS_MASTER = {
    "INCOTERMS": {"CIF": "Cost, Insurance & Freight", "CFR": "Cost and Freight", "FOB": "Free on Board", "EXW": "Ex Works", "CIP": "Carriage & Insurance Paid To", "CPT": "Carriage Paid To"},
    "PAYMENT_METHODS": {"1": "Cash Payment", "2": "T/T - Telex Transfer", "3": "L/C - Letter of Credit", "4": "EP - Electronic Payment", "7": "Bank Transfer"},
    "VEHICLE_BRANDS": {"5": "BENTLEY", "32": "MITSUBISHI", "44": "TOYOTA", "6": "BMW", "30": "MERCEDES-BENZ", "33": "NISSAN", "19": "HYUNDAI"},
    "VEHICLE_TYPES": {"CAR": "Passenger Car", "PIK": "Pick Up Truck", "VAN": "Van / Delivery Vehicle", "TRK": "Commercial Truck", "BUS": "Bus / Transport"}
}

# Sidebar Branding and Admin Info
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

# 2. Login Security Protocol Layer
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown("<h2 style='text-align: center;'>🔐 Customs Portal Authentication</h2>", unsafe_allow_html=True)
    _, col2, _ = st.columns([1, 2, 1])
    with col2:
        with st.form("Login Form"):
            username = st.text_input("Username", value="admin")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Access Workspace")
            if submit:
                if username == "admin" and password == "DubaiCustoms2026":
                    st.session_state["authenticated"] = True
                    st.success("Access Granted! Loading profile...")
                    st.rerun()
                else:
                    st.error("❌ Incorrect username or password configuration.")
    return False

if not check_password():
    st.stop()

# --- MAIN WORKSPACE APP ---
st.title("🇦🇪 Dubai Customs EDI Flat File Platform")
st.subheader("Automated Spreadsheet-to-EDI Translation Studio")
st.caption("🔒 Secured Workspace Session")
st.markdown("---")

# 3. Customs Reference Definition Lookup Tabs
st.markdown("### 🔍 Live Customs Definition Lookup Desk")
tab1, tab2, tab3 = st.tabs(["📦 HS Code / Item Rules", "🚘 Vehicle Specifications", "🌍 Country & Shipping Terms"])

with tab1:
    c1, col_hs = st.columns([2, 3])
    with c1:
        st.write("**HS Commodity Rules:**")
        st.caption("• Must be a minimum of 8 digits.\n• Cannot contain periods, spaces, or letters.")
    with col_hs:
        search_hs = st.text_input("Test an HS Code for verification:", placeholder="e.g., 84133000")
        if search_hs:
            clean_hs = re.sub(r'[^0-9]', '', search_hs)
            if len(clean_hs) >= 8:
                st.success(f"✅ Valid format pattern structure detected ({len(clean_hs)} digits).")
            else:
                st.error(f"❌ Pattern error: Code is only {len(clean_hs)} digits long. Minimum length required is 8 digits.")

with tab2:
    v_col1, v_col2 = st.columns(2)
    with v_col1:
        st.write("**Official Vehicle Brand Reference IDs:**")
        st.dataframe(pd.DataFrame(list(CUSTOMS_MASTER["VEHICLE_BRANDS"].items()), columns=["Brand Code", "Brand Name"]), hide_index=True)
    with v_col2:
        st.write("**Official Vehicle Type Classifications:**")
        st.dataframe(pd.DataFrame(list(CUSTOMS_MASTER["VEHICLE_TYPES"].items()), columns=["Type Code", "Description"]), hide_index=True)

with tab3:
    s_col1, s_col2 = st.columns(2)
    with s_col1:
        st.write("**Valid INCOTERMS Delivery Methods:**")
        st.dataframe(pd.DataFrame(list(CUSTOMS_MASTER["INCOTERMS"].items()), columns=["Incoterm", "Meaning"]), hide_index=True)
    with s_col2:
        st.write("**Payment Instrument Reference Lookups:**")
        st.dataframe(pd.DataFrame(list(CUSTOMS_MASTER["PAYMENT_METHODS"].items()), columns=["Instrument ID", "Method Name"]), hide_index=True)

st.markdown("---")

# 4. Diagnostic Data Validation Logic
def validate_customs_data(parts_df, vehicles_df):
    errors = []
    for idx, row in parts_df.iterrows():
        row_num = idx + 3
        inv = str(row.get("Invoice Number", "Unknown"))
        
        hs = str(row.get("HS Code", "")).strip()
        if pd.isna(row.get("HS Code")) or hs == "":
            errors.append(f"⚠️ Row {row_num} (Inv: {inv}): Missing Commodity HS Code.")
        elif not hs.replace('.','').isdigit() or len(hs.replace('.','')) < 8:
            errors.append(f"⚠️ Row {row_num} (Inv: {inv}): HS Code '{hs}' must contain at least 8 numeric digits.")
            
        coo = str(row.get("Country of Origin (2 Letter)", "")).strip()
        if pd.isna(row.get("Country of Origin (2 Letter)")) or len(coo) != 2:
            errors.append(f"⚠️ Row {row_num} (Inv: {inv}): Country of Origin '{coo}' must be exactly a 2-letter ISO code (e.g., JP, TH).")

    if not vehicles_df.empty:
        for idx, row in vehicles_df.iterrows():
            row_num = idx + 3
            v_inv = str(row.get("Invoice Number Link", "Unknown"))
            chassis = str(row.get("Vehicle Chassis Number", "")).strip()
            if pd.isna(row.get("Vehicle Chassis Number")) or chassis == "":
                errors.append(f"❌ Row {row_num} (Vehicle Tab): Missing Vehicle Chassis Number.")
            elif len(chassis) != 17:
                errors.append(f"❌ Row {row_num} (Vehicle Tab, Inv: {v_inv}): Chassis Number '{chassis}' is {len(chassis)} characters. Must be exactly 17 characters.")
    return errors

# 5. Core EDI Parsing & Translation Logic Engine
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

    parts_df['Invoice Number'] = parts_df['Invoice Number'].astype(str).str.strip()
    if not vehicles_df.empty:
        vehicles_df['Invoice Number Link'] = vehicles_df['Invoice Number Link'].astype(str).str.strip()

    validation_logs = validate_customs_data(parts_df, vehicles_df)
    unique_invoices = [inv for inv in parts_df['Invoice Number'].dropna().unique() if str(inv).lower() != 'nan' and str(inv).strip() != '']
    edi_outputs = {}

    for inv_no in unique_invoices:
        safe_inv_name = re.sub(r'[\\/*?:"<>|]', "", inv_no)
        output_filename = f"Invoice_{safe_inv_name}_Declaration.txt"
        current_items = parts_df[parts_df['Invoice Number'] == inv_no]
        first_row = current_items.iloc

        string_buffer = io.StringIO()
        writer = csv.writer(string_buffer, delimiter=',', quoting=csv.QUOTE_ALL)

        try:
            inv_val = f"{float(first_row.get('Total Invoice Value', 0)):.2f}"
        except Exception:
            inv_val = "0.00"

        # Write Invoice Header (IH)
        ih_row = [
            "IH", inv_no, str(first_row.get("Invoice Date (YYYY-MM-DD)", "")).split()[0] if pd.notna(first_row.get("Invoice Date (YYYY-MM-DD)")) else "",
            "1", "1", str(first_row.get("Seller Name", "")), "1", "1",
            str(first_row.get("Invoice Currency", "AED")), inv_val,
            str(first_row.get("INCO Terms", "CIF")), "", "", "", ""
        ]
        writer.writerow(ih_row)

        # Write Line Item (ID)
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
                str(item.get("Goods Description", "")), str(item.get("Goods Condition (N/U)", "N")),
                str(item.get("Qty Unit", "kg")), str(item.get("Quantity", "1")),
                "kg", net_wt, "", "", line_val,
                str(item.get("Country of Origin (2 Letter)", "JP")).strip().upper(),
                "", "", "", "", ""
            ]
            writer.writerow(id_row)

            # Write Connected Vehicle Sub-Lines (VD)
            if not vehicles_df.empty:
                matching_vds = vehicles_df[(vehicles_df['Invoice Number Link'] == inv_no) & (vehicles_df['Invoice Line Number Link'] == line_no)]
                for _, v in matching_vds.iterrows():
                    vd_row = [
                        "VD", str(v.get("Vehicle Chassis Number", "")).strip().upper(), 
                        str(int(v.get("Vehicle Brand Code", 5))) if pd.notna(v.get("Vehicle Brand Code")) else "5",
                        str(v.get("Vehicle Model", "")), str(v.get("Vehicle Engine Number", "")),
                        f"{float(v.get('Engine Capacity (Liters)', 0)):.2f}" if pd.notna(v.get('Engine Capacity (Liters)')) else "",
