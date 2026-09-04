import streamlit as st
import pandas as pd
import csv
import io
import re

# 1. Page & Personal Branding Configurations
st.set_page_config(
    page_title="Dubai Customs EDI Generator",
    page_icon="🇦🇪",
    layout="wide"
)

# Sidebar with Developer Info & Custom Theme Guide
with st.sidebar:
    st.title("👨‍💻 System Administrator")
    st.info("💡 **Developed by: Lahiru**")
    st.markdown("---")
    st.markdown("""
    ### 🎨 How to Change Themes:
    1. Click the **three dots (⋮)** in the top-right corner of the webpage.
    2. Go to **Settings** → **Theme**.
    3. Switch between **Light**, **Dark**, or **Custom System** options to instantly change colors.
    """)

# 2. Secure Login Authentication Control Layer
def check_password():
    """Returns True if the user entered the correct password."""
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    # Render login form window
    st.markdown("<h2 style='text-align: center;'>🔐 Customs Portal Authentication</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        with st.form("Login Form"):
            username = st.text_input("Username", value="admin")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Access Workspace")
            
            # Change these credentials to your preference
            if submit:
                if username == "admin" and password == "DubaiCustoms2026":
                    st.session_state["authenticated"] = True
                    st.success("Access Granted! Loading profile...")
                    st.rerun()
                else:
                    st.error("❌ Incorrect username or password configuration.")
    return False

# Stop execution if user is not authenticated
if not check_password():
    st.stop()

# --- MAIN WORKSPACE INTERFACE LOGIC (Runs only after typing correct password) ---

st.title("🇦🇪 Dubai Customs EDI Flat File Platform")
st.subheader("Automated Spreadsheet-to-EDI Translation Studio")
st.caption("🔒 Secured Workspace Session")
st.markdown("---")

def convert_excel_to_edi_dict(excel_file):
    """Processes uploaded Excel data buffers and formats divided invoice strings."""
    try:
        parts_df = pd.read_excel(excel_file, sheet_name="Invoices & Spare Parts", skiprows=1)
    except Exception as e:
        st.error(f"❌ Failed to parse 'Invoices & Spare Parts' tab. Error: {str(e)}")
        return None

    try:
        vehicles_df = pd.read_excel(excel_file, sheet_name="Vehicle Details", skiprows=1)
    except Exception:
        vehicles_df = pd.DataFrame()

    parts_df['Invoice Number'] = parts_df['Invoice Number'].astype(str).str.strip()
    if not vehicles_df.empty:
        vehicles_df['Invoice Number Link'] = vehicles_df['Invoice Number Link'].astype(str).str.strip()

    unique_invoices = parts_df['Invoice Number'].dropna().unique()
    unique_invoices = [inv for inv in unique_invoices if str(inv).lower() != 'nan' and str(inv).strip() != '']

    edi_outputs = {}

    for inv_no in unique_invoices:
        safe_inv_name = re.sub(r'[\\/*?:"<>|]', "", inv_no)
        output_filename = f"Invoice_{safe_inv_name}_Declaration.txt"
        
        current_items = parts_df[parts_df['Invoice Number'] == inv_no]
        first_row = current_items.iloc[0]

        string_buffer = io.StringIO()
        writer = csv.writer(string_buffer, delimiter=',', quoting=csv.QUOTE_ALL)

        try:
            inv_val = f"{float(first_row.get('Total Invoice Value', 0)):.2f}"
        except Exception:
            inv_val = "0.00"

        ih_row = [
            "IH", inv_no, 
            str(first_row.get("Invoice Date (YYYY-MM-DD)", "")).split()[0] if pd.notna(first_row.get("Invoice Date (YYYY-MM-DD)")) else "",
            "1", "1", str(first_row.get("Seller Name", "")), "1", "1",
            str(first_row.get("Invoice Currency", "AED")), inv_val,
            str(first_row.get("INCO Terms", "CIF")), "", "", "", ""
        ]
        writer.writerow(ih_row)

        for idx, (_, item) in enumerate(current_items.iterrows(), start=1):
            line_no = item.get("Line Number", idx)
            
            try:
                net_wt = f"{float(item.get('Net Weight (kg)', 0)):.4f}"
            except Exception:
                net_wt = "0.0000"
                
            try:
                line_val = f"{float(item.get('Line Total Value', 0)):.2f}"
            except Exception:
                line_val = "0.00"

            id_row = [
                "ID", str(int(line_no)), str(item.get("HS Code", "")),
                str(item.get("Goods Description", "")), str(item.get("Goods Condition (N/U)", "N")),
                str(item.get("Qty Unit", "kg")), str(item.get("Quantity", "1")),
                "kg", net_wt, "", "", line_val,
                str(item.get("Country of Origin (2 Letter)", "JP")), 
                "", "", "", "", ""
            ]
            writer.writerow(id_row)

            if not vehicles_df.empty:
                matching_vds = vehicles_df[
                    (vehicles_df['Invoice Number Link'] == inv_no) & 
                    (vehicles_df['Invoice Line Number Link'] == line_no)
                ]
                for _, v in matching_vds.iterrows():
                    vd_row = [
                        "VD", str(v.get("Vehicle Chassis Number", "")), 
                        str(int(v.get("Vehicle Brand Code", 5))) if pd.notna(v.get("Vehicle Brand Code")) else "5",
                        str(v.get("Vehicle Model", "")), str(v.get("Vehicle Engine Number", "")),
                        f"{float(v.get('Engine Capacity (Liters)', 0)):.2f}" if pd.notna(v.get('Engine Capacity (Liters)')) else "",
                        str(int(v.get("Passenger Capacity", 0))) if pd.notna(v.get('Passenger Capacity')) else "",
                        f"{float(v.get('Carriage Capacity (Tons)', 0)):.2f}" if pd.notna(v.get('Carriage Capacity (Tons)')) else "",
                        str(int(v.get("Year of Built (YYYY)", 2024))) if pd.notna(v.get("Year of Built (YYYY)")) else "", 
                        str(v.get("Vehicle Color", "")), str(v.get("Vehicle Condition (New/Used)", "New")), 
                        str(v.get("Vehicle Type", "CAR")), str(v.get("Vehicle Drive (L/R)", "L")), 
                        str(int(v.get("Specification (1=GCC, 2=Non-GCC)", 1))) if pd.notna(v.get("Specification (1=GCC, 2=Non-GCC)")) else "1"
                    ]
                    writer.writerow(vd_row)

        edi_outputs[output_filename] = string_buffer.getvalue()

    return edi_outputs

col1, col2 = st.columns([1, 2])

with col1:
    st.info("💡 **Step 1: Document Upload**")
    uploaded_file = st.file_uploader("Upload your Customs_Template.xlsx file here", type=["xlsx"])
    
    st.markdown("---")
    st.write("📋 **Quick Compliance Lookup Indicators:**")
    st.caption("• Brand ID 5: BENTLEY | Brand ID 32: MITSUBISHI")
    st.caption("• Condition: N = New, U = Used")
    st.caption("• Specification Standard: 1 = GCC, 2 = Non-GCC")

with col2:
    st.success("⚡ **Step 2: Split Invoice Outputs**")
    if uploaded_file is not None:
        with st.spinner("Processing template tabs..."):
            output_files = convert_excel_to_edi_dict(uploaded_file)
            
        if output_files:
            st.metric(label="Detected Independent Commercial Invoices", value=len(output_files))
            
            for filename, text_data in output_files.items():
                with st.expander(f"📁 {filename}", expanded=True):
                    st.text_area("File preview window", text_data[:400] + "\n...", height=120, disabled=True)
                    st.download_button(
                        label=f"⬇️ Download File Stream",
                        data=text_data,
                        file_name=filename,
                        mime="text/plain",
                        key=filename
                    )
        else:
            st.warning("Workbook columns read successfully but zero transaction lines were isolated.")
    else:
        st.info("Awaiting entry... please drag and drop your active data master into the panel on the left.")
