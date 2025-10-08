import streamlit as st
import pandas as pd
import os
import re
import io
from datetime import datetime

# Wyoming counties list
WY_COUNTIES = [
    "Albany", "Big Horn", "Campbell", "Carbon", "Converse", "Crook", "Fremont", "Goshen",
    "Hot Springs", "Johnson", "Laramie", "Lincoln", "Natrona", "Niobrara", "Park", "Platte",
    "Sheridan", "Sublette", "Sweetwater", "Teton", "Uinta", "Washakie", "Weston"
]

# Your existing functions (unchanged)
def parse_filer_name(full_name):
    full_name = full_name.strip()
    if not full_name:
        return ""
    parts = full_name.split()
    if len(parts) == 0:
        return ""
    last = parts[0]
    first = ' '.join(parts[1:]) if len(parts) > 1 else ""
    return f"{last}, {first}"

def find_account_col(df):
    account_pattern = re.compile(r'^[MR]\d{7}$')
    for col in df.columns:
        if df[col].astype(str).str.match(account_pattern, na=False).any():
            return col
    return None

def find_name_col(df):
    for col in df.columns:
        if re.search(r'name|owner', col, re.I):
            return col
    return None

def find_phone_col(df):
    for col in df.columns:
        if re.search(r'phone', col, re.I):
            return col
    return None

def get_address(row, original_df):
    parts = []
    predir = str(row.get('Predirection', pd.NA)).strip() if pd.notna(row.get('Predirection', pd.NA)) else ""
    street_no = str(row.get('Street Number', pd.NA)).strip() if pd.notna(row.get('Street Number', pd.NA)) else ""
    street_name = str(row.get('Street Name', pd.NA)).strip() if pd.notna(row.get('Street Name', pd.NA)) else ""
    street_type = str(row.get('Street Type', pd.NA)).strip() if pd.notna(row.get('Street Type', pd.NA)) else ""
    if predir: parts.append(predir)
    if street_no: parts.append(street_no)
    if street_name: parts.append(street_name)
    if street_type: parts.append(street_type)
    return ' '.join(parts)

def compare_excels(df1_bytes, df2_path):
    try:
        df1_orig = pd.read_excel(io.BytesIO(df1_bytes))
        df2_orig = pd.read_excel(df2_path)

        if df1_orig.empty or df2_orig.empty:
            return None, "One or both files are empty."

        key_col1 = find_account_col(df1_orig)
        key_col2 = find_account_col(df2_orig)
        if not key_col1 or not key_col2:
            return None, "Could not identify account number column (M/R + 7 digits) in one or both files."

        name_col1 = find_name_col(df1_orig)
        phone_col1 = find_phone_col(df1_orig)
        filer_address_col1 = next((col for col in df1_orig.columns if 'Filer Address' in col), None)

        if not name_col1:
            st.warning("Name column not found. Will skip name comparison.")
        if not phone_col1:
            st.warning("Phone column not found. Will skip phone comparison.")
        if not filer_address_col1:
            st.warning("Filer Address column not found. Will skip filer address.")

        account_pattern = re.compile(r'^[MR]\d{7}$')
        df1 = df1_orig[df1_orig[key_col1].astype(str).str.match(account_pattern, na=False)].copy()
        df2 = df2_orig[df2_orig[key_col2].astype(str).str.match(account_pattern, na=False)].copy()

        if df1.empty or df2.empty:
            return None, "No valid account numbers found in one or both files after filtering."

        df1.set_index(key_col1, inplace=True)
        df2.set_index(key_col2, inplace=True)
        common = df1[df1.index.isin(df2.index)]

        common_display = []
        for name, group in common.groupby(level=0):
            count = len(group)
            if count > 1:
                note_row = {
                    'Account Number': f"*** The below account has {count} entries ***",
                    'Name': '', 'Address': '', 'Filer Name': '', 'Filer Address': '', 'Filer Phone': ''
                }
                common_display.append(note_row)
            for _, sub_row in group.iterrows():
                name_f1 = sub_row.get(name_col1, pd.NA) if name_col1 else pd.NA
                phone_f1 = sub_row.get(phone_col1, pd.NA) if phone_col1 else pd.NA
                addr_f1 = get_address(sub_row, df1_orig)
                filer_addr_f1 = sub_row.get(filer_address_col1, pd.NA) if filer_address_col1 else pd.NA

                display_row = {
                    'Account Number': name,
                    'Name': str(name_f1) if pd.notna(name_f1) else '',
                    'Address': addr_f1,
                    'Filer Name': parse_filer_name(str(name_f1) if pd.notna(name_f1) else ''),
                    'Filer Address': str(filer_addr_f1) if pd.notna(filer_addr_f1) else '',
                    'Filer Phone': str(phone_f1) if pd.notna(phone_f1) else ''
                }
                common_display.append(display_row)

        common_all = pd.DataFrame(common_display)
        return common_all, None
    except Exception as e:
        return None, f"Failed to compare files: {str(e)}"

def generate_txt_output(common_all):
    if common_all is None or common_all.empty:
        return "No matching accounts found."
    
    output = io.StringIO()
    output.write("ALL MATCHING ACCOUNTS WITH DATA FROM HO APPLICANT FILE\n\n")
    header_fmt = "{:<15} {:<40} {:<30} {:<40} {:<30} {:<20}\n"
    output.write(header_fmt.format('Account Number', 'Name', 'Address', 'Filer Name', 'Filer Address', 'Filer Phone #'))
    sep_fmt = "{:<15} {:<40} {:<30} {:<40} {:<30} {:<20}\n"
    output.write(sep_fmt.format('-'*15, '-'*40, '-'*30, '-'*40, '-'*30, '-'*20))
    data_fmt = "{:<15} {:<40} {:<30} {:<40} {:<30} {:<20}\n"
    for _, row in common_all.iterrows():
        acc = str(row['Account Number']) if pd.notna(row['Account Number']) else ''
        name = str(row['Name']) if pd.notna(row['Name']) else ''
        addr = str(row['Address']) if pd.notna(row['Address']) else ''
        filer_name = str(row['Filer Name']) if pd.notna(row['Filer Name']) else ''
        filer_addr = str(row['Filer Address']) if pd.notna(row['Filer Address']) else ''
        filer_phone = str(row['Filer Phone']) if pd.notna(row['Filer Phone']) else ''
        output.write(data_fmt.format(acc, name, addr, filer_name, filer_addr, filer_phone))
    return output.getvalue()

def get_master_path(county):
    return f"master_lists/{county}/master.xlsx"

# Streamlit App
st.set_page_config(page_title="WY County Excel Comparison Tool", layout="wide")
st.title("Wyoming County Excel Comparison Tool")

# Back to Home button (styled, same tab)
st.markdown(
    """
    <a href="https://assessortools.com" target="_self" rel="noopener noreferrer" style="
        text-decoration: none;
        display: inline-block;
        padding: 8px 16px;
        background-color: #3B82F6;
        color: white;
        border-radius: 6px;
        border: 1px solid #3B82F6;
        font-weight: 500;
        cursor: pointer;
        margin-bottom: 20px;
    " onmouseover="this.style.backgroundColor='#2563EB'; this.style.borderColor='#2563EB';" 
       onmouseout="this.style.backgroundColor='#3B82F6'; this.style.borderColor='#3B82F6';">
        ← Back to Home
    </a>
    """,
    unsafe_allow_html=True
)

# Initialize session state
if 'county' not in st.session_state:
    st.session_state.county = None
if 'master_uploaded' not in st.session_state:
    st.session_state.master_uploaded = False

# Get logged-in county from auth
logged_in_county = os.environ.get('REMOTE_USER', '').strip()
default_county = logged_in_county if logged_in_county in WY_COUNTIES else None

st.subheader("Select Your County")
county = st.selectbox("Choose a county:", WY_COUNTIES, index=WY_COUNTIES.index(default_county) if default_county else 0, key="county_select")
if county != st.session_state.county:
    st.session_state.county = county
    st.session_state.docs_indexed = {}  # Or equivalent for app.py
    st.session_state.search_results = None
    st.session_state.selected_res = None
    st.rerun()

if not county:
    st.warning("Please select a county to proceed.")
    st.stop()

master_path = get_master_path(county)
master_dir = os.path.dirname(master_path)
os.makedirs(master_dir, exist_ok=True)  # Create folder if needed

# Master List Section
st.subheader(f"LTHO Master List for {county} County")
if os.path.exists(master_path):
    st.success(f"✓ Master list loaded from server: {os.path.basename(master_path)}")
    st.info("You can upload a new one below to overwrite.")
    if st.button("Refresh Comparison (Reload Master)", type="secondary"):
        st.session_state.master_uploaded = True  # Trigger reload
        st.rerun()
else:
    st.warning("No master list found for this county. Please upload one.")

master_upload = st.file_uploader("Upload/Update Master List (Excel)", type=['xlsx', 'xls'], key="master_upload")
if master_upload is not None and st.button("Save Master List to Server", type="primary"):
    try:
        with st.spinner("Saving master list..."):
            df = pd.read_excel(master_upload)
            df.to_excel(master_path, index=False, engine='openpyxl')
        st.success(f"Master list saved for {county} County!")
        st.session_state.master_uploaded = True
        st.rerun()
    except Exception as e:
        st.error(f"Failed to save: {str(e)}")

if not os.path.exists(master_path):
    st.stop()  # Can't proceed without master

# Applicant List Section (ephemeral)
st.subheader("HO Applicant List (Temporary - Session Only)")
applicant_upload = st.file_uploader("Upload Applicant List (Excel)", type=['xlsx', 'xls'], key="applicant_upload")

# Compare Button
if st.button("Compare Lists", type="primary") and applicant_upload is not None:
    with st.spinner("Comparing files..."):
        common_all, error = compare_excels(applicant_upload.read(), master_path)
        if error:
            st.error(error)
        else:
            st.success("Comparison complete!")
            st.dataframe(common_all, use_container_width=True)
            
            txt_content = generate_txt_output(common_all)
            st.download_button(
                label="Download Matches as .TXT",
                data=txt_content,
                file_name=f"{county}_LTHO_Matches.txt",
                mime="text/plain"
            )

# Sidebar: Info/Reset
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    - Select your county (hs).
    - Upload/save your master list once (persists on server).
    - Upload applicant list for each session (auto-deletes after).
    - Results downloadable per run.
    """)
    if st.button("Clear Session (Forget County)"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]

        st.rerun()
