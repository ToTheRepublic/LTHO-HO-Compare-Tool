import streamlit as st
import pandas as pd
import os
import re
import io
import json
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

def load_blacklist(county):
    blacklist_path = f"master_lists/{county}/blacklist.json"
    if os.path.exists(blacklist_path):
        with open(blacklist_path, 'r') as f:
            return set(json.load(f))
    return set()

def save_blacklist(county, blacklist_set):
    blacklist_path = f"master_lists/{county}/blacklist.json"
    os.makedirs(os.path.dirname(blacklist_path), exist_ok=True)
    with open(blacklist_path, 'w') as f:
        json.dump(list(blacklist_set), f)

def compare_excels(df1_bytes, df2_path, blacklist_set):
    try:
        df1_orig = pd.read_excel(io.BytesIO(df1_bytes), engine='openpyxl')
        df2_orig = pd.read_excel(df2_path, engine='openpyxl')

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
        # Filter out blacklisted accounts
        common = common[~common.index.isin(blacklist_set)]

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

def compare_addresses(df1_orig, accounts_path, blacklist_set):
    try:
        accounts_df = pd.read_excel(accounts_path, engine='openpyxl')
        if accounts_df.empty:
            return None, "Accounts file is empty."

        account_col = find_account_col(accounts_df)
        if not account_col:
            return None, "Could not identify account number column in accounts file."

        # Filter for M and R accounts
        mr_df = accounts_df[accounts_df[account_col].astype(str).str.match(r'^[MR]\d{7}$', na=False)].copy()
        # Filter out blacklisted
        mr_df = mr_df[~mr_df[account_col].isin(blacklist_set)]

        if mr_df.empty:
            return pd.DataFrame(), None

        potentials = []
        app_account_col = find_account_col(df1_orig)
        for _, app_row in df1_orig.iterrows():
            app_account = str(app_row.get(app_account_col, '')) if app_account_col else 'N/A'

            app_predir = str(app_row.get('Predirection', '')) if pd.notna(app_row.get('Predirection', '')) else ""
            app_stno = str(app_row.get('Street Number', '')) if pd.notna(app_row.get('Street Number', '')) else ""
            app_stname = str(app_row.get('Street Name', '')) if pd.notna(app_row.get('Street Name', '')) else ""
            app_sttype = str(app_row.get('Street Type', '')) if pd.notna(app_row.get('Street Type', '')) else ""
            app_addr_parts = [app_predir, app_stno, app_stname, app_sttype]
            app_addr = ' '.join(part for part in app_addr_parts if part.strip())
            if not app_addr:
                continue
            app_addr_lower = app_addr.lower().strip()

            for _, mr_row in mr_df.iterrows():
                mr_account = mr_row[account_col]

                mr_addr = str(mr_row.get('Address', '')) if pd.notna(mr_row.get('Address', '')) else ""
                mr_addr_lower = mr_addr.lower().strip()

                if mr_addr and app_addr_lower == mr_addr_lower and app_account != mr_account:
                    potentials.append({
                        'MR Account': mr_account,
                        'MR Address': mr_addr,
                        'Matched Applicant Account': app_account,
                        'Applicant Address': app_addr
                    })

        # Remove duplicates based on MR Account
        unique_potentials = []
        seen_mr = set()
        for p in potentials:
            if p['MR Account'] not in seen_mr:
                unique_potentials.append(p)
                seen_mr.add(p['MR Account'])
        return pd.DataFrame(unique_potentials), None
    except Exception as e:
        return None, f"Failed to compare addresses: {str(e)}"

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

def get_accounts_path(county):
    return f"master_lists/{county}/accounts.xlsx"

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
if 'accounts_uploaded' not in st.session_state:
    st.session_state.accounts_uploaded = False
if 'blacklist' not in st.session_state:
    st.session_state.blacklist = set()

# Get logged-in county from auth
logged_in_county = os.environ.get('REMOTE_USER', '').strip()
default_county = logged_in_county if logged_in_county in WY_COUNTIES else None

st.subheader("Select Your County")
county = st.selectbox("Choose a county:", WY_COUNTIES, index=WY_COUNTIES.index(default_county) if default_county else 0, key="county_select")
if county != st.session_state.county:
    st.session_state.county = county
    st.session_state.master_uploaded = False
    st.session_state.accounts_uploaded = False
    st.session_state.blacklist = load_blacklist(county)
    st.rerun()

if not county:
    st.warning("Please select a county to proceed.")
    st.stop()

master_path = get_master_path(county)
accounts_path = get_accounts_path(county)
blacklist_path = f"master_lists/{county}/blacklist.json"
master_dir = os.path.dirname(master_path)
os.makedirs(master_dir, exist_ok=True)  # Create folder if needed

# Load blacklist into session
st.session_state.blacklist = load_blacklist(county)

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
            df = pd.read_excel(master_upload, engine='openpyxl')
            df.to_excel(master_path, index=False, engine='openpyxl')
        st.success(f"Master list saved for {county} County!")
        st.session_state.master_uploaded = True
        st.rerun()
    except Exception as e:
        st.error(f"Failed to save: {str(e)}")

# Accounts List Section
st.subheader(f"Master Accounts List for {county} County")
if os.path.exists(accounts_path):
    st.success(f"✓ Accounts list loaded from server: {os.path.basename(accounts_path)}")
    st.info("You can upload a new one below to overwrite.")
    if st.button("Refresh Comparison (Reload Accounts)", type="secondary"):
        st.session_state.accounts_uploaded = True  # Trigger reload
        st.rerun()
else:
    st.warning("No accounts list found for this county. Please upload one.")

accounts_upload = st.file_uploader("Upload/Update Master Accounts List (Excel)", type=['xlsx', 'xls'], key="accounts_upload")
if accounts_upload is not None and st.button("Save Accounts List to Server", type="primary"):
    try:
        with st.spinner("Saving accounts list..."):
            df = pd.read_excel(accounts_upload, engine='openpyxl')
            df.to_excel(accounts_path, index=False, engine='openpyxl')
        st.success(f"Accounts list saved for {county} County!")
        st.session_state.accounts_uploaded = True
        st.rerun()
    except Exception as e:
        st.error(f"Failed to save: {str(e)}")

if not os.path.exists(master_path) or not os.path.exists(accounts_path):
    st.stop()  # Can't proceed without both

# Applicant List Section (ephemeral)
st.subheader("HO Applicant List (Temporary - Session Only)")
applicant_upload = st.file_uploader("Upload Applicant List (Excel)", type=['xlsx', 'xls'], key="applicant_upload")

# Compare Button
if st.button("Compare Lists", type="primary") and applicant_upload is not None:
    with st.spinner("Comparing files..."):
        applicant_bytes = applicant_upload.read()
        common_all, error = compare_excels(applicant_bytes, master_path, st.session_state.blacklist)
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

            # Second comparison: Address matches for M/R
            st.subheader("Potential M/R Accounts")
            df1_orig = pd.read_excel(io.BytesIO(applicant_bytes), engine='openpyxl')
            mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
            if mr_error:
                st.error(mr_error)
            elif not mr_potentials.empty:
                st.dataframe(mr_potentials, use_container_width=True)
                
                # Add to blacklist
                potential_accounts = mr_potentials['MR Account'].tolist()
                selected_to_blacklist = st.multiselect("Select M/R Accounts to Add to Blacklist:", potential_accounts, key="add_blacklist")
                if st.button("Add Selected to Blacklist"):
                    st.session_state.blacklist.update(selected_to_blacklist)
                    save_blacklist(county, st.session_state.blacklist)
                    st.success(f"Added {len(selected_to_blacklist)} accounts to blacklist.")
                    st.rerun()
            else:
                st.info("No potential M/R address matches found.")

# Blacklist Management
with st.expander("Blacklist Management", expanded=False):
    st.write(f"Current Blacklist ({len(st.session_state.blacklist)} accounts):")
    blacklist_list = list(st.session_state.blacklist)
    if blacklist_list:
        st.write(blacklist_list)
        selected_to_remove = st.multiselect("Select Accounts to Remove from Blacklist:", blacklist_list, key="remove_blacklist")
        if st.button("Remove Selected from Blacklist"):
            for acc in selected_to_remove:
                st.session_state.blacklist.discard(acc)
            save_blacklist(county, st.session_state.blacklist)
            st.success(f"Removed {len(selected_to_remove)} accounts from blacklist.")
            st.rerun()
    else:
        st.info("Blacklist is empty.")

# Sidebar: Info/Reset
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    - Select your county (hs).
    - Upload/save your master list and accounts list once (persist on server).
    - Upload applicant list for each session (auto-deletes after).
    - Results downloadable per run.
    - Use Blacklist Management to add/remove accounts to ignore in future comparisons.
    """)
    if st.button("Clear Session (Forget County)"):
        for key in list(st.session_state.keys()):
            if key != 'county':  # Keep county
                del st.session_state[key]
        st.session_state.blacklist = load_blacklist(county)
        st.rerun()