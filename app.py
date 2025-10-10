import streamlit as st
import streamlit_javascript as st_js  # New import for JS detection
import pandas as pd
import os
import re
import io
import json
from datetime import datetime
from typing import Optional

# Wyoming counties list
WY_COUNTIES = [
    "Albany", "Big Horn", "Campbell", "Carbon", "Converse", "Crook", "Fremont", "Goshen",
    "Hot Springs", "Johnson", "Laramie", "Lincoln", "Natrona", "Niobrara", "Park", "Platte",
    "Sheridan", "Sublette", "Sweetwater", "Teton", "Uinta", "Washakie", "Weston"
]

# Subdomain to county mapping (based on slugs from county-landing.html)
SUBDOMAIN_TO_COUNTY = {
    'albany': 'Albany',
    'big-horn': 'Big Horn',
    'campbell': 'Campbell',
    'carbon': 'Carbon',
    'converse': 'Converse',
    'crook': 'Crook',
    'fremont': 'Fremont',
    'goshen': 'Goshen',
    'hot-springs': 'Hot Springs',
    'johnson': 'Johnson',
    'laramie': 'Laramie',
    'lincoln': 'Lincoln',
    'natrona': 'Natrona',
    'niobrara': 'Niobrara',
    'park': 'Park',
    'platte': 'Platte',
    'sheridan': 'Sheridan',
    'sublette': 'Sublette',
    'sweetwater': 'Sweetwater',
    'teton': 'Teton',
    'uinta': 'Uinta',
    'washakie': 'Washakie',
    'weston': 'Weston'
}

# Detect subdomain via JS and set county (no cache - called once per run)
def detect_county():
    try:
        # JS expression to get subdomain (first part of hostname) - no 'return' needed
        subdomain = st_js.st_javascript("window.location.hostname.split('.')[0]")
        county = SUBDOMAIN_TO_COUNTY.get(subdomain.lower(), WY_COUNTIES[0])
        if subdomain.lower() not in SUBDOMAIN_TO_COUNTY:
            st.warning(f"Subdomain '{subdomain}' not recognized; defaulting to '{county}'.")
        return county
    except:
        # Fallback if JS fails
        return WY_COUNTIES[0]

county = detect_county()

# Early session state init for county (avoids flash)
if 'detected_county' not in st.session_state:
    st.session_state.detected_county = None

# Detect and store county
if st.session_state.detected_county is None:
    st.session_state.detected_county = detect_county()
    st.rerun()  # Immediate rerun to apply county-specific title/config

# county = st.session_state.detected_county

# Now set county-specific title/config on rerun (overrides placeholder)
st.set_page_config(page_title=f"LTHO-HO Compare Tool - {county} County", layout="wide")
st.title(f"{county} LTHO-HO Comparison Tool")

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
            data = json.load(f)
            if isinstance(data, list) and data and isinstance(data[0], str):
                # Migrate old format: list of strings to list of dicts
                return [{'account': acc, 'applicant_address': '', 'norm_addr': ''} for acc in data]
            else:
                return data
    return []

def save_blacklist(county, blacklist_list):
    blacklist_path = f"master_lists/{county}/blacklist.json"
    os.makedirs(os.path.dirname(blacklist_path), exist_ok=True)
    with open(blacklist_path, 'w') as f:
        json.dump(blacklist_list, f)

def normalize_address(addr):
    if not addr:
        return ''
    addr = addr.lower().strip()
    # Common replacements: full to abbr
    replacements = {
        r'\bstreet\b': 'st',
        r'\bavenue\b': 'ave',
        r'\boulevard\b': 'blvd',
        r'\bdrive\b': 'dr',
        r'\broad\b': 'rd',
        r'\bcircle\b': 'cir',
        r'\bcourt\b': 'ct',
        r'\blane\b': 'ln',
        r'\bplace\b': 'pl',
        r'\balley\b': 'aly',
        r'\bcenter\b': 'ctr',
        r'\bhighway\b': 'hwy',
        # Add more as needed
    }
    for full, abbr in replacements.items():
        addr = re.sub(full, abbr, addr)
    # Remove extra spaces
    addr = re.sub(r'\s+', ' ', addr).strip()
    return addr

def compare_excels(df1_bytes, df2_path, blacklist_list):
    blacklist_accounts = {d['account'] for d in blacklist_list if isinstance(d, dict) and 'account' in d}
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
        common = common[~common.index.isin(blacklist_accounts)]

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

def compare_addresses(df1_orig, accounts_path, blacklist_list):
    blacklist_norms = {d['norm_addr'] for d in blacklist_list if isinstance(d, dict) and 'norm_addr' in d}
    try:
        accounts_df = pd.read_excel(accounts_path, engine='openpyxl')
        if accounts_df.empty:
            return None, "Accounts file is empty."

        account_col = find_account_col(accounts_df)
        if not account_col:
            return None, "Could not identify account number column in accounts file."

        blacklist_accounts = {d['account'] for d in blacklist_list if isinstance(d, dict) and 'account' in d}
        # Filter for M and R accounts
        mr_df = accounts_df[accounts_df[account_col].astype(str).str.match(r'^[MR]\d{7}$', na=False)].copy()
        # Filter out blacklisted accounts
        mr_df = mr_df[~mr_df[account_col].isin(blacklist_accounts)]

        if mr_df.empty:
            return pd.DataFrame(), None

        # Normalize applicant addresses
        applicant_addrs = {}
        app_account_col = find_account_col(df1_orig)
        for _, app_row in df1_orig.iterrows():
            app_account = str(app_row.get(app_account_col, '')) if app_account_col else 'N/A'

            app_predir = str(app_row.get('Predirection', '')) if pd.notna(app_row.get('Predirection', '')) else ""
            app_stno = str(app_row.get('Street Number', '')) if pd.notna(app_row.get('Street Number', '')) else ""
            app_stname = str(app_row.get('Street Name', '')) if pd.notna(app_row.get('Street Name', '')) else ""
            app_sttype = str(app_row.get('Street Type', '')) if pd.notna(app_row.get('Street Type', '')) else ""
            app_addr_parts = [p.strip() for p in [app_predir, app_stno, app_stname, app_sttype]]
            app_addr = ' '.join(part for part in app_addr_parts if part)
            if not app_addr:
                continue
            app_addr_norm = normalize_address(app_addr)

            # Skip if address is blacklisted
            if app_addr_norm in blacklist_norms:
                continue

            if app_addr_norm:
                if app_addr_norm not in applicant_addrs:
                    applicant_addrs[app_addr_norm] = []
                applicant_addrs[app_addr_norm].append({
                    'Account': app_account,
                    'Address': app_addr
                })

        # Normalize MR addresses
        mr_addrs = {}
        for _, mr_row in mr_df.iterrows():
            mr_account = mr_row[account_col]

            mr_addr = str(mr_row.get('ADDRESS', '')) if pd.notna(mr_row.get('ADDRESS', '')) else ""
            if not mr_addr:
                continue
            mr_addr_norm = normalize_address(mr_addr)

            if mr_addr_norm:
                if mr_addr_norm not in mr_addrs:
                    mr_addrs[mr_addr_norm] = []
                mr_addrs[mr_addr_norm].append({
                    'Account': mr_account,
                    'Address': mr_addr
                })

        # Find matches - unique per applicant address, using first applicant as representative
        potentials = []
        for norm_addr, app_list in applicant_addrs.items():
            if norm_addr in mr_addrs:
                mr_list = mr_addrs[norm_addr]
                # Use the first applicant as representative
                rep_app = app_list[0]
                for mr in mr_list:
                    if rep_app['Account'] != mr['Account']:
                        potentials.append({
                            'Applicant Account': rep_app['Account'],
                            'Applicant Address': rep_app['Address'],
                            'Matching Account': mr['Account'],
                            'Matching Address': mr['Address']
                        })

        potentials_df = pd.DataFrame(potentials)
        if not potentials_df.empty:
            potentials_df = potentials_df.sort_values(['Applicant Address', 'Matching Account'])

        return potentials_df, None
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

def get_file_status(file_path):
    if os.path.exists(file_path):
        size_mb = os.path.getsize(file_path) / (1024 * 1024)  # MB
        return f"✅ Exists ({size_mb:.1f} MB): {os.path.basename(file_path)}"
    return f"❌ Missing"

# User preference functions (server-side persistence)
def get_user_prefs_path():
    username = os.environ.get('REMOTE_USER', 'anonymous').strip().replace(' ', '_')
    prefs_dir = 'user_prefs'
    os.makedirs(prefs_dir, exist_ok=True)
    return os.path.join(prefs_dir, f"{username}_prefs.json")

def load_user_pref(key: str, default=None):
    prefs_path = get_user_prefs_path()
    if os.path.exists(prefs_path):
        with open(prefs_path, 'r') as f:
            prefs = json.load(f)
            return prefs.get(key, default)
    return default

def save_user_pref(key: str, value):
    prefs_path = get_user_prefs_path()
    prefs = {}
    if os.path.exists(prefs_path):
        with open(prefs_path, 'r') as f:
            prefs = json.load(f)
    prefs[key] = value
    with open(prefs_path, 'w') as f:
        json.dump(prefs, f)

# Auto-set session state for county
if 'last_county' not in st.session_state:
    st.session_state.last_county = county
save_user_pref('last_county', county)  # Persist for fallback

# Back to Home button (to subdomain index.html)
st.markdown(
    """
    <style>
    .back-to-home {
        text-decoration: none;
        display: inline-block;
        padding: 8px 16px;
        background-color: #3B82F6;
        color: white !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        border-radius: 6px;
        border: 1px solid #3B82F6;
        cursor: pointer;
        margin-bottom: 20px;
        transition: background-color 0.2s, border-color 0.2s, color 0.2s;
        text-shadow: 0 1px 2px rgba(0,0,0,0.1);
        opacity: 1 !important;
    }
    .back-to-home:hover {
        background-color: #2563EB;
        border-color: #2563EB;
        color: white !important;
        text-shadow: 0 1px 2px rgba(0,0,0,0.2);
    }
    </style>
    <a href="/" target="_self" rel="noopener noreferrer" class="back-to-home">
        ← Back to Tools
    </a>
    """,
    unsafe_allow_html=True
)

# Initialize session state
if 'applicant_bytes' not in st.session_state:
    st.session_state.applicant_bytes = None
if 'comparison_results' not in st.session_state:
    st.session_state.comparison_results = None
if 'mr_potentials' not in st.session_state:
    st.session_state.mr_potentials = pd.DataFrame()
if 'blacklist' not in st.session_state:
    st.session_state.blacklist = load_blacklist(county)
if 'master_uploaded' not in st.session_state:
    st.session_state.master_uploaded = os.path.exists(get_master_path(county))
if 'accounts_uploaded' not in st.session_state:
    st.session_state.accounts_uploaded = os.path.exists(get_accounts_path(county))
if 'clear_password' not in st.session_state:
    st.session_state.clear_password = ""

master_path = get_master_path(county)
accounts_path = get_accounts_path(county)

# Sidebar with county display
with st.sidebar:
    st.write(f"**Current County:** {county}")

# Tabs
tab1, tab2 = st.tabs(["Compare", "Settings"])

with tab1:
    st.subheader("Upload Applicant List and Compare")
    
    uploaded_applicant = st.file_uploader("Upload HO Applicant Excel", type=['xlsx', 'xls'])
    if uploaded_applicant is not None:
        st.session_state.applicant_bytes = uploaded_applicant.getvalue()
        st.success("Applicant file loaded!")
    
    if st.button("Compare") and st.session_state.applicant_bytes:
        with st.spinner("Comparing..."):
            common_all, error = compare_excels(st.session_state.applicant_bytes, master_path, st.session_state.blacklist)
            if error:
                st.error(error)
            else:
                st.session_state.comparison_results = common_all
            
            df1_orig = pd.read_excel(io.BytesIO(st.session_state.applicant_bytes), engine='openpyxl')
            mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
            if mr_error:
                st.error(mr_error)
            else:
                st.session_state.mr_potentials = mr_potentials
        
        st.rerun()
    
    if st.session_state.comparison_results is not None:
        df = st.session_state.comparison_results
        st.dataframe(df, use_container_width=True)
        
        # Download TXT
        txt_output = generate_txt_output(df)
        st.download_button(
            label="Download as TXT",
            data=txt_output,
            file_name=f"{county}_comparison.txt",
            mime="text/plain"
        )
    
    # Potential M/R Matches
    with st.expander("Potential M/R Address Matches", expanded=True):
        if not st.session_state.mr_potentials.empty:
            potentials_df = st.session_state.mr_potentials.copy()
            potentials_df['Select'] = False
            edited_df = st.data_editor(
                potentials_df,
                column_config={
                    "Select": st.column_config.CheckboxColumn(
                        "Select to Blacklist",
                        help="Check to add this matching account to blacklist",
                        default=False,
                    )
                },
                use_container_width=True,
                hide_index=False,
            )
            
            if st.button("Add Selected to Blacklist"):
                selected_rows = edited_df[edited_df['Select'] == True]
                selected_to_blacklist = []
                for _, row in selected_rows.iterrows():
                    app_addr_norm = normalize_address(row['Applicant Address'])
                    selected_to_blacklist.append({
                        'applicant_account': row['Applicant Account'],
                        'account': row['Matching Account'],
                        'applicant_address': row['Applicant Address'],
                        'norm_addr': app_addr_norm
                    })
                if selected_to_blacklist:
                    st.session_state.blacklist.extend(selected_to_blacklist)
                    save_blacklist(county, st.session_state.blacklist)
                    
                    # Re-run comparisons with updated blacklist
                    common_all, error = compare_excels(st.session_state.applicant_bytes, master_path, st.session_state.blacklist)
                    if not error:
                        st.session_state.comparison_results = common_all
                    
                    df1_orig = pd.read_excel(io.BytesIO(st.session_state.applicant_bytes), engine='openpyxl')
                    mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
                    if not mr_error:
                        st.session_state.mr_potentials = mr_potentials
                    
                    st.success(f"Added {len(selected_to_blacklist)} accounts to blacklist. Results updated.")
                    st.rerun()
                else:
                    st.warning("No accounts selected.")
        else:
            st.info("No potential M/R address matches found.")

    # Blacklist Management
    with st.expander("Blacklist Management", expanded=False):
        st.write(f"Current Blacklist ({len(st.session_state.blacklist)} entries):")
        if st.session_state.blacklist:
            blacklist_df = pd.DataFrame(st.session_state.blacklist)
            if 'applicant_account' in blacklist_df.columns:
                blacklist_display_df = blacklist_df[['applicant_account', 'account', 'applicant_address']].copy()
                blacklist_display_df.columns = ['Applicant Account', 'Matching Account', 'Address']
            else:
                # Fallback for old format
                blacklist_display_df = pd.DataFrame({
                    'Applicant Account': [''] * len(blacklist_df),
                    'Matching Account': blacklist_df['account'],
                    'Address': blacklist_df['applicant_address']
                })
            blacklist_display_df['Select'] = False
            edited_blacklist = st.data_editor(
                blacklist_display_df,
                column_config={
                    "Select": st.column_config.CheckboxColumn(
                        "Select to Remove",
                        help="Check to remove this entry from blacklist",
                        default=False,
                    )
                },
                use_container_width=True,
                hide_index=False,
            )
            
            if st.button("Remove Selected from Blacklist"):
                selected_rows = edited_blacklist[edited_blacklist['Select'] == True]
                if not selected_rows.empty:
                    indices_to_remove = []
                    for _, row in selected_rows.iterrows():
                        for idx, entry in enumerate(st.session_state.blacklist):
                            match_acc = entry.get('account') == row['Matching Account']
                            match_addr = entry.get('applicant_address') == row['Address']
                            if 'applicant_account' in entry:
                                match_app = entry.get('applicant_account') == row['Applicant Account']
                                if match_acc and match_addr and match_app:
                                    indices_to_remove.append(idx)
                                    break
                            else:
                                if match_acc and match_addr:
                                    indices_to_remove.append(idx)
                                    break
                    # Remove in reverse order
                    for i in sorted(indices_to_remove, reverse=True):
                        del st.session_state.blacklist[i]
                    save_blacklist(county, st.session_state.blacklist)
                    
                    # Re-run comparisons with updated blacklist
                    if st.session_state.applicant_bytes:
                        common_all, error = compare_excels(st.session_state.applicant_bytes, master_path, st.session_state.blacklist)
                        if not error:
                            st.session_state.comparison_results = common_all
                        
                        df1_orig = pd.read_excel(io.BytesIO(st.session_state.applicant_bytes), engine='openpyxl')
                        mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
                        if not mr_error:
                            st.session_state.mr_potentials = mr_potentials
                    
                    st.success(f"Removed {len(indices_to_remove)} entries from blacklist. Results updated.")
                    st.rerun()
                else:
                    st.warning("No entries selected.")
        else:
            st.info("Blacklist is empty.")

    if not os.path.exists(master_path) or not os.path.exists(accounts_path):
        st.warning("Please upload master and accounts lists in Settings tab to proceed.")

with tab2:
    st.subheader("Settings: Upload Persistent Files")
    with st.expander("Upload or Manage Files", expanded=True):
        col1, col2 = st.columns(2)
        
        with col1:
            st.write(f"**LTHO Master List**")
            
            # Master Status and Replace
            master_status = get_file_status(master_path)
            st.write(f"**Status:** {master_status}")
            uploaded_master = st.file_uploader("Replace Master List", type=['xlsx', 'xls'], key="master_upload")
            if uploaded_master is not None and st.button("Save Master List to Server", type="primary", key="save_master"):
                try:
                    with st.spinner("Saving master list..."):
                        df = pd.read_excel(uploaded_master, engine='openpyxl')
                        df.to_excel(master_path, index=False, engine='openpyxl')
                    st.success(f"Master list saved for {county} County!")
                    st.session_state.master_uploaded = True
                    # Re-run comparison if applicant loaded
                    if st.session_state.applicant_bytes:
                        common_all, error = compare_excels(st.session_state.applicant_bytes, master_path, st.session_state.blacklist)
                        if not error:
                            st.session_state.comparison_results = common_all
                        df1_orig = pd.read_excel(io.BytesIO(st.session_state.applicant_bytes), engine='openpyxl')
                        mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
                        if not mr_error:
                            st.session_state.mr_potentials = mr_potentials
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to save: {str(e)}")
            
            if st.button("Refresh Comparison (Reload Master)", type="secondary", key="refresh_master"):
                if st.session_state.applicant_bytes:
                    # Re-run comparison if applicant loaded
                    common_all, error = compare_excels(st.session_state.applicant_bytes, master_path, st.session_state.blacklist)
                    if not error:
                        st.session_state.comparison_results = common_all
                    df1_orig = pd.read_excel(io.BytesIO(st.session_state.applicant_bytes), engine='openpyxl')
                    mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
                    if not mr_error:
                        st.session_state.mr_potentials = mr_potentials
                st.rerun()
        
        with col2:
            st.write(f"**Master Accounts List**")
            
            # Accounts Status and Replace
            accounts_status = get_file_status(accounts_path)
            st.write(f"**Status:** {accounts_status}")
            uploaded_accounts = st.file_uploader("Replace Accounts List", type=['xlsx', 'xls'], key="accounts_upload")
            if uploaded_accounts is not None and st.button("Save Accounts List to Server", type="primary", key="save_accounts"):
                try:
                    with st.spinner("Saving accounts list..."):
                        df = pd.read_excel(uploaded_accounts, engine='openpyxl')
                        df.to_excel(accounts_path, index=False, engine='openpyxl')
                    st.success(f"Accounts list saved for {county} County!")
                    st.session_state.accounts_uploaded = True
                    # Re-run comparison if applicant loaded
                    if st.session_state.applicant_bytes:
                        common_all, error = compare_excels(st.session_state.applicant_bytes, master_path, st.session_state.blacklist)
                        if not error:
                            st.session_state.comparison_results = common_all
                        df1_orig = pd.read_excel(io.BytesIO(st.session_state.applicant_bytes), engine='openpyxl')
                        mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
                        if not mr_error:
                            st.session_state.mr_potentials = mr_potentials
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to save: {str(e)}")
            
            if st.button("Refresh Comparison (Reload Accounts)", type="secondary", key="refresh_accounts"):
                if st.session_state.applicant_bytes:
                    # Re-run comparison if applicant loaded
                    common_all, error = compare_excels(st.session_state.applicant_bytes, master_path, st.session_state.blacklist)
                    if not error:
                        st.session_state.comparison_results = common_all
                    df1_orig = pd.read_excel(io.BytesIO(st.session_state.applicant_bytes), engine='openpyxl')
                    mr_potentials, mr_error = compare_addresses(df1_orig, accounts_path, st.session_state.blacklist)
                    if not mr_error:
                        st.session_state.mr_potentials = mr_potentials
                st.rerun()

    # Check file status
    st.subheader("File Status")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Master List:** {get_file_status(master_path)}")
    with col2:
        st.write(f"**Accounts List:** {get_file_status(accounts_path)}")

# Sidebar: Info/Reset (with collapsible content and protected clear button)
with st.sidebar:
    with st.expander("Instructions & Reset", expanded=False):
        st.header("Instructions")
        st.markdown("""
        - County is auto-detected from your subdomain.
        - Go to Settings tab to upload the Master List and Accounts List for your county (persist on server).
        - Back to Compare tab: Upload applicant list, click Compare to query and view matches.
        - Use Blacklist Management in Compare tab to add/remove accounts to ignore in future comparisons.
        - Files are stored server-side per county for reuse.
        """)
        
        # Protected Clear Session button
        st.subheader("Reset Session")
        if 'clear_password' not in st.session_state:
            st.session_state.clear_password = ""
        
        clear_password = st.text_input("Enter password to confirm:", type="password", key="clear_pwd_input")
        st.session_state.clear_password = clear_password
        
        if st.button("Clear Session (Forget County)", disabled=not clear_password):
            if clear_password == "admin":  # Change this to your desired password
                save_user_pref('last_county', None)  # Clear user pref
                st.query_params.clear()  # Clears URL params
                for key in list(st.session_state.keys()):
                    if key != 'clear_password':  # Preserve password input state
                        del st.session_state[key]
                st.session_state.last_county = county  # Reset to detected
                st.success("Session cleared! Reloading...")
                st.rerun()
            else:
                st.error("Incorrect password. Try again.")
                st.session_state.clear_password = ""  # Clear input on error