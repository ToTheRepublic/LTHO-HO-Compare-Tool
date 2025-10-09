import streamlit as st
import pandas as pd
import os
import re
import io
import fitz  # PyMuPDF
import json
import time
from datetime import datetime
import base64
import urllib.parse
from streamlit_pdf_viewer import pdf_viewer
import streamlit.components.v1 as components
from typing import Optional

# Wyoming counties list
WY_COUNTIES = [
    "Albany", "Big Horn", "Campbell", "Carbon", "Converse", "Crook", "Fremont", "Goshen",
    "Hot Springs", "Johnson", "Laramie", "Lincoln", "Natrona", "Niobrara", "Park", "Platte",
    "Sheridan", "Sublette", "Sweetwater", "Teton", "Uinta", "Washakie", "Weston"
]

# Document types
DOC_TYPES = ["Notice of Value", "Declaration", "Tax Notice"]

# Base directory for county data
BASE_DIR = "county_docs"
os.makedirs(BASE_DIR, exist_ok=True)

def get_file_status(county_dir, doc_type, extension):
    file_path = get_doc_path(county_dir, doc_type, extension)
    if os.path.exists(file_path):
        size_mb = os.path.getsize(file_path) / (1024 * 1024)  # MB
        return f"✅ Exists ({size_mb:.1f} MB): {os.path.basename(file_path)}"
    return f"❌ Missing: {doc_type}.{extension}"

def get_county_path(county):
    county_dir = os.path.join(BASE_DIR, county.replace(" ", "_"))
    os.makedirs(county_dir, exist_ok=True)
    return county_dir

def get_doc_path(county_dir, doc_type, extension):
    return os.path.join(county_dir, f"{doc_type.replace(' ', '_').lower()}.{extension}")

def extract_nov_info(text):
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    account = ""
    local_number = ""

    normalized_lines = [re.sub(r'\s+', ' ', line).strip() for line in lines]

    account_pattern = re.compile(r'[RMPO]000\d{4,5}', re.I)
    account_index = -1
    for i, line in enumerate(normalized_lines):
        match = account_pattern.search(line)
        if match:
            account = match.group().upper()
            account_index = i
            break

    if account_index != -1 and account_index + 1 < len(normalized_lines):
        local_number_candidate = normalized_lines[account_index + 1].strip()
        if re.match(r'^\d{4,6}$', local_number_candidate):
            local_number = local_number_candidate.lstrip('0').zfill(4)

    return account, local_number

def extract_declaration_info(text):
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    account = ""
    local_number = ""

    acc_pattern = re.compile(r'[RMPO]000\d{4,5}', re.I)
    for line in lines:
        acc_match = acc_pattern.search(line)
        if acc_match:
            account = acc_match.group().upper()
            break

    for i, line in enumerate(lines):
        if "January 1, 2025" in line:
            if i + 1 < len(lines) and re.match(r'^\d{4}$', lines[i + 1]):
                local_number = lines[i + 1]
                break

    return account, local_number

def extract_tax_notice_info(text):
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    account = ""
    local_number = ""

    for line in lines:
        if "LOCAL/REALWARE ID #" in line:
            id_match = re.search(r'LOCAL/REALWARE ID #\s*(\d+)/([RMPO]000\d{4,5})', line, re.I)
            if id_match:
                local_number = id_match.group(1).lstrip('0').zfill(4)
                account = id_match.group(2).upper()
            break

    return account, local_number

def extract_info_from_text(text, search_type):
    if search_type == "Notice of Value":
        return extract_nov_info(text)
    elif search_type == "Declaration":
        return extract_declaration_info(text)
    elif search_type == "Tax Notice":
        return extract_tax_notice_info(text)
    return "", ""

@st.cache_data
def index_pdf(pdf_path, excel_path, search_type):
    index_data = {}
    first_page = {}
    debug_accounts = ["R0007425", "P0007419"]

    excel_df = None
    if pd is not None and excel_path and os.path.isfile(excel_path):
        try:
            excel_df = pd.read_excel(excel_path, engine='openpyxl')
            required_columns = ['ACCOUNTNO', 'NAME1', 'BUSINESSNAME', 'PREDIRECTION', 'STREETNO', 'POSTDIRECTION', 'STREETNAME', 'STREETTYPE']
            if all(col in excel_df.columns for col in required_columns):
                excel_df.set_index('ACCOUNTNO', inplace=True)
            else:
                excel_df = None
        except:
            excel_df = None

    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        for page_num in range(total_pages):
            text = doc[page_num].get_text()
            if not text:
                continue
            account, local_number = extract_info_from_text(text, search_type)
            
            if account in debug_accounts:
                st.write(f"Debug for {account} on page {page_num + 1}")

            if account:
                ownership_name = ""
                property_address = ""
                business_name = ""
                if excel_df is not None and account in excel_df.index:
                    row = excel_df.loc[account]
                    ownership_name = str(row.get('NAME1', '')) if pd.notna(row.get('NAME1')) else ""
                    business_name = str(row.get('BUSINESSNAME', '')) if pd.notna(row.get('BUSINESSNAME')) else ""
                    address_parts = [
                        str(row.get('PREDIRECTION', '')) if pd.notna(row.get('PREDIRECTION')) else "",
                        str(row.get('STREETNO', '')) if pd.notna(row.get('STREETNO')) else "",
                        str(row.get('POSTDIRECTION', '')) if pd.notna(row.get('POSTDIRECTION')) else "",
                        str(row.get('STREETNAME', '')) if pd.notna(row.get('STREETNAME')) else "",
                        str(row.get('STREETTYPE', '')) if pd.notna(row.get('STREETTYPE')) else ""
                    ]
                    property_address = ' '.join(part for part in address_parts if part)
                    excel_local_number = str(row.get('Local Number', '')) if pd.notna(row.get('Local Number')) else ""
                    if excel_local_number and re.match(r'^\d{4,6}$', excel_local_number):
                        local_number = excel_local_number.lstrip('0').zfill(4)

                if account not in index_data:
                    index_data[account] = {
                        "local_number": local_number,
                        "business_name": business_name,
                        "address": property_address,
                        "ownership_name": ownership_name,
                        "pages": [page_num + 1]
                    }
                    first_page[account] = page_num + 1
                else:
                    index_data[account]["pages"].append(page_num + 1)
                    if page_num + 1 == first_page[account]:
                        if not index_data[account]["business_name"] and business_name:
                            index_data[account]["business_name"] = business_name
                        if not index_data[account]["address"] and property_address:
                            index_data[account]["address"] = property_address
                        if not index_data[account]["ownership_name"] and ownership_name:
                            index_data[account]["ownership_name"] = ownership_name
        doc.close()
    except Exception as e:
        st.error(f"Error indexing: {str(e)}")
    return index_data

def save_index(county_dir, search_type, index_data):
    index_file = get_doc_path(county_dir, search_type, "json")
    with open(index_file, 'w', encoding='utf-8') as f:
        json.dump(index_data, f, indent=4)

def load_index(county_dir, search_type):
    index_file = get_doc_path(county_dir, search_type, "json")
    if os.path.exists(index_file):
        with open(index_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def search_matches(index_data, query, search_type):
    query_lower = query.lower().strip()
    results = []

    # Exact account match
    if re.match(r'^[RMPO]000\d{4,5}$', query, re.I):
        q_upper = query.upper()
        if q_upper in index_data:
            data = index_data[q_upper]
            results.append({
                'acc': q_upper,
                'local_number': data.get("local_number", "").lstrip('0'),
                'ownership_name': data.get("ownership_name", ""),
                'address': data.get("address", ""),
                'business_name': data.get("business_name", ""),
                'pages': data['pages']
            })
    # Exact local number match
    elif re.match(r'^\d{4,}$', query):
        normalized_query = query.lstrip('0')
        for acc, data in index_data.items():
            local_number = data.get("local_number", "").lstrip('0')
            if normalized_query == local_number:
                results.append({
                    'acc': acc,
                    'local_number': local_number,
                    'ownership_name': data.get("ownership_name", ""),
                    'address': data.get("address", ""),
                    'business_name': data.get("business_name", ""),
                    'pages': data['pages']
                })
    # Partial name/address match
    else:
        for acc, data in index_data.items():
            ownership_name = data.get("ownership_name", "").lower()
            business_name = data.get("business_name", "").lower()
            address = data.get("address", "").lower()
            if (query_lower in ownership_name or 
                query_lower in business_name or 
                query_lower in address):
                results.append({
                    'acc': acc,
                    'local_number': data.get("local_number", "").lstrip('0'),
                    'ownership_name': data.get("ownership_name", ""),
                    'address': data.get("address", ""),
                    'business_name': data.get("business_name", ""),
                    'pages': data['pages']
                })
    return results

def get_business_name(res):
    return res.get('business_name', '') or 'N/A'

def get_ownership_name(res):
    return res.get('ownership_name', '') or 'N/A'

def get_address_from_index(res):
    return res.get('address', '') or 'N/A'

def extract_pdf(pdf_path, selected_res):
    try:
        doc = fitz.open(pdf_path)
        pages = selected_res['pages']
        output = fitz.open()
        for page_num in sorted(pages):
            page = doc[page_num - 1]  # 1-based to 0-based
            output.insert_pdf(doc, from_page=page.number, to_page=page.number)
        doc.close()
        output_bytes = io.BytesIO()
        output.save(output_bytes, garbage=4, deflate=True, clean=True)
        output.close()
        return output_bytes
    except Exception as e:
        return (None, f"Error extracting PDF: {str(e)}")

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

# Helper to get county from query params (for sharing)
def get_persistent_county() -> Optional[str]:
    query_params = st.query_params
    county_param = query_params.get("county", [None])[0]
    if county_param and county_param in WY_COUNTIES:
        return county_param
    return None

# Page config
st.set_page_config(page_title="WY County Document Search", layout="wide")

# Back to Home button (styled, same tab) - Fixed hover with CSS
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
        text-shadow: 0 1px 2px rgba(0,0,0,0.1);  /* Subtle shadow for readability */
        opacity: 1 !important;  /* Prevent fading */
    }
    .back-to-home:hover {
        background-color: #2563EB;
        border-color: #2563EB;
        color: white !important;
        text-shadow: 0 1px 2px rgba(0,0,0,0.2);  /* Slightly stronger on hover */
    }
    </style>
    <a href="https://assessortools.com" target="_self" rel="noopener noreferrer" class="back-to-home">
        ← Back to Home
    </a>
    """,
    unsafe_allow_html=True
)

# Initialize session state
if 'last_county' not in st.session_state:
    st.session_state.last_county = None
if 'docs_indexed' not in st.session_state:
    st.session_state.docs_indexed = {}
if 'search_results' not in st.session_state:
    st.session_state.search_results = None
if 'selected_res' not in st.session_state:
    st.session_state.selected_res = None
if 'clear_password' not in st.session_state:
    st.session_state.clear_password = ""

# Determine default county: user pref > URL param > logged-in county > first
user_pref_county = load_user_pref('last_county')
url_county = get_persistent_county()
logged_in_county = os.environ.get('REMOTE_USER', '').strip()
default_county = user_pref_county if user_pref_county in WY_COUNTIES else \
                url_county if url_county else \
                logged_in_county if logged_in_county in WY_COUNTIES else WY_COUNTIES[0]

# County selection (simple dropdown, persists in session)
st.subheader("Select Your County")
default_index = WY_COUNTIES.index(default_county)
county = st.selectbox("Choose a county:", WY_COUNTIES, index=default_index)
if county != st.session_state.last_county:
    st.session_state.last_county = county
    save_user_pref('last_county', county)  # Save to server-side prefs
    if url_county != county:  # Update URL only if different (for sharing)
        st.query_params["county"] = [county]
    st.session_state.docs_indexed = {}  # Reset indexing on county change
    st.session_state.search_results = None
    st.session_state.selected_res = None
    st.rerun()

if not county:
    st.warning("Please select a county to proceed.")
    st.stop()

st.title(f"WY County Document Search - {county} County")
county_dir = get_county_path(county)

# Sidebar with county display
with st.sidebar:
    st.write(f"**Current County:** {county}")

# Auto-load indexed status from disk
if county and county_dir:
    for doc_type in DOC_TYPES:
        index_file = get_doc_path(county_dir, doc_type, "json")
        if doc_type not in st.session_state.docs_indexed:
            st.session_state.docs_indexed[doc_type] = os.path.exists(index_file)

# Refresh indexed status if needed
if county and county_dir:
    for doc_type in DOC_TYPES:
        index_file = get_doc_path(county_dir, doc_type, "json")
        st.session_state.docs_indexed[doc_type] = os.path.exists(index_file)

# Sidebar: Instructions & Reset (with collapsible content and protected clear button)
with st.sidebar:
    with st.expander("Instructions & Reset", expanded=False):
        st.header("Instructions")
        st.markdown("""
        - Select your county above (remembers your last choice for next time).
        - Go to Settings tab to upload the 3 PDFs and 3 Excel files for your county.
        - Click "Index" for each document type in Settings.
        - Back to Search tab: Enter query and hit Enter or click Search to query and select from matches to download extracted PDFs.
        - Files are stored server-side per county for reuse.
        """)
        
        # Protected Clear Session button
        st.subheader("Reset Session")
        clear_password = st.text_input("Enter password to confirm:", type="password", value=st.session_state.clear_password, key="clear_pwd_input")
        st.session_state.clear_password = clear_password
        
        if st.button("Clear Session (Forget County)", disabled=not clear_password):
            if clear_password == "reset123":  # Change this to your desired password
                save_user_pref('last_county', None)  # Clear user pref
                st.query_params.clear()  # Clears URL params
                for key in list(st.session_state.keys()):
                    if key != 'clear_password':  # Preserve password input state
                        del st.session_state[key]
                st.session_state.last_county = None
                st.success("Session cleared! Reloading...")
                st.rerun()
            else:
                st.error("Incorrect password. Try again.")
                st.session_state.clear_password = ""  # Clear input on error

# Tabs
tab1, tab2 = st.tabs(["Search", "Settings"])

with tab1:
    st.subheader("Search Documents")
    
    if all(st.session_state.docs_indexed.get(doc_type, False) for doc_type in DOC_TYPES):
        with st.form("search_form"):
            type_var = st.selectbox("Document Type:", DOC_TYPES, key="doc_type")
            query = st.text_input("Search (Account/Local/Name/Address):", key="search_query", placeholder="e.g., R0001234 or 1234 or 'Smith' or 'Main St'")
            submitted = st.form_submit_button("Search Matches")

        # Define pdf_path here so it's always available (uses current type_var)
        pdf_path = get_doc_path(county_dir, type_var, "pdf")
        if not os.path.exists(pdf_path):
            st.warning("PDF not found. Please upload in Settings.")

        if submitted:
            index_data = load_index(county_dir, type_var)
            with st.spinner("Searching..."):
                results = search_matches(index_data, query, type_var)
                if not results:
                    st.error("No matches found.")
                    st.session_state.search_results = None
                else:
                    st.success(f"Found {len(results)} match(es).")
                    st.session_state.search_results = results
                    st.session_state.selected_res = None  # Reset selection
            st.rerun()

        # Display results as radio list if available
        if st.session_state.search_results:
            results = st.session_state.search_results
            display_options = [f"{r['acc']} - {r['ownership_name'][:30]}{'...' if len(r['ownership_name']) > 30 else ''} ({r['address'][:20]}{'...' if len(r['address']) > 20 else ''})" for r in results]
            selected_idx = st.radio("Select a match to extract:", range(len(display_options)), format_func=lambda idx: display_options[idx], key="match_radio")
            selected_res = results[selected_idx]
            st.session_state.selected_res = selected_res

            # Show details of selected
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.write("**Account:**")
                st.write(f"{selected_res['acc']} (Local: {selected_res['local_number']})")
            with col2:
                st.write("**Business Name:**")
                st.write(get_business_name(selected_res))
            with col3:
                st.write("**Ownership Name:**")
                st.write(get_ownership_name(selected_res))
            with col4:
                st.write("**Address:**")
                st.write(get_address_from_index(selected_res))

            # Extract and download button (single button, inside the if)
            if st.button("Extract Selected PDF", key="extract_pdf"):
                pdf_bytes = extract_pdf(pdf_path, selected_res)
                if isinstance(pdf_bytes, tuple):  # Error case
                    st.error(pdf_bytes[1])
                else:
                    pdf_data = pdf_bytes.getvalue()
                    st.download_button(
                        label="Download Extracted PDF",
                        data=pdf_data,
                        file_name=f"{county}_{type_var}_{selected_res['acc']}.pdf",
                        mime="application/pdf"
                    )

                    # Inline PDF Viewer
                    st.markdown("### Full PDF Preview:")
                    try:
                        pdf_viewer(pdf_data, height=800)
                    except Exception as e:
                        st.warning(f"Could not render PDF viewer: {e}. Falling back to first-page image preview.")
                        # Fallback image code
                        doc = fitz.open(stream=pdf_data, filetype="pdf")
                        if len(doc) > 0:
                            page = doc.load_page(0)
                            mat = fitz.Matrix(2, 2)
                            pix = page.get_pixmap(matrix=mat)
                            img_bytes = pix.tobytes("png")
                            st.image(img_bytes, caption=f"Preview of {selected_res['acc']} - Page 1", width='stretch')
                        doc.close()

    else:
        st.warning("Please index all document types in Settings before searching.")

with tab2:
    st.subheader("Settings: Upload and Index Documents")
    with st.expander("Upload or Manage Files", expanded=True):
        col1, col2, col3 = st.columns(3)
        for i, doc_type in enumerate(DOC_TYPES):
            col = [col1, col2, col3][i]
            with col:
                st.write(f"**{doc_type}**")
                
                # PDF Status and Replace
                pdf_status = get_file_status(county_dir, doc_type, "pdf")
                st.write(f"**PDF:** {pdf_status}")
                uploaded_pdf = st.file_uploader(f"Replace {doc_type} PDF", type=['pdf'], key=f"{doc_type.replace(' ', '_').lower()}_pdf_replace_{county}")
                if uploaded_pdf is not None:
                    pdf_path = get_doc_path(county_dir, doc_type, "pdf")
                    with open(pdf_path, "wb") as f:
                        f.write(uploaded_pdf.getbuffer())
                    st.success(f"{doc_type} PDF replaced!")
                    st.session_state.docs_indexed[doc_type] = False  # Mark as needs re-index
                    st.rerun()
                
                # Excel Status and Replace
                excel_status = get_file_status(county_dir, doc_type, "xlsx")
                st.write(f"**Excel:** {excel_status}")
                uploaded_excel = st.file_uploader(f"Replace {doc_type} Excel", type=['xlsx', 'xls'], key=f"{doc_type.replace(' ', '_').lower()}_excel_replace_{county}")
                if uploaded_excel is not None:
                    excel_path = get_doc_path(county_dir, doc_type, "xlsx")
                    with open(excel_path, "wb") as f:
                        f.write(uploaded_excel.getbuffer())
                    st.success(f"{doc_type} Excel replaced!")
                    st.session_state.docs_indexed[doc_type] = False  # Mark as needs re-index
                    st.rerun()
                
                # Index/Re-Index Button
                index_text = "Re-Index" if st.session_state.docs_indexed.get(doc_type, False) else "Index"
                if st.button(f"{index_text} {doc_type}", key=f"index_{doc_type}_{county}"):
                    pdf_path = get_doc_path(county_dir, doc_type, "pdf")
                    excel_path = get_doc_path(county_dir, doc_type, "xlsx")
                    if os.path.exists(pdf_path):
                        with st.spinner(f"Indexing {doc_type}..."):
                            index_data = index_pdf(pdf_path, excel_path if os.path.exists(excel_path) else None, doc_type)
                            save_index(county_dir, doc_type, index_data)
                            st.session_state.docs_indexed[doc_type] = True
                            st.success(f"{doc_type} indexed successfully!")
                            st.rerun()
                    else:
                        st.warning(f"Please upload {doc_type} PDF first.")

    # Check indexing status
    st.subheader("Indexing Status")
    for doc_type in DOC_TYPES:
        index_file = get_doc_path(county_dir, doc_type, "json")
        status = "✅ Indexed" if os.path.exists(index_file) else "❌ Not Indexed"
        st.write(f"{doc_type}: {status}")