import streamlit as st
import streamlit_javascript as st_js  # New import for JS detection
import pandas as pd
import os
import re
import io
import fitz  # PyMuPDF
import json
import time
import shutil
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

import time
st.write(f"Rerun at {time.strftime('%H:%M:%S')}")

# Detect subdomain via JS and set county (no cache - called once per run)
def detect_county():
    try:
        import streamlit.runtime as runtime  # Ensure import (add if missing)
        session_mgr = runtime.get_instance()._session_mgr
        active_sessions = session_mgr.list_active_sessions()
        if not active_sessions:
            raise ValueError("No active session found")
        
        # Get the first (typically only) active session's request
        request = active_sessions[0].client.request
        host = request.host.lower().strip()  # e.g., 'laramie.assessortools.com'
        
        if not host:
            raise ValueError("No host available in request")
        
        if 'assessortools.com' not in host:
            raise ValueError(f"Invalid host '{host}' (must include 'assessortools.com')")
        
        subdomain = host.split('.')[0]
        county = SUBDOMAIN_TO_COUNTY.get(subdomain)
        if not county:
            raise ValueError(f"Subdomain '{subdomain}' from host '{host}' not mapped to a Wyoming county")
        
        return county
    except Exception as e:
        st.error(f"County detection failed: {str(e)}")
        st.info("Please access the app via a valid subdomain (e.g., https://laramie.assessortools.com).")
        return None

# county = detect_county()

# Early session state init for county (avoids flash)
if 'detected_county' not in st.session_state:
    st.session_state.detected_county = None
#if 'uploading' not in st.session_state:  # Add lock/flag for upload
#    st.session_state.uploading = False

# Detect and store county
if st.session_state.detected_county is None:
    st.session_state.detected_county = detect_county()
    st.rerun()  # Immediate rerun to apply county-specific title/config

county = st.session_state.detected_county

if county is None:
    st.stop()  # Or st.error("No county detected—exiting.") if you prefer a message before stop

# Now set county-specific title/config on rerun (overrides placeholder)
if county:
    st.set_page_config(page_title=f"Document Search Tool - {county} County", layout="wide")
    st.title(f"{county} County Document Search Tool")
else:
    st.title("Document Search Tool (No County)")

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

    account_pattern = re.compile(r'[RMPO]\d{7}', re.I)
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

    acc_pattern = re.compile(r'[RMPO]\d{7}', re.I)
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
            id_match = re.search(r'LOCAL/REALWARE ID #\s*(\d+)/([RMPO]\d{7})', line, re.I)
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
    """Cached version for backward compatibility"""
    return index_pdf_with_progress(pdf_path, excel_path, search_type, None)

def index_pdf_with_progress(pdf_path, excel_path, search_type, progress_bar=None):
    """Index PDF with optional progress bar display"""
    index_data = {}
    first_page = {}
    debug_accounts = []

    excel_df = None
    if pd is not None and excel_path and os.path.isfile(excel_path):
        try:
            excel_df = pd.read_excel(excel_path, engine='openpyxl')
            required_columns = ['ACCOUNTNO', 'NAME1', 'ADDRESS']
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
            try:
                # Update progress bar if provided
                if progress_bar is not None:
                    progress = (page_num + 1) / total_pages
                    progress_bar.progress(progress, text=f"Processing page {page_num + 1} of {total_pages}")
                
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
                    local_number = ""
                    if excel_df is not None and account in excel_df.index:
                        try:
                            row = excel_df.loc[account]
                            # Handle case where there are duplicate account numbers (returns Series)
                            if isinstance(row, pd.Series):
                                # Single row found
                                ownership_name = str(row.get('NAME1', '')) if pd.notna(row.get('NAME1')) else ""
                                property_address = str(row.get('ADDRESS', '')) if pd.notna(row.get('ADDRESS')) else ""
                                business_name = str(row.get('BUSINESSNAME', '')) if pd.notna(row.get('BUSINESSNAME')) else ""
                                excel_local_number = str(row.get('Local Number', '')) if pd.notna(row.get('Local Number')) else ""
                            else:
                                # Multiple rows found (DataFrame), take the first one
                                first_row = row.iloc[0]
                                ownership_name = str(first_row.get('NAME1', '')) if pd.notna(first_row.get('NAME1')) else ""
                                property_address = str(first_row.get('ADDRESS', '')) if pd.notna(first_row.get('ADDRESS')) else ""
                                business_name = str(first_row.get('BUSINESSNAME', '')) if pd.notna(first_row.get('BUSINESSNAME')) else ""
                                excel_local_number = str(first_row.get('Local Number', '')) if pd.notna(first_row.get('Local Number')) else ""
                            
                            if excel_local_number and re.match(r'^\d{4,6}$', excel_local_number):
                                local_number = excel_local_number.lstrip('0').zfill(4)
                        except Exception as excel_error:
                            # If there's any issue with Excel lookup, continue without it
                            ownership_name = ""
                            property_address = ""
                            business_name = ""

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
            except Exception as page_error:
                # Continue processing even if a single page fails
                if progress_bar is not None:
                    progress = (page_num + 1) / total_pages
                    progress_bar.progress(progress, text=f"Error on page {page_num + 1}, continuing... ({str(page_error)[:50]})")
                # Log the error for debugging (you can remove this line if not needed)
                print(f"Error processing page {page_num + 1}: {str(page_error)}")
                continue
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
    if re.match(r'^[RMPO]\d{7}$', query, re.I):
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
    <a href="/" target="_self" rel="noopener noreferrer" class="back-to-home">
        ← Back to Tools
    </a>
    """,
    unsafe_allow_html=True
)

# Initialize session state
if 'docs_indexed' not in st.session_state:
    st.session_state.docs_indexed = {}
if 'search_results' not in st.session_state:
    st.session_state.search_results = None
if 'selected_res' not in st.session_state:
    st.session_state.selected_res = None
if 'clear_password' not in st.session_state:
    st.session_state.clear_password = ""

st.title(f"Document Search Tool - {county} County")
county_dir = get_county_path(county)

# Sidebar with county display
with st.sidebar:
    if county:
        st.write(f"**Current County:** {county}")
    else:
        st.error("**No County Detected**")

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
        - County is auto-detected from your subdomain.
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
                for key in list(st.session_state.keys()):
                    if key != 'clear_password':  # Preserve password input state
                        del st.session_state[key]
                st.session_state.last_county = county  # Reset to detected
                st.success("Session cleared! Reloading...")
                st.rerun()
            else:
                st.error("Incorrect password. Try again.")
                st.session_state.clear_password = ""  # Clear input on error

# Tabs
tab1, tab2 = st.tabs(["Search", "Settings"])

with tab1:
    st.subheader("Search Documents")
    
    # Get list of indexed document types
    indexed_doc_types = [doc_type for doc_type in DOC_TYPES if st.session_state.docs_indexed.get(doc_type, False)]
    
    # Show status of indexed documents
    if indexed_doc_types:
        st.success(f"✅ **Available document types:** {', '.join(indexed_doc_types)}")
        not_indexed = [doc_type for doc_type in DOC_TYPES if not st.session_state.docs_indexed.get(doc_type, False)]
        if not_indexed:
            st.info(f"💭 **Not yet indexed:** {', '.join(not_indexed)} (go to Settings to add these)")
    
    if indexed_doc_types:
        with st.form("search_form"):
            type_var = st.selectbox("Document Type:", indexed_doc_types, key="doc_type")
            query = st.text_input("Search (Account/Local/Name/Address):", key="search_query", placeholder="Minimum 3 characters. e.g., R0001234 or 1234 or 'Smith' or 'Main St'")
            submitted = st.form_submit_button("Search Matches")

        # Define pdf_path here so it's always available (uses current type_var)
        pdf_path = get_doc_path(county_dir, type_var, "pdf")
        if not os.path.exists(pdf_path):
            st.warning("PDF not found. Please upload in Settings.")

        if submitted:
            # Validate minimum search length
            if len(query.strip()) < 3:
                st.error("Please enter at least 3 characters to search.")
            else:
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

                    # Inline PDF Viewer with dynamic height
                    st.markdown("### Full PDF Preview:")
                    try:
                        # Calc total height to fit content (no inner scroll)
                        doc = fitz.open(stream=pdf_data, filetype="pdf")
                        target_width = 800  # px; adjust for your layout
                        total_height = 0
                        for page in doc:
                            rect = page.rect
                            if rect.width > 0:  # Avoid div0
                                scaled_height = rect.height * (target_width / rect.width)
                                total_height += scaled_height + 20  # +20px margin between pages
                        doc.close()
                        
                        # Cap at viewport-friendly max (e.g., 1200px) to avoid overflow; ensure int
                        viewer_height = int(min(max(total_height, 200), 1200))  # Floor 200px, int-cast
                        
                        pdf_viewer(pdf_data, height=viewer_height, width=target_width)
                    except Exception as e:
                        st.warning(f"Could not render PDF viewer: {e}. Falling back to first-page image preview.")
                        # Fallback image code (unchanged)
                        doc = fitz.open(stream=pdf_data, filetype="pdf")
                        if len(doc) > 0:
                            page = doc.load_page(0)
                            mat = fitz.Matrix(2, 2)
                            pix = page.get_pixmap(matrix=mat)
                            img_bytes = pix.tobytes("png")
                            st.image(img_bytes, caption=f"Preview of {selected_res['acc']} - Page 1", width='stretch')
                        doc.close()

    else:
        st.info("📋 **No indexed documents available**")
        st.markdown("To get started:")
        st.markdown("1. Go to the **Settings** tab")
        st.markdown("2. Upload PDF and Excel files for any document type you want to search") 
        st.markdown("3. Click the **Index** button for that document type")
        st.markdown("4. Return here to search your indexed documents")
        
        # Show which document types are available but not indexed
        not_indexed = [doc_type for doc_type in DOC_TYPES if not st.session_state.docs_indexed.get(doc_type, False)]
        if not_indexed:
            st.markdown(f"**Available document types to index:** {', '.join(not_indexed)}")

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
                
                # PDF Upload
                upload_status = st.empty()  # For dynamic feedback
                uploaded_pdf = st.file_uploader(f"Replace {doc_type} PDF", type=['pdf'], key=f"{doc_type.replace(' ', '_').lower()}_pdf_replace_{county}")
                if uploaded_pdf is not None:
                    upload_status.text("Uploading...")
                    try:
                        pdf_path = get_doc_path(county_dir, doc_type, "pdf")
                        with open(pdf_path, "wb") as f:
                            shutil.copyfileobj(uploaded_pdf, f, length=1024 * 1024)
                        upload_status.success(f"{doc_type} PDF replaced!")
                        st.session_state.docs_indexed[doc_type] = False  # Mark as needs re-index
                    except Exception as e:
                        upload_status.error(f"Upload failed: {str(e)}")
                        with open("/tmp/streamlit_upload_error.log", "a") as logf:
                            logf.write(f"{datetime.now()}: {str(e)}\n")
                
                # PDF Delete Button with confirmation
                pdf_path = get_doc_path(county_dir, doc_type, "pdf")
                if os.path.exists(pdf_path):
                    # Add confirmation state to session state if not exists
                    confirm_key = f"confirm_delete_pdf_{doc_type}_{county}"
                    if confirm_key not in st.session_state:
                        st.session_state[confirm_key] = False
                    
                    if not st.session_state[confirm_key]:
                        if st.button(f"🗑️ Delete {doc_type} PDF", key=f"delete_pdf_{doc_type}_{county}", help="Permanently delete this PDF file"):
                            st.session_state[confirm_key] = True
                            st.rerun()
                    else:
                        st.warning(f"⚠️ Are you sure you want to delete the {doc_type} PDF? This cannot be undone!")
                        col_confirm1, col_confirm2 = st.columns(2)
                        with col_confirm1:
                            if st.button(f"✅ Yes, Delete", key=f"confirm_yes_pdf_{doc_type}_{county}"):
                                try:
                                    os.remove(pdf_path)
                                    st.session_state.docs_indexed[doc_type] = False  # Mark as needs re-index
                                    st.session_state[confirm_key] = False  # Reset confirmation
                                    st.success(f"{doc_type} PDF deleted successfully!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Failed to delete PDF: {str(e)}")
                        with col_confirm2:
                            if st.button(f"❌ Cancel", key=f"confirm_no_pdf_{doc_type}_{county}"):
                                st.session_state[confirm_key] = False
                                st.rerun()
                
                # Excel Status and Replace
                excel_status = get_file_status(county_dir, doc_type, "xlsx")
                st.write(f"**Excel:** {excel_status}")
                
                # Excel Upload
                upload_status_excel = st.empty()  # For dynamic feedback
                uploaded_excel = st.file_uploader(f"Replace {doc_type} Excel", type=['xlsx', 'xls'], key=f"{doc_type.replace(' ', '_').lower()}_excel_replace_{county}")
                if uploaded_excel is not None:
                    upload_status_excel.text("Uploading...")
                    try:
                        excel_path = get_doc_path(county_dir, doc_type, "xlsx")
                        with open(excel_path, "wb") as f:
                            shutil.copyfileobj(uploaded_excel, f, length=1024 * 1024)
                        upload_status_excel.success(f"{doc_type} Excel replaced!")
                        st.session_state.docs_indexed[doc_type] = False  # Mark as needs re-index
                    except Exception as e:
                        upload_status_excel.error(f"Upload failed: {str(e)}")
                        with open("/tmp/streamlit_upload_error.log", "a") as logf:
                            logf.write(f"{datetime.now()}: {str(e)}\n")
                
                # Excel Delete Button with confirmation
                excel_path = get_doc_path(county_dir, doc_type, "xlsx")
                if os.path.exists(excel_path):
                    # Add confirmation state to session state if not exists
                    confirm_key = f"confirm_delete_excel_{doc_type}_{county}"
                    if confirm_key not in st.session_state:
                        st.session_state[confirm_key] = False
                    
                    if not st.session_state[confirm_key]:
                        if st.button(f"🗑️ Delete {doc_type} Excel", key=f"delete_excel_{doc_type}_{county}", help="Permanently delete this Excel file"):
                            st.session_state[confirm_key] = True
                            st.rerun()
                    else:
                        st.warning(f"⚠️ Are you sure you want to delete the {doc_type} Excel? This cannot be undone!")
                        col_confirm1, col_confirm2 = st.columns(2)
                        with col_confirm1:
                            if st.button(f"✅ Yes, Delete", key=f"confirm_yes_excel_{doc_type}_{county}"):
                                try:
                                    os.remove(excel_path)
                                    st.session_state.docs_indexed[doc_type] = False  # Mark as needs re-index
                                    st.session_state[confirm_key] = False  # Reset confirmation
                                    st.success(f"{doc_type} Excel deleted successfully!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Failed to delete Excel: {str(e)}")
                        with col_confirm2:
                            if st.button(f"❌ Cancel", key=f"confirm_no_excel_{doc_type}_{county}"):
                                st.session_state[confirm_key] = False
                                st.rerun()
                
                # Index/Re-Index and Delete Index Buttons
                col_idx1, col_idx2 = st.columns(2)
                
                with col_idx1:
                    index_text = "Re-Index" if st.session_state.docs_indexed.get(doc_type, False) else "Index"
                    if st.button(f"{index_text} {doc_type}", key=f"index_{doc_type}_{county}"):
                        pdf_path = get_doc_path(county_dir, doc_type, "pdf")
                        excel_path = get_doc_path(county_dir, doc_type, "xlsx")
                        if os.path.exists(pdf_path):
                            # Create progress bar
                            progress_placeholder = st.empty()
                            progress_bar = progress_placeholder.progress(0, text=f"Starting to index {doc_type}...")
                            
                            try:
                                index_data = index_pdf_with_progress(
                                    pdf_path, 
                                    excel_path if os.path.exists(excel_path) else None, 
                                    doc_type, 
                                    progress_bar
                                )
                                progress_bar.progress(1.0, text="Saving index...")
                                save_index(county_dir, doc_type, index_data)
                                st.session_state.docs_indexed[doc_type] = True
                                progress_placeholder.empty()  # Remove progress bar
                                st.success(f"{doc_type} indexed successfully!")
                                st.rerun()
                            except Exception as e:
                                progress_placeholder.empty()  # Remove progress bar on error
                                st.error(f"Indexing failed: {str(e)}")
                        else:
                            st.warning(f"Please upload {doc_type} PDF first.")
                
                with col_idx2:
                    # Delete Index Button
                    index_path = get_doc_path(county_dir, doc_type, "json")
                    if os.path.exists(index_path):
                        if st.button(f"🗑️ Clear {doc_type} Index", key=f"delete_index_{doc_type}_{county}", help="Delete the search index (keeps files)"):
                            try:
                                os.remove(index_path)
                                st.session_state.docs_indexed[doc_type] = False
                                st.success(f"{doc_type} index cleared!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Failed to clear index: {str(e)}")

    # Check indexing status
    st.subheader("📊 Indexing Status")
    
    indexed_count = 0
    col1, col2, col3 = st.columns(3)
    
    for i, doc_type in enumerate(DOC_TYPES):
        col = [col1, col2, col3][i]
        index_file = get_doc_path(county_dir, doc_type, "json")
        is_indexed = os.path.exists(index_file)
        
        if is_indexed:
            indexed_count += 1
            # Get index stats
            try:
                with open(index_file, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)
                    account_count = len(index_data)
                status_text = f"✅ **{doc_type}**\n\n{account_count:,} accounts indexed"
                col.success(status_text)
            except:
                col.success(f"✅ **{doc_type}**\n\nIndexed (stats unavailable)")
        else:
            col.error(f"❌ **{doc_type}**\n\nNot indexed")
    
    # Summary message
    if indexed_count == 0:
        st.warning("🚨 **No documents indexed yet.** Upload and index at least one document type to start searching.")
    elif indexed_count == len(DOC_TYPES):
        st.success(f"🎉 **All document types indexed!** You can search across all {len(DOC_TYPES)} document types.")
    else:
        st.info(f"📈 **{indexed_count} of {len(DOC_TYPES)} document types indexed.** You can search the indexed types and add more anytime.")
    
    # Bulk Actions
    if indexed_count > 0:
        with st.expander("🔧 Bulk Actions", expanded=False):
            col_bulk1, col_bulk2 = st.columns(2)
            
            with col_bulk1:
                st.write("**Clear All Indexes:**")
                if st.button("🗑️ Clear All Indexes", key=f"clear_all_indexes_{county}", help="Remove all search indexes (keeps PDF/Excel files)"):
                    cleared_count = 0
                    for doc_type in DOC_TYPES:
                        index_file = get_doc_path(county_dir, doc_type, "json")
                        if os.path.exists(index_file):
                            try:
                                os.remove(index_file)
                                st.session_state.docs_indexed[doc_type] = False
                                cleared_count += 1
                            except Exception as e:
                                st.error(f"Failed to clear {doc_type} index: {str(e)}")
                    if cleared_count > 0:
                        st.success(f"Cleared {cleared_count} index(es) successfully!")
                        st.rerun()
            
            with col_bulk2:
                st.write("**Delete All Files:**")
                if st.button("⚠️ Delete All Files", key=f"delete_all_files_{county}", help="PERMANENTLY delete all PDF, Excel, and index files"):
                    deleted_files = 0
                    for doc_type in DOC_TYPES:
                        # Delete PDF
                        pdf_path = get_doc_path(county_dir, doc_type, "pdf")
                        if os.path.exists(pdf_path):
                            try:
                                os.remove(pdf_path)
                                deleted_files += 1
                            except Exception as e:
                                st.error(f"Failed to delete {doc_type} PDF: {str(e)}")
                        
                        # Delete Excel
                        excel_path = get_doc_path(county_dir, doc_type, "xlsx")
                        if os.path.exists(excel_path):
                            try:
                                os.remove(excel_path)
                                deleted_files += 1
                            except Exception as e:
                                st.error(f"Failed to delete {doc_type} Excel: {str(e)}")
                        
                        # Delete Index
                        index_path = get_doc_path(county_dir, doc_type, "json")
                        if os.path.exists(index_path):
                            try:
                                os.remove(index_path)
                                deleted_files += 1
                            except Exception as e:
                                st.error(f"Failed to delete {doc_type} index: {str(e)}")
                        
                        # Reset session state
                        st.session_state.docs_indexed[doc_type] = False
                    
                    if deleted_files > 0:
                        st.success(f"Deleted {deleted_files} file(s) successfully!")
                        st.rerun()
                    else:
                        st.info("No files found to delete.")
            
            st.warning("⚠️ **Important:** Delete operations are permanent and cannot be undone!")