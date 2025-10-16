import streamlit as st
import os

# Simple test version to check if basic Streamlit works
st.title("🧪 Test - Public Docs Portal")

st.write("If you can see this, Streamlit is working!")

# Test subdomain detection
try:
    import streamlit.runtime as runtime
    session_mgr = runtime.get_instance()._session_mgr
    active_sessions = session_mgr.list_active_sessions()
    if active_sessions:
        request = active_sessions[0].client.request
        host = request.host.lower().strip()
        st.write(f"**Detected host:** {host}")
    else:
        st.write("**No active session found**")
except Exception as e:
    st.error(f"**Error detecting host:** {str(e)}")

# Test file system
st.write(f"**Current directory:** {os.getcwd()}")
st.write(f"**Files in directory:** {os.listdir('.')}")

# Test county directories
if os.path.exists('county_docs'):
    st.write(f"**County docs directory exists:** {os.listdir('county_docs')}")
else:
    st.write("**County docs directory missing**")

st.success("✅ Basic test completed!")