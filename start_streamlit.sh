#!/bin/bash

# Kill any existing Streamlit processes
echo "Stopping any existing Streamlit processes..."
pkill -f streamlit

# Wait a moment
sleep 2

# Start Streamlit with correct binding for external access
echo "Starting Streamlit on all interfaces..."
cd /path/to/your/app  # Change this to the actual path where public_docs.py is located

streamlit run public_docs.py \
    --server.port=8503 \
    --server.address=0.0.0.0 \
    --server.headless=true \
    --server.enableCORS=false \
    --server.enableXsrfProtection=false \
    --browser.gatherUsageStats=false

# Note: Change --server.address=0.0.0.0 to bind to all interfaces (not just 127.0.0.1)