#!/bin/bash

echo "=== STREAMLIT DEBUG SCRIPT ==="
echo

echo "1. Checking if Streamlit process is running:"
ps aux | grep streamlit | grep -v grep
echo

echo "2. Checking what's listening on port 8503:"
sudo netstat -tulpn | grep :8503
echo

echo "3. Checking if port 8503 is accessible locally:"
curl -I http://127.0.0.1:8503
echo

echo "4. Checking current working directory and files:"
pwd
ls -la public_docs.py 2>/dev/null || echo "public_docs.py not found in current directory"
echo

echo "5. Checking firewall status:"
sudo ufw status
echo

echo "6. Checking security groups (if this fails, you're not on AWS):"
curl -s http://169.254.169.254/latest/meta-data/security-groups 2>/dev/null || echo "Not on AWS or metadata service unavailable"
echo

echo "7. Checking system resources:"
free -h
df -h
echo

echo "=== END DEBUG ==="