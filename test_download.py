
import requests
import sys

# URL
LOGIN_URL = 'http://127.0.0.1:5000/login'
DOWNLOAD_URL = 'http://127.0.0.1:5000/download_excel'

try:
    s = requests.Session()
    
    # 1. Login
    print("Attempting login...")
    r = s.post(LOGIN_URL, data={'username': 'admin', 'password': 'nitesh2025'})
    print(f"Login Response: {r.status_code}")
    if 'Dashboard' not in r.text and r.url != 'http://127.0.0.1:5000/dashboard':
        print("Login might have failed or not redirected to dashboard.")
        # But app redirects to dashboard on success.
    
    # 2. Download
    print("Attempting download...")
    r = s.get(DOWNLOAD_URL)
    print(f"Download Response Code: {r.status_code}")
    print(f"Content Type: {r.headers.get('Content-Type')}")
    print(f"Content Disposition: {r.headers.get('Content-Disposition')}")
    
    content = r.content
    print(f"Content Length: {len(content)}")
    
    if len(content) > 4:
        # Check signature for zip (xlsx is a zip)
        # PK\x03\x04
        if content.startswith(b'PK\x03\x04'):
            print("SUCCESS: File appears to be a valid ZIP/XLSX.")
        elif content.startswith(b'<!DOCTYPE html'):
            print("FAILURE: File appears to be HTML (Login page redirection?).")
        else:
            print(f"UNKNOWN: First bytes: {content[:10]}")

except Exception as e:
    print(f"Error: {e}")
