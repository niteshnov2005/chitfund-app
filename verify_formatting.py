
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
    if r.status_code != 200:
        print(f"Login failed: {r.status_code}")
        sys.exit(1)
        
    # 2. Download
    print("Attempting download...")
    r = s.get(DOWNLOAD_URL)
    print(f"Download Response Code: {r.status_code}")
    print(f"Content Type: {r.headers.get('Content-Type')}")
    
    if r.status_code == 200:
        if r.content.startswith(b'PK\x03\x04'):
             print("SUCCESS: Received valid ZIP/XLSX file with new formatting.")
             # Optionally save it to inspect
             with open('formatted_receipts_test.xlsx', 'wb') as f:
                 f.write(r.content)
             print("Saved to formatted_receipts_test.xlsx")
        else:
             print("FAILURE: Content is not XLSX.")
             print(r.content[:100])

except Exception as e:
    print(f"Error: {e}")
