
import sys
import os
sys.path.append(os.getcwd())
try:
    from app import get_excel_data
    
    # 1. Default (should be '20')
    data_default = get_excel_data()
    print(f"Default Sheet Members: {len(data_default)}")
    if data_default:
        print(f"Sample: {data_default[0]['name']} - Total: {data_default[0]['total']}")
        
    # 2. Key Check: Does it have items?
    items_count = sum(len(m['items']) for m in data_default)
    print(f"Total Items found in Default Sheet: {items_count}")
    
except Exception as e:
    print(e)
