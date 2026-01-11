
import pandas as pd
import os
import sys

# Add current directory to path
sys.path.append(os.getcwd())

from app import get_excel_data

try:
    data = get_excel_data()
    print(f"Successfully retrieved {len(data)} members.")
    if len(data) > 0:
        print("First member sample:", data[0])
    
    # Check if we can create the DataFrame as intended
    export_data = []
    for m in data:
        base_info = {
            'Name': m['name'],
            'Area': m['area'],
            'Total Amount': m['total'],
            'Payment ID': m['payment_id'],
            'Fully Paid?': 'Yes' if m['is_paid'] else 'No',
            'Paid Date': m['paid_date'] if m['paid_date'] else '-'
        }
        if m['items']:
            for item in m['items']:
                row = base_info.copy()
                row.update({
                    'Month': item.get('month', '-'),
                    'Plan': item.get('plan', '-'),
                    'Commission': item.get('commission', '-'),
                    'Item Amount': item.get('amount', 0)
                })
                export_data.append(row)
        else:
            row = base_info.copy()
            row.update({
                'Month': 'Summary/Total',
                'Plan': '-',
                'Commission': '-',
                'Item Amount': m['total']
            })
            export_data.append(row)
            
    df = pd.DataFrame(export_data)
    print(f"DataFrame shape: {df.shape}")
    print("DataFrame Head:")
    print(df.head())
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
