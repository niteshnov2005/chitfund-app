
import openpyxl
import pandas as pd
import os

FILE_NAME = 'sample_gemini.xlsx'

if os.path.exists(FILE_NAME):
    try:
        # Check Sheet Names using openpyxl (Order matters)
        wb = openpyxl.load_workbook(FILE_NAME, read_only=True)
        print("Openpyxl Sheet Names (Order):", wb.sheetnames)
        wb.close()
        
        # Check what Pandas reads by default
        df = pd.read_excel(FILE_NAME, header=None, engine='openpyxl')
        print("Pandas Default Read - Cell A1:", df.iloc[0, 0])
        print("Pandas Default Read - Cell B1:", df.iloc[0, 1])
        
    except Exception as e:
        print(f"Error: {e}")
else:
    print("File not found.")
