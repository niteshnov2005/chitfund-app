
import pandas as pd
import io
import traceback

try:
    print("Testing ExcelWriter with BytesIO...")
    output = io.BytesIO()
    df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        print("Written to writer inside context.")
    
    print("Exited context.")
    
    try:
        output.seek(0)
        print("Seek successful.")
        content = output.read()
        print(f"Read {len(content)} bytes.")
    except Exception as e:
        print("FAILED to read output after context exit:")
        print(e)

except Exception as e:
    traceback.print_exc()
