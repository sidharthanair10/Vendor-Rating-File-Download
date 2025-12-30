import os
import pandas as pd

# === Configuration ===
input_folder = r"C:\Users\Sidharth\Desktop\Vendor_Rating"  
output_file = r"C:\Users\Sidharth\Desktop\combined_vendor_rating.xlsx"

# === Process ===
excel_files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls'))]
excel_files.sort()  

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        df = pd.read_excel(file_path)
        
        sheet_name = os.path.splitext(file)[0][:31]
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"✅ Combined {len(excel_files)} files into '{output_file}'")
