import os
import pandas as pd

# === Configuration ===
input_folder = r"C:\Users\Sidharth\Desktop\vendor rating"
output_file = r"C:\Users\Sidharth\Desktop\combined_vendor_rating.xlsx"

# === Specify the desired order (partial names allowed) ===
# Example: if folder contains "summary(1).xlsx", just write "summary"
ordered_keywords = [
    "summary",
    "Unit data",
    "cc box",
    "new btl",
    "pet btl",
    "old btl",
    "cap",
    "monocarton",
    "sticker",
    "lable",
    "pp seal",
    "cannister",
    "Others"
]

# === Process ===
# Get all Excel files from the folder
all_files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls'))]

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for keyword in ordered_keywords:
        matched = None
        for file in all_files:
            if keyword.lower() in file.lower():
                matched = file
                break

        if matched:
            file_path = os.path.join(input_folder, matched)
            print(f"✅ Matched '{keyword}' → '{matched}'")
            df = pd.read_excel(file_path)
            sheet_name = os.path.splitext(matched)[0][:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            print(f"⚠️ No match found for keyword: '{keyword}'")

print(f"\n✅ Combined Excel created successfully at:\n{output_file}")
