import os
from readers import read_raw_email, extract_body_text

INPUT_PATH = r"C:\Users\uik11822\Downloads\301_Contract_Ingestion_AP\301_Contract_Ingestion_AP\Input"

count = 0
for folder_name in sorted(os.listdir(INPUT_PATH)):
    folder_path = os.path.join(INPUT_PATH, folder_name)
    if os.path.isdir(folder_path):
        email_struct = read_raw_email(folder_path)
        body_string = extract_body_text(email_struct)
        count += 1

print(f"\n=== Total folders processed: {count} ===")
