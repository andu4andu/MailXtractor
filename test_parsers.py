from parsers import parse_zip

ZIP_PATH = r"C:\Users\uik11822\301_Contract_Ingestion_AP\Input\_Toyota 12.3 Meter display --- Waichi BLU Technical review\Fw_ RE_ Toyota 12.3_ Meter display --- Waichi BLU Technical review - Sample 3.zip"

result = parse_zip(ZIP_PATH)

with open("test_output.txt", "w", encoding="utf-8") as f:
    f.write(result)

print("Output written to test_output.txt")
