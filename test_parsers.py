from parsers import parse_xlsx

XLSX_PATH = r"C:\Users\uik11822\301_Contract_Ingestion_AP\Input\GW B01-5_Backlight Assy\ALCF version---CAF Cost Analysis Form for 10.25 from HADBEST 20251219 update.xlsx CNY报价.xlsx"

result = parse_xlsx(XLSX_PATH)

with open("test_output.txt", "w", encoding="utf-8") as f:
    f.write(result)

print("Output written to test_output.txt")
