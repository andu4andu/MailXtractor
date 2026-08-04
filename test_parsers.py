from parsers import parse_pdf

PDF_PATH = r"C:\Users\uik11822\301_Contract_Ingestion_AP\Input\Chery 17.3 roof display_PCB Cover\CAF_DCLQ250021R2_AAA2835010000_20250912_Mandfield.pdf"

result = parse_pdf(PDF_PATH)

with open("test_output.txt", "w", encoding="utf-8") as f:
    f.write(result)

print("Output written to test_output.txt")
