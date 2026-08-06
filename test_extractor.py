from extractor import load_rules, apply_extraction_rules

RULES_PATH = "rules.json"

body_string = ""

attachment_string = """
[CAF_2031939049_UX_S1_Chery 17.3 DS_PCB Cover_5754 H111_20250912_Supplier Miyoshi.pdf]
Cost Analysis Form
Part no.: AAA2835010000
Part name: PCBA COVER
Drawing no. and Index: 42176232 DRW 000 AA
Average Parts / year: 37,500
Supplier name: Wuxi Miyoshi Precision Co.,Ltd
Prod.Location: No.108 Hongda Road,Hongshan street,Wuxi,China
Shifts/week: 15
Currency: CNY
Date: 28.07.2025
Total B - Price [CNY] / Part 3.450
Total Tooling Costs [CNY] 192,000.00
"""

rules = load_rules(RULES_PATH)
result = apply_extraction_rules(body_string, attachment_string, rules)

print("\n--- EXTRACTED FIELDS ---")
for field, value in result.items():
    print(f"  {field}: {repr(value)}")
