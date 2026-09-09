import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('audit_results.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print(f"Total items in audit database: {len(data)}")

c3_failed = [d for d in data if d['c3'] == 'X']
c4_needs = [d for d in data if d['c4'] == 'V']
c5_needs = [d for d in data if d['c5'] == 'V']

print(f"Format issues (Condition 3 = X): {len(c3_failed)}")
print(f"Over 2 years needs proposal (Condition 4 = V): {len(c4_needs)}")
print(f"Org structure change proposal (Condition 5 = V): {len(c5_needs)}")

print("\n--- Sample R&D Office Audit Items ---")
rd_items = [d for d in data if '研發' in d['sheet']]
for item in rd_items[:12]:
    print(f"[{item['seq']}] Sheet:{item['sheet']} | Name:{item['reg_name'][:30]}")
    print(f"    Conditions: C1={item['c1']}, C2={item['c2']}, C3={item['c3']}, C4={item['c4']}, C5={item['c5']}")
    print(f"    Date: {item['last_date']} | Notes: {item['notes']}\n")
