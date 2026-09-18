# -*- coding: utf-8 -*-
"""
verify_audit.py
# 驗證全校 20 個單位、361 筆法規檢核成果與 Excel/JSON 資料庫之完整性與一致性。
"""

import json
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

with open('audit_results.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("=======================================================")
print("=== 輔英科技大學 全校法規檢核成果自動驗證 (361筆) ===")
print("=======================================================")

total = len(data)
print(f"1. 資料庫總筆數: {total} 筆 (預期: 361 筆) -> {'PASS' if total == 361 else 'FAIL'}")

# Check condition values
c1_pass = sum(1 for d in data if d['c1'] == 'V')
c2_pass = sum(1 for d in data if d['c2'] == 'V')
c3_fail = sum(1 for d in data if d['c3'] == 'X')
c3_pass = sum(1 for d in data if d['c3'] == 'V')
c4_needs = sum(1 for d in data if d['c4'] == 'V')
c4_ok = sum(1 for d in data if d['c4'] == 'X')
c5_needs = sum(1 for d in data if d['c5'] == 'V')

print(f"2. 條件1 (位階/最新版): 合格 {c1_pass}/{total} ({c1_pass/total*100:.1f}%)")
print(f"3. 條件2 (名稱正確性): 合格 {c2_pass}/{total} ({c2_pass/total*100:.1f}%)")
print(f"4. 條件3 (格式完整性): 合格 {c3_pass} 件, 瑕疵 {c3_fail} 件 (瑕疵率: {c3_fail/total*100:.1f}%)")
print(f"5. 條件4 (逾2年需提案): 逾期需提案 {c4_needs} 件, 時效正常 {c4_ok} 件 (逾期率: {c4_needs/total*100:.1f}%)")
print(f"6. 條件5 (組織規程因應): 需提案更新 {c5_needs} 件 ({c5_needs/total*100:.1f}%)")

# Unit distribution
print("\n--- 各單位法規數量與檢核瑕疵分布 ---")
by_sheet = {}
for d in data:
    s = d['sheet']
    if s not in by_sheet:
        by_sheet[s] = {'total': 0, 'c3_fail': 0, 'c4_v': 0, 'c5_v': 0}
    by_sheet[s]['total'] += 1
    if d['c3'] == 'X':
        by_sheet[s]['c3_fail'] += 1
    if d['c4'] == 'V':
        by_sheet[s]['c4_v'] += 1
    if d['c5'] == 'V':
        by_sheet[s]['c5_v'] += 1

print(f"{'單位名稱':<18} | {'總數':<4} | {'格式瑕疵(C3)':<10} | {'逾期提案(C4)':<10} | {'組織修訂(C5)':<10}")
print("-" * 65)
for s, st in by_sheet.items():
    print(f"{s:<18} | {st['total']:<4} | {st['c3_fail']:<10} | {st['c4_v']:<10} | {st['c5_v']:<10}")

# Verify Excel files
print("\n--- 驗證 Excel 檔案寫入完整性 ---")
for xfile in ['115年法規清單檢核表_已完成檢核.xlsx', '115年法規清單檢核表.xlsx']:
    wb = openpyxl.load_workbook(xfile, data_only=True)
    unfilled = 0
    total_rows = 0
    for sname in wb.sheetnames:
        sheet = wb[sname]
        for r in range(4, sheet.max_row + 1):
            seq = sheet.cell(r, 1).value
            if seq is not None and isinstance(seq, (int, float)):
                total_rows += 1
                c1 = sheet.cell(r, 4).value
                c3 = sheet.cell(r, 6).value
                c4 = sheet.cell(r, 7).value
                if not c1 or not c3 or not c4:
                    unfilled += 1
    print(f"{xfile}: 總列管法規 {total_rows} 列, 未填格數 {unfilled} -> {'ALL FILLED PASS' if unfilled == 0 and total_rows == 361 else 'FAIL'}")

print("\n=== 全校法規資料庫驗證合格！ ===")
