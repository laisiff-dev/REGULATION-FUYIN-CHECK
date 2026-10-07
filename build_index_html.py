# -*- coding: utf-8 -*-
"""
build_index_html.py
讀取 templates/index.html，將全校 361 筆法規檢核成果 (audit_results.json)
內嵌至 index.html 與 templates/index.html，
確保在純本機雙擊開啟 (file:// 協議，受瀏覽器 CORS 限制無伺服器環境) 或 Flask 伺服器環境下，
均能 100% 正常渲染 361 筆法規、即時查詢搜尋、母子法位階診斷與在線修訂。
"""

import json
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 1. 讀取 audit_results.json
with open(os.path.join(BASE_DIR, 'audit_results.json'), 'r', encoding='utf-8') as f:
    audit_data = json.load(f)

json_embed = json.dumps(audit_data, ensure_ascii=False)

# 2. 讀取 templates/index.html 母版
template_path = os.path.join(BASE_DIR, 'templates', 'index.html')
with open(template_path, 'r', encoding='utf-8') as f:
    html_content = f.read()

# 3. 將 DEFAULT_AUDIT_DATA 替換為實際內嵌資料
html_embedded = re.sub(
    r'const DEFAULT_AUDIT_DATA\s*=\s*\[.*?\];',
    lambda m: f'const DEFAULT_AUDIT_DATA = {json_embed};',
    html_content,
    flags=re.DOTALL
)

# 4. 寫入根目錄 index.html
output_root = os.path.join(BASE_DIR, 'index.html')
with open(output_root, 'w', encoding='utf-8') as f:
    f.write(html_embedded)

# 5. 同步寫入 templates/index.html
with open(template_path, 'w', encoding='utf-8') as f:
    f.write(html_embedded)

print(f"成功構建並內嵌數據！檔案大小: {os.path.getsize(output_root)} bytes (共內嵌 {len(audit_data)} 筆檢核數據)")
