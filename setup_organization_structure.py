# -*- coding: utf-8 -*-
"""
setup_organization_structure.py
依據輔英科技大學官方網站最新組織架構與 115 年法規清單檢核表，
自動建立「輔英科大各單位法規彙整」的完整階層目錄樹，
並將現有 38 筆法規自動歸檔複製至對應子目錄中，產出各單位索引 README.md。
"""

import os
import shutil
import re

BASE_DIR = os.path.join(os.getcwd(), '輔英科大各單位法規彙整')

# Directory hierarchy definition
STRUCTURE = {
    '01_行政單位': [
        '01_董事會',
        '02_校長室與秘書處',
        '03_教務處',
        '04_學生事務處',
        '05_總務處',
        '06_研發與永續發展處',
        '07_圖書暨資訊處',
        '08_國際暨兩岸事務處',
        '09_人事室',
        '10_會計室',
        '11_稽核室',
        '12_環境安全衛生中心',
        '13_健康與體育發展處',
        '14_推廣教育中心'
    ],
    '02_教學學術單位': {
        '01_護理學院': [
            '護理系(所、科)',
            '高齡及長期照護事業系',
            '健康事業管理系',
            '助產與婦嬰健康照護系',
            '學士後護理系'
        ],
        '02_醫學與健康學院': [
            '醫學檢驗生物技術系(科)',
            '物理治療系',
            '保健營養系',
            '健康美容系'
        ],
        '03_環境與生命學院': [
            '環境工程與科學系(科)',
            '職業安全衛生系',
            '應用化學及材料科學系',
            '生物科技系'
        ],
        '04_人文與管理學院': [
            '資訊科技與管理系',
            '休閒與遊憩事業管理系',
            '幼兒保育暨產業系',
            '應用外語系(科)'
        ],
        '05_共同教育中心': []
    },
    '03_研究與中心機構': [
        '01_老化及疾病預防研究中心',
        '02_校務研究與永續發展中心',
        '03_精準健康與環境檢驗中心',
        '04_任務導向型研究中心'
    ],
    '04_專門委員會與附設機構': [
        '01_產學合作暨智財管理委員會',
        '02_學術誠信與研究倫理委員會',
        '03_永續發展委員會',
        '04_附設醫院與幼兒園'
    ]
}

def create_directories():
    os.makedirs(BASE_DIR, exist_ok=True)
    created_paths = []

    # 1. 行政單位
    for u in STRUCTURE['01_行政單位']:
        p = os.path.join(BASE_DIR, '01_行政單位', u)
        os.makedirs(p, exist_ok=True)
        created_paths.append(p)

    # 2. 教學學術單位
    for col, depts in STRUCTURE['02_教學學術單位'].items():
        col_p = os.path.join(BASE_DIR, '02_教學學術單位', col)
        os.makedirs(col_p, exist_ok=True)
        created_paths.append(col_p)
        for dept in depts:
            dept_p = os.path.join(col_p, dept)
            os.makedirs(dept_p, exist_ok=True)
            created_paths.append(dept_p)

    # 3. 研究與中心機構
    for c in STRUCTURE['03_研究與中心機構']:
        p = os.path.join(BASE_DIR, '03_研究與中心機構', c)
        os.makedirs(p, exist_ok=True)
        created_paths.append(p)

    # 4. 專門委員會與附設機構
    for com in STRUCTURE['04_專門委員會與附設機構']:
        p = os.path.join(BASE_DIR, '04_專門委員會與附設機構', com)
        os.makedirs(p, exist_ok=True)
        created_paths.append(p)

    print(f"成功建立 {len(created_paths)} 個階層子目錄。")
    return created_paths

def copy_regulations():
    """Copy the 38 regulation files from root to the corresponding subdirectories."""
    rd_dir = os.path.join(BASE_DIR, '01_行政單位', '06_研發與永續發展處')
    ir_dir = os.path.join(BASE_DIR, '03_研究與中心機構', '02_校務研究與永續發展中心')
    comm_ip_dir = os.path.join(BASE_DIR, '04_專門委員會與附設機構', '01_產學合作暨智財管理委員會')
    comm_eth_dir = os.path.join(BASE_DIR, '04_專門委員會與附設機構', '02_學術誠信與研究倫理委員會')
    comm_sdg_dir = os.path.join(BASE_DIR, '04_專門委員會與附設機構', '03_永續發展委員會')

    copied_count = 0
    for f in os.listdir('.'):
        m = re.match(r'^(\d+)(.*)\.(docx|doc|pdf)$', f)
        if m and not f.startswith('~') and '條件' not in f:
            num = int(m.group(1))
            
            # Destination mapping
            if 1 <= num <= 35:
                dest = os.path.join(rd_dir, f)
                shutil.copy2(f, dest)
                copied_count += 1
                # Cross-reference copies for committees
                if num == 27: # 智財委員會辦法
                    shutil.copy2(f, os.path.join(comm_ip_dir, f))
                elif num == 34: # 倫理委員會辦法
                    shutil.copy2(f, os.path.join(comm_eth_dir, f))
            elif 36 <= num <= 38:
                dest = os.path.join(ir_dir, f)
                shutil.copy2(f, dest)
                copied_count += 1
                if num == 37: # 永續發展委員會要點
                    shutil.copy2(f, os.path.join(comm_sdg_dir, f))

    print(f"已成功歸檔複製 {copied_count} 份法規文件至專屬單位子目錄！")

def generate_readmes():
    """Generate index README.md files for master and key units."""
    # Master README
    master_readme = [
        "# 輔英科技大學各單位法規彙整體系",
        "",
        "> **依據標準**：依據輔英科技大學官方網站最新組織架構、115 年法規清單檢核表及《法規檢核五個條件_更新.docx》建立。",
        "> **更新時間**：民國 115 年 9 月",
        "",
        "---",
        "",
        "## 一、架構目錄索引",
        "",
        "### 1. [01_行政單位](./01_行政單位/)",
        "- `01_董事會/`：董事會組織章程與運作法規",
        "- `02_校長室與秘書處/`：秘書處、公關校友、全校性會議規程",
        "- `03_教務處/`：學則、選課、成績、教師評鑑等教學法規",
        "- `04_學生事務處/`：學生獎懲、就學貸款、社團輔導法規",
        "- `05_總務處/`：採購、財產、營繕、出納管理法規",
        "- `06_研發與永續發展處/`：**【現有法規歸檔主力】** 包含 01~35 號研發獎勵、專利、技轉、產學合作法規",
        "- `07_圖書暨資訊處/`：圖書借閱、校園網路、資安政策規章",
        "- `08_國際暨兩岸事務處/`：境外生招生、國際交流、雙聯學位要點",
        "- `09_人事室/`：教職員聘任、升等、請假、考核及退撫法規",
        "- `10_會計室/`：經費編列、核銷、主計審計規約",
        "- `11_稽核室/`：內部控制與稽核作業規範",
        "- `12_環境安全衛生中心/`：環安衛管委會、實驗室安全、承攬商環安規則",
        "- `13_健康與體育發展處/`：運動場館管理、體育課程與賽事規約",
        "- `14_推廣教育中心/`：推廣教育實施要點、推廣教師聘任辦法",
        "",
        "### 2. [02_教學學術單位](./02_教學學術單位/)",
        "- `01_護理學院/`（含護理系所科、高照系、健管系、助產系、學士後護理系）",
        "- `02_醫學與健康學院/`（含醫技系科、物治系、保營系、健康美容系）",
        "- `03_環境與生命學院/`（含環工系科、職安系、應化系、生技系）",
        "- `04_人文與管理學院/`（含資管系、休閒系、幼保系、應外系）",
        "- `05_共同教育中心/`（通識博雅、外語課程、專業倫理法規）",
        "",
        "### 3. [03_研究與中心機構](./03_研究與中心機構/)",
        "- `01_老化及疾病預防研究中心/`：共享儀器設備管理、中心運作要點",
        "- `02_校務研究與永續發展中心/`：**【現有法規歸檔主力】** 第 36~38 號校務研究設置辦法、永續委員會要點、校務資料庫管理要點",
        "- `03_精準健康與環境檢驗中心/`：檢驗認證規約與服務要點",
        "- `04_任務導向型研究中心/`：各型研究中心設置與退場作業要點",
        "",
        "### 4. [04_專門委員會與附設機構](./04_專門委員會與附設機構/)",
        "- `01_產學合作暨智財管理委員會/`（對應第 27 號辦法）",
        "- `02_學術誠信與研究倫理委員會/`（對應第 34 號辦法）",
        "- `03_永續發展委員會/`（對應第 37 號要點）",
        "- `04_附設醫院與幼兒園/`",
        "",
        "---",
        "",
        "## 二、全校檢核成果連結",
        "- 📊 [檢核結果總表 (HTML 儀表板)](../audit_report.html)",
        "- 🌲 [法規位階關係檔 (Markdown)](../法規位階關係檔.md)",
        "- 🔍 [法規格式檢核報告 (Markdown)](../法規格式檢核報告.md)",
        "- ⏰ [法規逾期未更新標註檔 (Markdown)](../法規逾期未更新標註檔.md)",
        "- 📋 [法規檢核結果總表 (Markdown)](../法規檢核結果總表.md)"
    ]

    with open(os.path.join(BASE_DIR, 'README.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(master_readme))

    # RD Dept README
    rd_dir = os.path.join(BASE_DIR, '01_行政單位', '06_研發與永續發展處')
    rd_readme = [
        "# 研發與永續發展處 法規彙整清單",
        "",
        "本目錄收錄輔英科技大學研發與永續發展處（原研究發展處）現行主管之法規文件（編號 01 至 35）：",
        "",
        "| 編號 | 法規名稱 | 位階等級 | 最後修訂日期 | 格式檢核 | 逾期提案標註 |",
        "| :---: | :--- | :---: | :---: | :---: | :---: |"
    ]
    # read full_audit_data.json
    import json
    with open('full_audit_data.json', 'r', encoding='utf-8') as fp:
        docs = json.load(fp)

    for d in docs:
        if 1 <= d['num'] <= 35:
            c3_icon = '🟢 合格' if not d['has_fmt_issues'] else '🔴 含瑕疵'
            c4_icon = '🔴 逾期需提案' if d['is_over_2_years'] else '🟢 正常'
            rd_readme.append(f"| {d['num']:02d} | **{d['raw_name']}** | {d['rank']} | `{d['last_date']}` | {c3_icon} | {c4_icon} |")

    rd_readme.extend([
        "",
        "---",
        "詳細格式瑕疵與修訂指南請參見根目錄 [法規格式檢核報告.md](../../法規格式檢核報告.md)。"
    ])
    with open(os.path.join(rd_dir, 'README.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(rd_readme))

    # IR Center README
    ir_dir = os.path.join(BASE_DIR, '03_研究與中心機構', '02_校務研究與永續發展中心')
    ir_readme = [
        "# 校務研究與永續發展中心 法規彙整清單",
        "",
        "本目錄收錄校務研究與永續發展中心（原校務研究暨規劃室）現行主管之法規文件（編號 36 至 38）：",
        "",
        "| 編號 | 法規名稱 | 位階等級 | 最後修訂日期 | 格式檢核 | 逾期提案標註 |",
        "| :---: | :--- | :---: | :---: | :---: | :---: |"
    ]
    for d in docs:
        if 36 <= d['num'] <= 38:
            c3_icon = '🟢 合格' if not d['has_fmt_issues'] else '🔴 含瑕疵'
            c4_icon = '🔴 逾期需提案' if d['is_over_2_years'] else '🟢 正常'
            ir_readme.append(f"| {d['num']:02d} | **{d['raw_name']}** | {d['rank']} | `{d['last_date']}` | {c3_icon} | {c4_icon} |")

    with open(os.path.join(ir_dir, 'README.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(ir_readme))

    print("已成功建立總目錄及各重點單位之 README.md 索引檔案！")

if __name__ == '__main__':
    create_directories()
    copy_regulations()
    generate_readmes()
