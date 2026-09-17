# -*- coding: utf-8 -*-
"""
generate_reports.py
依據《法規檢核五個條件_更新.docx》及本校《115年法規清單檢核表》，
全面解析全校 21 個單位、361 筆法規檢核數據，自動產出四大專業 Markdown 專卷：
1. 法規位階關係檔.md (涵蓋六大核心拓撲圖與全校母子法位階清冊)
2. 法規格式檢核報告.md (全校格式瑕疵清單與公文定稿清理SOP)
3. 法規逾期未更新標註檔.md (207筆逾期未修訂法規、三級急迫性與115提案排程)
4. 法規檢核結果總表.md (全校KPI儀表板、21單位橫向對照表與361案總表)
"""

import json
import os
import re
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, 'full_audit_data.json'), 'r', encoding='utf-8') as f:
    docs = json.load(f)

print(f"載入 full_audit_data.json 成功，共 {len(docs)} 筆受檢法規。")

# -----------------------------------------------------------------------------
# 1. 產生《法規位階關係檔.md》
# -----------------------------------------------------------------------------
def generate_hierarchy_md():
    out_path = os.path.join(BASE_DIR, '法規位階關係檔.md')
    lines = []
    lines.append("# 輔英科技大學法規位階關係檔\n")
    lines.append("> **檢核依據**：《法規檢核五個條件_更新.docx》第一項條件：「檢測各法規之間是否有母法與子法的位階關係, 建立位階關係檔。」  ")
    lines.append(f"> **受檢對象**：輔英科技大學全校 21 個單位、全體 {len(docs)} 筆法規文檔  ")
    lines.append("> **基準日期**：民國 115 年 9 月\n")
    lines.append("---\n")

    lines.append("## 一、全校法規位階體系總綱\n")
    lines.append("依據我國大學自治憲法法理、《中央法規標準法》及《輔英科技大學組織規程》，本校規章體系依法律位階與審議會議層級，嚴格分為五個授權位階：\n")
    lines.append("| 位階層級 | 規章類型 | 審議／核定會議層級 | 法律效力與規範特徵 | 全校代表性規章 |")
    lines.append("| :--- | :--- | :--- | :--- | :--- |")
    lines.append("| **第一位階 (上位法源)** | 中央法律／部會法規 | 立法院通過、總統公布／教育部、國科會發布 | 全校法規訂定之最高法源，牴觸者無效 | 《大學法》、《私立學校法》、《學位授予法》、《教師法》、《專利法》 |")
    lines.append("| **第二位階 (根本規程)** | 組織規程 | **校務會議**通過、教育部核定 | 本校立校根本大法，設立各行政與學術單位之法源 | 《輔英科技大學組織規程》 |")
    lines.append("| **第三位階 (校級核心母法)** | **辦法 (Regulation)**／**學則** | **校務會議**審議通過、校長公布發布 | 涉及師生重大權益、全校性體系及各委員會設置母法 | 《學則》、《產學合作實施辦法》、《教師聘任及升等辦法》、《學生獎懲辦法》 |")
    lines.append("| **第四位階 (業務執行子法)** | **要點 (Directions)**／**準則** | **行政會議**／**校教評會**審議通過 | 依據母法或組織規程授權訂定之具體作業要件與審查基準 | 《專利申請要點》、《教師合聘要點》、《兼職回饋金要點》、《場地管理要點》 |")
    lines.append("| **第五位階 (作業規範)** | **實施細則／規約／須知** | **處務會議**／**中心會議**／**館務會議**通過 | 執行技術細節、設備借用規則、內部管理標準表單 | 《培育室使用管理要點》、《圖書館閱覽服務要點》、《實驗室安全規則》 |")
    lines.append("\n---\n")

    lines.append("## 二、全校六大核心領域拓撲架構圖 (Mermaid 架構圖)\n")
    lines.append("為釐清全校跨單位法規之母子衍生關聯，以下繪製六大核心行政與教學領域之法規位階拓撲結構：\n")

    # 2.1 學校治理與根本組織規程體系
    lines.append("### 2.1 學校治理與組織根本規程體系\n")
    lines.append("```mermaid")
    lines.append("graph TD")
    lines.append("    subgraph Level1 [國家中央上位法源]")
    lines.append("        L1_UNI[\"大學法 / 私立學校法\"]")
    lines.append("    end")
    lines.append("    subgraph Level2 [校級根本規程]")
    lines.append("        L2_ORG[\"輔英科技大學組織規程 (114.12.5教育部核定本)\"]")
    lines.append("    end")
    lines.append("    subgraph Level3 [校級治理核心母法 - 校務會議]")
    lines.append("        M_COUNCIL[\"校務會議議事規則\"]")
    lines.append("        M_PRESIDENT[\"校長選聘及解聘辦法\"]")
    lines.append("        M_SELF_EVAL[\"自我評鑑實施辦法\"]")
    lines.append("    end")
    lines.append("    subgraph Level4 [行政業務子法 - 行政會議]")
    lines.append("        S_ADMIN[\"行政會議議事要點\"]")
    lines.append("        S_CAMPUS_SEC[\"校園安全防救委員會設置辦法\"]")
    lines.append("        S_INTER_AUDIT[\"內部控制實施要點\"]")
    lines.append("    end")
    lines.append("    L1_UNI --> L2_ORG")
    lines.append("    L2_ORG --> M_COUNCIL")
    lines.append("    L2_ORG --> M_PRESIDENT")
    lines.append("    L2_ORG --> M_SELF_EVAL")
    lines.append("    L2_ORG --> S_ADMIN")
    lines.append("    L2_ORG --> S_CAMPUS_SEC")
    lines.append("    L2_ORG --> S_INTER_AUDIT")
    lines.append("```\n")

    # 2.2 教務教學與學則體系
    lines.append("### 2.2 教務教學與學則體系\n")
    lines.append("```mermaid")
    lines.append("graph TD")
    lines.append("    subgraph Level1 [國家中央上位法源]")
    lines.append("        L1_DEGREE[\"學位授予法 / 大學法施行細則\"]")
    lines.append("    end")
    lines.append("    subgraph Level3 [教務根本母法 - 校務會議]")
    lines.append("        M_ACAD[\"輔英科技大學學則\"]")
    lines.append("        M_ADMIS[\"各項招生規定 (單招/甄試/申請入學)\"]")
    lines.append("        M_DEGREE_EXAM[\"研究所學位考試辦法\"]")
    lines.append("    end")
    lines.append("    subgraph Level4 [教務執行子法 - 行政會議/教務會議]")
    lines.append("        S_SELECT[\"學生選課要點\"]")
    lines.append("        S_CREDIT[\"學生抵免學分要點\"]")
    lines.append("        S_MINOR[\"學生修讀雙主修/輔系(科)要點\"]")
    lines.append("        S_EARLY[\"學生提前畢業要點\"]")
    lines.append("        S_INTER[\"跨領域學分學程實施要點\"]")
    lines.append("        S_ACAD_ETHIC[\"碩士學位論文違反學術倫理處理要點\"]")
    lines.append("    end")
    lines.append("    L1_DEGREE --> M_ACAD")
    lines.append("    M_ACAD --> M_ADMIS")
    lines.append("    M_ACAD --> M_DEGREE_EXAM")
    lines.append("    M_ACAD --> S_SELECT")
    lines.append("    M_ACAD --> S_CREDIT")
    lines.append("    M_ACAD --> S_MINOR")
    lines.append("    M_ACAD --> S_EARLY")
    lines.append("    M_ACAD --> S_INTER")
    lines.append("    M_DEGREE_EXAM --> S_ACAD_ETHIC")
    lines.append("```\n")

    # 2.3 學生事務與校園生活體系
    lines.append("### 2.3 學生事務與校園生活權益體系\n")
    lines.append("```mermaid")
    lines.append("graph TD")
    lines.append("    subgraph Level1 [國家中央上位法源]")
    lines.append("        L1_STU_AID[\"學生輔導法 / 性別平等教育法\"]")
    lines.append("    end")
    lines.append("    subgraph Level3 [學務核心母法 - 校務會議]")
    lines.append("        M_STU_REWARD[\"學生獎懲辦法\"]")
    lines.append("        M_GENDER[\"性別平等教育委員會設置辦法\"]")
    lines.append("        M_STU_APPEAL[\"學生申訴評議委員會組織及運作辦法\"]")
    lines.append("        M_SCHOLAR[\"學生就學獎助學金實施辦法\"]")
    lines.append("    end")
    lines.append("    subgraph Level4 [學務執行子法 - 行政會議/學務會議]")
    lines.append("        S_CLUB[\"學生社團輔導與活動要點\"]")
    lines.append("        S_DORM[\"學生宿舍輔導與管理要點\"]")
    lines.append("        S_LOAN[\"就學貸款作業要點\"]")
    lines.append("        S_PEACE[\"學生安心就學專案助學金要點\"]")
    lines.append("    end")
    lines.append("    L1_STU_AID --> M_GENDER")
    lines.append("    L1_STU_AID --> M_STU_APPEAL")
    lines.append("    M_STU_REWARD --> S_CLUB")
    lines.append("    M_STU_REWARD --> S_DORM")
    lines.append("    M_SCHOLAR --> S_LOAN")
    lines.append("    M_SCHOLAR --> S_PEACE")
    lines.append("```\n")

    # 2.4 人事師資與聘任升等體系
    lines.append("### 2.4 人事師資、聘任與升等體系\n")
    lines.append("```mermaid")
    lines.append("graph TD")
    lines.append("    subgraph Level1 [國家中央上位法源]")
    lines.append("        L1_TEACHER[\"教師法 / 勞動基準法 / 私校退撫條例\"]")
    lines.append("    end")
    lines.append("    subgraph Level3 [人事核心母法 - 校務會議]")
    lines.append("        M_T_HIRE[\"教師聘任及升等辦法\"]")
    lines.append("        M_EVAL[\"教師評鑑辦法\"]")
    lines.append("        M_APPEAL[\"教師申訴評議委員會組織及評議要點\"]")
    lines.append("        M_LABOR[\"適用勞基法人員工作規則\"]")
    lines.append("    end")
    lines.append("    subgraph Level4 [人事執行子法 - 校教評會/行政會議]")
    lines.append("        S_EXTEND[\"專任教師年滿65歲延長服務要點\"]")
    lines.append("        S_COMBINE[\"教師合聘要點\"]")
    lines.append("        S_UPGRADE[\"兼任教師聘任及升等要點\"]")
    lines.append("        S_TRAIN[\"專任教師進修要點\"]")
    lines.append("    end")
    lines.append("    L1_TEACHER --> M_T_HIRE")
    lines.append("    L1_TEACHER --> M_EVAL")
    lines.append("    L1_TEACHER --> M_LABOR")
    lines.append("    M_T_HIRE --> S_EXTEND")
    lines.append("    M_T_HIRE --> S_COMBINE")
    lines.append("    M_T_HIRE --> S_UPGRADE")
    lines.append("    M_T_HIRE --> S_TRAIN")
    lines.append("```\n")

    # 2.5 研發成果、產學合作與智財體系
    lines.append("### 2.5 研發成果、產學合作與智財管理體系\n")
    lines.append("```mermaid")
    lines.append("graph TD")
    lines.append("    subgraph Level1 [國家中央上位法源]")
    lines.append("        L1_PATENT[\"科學技術基本法 / 專利法 / 產學合作實施辦法\"]")
    lines.append("    end")
    lines.append("    subgraph Level3 [研發核心母法 - 校務會議]")
    lines.append("        M_INDUSTRY[\"產學合作實施辦法\"]")
    lines.append("        M_IPR[\"研發成果管理辦法\"]")
    lines.append("        M_DERIVE[\"校園衍生企業申設管理辦法\"]")
    lines.append("        M_ETHIC[\"學術誠信與研究倫理委員會設置辦法\"]")
    lines.append("    end")
    lines.append("    subgraph Level4 [研發執行子法 - 行政會議/研發會議]")
    lines.append("        S_PAT_APP[\"專利申請與維護作業要點\"]")
    lines.append("        S_TECH_TRANS[\"技術移轉作業要點\"]")
    lines.append("        S_INCENTIVE[\"教師學術及專業研發成果獎勵要點\"]")
    lines.append("        S_MOU[\"產官學合作備忘錄簽訂作業要點\"]")
    lines.append("        S_COI[\"研發成果利益衝突迴避及資訊揭露要點\"]")
    lines.append("    end")
    lines.append("    L1_PATENT --> M_INDUSTRY")
    lines.append("    L1_PATENT --> M_IPR")
    lines.append("    M_INDUSTRY --> M_DERIVE")
    lines.append("    M_IPR --> S_PAT_APP")
    lines.append("    M_IPR --> S_TECH_TRANS")
    lines.append("    M_IPR --> S_COI")
    lines.append("    M_INDUSTRY --> S_MOU")
    lines.append("    M_INDUSTRY --> S_INCENTIVE")
    lines.append("```\n")

    # 2.6 圖資、總務與行政支援體系
    lines.append("### 2.6 圖書資訊、總務資產與行政支援體系\n")
    lines.append("```mermaid")
    lines.append("graph TD")
    lines.append("    subgraph Level1 [國家上位法源]")
    lines.append("        L1_INFO[\"資通安全管理法 / 個人資料保護法 / 檔案法\"]")
    lines.append("    end")
    lines.append("    subgraph Level3 [行政支援核心母法 - 校務會議]")
    lines.append("        M_LIB[\"圖書館規程 / 館藏發展政策\"]")
    lines.append("        M_SEC_POLICY[\"資訊安全政策\"]")
    lines.append("        M_PROPERTY[\"財產管理辦法\"]")
    lines.append("    end")
    lines.append("    subgraph Level4 [支援執行子法 - 行政會議]")
    lines.append("        S_BORROW[\"圖書資源設備借用管理要點\"]")
    lines.append("        S_NET[\"校園網路使用規範\"]")
    lines.append("        S_PLACE[\"場地管理要點\"]")
    lines.append("        S_SAFE[\"實驗室安全衛生管理要點\"]")
    lines.append("    end")
    lines.append("    L1_INFO --> M_SEC_POLICY")
    lines.append("    M_LIB --> S_BORROW")
    lines.append("    M_SEC_POLICY --> S_NET")
    lines.append("    M_PROPERTY --> S_PLACE")
    lines.append("```\n")

    lines.append("---\n")

    # 3. 全校法規位階與母法對照全冊
    lines.append("## 三、全校法規位階與母法參照全清冊 (全校 361 筆)\n")
    lines.append("下表彙整全校 21 個單位全數 361 筆受檢法規之位階層級、核定會議、第一位階上位法源與校內母法依據：\n")
    lines.append("| 序號 | 業管單位 | 法規名稱 | 位階層級 | 審議／核定會議 | 母法依據／上位法源 | 第一條授權／訂定意旨摘要 |")
    lines.append("| :---: | :--- | :--- | :--- | :--- | :--- | :--- |")

    for idx, d in enumerate(docs, 1):
        sheet = d['sheet']
        name = d['reg_name']
        rank = d.get('rank', '業務執行子法 (Level 4: 行政會議)')
        body = d.get('approving_body', '行政會議')
        mother = "、".join(d.get('mother_laws', [])) if d.get('mother_laws') else "依組織規程/校務行政需要"
        art1 = d.get('art1_summary', '依相關法令及本校規章訂定之')
        # clean art1 text
        art1_clean = re.sub(r'[\r\n\t]+', ' ', art1).strip()[:70]
        lines.append(f"| {idx:03d} | {sheet} | **{name}** | {rank} | {body} | {mother} | {art1_clean}... |")

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"[MD 1/4] 已產出：{out_path}")

# -----------------------------------------------------------------------------
# 2. 產生《法規格式檢核報告.md》
# -----------------------------------------------------------------------------
def generate_format_report_md():
    out_path = os.path.join(BASE_DIR, '法規格式檢核報告.md')
    lines = []
    lines.append("# 輔英科技大學法規格式檢核報告\n")
    lines.append("> **檢核依據**：《法規檢核五個條件_更新.docx》第二項條件：「法規格式是否完整(須不含底線、刪除線及各式顏色標示等)。」  ")
    lines.append(f"> **受檢對象**：輔英科技大學全校 21 個單位、全體 {len(docs)} 筆法規文檔  ")
    lines.append("> **基準日期**：民國 115 年 9 月\n")
    lines.append("---\n")

    fmt_fails = [d for d in docs if d['c3'] == 'X']
    fmt_passes = [d for d in docs if d['c3'] == 'V']

    total = len(docs)
    pass_cnt = len(fmt_passes)
    fail_cnt = len(fmt_fails)
    fail_rate = (fail_cnt / total * 100) if total else 0

    lines.append("## 一、全校法規格式檢核總覽統計\n")
    lines.append("依據正式公文發布規格與標準規範，正式公布之法規定稿須為乾淨統一格式，**不得殘存修訂草案階段之底線（代表新增）、刪除線（代表刪除）、校對用彩色字（紅/藍字）或 Word 追蹤修訂 XML 標記及 PDF 註解**。\n")
    lines.append("| 檢核指標項目 | 受檢總數 | 格式合規件數 | 格式瑕疵件數 | 瑕疵比率 | 處置要求 |")
    lines.append("| :--- | :---: | :---: | :---: | :---: | :--- |")
    lines.append(f"| **全校法規格式完整性 (總評)** | **{total} 件** | **{pass_cnt} 件** | **{fail_cnt} 件** | **{fail_rate:.1f}%** | 須由業管單位清理標記後重新上傳定稿 |")

    # Detail bug counts
    c_under = sum(1 for d in fmt_fails if any('底線' in b for b in d.get('format_bugs', [])))
    c_strike = sum(1 for d in fmt_fails if any('刪除線' in b for b in d.get('format_bugs', [])))
    c_color = sum(1 for d in fmt_fails if any('顏色' in b for b in d.get('format_bugs', [])))
    c_track = sum(1 for d in fmt_fails if any('修訂標籤' in b for b in d.get('format_bugs', [])))
    c_annot = sum(1 for d in fmt_fails if any('註解' in b for b in d.get('format_bugs', [])))

    lines.append(f"| 殘留**文字顏色標示 (RGB Color)** | {total} 件 | {total - c_color} 件 | **{c_color} 件** | {c_color/total*100:.1f}% | 含紅色(#FF0000)、藍色(#0000FF)等非純黑字體 |")
    lines.append(f"| 殘留**底線文字 (Underline)** | {total} 件 | {total - c_under} 件 | **{c_under} 件** | {c_under/total*100:.1f}% | 殘存修訂草案劃記之底線 |")
    lines.append(f"| 殘留**追蹤修訂 XML 標籤 (Track Changes)** | {total} 件 | {total - c_track} 件 | **{c_track} 件** | {c_track/total*100:.1f}% | 文檔內嵌 `<w:del>` / `<w:ins>` 追蹤修訂標籤 |")
    lines.append(f"| 殘留**刪除線 (Strikethrough)** | {total} 件 | {total - c_strike} 件 | **{c_strike} 件** | {c_strike/total*100:.1f}% | 舊條文未清除乾淨之劃線 |")
    lines.append(f"| 殘留**PDF 批註／註解標記 (Annotation)** | {total} 件 | {total - c_annot} 件 | **{c_annot} 件** | {c_annot/total*100:.1f}% | 包含便利貼批註或劃記註解 |")
    lines.append("\n---\n")

    # Per-unit statistics
    lines.append("## 二、各單位法規格式合格率橫向對比表\n")
    lines.append("| 序號 | 業管單位名稱 | 列管法規總數 | 格式合規件數 | 格式瑕疵件數 | 單位瑕疵率 | 優先處置建議 |")
    lines.append("| :---: | :--- | :---: | :---: | :---: | :---: | :--- |")

    by_unit = defaultdict(list)
    for d in docs:
        by_unit[d['sheet']].append(d)

    for u_idx, (unit_name, u_docs) in enumerate(sorted(by_unit.items(), key=lambda x: -len([d for d in x[1] if d['c3'] == 'X'])), 1):
        u_tot = len(u_docs)
        u_fail = sum(1 for d in u_docs if d['c3'] == 'X')
        u_pass = u_tot - u_fail
        u_rate = u_fail / u_tot * 100 if u_tot else 0
        status_tag = "🔴 立即排程清理" if u_fail >= 5 else ("🟡 列入例行整新" if u_fail > 0 else "🟢 格式維持良好")
        lines.append(f"| {u_idx:02d} | {unit_name} | {u_tot} 件 | {u_pass} 件 | **{u_fail} 件** | {u_rate:.1f}% | {status_tag} |")

    lines.append("\n---\n")

    # Detailed non-compliant registry
    lines.append(f"## 三、格式瑕疵法規全面檢核清冊 (共 {fail_cnt} 案)\n")
    lines.append("下表列出全校受檢文檔中具體存在格式標記之法規項目，業管承辦人可依「瑕疵類型」進行逐項排查：\n")
    lines.append("| 序號 | 業管單位 | 法規名稱 | 文件格式 | 檢核出的格式瑕疵項目 | 建議處置行動 |")
    lines.append("| :---: | :--- | :--- | :---: | :--- | :--- |")

    for f_idx, d in enumerate(fmt_fails, 1):
        unit = d['sheet']
        name = d['reg_name']
        ftype = d.get('file_type', 'PDF')
        bugs_str = "、".join(d.get('format_bugs', []))
        fix_action = "移除紅藍字色彩，套用自動(黑色)字體" if '顏色' in bugs_str else "接受修訂並移除底線/刪除線"
        if '標籤' in bugs_str:
            fix_action = "接受所有修訂並停止追蹤修訂功能"
        lines.append(f"| {f_idx:02d} | {unit} | **{name}** | {ftype} | 🔴 {bugs_str} | {fix_action} |")

    lines.append("\n---\n")
    lines.append("## 四、法規文本格式標準清理作業程序 (SOP)\n")
    lines.append("為確保本校法規文檔格式符合教育部評鑑標準及行政公文正式發布準則，請各單位依以下三步驟執行定稿清洗：\n")
    lines.append("1. **清除追蹤修訂**：開啟 Word 檔 -> 點選「校閱」標籤頁 -> 於「變更」區塊點選「接受」向下選單 -> 選擇「**接受所有變更並停止追蹤**」。\n")
    lines.append("2. **統一字體色彩**：按 `Ctrl + A` 全選整份文件 -> 常用標籤頁 -> 字型色彩下拉選單選擇「**自動**」（或純黑 `#000000`），避免留存審稿用紅字、藍字。\n")
    lines.append("3. **清除未結底線與刪除線**：檢查條文各款項，按 `Ctrl + U` 兩次確認取消全選底線，按 `Ctrl + D` 檢查取消「刪除線」勾選，最後「另存新檔」轉存標準 PDF 上網公告。\n")

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"[MD 2/4] 已產出：{out_path}")

# -----------------------------------------------------------------------------
# 3. 產生《法規逾期未更新標註檔.md》
# -----------------------------------------------------------------------------
def generate_overdue_report_md():
    out_path = os.path.join(BASE_DIR, '法規逾期未更新標註檔.md')
    lines = []
    lines.append("# 輔英科技大學法規逾期未更新標註檔\n")
    lines.append("> **檢核依據**：《法規檢核五個條件_更新.docx》第三項條件：「以目前的年份為基準, 若法規超過2年以上未更新，請建立標註檔。」  ")
    lines.append("> **基準年份**：民國 **115 年**（西元 2026 年）  ")
    lines.append("> **逾期門檻**：最後修訂年份 $\le$ 民國 **112 年**（超過 2 年未更新）或未載明修訂日期者  ")
    lines.append("> **標註狀態**：【**逾期需檢討更新 - 建議排入 115 學年度行政會議／校務會議提案**】\n")
    lines.append("---\n")

    overdue_docs = [d for d in docs if d['c4'] == 'V']
    normal_docs = [d for d in docs if d['c4'] == 'X']

    total = len(docs)
    overdue_cnt = len(overdue_docs)
    normal_cnt = len(normal_docs)
    overdue_rate = (overdue_cnt / total * 100) if total else 0

    lines.append("## 一、全校逾期未更新法規整體概況\n")
    lines.append("為落實法規動態檢討管理機制，避免規章條文與當前外部法律、內部組織架構脫節，原則上規章應每 2 年配合時空環境進行通盤檢視。\n")
    lines.append("| 項目類別 | 法規件數 | 佔全校法規比例 | 管理意涵與處置措施 |")
    lines.append("| :--- | :---: | :---: | :--- |")
    lines.append(f"| **超過 2 年以上未更新 (本標註檔對象)** | **{overdue_cnt} 件** | **{overdue_rate:.1f}%** | 🔴 **全面建立標註檔**，分二梯次排入 115 學年度提案修訂 |")
    lines.append(f"| 近 2 年內曾更新法規 (113~115 年) | {normal_cnt} 件 | {100.0 - overdue_rate:.1f}% | 🟢 時效符合規定，維持常態維護運作 |")
    lines.append(f"| **全校列管法規總數** | **{total} 件** | 100.0% | 全校 21 個單位受檢法規 |")
    lines.append("\n---\n")

    # Priority definitions
    lines.append("## 二、逾期未更新急迫性三級標註標準\n")
    lines.append("本標註檔依據法規最後修訂距今年數、是否涉及中央部會廢止機關名稱（如科技部、衛生署）或本校組織規程更名，劃分三級提案急迫性：\n")
    lines.append("- 🔴 **高急迫性（Priority High）**：未更新年數 $\ge$ 7 年（最後修訂在 108 年以前），或條文中含有「科技部」、「行政院衛生署」、「研究發展處」等已被組織規程與中央部會更名之過時稱謂。此類規章極易引發法源瑕疵，應排入 **115學年度第一學期 (115-1)** 提案修訂。\n")
    lines.append("- 🟡 **中急迫性（Priority Medium）**：未更新年數介於 4 至 6 年（修訂於 109 ~ 111 年間），部分金額補助標準、表單作業程序已與現況落差，建議排入 **115學年度第二學期 (115-2)** 提案。\n")
    lines.append("- 🟢 **低急迫性／常態檢討（Priority Low）**：未更新年數為 3 年（修訂於 112 年），運作尚屬穩定，建議由各單位於本次例行巡檢中併案檢視文字即可。\n")
    lines.append("\n---\n")

    # Breakdown by urgency
    high_urgency = []
    med_urgency = []
    low_urgency = []

    for d in overdue_docs:
        yr = d.get('last_year')
        has_org = d.get('has_org_change', False)
        if yr is None or yr <= 108 or has_org:
            high_urgency.append(d)
        elif yr in [109, 110, 111]:
            med_urgency.append(d)
        else:
            low_urgency.append(d)

    lines.append("## 三、逾期急迫性分級統計看板\n")
    lines.append("| 急迫性分級 | 法規件數 | 佔逾期比例 | 建議提案會議排程 | 提案主責單位主力 |")
    lines.append("| :--- | :---: | :---: | :--- | :--- |")
    lines.append(f"| 🔴 **高急迫性 (High)** | **{len(high_urgency)} 件** | {len(high_urgency)/overdue_cnt*100:.1f}% | **115 學年度第 1 學期** 行政會議／校務會議 | 教務處、學務處、研發處、人事室 |")
    lines.append(f"| 🟡 **中急迫性 (Medium)** | **{len(med_urgency)} 件** | {len(med_urgency)/overdue_cnt*100:.1f}% | **115 學年度第 2 學期** 行政會議 | 教務處、學務處、圖資處、人事室 |")
    lines.append(f"| 🟢 **低急迫性 (Low)** | **{len(low_urgency)} 件** | {len(low_urgency)/overdue_cnt*100:.1f}% | 併入各處室 115 學年度常態會期檢討 | 全校各單位常態檢討 |")
    lines.append("\n---\n")

    # Detailed inventory
    lines.append(f"## 四、全校逾期未更新法規全清冊 (共 {overdue_cnt} 案)\n")
    lines.append("| 序號 | 業管單位 | 法規名稱 | 最後修訂日期 | 未修訂年數 | 急迫性分級 | 建議提案排程 | 備註說明 |")
    lines.append("| :---: | :--- | :--- | :---: | :---: | :---: | :--- | :--- |")

    for o_idx, d in enumerate(sorted(overdue_docs, key=lambda x: (x.get('last_year') or 0)), 1):
        unit = d['sheet']
        name = d['reg_name']
        dt = d.get('last_date', 'N/A')
        yr = d.get('last_year')
        diff = (115 - yr) if yr else '未詳'
        has_org = d.get('has_org_change', False)
        
        if yr is None or yr <= 108 or has_org:
            p_badge = "🔴 高急迫"
            sched = "115-1 提案"
        elif yr in [109, 110, 111]:
            p_badge = "🟡 中急迫"
            sched = "115-2 提案"
        else:
            p_badge = "🟢 低急迫"
            sched = "常規檢視"
            
        note_snippet = d.get('notes', '無特別說明')
        lines.append(f"| {o_idx:03d} | {unit} | **{name}** | {dt} | {diff} 年 | {p_badge} | {sched} | {note_snippet} |")

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"[MD 3/4] 已產出：{out_path}")

# -----------------------------------------------------------------------------
# 4. 產生《法規檢核結果總表.md》
# -----------------------------------------------------------------------------
def generate_summary_report_md():
    out_path = os.path.join(BASE_DIR, '法規檢核結果總表.md')
    lines = []
    lines.append("# 輔英科技大學法規檢核結果總表\n")
    lines.append("> **檢核依據**：《法規檢核五個條件_更新.docx》及本校《115年法規清單檢核表》作業標準  ")
    lines.append(f"> **受檢對象**：輔英科技大學全校 21 個單位、全體 {len(docs)} 筆法規項目（編號 001 ~ {len(docs):03d}）  ")
    lines.append("> **基準日期**：民國 115 年 9 月\n")
    lines.append("---\n")

    total = len(docs)
    c1_cnt = sum(1 for d in docs if d['c1'] == 'V')
    c2_cnt = sum(1 for d in docs if d['c2'] == 'V')
    c3_pass = sum(1 for d in docs if d['c3'] == 'V')
    c3_fail = sum(1 for d in docs if d['c3'] == 'X')
    c4_v = sum(1 for d in docs if d['c4'] == 'V')
    c5_v = sum(1 for d in docs if d['c5'] == 'V')

    lines.append("## 一、五大檢核條件意涵與判定標準說明\n")
    lines.append("依據教育部大學評鑑規範與本校自我檢核標準，五大條件定義如下：\n")
    lines.append("1. **條件 1 (位階與現行版本)**：檢測母法與子法位階關係是否明確，且是否已更新為現行最新版本（建立《法規位階關係檔.md》）。\n")
    lines.append("2. **條件 2 (法規名稱正確性)**：法規名稱是否正確，核對官網最新公布名稱與歷史更名（原：舊有名稱）對照。\n")
    lines.append("3. **條件 3 (格式完整性)**：法規定稿格式是否完整，**須不含底線、刪除線及各式顏色標示、Word 追蹤修訂標籤**（詳見《法規格式檢核報告.md》）。\n")
    lines.append("4. **條件 4 (時效性與提案標註)**：以 115 年為基準，**若法規超過 2 年以上未更新（最後修訂 $\le$ 112 年），是否需要更新並建立標註檔**（詳見《法規逾期未更新標註檔.md》）。\n")
    lines.append("5. **條件 5 (組織規程因應)**：因應校務會議與董事會通過之組織規程修訂案（如研發處更名為研究暨產學發展處、電算中心更名為圖資處）及外部中央主管機關（如科技部改制為國科會），是否需要修正條文。\n")
    lines.append("\n**符號定義說明**：\n")
    lines.append("- `V`：符合檢核條件，或該項條件評估為「需要更新提案／已建立標註」。\n")
    lines.append("- `X`：不符合條件（如格式存在瑕疵），或該項條件為「不需要更新提案／時效正常」。\n")
    lines.append("\n---\n")

    lines.append("## 二、全校法規關鍵檢核指標看板 (KPI Dashboard)\n")
    lines.append("| 關鍵指標名稱 | 統計件數 | 佔比率 | 狀態評估與風險層級 | 管理處置方向 |")
    lines.append("| :--- | :---: | :---: | :---: | :--- |")
    lines.append(f"| **受檢法規總量** | **{total} 案** | 100.0% | 完成全文本掃描檢核 | 全校 21 個單位全面列管 |")
    lines.append(f"| **條件1 最新版本與位階 (C1)** | {c1_cnt} 案 | {c1_cnt/total*100:.1f}% | 🟢 位階體系已完備梳理 | 發布全校位階拓撲結構檔 |")
    lines.append(f"| **條件2 法規名稱正確 (C2)** | {c2_cnt} 案 | {c2_cnt/total*100:.1f}% | 🟢 官方正式名稱無誤 | 具備更名歷史脈絡紀錄 |")
    lines.append(f"| **條件3 格式完全合規 (C3=V)** | {c3_pass} 案 | {c3_pass/total*100:.1f}% | 🟢 乾淨正式文本 | 維持常態發布品質 |")
    lines.append(f"| **條件3 格式存在瑕疵 (C3=X)** | **{c3_fail} 案** | **{c3_fail/total*100:.1f}%** | 🔴 含底線/顏色/修訂標記 | **即刻執行文本定稿清理** |")
    lines.append(f"| **條件4 逾 2 年未更新標註 (C4=V)** | **{c4_v} 案** | **{c4_v/total*100:.1f}%** | 🟡 須於 115 學年度提案修訂 | **分二梯次排入行政會議提案** |")
    lines.append(f"| **條件5 組織規程修訂提案 (C5=V)** | **{c5_v} 案** | **{c5_v/total*100:.1f}%** | 🔴 條文留存廢止機關名稱 | **併入 115-1 優先提案修訂** |")
    lines.append("\n---\n")

    # Unit comparison table
    lines.append("## 三、全校 21 個單位檢核成果橫向對比表\n")
    lines.append("| 序號 | 業管單位名稱 | 法規總數 | 格式合規 (C3=V) | 格式瑕疵 (C3=X) | 逾期需提案 (C4=V) | 組織因應提案 (C5=V) | 整體改善重點 |")
    lines.append("| :---: | :--- | :---: | :---: | :---: | :---: | :---: | :--- |")

    by_unit = defaultdict(list)
    for d in docs:
        by_unit[d['sheet']].append(d)

    for u_idx, (unit_name, u_docs) in enumerate(sorted(by_unit.items(), key=lambda x: -len(x[1])), 1):
        u_tot = len(u_docs)
        u_c3_x = sum(1 for d in u_docs if d['c3'] == 'X')
        u_c3_v = u_tot - u_c3_x
        u_c4_v = sum(1 for d in u_docs if d['c4'] == 'V')
        u_c5_v = sum(1 for d in u_docs if d['c5'] == 'V')
        
        focus_pts = []
        if u_c3_x > 0: focus_pts.append(f"清理{u_c3_x}件格式")
        if u_c4_v > 0: focus_pts.append(f"規劃{u_c4_v}件提案")
        if u_c5_v > 0: focus_pts.append(f"更正{u_c5_v}件舊機關名")
        focus_str = "；".join(focus_pts) if focus_pts else "常規維護"

        lines.append(f"| {u_idx:02d} | {unit_name} | {u_tot} 筆 | {u_c3_v} 筆 | **{u_c3_x} 筆** | **{u_c4_v} 筆** | {u_c5_v} 筆 | {focus_str} |")

    lines.append("\n---\n")

    # Full master table
    lines.append(f"## 四、全校法規檢核總結果全清冊 (全校 361 案)\n")
    lines.append("| 序號 | 單位 | 法規名稱 | C1(版本) | C2(名稱) | C3(格式) | C4(逾期) | C5(組織) | 最後修訂日期 | 備註說明與標註摘要 |")
    lines.append("| :---: | :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :--- |")

    for idx, d in enumerate(docs, 1):
        unit = d['sheet']
        name = d['reg_name']
        c1 = d['c1']
        c2 = d['c2']
        c3 = d['c3']
        c4 = d['c4']
        c5 = d['c5']
        dt = d.get('last_date', 'N/A')
        note = d.get('notes', '無特別說明')
        lines.append(f"| {idx:03d} | {unit} | **{name}** | {c1} | {c2} | {c3} | {c4} | {c5} | {dt} | {note} |")

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"[MD 4/4] 已產出：{out_path}")

if __name__ == '__main__':
    print("=== 開始依據全校資料庫編譯產出四大 Markdown 專卷 ===")
    generate_hierarchy_md()
    generate_format_report_md()
    generate_overdue_report_md()
    generate_summary_report_md()
    print("=== 四大 Markdown 專卷產出完成！ ===")
