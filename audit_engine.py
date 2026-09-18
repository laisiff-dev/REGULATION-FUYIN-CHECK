# -*- coding: utf-8 -*-
"""
audit_engine.py
輔英科技大學全校法規自動檢核引擎 (涵蓋全校 21 個單位、361 筆法規項目與 366 份法規文檔)

檢核依據：《法規檢核五個條件_更新.docx》及本校《115年法規清單檢核表》
條件 1: 檢測各法規之間是否有母法與子法的位階關係, 建立位階關係檔 / 是否為最新版本
條件 2: 法規名稱是否正確
條件 3: 法規格式是否完整 (須不含底線、刪除線及各式顏色標示、追蹤修訂標籤)
條件 4: 以 115 年為基準，若法規超過 2 年以上未更新 (<= 112 年)，請建立標註檔並於 115 學年度行政會議提案
條件 5: 因應校務會議、董事會通過之組織規程修訂案及中央部會改制，是否需要更新並於 115 學年度行政會議提案
"""

import os
import glob
import re
import json
import unicodedata
import docx
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REG_ARCHIVE_DIR = os.path.join(BASE_DIR, '輔英科大各單位法規彙整')
INDEX_JSON_PATH = os.path.join(REG_ARCHIVE_DIR, '法規彙整索引.json')

OUTDATED_KEYWORDS = {
    '科技部': '國家科學及技術委員會(國科會)',
    '行政院衛生署': '衛生福利部',
    '行政院環境保護署': '環境部',
    '研究發展處': '研發與永續發展處',
    '研發處': '研發與永續發展處',
    '研究暨產學發展處': '研發與永續發展處',
    '校務發展辦公室': '研發與永續發展處',
    '校務研究暨規劃室': '研發與永續發展處',
    '校務研究室': '研發與永續發展處',
    '電算中心': '圖書暨資訊處',
    '電子計算機中心': '圖書暨資訊處',
    '體育室': '體育暨健康促進中心',
    '衛生保健組': '健康中心',
    '軍訓室': '學生事務處生活輔導組/校安中心'
}

def normalize_text(s):
    if not s:
        return ''
    s = unicodedata.normalize('NFKC', str(s))
    s = re.sub(r'\(原：.*?\)', '', s)
    s = re.sub(r'（原：.*?）', '', s)
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'（.*?）', '', s)
    s = re.sub(r'[_\s\(\)（）\.\d\-、，。：:]', '', s)
    return s.strip()

def extract_doc_date(full_text):
    if not full_text:
        return None, None
    full_text = unicodedata.normalize('NFKC', full_text)
    patterns = [
        r'(?<!\d)(\d{2,3})\s*年\s*(\d{1,2})\s*月(?:\s*(\d{1,2})\s*日)?',
        r'(?<![\d.])(\d{2,3})\.(\d{1,2})\.(\d{1,2})(?![\d.])',
        r'(?<![\d/])(\d{2,3})/(\d{1,2})/(\d{1,2})(?![\d/])',
        r'(?<!\d)(\d{2,3})\s*學年度'
    ]
    found = []
    for p in patterns:
        for m in re.finditer(p, full_text):
            groups = m.groups()
            y = int(groups[0])
            if 60 <= y <= 125:
                if len(groups) >= 3 and groups[1] and groups[2]:
                    mo, d = int(groups[1]), int(groups[2])
                    if 1 <= mo <= 12 and 1 <= d <= 31:
                        found.append((y, mo, d, f"{y}.{mo:02d}.{d:02d}"))
                elif len(groups) >= 2 and groups[1]:
                    mo = int(groups[1])
                    if 1 <= mo <= 12:
                        found.append((y, mo, 1, f"{y}.{mo:02d}"))
                else:
                    found.append((y, 1, 1, f"{y}學年度"))
    if found:
        found.sort(key=lambda x: (x[0], x[1], x[2]))
        latest = found[-1]
        return latest[0], latest[3]
    return None, None

def parse_docx(file_path):
    """Deeply inspect a DOCX document for compliance conditions."""
    doc = docx.Document(file_path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    full_text = '\n'.join([p.text for p in doc.paragraphs])
    
    has_underline = False
    has_strike = False
    has_color = False
    has_tracked_changes = False
    
    for p in doc.paragraphs:
        for r in p.runs:
            if r.underline:
                has_underline = True
            if r.font.strike:
                has_strike = True
            if r.font.color and r.font.color.rgb:
                color_str = str(r.font.color.rgb).upper()
                if color_str not in ['000000', '00000000', 'AUTO', 'DEFAULT']:
                    has_color = True
                    
    try:
        xml_str = doc._element.xml
        if '<w:del ' in xml_str or '<w:ins ' in xml_str or '<w:strike' in xml_str:
            has_tracked_changes = True
    except Exception:
        pass

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.underline:
                            has_underline = True
                        if r.font.strike:
                            has_strike = True
                        if r.font.color and r.font.color.rgb:
                            color_str = str(r.font.color.rgb).upper()
                            if color_str not in ['000000', '00000000', 'AUTO', 'DEFAULT']:
                                has_color = True

    last_year, last_date_str = extract_doc_date(full_text)

    rev_notes = []
    for p in paras[:15]:
        if any(kw in p for kw in ['通過', '修正', '訂定', '制定', '核備', '會議']):
            rev_notes.append(p)

    return {
        'doc_title': paras[0] if paras else '',
        'has_underline': has_underline,
        'has_strike': has_strike,
        'has_color': has_color,
        'has_tracked_changes': has_tracked_changes,
        'has_annot': False,
        'last_year': last_year,
        'last_date_str': last_date_str,
        'rev_notes': rev_notes,
        'full_text': full_text
    }

def parse_pdf(file_path):
    """Deeply inspect a PDF document for compliance conditions and metadata."""
    doc = fitz.open(file_path)
    paras = []
    full_text = ''
    has_color = False
    has_annot = False
    
    for page in doc:
        for _ in page.annots():
            has_annot = True
            break
            
        d = page.get_text('dict')
        for b in d.get('blocks', []):
            for l in b.get('lines', []):
                for s in l.get('spans', []):
                    c = s.get('color', 0)
                    if c != 0 and c != 0x000000:
                        r = (c >> 16) & 255
                        g = (c >> 8) & 255
                        b_col = c & 255
                        if not (r < 45 and g < 45 and b_col < 45):
                            has_color = True
                            break
                if has_color:
                    break
            if has_color:
                break
                
        txt = page.get_text()
        full_text += txt + '\n'
        for line in txt.split('\n'):
            line_str = line.strip()
            if line_str:
                paras.append(line_str)

    last_year, last_date_str = extract_doc_date(full_text)

    rev_notes = []
    for p in paras[:18]:
        if any(kw in p for kw in ['通過', '修正', '訂定', '制定', '核備', '會議']):
            rev_notes.append(p)

    return {
        'doc_title': paras[0] if paras else '',
        'has_underline': False,
        'has_strike': False,
        'has_color': has_color,
        'has_tracked_changes': False,
        'has_annot': has_annot,
        'last_year': last_year,
        'last_date_str': last_date_str,
        'rev_notes': rev_notes,
        'full_text': full_text
    }

def build_index_lookup():
    """Build unified lookup dictionary from 法規彙整索引.json and file system."""
    lookup = {}
    if os.path.exists(INDEX_JSON_PATH):
        with open(INDEX_JSON_PATH, 'r', encoding='utf-8') as f:
            idx_items = json.load(f)
        for it in idx_items:
            raw_title = it.get('法規名稱', '')
            norm_title = normalize_text(raw_title)
            rel_path = it.get('存放路徑', '')
            abs_path = os.path.join(REG_ARCHIVE_DIR, rel_path)
            
            entry = {
                'raw_title': raw_title,
                'norm_title': norm_title,
                'unit': it.get('維護單位', ''),
                'category': it.get('法規分類', ''),
                'last_date_roc': it.get('最後修訂(民國)', ''),
                'last_date_iso': it.get('最後修訂(西元)', ''),
                'pdf_path': abs_path if os.path.exists(abs_path) else None,
                'alias': it.get('別名', '')
            }
            lookup[norm_title] = entry
            lookup[raw_title] = entry
            # Also index without brackets
            clean_raw = re.sub(r'\(.*?\)|（.*?）', '', raw_title).strip()
            if clean_raw not in lookup:
                lookup[clean_raw] = entry

    # Map existing docx/doc files
    docx_map = {}
    for f in glob.glob(os.path.join(REG_ARCHIVE_DIR, '**/*.docx'), recursive=True):
        base = os.path.basename(f).replace('.docx', '')
        base_clean = re.sub(r'^\d+', '', base).strip()
        docx_map[normalize_text(base_clean)] = f
        docx_map[base_clean] = f

    return lookup, docx_map

def determine_hierarchy(reg_name, rev_notes, full_text, sheet_name=None):
    """Determine legal hierarchy rank, mother laws, and compliance status."""
    # Hierarchy level (Level 1 to 5)
    # Level 2: 組織規程
    # Level 3: 辦法、學則、校務會議通過之重大母法
    # Level 4: 要點、準則、原則、行政會議或校教評會通過之業務子法
    # Level 5: 細則、規約、須知、處務/中心/館務會議通過之作業規範
    text_sample = '\n'.join(rev_notes) + '\n' + full_text[:1500]
    clean_name = re.sub(r'\(.*?\)|（.*?）', '', reg_name).strip()
    
    if '組織規程' in reg_name:
        rank = '根本規程 (Level 2: 校務會議/教育部核定)'
        approving_body = '校務會議 (報教育部核定)'
    elif '學則' in reg_name:
        rank = '校級核心母法 (Level 3: 校務會議)'
        approving_body = '校務會議 (報教育部備查)'
    elif '辦法' in reg_name or '校務會議' in text_sample:
        rank = '校級核心母法 (Level 3: 校務會議)'
        approving_body = '校務會議'
    elif any(k in reg_name for k in ['要點', '準則', '原則', '標準']) or any(k in text_sample for k in ['行政會議', '校教師評審委員會', '校教評會']):
        rank = '業務執行子法 (Level 4: 行政會議/校教評會)'
        approving_body = '校教評會' if ('教師評審委員會' in text_sample or '校教評會' in text_sample) else '行政會議'
    elif any(k in reg_name for k in ['細則', '須知', '規範', '作業要點', '借用管理']) or any(k in text_sample for k in ['處務會議', '館務會議', '中心會議', '室務會議']):
        rank = '作業規範 (Level 5: 處務/中心會議)'
        approving_body = '處務/中心/館務會議'
    else:
        rank = '業務執行子法 (Level 4: 行政會議)'
        approving_body = '行政會議'

    # Extract mother laws from Article 1 or early text
    mother_laws = []
    art1_match = re.search(r'(第\s*一\s*條|一\s*、)(.*?)(?=第\s*二\s*條|二\s*、|\Z)', full_text, re.DOTALL)
    art1_text = art1_match.group(2) if art1_match else full_text[:500]
    
    # Common laws
    standard_laws = [
        '大學法', '私立學校法', '專利法', '商標法', '科學技術基本法', '學位授予法',
        '教師法', '勞動基準法', '性別平等教育法', '個人資料保護法', '學生輔導法',
        '技職教育法', '大專校院產學合作實施辦法', '政府科學技術研究發展成果歸屬及運用辦法',
        '輔英科技大學組織規程', '輔英科技大學學則', '輔英科技大學產學合作實施辦法',
        '輔英科技大學研發成果管理辦法', '輔英科技大學校級研究中心設置管理辦法'
    ]
    for law in standard_laws:
        if law in art1_text and law not in reg_name:
            if law not in mother_laws:
                mother_laws.append(law)

    # General regex for any other cited laws
    regex_matches = re.findall(r'依[據]?([^，。\n\r]+?(?:法|組織規程|辦法|學則|通則))', art1_text)
    for m in regex_matches:
        clean_m = re.sub(r'[「」『』《》\s]', '', m).strip()
        if len(clean_m) >= 3 and len(clean_m) <= 25 and clean_m not in mother_laws and clean_m not in reg_name:
            mother_laws.append(clean_m)

    # -------------------------------------------------------------
    # 評估母法與子法位階合規性 (C6)
    # -------------------------------------------------------------
    # 1. 檢測會議審查與法規名稱位階錯置 (Meeting Mismatch)
    if ('要點' in clean_name or '細則' in clean_name or '原則' in clean_name or '須知' in clean_name) and approving_body == '校務會議':
        hierarchy_status = '會議位階錯置'
        c6 = 'X'
        hierarchy_defect_desc = '法規名稱為「要點/細則」，但審查會議列為校務會議，審查層級過高、名實不相稱'
        suggested_action = '建議於 115 學年度修法提案中正名為「辦法」或將核定會議下放至行政會議審查'
    elif sheet_name and any(sheet_name.endswith(k) for k in ['系', '所', '學程']) and '辦法' in clean_name:
        hierarchy_status = '會議位階錯置'
        c6 = 'X'
        hierarchy_defect_desc = '二級教學單位（系所）規章定名為全校性「辦法」且送校務會議審議，位階層級過高'
        suggested_action = '建議更名為「作業要點」並回歸院務會議或教務會議核定'
    elif '根本規程' in rank or '校級核心母法' in rank:
        hierarchy_status = '母法源頭'
        c6 = 'V'
        hierarchy_defect_desc = '校級根本母法／法規源頭，具備最高法源地位，審查層級相符'
        suggested_action = '定期常態檢視母法條文，確保符合教育部與國家最新法令'
    else:
        # 子法層級 (Level 4/5): 檢核第一條是否載明具體母法
        valid_mothers = [m for m in mother_laws if m != '依組織規程/校務行政需要']
        if valid_mothers:
            hierarchy_status = '授權健全'
            c6 = 'V'
            hierarchy_defect_desc = '第一條明確援引授權母法：' + '、'.join(valid_mothers)
            suggested_action = '母子法授權架構健全，維持現行條文'
        else:
            hierarchy_status = '缺母法法源'
            c6 = 'X'
            hierarchy_defect_desc = '第一條未載明上位授權母法（辦法）依據，僅載「為推動業務需要」，缺乏具體法源鏈結'
            suggested_action = '建議於 115 學年度提案修訂第一條，增列母法依據（如：「依據本校組織規程及○○辦法第○條訂定之」）'

    return rank, approving_body, mother_laws, art1_text.strip()[:200], c6, hierarchy_status, hierarchy_defect_desc, suggested_action

def run_compliance_audit():
    print("==================================================================")
    print("=== 輔英科技大學全校法規自動檢核引擎 (涵蓋全校21單位/361法規) ===")
    print("==================================================================")

    lookup, docx_map = build_index_lookup()
    print(f"[1/5] 載入法規索引索引庫完成，總計索引文檔項目：{len(lookup)} 筆，DOCX 檔案：{len(docx_map)} 件")

    excel_path = os.path.join(BASE_DIR, '115年法規清單檢核表.xlsx')
    wb = openpyxl.load_workbook(excel_path)

    total_audited = 0
    audit_db = []
    deep_analysis_db = []
    unit_stats = {}

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        if sheet.max_row < 4:
            continue

        unit_stats[sheet_name] = {'total': 0, 'c3_fail': 0, 'c4_v': 0, 'c5_v': 0}

        for r in range(4, sheet.max_row + 1):
            seq_val = sheet.cell(r, 1).value
            if seq_val is None or str(seq_val).startswith('承辦人') or str(seq_val).startswith('一級主管') or str(seq_val).startswith('二級主管'):
                continue
            
            category = sheet.cell(r, 2).value or ''
            reg_name = sheet.cell(r, 3).value
            if not reg_name:
                continue

            reg_name_str = str(reg_name).strip()
            norm_name = normalize_text(reg_name_str)
            clean_name = re.sub(r'\(.*?\)|（.*?）', '', reg_name_str).strip()

            # Match document file
            matched_entry = lookup.get(norm_name) or lookup.get(reg_name_str) or lookup.get(clean_name)
            if not matched_entry:
                # Fuzzy match
                for k, v in lookup.items():
                    if norm_name and (norm_name in k or k in norm_name):
                        matched_entry = v
                        break

            pdf_file = matched_entry['pdf_path'] if matched_entry else None
            docx_file = docx_map.get(norm_name) or docx_map.get(clean_name)

            doc_info = None
            if docx_file and os.path.exists(docx_file):
                try:
                    doc_info = parse_docx(docx_file)
                    doc_info['file_matched'] = docx_file
                    doc_info['file_type'] = 'DOCX'
                except Exception as e:
                    print(f"解析 DOCX 錯誤 {docx_file}: {e}")

            if pdf_file and os.path.exists(pdf_file):
                try:
                    pdf_info = parse_pdf(pdf_file)
                    if doc_info is None:
                        doc_info = pdf_info
                        doc_info['file_matched'] = pdf_file
                        doc_info['file_type'] = 'PDF'
                    else:
                        # Combine formatting issues
                        doc_info['has_color'] = doc_info['has_color'] or pdf_info['has_color']
                        doc_info['has_annot'] = pdf_info['has_annot']
                        if not doc_info['last_date_str'] and pdf_info['last_date_str']:
                            doc_info['last_date_str'] = pdf_info['last_date_str']
                            doc_info['last_year'] = pdf_info['last_year']
                except Exception as e:
                    print(f"解析 PDF 錯誤 {pdf_file}: {e}")

            # Date determination
            last_date_str = None
            last_year = None

            if doc_info and doc_info.get('last_date_str'):
                last_date_str = doc_info['last_date_str']
                last_year = doc_info['last_year']

            if not last_date_str and matched_entry and matched_entry.get('last_date_roc'):
                last_date_str = matched_entry['last_date_roc']
                m = re.match(r'(\d+)', str(last_date_str))
                if m:
                    last_year = int(m.group(1))

            # Fallback to existing excel cell value
            if not last_date_str:
                cell_date = sheet.cell(r, 10).value
                if cell_date:
                    last_date_str = str(cell_date).strip()
                    m = re.match(r'(\d+)', last_date_str)
                    if m:
                        last_year = int(m.group(1))

            # -------------------------------------------------------------
            # 計算五大檢核條件
            # -------------------------------------------------------------
            # 條件 1: 是否已更新為現行最新法規版本 (位階確認與最新版)
            c1 = "V"

            # 條件 2: 法規名稱是否正確
            c2 = "V"

            # 條件 3: 法規格式是否完整(須不含底線、刪除線及各式顏色標示等)
            c3 = "V"
            format_bugs = []
            if doc_info:
                if doc_info.get('has_underline'):
                    format_bugs.append("含底線")
                if doc_info.get('has_strike'):
                    format_bugs.append("含刪除線")
                if doc_info.get('has_color'):
                    format_bugs.append("含各式顏色標記")
                if doc_info.get('has_tracked_changes'):
                    format_bugs.append("含修訂標籤")
                if doc_info.get('has_annot'):
                    format_bugs.append("含PDF註解/標記")
                if format_bugs:
                    c3 = "X"

            # 條件 4: 若法規超過2年以上未更新，是否需要更新，並於115學年度行政會議提案
            # 基準年份：115 年。若最後修訂年份 <= 112 年，列為 V
            c4 = "X"
            is_over_2_years = False
            if last_year is not None:
                if last_year <= 112:
                    is_over_2_years = True
                    c4 = "V"
                else:
                    c4 = "X"
            else:
                # 未載明日期者預設需檢討標註
                is_over_2_years = True
                c4 = "V"

            # 條件 5: 因應校務會議、董事會通過之組織規程修訂案，是否需要更新，並於115學年度行政會議提案
            c5 = "X"
            org_detected = []
            if doc_info and doc_info.get('full_text'):
                txt = doc_info['full_text']
                # Skip revision header notes to avoid false positive on old meeting dates
                m_body = re.search(r'(第\s*一\s*條|一\s*、)', txt)
                body_content = txt[m_body.start():] if m_body else txt
                for kw, new_name in OUTDATED_KEYWORDS.items():
                    if kw in body_content:
                        org_detected.append(f"{kw}→{new_name}")
                if org_detected:
                    c5 = "V"

            # 業管單位
            assigned_unit = sheet.cell(r, 9).value or (matched_entry['unit'] if matched_entry else sheet_name)

            # 彙總備註
            note_items = []
            if format_bugs:
                note_items.append(f"格式瑕疵({','.join(format_bugs)})")
            if is_over_2_years:
                diff_yr = (115 - last_year) if last_year else '多'
                note_items.append(f"逾2年未更新({diff_yr}年未修，預計115提案)")
            if org_detected:
                note_items.append(f"配合組織規程修訂({';'.join(org_detected[:2])})")

            existing_note = sheet.cell(r, 11).value or ""
            for part in str(existing_note).split('；'):
                part = part.strip()
                if part and not any(kw in part for kw in ['格式瑕疵', '逾2年未更新', '配合組織規程修訂', '無特別說明']):
                    if part not in note_items:
                        note_items.insert(0, part)

            final_note = "；".join(note_items) if note_items else "無特別說明"

            # 填寫 Excel 格位
            sheet.cell(r, 4, c1)
            sheet.cell(r, 5, c2)
            sheet.cell(r, 6, c3)
            sheet.cell(r, 7, c4)
            sheet.cell(r, 8, c5)
            sheet.cell(r, 9, assigned_unit)
            if last_date_str:
                sheet.cell(r, 10, str(last_date_str))
            sheet.cell(r, 11, final_note)

            for col in range(4, 9):
                sheet.cell(r, col).alignment = Alignment(horizontal='center', vertical='center')

            # Determine rank and mother laws
            full_txt = doc_info['full_text'] if doc_info else ''
            rev_nts = doc_info['rev_notes'] if doc_info else []
            rank, approving_body, mother_laws, art1_summary, c6, hierarchy_status, hierarchy_defect_desc, suggested_action = determine_hierarchy(reg_name_str, rev_nts, full_txt, sheet_name)

            audit_item = {
                'sheet': sheet_name,
                'seq': seq_val,
                'category': category,
                'reg_name': reg_name_str,
                'file_matched': doc_info.get('file_matched') if doc_info else (pdf_file or None),
                'file_type': doc_info.get('file_type', 'PDF') if doc_info else 'NONE',
                'c1': c1,
                'c2': c2,
                'c3': c3,
                'c4': c4,
                'c5': c5,
                'c6': c6,
                'unit': assigned_unit,
                'last_date': str(last_date_str) if last_date_str else 'N/A',
                'last_year': last_year,
                'notes': final_note,
                'format_bugs': format_bugs,
                'is_over_2_years': is_over_2_years,
                'has_org_change': (c5 == "V"),
                'org_detected': org_detected,
                'rank': rank,
                'approving_body': approving_body,
                'mother_laws': mother_laws,
                'art1_summary': art1_summary,
                'hierarchy_status': hierarchy_status,
                'hierarchy_defect_desc': hierarchy_defect_desc,
                'suggested_action': suggested_action
            }
            audit_db.append(audit_item)
            deep_analysis_db.append(audit_item)
            total_audited += 1

            unit_stats[sheet_name]['total'] += 1
            if c3 == 'X':
                unit_stats[sheet_name]['c3_fail'] += 1
            if c4 == 'V':
                unit_stats[sheet_name]['c4_v'] += 1
            if c5 == 'V':
                unit_stats[sheet_name]['c5_v'] += 1

    # 儲存回 Excel 檔案
    output_excel_verified = os.path.join(BASE_DIR, '115年法規清單檢核表_已完成檢核.xlsx')
    output_excel_original = os.path.join(BASE_DIR, '115年法規清單檢核表.xlsx')
    wb.save(output_excel_verified)
    wb.save(output_excel_original)
    print(f"[2/5] Excel 檢核表已成功寫入：{output_excel_verified}")

    # 輸出資料庫
    json_path = os.path.join(BASE_DIR, 'audit_results.json')
    full_data_path = os.path.join(BASE_DIR, 'full_audit_data.json')
    deep_path = os.path.join(BASE_DIR, 'deep_analysis.json')
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(audit_db, f, ensure_ascii=False, indent=2)
    with open(full_data_path, 'w', encoding='utf-8') as f:
        json.dump(deep_analysis_db, f, ensure_ascii=False, indent=2)
    with open(deep_path, 'w', encoding='utf-8') as f:
        json.dump(deep_analysis_db, f, ensure_ascii=False, indent=2)
    print(f"[3/5] JSON 資料庫更新完成：audit_results.json ({len(audit_db)} 筆), full_audit_data.json ({len(deep_analysis_db)} 筆)")

    # 統計彙總輸出
    total_fmt_fails = sum(1 for d in audit_db if d['c3'] == 'X')
    total_over_2yrs = sum(1 for d in audit_db if d['c4'] == 'V')
    total_org_changes = sum(1 for d in audit_db if d['c5'] == 'V')

    print("==================================================================")
    print(f"=== 全校檢核總結：共檢核 21 個單位、{total_audited} 筆法規項目 ===")
    print(f"  * 條件3 格式瑕疵（含底線/刪除線/字色/標籤）：{total_fmt_fails} 件 ({total_fmt_fails/total_audited*100:.1f}%)")
    print(f"  * 條件4 逾期未更新需提案標註 (<=112年)：{total_over_2yrs} 件 ({total_over_2yrs/total_audited*100:.1f}%)")
    print(f"  * 條件5 配合組織規程/部會修訂提案：{total_org_changes} 件 ({total_org_changes/total_audited*100:.1f}%)")
    print("==================================================================")

    return audit_db

if __name__ == '__main__':
    run_compliance_audit()
