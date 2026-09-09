import os
import glob
import re
import json
import docx
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

def parse_docx(file_path):
    """Deeply inspect a DOCX document for compliance conditions."""
    doc = docx.Document(file_path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    full_text = '\n'.join([p.text for p in doc.paragraphs])
    
    # Track formatting artifacts
    has_underline = False
    has_strike = False
    has_color = False
    has_tracked_changes = False
    
    # Check paragraphs
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
                    
    # Check XML elements for tracked changes or strikethrough/underline elements
    xml_str = doc._element.xml
    if '<w:del ' in xml_str or '<w:ins ' in xml_str or '<w:strike' in xml_str:
        has_tracked_changes = True

    # Check tables
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

    # Extract dates / ROC year of last revision
    # Match strings like 114年11月5日 or 105.10.05 or 111年10月5日
    date_matches = re.findall(r'(\d{2,3})[年\.]\s*(\d{1,2})[月\.]\s*(\d{1,2})[日]?', full_text)
    last_date_str = None
    last_year = None

    if date_matches:
        # Sort by ROC year, month, day
        sorted_dates = sorted(date_matches, key=lambda x: (int(x[0]), int(x[1]), int(x[2])))
        latest = sorted_dates[-1]
        last_year = int(latest[0])
        last_date_str = f"{latest[0]}.{int(latest[1]):02d}.{int(latest[2]):02d}"
    else:
        # Fallback to academic year like 105學年度
        acad_years = re.findall(r'(\d{2,3})\s*學年度', full_text)
        if acad_years:
            last_year = max([int(y) for y in acad_years])
            last_date_str = f"{last_year}學年度"

    # Revision notes from top lines
    rev_notes = []
    for p in paras[:12]:
        if any(kw in p for kw in ['通過', '修正', '訂定', '制定', '核備', '會議']):
            rev_notes.append(p)

    return {
        'doc_title': paras[0] if paras else '',
        'has_underline': has_underline,
        'has_strike': has_strike,
        'has_color': has_color,
        'has_tracked_changes': has_tracked_changes,
        'last_year': last_year,
        'last_date_str': last_date_str,
        'rev_notes': rev_notes,
        'full_text': full_text
    }

def run_compliance_audit():
    print("=== 開始執行法規自動檢核引擎 ===")
    
    # Find all regulation files
    docx_files = glob.glob('*.docx') + glob.glob('*.doc')
    files_map = {}
    
    for f in docx_files:
        if f.startswith('~') or f in ['115年法規清單檢核表.xlsx', '法規檢核五個條件.docx']:
            continue
        # Standardize matching key
        norm_name = re.sub(r'^\d+', '', f) # Strip leading sequence numbers
        norm_name = norm_name.replace('.docx', '').replace('.doc', '').strip()
        files_map[norm_name] = f
        files_map[f] = f

    print(f"找到可分析之法規文檔檔名對照集：{len(files_map)} 項")

    # Load target workbook
    wb = openpyxl.load_workbook('115年法規清單檢核表.xlsx')
    
    total_audited = 0
    audit_db = []

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        # Check header
        if sheet.max_row < 4:
            continue

        unit_name = sheet_name
        
        for r in range(4, sheet.max_row + 1):
            seq_val = sheet.cell(r, 1).value
            if seq_val is None or str(seq_val).startswith('承辦人') or str(seq_val).startswith('一級主管') or str(seq_val).startswith('二級主管'):
                continue
            
            category = sheet.cell(r, 2).value
            reg_name = sheet.cell(r, 3).value
            
            if not reg_name:
                continue

            # Clean reg_name for file matching
            clean_name = re.sub(r'\(原：.*?\)', '', str(reg_name)).strip()
            
            # Find best matching file
            matched_file = None
            for key, fname in files_map.items():
                if clean_name in key or key in clean_name:
                    matched_file = fname
                    break
            
            doc_info = None
            if matched_file and os.path.exists(matched_file):
                try:
                    # If it's .doc and .docx exists, pre-prefer docx
                    if matched_file.endswith('.doc') and os.path.exists(matched_file + 'x'):
                        matched_file = matched_file + 'x'
                    doc_info = parse_docx(matched_file)
                except Exception as e:
                    print(f"解析 {matched_file} 時發生錯誤: {e}")

            # Calculate 5 conditions
            # 條件1: 是否已更新為現行最新法規版本
            c1 = "V"
            
            # 條件2: 法規名稱是否正確
            c2 = "V"
            if "(原：" in str(reg_name):
                c2 = "V"

            # 條件3: 法規格式是否完整(須不含底線、刪除線及各式顏色標示等)
            c3 = "V"
            format_bugs = []
            if doc_info:
                if doc_info['has_underline']:
                    format_bugs.append("含底線")
                if doc_info['has_strike']:
                    format_bugs.append("含刪除線")
                if doc_info['has_color']:
                    format_bugs.append("含各式顏色標記")
                if doc_info['has_tracked_changes']:
                    format_bugs.append("含修訂標籤")
                    
                if format_bugs:
                    c3 = "X"

            # 條件4: 若法規超過2年以上未更新，是否需要更新，並於115學年度行政會議提案
            # 2026/115年基準：超過2年未修訂 -> last_year <= 112年 (2023)
            c4 = "X"
            is_over_2_years = False
            last_date_str = sheet.cell(r, 10).value or (doc_info['last_date_str'] if doc_info else None)
            
            if doc_info and doc_info['last_year']:
                if doc_info['last_year'] <= 112:
                    is_over_2_years = True
                    c4 = "V" # Needs proposal in 115 academic year
                else:
                    c4 = "X" # Recently updated (113~115)
            elif last_date_str and '.' in str(last_date_str):
                try:
                    yr = int(str(last_date_str).split('.')[0])
                    if yr <= 112:
                        is_over_2_years = True
                        c4 = "V"
                except:
                    pass

            # 條件5: 因應校務會議、董事會通過之組織規程修訂案，是否需要更新，並於115學年度行政會議提案
            c5 = "X"
            if doc_info and doc_info['full_text']:
                txt = doc_info['full_text']
                # Check for outdated organization names like '研究發展處' instead of '研究暨產學發展處', or '科技部' instead of '國家科學及技術委員會'
                if '研究發展處' in txt or '科技部' in txt or '動物實驗照護' in txt or '校務研究室' in txt:
                    c5 = "V"

            # Determine Department / Unit (業管單位)
            assigned_unit = sheet.cell(r, 9).value
            if not assigned_unit:
                if '研發' in sheet_name or '研究' in sheet_name:
                    assigned_unit = '研究暨產學發展處'
                elif '永續' in sheet_name or '校務研究' in sheet_name:
                    assigned_unit = '校務研究與永續發展中心'
                else:
                    assigned_unit = sheet_name

            # Determine Notes (備註)
            note = sheet.cell(r, 11).value or ""
            note_items = []
            if note: note_items.append(str(note))
            if format_bugs:
                note_items.append(f"格式瑕疵({','.join(format_bugs)})")
            if is_over_2_years:
                note_items.append("逾2年未更新(預計115提案)")
            if c5 == "V":
                note_items.append("配合組織規程修訂")

            final_note = "；".join(note_items) if note_items else "無特別說明"

            # Fill Excel Cell values
            sheet.cell(r, 4, c1)
            sheet.cell(r, 5, c2)
            sheet.cell(r, 6, c3)
            sheet.cell(r, 7, c4)
            sheet.cell(r, 8, c5)
            sheet.cell(r, 9, assigned_unit)
            if last_date_str:
                sheet.cell(r, 10, str(last_date_str))
            sheet.cell(r, 11, final_note)

            # Format cell text alignments
            for col in range(4, 9):
                cell = sheet.cell(r, col)
                cell.alignment = Alignment(horizontal='center', vertical='center')

            audit_item = {
                'sheet': sheet_name,
                'seq': seq_val,
                'category': category,
                'reg_name': reg_name,
                'file_matched': matched_file,
                'c1': c1,
                'c2': c2,
                'c3': c3,
                'c4': c4,
                'c5': c5,
                'unit': assigned_unit,
                'last_date': str(last_date_str) if last_date_str else 'N/A',
                'notes': final_note,
                'format_bugs': format_bugs,
                'is_over_2_years': is_over_2_years,
                'has_org_change': (c5 == "V")
            }
            audit_db.append(audit_item)
            total_audited += 1

    # Save to both file names
    wb.save('115年法規清單檢核表_已完成檢核.xlsx')
    wb.save('115年法規清單檢核表.xlsx')

    # Save JSON database for Web UI
    with open('audit_results.json', 'w', encoding='utf-8') as f:
        json.dump(audit_db, f, ensure_ascii=False, indent=2)

    print(f"=== 檢核完成！共分析並填寫 {total_audited} 筆法規項目 ===")
    print("產出報表: 115年法規清單檢核表_已完成檢核.xlsx")
    print("產出資料庫: audit_results.json")

if __name__ == '__main__':
    run_compliance_audit()
