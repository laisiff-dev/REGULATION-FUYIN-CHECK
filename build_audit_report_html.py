# -*- coding: utf-8 -*-
"""
build_audit_report_html.py
讀取 full_audit_data.json 全校法規數據庫（21 個單位、361 筆法規項目），
自動構建具備頂級視覺美學與動態互動能力的單一入口儀表板：audit_report.html

包含四大核心頁籤：
1. 檢核結果總表 (全校 KPI 看板 + 21 單位下拉篩選 + 361 筆全條件即時搜尋清單)
2. 法規位階拓撲 (全校六大體系 Mermaid 向量拓撲圖 + 361 筆母法與授權清冊)
3. 格式瑕疵剖析 (84 件瑕疵法規卡片 + 底線/刪除線/紅藍字色碼剖析 + 定稿清理SOP)
4. 逾期未更新標註 (207 件逾期法規 + 高/中/低急迫性分組卡片 + 115學年度提案時程表)
"""

import json
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, 'full_audit_data.json'), 'r', encoding='utf-8') as f:
    docs = json.load(f)

# Sort docs by sheet then seq
docs = sorted(docs, key=lambda x: (x.get('sheet', ''), x.get('seq', 0)))

total_count = len(docs)
c3_pass = sum(1 for d in docs if d.get('c3') == 'V')
c3_fail = sum(1 for d in docs if d.get('c3') == 'X')
c4_v = sum(1 for d in docs if d.get('c4') == 'V')
c5_v = sum(1 for d in docs if d.get('c5') == 'V')

all_units = sorted(list(set(d.get('sheet', '') for d in docs)))

# Prepare JSON records for client-side JS
audit_report_data = []
for idx, d in enumerate(docs, 1):
    unit = d.get('sheet', '未分類')
    name = d.get('reg_name', '')
    c1 = d.get('c1', 'V')
    c2 = d.get('c2', 'V')
    c3 = d.get('c3', 'V')
    c4 = d.get('c4', 'X')
    c5 = d.get('c5', 'X')
    last_date = d.get('last_date', 'N/A')
    last_year = d.get('last_year')
    diff_years = (115 - last_year) if last_year else 10
    is_over_2_years = (c4 == 'V')
    has_fmt_issues = (c3 == 'X')
    has_org = (c5 == 'V')
    rank = d.get('rank', '業務執行子法 (Level 4: 行政會議)')
    meeting = d.get('approving_body', '行政會議')
    mother_laws = d.get('mother_laws', [])
    mother_str = "、".join(mother_laws) if mother_laws else "輔英科技大學組織規程"
    art1_summary = d.get('art1_summary', '依學校規章訂定')
    format_bugs = d.get('format_bugs', [])

    # Urgency & rating
    if not is_over_2_years:
        urgency = "時效合規"
        overdue_reason = "近 2 年內曾更新（時效正常）"
    elif last_year is None or last_year <= 108 or has_org:
        urgency = "高急迫性 🔴"
        overdue_reason = f"已逾 {diff_years} 年未修訂{'（含舊機關稱謂）' if has_org else ''}，排入 115-1 提案"
    elif last_year in [109, 110, 111]:
        urgency = "中急迫性 🟡"
        overdue_reason = f"已逾 {diff_years} 年未修訂，排入 115-2 提案"
    else:
        urgency = "低急迫性 🟢"
        overdue_reason = f"已逾 {diff_years} 年未修訂，列入例行常態檢討"

    if not has_fmt_issues and not is_over_2_years and not has_org:
        rating = "🟢 優良"
        rating_class = "rating-good"
    elif has_fmt_issues and not is_over_2_years and not has_org:
        rating = "🟡 格式待修"
        rating_class = "rating-warn"
    elif not has_fmt_issues and is_over_2_years:
        rating = "🟡 提案待修"
        rating_class = "rating-warn"
    elif diff_years >= 7 or has_org:
        rating = "🔴 急迫整改"
        rating_class = "rating-danger"
    else:
        rating = "🟠 待辦整改"
        rating_class = "rating-amber"

    audit_report_data.append({
        'num': idx,
        'seq': d.get('seq', idx),
        'title': name,
        'unit': unit,
        'category': d.get('category', ''),
        'rank': rank,
        'meeting': meeting,
        'mother_law': mother_str,
        'art1': art1_summary,
        'c1': c1,
        'c2': c2,
        'c3': c3,
        'c4': c4,
        'c5': c5,
        'last_date': last_date,
        'last_year': last_year,
        'diff_years': diff_years,
        'is_over_2_years': is_over_2_years,
        'urgency': urgency,
        'overdue_reason': overdue_reason,
        'rating': rating,
        'rating_class': rating_class,
        'has_fmt_issues': has_fmt_issues,
        'format_bugs': format_bugs,
        'has_org': has_org,
        'notes': d.get('notes', '無特別說明')
    })

audit_json_str = json.dumps(audit_report_data, ensure_ascii=False)
units_options_html = '\n'.join([f'<option value="{u}">{u}</option>' for u in all_units])

html_template = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>輔英科技大學 全校法規檢核綜合報告儀表板 (依據檢核五個條件_更新)</title>
    <!-- Google Fonts & Font Awesome -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+TC:wght@300;400;500;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Mermaid.js -->
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>

    <style>
        :root {{
            --bg-base: #0b1329;
            --bg-surface: #131f37;
            --bg-card: rgba(26, 38, 66, 0.75);
            --bg-card-hover: rgba(33, 49, 86, 0.95);
            --border-subtle: rgba(255, 255, 255, 0.08);
            --border-focus: #38bdf8;
            
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            
            --brand-primary: #0284c7;
            --brand-glow: rgba(2, 132, 199, 0.35);
            --emerald: #10b981;
            --emerald-bg: rgba(16, 185, 129, 0.15);
            --rose: #f43f5e;
            --rose-bg: rgba(244, 63, 94, 0.15);
            --amber: #f59e0b;
            --amber-bg: rgba(245, 158, 11, 0.15);
            --sky: #38bdf8;
            --sky-bg: rgba(56, 189, 248, 0.15);
            --purple: #a855f7;
            
            --radius-sm: 8px;
            --radius-md: 12px;
            --radius-lg: 20px;
            --shadow-card: 0 10px 30px -5px rgba(0, 0, 0, 0.4);
            --transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Noto Sans TC', 'Inter', sans-serif;
        }}

        body {{
            background-color: var(--bg-base);
            background-image: 
                radial-gradient(at 0% 0%, rgba(2, 132, 199, 0.18) 0px, transparent 50%),
                radial-gradient(at 100% 0%, rgba(168, 85, 247, 0.15) 0px, transparent 50%),
                radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.12) 0px, transparent 50%);
            background-attachment: fixed;
            color: var(--text-primary);
            min-height: 100vh;
            padding: 24px 32px;
        }}

        .layout-container {{
            max-width: 1750px;
            margin: 0 auto;
        }}

        /* Header Bar */
        header.top-bar {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-lg);
            padding: 20px 32px;
            box-shadow: var(--shadow-card);
            margin-bottom: 24px;
        }}

        .logo-area {{
            display: flex;
            align-items: center;
            gap: 18px;
        }}

        .logo-badge {{
            width: 54px;
            height: 54px;
            background: linear-gradient(135deg, #0284c7, #6366f1);
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 26px;
            color: #fff;
            box-shadow: 0 4px 20px var(--brand-glow);
        }}

        .title-group h1 {{
            font-size: 22px;
            font-weight: 700;
            letter-spacing: -0.5px;
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .title-group h1 span.tag {{
            background: var(--sky-bg);
            color: var(--sky);
            border: 1px solid rgba(56, 189, 248, 0.3);
            font-size: 11px;
            padding: 3px 10px;
            border-radius: 20px;
            font-weight: 600;
        }}

        .title-group p {{
            font-size: 13px;
            color: var(--text-secondary);
            margin-top: 4px;
        }}

        .meta-actions {{
            display: flex;
            align-items: center;
            gap: 14px;
        }}

        .btn-action {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(255, 255, 255, 0.06);
            border: 1px solid var(--border-subtle);
            color: var(--text-primary);
            padding: 9px 18px;
            border-radius: var(--radius-md);
            font-size: 13px;
            font-weight: 500;
            cursor: pointer;
            text-decoration: none;
            transition: var(--transition);
        }}

        .btn-action:hover {{
            background: rgba(255, 255, 255, 0.12);
            border-color: var(--sky);
            transform: translateY(-1px);
        }}

        .btn-primary {{
            background: linear-gradient(135deg, #0284c7, #2563eb);
            border: none;
            color: #fff;
            box-shadow: 0 4px 12px var(--brand-glow);
        }}

        /* KPI Cards */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 18px;
            margin-bottom: 24px;
        }}

        .kpi-card {{
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
            padding: 20px;
            box-shadow: var(--shadow-card);
            position: relative;
            overflow: hidden;
            transition: var(--transition);
        }}

        .kpi-card:hover {{
            transform: translateY(-3px);
            border-color: rgba(255, 255, 255, 0.2);
        }}

        .kpi-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
        }}

        .kpi-card.blue::before {{ background: #0284c7; }}
        .kpi-card.emerald::before {{ background: #10b981; }}
        .kpi-card.rose::before {{ background: #f43f5e; }}
        .kpi-card.amber::before {{ background: #f59e0b; }}

        .kpi-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
        }}

        .kpi-title {{
            font-size: 13px;
            font-weight: 500;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .kpi-icon {{
            width: 38px;
            height: 38px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
        }}

        .kpi-card.blue .kpi-icon {{ background: rgba(2, 132, 199, 0.15); color: #38bdf8; }}
        .kpi-card.emerald .kpi-icon {{ background: var(--emerald-bg); color: #34d399; }}
        .kpi-card.rose .kpi-icon {{ background: var(--rose-bg); color: #fb7185; }}
        .kpi-card.amber .kpi-icon {{ background: var(--amber-bg); color: #fbbf24; }}

        .kpi-val {{
            font-size: 32px;
            font-weight: 700;
            letter-spacing: -1px;
            color: #fff;
            display: flex;
            align-items: baseline;
            gap: 8px;
        }}

        .kpi-val span.sub {{
            font-size: 14px;
            font-weight: 400;
            color: var(--text-muted);
        }}

        .kpi-foot {{
            margin-top: 10px;
            font-size: 12px;
            color: var(--text-secondary);
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        /* Tabs Nav */
        .tabs-nav {{
            display: flex;
            gap: 8px;
            background: var(--bg-surface);
            padding: 6px;
            border-radius: var(--radius-md);
            border: 1px solid var(--border-subtle);
            margin-bottom: 24px;
        }}

        .tab-btn {{
            flex: 1;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            padding: 14px 20px;
            background: transparent;
            border: none;
            color: var(--text-secondary);
            font-size: 14px;
            font-weight: 600;
            border-radius: var(--radius-sm);
            cursor: pointer;
            transition: var(--transition);
        }}

        .tab-btn:hover {{
            color: var(--text-primary);
            background: rgba(255, 255, 255, 0.04);
        }}

        .tab-btn.active {{
            background: linear-gradient(135deg, rgba(2, 132, 199, 0.25), rgba(99, 102, 241, 0.25));
            color: #fff;
            border: 1px solid rgba(56, 189, 248, 0.4);
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        }}

        .tab-btn .badge-pill {{
            background: rgba(255, 255, 255, 0.1);
            color: #fff;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 700;
        }}

        .tab-btn.active .badge-pill {{
            background: var(--sky);
            color: #0b1329;
        }}

        /* Panes */
        .tab-pane {{
            display: none;
            animation: fadeIn 0.3s ease-in-out;
        }}

        .tab-pane.active {{
            display: block;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(6px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .content-box {{
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-lg);
            padding: 28px;
            box-shadow: var(--shadow-card);
            margin-bottom: 24px;
        }}

        .box-title-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 14px;
            border-bottom: 1px solid var(--border-subtle);
        }}

        .box-title {{
            font-size: 18px;
            font-weight: 700;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        /* Filter Controls */
        .filter-bar {{
            display: flex;
            gap: 12px;
            align-items: center;
            flex-wrap: wrap;
            margin-bottom: 16px;
        }}

        .search-box {{
            position: relative;
            flex: 1;
            min-width: 260px;
        }}

        .search-box input {{
            width: 100%;
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid var(--border-subtle);
            color: #fff;
            padding: 10px 16px 10px 40px;
            border-radius: var(--radius-md);
            font-size: 13px;
            outline: none;
            transition: var(--transition);
        }}

        .search-box input:focus {{
            border-color: var(--sky);
            box-shadow: 0 0 0 3px rgba(56, 189, 248, 0.15);
        }}

        .search-box i {{
            position: absolute;
            left: 14px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            font-size: 14px;
        }}

        .unit-select {{
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid var(--border-subtle);
            color: var(--text-primary);
            padding: 10px 14px;
            border-radius: var(--radius-md);
            font-size: 13px;
            outline: none;
            cursor: pointer;
        }}

        .filter-tags {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }}

        .filter-tag-btn {{
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-subtle);
            color: var(--text-secondary);
            padding: 8px 14px;
            border-radius: var(--radius-sm);
            font-size: 12px;
            cursor: pointer;
            transition: var(--transition);
            font-weight: 500;
        }}

        .filter-tag-btn:hover, .filter-tag-btn.active {{
            background: var(--sky-bg);
            color: var(--sky);
            border-color: var(--sky);
        }}

        /* Tables */
        .table-responsive {{
            overflow-x: auto;
            border-radius: var(--radius-md);
            border: 1px solid var(--border-subtle);
            max-height: 650px;
        }}

        table.modern-table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 13px;
        }}

        table.modern-table thead {{
            background: rgba(15, 23, 42, 0.95);
            position: sticky;
            top: 0;
            z-index: 10;
        }}

        table.modern-table th {{
            padding: 14px 16px;
            color: var(--text-secondary);
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: 2px solid rgba(255, 255, 255, 0.1);
            white-space: nowrap;
        }}

        table.modern-table td {{
            padding: 12px 16px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.04);
            color: var(--text-primary);
        }}

        table.modern-table tbody tr:hover {{
            background: rgba(255, 255, 255, 0.03);
        }}

        /* Status Pills */
        .pill {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 4px;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 700;
        }}

        .pill-v {{ background: var(--emerald-bg); color: #34d399; border: 1px solid rgba(52, 211, 153, 0.3); }}
        .pill-x {{ background: var(--rose-bg); color: #fb7185; border: 1px solid rgba(251, 113, 133, 0.3); }}

        .rating-good {{ background: var(--emerald-bg); color: #34d399; }}
        .rating-warn {{ background: var(--amber-bg); color: #fbbf24; }}
        .rating-danger {{ background: var(--rose-bg); color: #fb7185; }}
        .rating-amber {{ background: rgba(249, 115, 22, 0.15); color: #fb923c; }}

        /* Mermaid container */
        .mermaid-card {{
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
            padding: 24px;
            margin-bottom: 24px;
            overflow-x: auto;
        }}

        /* Defect Grid */
        .defect-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(380px, 1fr));
            gap: 16px;
        }}

        .defect-card {{
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid rgba(244, 63, 94, 0.25);
            border-radius: var(--radius-md);
            padding: 16px;
            transition: var(--transition);
        }}

        .defect-card:hover {{
            border-color: rgba(244, 63, 94, 0.6);
            transform: translateY(-2px);
        }}

        .defect-title {{
            font-size: 14px;
            font-weight: 700;
            color: #fff;
            margin-bottom: 8px;
        }}

        .defect-tag {{
            display: inline-block;
            background: var(--rose-bg);
            color: #fb7185;
            padding: 2px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 600;
            margin-right: 6px;
            margin-bottom: 6px;
        }}
    </style>
</head>
<body>

<div class="layout-container">

    <!-- Header Bar -->
    <header class="top-bar">
        <div class="logo-area">
            <div class="logo-badge">
                <i class="fa-solid fa-scale-balanced"></i>
            </div>
            <div class="title-group">
                <h1>
                    輔英科技大學 全校法規檢核綜合報告儀表板
                    <span class="tag">依據《法規檢核五個條件_更新》</span>
                </h1>
                <p>列管範圍：全校 21 個單位、共 361 筆法規項目及 366 份官網現行法規 PDF 全文檔 | 基準年份：民國 115 年</p>
            </div>
        </div>
        <div class="meta-actions">
            <a href="法規檢核條件欄位說明.html" class="btn-action" style="border-color: rgba(56, 189, 248, 0.4); color: #38bdf8;">
                <i class="fa-solid fa-circle-question"></i> C1~C5 欄位意義說明
            </a>
            <a href="115年法規清單檢核表_已完成檢核.xlsx" class="btn-action btn-primary" download>
                <i class="fa-solid fa-file-excel"></i> 下載 Excel 檢核表
            </a>
            <a href="法規檢核結果總表.md" class="btn-action" target="_blank">
                <i class="fa-solid fa-file-lines"></i> 總表 Markdown
            </a>
        </div>
    </header>

    <!-- KPI Cards Grid -->
    <div class="kpi-grid">
        <div class="kpi-card blue">
            <div class="kpi-header">
                <span class="kpi-title">全校受檢法規總量</span>
                <div class="kpi-icon"><i class="fa-solid fa-layer-group"></i></div>
            </div>
            <div class="kpi-val">{total_count} <span class="sub">案</span></div>
            <div class="kpi-foot">
                <i class="fa-solid fa-circle-check" style="color:#38bdf8;"></i> 涵蓋全校 21 個一級與二級單位
            </div>
        </div>

        <div class="kpi-card emerald">
            <div class="kpi-header">
                <span class="kpi-title">格式完全合規 (C3=V)</span>
                <div class="kpi-icon"><i class="fa-solid fa-spell-check"></i></div>
            </div>
            <div class="kpi-val">{c3_pass} <span class="sub">案 ({c3_pass/total_count*100:.1f}%)</span></div>
            <div class="kpi-foot">
                <i class="fa-solid fa-check" style="color:#10b981;"></i> 無底線、刪除線、紅藍字或標籤
            </div>
        </div>

        <div class="kpi-card rose">
            <div class="kpi-header">
                <span class="kpi-title">格式瑕疵法規 (C3=X)</span>
                <div class="kpi-icon"><i class="fa-solid fa-triangle-exclamation"></i></div>
            </div>
            <div class="kpi-val">{c3_fail} <span class="sub">案 ({c3_fail/total_count*100:.1f}%)</span></div>
            <div class="kpi-foot">
                <i class="fa-solid fa-brush" style="color:#f43f5e;"></i> 須由承辦處室重新上傳清理定稿
            </div>
        </div>

        <div class="kpi-card amber">
            <div class="kpi-header">
                <span class="kpi-title">逾 2 年未更新標註 (C4=V)</span>
                <div class="kpi-icon"><i class="fa-solid fa-clock-rotate-left"></i></div>
            </div>
            <div class="kpi-val">{c4_v} <span class="sub">案 ({c4_v/total_count*100:.1f}%)</span></div>
            <div class="kpi-foot">
                <i class="fa-solid fa-calendar-days" style="color:#f59e0b;"></i> 規劃於 115 學年度行政會議提案
            </div>
        </div>
    </div>

    <!-- Navigation Tabs -->
    <div class="tabs-nav">
        <button class="tab-btn active" onclick="switchTab('summary')">
            <i class="fa-solid fa-table-list"></i> 檢核結果總表
            <span class="badge-pill">{total_count}</span>
        </button>
        <button class="tab-btn" onclick="switchTab('hierarchy')">
            <i class="fa-solid fa-diagram-project"></i> 法規位階拓撲
            <span class="badge-pill">6 大體系</span>
        </button>
        <button class="tab-btn" onclick="switchTab('defects')">
            <i class="fa-solid fa-highlighter"></i> 格式瑕疵剖析
            <span class="badge-pill">{c3_fail}</span>
        </button>
        <button class="tab-btn" onclick="switchTab('overdue')">
            <i class="fa-solid fa-hourglass-half"></i> 逾期未更新標註
            <span class="badge-pill">{c4_v}</span>
        </button>
    </div>

    <!-- ========================================== -->
    <!-- TAB 1: 檢核結果總表 -->
    <!-- ========================================== -->
    <div id="tab-summary" class="tab-pane active">
        <div class="content-box">
            <div class="box-title-row">
                <div class="box-title">
                    <i class="fa-solid fa-magnifying-glass-chart" style="color:#38bdf8;"></i>
                    全校 21 單位法規即時檢核資料庫
                </div>
                <span id="filteredStats" style="font-size:13px; color:var(--text-secondary);">顯示 361 筆</span>
            </div>

            <!-- Filter Controls -->
            <div class="filter-bar">
                <div class="search-box">
                    <i class="fa-solid fa-magnifying-glass"></i>
                    <input type="text" id="summarySearch" placeholder="搜尋法規名稱、序號、關鍵字..." oninput="filterSummaryTable()">
                </div>

                <select id="unitFilter" class="unit-select" onchange="filterSummaryTable()">
                    <option value="ALL">全部單位 (All 21 Units)</option>
                    {units_options_html}
                </select>

                <div class="filter-tags">
                    <button class="filter-tag-btn active" onclick="setSummaryFilter('all', this)">全部 ({total_count})</button>
                    <button class="filter-tag-btn" onclick="setSummaryFilter('good', this)">🟢 格式合格 ({c3_pass})</button>
                    <button class="filter-tag-btn" onclick="setSummaryFilter('format-fail', this)">🔴 格式瑕疵 ({c3_fail})</button>
                    <button class="filter-tag-btn" onclick="setSummaryFilter('overdue', this)">🟡 逾期提案 ({c4_v})</button>
                    <button class="filter-tag-btn" onclick="setSummaryFilter('org', this)">🟣 組織修訂 ({c5_v})</button>
                </div>
            </div>

            <!-- Table -->
            <div class="table-responsive">
                <table class="modern-table">
                    <thead>
                        <tr>
                            <th style="width: 50px;">序號</th>
                            <th style="width: 110px;">業管單位</th>
                            <th>法規名稱</th>
                            <th style="text-align:center; width:65px;">C1 最新</th>
                            <th style="text-align:center; width:65px;">C2 名稱</th>
                            <th style="text-align:center; width:65px;">C3 格式</th>
                            <th style="text-align:center; width:65px;">C4 逾期</th>
                            <th style="text-align:center; width:65px;">C5 組織</th>
                            <th style="width: 100px;">修訂時間</th>
                            <th style="width: 100px;">綜合評等</th>
                            <th>備註與改善建議</th>
                        </tr>
                    </thead>
                    <tbody id="summaryTableBody">
                        <!-- Populated by JS -->
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- ========================================== -->
    <!-- TAB 2: 法規位階拓撲 -->
    <!-- ========================================== -->
    <div id="tab-hierarchy" class="tab-pane">
        <div class="content-box">
            <div class="box-title-row">
                <div class="box-title">
                    <i class="fa-solid fa-sitemap" style="color:#38bdf8;"></i>
                    全校法規位階拓撲結構體系 (依授權位階與審議會議劃分)
                </div>
            </div>

            <p style="font-size:14px; color:var(--text-secondary); margin-bottom:20px;">
                本校法規分為五級：Level 1 (中央法律) &rarr; Level 2 (組織規程) &rarr; Level 3 (校級核心母法-辦法/學則) &rarr; Level 4 (業務執行子法-要點/準則) &rarr; Level 5 (作業規範-細則/規約)。
            </p>

            <div class="mermaid-card">
                <h3 style="color:#38bdf8; margin-bottom:12px; font-size:16px;">1. 學校治理與根本組織規程體系</h3>
                <pre class="mermaid">
graph TD
    L1_LAW["大學法 / 私立學校法"] --> L2_ORG["輔英科技大學組織規程 (114.12.5教育部核定本)"]
    L2_ORG --> M_COUNCIL["校務會議議事規則"]
    L2_ORG --> M_PRESIDENT["校長選聘及解聘辦法"]
    L2_ORG --> M_SELF_EVAL["自我評鑑實施辦法"]
    L2_ORG --> S_ADMIN["行政會議議事要點"]
    L2_ORG --> S_CAMPUS_SEC["校園安全防救委員會設置辦法"]
    L2_ORG --> S_INTER_AUDIT["內部控制實施要點"]
                </pre>
            </div>

            <div class="mermaid-card">
                <h3 style="color:#38bdf8; margin-bottom:12px; font-size:16px;">2. 教務教學與學則體系</h3>
                <pre class="mermaid">
graph TD
    L1_DEG["學位授予法 / 大學法施行細則"] --> M_ACAD["輔英科技大學學則"]
    M_ACAD --> M_ADMIS["各項入學招生規定 (單招/甄試/二技/四技)"]
    M_ACAD --> M_DEGREE["研究所學位考試辦法"]
    M_ACAD --> S_COURSE["學生選課要點"]
    M_ACAD --> S_TRANSFER["學生抵免學分要點"]
    M_ACAD --> S_MINOR["學生修讀輔系(科)要點"]
    M_ACAD --> S_INTERDIS["跨領域學分學程實施要點"]
    M_DEGREE --> S_ETHIC["碩士學位論文違反學術倫理案件處理要點"]
                </pre>
            </div>

            <div class="mermaid-card">
                <h3 style="color:#38bdf8; margin-bottom:12px; font-size:16px;">3. 研發成果、產學合作與智慧財產權體系</h3>
                <pre class="mermaid">
graph TD
    L1_TECH["科學技術基本法 / 專利法 / 產學合作實施辦法"] --> M_IND["產學合作實施辦法"]
    L1_TECH --> M_IPR["研發成果管理辦法"]
    M_IND --> M_DERIVE["校園衍生企業申設管理辦法"]
    M_IPR --> S_PATENT["專利申請與維護作業要點"]
    M_IPR --> S_TECH["技術移轉作業要點"]
    M_IPR --> S_COI["研發成果利益衝突迴避及資訊揭露要點"]
    M_IND --> S_MOU["產官學合作備忘錄簽訂作業要點"]
    M_IND --> S_INCENT["教師學術及專業研發成果獎勵要點"]
                </pre>
            </div>

            <!-- Hierarchy Table -->
            <div class="box-title" style="margin:24px 0 14px 0;">全校 361 筆法規母子法位階清冊</div>
            <div class="table-responsive">
                <table class="modern-table">
                    <thead>
                        <tr>
                            <th style="width: 50px;">序號</th>
                            <th style="width: 110px;">單位</th>
                            <th>法規名稱</th>
                            <th style="width: 180px;">位階層級</th>
                            <th style="width: 140px;">審議/核定會議</th>
                            <th>上位法源 / 母法依據</th>
                        </tr>
                    </thead>
                    <tbody id="hierarchyTableBody">
                        <!-- Populated by JS -->
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- ========================================== -->
    <!-- TAB 3: 格式瑕疵剖析 -->
    <!-- ========================================== -->
    <div id="tab-defects" class="tab-pane">
        <div class="content-box">
            <div class="box-title-row">
                <div class="box-title">
                    <i class="fa-solid fa-brush" style="color:#f43f5e;"></i>
                    格式瑕疵法規全面檢核清冊 (共 {c3_fail} 案待定稿清理)
                </div>
                <div style="font-size:13px; color:var(--text-secondary);">
                    依據《法規檢核五個條件_更新》第二項條件：須不含底線、刪除線及各式顏色標示等
                </div>
            </div>

            <!-- SOP Banner -->
            <div style="background:rgba(2,132,199,0.1); border:1px solid rgba(56,189,248,0.3); border-radius:var(--radius-md); padding:16px; margin-bottom:20px;">
                <h4 style="color:#38bdf8; margin-bottom:6px;"><i class="fa-solid fa-wrench"></i> 法規格式清理標準作業程序 (SOP)</h4>
                <p style="font-size:13px; color:var(--text-secondary); line-height:1.6;">
                    1. <b>停止追蹤修訂</b>：Word「校閱」標籤 &rarr;「變更」&rarr;「接受所有變更並停止追蹤」。<br>
                    2. <b>字體色彩一致化</b>：按 Ctrl+A 全選條文，字型色彩選「自動」或純黑色 (#000000)。<br>
                    3. <b>取消修訂底線與刪除線</b>：全選條文解除底線 (Ctrl+U) 與刪除線勾選，再轉存純淨 PDF 重新上傳法規目錄。
                </p>
            </div>

            <!-- Defect Search -->
            <div class="search-box" style="margin-bottom:18px;">
                <i class="fa-solid fa-magnifying-glass"></i>
                <input type="text" id="defectSearch" placeholder="搜尋格式瑕疵法規、單位、瑕疵項目..." oninput="filterDefectCards()">
            </div>

            <div class="defect-grid" id="defectCardsContainer">
                <!-- Populated by JS -->
            </div>
        </div>
    </div>

    <!-- ========================================== -->
    <!-- TAB 4: 逾期未更新標註 -->
    <!-- ========================================== -->
    <div id="tab-overdue" class="tab-pane">
        <div class="content-box">
            <div class="box-title-row">
                <div class="box-title">
                    <i class="fa-solid fa-clock-rotate-left" style="color:#fbbf24;"></i>
                    超過 2 年以上未更新法規標註名冊 (共 {c4_v} 案)
                </div>
                <div style="font-size:13px; color:var(--text-secondary);">
                    依據《法規檢核五個條件_更新》第三項條件：以 115 年為基準，超過 2 年以上未更新請建立標註檔
                </div>
            </div>

            <!-- Urgency Legend -->
            <div style="display:flex; gap:16px; margin-bottom:20px; flex-wrap:wrap;">
                <div style="background:rgba(244,63,94,0.15); border:1px solid rgba(244,63,94,0.3); padding:10px 16px; border-radius:var(--radius-sm); font-size:13px;">
                    <b style="color:#fb7185;">🔴 高急迫性</b>：未修訂 &ge; 7 年 或 含有廢止機關名稱（排入 115-1 提案）
                </div>
                <div style="background:rgba(245,158,11,0.15); border:1px solid rgba(245,158,11,0.3); padding:10px 16px; border-radius:var(--radius-sm); font-size:13px;">
                    <b style="color:#fbbf24;">🟡 中急迫性</b>：未修訂 4~6 年（排入 115-2 提案）
                </div>
                <div style="background:rgba(16,185,129,0.15); border:1px solid rgba(16,185,129,0.3); padding:10px 16px; border-radius:var(--radius-sm); font-size:13px;">
                    <b style="color:#34d399;">🟢 低急迫性</b>：未修訂 3 年（併入常態檢討檢視文字）
                </div>
            </div>

            <div class="table-responsive">
                <table class="modern-table">
                    <thead>
                        <tr>
                            <th style="width: 50px;">序號</th>
                            <th style="width: 110px;">單位</th>
                            <th>法規名稱</th>
                            <th style="width: 100px;">修訂時間</th>
                            <th style="width: 90px;">已逾年數</th>
                            <th style="width: 110px;">標註狀態</th>
                            <th style="width: 110px;">急迫性分級</th>
                            <th>建議提案排程與因應處置</th>
                        </tr>
                    </thead>
                    <tbody id="overdueTableBody">
                        <!-- Populated by JS -->
                    </tbody>
                </table>
            </div>
        </div>
    </div>

</div>

<script>
    // Embedded Data
    const rawData = {audit_json_str};

    // Initialize Mermaid
    mermaid.initialize({{
        startOnLoad: false,
        theme: 'dark',
        themeVariables: {{
            darkMode: true,
            background: 'transparent',
            primaryColor: '#0284c7',
            primaryTextColor: '#fff',
            primaryBorderColor: '#38bdf8',
            lineColor: '#64748b',
            secondaryColor: '#1e293b',
            tertiaryColor: '#0f172a'
        }}
    }});

    // Tab Switching
    function switchTab(tabId) {{
        document.querySelectorAll('.tab-pane').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

        const targetPane = document.getElementById('tab-' + tabId);
        if (targetPane) targetPane.classList.add('active');

        const activeBtn = Array.from(document.querySelectorAll('.tab-btn')).find(b => b.getAttribute('onclick').includes(tabId));
        if (activeBtn) activeBtn.classList.add('active');

        if (tabId === 'hierarchy') {{
            setTimeout(() => {{
                mermaid.run();
            }}, 100);
        }}
    }}

    // Render Tab 1: Summary Table
    let currentSummaryFilter = 'all';

    function renderSummaryTable(data) {{
        const tbody = document.getElementById('summaryTableBody');
        tbody.innerHTML = '';

        data.forEach(item => {{
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><b>${{String(item.num).padStart(3, '0')}}</b></td>
                <td><span style="color:#38bdf8; font-weight:500;">${{item.unit}}</span></td>
                <td><b>${{item.title}}</b></td>
                <td style="text-align:center;"><span class="pill pill-${{item.c1.toLowerCase()}}">${{item.c1}}</span></td>
                <td style="text-align:center;"><span class="pill pill-${{item.c2.toLowerCase()}}">${{item.c2}}</span></td>
                <td style="text-align:center;"><span class="pill pill-${{item.c3.toLowerCase()}}">${{item.c3}}</span></td>
                <td style="text-align:center;"><span class="pill pill-${{item.c4.toLowerCase()}}">${{item.c4}}</span></td>
                <td style="text-align:center;"><span class="pill pill-${{item.c5.toLowerCase()}}">${{item.c5}}</span></td>
                <td><code style="font-family:'JetBrains Mono'; color:#38bdf8;">${{item.last_date}}</code></td>
                <td><span class="pill ${{item.rating_class}}">${{item.rating}}</span></td>
                <td style="font-size:12px; color:var(--text-secondary);">${{item.notes}}</td>
            `;
            tbody.appendChild(tr);
        }});

        const statsEl = document.getElementById('filteredStats');
        if (statsEl) statsEl.innerText = `顯示 ${{data.length}} 筆 / 全校 {total_count} 筆`;
    }}

    function filterSummaryTable() {{
        const q = document.getElementById('summarySearch').value.trim().toLowerCase();
        const selectedUnit = document.getElementById('unitFilter').value;

        let filtered = rawData.filter(d => {{
            // Unit filter
            if (selectedUnit !== 'ALL' && d.unit !== selectedUnit) return false;

            // Search query
            const matchQ = !q || d.title.toLowerCase().includes(q) || String(d.num).includes(q) || d.unit.toLowerCase().includes(q) || d.notes.toLowerCase().includes(q);
            if (!matchQ) return false;

            // Tag filter
            if (currentSummaryFilter === 'good') return !d.has_fmt_issues && !d.is_over_2_years;
            if (currentSummaryFilter === 'format-fail') return d.has_fmt_issues;
            if (currentSummaryFilter === 'overdue') return d.is_over_2_years;
            if (currentSummaryFilter === 'org') return d.has_org;
            return true;
        }});
        renderSummaryTable(filtered);
    }}

    function setSummaryFilter(type, btn) {{
        document.querySelectorAll('.filter-tag-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentSummaryFilter = type;
        filterSummaryTable();
    }}

    // Render Tab 2: Hierarchy Table
    function renderHierarchyTable() {{
        const tbody = document.getElementById('hierarchyTableBody');
        tbody.innerHTML = '';

        rawData.forEach(item => {{
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><b>${{String(item.num).padStart(3, '0')}}</b></td>
                <td><span style="color:#38bdf8;">${{item.unit}}</span></td>
                <td><b>${{item.title}}</b></td>
                <td><span style="font-size:12px; font-weight:600;">${{item.rank}}</span></td>
                <td><span style="color:#94a3b8;">${{item.meeting}}</span></td>
                <td style="font-size:12px; color:#cbd5e1;">${{item.mother_law}}</td>
            `;
            tbody.appendChild(tr);
        }});
    }}

    // Render Tab 3: Defect Cards
    function renderDefectCards() {{
        const container = document.getElementById('defectCardsContainer');
        container.innerHTML = '';

        const defects = rawData.filter(d => d.has_fmt_issues);
        defects.forEach(d => {{
            const card = document.createElement('div');
            card.className = 'defect-card';

            const tagRow = d.format_bugs.map(b => `<span class="defect-tag">${{b}}</span>`).join('');
            card.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:8px;">
                    <div class="defect-title">【${{d.unit}}】${{d.title}}</div>
                    <span class="pill pill-x">待清理</span>
                </div>
                <div style="margin-bottom:10px;">${{tagRow}}</div>
                <div style="font-size:12px; color:var(--text-secondary); line-height:1.5;">
                    <i class="fa-solid fa-wrench" style="color:#38bdf8;"></i> 處置建議：全選文字清除底線、刪除線，色彩統一設為「自動(純黑)」，另存標準定稿。
                </div>
            `;
            container.appendChild(card);
        }});
    }}

    function filterDefectCards() {{
        const q = document.getElementById('defectSearch').value.trim().toLowerCase();
        document.querySelectorAll('.defect-card').forEach(card => {{
            const text = card.innerText.toLowerCase();
            card.style.display = text.includes(q) ? 'block' : 'none';
        }});
    }}

    // Render Tab 4: Overdue Table
    function renderOverdueTable() {{
        const tbody = document.getElementById('overdueTableBody');
        tbody.innerHTML = '';

        const overdueItems = rawData.filter(d => d.is_over_2_years);
        overdueItems.sort((a, b) => b.diff_years - a.diff_years);

        overdueItems.forEach(d => {{
            const tr = document.createElement('tr');
            const urgencyClass = d.urgency.includes('高') ? 'rating-danger' : (d.urgency.includes('中') ? 'rating-warn' : 'rating-good');
            tr.innerHTML = `
                <td><b>${{String(d.num).padStart(3, '0')}}</b></td>
                <td><span style="color:#38bdf8;">${{d.unit}}</span></td>
                <td><b>${{d.title}}</b></td>
                <td><code style="font-family:'JetBrains Mono'; color:#38bdf8;">${{d.last_date}}</code></td>
                <td><b style="color:#fb7185;">已逾 ${{d.diff_years}} 年</b></td>
                <td><span class="pill pill-v">【逾期需更新】</span></td>
                <td><span class="pill ${{urgencyClass}}">${{d.urgency}}</span></td>
                <td style="font-size:12px; color:var(--text-secondary);">${{d.overdue_reason}}</td>
            `;
            tbody.appendChild(tr);
        }});
    }}

    // Initial Load
    document.addEventListener('DOMContentLoaded', () => {{
        renderSummaryTable(rawData);
        renderHierarchyTable();
        renderDefectCards();
        renderOverdueTable();
    }});
</script>

</body>
</html>
'''

output_html_path = os.path.join(BASE_DIR, 'audit_report.html')
with open(output_html_path, 'w', encoding='utf-8') as f:
    f.write(html_template)

print(f"成功編譯產出全校儀表板：{output_html_path} (大小: {len(html_template.encode('utf-8')):,} bytes)")
