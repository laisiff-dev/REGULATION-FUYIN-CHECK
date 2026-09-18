# -*- coding: utf-8 -*-
"""
輔英科技大學法規目錄自動抓取工具

來源：https://ese.fy.edu.tw/var/file/57/1057/img/125/law.htm  (Big5 編碼之 Word 匯出 HTML)

流程：
  1. 下載法規目錄頁並由 Big5 轉為 UTF-8
  2. 解析 <table> 中「法案編號 | 法規名稱 | 維護單位」三欄列與其 PDF 連結
  3. 依維護單位對應到「輔英科大各單位法規彙整」既有子目錄
  4. 下載 PDF（同檔名去重、失敗重試、節流）
  5. 以 PyMuPDF 擷取條文全文沿革，取最新一筆民國日期作為「最後一次修訂時間」
  6. 輸出 法規彙整索引.csv / .json

用法：
  python fetch_fy_regulations.py            # 完整抓取
  python fetch_fy_regulations.py --dry-run  # 只解析不下載
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
from collections import OrderedDict
from html import unescape
from pathlib import Path

INDEX_URL = "https://ese.fy.edu.tw/var/file/57/1057/img/125/law.htm"
BASE_URL = "https://ese.fy.edu.tw/var/file/57/1057/img/125/"
SITE_ROOT = "https://ese.fy.edu.tw/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")

# 該站前置 F5 WAF：Python 的 urllib/requests 一律被擋，只有 curl（Schannel）可通過；
# 且短時間大量請求會觸發 IP 層暫時封鎖，回傳「Request Rejected」小頁面。
WAF_REJECT = b"Request Rejected"

ROOT = Path(__file__).resolve().parent
ARCHIVE = ROOT / "輔英科大各單位法規彙整"

# 維護單位（網站用字） -> 彙整目錄相對路徑
# 對應依據：輔英科大各單位法規彙整/README.md 所載組織架構
UNIT_DIRS = {
    "董事會": "01_行政單位/01_董事會",
    "秘書室": "01_行政單位/02_校長室與秘書處",
    "公共事務暨校友服務室": "01_行政單位/02_校長室與秘書處",   # README：公關校友歸此
    "教務處": "01_行政單位/03_教務處",
    "學務處": "01_行政單位/04_學生事務處",
    "總務處": "01_行政單位/05_總務處",
    "研究暨產學發展處": "01_行政單位/06_研發與永續發展處",
    "研發與永續發展處": "01_行政單位/06_研發與永續發展處",
    "研發處": "01_行政單位/06_研發與永續發展處",
    "研究發展處": "01_行政單位/06_研發與永續發展處",
    "校務發展辦公室": "01_行政單位/06_研發與永續發展處",
    "校務研究暨規劃室": "01_行政單位/06_研發與永續發展處",
    "圖書暨資訊處": "01_行政單位/07_圖書暨資訊處",
    "國際暨兩岸事務處": "01_行政單位/08_國際暨兩岸事務處",
    "人事室": "01_行政單位/09_人事室",
    "福委會/人事室": "01_行政單位/09_人事室",
    "會計室": "01_行政單位/10_會計室",
    "稽核室": "01_行政單位/11_稽核室",
    "環境安全衛生中心": "01_行政單位/12_環境安全衛生中心",
    "體育暨健康促進中心": "01_行政單位/13_健康與體育發展處",
    "推廣教育中心": "01_行政單位/14_推廣教育中心",
    "護理學院": "02_教學學術單位/01_護理學院",
    "高齡及長期照護事業系": "02_教學學術單位/01_護理學院",    # README：高照系隸屬護理學院
    "健康美容系": "02_教學學術單位/02_醫學與健康學院",         # README：健康美容系隸屬醫健學院
    "幼兒保育暨產業系": "02_教學學術單位/04_人文與管理學院",    # README：幼保系隸屬人管學院
    "共同教育中心": "02_教學學術單位/05_共同教育中心",
    "老化及疾病預防研究中心": "03_研究與中心機構/01_老化及疾病預防研究中心",
    "校務研究暨規劃室": "03_研究與中心機構/02_校務研究與永續發展中心",
    "附設醫院": "04_專門委員會與附設機構/04_附設醫院與幼兒園",
    "附設幼兒園": "04_專門委員會與附設機構/04_附設醫院與幼兒園",
}

UNCLASSIFIED = "99_未對應單位"

ILLEGAL_FS = r'[\\/:*?"<>|\r\n\t]'


# --------------------------------------------------------------------------- #
# 抓取與解析
# --------------------------------------------------------------------------- #
class WafBlocked(Exception):
    """WAF 回傳 Request Rejected 攔截頁。"""


def _curl_bin():
    for name in ("curl", "curl.exe"):
        path = shutil.which(name)
        if path:
            return path
    raise RuntimeError("找不到 curl，本站只能透過 curl 存取（見檔頭說明）")


def fetch(url, timeout=90, referer=SITE_ROOT):
    """以 curl 取得內容。遇到 WAF 攔截頁時丟出 WafBlocked。"""
    # mkstemp 回傳「已開啟」的 fd，Windows 下若不先關閉，curl 無法寫入該檔（WinError 32）
    fd, tmp_name = tempfile.mkstemp(suffix=".part")
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        cmd = [_curl_bin(), "-sS", "-L", "--max-time", str(timeout),
               "-A", UA, "-H", "Referer: " + referer,
               "-H", "Accept: text/html,application/xhtml+xml,application/pdf,*/*;q=0.8",
               "-o", str(tmp), "-w", "%{http_code}", url]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0:
            raise OSError("curl 失敗 (exit %d): %s" % (proc.returncode, proc.stderr.strip()[:200]))
        code = (proc.stdout or "").strip()
        data = tmp.read_bytes()
        if WAF_REJECT in data[:400]:
            raise WafBlocked("WAF 攔截（HTTP %s, %d bytes）" % (code, len(data)))
        if code != "200":
            raise OSError("HTTP %s" % code)
        return data
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass  # 暫存檔清不掉不該讓整批下載失敗


def strip_tags(fragment):
    return re.sub(r"\s+", "", unescape(re.sub(r"<[^>]+>", "", fragment)))


def load_index_html(cache, refresh):
    if cache.exists() and not refresh:
        return cache.read_text(encoding="utf-8")
    raw = fetch(INDEX_URL)
    text = raw.decode("big5", errors="replace")
    cache.parent.mkdir(parents=True, exist_ok=True)
    cache.write_text(text, encoding="utf-8")
    return text


def section_map(html):
    """回傳 [(位置, 錨點代號, 分類名稱)]，供依文件位置判定法規分類。"""
    cats = []
    for m in re.finditer(r'<a[^>]*name="?([A-Za-z][A-Za-z0-9]{0,3})"?[^>]*>(.*?)</a>', html, re.S | re.I):
        code = m.group(1)
        if code.upper().startswith(("OLE", "GENE", "ORIG", "PROG")):
            continue
        title = strip_tags(m.group(2))
        if title:
            cats.append((m.start(), code, title))
    cats.sort()
    return cats


def parse_rows(html):
    """解析所有含 PDF 連結的表格列。"""
    cats = section_map(html)

    def category_of(pos):
        cur = ("?", "未分類")
        for cpos, code, title in cats:
            if cpos <= pos:
                cur = (code, title)
            else:
                break
        return cur

    rows = []

    def emit(serial, name, unit, chunk, base):
        if name in ("法案編號", "法規名稱", "維護單位", ""):
            return
        links = list(re.finditer(r'<a[^>]*href="([^"]+\.pdf)"[^>]*>(.*?)</a>', chunk, re.S | re.I))
        if not links:  # 極少數連結未包在 <a> 標籤內
            links = list(re.finditer(r'()href="([^"]+\.pdf)"', chunk, re.I))
            links = [(m, m.group(2), "") for m in links]
        else:
            links = [(m, m.group(1), strip_tags(m.group(2))) for m in links]

        # 同一個 href 在一格內出現多次，代表 Word 把同一個標題切成數段超連結，
        # 此時整格文字才是完整名稱，只取一筆。
        seen_href = set()
        links = [l for l in links if not (l[1] in seen_href or seen_href.add(l[1]))]

        # 一列可能掛兩個不同連結（本文 + 附件、或同要點的各學年度版本），
        # 此時儲存格文字是數段標題相接，改以各自的錨點文字命名；
        # 看起來不像完整名稱者（例如「(…附件)」）則接在主名稱後。
        primary = links[0][2] if links else name
        for m, href, atext in links:
            label = name
            if len(links) > 1 and atext:
                # 以「(」開頭者是附件或版本標註，屬於主名稱的一部分；其餘自成完整名稱
                label = primary + atext if atext[0] in "(（[［" else atext
            pos = base + m.start()
            code, cat = category_of(pos)
            rows.append({
                "serial": serial,
                "name": label,
                "unit": unit,
                "href": unescape(href),
                "anchor": code,
                "category": cat,
                "pos": pos,
            })

    for tm in re.finditer(r"<table.*?</table>", html, re.S | re.I):
        table = tm.group(0)
        covered = []  # 已被 <tr> 涵蓋的區間，供後續補抓孤兒 <td> 使用
        for rm in re.finditer(r"<tr.*?</tr>", table, re.S | re.I):
            tr = rm.group(0)
            covered.append((rm.start(), rm.end()))
            cells = [(c.start(), c.group(0)) for c in re.finditer(r"<t[dh].*?</t[dh]>", tr, re.S | re.I)]
            # 正常為 3 欄；原始頁面有少數列被 Word 合併成 6、9 欄（多筆法規擠在同一 <tr>），
            # 只要是 3 的倍數就依序切成每 3 欄一筆。
            if not cells or len(cells) % 3 != 0:
                continue
            for g in range(0, len(cells), 3):
                group = cells[g:g + 3]
                serial, name, unit = (strip_tags(c[1]) for c in group)
                emit(serial, name, unit, "".join(c[1] for c in group),
                     tm.start() + rm.start() + group[0][0])

        # 補抓：Word 匯出的 HTML 有少數 <td> 掉在 <tr> 之外（缺 <tr> 標籤），
        # 於未被任何 <tr> 涵蓋的區段再依每 3 欄一筆掃一次。
        orphans = [(c.start(), c.group(0))
                   for c in re.finditer(r"<t[dh].*?</t[dh]>", table, re.S | re.I)
                   if not any(s <= c.start() < e for s, e in covered)]
        if orphans and len(orphans) % 3 == 0:
            for g in range(0, len(orphans), 3):
                group = orphans[g:g + 3]
                serial, name, unit = (strip_tags(c[1]) for c in group)
                emit(serial, name, unit, "".join(c[1] for c in group),
                     tm.start() + group[0][0])
    return rows


def pdf_filename(href):
    """由連結取出 PDF 檔名。"""
    return href.split("?")[0].rstrip("/").split("/")[-1]


def dedupe(rows):
    """
    同一份 PDF 可能在多個分類重複列出（例如設置辦法同列綜合性與教務法規）。
    以檔名為鍵合併：取最長的法規名稱為正式名稱，其餘存為 aliases；
    分類與維護單位若有歧異則全數保留。
    """
    merged = OrderedDict()
    for r in sorted(rows, key=lambda x: x["pos"]):
        key = pdf_filename(r["href"]).lower()
        if key not in merged:
            merged[key] = {
                "file": pdf_filename(r["href"]),
                "name": r["name"],
                "aliases": [],
                "units": [r["unit"]],
                "categories": ["%s(#%s)" % (r["category"], r["anchor"])],
                "serials": [r["serial"]],
                "source": r["href"] if r["href"].startswith("http") else BASE_URL + pdf_filename(r["href"]),
            }
            continue
        e = merged[key]
        if r["name"] != e["name"] and r["name"] not in e["aliases"]:
            e["aliases"].append(r["name"])
        if r["unit"] not in e["units"]:
            e["units"].append(r["unit"])
        cat = "%s(#%s)" % (r["category"], r["anchor"])
        if cat not in e["categories"]:
            e["categories"].append(cat)
        if r["serial"] not in e["serials"]:
            e["serials"].append(r["serial"])

    # 標題被拆列的情形（如「輔英科技大學」+「數位學習平台管理要點」）：取最長者為正式名稱
    for e in merged.values():
        candidates = [e["name"]] + e["aliases"]
        best = max(candidates, key=len)
        if best != e["name"]:
            e["aliases"] = [c for c in candidates if c != best]
            e["name"] = best
    return list(merged.values())


# --------------------------------------------------------------------------- #
# 檔名與路徑
# --------------------------------------------------------------------------- #
def safe_name(name, stem, limit=70):
    n = unicodedata.normalize("NFC", name)
    n = re.sub(ILLEGAL_FS, "_", n).strip(" .")
    n = re.sub(r"_+", "_", n)
    if len(n) > limit:
        n = n[:limit].rstrip(" ._")
    if not n:
        n = "未命名法規"
    return "%s_%s.pdf" % (n, stem)


def target_dir(units):
    for u in units:
        if u in UNIT_DIRS:
            return ARCHIVE / UNIT_DIRS[u]
    return ARCHIVE / UNCLASSIFIED


# --------------------------------------------------------------------------- #
# 修訂日期擷取
# --------------------------------------------------------------------------- #
ROC_DATE = re.compile(r"(?<!\d)(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
# 少數法規的沿革寫成「91.01.25第一次校務會議通過」這種點號式民國日期
ROC_DOT = re.compile(r"(?<![\d.])(\d{2,3})\.(\d{1,2})\.(\d{1,2})(?![\d.])")
# 條文本體的起點，用來切出前面的「修正沿革」區塊。
# 必須錨在行首，否則會被沿革內的「除第三條第二款外…」這類字句誤切。
BODY_START = re.compile(
    r"^[ \t]*(第\s*[一二三四五六七八九十百]+\s*[篇章條]|一、|【餘修正歷程)", re.M)


def _max_roc_date(text):
    """取一段文字中最新的民國日期，回傳 (年, 月, 日) 或 None。"""
    best = None
    for pattern in (ROC_DATE, ROC_DOT):
        for m in pattern.finditer(text):
            y, mo, d = (int(x) for x in m.groups())
            if not (60 <= y <= 130 and 1 <= mo <= 12 and 1 <= d <= 31):
                continue
            key = (y, mo, d)
            if best is None or key > best:
                best = key
        if best is not None:
            break  # 「年月日」格式優先，找到就不再套用較易誤判的點號式
    return best


def latest_revision(pdf_path):
    """讀取 PDF 全文，取最新一筆民國日期，回傳 (民國表示, 西元 ISO)。"""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return None, None
    try:
        with fitz.open(pdf_path) as doc:
            text = "".join(page.get_text() for page in doc)
    except Exception:
        return None, None

    # 部分 PDF 的字型把「年」對映成 CJK 相容表意文字（如 U+F98E 而非 U+5E74），
    # 不先正規化會整份抓不到日期，或漏掉最新的一筆沿革。
    text = unicodedata.normalize("NFKC", text)

    # 先只看條文本體之前的沿革區塊；直接取全文最大值會誤抓條文裡的合約期間、
    # 施行日等未來日期（例：law01051 的合作期間至 116.6.30）。
    header = text[:BODY_START.search(text).start()] if BODY_START.search(text) else text[:1200]
    best = _max_roc_date(header) or _max_roc_date(text)
    if best is None:
        return None, None
    y, mo, d = best
    return "%d.%d.%d" % (y, mo, d), "%04d-%02d-%02d" % (y + 1911, mo, d)


# --------------------------------------------------------------------------- #
# 下載與索引
# --------------------------------------------------------------------------- #
def download(url, dest, retries=4, pause=2.0, cooldown=180):
    """
    下載單一 PDF。遇 WAF 攔截時採指數退避並拉長冷卻時間，
    因為該站是以來源 IP 做短時封鎖，硬重試只會延長封鎖。
    """
    for attempt in range(1, retries + 1):
        try:
            data = fetch(url, timeout=90, referer=INDEX_URL)
            if not data.startswith(b"%PDF"):
                return False, "回應非 PDF（前 8 bytes: %r）" % data[:8]
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(data)
            time.sleep(pause)
            return True, "%d bytes" % len(data)
        except WafBlocked as exc:
            if attempt == retries:
                return False, str(exc)
            wait = cooldown * attempt
            print("    ⏸ %s，冷卻 %d 秒後重試（第 %d 次）" % (exc, wait, attempt))
            time.sleep(wait)
        except (TimeoutError, OSError) as exc:
            if attempt == retries:
                return False, "%s: %s" % (type(exc).__name__, exc)
            time.sleep(2.0 * attempt)
    return False, "unknown"


def write_index(entries):
    csv_path = ARCHIVE / "法規彙整索引.csv"
    json_path = ARCHIVE / "法規彙整索引.json"
    cols = ["法規名稱", "維護單位", "法規分類", "最後修訂(民國)", "最後修訂(西元)",
            "存放路徑", "原始檔名", "來源網址", "別名", "狀態"]
    rows = []
    for e in entries:
        rows.append({
            "法規名稱": e["name"],
            "維護單位": "/".join(e["units"]),
            "法規分類": "；".join(e["categories"]),
            "最後修訂(民國)": e.get("revision_roc") or "",
            "最後修訂(西元)": e.get("revision_iso") or "",
            "存放路徑": str((e["dir"] / e["filename"]).relative_to(ARCHIVE)).replace("\\", "/"),
            "原始檔名": e["file"],
            "來源網址": e["source"],
            "別名": "；".join(e["aliases"]),
            "狀態": e.get("status", ""),
        })
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    json_path.write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    print("→ 索引已輸出：%s、%s" % (csv_path.name, json_path.name))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="只解析與列印，不下載")
    ap.add_argument("--refresh", action="store_true", help="強制重新下載目錄頁")
    ap.add_argument("--limit", type=int, default=0, help="只處理前 N 筆（測試用）")
    ap.add_argument("--pause", type=float, default=2.0,
                    help="每次下載間隔秒數（過短會觸發 WAF 封鎖，建議 >= 1.5）")
    ap.add_argument("--force", action="store_true", help="已存在的檔案也重新下載")
    args = ap.parse_args()

    cache = ROOT / ".cache" / "law_index_utf8.htm"
    print("→ 取得法規目錄頁：%s" % INDEX_URL)
    try:
        html = load_index_html(cache, args.refresh)
    except WafBlocked as exc:
        print("  ✗ %s — 來源站台目前拒絕連線，請稍後再試" % exc)
        return 2
    print("  頁面大小 %s 字元" % format(len(html), ","))

    rows = parse_rows(html)
    entries = dedupe(rows)
    print("  解析到 %d 個表格連結 → 去重後 %d 份法規" % (len(rows), len(entries)))

    for e in entries:
        e["dir"] = target_dir(e["units"])
        e["filename"] = safe_name(e["name"], Path(e["file"]).stem)

    unmapped = sorted(set(u for e in entries for u in e["units"] if u not in UNIT_DIRS))
    if unmapped:
        print("  ⚠ 未對應單位（將歸入 %s）：%s" % (UNCLASSIFIED, unmapped))

    if args.limit:
        entries = entries[: args.limit]

    by_dir = {}
    for e in entries:
        by_dir.setdefault(str(e["dir"].relative_to(ARCHIVE)).replace("\\", "/"), []).append(e)
    print("  歸檔分布：")
    for d in sorted(by_dir):
        print("    %-45s %3d 筆" % (d, len(by_dir[d])))

    if args.dry_run:
        return 0

    print("→ 開始下載 %d 份 PDF（間隔 %.1f 秒，可中斷後重跑續傳）…" % (len(entries), args.pause))
    ok = failed = skipped = 0
    for i, e in enumerate(entries, 1):
        dest = e["dir"] / e["filename"]
        if dest.exists() and dest.stat().st_size > 0 and not args.force:
            skipped += 1
            e["status"] = "已存在"
            e["bytes"] = dest.stat().st_size
        else:
            success, info = download(e["source"], dest, pause=args.pause)
            if success:
                ok += 1
                e["status"] = "OK"
                e["bytes"] = dest.stat().st_size
            else:
                failed += 1
                e["status"] = "FAILED: %s" % info
                print("  ✗ [%d/%d] %s — %s" % (i, len(entries), e["name"][:36], info))
        if dest.exists() and dest.stat().st_size > 0:
            roc, iso = latest_revision(dest)
            e["revision_roc"] = roc
            e["revision_iso"] = iso
        if i % 25 == 0 or i == len(entries):
            print("  … %d/%d（新下載 %d / 沿用 %d / 失敗 %d）" % (i, len(entries), ok, skipped, failed))

    write_index(entries)
    print("\n完成：新下載 %d 份、沿用既有 %d 份、失敗 %d 份" % (ok, skipped, failed))
    if failed:
        print("  失敗項目可直接重跑本程式續傳（已下載者會自動跳過）")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
