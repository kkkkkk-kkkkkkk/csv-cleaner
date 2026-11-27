#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
招投标公开信息采集器（合规版，标准库实现）
- 输入：种子列表页 URL（或本地 fixture）+ 字段配置
- 处理：列表抽取 -> 详情抓取 -> 字段清洗 -> 指纹去重 -> JSONL/SQLite 入库 -> 运行日志
- 只用公开页面，遵守 robots 与频率限制；不处理登录态、不碰个人隐私字段
"""
import os
import re
import json
import sqlite3
import time
import hashlib
import logging
from html.parser import HTMLParser
from urllib.parse import urljoin

LOG = logging.getLogger("tender")


class _Text(HTMLParser):
    """把 HTML 拍平成纯文本，保留 block 级换行，便于正则抽取。"""

    BLOCK = {"p", "div", "li", "tr", "td", "br", "h1", "h2", "h3", "span"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in self.BLOCK:
            self.parts.append("\n")
        if tag == "a":
            self.parts.append("\n[LINK]")

    def handle_endtag(self, tag):
        if tag in self.BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        self.parts.append(data)

    def text(self):
        raw = "".join(self.parts)
        lines = [re.sub(r"\s+", " ", l).strip() for l in raw.split("\n")]
        return "\n".join(l for l in lines if l)


def flatten(html):
    p = _Text()
    p.feed(html)
    return p.text()


def extract_links(html, base_url, pattern=r"/tender/|/notice/|/bid/"):
    """从列表页取详情链接，绝对化并去重（保留顺序）。"""
    hrefs = re.findall(r'href=["\']([^"\']+)["\']', html)
    out, seen = [], set()
    for h in hrefs:
        if not re.search(pattern, h):
            continue
        u = urljoin(base_url, h)
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


FIELD_RULES = {
    "title": r"(?:项目名称|招标项目名称|标题)\s*[:：]\s*(.+)",
    "buyer": r"(?:采购人|招标人|采购单位)\s*[:：]\s*(.+)",
    "deadline": r"(?:投标截止时间|报名截止时间|开标时间)\s*[:：]\s*(.+)",
    "region": r"(?:项目所在地|所属地区|地区)\s*[:：]\s*(.+)",
    "notice_type": r"(?:公告类型|信息类型)\s*[:：]\s*(.+)",
    "publish_time": r"(?:发布时间|发布日期)\s*[:：]\s*(.+)",
}
BUDGET_RULE = r"(?:预算金额|预算|控制价)\s*[:：]\s*([\d,，.]+)\s*(万元|元)?"


def parse_detail(html):
    """按 FIELD_RULES 抽字段；金额统一换算为元（整数）。"""
    text = flatten(html)
    rec = {}
    for k, rx in FIELD_RULES.items():
        m = re.search(rx, text)
        rec[k] = m.group(1).strip() if m else None
    b = re.search(BUDGET_RULE, text)
    if b:
        amount = float(b.group(1).replace(",", "").replace("，", ""))
        unit = b.group(2) or "元"
        rec["budget_yuan"] = int(round(amount * 10000)) if unit == "万元" else int(round(amount))
    else:
        rec["budget_yuan"] = None
    rec["title"] = rec["title"] or _first_heading(html)
    return rec


def _first_heading(html):
    m = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.S | re.I)
    return re.sub(r"<[^>]+>", "", m.group(1)).strip() if m else None


def fingerprint(rec):
    """业务指纹：标题+采购人+预算，标准化后哈希；同标的重复公告会被判重。"""
    key = "|".join(_norm(rec.get(k)) for k in ("title", "buyer", "budget_yuan"))
    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:16]


def _norm(s):
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"（.*?）|\(.*?\)", "", s)
    s = re.sub(r"[^\w\u4e00-\u9fff]", "", s)
    return s.lower()


def store_jsonl(path, records):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def store_sqlite(path, records):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute("""CREATE TABLE IF NOT EXISTS tenders(
        fp TEXT PRIMARY KEY, title TEXT, buyer TEXT, budget_yuan INTEGER,
        deadline TEXT, region TEXT, notice_type TEXT, publish_time TEXT, url TEXT)""")
    for r in records:
        conn.execute(
            "INSERT OR REPLACE INTO tenders VALUES (?,?,?,?,?,?,?,?,?)",
            (r["fp"], r["title"], r["buyer"], r["budget_yuan"], r["deadline"],
             r["region"], r["notice_type"], r["publish_time"], r["url"]),
        )
    conn.commit()
    conn.close()


def fetch_with_retry(fetch, url, retries=3, backoff=0.4, stats=None, sleeper=time.sleep):
    """带退避重试的抓取。失败不抛异常，返回 (html, ok, last_error)，让失败进入账而不是炸掉整轮。"""
    last = None
    for attempt in range(1, retries + 1):
        try:
            return fetch(url), True, None
        except Exception as e:  # noqa: BLE001 —— 抓取层要吞掉一切异常
            last = "%s: %s" % (type(e).__name__, e)
            if stats is not None:
                stats["retries"] = stats.get("retries", 0) + 1
            LOG.warning("fetch fail %s (attempt %d/%d) %s", url, attempt, retries, last)
            if attempt < retries:
                sleeper(backoff * (2 ** (attempt - 1)))
    if stats is not None:
        stats["errors"] = stats.get("errors", 0) + 1
    return None, False, last


def store_xlsx(path, records, failures=None, stats=None):
    """导出给人看的 xlsx：主表 + 失败清单 + 字段填充率。

    和 jsonl/sqlite 的区别：这三样是给程序看的，xlsx 是给甲方看的——
    行高、列宽、筛选、冻结首行都调好，甲方打开就能用，不用再格式化。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "tenders"
    cols = [
        ("序号", 6), ("项目名称", 44), ("采购人", 26), ("预算(元)", 14),
        ("预算(万元)", 12), ("投标截止", 20), ("地区", 14),
        ("公告类型", 12), ("发布时间", 16), ("字段完整度", 12), ("来源链接", 58),
    ]
    head_fill = PatternFill("solid", fgColor="1F3864")
    head_font = Font(color="FFFFFF", bold=True, size=10)
    for i, (name, width) in enumerate(cols, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.fill = head_fill
        c.font = head_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 24

    need = ("title", "buyer", "budget_yuan", "deadline", "region", "notice_type", "publish_time")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    for idx, r in enumerate(records, start=1):
        row = idx + 1
        filled = sum(1 for k in need if r.get(k) not in (None, ""))
        vals = [
            idx, r.get("title"), r.get("buyer"), r.get("budget_yuan"),
            (round(r["budget_yuan"] / 10000, 2) if r.get("budget_yuan") else None),
            r.get("deadline"), r.get("region"), r.get("notice_type"),
            r.get("publish_time"), "%d/%d" % (filled, len(need)), r.get("url"),
        ]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.alignment = Alignment(vertical="center", wrap_text=(i in (2, 11)))
        ws.cell(row=row, column=4).number_format = "#,##0"
        ws.cell(row=row, column=6).alignment = Alignment(horizontal="center")
        if filled < len(need):
            ws.cell(row=row, column=10).fill = warn_fill
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cols)), len(records) + 1)

    # 第二页：没抓到的链接。甲方最想知道"少了什么"，而不是只看"抓到了什么"
    ws2 = wb.create_sheet("failed")
    for i, name in enumerate(("URL", "失败原因"), start=1):
        c = ws2.cell(row=1, column=i, value=name)
        c.fill = head_fill
        c.font = head_font
    ws2.column_dimensions["A"].width = 62
    ws2.column_dimensions["B"].width = 56
    for idx, f in enumerate(failures or [], start=2):
        ws2.cell(row=idx, column=1, value=f.get("url"))
        ws2.cell(row=idx, column=2, value=f.get("reason"))
    if not (failures or []):
        ws2.cell(row=2, column=1, value="(本轮无失败链接)")

    # 第三页：字段填充率。判断"这站值不值得长期抓"的唯一客观依据
    ws3 = wb.create_sheet("fill_rate")
    ws3.cell(row=1, column=1, value="字段").font = Font(bold=True)
    ws3.cell(row=1, column=2, value="填充率").font = Font(bold=True)
    ws3.column_dimensions["A"].width = 18
    ws3.column_dimensions["B"].width = 12
    fr = (stats or {}).get("fill_rate", {})
    for idx, (k, v) in enumerate(fr.items(), start=2):
        ws3.cell(row=idx, column=1, value=k)
        c = ws3.cell(row=idx, column=2, value=v)
        c.number_format = "0.0%"
    ws3.cell(row=len(fr) + 3, column=1, value="条数").font = Font(bold=True)
    ws3.cell(row=len(fr) + 3, column=2, value=len(records))
    ws3.cell(row=len(fr) + 4, column=1, value="重试次数").font = Font(bold=True)
    ws3.cell(row=len(fr) + 4, column=2, value=(stats or {}).get("retries", 0))

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


def crawl(pages, delay=1.0, retries=3, sleeper=time.sleep, fetch=None):
    """pages: [(url, html)]，返回 (records, stats, failures)。

    stats 里除了计数还有 fill_rate：每个字段的填充率，用来判断"这站值不值得长期抓"。
    """
    t0 = time.time()
    fetcher = fetch or FETCH
    records, seen = [], set()
    stats = {"pages": 0, "details": 0, "dup": 0, "missing": 0, "retries": 0, "errors": 0}
    failures = []
    for list_url, list_html in pages:
        stats["pages"] += 1
        for detail_url in extract_links(list_html, list_url):
            stats["details"] += 1
            html, ok, err = fetch_with_retry(fetcher, detail_url, retries=retries,
                                             stats=stats, sleeper=sleeper)
            if not ok:
                failures.append({"url": detail_url, "reason": err})
                continue
            rec = parse_detail(html)
            rec["url"] = detail_url
            rec["fp"] = fingerprint(rec)
            if rec["fp"] in seen:
                stats["dup"] += 1
                LOG.info("dup skip %s %s", rec["fp"], rec["title"])
                continue
            if not rec["title"] or rec["budget_yuan"] is None:
                stats["missing"] += 1
            seen.add(rec["fp"])
            records.append(rec)
            sleeper(delay)  # 频率控制
    fields = ("title", "buyer", "deadline", "region", "notice_type", "publish_time", "budget_yuan")
    n = len(records) or 1
    stats["fill_rate"] = {f: round(sum(1 for r in records if r.get(f) not in (None, "")) / n, 3)
                          for f in fields}
    stats["failed"] = len(failures)
    stats["elapsed"] = round(time.time() - t0, 2)
    return records, stats, failures


def local_fetch(root):
    """把 fixture 目录当站点，URL 直接映射到本地文件（离线跑测试用）。"""

    def _f(url):
        path = os.path.join(root, url.split("/")[-1])
        with open(path, encoding="utf-8") as fh:
            return fh.read()

    return _f


FETCH = None
