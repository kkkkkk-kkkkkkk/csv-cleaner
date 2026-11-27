# -*- coding: utf-8 -*-
"""发票/单据 PDF 批量抽取 + 校验 + 台账导出。

交付物三样：
  1) 台账 xlsx（汇总 / 复核清单 / 统计 三页）
  2) jsonl 全量留档
  3) 每条不确定都进复核清单，不静默丢弃

只用公开可得的单据格式，不碰任何加密/受保护 PDF，不做扫描件 OCR 之外的绕过。
"""
import os
import re
import json
import time
import datetime
import pymupdf

FIELDS = ("invoice_no", "date", "buyer", "seller", "item", "amount", "tax", "total")

RE_NO = re.compile(r"发票号码\s*[:：]\s*([0-9]{6,20})")
RE_DATE = re.compile(r"开票日期\s*[:：]\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
RE_BUYER = re.compile(r"购买方名称\s*[:：]\s*(.+)")
RE_SELLER = re.compile(r"销售方名称\s*[:：]\s*(.+)")
RE_ITEM = re.compile(r"项目名称\s*[:：]\s*(.+)")
RE_AMOUNT = re.compile(r"(?<!合计)金额\s*[:：]\s*[¥￥]?\s*([\d,]+(?:\.\d+)?)")
RE_TAX = re.compile(r"税额\s*[:：]\s*[¥￥]?\s*([\d,]+(?:\.\d+)?)")
RE_TOTAL = re.compile(r"价税合计\s*[:：]\s*[¥￥]?\s*([\d,]+(?:\.\d+)?)")

CENT = 0.01


def parse_amount(s):
    if s is None:
        return None
    try:
        return round(float(str(s).replace(",", "").replace("，", "").strip()), 2)
    except ValueError:
        return None


def normalize_date(y, m, d):
    """返回 (iso, err)。非法日期不猜、不填 0，直接标错误进复核清单。"""
    try:
        dt = datetime.date(int(y), int(m), int(d))
    except ValueError as e:
        return None, "非法日期 %s-%s-%s (%s)" % (y, m, d, e)
    return dt.isoformat(), None


def read_pdf_text(path):
    doc = pymupdf.open(path)
    try:
        pages = [p.get_text() for p in doc]
    finally:
        doc.close()
    return "\n".join(pages)


def extract_one(path):
    """抽一张单据。返回记录 + 抽取阶段问题（不抛异常）。"""
    rec = {k: None for k in FIELDS}
    rec["source"] = os.path.basename(path)
    rec["path"] = path
    problems = []
    try:
        text = read_pdf_text(path)
    except Exception as e:  # noqa: BLE001
        problems.append(("EXTRACT_FAILED", "high", "text", "读不出文本：%s: %s" % (type(e).__name__, e)))
        rec["char_count"] = 0
        return rec, problems

    rec["char_count"] = len(text)
    if len(text.strip()) < 20:
        problems.append(("EMPTY_TEXT", "high", "text", "文本层几乎为空（可能是扫描件，需 OCR）"))

    m = RE_NO.search(text)
    if m:
        rec["invoice_no"] = m.group(1)
    m = RE_DATE.search(text)
    if m:
        iso, err = normalize_date(*m.groups())
        rec["date"] = iso or m.group(0)
        if err:
            problems.append(("BAD_DATE", "high", "date", err))
    for key, rx in (("buyer", RE_BUYER), ("seller", RE_SELLER), ("item", RE_ITEM)):
        m = rx.search(text)
        if m:
            rec[key] = m.group(1).strip()
    for key, rx in (("amount", RE_AMOUNT), ("tax", RE_TAX), ("total", RE_TOTAL)):
        m = rx.search(text)
        if m:
            rec[key] = parse_amount(m.group(1))
    rec["pages"] = rec.get("pages", 0)
    return rec, problems


def validate(rec, seen_nos=None, tol=0.02):
    """业务校验。返回 [(code, level, field, msg)]。"""
    issues = []
    if seen_nos is None:
        seen_nos = {}

    if not rec.get("invoice_no"):
        issues.append(("MISSING_NO", "high", "invoice_no", "缺发票号码：无法核验真伪，也无法查重"))
    else:
        no = rec["invoice_no"]
        if no in seen_nos and seen_nos[no] != rec.get("source"):
            issues.append(("DUP_NO", "high", "invoice_no",
                           "号码与 %s 重复：可能是补开、重开或重复报销" % seen_nos[no]))
        seen_nos.setdefault(no, rec.get("source"))

    if not rec.get("date") or not re.match(r"^\d{4}-\d{2}-\d{2}$", str(rec.get("date"))):
        issues.append(("BAD_DATE", "mid", "date", "日期缺失或格式不规范，无法归入账期"))
    if not rec.get("buyer"):
        issues.append(("MISSING_BUYER", "mid", "buyer", "缺购买方名称：无法判断是否本公司抬头"))
    if not rec.get("seller"):
        issues.append(("MISSING_SELLER", "mid", "seller", "缺销售方名称：无法做供应商对账"))

    amount, tax, total = rec.get("amount"), rec.get("tax"), rec.get("total")
    if amount is None and tax is None and total is None:
        issues.append(("MISSING_AMOUNT", "high", "amount", "三个金额字段全缺"))
    else:
        if amount == 0 and tax == 0:
            issues.append(("ZERO_AMOUNT", "high", "amount", "金额与税额均为 0：空票或占位票，不应入账"))
        if None not in (amount, tax, total) and abs((amount + tax) - total) > tol:
            issues.append(("TOTAL_MISMATCH", "high", "total",
                           "价税合计 %.2f ≠ 金额 %.2f + 税额 %.2f（差 %.2f）"
                           % (total, amount, tax, total - (amount + tax))))
    return issues


def extract_batch(paths):
    t0 = time.time()
    rows, issues, seen = [], [], {}
    for p in paths:
        rec, probs = extract_one(p)
        issues.extend([{"code": c, "level": lv, "field": f, "msg": m, "source": rec["source"]}
                       for c, lv, f, m in probs])
        for it in validate(rec, seen):
            issues.append({"code": it[0], "level": it[1], "field": it[2],
                           "msg": it[3], "source": rec["source"]})
        rows.append(rec)
    hard = {i["source"] for i in issues if i["level"] == "high"}
    for r in rows:
        r["status"] = "需复核" if r["source"] in hard else "通过"
        r["issue_count"] = sum(1 for i in issues if i["source"] == r["source"])
    stats = {
        "files": len(paths),
        "rows": len(rows),
        "passed": sum(1 for r in rows if r["status"] == "通过"),
        "flagged": sum(1 for r in rows if r["status"] == "需复核"),
        "issues": len(issues),
        "issue_high": sum(1 for i in issues if i["level"] == "high"),
        "amount_sum": round(sum(r["amount"] or 0 for r in rows), 2),
        "total_sum": round(sum(r["total"] or 0 for r in rows), 2),
        "elapsed": round(time.time() - t0, 2),
    }
    return rows, issues, stats


def store_jsonl(path, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def store_xlsx(path, rows, issues, stats):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "发票台账"
    cols = [("序号", 6), ("源文件", 20), ("发票号码", 18), ("开票日期", 13), ("购买方", 26),
            ("销售方", 26), ("项目", 22), ("金额", 12), ("税额", 12), ("价税合计", 12),
            ("状态", 10), ("问题数", 8)]
    head = PatternFill("solid", fgColor="1F3864")
    hf = Font(color="FFFFFF", bold=True, size=10)
    for i, (n, w) in enumerate(cols, 1):
        c = ws.cell(row=1, column=i, value=n)
        c.fill, c.font = head, hf
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cols)), len(rows) + 1)

    red = PatternFill("solid", fgColor="FCE4E4")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    for idx, r in enumerate(rows, 1):
        row = idx + 1
        vals = [idx, r["source"], r["invoice_no"], r["date"], r["buyer"], r["seller"], r["item"],
                r["amount"], r["tax"], r["total"], r["status"], r["issue_count"]]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=i, value=v)
            c.alignment = Alignment(vertical="center", wrap_text=(i in (5, 6, 7)))
            if i in (8, 9, 10):
                c.number_format = "#,##0.00"
        if r["status"] == "需复核":
            ws.cell(row=row, column=11).fill = yellow
        for i in (8, 9, 10):
            if isinstance(vals[i - 1], float) and vals[i - 1] == 0:
                ws.cell(row=row, column=i).fill = red

    ws2 = wb.create_sheet("复核清单")
    heads = [("源文件", 20), ("级别", 8), ("字段", 14), ("问题", 78), ("代码", 18)]
    for i, (n, w) in enumerate(heads, 1):
        c = ws2.cell(row=1, column=i, value=n)
        c.fill, c.font = head, hf
        ws2.column_dimensions[get_column_letter(i)].width = w
    if not issues:
        ws2.cell(row=2, column=1, value="(本轮无问题)")
    for idx, it in enumerate(issues, 2):
        ws2.cell(row=idx, column=1, value=it["source"])
        ws2.cell(row=idx, column=2, value=it["level"])
        ws2.cell(row=idx, column=3, value=it["field"])
        ws2.cell(row=idx, column=4, value=it["msg"]).alignment = Alignment(wrap_text=True)
        ws2.cell(row=idx, column=5, value=it["code"])
        if it["level"] == "high":
            ws2.cell(row=idx, column=2).fill = red

    ws3 = wb.create_sheet("统计")
    pairs = [("文件数", stats["files"]), ("通过", stats["passed"]), ("需复核", stats["flagged"]),
             ("问题条数", stats["issues"]), ("其中高危", stats["issue_high"]),
             ("金额合计", stats["amount_sum"]), ("价税合计汇总", stats["total_sum"]),
             ("耗时(秒)", stats["elapsed"])]
    ws3.cell(row=1, column=1, value="指标").font = Font(bold=True)
    ws3.cell(row=1, column=2, value="值").font = Font(bold=True)
    ws3.column_dimensions["A"].width = 20
    ws3.column_dimensions["B"].width = 16
    for i, (k, v) in enumerate(pairs, 2):
        ws3.cell(row=i, column=1, value=k)
        ws3.cell(row=i, column=2, value=v)

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path
