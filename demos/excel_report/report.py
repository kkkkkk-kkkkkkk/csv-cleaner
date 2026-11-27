# -*- coding: utf-8 -*-
"""Excel/CSV 清洗 + 公式化汇总 + 自校验 + 版本对比（弹药包 #4）

用法：
    python report.py 输入.csv 输出.xlsx                        # 出报表
    python report.py 输入.csv 输出.xlsx --baseline 上期.xlsx   # 出报表 + 差异页
    python report.py 输入.csv 输出.xlsx --json out.json        # 另存汇总 json（下次做 baseline 用）

做什么：
  1) 自动识别编码（utf-8-sig / utf-8 / gbk / gb18030）
  2) 清洗：去空行、去完全重复行、字段去空格、区域别名归一、多格式日期、
           金额剥离 ￥/元/千分位、非法数量剔除；金额/数量缺失单独标记
  3) 汇总：用 Excel 原生 SUMIFS 公式引用明细 —— 改明细，汇总自动跟着变
  4) 核对：公式写的三向交叉校验（明细合计 vs 区域分组 vs 月度分组），
           差异为 0 才 PASS，不是漂亮话，是可复核的公式
  5) 口径页：源文件、行数、剔除数、清洗规则、生成时间，一封邮件讲清楚
  6) 差异页：与上期报表对比金额/行数/区域变动
"""
import sys, os, re, csv, io, json, datetime, collections

DETAIL = "明细"  # 表名不带括号/空格，SUMIFS 里可直接引用，避免引号问题


def sniff_text(path):
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return f.read(), enc
        except UnicodeDecodeError:
            continue
    raise SystemExit("无法识别编码: %s" % path)


REGION_MAP = {
    "华东": "华东", "华东区": "华东", "华南": "华南", "华北": "华北", "华北区": "华北",
    "西南": "西南", "SW": "西南", "North": "华北",
}
DATE_FMTS = ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%Y.%m.%d")


def norm_region(v):
    v = (v or "").strip()
    return REGION_MAP.get(v, v or "未标注")


def parse_date(v):
    v = (v or "").strip()
    if not v:
        return None, "missing"
    for fmt in DATE_FMTS:
        try:
            return datetime.datetime.strptime(v, fmt).date(), "ok"
        except ValueError:
            continue
    return None, "bad"


def parse_money(v):
    """金额解析：剥离 ￥/¥/元/空格，区分“千分位逗号”与“小数逗号”。

    - `1,234.50` / `1，234.50` → 1234.50（逗号按千分位）
    - `1234，50`              → 1234.50（逗号按小数点，中文导出常见）
    - `￥ 1,234`              → 1234.00
    """
    raw = (v or "").strip()
    if not raw:
        return None, False
    s = re.sub(r"[￥¥元\s]", "", raw).replace("，", ",")
    if not s:
        return None, False
    dirty = s != raw
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", s):
        return float(s.replace(",", "")), dirty
    if re.fullmatch(r"-?\d+,\d{1,2}", s):
        return float(s.replace(",", ".")), True
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return float(s), dirty
    try:
        return float(s.replace(",", "")), True
    except ValueError:
        return None, dirty


def parse_int(v):
    v = (v or "").strip()
    if not v:
        return None
    try:
        return int(float(v))
    except Exception:
        return None


def load_rows(path):
    text, enc = sniff_text(path)
    rd = csv.reader(io.StringIO(text))
    header = [h.strip() for h in next(rd)]
    raw = []
    for r in rd:
        r = list(r) + [""] * (len(header) - len(r))
        raw.append(r[:len(header)])
    return header, raw, enc


def clean(header, raw):
    stats = collections.Counter()
    out = []
    for i, r in enumerate(raw, start=2):
        if not any((c or "").strip() for c in r):
            stats["空行"] += 1
            continue
        stats["源行"] += 1
        d = dict(zip(header, [(c or "").strip() for c in r]))
        raw_region = (d.get("区域", "") or "").strip()
        dt, dt_state = parse_date(d.get("下单日期", ""))
        amt, amt_fixed = parse_money(d.get("金额", ""))
        qty = parse_int(d.get("数量", ""))

        flags, drop = [], False
        if dt_state == "missing":
            flags.append("日期缺失")
            stats["日期缺失"] += 1
        elif dt_state == "bad":
            flags.append("日期不可解析")
            stats["日期异常"] += 1
        if amt_fixed:
            flags.append("金额已剥离符号")
            stats["金额符号修正"] += 1
        if amt is None:
            flags.append("金额缺失·未计入合计")
            stats["金额缺失"] += 1
            amt, drop = 0.0, True
        if qty is None:
            flags.append("数量缺失")
            stats["数量缺失"] += 1
            qty = 0
        elif qty <= 0:
            flags.append("数量非正·已剔除")
            stats["数量非正"] += 1
            drop = True
        region = norm_region(raw_region)
        if raw_region != region:
            flags.append("区域已归一")
            stats["区域归一"] += 1

        out.append({"row": i, "order": d.get("订单号", ""), "date": dt, "region": region,
                    "product": (d.get("产品", "") or "").strip(), "qty": qty,
                    "amount": amt, "sales": (d.get("业务员", "") or "").strip(),
                    "flag": "、".join(flags), "drop": drop})

    seen, dedup = set(), []
    for r in out:
        key = (r["order"], r["date"], r["region"], r["product"], r["qty"], round(r["amount"], 2), r["sales"])
        if r["order"] and key in seen:
            stats["重复行"] += 1
            continue
        seen.add(key)
        dedup.append(r)
    stats["有效行"] = sum(1 for r in dedup if not r["drop"])
    stats["剔除行"] = sum(1 for r in dedup if r["drop"])
    return dedup, stats


def summarize(rows):
    good = [r for r in rows if not r["drop"]]
    by_region, by_month, by_product = (collections.defaultdict(lambda: [0, 0.0]) for _ in range(3))
    for r in good:
        by_region[r["region"]][0] += r["qty"]
        by_region[r["region"]][1] += r["amount"]
        k = r["date"].strftime("%Y-%m") if r["date"] else "日期未知"
        by_month[k][0] += r["qty"]
        by_month[k][1] += r["amount"]
        by_product[r["product"]][0] += r["qty"]
        by_product[r["product"]][1] += r["amount"]
    return good, by_region, by_month, by_product


NCOLS = 11  # 明细列数（末尾多一列“月份”，供 SUMIFS 按月汇总）


def write_xlsx(out_path, rows, by_region, by_month, by_product, stats, meta=None, baseline=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference

    HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
    HEAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    BODY_FONT = Font(name="微软雅黑", size=10)
    TITLE_FONT = Font(name="微软雅黑", size=12, bold=True)
    THIN = Side(style="thin", color="D9D9D9")
    OK_FONT = Font(name="微软雅黑", size=10, bold=True, color="1E7B34")
    BAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="B00020")

    def head_row(ws, row, values):
        for i, v in enumerate(values, start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.fill, c.font = HEAD_FILL, HEAD_FONT
            c.alignment = Alignment(horizontal="center", vertical="center")

    wb = Workbook()
    tot_amt = round(sum(r["amount"] for r in rows if not r["drop"]), 2)
    tot_qty = sum(r["qty"] for r in rows if not r["drop"])
    good_rows = stats["有效行"]

    # ---------- Sheet 0 口径 ----------
    ws0 = wb.active
    ws0.title = "口径"
    ws0["A1"] = (meta or {}).get("title", "销售报表")
    ws0["A1"].font = Font(name="微软雅黑", size=14, bold=True)
    info = [
        ("源文件", (meta or {}).get("src", "")),
        ("生成时间", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("源文件编码", (meta or {}).get("enc", "")),
        ("原始数据行", stats.get("源行", 0)),
        ("有效行", good_rows),
        ("剔除行", stats.get("剔除行", 0)),
        ("重复行", stats.get("重复行", 0)),
        ("有效金额合计", tot_amt),
        ("有效数量合计", tot_qty),
        ("", ""),
    ]
    rules = [
        "1) 空行、完全重复行（同订单号+日期+区域+产品+数量+金额+业务员）剔除，重复行不计入合计",
        "2) 金额缺失（空值或非数字）的行保留在明细、标记“未计入合计”，不计入任何汇总 —— 避免把缺失当 0 拉低总额",
        "3) 数量非正（0 或负数）视为无效行剔除",
        "4) 区域别名归一（华东区→华东、SW/西南→西南、North→华北），未填区域记“未标注”",
        "5) 日期支持 2026-01-05 / 2026/01/05 / 05/01/2026 / 2026.01.05，不可解析的留空并标记，不猜",
        "6) 金额剥离 ￥ ¥ 元 千分位逗号与空格后再转数字",
        "7) 汇总页全部为 SUMIFS 公式引用明细，改明细汇总自动更新；核对页为三向交叉校验公式",
    ]
    r = 3
    for k, v in info:
        ws0.cell(row=r, column=1, value=k).font = Font(name="微软雅黑", size=10, bold=True)
        c = ws0.cell(row=r, column=2, value=v)
        c.font = BODY_FONT
        if isinstance(v, float):
            c.number_format = "#,##0.00"
        r += 1
    r += 1
    ws0.cell(row=r, column=1, value="清洗与计算口径").font = TITLE_FONT
    r += 1
    for i, t in enumerate(rules, start=1):
        ws0.cell(row=r, column=1, value="%d." % i).font = BODY_FONT
        ws0.cell(row=r, column=2, value=t).font = BODY_FONT
        r += 1
    ws0.column_dimensions["A"].width = 18
    ws0.column_dimensions["B"].width = 96

    # ---------- Sheet 1 明细 ----------
    ws = wb.create_sheet(DETAIL)
    head = ["源行号", "订单号", "下单日期", "区域", "产品", "数量", "金额", "业务员", "清洗标记", "是否剔除", "月份"]
    ws.append(head)
    for x in rows:
        ws.append([x["row"], x["order"],
                   x["date"].strftime("%Y-%m-%d") if x["date"] else "",
                   x["region"], x["product"], x["qty"], round(x["amount"], 2),
                   x["sales"], x["flag"], "是" if x["drop"] else "",
                   x["date"].strftime("%Y-%m") if x["date"] else "日期未知"])
    head_row(ws, 1, head)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22
    last = max(ws.max_row, 2)
    for row in ws.iter_rows(min_row=2, max_row=last, max_col=NCOLS):
        for c in row:
            c.font = BODY_FONT
            c.border = Border(bottom=THIN)
            if c.column in (6, 7):
                c.number_format = "#,##0" if c.column == 6 else "#,##0.00"
    for i, w in enumerate([8, 10, 12, 9, 13, 7, 11, 9, 28, 9, 10], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(NCOLS), last)

    # ---------- Sheet 2 汇总（SUMIFS 公式） ----------
    s2 = wb.create_sheet("汇总")
    rng = lambda col: "%s!$%s$2:$%s$%d" % (DETAIL, col, col, last)
    kept = '%s!$J$2:$J$%d,"<>是"' % (DETAIL, last)

    KEYCOL = {"区域": "D", "产品": "E", "月份": "K"}

    def block(title, keylabel, d, order):
        rowstart = block.rowstart
        s2.cell(row=rowstart, column=1, value=title).font = TITLE_FONT
        head_row(s2, rowstart + 1, [keylabel, "数量", "金额", "占比"])
        kc = rng(KEYCOL[keylabel])
        keys = order(d) if callable(order) else order
        for j, k in enumerate(keys, start=1):
            rr = rowstart + 1 + j
            s2.cell(row=rr, column=1, value=k).font = BODY_FONT
            s2.cell(row=rr, column=2, value='=SUMIFS(%s,%s,$A%d,%s)' % (rng("F"), kc, rr, kept))
            amt = s2.cell(row=rr, column=3, value='=SUMIFS(%s,%s,$A%d,%s)' % (rng("G"), kc, rr, kept))
            amt.number_format = "#,##0.00"
            pct = s2.cell(row=rr, column=4, value='=IF(SUM($C$%d:$C$%d)=0,"",C%d/SUM($C$%d:$C$%d))' % (rowstart + 2, rowstart + 1 + len(keys), rr, rowstart + 2, rowstart + 1 + len(keys)))
            pct.number_format = "0.0%"
        endr = rowstart + 2 + len(keys)
        s2.cell(row=endr, column=1, value="小计").font = Font(name="微软雅黑", size=10, bold=True)
        for col in (2, 3):
            c = s2.cell(row=endr, column=col, value="=SUM(%s%d:%s%d)" % (get_column_letter(col), rowstart + 2, get_column_letter(col), endr - 1))
            c.font = Font(name="微软雅黑", size=10, bold=True)
            c.number_format = "#,##0" if col == 2 else "#,##0.00"
        return endr

    by_region_order = lambda d: sorted(d, key=lambda x: -d[x][1])
    by_product_order = lambda d: sorted(d, key=lambda x: -d[x][1])
    by_month_order = lambda d: sorted(d)
    block.rowstart = 1
    e1 = block("区域汇总", "区域", by_region, by_region_order)
    block.rowstart = e1 + 2
    e2 = block("月度汇总", "月份", by_month, by_month_order)
    block.rowstart = e2 + 2
    e3 = block("产品汇总", "产品", by_product, by_product_order)
    for w, col in ((26, "A"), (10, "B"), (14, "C"), (9, "D")):
        s2.column_dimensions[col].width = w

    # ---------- Sheet 3 核对（公式三向校验） ----------
    s3 = wb.create_sheet("核对")
    s3["A1"] = "三向交叉校验（公式实时计算，差异必须为 0）"
    s3["A1"].font = TITLE_FONT
    checks = [
        ("金额合计：明细 vs 区域分组", '=SUMIFS(%s,%s)' % (rng("G"), kept), "=汇总!C%d" % (e1,), "金额"),
        ("金额合计：明细 vs 月度分组", '=SUMIFS(%s,%s)' % (rng("G"), kept), "=汇总!C%d" % (e2,), "金额"),
        ("金额合计：明细 vs 产品分组", '=SUMIFS(%s,%s)' % (rng("G"), kept), "=汇总!C%d" % (e3,), "金额"),
        ("数量合计：明细 vs 区域分组", '=SUMIFS(%s,%s)' % (rng("F"), kept), "=汇总!B%d" % (e1,), "数量"),
        ("数量合计：明细 vs 月度分组", '=SUMIFS(%s,%s)' % (rng("F"), kept), "=汇总!B%d" % (e2,), "数量"),
        ("数量合计：明细 vs 产品分组", '=SUMIFS(%s,%s)' % (rng("F"), kept), "=汇总!B%d" % (e3,), "数量"),
        ("有效行数：明细 vs 脚本统计", '=COUNTIFS(%s!$J$2:$J$%d,"<>是")' % (DETAIL, last), "=%d" % good_rows, "行"),
    ]
    head_row(s3, 3, ["校验项", "明细口径", "汇总口径", "差异", "判定"])
    for i, (name, f1, f2, kind) in enumerate(checks, start=1):
        rr = 3 + i
        s3.cell(row=rr, column=1, value=name).font = BODY_FONT
        c1 = s3.cell(row=rr, column=2, value=f1)
        c2 = s3.cell(row=rr, column=3, value=f2)
        for c in (c1, c2):
            c.font = BODY_FONT
            c.number_format = "#,##0" if kind == "行" else "#,##0.00"
        c3 = s3.cell(row=rr, column=4, value="=B%d-C%d" % (rr, rr))
        c3.font = BODY_FONT
        c3.number_format = "#,##0" if kind == "行" else "#,##0.000"
        c4 = s3.cell(row=rr, column=5, value='=IF(ABS(D%d)<0.005,"PASS","FAIL")' % rr)
        c4.font = OK_FONT
    for w, col in ((30, "A"), (16, "B"), (16, "C"), (12, "D"), (10, "E")):
        s3.column_dimensions[col].width = w

    # ---------- Sheet 4 差异（可选） ----------
    if baseline:
        s4 = wb.create_sheet("差异")
        s4["A1"] = "与上期对比"
        s4["A1"].font = TITLE_FONT
        head_row(s4, 3, ["项目", "上期", "本期", "差额", "变动率"])
        rows_d = [
            ("有效金额合计", baseline.get("tot_amount", 0), tot_amt),
            ("有效行数", baseline.get("rows", 0), good_rows),
            ("剔除行", baseline.get("dropped", 0), stats.get("剔除行", 0)),
            ("重复行", baseline.get("dupes", 0), stats.get("重复行", 0)),
        ]
        for i, (name, old, new) in enumerate(rows_d, start=1):
            rr = 3 + i
            s4.cell(row=rr, column=1, value=name).font = BODY_FONT
            s4.cell(row=rr, column=2, value=old).number_format = "#,##0.00" if name == "有效金额合计" else "#,##0"
            s4.cell(row=rr, column=3, value=new).number_format = "#,##0.00" if name == "有效金额合计" else "#,##0"
            c = s4.cell(row=rr, column=4, value="=C%d-B%d" % (rr, rr))
            c.number_format = "#,##0.00"
            c5 = s4.cell(row=rr, column=5, value='=IF(B%d=0,"新增",D%d/B%d)' % (rr, rr, rr))
            c5.number_format = "0.0%"
        s4.cell(row=3 + len(rows_d) + 2, column=1, value="区域变动").font = TITLE_FONT
        r0 = 3 + len(rows_d) + 3
        head_row(s4, r0, ["区域", "上期", "本期", "差额"])
        old_rg = baseline.get("by_region", {})
        for i, k in enumerate(sorted(set(list(old_rg) + list(by_region))), start=1):
            rr = r0 + i
            s4.cell(row=rr, column=1, value=k).font = BODY_FONT
            s4.cell(row=rr, column=2, value=round(old_rg.get(k, 0), 2)).number_format = "#,##0.00"
            s4.cell(row=rr, column=3, value=round(by_region.get(k, [0, 0])[1], 2)).number_format = "#,##0.00"
            s4.cell(row=rr, column=4, value="=C%d-B%d" % (rr, rr)).number_format = "#,##0.00"
        for w, col in ((16, "A"), (14, "B"), (14, "C"), (12, "D"), (10, "E")):
            s4.column_dimensions[col].width = w

    # ---------- Sheet 图表 ----------
    s5 = wb.create_sheet("图表")
    bar = BarChart(); bar.type = "col"; bar.title = "各区域销售额"; bar.y_axis.title = "金额"
    bar.add_data(Reference(s2, min_col=3, min_row=2, max_row=e1 - 1), titles_from_data=True)
    bar.set_categories(Reference(s2, min_col=1, min_row=3, max_row=e1 - 1))
    bar.height, bar.width = 7.5, 17
    s5.add_chart(bar, "B2")

    line = LineChart(); line.title = "月度销售趋势"; line.y_axis.title = "金额"
    mstart = e1 + 2
    line.add_data(Reference(s2, min_col=3, min_row=mstart + 1, max_row=e2 - 1), titles_from_data=True)
    line.set_categories(Reference(s2, min_col=1, min_row=mstart + 2, max_row=e2 - 1))
    line.height, line.width = 7.5, 17
    line.series[0].smooth = False
    s5.add_chart(line, "B17")

    s6 = wb.create_sheet("图表-产品占比")
    pie = PieChart(); pie.title = "产品销售额占比"
    pstart = e2 + 2
    pie.add_data(Reference(s2, min_col=3, min_row=pstart + 1, max_row=e3 - 1), titles_from_data=True)
    pie.set_categories(Reference(s2, min_col=1, min_row=pstart + 2, max_row=e3 - 1))
    pie.height, pie.width = 7.5, 13
    s6.add_chart(pie, "B2")

    wb.save(out_path)
    return {"tot_amount": tot_amt, "rows": good_rows, "dropped": stats.get("剔除行", 0),
            "dupes": stats.get("重复行", 0),
            "by_region": {k: round(v[1], 2) for k, v in by_region.items()}}


def load_baseline(path):
    if not path:
        return None
    if path.endswith(".json"):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    # 从旧 xlsx 读：汇总页区域块的值（公式未缓存时值为空 → 退回 json）
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
        s2 = wb["汇总"] if "汇总" in wb.sheetnames else wb.active
        vals = {}
        for r in s2.iter_rows(min_row=3, max_col=3):
            if isinstance(r[0].value, str) and isinstance(r[2].value, (int, float)):
                vals[r[0].value] = float(r[2].value)
        return {"by_region": vals, "tot_amount": 0, "rows": 0, "dropped": 0, "dupes": 0}
    except Exception:
        return None


def main():
    args = [a for a in sys.argv[1:]]
    if "--help" in args or "-h" in args or len(args) < 2:
        print(__doc__)
        raise SystemExit(0 if ("--help" in args or "-h" in args) else 1)
    src, dst = args[0], args[1]
    baseline_path = json_path = None
    title = "销售报表"
    i = 2
    while i < len(args):
        if args[i] == "--baseline" and i + 1 < len(args):
            baseline_path = args[i + 1]; i += 2
        elif args[i] == "--json" and i + 1 < len(args):
            json_path = args[i + 1]; i += 2
        elif args[i] == "--title" and i + 1 < len(args):
            title = args[i + 1]; i += 2
        else:
            i += 1

    header, raw, enc = load_rows(src)
    print("源文件 %s  编码 %s  原始行 %d  列 %s" % (os.path.basename(src), enc, len(raw), "/".join(header)))
    rows, stats = clean(header, raw)
    good, by_region, by_month, by_product = summarize(rows)

    print("\n== 清洗 ==")
    for k in ("源行", "空行", "重复行", "日期缺失", "日期异常", "金额符号修正", "金额缺失",
              "数量缺失", "数量非正", "区域归一", "有效行", "剔除行"):
        print("  %-10s %d" % (k, stats.get(k, 0)))

    tot_amt = round(sum(r["amount"] for r in good), 2)
    tot_qty = sum(r["qty"] for r in good)
    print("\n== 汇总 ==")
    print("  有效金额 %.2f  有效数量 %d  有效行 %d" % (tot_amt, tot_qty, len(good)))
    for k in sorted(by_region, key=lambda x: -by_region[x][1]):
        print("  %-8s %10.2f  %6.1f%%" % (k, by_region[k][1],
                                          (by_region[k][1] / tot_amt * 100) if tot_amt else 0))

    baseline = None
    if baseline_path:
        baseline = load_baseline(baseline_path)
        if baseline is None:
            print("\n[warn] baseline 读取失败，跳过差异页")
        else:
            print("\n== 差异 ==")
            print("  金额 %+.2f   行数 %+d" % (tot_amt - baseline.get("tot_amount", 0),
                                              len(good) - baseline.get("rows", 0)))

    res = write_xlsx(dst, rows, by_region, by_month, by_product, stats,
                     meta={"src": os.path.basename(src), "enc": enc, "title": title},
                     baseline=baseline)
    if json_path:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)
    print("\n已生成 %s  %.1f KB" % (dst, os.path.getsize(dst) / 1024))
    if json_path:
        print("汇总快照 %s" % json_path)


if __name__ == "__main__":
    main()
