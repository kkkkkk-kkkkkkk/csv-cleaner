# -*- coding: utf-8 -*-
"""csv-cleaner: turn a messy CSV into a clean Excel workbook plus an issue report.

Design notes (deliberate, not accidental):
  * Column names are NOT hard-coded. A schema JSON maps the semantic fields
    (date / amount / qty / id / group / label / owner) to whatever the client's
    header row happens to say. Without a schema the tool still runs: it cleans
    whitespace, blank rows and duplicates, and reports what it could not parse.
  * Date ambiguity is never guessed. ISO (yyyy-mm-dd) is tried first, then the
    format named by --date-order (mdy default, dmy optional). Anything else is
    reported, not silently coerced.
  * Money parsing strips currency symbols, thousands separators and stray
    spaces, and records that it did so, so the client can see what changed.
  * Every dropped or flagged row keeps its source line number. The issue report
    is per-cell where possible.
  * --strict exits 1 when any row was dropped or any field failed to parse, so
    it can sit in a client's pipeline.

Usage:
    python clean.py messy.csv out.xlsx [--schema schema.json] [--date-order dmy] [--strict]
"""
import argparse
import collections
import csv
import datetime
import io
import json
import re
import sys

FIELDS = ("date", "amount", "qty", "id", "group", "label", "owner")
# Fields that are free text: the only safe normalisation is whitespace + case.
# Semantics ("Eastern" == "East"?) are NOT ours to guess.
TEXT_FIELDS = ("id", "group", "label", "owner")
NUMERIC_TEXT = re.compile(r"^[0-9][0-9.,+\-]*$")

# ISO first because it is unambiguous. The slash/dot forms follow the
# caller-declared order; we never infer which of 03/04 is March or April.
ISO_FMTS = ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d")
MDY_FMTS = ("%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%y")
DMY_FMTS = ("%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%y")

MONEY_STRIP = re.compile(r"[^0-9.,\-]")


def sniff_text(path):
    """Return (text, encoding). BOM and GBK are both common in exported files."""
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk", "latin-1"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return f.read(), enc
        except (UnicodeDecodeError, LookupError):
            continue
    raise RuntimeError("cannot decode %s" % path)


def read_csv_text(text):
    rd = csv.reader(io.StringIO(text))
    try:
        header = [h.strip() for h in next(rd)]
    except StopIteration:
        raise RuntimeError("empty file")
    rows = []
    for r in rd:
        if len(r) < len(header):
            r = list(r) + [""] * (len(header) - len(r))
        rows.append(r[: len(header)])
    return header, rows


def load_schema(path):
    if not path:
        return {}
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    # Accept both {"date": "Order Date"} and {"date": ["Order Date", "date"]}
    schema = {}
    for k in FIELDS:
        if k in raw:
            v = raw[k]
            schema[k] = list(v) if isinstance(v, (list, tuple)) else [v]
    unknown = [k for k in raw if k not in FIELDS]
    if unknown:
        raise RuntimeError("schema has unknown keys: %s (allowed: %s)"
                           % (", ".join(unknown), ", ".join(FIELDS)))
    return schema


def resolve_columns(header, schema):
    """Map semantic field -> header index. Aliases are tried in order.

    Accepts a bare string as well as a list, so callers cannot accidentally
    iterate over the characters of the column name.
    """
    idx = {}
    for field, aliases in (schema or {}).items():
        if isinstance(aliases, str):
            aliases = [aliases]
        for a in aliases:
            if a in header:
                idx[field] = header.index(a)
                break
    return idx


def parse_date(value, date_order="mdy"):
    v = (value or "").strip()
    if not v:
        return None, "empty"
    # fast path: a real ISO-ish string
    for fmt in ISO_FMTS:
        try:
            return datetime.datetime.strptime(v, fmt).date(), ""
        except ValueError:
            pass
    for fmt in (DMY_FMTS if date_order == "dmy" else MDY_FMTS):
        try:
            return datetime.datetime.strptime(v, fmt).date(), ""
        except ValueError:
            pass
    return None, "unparsed"


def parse_money(value):
    """Return (amount, changed). '¥12,500.00 ' -> (12500.0, True).

    1.234,56 (European) and 1,234.56 (US) both work; the rule is positional,
    not a guess: the rightmost separator is the decimal one.
    """
    v = (value or "").strip()
    if not v:
        return None, False
    cleaned = MONEY_STRIP.sub("", v)
    has_dot, has_comma = "." in cleaned, "," in cleaned
    if has_dot and has_comma:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif has_comma:
        if cleaned.count(",") == 1 and len(cleaned.split(",")[-1]) in (1, 2):
            cleaned = cleaned.replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    if cleaned in ("", "-", "."):
        return None, False
    try:
        return float(cleaned), cleaned != v
    except ValueError:
        return None, False


def parse_qty(value):
    v = (value or "").strip()
    if not v:
        return None, "empty"
    try:
        f = float(MONEY_STRIP.sub("", v))
    except ValueError:
        return None, "unparsed"
    if f != int(f):
        return None, "not-integer"
    return int(f), ""


def writing_rank(v):
    """How much a spelling looks like something a human meant to type.

    When 'north', 'North' and 'NORTH' each appear once we still have to pick one
    spelling to write. Preferring mixed case ('North', 'iPhone') over ALL CAPS
    ('USA') over all lower case ('north') is a convention, not knowledge: it
    never invents a value, it only decides which existing one to reuse.
    """
    if v.isupper():
        return 1
    if any(ch.isupper() for ch in v):
        return 2
    return 0


def add_flag(row, note):
    row["flags"] = "; ".join([x for x in (row["flags"], note) if x])


def normalise_text(rows, header, col, stats, issues):
    """Collapse case variants of the same value; flag (never auto-merge) truncation.

    What is safe: "North" and "North  " and "north" are the same value written
    differently, so we pick one spelling and rewrite the rest. What is NOT safe:
    "K" might be an abbreviaton of "Kim", or it might be a different person, and
    "East" might be "Eastern". Those are reported, left untouched, and the
    client decides. Guessing here would silently corrupt their reporting.
    """
    for field in TEXT_FIELDS:
        if field not in col:
            continue
        colname = header[col[field]]
        groups = collections.defaultdict(collections.Counter)
        first_line = {}
        for r in rows:
            v = r[field]
            if not v:
                continue
            groups[v.casefold()][v] += 1
            first_line.setdefault(v, r["line"])

        canon = {}
        renamed = 0
        for cf, counter in groups.items():
            if len(counter) == 1:
                canon[cf] = next(iter(counter))
                continue
            # most frequent spelling wins; ties go to the most "human-looking"
            # spelling, then to whichever appeared first.
            best = sorted(counter.items(),
                          key=lambda kv: (-kv[1], -writing_rank(kv[0]), first_line[kv[0]]))[0][0]
            canon[cf] = best
            variants = ", ".join("%r x%d" % (k, v) for k, v in counter.items())
            stats["text_case_variants"] += 1
            issues.append({"line": "-", "column": colname,
                           "problem": "same value in different case, collapsed to %r" % best,
                           "value": variants})

        for r in rows:
            v = r[field]
            if not v:
                continue
            c = canon.get(v.casefold())
            if c and c != v:
                r[field] = c
                renamed += 1
                add_flag(r, "%s case-normalised to %s" % (field, c))
                issues.append({"line": r["line"], "column": colname,
                               "problem": "case variant rewritten", "value": v})
        stats["text_case_rows"] += renamed

    # Truncation / near-duplicate suspects: report only, never rewrite.
    for field in TEXT_FIELDS:
        if field not in col:
            continue
        colname = header[col[field]]
        distinct = sorted({r[field] for r in rows if r[field]})
        cf_map = {}
        for v in distinct:
            cf_map.setdefault(v.casefold(), []).append(v)
        for a in distinct:
            for b in distinct:
                if a == b or len(b) - len(a) < 2:
                    continue
                if b.casefold().startswith(a.casefold()) and a.casefold() != b.casefold():
                    stats["text_truncation_suspects"] += 1
                    issues.append({"line": "-", "column": colname,
                                   "problem": "possible truncation or abbreviation, left as-is",
                                   "value": "%r vs %r" % (a, b)})
                    break

    # Numbers sitting in a text column usually mean a shifted row.
    for field in ("group", "label", "owner"):
        if field not in col:
            continue
        for r in rows:
            v = r[field]
            if v and NUMERIC_TEXT.match(v):
                stats["numeric_in_text"] += 1
                add_flag(r, "%s looks numeric" % field)
                issues.append({"line": r["line"], "column": header[col[field]],
                               "problem": "numeric value in a text column (row may be shifted)",
                               "value": v})


def clean_rows(header, rows, schema, date_order="mdy"):
    """Pure function: no IO, so it is cheap to test.

    Returns (clean_rows, stats, issues). Each clean row is a dict with the
    normalised values plus `line` (1-based source line, header is line 1) and
    `flags` (human-readable notes). `dropped` marks rows that are excluded from
    aggregates.
    """
    col = resolve_columns(header, schema)
    stats = collections.Counter()
    issues = []
    out = []

    for i, r in enumerate(rows, start=2):
        cells = [(c or "").strip() for c in r]
        if not any(cells):
            stats["blank_rows"] += 1
            continue

        def cell(field):
            j = col.get(field)
            return cells[j] if j is not None and j < len(cells) else ""

        flags = []
        dropped = False

        dt = None
        if "date" in col:
            dt, why = parse_date(cell("date"), date_order)
            if why == "unparsed":
                flags.append("date unparsed")
                stats["date_unparsed"] += 1
                issues.append({"line": i, "column": header[col["date"]],
                               "problem": "date unparsed", "value": cell("date")})
            elif why == "empty":
                stats["date_missing"] += 1

        amount = None
        if "amount" in col:
            amount, changed = parse_money(cell("amount"))
            if changed:
                flags.append("amount symbols stripped")
                stats["amount_normalised"] += 1
                issues.append({"line": i, "column": header[col["amount"]],
                               "problem": "amount had currency/thousand separators",
                               "value": cell("amount")})
            if amount is None:
                flags.append("amount missing, counted as 0")
                stats["amount_missing"] += 1
                issues.append({"line": i, "column": header[col["amount"]],
                               "problem": "amount missing", "value": cell("amount")})
                amount = 0.0
            elif amount < 0:
                flags.append("negative amount kept")
                stats["amount_negative"] += 1

        qty = None
        if "qty" in col:
            qty, why = parse_qty(cell("qty"))
            if why == "empty":
                stats["qty_missing"] += 1
            elif why:
                flags.append("qty %s" % why)
                stats["qty_unparsed"] += 1
                issues.append({"line": i, "column": header[col["qty"]],
                               "problem": "qty %s" % why, "value": cell("qty")})
            elif qty <= 0:
                flags.append("qty non-positive, row dropped")
                stats["qty_non_positive"] += 1
                issues.append({"line": i, "column": header[col["qty"]],
                               "problem": "qty non-positive", "value": cell("qty")})
                dropped = True

        out.append({
            "line": i,
            "date": dt,
            "amount": amount,
            "qty": qty,
            "id": cell("id"),
            "group": cell("group") or "unlabelled",
            "label": cell("label"),
            "owner": cell("owner"),
            "flags": "; ".join(flags),
            "dropped": dropped,
        })

    # Text normalisation runs before duplicate detection on purpose: if "North"
    # and "north" are two spellings of one value, two otherwise identical rows
    # must collapse into one duplicate, not survive as a fake second record.
    normalise_text(out, header, col, stats, issues)

    # Duplicates: compare on the normalised business values, never on the line
    # number, otherwise nothing is ever a duplicate.
    seen = set()
    kept = []
    for row in out:
        key = (row["id"], row["date"], row["group"], row["label"],
               row["qty"], None if row["amount"] is None else round(row["amount"], 2),
               row["owner"])
        if key in seen:
            stats["duplicate_rows"] += 1
            issues.append({"line": row["line"], "column": "-",
                           "problem": "duplicate row dropped",
                           "value": row["id"] or row["label"]})
            continue
        seen.add(key)
        kept.append(row)

    stats["rows_in"] = len(rows)
    stats["rows_kept"] = len([r for r in kept if not r["dropped"]])
    stats["rows_dropped"] = len([r for r in kept if r["dropped"]])
    # stats already carries the counter() additions; merge the final tallies
    return kept, stats, issues


def aggregate(rows):
    good = [r for r in rows if not r["dropped"]]
    by_group = collections.defaultdict(lambda: [0, 0.0])
    by_month = collections.defaultdict(lambda: [0, 0.0])
    by_label = collections.defaultdict(lambda: [0, 0.0])
    for r in good:
        g = r["group"]
        by_group[g][0] += r["qty"] or 0
        by_group[g][1] += r["amount"] or 0.0
        m = r["date"].strftime("%Y-%m") if r["date"] else "unknown date"
        by_month[m][0] += r["qty"] or 0
        by_month[m][1] += r["amount"] or 0.0
        lab = r["label"] or "unknown"
        by_label[lab][0] += r["qty"] or 0
        by_label[lab][1] += r["amount"] or 0.0
    return good, by_group, by_month, by_label


def write_xlsx(path, header, schema, rows, issues, by_group, by_month, by_label):
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=10)
    thin = Side(style="thin", color="D9D9D9")

    def style_head(ws, ncol, row=1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill, cell.font = head_fill, head_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb = Workbook()
    ws = wb.active
    ws.title = "cleaned"
    fields = [k for k in FIELDS if k in schema]
    out_head = ["source_line"] + fields + ["flags", "dropped"]
    ws.append(out_head)
    for r in rows:
        row = [r["line"]]
        for k in fields:
            v = r[k]
            if k == "date":
                v = v.strftime("%Y-%m-%d") if v else ""
            elif k == "amount" and isinstance(v, (int, float)):
                v = round(v, 2)
            row.append("" if v is None else v)
        row += [r["flags"], "yes" if r["dropped"] else ""]
        ws.append(row)
    style_head(ws, len(out_head))
    flag_col = len(out_head) - 1
    flag_wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(out_head)):
        for c in row:
            c.font = body_font
            c.border = Border(bottom=thin)
            if c.column == flag_col:
                c.alignment = flag_wrap
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(out_head)), max(1, ws.max_row))
    for i, w in enumerate([12] + [16] * len(fields) + [52, 9], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    s2 = wb.create_sheet("summary")
    blocks = [("By %s" % (schema.get("group", ["group"])[0]), by_group),
              ("By month", by_month),
              ("By %s" % (schema.get("label", ["label"])[0]), by_label)]
    cursor = 1
    anchors = {}
    for title, data in blocks:
        s2.cell(row=cursor, column=1, value=title).font = Font(name="Calibri", size=12, bold=True)
        hdr = cursor + 1
        for i, v in enumerate(["key", "qty", "amount"], start=1):
            s2.cell(row=hdr, column=i, value=v)
        style_head(s2, 3, hdr)
        items = sorted(data.items(), key=lambda kv: -kv[1][1])
        for j, (k, (q, a)) in enumerate(items, start=1):
            s2.cell(row=hdr + j, column=1, value=k).font = body_font
            s2.cell(row=hdr + j, column=2, value=q).font = body_font
            c = s2.cell(row=hdr + j, column=3, value=round(a, 2))
            c.font = body_font
            c.number_format = "#,##0.00"
        anchors[title] = (hdr, hdr + len(items))
        cursor = hdr + len(items) + 2
    for col, w in (("A", 24), ("B", 10), ("C", 16)):
        s2.column_dimensions[col].width = w

    s3 = wb.create_sheet("charts")
    (g_hdr, g_end) = anchors[blocks[0][0]]
    bar = BarChart()
    bar.type = "col"
    bar.title = blocks[0][0]
    bar.y_axis.title = "amount"
    bar.add_data(Reference(s2, min_col=3, min_row=g_hdr, max_row=g_end), titles_from_data=True)
    bar.set_categories(Reference(s2, min_col=1, min_row=g_hdr + 1, max_row=g_end))
    bar.height, bar.width = 7.5, 17
    s3.add_chart(bar, "B2")

    (m_hdr, m_end) = anchors[blocks[1][0]]
    line = LineChart()
    line.title = blocks[1][0]
    line.add_data(Reference(s2, min_col=3, min_row=m_hdr, max_row=m_end), titles_from_data=True)
    line.set_categories(Reference(s2, min_col=1, min_row=m_hdr + 1, max_row=m_end))
    if line.series:
        line.series[0].smooth = False
    line.height, line.width = 7.5, 17
    s3.add_chart(line, "B17")

    s4 = wb.create_sheet("share")
    (l_hdr, l_end) = anchors[blocks[2][0]]
    pie = PieChart()
    pie.title = blocks[2][0]
    pie.add_data(Reference(s2, min_col=3, min_row=l_hdr, max_row=l_end), titles_from_data=True)
    pie.set_categories(Reference(s2, min_col=1, min_row=l_hdr + 1, max_row=l_end))
    pie.height, pie.width = 7.5, 13
    s4.add_chart(pie, "B2")

    s5 = wb.create_sheet("issues")
    s5.append(["source_line", "column", "problem", "value"])
    for it in issues:
        s5.append([it["line"], it["column"], it["problem"], str(it["value"])[:80]])
    style_head(s5, 4)
    for row in s5.iter_rows(min_row=2, max_row=s5.max_row, max_col=4):
        for c in row:
            c.font = body_font
    for col, w in (("A", 12), ("B", 22), ("C", 58), ("D", 30)):
        s5.column_dimensions[col].width = w

    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Clean a messy CSV into an Excel workbook + issue report.")
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--schema", default=None, help="JSON mapping semantic fields to your header names")
    ap.add_argument("--date-order", default="mdy", choices=["mdy", "dmy"],
                    help="interpretation for 03/04/2026 style dates (ISO is always tried first)")
    ap.add_argument("--strict", action="store_true", help="exit 1 if any row dropped or field unparsed")
    a = ap.parse_args(argv)

    text, enc = sniff_text(a.src)
    header, rows = read_csv_text(text)
    schema = load_schema(a.schema)
    unknown_fields = [f for f in FIELDS if f not in schema]
    rows, stats, issues = clean_rows(header, rows, schema, a.date_order)
    good, by_group, by_month, by_label = aggregate(rows)

    print("encoding: %s" % enc)
    print("columns:  %s" % ", ".join(header))
    print("mapped:   %s" % (", ".join("%s->%s" % (k, v) for k, v in resolve_columns(header, schema).items()) or "none"))
    print("")
    print("rows in     %d" % stats["rows_in"])
    print("rows kept   %d" % stats["rows_kept"])
    print("rows dropped %d" % stats["rows_dropped"])
    for k in ("blank_rows", "duplicate_rows", "date_unparsed", "date_missing",
              "amount_normalised", "amount_missing", "amount_negative",
              "qty_missing", "qty_unparsed", "qty_non_positive",
              "text_case_variants", "text_case_rows", "text_truncation_suspects",
              "numeric_in_text"):
        if stats.get(k):
            print("  %-20s %d" % (k, stats[k]))
    if unknown_fields and not schema:
        print("note: no schema supplied, only structural cleaning was possible")

    write_xlsx(a.dst, header, schema, rows, issues, by_group, by_month, by_label)
    print("")
    print("written: %s" % a.dst)

    if a.strict and (stats["rows_dropped"] or stats["date_unparsed"]
                     or stats["qty_unparsed"] or stats["amount_missing"]):
        print("strict: problems found, exit 1")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
