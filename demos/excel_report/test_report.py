# -*- coding: utf-8 -*-
"""report.py 单元测试（弹药包 #4）"""
import os, sys, io, json, csv, tempfile, unittest
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import report


HEADER = ["订单号", "下单日期", "区域", "产品", "数量", "金额", "业务员"]


def rows_of(data):
    return [dict(zip(HEADER, r)) for r in data]


def build(raw_rows):
    header = HEADER
    raw = [[r.get(h, "") for h in header] for r in raw_rows]
    return report.clean(header, raw)


class TestParsing(unittest.TestCase):
    def test_money_strips_symbols(self):
        self.assertEqual(report.parse_money("￥1,234.50")[0], 1234.50)
        self.assertEqual(report.parse_money("1，234.50")[0], 1234.50)   # 全角逗号按千分位
        self.assertEqual(report.parse_money("1234，50")[0], 1234.5)    # 两位小数 → 逗号按小数点
        self.assertEqual(report.parse_money(" 88 元 ")[0], 88.0)

    def test_money_flag_only_when_dirty(self):
        self.assertFalse(report.parse_money("88")[1])
        self.assertTrue(report.parse_money("￥88")[1])

    def test_money_bad(self):
        self.assertIsNone(report.parse_money("待定")[0])
        self.assertIsNone(report.parse_money("")[0])

    def test_date_formats(self):
        for s in ("2026-03-05", "2026/03/05", "05/03/2026", "2026.03.05"):
            d, st = report.parse_date(s)
            self.assertEqual(st, "ok", s)
            self.assertEqual((d.month, d.day), (3, 5), s)

    def test_date_states(self):
        self.assertEqual(report.parse_date("")[1], "missing")
        self.assertEqual(report.parse_date("2026-02-31")[1], "bad")   # 不猜，不纠正
        self.assertEqual(report.parse_date("2026年3月5日")[1], "bad")

    def test_region_alias(self):
        self.assertEqual(report.norm_region("华东区"), "华东")
        self.assertEqual(report.norm_region("SW"), "西南")
        self.assertEqual(report.norm_region("North"), "华北")
        self.assertEqual(report.norm_region(""), "未标注")


class TestClean(unittest.TestCase):
    def test_empty_and_dup(self):
        data = [
            {"订单号": "A1", "下单日期": "2026-01-01", "区域": "华东", "产品": "P1", "数量": "2", "金额": "100", "业务员": "张三"},
            {"订单号": "A1", "下单日期": "2026-01-01", "区域": "华东", "产品": "P1", "数量": "2", "金额": "100", "业务员": "张三"},
            {"订单号": "", "下单日期": "", "区域": "", "产品": "", "数量": "", "金额": "", "业务员": ""},
        ]
        rows, stats = build(data)
        self.assertEqual(stats["重复行"], 1)
        self.assertEqual(stats["空行"], 1)
        self.assertEqual(len(rows), 1)  # 重复行剔除 + 全空行按空行剔除，剩 1 条
        self.assertEqual(stats["有效行"], 1)

    def test_missing_amount_not_counted(self):
        """金额缺失必须剔除，不能当 0 计入合计（否则总额被拉低）"""
        data = [
            {"订单号": "A1", "下单日期": "2026-01-01", "区域": "华东", "产品": "P1", "数量": "2", "金额": "100", "业务员": "张三"},
            {"订单号": "A2", "下单日期": "2026-01-02", "区域": "华东", "产品": "P1", "数量": "3", "金额": "", "业务员": "张三"},
            {"订单号": "A3", "下单日期": "2026-01-03", "区域": "华东", "产品": "P1", "数量": "4", "金额": "待定", "业务员": "张三"},
        ]
        rows, stats = build(data)
        self.assertEqual(stats["金额缺失"], 2)
        good, _, _, _ = report.summarize(rows)
        self.assertEqual(sum(r["amount"] for r in good), 100.0)  # 只有 100，不是 100+0+0 的"看起来 100"
        self.assertEqual(len(good), 1)
        self.assertTrue(all(r["drop"] for r in rows if r["amount"] == 0.0))

    def test_nonpositive_qty_dropped(self):
        data = [
            {"订单号": "A1", "下单日期": "2026-01-01", "区域": "华东", "产品": "P1", "数量": "0", "金额": "50", "业务员": "李四"},
            {"订单号": "A2", "下单日期": "2026-01-01", "区域": "华东", "产品": "P1", "数量": "-2", "金额": "60", "业务员": "李四"},
            {"订单号": "A3", "下单日期": "2026-01-01", "区域": "华东", "产品": "P1", "数量": "1", "金额": "70", "业务员": "李四"},
        ]
        rows, stats = build(data)
        self.assertEqual(stats["数量非正"], 2)
        good, *_ = report.summarize(rows)
        self.assertEqual(len(good), 1)
        self.assertEqual(sum(r["amount"] for r in good), 70.0)

    def test_qty_missing_kept_but_flagged(self):
        data = [{"订单号": "A1", "下单日期": "2026-01-01", "区域": "华东", "产品": "P1", "数量": "", "金额": "90", "业务员": "王五"}]
        rows, stats = build(data)
        self.assertEqual(stats["数量缺失"], 1)
        self.assertFalse(rows[0]["drop"])
        self.assertIn("数量缺失", rows[0]["flag"])

    def test_unknown_date_goes_to_bucket(self):
        data = [{"订单号": "A1", "下单日期": "", "区域": "华东", "产品": "P1", "数量": "1", "金额": "10", "业务员": "王五"}]
        rows, stats = build(data)
        good, by_region, by_month, by_product = report.summarize(rows)
        self.assertIn("日期未知", by_month)
        self.assertEqual(stats["日期缺失"], 1)

    def test_region_normalized_and_flagged(self):
        data = [{"订单号": "A1", "下单日期": "2026-01-01", "区域": "华东区", "产品": "P1", "数量": "1", "金额": "10", "业务员": "王五"}]
        rows, stats = build(data)
        self.assertEqual(rows[0]["region"], "华东")
        self.assertEqual(stats["区域归一"], 1)
        self.assertIn("区域已归一", rows[0]["flag"])

    def test_three_way_totals_agree(self):
        base = rows_of([
            ("A1", "2026-01-01", "华东", "P1", "2", "100", "张三"),
            ("A2", "2026-01-02", "华东区", "P2", "3", "150", "张三"),
            ("A3", "2026-02-01", "华北", "P1", "1", "80", "李四"),
            ("A4", "2026-02-05", "North", "P3", "4", "200", "李四"),
        ])
        rows, stats = build(base)
        good, by_region, by_month, by_product = report.summarize(rows)
        tot = sum(r["amount"] for r in good)
        self.assertEqual(sum(v[1] for v in by_region.values()), tot)
        self.assertEqual(sum(v[1] for v in by_month.values()), tot)
        self.assertEqual(sum(v[1] for v in by_product.values()), tot)
        self.assertEqual(sum(v[0] for v in by_region.values()), sum(r["qty"] for r in good))


class TestWorkbook(unittest.TestCase):
    def test_sheets_and_formulas(self):
        from openpyxl import load_workbook
        base = rows_of([
            ("A1", "2026-01-01", "华东", "P1", "2", "￥100", "张三"),
            ("A2", "2026-01-02", "华东区", "P2", "3", "150", "张三"),
            ("A3", "2026-02-01", "华北", "P1", "1", "80", "李四"),
        ])
        rows, stats = build(base)
        good, br, bm, bp = report.summarize(rows)
        tmpdir = tempfile.mkdtemp()
        out = os.path.join(tmpdir, "t.xlsx")
        res = report.write_xlsx(out, rows, br, bm, bp, stats, meta={"src": "x.csv", "enc": "utf-8", "title": "T"})
        wb = load_workbook(out)
        self.assertEqual(wb.sheetnames[:4], ["口径", "明细", "汇总", "核对"])
        self.assertIn("图表", wb.sheetnames)
        s2 = wb["汇总"]
        self.assertTrue(str(s2["B3"].value).startswith("=SUMIFS("))
        self.assertIn("汇总!C", str(wb["核对"]["C4"].value))
        self.assertTrue(str(wb["核对"]["E4"].value).startswith("=IF(ABS("))
        self.assertEqual(res["tot_amount"], 330.0)
        self.assertEqual(res["rows"], 3)

    def test_diff_sheet(self):
        from openpyxl import load_workbook
        base = rows_of([
            ("A1", "2026-01-01", "华东", "P1", "2", "100", "张三"),
            ("A2", "2026-01-02", "华北", "P2", "3", "200", "张三"),
        ])
        rows, stats = build(base)
        good, br, bm, bp = report.summarize(rows)
        out = os.path.join(tempfile.mkdtemp(), "t2.xlsx")
        report.write_xlsx(out, rows, br, bm, bp, stats, meta={"src": "x.csv", "enc": "utf-8", "title": "T"},
                          baseline={"tot_amount": 100.0, "rows": 1, "dropped": 0, "dupes": 0,
                                    "by_region": {"华东": 100.0}})
        wb = load_workbook(out)
        self.assertIn("差异", wb.sheetnames)
        s4 = wb["差异"]
        self.assertEqual(s4["B4"].value, 100.0)
        self.assertEqual(s4["C4"].value, 300.0)
        self.assertEqual(str(s4["D4"].value), "=C4-B4")
        labels = [s4.cell(row=r, column=1).value for r in range(1, 20)]
        self.assertIn("华北", labels)


class TestCli(unittest.TestCase):
    def test_end_to_end(self):
        tmpdir = tempfile.mkdtemp()
        src = os.path.join(tmpdir, "in.csv")
        with open(src, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(HEADER)
            w.writerow(["A1", "2026-01-01", "华东", "P1", "2", "￥1,000", "张三"])
            w.writerow(["A2", "2026-01-02", "华东区", "P2", "3", "", "张三"])
            w.writerow(["A3", "2026/01/03", "华南", "P1", "5", "500", "李四"])
        out = os.path.join(tmpdir, "out.xlsx")
        js = os.path.join(tmpdir, "out.json")
        argv = sys.argv
        sys.argv = ["report.py", src, out, "--json", js, "--title", "单测"]
        try:
            report.main()
        finally:
            sys.argv = argv
        self.assertTrue(os.path.exists(out))
        with open(js, encoding="utf-8") as f:
            snap = json.load(f)
        self.assertEqual(snap["tot_amount"], 1500.0)
        self.assertEqual(snap["rows"], 2)
        self.assertEqual(snap["dropped"], 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
