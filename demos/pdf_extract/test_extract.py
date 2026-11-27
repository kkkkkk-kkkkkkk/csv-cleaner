# -*- coding: utf-8 -*-
"""单测：金额解析 / 日期归一 / 单张抽取 / 合计校验 / 空票 / 缺号 / 重复号 / 整批台账

fixtures 由 make_fixtures.py 生成（5 张 PDF，含 1 张税额不符、1 张缺号且金额 0、1 张重复号）。
"""
import os
import sys
import glob
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import invoice_extract as ix

FIX = os.path.join(HERE, "fixtures")


def pdfs():
    return sorted(glob.glob(os.path.join(FIX, "*.pdf")))


class TestParse(unittest.TestCase):
    def test_amount_formats(self):
        self.assertEqual(ix.parse_amount("1,000.00"), 1000.0)
        self.assertEqual(ix.parse_amount("1，234.56"), 1234.56)   # 全角逗号
        self.assertEqual(ix.parse_amount("0.00"), 0.0)
        self.assertIsNone(ix.parse_amount(None))
        self.assertIsNone(ix.parse_amount("abc"))

    def test_date_norm_and_illegal(self):
        self.assertEqual(ix.normalize_date("2026", "11", "05"), ("2026-11-05", None))
        iso, err = ix.normalize_date("2026", "2", "31")
        self.assertIsNone(iso)
        self.assertIn("非法日期", err)


class TestExtractOne(unittest.TestCase):
    def test_fields_extracted(self):
        rec, probs = ix.extract_one(os.path.join(FIX, "inv_a.pdf"))
        self.assertEqual(rec["invoice_no"], "44201234")
        self.assertEqual(rec["date"], "2026-11-05")
        self.assertEqual(rec["buyer"], "阳江某贸易有限公司")
        self.assertEqual(rec["seller"], "广州某科技有限公司")
        self.assertEqual(rec["amount"], 1000.0)
        self.assertEqual(rec["tax"], 130.0)
        self.assertEqual(rec["total"], 1130.0)
        self.assertEqual(probs, [])

    def test_amount_not_confused_with_total(self):
        """「金额」正则必须排除「合计金额 / 价税合计」。"""
        self.assertEqual(ix.parse_amount("1,000.00"), 1000.0)
        text = "金额：¥1,000.00\n价税合计：¥1,130.00"
        m = ix.RE_AMOUNT.search(text)
        self.assertEqual(m.group(1), "1,000.00")

    def test_missing_file_is_recorded_not_raised(self):
        rec, probs = ix.extract_one(os.path.join(FIX, "nope.pdf"))
        self.assertEqual(rec["invoice_no"], None)
        self.assertEqual(probs[0][0], "EXTRACT_FAILED")


class TestValidate(unittest.TestCase):
    def test_total_mismatch_detected(self):
        rec, _ = ix.extract_one(os.path.join(FIX, "inv_c.pdf"))   # 800 + 104 ≠ 914
        codes = [i[0] for i in ix.validate(rec)]
        self.assertIn("TOTAL_MISMATCH", codes)
        self.assertEqual(rec["total"] - (rec["amount"] + rec["tax"]), 10.0)

    def test_zero_amount_and_missing_no(self):
        rec, _ = ix.extract_one(os.path.join(FIX, "inv_d.pdf"))
        codes = [i[0] for i in ix.validate(rec)]
        self.assertIn("MISSING_NO", codes)
        self.assertIn("ZERO_AMOUNT", codes)

    def test_duplicate_number_needs_shared_dict(self):
        """重复号检测依赖调用方传同一个 dict —— 空 dict 不能触发 `or {}` 被换掉。"""
        rec_a, _ = ix.extract_one(os.path.join(FIX, "inv_a.pdf"))
        rec_e, _ = ix.extract_one(os.path.join(FIX, "inv_e.pdf"))
        seen = {}
        self.assertEqual([i[0] for i in ix.validate(rec_a, seen)], [])
        self.assertIn("DUP_NO", [i[0] for i in ix.validate(rec_e, seen)])
        self.assertEqual(len(seen), 1)   # 同一个 dict 被写入，而不是被换新


class TestBatch(unittest.TestCase):
    def test_stats_and_flags(self):
        rows, issues, stats = ix.extract_batch(pdfs())
        self.assertEqual(stats["files"], 5)
        self.assertEqual(stats["rows"], 5)
        self.assertEqual(stats["flagged"], 3)      # c / d / e
        self.assertEqual(stats["passed"], 2)       # a / b
        self.assertEqual(stats["issues"], 4)
        self.assertEqual(stats["amount_sum"], 5300.0)
        self.assertEqual(stats["total_sum"], 5999.0)
        self.assertEqual([r["source"] for r in rows if r["status"] == "需复核"],
                         ["inv_c.pdf", "inv_d.pdf", "inv_e.pdf"])

    def test_xlsx_three_sheets(self):
        from openpyxl import load_workbook
        rows, issues, stats = ix.extract_batch(pdfs())
        d = tempfile.mkdtemp()
        p = os.path.join(d, "inv.xlsx")
        ix.store_xlsx(p, rows, issues, stats)
        wb = load_workbook(p)
        self.assertEqual(wb.sheetnames, ["发票台账", "复核清单", "统计"])
        ws = wb["发票台账"]
        self.assertEqual(ws.max_row, 6)
        self.assertEqual(ws["C2"].value, "44201234")
        self.assertEqual(ws["H2"].value, 1000.0)
        self.assertEqual(ws["H2"].number_format, "#,##0.00")
        self.assertIsNotNone(ws.auto_filter.ref)
        self.assertEqual(wb["复核清单"]["A2"].value, "inv_c.pdf")
        self.assertEqual(wb["统计"]["B2"].value, 5)   # 文件数（第 2 行起才是指标）


if __name__ == "__main__":
    unittest.main(verbosity=2)
