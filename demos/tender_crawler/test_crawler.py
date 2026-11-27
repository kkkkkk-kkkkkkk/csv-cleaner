# -*- coding: utf-8 -*-
"""单测：链接抽取 / 字段清洗 / 金额换算 / 去重 / 入库"""
import os
import sys
import sqlite3
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import crawler


def read(name):
    with open(os.path.join(HERE, "fixtures", name), encoding="utf-8") as f:
        return f.read()


class TestCrawler(unittest.TestCase):
    def setUp(self):
        crawler.FETCH = crawler.local_fetch(os.path.join(HERE, "fixtures"))

    def test_extract_links_absolute_and_dedup(self):
        links = crawler.extract_links(read("list.html"), "https://demo.example.gov.cn/list/1")
        self.assertEqual(len(links), 4)  # about.html 被过滤，重复 detail_1 被去重
        self.assertTrue(links[0].startswith("https://demo.example.gov.cn/tender/"))

    def test_parse_budget_units(self):
        r1 = crawler.parse_detail(read("detail_1.html"))
        self.assertEqual(r1["budget_yuan"], 1285000)
        r2 = crawler.parse_detail(read("detail_2.html"))
        self.assertEqual(r2["budget_yuan"], 460000)
        self.assertEqual(r1["buyer"], "某市大数据管理局")
        self.assertEqual(r1["notice_type"], "公开招标")

    def test_missing_budget_counted(self):
        r4 = crawler.parse_detail(read("detail_4.html"))
        self.assertIsNone(r4["budget_yuan"])

    def test_fingerprint_dedup(self):
        r1 = crawler.parse_detail(read("detail_1.html"))
        r3 = crawler.parse_detail(read("detail_3.html"))
        self.assertEqual(crawler.fingerprint(r1), crawler.fingerprint(r3))

    def test_full_run_stats_and_storage(self):
        records, stats, failures = crawler.crawl([("file:///list.html", read("list.html"))], delay=0)
        self.assertEqual(stats["details"], 4)
        self.assertEqual(stats["dup"], 1)       # detail_3 判重
        self.assertEqual(stats["missing"], 1)   # detail_4 无预算
        self.assertEqual(len(records), 3)
        self.assertEqual(stats["failed"], 0)

        d = tempfile.mkdtemp()
        db = os.path.join(d, "t.sqlite3")
        crawler.store_sqlite(db, records)
        conn = sqlite3.connect(db)
        rows = conn.execute("SELECT COUNT(*), SUM(budget_yuan) FROM tenders").fetchone()
        conn.close()
        self.assertEqual(rows[0], 3)
        self.assertEqual(rows[1], 1285000 + 460000)

    # ---- r149 新增：重试、失败入账、xlsx 导出 ----

    def test_retry_then_success(self):
        """前两次抛异常、第三次成功：结果不丢，重试被计数。"""
        calls = {"n": 0}

        def flaky(url):
            calls["n"] += 1
            if calls["n"] <= 2:
                raise TimeoutError("read timed out")
            return read("detail_1.html")

        stats = {}
        html, ok, err = crawler.fetch_with_retry(flaky, "x", retries=3, stats=stats,
                                                 sleeper=lambda s: None)
        self.assertTrue(ok)
        self.assertIsNone(err)
        self.assertEqual(stats["retries"], 2)
        self.assertIn("某市大数据管理局", crawler.parse_detail(html)["buyer"])

    def test_all_retries_fail_goes_to_failures(self):
        """三次都失败：不抛异常，进 failures 账，stats 计 errors，其他链接照抓。"""
        def always_fail(url):
            raise ConnectionResetError("connection reset by peer")

        records, stats, failures = crawler.crawl(
            [("file:///list.html", read("list.html"))], delay=0, retries=3,
            sleeper=lambda s: None, fetch=always_fail)
        self.assertEqual(len(records), 0)
        self.assertEqual(len(failures), 4)
        self.assertEqual(stats["errors"], 4)
        self.assertEqual(stats["retries"], 4 * 3)
        self.assertIn("connection reset", failures[0]["reason"])

    def test_xlsx_has_three_sheets_and_numbers(self):
        from openpyxl import load_workbook
        records, stats, failures = crawler.crawl([("file:///list.html", read("list.html"))], delay=0)
        d = tempfile.mkdtemp()
        p = os.path.join(d, "t.xlsx")
        crawler.store_xlsx(p, records, failures, stats)
        wb = load_workbook(p)
        self.assertEqual(wb.sheetnames, ["tenders", "failed", "fill_rate"])
        ws = wb["tenders"]
        self.assertEqual(ws.max_row, 1 + len(records))
        self.assertEqual(ws["A2"].value, 1)
        self.assertEqual(ws["D2"].value, 1285000)          # 预算元
        self.assertEqual(ws["E2"].value, 128.5)            # 预算万元
        self.assertEqual(ws["D2"].number_format, "#,##0")
        self.assertIsNotNone(ws.auto_filter.ref)
        self.assertEqual(wb["failed"]["A2"].value, "(本轮无失败链接)")
        self.assertGreater(wb["fill_rate"]["B2"].value, 0)

    def test_store_xlsx_lists_failures(self):
        from openpyxl import load_workbook
        d = tempfile.mkdtemp()
        p = os.path.join(d, "t2.xlsx")
        crawler.store_xlsx(p, [], [{"url": "https://a/tender/9", "reason": "TimeoutError: x"}], {})
        wb = load_workbook(p)
        self.assertEqual(wb["failed"]["A2"].value, "https://a/tender/9")
        self.assertEqual(wb["failed"]["B2"].value, "TimeoutError: x")


if __name__ == "__main__":
    unittest.main(verbosity=2)
