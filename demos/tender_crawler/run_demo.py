#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""离线演示：列表 -> 详情 -> 清洗 -> 去重 -> 入库 -> 导出 xlsx -> 日志

一条命令重跑：
    python run_demo.py
    python run_demo.py --retries 4 --xlsx out\tenders.xlsx
"""
import os
import sys
import argparse
import logging

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import crawler

OUT = os.path.join(HERE, "out")


def main(argv=None):
    ap = argparse.ArgumentParser(description="招投标公开信息采集演示（离线 fixture）")
    ap.add_argument("--fixtures", default=os.path.join(HERE, "fixtures"))
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--retries", type=int, default=3)
    ap.add_argument("--xlsx", default=None)
    a = ap.parse_args(argv)

    os.makedirs(a.out, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(os.path.join(a.out, "run.log"), encoding="utf-8"),
                  logging.StreamHandler()],
    )
    crawler.LOG = logging.getLogger("tender")
    crawler.FETCH = crawler.local_fetch(a.fixtures)

    with open(os.path.join(a.fixtures, "list.html"), encoding="utf-8") as fh:
        list_html = fh.read()

    records, stats, failures = crawler.crawl([("file:///list.html", list_html)],
                                            delay=0, retries=a.retries)
    crawler.store_jsonl(os.path.join(a.out, "tenders.jsonl"), records)
    crawler.store_sqlite(os.path.join(a.out, "tenders.sqlite3"), records)
    xlsx = crawler.store_xlsx(a.xlsx or os.path.join(a.out, "tenders.xlsx"),
                              records, failures, stats)

    print()
    print("抓取统计:", {k: v for k, v in stats.items() if k != "fill_rate"})
    print("入库条数:", len(records), "| 失败链接:", len(failures), "| xlsx:", xlsx)
    print("字段填充率:", stats["fill_rate"])
    for r in records:
        print("  - %-8s %-28s %-10s %s" % (r["notice_type"], (r["title"] or "")[:26],
                                           r["budget_yuan"], (r["deadline"] or "")[:16]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
