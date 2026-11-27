#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批量跑：扫 fixtures 目录里的 PDF -> 抽取 -> 校验 -> 台账 xlsx + jsonl + 日志

一条命令：
    python run_batch.py
    python run_batch.py --dir fixtures --out out
"""
import os
import sys
import glob
import argparse
import logging

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import invoice_extract as ix


def main(argv=None):
    ap = argparse.ArgumentParser(description="发票 PDF 批量抽取与校验")
    ap.add_argument("--dir", default=os.path.join(HERE, "fixtures"))
    ap.add_argument("--out", default=os.path.join(HERE, "out"))
    a = ap.parse_args(argv)
    os.makedirs(a.out, exist_ok=True)
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.FileHandler(os.path.join(a.out, "run.log"), encoding="utf-8"),
                                  logging.StreamHandler()])

    paths = sorted(glob.glob(os.path.join(a.dir, "*.pdf")))
    if not paths:
        print("没有找到 PDF，先跑 python make_fixtures.py")
        return 1

    rows, issues, stats = ix.extract_batch(paths)
    ix.store_jsonl(os.path.join(a.out, "invoices.jsonl"), rows)
    xlsx = ix.store_xlsx(os.path.join(a.out, "invoices.xlsx"), rows, issues, stats)

    print()
    print("统计:", stats)
    print("台账:", xlsx)
    for r in rows:
        print("  %-10s %-9s %-11s %10.2f %10.2f %8s  %s" % (
            r["source"], r["invoice_no"] or "-", r["date"] or "-",
            r["amount"] or 0, r["total"] or 0, r["status"], r["issue_count"]))
    print()
    print("复核清单 %d 条：" % len(issues))
    for it in issues:
        print("  [%s] %s / %s / %s" % (it["level"], it["source"], it["field"], it["msg"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
