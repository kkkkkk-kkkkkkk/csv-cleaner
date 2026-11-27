# -*- coding: utf-8 -*-
"""生成 5 张样例发票 PDF（含 1 张税额不符、1 张缺号且金额为 0、1 张重复号）。
用 PyMuPDF 内置简体中文字体 china-s，不依赖系统字体。
"""
import os
import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
FIX = os.path.join(HERE, "fixtures")

SAMPLES = [
    dict(fn="inv_a.pdf", no="44201234", date=("2026", "11", "05"),
         buyer="阳江某贸易有限公司", seller="广州某科技有限公司",
         item="技术服务费（数据清洗）", amount="1,000.00", tax="130.00", total="1,130.00"),
    dict(fn="inv_b.pdf", no="44201250", date=("2026", "11", "12"),
         buyer="阳江某五金制品厂", seller="深圳某信息技术有限公司",
         item="软件开发服务", amount="2,500.00", tax="325.00", total="2,825.00"),
    dict(fn="inv_c.pdf", no="44201261", date=("2026", "11", "18"),
         buyer="某市文化传媒有限公司", seller="广州某广告有限公司",
         item="设计服务", amount="800.00", tax="104.00", total="914.00"),  # 合计 ≠ 金额+税额
    dict(fn="inv_d.pdf", no=None, date=("2026", "11", "20"),
         buyer="某县农业农村局", seller="阳江某供应链管理有限公司",
         item="咨询服务", amount="0.00", tax="0.00", total="0.00"),       # 缺号 + 金额为 0
    dict(fn="inv_e.pdf", no="44201234", date=("2026", "11", "22"),
         buyer="阳江某贸易有限公司", seller="广州某科技有限公司",
         item="技术服务费（数据清洗）", amount="1,000.00", tax="130.00", total="1,130.00"),  # 重复号
]


def build(s, path):
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=420)  # A5 横向
    y = 60
    lines = [
        ("增值税电子普通发票", 16),
        ("发票号码：%s" % (s["no"] or ""), 11),
        ("开票日期：%s年%s月%s日" % s["date"], 11),
        ("购买方名称：%s" % s["buyer"], 11),
        ("销售方名称：%s" % s["seller"], 11),
        ("项目名称：%s" % s["item"], 11),
        ("金额：¥%s" % s["amount"], 11),
        ("税额：¥%s" % s["tax"], 11),
        ("价税合计：¥%s" % s["total"], 11),
    ]
    for txt, size in lines:
        page.insert_text((50, y), txt, fontsize=size, fontname="china-s")
        y += 30 if size > 12 else 26
    doc.save(path)
    doc.close()


def main():
    os.makedirs(FIX, exist_ok=True)
    for s in SAMPLES:
        build(s, os.path.join(FIX, s["fn"]))
    print("fixtures:", sorted(os.listdir(FIX)))


if __name__ == "__main__":
    main()
