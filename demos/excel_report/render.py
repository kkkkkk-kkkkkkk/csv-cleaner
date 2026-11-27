# -*- coding: utf-8 -*-
"""把 out_report.pdf 渲染成 PNG（挑图表页/汇总页/明细页）"""
import os, sys
import fitz

DIR = os.path.dirname(os.path.abspath(__file__))
pdf = os.path.join(DIR, "out_report.pdf")
doc = fitz.open(pdf)
print("PAGES", doc.page_count)
shots = []
for i in range(doc.page_count):
    pg = doc[i]
    txt = pg.get_text().strip()
    pix = pg.get_pixmap(dpi=150)
    name = os.path.join(DIR, "page_%02d.png" % (i + 1))
    pix.save(name)
    shots.append((i + 1, txt[:60].replace("\n", " "), os.path.getsize(name)))
for s in shots:
    print(s)
