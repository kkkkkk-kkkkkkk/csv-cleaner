# -*- coding: utf-8 -*-
"""生成一份"脏"的销售明细 CSV，用来演示清洗能力。
故意埋的坑：日期 3 种写法、金额带￥和千分位逗号、区域名带空格/大小写/别名、
                  产品名带空格、完全重复行、空行、金额缺失、数量为 0 或负数。
"""
import csv, random, io, os

random.seed(7)
DIR = os.path.dirname(os.path.abspath(__file__))

REGIONS = ["华东", "华南", "华北", "西南", "华东 ", " 华南", "华北区", "华东区", "SW", "North"]
PRODUCTS = ["智能手环", "蓝牙耳机", "移动电源", "车载支架", "无线充电器", " 蓝牙耳机 ", "移动电源 "]
DATES = ["2026-01-05", "2026/1/5", "5/1/2026", "2026-02-11", "2026/2/11", "11/2/2026",
         "2026-03-02", "2026-03-18", "2026/3/18", "2026-04-07", "7/4/2026",
         "2026-05-09", "2026/5/9", "9/5/2026", "2026-06-15", "2026/6/15", "15/6/2026"]
PRICE = {"智能手环": 199, "蓝牙耳机": 259, "移动电源": 129, "车载支架": 69, "无线充电器": 99}

rows = []
for i in range(260):
    p = random.choice(PRODUCTS)
    base = PRICE[p.strip()]
    qty = random.choice([1, 1, 2, 3, 5, 8, -1, 0])          # 有 0 和负数脏值
    amt = base * qty
    if random.random() < 0.12:
        amt_s = ""                                          # 缺失金额
    elif random.random() < 0.5:
        amt_s = "￥{:,.2f}".format(amt)                      # 带符号千分位
    elif random.random() < 0.15:
        amt_s = " {:,.0f} 元 ".format(amt)                   # 带中文单位
    else:
        amt_s = str(amt)
    rows.append([str(1000 + i), random.choice(DATES), random.choice(REGIONS),
                 p, str(qty) if random.random() > 0.05 else "", amt_s, random.choice(["张三", "李四", "王五", " 赵六 "])])

# 完全重复 8 条
for r in random.sample(rows, 8):
    rows.append(list(r))
# 空行 3 条
rows.extend([["", "", "", "", "", "", ""]] * 3)

buf = io.StringIO()
w = csv.writer(buf)
w.writerow(["订单号", "下单日期", "区域", "产品", "数量", "金额", "业务员"])
w.writerows(rows)
path = os.path.join(DIR, "sample_sales.csv")
with open(path, "w", encoding="utf-8-sig", newline="") as f:
    f.write(buf.getvalue())
print("WROTE", path, "rows", len(rows))
