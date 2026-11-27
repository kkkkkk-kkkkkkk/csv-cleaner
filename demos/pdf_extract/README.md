# 发票 / 单据 PDF 批量抽取与校验（演示版）

针对"一堆 PDF 发票或单据要录进台账、还要保证数字对得上"这类需求的可运行原型。

## 跑起来

```bash
python make_fixtures.py                 # 生成 5 张样例发票 PDF（含 3 种典型问题）
python -m unittest test_extract -v      # 10 个用例
python run_batch.py                     # 批量抽取 -> out/invoices.jsonl + out/invoices.xlsx
```

## 组成

- `invoice_extract.py`
  - `extract_one()`：PyMuPDF 取文本层 → 正则抽 8 个字段（号码/日期/购买方/销售方/项目/金额/税额/价税合计）；读不出文件不抛异常，返回 `EXTRACT_FAILED`
  - `parse_amount()`：自动处理 `1,000.00`、`1，234.56`（全角逗号）、`0.00`
  - `normalize_date()`：`2026年11月05日 → 2026-11-05`；非法日期（如 2 月 31 日）不猜、不填 0，直接进复核清单
  - `validate()`：五类业务校验 —— 缺号码、**号码重复**、日期异常、合计不平（`价税合计 ≠ 金额 + 税额`）、空票（金额与税额均为 0）
  - `extract_batch()`：整批跑，返回 `rows / issues / stats`（含通过数、需复核数、高危问题数、金额与价税合计汇总、耗时）
  - `store_xlsx()`：三页台账 —— `发票台账`（冻结首行 + 筛选 + 千分位 + 状态色块）、`复核清单`（每条问题带源文件、级别、字段、原因、代码）、`统计`
- `make_fixtures.py`：用 PyMuPDF 内置简体中文字体生成 5 张样例发票，可重复生成
- `run_batch.py`：一条命令跑完，控制台打印统计与复核清单

## 实测结果（2026-11-27）

```
Ran 10 tests in 1.8s
OK

统计: {'files': 5, 'rows': 5, 'passed': 2, 'flagged': 3, 'issues': 4,
       'issue_high': 4, 'amount_sum': 5300.0, 'total_sum': 5999.0, 'elapsed': 0.02}
复核清单 4 条：
  [high] inv_c.pdf / total /      价税合计 914.00 ≠ 金额 800.00 + 税额 104.00（差 10.00）
  [high] inv_d.pdf / invoice_no / 缺发票号码：无法核验真伪，也无法查重
  [high] inv_d.pdf / amount /     金额与税额均为 0：空票或占位票，不应入账
  [high] inv_e.pdf / invoice_no / 号码与 inv_a.pdf 重复：可能是补开、重开或重复报销
```

**为什么值得看**：甲方最怕的不是"没抽到"，是"抽错了还不知道"。这一版把四类问题在 0.02 秒内挑出来，并逐条写清源文件、字段、差多少钱——人工只需看 4 行，而不是翻 5 个 PDF。

## 接真实单据时怎么改

1. **扫描件**：文本层为空会报 `EMPTY_TEXT`，此时接 OCR（PaddleOCR / 云 OCR），抽取层不用动；
2. **版式多变**：`FIELDS` 的正则改成"标签 + 值"的宽松规则，或按票种分模板（增值税专票/普票/收据各一套正则）；
3. **多页单据**：`read_pdf_text()` 已合并全部页，跨页字段可直接抽；
4. **要对接财务系统**：`store_jsonl()` 的输出就是稳定 schema，直接喂给入库脚本；
5. **量级**：单张 0.004 秒（含 PDF 打开），1000 张约 4 秒，瓶颈在磁盘不在解析。

合规边界：只处理用户提供的单据文件，不上传第三方、不破解加密 PDF、不做签名/真伪鉴别（那是税务平台的事）。
