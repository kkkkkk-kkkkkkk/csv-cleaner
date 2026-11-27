# Trial pack — messy CSV in, auditable workbook out

One command. About three seconds. No install beyond `openpyxl`.

```bash
python trial.py            # runs on the bundled sample (271 raw rows)
python trial.py your.csv   # runs on your own export
```

## What you get — `trial_out.xlsx`, six sheets

| Sheet | What it answers |
|---|---|
| `口径` | What file, how many rows, what was dropped and why, when generated |
| `明细` | Every surviving row, cleaned — amounts parsed, regions normalised, dates standardised |
| `汇总` | Four SUMIFS blocks; change a row in `明细` and the totals move with it |
| `核对` | 7 formulas that cross-check the total three ways; PASS only when diff is 0.000 |
| `差异` | Optional: pass `--baseline last_period.xlsx` and get amount / row-count deltas |
| `图表` | Region and month bars reading off the same live ranges |

Sample run on the bundled file:

```
271 raw rows  ->  171 valid  (89 dropped, each one listed)
Total 97,735.00   Quantity 552
Elapsed 2.75s
```

## Two rules this pack is built on

1. **A missing value is not a zero.** Empty amounts are flagged and listed, never averaged into the total.
2. **The total is a formula, not a number.** Every total cell is `SUMIFS` over `明细`, so a changed row changes the total, and `核对` recomputes it three independent ways.

## What it deliberately does not do

- It does not guess your column meanings. If your headers are unusual, send one sample row and it takes one extra pass.
- It does not fabricate missing cells. Blank stays blank and gets reported.
- It does not touch your original file — input is opened read-only.

## Sizing

Up to ~1,000 rows runs in seconds on a laptop. The logic is pure Python + openpyxl, so there is no cloud account, no upload, no per-row fee.

---

Author: Jixin Tan — Python, Go, openpyxl. Workbook logic, not spreadsheets-by-hand.
Repo: https://github.com/kkkkkk-kkkkkkk/csv-cleaner
