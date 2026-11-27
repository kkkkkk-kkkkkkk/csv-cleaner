**Try it in 3 seconds, no install beyond openpyxl:** [download dist/trial_pack.zip](https://github.com/kkkkkk-kkkkkkk/csv-cleaner/raw/main/dist/trial_pack.zip) and run python trial.py. Sample: 271 raw rows -> 171 valid, 97,735.00, 7 self-check formulas, 2.75s.

**Second trial pack — invoices:** [download dist/pdf_trial_pack.zip](https://github.com/kkkkkk-kkkkkkk/csv-cleaner/raw/main/dist/pdf_trial_pack.zip) and run python trial.py. Sample: 5 invoice PDFs -> 5 ledger rows, 8 fields each, 4 issues flagged with field+reason, 1.30s, 10 unit tests.

**[Live demo & sample output — see what “clean” actually looks like →](https://kkkkkk-kkkkkkk.github.io/csv-cleaner/)**

---
## What's in this repo

Four small tools, one principle: **never guess, never silently drop a row, always hand back a receipt.** Each runs offline on your own machine — standard library plus `openpyxl` (and `PyMuPDF` for the invoice reader).

| | tool | turns | the number it reports |
|---|---|---|---|
| 1 | **csv-cleaner** — `clean.py`, repo root | a messy CSV export → clean workbook + row-by-row issue report | 28 unit tests; every dropped row is counted and named |
| 2 | **crawl → table** — `demos/tender_crawler` | list pages → a 3-sheet workbook: records / failures / field fill-rate | 4 pages in → 3 stored, 1 deduplicated, fill-rate printed as 0.667 |
| 3 | **invoice reader** — `demos/pdf_extract` | a folder of invoice PDFs → a ledger workbook with a review queue | 5 invoices → 40 fields, 3 flagged for review, 4 hard issues caught (duplicate number, tax mismatch, zero amount) |
| 4 | **report builder** — `demos/excel_report` | a messy monthly sales file → a 6-sheet workbook where every total is a formula | 268 rows → 171 kept, 89 dropped, total 97,735.00, 7 of 7 cross-checks PASS at variance 0.000 |

The common thread is the **receipt**: for each job the output says what came in, what went out, what was rejected and why, and what still needs a human decision. That is the difference between "my script cleaned your file" and a workbook you can defend in a meeting.

Two rules hold across all four:

- **Missing is not zero.** A blank amount stays blank, gets flagged, and stays out of the totals — the classic way a report quietly comes out too low.
- **Totals are formulas, not numbers.** In the report builder, the summary and the check page are `SUMIFS` formulas over the detail sheet, so editing one row moves the totals and a PASS/FAIL page tells you whether they still tie. Verified by recalculating the workbook in LibreOffice, not by the script asserting it is right.

Each folder has its own README with the exact command, the output layout, and the known limits. Dependencies are deliberately boring: `pip install openpyxl pymupdf`.

---

# csv-cleaner

Turn a messy CSV export into a clean Excel workbook **and** a line-by-line issue report. Standard library plus `openpyxl`, no service, no upload, runs on your machine or a cron job.

Written because "clean my spreadsheet" work almost always fails the same way: the tool guesses. It reads `03/04/2026` as March 4th when the client meant April 3rd, quietly drops the rows it could not parse, and hands back a workbook nobody can audit. This one never guesses and never hides what it changed.

## Usage

```bash
python clean.py messy.csv clean.xlsx --schema schema.example.json
```

Optional:

| flag | meaning |
|---|---|
| `--date-order mdy\|dmy` | interpretation for `03/04/2026`-style dates. ISO (`2026-03-04`) is always tried first and is unambiguous. Default `mdy`. |
| `--strict` | exit code `1` if any row was dropped or any field failed to parse, so the tool can gate a pipeline. |
| `--schema` | JSON mapping semantic fields to your header names. Omit it and you still get structural cleaning (whitespace, blank rows, duplicates). |

### Schema

Header names are yours, not mine:

```json
{
  "id": "order_id",
  "date": ["order_date", "date"],
  "amount": ["amount", "total"],
  "qty": "qty",
  "group": "region",
  "label": "product",
  "owner": "rep"
}
```

`date`, `amount`, `qty`, `id`, `group`, `label`, `owner` are the only recognised keys; anything else is rejected instead of silently ignored. A field whose header is absent from the file is skipped, not treated as an error — real exports drop columns all the time.

## What it does to the data

- strips whitespace, drops blank rows, drops exact duplicates **on the normalised business values**, not on the raw text;
- reads `2026-01-05`, `2026/01/05`, `2026.01.05` as unambiguous, then applies `--date-order` for slash forms;
- unparsed dates stay visible with a flag and land in `issues` — they are never silently nulled into a fake aggregate;
- money: `$1,250.00`, `¥780`, `1.234,56` (European), `1,234.56` (US), `$2 400`, `3500元` all parse. The rule is positional: the rightmost separator is the decimal one. Symbols that had to be stripped are recorded, so the client can see what changed;
- quantity must be a whole number; `0` or negative drops the row (discount lines and returns in a sales export are not sales);
- text fields (`id`, `group`, `label`, `owner`) get the one normalisation that is safe and the rest reported — see below;
- every row keeps its source line number, so `issues` points at a line you can open in Excel and fix.

## Text fields: what gets merged and what does not

`north`, `North` and `  NORTH  ` are one value written three ways, so they are collapsed to a single spelling and every rewritten cell is flagged `group case-normalised to North`. This runs **before** duplicate detection, otherwise two identical sales recorded once as `North` and once as `north` would survive as two records — the classic inflated-total bug.

Which spelling wins: the most frequent one; on a tie, the one that looks typed by a human (mixed case `North` > all caps `USA` > all lower case `north`); then the earliest. The tool never invents a spelling that was not already in the file.

What is **never** merged, only reported:

- `East` vs `Eastern` — plausibly the same region, plausibly a parent and a child. Reported as *possible truncation or abbreviation, left as-is*; the client decides.
- `K` vs `Kim` — could be the same rep abbreviated, could be a different person. Guessing here silently corrupts a per-rep commission report.
- `5` sitting in a text column — almost always a shifted row rather than a real category. Flagged `group looks numeric`, value left untouched.

## Output workbook

| sheet | contents |
|---|---|
| `cleaned` | source line, mapped fields, `flags`, `dropped`; filter enabled |
| `summary` | grouped by group / by month / by label, each with qty and amount |
| `charts` | native Excel bar and line charts (editable, not images) |
| `share` | pie chart of label share |
| `issues` | one row per problem: source line, column, problem, offending value |

Native charts matter: the client can restyle them, and they do not turn into a 300-dpi PNG the moment the data changes.

## Tests

```bash
python -m unittest test_clean -v
```

28 tests, no network, no fixtures on disk except a temp dir. They pin the things that quietly break real jobs:

- `03/04/2026` resolves differently under `mdy` and `dmy`, deliberately;
- `1.234,56` is 1234.56 and `1,234.56` is 1234.56;
- a duplicate pair collapses to one row while the dropped copy is still reported;
- aggregate totals equal the sum of surviving rows (checked, not assumed);
- `--strict` returns exit `1` when a date cannot be parsed, and writes the workbook anyway so the issue is inspectable;
- `north`/`North`/`NORTH` collapse to one spelling while `East`/`Eastern` and `K`/`Kim` deliberately do not;
- a row written once as `North` and once as `north` is detected as a duplicate, not kept twice.

## Sample

`sample_messy.csv` is 21 lines of deliberately broken export: a blank row, an exact duplicate, `¥780` with qty `0`, an `?` amount, `$2 400`, `1.234,56`, `25 Nov 2026`, `north` vs `  North  ` vs `Eastern`. Running it produces:

```
rows in     21
rows kept   17
rows dropped 2
  blank_rows           1
  duplicate_rows       1
  date_unparsed        1
  amount_normalised    4
  amount_missing       3
  qty_non_positive     2
  text_case_variants   2
  text_case_rows       2
  text_truncation_suspects 2
  numeric_in_text      1
strict: problems found, exit 1
```

21 = 1 blank + 1 duplicate + 2 dropped + 17 kept. Counters are counted **before** de-duplication, so a repeated bad row is reported twice and once as a duplicate — that is deliberate, because the client needs to fix the source, not just the count.

## Deliberate limits

- No fuzzy name matching between duplicate-looking rows (`ACME` vs `Acme Ltd`). Merging customers requires a rule the client has to state, and getting it wrong is worse than leaving it alone.
- No currency conversion. Amounts are reported as numbers; mixing currencies in one column is an input error, and the tool will happily add them up unless the client separates them.
- Dates that are genuinely ambiguous in the file (`03/04/2026` with no stated order) are resolved by the flag you pass, never by inference from other rows.
- Output is `.xlsx`; the report is a sheet, not a separate PDF. If a client wants a PDF, that is a rendering step on top, not a second source of truth.

## Files

```
clean.py               CLI + pure functions (clean_rows is IO-free, hence testable)
test_clean.py          28 unittest cases
schema.example.json    semantic field -> header name mapping
sample_messy.csv       21 broken lines + expected summary above
```
