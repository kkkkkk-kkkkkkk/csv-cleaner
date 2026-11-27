# -*- coding: utf-8 -*-
"""Tests for csv-cleaner. Run: python -m unittest test_clean -v"""
import csv
import datetime
import io
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import clean as C


class TestParseDate(unittest.TestCase):
    def test_iso_is_unambiguous(self):
        self.assertEqual(C.parse_date("2026-11-25"), (datetime.date(2026, 11, 25), ""))
        self.assertEqual(C.parse_date("2026/11/25")[0], datetime.date(2026, 11, 25))

    def test_mdy_and_dmy_disagree_on_purpose(self):
        # 03/04/2026 is either March 4th or April 3rd; the caller decides, we never guess.
        self.assertEqual(C.parse_date("03/04/2026", "mdy")[0], datetime.date(2026, 3, 4))
        self.assertEqual(C.parse_date("03/04/2026", "dmy")[0], datetime.date(2026, 4, 3))

    def test_unparsed_is_reported_not_coerced(self):
        d, why = C.parse_date("25 Nov 2026")
        self.assertIsNone(d)
        self.assertEqual(why, "unparsed")

    def test_empty(self):
        self.assertEqual(C.parse_date("  "), (None, "empty"))


class TestParseMoney(unittest.TestCase):
    def test_yen_and_thousands(self):
        self.assertEqual(C.parse_money("¥12,500.00"), (12500.0, True))
        self.assertEqual(C.parse_money("12500"), (12500.0, False))

    def test_european_decimal_comma(self):
        self.assertEqual(C.parse_money("1.234,56")[0], 1234.56)
        self.assertEqual(C.parse_money("1.234,56")[1], True)

    def test_us_mixed(self):
        self.assertEqual(C.parse_money("1,234.56")[0], 1234.56)

    def test_trailing_word(self):
        self.assertEqual(C.parse_money("3500元")[0], 3500.0)

    def test_junk(self):
        self.assertEqual(C.parse_money("n/a"), (None, False))
        self.assertEqual(C.parse_money(""), (None, False))


class TestParseQty(unittest.TestCase):
    def test_integer_strings(self):
        self.assertEqual(C.parse_qty("3"), (3, ""))
        self.assertEqual(C.parse_qty("3.0"), (3, ""))

    def test_float_is_flagged(self):
        self.assertEqual(C.parse_qty("3.5"), (None, "not-integer"))

    def test_junk(self):
        self.assertEqual(C.parse_qty("two"), (None, "unparsed"))


class TestCleanRows(unittest.TestCase):
    HEADER = ["order_id", "order_date", "amount", "qty", "region", "product", "rep"]
    SCHEMA = {"id": "order_id", "date": "order_date", "amount": "amount",
              "qty": "qty", "group": "region", "label": "product", "owner": "rep"}

    def run_clean(self, rows, date_order="mdy"):
        return C.clean_rows(self.HEADER, rows, self.SCHEMA, date_order)

    def test_blank_rows_are_counted_not_kept(self):
        rows, stats, _ = self.run_clean([["", "", "", "", "", "", ""]])
        self.assertEqual(stats["blank_rows"], 1)
        self.assertEqual(rows, [])

    def test_duplicates_dropped_once(self):
        r = ["A1", "2026-01-05", "100", "2", "North", "Widget", "Kim"]
        rows, stats, _ = self.run_clean([r, list(r)])
        self.assertEqual(stats["duplicate_rows"], 1)
        self.assertEqual(len(rows), 1)

    def test_non_positive_qty_drops_row(self):
        rows, stats, issues = self.run_clean([["A1", "2026-01-05", "100", "0", "N", "W", "K"]])
        self.assertTrue(rows[0]["dropped"])
        self.assertEqual(rows[0]["qty"], 0)
        self.assertIn("qty non-positive", issues[0]["problem"])

    def test_line_numbers_are_source_lines(self):
        rows, _, _ = self.run_clean([["A1", "2026-01-05", "100", "2", "N", "W", "K"]])
        self.assertEqual(rows[0]["line"], 2)

    def test_sum_of_kept_rows_matches_aggregate(self):
        rows, _, _ = self.run_clean([
            ["A1", "2026-01-05", "100", "2", "North", "W", "K"],
            ["A2", "2026-02-05", "¥1,500", "3", "North", "W2", "K"],
            ["A3", "2026-02-09", "200", "-1", "South", "W", "K"],
        ])
        good, by_group, by_month, _ = C.aggregate(rows)
        self.assertEqual(len(good), 2)
        self.assertEqual(by_group["North"][1], 1600.0)
        self.assertEqual(by_month["2026-02"][1], 1500.0)

    def test_unparsed_date_does_not_become_a_row(self):
        rows, stats, issues = self.run_clean([["A1", "25 Nov 2026", "100", "2", "N", "W", "K"]])
        self.assertEqual(rows[0]["date"], None)
        self.assertEqual(stats["date_unparsed"], 1)
        self.assertIn("date unparsed", rows[0]["flags"])

    def test_no_schema_only_structural_cleaning(self):
        rows, stats, _ = C.clean_rows(["a", "b"], [["1", "2"], ["1", "2"]], {})
        self.assertEqual(len(rows), 1)
        self.assertEqual(stats["duplicate_rows"], 1)


class TestSchemaAndEndToEnd(unittest.TestCase):
    def test_unknown_schema_key_is_an_error(self):
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
            f.write('{"date": "d", "totally_made_up": "x"}')
            path = f.name
        try:
            with self.assertRaises(RuntimeError):
                C.load_schema(path)
        finally:
            os.unlink(path)

    def test_header_absent_from_file_is_skipped(self):
        idx = C.resolve_columns(["order_date", "amount"], {"date": "order_date", "qty": "quantity"})
        self.assertEqual(idx, {"date": 0})
        self.assertNotIn("qty", idx)

    def test_end_to_end_writes_workbook(self):
        tmp = tempfile.mkdtemp()
        src = os.path.join(tmp, "messy.csv")
        dst = os.path.join(tmp, "out.xlsx")
        sch = os.path.join(tmp, "schema.json")
        with open(src, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["order_id", "order_date", "amount", "qty"])
            w.writerow(["A1", "2026-01-05", "¥100", "2"])
            w.writerow(["A1", "2026-01-05", "¥100", "2"])
            w.writerow(["A2", "25 Nov 2026", "500", "1"])
        with open(sch, "w", encoding="utf-8") as f:
            f.write('{"id": "order_id", "date": "order_date", "amount": "amount", "qty": "qty"}')

        rc = C.main([src, dst, "--schema", sch, "--strict"])
        self.assertEqual(rc, 1, "strict must fail when a date is unparsed")
        self.assertTrue(os.path.exists(dst))
        from openpyxl import load_workbook
        wb = load_workbook(dst)
        for name in ("cleaned", "summary", "charts", "issues"):
            self.assertIn(name, wb.sheetnames)
        ws = wb["cleaned"]
        self.assertEqual(ws.max_row, 3)          # header + 2 rows surviving dedupe
        problems = [wb["issues"].cell(row=r, column=3).value
                    for r in range(2, wb["issues"].max_row + 1)]
        self.assertIn("date unparsed", problems)
        self.assertTrue(any("currency" in (p or "") for p in problems))


class TestTextNormalisation(unittest.TestCase):
    """Case collapsing is safe; semantic merging is not, so it must never happen."""

    HEAD = ["id", "dt", "amt", "qty", "region", "rep"]
    SCH = {"id": ["id"], "date": ["dt"], "amount": ["amt"],
           "qty": ["qty"], "group": ["region"], "owner": ["rep"]}

    def run_rows(self, rows, order="mdy"):
        return C.clean_rows(self.HEAD, rows, self.SCH, order)

    def test_case_variants_collapse_to_the_majority_spelling(self):
        rows = [["A", "2026-01-05", "10", "1", "north", "Kim"],
                ["B", "2026-01-06", "20", "2", "North", "Kim"],
                ["C", "2026-01-07", "30", "3", "NORTH", "Kim"]]
        got, stats, issues = self.run_rows(rows)
        self.assertEqual([r["group"] for r in got], ["North"] * 3)
        self.assertEqual(stats["text_case_variants"], 1)
        self.assertEqual(stats["text_case_rows"], 2)
        self.assertTrue(any("case-normalised" in r["flags"] for r in got))

    def test_collapsing_case_then_finding_duplicates(self):
        # Same business row written twice, differing only in case: without the
        # normalisation step this would survive as two fake records.
        rows = [["A", "2026-01-05", "10", "1", "north", "Kim"],
                ["A", "2026-01-05", "10", "1", "North", "Kim"]]
        got, stats, _ = self.run_rows(rows)
        self.assertEqual(stats["duplicate_rows"], 1)
        self.assertEqual(len(got), 1)

    def test_abbreviation_is_flagged_but_not_merged(self):
        # 'K' may be 'Kim' or may be a different person; guessing corrupts data.
        rows = [["A", "2026-01-05", "10", "1", "North", "Kim"],
                ["B", "2026-01-06", "20", "2", "North", "K"]]
        got, stats, issues = self.run_rows(rows)
        self.assertEqual({r["owner"] for r in got}, {"Kim", "K"})
        self.assertEqual(stats["text_truncation_suspects"], 1)
        self.assertTrue(any("truncation" in i["problem"] for i in issues))

    def test_east_and_eastern_stay_apart(self):
        rows = [["A", "2026-01-05", "10", "1", "East", "Kim"],
                ["B", "2026-01-06", "20", "2", "Eastern", "Kim"]]
        got, stats, issues = self.run_rows(rows)
        self.assertEqual({r["group"] for r in got}, {"East", "Eastern"})
        self.assertEqual(stats["text_case_variants"], 0)
        self.assertTrue(any("truncation" in i["problem"] for i in issues))

    def test_number_in_text_column_is_flagged(self):
        rows = [["A", "2026-01-05", "10", "1", "5", "Widget C"]]
        got, stats, issues = self.run_rows(rows)
        self.assertEqual(stats["numeric_in_text"], 1)
        self.assertEqual(got[0]["group"], "5", "value is reported, not rewritten")
        self.assertIn("looks numeric", got[0]["flags"])

    def test_empty_group_does_not_join_the_collapse(self):
        rows = [["A", "2026-01-05", "10", "1", "", "Kim"],
                ["B", "2026-01-06", "20", "2", "", "Kim"]]
        got, stats, _ = self.run_rows(rows)
        self.assertEqual(stats["text_case_variants"], 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
