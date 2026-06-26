#!/usr/bin/env python3
"""
Tests for sheetblix.

Run from the project folder with:

    python3 -m unittest -v

Uses only Python's standard library, matching the tool itself. No pytest,
no installs.
"""

import os
import tempfile
import unittest
from datetime import date

import sheetblix


class TestValueParsing(unittest.TestCase):

    def test_parse_date_common_formats(self):
        self.assertIsNotNone(datablix.parse_date("2026-06-26"))
        self.assertIsNotNone(datablix.parse_date("26/06/2026"))
        self.assertIsNotNone(datablix.parse_date("June 26, 2026"))

    def test_parse_date_rejects_non_dates(self):
        self.assertIsNone(datablix.parse_date("not a date"))
        self.assertIsNone(datablix.parse_date(""))

    def test_parse_number_handles_symbols(self):
        self.assertEqual(datablix.parse_number("1,250"), 1250.0)
        self.assertEqual(datablix.parse_number("$45.50"), 45.50)
        self.assertEqual(datablix.parse_number("80%"), 80.0)

    def test_parse_number_rejects_text(self):
        self.assertIsNone(datablix.parse_number("hello"))
        self.assertIsNone(datablix.parse_number(""))

    def test_normalize_label_collapses_case_and_space(self):
        self.assertEqual(datablix.normalize_label("  Active "), "active")
        self.assertEqual(datablix.normalize_label("YES"), "yes")
        self.assertEqual(
            datablix.normalize_label("British   Columbia"),
            "british columbia",
        )


class TestColumnClassification(unittest.TestCase):

    def setUp(self):
        self.headers = ["id", "region", "joined", "score"]
        self.rows = [
            {"id": "1", "region": "Ontario", "joined": "2026-01-01", "score": "10"},
            {"id": "2", "region": "Quebec", "joined": "2026-02-01", "score": "20"},
            {"id": "3", "region": "Ontario", "joined": "2026-03-01", "score": "30"},
            {"id": "4", "region": "Quebec", "joined": "2026-04-01", "score": "40"},
        ]

    def test_types_detected(self):
        types = datablix.classify_columns(self.headers, self.rows, max_categories=20)
        self.assertEqual(types["joined"], "date")
        self.assertEqual(types["region"], "categorical")
        # Every id is unique, so it should not be treated as categorical.
        self.assertIn(types["id"], ("text", "numeric"))


class TestFreshness(unittest.TestCase):

    def test_status_buckets(self):
        headers = ["updated"]
        rows = [
            {"updated": "2026-06-20"},  # recent
            {"updated": "2026-01-01"},  # old
        ]
        types = datablix.classify_columns(headers, rows, max_categories=20)
        report = datablix.build_freshness(
            headers, rows, types,
            today=date(2026, 6, 26),
            fresh_days=30, stale_days=90,
        )
        overall = report["overall"]
        self.assertEqual(overall["status"], "Fresh")
        self.assertEqual(overall["most_recent_value"], "2026-06-20")
        col = report["date_columns"][0]
        self.assertEqual(col["stale_rows"], 1)

    def test_no_date_columns(self):
        headers = ["name"]
        rows = [{"name": "Alex"}, {"name": "Sam"}]
        types = datablix.classify_columns(headers, rows, max_categories=20)
        report = datablix.build_freshness(
            headers, rows, types,
            today=date(2026, 6, 26),
            fresh_days=30, stale_days=90,
        )
        self.assertIsNone(report["overall"])


class TestGlossary(unittest.TestCase):

    def test_counts_and_inconsistencies(self):
        headers = ["status"]
        rows = [
            {"status": "Active"},
            {"status": "active"},
            {"status": "Active"},
            {"status": "Inactive"},
            {"status": ""},
        ]
        types = datablix.classify_columns(headers, rows, max_categories=20)
        glossary = datablix.build_glossary(
            headers, rows, types, max_categories=20, top_n=10,
        )
        self.assertEqual(len(glossary), 1)
        col = glossary[0]
        self.assertEqual(col["column"], "status")
        self.assertEqual(col["missing_cells"], 1)
        # "Active" and "active" should be flagged as the same value.
        flagged = col["inconsistent_value_groups"]
        self.assertTrue(any("Active" in g and "active" in g for g in flagged))


class TestReadCsv(unittest.TestCase):

    def test_reads_headers_and_rows(self):
        content = "a,b\n1,2\n3,4\n"
        tmp = tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, encoding="utf-8"
        )
        try:
            tmp.write(content)
            tmp.close()
            headers, rows = datablix.read_csv(tmp.name)
            self.assertEqual(headers, ["a", "b"])
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["a"], "1")
        finally:
            os.unlink(tmp.name)

    def test_missing_file_raises(self):
        with self.assertRaises(FileNotFoundError):
            datablix.read_csv("definitely_not_here.csv")


if __name__ == "__main__":
    unittest.main(verbosity=2)
