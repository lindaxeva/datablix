#!/usr/bin/env python3
"""
sheetblix: a tiny, dependency-free helper for understanding your spreadsheet.

It reads a CSV file and produces two plain-language reports:

  1. Data freshness  : how old is this data, and how much of it is stale?
  2. Categorical glossary : what categories live in each column, how common
                            are they, and which values look inconsistent?

The goal is to turn a raw CSV into an "at a glance" summary that a
user can read and act on, without opening a spreadsheet
or writing any code.

Only Python's standard library is used, so the script runs anywhere
Python 3.8+ is installed. No "pip install" required.

Usage:
    python3 sheetblix.py data.csv
    python3 sheetblix.py data.csv --format markdown --output report.md
    python3 sheetblix.py data.csv --glossary-only
    python3 sheetblix.py data.csv --stale-days 60 --today 2026-06-26

Author: built for the GLOCAL Foundation "Automate a Useful Task" volunteer task.
License: MIT
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import Counter, OrderedDict
from datetime import datetime, date


# Date formats sheetblix will try, in order, when guessing date columns.
DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%m-%d-%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%d %b %Y",
    "%d %B %Y",
    "%b %d, %Y",
    "%B %d, %Y",
    "%Y%m%d",
)

# A value must parse under this share of non-empty cells for a column to be
# treated as a "date" or "numeric" column.
TYPE_CONFIDENCE = 0.80


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def read_csv(path):
    """Read a CSV into (headers, rows). Rows are lists of dict values.

    Tries utf-8 first, then falls back to latin-1 so messy real-world files
    still load. Raises a friendly error if the file is missing or empty.
    """
    if not os.path.isfile(path):
        raise FileNotFoundError("File not found: %s" % path)

    for encoding in ("utf-8-sig", "latin-1"):
        try:
            with open(path, newline="", encoding=encoding) as handle:
                reader = csv.reader(handle)
                rows = list(reader)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("Could not decode file with utf-8 or latin-1.")

    if not rows:
        raise ValueError("The CSV file is empty.")

    headers = [h.strip() for h in rows[0]]
    data = []
    for raw in rows[1:]:
        # Pad or trim ragged rows so every record lines up with the headers.
        record = OrderedDict()
        for i, header in enumerate(headers):
            record[header] = raw[i].strip() if i < len(raw) else ""
        data.append(record)
    return headers, data


# ---------------------------------------------------------------------------
# Value parsing helpers
# ---------------------------------------------------------------------------

def parse_date(value):
    """Return a datetime if value matches a known date format, else None."""
    text = value.strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_number(value):
    """Return a float if value looks numeric (handles $, %, commas), else None."""
    text = value.strip()
    if not text:
        return None
    cleaned = text.replace(",", "").replace("$", "").replace("%", "").strip()
    if cleaned in ("", "-", "."):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def normalize_label(value):
    """Collapse a value to a comparison key for spotting inconsistent entries.

    'Yes', ' yes ', and 'YES' all normalize to 'yes', so sheetblix can flag
    them as the same category written three different ways.
    """
    return re.sub(r"\s+", " ", value.strip().lower())


# ---------------------------------------------------------------------------
# Column type detection
# ---------------------------------------------------------------------------

def classify_columns(headers, rows, max_categories):
    """Decide a type for each column: date, numeric, categorical, text, empty."""
    types = OrderedDict()
    for header in headers:
        values = [r[header] for r in rows]
        non_empty = [v for v in values if v != ""]

        if not non_empty:
            types[header] = "empty"
            continue

        date_hits = sum(1 for v in non_empty if parse_date(v) is not None)
        if date_hits / len(non_empty) >= TYPE_CONFIDENCE:
            types[header] = "date"
            continue

        number_hits = sum(1 for v in non_empty if parse_number(v) is not None)
        is_numeric = number_hits / len(non_empty) >= TYPE_CONFIDENCE

        distinct = len(set(non_empty))
        # A column is "categorical" when its distinct values are few enough to
        # list and there is real repetition (so it is not an ID or free text).
        looks_categorical = distinct <= max_categories and distinct < len(non_empty)

        if is_numeric and not looks_categorical:
            types[header] = "numeric"
        elif looks_categorical:
            types[header] = "categorical"
        else:
            types[header] = "text"
    return types


# ---------------------------------------------------------------------------
# Freshness report
# ---------------------------------------------------------------------------

def build_freshness(headers, rows, types, today, fresh_days, stale_days):
    """Summarize how recent the data is, per detected date column."""
    date_columns = [h for h in headers if types[h] == "date"]
    columns_summary = []

    for header in date_columns:
        parsed = [parse_date(r[header]) for r in rows]
        parsed = [d for d in parsed if d is not None]
        if not parsed:
            continue

        newest = max(parsed)
        oldest = min(parsed)
        days_since = (today - newest.date()).days

        if days_since <= fresh_days:
            status = "Fresh"
        elif days_since <= stale_days:
            status = "Aging"
        else:
            status = "Stale"

        stale_rows = sum(
            1 for d in parsed if (today - d.date()).days > stale_days
        )

        columns_summary.append(OrderedDict([
            ("column", header),
            ("newest", newest.date().isoformat()),
            ("oldest", oldest.date().isoformat()),
            ("days_since_newest", days_since),
            ("status", status),
            ("rows_with_dates", len(parsed)),
            ("stale_rows", stale_rows),
            ("stale_share_pct", round(100 * stale_rows / len(parsed), 1)),
        ]))

    overall = None
    if columns_summary:
        # Overall dataset freshness uses the most recently touched date column.
        freshest = min(columns_summary, key=lambda c: c["days_since_newest"])
        overall = OrderedDict([
            ("reference_date", today.isoformat()),
            ("most_recent_value", freshest["newest"]),
            ("days_since", freshest["days_since_newest"]),
            ("status", freshest["status"]),
            ("driven_by_column", freshest["column"]),
        ])

    return OrderedDict([
        ("reference_date", today.isoformat()),
        ("fresh_threshold_days", fresh_days),
        ("stale_threshold_days", stale_days),
        ("overall", overall),
        ("date_columns", columns_summary),
    ])


# ---------------------------------------------------------------------------
# Categorical glossary
# ---------------------------------------------------------------------------

def build_glossary(headers, rows, types, max_categories, top_n):
    """Describe each categorical column: its values, counts, and inconsistencies."""
    glossary = []
    for header in headers:
        if types[header] != "categorical":
            continue

        values = [r[header] for r in rows]
        non_empty = [v for v in values if v != ""]
        missing = len(values) - len(non_empty)

        counts = Counter(non_empty)
        ranked = counts.most_common()
        total = len(non_empty)

        entries = []
        for value, count in ranked[:top_n]:
            entries.append(OrderedDict([
                ("value", value),
                ("count", count),
                ("share_pct", round(100 * count / total, 1) if total else 0.0),
            ]))

        # Detect inconsistent spellings: raw values that share a normalized key.
        groups = {}
        for value in counts:
            groups.setdefault(normalize_label(value), set()).add(value)
        inconsistencies = [
            sorted(variants) for variants in groups.values() if len(variants) > 1
        ]

        glossary.append(OrderedDict([
            ("column", header),
            ("distinct_values", len(counts)),
            ("missing_cells", missing),
            ("missing_share_pct",
             round(100 * missing / len(values), 1) if values else 0.0),
            ("top_values", entries),
            ("more_values_not_shown", max(0, len(ranked) - top_n)),
            ("inconsistent_value_groups", inconsistencies),
        ]))
    return glossary


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def render_text(meta, freshness, glossary):
    out = []
    out.append("=" * 64)
    out.append("SHEETBLIX REPORT")
    out.append("=" * 64)
    out.append("File      : %s" % meta["file"])
    out.append("Rows      : %d" % meta["rows"])
    out.append("Columns   : %d" % meta["columns"])
    out.append("")

    if freshness is not None:
        out.append("-" * 64)
        out.append("DATA FRESHNESS")
        out.append("-" * 64)
        overall = freshness["overall"]
        if overall is None:
            out.append("No date columns detected.")
        else:
            out.append("Reference date    : %s" % overall["reference_date"])
            out.append("Most recent value : %s" % overall["most_recent_value"])
            out.append("Days since        : %d" % overall["days_since"])
            out.append("Overall status    : %s" % overall["status"])
            out.append("Driven by column  : %s" % overall["driven_by_column"])
            out.append("")
            for col in freshness["date_columns"]:
                out.append("  Column: %s" % col["column"])
                out.append("    range        : %s to %s"
                           % (col["oldest"], col["newest"]))
                out.append("    days since   : %d (%s)"
                           % (col["days_since_newest"], col["status"]))
                out.append("    stale rows   : %d of %d (%.1f%%)"
                           % (col["stale_rows"], col["rows_with_dates"],
                              col["stale_share_pct"]))
                out.append("")

    if glossary is not None:
        out.append("-" * 64)
        out.append("CATEGORICAL GLOSSARY")
        out.append("-" * 64)
        if not glossary:
            out.append("No categorical columns detected.")
            out.append("")
        for col in glossary:
            out.append("Column: %s" % col["column"])
            out.append("  distinct values : %d" % col["distinct_values"])
            out.append("  missing cells   : %d (%.1f%%)"
                       % (col["missing_cells"], col["missing_share_pct"]))
            for entry in col["top_values"]:
                out.append("    %-24s %6d  (%.1f%%)"
                           % (entry["value"][:24], entry["count"],
                              entry["share_pct"]))
            if col["more_values_not_shown"]:
                out.append("    ... and %d more value(s)"
                           % col["more_values_not_shown"])
            if col["inconsistent_value_groups"]:
                out.append("  possible inconsistencies:")
                for group in col["inconsistent_value_groups"]:
                    out.append("    -> %s" % ", ".join(group))
            out.append("")

    return "\n".join(out).rstrip() + "\n"


def render_markdown(meta, freshness, glossary):
    out = []
    out.append("# sheetblix report")
    out.append("")
    out.append("| Field | Value |")
    out.append("| --- | --- |")
    out.append("| File | %s |" % meta["file"])
    out.append("| Rows | %d |" % meta["rows"])
    out.append("| Columns | %d |" % meta["columns"])
    out.append("")

    if freshness is not None:
        out.append("## Data freshness")
        out.append("")
        overall = freshness["overall"]
        if overall is None:
            out.append("_No date columns detected._")
            out.append("")
        else:
            out.append("**Overall status: %s** "
                       "(most recent value %s, %d day(s) ago, reference %s)"
                       % (overall["status"], overall["most_recent_value"],
                          overall["days_since"], overall["reference_date"]))
            out.append("")
            out.append("| Column | Oldest | Newest | Days since | Status | Stale rows |")
            out.append("| --- | --- | --- | --- | --- | --- |")
            for col in freshness["date_columns"]:
                out.append("| %s | %s | %s | %d | %s | %d (%.1f%%) |"
                           % (col["column"], col["oldest"], col["newest"],
                              col["days_since_newest"], col["status"],
                              col["stale_rows"], col["stale_share_pct"]))
            out.append("")

    if glossary is not None:
        out.append("## Categorical glossary")
        out.append("")
        if not glossary:
            out.append("_No categorical columns detected._")
            out.append("")
        for col in glossary:
            out.append("### %s" % col["column"])
            out.append("")
            out.append("Distinct values: %d. Missing cells: %d (%.1f%%)."
                       % (col["distinct_values"], col["missing_cells"],
                          col["missing_share_pct"]))
            out.append("")
            out.append("| Value | Count | Share |")
            out.append("| --- | --- | --- |")
            for entry in col["top_values"]:
                out.append("| %s | %d | %.1f%% |"
                           % (entry["value"], entry["count"], entry["share_pct"]))
            out.append("")
            if col["more_values_not_shown"]:
                out.append("_... and %d more value(s) not shown._"
                           % col["more_values_not_shown"])
                out.append("")
            if col["inconsistent_value_groups"]:
                out.append("**Possible inconsistencies (same value, different spelling):**")
                out.append("")
                for group in col["inconsistent_value_groups"]:
                    out.append("- %s" % ", ".join(group))
                out.append("")

    return "\n".join(out).rstrip() + "\n"


def render_json(meta, freshness, glossary):
    payload = OrderedDict([
        ("meta", meta),
        ("freshness", freshness),
        ("glossary", glossary),
    ])
    return json.dumps(payload, indent=2) + "\n"


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_today(value):
    if value is None:
        return date.today()
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        raise argparse.ArgumentTypeError(
            "--today must look like YYYY-MM-DD, got: %s" % value
        )


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="sheetblix",
        description="Summarize data freshness and categorical values in a CSV.",
    )
    parser.add_argument("csv_file", help="Path to the CSV file to analyze.")
    parser.add_argument("--format", choices=("text", "markdown", "json"),
                        default="text", help="Output format (default: text).")
    parser.add_argument("--output", help="Write the report to this file "
                        "instead of the screen.")
    parser.add_argument("--freshness-only", action="store_true",
                        help="Show only the freshness section.")
    parser.add_argument("--glossary-only", action="store_true",
                        help="Show only the categorical glossary section.")
    parser.add_argument("--max-categories", type=int, default=20,
                        help="Max distinct values for a column to count as "
                             "categorical (default: 20).")
    parser.add_argument("--top", type=int, default=10,
                        help="How many top values to list per column "
                             "(default: 10).")
    parser.add_argument("--fresh-days", type=int, default=30,
                        help="Data newer than this many days is Fresh "
                             "(default: 30).")
    parser.add_argument("--stale-days", type=int, default=90,
                        help="Data older than this many days is Stale "
                             "(default: 90).")
    parser.add_argument("--today", type=parse_today, default=None,
                        help="Reference date as YYYY-MM-DD (default: today).")

    args = parser.parse_args(argv)
    today = args.today if args.today is not None else date.today()

    if args.freshness_only and args.glossary_only:
        parser.error("Choose at most one of --freshness-only / --glossary-only.")

    try:
        headers, rows = read_csv(args.csv_file)
    except (FileNotFoundError, ValueError) as exc:
        print("Error: %s" % exc, file=sys.stderr)
        return 1

    types = classify_columns(headers, rows, args.max_categories)

    meta = OrderedDict([
        ("file", os.path.basename(args.csv_file)),
        ("rows", len(rows)),
        ("columns", len(headers)),
        ("column_types", types),
    ])

    show_freshness = not args.glossary_only
    show_glossary = not args.freshness_only

    freshness = (build_freshness(headers, rows, types, today,
                                 args.fresh_days, args.stale_days)
                 if show_freshness else None)
    glossary = (build_glossary(headers, rows, types,
                               args.max_categories, args.top)
                if show_glossary else None)

    if args.format == "markdown":
        report = render_markdown(meta, freshness, glossary)
    elif args.format == "json":
        report = render_json(meta, freshness, glossary)
    else:
        report = render_text(meta, freshness, glossary)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as handle:
            handle.write(report)
        print("Report written to %s" % args.output)
    else:
        sys.stdout.write(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
