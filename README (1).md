# datablix

A tiny, dependency-free command-line tool that turns a raw CSV into an
"at a glance" summary a non-technical person can actually read and act on.

It answers two everyday questions about a spreadsheet:

1. **Is this data still fresh?** How recent is it, and how much of it is stale?
2. **What categories are in here?** What values appear in each column, how
   common are they, and which ones look like the same thing typed two
   different ways?

datablix uses **only the Python standard library**, so there is nothing to
install. If you have Python 3.8 or newer, it runs.

## The problem it solves

Lists and records drift over time. Two common failure modes show up again
and again in volunteer, community, and small-office settings:

- A contact list, roster, or tracker quietly goes out of date, and nobody
  notices until it causes a problem.
- The same category gets typed inconsistently ("Active", "active", "ACTIVE"),
  which breaks filters, counts, and reports without anyone realizing.

Checking for these by eye is slow and error-prone. datablix does it in one
command and explains what it found in plain language.

## Install

No installation needed. Clone the repository and run the script directly:

```bash
git clone https://github.com/<your-username>/datablix.git
cd datablix
python3 datablix.py sample_data.csv
```

## Usage

```bash
python3 datablix.py YOUR_FILE.csv [options]
```

### Common examples

```bash
# Full report (freshness + glossary) to the screen
python3 datablix.py sample_data.csv

# Only the freshness section
python3 datablix.py sample_data.csv --freshness-only

# Only the categorical glossary
python3 datablix.py sample_data.csv --glossary-only

# Save a shareable Markdown report
python3 datablix.py sample_data.csv --format markdown --output report.md

# Get machine-readable JSON (for piping into other tools)
python3 datablix.py sample_data.csv --format json

# Treat anything older than 60 days as stale, using a fixed reference date
python3 datablix.py sample_data.csv --stale-days 60 --today 2026-06-26
```

## Options

| Option | What it does | Default |
| --- | --- | --- |
| `--format` | Output style: `text`, `markdown`, or `json` | `text` |
| `--output` | Write the report to a file instead of the screen | screen |
| `--freshness-only` | Show only the freshness section | off |
| `--glossary-only` | Show only the categorical glossary | off |
| `--max-categories` | Max distinct values for a column to count as categorical | 20 |
| `--top` | How many top values to list per column | 10 |
| `--fresh-days` | Data newer than this many days is "Fresh" | 30 |
| `--stale-days` | Data older than this many days is "Stale" | 90 |
| `--today` | Reference date (YYYY-MM-DD) for freshness math | system date |

## How it works

datablix reads the CSV and gives each column a type by looking at its values:

| Type | How it is detected |
| --- | --- |
| date | At least 80% of non-empty cells match a known date format |
| numeric | At least 80% of non-empty cells parse as numbers ($, %, commas allowed) |
| categorical | Few distinct values (within `--max-categories`) with real repetition |
| text | Many distinct values or long free text (treated as names, IDs, notes) |
| empty | The column has no values |

**Freshness** is computed per date column: the oldest and newest values,
how many days have passed since the newest, an overall status (Fresh,
Aging, or Stale), and the share of rows that are past the stale threshold.

**The glossary** lists each categorical column's values with counts and
percentages, flags missing cells, and groups values that normalize to the
same key (after trimming spaces and ignoring case) so inconsistent spellings
surface automatically.

## Supported date formats

`YYYY-MM-DD`, `YYYY/MM/DD`, `DD-MM-YYYY`, `DD/MM/YYYY`, `MM/DD/YYYY`,
`MM-DD-YYYY`, `YYYY-MM-DD HH:MM:SS`, ISO timestamps, `12 Jan 2026`,
`January 12, 2026`, and `YYYYMMDD`.

## Limitations

- Designed for CSV input. Convert other formats to CSV first.
- Date guessing is format-based, so an unusual format may be read as text.
- Inconsistency detection catches case and whitespace differences, not
  spelling mistakes or synonyms (for example "BC" versus "British Columbia").

## License

MIT. See `LICENSE`.
