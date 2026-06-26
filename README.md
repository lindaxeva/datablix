# Datablix

Meet Datablix. You hand it a CSV file (that is a spreadsheet saved as plain
text), and it answers two questions you probably should be asking but never
quite have time for:

1. **Is this data still fresh, or has it gone off?**
2. **What is actually in each column, and is anything secretly typed five
   different ways?**

The best part: Datablix only uses tools that already come with Python, so
there is nothing to install. If you have Python 3.8 or newer, you are good to
go. And if the command line makes you nervous, relax, there is a web version
that runs right in your browser.

**Want to just try it?** Here you go: https://lindaxeva.github.io/datablix/

## Why you might want this

Here is a scene you may recognize. You open a spreadsheet you have not touched
in months to send out an update. It looks fine. It always looks fine. But half
the entries are quietly out of date, and nothing on the screen warns you. You
find out the hard way, right after you hit send.

Or you try to count how many people are marked "Active". The total comes out
weird. Turns out one row says "Active", another says "active", and a third
went all in with "ACTIVE". To the spreadsheet, those are three completely
different things, so your count splits across all of them. The spreadsheet is
not wrong, exactly. It is just being extremely literal at the worst possible
moment.

This stuff is easy to do and almost impossible to catch by eye. So let Datablix
do the squinting. It reads the file, tells you how old the data is, and flags
the values that are really the same thing wearing different hats. One step,
plain words, no detective work required.

## Getting started

There is nothing to install, so this is quick. Copy the project to your
computer and run it:

```bash
git clone https://github.com/lindaxeva/datablix.git
cd datablix
python3 datablix.py sample_data.csv
```

The `sample_data.csv` that comes with it has a few old dates and a few
mismatched spellings planted on purpose, like a tiny crime scene, so you can
watch Datablix solve it on your very first run.

## How to use it

The shape of every command is the same:

```bash
python3 datablix.py YOUR_FILE.csv [options]
```

And here are the things you will actually reach for:

```bash
# The full report (freshness and categories)
python3 datablix.py sample_data.csv

# Just the freshness check
python3 datablix.py sample_data.csv --freshness-only

# Just the list of categories
python3 datablix.py sample_data.csv --glossary-only

# Save it as a Markdown file you can share
python3 datablix.py sample_data.csv --format markdown --output report.md

# Get it as JSON, if you want to feed it into something else
python3 datablix.py sample_data.csv --format json

# Count anything older than 60 days as old, measured from a set date
python3 datablix.py sample_data.csv --stale-days 60 --today 2026-06-26
```

## The options, if you want to tweak things

| Option | What it does | Default |
| --- | --- | --- |
| `--format` | Pick the output style: `text`, `markdown`, or `json` | `text` |
| `--output` | Save the report to a file instead of showing it on screen | screen |
| `--freshness-only` | Show only the freshness check | off |
| `--glossary-only` | Show only the list of categories | off |
| `--max-categories` | How many different values a column can have and still count as a category | 20 |
| `--top` | How many of the most common values to show per column | 10 |
| `--fresh-days` | Data newer than this many days counts as fresh | 30 |
| `--stale-days` | Data older than this many days counts as old | 90 |
| `--today` | The date to measure freshness from (defaults to today) | today |

## What it is doing behind the scenes

First, Datablix takes a look at each column and figures out what kind of value
it holds:

| Kind | How it decides |
| --- | --- |
| date | Most of the values look like dates |
| number | Most of the values are numbers (it ignores $, %, and commas) |
| category | Only a few different values, and they keep repeating |
| text | Lots of different values or long text, like names or notes |
| empty | The column has nothing in it |

Then, for the **freshness check**, it goes through each date column and tells
you the oldest and newest dates, how many days have gone by since that newest
one, an overall label (Fresh, Aging, or Old), and how many rows have slipped
past the cutoff for old.

And for the **list of categories**, it shows you each category column's values
and how often each one turns up, points out any blank cells, and quietly
groups together the values that are really the same once you stop caring about
capital letters and stray spaces. That is how the mismatched spellings give
themselves away.

## The web version (no command line, promise)

If typing commands is not your thing, the project includes a browser version
anyone can use:

- `index.html` is the page you open.
- `datablix.js` runs the same checks as the Python tool, just written for the
  browser.

Open `index.html` in any browser, or use the online version at
https://lindaxeva.github.io/datablix/. Want your own copy online for free? In
the project on GitHub, open **Settings**, then **Pages**, point the source at
the `main` branch, and save. Give it a minute and it goes live at
`https://lindaxeva.github.io/datablix/`.

One thing worth knowing: the web version does everything inside your own
browser. Whatever file you open stays on your device and is never sent
anywhere. No uploads, no servers, no funny business.

## Dates it can read

`YYYY-MM-DD`, `YYYY/MM/DD`, `DD-MM-YYYY`, `DD/MM/YYYY`, `MM/DD/YYYY`,
`MM-DD-YYYY`, `YYYY-MM-DD HH:MM:SS`, ISO timestamps, `12 Jan 2026`,
`January 12, 2026`, and `YYYYMMDD`.

## Want to check it really works?

It comes with its own set of tests, and again, nothing to install. From the
project folder, run:

```bash
python3 -m unittest -v
```

Every test should come back with `OK`. These same tests also run on their own
over on GitHub every time you upload a change, and you will see a little
checkmark on the project page when they pass.

## A few honest limits

- It works with CSV files, so save other formats as CSV first.
- It spots dates by their format, so a really unusual date style might get read
  as plain text.
- It catches spacing and capital-letter mismatches, but not typos or different
  words for the same thing. It will not guess that "BC" and "British Columbia"
  are the same place. It is thorough, not psychic.

## License

MIT. See `LICENSE`.
