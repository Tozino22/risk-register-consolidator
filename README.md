# Risk Register Consolidator

A small, dependency-light tool that ingests **multiple risk registers** (Excel files with
different layouts and scoring scales) into **one normalized staging register**, converts
scores to a master 1–5 scale, and **flags likely duplicate risks** for human review.

The output is shaped for onward import into GRC platforms such as **Drata**
(Risk Management → Risks → *Upload a file*), Vanta, or any tool that accepts CSV.

Built by **Tosin Aina — Northbridge**. MIT licensed — use it, share it, adapt it.

---

## Why this exists

Organizations accumulate risk registers the way attics accumulate boxes: an IT security
spreadsheet on a 5×5 scale, an ops log rated High/Medium/Low, a vendor list on 3×3,
audit findings in Word. Consolidating them by hand means days of copy-paste and
silent mistakes. This tool automates the mechanical part — **moving and matching** —
so the humans can focus on the judgment part: **what scores mean, and what to merge**.

## What it does / doesn't do

| Automated by this tool | Stays human (on purpose) |
|---|---|
| Reading every register per a mapping you define once | Agreeing what "High" means with the risk owner |
| Normalizing columns into one staging table | Rewriting risks into a clear standard format |
| Converting 3×3 / 4×4 / High-Med-Low to a 1–5 starting point | Confirming converted scores (rows are flagged) |
| Flagging similar risks as duplicate candidates | Deciding what actually merges |
| Traceability (source register + reference on every row) | Treatment decisions |

## Quick start

Requires Python 3.9+ and `openpyxl`:

```bash
pip install openpyxl
python consolidate_risks.py init
```

`init` creates:
- `mapping.xlsx` — one row per source register: file name, which column letter holds
  the title / description / scores / owner, and the scale used (`5x5`, `4x4`, `3x3`, or `HML`)
- `sources/` — drop your register files in here

Fill in `mapping.xlsx` (row 2 holds hints; two example rows are included), then:

```bash
python consolidate_risks.py run
```

`run` produces **`staging_output.xlsx`**:
- **STAGING** — every risk, one row each, with sequential IDs (R-001…), converted 1–5
  scores, owner, source traceability. Converted rows are marked **"Needs owner review"**;
  possible duplicates are shaded amber.
- **DEDUPE CANDIDATES** — pairs of similar risks with a similarity score and a blank
  *Decision* column for your review session.
- **SUMMARY** — counts per register, review flags, duplicate pairs.

## Try it with the examples

The `examples/` folder contains two intentionally mismatched registers — a clean IT
register on a 5×5 scale and a messy ops log rated High/Medium/Low — plus a filled-in
mapping file. Copy them into place and run:

```bash
copy examples\mapping_example.xlsx mapping.xlsx
copy examples\sources\* sources\
python consolidate_risks.py run
```

You should see 10 risks staged, converted scores flagged for review, and the
laptop/cyber overlaps appear in DEDUPE CANDIDATES.

## The full method (the part the tool can't do)

1. **Inventory** every register before merging anything.
2. **Agree scale conversions with the risk owners** — never mechanically. "Low" often
   means *rare but severe*: one word hiding two numbers.
3. Run this tool to **stage everything** with traceability.
4. **Review duplicate candidates together** — merging is combining knowledge, not deleting rows.
5. Validate scores and owners in short workshops; then export to your GRC platform's
   import template.
6. Make the old registers **read-only** and declare the new register the single source of truth.

## Drata import notes

Drata's bulk import (Risk Management → Risks → *Upload a file*) is a guided CSV flow:
you map your columns to their fields on screen and fix validation errors before
submitting. Key rules (as of mid-2026): Title ≤ 191 chars, status `active`/`closed`/`archived`,
treatment one of Untreated/Accept/Transfer/Avoid/Mitigate, scores 1–5 (Standard tier),
owners must already exist as Drata users with Risk Manager or Admin roles, dates
`yyyy-mm-dd`. Leaving Risk ID blank creates new risks; re-importing with Drata's IDs
updates them in place.

## License

MIT — see [LICENSE](LICENSE).
