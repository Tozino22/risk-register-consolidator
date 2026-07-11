#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Risk Register Consolidator
==========================
Ingests multiple risk registers (Excel) into one normalized staging register,
converts scoring scales, and flags likely duplicate risks for human review.
Output is shaped for onward import into GRC platforms such as Drata.

Built by Tosin Aina - Northbridge (https://github.com/Tozino22)
License: MIT

Usage:
    python consolidate_risks.py init     # create mapping.xlsx template + sources/ folder
    python consolidate_risks.py automap  # scan sources/ and auto-guess the mapping (confirm before running!)
    python consolidate_risks.py run      # ingest everything per mapping.xlsx -> staging_output.xlsx

Requires: Python 3.9+ and openpyxl  (pip install openpyxl)
"""

import os
import sys
from difflib import SequenceMatcher

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")

MAPPING_FILE = "mapping.xlsx"
SOURCES_DIR = "sources"
OUTPUT_FILE = "staging_output.xlsx"
DEDUPE_THRESHOLD = 0.5  # similarity score above which a pair is flagged for human review

TEAL = "0B8F70"
AMBER = "FFF3D6"

# Discrete-scale conversions to the master 1-5 scale.
# These are STARTING POINTS - every converted row is flagged "Needs owner review".
SCALE_MAPS = {
    "3x3": {1: 1, 2: 3, 3: 5},
    "4x4": {1: 1, 2: 3, 3: 4, 4: 5},
    "5x5": {1: 1, 2: 2, 3: 3, 4: 4, 5: 5},
}
# Qualitative ratings -> (likelihood, impact) starting points
HML_MAP = {
    "HIGH": (4, 4), "H": (4, 4),
    "MEDIUM": (3, 3), "MED": (3, 3), "M": (3, 3),
    "LOW": (2, 2), "L": (2, 2),
    "CRITICAL": (5, 5), "SEVERE": (5, 5), "VERY HIGH": (5, 5),
    "VERY LOW": (1, 1), "NEGLIGIBLE": (1, 1),
}

MAPPING_HEADERS = [
    ("register_name", 26, "Friendly name, e.g. 'IT Security Register'"),
    ("file", 30, "File name inside the sources/ folder"),
    ("sheet", 16, "Sheet name (blank = first sheet)"),
    ("header_row", 12, "Row number of the header (default 1)"),
    ("col_title", 10, "Column LETTER holding the risk title"),
    ("col_description", 16, "Column letter for description (blank = reuse title)"),
    ("col_likelihood", 14, "Column letter for likelihood (numeric scales)"),
    ("col_impact", 12, "Column letter for impact (numeric scales)"),
    ("col_rating", 12, "Column letter for a single High/Med/Low rating"),
    ("col_owner", 12, "Column letter for owner / assigned to"),
    ("col_ref", 10, "Column letter for the source's own ID (blank = row number)"),
    ("scale", 10, "One of: 5x5, 4x4, 3x3, HML"),
]

STAGING_HEADERS = [
    "Risk ID", "Title", "Description", "Category", "Source Register", "Source Ref",
    "Likelihood (1-5)", "Impact (1-5)", "Inherent Score", "Owner",
    "Treatment", "Status", "Duplicate Of", "Notes",
]


def cmd_init():
    if os.path.exists(MAPPING_FILE):
        sys.exit(f"{MAPPING_FILE} already exists - delete it first if you want a fresh one.")
    wb = Workbook()
    ws = wb.active
    ws.title = "mapping"
    for i, (name, width, _) in enumerate(MAPPING_HEADERS, 1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=TEAL)
        ws.column_dimensions[get_column_letter(i)].width = width
    for i, (_, _, hint) in enumerate(MAPPING_HEADERS, 1):
        c = ws.cell(row=2, column=i, value=hint)
        c.font = Font(italic=True, size=9, color="7C8B86")
    # example rows (delete before running)
    ws.append(["IT Security Register", "IT_Security_Register.xlsx", "", 1,
               "B", "C", "D", "E", "", "G", "A", "5x5"])
    ws.append(["Ops Risk Log", "Ops_Risk_Log.xlsx", "", 1,
               "B", "", "", "", "C", "D", "A", "HML"])
    wb.save(MAPPING_FILE)
    os.makedirs(SOURCES_DIR, exist_ok=True)
    print(f"Created {MAPPING_FILE} (row 2 = hints, rows 3-4 = examples; replace with your registers)")
    print(f"Created {SOURCES_DIR}/ - drop your register files in there, then run:")
    print("    python consolidate_risks.py run")


def _col(letter):
    return column_index_from_string(letter.strip().upper()) if letter and str(letter).strip() else None


def _convert_numeric(value, scale):
    try:
        v = int(round(float(value)))
    except (TypeError, ValueError):
        return None
    mapped = SCALE_MAPS.get(scale, {}).get(v)
    return mapped if mapped else max(1, min(5, v))


def read_mapping():
    wb = load_workbook(MAPPING_FILE, data_only=True)
    ws = wb["mapping"] if "mapping" in wb.sheetnames else wb.active
    keys = [h[0] for h in MAPPING_HEADERS]
    entries = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(keys, row))
        name = str(rec.get("register_name") or "").strip()
        file_ = str(rec.get("file") or "").strip()
        # skip hint row and blanks
        if not name or not file_ or file_.lower().startswith("file name"):
            continue
        entries.append(rec)
    return entries


def ingest(entry):
    path = os.path.join(SOURCES_DIR, str(entry["file"]).strip())
    if not os.path.exists(path):
        print(f"  !! SKIPPED - file not found: {path}")
        return []
    wb = load_workbook(path, data_only=True)
    sheet = str(entry.get("sheet") or "").strip()
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
    header_row = int(entry.get("header_row") or 1)
    scale = str(entry.get("scale") or "5x5").strip().upper()

    c_title = _col(entry.get("col_title"))
    c_desc = _col(entry.get("col_description"))
    c_like = _col(entry.get("col_likelihood"))
    c_imp = _col(entry.get("col_impact"))
    c_rate = _col(entry.get("col_rating"))
    c_owner = _col(entry.get("col_owner"))
    c_ref = _col(entry.get("col_ref"))

    records = []
    for r in range(header_row + 1, ws.max_row + 1):
        title = ws.cell(row=r, column=c_title).value if c_title else None
        if title is None or not str(title).strip():
            continue
        title = str(title).strip()
        desc = str(ws.cell(row=r, column=c_desc).value or title).strip() if c_desc else title
        owner = str(ws.cell(row=r, column=c_owner).value or "").strip() if c_owner else ""
        ref = str(ws.cell(row=r, column=c_ref).value or f"Row {r}").strip() if c_ref else f"Row {r}"

        likelihood = impact = None
        needs_review = False
        note = ""

        if scale == "HML" and c_rate:
            raw = str(ws.cell(row=r, column=c_rate).value or "").strip().upper()
            pair = HML_MAP.get(raw)
            if pair:
                likelihood, impact = pair
                needs_review = True
                note = f"Converted from '{raw}' - starting point only, confirm with owner"
            else:
                note = f"Unrecognized rating '{raw}' - score manually"
                needs_review = True
        else:
            if c_like:
                likelihood = _convert_numeric(ws.cell(row=r, column=c_like).value, scale)
            if c_imp:
                impact = _convert_numeric(ws.cell(row=r, column=c_imp).value, scale)
            if scale != "5x5" and (likelihood or impact):
                needs_review = True
                note = f"Converted from {scale} - confirm with owner"

        records.append({
            "title": title, "description": desc,
            "register": str(entry["register_name"]).strip(), "ref": ref,
            "likelihood": likelihood, "impact": impact, "owner": owner,
            "status": "Needs owner review" if needs_review else "In staging",
            "note": note,
        })
    print(f"  {entry['register_name']}: {len(records)} risks ingested")
    return records


STOPWORDS = {"the", "a", "an", "of", "to", "and", "or", "in", "on", "at", "with", "for",
             "our", "we", "us", "them", "they", "it", "is", "are", "get", "gets", "cannot",
             "not", "by", "via", "from", "into", "risk", "that"}


def _stem(word):
    for suffix in ("ing", "ed", "es", "s"):
        if len(word) > 4 and word.endswith(suffix):
            return word[: -len(suffix)]
    return word


def normalize(text):
    words = "".join(ch.lower() for ch in text if ch.isalnum() or ch.isspace()).split()
    return [_stem(w) for w in words if w not in STOPWORDS]


def _overlap(a, b):
    sa, sb = set(a), set(b)
    return len(sa & sb) / min(len(sa), len(sb)) if sa and sb else 0.0


def similarity(rec_a, rec_b):
    """Max of: sequence similarity on full text, content-word overlap on full
    text, and content-word overlap on titles alone. Titles are compared
    separately because long descriptions dilute overlap on the same concept."""
    full_a = normalize(rec_a["title"] + " " + rec_a["description"])
    full_b = normalize(rec_b["title"] + " " + rec_b["description"])
    seq = SequenceMatcher(None, " ".join(full_a), " ".join(full_b)).ratio()
    return max(seq,
               _overlap(full_a, full_b),
               _overlap(normalize(rec_a["title"]), normalize(rec_b["title"])))


def find_dupes(records):
    pairs = []
    for i in range(len(records)):
        for j in range(i + 1, len(records)):
            score = similarity(records[i], records[j])
            if score >= DEDUPE_THRESHOLD:
                pairs.append((i, j, score))
    return sorted(pairs, key=lambda p: -p[2])


# --- automap: guess column mappings from headers + data samples ---

def _header_row(ws):
    for r in range(1, min(6, ws.max_row + 1)):
        texts = [c for c in next(ws.iter_rows(min_row=r, max_row=r, values_only=True))
                 if isinstance(c, str) and c.strip()]
        if len(texts) >= 2:
            return r
    return 1


def _sample_values(ws, col, header_row, n=10):
    vals = []
    for r in range(header_row + 1, min(header_row + 1 + n, ws.max_row + 1)):
        v = ws.cell(row=r, column=col).value
        if v is not None and str(v).strip():
            vals.append(v)
    return vals


def _mostly_numeric(vals):
    if not vals:
        return False
    hits = sum(1 for v in vals if str(v).strip().replace(".", "", 1).isdigit())
    return hits >= max(1, len(vals) * 0.7)


def _mostly_hml(vals):
    if not vals:
        return False
    hits = sum(1 for v in vals if str(v).strip().upper() in HML_MAP)
    return hits >= max(1, len(vals) * 0.7)


def guess_mapping(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hr = _header_row(ws)
    guess = {"sheet": "", "header_row": hr, "col_title": "", "col_description": "",
             "col_likelihood": "", "col_impact": "", "col_rating": "",
             "col_owner": "", "col_ref": "", "scale": ""}
    numeric_max = 0

    headers = next(ws.iter_rows(min_row=hr, max_row=hr, values_only=True))
    for idx, h in enumerate(headers, 1):
        if not isinstance(h, str) or not h.strip():
            continue
        name = h.strip().lower()
        letter = get_column_letter(idx)
        vals = _sample_values(ws, idx, hr)

        if any(k in name for k in ("descript", "detail")):
            guess["col_description"] = guess["col_description"] or letter
        elif any(k in name for k in ("likelihood", "probability", "prob")):
            guess["col_likelihood"] = guess["col_likelihood"] or letter
            numeric_max = max(numeric_max, *[int(float(v)) for v in vals if str(v).strip().replace(".", "", 1).isdigit()] or [0])
        elif any(k in name for k in ("impact", "consequence")):
            guess["col_impact"] = guess["col_impact"] or letter
            numeric_max = max(numeric_max, *[int(float(v)) for v in vals if str(v).strip().replace(".", "", 1).isdigit()] or [0])
        elif any(k in name for k in ("severity", "rating", "risk level", "priority")):
            if _mostly_hml(vals):
                guess["col_rating"] = guess["col_rating"] or letter
            elif _mostly_numeric(vals):
                guess["col_impact"] = guess["col_impact"] or letter
                numeric_max = max(numeric_max, *[int(float(v)) for v in vals if str(v).strip().replace(".", "", 1).isdigit()] or [0])
        elif any(k in name for k in ("owner", "assigned", "responsible", "who")):
            guess["col_owner"] = guess["col_owner"] or letter
        elif name in ("id", "ref", "no", "#", "number") or name.endswith(" id"):
            guess["col_ref"] = guess["col_ref"] or letter
        elif any(k in name for k in ("risk", "finding", "issue", "title", "name")):
            guess["col_title"] = guess["col_title"] or letter

    if guess["col_rating"] and not (guess["col_likelihood"] or guess["col_impact"]):
        guess["scale"] = "HML"
    elif numeric_max and numeric_max <= 3:
        guess["scale"] = "3x3"
    elif numeric_max == 4:
        guess["scale"] = "4x4"
    else:
        guess["scale"] = "5x5"
    return guess


def cmd_automap():
    if not os.path.isdir(SOURCES_DIR):
        sys.exit(f"No {SOURCES_DIR}/ folder found. Run:  python consolidate_risks.py init")
    files = [f for f in sorted(os.listdir(SOURCES_DIR)) if f.lower().endswith((".xlsx", ".xlsm"))]
    if not files:
        sys.exit(f"No Excel files found in {SOURCES_DIR}/")
    target = MAPPING_FILE if not os.path.exists(MAPPING_FILE) else "mapping_auto.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "mapping"
    for i, (name, width, _) in enumerate(MAPPING_HEADERS, 1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=TEAL)
        ws.column_dimensions[get_column_letter(i)].width = width

    print(f"Scanning {len(files)} file(s) in {SOURCES_DIR}/ ...")
    for f in files:
        g = guess_mapping(os.path.join(SOURCES_DIR, f))
        register_name = os.path.splitext(f)[0].replace("_", " ")
        ws.append([register_name, f, g["sheet"], g["header_row"], g["col_title"],
                   g["col_description"], g["col_likelihood"], g["col_impact"],
                   g["col_rating"], g["col_owner"], g["col_ref"], g["scale"]])
        missing = " <- CHECK: no title column found!" if not g["col_title"] else ""
        print(f"  {f}: title={g['col_title'] or '?'} scale={g['scale']}{missing}")

    wb.save(target)
    print(f"\nGuesses written to {target}")
    print("IMPORTANT: open it and CONFIRM every row before running - guesses are starting points.")
    if target != MAPPING_FILE:
        print(f"(An existing {MAPPING_FILE} was left untouched - compare and merge manually.)")


def cmd_run():
    if not os.path.exists(MAPPING_FILE):
        sys.exit(f"No {MAPPING_FILE} found. Run:  python consolidate_risks.py init")
    entries = read_mapping()
    if not entries:
        sys.exit("mapping.xlsx has no usable rows - fill in one row per source register.")
    print(f"Ingesting {len(entries)} register(s) from {SOURCES_DIR}/ ...")
    records = []
    for entry in entries:
        records.extend(ingest(entry))
    if not records:
        sys.exit("No risks ingested - check file names and column letters in mapping.xlsx")

    for n, rec in enumerate(records, 1):
        rec["id"] = f"R-{n:03d}"
    dupes = find_dupes(records)

    wb = Workbook()
    ws = wb.active
    ws.title = "STAGING"
    widths = [9, 34, 52, 14, 22, 10, 14, 12, 12, 20, 12, 18, 12, 40]
    for i, (h, w) in enumerate(zip(STAGING_HEADERS, widths), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=TEAL)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    flagged = {i for pair in dupes for i in pair[:2]}
    for idx, rec in enumerate(records):
        r = idx + 2
        score = rec["likelihood"] * rec["impact"] if rec["likelihood"] and rec["impact"] else ""
        row = [rec["id"], rec["title"], rec["description"], "", rec["register"], rec["ref"],
               rec["likelihood"] or "", rec["impact"] or "", score, rec["owner"],
               "", rec["status"], "", rec["note"]]
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if idx in flagged:
                cell.fill = PatternFill("solid", fgColor=AMBER)

    ws2 = wb.create_sheet("DEDUPE CANDIDATES")
    for i, (h, w) in enumerate(zip(["Risk A", "Title A", "Risk B", "Title B", "Similarity", "Decision (keep / merge into / not duplicates)"],
                                   [9, 40, 9, 40, 12, 34]), 1):
        c = ws2.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=TEAL)
        ws2.column_dimensions[get_column_letter(i)].width = w
    for k, (i, j, score) in enumerate(dupes, 2):
        ws2.append([records[i]["id"], records[i]["title"],
                    records[j]["id"], records[j]["title"], round(score, 2), ""])

    ws3 = wb.create_sheet("SUMMARY")
    ws3.column_dimensions["A"].width = 40
    ws3.column_dimensions["B"].width = 14
    per = {}
    for rec in records:
        per[rec["register"]] = per.get(rec["register"], 0) + 1
    rows = [("Registers ingested", len(entries)), ("Total risks staged", len(records)),
            ("Rows needing owner review (converted scores)", sum(1 for r in records if r["status"] == "Needs owner review")),
            ("Duplicate candidate pairs flagged", len(dupes))] + sorted(per.items())
    for a, b in rows:
        ws3.append([a, b])

    wb.save(OUTPUT_FILE)
    print(f"\nDone -> {OUTPUT_FILE}")
    print(f"  STAGING: {len(records)} risks (amber rows = possible duplicates)")
    print(f"  DEDUPE CANDIDATES: {len(dupes)} pair(s) to review by hand")
    print("  Next: review scores with owners, resolve duplicates, then map to your GRC import template.")


if __name__ == "__main__":
    cmd = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    if cmd == "init":
        cmd_init()
    elif cmd == "automap":
        cmd_automap()
    elif cmd == "run":
        cmd_run()
    else:
        print(__doc__)
