# -*- coding: utf-8 -*-
"""
concatenate_dckj_po.py
======================
Concatenate every DCKJ purchase-order CSV in
`DCKJ_PO_sept_2025_to_sept_2026/` vertically into one table:

    purchase_orders_b2b/sept_2025_to_2026_concatenated.xlsx

Each source file is a batch export with the same shape:

    line 1   START,,,,...                     <- marker, dropped
    line 2   PO_NUMBER,ORG_NUMBER,...         <- the real header (24 cols)
    line 3+  data
    last     END,,,,...                       <- marker, dropped

`PO_NUMBER` appears TWICE in the header (column 1 and column 10); that is how
the export ships and it is preserved as-is, so the table keeps its 24 columns.

Values are written **exactly as exported, as text**. That is deliberate: the
source mixes representations that coercion would silently damage --
`SEQ_NUMBER` is zero-padded ('0001'), `BARCODE` has values already flattened to
scientific notation by the export ('1.92215E+11'), and `ITEM_TOTAL_EXCL_GST`
carries thousands separators ('1,000.000'). Nothing is invented or truncated.

Files are concatenated in filename order, which is their export order.

Usage:
    python purchase_orders_b2b/concatenate_dckj_po.py
"""

import csv
import glob
import os
import re
import collections

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, "purchase_orders_b2b", "DCKJ_PO_sept_2025_to_sept_2026")
OUT = os.path.join(ROOT, "purchase_orders_b2b", "sept_2025_to_2026_concatenated.xlsx")

SCI = re.compile(r"^-?\d+(\.\d+)?[eE][+-]?\d+$")


def read_batch(path):
    """-> (header, data_rows) for one batch export."""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    if not rows or rows[0][0].strip().upper() != "START":
        raise SystemExit("ABORT: %s does not start with START" % path)
    header = rows[1]
    data = []
    for row in rows[2:]:
        if row and row[0].strip().upper() == "END":
            break
        if not any(c.strip() for c in row):
            continue
        if len(row) != len(header):
            raise SystemExit("ABORT: %s has a %d-field row, header has %d"
                             % (path, len(row), len(header)))
        data.append(row)
    return header, data


def main():
    files = sorted(glob.glob(os.path.join(SRC_DIR, "*.csv")))
    if not files:
        raise SystemExit("ABORT: no CSVs in %s" % SRC_DIR)

    header = None
    table = []
    per_file = []
    for path in files:
        h, data = read_batch(path)
        if header is None:
            header = h
        elif h != header:
            raise SystemExit("ABORT: %s header differs from the first file"
                             % os.path.basename(path))
        table.extend(data)
        per_file.append((os.path.basename(path), len(data)))

    ncol = len(header)
    print("files read        : %d" % len(files))
    print("columns           : %d" % ncol)
    print("data rows         : %d" % len(table))

    # ---- integrity: exact duplicate rows ---------------------------------
    exact = collections.Counter(tuple(r) for r in table)
    dup_rows = sum(v - 1 for v in exact.values() if v > 1)
    dup_keys = collections.Counter((r[0], r[9], r[11], r[12]) for r in table)
    print("exact duplicate rows: %d (across %d POs)"
          % (dup_rows, len({k[0] for k, v in dup_keys.items() if v > 1})))

    # ---- source hazards, reported not "fixed" ---------------------------
    bi = header.index("BARCODE")
    ti = header.index("ITEM_TOTAL_EXCL_GST")
    gi = header.index("GST_AMOUNT")
    sci = [r[bi] for r in table if SCI.match(r[bi].strip())]
    comma = [r[ti] for r in table if "," in r[ti]]
    gcomma = [r[gi] for r in table if "," in r[gi]]
    print("BARCODE in scientific notation : %d  e.g. %s"
          % (len(sci), sorted(set(sci))[:3]))
    print("ITEM_TOTAL_EXCL_GST with commas: %d  e.g. %s"
          % (len(comma), sorted(set(comma))[:2]))
    print("GST_AMOUNT with commas         : %d" % len(gcomma))

    ed = header.index("ENTRY_DATE")
    months = collections.Counter(r[ed][:6] for r in table if r[ed].strip())
    print("ENTRY_DATE months : %s -> %s"
          % (min(months), max(months)))

    # ---- write ----------------------------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PO"
    ws.append(header)
    for c in range(1, ncol + 1):
        ws.cell(1, c).font = Font(bold=True)
    for row in table:
        ws.append(row)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(ncol), len(table) + 1)
    for c, name in enumerate(header, 1):
        width = max(len(str(name)), *(len(str(r[c - 1])) for r in table[:400]))
        ws.column_dimensions[get_column_letter(c)].width = min(max(width + 2, 8), 46)
    wb.save(OUT)
    print()
    print("written: %s" % OUT)
    print("  sheet 'PO': 1 header row + %d data rows, %d columns"
          % (len(table), ncol))


if __name__ == "__main__":
    main()
