# -*- coding: utf-8 -*-
"""
dedupe_dckj_products.py
=======================
Add a `Unique_Products` sheet to

    purchase_orders_b2b/sept_2025_to_2026_concatenated.xlsx

holding the non-duplicate rows of the trimmed `PO` sheet, keyed on the four
product-attribute columns the owner named (2026-09-26, Giant is the customer):

    PRD_NUMBER  Giant's item code
    PRD_DESC    Giant's product description
    PACK_QTY    punnets per carton OR no. of pcs
    BARCODE     barcode

One row survives per distinct 4-key -- the first occurrence in the sheet's own
order -- so every surviving row is verbatim source data, with its original
PO_NUMBER / UOM / ORDER_QTY intact.

`BARCODE` is part of the key as instructed. Note it is exported two ways for
the same product (`9.55515E+12` as well as `9555149202323` -- see
concatenate_dckj_po.py), which splits some products into two keys; the run
report quantifies that, and `--ignore-barcode` produces the product-level
variant (PRD_NUMBER + PRD_DESC + PACK_QTY only) for comparison.

Usage:
    python purchase_orders_b2b/dedupe_dckj_products.py                 # writes Unique_Products
    python purchase_orders_b2b/dedupe_dckj_products.py --ignore-barcode
    python purchase_orders_b2b/dedupe_dckj_products.py --dry-run
"""

import os
import sys
from decimal import Decimal, InvalidOperation
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOK = os.path.join(ROOT, "purchase_orders_b2b", "sept_2025_to_2026_concatenated.xlsx")
SRC_SHEET = "PO"
NEW_SHEET = "Unique_Products"

# column positions in the trimmed PO sheet (1-based)
PO_NUMBER, PRD_NUMBER, PRD_DESC, UOM, ORDER_QTY, PACK_QTY, BARCODE = range(1, 8)
KEY_COLS = [PRD_NUMBER, PRD_DESC, PACK_QTY, BARCODE]


def sci_parts(value):
    """('9.55515E+12') -> (9.55515, 12); returns None if not scientific."""
    try:
        mant, exp = str(value).upper().split("E")
        return round(float(mant), 5), int(exp)
    except (ValueError, AttributeError):
        return None


def plain_parts(value):
    """('9555149202323') -> (9.55515, 12) at the same order of magnitude."""
    try:
        d = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None
    if d == 0:
        return None
    exp = d.adjusted()                      # exponent of the leading digit
    return round(float(d.scaleb(-exp)), 5), exp


def same_barcode(a, b):
    """True when two spellings denote the same barcode, allowing the export's
    scientific-notation rounding (it keeps only 6 significant digits)."""
    if a == b:
        return True
    pa, pb = sci_parts(a), sci_parts(b)
    if (pa is None) == (pb is None):
        return False                        # both plain, or both scientific
    sci, other = (pa, b) if pa is not None else (pb, a)
    return plain_parts(other) == sci


def read_rows(ws):
    return [[ws.cell(r, c).value if ws.cell(r, c).value is not None else ""
             for c in range(1, 8)]
            for r in range(2, ws.max_row + 1)]


def main():
    ignore_barcode = "--ignore-barcode" in sys.argv
    dry = "--dry-run" in sys.argv
    key_cols = [c for c in KEY_COLS if not (ignore_barcode and c == BARCODE)]

    wb = openpyxl.load_workbook(BOOK)
    if SRC_SHEET not in wb.sheetnames:
        raise SystemExit("ABORT: no %r sheet in %s" % (SRC_SHEET, BOOK))
    ws = wb[SRC_SHEET]
    header = [ws.cell(1, c).value for c in range(1, 8)]
    rows = read_rows(ws)
    print("source sheet %r: %d data rows, %d columns" % (SRC_SHEET, len(rows), len(header)))

    def key(row):
        return tuple(row[c - 1] for c in key_cols)

    seen = {}
    for row in rows:
        seen.setdefault(key(row), row)
    unique = list(seen.values())
    if ignore_barcode:
        unique.sort(key=lambda r: (r[PRD_NUMBER - 1], r[PRD_DESC - 1]))
    print("key            : %s" % ", ".join(header[c - 1] for c in key_cols))
    print("distinct rows  : %d   (collapsed %d)" % (len(unique), len(rows) - len(unique)))

    # ---- how much of the split is the barcode's two spellings? -----------
    by4 = defaultdict(list)
    for row in rows:
        by4[tuple(row[c - 1] for c in KEY_COLS)].append(row)
    by3 = defaultdict(set)
    for row in rows:
        by3[tuple(row[c - 1] for c in (PRD_NUMBER, PRD_DESC, PACK_QTY))].add(row[BARCODE - 1])
    split_keys = [(k, v) for k, v in by3.items() if len(v) > 1]
    spelling, genuinely = 0, 0
    for k, variants in split_keys:
        vs = sorted(variants)
        merged = []                               # groups of equivalent spellings
        for v in vs:
            for grp in merged:
                if any(same_barcode(v, o) for o in grp):
                    grp.add(v)
                    break
            else:
                merged.append({v})
        spelling += len(vs) - len(merged)
        genuinely += len(merged) - 1
    print()
    print("distinct (PRD_NUMBER, PRD_DESC, PACK_QTY) ignoring BARCODE: %d" % len(by3))
    print("  extra keys caused by BARCODE differing: %d" % (len(by4) - len(by3)))
    print("     of which the export's sci-notation/plain spelling of ONE barcode: %d" % spelling)
    print("     of which a genuinely different barcode:                           %d" % genuinely)
    if split_keys:
        print("  examples of a product carrying >1 barcode:")
        for k, variants in sorted(split_keys)[:6]:
            print("     %-30s pack=%-5s %s" % (k[1][:30], k[2], sorted(variants)))

    if dry:
        print("\n--dry-run: nothing written")
        return

    # ---- write the sheet, matching the PO sheet's formatting -------------
    if NEW_SHEET in wb.sheetnames:
        del wb[NEW_SHEET]
    out = wb.create_sheet(NEW_SHEET)
    out.append(header)
    for c in range(1, len(header) + 1):
        out.cell(1, c).font = Font(bold=True)
    for row in unique:
        out.append(row)
    out.freeze_panes = "A2"
    out.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(header)), len(unique) + 1)
    for c in range(1, len(header) + 1):
        letter = get_column_letter(c)
        src_w = ws.column_dimensions[letter].width
        out.column_dimensions[letter].width = src_w or 14
    wb.save(BOOK)
    print("\nwritten: %s" % BOOK)
    print("  sheet %r: 1 header row + %d rows" % (NEW_SHEET, len(unique)))


if __name__ == "__main__":
    main()
