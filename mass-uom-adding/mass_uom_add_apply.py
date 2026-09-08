# -*- coding: utf-8 -*-
"""
mass_uom_add_apply.py
=====================
Apply the APPROVED preview (mass_uom_add_preview_<ts>.xlsx, sheet Rows,
Action = ADD) to imported_item_codes.xlsx.

For every ADD row a new UOM-conversion row is inserted into the code's group,
immediately before its CTN row.  The new row is cloned from the code's own
first row (item attributes and system columns stay identical) and only the
UOM block is replaced:
    UOM = the added UOM      UOMDesc = provisional text from the preview
    Rate = SizeRate = preview Rate      UserUOM = PKTS      ADJdedcut = T

Safety: the master is first copied to a timestamped backup, the result is
written to a staged file (imported_item_codes_massUOM_<ts>.xlsx) and validated,
then promoted over the live master.

Run from project root:  python python_scripts/mass_uom_add_apply.py <preview.xlsx>
"""

import os
import re
import sys
import shutil
import datetime
from collections import defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER = os.path.join(ROOT, "imported_item_codes.xlsx")
HEADERS = None  # filled from master row 1


def uom_digits(uom):
    m = re.match(r"PKT\s*(\d+)", str(uom), re.I)
    return int(m.group(1)) if m else None


def sort_key(uom):
    d = uom_digits(uom)
    return (0, d) if d is not None else (1, str(uom))


def read_adds(preview_path):
    wb = openpyxl.load_workbook(preview_path, read_only=True, data_only=True)
    ws = wb["Rows"]
    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    gi = {h: i for i, h in enumerate(hdr)}
    adds = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[gi["Action"]] != "ADD":
            continue
        code = row[gi["ItemCode"]]
        uom = (row[gi["UOM"]] or "").strip().upper()
        rate = row[gi["Rate"]]
        ud = row[gi["Provisional_UOMDesc"]]
        if rate in (None, "") or not (ud or "").strip():
            print("!! skipping incomplete ADD", code, uom, rate, ud)
            continue
        adds.append({"code": code, "uom": uom, "rate": rate, "uomdesc": str(ud).strip()})
    wb.close()
    return adds


def load_rows():
    wb = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    rows = [r for r in it]
    # preserve the (scratch) Sheet3 used range, cell-by-cell
    sheet3 = []
    if "Sheet3" in wb.sheetnames:
        w3 = wb["Sheet3"]
        for rr in w3.iter_rows(values_only=True):
            sheet3.append(list(rr))
    wb.close()
    return header, rows, sheet3


def build(header, rows, adds):
    by_code = defaultdict(list)
    for i, r in enumerate(rows):
        by_code[r[0]].append(i)
    add_groups = defaultdict(list)
    for a in adds:
        if a["code"] not in by_code:
            print("!! code not found in master:", a["code"])
            continue
        add_groups[a["code"]].append(a)

    new_rows = []
    idx = 0
    n = len(rows)
    while idx < n:
        code = rows[idx][0]
        end = idx
        while end < n and rows[end][0] == code:
            end += 1
        group = rows[idx:end]
        existing_uoms = {(r[18] or "").upper() for r in group}
        adds_here = [a for a in sorted(add_groups.pop(code, []), key=lambda a: sort_key(a["uom"]))
                     if a["uom"] not in existing_uoms]
        # find first CTN row position
        ctn_at = next((j for j, r in enumerate(group) if (r[18] or "").upper() == "CTN"), len(group))
        out = group[:ctn_at]
        for a in adds_here:
            base = list(group[0])
            base[2] = a["uomdesc"]        # UOMDesc
            base[15] = a["rate"]          # UDF_UOM_SizeRate
            base[16] = a["rate"]          # Rate
            base[17] = "PKTS"             # UDF_UOM_UserUOM
            base[18] = a["uom"]           # UOM
            base[19] = "T"                # UDF_UOM_ADJdedcut
            out.append(tuple(base))
        out.extend(group[ctn_at:])
        new_rows.extend(out)
        idx = end

    if add_groups:
        print("!! codes with adds that were never placed:", list(add_groups))
    return new_rows


def validate(rows):
    """Return the set of (code, UOM) pairs that appear more than once in a
    group.  The live master already carries a handful of such anomalies, so the
    caller compares before/after and only rejects NEW ones."""
    g = defaultdict(list)
    for r in rows:
        g[r[0]].append(r)
    dups = set()
    for code, grp in g.items():
        cnt = defaultdict(int)
        for r in grp:
            cnt[(r[18] or "").upper()] += 1
        for (u, c) in cnt.items():
            if c > 1:
                dups.add((code, u))
    return dups


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    preview = os.path.join(ROOT, sys.argv[1]) if not os.path.isabs(sys.argv[1]) else sys.argv[1]
    adds = read_adds(preview)
    print("ADD rows to apply:", len(adds))

    stamp = datetime.datetime.now().strftime("%b-%d_%I%M%p").lower()
    backup = os.path.join(ROOT, "imported_item_codes_backup_%s.xlsx" % stamp)
    shutil.copyfile(MASTER, backup)
    print("backup written:", os.path.basename(backup))

    # merge duplicate instructions: a code gets each UOM only once
    seen = set()
    uniq = []
    for a in adds:
        k = (a["code"], a["uom"])
        if k in seen:
            print("  merged duplicate instruction:", k)
            continue
        seen.add(k)
        uniq.append(a)
    adds = uniq
    print("unique (code,UOM) additions:", len(adds))

    header, rows, sheet3 = load_rows()
    print("master rows before:", len(rows))
    dups_before = validate(rows)
    print("pre-existing duplicate (code,UOM) pairs:", len(dups_before), sorted(dups_before)[:8])
    new_rows = build(header, rows, adds)
    print("master rows after :", len(new_rows))
    dups_after = validate(new_rows)
    new_dups = dups_after - dups_before
    print("NEW duplicate (code,UOM) pairs:", len(new_dups), sorted(new_dups))
    if new_dups:
        print("!! ABORT: new duplicates would be written; nothing was changed.")
        return 3

    staged = os.path.join(ROOT, "imported_item_codes_massUOM_%s.xlsx" % stamp)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(list(header))
    for r in new_rows:
        ws.append(list(r))
    # preserve the scratch Sheet3 cells (values only)
    ws3 = wb.create_sheet("Sheet3")
    for ri, rr in enumerate(sheet3, start=1):
        for ci, v in enumerate(rr, start=1):
            if v is not None:
                ws3.cell(row=ri, column=ci, value=v)
    wb.save(staged)
    print("staged file written:", os.path.basename(staged))

    # promote to live master
    try:
        shutil.copyfile(staged, MASTER)
        print("PROMOTED to live master: imported_item_codes.xlsx")
    except Exception as e:
        print("!! could not promote (is the master open in Excel?):", e)
        print("   staged file is ready at:", os.path.basename(staged))
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
