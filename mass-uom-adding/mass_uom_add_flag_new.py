# -*- coding: utf-8 -*-
"""
mass_uom_add_flag_new.py
========================
Flag every row that the mass-UOM apply ADDED to imported_item_codes.xlsx.

Rows are identified by (ItemCode, UOM, UOMDesc) taken from the APPROVED preview
(its Action = ADD rows), not by diffing against the backup -- so any later
manual edits to the master (e.g. re-basing the bagged Packham codes to PKT) do
not pollute the output.  Rows are located in the current master.

Output: mass_uom_add_new_rows_<ts>.xlsx  (root).
Run: python python_scripts/mass_uom_add_flag_new.py [preview.xlsx]
"""

import os
import sys
import datetime
from collections import defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER = os.path.join(ROOT, "imported_item_codes.xlsx")


def read_added(preview):
    """{(code, UOM): {'uomdesc','rate','layout':[..],'products':[..]}} from ADD rows."""
    wb = openpyxl.load_workbook(preview, read_only=True, data_only=True)
    ws = wb["Rows"]
    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    gi = {h: i for i, h in enumerate(hdr)}
    added = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[gi["Action"]] != "ADD":
            continue
        code = row[gi["ItemCode"]]
        uom = (row[gi["UOM"]] or "").strip().upper()
        ud = (row[gi["Provisional_UOMDesc"]] or "").strip()
        if not ud:
            continue
        key = (code, uom)
        rec = added.setdefault(key, {"uomdesc": ud, "layout": [], "products": []})
        if row[gi["LayoutRow"]] not in rec["layout"]:
            rec["layout"].append(row[gi["LayoutRow"]])
        if row[gi["Product"]] not in rec["products"]:
            rec["products"].append(row[gi["Product"]])
    wb.close()
    return added


def main():
    preview = os.path.join(ROOT, sys.argv[1]) if len(sys.argv) > 1 else os.path.join(
        ROOT, "mass_uom_add_preview_sep-09_1201am.xlsx")
    added = read_added(preview)
    print("expected added rows:", len(added))

    wb = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    it = ws.iter_rows(values_only=True)
    header = list(next(it))
    found, missing, extra = [], [], []
    seen = set()
    for r in it:
        key = (r[0], (r[18] or "").strip().upper())
        rec = added.get(key)
        if rec and (r[2] or "").strip() == rec["uomdesc"]:
            if key in seen:
                extra.append(r[0:2])
                continue
            seen.add(key)
            found.append((rec, r))
        elif key in added and (r[2] or "").strip() != rec["uomdesc"]:
            missing.append((key, (r[2] or "")))
    wb.close()

    for k, rec in added.items():
        if k not in seen:
            missing.append((k[0], k[1], rec["uomdesc"], "not found in master"))
    print("found in master:", len(found), "| missing:", len(missing), "| duplicated in master:", len(extra))
    for m in missing[:15]:
        print("  MISSING:", m)

    stamp = datetime.datetime.now().strftime("%b-%d_%I%M%p").lower()
    out = os.path.join(ROOT, "mass_uom_add_new_rows_%s.xlsx" % stamp)
    out_wb = openpyxl.Workbook()
    ows = out_wb.active
    ows.title = "NewRows"
    ows.append(["FLAG", "LayoutRows", "Product(s)"] + header)
    for rec, r in found:
        ows.append(["NEW",
                    ";".join(str(x) for x in rec["layout"]),
                    ";".join(rec["products"])] + list(r))
    for i, cw in enumerate([8, 12, 30] + [16] * len(header), start=1):
        ows.column_dimensions[openpyxl.utils.get_column_letter(i)].width = cw
    if missing:
        wsn = out_wb.create_sheet("NotFlagged")
        wsn.append(["ItemCode", "UOM", "Expected UOMDesc", "Status"])
        for code, uom, ud, status in missing:
            wsn.append([code, uom, ud, status])
    out_wb.save(out)
    print("output written:", os.path.basename(out), "| rows flagged:", len(found))
    return 0


if __name__ == "__main__":
    sys.exit(main())
