# -*- coding: utf-8 -*-
"""
build_open_items_sheet.py
=========================
Add an `Open_Items` sheet to `purchase_orders_b2b/giant_layout.xlsx` listing
everything from the Giant work that is still unresolved, so it travels with the
layout instead of living in chat scrollback.

`mass_uom_add_preview.py` reads only `Sheet1`, so extra sheets are safe.

Rows are one flat table with a Category, so the sheet can be filtered:

  No match        the instruction matches no master code (the preview's verdict)
  Verify rate     the preview computed no rate and needs the owner's number
  Opportunity     a one-line repair would turn a no-match into a match
  Master cell     a cell in imported_item_codes.xlsx that disagrees with itself
  Rule bucket     a pre-existing UDF_UOM_UOMDesc rule violation, by rule

The preview is regenerated and consumed inside this script: the xlsx in this
folder is synced, so a freshly written file is briefly unreadable by a second
process.
"""

import os
import re
import sys
import time
import shutil
import subprocess
import collections

import openpyxl
from openpyxl.styles import Font

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LAYOUT = os.path.join(ROOT, "purchase_orders_b2b", "giant_layout.xlsx")
PREVIEW_SCRIPT = os.path.join(ROOT, "mass-uom-adding", "mass_uom_add_preview.py")
PREVIEW_GLOB = "mass_uom_add_preview_"
WORK = os.path.join(ROOT, "giant_preview_working.xlsx")
MASTER = os.path.join(ROOT, "imported_item_codes.xlsx")

# why a given Giant phrase has no master counterpart, and what to do about it
CAUSE = [
    (r"JUMBO", "GN", "JUMBO reads as a size, not a variety, and the master's only "
                    "Jumbo blueberry is Australian",
               "Get the variety from Giant, or treat JUMBO as a droppable size word"),
    (r"MYOCA BLACK", "GN", "no Moyca code of that variety exists",
               "Confirm with Giant which Moyca variety this is"),
    (r"PIQA BOO", "OP", "master spells the variety Piqa, not Piqa Boo",
               "add a PIQA BOO -> Piqa synonym to build_giant_layout.SYNONYM"),
    (r"^ORANGE MANDARIN", "OP", "the leading word ORANGE blocks the fruit word the "
                                "matcher uses to strip the variety",
               "drop a leading fruit word when a second fruit word remains"),
    (r"HONEY GOLD|RED BEAUTY|DONUT|FLAT YELLOW|^NECTARINE WHITE|MANDARIN WOGAN",
     "GN", "Giant's name is not a master ItemCategory",
               "Confirm the master category with the owner, or add a synonym"),
    (r"BLACKBERRY|KIWIBERRY|PLUM RED|RED PLUM CHERRY|CHERRY PLUM|SUNGOLD KIWI ORGANIC",
     "VA", "no such variety (or organic variant) in the master",
               "Needs a new item code via create_new_item_code/"),
    (r"APPLE-CHINA FUJI|^PEAR BLUSH", "GB", "description is garbled or the word is "
                                            "not a variety at all",
               "Ask Giant for a clean description"),
]
DEFAULT_CAUSE = ("GN", "no master ItemCategory matches the derived variety",
                 "Confirm with the owner")

CATEGORY_LABEL = {"GN": "Confirm variety", "OP": "Opportunity", "VA": "Needs new code",
                  "GB": "Bad input"}


def newest_preview():
    """Run the preview and read the file it just wrote, before the sync agent
    gets to it."""
    subprocess.run([sys.executable, PREVIEW_SCRIPT,
                    os.path.join(ROOT, "purchase_orders_b2b", "giant_layout.xlsx")],
                   cwd=ROOT, check=True, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL)
    import glob
    cands = sorted(glob.glob(os.path.join(ROOT, PREVIEW_GLOB + "*.xlsx")),
                   key=os.path.getmtime, reverse=True)
    for _ in range(60):
        for c in cands[:4]:
            try:
                shutil.copyfile(c, WORK)
                wb = openpyxl.load_workbook(WORK)
                if wb["Rows"].max_row > 100:
                    return wb
            except Exception:
                continue
        time.sleep(0.5)
    raise SystemExit("ABORT: could not obtain a readable preview")


def preview_rows(wb, sheet):
    ws = wb[sheet]
    h = [c.value for c in ws[1]]
    return [{h[i]: ws.cell(r, i + 1).value for i in range(len(h))}
            for r in range(2, ws.max_row + 1)]


def main():
    pv = newest_preview()
    summary = preview_rows(pv, "Summary")
    rows = preview_rows(pv, "Rows")

    items = []
    # ---- no-match instructions ------------------------------------------
    for d in sorted(summary, key=lambda x: str(x["Product"])):
        if d["MatchedCodes"]:
            continue
        prod = str(d["Product"])
        code, why, action = DEFAULT_CAUSE
        for pat, c, w, a in CAUSE:
            if re.search(pat, prod, re.I):
                code, why, action = c, w, a
                break
        items.append(["No match", "%s  ->  %s" % (prod, d["UOM"]),
                      "country=%s  brand=%s" % (d["Countries"], d["Brand"]),
                      why, CATEGORY_LABEL.get(code, code) + " - " + action])

    # ---- VERIFY-RATE rows ------------------------------------------------
    for x in rows:
        if x.get("Action") != "VERIFY-RATE":
            continue
        items.append(["Verify rate", x["ItemCode"],
                      "%s on a %s base" % (x["UOM"], x["BaseUOM"]),
                      str(x.get("Reason") or "")[:180],
                      "owner supplies the Rate by hand"])

    # ---- opportunities / master cells / rule buckets ---------------------
    items += [
        ["Opportunity", "PEAR PIQA BOO NZ",
         "PEAR PIQA BOO NZ -> PKT3",
         "master spells the variety Piqa; only the BOO differs",
         "add PIQA BOO -> Piqa to build_giant_layout.SYNONYM (one line)"],
        ["Opportunity", "ORANGE MANDARIN WOGAN",
         "ORANGE MANDARIN WOGAN -> PKT800",
         "the master has Wogan; the leading word ORANGE blocks the fruit strip",
         "drop a leading fruit word when a second fruit word remains (one line)"],
        ["Master cell", "PR-PCKHM-SA-0033",
         "PCS base row UDF_UOM_SizeUOM = 10",
         "owner confirmed SizeUOM is meant to be 112; 10 is the punnet count",
         "correct the cell to 112 (previewed separately)"],
        ["Master cell", "PR-PCKHM-SA-0033",
         "CTN row Rate / SizeRate = 10",
         "on a PCS base the CTN rate should be the piece count, 112, and it "
         "disagrees with that row's own SizeUOM = 112",
         "correct to 112 - feeds stock conversion, so preview on its own"],
    ]

    # run the rule checker fresh rather than trusting an earlier report
    fresh = os.path.join(ROOT, "uomdesc_check_current.xlsx")
    subprocess.run([sys.executable,
                    os.path.join(ROOT, "customer-desc-column-curation",
                                 "check_uomdesc_rules.py"),
                    "--out", fresh],
                   cwd=ROOT, check=True, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL)
    wbv = openpyxl.load_workbook(fresh)
    ws = wbv["Summary"]
    counts = {}
    for r in range(2, ws.max_row + 1):
        k, v = ws.cell(r, 1).value, ws.cell(r, 2).value
        if isinstance(v, int):
            counts[k] = v
    os.remove(fresh)
    BUCKET = {"PKTa": "count packs (PKT<n<100>) not ending '<n>s'",
              "PKTb": "gram punnets (PKT>=100) not ending '<n>G'",
              "PKTpre": "PKT row missing the 'Pre Pack' prefix (audit script rule)",
              "R11": "fig variety not dropped",
              "R12": "house brand not leading the phrase (generation-time rule)",
              "R3ba": "KG row ending 'P/KGS' instead of 'P/KG'",
              "R1": "grade shown on a PKT row",
              "R5c": "packing form not in the compact config spelling"}
    for rule, n in sorted(counts.items()):
        if rule in ("Rows checked", "Rows with violations"):
            continue
        items.append(["Rule bucket", rule, "%d row(s)" % n,
                      BUCKET.get(rule, ""),
                      "not applied - sweep pending owner decision"])

    # ---- write the sheet -------------------------------------------------
    wb = openpyxl.load_workbook(LAYOUT)
    if "Open_Items" in wb.sheetnames:
        del wb["Open_Items"]
    out = wb.create_sheet("Open_Items")
    hdr = ["Category", "Item", "Detail", "Why it is open", "Suggested action"]
    out.append(hdr)
    for c in range(1, 6):
        out.cell(1, c).font = Font(bold=True)
    for it in items:
        out.append(it)
    out.freeze_panes = "A2"
    out.auto_filter.ref = "A1:E%d" % (len(items) + 1)
    for col, w in zip("ABCDE", (14, 46, 46, 62, 62)):
        out.column_dimensions[col].width = w
    wb.save(LAYOUT)

    print("Open_Items written to %s" % os.path.basename(LAYOUT))
    print("  rows: %d" % len(items))
    print("  by category: %s"
          % dict(collections.Counter(i[0] for i in items)))
    try:
        os.remove(WORK)
    except Exception:
        pass


if __name__ == "__main__":
    main()
