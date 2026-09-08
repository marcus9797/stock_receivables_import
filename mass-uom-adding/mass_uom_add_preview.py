# -*- coding: utf-8 -*-
"""
mass_uom_add_preview.py
=======================
Read `mass_UOM_edit_layout.xlsx` (an instruction sheet: one row per "add this
PKT/box UOM to existing item codes that match product/country/brand/size"),
match against `imported_item_codes.xlsx`, and produce a REVIEWABLE PREVIEW of
the rows that would be added -- WITHOUT touching the master.

Intent per layout row (col A=product phrase, B=UOM to add, C=countries,
D=brand, E=sizes; "does not matter" = wildcard):
  For every existing item code that matches the row's filters and does NOT
  already carry that UOM, a new UOM row would be added (a "Pre Pack" row).

Matching: the product phrase in col A (minus country/brand/fruit words) is the
**variety**, compared against the code's ItemCategory -- exact match first,
falling back to "contains" only when the exact form finds nothing (covers
categories such as `Navelate Navel`).  Country = ItemClass, brand = ItemBrand,
sizes = UDF_UOM_Size.  A product phrase that does NOT say "Organic" will not
match an "Organic ..." category.

Rate rules (confirmed by project owner):
  * count pre-packs  PKT<n> on PCS/PKT-base codes        -> Rate = n
  * weight punnets on GM-base codes (e.g. PKT500)        -> Rate = grams
  * weight punnets on PCS-base codes (e.g. PKT800/750)   -> Rate = ceil(grams / (carton_g / pieces_per_carton))
    grams-per-piece = UDF_UOM_WeightPerCtn / UDF_UOM_SizeUOM on the code's CTN
    row; if WeightPerCtn is blank, net kg is inferred from the Description.

Ambrosia row is widened to NZ + Italy per the owner (Ambrosia also comes from
Italy although col C says NZ).  UOMDesc shown is PROVISIONAL.

Output: `mass_uom_add_preview_<MMM-dd>_<hhmm>.xlsx` next to the layout file
(Summary + Rows sheets).  Nothing is written to imported_item_codes.xlsx.
Run from project root:  python python_scripts/mass_uom_add_preview.py
"""

import math
import os
import re
import sys
import datetime
import json
from collections import defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER = os.path.join(ROOT, "imported_item_codes.xlsx")
LAYOUT = os.path.join(ROOT, "mass_UOM_edit_layout.xlsx")

# ---------------------------------------------------------------------------
# Country vocabulary
# ---------------------------------------------------------------------------
def _load_country_map():
    cmap = {}
    hp = os.path.join(ROOT, "create_new_item_code", "item_code_hashmap.json")
    try:
        with open(hp, encoding="utf-8") as f:
            data = json.load(f)
        for full, short in (data.get("country") or {}).items():
            cmap[str(full).lower()] = full
            cmap[str(short).lower()] = full
    except Exception as e:
        print("warn: could not read item_code_hashmap.json (%s)" % e)
    cmap.update({
        "sa": "South Africa", "south africa": "South Africa",
        "us": "USA", "usa": "USA",
        "nz": "New Zealand", "new zealand": "New Zealand",
        "aust": "Australia", "aus": "Australia", "australia": "Australia",
        "cn": "China", "china": "China", "ita": "Italy", "italy": "Italy",
        "egy": "Egypt", "egypt": "Egypt",
    })
    return cmap

COUNTRY = _load_country_map()
COUNTRY_SHORT_TOKENS = {"usa", "us", "sa", "nz", "aust", "aus", "cn", "ita", "egy"}

BRAND_WHITELIST = {"zespri", "rockit", "enza", "sunkist", "barnfield",
                   "soluna", "juliet", "organic juliet"}
# Our own (house) brands.  They may sit in UDF_OurBrand or (for a few items)
# in ItemBrand; either way they lead the UOMDesc phrase before the country.
HOUSE_BRANDS = {"garden basket", "tian tian", "sugarball", "my love bite", "lillie"}
FRUIT_WORDS = ["apples", "apple", "pears", "pear", "oranges", "orange",
               "kiwis", "kiwi", "cherries", "cherry"]
FRUIT_ITEMTYPE = {"apple": "Apples", "apples": "Apples", "pear": "Pears", "pears": "Pears",
                  "orange": "Oranges", "oranges": "Oranges", "kiwi": "Kiwis", "kiwis": "Kiwis",
                  "cherry": "Cherries", "cherries": "Cherries"}


def fruit_from_text(text):
    low = " " + str(text or "").lower() + " "
    for w in FRUIT_WORDS:
        if re.search(r"\b" + w + r"\b", low):
            return FRUIT_ITEMTYPE[w]
    return None


def split_tokens(text):
    return [t for t in re.split(r"[\s,;./-]+", str(text or "").strip().lower()) if t]


def _strip_tok(txt, tok):
    if tok in COUNTRY_SHORT_TOKENS or len(tok) > 3:
        txt = re.sub(r"\b" + re.escape(tok) + r"\b", " ", txt)
    return txt


def strip_country_tokens(text):
    t = " " + str(text or "").strip().lower() + " "
    for tok in sorted(COUNTRY, key=len, reverse=True):
        t = _strip_tok(t, tok)
    return re.sub(r"\s+", " ", t).strip()


def norm(text):
    return re.sub(r"\s+", " ", str(text or "").lower()).strip()


def variety_phrase(product, brand_req):
    """col A minus country words, minus the requested brand, minus fruit words."""
    t = strip_country_tokens(product)
    if brand_req:
        for w in brand_req.split():
            t = re.sub(r"\b" + re.escape(w) + r"\b", " ", t)
    for fw in FRUIT_WORDS:
        t = re.sub(r"\b" + fw + r"\b", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def requested_countries(cell):
    """col C -> list of full ItemClass country names (None = any / none found).
    Only known country names/short-forms are accepted (never arbitrary words)."""
    if not cell or cell.lower() in ("does not matter", "dm", "-"):
        return None
    out = []
    for tok in split_tokens(cell):
        full = COUNTRY.get(tok)
        if full and full not in out:
            out.append(full)
    return out or None


def uom_digits(uom):
    m = re.match(r"PKT\s*(\d+)", str(uom), re.I)
    return int(m.group(1)) if m else None


def has_box_config(desc):
    """Carton described as N x M packs/bags (e.g. 10x850G, 8X6PCS, 20X450G):
    SizeUOM is then the pack count, NOT pieces per carton."""
    return bool(re.search(r"\d+\s*[xX]\s*\d+\s*(?:g|kg|lb|pcs)\b", str(desc or ""), re.I))


# Weight-pack rates the owner confirmed by hand on 2026-09-08 (review of the
# VERIFY-RATE rows in the preview).  These take precedence over auto-derivation.
OVERRIDES = {
    # Bagged Packham cartons (base PCS but carton = 10 inner 800-850g bags):
    # one PKT800 consumer punnet ~= one bag  ->  Rate = 1.
    "PR-PCKHM-SA-0032": ("ADD", 1, "owner: 10x850G bagged carton - PKT800 ~= 1 bag -> rate 1"),
    "PR-PCKHM-SA-0042": ("ADD", 1, "owner: 10x800G bagged carton - PKT800 ~= 1 bag -> rate 1"),
    # Weight punnets on PKT-base boxed items are not convertible -> do NOT add.
    "PR-FRGNT-CN-0015": ("SKIP", None, "owner: base PKT - do not add PKT750"),
    "PR-FRGNT-CN-0016": ("SKIP", None, "owner: base PKT - do not add PKT750"),
    "CHRY-CHRY-USA-0001": ("SKIP", None, "owner: base PKT - do not add PKT500"),
}


def norm_size(s):
    return str(s or "").strip().lower()


def sizes_match(code_size, req_tokens):
    if req_tokens is None:
        return True
    cs = norm_size(code_size)
    if not cs:
        return False

    def num(x):
        try:
            return float(re.sub(r"[^0-9.]", "", x))
        except ValueError:
            return None

    for req in req_tokens:
        r = norm_size(req)
        if r == cs:
            return True
        cn, rn = num(cs), num(r)
        if cn is not None and rn is not None and cn == rn:
            return True
    return False


# ---------------------------------------------------------------------------
# Master / layout loading
# ---------------------------------------------------------------------------
def load_master():
    wb = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    groups = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        groups[r[0]].append(r)
    wb.close()
    return groups


def parse_layout():
    wb = openpyxl.load_workbook(LAYOUT, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    for rr in ws.iter_rows(min_row=2, values_only=True):
        if rr[0] is None and rr[1] is None:
            continue
        vals = [str(x or "").strip() for x in rr[:5]]
        if vals[1]:
            rows.append({"product": vals[0], "uom": vals[1], "countries": vals[2],
                         "brand": vals[3], "sizes": vals[4]})
    wb.close()
    return rows


# ---------------------------------------------------------------------------
# Row-value builders
# ---------------------------------------------------------------------------
def weight_per_ctn(row):
    v = row[12]
    if v not in (None, ""):
        try:
            return float(v)
        except (TypeError, ValueError):
            pass
    d = str(row[1] or "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*kg", d, re.I)
    if m:
        return float(m.group(1)) * 1000.0
    m = re.search(r"(\d+)\s*[xX]\s*(\d+)\s*g", d, re.I)
    if m:
        return float(m.group(1)) * float(m.group(2))
    return None


def size_uom_count(row):
    v = row[11]
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def product_prefix(group):
    """Product phrase (country [brand] variety fruit) from the code's own base
    (PCS/GM) row description, with its P/PCS / P/KGS tail removed."""
    for r in group:
        if r[18] in ("PCS", "GM") and r[16] in (None, "", 1, "1"):
            ud = str(r[2] or "")
            ud = re.sub(r"\s+P/(PCS|KGS|PKG|PKT|GM|KG)\s*$", "", ud, flags=re.I)
            return ud.strip()
    for r in group:
        if r[18] == "CTN":
            d = re.sub(r"\([^)]*\)", " ", str(r[1] or ""))
            d = re.sub(r"\s*\d+(?:\.\d+)?\s*kg\s*$", "", d, flags=re.I)
            d = re.sub(r"\s*\d+\s*lb\s*$", "", d, flags=re.I)
            d = re.sub(r"\s+\d+([./]\d+)?\s*s\s*$", "", d)
            return d.strip()
    return (group[0][1] or "").strip()


def strip_nonwhitelist_brand(prefix, r0):
    """Drop a non-whitelist grower/brand word from the product phrase (unless
    the brand IS the variety).  Whitelist brands are handled in curated_."""
    brand = str(r0[7] or "").strip()
    if not brand or brand.lower() in BRAND_WHITELIST or brand.lower() in HOUSE_BRANDS:
        return prefix
    if norm(brand).lower() and norm(brand) == norm(r0[4]):  # brand == variety
        return prefix
    out = " " + prefix + " "
    for w in split_tokens(brand):
        out = re.sub(r"\b" + re.escape(w) + r"\b", " ", out, flags=re.I)
    return re.sub(r"\s+", " ", out).strip()


def curated_uomdesc(group, prefix, uom_d, kind):
    """Provisional 'Pre Pack ...' UOMDesc (whitelist brand shown, Envy keeps
    its grade)."""
    r0 = group[0]
    desc = (prefix or "").strip()
    desc = re.sub(r"\s+[-]\s+", " ", desc)  # drop stray dashes left after brand strip
    # Our own (house) brand leads the phrase, before the country:
    #   "Pre Pack Garden Basket China Fragrant Pears 750G"
    our = str(r0[8] or "").strip() or (
        str(r0[7] or "").strip() if str(r0[7] or "").strip().lower() in HOUSE_BRANDS else "")
    if our:
        desc = re.sub(r"(?i)\s*\b" + re.escape(our) + r"\b\s*", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()
        desc = (our + " " + desc).strip()
    lower = desc.lower()
    brand = str(r0[7] or "").strip()
    if brand.lower() in BRAND_WHITELIST and brand.lower() not in lower:
        bits = desc.split(" ", 1)
        if len(bits) == 2 and bits[0]:
            desc = bits[0] + " " + brand + " " + bits[1]
        else:
            desc = (brand + " " + bits[0]).strip()
    var = (r0[4] or "").strip().lower()
    grade = str(r0[6] or "").strip()
    if var in ("envy", "organic envy") and grade:
        fruit = r0[3] or "Apples"
        if grade.lower() not in lower:
            desc = desc.replace(fruit, "(%s) %s" % (grade, fruit), 1)
    label = ("1KG" if (kind == "weight" and uom_d == 1000)
             else str(uom_d) + ("G" if kind == "weight" else "s"))
    return ("Pre Pack " + desc + " " + label).strip()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    groups = load_master()
    by_code = list(groups.items())
    layouts = parse_layout()

    summary_rows, detail_rows, notes = [], [], []

    for li, row in enumerate(layouts, start=2):
        uom = row["uom"].strip().upper()
        ud = uom_digits(uom)
        product = row["product"]
        brand = row["brand"]
        brand_req = None if (not brand or brand.lower() in ("does not matter", "dm", "-")) else brand.lower()
        sizes_req = None if (row["sizes"].lower() in ("does not matter", "dm", "-")) else split_tokens(row["sizes"])

        # country set (col C; widened for Ambrosia per owner)
        countries = requested_countries(row["countries"])
        if countries is None:
            countries = requested_countries(" , ".join(split_tokens(product)))
        phrase = variety_phrase(product, brand_req)
        if "ambrosia" in phrase:
            countries = ["New Zealand", "Italy"]

        want_organic = "organic" in phrase.split()
        fruit_from_a = fruit_from_text(product)
        cat = norm(phrase)
        if not cat and not fruit_from_a:
            notes.append("row %d %r: product phrase has no variety and no fruit word; skipped" % (li, product))
            summary_rows.append({"LayoutRow": li, "Product": product, "UOM": uom, "MatchMode": "skipped",
                                 "Countries": row["countries"], "CountriesMatched": "", "Brand": row["brand"],
                                 "Sizes": row["sizes"], "MatchedCodes": 0, "ToAdd": 0, "Verify": 0,
                                 "Skipped": 0, "AlreadyHas": 0})
            continue

        def matches(item, fallback_contains):
            _code, grp = item
            r0 = grp[0]
            if fruit_from_a and str(r0[3]).strip().lower() != fruit_from_a.lower():
                return False
            if countries and str(r0[5]).strip().lower() not in [x.lower() for x in countries]:
                return False
            if brand_req and str(r0[7] or "").strip().lower() != brand_req:
                return False
            if not sizes_match(r0[10], sizes_req):
                return False
            if not cat:
                # product phrase is fruit-only (e.g. "USA Cherries"): variety is
                # open; the fruit gate above already restricts the item type.
                return True
            varcat = norm(r0[4])
            if fallback_contains:
                return cat in varcat
            return varcat == cat

        # decide mode: prefer exact-category; fall back to contains if exact finds none
        exact = [g for g in by_code if matches(g, False)]
        mode = "category-exact"
        pool = exact
        if not exact:
            pool = [g for g in by_code if matches(g, True)]
            if pool:
                mode = "category-contains-fallback"
                notes.append("row %d %r: no exact ItemCategory '%s'; used 'contains' -> %d codes"
                             % (li, product, cat, len(pool)))

        pool.sort(key=lambda g: g[0])
        if not pool:
            notes.append("row %d %r: NO codes matched (variety '%s', countries=%s, brand=%s, sizes=%s)"
                         % (li, product, cat, countries, brand_req, sizes_req))

        n_add = n_has = n_verify = n_skip = 0
        for code, g in pool:
            r0 = g[0]
            base = str(r0[20] or "")
            prefix = strip_nonwhitelist_brand(product_prefix(g), r0)
            rate = size_rate = gpp = None
            rate_src = reason = ""
            kind = "weight" if (ud and ud >= 500) or base == "GM" else "count"
            provisional = ""

            if any((r[18] or "").upper() == uom for r in g):
                action = "ALREADY-HAS"
                reason = ""
            elif (ov := OVERRIDES.get(code)) is not None:
                o_act, o_rate, o_reason = ov
                reason = o_reason
                if o_act == "SKIP":
                    action = "SKIP"
                else:                       # owner-set rate (e.g. rate = 1)
                    action = "ADD"
                    rate = size_rate = o_rate
                    rate_src = "owner override"
                    kind = "weight"
                    provisional = curated_uomdesc(g, prefix, ud, kind)
            elif ud is None:
                action = "NO-RATE"; reason = "cannot parse UOM " + uom
            elif base == "GM":
                action = "ADD"; kind = "weight"; rate = size_rate = ud
                rate_src = "GM-grams"
                provisional = curated_uomdesc(g, prefix, ud, kind)
            elif ud >= 500:                # weight punnets on PCS/PKT-base codes
                kind = "weight"
                w = weight_per_ctn(r0); szu = size_uom_count(r0)
                if has_box_config(r0[1]):
                    action = "VERIFY-RATE"
                    reason = ("verify rate: carton is boxed/multi-pack (%s); "
                              "need real pieces-per-carton to convert %s" % (r0[1], uom))
                elif w and szu:
                    gpp = w / szu
                    if 40 <= gpp <= 500:
                        action = "ADD"
                        rate = size_rate = int(math.ceil(ud / gpp))
                        rate_src = "weight: ceil(%d / %.2fg-per-pc)" % (ud, gpp)
                        provisional = curated_uomdesc(g, prefix, ud, kind)
                    else:
                        action = "VERIFY-RATE"
                        reason = ("verify rate: carton g/pc = %.2f (WeightPerCtn=%s / "
                                  "SizeUOM=%s) implausible - is SizeUOM pieces-per-carton?"
                                  % (gpp, w, szu))
                else:
                    action = "VERIFY-RATE"
                    reason = ("verify rate: no WeightPerCtn/SizeUOM to compute %s on %s base"
                              % (uom, base))
            else:
                action = "ADD"; kind = "count"; rate = size_rate = ud; rate_src = "count"
                provisional = curated_uomdesc(g, prefix, ud, kind)

            if action == "ADD":
                n_add += 1
            elif action == "VERIFY-RATE":
                n_verify += 1
            elif action == "ALREADY-HAS":
                n_has += 1
            elif action == "SKIP":
                n_skip += 1
            detail_rows.append({
                "LayoutRow": li, "Product": product, "UOM": uom,
                "MatchMode": mode, "VarietyPhrase": cat,
                "CountryFilter": row["countries"], "BrandFilter": row["brand"],
                "SizeFilter": row["sizes"],
                "MatchedCountries": ", ".join(countries) if countries else "any",
                "ItemCode": code, "Description": r0[1],
                "BaseUOM": base, "Fruit": r0[3], "Variety": r0[4],
                "Country": r0[5], "Brand": r0[7], "Size": r0[10],
                "Action": action, "Rate": rate, "SizeRate": size_rate,
                "RateSource": rate_src, "g-per-piece": round(gpp, 2) if gpp else None,
                "Provisional_UOMDesc": provisional, "Reason": reason,
            })

        summary_rows.append({
            "LayoutRow": li, "Product": product, "UOM": uom, "MatchMode": mode,
            "Countries": row["countries"],
            "CountriesMatched": ", ".join(countries) if countries else "any",
            "Brand": row["brand"], "Sizes": row["sizes"],
            "MatchedCodes": len(pool), "ToAdd": n_add, "Verify": n_verify,
            "Skipped": n_skip, "AlreadyHas": n_has,
        })

    # ---------------- write preview workbook ----------------
    stamp = datetime.datetime.now().strftime("%b-%d_%I%M%p").lower()
    out = os.path.join(ROOT, "mass_uom_add_preview_%s.xlsx" % stamp)
    wb = openpyxl.Workbook()
    ws_s = wb.active
    ws_s.title = "Summary"
    s_headers = ["LayoutRow", "Product", "UOM", "MatchMode", "Countries",
                 "CountriesMatched", "Brand", "Sizes", "MatchedCodes", "ToAdd", "Verify",
                 "Skipped", "AlreadyHas"]
    ws_s.append(s_headers)
    for r in summary_rows:
        ws_s.append([r[h] for h in s_headers])

    ws_r = wb.create_sheet("Rows")
    r_headers = ["LayoutRow", "Product", "UOM", "MatchMode", "VarietyPhrase",
                 "CountryFilter", "BrandFilter", "SizeFilter", "MatchedCountries",
                 "ItemCode", "Description", "BaseUOM", "Fruit", "Variety", "Country",
                 "Brand", "Size", "Action", "Rate", "SizeRate", "RateSource",
                 "g-per-piece", "Provisional_UOMDesc", "Reason"]
    ws_r.append(r_headers)
    for r in detail_rows:
        ws_r.append([r[h] for h in r_headers])

    for sn, wl in {
        "Summary": [9, 24, 8, 24, 12, 20, 10, 20, 13, 8, 11],
        "Rows": [9, 24, 8, 24, 18, 12, 12, 18, 18, 20, 55, 8, 12, 18, 14, 12, 10,
                 12, 8, 9, 26, 12, 55, 36],
    }.items():
        w = wb[sn]
        for i, cw in enumerate(wl, start=1):
            w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = cw
    wb.save(out)

    print("Preview written to: %s\n" % out)
    total_add = sum(r["ToAdd"] for r in summary_rows)
    total_verify = sum(r.get("Verify", 0) for r in summary_rows)
    total_skip = sum(r.get("Skipped", 0) for r in summary_rows)
    print("layout rows: %d  | rows to add: %d  | verify-rate: %d  | owner-skipped: %d\n"
          % (len(layouts), total_add, total_verify, total_skip))
    print("%-3s %-24s %-7s %5s %5s %5s %5s %5s  %s" % ("#", "product", "uom", "match", "add", "chk", "skp", "have", "mode"))
    for r in summary_rows:
        flag = ""
        if r["MatchedCodes"] == 0:
            flag = "  NO MATCH"
        elif r["ToAdd"] == 0 and r.get("Verify", 0) == 0 and r.get("Skipped", 0) == 0:
            flag = "  all have it"
        print("%-3d %-24s %-7s %5d %5d %5d %5d %5d  %s%s" % (r["LayoutRow"], r["Product"][:24],
              r["UOM"], r["MatchedCodes"], r["ToAdd"], r.get("Verify", 0),
              r.get("Skipped", 0), r["AlreadyHas"], r["MatchMode"], flag))
    if notes:
        print("\nNOTES / things to review:")
        for n in notes:
            print("  - " + n)


if __name__ == "__main__":
    sys.exit(main())
