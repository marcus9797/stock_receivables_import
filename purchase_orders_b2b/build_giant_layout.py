# -*- coding: utf-8 -*-
"""
build_giant_layout.py
=====================
Turn Giant's cleaned product list into a `mass-uom-adding` instruction sheet.

Input : purchase_orders_b2b/sept_2025_to_2026_concatenated.xlsx, sheet
        `Unique_Products` (PO_NUMBER, PRD_NUMBER, PRD_DESC, UOM, ORDER_QTY,
        PACK_QTY, BARCODE).
Output: purchase_orders_b2b/giant_layout.xlsx, sheet `Sheet1` in the 5-column
        layout `mass_uom_add_preview.py` expects
        (Product description | UOM | Countries | Brand | Sizes), plus two
        reference sheets the preview ignores:
          `Source`    -- layout row -> PRD_NUMBER + original PRD_DESC + barcode
          `Report`    -- every row held back, with the reason

Owner rules applied (2026-09-26):

  PACK_QTY > 20   -> pieces per carton: NOT a pre-pack, so no instruction in
                     this batch. Held back, reported.
  PACK_QTY <= 20  -> punnets per carton. The per-punnet quantity comes from
                     PRD_DESC: a weight token (500GM, 1KG, 800G) -> PKT<grams>,
                     else a count token (6S, 8'S, 4PC/PKT) -> PKT<count>.
  BARCODE         -> a value of <= 10 characters is invalid and is blanked in
                     the sheet (owner: apply the character test literally, so
                     no scientific-notation value is singled out).

Multi-country descriptions produce ONE instruction per country (owner
2026-09-26). Descriptions naming no country at all get the literal wildcard
`does not matter` (owner chose "match every country"); `mass-uom-adding/CLAUDE.md`
records that this silently means *any* country, so those rows are pinned in the
Report and the preview's `Rows` sheet must be read per country.

Usage:
    python purchase_orders_b2b/build_giant_layout.py [--no-blank-barcodes]
"""

import os
import re
import sys
import collections

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOK = os.path.join(ROOT, "purchase_orders_b2b", "sept_2025_to_2026_concatenated.xlsx")
OUT = os.path.join(ROOT, "purchase_orders_b2b", "giant_layout.xlsx")
MASTER = os.path.join(ROOT, "imported_item_codes.xlsx")

# Reuse the preview's own country/variety/fruit parsing so a "repair" is judged
# by exactly the logic that will later do the matching.
sys.path.insert(0, os.path.join(ROOT, "mass-uom-adding"))
sys.path.insert(0, os.path.join(ROOT, "prepack_items_list"))
import mass_uom_add_preview as P
import build_zhengheng_layout as Z          # the owner-blessed grape mapping

# colour+seedless -> concrete master varieties, from
# prepack_items_list/grapes_color_seed_mapping.xls (owner rulings included).
GRAPE_BUCKETS = Z.grape_buckets()
COLOUR_BUCKETS = Z.COLOUR_BUCKETS

SRC_SHEET = "Unique_Products"
BARCODE_COL = 7                                   # G
PACK_COL = 6                                      # F
DESC_COL = 3                                      # C
PRD_COL = 2                                       # B
WILDCARD = "does not matter"
BARCODE_MAX = 10                                  # <= this many chars = invalid

# owner's country vocabulary (countries.yaml + the master's ItemClass list)
COUNTRIES = {
    "Australia": r"AUST?\b|AUSTRALIA", "China": r"CN\b|CHINA",
    "Egypt": r"EGY?\b|EGYPT", "France": r"FRANCE\b", "Italy": r"ITA\b|ITALY",
    "New Zealand": r"NZ\b", "South Africa": r"SA\b|ZA\b", "USA": r"USA?\b",
    "Zimbabwe": r"ZIMB\b|ZIMBABWE", "Korea": r"KOR\b|KOREA",
    "Turkey": r"TURKEY", "Spain": r"SPAIN", "India": r"INDIA", "Peru": r"PERU",
    "Chile": r"CHILE", "Ecuador": r"ECUADOR", "Pakistan": r"PAKISTAN",
    "Vietnam": r"VIETNAM", "Mexico": r"MEXICO", "Morocco": r"MOROCCO",
    "Greece": r"GREECE", "Japan": r"JAPAN", "Kenya": r"KENYA",
    "Lebanon": r"LEBANON", "Namibia": r"NAMIBIA", "Poland": r"POLAND",
    "Portugal": r"PORTUGAL", "Serbia": r"SERBIA", "Taiwan": r"TAIWAN",
    "Ukraine": r"UKRAINE", "Argentina": r"ARGENTINA", "Thailand": r"THAILAND",
}
WHITELIST_BRANDS = ["ZESPRI", "ROCKIT", "ENZA", "SUNKIST", "BARNFIELD",
                    "SOLUNA", "JULIET"]

# The per-punnet quantity, in the two forms the descriptions use.
WEIGHT = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|GM|G)\b", re.I)
COUNT = re.compile(r"(\d+)\s*'?\s*(S|PCS|PC)\b(?!\w)", re.I)
# carton-count notation Giant uses for piece-count items (C88, C100/113, C30)
CARTON_COUNT = re.compile(r"\bC\s*\d+(?:\s*/\s*\d+)?\b", re.I)


def clean(s):
    """Repair the export's encoding damage and collapse whitespace."""
    s = str(s or "")
    s = s.replace("u00a0", " ").replace(" ", " ")
    s = s.replace("’", "'").replace("‘", "'")
    s = re.sub(r"[�]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def split_x_runs(s):
    """Space out the X that Giant uses as pack notation ('800GX10PKT',
    '6'SX 12PKT', '10X500G') so the pack token is visible.

    ONLY pack notation -- an X that is a letter inside a word must survive
    ('PIXIE', 'MEXICO').  Every pattern below therefore requires a digit on at
    least one side, which is what distinguishes the two.
    """
    s = re.sub(r"([0-9])\s*'?\s*(KG|GM|G|S|PCS|PC)\s*[xX]\s*(?=[0-9])",
               r"\1\2 X ", s, flags=re.I)       # 800GX10PKT, 6'SX 12PKT
    s = re.sub(r"([0-9])\s*'?\s*(KG|GM|G|S|PCS|PC)\s*[xX]\s*(?=PKT)",
               r"\1\2 ", s, flags=re.I)         # 250GMXPKT
    s = re.sub(r"(?<=[0-9])\s*[xX]\s*(?=[0-9])", " X ", s)   # 10X500G
    return s


def countries_of(desc):
    up = desc.upper()
    return sorted({c for c, pat in COUNTRIES.items()
                   if re.search(r"\b(?:" + pat + ")", up)})


def brands_of(desc):
    up = desc.upper()
    return [b for b in WHITELIST_BRANDS if re.search(r"\b" + b + r"\b", up)]


def per_pack(desc):
    """-> (uom, token) e.g. ('PKT800','800G'); (None, None) when underivable."""
    d = split_x_runs(desc)
    m = WEIGHT.search(d)
    if m:
        grams = float(m.group(1))
        if m.group(2).upper() == "KG":
            grams *= 1000
        return "PKT%d" % round(grams), m.group(0).strip()
    m = COUNT.search(d)
    if m:
        return "PKT%d" % int(m.group(1)), m.group(0).strip()
    return None, None


BOX_TOKEN = re.compile(r"(\d+(?:\.\d+)?)\s*'?\s*(KG|GM|G|S|PCS|PC)\b", re.I)


def box_uom(pack, token):
    """The BOX config a row implies: PACK_QTY punnets x the per-punnet quantity.

    Owner rule (2026-09-26): `PACK_QTY` is the number of pkts in a carton, so
    `PACK_QTY=14` with a per-pkt `5s` is the UOM `14x5s` -- a box form, not a
    `PKT5` punnet.  Spelling is the compact lower-case form the curation doc
    prescribes for new packing-form rows (`14x6s`, not `14X6PCS`).
    """
    if not str(pack).isdigit() or not token:
        return None
    m = BOX_TOKEN.search(token)
    if not m:
        return None
    n = re.sub(r"\.0+$", "", m.group(1))
    u = m.group(2).upper()
    unit = "kg" if u == "KG" else ("g" if u in ("G", "GM") else "s")
    return "%sx%s%s" % (pack, n, unit)


def product_phrase(desc, uom, token, pack=""):
    """Strip the pack/count tokens and Giant's noise words, leaving a phrase the
    preview can match on (it does its own country/brand/fruit-word stripping).

    '/' and '-' are deliberately LEFT in place here: a slash can be a country
    list OR a variety list, and a hyphen can be a separator, so both are
    resolved by `phrase_candidates` rather than blindly collapsed."""
    d = split_x_runs(desc)
    if token:
        d = d.replace(token, " ")
    d = CARTON_COUNT.sub(" ", d)
    d = re.sub(r"\b\d+\s*PKTS?\b", " ", d, flags=re.I)   # '10PKT' left by the split
    d = re.sub(r"\bPKTS?\b", " ", d, flags=re.I)
    d = re.sub(r"\bEA\b|\bEACH\b", " ", d, flags=re.I)
    d = re.sub(r"\bX\b", " ", d)
    d = re.sub(r"\s*\+\s*", " ", d)                      # '4+1' pack
    d = re.sub(r"\bKG\b", " ", d, flags=re.I)
    d = re.sub(r"\(\s*\)|\[\s*\]", " ", d)      # empty brackets left by a token
    # A bare integer equal to PACK_QTY is the punnets-per-carton count, not part
    # of the product ('ZYPHYR 10X500G' -> 'ZYPHYR 10').  Only strip that exact
    # value, so variety numbers ('86 Wang Melon', 'M7') survive.
    if pack.isdigit():
        d = re.sub(r"(?<![\w])" + pack + r"(?![\w])", " ", d)
    return re.sub(r"\s+", " ", d).strip(" ,")


# Descriptor words Giant writes that are never an ItemCategory in the master.
# DELIBERATELY EXCLUDES 'SEEDLESS', 'JUMBO', 'GIANT', 'BLACK', 'RED', 'GREEN':
# those ARE real ItemCategory values ('Black Seedless', 'Jumbo'), so stripping
# them would turn an honest no-match into a WRONG match.  Rows blocked by one of
# those are reported for a human ruling instead of being repaired blindly.
NOISE = ["PREMIUM", "LARGE", "XXL", "AIR FLOWN", "FAMILY", "SMALL", "FRESH"]

# Class / size words that sit in FRONT of a variety.  Dropped only by the
# "specificity" repair, which accepts the result ONLY when it is a non-empty
# EXACT master ItemCategory that is not itself one of these words -- so
# 'BLUEBERRY JUMBO' can never collapse to a fruit-only phrase.
CLASS_WORDS = {"SEEDLESS", "SEEDED", "RED", "GREEN", "BLACK", "WHITE", "JUMBO",
               "GIANT", "PREMIUM", "LARGE", "SMALL", "XL", "XXL", "AIR",
               "FLOWN", "FAMILY", "FRESH"}

# Grower words that are neither a master ItemBrand nor part of the variety.
# (Empty at present: `MYOCA` turned out to be a misspelling of the master's
# `Moyca` brand and is handled by SYNONYM instead, so it lands in the Brand
# column rather than being thrown away -- see the brand-move step.)
GROWER_WORDS = set()

# Spelling variants Giant uses against the master's own spelling.  Owner
# confirmed 2026-09-26 (FORELLA->Forelle, SHINGO->Singo); the rest mirror
# prepack_items_list/build_zhengheng_layout.py's MAPPING_RENAME.
SYNONYM = {
    "forella": "Forelle", "shingo": "Singo", "sonaka": "Sonaka",
    "sweanta blanc": "Sweeta Blanc", "thompson jumbo": "Thompson",
    "red prime seedless": "Prime",
    # Owner 2026-09-26: MYOCA is Giant's misspelling of the master's grower
    # brand `Moyca` (6 Spain grape codes).  Correcting it lets the brand-move
    # step put it in the Brand column, which also narrows the match to Moyca's
    # own codes instead of every Spain code of that variety.
    "myoca": "Moyca",
}

# Country vocabulary, unioned from three sources so a slash group like 'NZ/US'
# or 'ZA/ZIMB/CHINA' is recognised as a COUNTRY list and not split into
# varieties: our own full names, the short forms inside our patterns (ZIMB, AUS,
# AUST...), and the preview's own COUNTRY_TOKENS (which supplies US, EG).
_COUNTRY_WORDS = {c.upper() for c in COUNTRIES}
for _pat in COUNTRIES.values():
    _COUNTRY_WORDS |= {t.upper() for t in re.findall(r"[A-Za-z]{2,}", _pat)}
_COUNTRY_WORDS |= {str(t).strip().upper() for t, _full in P.COUNTRY_TOKENS}
_COUNTRY_WORDS |= {"EG", "ZA", "US"}

SLASH_GROUP = re.compile(r"\b([A-Za-z][A-Za-z.]*(?:/[A-Za-z][A-Za-z.]*)+)")


def slash_variants(phrase):
    """A slash list of varieties becomes one phrase per variety:
    'PEAR WILLIAM/PACKHAM SA' -> ['PEAR WILLIAM SA', 'PEAR PACKHAM SA'].

    A group whose every part is a country ('SA/EGYPT/CHINA') is a COUNTRY list --
    already expanded into one instruction per country -- so it is left alone.
    Groups containing digits are pack/count notation ('C100/113', '500G/10PKT')
    and never match the SLASH_GROUP pattern anyway."""
    m = SLASH_GROUP.search(phrase)
    if not m:
        return [phrase]
    parts = m.group(1).split("/")
    if all(p.upper().strip(".") in _COUNTRY_WORDS for p in parts):
        return [phrase]
    return [(phrase[:m.start(1)] + p + phrase[m.end(1):]).strip() for p in parts]


def phrase_candidates(desc, uom, token, pack):
    """Every phrase spelling worth trying, most faithful first."""
    base = product_phrase(desc, uom, token, pack)
    cand = []

    def add(p):
        p = re.sub(r"\s+", " ", str(p)).strip(" -,")
        if p and p not in cand:
            cand.append(p)

    for ph in (base, base.replace("-", " ")):     # hyphen as a separator
        add(ph)
        for v in slash_variants(ph):              # variety lists
            add(v)
    for ph in list(cand):                         # safe descriptor words
        for n in NOISE:
            add(re.sub(r"\b" + n + r"\b", " ", ph, flags=re.I))
    return cand


def load_master_vocab():
    """ItemCategory / ItemBrand / every (category, country, brand, itemtype)
    tuple from the master, so a repair is judged by exactly the gates the
    preview applies -- variety AND country AND brand AND fruit."""
    wb = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    cats, brands, entries = set(), set(), set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        cat = P.norm(r[4])
        brand = str(r[7] or "").strip().lower()
        cats.add(cat)
        brands.add(brand)
        entries.add((cat, str(r[5] or "").strip().lower(), brand, P.norm(r[3])))
    wb.close()
    return cats, brands, entries


def phrase_matches(phrase, brand_req, countries, cats, entries):
    """Would the preview find a code for this phrase?  Mirrors its gates: the
    variety phrase against ItemCategory (exact first, then 'phrase in category'),
    plus country, brand and the fruit of the product phrase."""
    cat = P.norm(P.variety_phrase(phrase, brand_req))
    if not cat:
        return True                     # fruit-only phrase: gated on fruit alone
    want = {c.lower() for c in countries} if countries else None
    fruit = P.fruit_from_text(phrase)
    fruit = P.norm(fruit) if fruit else None
    for e_cat, e_ctry, e_brand, e_type in entries:
        if want is not None and e_ctry not in want:
            continue
        if brand_req is not None and e_brand != brand_req:
            continue
        if fruit is not None and e_type != fruit:
            continue
        if e_cat == cat or cat in e_cat:
            return True
    return False


def choose_phrases(desc, uom, token, pack, brand, countries, vocab):
    """Every instruction this description should produce, most faithful first.

    A slash list of VARIETIES expands into one instruction per variety -- it is
    a list of products, so dropping any of them would silently lose the adds
    ('PEAR WILLIAM/PACKHAM' must reach both Williams and Packham).  Everything
    else is an ALTERNATIVE: one spelling is chosen, the first that lands."""
    cats, brands, entries = vocab
    brand_req = None if brand == WILDCARD else str(brand).lower()

    def ok(ph, br=None):
        return phrase_matches(ph, brand_req if br is None else br,
                              countries, cats, entries)

    def tidy(x):
        """Collapse the separators the preview cannot use.  '/' has already
        done its job by the time this runs, so a leftover one is noise."""
        x = re.sub(r"[/+]", " ", str(x))
        return re.sub(r"\s+", " ", x).strip(" -,")

    def resolve(ph):
        """(phrase, brand, repair, matched) for one spelling."""
        if ok(ph):
            return (ph, brand, "", True)
        for n in NOISE:
            v = tidy(re.sub(r"\b" + n + r"\b", " ", ph, flags=re.I))
            if v and ok(v):
                return (v, brand, "dropped %r" % n, True)
        return (ph, brand, "", False)

    def sharpen(ph):
        """Strip spelling synonyms and class / grower words, and accept the
        result ONLY when what remains is a non-empty EXACT master ItemCategory
        that is not itself a class word -- so this can sharpen a match but can
        never broaden one to 'red' / 'blueberry'."""
        cand = ph
        for w, rep in SYNONYM.items():
            cand = re.sub(r"\b" + re.escape(w) + r"\b", " " + rep + " ", cand,
                          flags=re.I)
        dropped = CLASS_WORDS | GROWER_WORDS
        stripped = tidy(re.sub(r"\b(?:" + "|".join(sorted(dropped)) + r")\b",
                               " ", cand, flags=re.I))
        if not stripped or stripped == ph:
            return None
        var = P.norm(P.variety_phrase(stripped, brand_req))
        if not var or var in {w.lower() for w in CLASS_WORDS}:
            return None
        if var in cats and ok(stripped):
            return (stripped, brand, "sharpened to %r" % var, True)
        return None

    def expand(ph):
        """colour + seedless: the master names the VARIETY (Crimson, Flame, Red
        Globe...), not the class, so expand through the owner's own mapping
        (prepack_items_list) rather than collapsing to a colour word."""
        var = P.norm(P.variety_phrase(ph, brand_req))
        if var not in COLOUR_BUCKETS or " " not in var:
            return None
        colour, seed = var.split()
        hits = []
        for variety in GRAPE_BUCKETS.get((colour, seed), []):
            new = tidy(re.sub(r"\b" + colour + r"\s+" + seed + r"\b", variety,
                              ph, flags=re.I))
            if new and ok(new):
                hits.append((new, brand, "expanded %r -> %s" % (var, variety), True))
        return hits or None

    def chain(ph):
        """Every instruction one spelling should produce, best first: as-is ->
        noise -> sharpen -> colour-class expansion; else the unmatched phrase,
        so the preview still reports it instead of it vanishing."""
        r = resolve(ph)
        if r[3]:
            return [r]
        for repair in (sharpen, expand):
            out = repair(ph)
            if out:
                return out if isinstance(out, list) else [out]
        return [r]

    base = product_phrase(desc, uom, token, pack)

    # A slash list of VARIETIES is a list of products -> one instruction each,
    # and each variant still goes through the full repair chain (a variant can
    # itself need a synonym: 'FORELLA' -> 'Forelle').
    variants = [v for v in (tidy(v) for v in slash_variants(base)) if v]
    if len(variants) > 1:
        return [x for v in variants for x in chain(v)]

    bases = []
    for b in (tidy(base), tidy(base.replace("-", " "))):
        if b and b not in bases:
            bases.append(b)
    if not bases:
        return []

    for b in bases:                       # plain, then hyphen-as-separator
        out = chain(b)
        if out[0][3]:
            return out

    # a phrase word that is a BRAND in the master ('Joya' on a 'Cripps Red'
    # code, 'Moyca' on a Spain Cotton Candy code) moves to the Brand column,
    # leaving a fruit-only phrase.  Synonyms are applied first so a misspelt
    # brand is recognised as the master's own spelling.
    if brand == WILDCARD:
        corr = bases[0]
        for w, rep in SYNONYM.items():
            corr = re.sub(r"\b" + re.escape(w) + r"\b", " " + rep + " ", corr,
                          flags=re.I)
        corr = tidy(corr)
        for w in [x for x in corr.split() if len(x) > 2]:
            if w.lower() in brands:
                alt = tidy(re.sub(r"\b" + re.escape(w) + r"\b", " ", corr,
                                  flags=re.I))
                if alt and ok(alt, w.lower()):
                    return [(alt, w, "moved %r to the Brand column" % w, True)]
    return chain(bases[0])


def main():
    blank_barcodes = "--no-blank-barcodes" not in sys.argv
    wb = openpyxl.load_workbook(BOOK)
    if SRC_SHEET not in wb.sheetnames:
        raise SystemExit("ABORT: no %r sheet in %s" % (SRC_SHEET, BOOK))
    ws = wb[SRC_SHEET]
    header = [ws.cell(1, c).value for c in range(1, 8)]
    raw = [[ws.cell(r, c).value if ws.cell(r, c).value is not None else ""
            for c in range(1, 8)]
           for r in range(2, ws.max_row + 1)]

    # ---- 1. blank invalid barcodes, in the sheet -------------------------
    blanked = 0
    if blank_barcodes:
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, BARCODE_COL)
            v = str(cell.value or "").strip()
            if v and len(v) <= BARCODE_MAX:
                cell.value = ""
                blanked += 1
        wb.save(BOOK)
    print("barcodes blanked in %r: %d (rule: <=%d characters)"
          % (SRC_SHEET, blanked, BARCODE_MAX))

    rows = [[("" if v is None else str(v)) for v in r] for r in raw]

    # ---- 2a. derive a candidate instruction per source row ---------------
    VOCAB = load_master_vocab()
    print("master vocabulary: %d ItemCategories, %d ItemBrands"
          % (len(VOCAB[0]), len({b for b in VOCAB[1] if b})))

    cand, held = [], []
    for r in rows:
        prd, desc, pack, barcode = r[1], clean(r[2]), r[5].strip(), r[6].strip()
        orig = r[2]
        if pack.isdigit() and int(pack) > 20:
            held.append([prd, desc, pack, "PACK_QTY>20 = pcs/ctn (not a pre-pack)",
                         "", ""])
            continue
        if not pack.isdigit():
            held.append([prd, desc, pack,
                         "PACK_QTY is a carton weight, not a pack count", "", ""])
            continue
        uom, token = per_pack(desc)
        if not uom:
            held.append([prd, desc, pack,
                         "no per-punnet weight or count token in PRD_DESC", "", ""])
            continue
        brand_raw = ", ".join(brands_of(desc)) or WILDCARD
        ctry = countries_of(desc)
        chosen = choose_phrases(desc, uom, token, pack, brand_raw, ctry, VOCAB)
        if not chosen:
            held.append([prd, desc, pack, "empty product phrase after cleaning",
                         uom, token])
            continue
        box = box_uom(pack, token)
        for phrase, brand_out, repair, _matched in chosen:
            cand.append({"prd": prd, "orig": orig, "pack": pack, "uom": uom,
                         "token": token, "phrase": phrase, "brand": brand_out,
                         "repair": repair, "brand_raw": brand_raw,
                         "countries": ctry, "barcode": barcode})
            # ...and the CARTON its PACK_QTY describes, as a box form: the owner
            # rule is that PACK_QTY punnets of that per-punnet size IS the UOM.
            if box:
                cand.append({"prd": prd, "orig": orig, "pack": pack, "uom": box,
                             "token": token, "phrase": phrase, "brand": brand_out,
                             "repair": "box form of the %s punnet" % uom,
                             "brand_raw": brand_raw,
                             "countries": ctry, "barcode": barcode})

    # ---- 2b. a country-less description is subsumed when the SAME
    #          PRD_NUMBER also carries one that names a country -----------
    with_country = {c["prd"] for c in cand if c["countries"]}
    dropped_subsumed = []
    kept_cand = []
    for c in cand:
        if not c["countries"] and c["prd"] in with_country:
            dropped_subsumed.append([c["prd"], c["orig"],
                                     "country-less variant of a PRD_NUMBER that "
                                     "also names countries", c["uom"]])
            continue
        kept_cand.append(c)

    # ---- 2c. expand to one instruction per country, then dedupe ----------
    layout, source = [], []
    seen = set()
    dupes = 0
    for c in kept_cand:
        ctry = c["countries"] or [WILDCARD]
        note = ("country from description" if c["countries"]
                else "NO COUNTRY in description -> matches every country")
        for country in ctry:
            k = (c["phrase"], c["uom"], country, c["brand"])
            if k in seen:
                dupes += 1
                continue
            seen.add(k)
            layout.append({"product": c["phrase"], "uom": c["uom"],
                           "countries": country, "brand": c["brand"],
                           "sizes": WILDCARD})
            why = note + ("; " + c["repair"] if c.get("repair") else "")
            source.append([len(layout), c["prd"], c["orig"], c["pack"],
                           c["uom"], c["token"], country, c["brand"], why,
                           c["barcode"]])

    repaired = collections.Counter(c["repair"] for c in kept_cand if c.get("repair"))
    print("instructions derived : %d layout rows" % len(layout))
    print("deduped              : %d identical instruction(s) dropped" % dupes)
    print("subsumed & dropped   : %d country-less row(s)" % len(dropped_subsumed))
    print("held back            : %d rows" % len(held))
    print("phrases repaired     : %d" % sum(repaired.values()))
    for k, v in repaired.most_common(8):
        print("    %-52s %d" % (k[:52], v))
    per_uom = collections.Counter(x["uom"] for x in layout)
    print("UOMs                 : %s"
          % ", ".join("%s x%d" % (k, v) for k, v in sorted(per_uom.items())))
    wild = sum(1 for x in source if x[6] == WILDCARD)
    print("rows with wildcard country (match-everything): %d layout rows from %d source rows"
          % (wild, len({s[1] for s in source if s[6] == WILDCARD})))

    # ---- 3. write the layout workbook -----------------------------------
    out = openpyxl.Workbook()
    sh = out.active
    sh.title = "Sheet1"
    sh.append(["Product description", "UOM", "Countries", "Brand", "Sizes"])
    for x in layout:
        sh.append([x["product"], x["uom"], x["countries"], x["brand"],
                   x["sizes"]])
    for c in range(1, 6):
        sh.cell(1, c).font = Font(bold=True)
    sh.freeze_panes = "A2"

    s2 = out.create_sheet("Source")
    s2.append(["LayoutRow", "PRD_NUMBER", "PRD_DESC (original)", "PACK_QTY",
               "UOM", "PackToken", "Country", "Brand", "Note",
               "BARCODE (original value, blanked in Unique_Products if <=10 chars)"])
    for row in source:
        s2.append(row)
    for c in range(1, 11):
        s2.cell(1, c).font = Font(bold=True)
    s2.freeze_panes = "A2"

    s3 = out.create_sheet("Report")
    s3.append(["PRD_NUMBER", "PRD_DESC (cleaned)", "PACK_QTY", "Why no instruction",
               "UOM that would apply", "PackToken"])
    for row in held:
        s3.append(row)
    for row in dropped_subsumed:
        s3.append([row[0], row[1], "", row[2], row[3], ""])
    for c in range(1, 7):
        s3.cell(1, c).font = Font(bold=True)
    s3.freeze_panes = "A2"

    out.save(OUT)
    print("\nwritten: %s" % OUT)
    print("  Sheet1 : %d instructions" % len(layout))
    print("  Source : traceability per layout row")
    print("  Report : %d row(s) with no instruction"
          % (len(held) + len(dropped_subsumed)))


if __name__ == "__main__":
    main()
