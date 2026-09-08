# UDF_UOM_UOMDesc curation rules

This folder records the rules we apply to the **`UDF_UOM_UOMDesc`** column (column D) of **`imported_item_codes.xlsx`**, the master list of item codes already imported into Autocount.

Each rule keys off the **`UOM`** column (column T) of the same row. When cleaning a description, always strip only the offending tokens and keep the product phrase (`Pre Pack <Country> [Brand] <Variety> <Fruit> <pack info>` structure) intact.

| Our language | Column |
|---|---|
| Description under curation | `UDF_UOM_UOMDesc` |
| UOM of the row | `UOM` |
| Fruit | `ItemType` |
| Variety | `ItemCategory` |
| Country | `ItemClass` |
| Grade | `UDF_Grade` |
| Brand | `ItemBrand` |

---

## Rule 1 — PKT rows: keep grade only for Envy apples

**Scope:** rows whose `UOM` starts with `PKT`.

A PKT description must **not** carry the item's grade (the value in `UDF_Grade`, e.g. `(PG)`, `(HG)`, `SG`, `BLUSH`, `Class 1`, `PREMIUM`) — **except** for **Envy apples**.

- **Envy apples** = `ItemType` = `Apples` and `ItemCategory` = `Envy` or `Organic Envy` (covers `Envy`, `Organic Envy`, and either with a trailing space). These **keep** (and should show) their grade: `Pre Pack NZ Enza Envy (HG) Apples 2s`, `Pre Pack USA Organic Envy (AG) Apples 4s`.
- **Every other PKT row**: remove the grade token(s) from the description (`Pre Pack NZ Cherish (PG) Apples 4s` → `Pre Pack NZ Cherish Apples 4s`).
- Note: Envy / Organic Envy rows that genuinely have no grade (`UDF_Grade` empty) are fine without one — there is nothing to add.

---

## Rule 2 — PCS rows: single P/PCS ending, no carton weight, no size

**Scope:** rows where `UOM` = `PCS`.

A PCS description must:
- **(a)** end with **`P/PCS`** — exactly once (no repeats, no `P/PKG`/`P/PCS P/PCS`, nothing after it). Appended if missing.
- **(b)** not contain **fruit/carton sizes** — `88s`/`125s` (`\d+s`), `10ROW`, count ranges like `60/64`, pack weights like `800G`/`300G`, or trailing size letters like `S`.
- **(c)** not contain **carton weights** — `17.5kg`/`18kg` (`\d+(\.\d+)?kg`, case-insensitive; also `LB`).

Examples:
- `China Fuji Apples 2.5kg 8s P/PCS` → `China Fuji Apples P/PCS`
- `Aust Evo Hass Avocados 5kg 28s` → `Aust Evo Hass Avocados P/PCS`
- `SA Packham Pears P/PKG` → `SA Packham Pears P/PCS`
- `USA Summertime White Peaches 60/64` → `USA Summertime White Peaches P/PCS`

Keep digits that belong to the product itself (grade `XF1`/`XF2`, variety `86 Wang Melons`, `M7`, `C5`).

---

## Rule 3 — GM rows: single P/KGS ending, no carton weight, no fruit size

**Scope:** rows where `UOM` = `GM`.

A GM description must:
- **(a)** end with **`P/KGS`** — exactly once (replace `P/PKG`, `P/PKT`, `P/GM`; append if missing).
- **(b)** not contain **carton weights** — `9kg`/`17.5kg` (`\d+(\.\d+)?kg`, case-insensitive; also `19LB`).
- **(c)** not contain **fruit sizes** — `9.5ROW`/`10 ROW`, `23MM`/`28mm`, cherry size letters (`J`/`JJ`/`SJ`/`M`), counts (`4s`), `20B`.

Examples:
- `Vietnam Cavendish Bananas 13kg 4s` → `Vietnam Cavendish Bananas P/KGS`
- `USA Borton Fruit Sweet Cherries 9kg 9.5 ROW` → `USA Borton Fruit Sweet Cherries P/KGS`
- `China Fragrant Pears P/PKG` → `China Fragrant Pears P/KGS`

A clean GM description contains **no digits at all** — use that as a hard check after cleaning.

---

## Rule 4 — PKT rows: brands limited to whitelist

**Scope:** rows whose `UOM` starts with `PKT`.

A PKT description may carry a brand only when the item's `ItemBrand` is in the **whitelist**:

- **Zespri, Rockit, Enza, Sunkist, Barnfield**
- plus (user addendum): **Soluna, Juliet, Organic Juliet**

**Whitelist-brand items should actually show the brand.** When a whitelisted brand is missing from the description, insert it right after the country token — `Pre Pack SA Lemons 3s` → `Pre Pack SA Sunkist Lemons 3s` (349 Sunkist rows were added this way).

**Non-whitelist brands already present in the description are removed** (keep the variety):
- `Pre Pack SA Joya Cripps Red Apples 3s` → `Pre Pack SA Cripps Red Apples 3s` (removed `Joya`)

**Exceptions (leave as-is):**
- The recorded brand is **also the variety name** and appears once in the description — e.g. `Red Pop`, `Pre Pack Italy Red Pop Apples 4s`. This reads as a variety, not a brand.
- Words appearing twice as brand+variety duplication — e.g. `Pre Pack NZ Smitten Smitten Apples 4s`, `Pre Pack NZ Piqa Piqa Pears 3s` — remove **one** copy so the variety stays.

**Deliberately open (do not touch without asking):** Enza (77 rows) and Barnfield (44 rows) PKT descriptions that are missing the brand were left as-is. Barnfield rows are messy — several already contain `Sunkist` (also whitelisted) instead.

---

## Rule 5 — Packing-form UOM rows (`16X6PCS`, `8X500G`, `7X1KG`): no `Pre Pack`, compact `<n>x<m>` box-config ending

**Scope:** rows whose `UOM` is a packing form — `<count>x<size><unit>` e.g. `16x5pcs`, `14X6PCS`, `8x500g`, `12X250G`, `7x1kg`. These describe a box/repack of N inner packs each of size M, **not** a consumer pre-pack (which is what the `Pre Pack` prefix and the `PKT` rules are for). They currently appear only on the export side of a merge (`imported_item_codes.xlsx` itself has no such UOM rows yet).

A packing-form description must:
- **(a)** **not** start with `Pre Pack` — drop the prefix if present.
- **(b)** carry the item's product phrase as in its **`Description`** column — `<Country> [Brand] <Variety> <Fruit>` — which fixes brand/variety typos (`Glod` → `Gold`, missing `Organic`, singular/plural fruit-word mismatches) and drops any stray CTN carton weight (`SA NVF Granny Smith Apples 18kg 16x1kg` → `SA NVF Granny Smith Apples 16x1kg`). Cherries / organic packing rows are then subject to Rules 6 and 7 too — do **not** pull a cherry cultivar back in from `Description`.
- **(c)** end with the box config spelled compact, like the UOM value: `<count>x<size>` with a lower-case `x`, and the size unit as `s` (pieces), `g`, or `kg` (`14X6PCS` → `14x6s`, `16X1KG` → `16x1kg`, `12X250G` → `12x250g`), with no CTN carton-weight/count tail. (The long `NPktsXM` spelling is not used.)

Examples:
- `Pre Pack NZ Enza Envy (AG) Apples 14PktsX6Pcs` → `NZ Enza Envy (AG) Apples 14x6s`
- `Pre Pack SA Dutoit Granny Smith Apples16PktsX1kg` → `SA Dutoit Granny Smith Apples 16x1kg`
- `SA SY Eureka Lemon 18PktsX6Pcs` → `SA SY Eureka Lemons 18x6s` (fruit word pluralised to match `Description`)

Applied by `python_scripts/curate_uomdesc_packing_rows.py`.

---

## Rule 6 — Cherries: cultivar dropped from every UOM row

**Scope:** rows whose fruit (`ItemType`) is `Cherries` — this applies to **all** their UOM rows (`GM` base, `PKT` pre-packs, `CTN`, packing forms), not just one type of row.

A cherries description never carries the cultivar (the value in `ItemCategory` — `Bing`, `Skeena`, `Lapin`, `Santina`, `Black Pearl`, `Sweetheart`, `Royal Lynn`, `Sweet`, `Red`, `White`, `Dark Sweet`, `Ziraat 900`, …). The cultivar lives in the item code and `Description`; it is not repeated in `UDF_UOM_UOMDesc`.

- `USA Gold Prize Skeena Cherries 5kg 9.5 ROW` → `USA Gold Prize Cherries 5kg 9.5 ROW`
- `Chile Bing Cherries 5kg 28-30MM` → `Chile Cherries 5kg 28-30MM`
- `USA Skeena Cherries 12x250g` → `USA Cherries 12x250g`

Keep country and brand; drop only the cultivar/type words.

---

## Rule 7 — Organic items: the word `Organic` on every UOM row

**Scope:** items whose `Description` (or `ItemCategory`) marks them `Organic` (e.g. `Organic Green Kiwis`, `Organic Juliet`).

Every UOM row of an organic item must carry the word `Organic` in `UDF_UOM_UOMDesc`, spliced in at the same position it holds in the `Description` (`<Country> <Brand> Organic <Variety> <Fruit>`).

- `France Juliet Apples P/PCS` → `France Organic Juliet Apples P/PCS`
- `France Juliet Apples 18kg 163s` → `France Organic Juliet Apples 18kg 163s`

Applied together with Rule 6 by `python_scripts/curate_uomdesc_cherries_organic.py`.

---

## Rule 8 — PKT·PCS·GM rows: pack quantity only, no fruit calibre (MM / ROW)

**Scope:** rows whose `UOM` is `PKT` (bare or numbered), `PCS`, or `GM`.

These rows describe the pack sold (a pre-pack or the base unit), so their UOMDesc carries the pack quantity (`125G`, `1KG`, `4s`, `P/PCS`, `P/KGS`) but **never** the fruit calibre. Drop any MM-size / ROW-count token (export-only rows can still carry them, e.g. `125G 14MM+`):

- `Zimbabwe Atlas Blue Blueberries 125G 14MM+` → `Zimbabwe Atlas Blue Blueberries 125G`
- `Spain Conference Pears 1KG 50-55MM` → `Spain Conference Pears 1KG`

The calibre belongs on the `CTN` row / `Description`, which keeps it.

Applied by `python_scripts/curate_uomdesc_no_fruit_size.py`.

---

## Rule 9 — Blueberries: cultivar dropped from every UOM row

**Scope:** rows whose fruit (`ItemType`) is `Blueberries` — applies to **all** their UOM rows (`PKT` pre-pack / base, `CTN`, …).

A blueberries description never carries the cultivar (the value in `ItemCategory` — `Atlas Blue`, `Azrablue`, `Biloxi`, `Eureka`, `Eureka Sunrise`, `Sekoya Pop`, `Snowchaser`, `Splash`, `Ventura`, …). The cultivar lives in the item code / `Description`; it is not repeated in `UDF_UOM_UOMDesc`.

- `SA Gen Azrablue Blueberries 12X125G 14MM+` → `SA Gen Blueberries 12X125G 14MM+`
- `Zimbabwe Eureka Sunrise Blueberries 200G` → `Zimbabwe Blueberries 200G`

Keep country and brand. **`Organic` is preserved**: an `Organic X` item keeps the `Organic` and drops only the cultivar (`Peru Divino Organic Biloxi Blueberries …` → `Peru Divino Organic Blueberries …`), and a generic `Organic` row (no cultivar) is untouched (Rule 7).

Applied by `python_scripts/curate_uomdesc_blueberry_cultivars.py`.

---

## Rule 10 — Collapse repeated word/phrase runs (any UOM)

**Scope:** the whole `UDF_UOM_UOMDesc` column.

Some descriptions repeat a word because the same token is both the brand and the variety (e.g. `Piqa`), or a grade / phrase is written twice (`BLUSH BLUSH`, `Class 1 Class 1`). Collapse the immediate adjacent repeat, on any UOM row:
- `NZ Piqa Piqa Pears P/PCS` → `NZ Piqa Pears P/PCS`
- `Pre Pack NZ Evercrisp Evercrisp Apples 4s` → `Pre Pack NZ Evercrisp Apples 4s`
- `China Fuji BLUSH BLUSH Apples 17kg 64s` → `China Fuji BLUSH Apples 17kg 64s`

Leave genuinely doubled names alone when the repeated phrase equals the item's variety or brand — `Cara Cara` oranges stays `Cara Cara`, `Crane & Crane` is untouched. Interleaved / case-variant rows (e.g. `Evercrisp (Pg) Evercrisp (PG) …`) are also left as-is rather than mangled.

Applied by `python_scripts/curate_uomdesc_dupe_words.py`.

---

## Rule 11 — Figs: variety dropped from every UOM row

**Scope:** rows whose fruit (`ItemType`) is `Figs` — applies to **all** their UOM rows (`PKT` pre-pack / base, `CTN`, …).

A figs description never carries the variety (the value in `ItemCategory` — `Evita`, `Parisian`, `Madeira`, …). The variety lives in the item code / `Description`; it is not repeated in `UDF_UOM_UOMDesc`.

- `Pre Pack SA Evita Figs 160G` → `Pre Pack SA Figs 160G`
- `Pre Pack SA Parisian Figs 160G` → `Pre Pack SA Figs 160G`
- `Pre Pack Peru Organic Madeira Figs 160G` → `Pre Pack Peru Organic Figs 160G`

Keep country and brand. **`Organic` is preserved**: an `Organic X` item keeps the `Organic` and drops only the variety (Rule 7).

Applied by `python_scripts/curate_uomdesc_unmatched_new.py` (R11).

---

## Workflow note

`imported_item_codes.xlsx` is the live master and is git-tracked. Before any bulk edit, save a timestamped backup (`imported_item_codes_backup_<MMM-dd>_<hhmmam/pm>.xlsx`) next to it, apply changes, then re-run the rule checks to confirm 0 remaining violations and that only the intended rows changed.
