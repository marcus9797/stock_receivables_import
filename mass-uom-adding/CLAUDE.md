# mass-uom-adding

Bulk-add a **new pre-pack UOM** (e.g. `PKT8`, `PKT500`, `PKT800`) to **many existing item codes** in `imported_item_codes.xlsx` at once — safely, behind a reviewable preview and a human approval gate. It never guesses: every candidate row is surfaced first, the owner reviews/annotates, and only after approval is the master touched (behind a backup).

## The input — `mass_UOM_edit_layout.xlsx` (project root)

A human-edited instruction sheet. **One row = "add this UOM to every existing code that matches, unless it already has it."** Columns:

| Col | Header | Meaning | Wildcard |
|---|---|---|---|
| A | Product description | loose product phrase, e.g. `SA Granny Smith`, `USA Cherries`, `Organic Royal Gala`, `Zespri sungold kiwi` | — |
| B | UOM | the pre-pack to add (`PKT2…PKT12` count packs; `PKT500/750/800/1000` weight punnets) | — |
| C | Countries | which countries of origin to target (ItemClass) | `does not matter` |
| D | Brand | grower/house brand to target (ItemBrand) | `does not matter` |
| E | Sizes | which `UDF_UOM_Size` values to target | `does not matter` |

Write sizes comma-separated (`198, 216, 237`) — numeric and text sizes both work (`100`, `36s`, `120G`, `150G+`). Case is ignored. A row that spells a product **without** the word `Organic` will not match `Organic …` codes (organic is curated separately).

## Scripts (this folder) — run from the project root

| Script | What it does | Writes |
|---|---|---|
| `mass_uom_add_preview.py` | Reads the layout + master, matches codes, computes each candidate row's `Rate`/`SizeRate`/`UOMDesc` (provisional) and marks ADD / ALREADY-HAS / VERIFY-RATE / SKIP | `mass_uom_add_preview_<ts>.xlsx` (Summary + Rows). Never touches the master. |
| `mass_uom_add_apply.py` | Applies the APPROVED preview: dedupes, backs up the master, inserts the ADD rows before each code's `CTN` row, validates, writes a staged file, then promotes over the live master | `imported_item_codes_backup_<ts>.xlsx`, `imported_item_codes_massUOM_<ts>.xlsx`, and the promoted `imported_item_codes.xlsx` |
| `mass_uom_add_flag_new.py` | Flags the rows the apply added, by identity (`ItemCode`, `UOM`, `UOMDesc` from the approved preview) — robust to later manual edits | `mass_uom_add_new_rows_<ts>.xlsx` (NewRows + NotFlagged sheets) |

## The workflow (canonical run)

1. **Owner** edits `mass_UOM_edit_layout.xlsx` (one product/UOM/… combination per row).
2. Run preview → read **Summary** (matched / to-add / already-have / skipped per layout row) and **Rows** (per-code detail: matched code, `BaseUOM`, size, proposed `Rate`, `RateSource`, `g-per-piece`, provisional `UOMDesc`, `Reason`).
3. **Owner reviews Rows**, especially:
   - `VERIFY-RATE` rows → type the decision into a `my response` column (rate, or skip, or note a rule), and
   - rows that show "all have it" (no-op) to confirm they're redundant or a size list should widen.
4. **Convert** the owner's decisions into `OVERRIDES` in `mass_uom_add_preview.py` (each entry = ADD-with-rate or SKIP, with the reason on the line) and re-run the preview until it shows 0 `VERIFY-RATE` and the intended adds. Record any new `UOMDesc` rule in `customer-desc-column-curation/CLAUDE.md`.
5. **Owner approves** the final preview.
6. Run `mass_uom_add_apply.py <approved-preview.xlsx>` — it refuses to promote if it would create a duplicate it didn't already have.
7. Optionally run `mass_uom_add_flag_new.py` to get a workbook flagging exactly the newly added rows.

## Matching rules

- **Variety = col A minus country/brand/fruit words**, compared against the code's `ItemCategory`: **exact first**, falling back to "contains" only when exact finds nothing (e.g. `Navelate` lives in `Navelate Navel` / `Cambria Navelate`).
- Country → `ItemClass` (full name; short forms `SA`, `NZ`, `USA`, `CN`, `AUS`… translated). Brand → `ItemBrand`. Sizes → `UDF_UOM_Size`.
- When col A names a fruit but no variety (`USA Cherries`), it matches **every variety** of that fruit in the target countries (`ItemType` = Cherries) — the "all cherries" case.
- Ambrosia row is special-cased to NZ **+ Italy** (owner decision — Ambrosia is grown in both).

## Rate rules (owner-confirmed)

| Scenario | Rate |
|---|---|
| Count pack `PKT<n>` on PCS/PKT-base code | `n` |
| Weight pack `PKT<g>` on a **GM**-base code | grams (`PKT500` → 500) |
| Weight pack on a **loose-count PCS** code | `ceil( pack_g / grams_per_piece )`, where `grams_per_piece = UDF_UOM_WeightPerCtn ÷ UDF_UOM_SizeUOM` (WeightPerCtn inferred from `…kg` in the Description if blank) |
| Weight pack on a **boxed/multi-pack carton** (`10x850G`, `8X6PCS`, `20X450G`) or on a **PKT**-base code | not auto-derivable → `VERIFY-RATE` for the owner; decided per code (examples: bagged Packham `10x800G` → `PKT800` `Rate = 1`; `8X6PCS`/`20X450G` PKT-base → **skip, don't add**) |

Every added row: `UOM` = the pack, `UOMDesc` = `Pre Pack …`, `Rate` = `SizeRate`, `UDF_UOM_UserUOM` = `PKTS`, `UDF_UOM_ADJdedcut` = `T`, inserted inside the code group just before its `CTN` row. `UOMDesc` is generated per the curation rules in `customer-desc-column-curation/` — whitelist grower brand after the country, **house brand first** (Rule 12, e.g. `Pre Pack Garden Basket China Fragrant Pears 750G`), Envy keeps its grade, cultivar dropped for cherries/blueberries/figs.

## Safety & pitfalls (learned the hard way)

- The master is git-tracked and lives **open in Excel on a schedule** — an Excel re-save can silently overwrite a promotion or change rows. Before promoting, ensure `imported_item_codes.xlsx` is closed. The apply script always writes a staged file first; if promotion fails (file locked) the staged file is the result.
- **Layout rows overlap.** Two layout rows can both target the same `(code, UOM)` (e.g. SA Granny `PKT8` sizes 198–237 *and* 135–198). Preview counts both; apply must **dedupe to one row per (code, UOM)** (337 instructions → 318 unique rows in the first real run).
- The master already contains a few pre-existing duplicate rows (`BNA-NA-VIE-0001` etc.) — the apply's duplicate guard compares **before vs after**, and only rejects *new* duplicates.
- After an apply, the owner may hand-edit codes (e.g. re-base bagged Packham `PCS` → `PKT`). The flag script identifies added rows **by identity**, not by backup-diff, so those later edits don't pollute it.
- `UDF_UOM_ADJdedcut` should be `T` for UOMs starting with `PKT` and box forms, `F` otherwise — but ~400 bare-`PKT` **base** rows in the master are still `F` (a known, deliberately-left grey area; see `item-code-importing-table-structure/CLAUDE.md`).

## Related docs
- `item-code-importing-table-structure/CLAUDE.md` — the master's column layout, row-role model and `UOM` block (read this first if unsure what a column means).
- `customer-desc-column-curation/CLAUDE.md` — the `UDF_UOM_UOMDesc` rules the generated text follows.
- `create_new_item_code/CLAUDE.md` — how brand-new codes/rows are normally generated when an item doesn't exist at all.
