# Restructure Autocount-exported items (autocount_items.xlsx)

## Grand Purpose (Why)
A direct export from Autocount (e.g. **autocount_items.xlsx** at the repo root) names its columns
with Autocount's display headers — spaces, no `UDF_` prefix, e.g. `Fruit Size (Size)`,
`Amount per CTN (SizeUOM)`, `Need to re-pack ?` — and arranges them in Autocount's own order. Our
curated master **imported_item_codes.xlsx** uses a different naming and column order. Before such an
export can be diffed, merged, or placed side-by-side with the master, it must be restructured so the
two line up column-for-column: same order, same header names, readable widths.

Note: the two files hold the *same* item data — `Description`/Item Code values match across ~16.6k
rows (e.g. `APRCT-NA-CN-0001` appears identically in both). Only the header layer differs.

## What's in here
| File | Purpose |
|---|---|
| `header_map_imported_item_codes_to_autocount_items.json` | Header↔header map between the two workbooks, both directions (objects `imported_item_codes_to_autocount_items` and `autocount_items_to_imported_item_codes`). A `null` value = no counterpart. Built by reading the actual header rows; every header is covered and every non-null target is a real header in the other file. |
| `reorder_autocount_columns.py` | Reorder `autocount_items.xlsx` columns to mirror `imported_item_codes.xlsx` column order. Autocount-only columns are appended at the end in their original relative order. Cell values **and** styles travel with their header. Writes a timestamped backup, prints the plan, then verifies all ~807k cells by full re-read. |
| `rename_autocount_headers.py` | Rename each column header of `autocount_items.xlsx` to its `imported_item_codes.xlsx` name via the reverse map. Headers with no counterpart (`Bar Code`, `Measurement`) keep their name. Touches row 1 only; writes a backup and verifies the data rows are unchanged. |
| `autofit_autocount_columns.py` | Autofit column widths to the widest header/content (+2 units), counting 2-char-wide glyphs (CJK etc.) double. Writes a backup and verifies widths persisted. |
| `sort_itemcode_rate.py` | Group data rows by `ItemCode` (group order = first occurrence) and sort each group ascending by `Rate`, mirroring the imported master (base unit Rate 1, prepacks, CTN last). Stable; equal Rates keep original relative order. Writes a backup and verifies by full re-read. |
| `merge_matched_unmatched.py` | Split a restructured export against `imported_item_codes.xlsx` on `(ItemCode, Description)` into three files in **`outputs/`**: matched (in both; master values win, export-only UOM variants left as-is), unmatched-new (export-only, fruit-SKU code shape), imported-only (master rows missing from the export). See "Merging…" below. |

## How to apply to a freshly exported autocount items file
Run from anywhere (the scripts use absolute paths). Each script takes the target filename as an
optional first argument; it defaults to `autocount_items.xlsx` at the repo root. The target is
restructured in place. Steps in order:
1. `python restructure-autocount-exported-items/reorder_autocount_columns.py <file>.xlsx`
2. `python restructure-autocount-exported-items/rename_autocount_headers.py <file>.xlsx`
3. `python restructure-autocount-exported-items/autofit_autocount_columns.py <file>.xlsx`
4. `python restructure-autocount-exported-items/sort_itemcode_rate.py <file>.xlsx`

e.g. for a fresh export saved as `autocount_items_sep2nd_1125pm.xlsx`:
1. `python restructure-autocount-exported-items/reorder_autocount_columns.py autocount_items_sep2nd_1125pm.xlsx`
2. `python restructure-autocount-exported-items/rename_autocount_headers.py autocount_items_sep2nd_1125pm.xlsx`
3. `python restructure-autocount-exported-items/autofit_autocount_columns.py autocount_items_sep2nd_1125pm.xlsx`
4. `python restructure-autocount-exported-items/sort_itemcode_rate.py autocount_items_sep2nd_1125pm.xlsx`

Each step self-verifies against a pre-step snapshot and prints its plan; each writes a timestamped
backup `<target-stem>_backup_<MMM-dd_HHMMSSpm>.xlsx` beside the source before changing it. Run the
steps on a **raw export**: `reorder` expects the original display-header layout, so apply it only
once per export (re-running it on an already-restructured file will abort on its layout assertion).
`rename`, `autofit` and `sort` are safe to re-run. (Backups only ever overwrite when two steps run in
the same minute — the timestamp includes seconds to avoid that.)

If a target's headers differ from the ones in the JSON, regenerate the JSON first — see below.

## Sorting data rows within an item code (ItemCode + Rate)
Each `ItemCode` normally has several UOM rows (base unit, prepacks, CTN). The imported master lists
a code's rows base-unit first (Rate = 1), then prepacks (Rate 3/6/8…), with the CTN row
(Rate = carton count) **last**. A raw Autocount export leads with the CTN row instead. To match the
master, `sort_itemcode_rate.py`:
- Groups rows by `ItemCode`, preserving the group order of the file (first occurrence), so the
  overall sequence of item codes is not alphabetised or otherwise shuffled.
- Within each group, stably sorts rows ascending by `Rate` (numeric). Equal Rates keep their
  original relative order; a missing/unparseable `Rate` sorts to the end of its group.
- Reorders data rows only — the header row is untouched, and each row's values + cell styles move
  together. Header names and column order (from the earlier steps) are unaffected.

## Check: `UDF_UOM_ADJdedcut` (`Need to re-pack ?`)
Each UOM row's `Need to re-pack ?` flag becomes `UDF_UOM_ADJdedcut` after rename. Validate the
column per row: it may only hold `T` or `F`, and the value must agree with the row's `UOM`:
- `UOM` = `PCS`, `CTN`, `GM`, or bare `PKT` (the base unit, e.g. a `10X500G` item's `PKT` row) → value must be `F`.
- `UOM` starts with `PKT` but is not bare (e.g. `PKT3`, `PKT10`, `PKT500`) → value must be `T`.
- `UOM` is a packing-form unit such as `16x5s`, `16x5pcs`, `10x500g`, `8x1kg` → value must be `T`.

Any row whose flag is blank, not `T`/`F`, or disagrees with its `UOM` under these rules fails the
check. Apply it to a restructured export (and to the merged `outputs/`) after `sort_itemcode_rate.py`.

## Merging an export against imported_item_codes.xlsx
`python merge_matched_unmatched.py <file>.xlsx` inner-joins a **restructured** export to the master on
`(ItemCode, Description)` and writes three files (45-column template = imported headers minus
`Column1`) into the **`outputs/`** subfolder here:
- `<stem>_matched_imported_<stamp>.xlsx` — export rows whose key exists in the master.
  `ItemCode`/`Description` are identical on both sides (the join key). For the **other** columns the
  **master's values win**: each row is paired to the master row with the same `(ItemCode,
  Description, UOM)` and copies its columns. Export rows whose UOM variant has no twin in the
  master (the export has more UOM rows than the master for that key) are left as-is with the
  export's values, and only those rows get `UNK` in the four columns the export has no header for
  (`IsCalcBonusPoint`, `StockControl`, `HasSerialNo`, `CostingMethod`).
- `<stem>_unmatched_new_<stamp>.xlsx` — export rows whose key is absent from the master, kept only
  for ItemCodes of the fruit-SKU shape "3 letter groups + 4-digit number" (e.g. `APL-GLA-TUR-0001`,
  `BLUBRY-ORG-USA-0001`); numeric / `ADJ-*` / `BRD-*` / `BM-*` codes are dropped.
- `imported_rows_missing_from_<stem>_<stamp>.xlsx` — master rows whose key is absent from the
  export (values from the master, so no `UNK`).

As of the Sep-02 11:25pm export these were 16,421 matched (15,997 copied from master, 424
export-only UOM variants left as-is) / 1,785 unmatched-new / 842 imported-only rows. (If a desired
output is "whole codes absent from Autocount" rather than exact `(code, description)` pairs, switch
the join to ItemCode alone.)

## Column-order rules
- Follow the header order of `Sheet1` in `imported_item_codes.xlsx`.
- A column present in both files is placed at the position of its imported counterpart.
- Autocount-only columns are appended at the end in their original relative order.
- Notable moves vs. the raw export: `PkgDesign` moves up near the front; `UserUOM` sits just before
  `UOM`; `Classification`, `HasBatchNo`, `DutyRate`, the tax-code columns and `TariffCode` are
  regrouped near the tail to mirror the imported master.

## Headers with no counterpart
- In `imported_item_codes.xlsx` only (their JSON value is `null` going export→imported): `Column1`
  (always empty), `IsCalcBonusPoint`, `StockControl`, `HasSerialNo`, `CostingMethod`.
- In `autocount_items.xlsx` only (kept as-is when renaming): `Bar Code`, `Measurement` (empty).

## Headers that look unchanged but are mapped
`Description`, `Rate`, `UOM` and `Classification` are the same literal string in both files, so a
rename leaves them looking identical — they *do* have counterparts.

## Mapping caveats
- Some column pairs hold overlapping values (a size like `88` equals its SizeUOM; `UOM` equals
  `Base UOM` on certain rows), so pair headers by name/semantics — including Autocount's parenthetical
  hints (`Fruit Size (Size)`, `Amount per CTN (SizeUOM)`, `Prepacked Config (Packing)`, `Need to
  re-pack ?`) — rather than by matching cell values.
- A few value dialects differ even after alignment: `HasBatchNo` spells its flag `Checked`/`Unchecked`
  in Autocount vs `T`/`F` in the master; price columns use `-1` sentinels in the export. Normalise as
  needed after alignment.
- `imported_item_codes.xlsx` has no native autofit either; widths are estimated from text length
  (wide glyphs count double), so column width ≈ longest visible string + 2.

## Regenerating the header map
If either file's headers change, re-create the JSON by reading both header rows and aligning by
semantics, then re-validate (every source header covered; every non-null target is a real header in
the other file). This is what the current JSON was checked against.
