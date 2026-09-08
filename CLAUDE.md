# Autocount SKU creation

## Grand Purpose (Why)
We need to import existing in-stock items into Autocount Malaysia (ERP software). Prior to that, a template must first be followed and filled with the relevant info.

# Directory structure

stock_receivables_import/                    # Root of all workspaces
|--- CLAUDE.md
|--- Import_stock_receive_template.xlsx
|--- imported_item_codes.xlsx
|--- mass_UOM_edit_layout.xlsx        # instruction sheet: "add this UOM to every code matching product/country/brand/size"
|--- varieties.yaml
|--- fruits.yaml
|--- countries.yaml
|--- search_item_code/
    |--- CLAUDE.md
│--- create_new_item_code/
    |--- CLAUDE.md
    |--- item_code_hashmap.json
    |--- create_new_codes.py            # generate new item codes → New_Codes sheet in imported_item_codes.xlsx
    |--- process_round2.py              # resolve Ambiguous_Matches corrections; create new codes → to_add_item_codes_*.xlsx
    |--- rebuild_all.py                 # regenerate new codes + apply to Sheet1 + clean Ambiguous_Matches
    |--- generate_to_add_item_codes_aug06.py  # build to_add_item_codes_*.xlsx for a specific new-code batch (e.g. Aug-06)
|--- current_stock_lists/
    |--- USJ/
        |--- USJ_stock_list.xlsx
    |--- BM/
        |--- BM_stock_list.xlsx
    |--- GM/
        |--- GM_stock_list.xlsx
|--- get_full_container_no/
    |--- CLAUDE.md
    |--- ETA_containers.xlsx
|--- restructure-autocount-exported-items/   # make an Autocount-exported items file line up with imported_item_codes.xlsx, then split it against the master
    |--- CLAUDE.md
    |--- header_map_imported_item_codes_to_autocount_items.json
    |--- reorder_autocount_columns.py
    |--- rename_autocount_headers.py
    |--- autofit_autocount_columns.py
    |--- sort_itemcode_rate.py
    |--- merge_matched_unmatched.py
    |--- outputs/                          # matched / unmatched-new / imported-only merge files
|--- item-code-importing-table-structure/   # structural reference for imported_item_codes.xlsx (columns, row grouping, UOM roles)
    |--- CLAUDE.md
|--- mass-uom-adding/                   # bulk-add pre-pack UOMs to existing codes (mass_UOM_edit_layout.xlsx -> preview -> approval -> apply)
    |--- CLAUDE.md
    |--- mass_uom_add_preview.py
    |--- mass_uom_add_apply.py
    |--- mass_uom_add_flag_new.py
|--- python_scripts/                    # stock-list → import-template processing (item-code-creation scripts live in create_new_item_code/)

# Floor plan
## You want to .. | Go here.. |
View ultimate template to follow prior to inserting into Autocount | **Import_stock_receive_template.xlsx**
View existing item codes already imported in Autocount | **imported_item_codes.xlsx**
Find an existing item code, given an item description | **search_item_code/**
Create new item codes | **create_new_item_code/**
Generate new item codes via python scripts (→ to_add_item_codes_*.xlsx / New_Codes sheets) | **create_new_item_code/** (see `create_new_codes.py`, `process_round2.py`, `rebuild_all.py`)
View existing stock lists to understand their input templates | **current_stock_lists/**
View existing fruits we have in Autocount | **fruits.yaml**
View existing varieties we have in Autocount | **varieties.yaml**
View existing countries we have in Autocount | **countries.yaml**
View existing grades we have in Autocount | **grades.yaml**
View existing brands we have in Autocount | **brands.yaml**
Get full container numbers and reflect them into stock_receive_template.xlsx | **get_full_container_no/**
View the curation rules for the UDF_UOM_UOMDesc column in imported_item_codes.xlsx | **customer-desc-column-curation/**
Restructure an Autocount-exported items file (reorder columns, rename headers, autofit) so it lines up with imported_item_codes.xlsx | **restructure-autocount-exported-items/**
Understand the full table structure of imported_item_codes.xlsx (columns, row grouping, UOM roles, conventions) | **item-code-importing-table-structure/**
Add a pre-pack UOM (PKT…) in bulk across many existing item codes (mass_UOM_edit_layout.xlsx input → preview → approval → apply) | **mass-uom-adding/**

# imported_item_codes.xlsx
These show the currently imported item codes within Autocount. Use this to find corresponding field values for Variety, Country, Fruit etc.

## In the following key-value pairs, you'll find the respective field columns inside the xlsx for the relevant info. 

Key = our language, Value = respective column inside the xlsx.

Variety: ItemCategory
Country of Origin: ItemClass
Fruit: ItemType
Brand: ItemBrand
Size: UDF_UOM_Size
Packing (e.g. 16x800G): UDF_OriginallyPrepacked
Grade: UDF_Grade
Carton Weight (e.g. 17kg): UDF_UOM_WeightPerCtn, expressed in grams.
BaseUOM: UOM. One of three values:
  - **PCS** — piece count per carton (e.g. apples size 80 = 80 pcs/ctn)
  - **PKT** — inner packets per carton (e.g. blueberries 10×200G)
  - **GM** — grams, weight-based (e.g. cherries, grapes, and other fruits sold loose by weight)
  
  Every item code must have at least one CTN row, plus a row for its base UOM. The baseUOM for an existing item is whatever `imported_item_codes.xlsx` records — don't guess it from the fruit type. For new items, infer it from the closest matching existing item or from stock list context (weight → GM, packing → PKT, piece count → PCS).

**item_code_hashmap.json** (in **create_new_item_code/**): Knowledge base that stores our countries, fruits and varieties. The important parts here are the Keys, not values. Use this when looking at a description and deciphering which word is a country/fruit/variety. As for what a description means, refer to **BM_stock_list.xlsx**.

# Import_stock_receive_template.xlsx
This is the template to adhere to before finally submitting to Autocount. 

NOTE: values inside are for illustration purposes only.

-   Columns A to J (i.e. DocNo to UDF_RCV_MArrivalDate) denote one container. Within one container, there are many items, and each of these items pertain to an ItemCode (see column L header title).
-   Columns K to Y (i.e. Numbering to UDF_RCVDtl_STPrice) denote the details per item. 
-   Using the values in the file as reference, Doc1 under DocNo pertains to a container. The container no lies under column Location, which in this case is BM-OREU1854920. Notice how the container no is the same until the next container.
-   Column "Description" shows "OPENING BALANCE" for the two containers. For now, you may also write this under column "Description" for every container.
-   Column 'UDF_RCV_MArrivalDate' shows the actual arrival date of this container.
-   Column ItemCode. The output either after you have used **search_item_code/** , given the input column "DetailDescription".
-   Column DetailDescription. Gives you a clue on what the corresponding ItemCode should be.
-   Column Location. The container no.
-   Column UOM. For this use case, we can assign all UOM as CTN.
-   Column UDF_RCVDtl_DnChar. Duty & Charges. 
-   Column UDF_RCVDtl_Remarks. Remarks
-   Column UDF_RCVDtl_STPrice. This is the Price.

**P.S.** lb to grams conversion for packing items.
E.g. 10x1lb -> 10x454G. this is helpful when finding matching item codes in **imported_item_codes.xlsx**.

## Critical Rules
-   Output the final template into another xlsx file with the format file name: import_stock_receive_(date in MMM-dd)_(time in T.TTam/pm).xlsx. e.g. import_stock_receive_Apr-04_747am.xlsx
-   Add a column called Matched_Desc that corresponds to the item code's description matched from **imported_item_codes.xlsx**. Place Matched_Desc in the column **immediately to the right of ItemCode** (i.e. ItemCode = column L, Matched_Desc = column M).
-   DocNo column: put `<<New>>` in the DocNo cell of the **first row of each container**, and leave it **blank** on subsequent rows of that container. Each container = one new document; Autocount starts a new document when it sees `<<New>>`. Do **not** use `Doc1`/`Doc2`/etc. (those are only illustrative in **Import_stock_receive_template.xlsx**).
