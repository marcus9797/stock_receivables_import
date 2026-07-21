---
name: create-item-code
version: 1.0.0
category: data-processing
description: Use when you need to add new item codes and rows accordingly to ./imported_item_codes.xlsx.
disable-model-invocation: true
allowed-tools: Read, Write, Bash, Glob
---
## Project
Sometimes, no existing item code can be found using skill **find-item-code**. So we'll need to create new item codes accordingly to eventually add into **./imported_item_codes.xlsx**.

## Workflow
1. you use file **./to-add-items.txt** as input. There will be comments inside marked with # containing additional info. For every line, you'll need to parse the description info.

E.g. Parse the description for the following Autocount fields:
   - ItemType → Fruit
   - ItemBrand → Brand
   - ItemClass → Country
   - ItemCategory → Variety
   - UDF_Grade → Grade
   - UDF_UOM_Size → Fruit size (ROW, MM, etc.)
   - UDF_UOM_WeightPerCtn → Carton weight (kg)
   - UDF_Originally_Prepacked -> Packing (e.g. "16x800G", "12x1KG", leave blank if not specified)

   E.g.
   If description = AUST SKIST WASHINGTON NAVEL ORANGE 18KG 88,

   ItemType = Oranges
   ItemBrand = Sunkist (from **./common_brand_abbreviations.json**), but do also check **./unique_brands.json**.
   ItemClass = Australia
   ItemCategory = Navel
   UDF_Grade = Nil (none found from **./grades.txt**)
   UDF_UOM_Size = 88
   UDF_UOM_WeightPerCtn = 18000

2. ensure that you adhere to the table structure of **./imported_item_codes.xlsx**. 

These include, but are not limited to:
- For a given item code there must at least be one row with UDF_UOM_SizeRate=Rate=1, and its UOM=BaseUOM. 
- Row order and grouping are vital. For a given item code, a row with Rate=1 is at the top, the row with maximum Rate, and thus UOM=CTN, is at the bottom. All rows with the same item code are contiguous.
- Column UDF_UOM_ADJdedcut must be F for Rate=1 rows and UOM=CTN rows, else it'll be T
- Column Classification must be '022 for ALL rows. NOT 22 or 022. The ' at the front is important.
- To add an item code, you first find the new item's ItemType, ItemCategory and ItemClass. That will give you the first 3 short code combinations e.g. APL-GLA-TUR-xxxx. xxxx is the numbering. For a combination, make sure to increment by +1 from the highest number. E.g. to add APL-GLA-TUR-0099 if highest number at the end for this combination was APL-GLA-TUR-0098.
3. output the new item rows to be added to **./to_add_item_codes.xlsx**.
4. if there are any new brands not already in **./brands.yaml**, add them into the yaml file.