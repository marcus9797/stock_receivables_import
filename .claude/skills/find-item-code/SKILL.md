---
name: find-item-code
version: 1.0.0
category: data-processing
description: Use when you need to find the respective item code in ./imported_item_codes.xlsx, given a description like SA SY NAVEL ORANGE 15KG 60.
---
## Purpose
Descriptions contain information that you can break down, and in turn use to find the corresponding item code(s) in **./imported_item_codes.xlsx**.

## Workflow
1. Parse the description for the following Autocount fields:
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
2. **Query ./imported_item_codes.xlsx located at the project root (`./imported_item_codes.xlsx`)** using the parsed values against the column headers.

E.g. "CHINA FUJI APPLE 17KG 100s". 
   ItemType = Apples
   ItemBrand = Nil (not listed)
   ItemClass = China
   ItemCategory = Fuji
   UDF_Grade = Nil (none found from **./grades.txt**)
   UDF_UOM_Size = 100
   UDF_UOM_WeightPerCtn = 17000

Using the above parsed info, the item code is APL-FJI-CN-0010.

3. **Return the matching ItemCode**. If multiple matches exist, list all possible codes. If no matches exist, output UNKNOWN as the item code and flag for human review.

## Critical Rules
-   If you meet any words in the description that don't match any brands/countries etc. flag it for manual human review. Don't simply fill it with any item codes.