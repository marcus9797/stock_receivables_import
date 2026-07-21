# Autocount SKU creation

## Purpose (Why)
We need to import existing in-stock items into Autocount Malaysia (ERP software). But to do so, a template must first be followed

## What
**imported_item_codes.xlsx**: These show the currently imported item codes within Autocount. Use this to find corresponding field values for Variety, Country, Fruit etc.

### In the following key-value pairs, you'll find the respective field columns inside the xlsx for the relevant info. 

Key = our language, Value = respective column inside the xlsx.

Variety: ItemCategory
Country of Origin: ItemClass
Fruit: ItemType
Brand: ItemBrand
Size: UDF_UOM_Size
Packing (e.g. 16x800G): UDF_Originally_Prepacked
Grade: UDF_Grade
Carton Weight (e.g. 17kg): UDF_UOM_WeightPerCtn

**item_code_hashmap.json**: Knowledge base that stores our countries, fruits and varieties. The important parts here are the Keys, not values. Use this when looking at a description and deciphering which word is a country/fruit/variety. As for what a description means, refer to **BM_stock_list.xlsx**.

**Import_stock_receive_template.xlsx**: This is the template to adhere to before finally submitting to Autocount. NOTE: values inside are for illustration purposes only.

-   Columns A to J (i.e. DocNo to UDF_RCV_MArrivalDate) denote one container. Within one container, there are many items, and each of these items pertain to an ItemCode (see column L header title).
-   Columns K to Y (i.e. Numbering to UDF_RCVDtl_STPrice) denote the details per item. 
-   Using the values in the file as reference, Doc1 under DocNo pertains to a container. The container no lies under column Location, which in this case is BM-OREU1854920. Notice how the container no is the same until the next container.
-   Column "Description" shows "OPENING BALANCE" for the two containers. For now, you may also write this under column "Description" for every container.
-   Column 'UDF_RCV_MArrivalDate' shows the actual arrival date of this container.
-   Column ItemCode. The output after using skill **find-item-code** , given the input column "DetailDescription".
-   Column DetailDescription. Gives you a clue on what the corresponding ItemCode should be. See skill **find-item-code**.
-   Column Location. The container no.
-   Column UOM. For this use case, we can assign all UOM as CTN.
-   Column UDF_RCVDtl_DnChar. Duty & Charges. 
-   Column UDF_RCVDtl_Remarks. Remarks
-   Column UDF_RCVDtl_STPrice. This is the Price.

**BM/BM_stock_list.xlsx**: This is the stock list in which we have to pull out the info from and translate into the template in **Import_stock_receive_template.xlsx**.
-   Look at row 4, where it says FUJI APPLE under column A, SIZE under column B, QTY under column C etc. This is what I call variety headers. They mark the start of a certain fruit variety moving forward until the next variety header comes around. 
-   The values under each column give you a clue on what to map to within **Import_stock_receive_template.xlsx**:

    - ETA -> UDF_RCVDtl_ETA
    - D & CHAR -> UDF_RCVDtl_DnChar
    - PRICE -> UDF_RCVDtl_STPrice
    - ARV BM -> UDF_RCV_MArrivalDate
    - REMARK -> UDF_RCVDtl_Remarks
    - CONT -> Location
    - QTY -> Qty
    - SIZE, this needs to be added into the DetailDescription inside **Import_stock_receive_template.xlsx**.
    - Except the variety headers, Column A is where you get most of the description to be added into DetailDescription, you just need to include the SIZE part before filling into DetailDescription.
    - With the previous two points covered and compiled into a single description, you can then use DetailDescription to work out the ItemCode, from **imported_item_codes.xlsx** as explained before.
    - You may ignore filling in the other columns except for Description in **Import_stock_receive_template.xlsx**.

**brands.yaml**: This contains the brands for our items. 
-   They correspond to ItemBrand in **imported_item_codes.xlsx**.
-   Refer to this when scanning descriptions as they contain brands.
-   When checking to see if a brand exists, make sure it's case-insensitive.

**grades.txt**: Contains known grades of items. When you see any one of these when looking at descriptions in the stock list, know that they should be under column UDF_Grade.

**common_brand_abbreviations.json**: You may bump into abbreviated brands in stock list item descriptions. Refer to this to ensure you don't miss any brands. When you get the description from the input stock list xlsx, make sure to check for any brand abbreviations in this json. only then will you know the true brand to match with in imported_item_codes.xlsx.

## Critical Rules
-   Output the final template into another xlsx file with the format file name: import_stock_receive_<todays_date>_<time>.xlsx
Format the output xlsx like this: import_stock_receive_Jul19_230pm.xlsx, for e.g.
-   In any files ending with **_stock_list.xlsx**, if you come across any unknown fields in the description (e.g. unknown ItemBrand) that can't be broken down into fields inside **imported_item_codes.xlsx**, flag it.