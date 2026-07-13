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

**item_code_hashmap.json**: Knowledge base that stores our countries, fruits and varieties. The important parts here are the Keys, not values. Use this when looking at a description and deciphering which word is a country/fruit/variety. As for what a description means, refer to **BM_stock_list.xlsx**.

**Import_stock_receive_template.xlsx**: This is the template to adhere to before finally submitting to Autocount. NOTE: values inside are for illustration purposes only.

-   Columns A to J (i.e. DocNo to UDF_RCV_MArrivalDate) denote one container. Within one container, there are many items, and each of these items pertain to an ItemCode (see column L header title).
-   Columns K to Y (i.e. Numbering to UDF_RCVDtl_STPrice) denote the details per item. 
-   Using the values in the file as reference, Doc1 under DocNo pertains to a container. The container no lies under column Location, which in this case is BM-OREU1854920. Notice how the container no is the same until the next container.
-   Column "Description" shows "OPENING BALANCE" for the two containers. For now, you may also write this under column "Description" for every container.
-   Column 'UDF_RCV_MArrivalDate' shows the actual arrival date of this container.
-   Column ItemCode. This is the point where you'll need to refer to file **imported_item_codes.xlsx** to search for the respective ItemCode based on the values in other fields. But how do you determine such values? From column "DetailDescription".
-   Column DetailDescription. Gives you a clue on what the corresponding ItemCode should be. Let's pick apart cell M2. "CHINA FUJI APPLE 17KG 100s". Obviously China is a country, Apple is a Fruit, 17KG is the carton weight. But 100s is the size of the fruit. When you break them down you can use these individual components to filter the appropriate columns in **imported_item_codes.xlsx** to search for the respective ItemCode. Thus, in this example, the correct ItemCode is APL-FJI-CN-0010.
-   Column Location. The container no.
-   Column UOM. For this use case, we can assign all UOM as CTN.
-   Column UDF_RCVDtl_DnChar. Duty & Charges. 
-   Column UDF_RCVDtl_Remarks. Remarks
-   Column UDF_RCVDtl_STPrice. This is the Price.

**BM_stock_list.xlsx**: This is the stock list in which we have to pull out the info from and translate into the template in **Import_stock_receive_template.xlsx**.
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

**unique_brands.yaml**: This contains the brands for our items. 
-   They correspond to ItemBrand in **imported_item_codes.xlsx**.

## Critical Rules
-   Output the final template into another xlsx file with the format file name: import_stock_receive_<todays_date>_<time>.xlsx
-   In **BM_stock_list.xlsx**, if you come across any unknown fields in the description (e.g. unknown ItemBrand) that can't be broken down into fields inside **imported_item_codes.xlsx**, flag it.
