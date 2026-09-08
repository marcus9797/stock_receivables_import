---
name: get-full-container-no
description: Use when you need to update container no to its full version
---
# Purpose
In stock_list.xlsx files (e.g. BM_stock_list.xlsx, GM_stock_list.xlsx, USJ_stock_list.xlsx), container numbers are mostly displayed in their last 4 digits. But to prepare import_stock_receive____.xlsx files, we'll need their full container numbers.

# Workflow

## Inputs

- import_stock_receive_______.xlsx
- ___stock_list.xlsx

## Output

- import_stock_receive_______.xlsx with updated full respective container numbers.

**ETA_containers.xlsx** give you the full container number. 

## How to find the respective full container number?

1. For each container number inside any of the stock_list.xlsx files, get the last 4 digits.

   E.g. GL 1903 -> 1903

2. For the container number noted in step 1, take note of the product description.

   E.g. SA Red Fuji Apples 18kg 175s

3. The product description from step 2 is used to verify that it is the correct container for this particular item. Look through all the sheets inside **ETA_containers.xlsx** to make sure the previous two info match. Only when both conditions match do we assign the full container no from ETA_containers.xlsx into the output **import_stock_receive_______.xlsx**