# Excel Data Cleaning & Transformation

This project documents the step-by-step process of cleaning, standardizing, and transforming a dataset using Microsoft Excel. 
The goal was to improve data quality by handling missing values, correcting inconsistencies, removing duplicates, and formatting data for analysis.

---

📌 Step 1: Identify Missing Values (Blank Cells)

The **Filter** feature in Excel was used to locate missing values in the **Price** column.

- Applied filter on the Price column  
- Selected **(Blanks)** to display missing values  
- Identified records requiring imputation  

---

📌 Step 2: Replace Missing Price Values

Missing values were replaced using a **category-based median imputation formula**.

### Formula Used:
```excel
=IF(D2<>"",D2,MEDIAN(IF($F$2:$F$33=F2,$E$2:$E$33)))
```
Logic:
If Price is not blank → keep original value
If blank → replace using median price of same category

Why Median?
-> More robust than average
-> Less affected by outliers

Categories Used:
Electronics
Fashion
Kitchen
Outdoor
Accessories

---

📌 Step 3: Replace Missing Categories

Created a lookup table with:
-> Product Name
-> Category

Used Product Name as a matching key
Filled missing category values using reference mapping

---

📌 Step 4: Correct Inconsistent Data

The dataset contained spelling and formatting inconsistencies in Category and Product Name fields.

Fixes Applied:
Used PROPER() function for standardizing text
Used Find & Replace to fix category variations
Ex:
|Electroni|	
|Electrioni|
|Electronic|
|electronics|
|Electronics|

---

📌 Step 5: Remove Duplicates

Duplicate records were identified as rows with identical values across all columns.

Steps:
-> Selected full dataset
-> Navigated to Data → Remove Duplicates
-> Selected all columns
-> Clicked OK

---

📌 Step 6: Split and Merge Data

The Product ID column contained embedded information (date + country code).

Example:
28-JAN-US

Split Into:
-> Manufacturing Date → 28-JAN
-> Country Code → US

Formulas Used:
=LEFT(A2, 6)
=RIGHT(A2, 2)

Merge Data (Improved Readability)
Combined product-related fields:
=CONCATENATE(B2," ",A2)

---

📌 Step 7: Formatting & Conditional Formatting (optional)

Data Formatting:
Price → Currency format
Manufacturing Date → DD-MM-YYYY format

Conditional Formatting:
Applied color scales / data bars to Price column
Created rules to highlight categories visually

---

📊 Summary

This project improved dataset quality by:

-> Handling missing values
-> Standardizing inconsistent text
-> Removing duplicates
-> Structuring embedded data
-> Enhancing readability and analysis readiness

---

🚀 Tools Used
Microsoft Excel
Excel Functions: IF, MEDIAN, LEFT, RIGHT, CONCATENATE, PROPER
Data Cleaning Tools: Filter, Find & Replace, Remove Duplicates
Conditional Formatting
