# SWYNEX-Data-Cleaning-Preparation

Project Overview

This project focuses on cleaning and preparing a customer marketing campaign dataset using Microsoft Excel.
The dataset contains information about 2,240 customers, including their demographic details, income, household information, purchase behavior, spending across different product categories, website activity, and responses to previous marketing campaigns.
The main objective of this project was not to perform advanced analysis or build a dashboard. Instead, I focused on an important first step in any Data Analyst workflow:
Taking raw, inconsistent data and preparing it so that it can be used reliably for further analysis.
During the project, I worked on identifying data-quality issues, standardizing inconsistent categories, creating a new Age column, checking duplicate customer IDs, checking missing values, formatting dates, and validating the final dataset.

🎯 Project Objective

The main objectives of this project were:

Understand the structure of a raw customer marketing dataset
Identify data-quality problems
Check for duplicate customer records
Standardize inconsistent categorical values
Handle date formatting issues
Create a useful derived column such as Age
Identify missing values
Check unusual or potentially incorrect values
Prepare the dataset for further analysis
Create a simple data dictionary to understand the columns

The final cleaned dataset contains:

2,240 customer records and 30 columns.

📂 Dataset Information

The original dataset contained 2,240 customer records and 29 columns.

The cleaned dataset contains 2,240 records and 30 columns because I created an additional:

Age

column from the customer's birth year.

Main types of information in the dataset

The dataset contains information related to:

Customer ID
Year of birth
Age
Education
Marital status
Income
Number of children
Number of teenagers in the household
Customer enrollment date
Recency of purchase
Product spending
Purchase channels
Website visits
Previous campaign responses
Complaints
Latest campaign response
🛠️ Tools Used
Microsoft Excel

I used Excel for the complete data-cleaning process.

Main Excel features used:

Import
Filters
Find & Replace
Remove Duplicates
Text to Columns
Date Formatting
Excel formulas
Excel Tables
Data validation/checking
Sorting
Data Dictionary
🔄 Data Cleaning Process
1. Understanding the Raw Dataset

I first opened the raw marketing campaign dataset and checked its overall structure.

The original file contained 29 columns and 2,240 customer records.

Before making any changes, I kept the original data separately so that I could compare the raw data with the cleaned version.

I maintained two versions:

Raw Data
     ↓
Cleaning Process
     ↓
Cleaned Data

This helped me avoid accidentally losing the original information.

2. Fixing the CSV Import Problem

One of the first problems I faced was related to the CSV file itself.

When I initially opened the CSV directly in Excel, the data appeared incorrectly.

Instead of appearing like:

ID | Year_Birth | Education | Marital_Status | Income

the values were appearing together in a single column.

For example:

IDYear_BirthEducationMarital_StatusIncome...
How I fixed it

Instead of manually moving the data, I imported the CSV through:

Data → From Text/CSV

I selected the correct delimiter so that Excel could recognize each field as a separate column.

After importing it correctly, the dataset was structured properly.

What I learned

I learned that a CSV file can sometimes open incorrectly depending on its delimiter and Excel's import settings.

Before cleaning the actual data, it is important to make sure the dataset structure is correct.

3. Checking Duplicate Customer IDs

The ID column represents the unique customer identifier.

Before removing anything, I checked whether the same customer ID appeared more than once.

I used:

Data → Remove Duplicates

and selected only the ID column for checking.

I also used a COUNTIF formula to verify the IDs.

Example:

=COUNTIF($A:$A,A2)

If the result was:

1

the ID appeared only once.

If the result was:

2

the ID appeared twice and would need investigation.

Result

The dataset contained:

2,240 unique IDs

and therefore:

0 duplicate customer IDs were removed.

This was an important lesson for me:

Data cleaning does not always mean deleting data. Sometimes the correct result is to check for a problem and find that there is nothing to remove.

4. Cleaning the Education Column

The Education column contained inconsistent category names.

The original dataset contained values such as:

Graduation
2n Cycle
Master
PhD
Basic

I standardized the categories to:

Graduate
Master
PhD
Basic
Changes made
Graduation → Graduate
2n Cycle → Graduate

I used Excel's:

Find & Replace

feature to make these changes.

Why I did this

If the same type of customer is represented using different names, it can create problems during analysis.

For example:

Graduation
Graduate

could be treated as two separate categories.

Standardizing them makes future charts, PivotTables, and analysis more reliable.

5. Cleaning the Marital Status Column

The Marital_Status column had several inconsistent categories.

The raw dataset contained:

Single
Together
Married
Divorced
Widow
Alone
Absurd
YOLO

I standardized these into:

Single
Married
Divorced
Widowed
Changes made
Original Value	Standardized Value
Together	Married
Widow	Widowed
Alone	Single
Absurd	Single
YOLO	Single

I used Find & Replace in Excel to make these changes.

Why this was necessary

Having too many inconsistent categories can make analysis confusing.

For example, if Together and Married are considered the same group for the purpose of this cleaned dataset, keeping them separately would produce misleading category counts.

6. Creating the Age Column

The original dataset contained:

Year_Birth

but did not contain an Age column.

I created a new:

Age

column.

I calculated Age using:

=2025-Year_Birth

For example:

Year_Birth = 1957

Age = 2025 - 1957

Age = 68

Another example:

Year_Birth = 1984

Age = 2025 - 1984

Age = 41
Why I created Age

Year of birth is useful, but Age is easier to understand when performing customer segmentation.

For example, later I can analyze:

customer age groups
spending by age
campaign response by age
purchasing behavior by age
7. Checking Age Outliers

While checking the Age column, I found some unusual values.

There were 3 customers with ages above 100.

These came from unusually old birth years, including:

1893
1899
1900

For example, a birth year of 1893 results in:

2025 - 1893 = 132

These values are potential data-quality anomalies.

What I did

I did not automatically delete these customers.

Instead, I kept the records and identified them as potential outliers.

Why?

Deleting unusual data without understanding the reason can cause unnecessary data loss.

This taught me that:

An unusual value is not automatically a value that should be deleted.

8. Checking Missing Income Values

I checked the Income column using Excel filters.

I found:

24 missing Income values.

Instead of replacing these values with:

0

or the average income, I kept them blank.

Why?

Replacing missing values without a clear business or statistical reason can introduce assumptions into the dataset.

For this project, I decided to preserve the missing values and document them.

This means the cleaned dataset still contains:

24 missing Income values.

9. Formatting the Customer Date

The Dt_Customer column contains the date when the customer became a customer.

The original values were represented in a day-month-year format.

I converted the values into proper Excel dates and formatted them consistently as:

YYYY-MM-DD

For example:

04-09-2012

became:

2012-09-04
Why this was important

Dates need to be stored consistently so they can later be used for:

sorting
filtering
calculating time periods
trend analysis
customer tenure analysis
10. Checking Numerical Columns

I checked the numerical fields to make sure they contained appropriate numeric values.

Examples include:

Income
Recency
MntWines
MntFruits
MntMeatProducts
MntFishProducts
MntSweetProducts
MntGoldProds
NumWebPurchases
NumStorePurchases
NumCatalogPurchases
NumWebVisitsMonth

I also checked that legitimate 0 values were not accidentally treated as missing data.

For example:

MntWines = 0

doesn't necessarily mean the data is wrong.

It can simply mean that the customer did not spend anything on wine.

11. Checking Campaign Response Columns

I also checked the campaign-related fields such as:

AcceptedCmp1
AcceptedCmp2
AcceptedCmp3
AcceptedCmp4
AcceptedCmp5
Response

These columns use binary values:

0
1

where the values represent different campaign outcomes.

I kept these values in their original numeric format instead of changing them to text such as:

Yes
No

This makes the dataset easier to use for future calculations and analysis.

12. Creating an Excel Table

After cleaning the dataset, I converted the data range into an Excel Table.

This made it easier to:

filter data
sort data
identify columns
navigate through the dataset
maintain consistent formatting

It also made the final dataset easier to work with for future analysis.

13. Creating a Data Dictionary

I also created a separate:

Data_Dictionary

sheet.

The purpose of the data dictionary is to explain what each column represents.

For example:

Column	Description
ID	Unique customer identifier
Year_Birth	Customer's birth year
Age	Customer age calculated from birth year
Education	Standardized education category
Marital_Status	Standardized marital status
Income	Customer income
Kidhome	Number of children in household
Teenhome	Number of teenagers in household
Dt_Customer	Customer enrollment date
Recency	Number of days since last purchase
MntWines	Amount spent on wine
MntFruits	Amount spent on fruits
MntMeatProducts	Amount spent on meat products
MntFishProducts	Amount spent on fish products
MntSweetProducts	Amount spent on sweets
MntGoldProds	Amount spent on gold products
Response	Response to the latest campaign

This makes the dataset easier for another person to understand.

📊 Before vs After
Item	Raw Dataset	Cleaned Dataset
Customer Records	2,240	2,240
Columns	29	30
Duplicate IDs	Checked	0
Age Column	❌	✅
Education Standardized	❌	✅
Marital Status Standardized	❌	✅
Date Format	Inconsistent/raw format	YYYY-MM-DD
Missing Income	24	24 retained
Data Dictionary	❌	✅
🚧 Problems I Faced During the Project

This project was useful because I faced several practical problems rather than simply following a tutorial.

Problem 1: CSV data was appearing in one column

When I first opened the CSV file, all the fields were appearing together instead of being separated into columns.

Solution

I imported the CSV using Excel's From Text/CSV option and selected the correct delimiter.

What I learned

Before cleaning data, I need to make sure the file has been imported correctly.

Problem 2: Different names represented similar categories

The Education column contained:

Graduation
2n Cycle
Graduate

while the cleaned dataset needed consistent categories.

Solution

I used Find & Replace to standardize the values.

What I learned

Categorical data needs consistency before performing analysis.

Problem 3: Marital Status contained unusual categories

Values such as:

Together
Alone
Absurd
YOLO
Widow

made the category structure inconsistent.

Solution

I standardized these categories based on the cleaning rules used for the dataset.

What I learned

Before changing categorical values, I need to understand what each value represents rather than simply deleting unusual values.

Problem 4: Missing Income values

I found 24 customers without Income information.

Solution

I kept these values blank instead of making up values.

What I learned

Missing data should be handled based on the purpose of the analysis. Filling every blank is not always the correct approach.

Problem 5: Unusual customer ages

Some customers had calculated ages above 100.

Solution

I identified these records as potential data-quality issues but did not delete them automatically.

What I learned

Outlier detection and outlier removal are two different things.

First identify the unusual value, then decide whether there is a valid reason to remove or correct it.

Problem 6: Checking duplicates

I initially considered removing duplicate customer IDs.

After checking the data, I found that all 2,240 customer IDs were unique.

What I learned

A cleaning process should be based on actual checks rather than assuming that every dataset contains duplicate records.

🔍 Data Quality Validation

After completing the cleaning process, I performed a final validation.

Final dataset

Rows: 2,240

Columns: 30

Unique Customer IDs: 2,240

Duplicate IDs: 0

Missing Income values: 24

Education categories: 4

Basic
Graduate
Master
PhD

Marital Status categories: 4

Single
Married
Divorced
Widowed

Potential age outliers: 3

📁 Project Structure

The Excel workbook contains:

Customer-Marketing-Campaign/
│
├── marketing_campaign
│   └── Original raw dataset
│
├── Cleaned_Data
│   └── Final cleaned dataset
│
└── Data_Dictionary
    └── Explanation of dataset columns
💡 What I Learned From This Project

This project helped me understand that data cleaning is much more than deleting blanks and duplicates.

I learned how to:

Import CSV data correctly
Understand the structure of a dataset
Check duplicate records
Standardize categorical values
Work with dates in Excel
Create calculated columns
Identify missing values
Identify potential outliers
Preserve data when there is no valid reason to delete it
Use Excel filters for data-quality checks
Create a structured Excel table
Document columns using a data dictionary

Most importantly, I learned that every cleaning decision should have a reason.

For example:

A missing value does not automatically mean the row should be deleted.

and:

An unusual value does not automatically mean the value is incorrect.

🚀 Future Analysis

After cleaning, this dataset can be used for further analysis such as:

Customer Segmentation
Age groups
Income groups
Education
Marital status
Product Analysis
Which products generate the most spending?
Which customer groups spend the most?
Which products are popular among different customer segments?
Marketing Campaign Analysis
Which campaign received the highest response?
Which customer groups respond to campaigns?
Does income affect campaign response?
Does age affect campaign response?
Purchasing Behavior
Web purchases
Store purchases
Catalog purchases
Website visits
Deal purchases
Customer Behavior
Recency
Spending
Household characteristics
Campaign response

The cleaned dataset can later be used with Excel, SQL, Power BI, Tableau, or Python for deeper analysis.

📌 Key Takeaway

This project gave me practical experience with one of the most important parts of a Data Analyst's workflow:

Raw Data → Data Cleaning → Data Validation → Analysis-ready Data

Instead of directly building charts from raw data, I first focused on making sure the underlying dataset was structured, consistent, and understandable.

👨‍💻 Skills Demonstrated

Microsoft Excel Data Cleaning Data Preparation Data Validation Data Quality Excel Formulas Find & Replace Filters Duplicate Checking Date Formatting Data Standardization Data Documentation
