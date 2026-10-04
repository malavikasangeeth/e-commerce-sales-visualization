# E-Commerce Sales Analysis Dashboard

## 📊 Project Overview

This project analyzes an e-commerce sales dataset using **Microsoft
Excel** to demonstrate data cleaning, transformation, lookup-based data
integration, descriptive statistics, Pivot Tables, Pivot Charts, and
interactive dashboard development.

The workbook contains customer, product, store, and sales transaction
data. The data was cleaned and consolidated into a final analytical
table, which was then used to build an interactive sales dashboard and
perform statistical analysis.

## 🎯 Objectives

-   Organize raw e-commerce data into structured tables.
-   Clean missing, inconsistent, and incorrectly formatted data.
-   Combine customer, product, store, and sales information into a
    single analytical table.
-   Apply Excel formulas for data transformation and calculations.
-   Use `XLOOKUP` to integrate data from dimension tables.
-   Perform descriptive statistical analysis.
-   Create Pivot Tables and Pivot Charts.
-   Build an interactive sales dashboard for business analysis.

## 📁 Workbook Structure

  -----------------------------------------------------------------------
  Sheet                               Description
  ----------------------------------- -----------------------------------
  `Customer_Dim`                      Customer master data including
                                      customer ID, name, demographics,
                                      location, and loyalty level.

  `Product_Dim`                       Product master data including
                                      product name, category, brand,
                                      cost, and stock.

  `Store_Dim`                         Store information including store
                                      name, region, city, and store type.

  `Sales_Fact`                        Transaction-level sales data
                                      including dates, products,
                                      customers, quantities, prices,
                                      discounts, payment types, and sales
                                      amount.

  `Final Table`                       Consolidated analytical dataset
                                      created by combining the dimension
                                      and fact tables.

  `Dashboard`                         Interactive sales dashboard built
                                      using Pivot Tables/Pivot Charts.

  `Solution`                          Documentation of the data
                                      preparation, cleaning,
                                      transformation, and analysis steps.
  -----------------------------------------------------------------------

## 📈 Dataset Summary

The workbook contains:

-   **500 customers**
-   **100 products**
-   **20 stores**
-   **2,000 sales transactions**

The sales data covers transactions with fields such as:

-   Order Date
-   Customer
-   Product
-   Store
-   Quantity
-   Unit Price
-   Discount
-   Payment Type
-   Total Amount

## 🧹 Data Cleaning

The project includes several data-cleaning and preparation activities:

-   Removed/handled duplicate and inconsistent records.
-   Cleaned customer names using `TRIM`.
-   Filled missing loyalty levels with a default value of `Bronze`.
-   Standardized country values, including replacing
    `United States of America` with `USA`.
-   Standardized Customer IDs from formats such as `C-<id>` to
    `CUST<id>`.
-   Cleaned Product IDs.
-   Filled missing product costs using the average cost for the relevant
    category.
-   Filled missing stock values with `0`.
-   Filled missing quantities with `1`.
-   Filled missing unit prices using the average unit price for the
    corresponding product.
-   Filled missing discounts using the average discount for the
    corresponding product.
-   Removed irrelevant data such as the `Sub_Category` column where
    appropriate.

## 🔄 Data Transformation & Integration

The dimension tables were combined with the sales fact table to create a
consolidated `Final Table`.

`XLOOKUP` was used to retrieve customer, product, and related
attributes.

Example:

``` excel
=XLOOKUP([@[Customer_ID]],Customer[Customer_ID],Customer[Name])
```

The final table combines information such as:

-   Customer Name
-   Age
-   Gender
-   City
-   State
-   Country
-   Loyalty Level
-   Product Name
-   Category
-   Brand
-   Cost
-   Stock
-   Store Name
-   Region
-   Store Type
-   Quantity
-   Unit Price
-   Discount
-   Payment Type
-   Total Amount

## 📊 Dashboard

The `Dashboard` sheet provides an interactive view of sales performance
using **Pivot Tables and Pivot Charts**.

The dashboard can be used to explore sales data and identify patterns
across dimensions such as:

-   Product/category
-   Store/region
-   Customer characteristics
-   Loyalty level
-   Payment type
-   Sales performance

## 📐 Statistical Analysis

The project also demonstrates descriptive statistics using Excel
formulas, including:

-   Mean
-   Median
-   Mode
-   Standard deviation
-   Other summary statistics where applicable

These statistics provide additional insight into the distribution and
characteristics of the sales data.

## 🛠️ Tools & Techniques

**Tool:** - Microsoft Excel

**Excel features used:** - Excel Tables - Data Cleaning - Find &
Replace - Go To Special - `TRIM` - `XLOOKUP` - `AVERAGEIF` - Pivot
Tables - Pivot Charts - Interactive Dashboard - Descriptive Statistics -
Data Validation and formatting

## 🔍 Key Skills Demonstrated

This project demonstrates practical skills in:

-   Data cleaning and preprocessing
-   Data transformation
-   Relational data integration
-   Lookup functions
-   Business data analysis
-   Statistical analysis
-   Pivot-based analysis
-   Dashboard development
-   Data visualization
-   Excel-based reporting

## 🚀 How to Use

1.  Download `Ecommerce_Sales_Dataset_updated.xlsx`.
2.  Open the workbook in Microsoft Excel.
3.  Review the dimension tables (`Customer_Dim`, `Product_Dim`, and
    `Store_Dim`).
4.  Review the transaction-level data in `Sales_Fact`.
5.  Use `Final Table` as the consolidated analytical dataset.
6.  Open the `Dashboard` sheet to explore the visual analysis.
7.  Refer to the `Solution` sheet for details of the data preparation
    and analysis process.

## 📌 Project Outcome

The project transforms raw e-commerce transaction and master data into a
**clean, consolidated, and analysis-ready dataset**, followed by an
interactive dashboard and descriptive statistical analysis.

It demonstrates an end-to-end Excel data analytics workflow:

``` text
Raw Data
   ↓
Data Cleaning
   ↓
Data Standardization
   ↓
Missing-Value Treatment
   ↓
XLOOKUP / Data Integration
   ↓
Final Analytical Table
   ↓
Pivot Tables & Pivot Charts
   ↓
Interactive Dashboard
   ↓
Business Insights
```

## 👤 Author

**Sangeeth KS**

Senior Software Engineer \| Frontend / Full-Stack Developer
