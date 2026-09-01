# Real Estate Property Data Analysis Using Excel

## Overview

**Real Estate Property Data Analysis Using Excel** is a university
data-analysis project that uses Microsoft Excel to clean, analyze,
summarize, and visualize real estate property data.

The project starts with raw property records, performs data-quality
inspection and cleaning, creates calculated fields, applies Excel
formulas, builds PivotTables and dashboards, and finishes with findings
and recommendations.

## Objectives

-   Inspect the quality of the raw real estate dataset
-   Identify duplicate records and missing values
-   Remove duplicate records
-   Clean and standardize text and category values
-   Create calculated fields for further analysis
-   Analyze property prices, sizes, bedrooms, types, locations, and
    listing years
-   Build PivotTables for summarized analysis
-   Create two dashboards
-   Present key findings and recommendations

## Dataset

The original dataset contains **300 property records** and **7 main
columns**:

  Column            Description
  ----------------- -----------------------------------
  `Property_ID`     Unique property identifier
  `Listing_Date`    Date when the property was listed
  `Property_Type`   Type/category of the property
  `Location`        Property location
  `Size_Sqm`        Property size in square meters
  `Bedrooms`        Number of bedrooms
  `Price`           Property price

After data cleaning, **292 records** remained for analysis.

## Data Quality and Cleaning

The raw dataset was inspected for:

-   Duplicate records
-   Missing values
-   Blank rows
-   Inconsistent property types
-   Inconsistent location values
-   Text-format inconsistencies
-   Unknown values

### Cleaning Results

  Quality Check         Before Cleaning   After Cleaning
  ------------------- ----------------- ----------------
  Total Records                     300              292
  Duplicate Records                   8                0
  Blank Rows                          0                0
  Missing Values                    111                0

The cleaning process completed duplicate removal, missing-value
handling, inconsistent-value correction, and text cleaning.

## Data Preparation

The `Clean_Data` sheet contains the cleaned dataset together with
additional calculated columns:

-   `Price_Category`
-   `Days Since Listing`
-   `ID_Extracted`
-   `Location Length`
-   `Standardized`
-   `Standardized2`
-   `Year _Colomn`

Examples of formulas used in the workbook include:

``` excel
=IFS(G2>=200000,"High",G2>=100000,"Medium",G2<100000,"Low")
```

``` excel
=IF(OR(B2="Unknown",B2=""),"Unknown",TODAY()-B2)
```

``` excel
=RIGHT(A2,4)
```

``` excel
=LEN(D2)
```

``` excel
=PROPER(C2)
```

``` excel
=PROPER(D2)
```

``` excel
=IF(B2="Unknown","Unknown",YEAR(B2))
```

## Analysis

The `Analysis` sheet contains the main data-quality summary and
descriptive statistics.

### Main Results

  Metric                                       Result
  --------------------------------- -----------------
  Total Properties                                292
  Total Property Price                \$52,230,254.08
  Average Price                          \$178,870.73
  Minimum Price                           \$17,505.31
  Maximum Price                        \$2,000,022.00
  Average Size                             294.53 sqm
  Average Bedrooms                               2.63
  Price Range                          \$1,982,516.69
  Properties in 2025                              169
  Properties in 2026                              107
  Most Common Property Type                 Apartment
  Most Common Property Type Count                  96
  Top Location                              Mogadishu
  Top Location Count                              101
  High Price Properties                            85

### Excel Functions Used

The project uses several Excel functions for data analysis, including:

-   `COUNTA`
-   `COUNTBLANK`
-   `COUNTIF`
-   `COUNTIFS`
-   `SUM`
-   `AVERAGE`
-   `MIN`
-   `MAX`
-   `INDEX`
-   `MODE`
-   `MATCH`
-   `DATE`
-   `YEAR`
-   `TODAY`
-   `IFS`
-   `IF`
-   `OR`
-   `LEN`
-   `PROPER`
-   `RIGHT`

### Example Analysis Formulas

**Total Properties**

``` excel
=COUNTA(Clean_Data!A2:A1000)
```

**Total Property Price**

``` excel
=SUM(PropertyTable[Price])
```

**Average Price**

``` excel
=AVERAGE(Clean_Data!G2:G293)
```

**Minimum Price**

``` excel
=MIN(Clean_Data!G2:G293)
```

**Maximum Price**

``` excel
=MAX(Clean_Data!G2:G293)
```

**Missing Values**

``` excel
=COUNTBLANK(Raw_Data!A2:A301)
```

**Properties Listed in 2025**

``` excel
=COUNTIFS(PropertyTable[Listing_Date],">="&DATE(2025,1,1),PropertyTable[Listing_Date],"<"&DATE(2026,1,1))
```

**Properties Listed in 2026**

``` excel
=COUNTIFS(PropertyTable[Listing_Date],">="&DATE(2026,1,1),PropertyTable[Listing_Date],"<"&DATE(2027,1,1))
```

**Most Common Property Type**

``` excel
=INDEX(Clean_Data!C2:C293,MODE(MATCH(Clean_Data!C2:C293,Clean_Data!C2:C293,0)))
```

**Top Location**

``` excel
=INDEX(Clean_Data!D2:D293,MODE(MATCH(Clean_Data!D2:D293,Clean_Data!D2:D293,0)))
```

## PivotTable Analysis

The `PivotTables` sheet contains PivotTable-based summaries used to
explore the cleaned data.

The analysis covers areas such as:

-   Properties by location
-   Property type distribution
-   Average price by location
-   Total property value by type
-   Total property size by location
-   Other summarized property metrics

PivotTables were used to make comparisons easier and support the
dashboard visualizations.

## Dashboards

The workbook contains two dashboard sheets.

### Dashboard 1: Overview

`Dashbroad_1` provides a high-level summary of the dataset.

Main KPI values include:

-   Total Properties: **292**
-   Average Price: **\$178,870.73**
-   Average Size: **294.53 sqm**
-   Highest Price: **\$2,000,022**

### Dashboard 2: Analysis

`Dashbroad_2` provides a deeper view of the dataset.

Main KPI values include:

-   Total Properties: **292**
-   Total Property Price: **\$52,230,254.08**
-   Total Bedrooms: **767**
-   Average Bedrooms: **2.63**

The dashboards use Excel charts, PivotTable-based analysis, and
interactive filtering elements to present the results.

## Key Findings

1.  **Mogadishu has the highest number of properties**, with **101
    listings**, making it the largest location in the dataset.
2.  **Apartments are the most common property type**, with **96
    properties**, followed by Houses with 59.
3.  **Mogadishu has the highest average property price**, at
    approximately **\$209,588**, followed by Garowe at approximately
    **\$206,754**.
4.  The data-quality process improved the dataset by removing **8
    duplicate records** and handling **111 missing values**, leaving
    **292 clean records** for analysis.
5.  The dataset contains more listings from **2025 (169)** than **2026
    (107)**.

## Recommendations

-   Focus further market analysis on **Mogadishu**, which has the
    highest number of listings and a high average property price.
-   Monitor the **apartment market**, since apartments are the largest
    property-type group.
-   Compare property prices between locations, especially **Mogadishu
    and Garowe**, to understand differences in market value.
-   Continue regular data cleaning to prevent duplicate, missing, and
    inconsistent records.
-   Keep listing dates and property information updated so future
    analysis remains reliable.

## Project Workflow

``` text
Raw Data
    ↓
Data Quality Inspection
    ↓
Duplicate & Missing Value Analysis
    ↓
Data Cleaning
    ↓
Text & Category Standardization
    ↓
Calculated Fields
    ↓
Descriptive Analysis
    ↓
PivotTables
    ↓
Dashboards
    ↓
Findings & Recommendations
```

## Workbook Structure

``` text
Real Estate Property Data Analysis Using Excel.xlsx
│
├── Raw_Data
├── Clean_Data
├── Analysis
├── PivotTables
├── Dashbroad_1
├── Dashbroad_2
└── Findings
```

### Sheet Descriptions

  Sheet           Purpose
  --------------- -----------------------------------------------
  `Raw_Data`      Original real estate dataset
  `Clean_Data`    Cleaned dataset and calculated fields
  `Analysis`      Data-quality summary and descriptive analysis
  `PivotTables`   PivotTable-based summaries
  `Dashbroad_1`   Overview dashboard
  `Dashbroad_2`   Detailed analysis dashboard
  `Findings`      Key findings and recommendations

## Tools Used

-   Microsoft Excel
-   Excel Tables
-   Excel Formulas
-   Sorting and Filtering
-   Conditional Formatting
-   Data Cleaning
-   PivotTables
-   PivotCharts
-   Dashboard Design
-   Slicers

## Repository Structure

A recommended GitHub repository structure is:

``` text
Real-Estate-Property-Data-Analysis/
│
├── Dataset/
│   └── Real Estate Property Data Analysis Using Excel.xlsx
│
├── PPT_file/
│   ├── Presentation.pptx
│   └── Report.pdf
│
├── screenshots/
│   └── README.md
│
└── README.md
```

## How to Use

1.  Download the Excel workbook from the repository.
2.  Open the workbook in Microsoft Excel.
3.  Start with `Raw_Data` to review the original records.
4.  Open `Clean_Data` to review the cleaned dataset.
5.  Review `Analysis` for data-quality results and calculations.
6.  Review `PivotTables` for summarized analysis.
7.  Open `Dashbroad_1` for the overview dashboard.
8.  Open `Dashbroad_2` for the detailed dashboard.
9.  Open `Findings` to review the final conclusions and recommendations.

## Conclusion

This project demonstrates a complete real estate data-analysis workflow
using Microsoft Excel.

The dataset was inspected, cleaned, analyzed, and visualized. After
removing duplicate records and handling missing and inconsistent values,
**292 property records** were used for the final analysis.

The analysis provides useful insights into property prices, locations,
property types, sizes, bedrooms, and listing years, while the dashboards
make the results easier to understand and communicate.

## Author

**Real Estate Property Data Analysis Using Excel**

University Project
