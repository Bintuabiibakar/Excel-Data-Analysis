# Real Estate Property Data Analysis Using Excel

## Project Overview

This project analyzes real estate property data using Microsoft Excel.
The project focuses on data quality inspection, data cleaning,
descriptive analysis, PivotTables, dashboards, and findings.

The main goal is to transform raw property data into clean, reliable,
and useful information that can support real estate market analysis.

## Objectives

-   Inspect the quality of the raw dataset
-   Identify duplicate records and missing values
-   Clean and standardize the dataset
-   Create calculated fields for analysis
-   Analyze property prices, sizes, bedrooms, property types, locations,
    and listing dates
-   Build PivotTables and dashboard summaries
-   Identify important findings and provide recommendations

## Dataset

The dataset contains real estate property information with the following
main columns:

  Column          Description
  --------------- -------------------------------------
  Property_ID     Unique identifier for each property
  Listing_Date    Date the property was listed
  Property_Type   Type of property
  Location        Property location
  Size_Sqm        Property size in square meters
  Bedrooms        Number of bedrooms
  Price           Property price

### Dataset Size

-   **Raw records:** 300
-   **Duplicate records:** 8
-   **Clean records:** 292
-   **Blank rows:** 0
-   **Missing values before cleaning:** 111
-   **Missing values after cleaning:** 0

## Data Cleaning Process

The project follows a structured data-cleaning process.

### 1. Data Quality Inspection

The raw dataset was inspected to identify:

-   Duplicate records
-   Missing values
-   Blank rows
-   Inconsistent property types
-   Inconsistent location values
-   Extra spaces
-   Text-format inconsistencies
-   Unknown or invalid values

### 2. Duplicate Removal

Duplicate records were identified using the `Property_ID` field.

A total of **8 duplicate records** were removed from the raw dataset.

### 3. Missing Value Handling

Missing values were checked across the important dataset columns using
Excel functions such as `COUNTBLANK`.

A total of **111 missing values** were identified before cleaning. After
the cleaning process, the cleaned dataset contains **0 missing values**
according to the workbook's data-quality summary.

### 4. Text Cleaning and Standardization

Text fields were reviewed and standardized to improve consistency.

The project also used text-related Excel functions for checking and
extracting information, including:

-   `LEN`
-   `UPPER`
-   `RIGHT`

### 5. Calculated Fields

Additional fields were created in the `Clean_Data` sheet to support
analysis:

-   `Price_Category`
-   `Days Since Listing`
-   `ID_Extracted`
-   `Location Length`
-   `Standardized`
-   `Year`

## Excel Analysis

The `Analysis` sheet contains the main descriptive statistics and
data-quality results.

The project uses Excel functions including:

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

### Example Formulas

#### Total Properties

``` excel
=COUNTA(Clean_Data!A2:A1000)
```

#### Average Price

``` excel
=AVERAGE(Clean_Data!G2:G293)
```

#### Minimum Price

``` excel
=MIN(Clean_Data!G2:G293)
```

#### Maximum Price

``` excel
=MAX(Clean_Data!G2:G293)
```

#### Missing Values

``` excel
=COUNTBLANK(Raw_Data!A2:A301)
```

#### Properties Listed in 2025

``` excel
=COUNTIFS(PropertyTable[Listing_Date],">="&DATE(2025,1,1),PropertyTable[Listing_Date],"<"&DATE(2026,1,1))
```

#### Properties Listed in 2026

``` excel
=COUNTIFS(PropertyTable[Listing_Date],">="&DATE(2026,1,1),PropertyTable[Listing_Date],"<"&DATE(2027,1,1))
```

#### Most Common Property Type

``` excel
=INDEX(Clean_Data!C2:C293,MODE(MATCH(Clean_Data!C2:C293,Clean_Data!C2:C293,0)))
```

#### Top Location

``` excel
=INDEX(Clean_Data!D2:D293,MODE(MATCH(Clean_Data!D2:D293,Clean_Data!D2:D293,0)))
```

## Key Analysis Results

  Metric                                       Result
  --------------------------------- -----------------
  Total Properties                                292
  Total Property Value                \$52,230,254.08
  Average Price                          \$178,870.73
  Minimum Price                           \$17,505.31
  Maximum Price                        \$2,000,022.00
  Average Size                             294.53 sqm
  Average Bedrooms                               2.63
  Price Range                          \$1,982,516.69
  Properties Listed in 2025                       169
  Properties Listed in 2026                       107
  Most Common Property Type                 Apartment
  Most Common Property Type Count                  96
  Top Location                              Mogadishu
  Top Location Count                              101

## PivotTable Analysis

The `PivotTables` sheet is used to summarize the cleaned real estate
data.

The analysis includes comparisons such as:

-   Properties by location
-   Property size by location
-   Property prices by location
-   Property type distribution
-   Bedrooms by location
-   Property value and property type

PivotTables make it easier to compare different parts of the real estate
dataset and identify patterns.

## Dashboards

The workbook contains two dashboards:

### Dashboard 1: Overview

`Dashbroad_1` provides an overview of the real estate dataset and its
main metrics.

### Dashboard 2: Analysis

`Dashbroad_2` provides additional analysis of property characteristics
and market information.

The dashboards use Excel charts and summarized data to make the results
easier to understand.

## Key Findings

1.  **Mogadishu has the highest number of properties**, with 101
    listings, making it the largest market in the dataset.
2.  **Apartments are the most common property type**, with 96
    properties, followed by Houses with 59.
3.  **Mogadishu has the highest average property price**, at
    approximately \$209,588, followed by Garowe at approximately
    \$206,754.
4.  **Data quality improved after cleaning.** Eight duplicate records
    were removed and missing values were handled, leaving 292 clean
    records.

## Recommendations

1.  Focus more analysis on Mogadishu because it has the largest number
    of properties and strong average prices.
2.  Monitor the apartment market because apartments represent the
    largest property-type group.
3.  Compare pricing by location, especially Mogadishu and Garowe, to
    understand differences in average property values.
4.  Keep the dataset regularly updated and cleaned to reduce duplicate,
    missing, and inconsistent information.

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

  Sheet         Purpose
  ------------- -----------------------------------------------
  Raw_Data      Original dataset before cleaning
  Clean_Data    Cleaned dataset and calculated fields
  Analysis      Data-quality summary and descriptive analysis
  PivotTables   PivotTable-based analysis
  Dashbroad_1   Overview dashboard
  Dashbroad_2   Detailed analysis dashboard
  Findings      Key findings and recommendations

## Project Workflow

``` text
Raw Data
    ↓
Data Quality Inspection
    ↓
Duplicate and Missing Value Analysis
    ↓
Data Cleaning
    ↓
Text and Category Standardization
    ↓
Calculated Fields
    ↓
Descriptive Analysis
    ↓
PivotTables
    ↓
Dashboards
    ↓
Findings and Recommendations
```

## Tools Used

-   Microsoft Excel
-   Excel Tables
-   Excel Formulas
-   Conditional Formatting
-   Sorting and Filtering
-   Data Cleaning
-   PivotTables
-   PivotCharts
-   Dashboard Design

## How to Use This Project

1.  Download the Excel workbook from this repository.
2.  Open the workbook in Microsoft Excel.
3.  Start with the `Raw_Data` sheet to review the original dataset.
4.  Open `Clean_Data` to review the cleaned records.
5.  Review `Analysis` for the main calculations and data-quality
    results.
6.  Review `PivotTables` for summarized analysis.
7.  Open `Dashbroad_1` and `Dashbroad_2` to view the dashboards.
8.  Open `Findings` to review the final findings and recommendations.

## Conclusion

This project demonstrates how Microsoft Excel can be used as a complete
tool for a real estate data-analysis workflow, from raw-data inspection
and cleaning to analysis, visualization, and reporting.

The final dataset contains **292 cleaned property records** and provides
useful information about property prices, sizes, bedrooms, property
types, locations, and listing years.

## Author

**Real Estate Property Data Analysis Using Excel**

University Project
