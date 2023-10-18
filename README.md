# Coffee Sales Dashboard
This repository contains my end to end Microsoft Excel project.

This project showcases an in-depth data analysis of coffee sales using data gathering & transformation 
leveraging pivot tables and pivot charts to build a dynamic & interactive Coffee Sales Dashboard. The analysis
provides insightful information into coffee sales and customers relative to date and country.

Put link

Below are the steps followed.
### Step 1 - Data Gathering
* Populated Customers, Email and Country columns from raw data file (put link). Gathered data using
INDEX - MATCH formula.

   (e.g. for Customers column: =INDEX(customers!$B$1:$B$1001;MATCH(orders!C2;customers!$A$1:$A$1001;0))

* Eliminated number "0" from empty email cells when value was missing using IF function.

   (i.e. =IF(INDEX(customers!$C$1:$C$1001;MATCH(orders!C2;customers!$A$1:$A$1001;0))=0;"";
  (INDEX(customers!$C$1:$C$1001;MATCH(orders!C2;customers!$A$1:$A$1001;0))))

* Populated Coffee Type, Roast Type, Size and Unit Price columns. Gathered data using INDEX - MATCH - MATCH
  formula. MATCH to match row number and MATCH to match column number. This one formula populated all four
  columns at once.

     (e.g. for Coffee Type column: =INDEX(products!$A$1:$G$49;MATCH(orders!$D2;products!$A$1:$A$49;0);
  MATCH(orders!I$1;products!$A$1:$G$1;0))

* Populated Sales column. Gathered data using formula: Unit Price * Quantity

* Inserted new column named Loyalty card. Gathered data using INDEX - MATCH formula.

  (i.e. =INDEX(customers!$I$1:$I$1001;MATCH([@[Customer ID]];customers!$A$1:$A$1001;0))

### Step 2 - Data Transformation
* Formatted Coffee Type and Roast Type columns from abbreviations to full names in new columns Coffee Type
  Name and Roast Type Name respectively using multiple IF functions
  
* Formatted Order Date column to month letter abbreviation (Jun, Jul, Sep)
  
* Formatted Size column adding decimals and "kg"
  
* Formatted Unit Price and Sales adding currency $
  
* Converted range to table in order to update automatically when adding/removing data

### Step 3 - Pivot Tables & Pivot Charts
* Created Pivot Table *TotalSales* adding Years and Order Date fields into Rows, Coffee Type Name field
  into Columns and Sales field into Values

* Inserted and Formatted 2D Line Chart 

* Inserted and Formatted Timeline 

* Inserted and Formatted three Slicers (Size, Roast Type Name, Loyalty Card) 

* Created Pivot Table *CountryBarChart* adding Country field into Rows and Sales field into Values

* Inserted and Formatted 2D Bar Chart

* Created Pivot Table *Top5Customers* (by duplicating *CountryBarChart*) adding Customer Name field into Rows and Sales field into Values

* Filtered by Top 5 in Customer Name and sorted by Sum of Sales

* Formatted the Bar Chart

### Step 4 - Dashboard

* Created New Worksheet and c/p Pivot Charts, Timeline and Slicers

* Formatted the dashboard

* Report connections from Timeline and Slicers to *CountryBarChart* and *Top5Customers* to affect all visuals

Feel free to reach out if you have any questions or feedback. 

Thank you.

















