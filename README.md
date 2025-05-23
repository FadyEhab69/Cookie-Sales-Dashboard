# Overview
This report documents the data cleaning, transformation, and visualization process for the "Cookie By Company" dataset, originally provided in Cookies.xlsx. As a data analyst, I leveraged Microsoft Power Query and Power Pivot to preprocess the data, followed by creating an interactive Excel dashboard to derive actionable insights into profit trends and sales distribution across products and countries. The dashboard, developed using Excel's charting capabilities, provides a comprehensive view of the business performance for the year 2019.

# Data Cleaning and Transformation Process
Initial Data Assessment
The raw dataset contained multiple sheets within Cookies.xlsx, including pivot tables summarizing Profit and Units Sold by country, product, and month, as well as a detailed transactional dataset. The initial assessment revealed several data quality issues:

Missing Values: The Profit and Units Sold pivot tables included a (blank) category, indicating missing product or country data. The transactional data had sporadic missing entries for Cost and inconsistent Date formats (e.g., 43709, 43800).
Inconsistent Formatting: Column headers and values contained extra spaces or non-standard representations (e.g., Chocolate Chip vs. inconsistently cased variants).
Data Type Mismatches: Dates were stored as Excel serial numbers (e.g., 43709 corresponding to 2019-09-01), and numerical fields like Revenue, Cost, and Profit required type validation.
Outliers and Anomalies: Unusually high Profit values (e.g., 12753 for a single Chocolate Chip transaction) suggested potential data entry errors or bulk sales.
Incomplete Relationships: The transactional data lacked explicit links between Country, Product, and Date for aggregation consistency with pivot tables.
Cleaning Steps Using Power Query
Data Loading and Structuring:
Imported all sheets into Power Query to consolidate the pivot tables (Sum of Profit, Sum of Units Sold) and transactional data into a unified model.
Removed duplicate rows based on Country, Product, and Date combinations to ensure uniqueness.
Handling Missing Values:
Filtered out rows where Product or Country was (blank) in pivot tables to eliminate incomplete records.
In the transactional data, calculated Profit where missing by subtracting Cost from Revenue (e.g., Revenue 8625 - Cost 3450 = Profit 5175 for India, Chocolate Chip, Date 43770).
Removed rows with invalid or negative Profit values to address potential data corruption.
Standardization:
Trimmed and capitalized column headers (e.g., country → Country, product → Product).
Standardized product names (e.g., ensured Chocolate Chip, Fortune Cookie, etc., were consistently formatted across datasets).
Converted Excel serial dates (e.g., 43709) to a readable YYYY-MM-DD format (e.g., 2019-09-01) using Power Query’s date transformation.
Type Conversion:
Ensured Units Sold, Revenue, Cost, and Profit were numeric types, resolving any text-formatted numbers.
Parsed Date into a date type for time-series analysis.
Outlier Detection and Correction:
Identified outliers using statistical measures (e.g., Profit > 10,000 was flagged as potential bulk sales or errors).
Cross-verified high Profit values (e.g., 12753) against Units Sold and Revenue to confirm plausibility; retained valid bulk transactions and flagged others for review.
Relationship Establishment:
Used Power Pivot to create relationships between the transactional data and pivot tables using Country, Product, and Date as keys.
Aggregated transactional data to match pivot table summaries, ensuring consistency (e.g., total Profit for India, Chocolate Chip = 227529).
Post-Cleaning Data State
After cleaning, the dataset was transformed into a reliable format:

Completeness: All Profit and Units Sold values were populated or calculated, with (blank) categories eliminated.
Consistency: Uniform date formats (2019-01-01 to 2019-12-31) and standardized product/country names.
Accuracy: Outliers were validated, and relationships enabled accurate aggregations (e.g., total Profit across all countries and products = 2,763,364.45).
Example Transformation:
Before: Country: India, Product: Chocolate Chip, Date: 43770, Units Sold: 1725, Revenue: 8625, Cost: 3450, Profit: (null)
After: Country: India, Product: Chocolate Chip, Date: 2019-09-05, Units Sold: 1725, Revenue: 8625, Cost: 3450, Profit: 5175
# Dashboard Creation and Insights
Dashboard Design
The Excel dashboard, titled "Cookie By Company," 
features:

Profit by Product in Month: A line chart tracking monthly profit trends for each product (e.g., Chocolate Chip peaked at ~60,000 in December 2019).
Profit by Product & Country: A bar chart comparing profit across countries for each product, highlighting India’s dominance in Chocolate Chip (227,529).
Unit Sold by Country: A pie chart showing sales distribution (e.g., India 22%, Malaysia 21%).
Interactive Filters: Dropdowns for Date, Product, and Country to drill down into specific data segments.

# Key Insights
Profit Trends: Chocolate Chip consistently led in profit, with a sharp increase from January (23,599) to December (52,619), suggesting seasonal demand.
Geographical Performance: India and the United Kingdom were top performers, with profits of 606,331.65 and 595,071.60, respectively, driven by high Chocolate Chip and White Chocolate Macadamia Nut sales.
Product Popularity: Chocolate Chip and White Chocolate Macadamia Nut dominated units sold (335,894 and 162,426), indicating strong market preference.
Unexpected Finding: The (blank) category in profit data (7,047) suggests untracked sales, warranting further investigation into data collection processes.
# Analytical Skills Demonstrated
Data Cleaning: Effectively used Power Query to handle missing data, standardize formats, and resolve inconsistencies.
Data Modeling: Leveraged Power Pivot to establish relationships and validate aggregations, ensuring data integrity.
Visualization: Created an interactive, visually appealing dashboard using Excel charts, tailored to business needs.
Insight Generation: Identified trends, regional strengths, and anomalies, providing actionable recommendations.
# Conclusion
This analysis transformed a raw dataset into a robust foundation for business decision-making. The dashboard offers a clear view of profit and sales dynamics, with recommendations to focus marketing on high-performing products (e.g., Chocolate Chip) in key markets (e.g., India) and address data gaps (e.g., (blank) entries). Future work could include predictive modeling to forecast 2020 trends.
