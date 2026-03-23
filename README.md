# automated-financial-reporting

# Automated Financial Reporting System

A Python automation pipeline that pulls raw financial data, 
cleans it, and generates a formatted Excel report automatically.

## What it does
- Reads messy raw CSV data (currency symbols, missing values, inconsistent formats)
- Cleans and transforms the data using pandas
- Generates a multi-sheet Excel report with charts in under 1 second

## Sheets in the report
- Monthly P&L — Revenue, COGS, Gross Profit, Net Income by month
- Regional Breakdown — Sales performance by region with bar chart
- Expense Breakdown — Full cost breakdown by category

## Tech stack
- Python 3
- pandas — data cleaning and aggregation
- openpyxl — Excel generation and formatting
- numpy — missing value handling

## How to run

1. Install dependencies
pip install pandas openpyxl numpy

2. Run the report
python generate_report.py

3. Open Financial_Report_2024.xlsx