# VBA Stock Data Analysis Macro

## Overview
This project is a VBA Macro created to automate the analysis of stock data across multiple worksheets in an Excel file. 
The macro loops through each worksheet (each representing different stock quarters) and outputs key metrics like:

- Ticker Symbol
- Quarterly Change ($)
- Percent Change (%)
- Total Stock Volume

It also highlights positive changes in green and negative changes in red and identifies:
- Greatest % Increase
- Greatest % Decrease
- Greatest Total Stock Volume

## Files Included
- `alphabetical_testing.xlsx` — Test file with smaller dataset
- `Multiple_year_stock_data.xlsx` — Full dataset file for final run
- `Module1.bas` — VBA code file (exported)
- `Screenshots/` — Folder containing screenshots of output results

## How It Works
1. The macro loops through every worksheet using `For Each ws In Worksheets`.
2. It calculates the open price, close price, total volume, quarterly change, and percent change for each stock ticker.
3. It writes a summary table starting in column I (columns 9 to 12).
4. It uses conditional formatting to color-code the quarterly and percent changes.
5. It identifies and records the greatest percentage increase, decrease, and highest total volume for each worksheet.

## Technologies Used
- Microsoft Excel for Mac (Version 16.77.1)
- Visual Basic for Applications (VBA)

## Date
April 28, 2025

## Author
(Student Name Here)

---
