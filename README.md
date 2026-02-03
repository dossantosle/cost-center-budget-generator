# 📊 Budget Data Generator
Python automation script for generating structured budget control tables from Excel data.


## Description
This project automates the creation of structured tables used for budget and cost control.

It reads data from Excel sheets, applies predefined mappings and generates multiple output files ready for analysis.

## What it does
- Reads Excel input files
- Maps managers to cost centers
- Combines cost centers, months and cost classes
- Generates clean output tables automatically

## Output files
- CostCenter_Researcher_YEAR.xlsx
- CostCenter_Month_CostClass_YEAR.xlsx
- Budget_Researcher_Month_CostClass_YEAR.xlsx

All files are saved in the output folder.

## Tech
- Python
- pandas
- Excel

## How to use
1. Set the input Excel file path in the script
2. Update mappings and year if needed
3. Run the script
