# Price-Corrector

## Introduction
Openpyxl is a library for reading and writing Excel 2010 xlsx and xlsm files. This project provides a function to process a given workbook.

## Functionality
The function takes a filename as input and for each row in the workbook, it calculates the corrected price by multiplying the price by 0.9 and updates the corrected price in a new column. It also creates a bar chart and adds it to the workbook. The chart displays the corrected prices in the new column. The function saves the updated workbook with the same filename.

## Limitations
The function currently works for a single workbook, it will not work for thousands of workbooks.

## Usage
```python
from spreadsheet import process_workbook
process_workbook('filename.xlsx')

