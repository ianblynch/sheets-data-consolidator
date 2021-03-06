# sheets-data-consolidator

I made this Apps Script for Google Sheets.  It will take data from multiple source spreadsheets and consolidate it into one destination sheet.

It will set a long formula made of consecutive formulas of format =query(importrange(...),...)

# Example

Here is an example sheet setup for use:
https://docs.google.com/spreadsheets/d/1dts7pZruCYmABZZzQq7P12GSArqRXZRPqxuq9DtsgRg/edit?usp=sharing

Here is the example source sheet it draws from:
https://docs.google.com/spreadsheets/d/1XVGMySns6RMRNPHhhClrJ6zpCggq2U_4vo6YAFZac5U/edit?usp=sharing

# Setup

Configure the layout of your spreadsheet on the 'Package' sheet:

Destination Sheet - Name of sheet in this workbook you want the data routed to.  Data will be deposited starting in cell A2.

Array Sheet - The name of the sheet in this workbook holding the Name, Range, and Query parameters for each sheet to be imported. 

Reusable Variables Sheet - The name of the sheet holding reusable variables for Range and Query parameters.

# Usage

Once you've setup the Package sheet and it's Destination Sheet, Array Sheet, and Reusable Variables Sheet, you simply use the 'Data Import' -> 'Inclusive Data Import' menu item to import your data into the Destination Sheet.

The 'Inclusive Data Import' inconsistently trims the trailing rows so I have included a 'Trim Active Sheet Trailing Rows' utility for your convenience.