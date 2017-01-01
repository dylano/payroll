# Google Sheets payroll

A Google Sheets spreadsheet & script to do some simple automated payroll for an hourly employee. Handles calculation and YTD values for a handful of common taxes in US/California:
* Federal income tax [1]
* Social Security
* Medicare
* CA income tax [1]
* CA SDI

The spreadsheet is available here:  
https://docs.google.com/spreadsheets/d/1R4kWt5e5csR3IFdjX_085xPmb7GRWPxSfw964SpLJEY/edit#gid=0

Make a copy of the sheet in your own account and edit employer/employee name & address. 

(The sheet has a placeholder for deductions but they are not implemented. It would not be difficult to do so, I just don't have any need for them in my case.)

[1] Note that income tax calculations are not performed by the sheet and must be entered manually. http://www.paycheckcity.com/calculator/hourly/ is a good place to do that and then input into the sheet.
