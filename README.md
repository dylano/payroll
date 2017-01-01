# Google Sheets payroll

A Google Sheets spreadsheet & script to do some simple automated payroll for an hourly employee. Handles calculation and year-to-date values for a handful of common taxes in US/California:
* Federal income tax *(see note below)*
* Social Security
* Medicare
* CA income tax *(see note below)*
* CA SDI

#### Notes
1. Income tax calculations are not performed by the sheet and must be entered manually. Those calculations are complex and require several more inputs than we want to deal with on this sheet. http://www.paycheckcity.com/calculator/hourly/ is a good place to use for PIT calculation.
1. The sheet is limited to one paystub per pay date, which essentially means it is only useful for paying one person. 
1. The sheet has a placeholder for deductions but they are not implemented. It would not be difficult to do so, I just don't have any need for them in my case.
1. YTD is not super intelligent. It generally expects that you are creating checks in date order and it keeps a running tally of everything, but doesn't attempt to keep everything consistent on every paystub.
  * For example, if you create three paystubs and then delete the one in the middle, the third paystub will not be edited to reflect those now missing wages/taxes. However, the YTD sheet (which contains the cached values) will be updated and when you create a fourth check, it will show the proper YTD values.

#### Usage
##### Spreadsheet 
* https://docs.google.com/spreadsheets/d/1R4kWt5e5csR3IFdjX_085xPmb7GRWPxSfw964SpLJEY/edit#gid=0  
Make a copy of the sheet in your own account and edit employer/employee name & address. 

##### New paystubs
* Enter pay period, hours, and tax detail into the box on the Template sheet.
* Use "Run Payroll" from the Payroll menu to create the paystub.
* You'll be taken to a new sheet with the resulting paystub. From there you can print, save to PDF, etc. to provide a copy to your employee.
  * Each additional pay date will add a new sheet. Do not remove these sheets, unless you need to make an adjustment...

##### Adjustments
* If you need to edit a paystub **_for the same pay date_**, just create a new payroll from the Template sheet and you'll be prompted for confirmation that you are overwriting the old paystub. YTD values will be adjusted accordingly.
* If you need to change the pay date or just delete a paystub altogether, use the "Delete Paystub" function from the Payroll menu. This will back out that check's values from the YTD calculations so that the next check created will show the correct YTD.
