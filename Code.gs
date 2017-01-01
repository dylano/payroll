/**
 * Add menu items.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Run Payroll', functionName: 'runPayroll'},
    {name: 'Delete Paystub', functionName: 'deletePayroll'}, 
    {name: 'Rebuild YTD from Stubs', functionName: 'rebuildYTD'}
  ];
  spreadsheet.addMenu('Payroll', menuItems);
}

function deletePayroll() {
  // prompt for paystub date 
  var strDel = Browser.inputBox('Delete Payroll', 'Please enter the date of the paystub to remove (mm/dd/yyyy):', Browser.Buttons.OK_CANCEL);
  if (strDel == 'cancel') {
    return;
  }

  // search for sheet and confirm deletion
  var spreadsheet = SpreadsheetApp.getActive();
  var deleteSheet = spreadsheet.getSheetByName(strDel);
  if(deleteSheet) {
    var resp = Browser.msgBox('Are you sure?', Utilities.formatString("Located the payroll for %s -- continue with deletion?", strDel), Browser.Buttons.YES_NO);
    if(resp == 'no') {
       return;
    }
  } else {
    Browser.msgBox('Nothing to Delete', strDel, Browser.Buttons.OK);
    return;
  }
  
  // Subtract from YTD sheet
  removePayrollFromYTD(deleteSheet);
  
  // Delete tab and return to template
  spreadsheet.deleteSheet(deleteSheet);
  SpreadsheetApp.flush();
  spreadsheet.getSheetByName('Template').activate();
  
  // todo: is it possible to trap the Delete Sheet action from UI and trigger this function? or at least ask if we should adjust YTD before deleting
}

function rebuildYTD() {
  Browser.msgBox('Payroll', "Haven't gotten around to this feature yet...", Browser.Buttons.OK);
  // todo: 
  // copy current YTD to backup column 
  // zero out current YTD values
  // cycle through all paystubs to sum new values
}

function runPayroll() {
  var spreadsheet = SpreadsheetApp.getActive();
  var templateSheet = spreadsheet.getSheetByName('Template');
  
  // Validate weekly input
  if(!validateInput())
  {
    return;
  }
    
  // Create new sheet for pay date
  var sheetName = Utilities.formatDate(getNamedValue('TempPayDate'), "PST", "MM/dd/yyyy");
  var paystubSheet = spreadsheet.getSheetByName(sheetName);
  if (paystubSheet) {
    removePayrollFromYTD(paystubSheet);
    paystubSheet.clear();
    SpreadsheetApp.flush();  // ran into a race issue here with YTD -- flush() seems to fix it
  } else {
    paystubSheet = spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
  }  

  // Copy pay data from Template to new sheet 
  var dataRange = templateSheet.getRange(1, 1, 45, 6);
  dataRange.copyValuesToRange(paystubSheet, 1, 6, 1, 45);
  dataRange.copyFormatToRange(paystubSheet, 1, 6, 1, 45);
  paystubSheet.setRowHeight(1, 10);
  paystubSheet.setRowHeight(2, 10);
  paystubSheet.setRowHeight(43, 10);
  paystubSheet.setRowHeight(44, 10);
    
  // Update YTD sheet. Grab current values from template before we clean it up
  updateYTDFromCurrent();
    
  // Reset template data
  cleanup();
  
  // Activate new paystub and pre-set selection for easy printing
  paystubSheet.activate();
  paystubSheet.setActiveSelection(paystubSheet.getRange(1, 1, 45, 6));

//  Logger.log(sheetName);
//  dumpLog();
}

function validateInput() {

  // hours
  var hrsReg = getNamedValue('TempHoursReg');
  var hrsOT = getNamedValue('TempHoursOT');
  if(hrsReg + hrsOT <= 0) {
    Browser.msgBox('Error', 'Work harder. Enter some hours.', Browser.Buttons.OK);
    return 0;
  }

  // taxes
  var taxFed = getNamedValue('TempFedTax');
  var taxCa = getNamedValue('TempCaTax');
  if(taxFed <= 0 || taxCa <= 0) {
    Browser.msgBox('Error', 'Pay your fair share. Enter some taxes.', Browser.Buttons.OK);
    return 0;
  }
    
  // dates

  // very basic sanity on pay period
  var begin = getNamedValue('TempPeriodBegin');
  var end = getNamedValue('TempPeriodEnd');
  if(end < begin) {
    Browser.msgBox('Error', 'Pay period looks funky.', Browser.Buttons.OK);
    return 0;
  }

  // warn on duplicate payroll
  var payDate = getNamedValue('TempPayDate');
  var sheetName = Utilities.formatDate(payDate, "PST", "MM/dd/yyyy");
  var paystubSheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (paystubSheet) {
    var resp = Browser.msgBox('Do Over?', 'You already have a paycheck for this date -- are you sure you want to overwrite it?', Browser.Buttons.YES_NO);
    if(resp == 'no') {
       return 0;
    }
  }
  
  return 1;
}

function updateYTDFromCurrent() {
  incrementNamedValue('YTDFedTax', getNamedValue('CurFedTax'));
  incrementNamedValue('YTDSocSec', getNamedValue('CurSocSec'));
  incrementNamedValue('YTDMedicare', getNamedValue('CurMedicare'));
  incrementNamedValue('YTDCaTax', getNamedValue('CurCaTax'));
  incrementNamedValue('YTDCaSdi', getNamedValue('CurCaSdi'));
  incrementNamedValue('YTDWagesReg', getNamedValue('CurWagesReg'));
  incrementNamedValue('YTDWagesOT', getNamedValue('CurWagesOT'));
}

function cleanup() {
  clearNamedValue('TempHoursReg');
  clearNamedValue('TempHoursOT');
  clearNamedValue('TempFedTax');
  clearNamedValue('TempCaTax');
/*  
  clearNamedValue('TempPeriodBegin');
  clearNamedValue('TempPeriodEnd');
  clearNamedValue('TempPayDate');
*/
}

function getNamedValue(valueName) {
  var nr = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(valueName);
  return nr.getValue();
}

function incrementNamedValue(valueName, valueToAdd) {
  var nr = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(valueName);
  nr.setValue(nr.getValue() + valueToAdd);
}

function clearNamedValue(valueName) {
  var nr = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(valueName);
  if(nr) {
    nr.clearContent();
  }
}

function removePayrollFromYTD(payrollSheet) {
  // note: this will be fragile, we don't have named ranges for the individual payroll sheets so using specific cell IDs instead
  //       any future change in paystub structure would need to be applied to all past paystubs too. Or just implement named ranges for all paystubs.
  var wagesReg = payrollSheet.getRange('E30');
  incrementNamedValue('YTDWagesReg', -wagesReg.getValue());
  
  var wagesOT = payrollSheet.getRange('E31');
  incrementNamedValue('YTDWagesOT', -wagesOT.getValue());
  
  var fedTax = payrollSheet.getRange('E38');
  incrementNamedValue('YTDFedTax', -fedTax.getValue());

  var caTax = payrollSheet.getRange('E41');
  incrementNamedValue('YTDCaTax', -caTax.getValue());

  var caSdi = payrollSheet.getRange('E42');
  incrementNamedValue('YTDCaSdi', -caSdi.getValue());

  var socSec = payrollSheet.getRange('E39');
  incrementNamedValue('YTDSocSec', -socSec.getValue());

  var medicare = payrollSheet.getRange('E40');
  incrementNamedValue('YTDMedicare', -medicare.getValue());
}

function dumpLog() {
  var logSheet = SpreadsheetApp.getActive().getSheetByName('devlog');
  if(!logSheet) {
    logSheet = SpreadsheetApp.getActive().insertSheet('devlog', 0); 
  }

  var logData = Logger.getLog();
  var a1 = logSheet.getRange(1,1);
  a1.setValue(logData);
}
