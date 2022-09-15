// WITH THE NAME OF ALLAH, THE MOST MERCIFUL, THE MOST BENEFICIENT 

const resources = {

  // Important Spreadsheet IDs 

  strategicId: '1BL6_DpPiIqDv280F4Tq8zo2j9P3lyGj8ArxPiIjBrLA',
  cmykOperationsId: '1FBKUkaupI6kYf9R3TdLVkSNasRvCMZwBsjujU56X7Wk', 
  hrId: '1oj2VqcRLlnRgOm4Q8MQ7PrEHwhosJmL_QWN8kWsd0_A',
  salesId: '1AGzk6kwQXxRpvIHXEla3QXcMrUJnky9VfBF2jEQmDPc',

  // Getters For Spreadsheets, Sheets, And Some Standard Ranges

  strategicSS: function () {

    const ss = SpreadsheetApp.openById(this.strategicId); 
    const sheet = ss.getSheetByName('Strategic Management'); 
    const dataRange = sheet.getRange(2,1, sheet.getLastRow() -1, sheet.getLastColumn()); 
    const values = dataRange.getValues(); 
    const archivedTasksSheet = ss.getSheetByName('archivedTasks'); 
    const cliSheet = ss.getSheetByName('cli'); 

    return {ss, sheet, archivedTasksSheet, dataRange, values, cliSheet}; 

  },

  cmykOperationsSS: function () {

    const ss = SpreadsheetApp.openById(this.cmykOperationsId); 
    const teamsSheet = ss.getSheetByName('teams'); 
    const queSheet = ss.getSheetByName('Que'); 
    const teamsdataRange = teamsSheet.getRange(2,1, strategicTasksSheet.getLastRow() -1, strategicTasksSheet.getLastColumn()); 
    const queDataRange = queSheet.getRange(2,1, strategicTasksSheet.getLastRow() -1, strategicTasksSheet.getLastColumn()); 
    const teamsValues = teamsdataRange.getValues(); 
    const queValues = queDataRange.getValues();  

    return {ss, teamsSheet, queSheet, teamsdataRange, queDataRange, teamsValues, queValues}; 
    
  },

  hrSS: function () {

    const ss = SpreadsheetApp.openById(this.hrId); 
    const employeesSheet = ss.getSheetByName('employees'); 
    const employeesDataRange = employeesSheet.getRange(2,1, employeesSheet.getLastRow() -1, employeesSheet.getLastColumn());
    const employeesDataValues = employeesDataRange.getValues(); 

    return {ss, employeesSheet, employeesDataRange, employeesDataValues};  

  },

  salesSs: function () {
    const ss = SpreadsheetApp.openById(this.salesId); 
    const confirmedOrdersSheet = ss.getSheetByName('Confirmed Orders'); 

    return {ss, confirmedOrdersSheet}; 

  }, 




}

 

function printEmployeesData() {

  const employeesData = resources.hrSS().employeesDataValues[0][6]; 
  resources.hrSS().employeesDataRange.setBackground('green');  
  Logger.log(employeesData); 

}


