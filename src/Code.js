// WITH THE NAME OF ALLAH, THE MOST MERCIFUL, THE MOST BENEFICIENT! 
// PRAISE BE TO ALLAH WHO HAS TAUGHT US WITH PEN


function onOpen(e) {
  // This creates the Organize Menu Item
  var ss = SpreadsheetApp.getUi(); 
  ss.createMenu('Organize')
  .addItem('Archive Completed Tasks', 'archiveCompleted')
  .addToUi(); 

}

 

function editTrigger1() {

  ScriptApp.newTrigger('autoNumberAndDate')
  .forSpreadsheet(resources.strategicSS().ss)
  .onEdit()
  .create(); 

}

function editTrigger2() {

  // const ss = SpreadsheetApp.getActive(); 
  ScriptApp.newTrigger('detectAssignment')
  .forSpreadsheet(resources.strategicSS().ss)
  .onEdit()
  .create(); 

}; 


function autoNumberAndDate(e) {

  const strategicTasksSheet = resources.strategicSS().sheet; //Check

// var ss = SpreadsheetApp.getActiveSpreadsheet(); 
// var sheet = ss.getSheetByName("Strategic Management"); 
var autoNumberRange = strategicTasksSheet.getRange(e.range.getRow(), 1);
var upperNumberValue = strategicTasksSheet.getRange(e.range.getRow() -1, 1).getValue();
var dateRange = strategicTasksSheet.getRange(e.range.getRow(), 2); 
var taskHeaderValue = strategicTasksSheet.getRange(1, e.range.getColumn()).getValue(); 

// const autoNumberFirstPartValue = 'CEO'; 
// const autoNumberSecondpartValue = 0; 
// const autoNumberCurrentValues = strategicTasksSheet.getRange(2, 1, strategicTasksSheet.getLastRow(), 1).getValues();
// Logger.log(autoNumberCurrentValues); 

// const maxAutoNumberValue = Math.max(...autoNumberCurrentValues);
const newTaskIdNumber = Math.floor(Math.random() * 10000000000); 
// const uniqueTaskId = autoNumberFirstPartValue + "-" + newTaskIdNumber; 
// Logger.log(`The Max Autonumber Value is ${maxAutoNumberValue}`); 

if(taskHeaderValue === 'Task') {
autoNumberRange.setValue(newTaskIdNumber);

dateRange.setValue(new Date()); 
}; 
};



function archiveCompleted() {

  const strategicTasksSheet = resources.strategicSS().sheet; 

  // var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  // var sheet = ss.getSheetByName('Strategic Management'); 
  var range = strategicTasksSheet.getRange(2,1,strategicTasksSheet.getLastRow() -1, strategicTasksSheet.getLastColumn()); 
  var tasksArray = range.getValues(); 
  // var strategicSpreadSheet = resources.strategicSS(); 
  var archiveSheet = resources.strategicSS().archivedTasksSheet; 

  const employeesWithCompletedTasks = []; 

tasksArray.forEach(row => {

  if((row[6] != 'Me' || row[6] != 'me' || row[6] != '') && (row[7] == 'Completed' || row[7] == 'Cancelled')) {
    employeesWithCompletedTasks.push(row[6]); 
  }
}); 

Logger.log('Completed tasks employees will come here'); 
Logger.log(employeesWithCompletedTasks); 

const uniqueEmployees = employeesWithCompletedTasks.filter ((name, index, array) => {

// this statement returns unique values in an array! quite powerful! 
  return array.indexOf(name) === index; 


}); 

Logger.log('array with unique values will come here');
Logger.log(uniqueEmployees); 


uniqueEmployees.forEach( employee => {

  

  const employeeTargetSheetId = fun.findTargetEmployeeSpreadSheet(employee).employeeSpreadsheetId; 
  const ss = SpreadsheetApp.openById(employeeTargetSheetId); 
  const sheet = ss.getSheetByName('Tasks'); 

  const sheetA2CellData = sheet.getRange("A2").getValue(); 

  Logger.log ( "First Data Cell Value will come here"); 

  Logger.log ( sheetA2CellData); 

  if ( sheetA2CellData !== "") { 
    Logger.log("The employee task sheet has tasks data therefore code will execute! ")

const employeetaskSheetRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()); 


  const employeeTasksArray = employeetaskSheetRange.getValues(); 

  Logger.log('target employee Tasks array will come here'); 

  Logger.log(employeeTasksArray); 

  const employeeCompletedTasksArray = employeeTasksArray.filter ( task => {

    if(task[7] === 'Completed' || task[7] === 'Cancelled') {
      return task; 
    }; 
  }); 

  Logger.log('Completed employee Tasks array will come here'); 
  Logger.log(employeeCompletedTasksArray); // so far so good! 
  
  if (employeeCompletedTasksArray.length !== 0) {
    Logger.log ("The employee has completed tasks therefore code will execute from here!"); 

     const employeePendingTasksArray = employeeTasksArray.filter ( task => {

    return task[7] !== 'Completed'; 

  }); 

  Logger.log('Pending employee Tasks array will come here'); 
  Logger.log(employeePendingTasksArray); // pending tasks array seems to be having issue



  employeetaskSheetRange.clearContent(); 

   const employeeArchivedTasksSheet = ss.getSheetByName('archivedTasks'); 
  const archivedTasksRange = employeeArchivedTasksSheet.getRange(employeeArchivedTasksSheet.getLastRow() + 1, 1,  employeeCompletedTasksArray.length, employeeCompletedTasksArray[0].length); 

  archivedTasksRange.setValues(employeeCompletedTasksArray); 

  if(employeePendingTasksArray.length !== 0) {

    const pendingTasksRange = sheet.getRange(2, 1, employeePendingTasksArray.length, employeePendingTasksArray[0].length); 
  pendingTasksRange.setValues(employeePendingTasksArray); 

  } 

  } else if (employeeCompletedTasksArray.length === 0){

    Logger.log("Although employee tasks sheet is not empty but it has no completed tasks therefore code will skip from here! "); 

  }
 

  } else if (sheetA2CellData === "") {
    Logger.log("Employee has no tasks in its sheet therefore code will skip from here! ")
  }
  
   // Something Wrong Here! lets find out what's wrong here! 
  // Problem Found! If a selected sheet has empty rows, then due to the function sheet.getLastRow() error will be generated! because this function tries to find content and based on the content, tries to get the last row! When there is no content such problem gets in the way. 


  
  

})

  var completedTasksArray = tasksArray.filter( task => {
    return task[7] === "Completed"; 
  }); 

  if ( completedTasksArray.length !== 0) {

    Logger.log("The Primary SCCS has completed tasks therefore code will execute from here!")

    var archiveSheetRange = archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, completedTasksArray.length , completedTasksArray[0].length); 
  archiveSheetRange.setValues(completedTasksArray); 

  var pendingTasksArray = tasksArray.filter( task => {
    return task[7] !== "Completed"; 
  });

  range.clearContent(); 

  if (pendingTasksArray.length !== 0) {
    Logger.log("The Primary SCCS has pending tasks therefore the code will execute from here!"); 
    
  var pendingTasksRange = strategicTasksSheet.getRange(2,1, pendingTasksArray.length, pendingTasksArray[0].length); 
  pendingTasksRange.setValues(pendingTasksArray); 

  } else if ( pendingTasksArray.length === 0) {

    Logger.log("The primary SCCS has no pending tasks therefore code will skip from here!")
  }; 
  

  } else if ( completedTasksArray.length === 0) {
    Logger.log("There are no completed tasks, therefore the code will skip from here! ")
  }; 
  
}; 
