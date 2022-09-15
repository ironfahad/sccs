function personalTasksTrigger() {
  ScriptApp.newTrigger('detectPersonalTask')
  .forSpreadsheet(resources.strategicSS().ss)
  .onEdit()
  .create()
}

const copyDataToTargetSpreadsheet = (targetPersonData, valueOfInput, idOfTask) => {

  const inputCommunicationsValue = valueOfInput; 

  const taskId = idOfTask;
  
  const dataOfTargetPerson = targetPersonData;  
  Logger.log('data of the person will come below '); 

  Logger.log(dataOfTargetPerson); // so far so good! 

  Logger.log('why is the data below not coming? ')

  Logger.log(dataOfTargetPerson[0][6]); 


  const targetSpreadsheet = SpreadsheetApp.openById(dataOfTargetPerson[0][6]);
  const targetSheet = targetSpreadsheet.getSheetByName('Tasks'); 
  const targetSheetDataRange = targetSheet.getRange(2, 1, targetSheet.getLastRow() -1, targetSheet.getLastColumn()); 
  const targetSheetArray = targetSheetDataRange.getValues(); 
  Logger.log('targetsheet array is below')
  Logger.log(targetSheetArray)

  // const parentTasksSheetDataRange = parentTasksSheet.getDataRange(); 
  // const parentTasksSheetArray = parentTasksSheetDataRange.getValues(); 

  const relatedTaskRow = targetSheetArray.filter( task => {
    return task[0] === taskId; 
  }); 
  Logger.log('related task row starts below')
  Logger.log(relatedTaskRow); 
  const relatedRowIndex = targetSheetArray.indexOf(relatedTaskRow[0]); 
  Logger.log(relatedRowIndex); 
  targetSheetArray[relatedRowIndex][9] = inputCommunicationsValue; 
  
  targetSheetDataRange.clearContent(); 
  targetSheetDataRange.setValues(targetSheetArray); 
  
  // const rowForInputCommunicationDataRange = parentTasksSheet.getRange(relatedRowIndex + 1, 1, 1, parentTasksSheet.getLastColumn()); 


}


const findTargetSpreadSheet = (targetPersonName) => {

  const personName = targetPersonName; 
  Logger.log(personName); 

  const totalEmployeesArray = resources.hrSS().employeesSheet.getDataRange().getValues(); 

  Logger.log(totalEmployeesArray); 

  const targetPersonDataRow = totalEmployeesArray.filter( member => {
    return member[1].includes(targetPersonName); 
  })

  Logger.log(targetPersonDataRow); 

  const targetPersonSpreadSheetId = targetPersonDataRow[0][6]; 
  Logger.log(`the target person spreadsheet is ${targetPersonSpreadSheetId}`); 

  return {targetPersonDataRow, targetPersonSpreadSheetId}; 

}; 

const toggleBackgroundColorWhite = (e) => {
const sourceRowRange = resources.strategicSS().sheet.getRange(e.range.getRow(), 1, 1, resources.strategicSS().sheet.getLastColumn()); 

sourceRowRange.setBackground("white"); 

}; 

const sendNotificationEmail = (emailPersonName, personEmailAddress, spreadSheetId, taskType ) => {


  const teamsArray = resources.hrSS().employeesSheet.getDataRange().getValues(); 
  Logger.log('Teams Array For Send Notificaiton will come here!'); 
  Logger.log(teamsArray); 
  const selectedPersonDataRow = teamsArray.filter( employee => {
    return employee[1].includes(emailPersonName); 
  })

  Logger.log(selectedPersonDataRow); 

  const personName = emailPersonName; 
  const typeofTask = taskType; 
  const emailID = personEmailAddress; 
  Logger.log(emailID);  
  const subject = `New Task Assigned!`
  // const assignedPersonSpreadSheetId = spreadSheetId; 
  const assignedPersonSpreadsheet = SpreadsheetApp.openById(spreadSheetId); 
  const assignedPersonSpreadSheetUrl = assignedPersonSpreadsheet.getUrl(); 
  const personTasksSheet = assignedPersonSpreadsheet.getSheetByName('Tasks'); 
  const tasksSheetID = personTasksSheet.getSheetId(); 
  const targetedSheetUrl = assignedPersonSpreadSheetUrl + "#gid=" + tasksSheetID; 
  Logger.log(targetedSheetUrl); 
  Logger.log(assignedPersonSpreadSheetUrl); 
  const emailBody = `Dear ${personName},  You have received a new Task related to ${typeofTask}. Please click on the following link to directly goto the task sheet. ${targetedSheetUrl}`; 

  GmailApp.sendEmail(emailID, subject, emailBody); 


}

const assignTask = (taskArrayValues) => {

  const taskArray = taskArrayValues;  
  Logger.log('Assignment Task Array values will come here')
  Logger.log(taskArray); 

  const employeeName = taskArray[0][6]; 
  Logger.log(employeeName); 
  const employeesTotalDataArray = resources.hrSS().employeesSheet.getDataRange().getValues(); 

  const employeeDataArray = employeesTotalDataArray.filter( employeeData => {
    return employeeData[1].includes(employeeName); 
  }); 

  Logger.log(employeeDataArray); 

  const employeeSpreadSheetId = employeeDataArray[0][6]; 
  const employeeSheetName = 'Tasks'; 
  const employeeSpreadSheet = SpreadsheetApp.openById(employeeSpreadSheetId); 
  const employeeTasksSheet = employeeSpreadSheet.getSheetByName(employeeSheetName);
  const newTaskRowRange = employeeTasksSheet.getRange(employeeTasksSheet.getLastRow() + 1, 1, 1, employeeTasksSheet.getLastColumn());
  newTaskRowRange.setValues(taskArray); 

  const reformattedTask = taskArray[0][3].slice(4, taskArray[0][3].length); 
  const taskCellRange = employeeTasksSheet.getRange(employeeTasksSheet.getLastRow(), 4); 

  taskCellRange.setValue(reformattedTask); 

  //New Code Data Entry from here 

  // const targetSheetTotalTasksArray = SpreadsheetApp.openById(employeeSpreadSheetId).getSheetByName('Tasks').getRange(2, 1, employeeTasksSheet.getLastRow() - 1, employeeTasksSheet.getLastColumn()).getValues(); 

  // const uniqueTasksArray = fun.generateUniqueArray(targetSheetTotalTasksArray); 

  // const totalTasksRange = employeeTasksSheet.getRange(2, 1, employeeTasksSheet.getLastRow() -1, employeeTasksSheet.getLastColumn()); 

  // totalTasksRange.setValues(uniqueTasksArray); 

  // New code entry ends here! 

  return {employeeDataArray};  
  

}



function detectPersonalTask(e) {

  const strategicTasksSheet = resources.strategicSS().sheet; 

  const tasksRowPersonValue = strategicTasksSheet.getRange(e.range.getRow(), e.range.getColumn()).getValue();

  const tasksCellRange = strategicTasksSheet.getRange(e.range.getRow(), e.range.getColumn() - 3); 
  const tasksCellValue = tasksCellRange.getValue(); 
  const taskArray = strategicTasksSheet.getRange(e.range.getRow(), 1, 1, e.range.getColumn() + 3).getValues(); 
  

  Logger.log(taskArray); 


  Logger.log(tasksCellValue); 

  Logger.log(strategicTasksSheet.getRange(1, e.range.getColumn()).getValue()); 

  const headerCellValue = strategicTasksSheet.getRange(1, e.range.getColumn()).getValue(); 


  if(strategicTasksSheet.getRange(1, e.range.getColumn()).getValue() === 'Delegate' && fun.getEventData(e).taskID > 0 ) {

   fun.unHighLightColor(resources.strategicId, resources.strategicSS().sheet.getSheetName(), fun.findRowNumber(resources.strategicId, resources.strategicSS().sheet.getSheetName(), taskArray[0][0])); 

      if(tasksRowPersonValue !== null && tasksCellValue.includes("at:")) {

    Logger.log('Hello World - The value includes at:'); 

    

    const assignTsk = assignTask(taskArray); 

    sendNotificationEmail(assignTsk.employeeDataArray[0][1], assignTsk.employeeDataArray[0][7], assignTsk.employeeDataArray[0][6], taskArray[0][2]); 

    const employeeNameValue = strategicTasksSheet.getRange(e.range.getRow(), 7).getValue();
    
    const targetEmployeeSheetId = fun.findTargetEmployeeSpreadSheet(employeeNameValue).employeeSpreadsheetId; 

    const numRow = fun.findRowNumber(targetEmployeeSheetId, fun.tasksSheetName, resources.strategicSS().sheet.getRange(e.range.getRow(), 1).getValue()); 

    const targetedRowOfTask = fun.findtargetTaskRow(targetEmployeeSheetId, fun.tasksSheetName, numRow); 

    const assignTaskPriorityValue = fun.getEventData(e).priority; 

    const targetEmployeeSheet = SpreadsheetApp.openById(targetEmployeeSheetId).getSheetByName('Tasks'); 

    const targetSheetTotalTasksArray = targetEmployeeSheet.getRange(2, 1, targetEmployeeSheet.getLastRow() -1, targetEmployeeSheet.getLastColumn()).getValues(); 

    // fun.reprioritizeTaskRow(taskArray, assignTaskPriorityValue, targetSheetTotalTasksArray, targetEmployeeSheet, numRow - 2 );  // HERE IS THE BUG! SOLVE IT! 

  } 
 
  } else if(headerCellValue === 'Input - Comm' && fun.getEventData(e).taskID > 0){
    Logger.log('Condition Met!'); 


    const employeeNameValue = strategicTasksSheet.getRange(e.range.getRow(), 7).getValue(); 
    const inputCommValue = strategicTasksSheet.getRange(e.range.getRow(), e.range.getColumn()).getValue(); 
    Logger.log(`The input Comm value is ${inputCommValue}`); 
    const iDofTask = strategicTasksSheet.getRange(e.range.getRow(), 1).getValue(); 
    Logger.log(`The id of the task is ${iDofTask}`); 

    toggleBackgroundColorWhite(e); 
    const targetSheetData = findTargetSpreadSheet(employeeNameValue); 
    Logger.log('targetperson data row will com here'); 

    Logger.log(targetSheetData.targetPersonDataRow);// this data is coming correctly 

    copyDataToTargetSpreadsheet(targetSheetData.targetPersonDataRow, inputCommValue, iDofTask); 

    const targetEmployeeSheetId = fun.findTargetEmployeeSpreadSheet(employeeNameValue).employeeSpreadsheetId; 

    const numRow = fun.findRowNumber(targetEmployeeSheetId, fun.tasksSheetName, iDofTask);

    fun.toggleTargetSheetToActiveColor(targetEmployeeSheetId, fun.tasksSheetName, numRow); 

    // const taskSheetinStrategicSS = 

    // const valueOfPriority = resources.strategicSS().sheet.getRange(e.range.getRow(), 6).getValue(); 

    // const arrayOfTotalTasks = resources.strategicSS().sheet.getRange(2, 1, resources.strategicSS().sheet.getLastRow() - 1, resources.strategicSS().sheet.getLastColumn()).getValues(); 

    // const targetEmployeeSheet = SpreadsheetApp.openById(targetEmployeeSheetId).getSheetByName('Tasks'); 

    // const indexRelatedToDataRow = numRow - 2; 


    // fun.reprioritizeTaskRow(taskArray, valueOfPriority, arrayOfTotalTasks, targetEmployeeSheet, indexRelatedToDataRow); 

    


    // sendNotificationEmail(); 

  } else if ( strategicTasksSheet.getRange(1, e.range.getColumn()).getValue() === 'Priority' && fun.getEventData(e).taskID > 0) {


    Logger.log(`the value of taskID for priority is ${fun.getEventData(e).taskID}`); 

    const taskSheet = resources.strategicSS().sheet; 

    const rowOfTask = taskSheet.getRange(e.range.getRow(), 1, 1, taskSheet.getLastColumn()).getValues();
    Logger.log('row of task will come here')
    Logger.log(rowOfTask); 

    const priorityNumber = taskSheet.getRange(e.range.getRow(), e.range.getColumn()).getValue(); 
    Logger.log('proirity number will come here')
    Logger.log(priorityNumber);  

    const arrayOfTotalTasks = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).getValues(); 

    Logger.log('total tasks array will come here')

    Logger.log(arrayOfTotalTasks); 

    const taskId = rowOfTask[0][0];  
    

    const numRow = fun.findRowNumber(resources.strategicId, resources.strategicSS().sheet.getSheetName(), taskId);

    const rowIndex = numRow - 2; 

    Logger.log('Delegate Value will come here '); 
    Logger.log(fun.getEventData(e).subTask); // fun.getEventData is now online ... Congratulations! InshAllah this will be a powerful function with the Will of Allah!

    const eventEmployeeName = fun.getEventData(e).delegateValue; 
      Logger.log(`The checkEmployee function says that the name of the employee is ${eventEmployeeName}`); 

      const eventTaskId = fun.getEventData(e).taskID; 
      Logger.log(`The checkTaskId function says that the name of the employee is ${eventTaskId}`); 

      const eventActiveDataRow = fun.getEventData(e).activeDataRowArray; 
      const eventPriorityValue = fun.getEventData(e).priority; 


    fun.reprioritizeTaskRow(rowOfTask, priorityNumber, arrayOfTotalTasks, taskSheet, rowIndex ); 

    const numRow2 = fun.findRowNumber(resources.strategicId, resources.strategicSS().sheet.getSheetName(), taskId);

    fun.highLightColor(resources.strategicId, resources.strategicSS().sheet.getSheetName(), numRow2); 

        
     

    if(eventEmployeeName !== "Fahad" && eventEmployeeName !== "") {

      

      Logger.log('Last Condition Satisfied!'); 
      Logger.log(`the reason Last Condition has been satisfied is because ${eventEmployeeName}`); 

      const employeeTargetSheetId = fun.findTargetEmployeeSpreadSheet(eventEmployeeName).employeeSpreadsheetId; 

      Logger.log(`The EmployeetargetSheetID function which is a serious problem right now , the value is ${employeeTargetSheetId}`); // correct value is coming ... 

      const employeeTargetSheet = SpreadsheetApp.openById(employeeTargetSheetId).getSheetByName('Tasks'); 
      

      const employeeTotalTasksArray = employeeTargetSheet.getRange(2, 1, employeeTargetSheet.getLastRow() - 1, employeeTargetSheet.getLastColumn()).getValues();  // something is serious wrong here! debug! 

      Logger.log( 'testing purpose employee sheet name will come here '); 
      Logger.log(employeeTargetSheet.getSheetName()); // correct value is coming here! 
      Logger.log('last condition employee total array will come here')
      Logger.log(employeeTotalTasksArray); // there is a serious issue here as well !

      const targetRowNumber = fun.findRowNumber(employeeTargetSheetId, 'Tasks', eventTaskId); 
      Logger.log(`The targetRowNumber which is just for checking is ${targetRowNumber}`); 

      const _employeeTaskRowArray = fun.findtargetTaskRow(employeeTargetSheetId, 'Tasks', eventTaskId); // does this produce a single dimensional or two dimensional array? 
      Logger.log('_TaskRowArray will come here')
      Logger.log(_employeeTaskRowArray); 


      fun.reprioritizeTaskRow(_employeeTaskRowArray, eventPriorityValue, employeeTotalTasksArray, employeeTargetSheet, fun.findRowNumber(employeeTargetSheetId, 'Tasks', eventTaskId ) - 2);

      Logger.log('Looks like the reprioritize function executed successfully! Check! '); 

    }; 

    
    

  } 
   

}
