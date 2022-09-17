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

const sheet = resources.strategicSS().cliSheet; 
const cliRange = sheet.getRange(2, e.range.getColumn());  
const cliRangeValue = cliRange.getValue(); 
const cliColumnRange = sheet.getRange(1, e.range.getColumn()); 
const cliColumnValue = cliColumnRange.getValue(); 
const cliOutputRange = sheet.getRange(2, 2); 
const cliOutputRangeValue = cliOutputRange.getValue(); 
Logger.log(cliColumnValue); 
Logger.log(cliRangeValue); 

const cliValueArray = cliRangeValue.split(' '); 
Logger.log(cliValueArray); 

    if(cliColumnValue == 'Command') {

        Logger.log("Command line interface detected"); 

        if(cliValueArray[3] == "marketresearch" && cliValueArray[4] == "online") {

          Logger.log("Online market research command detected")
          cliOutputRange.clearContent(); 
          cliOutputRange.setValue('Generating Templates...'); 

          const marketResearchTemplate = '1waQ3iU9uXyZLpPCS5f04ZQ2ly8JHnKDDJLV8ZPLwjug'; 
          const projectTrackerTemplate = '12Q8jdMavETZQ0-B-LvrMo7kfruBweVhrAjCdB5pKCtE'; 
          const targetlistTemplate = '1a9srmldCbAfztqKl_pqWaGfN5QA7LzhyRHfJMXebKWo'; 
          const marketingFolder = '1u1F9ACnpd2hqKhV-loX5mSS8-Pf_7-OM'; 

          const newMarketProjectName = 'Online Market Research' + ' ' + cliValueArray[1] + ' ' + cliValueArray[5]
          Logger.log(newMarketProjectName); 

          const newMarketResearchFolder = DriveApp.getFolderById(marketingFolder).createFolder(newMarketProjectName); 
          const newMarketResearchTemplate = DriveApp.getFileById(marketResearchTemplate).makeCopy(`Market Research ${cliValueArray[1]}`, newMarketResearchFolder); 
          const newProjectTrackerTemplate = DriveApp.getFileById(projectTrackerTemplate).makeCopy(`MarketResearch Tracker ${cliValueArray[1]}`, newMarketResearchFolder); 
          const newTargetlistTemplate = DriveApp.getFileById(targetlistTemplate).makeCopy(`Targetlist ${cliValueArray[1]} - ${cliValueArray[5]}`, newMarketResearchFolder); 
          cliOutputRange.setValue('All files and folders generated successfully'); 
          const stratetgicSheetLastRowRange = strategicTasksSheet.getRange(strategicTasksSheet.getLastRow() + 1, 1, 1, 8); 
          const stratetgicSheetLastRowData = stratetgicSheetLastRowRange.getValues(); 
          stratetgicSheetLastRowData[0] = Math.floor(Math.random() * 10000000000); 
          stratetgicSheetLastRowData[1] = new Date(); 
          stratetgicSheetLastRowData[2] = cliValueArray[1]; 
          const newMarketResearchTemplateLink = newMarketResearchTemplate.getUrl(); 
          stratetgicSheetLastRowData[3] = `=HYPERLINK("${newMarketResearchTemplateLink}", "Update Market Research File Here")`; 
          stratetgicSheetLastRowData[4] = '15 mins'; 
          stratetgicSheetLastRowData[5] = '100'; 
          stratetgicSheetLastRowData[6] = 'Fahad'; 
          stratetgicSheetLastRowData[7] = 'Waiting for Update'; 

          Logger.log(stratetgicSheetLastRowData); 

          stratetgicSheetLastRowRange.setValues([stratetgicSheetLastRowData]); 
          cliOutputRange.setValue('Link Data Row Generated Hopefully! ;-'); 

          const operationsSheet = resources.strategicSS().ss.getSheetByName('Operations'); 
          const operationsLastRowRange = operationsSheet.getRange(operationsSheet.getLastRow() + 1, 1, 1, 9); 
          const operationsLastRowArray = operationsLastRowRange.getValues(); 
          operationsLastRowArray[0] = stratetgicSheetLastRowData[0]; 
          operationsLastRowArray[1] = stratetgicSheetLastRowData[1]; 
          operationsLastRowArray[2] = stratetgicSheetLastRowData[2]; 
          operationsLastRowArray[3] = `=HYPERLINK("${newMarketResearchFolder.getUrl()}", "Project Folder")`; 
          operationsLastRowArray[4] = `=HYPERLINK("${newTargetlistTemplate.getUrl()}","TargetList Link")`; 
          operationsLastRowArray[5] = `=HYPERLINK("${newProjectTrackerTemplate.getUrl()}", "Project Tracker Link")`; 
          operationsLastRowArray[6] = 'Marketing'; 
          operationsLastRowArray[7] = 'Online Market Research'; 
          operationsLastRowArray[8] = 'Waiting For Market Research File Update'; 

          operationsLastRowRange.setValues([operationsLastRowArray]); 
          cliOutputRange.setValue('All Data Rows added successfully!'); 

        } else if(cliValueArray[0] == "operations" && cliValueArray[2] == "-n" && cliValueArray[3] == "-lr"){

          Logger.log('Request for new Lab Research operations project creation received!')

          cliOutputRange.clearContent(); 

          const labResearchFolder = '12A9tA5crOip2ATiLRORhgjSNzGms3HQu'; 
          const researchExperimentationTemplate = '1ZrVRDX6KtcLeZ2esbDnS1y8o-hvcWQ7fmmGZcJzTMRY'; 
          const researchPlanningTemplate = '1MxDAJKBCNUorJ3zJTWmc4QmJ_QsQBvPUFnXzHspGLd4'; 
          const labResearchProjectTrackerTemplate = '1G-QtstPq2_5otOZreilKoUd0tKfXBsiuNptvm9pi9IQ'; 

          const newLabResearchProjectName = 'Lab Research Project' + '-' + cliValueArray[1]; 
          Logger.log(newLabResearchProjectName); 

          const newLabResearchFolder = DriveApp.getFolderById(labResearchFolder).createFolder(newLabResearchProjectName); 
          const newResearchExperimentationTemplate = DriveApp.getFileById(researchExperimentationTemplate).makeCopy(`Project Experimentation - ${cliValueArray[1]}`, newLabResearchFolder); 
          const newProjectTrackerTemplate = DriveApp.getFileById(labResearchProjectTrackerTemplate).makeCopy(`Project Tracker - ${cliValueArray[1]}`, newLabResearchFolder); 
          const newProjectPlanningTemplate = DriveApp.getFileById(researchPlanningTemplate).makeCopy(`Project Planning - ${cliValueArray[1]}`, newLabResearchFolder); 
          cliOutputRange.setValue('All files and folders generated successfully'); 
          const stratetgicSheetLastRowRange = strategicTasksSheet.getRange(strategicTasksSheet.getLastRow() + 1, 1, 1, 8); 
          const stratetgicSheetLastRowData = stratetgicSheetLastRowRange.getValues(); 
          stratetgicSheetLastRowData[0] = Math.floor(Math.random() * 10000000000); 
          stratetgicSheetLastRowData[1] = new Date(); 
          stratetgicSheetLastRowData[2] = "Project" + " " + cliValueArray[1]; 
          const projectPlanningTemplateLink = newProjectPlanningTemplate.getUrl(); 
          stratetgicSheetLastRowData[3] = `=HYPERLINK("${projectPlanningTemplateLink}", "Update lab research project plan")`; 
          stratetgicSheetLastRowData[4] = '15 mins'; 
          stratetgicSheetLastRowData[5] = '100'; 
          stratetgicSheetLastRowData[6] = 'Fahad'; 
          stratetgicSheetLastRowData[7] = 'Waiting for Update'; 

          Logger.log(stratetgicSheetLastRowData); 

          stratetgicSheetLastRowRange.setValues([stratetgicSheetLastRowData]); 
          cliOutputRange.setValue('Link Data Row Generated Hopefully! ;-'); 

          const operationsSheet = resources.strategicSS().ss.getSheetByName('Operations'); 
          const operationsLastRowRange = operationsSheet.getRange(operationsSheet.getLastRow() + 1, 1, 1, 9); 
          const operationsLastRowArray = operationsLastRowRange.getValues(); 
          operationsLastRowArray[0] = stratetgicSheetLastRowData[0]; 
          operationsLastRowArray[1] = stratetgicSheetLastRowData[1]; 
          operationsLastRowArray[2] = stratetgicSheetLastRowData[2]; 
          operationsLastRowArray[3] = `=HYPERLINK("${newLabResearchFolder.getUrl()}", "Project Folder")`; 
          operationsLastRowArray[4] = `=HYPERLINK("${newResearchExperimentationTemplate.getUrl()}","Project Experimentation Link")`; 
          operationsLastRowArray[5] = `=HYPERLINK("${newProjectTrackerTemplate.getUrl()}", "Project Tracker Link")`; 
          operationsLastRowArray[6] = 'Operations'; 
          operationsLastRowArray[7] = 'Lab Research'; 
          operationsLastRowArray[8] = 'Waiting For Project Planning Update'; 

          operationsLastRowRange.setValues([operationsLastRowArray]); 
          cliOutputRange.setValue('All Data Rows added successfully!');

        }

        else {


          Logger.log("Invalid command detected"); 
          cliOutputRange.setValue("Invalid Command Detected"); 

        }
    } else if (strategicTasksSheet.getRange(e.range.getRow(), 8).getValue() == "Completed" && strategicTasksSheet.getRange(1, e.range.getColumn()).getValue() == "Status" && strategicTasksSheet.getRange(e.range.getRow(), 4).getValue() == "Update Market Research File Here") {

      Logger.log("target link data row detected successfully!"); 

      // const statusCellValue = strategicTasksSheet.getRange(e.range.getRow(), 8).getValue(); 
      // const subTaskValue = strategicTasksSheet.getRange(e.range.getRow(), 4).getValue(); 
      // Logger.log(`status cell value is ${statusCellValue} and subtask value is ${subTaskValue} and task header value is ${taskHeaderValue}`);

      const marketResearchProjectID = fun.getEventData(e).taskID; 

      Logger.log("the marketresearch project ID is" + " " + marketResearchProjectID); 

      const operationsSheet = resources.strategicSS().operationsSheet; 
      const operationsProjectsRange = operationsSheet.getRange(2, 1, operationsSheet.getLastRow() - 1, 11); 
      const operationsProjectArray = operationsProjectsRange.getValues(); 

      Logger.log('Operations total projects array will come here'); 
      Logger.log(operationsProjectArray);


      const filteredMarketResearchProjectArray = operationsProjectArray.filter( project => {

        return project[0] == marketResearchProjectID; 

      }); 

      Logger.log('filtered projects array will come here'); 
      Logger.log(filteredMarketResearchProjectArray); 
      
      fun.loadBalancer("Operations Executive", fun.getEventData(e).taskID); 

      Logger.log('Load Balancer executed successfully! Alhumdulillah!'); 



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
