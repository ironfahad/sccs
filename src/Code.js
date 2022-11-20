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

          Logger.log('Request for new Lab Research operations project creation received!'); 

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

        } else if(cliValueArray[0] == "operations" && cliValueArray[2] == "-n" && cliValueArray[3] == "-pp"){

          Logger.log('Request for new "Pilot Project" creation received!'); 

          cliOutputRange.clearContent(); 

          const pilotProjectFolder = '11OYWHACViR9wawKTxCDIUcXgEe0CxoC6'; 
          const pilotExperimentationTemplate = '1r0IZznVpvDhQ9otgoZRnnMrDZcnjS_D28_DPTLQuJaw'; 
          const pilotPlanningTemplate = '1gCFVIa5oO2ZbO9fHPpco3bmz6re0ykb7xxhyVYpQ4Mg'; 
          const pilotProjectTrackerTemplate = '1HqKu2zQ72Ab1H1eK6YIe_UwytA6BNILKQxxdjsygy-Q'; 

          const newPilotProjectName = 'Pilot Project' + '-' + cliValueArray[1]; 
          Logger.log(newPilotProjectName); 

          const newPilotProjectFolder = DriveApp.getFolderById(pilotProjectFolder).createFolder(newPilotProjectName); 
          const newPilotExperimentationTemplate = DriveApp.getFileById(pilotExperimentationTemplate).makeCopy(`Project Experimentation - ${cliValueArray[1]}`, newPilotProjectFolder); 
          const newProjectTrackerTemplate = DriveApp.getFileById(pilotProjectTrackerTemplate).makeCopy(`Project Tracker - ${cliValueArray[1]}`, newPilotProjectFolder); 
          const newProjectPlanningTemplate = DriveApp.getFileById(pilotPlanningTemplate).makeCopy(`Project Planning - ${cliValueArray[1]}`, newPilotProjectFolder); 
          cliOutputRange.setValue('All files and folders generated successfully'); 
          const stratetgicSheetLastRowRange = strategicTasksSheet.getRange(strategicTasksSheet.getLastRow() + 1, 1, 1, 8); 
          const stratetgicSheetLastRowData = stratetgicSheetLastRowRange.getValues(); 
          stratetgicSheetLastRowData[0] = Math.floor(Math.random() * 10000000000); 
          stratetgicSheetLastRowData[1] = new Date(); 
          stratetgicSheetLastRowData[2] = "Project" + " " + cliValueArray[1]; 
          const projectPlanningTemplateLink = newProjectPlanningTemplate.getUrl(); 
          stratetgicSheetLastRowData[3] = `=HYPERLINK("${projectPlanningTemplateLink}", "Update pilot project plan")`; 
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
          operationsLastRowArray[3] = `=HYPERLINK("${newPilotProjectFolder.getUrl()}", "Project Folder")`; 
          operationsLastRowArray[4] = `=HYPERLINK("${newPilotExperimentationTemplate.getUrl()}","Project Experimentation Link")`; 
          operationsLastRowArray[5] = `=HYPERLINK("${newProjectTrackerTemplate.getUrl()}", "Project Tracker Link")`; 
          operationsLastRowArray[6] = 'Operations'; 
          operationsLastRowArray[7] = 'Pilot Project'; 
          operationsLastRowArray[8] = 'Waiting For Project Planning Update'; 

          operationsLastRowRange.setValues([operationsLastRowArray]); 
          cliOutputRange.setValue('All Data Rows added successfully!');

        } else if(cliValueArray[3] == "telecom") {

          Logger.log("Telecom marketing campaign command detected")
          cliOutputRange.clearContent(); 
          cliOutputRange.setValue('Generating Templates...');
          SpreadsheetApp.getActive().toast('Command Detected successfully!')

          // New telecom campaign ID that will be used everywhere 

          const newTelecomCampaignID = Math.floor(Math.random() * 10000000000); 

          const telecomExeCampTemplate = '1oKRQP1Ih8qcuxw1fTyzTWBk6Cx6cskZ3KZ_KSsVjQV0';
          // I believe a project tracker template is reasonable essentail but may make the project more or less complex
          const telecomCamPojectTrackerTemplate = '1aVNuaxoR8aD844YnNt2yWAh0P_5Qp0Xkfq9wP4M0zBM'; 
          //target list file either has to be filtered through a criteria or direct file ID needs to be provided. initially lets be simple and get the ID from the command line
          const targetlistFileId = cliValueArray[4]; 
          const insideSalesExeCamTemplate = '1IsSEbdU5fjzbT3pKIBqSVSHiZCL0JEJF1uwPwlnYFeQ';
          const marketingExeCamTemplate = '1TBubILpHU9YQhLR9uK-feV7tE5DIEI4D1PimGFHIf7w'; 
          const telecomCamPlanningTemplate = '1cAxVFDmvrsg8Re4OHfW_pcmhLIVU5dUsOYhFfdb84bc'; 
          const marketingFolderId = '1u1F9ACnpd2hqKhV-loX5mSS8-Pf_7-OM'; 


          const newTelecomCamName = 'Telecom Campaign' + ' ' + cliValueArray[1] + ' ' + cliValueArray[5]
          Logger.log(newTelecomCamName); 

          const newTelecomCamFolder = DriveApp.getFolderById(marketingFolderId).createFolder(newTelecomCamName); 
          // const newTelecomExeCamFile = DriveApp.getFileById(telecomExeCampTemplate).makeCopy(`Telecom Exe ${cliValueArray[1]}`, newTelecomCamFolder); // note that there will be multiple telecom executives. so this needs to be done in an iterator. but lets say make one then iterate over. 
          // const newInsideSalesExeCamFile = DriveApp.getFileById(insideSalesExeCamTemplate).makeCopy(`Insides Sales Exe ${cliValueArray[1]}`, newTelecomCamFolder); 
          // const newMarketingExeCamFile = DriveApp.getFileById(marketingExeCamTemplate).makeCopy(`Marketing Exe ${cliValueArray[1]}`, newTelecomCamFolder);
          const duplicateTargetListFile = DriveApp.getFileById(targetlistFileId).makeCopy('Telecom Duplicate TargetList', newTelecomCamFolder); 


          const newTelecomCamPlanFile = DriveApp.getFileById(telecomCamPlanningTemplate).makeCopy(`${cliValueArray[1]} Telecom Campaign Planning`, newTelecomCamFolder); 

          // Lets iterate over all employees dataset, create, update and share files with them 

          SpreadsheetApp.getActive().toast('Initial Files generated successfully!')

          const HrmSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HRM'); 
          const HrmSheetRange = HrmSheet.getRange(2, 1, HrmSheet.getLastRow() - 1, HrmSheet.getLastColumn()); 
          const HrmDataArray = HrmSheetRange.getValues(); 

          // Here Iterate over all employees, and for each matching employee run several functionals 

          HrmDataArray.forEach( employee => {

            if ( employee[4] == 'Telecom Executive' && employee[5] == 'Active' && employee[9] == 0) {

              // 1. create telecom campaign file 
              const employeeName = employee[1]; 

              const newTelecomCamFile = DriveApp.getFileById(telecomExeCampTemplate).makeCopy(`${cliValueArray[1]}-TE-${employee[1]}`, newTelecomCamFolder); 

              // 2. get employee telecom campaign file id

              const employeeTelecomFileId = newTelecomCamFile.getId(); 

              // 3. get email address of the telecom executive 

              const telecomExecEmail = employee[7]; 

              // 4. share telecom campaign file with the executive
              
              newTelecomCamFile.addEditor(telecomExecEmail); 


              // 5. get telem human resource employee file id 

              const telecomExecutiveOfficialFileId = employee[6]; 

              // 6. create new campaign record in campaigns sheet of the employee

              const telecomEmployeeCampaignSheet = SpreadsheetApp.openById(telecomExecutiveOfficialFileId).getSheetByName('Campaigns'); //sheet name verified 
              const campaignSheetLastRowRange = telecomEmployeeCampaignSheet.getRange(telecomEmployeeCampaignSheet.getLastRow() + 1, 1, 1, telecomEmployeeCampaignSheet.getLastColumn()); 
              const  campaignRowDataArray = campaignSheetLastRowRange.getValues();
              campaignRowDataArray[0] = newTelecomCampaignID; // campaign id column
              Logger.log(`telecom campaign ID check ${campaignRowDataArray[0]}`);
              campaignRowDataArray[1] = new Date(); // date column
              campaignRowDataArray[2] = newTelecomCamName; // campaign name column
              campaignRowDataArray[3] = 'General'; // type column 
              campaignRowDataArray[4] = `=HYPERLINK("${newTelecomCamFile.getUrl()}", "Campaign TargetList File Link")`; // link 1 column
              campaignRowDataArray[5] = 'check'; // link 2 column - here a link to the product knowledge and faqs should come
              campaignRowDataArray[6] = 'check'; // link 3 column - here the link to the telecom script should come
              campaignRowDataArray[7] = 'Waiting For Acceptance'; // status column 

              campaignSheetLastRowRange.setValues([campaignRowDataArray]); 


              

              // first create the telecom employee file. Add the campaign sheet. Add all the necessary columns and then start coding here further

              
              // 7. get duplicate target list file id 

              const duplicateTargetlistFileId = duplicateTargetListFile.getId(); 

              Logger.log(`Duplicate targetlist file ID check ${duplicateTargetlistFileId}`); 

              // 8. copy next 25 records to the designated telecom file

              // Source File 

              const sourceTargetlistSheet = SpreadsheetApp.openById(duplicateTargetlistFileId).getSheetByName('Sheet1'); //verified 
              const sourceTargetListRange = sourceTargetlistSheet.getRange(2, 1, 25, sourceTargetlistSheet.getLastColumn()); 
              const sourceTargetListArray = sourceTargetListRange.getValues(); 
              const sourceTargetListTotalDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow(), sourceTargetlistSheet.getLastColumn()); 

              const remainingTargetListRange = sourceTargetlistSheet.getRange(27, 1, sourceTargetlistSheet.getLastRow() - 25, sourceTargetlistSheet.getLastColumn()); 
              const remainingTargetListDataArray = remainingTargetListRange.getValues(); 
              const updatedTargetListRange = sourceTargetlistSheet.getRange(2, 1, remainingTargetListDataArray.length, remainingTargetListDataArray[0].length); 


              // Target File

              const targetEmployeeTelecomFileSheet = SpreadsheetApp.openById(employeeTelecomFileId).getSheetByName('Sheet1'); 
              const targetEmployeeTelecomSheetRange = targetEmployeeTelecomFileSheet.getRange(2, 1, 25, sourceTargetlistSheet.getLastColumn()); 
              const targetEmployeeTelecomDataArray = targetEmployeeTelecomSheetRange.getValues(); 
              

              targetEmployeeTelecomSheetRange.setValues(sourceTargetListArray);
               

              // 9. delete the same records from the duplicate targetlist file 

              const totaltargetlistDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow() + 1, sourceTargetlistSheet.getLastColumn()); 

              totaltargetlistDataRange.clearContent();

              updatedTargetListRange.setValues(remainingTargetListDataArray);
              

              // 10. Update load value of employee in the HRM sheet
              
              const employeeDataUpdateRange = HrmSheet.getRange(employee[0] + 1, 10, 1, 4);
              const  employeeDataUpdateArray = employeeDataUpdateRange.getValues(); 

              employeeDataUpdateArray[0][0] = 1; 
              employeeDataUpdateArray[0][1] = 1; 
              employeeDataUpdateArray[0][2] = newTelecomCampaignID; 
              employeeDataUpdateArray[0][3] = newTelecomCamName; 
              
              employeeDataUpdateRange.setValues(employeeDataUpdateArray); 

              

              // 10. Send email notification to the designated telecom executive about new campaign project 
              // 11. optional create training file and add the link to the campaign record in the campaign sheet of the employee 

            } else if ( employee[4] == 'Inside Sales Executive' && employee[5] == 'Active' && employee[9] == 0) {

              // 1. create telecom campaign file 
              const employeeName = employee[1]; 

              const newInsideSalesCamFile = DriveApp.getFileById(insideSalesExeCamTemplate).makeCopy(`${cliValueArray[1]}-ISE-${employee[1]}`, newTelecomCamFolder); 

              // 2. get employee telecom campaign file id

              const employeeTelecomFileId = newInsideSalesCamFile.getId(); 

              // 3. get email address of the telecom executive 

              const insideSalesExeEmail = employee[7]; 

              // 4. share telecom campaign file with the executive
              
              newInsideSalesCamFile.addEditor(insideSalesExeEmail);  


              // 5. get telem human resource employee file id 

              const insideSalesExecutiveOfficialFileId = employee[6]; 

              // 6. create new campaign record in campaigns sheet of the employee

              const telecomEmployeeCampaignSheet = SpreadsheetApp.openById(insideSalesExecutiveOfficialFileId).getSheetByName('Campaigns'); //sheet name verified 
              const campaignSheetLastRowRange = telecomEmployeeCampaignSheet.getRange(telecomEmployeeCampaignSheet.getLastRow() + 1, 1, 1, telecomEmployeeCampaignSheet.getLastColumn()); 
              const  campaignRowDataArray = campaignSheetLastRowRange.getValues();
              campaignRowDataArray[0] = newTelecomCampaignID; // campaign id column
              Logger.log(`telecom campaign ID check ${campaignRowDataArray[0]}`);
              campaignRowDataArray[1] = new Date(); // date column
              campaignRowDataArray[2] = newTelecomCamName; // campaign name column
              campaignRowDataArray[3] = 'General'; // type column 
              campaignRowDataArray[4] = `=HYPERLINK("${newInsideSalesCamFile.getUrl()}", "Campaign TargetList File Link")`; // link 1 column
              campaignRowDataArray[5] = 'check'; // link 2 column - here a link to the product knowledge and faqs should come
              campaignRowDataArray[6] = 'check'; // link 3 column - here the link to the telecom script should come
              campaignRowDataArray[7] = 'Waiting For Acceptance'; // status column 

              campaignSheetLastRowRange.setValues([campaignRowDataArray]); 


              

              // first create the telecom employee file. Add the campaign sheet. Add all the necessary columns and then start coding here further

              
              // 7. get duplicate target list file id 

              // const duplicateTargetlistFileId = duplicateTargetListFile.getId(); 

              // Logger.log(`Duplicate targetlist file ID check ${duplicateTargetlistFileId}`); 

              // // 8. copy next 25 records to the designated telecom file

              // // Source File 

              // const sourceTargetlistSheet = SpreadsheetApp.openById(duplicateTargetlistFileId).getSheetByName('Sheet1'); //verified 
              // const sourceTargetListRange = sourceTargetlistSheet.getRange(2, 1, 25, sourceTargetlistSheet.getLastColumn()); 
              // const sourceTargetListArray = sourceTargetListRange.getValues(); 
              // const sourceTargetListTotalDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow(), sourceTargetlistSheet.getLastColumn()); 

              // const remainingTargetListRange = sourceTargetlistSheet.getRange(27, 1, sourceTargetlistSheet.getLastRow() - 25, sourceTargetlistSheet.getLastColumn()); 
              // const remainingTargetListDataArray = remainingTargetListRange.getValues(); 
              // const updatedTargetListRange = sourceTargetlistSheet.getRange(2, 1, remainingTargetListDataArray.length, remainingTargetListDataArray[0].length); 


              // // Target File

              // const targetEmployeeTelecomFileSheet = SpreadsheetApp.openById(employeeTelecomFileId).getSheetByName('Sheet1'); 
              // const targetEmployeeTelecomSheetRange = targetEmployeeTelecomFileSheet.getRange(2, 1, 25, sourceTargetlistSheet.getLastColumn()); 
              // const targetEmployeeTelecomDataArray = targetEmployeeTelecomSheetRange.getValues(); 
              

              // targetEmployeeTelecomSheetRange.setValues(sourceTargetListArray);
               

              // // 9. delete the same records from the duplicate targetlist file 

              // const totaltargetlistDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow() + 1, sourceTargetlistSheet.getLastColumn()); 

              // totaltargetlistDataRange.clearContent();

              // updatedTargetListRange.setValues(remainingTargetListDataArray);
              

              // 10. Update load value of employee in the HRM sheet
              
              const employeeDataUpdateRange = HrmSheet.getRange(employee[0] + 1, 10, 1, 4);
              const  employeeDataUpdateArray = employeeDataUpdateRange.getValues(); 

              employeeDataUpdateArray[0][0] = 1; 
              employeeDataUpdateArray[0][1] = 1; 
              employeeDataUpdateArray[0][2] = newTelecomCampaignID; 
              employeeDataUpdateArray[0][3] = newTelecomCamName; 
              
              employeeDataUpdateRange.setValues(employeeDataUpdateArray); 

              

              // 10. Send email notification to the designated telecom executive about new campaign project 
              // 11. optional create training file and add the link to the campaign record in the campaign sheet of the employee 

            } else if ( employee[4] == 'Marketing Executive' && employee[5] == 'Active' && employee[9] == 0) {

              // 1. create telecom campaign file 
              const employeeName = employee[1]; 

              const newMarketingexeCamFile = DriveApp.getFileById(marketingExeCamTemplate).makeCopy(`${cliValueArray[1]}-ME-${employee[1]}`, newTelecomCamFolder); 

              // 2. get employee telecom campaign file id

              const employeeTelecomFileId = newMarketingexeCamFile.getId();  

              // 3. get email address of the telecom executive 

              const marketingExeEmail = employee[7]; 

              // 4. share telecom campaign file with the executive
              
              newMarketingexeCamFile.addEditor(marketingExeEmail);   


              // 5. get telem human resource employee file id 

              const marketingExecutiveOfficialFileId = employee[6]; 

              // 6. create new campaign record in campaigns sheet of the employee

              const telecomEmployeeCampaignSheet = SpreadsheetApp.openById(marketingExecutiveOfficialFileId).getSheetByName('Campaigns'); //sheet name verified 
              const campaignSheetLastRowRange = telecomEmployeeCampaignSheet.getRange(telecomEmployeeCampaignSheet.getLastRow() + 1, 1, 1, telecomEmployeeCampaignSheet.getLastColumn()); 
              const  campaignRowDataArray = campaignSheetLastRowRange.getValues();
              campaignRowDataArray[0] = newTelecomCampaignID; // campaign id column
              Logger.log(`telecom campaign ID check ${campaignRowDataArray[0]}`);
              campaignRowDataArray[1] = new Date(); // date column
              campaignRowDataArray[2] = newTelecomCamName; // campaign name column
              campaignRowDataArray[3] = 'General'; // type column 
              campaignRowDataArray[4] = `=HYPERLINK("${newMarketingexeCamFile.getUrl()}", "Campaign TargetList File Link")`; // link 1 column
              campaignRowDataArray[5] = 'check'; // link 2 column - here a link to the product knowledge and faqs should come
              campaignRowDataArray[6] = 'check'; // link 3 column - here the link to the telecom script should come
              campaignRowDataArray[7] = 'Waiting For Acceptance'; // status column 

              campaignSheetLastRowRange.setValues([campaignRowDataArray]); 


              

              // first create the telecom employee file. Add the campaign sheet. Add all the necessary columns and then start coding here further

              
              // 7. get duplicate target list file id 

              // const duplicateTargetlistFileId = duplicateTargetListFile.getId(); 

              // Logger.log(`Duplicate targetlist file ID check ${duplicateTargetlistFileId}`); 

              // // 8. copy next 25 records to the designated telecom file

              // // Source File 

              // const sourceTargetlistSheet = SpreadsheetApp.openById(duplicateTargetlistFileId).getSheetByName('Sheet1'); //verified 
              // const sourceTargetListRange = sourceTargetlistSheet.getRange(2, 1, 25, sourceTargetlistSheet.getLastColumn()); 
              // const sourceTargetListArray = sourceTargetListRange.getValues(); 
              // const sourceTargetListTotalDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow(), sourceTargetlistSheet.getLastColumn()); 

              // const remainingTargetListRange = sourceTargetlistSheet.getRange(27, 1, sourceTargetlistSheet.getLastRow() - 25, sourceTargetlistSheet.getLastColumn()); 
              // const remainingTargetListDataArray = remainingTargetListRange.getValues(); 
              // const updatedTargetListRange = sourceTargetlistSheet.getRange(2, 1, remainingTargetListDataArray.length, remainingTargetListDataArray[0].length); 


              // // Target File

              // const targetEmployeeTelecomFileSheet = SpreadsheetApp.openById(employeeTelecomFileId).getSheetByName('Sheet1'); 
              // const targetEmployeeTelecomSheetRange = targetEmployeeTelecomFileSheet.getRange(2, 1, 25, sourceTargetlistSheet.getLastColumn()); 
              // const targetEmployeeTelecomDataArray = targetEmployeeTelecomSheetRange.getValues(); 
              

              // targetEmployeeTelecomSheetRange.setValues(sourceTargetListArray);
               

              // // 9. delete the same records from the duplicate targetlist file 

              // const totaltargetlistDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow() + 1, sourceTargetlistSheet.getLastColumn()); 

              // totaltargetlistDataRange.clearContent();

              // updatedTargetListRange.setValues(remainingTargetListDataArray);
              

              // 10. Update load, max load, campaign id and campaign name values of employee in the HRM sheet
              
              const employeeDataUpdateRange = HrmSheet.getRange(employee[0] + 1, 10, 1, 4);
              const  employeeDataUpdateArray = employeeDataUpdateRange.getValues(); 

              employeeDataUpdateArray[0][0] = 1; 
              employeeDataUpdateArray[0][1] = 1; 
              employeeDataUpdateArray[0][2] = newTelecomCampaignID; 
              employeeDataUpdateArray[0][3] = newTelecomCamName; 
              
              employeeDataUpdateRange.setValues(employeeDataUpdateArray); 


              // 

              

              // 10. Send email notification to the designated telecom executive about new campaign project 
              // 11. optional create training file and add the link to the campaign record in the campaign sheet of the employee 

            } 


          })

          SpreadsheetApp.getActive().toast('Array Iterator ran successfully MashAllah!'); 

          cliOutputRange.setValue('All files and folders generated successfully'); 
          const stratetgicSheetLastRowRange = strategicTasksSheet.getRange(strategicTasksSheet.getLastRow() + 1, 1, 1, 8); 
          const stratetgicSheetLastRowData = stratetgicSheetLastRowRange.getValues(); 
          stratetgicSheetLastRowData[0] = newTelecomCampaignID;  
          stratetgicSheetLastRowData[1] = new Date(); 
          stratetgicSheetLastRowData[2] = cliValueArray[1]; 
          const newTelecomPlanningFileLink = newTelecomCamPlanFile.getUrl(); 
          stratetgicSheetLastRowData[3] = `=HYPERLINK("${newTelecomPlanningFileLink}", "Update Telecom Plan Data")`; 
          stratetgicSheetLastRowData[4] = '1 hr'; 
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
          operationsLastRowArray[3] = `=HYPERLINK("${newTelecomCamPlanFile.getUrl()}", "Telecom Campaign Plan")`; 
          operationsLastRowArray[4] = 'Check'; 
          operationsLastRowArray[5] = 'Check'; 
          operationsLastRowArray[6] = 'Marketing'; 
          operationsLastRowArray[7] = 'Online Market Research'; 
          operationsLastRowArray[8] = 'Waiting For Market Research File Update'; 

          operationsLastRowRange.setValues([operationsLastRowArray]); 

          SpreadsheetApp.getActive().toast('All Data Entries Added successfully!');
          SpreadsheetApp.getActive().toast('Operation Completed Successfully'); 

          cliOutputRange.setValue('All Telecom data Rows added successfully!');

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
