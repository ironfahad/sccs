const checkSpmLoad = (spmName) => {

 const spmDataRange = teamsSheet.getDataRange().getValues(); 
 const projectManagerName = spmName; 

 const spmArray = spmDataRange.filter( spmRow => {
   return spmRow[1] === spmName; 
 })

 Logger.log(spmArray); 

 const spmLoad = spmArray[0][5]; 

 let loadAvailable = false; 

 if(spmLoad < 10) {
   Logger.log(spmLoad); 
  Logger.log(loadAvailable);
   loadAvailable = true; 
   Logger.log(loadAvailable);
 }

 const spmId = spmArray[0][0]; 
 const spmSpreadSheetId = spmArray[0][6]; 

 return {loadAvailable, spmId, spmSpreadSheetId}; 

}

const assignProject = (projectManagerId, projectManagerSpreadSheetId, projId, projDate, projName, projType, projDesc, projClientName, projContact, projAddress, projStatus) => {


  const projectSpmId = projectManagerId;  
  const projectSpmSpreadSheetId = projectManagerSpreadSheetId;  
  const projectSpmSpreadSheet = SpreadsheetApp.openById(projectSpmSpreadSheetId); 
  const spmProjectsSheet = projectSpmSpreadSheet.getSheetByName('Projects'); 

  const spmNewProjectRange = spmProjectsSheet.getRange(spmProjectsSheet.getLastRow() + 1, 1, 1, 9); 

  const projectId = projId; 
  const projectDate = projDate; 
  const projectName = projName; 
  const projectType = projType; 
  const projectDescription = projDesc;  
  const clientName = projClientName; 
  const clientContact = projContact; 
  const clientAddress = projAddress; 
  const projectStatus = 'Order Confirmed'; 

  const newProjectDataArray = [projectId, projectDate, projectName, projectType, projectDescription, clientName, clientContact, clientAddress, projectStatus]; 

  spmNewProjectRange.setValues([newProjectDataArray]); 

 }

function createLabResearchProject(projectName, clientName, e) {
  Logger.log(salesSheetName); 
  Logger.log(orderSheetFirstCellValue); // now correct values are coming! 

  Logger.log('Creating Lab Research Project...Woomm!'); // so far so good! Alhumdulillah!
  const newProjectRowRange = resources.salesSs().confirmedOrdersSheet.getRange(resources.salesSs().confirmedOrdersSheet.getLastRow() + 1, 1, 1, resources.salesSs().confirmedOrdersSheet.getLastColumn());  

  const lastProjectId = resources.salesSs().confirmedOrdersSheet.getRange(resources.salesSs().confirmedOrdersSheet.getLastRow(), 1).getValue(); 
  Logger.log(lastProjectId); 
  const labProjectName = projectName;
  const projectType = 'Lab Research';
  const projectId = lastProjectId + 1;
  const projectDate = new Date();
  const projectStatus = 'Processed';
  const projectClientName = clientName; 
  // const customerName = 'Not Available'; 
  const customerContact = 'Not Available'; 
  const customerAddress = 'Not Available'; 
  const projectDescription = 'Not Available'; 
  const newProjectValues = [projectId, projectDate, labProjectName, projectType, projectDescription, projectClientName, customerContact, customerAddress, projectStatus]; 

  newProjectRowRange.setValues([newProjectValues]);

  return {projectId, projectDate, projectType, labProjectName, projectDescription, projectClientName, customerContact, customerAddress, projectStatus}; 

};

const copyProjectToOperations = (id, date, projectName, type, description, clientName, contact, address, status ) => {

  const queSheet = resources.cmykOperationsSS().queSheet; 

  const newProjectRowRange = queSheet.getRange(queSheet.getLastRow() + 1, 1, 1, queSheet.getLastColumn()); 
  const projectId = id; 
  const newProjectDate = date;  
  const newProjectName = projectName; 
  const newProjectType = type; 
  const newProjectDescription = description; 
  const newProjectClient = clientName; 
  const newProjectContact = contact; 
  const newProjectAddress = address; 
  const newProjectStatus = status; 

  const newProjectValues = [projectId, newProjectDate, newProjectName, newProjectType, newProjectDescription, newProjectClient, newProjectContact, newProjectAddress, newProjectStatus]; 
  Logger.log(newProjectValues); 
  // Utilities.sleep(2000); 

  newProjectRowRange.setValues([newProjectValues]); // performing operations successfully! Alhumdulillah!
  
}
  
const createNewProject = (strategicCellRange, strategicCellValue, e) => {
  
  // const strategicCellRange = strategicTasksSheet.getRange(e.range.getRow(), e.range.getColumn()); 
  // const strategicCellValue = strategicCellRange.getValue();
  const strategicTaskRange = strategicCellRange; 
  const strategicTaskValue = strategicCellValue; 

  Logger.log('Creating New Project ...'); 

  var indexOfFirstColon = strategicTaskValue.indexOf(':'); 
  var indexOfLastColon = strategicTaskValue.lastIndexOf(':'); // so far so good! 
  Logger.log(indexOfFirstColon); 
  Logger.log(indexOfLastColon); 
  var strategicCellInnerData = strategicTaskValue.slice(indexOfFirstColon + 2, indexOfLastColon - 5); 
  Logger.log(strategicCellInnerData); // Correct Values are received! 
  
  // var strategicCellProjectNameValue = strategicCellValue.slice(indexOfFirstColon + 2, indexOfLastColon -4); 
  var indexOfInnerDataFirstColumn = strategicCellInnerData.indexOf(':'); 
  var indexOfInnerDataLastColumn = strategicCellInnerData.lastIndexOf(':'); 
  var strategicCellProjectNameValue = strategicCellInnerData.slice(0, indexOfInnerDataFirstColumn); 
  Logger.log(strategicCellProjectNameValue); // so far so good! 
  var strategicCellClientNameValue = strategicCellInnerData.slice(indexOfInnerDataFirstColumn + 2, strategicCellInnerData.length +1 ); 
  Logger.log(strategicCellClientNameValue); // so far so good! Alhumdulillah! 
  var strategicCellProjectTypeValue = strategicCellValue.slice(indexOfLastColon -2, indexOfLastColon); 
  Logger.log(strategicCellProjectTypeValue); // so far so good Alhumdulillah!
  var strategicCellSpmValue = strategicTaskValue.slice(indexOfLastColon + 2, strategicTaskValue.length + 1); 
  Logger.log(strategicCellSpmValue); 

  var newLabResearchProject = 'lr';
  var newPilotPlantProject = 'pp'; 
  var newIndustrialProjet = 'ip'; 
  var newWebProject = 'wp'; 
  var newSoftwareProject = 'sp'; 
  var newSoftwareFeatureRequest = 'sf'; 
  var newWebProjectFeatureRequest = 'wf'; 


  if(strategicCellProjectTypeValue === newLabResearchProject) {
   const createLabProject =  createLabResearchProject(strategicCellProjectNameValue, strategicCellClientNameValue);

    copyProjectToOperations(createLabProject.projectId, createLabProject.projectDate, createLabProject.labProjectName, createLabProject.projectType, createLabProject.projectDescription, createLabProject.projectClientName, createLabProject.customerContact, createLabProject.customerAddress, createLabProject.projectStatus); 

    const projectSpmLoad = checkSpmLoad(strategicCellSpmValue); 

    if(projectSpmLoad.loadAvailable) {

      assignProject(projectSpmLoad.spmId, projectSpmLoad.spmSpreadSheetId, createLabProject.projectId, createLabProject.projectDate, createLabProject.labProjectName, createLabProject.projectType, createLabProject.projectDescription, createLabProject.projectClientName, createLabProject.customerContact, createLabProject.customerAddress, createLabProject.projectStatus);

      // updateSpmProjectLoad(); 
      
      // sendSpmNewProjectEmail(); 
 
    } else {

      Logger.log('Load is not Available'); // add functionality here 
    }; 
 

  } else if(strategicCellProjectTypeValue === newPilotPlantProject) {
    createPilotPlantProject(); 
  } else if( strategicCellProjectTypeValue === newIndustrialProjet) {
    createNewIndustrialProject(); 
  } else if ( strategicCellProjectTypeValue === newWebProject) {
    createNewWebProject(); 
  } else if( strategicCellProjectTypeValue === newSoftwareProject) {
    createNewSoftwareProject(); 
  } else if( strategicCellProjectTypeValue === newSoftwareFeatureRequest) {
    createNewSoftwareFeatureRequest();  
  } else if( strategicCellProjectTypeValue === newWebProjectFeatureRequest) {
    createNewWebProjectFeatureRequest(); 
  } else {
    Logger.log("Continue Working Buddy! ...")
  }

  return {strategicCellProjectNameValue, strategicCellProjectTypeValue, strategicCellSpmValue, strategicCellClientNameValue}; 

}; 

function detectAssignment(e) {

  if (fun.getEventData(e).subTask.includes('cp:')) {
const strategicCellRange = resources.strategicSS().sheet.getRange(e.range.getRow(), e.range.getColumn()); 
const strategicCellValue = strategicCellRange.getValue();
Logger.log(strategicCellValue); 
 
  const strategicCellCpValue = strategicCellValue.slice(0,3);

  Logger.log('Strategic Cell CP value will come here!') 
  Logger.log(strategicCellCpValue); 

  if(strategicCellValue.slice(0,3) === 'cp:') {
    createNewProject(strategicCellRange, strategicCellValue, e); 
  }
  const strategicCellArray = strategicCellValue.split(': '); 

}

}

