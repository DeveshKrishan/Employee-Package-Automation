function main() {
    makeFolder();
    grabEmployeeNames();
  }
  
  // Have to find if folder first exists and if it doesn't, create a foler caled "Employee Packages" 
  
  // find employees and check if the file exists in the folder, else make the files for the employee
  
  // collect data and update each employee package 
  
  // review permissions for each document created
  
  function makeFolder(){
    // https://stackoverflow.com/questions/26118809/how-to-create-a-folder-if-it-doesnt-exist/38024062#38024062
    const EMPLOYEE_FOLDER_NAME = "Employee Packages";
    // Creates a folder with the following name if it does not exist. 
  
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next()
    
    try {
      var newFdr = parentFolder.getFoldersByName(EMPLOYEE_FOLDER_NAME).next();
    }
    catch(e) {
      var newFdr = parentFolder.createFolder(EMPLOYEE_FOLDER_NAME);
    }
  
  }
  
  function grabEmployeeNames(){
    const EMPLOYEE_COLUMN_NAME = "Employee";
    const ROW = 1;
    // Iterate through the "Employees Column"
    const result = getByName(EMPLOYEE_COLUMN_NAME, ROW);
    if (result >= 0){
      
    }
    else{
      throw `A column header at ${ROW} with the name ${EMPLOYEE_COLUMN_NAME} does not exist.`;
    }
  
    // 
  }
  
  function getByName(colName, row) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var col = data[0].indexOf(colName);
    // if (col != -1) {
    //   return data[row-1][col];
    // }
    return col;
  }