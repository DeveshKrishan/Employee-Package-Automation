function main() {
  const folder_location = makeFolder();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getDataRange().getValues(); // 2D Array of Values
  const employeeMap = findEmployeeNamesAndHub(sheet);
  const dataToEmployee = mapDatatoEmployee(sheet);
  




  const file = createFile(employeeMap, dataToEmployee, folder_location);
  // console.log(file);

  
  // console.log(file);
  // console.log(LogDoc.getBody());


  // set_Employee_Names = grabEmployeeNames();

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
  return newFdr;

}

function createFile(employeeMap, dataToEmployee, folder_location){
  const file_array = [];
  for (const [employee, hub] of employeeMap) {
    // https://stackoverflow.com/questions/43566567/add-file-to-an-existing-folder-if-file-doesnt-exist-in-that-folder\


    var file = DriveApp.getFileById(PropertiesService.getScriptProperties().getProperty('TEMPLATE_FILE_ID'))

    // console.log(folder_location);


    let name = `[Draft-Automation] ${employee} | ${hub} Equipment Inventory`;
    if (!folder_location.getFilesByName(name).hasNext()) {
      let file_name = file.makeCopy(name, folder_location);   
      documentId = file_name.getId();

      var ss = SpreadsheetApp.openById(documentId).getSheets()[0];
      let employeeSheet = ss.getDataRange().getValues(); // 2D Array of Values
      var range = ss.getDataRange();

      employeeSheet[0][1] = `Equipment Package | ${employee} [UCINetID: staff]`;

      const arr = dataToEmployee.get(employee);
      console.log(arr)
      let itemRowStart = 5;
// manufacturer, serial_num, equipment_type, model, date_install, value
      for (let i = 0; i < arr.length;i++){
        let item = `${arr[i][0]} ${arr[i][2]} ${arr[i][3]}`;
        let quantity = 1;
        let value = arr[i][5];
        let serial_num = arr[i][1];
        let date_install = arr[i][4];
        if (date_install != ""){
          date_install = date_install.toLocaleDateString();
        }

        employeeSheet[itemRowStart][1] = item;
        employeeSheet[itemRowStart][2] = quantity;
        employeeSheet[itemRowStart][3] = value;
        employeeSheet[itemRowStart][4] = serial_num;
        //     Thu Feb 23 2023 00:00:00 GMT-0800 (Pacific Standard Time),
        employeeSheet[itemRowStart][5] = date_install;
        itemRowStart++;
      }

      loanTable = ss.getRange(6, 2, itemRowStart - 5, 5);
      loanTable.setBorder(true, true, true, true, true, true);
      loanTable.setBackground("white");

      let Array2D = (r,c) => [...Array(r)].map(_=>Array(c).fill(0));

      let horizontalAlignments = Array2D(itemRowStart - 5,5);

      let newArray = horizontalAlignments.map(number => [ "left", "center", "center", "left", "right"]);

      loanTable.setHorizontalAlignments(newArray);
      ss.getRange(6, 4, itemRowStart - 5, 1).setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');
      range.setValues(employeeSheet)
      file_array.push(file_name);
    }
    else{
      // for testing
    }
  }
  return file_array



}

function mapDatatoEmployee(sheet){
  // grab Employee, Manufactuer, Equipment Type, Model, Serial NUmber, Date of First Install, Value

  const labels = sheet[0];

  // if any column names change for any of these, update them to the correct one.

  const EMPLOYEE_COLUMN_NAME = "Employee";
  const MANUFACTURER_NAME = "Manufacturer";
  const SERIAL_NUMBER_NAME = "Serial Number";
  const EQUIPMENT_TYPE_NAME = "Equipment Type";
  const MODEL_NAME = "Model";
  const DATE_OF_FIRST_INSTALL_NAME = "Date of First Install";
  const VALUE_NAME  = "Value";

  let employee_index = labels.indexOf(EMPLOYEE_COLUMN_NAME);
  let manufacturer_index = labels.indexOf(MANUFACTURER_NAME);
  let serial_number_index = labels.indexOf(SERIAL_NUMBER_NAME);
  let equipment_type_index = labels.indexOf(EQUIPMENT_TYPE_NAME);
  let model_name_index = labels.indexOf(MODEL_NAME);
  let date_of_first_install_index = labels.indexOf(DATE_OF_FIRST_INSTALL_NAME);
  let value_index = labels.indexOf(VALUE_NAME);



  // console.log(labels);
  var myMap = new Map();
  for (const row of sheet){
    let employeeName = row[employee_index];
    let manufacturer = row[manufacturer_index];
    let equipment_type = row[equipment_type_index];
    let model = row[model_name_index];
    let date_install = row[date_of_first_install_index];
    let value = row[value_index];
    let serial_num = row[serial_number_index];

    // console.log(employeeName, manufacturer, serial_num, equipment_type, model, date_install, value);

    if (myMap.has(employeeName)){
      let temp_arr = myMap.get(employeeName);
      temp_arr.push([manufacturer, serial_num, equipment_type, model, date_install, value]);
      // console.log(temp_arr);
      myMap.set(employeeName, temp_arr)
    }
    else{
      myMap.set(employeeName, [[manufacturer, serial_num, equipment_type, model, date_install, value]]);

    }

  }


  return myMap;
}



function findEmployeeNamesAndHub(sheet){
  const labels = sheet[0];
  const EMPLOYEE_COLUMN_NAME = "Employee";
  const HUB_COLUMN_NAME = "Hub";

  let hub_index = labels.indexOf(HUB_COLUMN_NAME);
  let employee_index = labels.indexOf(EMPLOYEE_COLUMN_NAME);
  var myMap = new Map();

  for (const row of sheet){
    let hub = row[hub_index];
    let employeeName = row[employee_index];
    if (myMap.has(employeeName)){
      continue;
    }
    else{
      myMap.set(employeeName, hub);

    }
  }

// for (const [key, value] of myMap) {
//   console.log(`${key} = ${value}`);
// }

  // Iterate through the 2d array with the given indexes for column
  // Step 1: Grab all the employees and map each STAFF NAME to HUB
  // console.log(hub_index, employee_index, EMPLOYEE_COLUMN_NAME, HUB_COLUMN_NAME);

  return myMap;
}



