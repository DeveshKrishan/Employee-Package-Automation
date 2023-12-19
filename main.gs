function main() {
  const folder_location = makeFolder();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getDataRange().getValues(); // 2D Array of Values
  const employeeMap = findEmployeeNamesAndHub(sheet);
  const dataToEmployee = mapDatatoEmployee(sheet);
  createFile(employeeMap, dataToEmployee, folder_location);

}

function makeFolder(){
    /** 
   * Creates a folder that will contain the generated reports for each employee.
   */
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

function findQuantityFromNotes(note){
  /** 
   * If there is text containing "Quantity" in the Notes Column, it parses through to find the amount.
   * For example, finding "Quantity (x2)" will result in "2".
   */  
  if (note.includes("Quantity")){
    let noteIndex = note.indexOf("Quantity");
    if (noteIndex != -1){
      note = note.slice(noteIndex + 9)
      let parenthesisStartIndex = note.indexOf("(");
      let parenthesisEndIndex = note.indexOf(")");
      note = note.slice(parenthesisStartIndex, parenthesisEndIndex);
      note = note.replace("(", "");
      note = note.replace(")", "");
      note = note.replace("x", "");

      let quantityFromNotes = (parseInt(note));
      return quantityFromNotes;
    }
  }

  return -1; //ERROR

}

function createHorizontalAllignments(a, b){
  /** 
   * Creates a 2D array representing the horizontal allignment values to set in the sheet.
   */
    let Array2D = (r,c) => [...Array(r)].map(_=>Array(c).fill(0));

    let horizontalAlignments = Array2D(a, b); // + 1 represents the total row being included

    let newArray = horizontalAlignments.map(number => [ "left", "center", "center", "center", "right"]);

    return newArray;
}

function setUpEmployeeReturnTable(employeeSheet, employeeReturnRow){
  /** 
   * Sets up the Column Names for the Employee Return Table.
   */
  employeeSheet[employeeReturnRow - 1][1] = "Equipment returned by employee:";
  employeeSheet[employeeReturnRow][1] = "Item";
  employeeSheet[employeeReturnRow][2] = "Quantity";
  employeeSheet[employeeReturnRow][3] = "Value";
  employeeSheet[employeeReturnRow][4] = "Description / Number";
  employeeSheet[employeeReturnRow][5] = "Date Returned";
  return employeeSheet;
}

function formatLoanTable(loanTable, itemRowStart){
  /** 
   * Sets up the formatting for the Employee Loan Table.
   */
  loanTable.setBorder(true, true, true, true, true, true);
  loanTable.setBackground("white");
  loanTable.setHorizontalAlignments(createHorizontalAllignments(itemRowStart - 5 + 1,5));
}

function formatReturnTable(returnTable, returnRowCounter, employeeReturnRow){
  /** 
   * Sets up the formatting for the Employee Return Table.
   */
  returnTable.setBorder(true, true, true, true, true, true);
  returnTable.setBackground("white");
  returnTable.setHorizontalAlignments(createHorizontalAllignments(returnRowCounter - employeeReturnRow - 1, 5));
}

function setReturnTableTitleRowHeight(ss, employeeReturnRow){
  /** 
   * Sets up the row height for the Employee Return Table Titles.
   */
  // https://stackoverflow.com/questions/58620929/apps-script-to-force-specific-row-height-in-google-sheets
  // row height bug where it seems to only change  2 rows instead of 3?
  // trying to find why it doenst change the ERR + 1 row to 16 pixels and keeps it standard
  // can experiment with copy format and if last resort, manually format to copy the draft

  ss.setRowHeights(employeeReturnRow, 2, 30)
  ss.setRowHeight((employeeReturnRow), 38);
}


function createFile(employeeMap, dataToEmployee, folder_location){
  /** 
   * Creates a file for each employee setting up formatting and values.
   */
  const file_array = [];
  for (const [employee, hub] of employeeMap) {
    // https://stackoverflow.com/questions/43566567/add-file-to-an-existing-folder-if-file-doesnt-exist-in-that-folder\
    var file = DriveApp.getFileById(PropertiesService.getScriptProperties().getProperty('TEMPLATE_FILE_ID'))

    let name = `[Draft-Automation] ${employee} | ${hub} Equipment Inventory`; // names the file for each employee
    if (!folder_location.getFilesByName(name).hasNext()) { //if this file does not exist
      let fileName = file.makeCopy(name, folder_location);   
      let documentId = fileName.getId();

      var ss = SpreadsheetApp.openById(documentId).getSheets()[0];
      let employeeSheet = ss.getDataRange().getValues(); // 2D Array of Values
      let range = ss.getDataRange();

      employeeSheet[0][1] = `Equipment Package | ${employee} [UCINetID: ]`; //set title of employee report

      const arr = dataToEmployee.get(employee); //grab 2D array of items for each employee
      let itemRowStart = 5;
      let employeeSum = 0;

      for (let i = 0; i < arr.length;i++){
        let item = `${arr[i][0]} ${arr[i][2]} ${arr[i][3]}`, quantity = 1, value = arr[i][5];
        employeeSum += (parseInt(value));
        let serial_num = arr[i][1], date_install = arr[i][4], note = arr[i][7];

        let quantityFromNotes = findQuantityFromNotes(note);
        
        if (quantityFromNotes > 1){
          quantity = quantityFromNotes;
        }

        try{
          if (date_install != ""){
            date_install = date_install.toLocaleDateString();
          }
        }
        catch(e){

        }


        employeeSheet[itemRowStart][1] = item;
        employeeSheet[itemRowStart][2] = quantity;
        employeeSheet[itemRowStart][3] = value;
        employeeSheet[itemRowStart][4] = serial_num;
        employeeSheet[itemRowStart][5] = date_install;
        itemRowStart++;
      }


      totalRow = 6 + arr.length - 1;
      employeeSheet[totalRow][1] = "TOTAL VALUE";
      employeeSheet[totalRow][3] = employeeSum;

      loanTable = ss.getRange(6, 2, itemRowStart - 5 + 1, 5);
      formatLoanTable(loanTable, itemRowStart);

      ss.getRange(6, 4, itemRowStart - 5 + 1, 1).setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');
      let colorArr = [["#b7b7b7", "#b7b7b7", "#b7b7b7", "#b7b7b7", "#b7b7b7"]];
      ss.getRange(totalRow + 1, 2, 1, 5).setBackgrounds(colorArr).setFontWeight("bold");

      let employeeReturnRow = totalRow + 1 + 10;

      employeeSheet = setUpEmployeeReturnTable(employeeSheet, employeeReturnRow);

      range.setValues(employeeSheet)

      // https://developers.google.com/apps-script/reference/spreadsheet/range#copyformattorangesheet,-column,-columnend,-row,-rowend
      // https://www.sppatra.com/2022/08/how-to-copy-range-with-and-without.html
      // set up return table by copying format of loan table
      let employeeLoanRange = ss.getRange(4, 2, 2, 5);
      employeeLoanRange.copyFormatToRange(ss, 2, 7, employeeReturnRow, employeeReturnRow + 2);
      setReturnTableTitleRowHeight(ss, employeeReturnRow)
      employeeSheet = ss.getDataRange().getValues(); // 2D Array of Values

      let returnRowCounter = employeeReturnRow + 1;

      for (let i = 0; i < arr.length;i++){
        let item = `${arr[i][0]} ${arr[i][2]} ${arr[i][3]}`, quantity = 1, value = arr[i][5], serial_num = arr[i][1];
        let date_install = arr[i][4], returnDate = arr[i][6], note = arr[i][7];
        let quantityFromNotes = findQuantityFromNotes(note);

        if (quantityFromNotes > 1){
          quantity = quantityFromNotes;
        }

        if (returnDate != "" && new Date(returnDate) < new Date()){
          employeeSheet[returnRowCounter][1] = item;
          employeeSheet[returnRowCounter][2] = quantity;
          employeeSheet[returnRowCounter][3] = value;
          employeeSheet[returnRowCounter][4] = serial_num;
          employeeSheet[returnRowCounter][5] = returnDate.toLocaleDateString();
          returnRowCounter++;
        }
      }

      // covers case when there is 0 rows for return table
      try{
        returnTable = ss.getRange(employeeReturnRow + 2, 2, returnRowCounter - employeeReturnRow - 1, 5);
        formatReturnTable(returnTable, returnRowCounter, employeeReturnRow);
      }
      catch(e){
      }

      ss.getRange(employeeReturnRow +2, 4, itemRowStart - 5 + 1, 1).setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');

      range.setValues(employeeSheet)

      file_array.push(fileName);
    }
  }
  return file_array

}

function mapDatatoEmployee(sheet){
  /** 
   * Generates a map mapping each employee to each item containing its' Manufacturer, Serial Number, Equipment Type, Model, Install Date, Value of Product, and Note regarding the item. 
   */

  const labels = sheet[0];
  // if any column names change for any of these, update them to the correct one.
  const EMPLOYEE_COLUMN_NAME = "Employee";
  const MANUFACTURER_NAME = "Manufacturer";
  const SERIAL_NUMBER_NAME = "Serial Number";
  const EQUIPMENT_TYPE_NAME = "Equipment Type";
  const MODEL_NAME = "Model";
  const DATE_OF_FIRST_INSTALL_NAME = "Date of First Install";
  const VALUE_NAME  = "Value";
  const EMPLOYEE_RETURN_DATE = "Employee Return Date";
  const NOTES = "Notes";

  let employee_index = labels.indexOf(EMPLOYEE_COLUMN_NAME);
  let manufacturer_index = labels.indexOf(MANUFACTURER_NAME);
  let serial_number_index = labels.indexOf(SERIAL_NUMBER_NAME);
  let equipment_type_index = labels.indexOf(EQUIPMENT_TYPE_NAME);
  let model_name_index = labels.indexOf(MODEL_NAME);
  let date_of_first_install_index = labels.indexOf(DATE_OF_FIRST_INSTALL_NAME);
  let value_index = labels.indexOf(VALUE_NAME);
  let return_date_index = labels.indexOf(EMPLOYEE_RETURN_DATE);
  let notes_index = labels.indexOf(NOTES);

  // NOTE: Do not have to check for Employee Index or Hub Index as that was previously checked.
  if (manufacturer_index == -1){
    throw new Error(`The program could not find the COLUMN: ${MANUFACTURER_NAME} in the first row.`);
  }
  else if (serial_number_index == -1){
    throw new Error(`The program could not find the COLUMN: ${SERIAL_NUMBER_NAME} in the first row.`);
  }
  else if (equipment_type_index == -1){
    throw new Error(`The program could not find the COLUMN: ${EQUIPMENT_TYPE_NAME} in the first row.`);
  }
  else if (model_name_index == -1){
    throw new Error(`The program could not find the COLUMN: ${MODEL_NAME} in the first row.`);
  }
  else if (date_of_first_install_index == -1){
    throw new Error(`The program could not find the COLUMN: ${DATE_OF_FIRST_INSTALL_NAME} in the first row.`);
  }
  else if (value_index == -1){
    throw new Error(`The program could not find the COLUMN: ${VALUE_NAME} in the first row.`);
  }
  else if (return_date_index == -1){
    throw new Error(`The program could not find the COLUMN: ${EMPLOYEE_RETURN_DATE} in the first row.`);
  }
  else if (notes_index == -1){
    throw new Error(`The program could not find the COLUMN: ${NOTES} in the first row.`);
  }

  let myMap = new Map();

  for (const row of sheet){
    let employeeName = row[employee_index];
    let manufacturer = row[manufacturer_index];
    let equipment_type = row[equipment_type_index];
    let model = row[model_name_index];
    let date_install = row[date_of_first_install_index];
    let value = row[value_index];
    let serial_num = row[serial_number_index];
    let return_date = row[return_date_index];
    let note = row[notes_index];

    if (myMap.has(employeeName)){ // if an employee key has already been created
      let temp_arr = myMap.get(employeeName);
      temp_arr.push([manufacturer, serial_num, equipment_type, model, date_install, value, return_date, note]);
      myMap.set(employeeName, temp_arr)
    }
    else{
      myMap.set(employeeName, [[manufacturer, serial_num, equipment_type, model, date_install, value, return_date, note]]);
    }
  }

  return myMap;
}



function findEmployeeNamesAndHub(sheet){
  /** 
   * Given a sheet, grabs all the names and hub of each employee that is present in the Equipment Inventory Sheet.
   */
  // https://dev.to/stephencweiss/write-your-own-javascript-contracts-and-docstrings-42ho
  const labels = sheet[0]; //grab first row of sheet
  // You can change these constants to search if the Column Names are changed anytime
  const EMPLOYEE_COLUMN_NAME = "Employee";
  const HUB_COLUMN_NAME = "Hub";


  let hub_index = labels.indexOf(HUB_COLUMN_NAME);
  let employee_index = labels.indexOf(EMPLOYEE_COLUMN_NAME);

  if (hub_index == -1){
    throw new Error(`The program could not find the COLUMN: ${HUB_COLUMN_NAME} in the first row.`);
  }
  if (employee_index == -1){
    throw new Error(`The program could not find the COLUMN: ${EMPLOYEE_COLUMN_NAME} in the first row.`);
  }

  let myMap = new Map();

  for (const row of sheet){
    let hub = row[hub_index], employeeName = row[employee_index];
    if (myMap.has(employeeName)){
      continue;
    }
    else{
      myMap.set(employeeName, hub);
    }
  }
  return myMap;
}

