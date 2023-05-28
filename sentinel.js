// const { convertArrayToCSV } = require("convert-array-to-csv");
const converter = require("convert-array-to-csv");
const fs = require("fs");
const util = require("util");
var XLSX = require("xlsx");
const Blob = require('buffer');

let result = [];
let sccm = "SCCM";
let vmwareOriginCharacter = "VMware, Inc.";
let vmwareReplaceCharacter = "VMware. Inc.";
let vmware7OriginCharacter = "VMware7,1";
let vmware7ReplaceCharacter = "VMware7.1";
let vmware20ReplaceCharacter = "VMware7.1";
let vmware20OriginCharacter = "VMware7.1";
let regexVMware1 = /VMware, Inc./g;
let regexVMware2 = /VMware. Inc./g;
let regexVMware71 = /VMware7,1/g;
let regexVMware72 = /VMware7.1/g;
let regexVMware201 = /VMware20,1/g;
let regexVMware202 = /VMware20.1/g;

// const readFile = util.promisify(fs.readFile);

// function getContent() {
//   return readFile("./SCCM_0523.csv", "utf-8");
// }

function reconvertText(item) {
  if (!item || item.length == 0) return "";
  if (item.toLowerCase().includes(vmwareReplaceCharacter.toLowerCase())) {
    const newItem = item.replaceAll(regexVMware2, vmwareOriginCharacter);
    return newItem;
  } else if (
    item.toLowerCase().includes(vmware7ReplaceCharacter.toLowerCase())
  ) {
    const newItem = item.replaceAll(regexVMware72, vmware7OriginCharacter);
    return newItem;
  } else if (
    item.toLowerCase().includes(vmware20ReplaceCharacter.toLowerCase())
  ) {
    const newItem = item.replaceAll(regexVMware202, vmware20OriginCharacter);
    return newItem;
  } else {
    return item;
  }
}

var wb = XLSX.readFile("./sentinel_report.xlsx");

// var worksheet = wb.Sheets[wb.SheetNames];

// Define the column you want to extract values from (e.g., column A)
// var column = "B";
const sheet = wb.Sheets[wb.SheetNames];
const jsonData = XLSX.utils.sheet_to_json(sheet);
const endpointName = 'Endpoint Name';
const serialNumber = 'Serial Number';
// const manufacturer = 'Site';
const modelName = 'Model Name';
const os = 'OS';
const lastReportedIP = 'Last Reported IP';
// const location = 'location';
// const classType = 'classType';
const lastLoggedInUser = 'Last Logged In User';
const discoverySource = 'SentinelOne';
// const activity = 'SentinelOne';
const endPointValue = jsonData.map(row => row[endpointName]);
const serialValue = jsonData.map(row => row[serialNumber]);
const modelValue = jsonData.map(row => row[modelName]);
const osValue = jsonData.map(row => row[os]);
const lastReportedIPValue = jsonData.map(row => row[lastReportedIP]);
const lastLoggedInValue = jsonData.map(row => row[lastLoggedInUser]);



// Find the range of cells in the specified column
// var columnRange = XLSX.utils.decode_range(worksheet['!ref']);
// var columnValues = [];
// var column1Values = [];

// for (var rowNum = columnRange.s.r; rowNum <= columnRange.e.r; rowNum++) {
//   var cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: columnRange.s.c });
//   var cell = worksheet[cellAddress];

//   // Access the value of each cell
//   var cellValue = (cell && cell.v) || '';

//   // Add the value to the columnValues array
//   columnValues.push(cellValue);
// }



const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.aoa_to_sheet([]);

endPointValue.forEach(function (value, index) {
  var cellAddress = 'A' + (index + 1);
  var discoveryCellAddress = 'J' + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
  XLSX.utils.sheet_add_aoa(newWorksheet, [[discoverySource]], { origin: discoveryCellAddress });
});
serialValue.forEach(function (value, index) {
  var cellAddress = 'B' + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
modelValue.forEach(function (value, index) {
  var cellAddress = 'D' + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
osValue.forEach(function (value, index) {
  var cellAddress = 'E' + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
lastReportedIPValue.forEach(function (value, index) {
  var cellAddress = 'F' + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
lastLoggedInValue.forEach(function (value, index) {
  var cellAddress = 'I' + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});

XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

XLSX.writeFile(newWorkbook, 'output5.xlsx');



// var workbook = XLSX.utils.book_new();

// // Create a new worksheet
// var newWorksheet = XLSX.utils.aoa_to_sheet([]);

// columnValues.forEach(function (value, index) {
//   var cellAddress = column + (index + 1);
//   XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
// });

// // Add the worksheet to the workbook
// XLSX.utils.book_append_sheet(workbook, newWorksheet, 'Sheet 1');

// // Convert the workbook to an array buffer
// var excelData = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });

// // Save the Excel file
// function saveExcelFile(data, filename) {
//   var buffer = Buffer.from(data);
//   fs.writeFileSync(filename, buffer);
// }

// // Save the Excel file with a specific filename
// saveExcelFile(excelData, 'output4.xlsx');

