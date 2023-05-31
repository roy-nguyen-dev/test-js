// const { convertArrayToCSV } = require("convert-array-to-csv");
const converter = require("convert-array-to-csv");
const fs = require("fs");
const util = require("util");
var XLSX = require("xlsx");
const Blob = require("buffer");

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

var wb = XLSX.readFile("./vmware.csv");

// var worksheet = wb.Sheets[wb.SheetNames];

// Define the column you want to extract values from (e.g., column A)
// var column = "B";
const sheet = wb.Sheets[wb.SheetNames];
const jsonData = XLSX.utils.sheet_to_json(sheet);
const name = "name";
const bios_uid = "bios_uuid";
const guest_os_fullname = "guest_os_fullname";
const ip_address = "ip_address";
const location = "location";
const discoverySource = "VMWare Instance";
const last_discovered = "last_discovered";
// const manufacturer = 'Site';
// const lastReportedIP = 'Last Reported IP';
// const classType = 'classType';
// const activity = 'SentinelOne';
const nameValue = jsonData.map((row) => row[name]);
const biosValue = jsonData.map((row) => row[bios_uid]);
const guestOSValue = jsonData.map((row) => row[guest_os_fullname]);
const ipAddressValue = jsonData.map((row) => row[ip_address]);
const locationValue = jsonData.map((row) => row[location]);
const lastDiscoverValue = jsonData.map((row) => row[last_discovered]);

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

nameValue.forEach(function (value, index) {
  var cellAddress = "A" + (index + 1);
  var discoveryCellAddress = "J" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
  XLSX.utils.sheet_add_aoa(newWorksheet, [[discoverySource]], {
    origin: discoveryCellAddress,
  });
});
biosValue.forEach(function (value, index) {
  var cellAddress = "B" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
guestOSValue.forEach(function (value, index) {
  var cellAddress = "E" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
ipAddressValue.forEach(function (value, index) {
  var cellAddress = "F" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
locationValue.forEach(function (value, index) {
  var cellAddress = "G" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
lastDiscoverValue.forEach(function (value, index) {
  var cellAddress = "K" + (index + 1);
  const dateValue = new Date(value);
  const result =
    dateValue.getMonth() +
    1 +
    "/" +
    dateValue.getDate() +
    "/" +
    dateValue.getFullYear();
  XLSX.utils.sheet_add_aoa(newWorksheet, [[result]], { origin: cellAddress });
});

XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

XLSX.writeFile(newWorkbook, "output6.xlsx");
