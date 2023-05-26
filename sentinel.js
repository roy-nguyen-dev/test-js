const { convertArrayToCSV } = require("convert-array-to-csv");
const converter = require("convert-array-to-csv");
const fs = require("fs");
const util = require("util");
var XLSX = require("xlsx");

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

var wb = XLSX.readFile("./sentinel.xlsx");

var worksheet = wb.Sheets[wb.SheetNames[0]];

// Define the column you want to extract values from (e.g., column A)
var column = "A";

// Find the range of cells in the specified column
var columnRange =
  worksheet[column + "1:" + column + worksheet["!ref"].split(":")[1]];
console.log("columnRange", columnRange);
// Array to store the values in the column
var columnValues = [];

// Iterate through each cell in the column range
columnRange.forEach(function (cell) {
  // Access the value of each cell
  var cellValue = cell.v;

  // Add the value to the columnValues array
  columnValues.push(cellValue);
});

console.log(columnValues);

// try {
//   getContent().then((data) => {
//     if (!data) return;
//     let array = data.split("\n");

//     let temp = [];
//     for (let item of array) {
//       let newItem;
//       const isExistVMWare = item
//         .toLowerCase()
//         .includes(vmwareOriginCharacter.toLowerCase());
//       if (isExistVMWare) {
//         newItem = item
//           .replaceAll(regexVMware1, vmwareReplaceCharacter)
//           .replaceAll(regexVMware71, vmware7ReplaceCharacter)
//           .replaceAll(regexVMware201, vmware20ReplaceCharacter);
//       }
//       let converted = isExistVMWare
//         ? newItem.replaceAll('"', "").split(",")
//         : item.replaceAll('"', "").split(",");

//       let name = converted[0] ?? "";
//       let serialNumber =
//         converted[8] && converted[8].length > 0 ? converted[8] : converted[0];
//       let manufacturer = reconvertText(converted[6]) ?? "";
//       let modelID = reconvertText(converted[7]) ?? "";
//       let os = converted[10] ?? "";
//       let ipAddress = "";
//       let location = converted[5] ?? "";
//       let classType = "";
//       let assignedTo = converted[2] ?? "";
//       let discovery = sccm;
//       let dateOnly = converted[9] ?? "";

//       temp.push([
//         name,
//         serialNumber,
//         manufacturer,
//         modelID,
//         os,
//         ipAddress,
//         location,
//         classType,
//         assignedTo,
//         discovery,
//         dateOnly,
//       ]);
//     }

//     console.log(temp);

//     const csvFromArrayOfArrays = convertArrayToCSV(temp);
//     fs.writeFile("output2.csv", csvFromArrayOfArrays, (error) => {
//       if (error) {
//         console.log(10, error);
//       }
//       console.log("csv file saved successfully.");
//     });
//   });
// } catch (error) {
//   console.log(error);
// }
