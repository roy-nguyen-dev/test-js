var XLSX = require("xlsx");

var wb1 = XLSX.readFile("./cmdb_ci_output.xlsx");
var wb2 = XLSX.readFile("./done/ad.xlsx");

const sheet1 = wb1.Sheets[wb1.SheetNames];
const sheet2 = wb2.Sheets[wb2.SheetNames];
const jsonData1 = XLSX.utils.sheet_to_json(sheet1);
const jsonData2 = XLSX.utils.sheet_to_json(sheet2);
const osValue = jsonData1.map((row) => row['os']);

// function findDuplicates(array1, array2) {
//   const duplicates = [];

//   for (let i = 0; i < array1.length; i++) {
//     for (let j = 0; j < array2.length; j++) {
//       if (
//         array1[i].serial_number === array2[j.serial_number] &&
//         !duplicates.includes(array1[i].serial_number)
//       ) {
//         duplicates.push(array1[i]);
//         break;
//       }
//     }
//   }

//   return duplicates;
// }
// const result = findDuplicates(jsonData1, jsonData2);

function findUniqueItems(array) {
  const uniqueItems = array.filter((item, index) => array.indexOf(item) === index);
  return uniqueItems;
}

const result = findUniqueItems(osValue);


const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.aoa_to_sheet([]);

result.forEach(function (value, index) {
  var cellAddress = "A" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});

XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

XLSX.writeFile(newWorkbook, "duplicates.xlsx");
