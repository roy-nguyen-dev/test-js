var XLSX = require("xlsx");

var wb = XLSX.readFile("./shoreside.csv");
const sheet = wb.Sheets[wb.SheetNames];
const jsonData = XLSX.utils.sheet_to_json(sheet);
const name = "Name";
const os = "OperatingSystem";
const discoverySource = "Active Directory";
const passwordLastSet = "PasswordLastSet";

function filterOldData(jsonData) {
  const result = jsonData.filter((row) => {
    const currentDate = new Date();
    const dateValue = new Date(row[passwordLastSet]);

    const timeDiff = currentDate.getTime() - dateValue.getTime();

    // Calculate the number of days
    const daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24));

    return daysDiff <= 120;
  });

  return result;
}

const newData = filterOldData(jsonData);

const nameValue = newData.map((row) => row[name]);
const osValue = newData.map((row) => row[os]);
const passwordLastSetValue = newData.map((row) => row[passwordLastSet]);

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
osValue.forEach(function (value, index) {
  var cellAddress = "E" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
passwordLastSetValue.forEach(function (value, index) {
  const cellAddress = "K" + (index + 1);
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

XLSX.writeFile(newWorkbook, "output7.xlsx");
