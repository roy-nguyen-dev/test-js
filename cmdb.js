var XLSX = require("xlsx");

var wb = XLSX.readFile("./cmdb_ci.csv");
const sheet = wb.Sheets[wb.SheetNames];
const jsonData = XLSX.utils.sheet_to_json(sheet);
const name = "name";
const serialNumber = "serial_number";
const manufacturer = "manufacturer";
const modelID = "model_id";
const os = "os";
const ip = "ip_address";
const location = "location";
const classType = "sys_class_name";
const discovery = "discovery_source";
const update = "sys_updated_on";

const nameValue = jsonData.map((row) => row[name]);
const serialValue = jsonData.map((row) => row[serialNumber]);
const manufacturerValue = jsonData.map((row) => row[manufacturer]);
const modelIDValue = jsonData.map((row) => row[modelID]);
const osValue = jsonData.map((row) => row[os]);
const ipValue = jsonData.map((row) => row[ip]);
const locationValue = jsonData.map((row) => row[location]);
const classTypeValue = jsonData.map((row) => row[classType]);
const discoveryValue = jsonData.map((row) => row[discovery]);
const updateValue = jsonData.map((row) => row[update]);

const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.aoa_to_sheet([]);

nameValue.forEach(function (value, index) {
  var cellAddress = "A" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
serialValue.forEach(function (value, index) {
  var cellAddress = "B" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
manufacturerValue.forEach(function (value, index) {
  var cellAddress = "C" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
modelIDValue.forEach(function (value, index) {
  var cellAddress = "D" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
osValue.forEach(function (value, index) {
  var cellAddress = "E" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
ipValue.forEach(function (value, index) {
  var cellAddress = "F" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
locationValue.forEach(function (value, index) {
  var cellAddress = "G" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
classTypeValue.forEach(function (value, index) {
  var cellAddress = "H" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
discoveryValue.forEach(function (value, index) {
  var cellAddress = "J" + (index + 1);
  XLSX.utils.sheet_add_aoa(newWorksheet, [[value]], { origin: cellAddress });
});
updateValue.forEach(function (value, index) {
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

XLSX.writeFile(newWorkbook, "cmdb_ci_output.xlsx");
