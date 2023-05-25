const { convertArrayToCSV } = require("convert-array-to-csv");
const converter = require("convert-array-to-csv");
const fs = require("fs");
const util = require("util");

let result = [];
let unknown = "N/A";

const readFile = util.promisify(fs.readFile);

function getContent() {
  return readFile("./cmdb_ci1.csv", 'utf-8');
}

try {
  getContent().then((data) => {
    if (!data) return;
    let array = data.split("\n");

    let temp = [];
    for (let item of array) {
      let converted = item.split(",");
      let name = converted[0] ?? "";
      let serialNumber = converted[1] ?? "";
      let manufacturer = converted[2] ?? "";
      let modelID = converted[3] ?? "";
      let os = unknown;
      let ipAddress = converted[4] ?? "";
      let location = converted[6] ?? "";
      let classType = converted[8] ?? "";
      let assignedTo = unknown;
      let discovery = converted[7] ?? "";
      let dateOnly = converted[9]?.split(" ")?.[0] ?? "";

      temp.push([
        name,
        serialNumber,
        manufacturer,
        modelID,
        os,
        ipAddress,
        location,
        classType,
        assignedTo,
        discovery,
        dateOnly,
      ]);
    }

    console.log(temp);

    const csvFromArrayOfArrays = convertArrayToCSV(temp);
    fs.writeFile("output1.csv", csvFromArrayOfArrays, (error) => {
      if (error) {
        console.log(10, error);
      }
      console.log("csv file saved successfully.");
    });
  });
} catch (error) {
  console.log(error);
}

// fs.readFile("./cmdb_ci_text.txt", "utf8", async (err, data) => {
//   if (err) {
//     console.error(err);
//     return;
//   }
//   if (data) {
//     result.push(data);
//   }
// });

// console.log("result", result);

// fetch("/cmdb_ci_text.txt")
//   .then((response) => response.text())
//   .then((text) => console.log(text));
