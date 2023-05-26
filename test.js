const { convertArrayToCSV } = require("convert-array-to-csv");
const converter = require("convert-array-to-csv");
const fs = require("fs");
const util = require("util");

let result = [];
let unknown = "N/A";
let vmwareOriginCharacter = "VMware, Inc.";
let vmwareReplaceCharacter = "VMware. Inc.";
let vmware7OriginCharacter = "VMware7,1";
let vmware7ReplaceCharacter = "VMware7.1";
let regexVMware1 = /VMware, Inc./g;
let regexVMware2 = /VMware. Inc./g;
let regexVMware71 = /VMware7,1/g;
let regexVMware72 = /VMware7.1/g;

const readFile = util.promisify(fs.readFile);

function getContent() {
  return readFile("./cmdb_ci.csv", "utf-8");
}

function reconvertText(item) {
  if (!item || item.length == 0) return "";
  if (item.toLowerCase().includes(vmwareReplaceCharacter.toLowerCase())) {
    const newItem = item
      .replaceAll(regexVMware2, vmwareOriginCharacter)
      .replaceAll(regexVMware72, vmware7OriginCharacter);
    return newItem;
  } else {
    return item;
  }
}

try {
  getContent().then((data) => {
    if (!data) return;
    let array = data.split("\n");

    let temp = [];
    for (let item of array) {
      let newItem;
      const isExistVMWare = item
        .toLowerCase()
        .includes(vmwareOriginCharacter.toLowerCase());
      if (isExistVMWare) {
        newItem = item
          .replaceAll(regexVMware1, vmwareReplaceCharacter)
          .replaceAll(regexVMware71, vmware7ReplaceCharacter);
      }
      let converted = isExistVMWare
        ? newItem.replaceAll('"', "").split(",")
        : item.replaceAll('"', "").split(",");
      let name = converted[0] ?? "";
      let serialNumber =
        converted[1] && converted[1].length > 0 ? converted[1] : converted[0];
      let manufacturer = reconvertText(converted[2]) ?? "";
      let modelID = reconvertText(converted[3]) ?? "";
      let os = "";
      let ipAddress = converted[4] ?? "";
      let location = converted[6] ?? "";
      let classType = converted[8] ?? "";
      let assignedTo = "";
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
