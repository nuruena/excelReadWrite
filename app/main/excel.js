const { ipcMain } = require("electron");
const path = require("path");
const fs = require("fs-extra");
const os = require("os");
const open = require("open");
const Excel = require("exceljs");


//Resolved Excel directory
var time = new Date();
const resolvedExcelDir = path.resolve(os.homedir(), 'electron-app-files')+'\\liquidaciones\\'+ time.toISOString().substring(0, 7);

//Excel dependencies
const reader = require("xlsx");

//Write excel file
exports.excelWriteFile = (filePath, name) => {

  fs.ensureDirSync(resolvedExcelDir);
  
  let fileName = path.resolve(resolvedExcelDir + '\\' + name + " " + time.getDate().toString() + "-" + (time.getMonth() + 1).toString() + "-" + time.getFullYear().toString()) + ".xlsx";

  const workbook = new Excel.Workbook();

  workbook.xlsx.readFile(filePath).then(function() {
    const worksheet = workbook.getWorksheet("Hoja1");
    console.log(worksheet.id);
    //workbook.removeWorksheet(worksheet.id);
    return workbook.xlsx.writeFile(fileName)
  }).catch();

  return resolvedExcelDir;
};
