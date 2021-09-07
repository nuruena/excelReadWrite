const { ipcMain } = require("electron");
const path = require("path");
const fs = require("fs-extra");
const os = require("os");
const open = require("open");
const Excel = require("exceljs");

//Resolved Excel directory
var time = new Date();
const resolvedExcelDir =
  path.resolve(os.homedir(), "electron-app-files") +
  "\\liquidaciones\\" +
  time.toISOString().substring(0, 7);

//Excel dependencies
const reader = require("xlsx");

//Write excel file
exports.excelWriteFile = (filePath, name, extension) => {
  fs.ensureDirSync(resolvedExcelDir);

  let fileName =
    path.resolve(
      resolvedExcelDir +
        "\\" +
        name +
        " " +
        time.getDate().toString() +
        "-" +
        (time.getMonth() + 1).toString() +
        "-" +
        time.getFullYear().toString()
    ) +
    "." +
    extension;
  console.log(filePath);
  console.log(fileName);
  const workbook = new Excel.Workbook();

  workbook.xlsx
    .readFile(filePath)
    .then(function () {
      const worksheet = workbook.getWorksheet("Hoja1") ? workbook.getWorksheet("Hoja1") : workbook.getWorksheet("Sheet1");
      console.log(worksheet.id);
      const colNames = ["CANT", "ISBN", "COD", "TITULO"];
      worksheet.spliceRows(10, 0, colNames);
      var endRow = worksheet.rowCount;
      worksheet.getCell(`A${endRow + 1 }`).value = { formula: `SUM(A11:A${endRow})` };
      worksheet.spliceRows(endRow+2 , 0, ["Totales","","","Sin reposición"])
      worksheet.spliceRows(endRow+3 , 0, ["","","","Facturar a precio viejo"])

      //workbook.removeWorksheet(worksheet.id);
      worksheet.columns = [ { width: 15 }, { width: 20 }, { width: 15 }, { width: 45 }];
      worksheet.columns.forEach(col => {
        col.eachCell(cell => {
        if (cell) {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      })});
      worksheet.spliceColumns(5, 10);
      worksheet.spliceColumns(0, 0);
      return workbook.xlsx.writeFile(fileName);
    })
    .catch();

  return resolvedExcelDir;
};
