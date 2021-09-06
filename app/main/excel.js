const { ipcMain } = require("electron");
const path = require("path");
const fs = require("fs-extra");
const os = require("os");
const open = require("open");

//Resolved Excel directory
const resolvedExcelDir = path.resolve(os.homedir(), 'electron-app-files')+'\\liquidaciones\\'+ new Date().toISOString().substring(0, 7);

//Excel dependencies
const reader = require("xlsx");

//Read excel file
exports.excelReadFile = (path) => {
  const file = reader.readFile(path);

  let data = [];

  const sheets = file.SheetNames;

  console.clear();

  for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
      data.push(res);
    });
  }

  return data;
};

//Write excel file
exports.excelWriteFile = (data, name) => {

  fs.ensureDirSync(resolvedExcelDir);
  
  let fileName = resolvedExcelDir + '\\' + name + + new Date().toISOString().substring(0, 10) + '.xls';

  try {
    fs.utimesSync(fileName, new Date(), new Date());
  }
  catch (err){
    fs.closeSync(fs.openSync(fileName,'w'));
  }

  let file = reader.readFile(fileName);


  //reader.utils.book_new();
  
  const write = reader.utils.json_to_sheet(data);

  reader.utils.book_append_sheet(file, write, "Sheet1");

  reader.writeFile(file, fileName);

  return resolvedExcelDir;
};
