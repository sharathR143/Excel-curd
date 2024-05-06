const { app, BrowserWindow, ipcMain } = require("electron/main");
const path = require("node:path");
const fs = require("node:fs");
const XLSX = require("xlsx");

const createWindow = () => {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
    },
  });

  ipcMain.on("get-files", (event, title) => {
    let files = [];
    const directoryPath = path.join(__dirname, "xlData");
    fs.readdirSync(directoryPath).forEach((file, index) => {
      if (file) {
        files.push({
          name: file,
          path: directoryPath,
        });
      }
    });
    win.webContents.send("send-files", files);
  });

  ipcMain.on("get-file", (event, fileName) => {
    const directoryPath = path.join(__dirname, "xlData", fileName);
    const workbook = XLSX.readFile(directoryPath);
    const sheets = workbook.SheetNames;
    let excelData = [];
    sheets.forEach((sheetName) => {
      const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
      });
      excelData.push(jsonData);
    });
    win.webContents.send("send-file", excelData, sheets);
  });

  ipcMain.on("get-updatedData", (event, ExcelData, fileName) => {
    const directoryPath = path.join(__dirname, "xlData", fileName);
    const workbook = XLSX.readFile(directoryPath);
    const sheets = workbook.SheetNames;
    sheets.forEach((sheetName, sheetIndex) => {
      XLSX.utils.sheet_add_aoa(
        workbook.Sheets[sheetName],
        ExcelData[sheetIndex]
      );
    });
    try {
      XLSX.writeFile(workbook, directoryPath);
      win.webContents.send("send-updatedData", "successfully updated..");
    } catch (error) {
      win.webContents.send("send-updatedData", "failed..");
    }
  });
  win.loadFile("index.html");
};

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
