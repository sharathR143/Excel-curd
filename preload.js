const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("versions", {
  node: () => process.versions.node,
  chrome: () => process.versions.chrome,
  electron: () => process.versions.electron,

  getFiles: (title) => ipcRenderer.send("get-files", title),
  getFileList: (callback) => ipcRenderer.on("send-files", callback),

  sendFile: (fileName) => ipcRenderer.send("get-file", fileName),
  getFile: (callback) => ipcRenderer.on("send-file", callback),
  // we can also expose variables, not just functions
  sendUpdatedData: (ExcelData, fileName) =>
    ipcRenderer.send("get-updatedData", ExcelData, fileName),

  getUpdatedData: (callback) => ipcRenderer.on("send-updatedData", callback),
  // getUpdatedFailedData: (callback) =>
  //   ipcRenderer.on("send-updatedFailedData", callback),
});
