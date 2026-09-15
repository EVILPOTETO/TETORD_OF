const { contextBridge, ipcRenderer } = require('electron');

const storage = {
  readJSON: (relativePath) => ipcRenderer.invoke('tetord:storage-read-json', relativePath),
  writeJSON: (relativePath, value) => ipcRenderer.invoke('tetord:storage-write-json', { relativePath, value }),
  remove: (relativePath) => ipcRenderer.invoke('tetord:storage-remove', relativePath),
  mkdir: (relativePath) => ipcRenderer.invoke('tetord:storage-mkdir', relativePath),
  writeBytes: (relativePath, base64, mime) => ipcRenderer.invoke('tetord:storage-write-bytes', { relativePath, base64, mime }),
  readDataURL: (relativePath) => ipcRenderer.invoke('tetord:storage-read-data-url', relativePath),
  chooseFile: (options) => ipcRenderer.invoke('tetord:storage-choose-file', options || {}),
  list: (relativePath, recursive = true) => ipcRenderer.invoke('tetord:storage-list', { relativePath, recursive }),
  stat: (relativePath) => ipcRenderer.invoke('tetord:storage-stat', relativePath),
  rename: (from, to) => ipcRenderer.invoke('tetord:storage-rename', { from, to }),
  copy: (from, to) => ipcRenderer.invoke('tetord:storage-copy', { from, to })
};

contextBridge.exposeInMainWorld('TETORD_DESKTOP', {
  storage,
  saveFile: (options, base64) => ipcRenderer.invoke('tetord:save-file', { options, base64 }),
  savePdf: (options) => ipcRenderer.invoke('tetord:save-pdf', options)
});
