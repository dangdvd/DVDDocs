const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('dvdDesktop', {
  openFile: () => ipcRenderer.invoke('dvd:open'),
  saveFile: (opts) => ipcRenderer.invoke('dvd:save', opts),
});
