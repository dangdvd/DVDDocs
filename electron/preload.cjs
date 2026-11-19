const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('dvdDesktop', {
  openFile: () => ipcRenderer.invoke('dvd:open'),
  saveFile: (opts) => ipcRenderer.invoke('dvd:save', opts),
  getLaunchFile: () => ipcRenderer.invoke('dvd:launch-file'),
  onOpenFile: (cb) => {
    const listener = (_e, file) => cb(file);
    ipcRenderer.on('dvd:open-file', listener);
    return () => ipcRenderer.removeListener('dvd:open-file', listener);
  },
});
