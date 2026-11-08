// Vỏ ứng dụng máy tính (Windows / macOS / Linux) cho DVDDocs.
const { app, BrowserWindow, dialog, ipcMain, Menu, net, protocol, shell } = require('electron');
const { pathToFileURL } = require('node:url');
const fs = require('node:fs/promises');
const path = require('node:path');

const FILTERS = [
  { name: 'Tài liệu', extensions: ['pdf', 'docx', 'xlsx', 'csv', 'tsv'] },
  { name: 'PDF', extensions: ['pdf'] },
  { name: 'Word', extensions: ['docx'] },
  { name: 'Excel', extensions: ['xlsx', 'csv', 'tsv'] },
];

// Phục vụ giao diện qua giao thức riêng app:// (thay vì file://) để fetch, worker và IndexedDB hoạt động như trên web.
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true } },
]);
const DIST = path.join(__dirname, '..', 'dist');

function toArrayBuffer(buf) {
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 360,
    title: 'DVDDocs',
    icon: path.join(__dirname, '..', 'dist', 'icon.svg'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  // Liên kết ngoài mở bằng trình duyệt hệ thống, không mở trong ứng dụng.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  const devUrl = process.env.DVD_DEV_URL;
  if (devUrl) win.loadURL(devUrl);
  else win.loadURL('app://dvddocs/index.html');
}

ipcMain.handle('dvd:open', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const res = await dialog.showOpenDialog(win, { properties: ['openFile'], filters: FILTERS });
  if (res.canceled || !res.filePaths[0]) return null;
  const p = res.filePaths[0];
  const buf = await fs.readFile(p);
  return { name: path.basename(p), path: p, data: toArrayBuffer(buf) };
});

ipcMain.handle('dvd:save', async (event, { name, path: filePath, data, saveAs }) => {
  let target = filePath;
  if (!target || saveAs) {
    const win = BrowserWindow.fromWebContents(event.sender);
    const ext = path.extname(name).slice(1);
    const res = await dialog.showSaveDialog(win, {
      defaultPath: target || name,
      filters: [{ name: ext.toUpperCase(), extensions: [ext] }],
    });
    if (res.canceled || !res.filePath) return null;
    target = res.filePath;
  }
  await fs.writeFile(target, Buffer.from(data));
  return target;
});

app.whenReady().then(() => {
  protocol.handle('app', (req) => {
    const rel = decodeURIComponent(new URL(req.url).pathname).replace(/^\/+/, '');
    const file = path.normalize(path.join(DIST, rel || 'index.html'));
    if (!file.startsWith(DIST)) return new Response('Not found', { status: 404 });
    return net.fetch(pathToFileURL(file).toString());
  });
  if (process.platform !== 'darwin') Menu.setApplicationMenu(null);
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
