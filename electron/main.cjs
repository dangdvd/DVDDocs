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

const SUPPORTED = /\.(pdf|docx|xlsx|csv|tsv)$/i;
/** Tệp chờ mở: do nhấp đúp / "Mở bằng" trước khi giao diện sẵn sàng. */
let pendingPath = null;
let mainWindow = null;

/** Lấy đường dẫn tệp tài liệu từ tham số dòng lệnh (Windows, Linux truyền tệp theo cách này). */
function fileFromArgv(argv) {
  return argv.slice(1).find((a) => !a.startsWith('-') && SUPPORTED.test(a)) ?? null;
}

async function readDoc(p) {
  const buf = await fs.readFile(p);
  return { name: path.basename(p), path: p, data: toArrayBuffer(buf) };
}

/** Gửi tệp sang giao diện đang chạy; nếu chưa có cửa sổ thì giữ lại để giao diện hỏi khi khởi động. */
async function openInApp(p) {
  if (!p) return;
  if (!mainWindow || mainWindow.webContents.isLoading()) { pendingPath = p; return; }
  try {
    mainWindow.webContents.send('dvd:open-file', await readDoc(p));
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  } catch (e) {
    dialog.showErrorBox('Không mở được tệp', `${p}\n${e.message}`);
  }
}

// Chỉ chạy một phiên ứng dụng: mở thêm tệp thì gửi sang cửa sổ đang có.
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  pendingPath = fileFromArgv(process.argv);
  app.on('second-instance', (_e, argv) => openInApp(fileFromArgv(argv)));
}
// macOS: mở tệp từ Finder / kéo vào biểu tượng Dock (có thể đến trước khi ứng dụng sẵn sàng).
app.on('open-file', (e, p) => {
  e.preventDefault();
  if (app.isReady()) openInApp(p);
  else pendingPath = p;
});

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
  mainWindow = win;
  win.on('closed', () => { if (mainWindow === win) mainWindow = null; });
  // Còn thay đổi chưa lưu (giao diện chặn beforeunload): hỏi người dùng trước khi đóng cửa sổ.
  win.webContents.on('will-prevent-unload', (event) => {
    const choice = dialog.showMessageBoxSync(win, {
      type: 'warning',
      buttons: ['Đóng không lưu', 'Quay lại'],
      defaultId: 1,
      cancelId: 1,
      title: 'DVDDocs',
      message: 'Tài liệu có thay đổi chưa lưu.',
      detail: 'Nếu đóng bây giờ, thay đổi vẫn được giữ thành bản nháp và có thể khôi phục ở trang chủ lần mở sau.',
    });
    if (choice === 0) event.preventDefault();
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

ipcMain.handle('dvd:launch-file', async () => {
  const p = pendingPath;
  pendingPath = null;
  if (!p) return null;
  try {
    return await readDoc(p);
  } catch (e) {
    dialog.showErrorBox('Không mở được tệp', `${p}\n${e.message}`);
    return null;
  }
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
  if (!gotLock) return;
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
