const { app, BrowserWindow, Menu, shell, dialog, ipcMain } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');

const STORAGE_ROOT_NAME = 'TETORD';
let storageRoot = null;

function getStorageRoot() {
  if (!storageRoot) storageRoot = path.join(app.getPath('documents'), STORAGE_ROOT_NAME);
  return storageRoot;
}

async function ensureStorageRoot() { await fs.mkdir(getStorageRoot(), { recursive: true }); return getStorageRoot(); }

function safeStoragePath(relativePath = '') {
  const root = path.resolve(getStorageRoot());
  const clean = String(relativePath || '').replace(/\\/g, '/').replace(/^\/+/, '');
  if (clean.includes('\0')) throw new Error('Ruta de almacenamiento inválida.');
  if (!clean) return root;
  const target = path.resolve(root, clean);
  if (target !== root && !target.startsWith(`${root}${path.sep}`)) throw new Error('Ruta fuera del almacenamiento de TETORD.');
  return target;
}

async function atomicWrite(target, data) {
  await fs.mkdir(path.dirname(target), { recursive: true });
  const temp = `${target}.tmp-${process.pid}-${Date.now()}`;
  try { await fs.writeFile(temp, data); await fs.rename(temp, target); }
  finally { await fs.rm(temp, { force: true }).catch(() => {}); }
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1440, height: 900, minWidth: 980, minHeight: 650,
    title: 'TETORD', backgroundColor: '#f5f6f4',
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  win.loadFile(path.join(__dirname, '..', 'MiWord', 'index.html'));
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/i.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  return win;
}

ipcMain.handle('tetord:storage-read-json', async (_event, relativePath) => {
  try { const raw = await fs.readFile(safeStoragePath(relativePath), 'utf8'); return { ok: true, value: JSON.parse(raw) }; }
  catch (error) { if (error.code === 'ENOENT') return { ok: false, missing: true }; return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-write-json', async (_event, { relativePath, value }) => {
  try { await ensureStorageRoot(); await atomicWrite(safeStoragePath(relativePath), JSON.stringify(value, null, 2)); return { ok: true, path: relativePath }; }
  catch (error) { return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-remove', async (_event, relativePath) => {
  try { await fs.rm(safeStoragePath(relativePath), { recursive: true, force: true }); return { ok: true }; }
  catch (error) { return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-mkdir', async (_event, relativePath) => {
  try { await fs.mkdir(safeStoragePath(relativePath), { recursive: true }); return { ok: true, path: relativePath }; }
  catch (error) { return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-write-bytes', async (_event, { relativePath, base64 }) => {
  try { await ensureStorageRoot(); await atomicWrite(safeStoragePath(relativePath), Buffer.from(String(base64 || ''), 'base64')); return { ok: true, path: relativePath }; }
  catch (error) { return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-list', async (_event, { relativePath = '', recursive = true } = {}) => {
  try {
    await ensureStorageRoot();
    const root = safeStoragePath(relativePath || '');
    const entries = [];
    async function walk(dir, prefix) {
      const children = await fs.readdir(dir, { withFileTypes: true });
      for (const child of children) {
        const rel = prefix ? `${prefix}/${child.name}` : child.name;
        const target = path.join(dir, child.name);
        const st = await fs.stat(target);
        const item = { name: child.name, relativePath: rel, type: child.isDirectory() ? 'directory' : 'file', size: child.isFile() ? st.size : 0, mtimeMs: st.mtimeMs };
        entries.push(item);
        if (recursive && child.isDirectory()) await walk(target, rel);
      }
    }
    await fs.mkdir(root, { recursive: true });
    await walk(root, '');
    return { ok: true, entries };
  } catch (error) { return { ok: false, error: error.message, entries: [] }; }
});

ipcMain.handle('tetord:storage-stat', async (_event, relativePath) => {
  try { const st = await fs.stat(safeStoragePath(relativePath)); return { ok: true, size: st.size, mtimeMs: st.mtimeMs, isFile: st.isFile(), isDirectory: st.isDirectory() }; }
  catch (error) { return { ok: false, missing: error.code === 'ENOENT', error: error.message }; }
});

ipcMain.handle('tetord:storage-rename', async (_event, { from, to } = {}) => {
  try { await fs.mkdir(path.dirname(safeStoragePath(to)), { recursive: true }); await fs.rename(safeStoragePath(from), safeStoragePath(to)); return { ok: true, path: to }; }
  catch (error) { return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-copy', async (_event, { from, to } = {}) => {
  try { await fs.mkdir(path.dirname(safeStoragePath(to)), { recursive: true }); await fs.cp(safeStoragePath(from), safeStoragePath(to), { recursive: true, errorOnExist: true }); return { ok: true, path: to }; }
  catch (error) { return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-read-data-url', async (_event, relativePath) => {
  try {
    const target = safeStoragePath(relativePath); const data = await fs.readFile(target);
    const ext = path.extname(target).toLowerCase();
    const mime = ({'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.webp':'image/webp','.gif':'image/gif','.svg':'image/svg+xml'})[ext] || 'application/octet-stream';
    return { ok: true, dataUrl: `data:${mime};base64,${data.toString('base64')}` };
  } catch (error) { return { ok: false, error: error.message }; }
});

ipcMain.handle('tetord:storage-choose-file', async (_event, options = {}) => {
  const result = await dialog.showOpenDialog({ title: options.title || 'Seleccionar archivo', properties: ['openFile'], filters: Array.isArray(options.filters) ? options.filters : [{ name: 'Imágenes', extensions: ['png','jpg','jpeg','webp','gif','svg'] }] });
  return result.canceled || !result.filePaths[0] ? { canceled: true } : { canceled: false, path: result.filePaths[0], name: path.basename(result.filePaths[0]) };
});

ipcMain.handle('tetord:save-file', async (_event, payload) => {
  const { options = {}, base64 = '' } = payload || {}; const filters = Array.isArray(options.filters) ? options.filters : []; let filePath = options.filePath || null;
  if (!filePath) { const result = await dialog.showSaveDialog({ title: 'Guardar como', defaultPath: options.suggestedName || 'Mi documento', filters: filters.length ? filters : [{ name: 'Todos los archivos', extensions: ['*'] }] }); if (result.canceled || !result.filePath) return { canceled: true }; filePath = result.filePath; }
  await fs.writeFile(filePath, Buffer.from(base64, 'base64')); return { canceled: false, path: filePath };
});

ipcMain.handle('tetord:save-pdf', async (event, options = {}) => {
  let filePath = options.filePath || null; if (!filePath) { const result = await dialog.showSaveDialog({ title: 'Guardar PDF', defaultPath: options.suggestedName || 'Mi documento.pdf', filters: [{ name: 'Documento PDF', extensions: ['pdf'] }] }); if (result.canceled || !result.filePath) return { canceled: true, saved: false }; filePath = result.filePath; }
  const pdf = await event.sender.printToPDF({ printBackground: true, preferCSSPageSize: true }); await fs.writeFile(filePath, pdf); return { canceled: false, saved: true, path: filePath };
});

app.whenReady().then(async () => { Menu.setApplicationMenu(null); await ensureStorageRoot(); createWindow(); app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); }); });
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
